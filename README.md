Motivation
==========
Every once in a while customers run into issues with permissions on their file
servers. Manually analyzing permissions is quite tedious, even when using
standard tools like `cacls` or `xcacls`. The output of [`XCACLS.vbs`][1] isn't any
better, but since it's a script, I thought about modifying it to suit my needs.
After taking an actual look at the script I abandoned that thought.

I came across [`ntfsacl`][2], a nice little freeware utility, whose output is
much closer to what I desire. However, it's still not exactly there. For
instance, I'd have to run the command twice to get the permissions on the root
directory and only the non-inherited permissions below the root. I don't find
[SDDL][3] all that readable, too. And when using the simplified format, I got
the impression that for some items the permissions hadn't been displayed
accurately. I didn't look further into that last issue, though, so my
impression may have been wrong there.

Bottom line: I wasn't able to find a tool that would traverse a directory tree
for me, and show me the permissions of the root folder plus all non-inherited
permissions below the root folder in a format that I consider readable. Thus,
after a question over at <strike>visualbasicscript.com</strike> (RIP) got me started, I wrote
my own script for auditing ACLs on files and folders.


Copyright
=========
See COPYING.txt.


Usage
=====

    AuditACLs.vbs [/e] [/f] [/i] [/n] [/o] [/r] [/s] PATH [PATH ...]
    AuditACLs.vbs /?

      /?      Print this help and exit.
      /e      Show extended permissions (default is simple permissions).
      /f      Show security information of files as well (not only folders).
      /i      Show inherited permissions.
      /n      Show user/group names (default).
      /o      Show owner.
      /r      Recurse into subfolders.
      /s      Show SIDs. When used in combination with /n show SIDs
              alongside names.

    PATH is the absolute or relative path to a file or folder.


Output Format
=============
The script prints the ACEs of the ACL of a given file or folder either in
simple or extended form:

- Simple form:

      +? ??? Trustee
      || |   `------- user, group or security principal the ACE applies to
      || `----------- simple permissions
      ||                F   = full control
      ||                M   = modify
      ||                RWX = read, write & execute
      ||                RW  = read & write
      ||                RX  = read & execute
      ||                R   = read
      ||                W   = write
      ||                (S) = special (permissions do not match any of the above)
      |`------------- ACE type
      |                 A = Allow
      |                 D = Deny
      |                 I = Audit
      `-------------- indicator for non-inherited ACE

  Examples:

  - `+A F   FOO\Domain Admins` &rArr; admins of domain FOO full control (not inherited)
  - ` A RX  NT AUTHORITY\Authenticated Users` &rArr; allow authenticated users to read/execute files and to access/travers folders (inherited)
  - `+D W   BUILTIN\Guests [S-1-5-32-546]` &rArr; deny write access for local group Guests (not inherited), displaying name and SID

- Extended form:

      +? ?????? ?? ?? ??? (OI)(CI)(IO)(NP) Trustee
      || |      |  |  |   |                `-- user, group or security principal the
      || |      |  |  |   |                    ACE applies to
      || |      |  |  |   `------------------- inheritance settings
      || |      |  |  |                          (OI) = object inheritance
      || |      |  |  |                          (CI) = container inheritance
      || |      |  |  |                          (IO) = inherit only
      || |      |  |  |                          (NP) = no-propagate inheritance
      || |      |  |  `----------------------- access to security information
      || |      |  |                             r = read security descriptor and
      || |      |  |                                 ACLs
      || |      |  |                             w = write DAC
      || |      |  |                             o = take ownership
      || |      |  `-------------------------- access to extended attributes
      || |      |                                r = read extended attributes
      || |      |                                w = write extended attributes
      || |      `----------------------------- access to attributes
      || |                                       r = read attributes
      || |                                       w = write attributes
      || `------------------------------------ file/folder permissions
      ||                                         r = files:   read data
      ||                                             folders: list directory
      ||                                         w = files:   write data
      ||                                             folders: create file
      ||                                         a = files:   append data
      ||                                             folders: create subfolders
      ||                                         x = files:   execute file
      ||                                             folders: traverse folder
      ||                                         d = delete
      ||                                         c = delete child
      |`-------------------------------------- ACE type
      |                                          A = Allow   (DACL)
      |                                          D = Deny    (DACL)
      |                                          S = Success (SACL)
      |                                          F = Failure (SACL)
      `--------------------------------------- indicator for non-inherited ACE

  Examples:

  - ` A rwaxdc rw rw rwo (OI)(CI)         FOO\Domain Admins` &rArr; give admins of domain FOO full control (inherited)
  - `+A -wa--- -- -- ---                  BUILTIN\Users` &rArr; allow local users to create files and subfolders in the current folder (not inherited)
  - ` A rwaxdc rw rw rwo (OI)(CI)(IO)     CREATOR OWNER` &rArr; give full access to the owner of a file/subfolder
  - `+F -wa--- -w -w ---                  BUILTIN\Guests` &rArr; log failed write access for the local group Guests (not inherited)

[1]: https://support.microsoft.com/kb/825751
[2]: http://www.coopware.in2.info/_ntfsacl.htm
[3]: https://msdn.microsoft.com/en-us/library/aa379567.aspx
