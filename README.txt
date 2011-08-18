Motivation
----------
Every once in a while customers run into issues with permissions on their file
servers. Manually analyzing permissions is quite tedious, even when using
standard tools like cacls or xcacls. The output of XCACLS.vbs [1] isn't any
better, but since it's a script, I thought about modifying it to suit my needs.
After taking an actual look at the script I abandoned that thought.

I came across ntfsacls [2], a nice little freeware utility, whose output is
much closer to what I desire. However, it's still not exactly there. For
instance, I'd have to run the command twice to get the permissions on the root
directory and only the non-inherited permissions below the root. I don't find
SDDL [3] all that readable, too. And when using the simplified format, I got
the impression that some items weren't shown accurately. I didn't look further
into that last issue, though, so my impression may be wrong.

Bottom line: I wasn't able to find a tool that would traverse a directory tree
for me, and show me the permissions of the root folder plus all non-inherited
permissions below the root folder in a format that I consider readable. Thus,
after a question over at VisualBasicScript.com [4] got me started, I wrote my
own script for auditing ACLs on files and folders.


Copyright
---------
See COPYING.txt.


Output Format
-------------
The script prints the ACEs of the ACL of a given file or folder either in
simple or full form:

- simple form:
  ? ??? Trustee
  | |   `------- user, group or security principal the ACE applies to
  | `----------- simple permissions
  |                F   = full control
  |                M   = modify
  |                RWX = read, write & execute
  |                RW  = read & write
  |                RX  = read & execute
  |                R   = read
  |                W   = write
  |                (S) = special (permissions do not match any of the above)
  `------------- ACE type
                   A = Allow
                   D = Deny
                   I = Audit

  Examples:
    A F   FOO\Domain Admins                  => admins of domain FOO full
                                                control
    A RX  NT AUTHORITY\Authenticated Users   => allow authenticated users to
                                                read/execute files and to
                                                access/travers folders
    D W   BUILTIN\Guests                     => deny write access for local
                                                group Guests

- full form:
  ? ?????? ?? ?? ??? ???????????????? Trustee
  | |      |  |  |   |                `-- user, group or security principal the
  | |      |  |  |   |                    ACE applies to
  | |      |  |  |   `------------------- inheritance settings
  | |      |  |  |                          (OI) =
  | |      |  |  |                          (CI) =
  | |      |  |  |                          (IO) =
  | |      |  |  |                          (NP) =
  | |      |  |  `----------------------- access to security information
  | |      |  |                             r = read security descriptor and
  | |      |  |                                 ACLs
  | |      |  |                             w = write DAC
  | |      |  |                             o = take ownership
  | |      |  `-------------------------- access to extended attributes
  | |      |                                r = read extended attributes
  | |      |                                w = write extended attributes
  | |      `----------------------------- access to attributes
  | |                                       r = read attributes
  | |                                       w = write attributes
  | `------------------------------------ file/folder permissions
  |                                         r = read data (files)
  |                                             list directory (folders)
  |                                         w = write data (files)
  |                                             create file (folders)
  |                                         a = append data (files)
  |                                             create subfolders (folders)
  |                                         x = execute file (files)
  |                                             traverse folder (folders)
  |                                         d = delete
  |                                         c = delete child
  `-------------------------------------- ACE type
                                            A = Allow
                                            D = Deny
                                            I = Audit

  Examples:
    A rwaxdc rw rw rwo (OI)(CI)         FOO\Domain Admins => give admins of
                                                             domain FOO full
                                                             control
    A -wa--- -- -- ---                  BUILTIN\Users     => allow local users
                                                             to create files
                                                             and subfolders
    A rwaxdc rw rw rwo (OI)(CI)(IO)     CREATOR OWNER     => give full acceass
                                                             to the creator of
                                                             a file/folder
    D W   BUILTIN\Guests            deny write access for group Guests of host BAR


Usage
-----
AuditACLs.vbs [/f] [/i] [/o] [/s] [/r] FILE/FOLDER [FILE/FOLDER ...]
AuditACLs.vbs /?

  /?      Print this help and exit.
  /f      Show security information of files as well (not only folders).
  /i      Show inherited permissions.
  /o      Show owner.
  /r      Recurse into subfolders.
  /s      Show simple permissions.


References
----------
[1] http://support.microsoft.com/kb/825751
[2] http://www.coopware.in2.info/_ntfsacl.htm
[3] http://msdn.microsoft.com/en-us/library/aa379567.aspx
[4] http://www.visualbasicscript.com/
