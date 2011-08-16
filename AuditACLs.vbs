'! A script for auditing permissions of files and folders.
'!
'! @author  Ansgar Wiechers <ansgar.wiechers@planetcobalt.net>
'! @date    2011-08-16
'! @version 0.9

'! @todo add handling of SACLs

Option Explicit

' ControlFlags
' see <http://msdn.microsoft.com/en-us/library/aa394402.aspx>
Private Const SE_OWNER_DEFAULTED         = &h0001
Private Const SE_GROUP_DEFAULTED         = &h0002
Private Const SE_DACL_PRESENT            = &h0004
Private Const SE_DACL_DEFAULTED          = &h0008
Private Const SE_SACL_PRESENT            = &h0010
Private Const SE_SACL_DEFAULTED          = &h0020
Private Const SE_DACL_AUTO_INHERIT_REQ   = &h0100
Private Const SE_SACL_AUTO_INHERIT_REQ   = &h0200
Private Const SE_DACL_AUTO_INHERITED     = &h0400
Private Const SE_SACL_AUTO_INHERITED     = &h0800
Private Const SE_DACL_PROTECTED          = &h1000
Private Const SE_SACL_PROTECTED          = &h2000
Private Const SE_SELF_RELATIVE           = &h8000

' AccessMask flags
' see <http://msdn.microsoft.com/en-us/library/aa394063.aspx>
Private Const FILE_READ_DATA             = &h000001
Private Const FILE_LIST_DIRECTORY        = &h000001
Private Const FILE_WRITE_DATA            = &h000002
Private Const FILE_ADD_FILE              = &h000002
Private Const FILE_APPEND_DATA           = &h000004
Private Const FILE_ADD_SUBDIRECTORY      = &h000004
Private Const FILE_READ_EA               = &h000008
Private Const FILE_WRITE_EA              = &h000010
Private Const FILE_EXECUTE               = &h000020
Private Const FILE_TRAVERSE              = &h000020
Private Const FILE_DELETE_CHILD          = &h000040
Private Const FILE_READ_ATTRIBUTES       = &h000080
Private Const FILE_WRITE_ATTRIBUTES      = &h000100
Private Const DELETE                     = &h010000
Private Const READ_CONTROL               = &h020000
Private Const WRITE_DAC                  = &h040000
Private Const WRITE_OWNER                = &h080000
Private Const SYNCHRONIZE                = &h100000

' AceFlags
' see <http://msdn.microsoft.com/en-us/library/aa394063.aspx>
Private Const OBJECT_INHERIT_ACE         = &h01
Private Const CONTAINER_INHERIT_ACE      = &h02
Private Const NO_PROPAGATE_INHERIT_ACE   = &h04
Private Const INHERIT_ONLY_ACE           = &h08
Private Const INHERITED_ACE              = &h10
Private Const SUCCESSFUL_ACCESS_ACE_FLAG = &h40
Private Const FAILED_ACCESS_ACE_FLAG     = &h80

' AceType
' see <http://msdn.microsoft.com/en-us/library/aa394063.aspx>
Private Const ACCESS_ALLOWED = 0
Private Const ACCESS_DENIED  = 1
Private Const AUDIT          = 2

' simple permissions (for simplicity reasons defined as literal values, not as
' combinations of AccessMask flags)
Private Const FULL_CONTROL       = &h001f01ff
Private Const MODIFY             = &h001301bf
Private Const READ_WRITE_EXECUTE = &h001201bf
Private Const READ_EXECUTE       = &h001200a9
Private Const READ_WRITE         = &h0012019f
Private Const READ_ONLY          = &h00120089
Private Const WRITE_ONLY         = &h00100116

' maximum length of a string displaying aceFlags
Private Const MAXLEN_ACEFLAGS_STR = 16

Private accessMask : Set accessMask = CreateObject("Scripting.Dictionary")
	accessMask.Add 0                     , "-"
	accessMask.Add FILE_READ_DATA        , "r"
	accessMask.Add FILE_WRITE_DATA       , "w"
	accessMask.Add FILE_APPEND_DATA      , "a"
	accessMask.Add FILE_READ_EA          , "r"
	accessMask.Add FILE_WRITE_EA         , "w"
	accessMask.Add FILE_EXECUTE          , "x"
	accessMask.Add FILE_DELETE_CHILD     , "c"
	accessMask.Add FILE_READ_ATTRIBUTES  , "r"
	accessMask.Add FILE_WRITE_ATTRIBUTES , "w"
	accessMask.Add DELETE                , "d"
	accessMask.Add READ_CONTROL          , "r"
	accessMask.Add WRITE_DAC             , "w"
	accessMask.Add WRITE_OWNER           , "o"
	accessMask.Add SYNCHRONIZE           , "s"

Private aceFlags : Set aceFlags = CreateObject("Scripting.Dictionary")
	aceFlags.Add 0                       , ""
	aceFlags.Add OBJECT_INHERIT_ACE      , "(OI)"
	aceFlags.Add CONTAINER_INHERIT_ACE   , "(CI)"
	aceFlags.Add NO_PROPAGATE_INHERIT_ACE, "(NP)"
	aceFlags.Add INHERIT_ONLY_ACE        , "(IO)"
	aceFlags.Add INHERITED_ACE           , "(I)"

Private aceType : Set aceType = CreateObject("Scripting.Dictionary")
	aceType.Add ACCESS_ALLOWED, "A"
	aceType.Add ACCESS_DENIED , "D"
	aceType.Add AUDIT         , "I"

Private simplePermissions : Set simplePermissions = CreateObject("Scripting.Dictionary")
	simplePermissions.Add FULL_CONTROL      , "F  "
	simplePermissions.Add MODIFY            , "M  "
	simplePermissions.Add READ_WRITE_EXECUTE, "RWX"
	simplePermissions.Add READ_EXECUTE      , "RX "
	simplePermissions.Add READ_WRITE        , "RW "
	simplePermissions.Add READ_ONLY         , "R  "
	simplePermissions.Add WRITE_ONLY        , "W  "

' global configuration flags
Private showOwner, showInheritedPermissions, showSimplePermissions, showFiles, recurse

Private fso : Set fso = CreateObject("Scripting.FileSystemObject")

Main WScript.Arguments

'! Main procedure to handle commandline arguments and start the auditing.
'!
'! @param  args   The commandline arguments passed to the script.
Sub Main(args)
	Dim arg, path

	' check if script is run with cscript.exe, abort otherwise
	If LCase(Right(WScript.FullName, 11)) <> "cscript.exe" Then
		MsgBox WScript.ScriptName & " must be run with cscript.exe!" _
			, vbOKOnly Or vbCritical, WScript.ScriptName
		WScript.Quit 1
	End If

	' evaluate commandline options
	If args.Named.Exists("?") Then PrintUsage
	showOwner = args.Named.Exists("o")
	showInheritedPermissions = args.Named.Exists("i")
	showSimplePermissions = args.Named.Exists("s")
	showFiles = args.Named.Exists("f")
	recurse = args.Named.Exists("r")

	For Each arg In args.Unnamed
		path = fso.GetAbsolutePathName(arg)
		If fso.FileExists(path) Then
			' arg is a file
			PrintSecurityInformation fso.GetFile(path), "", ""
		ElseIf fso.FolderExists(path) Then
			' arg is a folder
			PrintSecurityInformation fso.GetFolder(path), "", ""
		Else
			' Print an error if arg doesn't exist and continue.
			WScript.StdErr.WriteLine "File or folder '" & path & "' does not exist."
		End If
	Next
End Sub

'! Print security information on the given object. The object can be either a
'! folder or a file. The amount of information displayed depends on the global
'! flags showOwner, showInheritedPermissions and showSimplePermissions.
'!
'! If the object is a folder and the flags recurse or showFiles are set to
'! True, security information on the subfolders or files is printed as well.
'!
'! @param  obj          The folder or file for which security information
'!                      should be printed.
'! @param  parentPrefix Indention prefix for the information on the parent
'!                      folder of the current file/folder.
'! @param  myPrefix     Additional prefix for the current file/folder.
Private Sub PrintSecurityInformation(obj, ByVal parentPrefix, ByVal myPrefix)
	Dim record, sd, owner, group, ace
	Dim indentString, i, sf, f

	record = parentPrefix & myPrefix & obj.Name

	' Adjust the indention string according to whether the current object has
	' subfolders or has files that should be displayed.
	If myPrefix = "+-" Then myPrefix = "| " ' "connect" subsequent object unless
	If myPrefix = "`-" Then myPrefix = "  " ' it's the last object (which has no successor)
	indentString = parentPrefix & myPrefix & "    "

	If TypeName(obj) = "Folder" Then
		record = record & "\" ' append a "\" to folder names
		' "connect" subsequent objects through the ACEs for the current object
		If (showFiles And obj.Files.Count > 0) Or obj.SubFolders.Count > 0 Then
			indentString = parentPrefix & myPrefix & "|   "
		End If
	End If

	Set sd = GetSecurityDescriptor(obj.Path)

	' display owner information
	owner = FormatTrustee(sd.Owner)
	If Not IsNull(sd.Group) Then owner = owner & " (" & FormatTrustee(sd.Group) & ")"
	If showOwner Then record = record & vbTab & owner

	' display DACLs
	If IsSet(sd.ControlFlags, SE_DACL_PRESENT) Then
		For Each ace In sd.DACL
			If showInheritedPermissions Or Not IsInherited(ace) Then record = record _
				& vbNewLine & indentString & FormatACE(ace)
		Next
	End If

	' Display SACLs
	'~ If IsSet(sd.ControlFlags, SE_SACL_PRESENT) Then
		'~ For Each ace In sd.SACL
			'~ ' do stuff
		'~ Next
	'~ End If

	WScript.StdOut.WriteLine record

	' If the given object is a folder and recurse is True or showFiles is True,
	' print security information of the contained folders or files respectively.
	If TypeName(obj) = "Folder" Then
		If recurse Then
			i = 0
			For Each sf In obj.SubFolders
				i = i + 1
				If i = obj.SubFolders.Count And (Not showFiles Or obj.Files.Count = 0) Then
					PrintSecurityInformation sf, parentPrefix & myPrefix, "`-"
				Else
					PrintSecurityInformation sf, parentPrefix & myPrefix, "+-"
				End If
			Next
		End If

		If showFiles Then
			i = 0
			For Each f In obj.Files
				i = i + 1
				If i = obj.Files.Count Then
					PrintSecurityInformation f, parentPrefix & myPrefix, "`-"
				Else
					PrintSecurityInformation f, parentPrefix & myPrefix, "+-"
				End If
			Next
		End If
	End If
End Sub

'! Get the security descriptor of the given file or folder. The function does
'! NOT check whether the object exists, so the caller MUST ensure this.
'!
'! @param  path   A relative or absolute path to a file or folder.
'! @return The security descriptor object of the file/folder.
'!
'! @see <http://msdn.microsoft.com/en-us/library/aa390773.aspx> (GetSecurityDescriptor Method of the Win32_LogicalFileSecuritySetting Class)
'! @see <http://msdn.microsoft.com/en-us/library/aa394577.aspx> (WMI Security Descriptor Objects)
Private Function GetSecurityDescriptor(ByVal path)
	Dim wmiFileSecSetting, wmiSD

	Set GetSecurityDescriptor = Nothing

	path = fso.GetAbsolutePathName(path)
	If fso.FileExists(path) Or fso.FolderExists(path) Then
		On Error Resume Next
		Set wmiFileSecSetting = GetObject("winmgmts:Win32_LogicalFileSecuritySetting.path='" _
			& Replace(path, "\", "\\") & "'")
		wmiFileSecSetting.GetSecurityDescriptor wmiSD
		If Err.Number = 0 Then
			Set GetSecurityDescriptor = wmiSD
		Else
			WScript.StdErr.WriteLine Err.Description & " (" & Hex(Err.Number) & ")"
		End If
		On Error Goto 0
	End If
End Function

'! Return a formatted string representing the given ACE. An ACE is presented
'! either in simple or in full form, depending on the showSimple paramenter.
'!
'! - Simple: A RX  Trustee
'! - Full:   D rwaxdc rw rw rwo [(OI)(CI)(IO)(NP)] Trustee
'!
'! First is always the ACE type (A = Allow, D = Deny, I = Audit). Last is
'! always the name of the trustee (user, group or security principal) to whom
'! the ACE applies.
'!
'! When showSimplePermissions is set to True, permissions can be either full
'! control (F), modify (M), read & write & execute (RWX), read & execute (RX),
'! read & write (RW), read-only (R), write-only (W). All other permissions are
'! displayed as "(S)" (special permissions).
'!
'! When showSimple is False, the second group shows the flags for read (r),
'! write (w), append (a), execute (x), delete (d) and delete child (c)
'! permissions. The next two groups specify read/write permissions for
'! attributes and extended attributes respectively. The fourth group shows
'! the rights to read permissions (r), write permissions (w) and take
'! ownership (o), followed by a group with inheritance settings. This fifth
'! group may be empty if permissions to the object are neither inherited nor
'! inheritable).
'!
'! @param  ace        The ACE.
'! @return A string representation of the given ACE.
Private Function FormatACE(ByVal ace)
	Dim inheritance

	FormatACE = aceType(ace.AceType) & " "
	If showSimplePermissions Then
		' When displaying simple permissions, all access masks that aren't exact
		' matches will be displayed as "(S)" (special permissions).
		If simplePermissions.Exists(ace.AccessMask) Then
			FormatACE = FormatACE & simplePermissions(ace.AccessMask)
		Else
			FormatACE = FormatACE & "(S)"
		End If
	Else
		' Display full permissions and inheritance settings.
		inheritance = aceFlags(ace.AceFlags And OBJECT_INHERIT_ACE) _
			& aceFlags(ace.AceFlags And CONTAINER_INHERIT_ACE) _
			& aceFlags(ace.AceFlags And INHERIT_ONLY_ACE) _
			& aceFlags(ace.AceFlags And NO_PROPAGATE_INHERIT_ACE)

		FormatACE = FormatACE _
			& accessMask(ace.AccessMask And FILE_READ_DATA) _
			& accessMask(ace.AccessMask And FILE_WRITE_DATA) _
			& accessMask(ace.AccessMask And FILE_APPEND_DATA) _
			& accessMask(ace.AccessMask And FILE_EXECUTE) _
			& accessMask(ace.AccessMask And DELETE) _
			& accessMask(ace.AccessMask And FILE_DELETE_CHILD) & " " _
			& accessMask(ace.AccessMask And FILE_READ_ATTRIBUTES) _
			& accessMask(ace.AccessMask And FILE_WRITE_ATTRIBUTES) & " " _
			& accessMask(ace.AccessMask And FILE_READ_EA) _
			& accessMask(ace.AccessMask And FILE_WRITE_EA) & " " _
			& accessMask(ace.AccessMask And READ_CONTROL) _
			& accessMask(ace.AccessMask And WRITE_DAC) _
			& accessMask(ace.AccessMask And WRITE_OWNER) & " " _
			& inheritance & String(MAXLEN_ACEFLAGS_STR - Len(inheritance), " ")
	End If
	FormatACE = FormatACE & " " & FormatTrustee(ace.Trustee)
End Function

'! Format user and group names. If the domain is an FQDN, the format
'! USER@DOMAIN is used, otherwise the format DOMAIN\USER.
'!
'! @param  trustee  A Win32_Trustee data structure.
'! @return The formatted user/group name for the given trustee.
'!
'! @see <http://msdn.microsoft.com/en-us/library/aa394501.aspx>
Private Function FormatTrustee(ByVal trustee)
	If InStr(trustee.Domain, ".") > 0 Then
		FormatTrustee = trustee.Name & "@" & trustee.Domain
	ElseIf IsNull(trustee.Domain) Or trustee.Domain = "" Then
		FormatTrustee = trustee.Name
	Else
		FormatTrustee = trustee.Domain & "\" & trustee.Name
	End If
End Function

'! Check if a flag is set in the given value.
'!
'! @param  val  An integer value.
'! @param  flag The flag to check.
'! @return True if the flag is set, otherwise False.
Private Function IsSet(ByVal val, ByVal flag)
	IsSet = ((val And flag) = flag)
End Function

'! Check if the given ACE is inherited from a parent object.
'!
'! @param  ace  The ACE.
'! @return True if the ACE is inherited, otherwise False.
Private Function IsInherited(ByVal ace)
	IsInherited = IsSet(ace.AceFlags, INHERITED_ACE)
End Function

'! Print usage information and exit.
Private Sub PrintUsage
	WScript.Echo "Display security information of a given file, folder or directory tree." & vbNewLine & vbNewLine _
		& "Usage:" & vbTab & WScript.ScriptName & " [/f] [/i] [/o] [/s] [/r] FILE/FOLDER [FILE/FOLDER ...]" & vbNewLine _
		& vbTab & WScript.ScriptName & " /?" & vbNewLine & vbNewLine _
		& vbTab & "/?" & vbTab & "Print this help and exit." & vbNewLine _
		& vbTab & "/f" & vbTab & "Show security information of files as well (not only folders)." & vbNewLine _
		& vbTab & "/i" & vbTab & "Show inherited permissions." & vbNewLine _
		& vbTab & "/o" & vbTab & "Show owner." & vbNewLine _
		& vbTab & "/r" & vbTab & "Recurse into subfolders." & vbNewLine _
		& vbTab & "/s" & vbTab & "Show simple permissions."
	WScript.Quit 0
End Sub
