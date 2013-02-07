'! A script for auditing permissions, ownership and access monitoring settings
'! of files and folders.
'!
'! @author  Ansgar Wiechers <ansgar.wiechers@planetcobalt.net>
'! @date    2012-05-06
'! @version 1.1.1

' This program is free software; you can redistribute it and/or
' modify it under the terms of the GNU General Public License
' as published by the Free Software Foundation; either version 2
' of the License, or (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program; if not, write to the Free Software
' Foundation, Inc., 59 Temple Place - Suite 330, Boston, MA  02111-1307, USA.

Option Explicit

'! ControlFlags
'! @see <http://msdn.microsoft.com/en-us/library/aa394402.aspx>
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

'! AccessMask flags
'! @see <http://msdn.microsoft.com/en-us/library/aa394063.aspx>
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

'! AceFlags
'! @see <http://msdn.microsoft.com/en-us/library/aa394063.aspx>
Private Const OBJECT_INHERIT_ACE         = &h01
Private Const CONTAINER_INHERIT_ACE      = &h02
Private Const NO_PROPAGATE_INHERIT_ACE   = &h04
Private Const INHERIT_ONLY_ACE           = &h08
Private Const INHERITED_ACE              = &h10
Private Const SUCCESSFUL_ACCESS_ACE_FLAG = &h40
Private Const FAILED_ACCESS_ACE_FLAG     = &h80

'! AceType
'! @see <http://msdn.microsoft.com/en-us/library/aa394063.aspx>
Private Const ACCESS_ALLOWED = 0
Private Const ACCESS_DENIED  = 1
Private Const AUDIT          = 2

'! simple permissions (for simplicity reasons defined as literal values, not as
'! combinations of AccessMask flags)
Private Const FULL_CONTROL       = &h001f01ff
Private Const MODIFY             = &h001301bf
Private Const READ_WRITE_EXECUTE = &h001201bf
Private Const READ_EXECUTE       = &h001200a9
Private Const READ_WRITE         = &h0012019f
Private Const READ_ONLY          = &h00120089
Private Const WRITE_ONLY         = &h00100116

'! The maximum length of a string displaying aceFlags.
Private Const MAXLEN_ACEFLAGS_STR = 16

'! Maximum path length (minus 1 to account for the terminating null character).
'! @see <http://msdn.microsoft.com/en-us/library/aa365247.aspx#maxpath>
Private Const MAX_PATH = 259

'! error constants
Private Const INVALID_PROCEDURE_CALL = 5  '! @see <http://msdn.microsoft.com/en-us/library/fsk1bk1y.aspx>
Private Const PERMISSION_DENIED      = 70 '! @see <http://msdn.microsoft.com/en-us/library/aa264534.aspx>
Private Const PATH_NOT_FOUND         = 76 '! @see <http://msdn.microsoft.com/en-us/library/aa264532.aspx>

'! Wbem impersonation levels
'! @see <http://msdn.microsoft.com/en-us/library/aa393618.aspx>
Private Const WbemImpersonationLevelAnonymous   = 1
Private Const WbemImpersonationLevelIdentify    = 2
Private Const WbemImpersonationLevelImpersonate = 3
Private Const WbemImpersonationLevelDelegate    = 4

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

' Technically SUCCESSFUL_ACCESS_ACE_FLAG and FAILED_ACCESS_ACE_FLAG are not
' ACE types, but ACE flags. However, since they're displayed similar to
' ACCESS_ALLOWED and ACCESS_DENIED types in the GUI, SACLs are easier to
' understand when the respective flags are displayed in the CLI in a similar
' way as the DACL ACE types.
Private aceType : Set aceType = CreateObject("Scripting.Dictionary")
	aceType.Add ACCESS_ALLOWED, "A"
	aceType.Add ACCESS_DENIED , "D"
	aceType.Add AUDIT         , "I"
	aceType.Add SUCCESSFUL_ACCESS_ACE_FLAG, "S"
	aceType.Add FAILED_ACCESS_ACE_FLAG    , "F"

Private simplePermissions : Set simplePermissions = CreateObject("Scripting.Dictionary")
	simplePermissions.Add FULL_CONTROL      , "F  "
	simplePermissions.Add MODIFY            , "M  "
	simplePermissions.Add READ_WRITE_EXECUTE, "RWX"
	simplePermissions.Add READ_EXECUTE      , "RX "
	simplePermissions.Add READ_WRITE        , "RW "
	simplePermissions.Add READ_ONLY         , "R  "
	simplePermissions.Add WRITE_ONLY        , "W  "

' Return codes of the GetSecurityDescriptor() method
' see <http://msdn.microsoft.com/en-us/library/aa390773.aspx>
Private rcDescription : Set rcDescription = CreateObject("Scripting.Dictionary")
	rcDescription.Add 0, "Success"
	rcDescription.Add 2, "Access Denied"
	rcDescription.Add 8, "Unknown Failure"
	rcDescription.Add 9, "Privilege Missing"
	rcDescription.Add 21, "Invalid Parameter"

' global script configuration flags
Private showOwner                 '! Display the owner of each object in the
                                  '! output. Global configuration flag.
Private showSID                   '! Display the SID of a trustee.
Private showName                  '! Display the name of a trustee. If a name
                                  '! cannot be resolved (e.g. the account was
                                  '! deleted, belongs to a different domain,
                                  '! etc.) the SID is displayed instead.
Private showInheritedPermissions  '! Display inherited permissions in the
                                  '! output. Otherwise only non-inherited
                                  '! permissions will be displayed. Global
                                  '! configuration flag.
Private showExtendedPermissions   '! Display extended permissions instead of
                                  '! simple permissions in the output. Global
                                  '! configuration flag.
Private showFiles                 '! Display files in the output. Otherwise
                                  '! only folders will be displayed. Global
                                  '! configuration flag.
Private recurse                   '! Recurse into subdirectories. Otherwise
                                  '! only the given object(s) (files or
                                  '! folders) will be displayed. Global
                                  '! configuration flag.

Private fso : Set fso = CreateObject("Scripting.FileSystemObject")
Private sh  : Set sh  = CreateObject("WScript.Shell")
Private wmiSvc

Main WScript.Arguments


'! Main procedure to handle commandline arguments and start the auditing.
'!
'! @param  args   The commandline arguments passed to the script.
Sub Main(args)
	Dim arg, path, locator

	' check if script is run with cscript.exe, abort otherwise
	If LCase(Right(WScript.FullName, 11)) <> "cscript.exe" Then
		MsgBox WScript.ScriptName & " must be run with cscript.exe!" _
			, vbOKOnly Or vbCritical, WScript.ScriptName
		WScript.Quit 1
	End If

	' evaluate commandline arguments
	If args.Count = 0 Then PrintUsage
	If args.Named.Exists("?") Then PrintUsage
	showOwner = args.Named.Exists("o")
	showSID = args.Named.Exists("s")
	If showSID Then
		showName = args.Named.Exists("n")
	Else
		showName = True
	End If
	showInheritedPermissions = args.Named.Exists("i")
	showExtendedPermissions = args.Named.Exists("e")
	showFiles = args.Named.Exists("f")
	recurse = args.Named.Exists("r")

	' get WMI service and adjust privileges so we can read SACLs
	On Error Resume Next
	Set locator = CreateObject("WbemScripting.SWbemLocator")
	If Err.Number <> 0 Then
		WScript.StdErr.WriteLine "Error creating SWbemLocator object: " _
			& Err.Description & " (0x" & Hex(Err.Number) & ")"
		WScript.Quit 1
	End If
	Set wmiSvc = locator.ConnectServer("", "root/cimv2")
	If Err.Number <> 0 Then
		WScript.StdErr.WriteLine "Error connecting to WMI service: " _
			& Err.Description & " (0x" & Hex(Err.Number) & ")"
		WScript.Quit 1
	End If
	wmiSvc.Security_.ImpersonationLevel = WbemImpersonationLevelImpersonate
	wmiSvc.Security_.Privileges.AddAsString "SeSecurityPrivilege", True
	If Err.Number <> 0 Then
		WScript.StdErr.WriteLine "Error adjusting privileges: " _
			& Err.Description & " (0x" & Hex(Err.Number) & ")"
		WScript.Quit 1
	End If
	Set locator = Nothing
	On Error Goto 0

	For Each arg In args.Unnamed
		path = fso.GetAbsolutePathName(arg)
		If fso.FileExists(path) Then
			' arg is a file
			PrintSecurityInformation fso.GetFile(path), True, "", ""
			WScript.StdOut.WriteLine
		ElseIf fso.FolderExists(path) Then
			' arg is a folder
			PrintSecurityInformation fso.GetFolder(path), True, "", ""
			WScript.StdOut.WriteLine
		Else
			' Print an error if arg doesn't exist and continue.
			WScript.StdErr.WriteLine "File or folder """ & path & """ does not exist."
		End If
	Next
End Sub

'! Print security information on the given object. The object can be either
'! a folder or a file. The amount of information displayed depends on the
'! global flags showOwner, showSID, showInheritedPermissions and
'! showExtendedPermissions.
'!
'! If the object is a folder and the flags recurse or showFiles are set to
'! True, security information on the subfolders or files is printed as well.
'!
'! @param  obj            The folder or file for which security information
'!                        should be printed.
'! @param  showInherited  Boolean value indicting whether or not inherited
'!                        permissions should be displayed. Must be passed as a
'!                        parameter, because inherited permissions of the root
'!                        object should always be displayed.
'! @param  parentPrefix   Indention prefix for the information on the parent
'!                        folder of the current file/folder.
'! @param  myPrefix       Additional prefix for the current file/folder.
Private Sub PrintSecurityInformation(obj, ByVal showInherited, ByVal parentPrefix, ByVal myPrefix)
	Dim record, sd, owner, group, ace
	Dim indentString, i, sf, f, skipFolder
	Dim numFiles, numFolders, newDrive

	skipFolder = False
	record = parentPrefix & myPrefix & obj.Name

	' Adjust the indention string according to whether the current object has
	' subfolders or has files that should be displayed.
	If myPrefix = "+-" Then myPrefix = "| "  ' "connect" subsequent object unless
	If myPrefix = "`-" Then myPrefix = "  "  ' it's the last object (does not have
	                                         ' a successor)
	indentString = parentPrefix & myPrefix & "    "

	If TypeName(obj) = "Folder" Then
		' If the full path of a file or subfolder in the current folder (obj) is
		' longer than MAX_PATH, counting them will fail with error 76 (Path not
		' found).
		' At this point we'll also take care of error 70 (Access denied) when the
		' user running the script does not have access to the contents of a folder
		' (e.g. "System Volume Information").
		On Error Resume Next
		numFiles   = obj.Files.Count
		numFolders = obj.SubFolders.Count
		If Err.Number <> 0 Then
			If Err.Number = PATH_NOT_FOUND Then
				newDrive = Subst(obj.Path)
				If IsNull(newDrive) Then Fail "Critical error. Terminating."
				WScript.StdErr.WriteLine "Warning: path might exceed " & MAX_PATH _
					& " characters. Substituting """ & obj.Path & """ with " & newDrive & "\."
				Err.Clear
				Set obj = fso.GetFolder(newDrive)
				If Err.Number <> 0 Then Fail "Unexpected error: " & Err.Description _
					& " (0x" & Hex(Err.Number) & "). Terminating."
				numFiles   = obj.Files.Count
				numFolders = obj.SubFolders.Count
			ElseIf Err.Number = PERMISSION_DENIED Then
				WScript.StdErr.WriteLine "Cannot enumerate subfolders of """ _
					& obj.Path & """: " & Err.Description & " (" & Err.Number & ")"
				skipFolder = True
			Else
				WScript.StdErr.WriteLine "Unexpected error: " & Err.Description _
					& " (0x" & Hex(Err.Number) & ")"
			End If
		End If
		On Error Goto 0

		record = record & "\" ' append a "\" to folder names
		' "connect" subsequent objects through the ACEs for the current object
		If recurse And ((showFiles And numFiles > 0) Or numFolders > 0) Then
			indentString = parentPrefix & myPrefix & "|   "
		End If
	End If

	Set sd = GetSecurityDescriptor(obj)

	If Not sd Is Nothing Then
		' display owner information
		owner = FormatTrustee(sd.Owner)
		If Not IsNull(sd.Group) Then owner = owner & " (" & FormatTrustee(sd.Group) & ")"
		If showOwner Then record = record & vbTab & owner

		' display DACLs
		If IsSet(sd.ControlFlags, SE_DACL_PRESENT) Then
			If IsNull(sd.DACL) Then
				' Null DACL found. This might be a security problem, because it will
				' grant full access to everyone. See for instance:
				' <http://blogs.technet.com/b/askds/archive/2009/06/02/what-occurs-when-the-security-group-policy-cse-encounters-a-null-dacl.aspx>
				WScript.StdErr.WriteLine "Warning: Null DACL found on file/folder """ & obj.Path & """!"
				record = record & vbNewLine & indentString & "(Null)"
			ElseIf showInherited Or HasNonInheritedACE(sd.DACL) Or Not _
					IsSet(sd.ControlFlags, SE_DACL_AUTO_INHERITED) Then
				For Each ace In sd.DACL
					record = record & vbNewLine & indentString & FormatACE(ace)
				Next
			End If
		End If

		' display SACLs
		If IsSet(sd.ControlFlags, SE_SACL_PRESENT) And Not IsNull(sd.SACL) Then
			For Each ace In sd.SACL
				record = record & vbNewLine & indentString & FormatACE(ace) & " "
			Next
		End If

		' printing the record might fail under some circumstances, e.g. when a
		' filename contains certain Unicode characters
		On Error Resume Next
		WScript.StdOut.WriteLine record
		If Err.Number = INVALID_PROCEDURE_CALL Then
			WScript.StdErr.WriteLine "Cannot display permissions for file """ _
				& Sanitize(obj.Path) & """. " & Err.Description & " (" & Err.Number _
				& ")"
		ElseIf Err.Number <> 0 Then
			WScript.StdErr.WriteLine "Unexpected error: " & Err.Description _
				& " (0x" & Hex(Err.Number) & ")"
		End If
		On Error Goto 0
	Else
		WScript.StdErr.WriteLine "Unable to retrieve security descriptor for """ _
			& obj.Path & """."
		skipFolder = True
	End If

	' If the given object is a folder and recurse is True or showFiles is True,
	' print security information of the contained folders or files respectively.
	If TypeName(obj) = "Folder" And Not skipFolder Then
		If recurse Then
			i = 0
			For Each sf In obj.SubFolders
				i = i + 1
				' When i = obj.SubFolders.Count and no files have to be displayed
				' (either because none should be displayed or because there aren't
				' any), the prefix for the next recursion level is "`-". Otherwise
				' it's "+-" (i.e. there are more folders in this parent folder).
				If i = numFolders And (Not showFiles Or numFiles = 0) Then
					PrintSecurityInformation sf, showInheritedPermissions _
						, parentPrefix & myPrefix, "`-"
				Else
					PrintSecurityInformation sf, showInheritedPermissions _
						, parentPrefix & myPrefix, "+-"
				End If
			Next
		End If

		If showFiles Then
			i = 0
			For Each f In obj.Files
				i = i + 1
				' When i = obj.Files.Count (i.e. the last file in this parent folder
				' is being processed), the prefix for the next recursion level is "`-".
				' Otherwise it's "+-".
				If i = numFiles Then
					PrintSecurityInformation f, showInheritedPermissions _
						, parentPrefix & myPrefix, "`-"
				Else
					PrintSecurityInformation f, showInheritedPermissions _
						, parentPrefix & myPrefix, "+-"
				End If
			Next
		End If
	End If

	If Not IsEmpty(newDrive) Then UnSubst(newDrive)
End Sub

'! Get the security descriptor of the given file or folder.
'!
'! @param  obj  A file or folder object.
'! @return The security descriptor object of the file/folder.
'!
'! @see <http://msdn.microsoft.com/en-us/library/aa390773.aspx> (GetSecurityDescriptor Method of the Win32_LogicalFileSecuritySetting Class)
'! @see <http://msdn.microsoft.com/en-us/library/aa394577.aspx> (WMI Security Descriptor Objects)
Private Function GetSecurityDescriptor(obj)
	Dim path, wmiFileSecSetting, sd, rc

	Set GetSecurityDescriptor = Nothing

	path = obj.Path
	If Not (fso.FileExists(path) Or fso.FolderExists(path)) Then Exit Function

	On Error Resume Next
	Set wmiFileSecSetting = wmiSvc.Get("Win32_LogicalFileSecuritySetting.Path=""" _
		& qq(path) & """")
	If Err.Number <> 0 Then
		WScript.StdErr.WriteLine "Error getting security settings for """ _
			& obj.Path & """: " & Err.Description & " (0x" & Hex(Err.Number) & ")"
	Else
		rc = wmiFileSecSetting.GetSecurityDescriptor(sd)
		If Err.Number = 0 Then
			If rc = 0 Then
				Set GetSecurityDescriptor = sd
			Else
				WScript.StdErr.WriteLine "GetSecurityDescriptor() returned code " _
					& rc & " (" & rcDescription(rc) & ")."
			End If
		Else
			WScript.StdErr.WriteLine "Error getting security descriptor: " _
				& Err.Description & " (0x" & Hex(Err.Number) & ")"
		End If
	End If
	On Error Goto 0
End Function

'! Return a formatted string representing the given ACE. An ACE is presented
'! either in simple or in full form, depending on the global configuration
'! flag "showExtendedPermissions".
'!
'! - Simple:  A RX  Trustee
'! - Full:   +D rwaxdc rw rw rwo (OI)(CI)(IO)     Trustee
'!
'! First comes always the ACE type (A = Allow, D = Deny, S = Success,
'! F = Failure). Last is always the name or SID of the trustee (user, group or
'! security principal) to whom the ACE applies. A leading plus sign indicates
'! that the ACE was not inherited from a parent object.
'!
'! Note: Technically "Success" and "Failure" are not ACE types, but ACE flags.
'! However, since they're displayed similar to "Allow" and "Deny" type in the
'! GUI, I thought that SACLs would be easier to read if those flags were
'! displayed in the CLI in a similar way as the DACL ACE types.
'!
'! If "showExtendedPermissions" is set to False, permissions can be either full
'! control (F), modify (M), read & write & execute (RWX), read & execute (RX),
'! read & write (RW), read-only (R), write-only (W). All other permissions are
'! displayed as "(S)" (special permissions).
'!
'! If "showExtendedPermissions" is set to True, the second group shows the
'! flags for read (r), write (w), append (a), execute (x), delete (d) and
'! delete child (c) permissions. The next two groups specify read/write
'! permissions for attributes and extended attributes respectively. The fourth
'! group shows the rights to read permissions (r), write permissions (w) and
'! take ownership (o), followed by a group with inheritance settings. This
'! fifth group may be empty if permissions to the object are not inheritable).
'!
'! @param  ace  The ACE.
'! @return A string representation of the given ACE.
Private Function FormatACE(ByVal ace)
	Dim inheritance, mask

	If IsInherited(ace) Then
		FormatACE = " "
	Else
		FormatACE = "+"
	End If

	If aceType(ace.AceType) = "I" Then
		FormatACE = FormatACE & aceType(ace.AceFlags And _
			(SUCCESSFUL_ACCESS_ACE_FLAG Or FAILED_ACCESS_ACE_FLAG)) & " "
	Else
		FormatACE = FormatACE & aceType(ace.AceType) & " "
	End If

	If showExtendedPermissions Then
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
	Else
		' When displaying simple permissions, all access masks that aren't exact
		' matches will be displayed as "(S)" (special permissions).
		mask = ace.AccessMask
		' Add SYNCHRONIZE flag to ACEs from SACLs. That way all ACEs can be treated
		' the same way.
		If aceType(ace.AceType) = "I" Then mask = mask Or SYNCHRONIZE
		If simplePermissions.Exists(mask) Then
			FormatACE = FormatACE & simplePermissions(mask)
		Else
			FormatACE = FormatACE & "(S)"
		End If
	End If

	FormatACE = FormatACE & " " & FormatTrustee(ace.Trustee)
End Function

'! Format user and group names. If the domain is an FQDN, the format
'! USER@DOMAIN is used, otherwise the format DOMAIN\USER. If the domain
'! is Null or an empty string, it's omitted. If the name is Null, the
'! SID (in string format) is returned instead.
'!
'! @param  trustee  A Win32_Trustee data structure.
'! @return The formatted user/group name for the given trustee.
'!
'! @see <http://msdn.microsoft.com/en-us/library/aa394501.aspx>
Private Function FormatTrustee(ByVal trustee)
	Dim displaySID

	If IsNull(trustee) Then
		FormatTrustee = ""
	ElseIf Not showName Or IsNull(trustee.Name) Then
		FormatTrustee = trustee.SIDString
	Else
		displaySID = ""
		If showSID Then displaySID = " [" & trustee.SIDString & "]"
		If InStr(trustee.Domain, ".") > 0 Then
			FormatTrustee = trustee.Name & "@" & trustee.Domain & displaySID
		ElseIf IsNull(trustee.Domain) Or trustee.Domain = "" Then
			FormatTrustee = trustee.Name & displaySID
		Else
			FormatTrustee = trustee.Domain & "\" & trustee.Name & displaySID
		End If
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

'! Check if any ACE in the given ACL is not inherited from a parent object.
'!
'! @param  acl  The ACL to inspect.
'! @return True if at least one ACE in the given ACL was not inherited,
'!         otherwise False.
Private Function HasNonInheritedACE(acl)
	Dim ace

	For Each ace In acl
		If Not IsInherited(ace) Then
			HasNonInheritedACE = True
			Exit Function
		End If
	Next

	HasNonInheritedACE = False
End Function

'! Substitute a given folder with a drive letter.
'!
'! @param  folder   The path to the folder.
'! @return Drive-substitution of the folder or Null if the substitution failed.
'!
'! @see UnSubst()
Private Function Subst(folder)
	Dim i, drive, rc

	Subst = Null

	' Get an unused drive letter from the range C..Z.
	For i = 67 To 90
		If Not fso.DriveExists(Chr(i)) Then
			drive = Chr(i) & ":"
			Exit For
		End If
	Next
	If IsEmpty(drive) Then
		WScript.StdErr.WriteLine "Cannot subst """ & folder _
			& """ to a drive letter. No unused drive letters found."
		Exit Function
	End If

	rc = sh.Run("%COMSPEC% /c subst " & drive & " """ & folder & """", 0, True)
	If rc = 0 Then
		Subst = drive
	Else
		WScript.StdErr.WriteLine "External command 'subst' returned error code " _
			& rc & "."
	End If
End Function

'! Delete the subst-mapping to the given drive letter.
'!
'! @param  The drive to un-substitute.
'!
'! @see Subst()
Private Sub UnSubst(drive)
	Dim rc

	rc = sh.Run("%COMSPEC% /c subst " & drive & " /d", 0, True)
	If rc <> 0 Then WScript.StdErr.WriteLine _
		"External command 'subst /d' returned error code " & rc & "."
End Sub

'! Escape characters in the given string for WMI queries.
'!
'! @param  str  The string to escape.
'! @return The escaped string.
Private Function qq(ByVal str)
	qq = Replace(str, "\", "\\")
End Function

'! Print an error message to STDERR and quit.
'!
'! @param  msg  The message to print.
Private Sub Fail(msg)
	WScript.StdErr.WriteLine msg
	WScript.Quit 1
End Sub

'! Sanitize a given string by replacing all but known-good characters with a
'! replacement character ("?").
'!
'! @param  str  The string to sanitize.
'! @return The sanitized string.
Private Function Sanitize(str)
	Dim re

	Set re = New RegExp
	re.Pattern    = "[^a-z0-9_\-.,;:!'+%&#~=∞^@$¢£•Ä©Æ\[\](){}µ‡·‚„‰ÂÁËÈÍÎÏÌÓÔÒÚÛÙıˆ˘˙˚¸ﬂ \\]"
	re.Global     = True
	re.IgnoreCase = True

	Sanitize = re.Replace(str, "?")
End Function

'! Print usage information and exit.
Private Sub PrintUsage
	WScript.Echo "Display security information of a given file, folder or directory tree." & vbNewLine & vbNewLine _
		& "Usage:" & vbTab & WScript.ScriptName & " [/e] [/f] [/i] [/o] [/r] [/s] PATH [PATH ...]" & vbNewLine _
		& vbTab & WScript.ScriptName & " /?" & vbNewLine & vbNewLine _
		& vbTab & "/?" & vbTab & "Print this help and exit." & vbNewLine _
		& vbTab & "/e" & vbTab & "Show extended permissions (default is simple permissions)." & vbNewLine _
		& vbTab & "/f" & vbTab & "Show security information of files as well (not only folders)." & vbNewLine _
		& vbTab & "/i" & vbTab & "Show inherited permissions." & vbNewLine _
		& vbTab & "/n" & vbTab & "Show user/group names (default)." & vbNewLine _
		& vbTab & "/o" & vbTab & "Show owner." & vbNewLine _
		& vbTab & "/r" & vbTab & "Recurse into subfolders." & vbNewLine _
		& vbTab & "/s" & vbTab & "Show SIDs. When used in combination with /n show SIDs alongside names." & vbNewLine & vbNewLine _
		& vbTab & "PATH is the absolute or relative path to a file or folder."
	WScript.Quit 0
End Sub
