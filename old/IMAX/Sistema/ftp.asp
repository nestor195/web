<%@ Language=VBScript %>
<%
' FTP via ASP without using 3rd-party components
' Ben Meghreblian 15th Jan 2002
' benmeg at benmeg dot com / http://benmeg.com/code/asp/ftp.asp.html
'
' This script assumes the file to be FTP'ed is in the same directory as this script.
' It should be obvious how to change this (*hint* change the lcd line).
' You may specify a wildcard in ftp_files_to_put (e.g. *.txt).

' NB: You need to have C:\winnt\system32\wshom.ocx registered to use the WSCRIPT.SHELL object.
' It is registered by default, but is sometimes removed for security reasons (no kidding!).
' You will also need cmd.exe in the path, which again is there, unless the box is locked down.
' Check with your web host/resident sysadmin if in doubt.
'
' NB: This script was originally written in response to a thread on a Wrox ASP mailing list.
' At the time, I was hosting on a shared NT4/IIS4 box and the script worked fine. Since I wrote
' it, several people have got in contact asking why it doesn't work on later versions of either
' Windows or IIS. The answer is probably either as mentioned in the above NB, or to do with
' firewalls restricting outbound traffic from and/or to certain ports. This said, many people
' have successfully used this code to FTP to/from Windows 2000/Windows XP boxes running IIS5/IIS6.
Dim objFSO, objTextFile, oScript, oScriptNet, oFileSys, oFile, strCMD, strTempFile, strCommandResult
Dim ftp_address, ftp_username, ftp_password, ftp_physical_path, ftp_files_to_put

' Edit these variables to match your specifications
ftp_address          = "www.imaxcba.com.ar"
ftp_username         = "ftp.imaxcba.com.ar"
ftp_password         = "123579"
ftp_remote_directory = "web" ' Leave blank if uploading to root directory
ftp_files_to_put     = "file.txt"     ' You can use wildcards here (e.g. *.txt)
On Error Resume Next  '*********************************************************************************
Set oScript = Server.CreateObject("WSCRIPT.SHELL")
Set oFileSys = Server.CreateObject("Scripting.FileSystemObject")
Set objFSO = CreateObject("Scripting.FileSystemObject")
' Build our ftp-commands file
Set objTextFile = objFSO.CreateTextFile(Server.MapPath("test.ftp"))
objTextFile.WriteLine "lcd " & Server.MapPath(".")
objTextFile.WriteLine "open " & ftp_address
objTextFile.WriteLine ftp_username
objTextFile.WriteLine ftp_password

' Check to see if we need to issue a 'cd' command
If ftp_remote_directory <> "" Then
   objTextFile.WriteLine "cd " & ftp_remote_directory
End If

objTextFile.WriteLine "prompt"

' If the file(s) is/are binary (i.e. .jpg, .mdb, etc..), uncomment the following line' objTextFile.WriteLine "binary"
' If there are multiple files to put, we need to use the command 'mput', instead of 'put'
If Instr(1, ftp_files_to_put, "*",1) Then
   objTextFile.WriteLine "mput " & ftp_files_to_put
Else
   objTextFile.WriteLine "put " & ftp_files_to_put
End If
objTextFile.WriteLine "bye"
objTextFile.Close
Set objTextFile = Nothing
' Use cmd.exe to run ftp.exe, parsing our newly created command file
strCMD = "ftp.exe -s:" & Server.MapPath("test.ftp")
strTempFile = "C:\" & oFileSys.GetTempName( )
' Pipe output from cmd.exe to a temporary file (Not :| Steve)
Call oScript.Run ("cmd.exe /c " & strCMD & " > " & strTempFile, 0, True)
Set oFile = oFileSys.OpenTextFile (strTempFile, 1, False, 0)

On Error Resume Next
' Grab output from temporary file
strCommandResult = Server.HTMLEncode( oFile.ReadAll )
oFile.Close
' Delete the temporary & ftp-command files
Call oFileSys.DeleteFile( strTempFile, True )
Call objFSO.DeleteFile( Server.MapPath("test.ftp"), True )
Set oFileSys = Nothing
Set objFSO = Nothing
' Print result of FTP session to screen
Response.Write( Replace( strCommandResult, vbCrLf, "<br>", 1, -1, 1) )
%>