' Setup shell
Set Shell = CreateObject("WScript.Shell")
Set FSO = CreateObject("Scripting.FileSystemObject")

' Get command line arguments
Set paths = Wscript.Arguments
argumentCount = paths.Count

' Check argument for path
If argumentCount > 0 Then
	' Use provided path
	Wscript.Echo "Using " + paths(0)
	path = paths(0)
Else
	' Use current folder
	Wscript.Echo "Using current folder"
	path = "."
End If

' Start looping recursively
Set RootFolder = FSO.GetFolder(path)
LoopFolfer RootFolder
	
Set Shell = Nothing

Sub LoopFolfer(Folder)
	' Loop all files
	For Each File in Folder.Files
		' Check hidden nxl extension
		If(fso.FileExists(File + ".nxl")) Then
			Wscript.Echo "Skipping NXL File: " + File
		' DOCX file
		ElseIf UCase(FSO.GetExtensionName(File)) = "DOCX" Then
			Wscript.Echo File
			Protect File
		' PDF file
		ElseIf UCase(FSO.GetExtensionName(File)) = "PDF" Then
			Wscript.Echo File
			Protect File
		' XLSX file
		ElseIf UCase(FSO.GetExtensionName(File)) = "XLSX" Then
			Wscript.Echo File
			Protect File
		' PPTX file
		ElseIf UCase(FSO.GetExtensionName(File)) = "PPTX" Then
			Wscript.Echo File
			Protect File
		' DWG file
		ElseIf UCase(FSO.GetExtensionName(File)) = "DWG" Then
			Wscript.Echo File
			Protect File
		' TXT file
		ElseIf UCase(FSO.GetExtensionName(File)) = "TXT" Then
			Wscript.Echo File
			Protect File
		' PRT file
		ElseIf UCase(FSO.GetExtensionName(File)) = "PRT" Then
			Wscript.Echo File
			Protect File
		End If
	Next
	
	' Loop subfolders recursively
	For Each Subfolder in Folder.SubFolders
		LoopFolfer SubFolder
	Next
End Sub

' Protect file with Rights Management Client command line interface
Sub Protect(File)
	Shell.Run """C:\Program Files\NextLabs\Rights Management\bin\nxrmconv.exe"" protect """ + File + """ /s /t ip_classification=confidential", 1, True
End Sub

'' EOF ''


