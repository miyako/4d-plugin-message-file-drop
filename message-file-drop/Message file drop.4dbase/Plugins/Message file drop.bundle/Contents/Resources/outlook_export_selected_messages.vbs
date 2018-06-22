CRLF = Chr(13) & Chr(10)
On Error Resume Next
	Set objOutlook = GetObject("", "Outlook.Application") 'empty string must be explicit
On Error GoTo 0

If Err.Number = 0 Then

	Set objFSO = CreateObject("Scripting.FileSystemObject")
	Set objShell = WScript.CreateObject("Shell.Application")
	Set objArgs = WScript.Arguments

	If objArgs.Count = 0 Then
		exportFolderPath = objShell.Namespace(0).Self.Path 'Desktop
		'https://msdn.microsoft.com/en-us/library/windows/desktop/bb774096.aspx
		WScript.StdOut.Write "no path specified, default to desktop" & CRLF
	Else
		exportFolderPath = objArgs(0)
	End if

	If Not objFSO.FolderExists(exportFolderPath) Then
		exportFolderPath = objFSO.CreateFolder(exportFolderPath).Path
		WScript.StdOut.Write "creating folder " & exportFolderPath & CRLF
	End If

	If Not Right(Trim(exportFolderPath), 1) = "\" Then
		exportFolderPath = exportFolderPath & "\"
	End If

	WScript.StdOut.Write "export to " & exportFolderPath & CRLF

	Set objSelection = objOutlook.ActiveExplorer().Selection

	For i = 1 To objSelection.Count
		Set selObject = objSelection.Item(i)
		exportPath = exportFolderPath & i & ".mht"
		'WScript.StdOut.Write selObject.Body
		On Error Resume Next
			selObject.SaveAs exportPath, 10 'olMHTML
			'https://msdn.microsoft.com/en-us/VBA/Outlook-VBA/articles/olsaveastype-enumeration-outlook
		On Error GoTo 0
		WScript.StdOut.Write "creating file " & exportPath & CRLF
		WScript.StdOut.Write "result code " & Err.Number & CRLF
	Next

	Set objSelection = Nothing
	Set objArgs = Nothing
	Set objShell = Nothing
	Set objFSO = Nothing
	Set objOutlook = Nothing

End If
