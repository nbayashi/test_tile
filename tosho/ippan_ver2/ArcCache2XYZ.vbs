Option Explicit

Dim fso
Dim folder

Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(".")

CheckZoomLevelFolders(folder)
RenameZoomLevelFolders(folder)

WScript.Echo("フォルダ・ファイル名を{z}/{x}/{y}.png形式に変更しました")

'-------------------------------------------------------------------------------
Function CheckZoomLevelFolders(parentFolder)
	Dim f
	Dim NewFolderName
	For Each f in parentFolder.SubFolders
		If Left(f.Name, 1) = "L" Then
			NewFolderName = CStr(CInt(Right(f.Name, Len(f.Name) - 1)))
			f.Name = NewFolderName
			CheckRowFolders(f)
		End If
	Next
End Function

'-------------------------------------------------------------------------------
Function CheckRowFolders(parentFolder)
	Dim f
	Dim NewFolderName
	For Each f in parentFolder.SubFolders
		If Left(f.Name, 1) = "R" Then
			NewFolderName = CStr(CLng("&H" + Right(f.Name, Len(f.Name) - 1)))
			f.Name = NewFolderName
			CheckImageFiles(f)
		End If
	Next
End Function

'-------------------------------------------------------------------------------
Function CheckImageFiles(parentFolder)
	Dim f
	Dim ext
	Dim baseName, NewFileName
	For Each f in parentFolder.Files
		If Left(f.Name, 1) = "C" Then
			ext = fso.GetExtensionName(f.Path)
			baseName = fso.GetBaseName(f.Path)
			baseName = Right(baseName, Len(baseName) - 1)
			NewFileName = CStr(CLng("&H" + baseName)) + ".png"
			f.Name = NewFileName
		End If
	Next
End Function

'-------------------------------------------------------------------------------
Function RenameZoomLevelFolders(parentFolder)
	Dim zFol, xFol, yFol
	For Each zFol in parentFolder.SubFolders
		For Each xFol in zFol.SubFolders
			RenameImageFiles zFol, xFol
		Next
		For Each xFol in zFol.SubFolders
			If Left(xFol.Name, 1) <> "X" Then
				xFol.Delete
			End If
		Next
		For Each yFol in zFol.SubFolders
			If Left(yFol.Name, 1) = "X" Then
				yFol.Name = Right(yFol.Name, Len(yFol.Name) - 1)
			End If
		Next
	Next
End Function

'-------------------------------------------------------------------------------
Function RenameImageFiles(zFol, xFol)
	Dim f
	Dim ext
	Dim newFolderPath
	Dim NewFilePath

	For Each f in xFol.Files
		newFolderPath = zFol.Path + "\X" + fso.GetBaseName(f.Path)
		if fso.FolderExists(newFolderPath) = False Then
			fso.CreateFolder(newFolderPath)
		End if
		ext = fso.GetExtensionName(f.Path)
		NewFilePath = newFolderPath + "\" + f.parentFolder.Name + "." + ext
		fso.CopyFile f.Path, NewFilePath
	Next
End Function
