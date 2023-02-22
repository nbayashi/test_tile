Option Explicit

Dim fso
Dim folder

Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder(".")

CheckZoomLevelFolders(folder)

WScript.Echo("フォルダ・ファイル名を{z}/{y}/{x}.png形式に変更しました")

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
			NewFileName = CStr(CInt("&H" + baseName)) + ".png"
			f.Name = NewFileName
		End If
	Next
End Function
