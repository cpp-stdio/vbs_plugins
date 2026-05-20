Sub directory_copy(ByVal beforeDirectoryName,ByVal afterDirectoryName)
    ' Copies a directory and all its contents to another location.
    ' ディレクトリとその中身をすべて別の場所にコピーする。
    '
    ' Parameters / パラメータ
    ' ----------
    ' beforeDirectoryName : String
    '   Source directory path.
    '   コピー元ディレクトリのパス。
    ' afterDirectoryName : String
    '   Destination directory path.
    '   コピー先ディレクトリのパス。

    if beforeDirectoryName = afterDirectoryName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If objFso.FolderExists(beforeDirectoryName) <> True Then
        WScript.Echo "Source directory not found: " + beforeDirectoryName
        Set objFso = Nothing
        Exit Sub
    End if

    If objFSO.FolderExists(afterDirectoryName) <> True Then
        Call creater(afterDirectoryName, objFso)
    End If

    WScript.Echo "Copying: " + beforeDirectoryName
    'Copy all contents, including read-only files (overwrites existing files at the destination)
    Call objFSO.CopyFolder(beforeDirectoryName, afterDirectoryName, True)
    
    If Err.Number = 0 Then
        WScript.Echo "Copied to: " + afterDirectoryName
    Else
        WScript.Echo "Error: " + Err.Description
    End if
    Set objFso = Nothing
End Sub
