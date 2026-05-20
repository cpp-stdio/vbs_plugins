Sub directory_move(ByVal beforeDirectoryName,ByVal afterDirectoryName)
    ' Moves a directory and all its contents to another location.
    ' ディレクトリとその中身をすべて別の場所に移動する。
    '
    ' Parameters / パラメータ
    ' ----------
    ' beforeDirectoryName : String
    '   Source directory path.
    '   移動元ディレクトリのパス。
    ' afterDirectoryName : String
    '   Destination directory path.
    '   移動先ディレクトリのパス。

    if beforeDirectoryName = afterDirectoryName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    If objFSO.FolderExists(beforeDirectoryName) <> True Then
        WScript.Echo "Source directory not found: " + beforeDirectoryName
        Set objFSO = Nothing
        Exit Sub
    End if

    If objFSO.FolderExists(afterDirectoryName) <> True Then
        Call creater(afterDirectoryName, objFSO)
    End If

    WScript.Echo "Moving: " + beforeDirectoryName
    'Move all contents, including read-only files
    Call objFSO.MoveFolder(beforeDirectoryName, afterDirectoryName)
    
    If Err.Number = 0 Then
        WScript.Echo "Moved to: " + afterDirectoryName
    Else
        WScript.Echo "Error: " + Err.Description
    End if
    Set objFSO = Nothing
End Sub
