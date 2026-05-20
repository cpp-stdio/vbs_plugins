Sub file_move(ByVal beforeFileName,ByVal afterFileName)
    ' Moves the specified file to another location.
    ' 指定したファイルを別の場所に移動する。
    '
    ' Parameters / パラメータ
    ' ----------
    ' beforeFileName : String
    '   Source file path.
    '   移動元ファイルのパス。
    ' afterFileName : String
    '   Destination file path.
    '   移動先ファイルのパス。

    if beforeFileName = afterFileName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

    If objFso.FileExists(beforeFileName) Then

        WScript.Echo "Moving: " + beforeFileName
        'Move the file, including read-only files
        Call objFSO.MoveFile(beforeFileName, afterFileName)
        WScript.Echo "Moved to: " + afterFileName
    Else
        WScript.Echo "Source file not found: " + beforeFileName
    End If
    Set objFso = Nothing
End Sub
