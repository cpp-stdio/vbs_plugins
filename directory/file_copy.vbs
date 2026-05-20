Sub file_copy(ByVal beforeFileName,ByVal afterFileName)
    ' Copies the specified file to another location.
    ' 指定したファイルを別の場所にコピーする。
    '
    ' Parameters / パラメータ
    ' ----------
    ' beforeFileName : String
    '   Source file path.
    '   コピー元ファイルのパス。
    ' afterFileName : String
    '   Destination file path.
    '   コピー先ファイルのパス。

    if beforeFileName = afterFileName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

    If objFso.FileExists(beforeFileName) Then

        WScript.Echo "Copying: " + beforeFileName
        'Copy the file, including read-only files (overwrites the destination if it exists)
        Call objFSO.CopyFile(beforeFileName, afterFileName, True)
        WScript.Echo "Copied to: " + afterFileName
    Else
        WScript.Echo "Source file not found: " + beforeFileName
    End If
    Set objFso = Nothing
End Sub
