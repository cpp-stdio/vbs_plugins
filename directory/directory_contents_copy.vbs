Sub directory_contents_copy(ByVal beforeDirectoryName,ByVal afterDirectoryName,ByVal deeper)
    ' Copies the contents of a directory to another directory.
    ' ディレクトリの中身を別のディレクトリにコピーする。
    '
    ' Parameters / パラメータ
    ' ----------
    ' beforeDirectoryName : String
    '   Source directory path.
    '   コピー元ディレクトリのパス。
    ' afterDirectoryName : String
    '   Destination directory path.
    '   コピー先ディレクトリのパス。
    ' deeper : Integer
    '   Controls how many levels of subdirectories to include.
    '   何階層まで処理するかを指定する。
    '   [Negative number] : Copies all subdirectories recursively.
    '                       すべてのサブディレクトリを再帰的にコピーする。
    '   [       0       ] : Copies only the direct contents (no subdirectories).
    '                       指定ディレクトリ直下の中身のみコピーする（サブディレクトリは含まない）。
    '   [Positive number] : Copies subdirectories up to the specified depth.
    '                       指定した階層数の深さまでコピーする。

    If beforeDirectoryName = afterDirectoryName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

    If Right(beforeDirectoryName,1) = "\" Then
        beforeDirectoryName = left(beforeDirectoryName, len(beforeDirectoryName) - 1)
    End If

    If Right(afterDirectoryName,1) = "\" Then
        afterDirectoryName = left(afterDirectoryName, len(afterDirectoryName) - 1)
    End If

    On Error Resume Next
    If objFso.FolderExists(beforeDirectoryName) <> True Then
        WScript.Echo "Source directory not found: " + beforeDirectoryName
        Set objFSO = Nothing
        Exit Sub
    End If

    If objFSO.FolderExists(afterDirectoryName) <> True Then
        Call creater(afterDirectoryName, objFso)
    End If

    Call deep_copy(beforeDirectoryName, afterDirectoryName, deeper, objFSO)

    If Err.Number = 0 Then
        WScript.Echo "Copied to: " + afterDirectoryName
    Else
        WScript.Echo "Error: " + Err.Description
    End if

    Set objFSO = Nothing
End Sub
