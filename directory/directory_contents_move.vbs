Sub directory_contents_move(ByVal beforeDirectoryName,ByVal afterDirectoryName,ByVal deeper,ByVal overwrite)
    ' Moves the contents of a directory to another directory.
    ' ディレクトリの中身を別のディレクトリに移動する。
    '
    ' Parameters / パラメータ
    ' ----------
    ' beforeDirectoryName : String
    '   Source directory path.
    '   移動元ディレクトリのパス。
    ' afterDirectoryName : String
    '   Destination directory path.
    '   移動先ディレクトリのパス。
    ' deeper : Integer
    '   Controls how many levels of subdirectories to include.
    '   何階層まで処理するかを指定する。
    '   [Negative number] : Moves all subdirectories recursively.
    '                       すべてのサブディレクトリを再帰的に移動する。
    '   [       0       ] : Moves only the direct contents (no subdirectories).
    '                       指定ディレクトリ直下の中身のみ移動する（サブディレクトリは含まない）。
    '   [Positive number] : Moves subdirectories up to the specified depth.
    '                       指定した階層数の深さまで移動する。
    ' overwrite : String
    '   Specifies what to do when a file with the same name already exists at the destination.
    '   移動先に同名のファイルが既に存在する場合の処理を指定する。
    '   "YES"     : Overwrite the existing file.
    '               上書きする。
    '   "DELETE"  : Delete the entire destination directory first, then move.
    '               移動先ディレクトリをすべて削除してから移動する。
    '   "ANOTHER" : Move with a different name (e.g., "file(1).txt").
    '               別の名前（例: "file(1).txt"）で移動する。
    '   "NO"      : Skip files that would conflict; do not overwrite.
    '               競合するファイルはスキップして移動しない。

    If beforeDirectoryName = afterDirectoryName Then Exit Sub
    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

    If Right(beforeDirectoryName,1) = "\" Then
        beforeDirectoryName = left(beforeDirectoryName, len(beforeDirectoryName) - 1)
    End If

    If Right(afterDirectoryName,1) = "\" Then
        afterDirectoryName = left(afterDirectoryName, len(afterDirectoryName) - 1)
    End If

    On Error Resume Next
    If objFSO.FolderExists(beforeDirectoryName) <> True Then
        WScript.Echo "Source directory not found: " + beforeDirectoryName
        Set objFSO = Nothing
        Exit Sub
    End If

    If StrComp(overwrite, "DELETE", 1) = 0 Then
        Dim delete_path
        delete_path = afterDirectoryName + Right(beforeDirectoryName,len(beforeDirectoryName) - len(objFSO.GetParentFolderName(beforeDirectoryName)))
        
        If objFSO.FolderExists(delete_path) = True Then 
            Call delete_directory(delete_path)
        End If
        overwrite = "YES"
    End If

    Call deep_move(beforeDirectoryName, afterDirectoryName, deeper, overwrite, objFSO)

    If Err.Number = 0 Then
        WScript.Echo "Moved to: " + afterDirectoryName
    Else
        WScript.Echo "Error: " + Err.Description
    End if

    Set objFSO = Nothing
End Sub
