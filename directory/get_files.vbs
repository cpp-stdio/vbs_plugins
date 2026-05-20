Function get_files(ByVal directoryName, ByVal deeper)
    ' Returns the full paths of all files found within the specified directory.
    ' 指定したディレクトリ内のすべてのファイルのフルパスを配列で返す。
    '
    ' Parameters / パラメータ
    ' ----------
    ' directoryName : String
    '   Path of the directory to search.
    '   検索対象ディレクトリのパス。
    ' deeper : Integer
    '   Controls how many levels of subdirectories to search.
    '   何階層まで検索するかを指定する。
    '   [Negative number] : Searches all subdirectories recursively.
    '                       すべてのサブディレクトリを再帰的に検索する。
    '   [       0       ] : Searches only the specified directory (no subdirectories).
    '                       指定ディレクトリ直下のみ検索する（サブディレクトリは含まない）。
    '   [Positive number] : Searches subdirectories up to the specified depth.
    '                       指定した階層数の深さまで検索する。
    '
    ' Return / 戻り値
    ' ----------
    ' Array
    '   Array of full paths of found files. Empty if none found.
    '   見つかったファイルのフルパスの配列。見つからなかった場合は空。

    If Right(directoryName,1) = "\" Then
        directoryName = left(directoryName, len(directoryName) - 1)
    End If

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next

    Dim files
    If objFso.FolderExists(directoryName) = True Then
        files = deep_items(directoryName, deeper, "FILE", objFso)

        If Err.Number <> 0 Then
            WScript.Echo "Error: " + Err.Description
        ElseIf IsEmpty(files) = True Then
            WScript.Echo "No files found."
        Else
            WScript.Echo "searched, " + Cstr(UBound(files) + 1) + " files found."
        End if
    Else
        WScript.Echo "Directory not found: " + directoryName
    End If

    get_files = files
    Set objFso = Nothing
End Function
