Function get_items(ByVal directoryName, ByVal deeper)
    ' Returns the full paths of all files and directories found within the specified directory.
    ' 指定したディレクトリ内のすべてのファイルとサブディレクトリのフルパスを配列で返す。
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
    '   Array of full paths of found files and directories. Empty if none found.
    '   見つかったファイルとディレクトリのフルパスの配列。見つからなかった場合は空。

    If Right(directoryName,1) = "\" Then
        directoryName = left(directoryName, len(directoryName) - 1)
    End If

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next

    Dim items
    If objFso.FolderExists(directoryName) = True Then
        items = deep_items(directoryName, deeper, "ALL", objFso)

        If Err.Number <> 0 Then
            WScript.Echo "Error: " + Err.Description
        ElseIf IsEmpty(items) = True Then
            WScript.Echo "No items found."
        Else
            WScript.Echo "searched, " + Cstr(UBound(items) + 1) + " items found."
        End if
    Else
        WScript.Echo "Directory not found: " + directoryName
    End If

    get_items = items
    Set objFso = Nothing
End Function
