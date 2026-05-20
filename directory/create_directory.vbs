Function create_directory(ByVal directoryName)
    ' Creates a directory recursively.
    ' If any parent directory in the path does not exist, it is also created automatically.
    ' ディレクトリを再帰的に作成する。
    ' パスの途中に存在しない親ディレクトリがあった場合も、自動的に作成される。
    '
    ' Parameters / パラメータ
    ' ----------
    ' directoryName : String
    '   Path of the directory to create.
    '   作成するディレクトリのパス。
    '
    ' Return / 戻り値
    ' ----------
    ' Boolean
    '   True if successful, False if an error occurred.
    '   成功した場合は True、エラーが発生した場合は False。
    create_directory = False

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    
    Call creater(directoryName,objFso)

    If Err.Number = 0 Then
        WScript.Echo "Directory created: " + directoryName
        create_directory = True
    Else
        WScript.Echo "Error: " + Err.Description
    End if
    Set objFso = Nothing
End Function
