Function delete_directory(ByVal directoryName)
    ' Deletes the specified directory along with all its contents (files and subdirectories).
    ' 指定したディレクトリを、その中身（ファイル・サブディレクトリ）ごとすべて削除する。
    '
    ' Parameters / パラメータ
    ' ----------
    ' directoryName : String
    '   Path of the directory to delete.
    '   削除するディレクトリのパス。
    '
    ' Return / 戻り値
    ' ----------
    ' Boolean
    '   True if successful, False if an error occurred.
    '   成功した場合は True、エラーが発生した場合は False。
    delete_directory = False

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    
    'Delete the directory and all its contents (forced, including read-only files and subdirectories)
    Call objFso.DeleteFolder(directoryName, True)

    If Err.Number <> 0 Then
        'Determine and display the reason for failure
        If objFso.FolderExists(directoryName) = True Then
            WScript.Echo "Cannot delete: the directory may be in use or locked."
        Else
            WScript.Echo "Directory not found: " + directoryName
        End if
    Else
        WScript.Echo "Directory deleted: " + directoryName
        delete_directory = True
    End If

    Set objFso = Nothing
End Function
