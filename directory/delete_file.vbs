Function delete_file(ByVal fileName)
    ' Deletes the specified file.
    ' 指定したファイルを削除する。
    '
    ' Parameters / パラメータ
    ' ----------
    ' fileName : String
    '   Path of the file to delete.
    '   削除するファイルのパス。
    '
    ' Return / 戻り値
    ' ----------
    ' Boolean
    '   True if successful, False if an error occurred.
    '   成功した場合は True、エラーが発生した場合は False。
    delete_file = False

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    
    If objFso.FileExists(fileName) = True Then
        'Delete the file (forced, regardless of read-only attribute)
        Call objFso.DeleteFile(fileName, True)

        If Err.Number <> 0 Then
            WScript.Echo "Cannot delete: the file may be in use or locked."
        Else
            WScript.Echo "File deleted: " + fileName
            delete_file = True
        End if
    Else
        WScript.Echo "File not found: " + fileName
    End If

    Set objFso = Nothing
End Function
