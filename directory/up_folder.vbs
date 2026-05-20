Function up_folder(ByVal folderName)
    ' Returns the path of the parent folder of the specified path.
    ' 指定したパスの 1 つ上の（親）フォルダパスを返す。
    '
    ' Parameters / パラメータ
    ' ----------
    ' folderName : String
    '   The path to get the parent folder of.
    '   親フォルダを取得したいパス。
    '
    ' Return / 戻り値
    ' ----------
    ' String
    '   Full path of the parent folder.
    '   親フォルダのフルパス。

    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    up_folder = objFso.GetParentFolderName(folderName)
    WScript.Echo "Parent folder: " + up_folder
    Set objFso = Nothing
End Function

'------------------------------------------------------------------------------------------------------------------------------
'   test code
'------------------------------------------------------------------------------------------------------------------------------
'thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))
'Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(thisPath + "VBS\__init__.vbs").ReadAll())
'
'Dim target_folder
'target_folder = up_folder(thisPath)
'WScript.Echo target_folder
