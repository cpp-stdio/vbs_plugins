Function up_folder(ByVal folderName)
    ' Get one upper folder
    '
    'Parameters
    '----------
    ' folderName : String
    '    this is folder name
    '
    'Return
    '----------
    'String
    '   one upper folder name

    Dim objFSO: Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")
    up_folder = objFso.GetParentFolderName(folderName)
    WScript.Echo "one upper folder / " + up_folder
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
