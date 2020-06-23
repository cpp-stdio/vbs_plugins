Function create_directory(ByVal directoryName)
    'Create folders recursively.
    '
    'Parameters
    '----------
    'directoryName : String
    '   directory name to be create
    '
    'Return
    '----------
    'boolen
    '   success(True) , failure(False)
    create_directory = False

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    
    call creater(directoryName,objFso)

    If Err.Number = 0 Then
        WScript.Echo "completed, Creation of " + directoryName
        create_directory = True
    Else
        WScript.Echo "Error " + Err.Description
    End if
    Set objFso = Nothing
End Function
