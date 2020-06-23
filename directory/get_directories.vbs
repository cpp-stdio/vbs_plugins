Function get_directories(ByVal directoryName, ByVal deeper)
    'Get deep directories from it directory.
    '
    'Parameters
    '----------
    'directoryName : String
    '   it directory name
    'deeper : int
    '   [Negative number] It search all directory contents
    '   [       0       ] It search a directory contents
    '   [positive number] It searches directory contents deep hierarchy for the according to number
    '
    'Return
    '----------
    'list
    '   It full path directories

    If Right(directoryName,1) = "\" Then
        directoryName = left(directoryName, len(directoryName) - 1)
    End If

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next

    Dim directories
    If objFso.FolderExists(directoryName) = True Then
        directories = deep_items(directoryName, deeper,"FOLDER", objFso)
        
        If Err.Number <> 0 Then
            WScript.Echo "Error " + Err.Description
        ElseIf IsEmpty(directories) = True Then
            WScript.Echo "searched, Not found deep directories"
        Else
            WScript.Echo "searched, " + Cstr(UBound(directories) + 1) + " directories found."
        End if
    Else
        WScript.Echo "Not exist, " + directoryName
    End If

    get_directories = directories
    Set objFso = Nothing
End Function
