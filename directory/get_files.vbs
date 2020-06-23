Function get_files(ByVal directoryName, ByVal deeper)
    'Get deep files from it directory.
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
    '   It full path files

    If Right(directoryName,1) = "\" Then
        directoryName = left(directoryName, len(directoryName) - 1)
    End If

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next

    Dim files
    If objFso.FolderExists(directoryName) = True Then
        files = deep_items(directoryName, deeper, "FILE", objFso)

        If Err.Number <> 0 Then
            WScript.Echo "Error " + Err.Description
        ElseIf IsEmpty(files) = True Then
            WScript.Echo "searched, Not found deep files"
        Else
            WScript.Echo "searched, " + Cstr(UBound(files) + 1) + " files found."
        End if
    Else
        WScript.Echo "Not exist, " + directoryName
    End If

    get_files = files
    Set objFso = Nothing
End Function
