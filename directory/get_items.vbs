Function get_items(ByVal directoryName, ByVal deeper)
    'Get deep items from it directory.
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
    '   It full path items

    If Right(directoryName,1) = "\" Then
        directoryName = left(directoryName, len(directoryName) - 1)
    End If

    Dim objFso: Set objFso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next

    Dim items
    If objFso.FolderExists(directoryName) = True Then
        items = deep_items(directoryName, deeper, "ALL", objFso)

        If Err.Number <> 0 Then
            WScript.Echo "Error " + Err.Description
        ElseIf IsEmpty(items) = True Then
            WScript.Echo "searched, Not found deep items"
        Else
            WScript.Echo "searched, " + Cstr(UBound(items) + 1) + " items found."
        End if
    Else
        WScript.Echo "Not exist, " + directoryName
    End If

    get_items = items
    Set objFso = Nothing
End Function
