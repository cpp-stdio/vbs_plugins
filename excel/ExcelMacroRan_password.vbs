Function excel_macro_ran_password(ByVal fileName, ByVal macroName, ByVal pass, ByVal writeResPass)
    'Running excel VBA
    '
    'Parameters
    '----------
    'fileName : String
    '   Excel file name
    'macroName : String
    '   will run function name or sub name the macro
    'pass : String
    '   Password needed to open excel. If it is empty, insert "Nothing".
    'writeResPass : String
    '   Password needed to read excel. If it is empty, insert "Nothing".
    '
    'Return
    '----------
    'boolen
    '   success(True) , failure(False)
    excel_macro_ran_password = False

    'start up excel application
    Dim excelApp :Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    'open file
    Dim excelWorkbook
    const vbString = 8
    If VarType(pass) = vbString And VarType(writeResPass) = vbString Then
        Set excelWorkbook = excelApp.Workbooks.Open(fileName,,,,pass,writeResPass,True)
    ElseIf VarType(pass) = vbString Then
        Set excelWorkbook = excelApp.Workbooks.Open(fileName,,,,pass,,True)
    ElseIf VarType(writeResPass) = vbString Then
        Set excelWorkbook = excelApp.Workbooks.Open(fileName,,,,,writeResPass,True)
    Else
        Set excelWorkbook = excelApp.Workbooks.Open(fileName)
    End If
    
    On Error Resume Next
    'start VBA
    WScript.Echo "Ran " + macroName
    Call excelApp.Run(macroName)

    If Err.Number <> 0 Then
        'We will leave the Excel without closing for the review of VBA.
        WScript.Echo "Error : " + macroName
    Else
        'Exit the Excel application.
        Call excelWorkbook.Save()
        Call excelWorkbook.Close(False)
        excelApp.Workbooks.Close
        excelApp.Quit
        WScript.Echo fileName + " of " + macroName + " was executed."
        excel_macro_ran_password = True
    End If
    Set excelWorkbook = Nothing
    Set excelApp = Nothing
End Function

'test code
'pathLen = len(wscript.scriptfullname) - len(wscript.scriptname)
'parPath = left(wscript.scriptfullname,pathLen)
'Call excel_macro_ran_password(parPath + "test.xlsm", "main", "AW", Nothing)
'Call excel_macro_ran_password(parPath + "test.xlsm", "error", Nothing, "AW")
'Call excel_macro_ran_password(parPath + "test.xlsm", "main", "AW", "AW")