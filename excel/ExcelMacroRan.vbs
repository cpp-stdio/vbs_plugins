Function excel_macro_ran(ByVal fileName, ByVal macroName)
    'Running excel VBA
    '
    'Parameters
    '----------
    'fileName : String
    '   Excel file name
    'macroName : String
    '   will run function name or sub name the macro
    '
    'Return
    '----------
    'boolen
    '   success(True) , failure(False)
    excel_macro_ran = False

    'start up excel application
    Dim excelApp :Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    'open file
    Dim excelWorkbook :Set excelWorkbook = excelApp.Workbooks.Open(fileName)

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
        excel_macro_ran = True
    End If
End Function

'test code
'pathLen = len(wscript.scriptfullname) - len(wscript.scriptname)
'parPath = left(wscript.scriptfullname,pathLen)
'Call excel_macro_ran(parPath + "test.xlsm", "main")
'Call excel_macro_ran(parPath + "test.xlsm", "error")
