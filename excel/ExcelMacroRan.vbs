Function excel_macro_ran(ByVal fileName, ByVal macroName)
    ' Executes a VBA macro in the specified Excel file.
    ' 指定した Excel ファイル内の VBA マクロを実行する。
    '
    ' Parameters / パラメータ
    ' ----------
    ' fileName : String
    '   Path of the Excel file to open.
    '   開く Excel ファイルのパス。
    ' macroName : String
    '   Name of the VBA macro (function or subroutine) to execute.
    '   実行する VBA マクロ名（Function または Sub の名前）。
    '
    ' Return / 戻り値
    ' ----------
    ' Boolean
    '   True if the macro executed successfully, False if an error occurred.
    '   マクロが正常に実行された場合は True、エラーが発生した場合は False。
    excel_macro_ran = False

    'Launch the Excel application
    Dim excelApp :Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    'Open the target Excel file
    Dim excelWorkbook :Set excelWorkbook = excelApp.Workbooks.Open(fileName)

    On Error Resume Next
    
    'Execute the specified macro
    WScript.Echo "Executing macro: " + macroName
    Call excelApp.Run(macroName)

    If Err.Number <> 0 Then
        'Leave Excel open so the error can be inspected
        WScript.Echo "Macro failed: " + macroName
    Else
        'Save the file, close it, and quit Excel
        Call excelWorkbook.Save()
        Call excelWorkbook.Close(False)
        excelApp.Workbooks.Close
        excelApp.Quit
        WScript.Echo "Macro executed: " + macroName + " in " + fileName
        excel_macro_ran = True
    End If
End Function

'test code
'pathLen = len(wscript.scriptfullname) - len(wscript.scriptname)
'parPath = left(wscript.scriptfullname,pathLen)
'Call excel_macro_ran(parPath + "test.xlsm", "main")
'Call excel_macro_ran(parPath + "test.xlsm", "error")
