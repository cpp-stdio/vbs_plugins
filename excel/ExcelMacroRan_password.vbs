Function excel_macro_ran_password(ByVal fileName, ByVal macroName, ByVal pass, ByVal writeResPass)
    ' Executes a VBA macro in the specified password-protected Excel file.
    ' パスワード付き Excel ファイル内の VBA マクロを実行する。
    '
    ' Parameters / パラメータ
    ' ----------
    ' fileName : String
    '   Path of the Excel file to open.
    '   開く Excel ファイルのパス。
    ' macroName : String
    '   Name of the VBA macro (function or subroutine) to execute.
    '   実行する VBA マクロ名（Function または Sub の名前）。
    ' pass : String
    '   Password required to open the Excel file. Pass Nothing if no password is set.
    '   Excel ファイルを開くためのパスワード。パスワードがない場合は Nothing を渡す。
    ' writeResPass : String
    '   Password required to edit the Excel file. Pass Nothing if no password is set.
    '   Excel ファイルを編集するためのパスワード。パスワードがない場合は Nothing を渡す。
    '
    ' Return / 戻り値
    ' ----------
    ' Boolean
    '   True if the macro executed successfully, False if an error occurred.
    '   マクロが正常に実行された場合は True、エラーが発生した場合は False。
    excel_macro_ran_password = False

    'Launch the Excel application
    Dim excelApp :Set excelApp = CreateObject("Excel.Application")
    excelApp.Visible = True
    'Open the target Excel file
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
