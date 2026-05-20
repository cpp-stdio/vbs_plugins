Dim fso, shell, scriptDir, vbsDir, parentDir, initPath
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

scriptDir = fso.GetParentFolderName(WScript.ScriptFullName)

If LCase(fso.GetFileName(scriptDir)) = "vbs" Then
    vbsDir = scriptDir
ElseIf fso.FolderExists(fso.BuildPath(scriptDir, "VBS")) Then
    vbsDir = fso.BuildPath(scriptDir, "VBS")
Else
    WScript.Echo "ERROR: Could not locate VBS folder."
    WScript.Echo "Run this script from the VBS folder or its parent folder."
    WScript.Quit 1
End If

parentDir = fso.GetParentFolderName(vbsDir)
shell.CurrentDirectory = parentDir

initPath = fso.BuildPath(vbsDir, "__init__.vbs")
Execute(fso.OpenTextFile(initPath).ReadAll())

'------------------------------------------------------------------------
' Character Code Detection Tests
'------------------------------------------------------------------------
WScript.Echo "========== Character Code Detection Tests =========="

Dim samplesDir, utf8File, utf8BomFile, sjisFile, unicodeFile, asciiFile
samplesDir = fso.BuildPath(vbsDir, "samples")

utf8File = fso.BuildPath(samplesDir, "utf8.txt")
utf8BomFile = fso.BuildPath(samplesDir, "utf8_bom.txt")
sjisFile = fso.BuildPath(samplesDir, "shift_jis.txt")
unicodeFile = fso.BuildPath(samplesDir, "unicode.txt")
asciiFile = fso.BuildPath(samplesDir, "ascii.txt")

WScript.Echo ""
WScript.Echo "--- UTF-8 (no BOM) detection ---"
If is_utf8(utf8File) Then
    WScript.Echo "? PASS: " + utf8File + " detected as UTF-8"
Else
    WScript.Echo "? FAIL: " + utf8File + " not detected as UTF-8"
End If

If is_utf8_nobom(utf8File) Then
    WScript.Echo "? PASS: " + utf8File + " detected as UTF-8 (no BOM)"
Else
    WScript.Echo "? FAIL: " + utf8File + " not detected as UTF-8 (no BOM)"
End If

WScript.Echo ""
WScript.Echo "--- UTF-8 (with BOM) detection ---"
If is_utf8_bom(utf8BomFile) Then
    WScript.Echo "? PASS: " + utf8BomFile + " detected as UTF-8 (with BOM)"
Else
    WScript.Echo "? FAIL: " + utf8BomFile + " not detected as UTF-8 (with BOM)"
End If

If is_utf8(utf8BomFile) Then
    WScript.Echo "? PASS: " + utf8BomFile + " detected as UTF-8 (either)"
Else
    WScript.Echo "? FAIL: " + utf8BomFile + " not detected as UTF-8 (either)"
End If

WScript.Echo ""
WScript.Echo "--- Shift_JIS detection ---"
If is_sjis(sjisFile) Then
    WScript.Echo "? PASS: " + sjisFile + " detected as Shift_JIS"
Else
    WScript.Echo "? FAIL: " + sjisFile + " not detected as Shift_JIS"
End If

WScript.Echo ""
WScript.Echo "--- Unicode detection ---"
If is_unicode(unicodeFile) Then
    WScript.Echo "? PASS: " + unicodeFile + " detected as Unicode"
Else
    WScript.Echo "? FAIL: " + unicodeFile + " not detected as Unicode"
End If

WScript.Echo ""
WScript.Echo "--- ASCII detection ---"
If is_ascii(asciiFile) Then
    WScript.Echo "? PASS: " + asciiFile + " detected as ASCII"
Else
    WScript.Echo "? FAIL: " + asciiFile + " not detected as ASCII"
End If

If Not is_ascii(utf8File) Then
    WScript.Echo "? PASS: " + utf8File + " correctly rejected as ASCII"
Else
    WScript.Echo "? FAIL: " + utf8File + " incorrectly detected as ASCII"
End If

If Not is_ascii(utf8BomFile) Then
    WScript.Echo "? PASS: " + utf8BomFile + " correctly rejected as ASCII"
Else
    WScript.Echo "? FAIL: " + utf8BomFile + " incorrectly detected as ASCII"
End If

If Not is_ascii(sjisFile) Then
    WScript.Echo "? PASS: " + sjisFile + " correctly rejected as ASCII"
Else
    WScript.Echo "? FAIL: " + sjisFile + " incorrectly detected as ASCII"
End If

If Not is_ascii(unicodeFile) Then
    WScript.Echo "? PASS: " + unicodeFile + " correctly rejected as ASCII"
Else
    WScript.Echo "? FAIL: " + unicodeFile + " incorrectly detected as ASCII"
End If


WScript.Echo ""
WScript.Echo "========== Test Complete =========="
