## VBS
This program is a VBscript (VBS) module that allows you to run excel macros,touch directories, and many other things in just one line.
One drawback of this program is that you have to name the repository "VBS", It only one rule will dramatically increase the maintainability and readability of your program.

## Installation
Add a this submodule with git.

## Usage
```vbscript
' get own path
thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))
' Include external modules.
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(thisPath + "VBS\__init__.vbs").ReadAll())
```

