## VBS
This program is a VBscript (VBS) module it allows you to run excel macros, touch directories, and many other things in just one line.
So you will dramatically increase the maintainability and readability of your program.

One drawback of this program is it you have to name the repository "VBS". It only one rule.



## Installation
Add a this submodule with git.

## Usage
The first two lines you might wonder about are.
```vbscript
' get own path
thisPath = left(wscript.scriptfullname, len(wscript.scriptfullname) - len(wscript.scriptname))
' Include external modules.
Execute(CreateObject("Scripting.FileSystemObject").OpenTextFile(thisPath + "VBS\__init__.vbs").ReadAll())
```
This program is like an "include" in other languages.
