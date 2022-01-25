@echo off

Rem [J] yyyymmddhhmmssŒ`®‚ÅŠÔ‚ğæ“¾‚µ‚Ü‚·B
Rem [E] Get the time in yyyymmddhhmmss format.

Set CURRENT_TIME=%time: =0%
Set CURRENT_TIME=%CURRENT_TIME:~0,2%%CURRENT_TIME:~3,2%%CURRENT_TIME:~6,2%
Set CURRENT_TIME=%date:~0,4%%date:~5,2%%date:~8,2%%CURRENT_TIME%

@echo on
Rem @echo %CURRENT_TIME%
Rem pause
