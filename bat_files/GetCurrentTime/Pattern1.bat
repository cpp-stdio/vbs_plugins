@echo off

Rem [J] yyyy/mm/dd hh:mm:sså`éÆÇ≈éûä‘ÇéÊìæÇµÇ‹Ç∑ÅB
Rem [E] Get the time in yyyy/mm/dd hh:mm:ss format.

Set CURRENT_TIME=%time: =0%
Set CURRENT_TIME=%CURRENT_TIME:~0,2%:%CURRENT_TIME:~3,2%:%CURRENT_TIME:~6,2%
Set CURRENT_TIME=%date% %CURRENT_TIME%

@echo on
Rem @echo %CURRENT_TIME%
Rem pause
