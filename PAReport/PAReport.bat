REM Â«ÏÌÝè
set ssUser=fsol07
set ssPwd=fsol07
set ssDir=\\ADSDEAI1\vss
set targetDir=D:\data\VssWork\04.Rrsdo\04.or`osHö\81.i»Ç
set targetPrj=$/04.Rrsdo/04.or`osHö/81.i»Ç/
set reviewDir1="X:\52_·úvæ\04.Rrsdo\04.or`osHö\85.Ztr[\oEx¥E¬Êx[g"
set reviewDir2="X:\52_·úvæ\04.Rrsdo\04.or`osHö\84.r["
set BaseDir=%~dp0

REM ^CX^vÌútðßé
for /F "tokens=1" %%a in ('date /t') do set orgdate=%%a
set yy=%orgdate:~0,4%
set mm=%orgdate:~5,2%
set dd=%orgdate:~8,2%
set today="%BaseDir%%yy%%mm%%dd%"

REM vWFNgfBNgðÄì¬
rmdir %targetDir% /S /Q
mkdir %targetDir%

REM útfBNgðÄì¬
rmdir %today% /S /Q
mkdir %today%

REM VSSæèÅVÅðæ¾
ss Get %targetPrj% -R  -I-Y -GL%targetDir% -GWR -GTU -GCK

REM i»Çt@CðRs[
mkdir %today%\i»Ç
robocopy %targetDir% %today%\i»Ç *.xls /E

REM r[\PðRs[
mkdir %today%\r[\1
robocopy %reviewDir1% %today%\r[\1 àr[L^[*.xls /E

REM r[\QðRs[
mkdir %today%\r[\2
robocopy %reviewDir2% %today%\r[\2 PTr[Ç[*_ox¥.xls /E

REM ¿ð
PAReport1.exe %today%\i»Ç %today%\i»Ç.xls
PAReport2.exe %today%\r[\1 %today%\r[\1.xls
PAReport3.exe %today%\r[\2 %today%\r[\2.xls
