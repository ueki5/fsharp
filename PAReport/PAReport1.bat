REM ŠÂ‹«•Ï”‚Ìİ’è
set ssUser=fsol07
set ssPwd=fsol07
set ssDir=\\ADSDEAI1\vss
set targetDir=D:\data\VssWork\04.‚RŸ‚r‚s‚d‚o\04.‚o‚r`‚o‚sH’ö\81.i’»ŠÇ—
set targetPrj=$/04.‚RŸ‚r‚s‚d‚o/04.‚o‚r`‚o‚sH’ö/81.i’»ŠÇ—/
set reviewDir1="X:\52_’†’·ŠúŒv‰æ\04.‚RŸ‚r‚s‚d‚o\04.‚o‚r`‚o‚sH’ö\85.ƒZƒ‹ƒtƒŒƒrƒ…[\Šo‘Ex•¥E—¬’ÊƒŠƒx[ƒg"
set reviewDir2="X:\52_’†’·ŠúŒv‰æ\04.‚RŸ‚r‚s‚d‚o\04.‚o‚r`‚o‚sH’ö\84.ƒŒƒrƒ…["
set BaseDir=%~dp0

REM ƒ^ƒCƒ€ƒXƒ^ƒ“ƒv‚Ì“ú•t‚ğ‹‚ß‚é
for /F "tokens=1" %%a in ('date /t') do set orgdate=%%a
set yy=%orgdate:~0,4%
set mm=%orgdate:~5,2%
set dd=%orgdate:~8,2%
set today="%BaseDir%%yy%%mm%%dd%"

REM ƒvƒƒWƒFƒNƒgƒfƒBƒŒƒNƒgƒŠ‚ğÄì¬
rmdir %targetDir% /S /Q
mkdir %targetDir%

REM “ú•tƒfƒBƒŒƒNƒgƒŠ‚ğÄì¬
rmdir %today% /S /Q
mkdir %today%

REM VSS‚æ‚èÅV”Å‚ğæ“¾
ss Get %targetPrj% -R  -I-Y -GL%targetDir% -GWR -GTU -GCK

REM i’»ŠÇ—ƒtƒ@ƒCƒ‹‚ğƒRƒs[
mkdir %today%\i’»ŠÇ—
robocopy %targetDir% %today%\i’»ŠÇ— *.xls /E

REM ‘—¿‚ğŒ‹‡
PAReport1.exe %today%\i’»ŠÇ— %today%\i’»ŠÇ—.xls
