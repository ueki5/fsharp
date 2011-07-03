REM 環境変数の設定
set ssUser=fsol07
set ssPwd=fsol07
set ssDir=\\ADSDEAI1\vss
set targetDir=D:\data\VssWork\04.３次ＳＴＥＰ\04.ＰＳ～ＰＴ工程\81.進捗管理
set targetPrj=$/04.３次ＳＴＥＰ/04.ＰＳ～ＰＴ工程/81.進捗管理/
set reviewDir1="X:\52_中長期計画\04.３次ＳＴＥＰ\04.ＰＳ～ＰＴ工程\85.セルフレビュー\覚書・支払・流通リベート"
set reviewDir2="X:\52_中長期計画\04.３次ＳＴＥＰ\04.ＰＳ～ＰＴ工程\84.レビュー"
set BaseDir=%~dp0

REM タイムスタンプの日付を求める
for /F "tokens=1" %%a in ('date /t') do set orgdate=%%a
set yy=%orgdate:~0,4%
set mm=%orgdate:~5,2%
set dd=%orgdate:~8,2%
set today="%BaseDir%%yy%%mm%%dd%"

REM プロジェクトディレクトリを再作成
rmdir %targetDir% /S /Q
mkdir %targetDir%

REM 日付ディレクトリを再作成
rmdir %today% /S /Q
mkdir %today%

REM VSSより最新版を取得
ss Get %targetPrj% -R  -I-Y -GL%targetDir% -GWR -GTU -GCK

REM 進捗管理ファイルをコピー
mkdir %today%\進捗管理
robocopy %targetDir% %today%\進捗管理 *.xls /E

REM レビュー表１をコピー
mkdir %today%\レビュー表1
robocopy %reviewDir1% %today%\レビュー表1 内部レビュー記録票*.xls /E

REM レビュー表２をコピー
mkdir %today%\レビュー表2
robocopy %reviewDir2% %today%\レビュー表2 PTレビュー管理票*_覚書支払.xls /E

REM 資料を結合
PAReport1.exe %today%\進捗管理 %today%\進捗管理.xls
PAReport2.exe %today%\レビュー表1 %today%\レビュー表1.xls
PAReport3.exe %today%\レビュー表2 %today%\レビュー表2.xls
