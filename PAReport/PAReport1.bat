REM ���ϐ��̐ݒ�
set ssUser=fsol07
set ssPwd=fsol07
set ssDir=\\ADSDEAI1\vss
set targetDir=D:\data\VssWork\04.�R���r�s�d�o\04.�o�r�`�o�s�H��\81.�i���Ǘ�
set targetPrj=$/04.�R���r�s�d�o/04.�o�r�`�o�s�H��/81.�i���Ǘ�/
set reviewDir1="X:\52_�������v��\04.�R���r�s�d�o\04.�o�r�`�o�s�H��\85.�Z���t���r���[\�o���E�x���E���ʃ��x�[�g"
set reviewDir2="X:\52_�������v��\04.�R���r�s�d�o\04.�o�r�`�o�s�H��\84.���r���["
set BaseDir=%~dp0

REM �^�C���X�^���v�̓��t�����߂�
for /F "tokens=1" %%a in ('date /t') do set orgdate=%%a
set yy=%orgdate:~0,4%
set mm=%orgdate:~5,2%
set dd=%orgdate:~8,2%
set today="%BaseDir%%yy%%mm%%dd%"

REM �v���W�F�N�g�f�B���N�g�����č쐬
rmdir %targetDir% /S /Q
mkdir %targetDir%

REM ���t�f�B���N�g�����č쐬
rmdir %today% /S /Q
mkdir %today%

REM VSS���ŐV�ł��擾
ss Get %targetPrj% -R  -I-Y -GL%targetDir% -GWR -GTU -GCK

REM �i���Ǘ��t�@�C�����R�s�[
mkdir %today%\�i���Ǘ�
robocopy %targetDir% %today%\�i���Ǘ� *.xls /E

REM ����������
PAReport1.exe %today%\�i���Ǘ� %today%\�i���Ǘ�.xls
