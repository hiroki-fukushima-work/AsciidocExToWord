@echo off
chcp 65001 >nul
setlocal

echo ASCII DOC CONVERTER

cd /d %~d0%~p0
echo CurrentDir: %~d0%~p0

set ADOC_NAME=AsciiDocSample
set ADOCBASEPATH=%~d0%~p0
set SCRIPTBASEPATH=%~d0%~p0

del /f /q /s "%ADOCBASEPATH%out" 2>nul

echo START %ADOC_NAME% Word Convert

powershell -NoProfile -ExecutionPolicy Bypass ^
  -File "%SCRIPTBASEPATH%\Convert-AsciiDocToWord.ps1" ^
  -AdocFullPath "%ADOCBASEPATH%%ADOC_NAME%.adoc" ^
  -OutputFullPath "%ADOCBASEPATH%out\%ADOC_NAME%.docx" ^
  -ConfigFullPath "%ADOCBASEPATH%conf/word-style.sample.json"

set ADOC_NAME=AsciiDocToWord
powershell -NoProfile -ExecutionPolicy Bypass ^
  -File "%SCRIPTBASEPATH%\Convert-AsciiDocToWord.ps1" ^
  -AdocFullPath "%ADOCBASEPATH%%ADOC_NAME%.adoc" ^
  -OutputFullPath "%ADOCBASEPATH%out\%ADOC_NAME%.docx" ^
  -ConfigFullPath "%ADOCBASEPATH%conf/word-style.sample.json"

echo FINISH %ADOC_NAME% Word Convert
endlocal
pause
