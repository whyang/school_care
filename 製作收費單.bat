REM """
REM Created on Saturday Jan. 22, 2022
REM @author: Wen-Hsin Yang
REM
REM 說明：
REM 執行批次檔(Execution Script of Batch file on Windows)
REM 使用 pandas, openpyxl, pywin32整理安親班學生收費資料，製作學生收費工作表及收費單(Excel worksheet, PDF file) 
REM 1. 讀入[安親斑_收費資料.xlsx]: 名冊、收費項目、收費期間 等內容
REM 2. 整理成[收費單.xlsx]：每位學生使用1個工作表
REM 3. 由[收費單.xlsx] 製作 [收費單_合併.xlsx]、[收費單_合併.pdf]：可列印用收費單(PDF)
REM 4. 用pyInstaller打包"製作收費單.py"成"製作收費單.exe"
REM 5. 編寫command scripts將"製作收費單.exe"，製作成執行批次檔("製作收費單.bat")
REM """

@echo off

REM 後續命令使用的是：UTF-8編碼
chcp 65001

REM 先清空螢幕
cls

REM 設定區域變數
setlocal

:BeginProcess
echo ...
echo 請輸入收費資料檔名，例如'ABC收費資料.xlsx' 或是 直接按 Enter 鍵(使用'收費資料.xlsx')
SET /P filename=
echo 你輸入的是：%filename%
pause

REM 檢查檔案是否存在
if exist %filename% goto runProcess
if "%filename%" == "" goto runDefault

:errorProcess
echo %filename% 不存在，請將檔案拷貝放在目前執行目錄下
set filename=
goto BeginProcess

:runDefault
REM 製作收費單(使用default指定的收費資料檔案 '收費資料.xlsx')
echo 使用default指定的收費資料檔案 '收費資料.xlsx'
set filename=收費資料.xlsx
if not exist %filename% goto errorProcess
pause
start 製作收費單.exe && goto EOF

:runProcess
REM 製作收費單(依據指定的收費資料檔案)
start 製作收費單.exe -f %filename%

:EOF
REM 有錯誤發生
if errorlevel 1 echo 有錯誤發生! && pause

REM 結束設定區域變數
endlocal