@echo off
chcp 65001
SET TaskName=提醒
:: 检测是否以管理员运行
openfiles >nul 2>&1
if %errorlevel% NEQ 0 (
    echo 提升权限中...
    powershell -Command "Start-Process '%~f0' -Verb runAs"
    exit /b
)

REM 检查是否存在任务
schtasks /Query /TN "%TaskName%" >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo 任务 "%TaskName%" 不存在，无需删除。
    pause
    exit /b
)

REM 删除任务
schtasks /Delete /TN "%TaskName%" /F

IF %ERRORLEVEL% EQU 0 (
    echo 任务 "%TaskName%" 已成功删除！
) ELSE (
    echo 删除任务 "%TaskName%" 失败，请检查权限。
)

pause
