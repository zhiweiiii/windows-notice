@echo off
chcp 65001
SET TaskName=提醒
SET VbsPath="%~dp0notice.vbs"
:: 检测是否以管理员运行
openfiles >nul 2>&1
if %errorlevel% NEQ 0 (
    echo 提升权限中...
    powershell -Command "Start-Process '%~f0' -Verb runAs"
    exit /b
)
REM 删除已有同名任务（防止重复）
schtasks /Delete /TN "%TaskName%" /F

REM 创建每天中午12:00触发的任务
schtasks /Create /SC DAILY /TN "%TaskName%" /TR "%VbsPath%" /ST 10:36 /RL HIGHEST /F
echo 任务计划创建完成！
pause
