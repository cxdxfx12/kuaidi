@echo off
chcp 65001 >nul
echo ========================================
echo   Windows 图标缓存清理工具
echo ========================================
echo.

echo [1/3] 结束 Explorer 进程...
taskkill /f /im explorer.exe >nul 2>&1
timeout /t 2 >nul

echo [2/3] 删除图标缓存文件...
del /f /s /q "%userprofile%\AppData\Local\IconCache.db" >nul 2>&1
del /f /s /q "%userprofile%\AppData\Local\Microsoft\Windows\Explorer\iconcache_*.db" >nul 2>&1
del /f /s /q "%userprofile%\AppData\Local\Microsoft\Windows\Explorer\thumbcache_*.db" >nul 2>&1

echo [3/3] 重启 Explorer...
start explorer.exe

echo.
echo ========================================
echo   图标缓存已清理完成！
echo   请在文件管理器中重新查看软件图标
echo ========================================
echo.
pause
