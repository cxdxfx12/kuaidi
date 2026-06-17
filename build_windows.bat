@echo off
cd /d "%~dp0"

echo ========================================
echo   大圣.快递物流派费结算系统V1.0
echo   Windows打包脚本
echo ========================================
echo.

python -m pip install PyInstaller==5.13.0

python -m PyInstaller ^
    --name="大圣派费结算系统" ^
    --windowed ^
    "--add-data=data/icons/*;data/icons/" ^
    "--add-data=data/config/*;data/config/" ^
    "--add-data=data/uploads/*;data/uploads/" ^
    --hidden-import=openpyxl ^
    --hidden-import=pandas ^
    --hidden-import=sqlalchemy ^
    --hidden-import=PyQt5.sip ^
    --hidden-import=PyQt5.QtCore ^
    --hidden-import=PyQt5.QtGui ^
    --hidden-import=PyQt5.QtWidgets ^
    -y ^
    --clean ^
    main.py

echo.
echo ========================================
echo   打包完成！
echo   产物目录: dist\大圣派费结算系统
echo ========================================
pause