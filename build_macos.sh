#!/bin/bash
# ============================================
#   大圣.快递物流派费结算系统V1.0 - macOS打包脚本
#   杭州喵喵至家网络有限公司 · 大圣智慧软件
# ============================================

echo ""
echo "========================================"
echo "   大圣.快递物流派费结算系统V1.0"
echo "   macOS打包脚本"
echo "========================================"
echo ""

# 检查PyInstaller
python3 -c "import PyInstaller" 2>/dev/null
if [ $? -ne 0 ]; then
    echo "[信息] 安装 PyInstaller..."
    pip3 install PyInstaller==5.13.0
fi

# 使用PyInstaller打包
python3 -m PyInstaller \
    --name="大圣派费结算系统" \
    --windowed \
    --icon="data/icons/monkey-icon.png" \
    --add-data="data/icons/*:data/icons/" \
    --add-data="data/config/*:data/config/" \
    --add-data="data/uploads/*:data/uploads/" \
    --hidden-import=openpyxl \
    --hidden-import=pandas \
    --hidden-import=sqlalchemy \
    --hidden-import=PyQt5.sip \
    --hidden-import=PyQt5.QtCore \
    --hidden-import=PyQt5.QtGui \
    --hidden-import=PyQt5.QtWidgets \
    -y \
    --clean \
    main.py

echo ""
echo "========================================"
echo "   打包完成！"
echo "   产物目录: dist/大圣派费结算系统.app"
echo "========================================"
echo ""