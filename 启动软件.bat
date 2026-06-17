@echo off
chcp 65001 >nul
title 大圣.快递物流派费结算系统 V1.0
cd /d "%~dp0"
start "" pythonw main.py
