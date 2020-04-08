@echo off
title handleExcel
set path=%cd%\jre\bin
%path%\java.exe -jar handleExcel.jar ./src.xls ./result.xls