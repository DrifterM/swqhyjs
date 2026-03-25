@echo off
echo 正在创建正确的Word文档...
echo 步骤1：检查文件...

if not exist "交易咨询专员设置方案.txt" (
    echo 错误：找不到文本文件
    pause
    exit /b 1
)

echo 步骤2：创建新的.docx文件（使用正确的扩展名和格式）...

REM 方法1：使用copy命令创建干净的文本，然后你可以用Word打开另存为.docx
copy "交易咨询专员设置方案.txt" "交易咨询专员设置方案_正确版.docx"

echo.
echo 完成！
echo.
echo 建议操作：
echo 1. 打开 "交易咨询专员设置方案_正确版.docx"（实际上是文本文件）
echo 2. 在Word中全选内容
echo 3. 另存为真正的Word文档格式（.docx）
echo 4. 应用标题样式、调整格式
echo.
echo 或者直接使用 "交易咨询专员设置方案.txt"，用Word打开后另存为.docx
echo.
pause