@echo off
pyinstaller --onefile --windowed --name="高校学生档案转递单批量生成工具" --add-data="template;template" --clean Generator-for-college-student-file-transfer-form.py
pause