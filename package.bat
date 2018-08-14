rd /s /q build\
rd /s /q dist\
pyinstaller --hidden-import=queue -F -w -i "img/calculator.ico" "user_interface.py"
md dist\platforms
xcopy platforms dist\platforms /s
copy user_interface.ui dist\
cd /d dist
ren user_interface.exe calculator.exe