"pyinstaller" -w -F -i "src\assets\unischeduler_icon.ico" -p "../" --specpath="./build" --distpath=. Scheduler.py
if not errorlevel 1 rmdir /S /Q "./build"
