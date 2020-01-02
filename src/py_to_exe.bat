"pyinstaller" -w -F -p "../" --hidden-import nameparser --specpath="./build" --distpath=. UScheduler.py
if not errorlevel 1 rmdir /S /Q "./build"
