pyinstaller --onefile --noconsole toranim.py
rmdir /s build
del toranim.spec
move dist\toranim.exe .
rmdir dist