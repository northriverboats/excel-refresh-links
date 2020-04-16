# NRB EXCEL REFRESH LINKS
## To initalize project with Python 3.7
1. Use Git Bash
2. `cd /c/Development/`
3. `git clone git@github.com:northriverboats/excel-refresh-links.git`
4. `cd /c/Development/excel-refresh-links/`
5. Switch to Windows Termial
6. `cd \Development\excel-refresh-links`
7. Download [PyQt4](https://download.lfd.uci.edu/pythonlibs/s2jqpv5t/PyQt4-4.11.4-cp37-cp37m-win_amd64.whl)
8. `\python37\python -m venv venv`
9. `venv\Scripts\activate.bat`
10. `python -m pip install pip --upgrade`
11. `pip install  "\Users\<username>\Downloads\PyQt4-4.11.4-cp37-cp37m-win_amd64"`
12. `pip install -r requirements.txt`

## To Edit Source Code and Work with GIT
1. Use Windows Terminal
2. `cd Development/NRB\ Excel\ Refresh\ Links/`
3. `venv\Scripts\activate.bat`
4. Remember to Create New Branch Before Doing Any Work

## Generate UI
1. Ues QT Creator
2. `venv\Lib\site-packages\PyQT4\designer.exe MainWindow.ui`
3. To create Python Code
4. `venv\Lib\site-packages\PyQt4\pyuic4.bat -x MainWindow.ui -o MainWindow.py`

## Build Executable
`pyinstaller.exe --onefile --windowed --icon relink.ico  "Excel Refresh Links FWW.spec"`