# NRB EXCEL REFRESH LINKS
## To Edit Source Code and Work with GIT
1. Use Git Bash
2. `cd /c/Program\ Files/NRB\ Excel\ Refresh\ Links/`
3. Remember to Create New Branch Before Doing Any Work

## Generate UI
1. Ues QT Creator
2. ReplacePartsXLS.ui
3. DialogResults.ui
4. `/c/Python27/Lib/site-packages/PyQt4/pyuic4.bat -x MainWindow.ui -o MainWindow.py`
5. `/c/Python27/Lib/site-packages/PyQt4/pyuic4.bat -x MainWindow.ui -o MainWindow.py`

## Build Executable
`pyinstaller.exe --onefile --windowed --icon forklift.ico  Excel/ Refresh/Links\ FWW.spec`
