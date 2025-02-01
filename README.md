Use the below command to convert to exe using pyinstaller.
pyinstaller --onedir --windowed --icon=FTP_Chart.ico --hidden-import=pandas --hidden-import=matplotlib.pyplot --hidden-import=os --hidden-import=tkinter --hidden-import=tkinter.filedialog --hidden-import=tkinter.messagebox --hidden-import=tkinter.toplevel --collect-all openpyxl --hidden-import=sys FTP_Chart.py

This application gives you plots based on engine/ other related data- such as steady, transient, bubble and pie charts
