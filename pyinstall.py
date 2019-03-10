import  os
if __name__ == '__main__':
    from PyInstaller.__main__ import run
    opts=['CallDemoSerial1.py','-w','-F','--icon=./images/cartoon.ico']
    run(opts)
