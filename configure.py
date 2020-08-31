import os
import sys
import winreg

CREATE_ACTION = "CREATE"
REMOVE_ACTION = "REMOVE"

ACTION: str = CREATE_ACTION

REG_PATH = r"Directory\Background\shell"

CREATE_LIST_PATH = REG_PATH + r"\Оглавление"
PRINT_LIST_PATH = REG_PATH + r"\Напечатать оглавление"

CREATE_LIST_KEY = CREATE_LIST_PATH + r"\command"
PRINT_LIST_KEY = PRINT_LIST_PATH + r"\command"

PROJECT_DIR = os.path.realpath(__file__).replace("configure.py", "")
PYTHON_EXE = PROJECT_DIR + r"venv\Scripts\python.exe"
GENERATOR = PROJECT_DIR + r"__init__.py"

CREATE_REG_VALUE = "\"" + PYTHON_EXE + "\" " + "\"" + GENERATOR + "\" " + "\"%V\""
PRINT_REG_VALUE = CREATE_REG_VALUE + " TRUE " + "TRUE"

if __name__ == "__main__":
    if len(sys.argv) > 1:
        ACTION = str(sys.argv[1].upper())

    if ACTION == CREATE_ACTION:
        try:
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, CREATE_LIST_KEY)
            winreg.CreateKey(winreg.HKEY_CLASSES_ROOT, PRINT_LIST_KEY)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, CREATE_LIST_KEY, 0, winreg.KEY_WRITE) as registry_key:
                winreg.SetValueEx(registry_key, None, 0, winreg.REG_SZ, CREATE_REG_VALUE)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, PRINT_LIST_KEY, 0, winreg.KEY_WRITE) as registry_key:
                winreg.SetValueEx(registry_key, None, 0, winreg.REG_SZ, PRINT_REG_VALUE)
            print("Keys added")
        except WindowsError as err:
            print(err)
    elif ACTION == REMOVE_ACTION:
        try:
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, CREATE_LIST_KEY)
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, CREATE_LIST_PATH)
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, PRINT_LIST_KEY)
            winreg.DeleteKey(winreg.HKEY_CLASSES_ROOT, PRINT_LIST_PATH)
            print("Keys removed")
        except WindowsError as err:
            print(err)
