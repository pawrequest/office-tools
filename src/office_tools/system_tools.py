import asyncio
import os
import shutil
import winreg as reg
from pathlib import Path


# acomment
def print_file(file_path: Path):
    try:
        os.startfile(str(file_path), 'print')
        return True
    except Exception as e:
        print(f'Failed to print: {e}')
        return False


async def wait_for_process(process):
    while True:
        res = process.poll()
        if res is not None:
            break
        await asyncio.sleep(3)
    print('Process has finished.')


def check_registry(reg_path: str) -> bool:
    try:
        key = reg.OpenKey(reg.HKEY_LOCAL_MACHINE, reg_path)
        reg.CloseKey(key)
        return True
    except FileNotFoundError:
        return False


def check_word() -> bool:
    return check_registry(
        r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\WINWORD.EXE'
    )


def check_excel() -> bool:
    return check_registry(
        r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\EXCEL.EXE'
    )


def check_lib2() -> bool:
    return check_registry(r'SOFTWARE\LibreOffice')


def check_libre() -> bool:
    return shutil.which(r'soffice.exe') is not None


def check_outlook() -> bool:
    return check_registry(
        r'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
    )
