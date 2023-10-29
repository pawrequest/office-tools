from pathlib import Path
from typing import Tuple
import PySimpleGUI as sg
from docxtpl import DocxTemplate


def get_template_and_path(tmplt, temp_file, context=None) -> Tuple[DocxTemplate, Path]:
    context = context or dict()
    template = DocxTemplate(tmplt)
    template.render(context)

    while True:
        try:
            template.save(temp_file)
            return template, temp_file
        except Exception as e:
            if sg.popup_ok_cancel("Close the template file and try again") == "OK":
                continue
            else:
                raise e


def fake():
    if 1 > 0:
        print("test")
