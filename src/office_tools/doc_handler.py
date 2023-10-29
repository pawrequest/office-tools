import platform
import subprocess
import webbrowser
from abc import ABC, abstractmethod
from pathlib import Path
from typing import Any, Tuple

from comtypes.client import CreateObject  # type: ignore
from docx2pdf import convert as convert_word  # type: ignore


class DocHandler(ABC):
    @abstractmethod
    def display_doc(self, doc_path: Path) -> Tuple[Any, Any]:
        raise NotImplementedError

    @abstractmethod
    def to_pdf(self, doc_file: Path) -> Path:
        raise NotImplementedError


class WordHandler(DocHandler):
    # todo use library
    def display_doc(self, doc_path: Path) -> Tuple[Any, Any]:
        try:
            word = CreateObject('Word.Application')
            word.Visible = True
            word_doc = word.Documents.Open(str(doc_path))
            return word, word_doc
        except OSError as e:
            print(f"Is Word installed? Failed to open {doc_path} with error: {e}")
            raise e
        except Exception as e:
            raise e

    def to_pdf(self, doc_file: Path) -> Path:
        try:
            convert_word(doc_file, output_path=doc_file.parent)
            outfile = doc_file.with_suffix('.pdf')
            print(f"Converted {outfile}")
            return outfile
        except Exception as e:
            raise e


class LibreHandler(DocHandler):
    def display_doc(self, doc_path: Path) -> Tuple[Any, Any]:
        try:
            process = subprocess.Popen(['soffice', str(doc_path)], stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        except Exception as e:
            raise e
        return process, None

    def to_pdf(self, doc_file: Path) -> Path:
        try:
            subprocess.run(f'soffice --headless --convert-to pdf {str(doc_file)} --outdir {str(doc_file.parent)}')
            outfile = doc_file.with_suffix('.pdf')
            print(f"Converted {outfile}")
            return outfile
        except FileNotFoundError:
            print('Is LibreOffice installed?')
            raise FileNotFoundError('LibreOffice not installed')


def get_libre_platform():
    system = platform.system()
    architecture = platform.architecture()[0]

    if system == 'Windows':
        if '64' in architecture:
            return 'win', 'x86_64', 'Win_x86-64.msi'
        else:
            return 'win', 'x86', 'Win_x86.msi'
    elif system == 'Darwin':  # macOS
        return 'mac', 'aarch64', 'MacOS_aarch64.dmg'
    elif system == 'Linux':
        # todo check debian vs other
        if '64' in architecture:
            return 'deb', 'x86_64', 'Linux_x86-64_deb.tar.gz'
        else:
            return 'deb', 'x86', 'Linux_x86_deb.tar.gz'
    else:
        raise EnvironmentError(f"Unsupported system: {system}")


def go_install_libre(version='7.6.2'):
    plat = get_libre_platform()
    dl = f'https://www.libreoffice.org/donate/dl/{plat[0]}-{plat[1]}/{version}/en-US/LibreOffice_{version}_{plat[2]}'
    webbrowser.open(dl)
    return True


def direct_libre_dl():
    libre_platform = get_libre_platform()
    libre_version = '7.6.2'
    dl = f'https://download.documentfoundation.org/libreoffice/stable/{libre_version}/{libre_platform[0]}/{libre_platform[1]}/LibreOffice_{libre_version}_{libre_platform[2]}'
    return dl
