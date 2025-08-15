import warnings
import os.path
import time
import sys
from PyQt6 import QtWidgets
from PyQt6.QtWidgets import QWidget, QPushButton
from create_epub_from_xlsx import translated_xlsx_to_epub, clear_files, epub_to_xlsx

warnings.simplefilter(action='ignore', category=UserWarning)
warnings.simplefilter(action='ignore', category=FutureWarning)


def show_dialog():
    index = 0
    for file in os.listdir('./to_translate'):
        print()
        epub_to_xlsx('./to_translate/{0}'.format(file), index)
        index += 1


def show_dialog2():
    print(len(os.listdir('./to_epub')))
    print(os.listdir('./to_epub'))
    index = 1
    ids_list = []
    for file in os.listdir('./to_epub'):
        if file.split('-')[0] not in ids_list:
            print(f'{index}: start work with file {file}')
            ids_list.append(file.split('-')[0])
            index += 1
            if (file.split('_')[0].split('-')[1] == '0'):
                translated_xlsx_to_epub('./to_epub/{0}'.format(file), '.'.join(file.split('.')[:-1]), index)
            else:
                print(f'skip file {file}')
            time.sleep(0.1)


class epubTranslatorForm(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.driver = None
        self._init_btn_epub_to_xlsx(parent)
        self._init_btn_xlsx_to_epub(parent)
        self._init_btn_clear_folders(parent)

    def _init_btn_epub_to_xlsx(self, parent):
        # Инициализирует кнопку перевода epub файла в xlsx из папки to_translate в папку temp
        self.epub_choose_button = QPushButton(parent)
        self.epub_choose_button.setGeometry(50, 25, 150, 50)
        self.epub_choose_button.setText("Epub -> Temp Xlsx")
        self.epub_choose_button.clicked.connect(show_dialog)

    def _init_btn_xlsx_to_epub(self, parent):
        # Инициализирует кнопку перевода xlsx файла в epub из папки to_epub в папку result
        self.translated_choose_button = QPushButton(parent)
        self.translated_choose_button.setGeometry(50, 125, 150, 50)
        self.translated_choose_button.setText("Rus Xlsx -> Rus Epub")
        self.translated_choose_button.clicked.connect(show_dialog2)

    def _init_btn_clear_folders(self, parent):
        # Инициализирует кнопку очищения папок result, temp, to_epub, to_Translate
        self.clear_folder_button = QPushButton(parent)
        self.clear_folder_button.setGeometry(250, 25, 100, 50)
        self.clear_folder_button.setText("Очистить папку")
        self.clear_folder_button.clicked.connect(clear_files)


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    try:
        widget = QtWidgets.QWidget()
        widget.resize(400, 200)
        win = epubTranslatorForm(widget)
        widget.setWindowTitle("Перевод epub с иностранного на русский")
        widget.show()
        exit(app.exec())
    except Exception as e:
        raise e
