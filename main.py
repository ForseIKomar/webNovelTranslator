import warnings
warnings.simplefilter(action='ignore', category=FutureWarning)

import io
import os.path
import sys
import time
import urllib
from collections import OrderedDict
from PIL import Image
from io import BytesIO
import requests as req
from PyQt6 import QtWidgets, QtCore
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QIntValidator, QFont, QDoubleValidator, QPixmap, QImage
from PyQt6.QtWidgets import QLineEdit, QWidget, QFormLayout, QPushButton, QFileDialog, QDialog, QMessageBox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
import ebooklib
from ebooklib import epub
import xlsxwriter
import datetime
import pandas as pd
import pyautogui, sys
import os
import keyboard
from selenium.webdriver.support.expected_conditions import text_to_be_present_in_element_attribute
from selenium.webdriver.support.wait import WebDriverWait


class lineEditDemo(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.driver = None
        self.code_name = "None"
        self.parent = parent

        self.epub_text_edit = QLineEdit(parent)
        self.epub_text_edit.setReadOnly(True)
        self.epub_text_edit.setFont(QFont("Arial", 10))
        self.epub_text_edit.setGeometry(50, 50, 400, 50)

        self.epub_choose_button = QPushButton(parent)
        self.epub_choose_button.setGeometry(450, 50, 150, 50)
        self.epub_choose_button.setText("Выбрать epub\n для трансформации")
        self.epub_choose_button.clicked.connect(self.show_dialog)

        self.epub_to_xlsx_button = QPushButton(parent)
        self.epub_to_xlsx_button.setGeometry(650, 50, 100, 50)
        self.epub_to_xlsx_button.setText("Epub TO Xlsx")
        self.epub_to_xlsx_button.clicked.connect(self.epub_to_xlsx)

        self.image_link_edit = QLineEdit(parent)
        self.image_link_edit.setReadOnly(True)
        self.image_link_edit.setFont(QFont("Arial", 10))
        self.image_link_edit.setGeometry(50, 150, 400, 50)

        self.image_clear_button = QPushButton(parent)
        self.image_clear_button.setGeometry(450, 150, 150, 50)
        self.image_clear_button.setText("Очистить")
        self.image_clear_button.clicked.connect(self.clear_image_link)

        self.image_clear_button = QPushButton(parent)
        self.image_clear_button.setGeometry(650, 150, 100, 50)
        self.image_clear_button.setText("Браузер")
        self.image_clear_button.clicked.connect(self.open_chrome_to_translate)

        self.translated_path_edit = QLineEdit(parent)
        self.translated_path_edit.setReadOnly(True)
        self.translated_path_edit.setFont(QFont("Arial", 10))
        self.translated_path_edit.setGeometry(50, 250, 400, 50)

        self.translated_choose_button = QPushButton(parent)
        self.translated_choose_button.setGeometry(450, 250, 150, 50)
        self.translated_choose_button.setText("Выбрать перевод для \nтрансформации в epub")
        self.translated_choose_button.clicked.connect(self.show_dialog2)

        self.translated_to_epub_button = QPushButton(parent)
        self.translated_to_epub_button.setGeometry(650, 250, 100, 50)
        self.translated_to_epub_button.setText("Translated\n TO EPUB")
        self.translated_to_epub_button.clicked.connect(self.translated_xlsx_to_epub)

    def clear_image_link(self):
        self.image_link_edit.setText('')

    def show_dialog(self):
        self.update_text()
        #fname = QFileDialog.getOpenFileName(self, 'Open file', '.')
        #self.epub_text_edit.setText(fname[0])
        for file in os.listdir('./to_translate'):
            self.epub_to_xlsx('./to_translate/{0}'.format(file))
        #self.open_chrome_to_translate()

    def show_dialog2(self):
        self.update_text()
        for file in os.listdir('./to_epub'):
            print(file)
            self.translated_xlsx_to_epub('./to_epub/{0}'.format(file), '.'.join(file.split('.')[:-1]))
            time.sleep(0.1)
        self.epub_to_xlsx_button.setText('Готово!')

    def epub_to_xlsx(self, file):
        self.update_text()
        book = epub.read_epub(file)
        if book.metadata[list(book.metadata.keys())[0]]['source'][0][1]['id'] == 'id.cover-image':
            self.image_link_edit.setText(book.metadata[list(book.metadata.keys())[0]]['source'][0][0])
        items = list(book.get_items_of_type(ebooklib.ITEM_DOCUMENT))
        self.code_name = '.'.join(file.split('/')[-1].split('.')[:-1])
        wb = xlsxwriter.Workbook('temp/{0}.xlsx'.format(self.code_name))
        sheet = wb.add_worksheet()
        new_it_index = 0
        for item_index, elem in enumerate(items[1:]):
            out = 0
            text = items[item_index].get_body_content().decode("utf-8")
            index = 0
            curr_text = ''
            result_text = ''
            prev = ''
            while index < len(text):
                if text[index] != '<' and text[index] != '>':
                    curr_text += str(text[index])
                elif text[index] == '<':
                    if text[index + 1] != '/':
                        prev = '<'
                        result_text += curr_text.strip() + ' | ' if curr_text.strip() != '' else ''
                        curr_text = ''
                    else:
                        prev = '</'
                        result_text += curr_text.strip() + ' | ' if curr_text.strip() != '' else ''
                        curr_text = ''
                        index += 1
                elif text[index] == '>':
                    if prev == '<':
                        # print(out*' ' + curr_text)
                        curr_text = ''
                        out += 1
                        prev = '>'
                    else:
                        # print(out*' ' + curr_text)
                        curr_text = ''
                        out -= 1
                        prev = '>'
                index += 1

            if len(result_text) > 30000:
                absatz = result_text.split('|')
                sp_text_1 = absatz[0]
                is_first = True
                for index, elem in enumerate(absatz[1:]):
                    if len(sp_text_1) + len(elem) > 30000:
                        if is_first:
                            sheet.write(new_it_index + 1, 0, '' + sp_text_1)
                            is_first = False
                        else:
                            sheet.write(new_it_index + 1, 0, '$$$ ' + sp_text_1)

                        new_it_index += 1
                        sp_text_1 = elem
                    else:
                        sp_text_1 = sp_text_1 + '|' + elem
                if len(sp_text_1) > 0:
                    sheet.write(new_it_index + 1, 0, '$$$ ' + sp_text_1)
                    new_it_index += 1
            else:
                sheet.write(new_it_index + 1, 0, result_text)
                new_it_index += 1
        wb.close()

    def open_chrome_to_translate(self):
        service = Service(executable_path="C:\\chromedriver\\chromedriver.exe")
        self.driver = webdriver.Chrome(service=service)
        self.driver.implicitly_wait(5)
        self.driver.get('https://translate.google.com/?sl=en&tl=ru&op=docs')
        time.sleep(0.2)
        self.driver.find_element(By.XPATH, '//*[@id="yDmH0d"]/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[3]/label').click()
        time.sleep(0.2)
        pyautogui.click(x=378, y=61)
        time.sleep(0.1)
        pyautogui.write('C:\\Users\\samit\\PycharmProjects\\webNovelTranslator\\temp\\')
        time.sleep(0.1)
        pyautogui.press('Enter')
        time.sleep(0.1)
        pyautogui.click(x=331, y=488)
        time.sleep(0.1)
        pyautogui.write(self.code_name + '.xlsx')
        time.sleep(0.1)
        pyautogui.press('Enter')
        time.sleep(0.5)
        while True:
            time.sleep(1)
        self.driver.find_element(By.XPATH, '/html/body/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[2]/div/div/button/span').click()
        time.sleep(0.1)
        try:
            WebDriverWait(self.driver, 60).until(
            text_to_be_present_in_element_attribute((By.XPATH, "/html/body/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[2]/div/div/button/span"), "Перевести")
            )
        except:
            pass
        time.sleep(0.1)
        self.driver.find_element(By.XPATH, '/html/body/c-wiz/div/div[2]/c-wiz/div[3]/c-wiz/div[2]/c-wiz/div/div[1]/div/div[2]/div/button/span[2]').click()
        time.sleep(1)
        self.driver.close()

    def translated_xlsx_to_epub(self, file, filename):
        xlsx_dataframe = pd.read_excel(io=file) #self.translated_path_edit.text())
        res_arr = xlsx_dataframe.to_numpy()
        book = epub.EpubBook()
        if True:#self.image_link_edit.text() != '':
            last_book = epub.read_epub('./to_translate/{0}.epub'.format(filename))
            if last_book.metadata[list(last_book.metadata.keys())[0]]['source'][0][1]['id'] == 'id.cover-image':
                image_link = last_book.metadata[list(last_book.metadata.keys())[0]]['source'][0][0]
                image_content = Image.open(BytesIO(req.get(image_link).content))
                b = io.BytesIO()
                image_content.save(b, 'jpeg')
                b_image1 = b.getvalue()
                image1_item = epub.EpubItem(uid='image_1', file_name='images/image1.jpeg', media_type='image/jpeg',
                                            content=b_image1)

                # add Image file
                book.set_cover(file_name='images/image1.jpeg', content=b_image1)
                book.add_item(image1_item)
        book.set_identifier(res_arr[0][0].split('|')[2].split('/')[-1])
        book.set_title(res_arr[0][0].split('|')[3])
        book.set_language('ru')
        book.add_author(res_arr[0][0].split('|')[9])
        style = '''BODY { text-align: justify;}'''

        default_css = epub.EpubItem(uid="style_default", file_name="style/default.css", media_type="text/css",
                                    content=style)
        book.add_item(default_css)
        chapters = []
        for index, element in enumerate(res_arr):
            if element[0][:3] != '$$$':
                c1 = epub.EpubHtml(title=element[0].split('|')[0],
                                   file_name='page{0}.xhtml'.format(index),
                                   lang='ru')
                txt = element[0]
                part_index = 1
                while index + part_index < len(res_arr) and res_arr[index + part_index][0][:3] == '$$$':
                    txt += '|' + res_arr[index + part_index][0][3:]
                    part_index += 1
                text = ' </p><p> '.join(txt.split('|')[1:])
                c1.set_content(u'<html><body><h3>{0}</h3><p>{1}</p></body></html>'.format(txt.split('|')[0], text))
                book.add_item(c1)
                chapters.append(c1)

        chapters.insert(0, epub.Link('intro.xhtml', 'Introduction', 'intro'))

        book.toc = chapters

        # add navigation files
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())
        chapters[0] = 'nav'
        book.spine = chapters
        print('./result/{0}.epub'.format(res_arr[0][0].split('|')[3].strip().replace(':', '').replace('/', ' ').replace('?', '')))
        epub.write_epub('./result/{0}.epub'.format(res_arr[0][0].split('|')[3].strip().replace(':', '').replace('/', ' ').replace('?', '')), book, {})

    def update_text(self):
        self.epub_to_xlsx_button.setText('Epub TO Xlsx')



if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    widget = QtWidgets.QWidget()
    widget.resize(800, 400)
    win = lineEditDemo(widget)
    widget.setWindowTitle("This is PyQt Widget example")
    widget.show()
    exit(app.exec())
