import io
import requests as req
import pandas as pd
from PIL import Image
from io import BytesIO
from ebooklib import epub
import os
import ebooklib
import xlsxwriter
from xlsxwriter.worksheet import Worksheet
from novel_class import BookParser
from utils.utils import *
import numpy as np


def clear_files():
    folder_names = ['./temp', './to_epub', './to_translate', './result']
    for i in range(0, 2):
        for file in os.listdir(folder_names[i]):
            os.remove(f'{folder_names[i]}/{file}')
    for file in os.listdir(folder_names[2]):
        os.replace(f'{folder_names[2]}/{file}', f'originals/{file}')
    for file in os.listdir(folder_names[3]):
        os.replace(f'{folder_names[3]}/{file}', f'all_result_books/{file}')


def translated_xlsx_to_epub(file, filename, enumerator):
    file_num = filename.split('-')[0]                       # Получаем номер книги
    res_arr = []
    print(os.listdir('./to_epub'))
    for file in os.listdir('./to_epub'):
        if file.split('-')[0] == file_num:
            print(f'{file} is a part of {filename}')
            xlsx_dataframe = pd.read_excel(io='./to_epub/' + file)
            new_arr = xlsx_dataframe.to_numpy()
            for elem in new_arr:
                res_arr.append(elem)

    book = epub.EpubBook()
    try:
        print(f"trying to get image from {'_'.join(filename.split('_')[1:])}")
        last_book = epub.read_epub('./to_translate/{0}.epub'.format('_'.join(filename.split('_')[1:])))
        if last_book.metadata[list(last_book.metadata.keys())[0]]['source'][0][1]['id'] == 'id.cover-image':
            image_link = last_book.metadata[list(last_book.metadata.keys())[0]]['source'][0][0]
            image_content = Image.open(BytesIO(req.get(image_link).content))
            b = io.BytesIO()
            image_content.save(b, 'jpeg')
            b_image1 = b.getvalue()
            image1_item = epub.EpubItem(uid='image_1', file_name='images/image1.jpeg', media_type='image/jpeg',
                                        content=b_image1)
            book.set_cover(file_name='images/image1.jpeg', content=b_image1)
            book.add_item(image1_item)
    except:
        pass

    book_name = set_book_meta_data(book, filename, res_arr[0][0])
    book.set_language('ru')
    style = '''BODY { text-align: justify;}'''

    default_css = epub.EpubItem(uid="style_default", file_name="style/default.css", media_type="text/css",
                                content=style)
    book.add_item(default_css)
    chapters = []

    for index, element in enumerate(res_arr):
        if isinstance(element[0], str) and element[0][:3] != '$$$':
            c1 = epub.EpubHtml(title=element[0].split('|')[0],
                               file_name='page{0}.xhtml'.format(index),
                               lang='ru')
            txt = element[0]
            part_index = 1
            try:
                while index + part_index < len(res_arr) and res_arr[index + part_index][0][:3] == '$$$':
                    txt += '|' + res_arr[index + part_index][0][3:]
                    part_index += 1
            except Exception as e:
                print(e)
                print(res_arr[index + part_index])
            text = ' </p><p> '.join(txt.split('|')[1:])
            c1.set_content(u'<html><body><h3>{0}</h3><p>{1}</p></body></html>'.format(txt.split('|')[0], text))
            book.add_item(c1)
            chapters.append(c1)

    chapters.insert(0, epub.Link('intro.xhtml', 'Introduction', 'intro'))

    book.toc = chapters

    book.add_item(epub.EpubNcx())
    book.add_item(epub.EpubNav())
    chapters[0] = 'nav'
    book.spine = chapters

    new_file_name = './result/{0}-{1}.epub'.format(book_name.strip().replace(':', '').replace('/', ' ').replace('?', '')
                                                   , enumerator)
    print(f',   result stored in {new_file_name}')
    print()
    epub.write_epub(new_file_name, book, {})


def set_book_meta_data(book, filename, meta_string):
    if meta_string.find('panda-novel') != -1:
        book.set_identifier(meta_string.split('|')[2].split('/')[-1])
        book.set_title(meta_string.split('|')[3])
        book_name = meta_string.split('|')[2].split('/')[-1]
        book.add_author(meta_string.split('|')[9])
    elif meta_string.find('fanfiction.net') != -1:
        book.set_identifier(meta_string.split('|')[3].split(': ')[-1])
        element = ''
        for elem in meta_string.split('|'):
            if elem.find('fanfiction.net') != -1:
                element = elem
        book.set_title(element.split('/')[-1])
        book_name = element.split('/')[-1]
        book.add_author(meta_string.split('|')[6])
    elif meta_string.find('hostednovel.com') != -1:
        book.set_identifier(meta_string.split('|')[2])
        book.set_title(meta_string.split('|')[2].split('/')[-1].split('#')[0])
        book_name = meta_string.split('|')[2].split('/')[-1].split('#')[0]
    elif meta_string.find('Custom-1') != -1:
        book.set_identifier(meta_string.split('|')[0])
        book.set_title(meta_string.split('|')[0])
        book_name = meta_string.split('|')[0]
    elif meta_string.find(' https://genesistls.com/') != -1:
        book.set_identifier(meta_string.split('|')[2])
        book.set_title(meta_string.split('|')[2].split('/')[-2])
        book_name = meta_string.split('|')[2].split('/')[-2]
    elif meta_string.find('https://www.isotls.com/') != -1:
        book.set_identifier(meta_string.split('|')[2])
        book.set_title(meta_string.split('|')[2].split('/')[-2])
        book_name = meta_string.split('|')[2].split('/')[-2]
    elif meta_string.find('https://noblemtl.com/series/') != -1:
        book.set_identifier(meta_string.split('|')[2])
        book.set_title(meta_string.split('|')[2].split('/')[-2])
        book_name = meta_string.split('|')[2].split('/')[-2]
    elif meta_string.find('webnovel.com') != -1:
        book.set_identifier(meta_string.split('|')[2][:30])
        book.set_title(meta_string.split('|')[3][:30])
        book_name = meta_string.split('|')[3][:30]
    elif meta_string.find('scribblehub.com') != -1:
        book.set_identifier(meta_string.split('|')[2])
        book.set_title(meta_string.split('|')[2].split('/')[-2])
        book_name = meta_string.split('|')[2].split('/')[-2]
    elif meta_string.find('novelbin.com') != -1:
        book.set_identifier(meta_string.split('|')[2])
        book.set_title(meta_string.split('|')[3])
        book_name = meta_string.split('|')[3]
    elif meta_string.find('Следов:') != -1:
        book.set_identifier(meta_string.split('|')[0].split(': ')[-1])
        book.set_title(meta_string.split('|')[0].split('/')[-1])
        book_name = meta_string.split('|')[0].split('/')[0]
        book.add_author(meta_string.split('|')[4])
    elif meta_string.find('wuxiabox.com') != -1:
        book.set_identifier(filename.split('.')[0])
        book.set_title(filename.split('.')[0])
        book_name = filename.split('.')[0]
    elif meta_string.find('allnovelbook.com') != -1:
        parsed_name = parsers['allnovelbook.com'](meta_string)
        if parsed_name:
            book.set_identifier(parsed_name[:30])
            book.set_title(parsed_name[:30])
            book_name = parsed_name[:30]
        else:
            book.set_identifier(meta_string.split('|')[2][:30])
            book.set_title(meta_string.split('|')[2][:30])
            book_name = meta_string.split('|')[2][:30]
    elif meta_string.find('fucknovelpia.com') != -1:
        parsed_name = parsers['fucknovelpia.com'](meta_string)
        if parsed_name:
            book.set_identifier(parsed_name[:30])
            book.set_title(parsed_name[:30])
            book_name = parsed_name[:30]
        else:
            book.set_identifier(meta_string.split('|')[2][:30])
            book.set_title(meta_string.split('|')[2][:30])
            book_name = meta_string.split('|')[2][:30]
    else:
        book.set_identifier(meta_string.split('|')[2][:30])
        book.set_title(meta_string.split('|')[2][:30])
        book_name = meta_string.split('|')[2][:30]
    return book_name


def epub_to_xlsx(file, index):
    import os
    book = epub.read_epub(file)
    items = list(book.get_items_of_type(ebooklib.ITEM_DOCUMENT))
    print(f'{file}-{len(items) // 300}')
    wbs = []
    chapter_num = 0
    i = 0

    # Гарантируем существование папки temp и корректный путь к файлу
    out_dir = "temp"
    os.makedirs(out_dir, exist_ok=True)
    base_name = os.path.splitext(os.path.basename(file))[0]

    while chapter_num < len(items):
        out_path = os.path.join(out_dir, f'{index}-{i}_{base_name}.xlsx')
        wb = xlsxwriter.Workbook(out_path)
        wbs.append(wb)
        sheet = wb.add_worksheet()
        new_it_index = 0
        cumulative_sum = 0
        print('ok')
        while cumulative_sum < 2_500_000 and chapter_num < len(items):
            result_text = erase_tags_from_text(items[chapter_num])
            new_it_index = split_xlsx_for_rows(sheet, new_it_index, result_text)
            cumulative_sum += len(result_text)
            chapter_num += 1
        print(f'file saving - {i} = {cumulative_sum} symbols')
        wb.close()
        i += 1

    '''
    for i in range(0, len(items) // 300 + 1):
        wb = xlsxwriter.Workbook(f'temp/{index}-{i}_{".".join(file.split("/")[-1].split(".")[:-1])}.xlsx')
        wbs.append(wb)
        sheet = wb.add_worksheet()
        new_it_index = 0
        print('ok')
        sum_text_len = 0
        for item_index, elem in enumerate(items[300*i+1:300*(i+1)]):
            result_text = erase_tags_from_text(items[item_index + 300*i])
            new_it_index = split_xlsx_for_rows(sheet, new_it_index, result_text)
            sum_text_len += len(result_text)
        print(f'file saving - {i} = {sum_text_len} symbols')
        wb.close()
    '''

def split_xlsx_for_rows(sheet: Worksheet, item_index: int, result_text: str, max_row_chars: int = 30000):
    if len(result_text) > max_row_chars:
        absatz = result_text.split('|')
        sp_text_1 = absatz[0]
        is_first = True
        for index, elem in enumerate(absatz[1:]):
            if len(sp_text_1) + len(elem) > max_row_chars:
                if is_first:
                    sheet.write(item_index + 1, 0, '' + sp_text_1)
                    is_first = False
                else:
                    sheet.write(item_index + 1, 0, '$$$ ' + sp_text_1)

                item_index += 1
                sp_text_1 = elem
            else:
                sp_text_1 = sp_text_1 + '|' + elem
        if len(sp_text_1) > 0:
            sheet.write(item_index + 1, 0, '$$$ ' + sp_text_1)
            item_index += 1
    else:
        sheet.write(item_index + 1, 0, result_text)
        item_index += 1
    return item_index


def erase_tags_from_text(epub_text_item):
    out = 0
    text = epub_text_item.get_body_content().decode("utf-8")
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
                curr_text = ''
                out += 1
                prev = '>'
            else:
                curr_text = ''
                out -= 1
                prev = '>'
        index += 1
    return result_text
