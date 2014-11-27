# -*- coding: utf-8 -*-
#-------------------------------------------------------------------------------
# Name:        pyall2txt
# Created:     23.05.2014
# Version:     0.3
# Copyright:   (c) 2014, Евгений
# Licence:     Apache License version 2.0
#-------------------------------------------------------------------------------

from sys import argv
from os.path import isfile, splitext
from zipfile import ZipFile
from xml.etree.ElementTree import parse, XML
from html.parser import HTMLParser


def help():
    print('Для конвертации .doc, .docx, .odt, .htm, .html, .fb2 документов:\n\
    pyall2txt.py имя_файла\n\
Если в имени файла используются пробелы, то возьмите его в кавычки')
    exit()


def ok():
    print('Преобразование выполнено!')


def is_ascii(s):
    '''is_ascii(byte) -> bool

    Проверяет один байт на наличие в русском или англ. алфавитах,
    а так же среди знаков припинания.
    '''
    if (ord(s) in range(32, 122) ) or (ord(s) in range(1039, 1104))\
    or (ord(s) in (9,10,13)):
        return True
    return False


class Converter:
    def doc_txt(self, ifname, ofname):
        with open(ifname, 'rb') as ffr:
            data = ffr.read()
        ## Определяет начало первого и конец последнего блоков текста
        start = data[:2800].rfind(b'\x00'*50) + 50
        end = start + data[start:].find(b'\x0d' + b'\x00'*10)
        if end < start: w95, end = True, start + data[start:].find(b'\x00'*10) ## Для 95 ворда
        else: w95 = False
        data = data[start:end]

        text, itemp = str(), 0

        ## Цикл читает получившийся текст по 2 байта и декодирует в читаемый вид
        while True:
            if w95 == False:
                try:
                    t = data[itemp:itemp+2]
                    if is_ascii(t.decode('utf-16le')) is True:
                        text += t.decode('utf-16le')

                except UnicodeDecodeError:
                    t = t.decode('cp1251', 'ignore')
                    if len (t) > 0:
                        for i in t:
                            if is_ascii(i) is True:
                                text += i
                itemp += 2
                if itemp >= len(data): break

        ## Если это кодировка cp1251, то декодирует весь текст в один приём
            else:
                t = data.decode('cp1251', 'ignore')
                for i in t:
                    if is_ascii(i) is True:
                        text += i
                break

        with open(ofname, "wb") as ffw:
            ffw.write(b'\xef\xbb\xbf' + text.encode())
        return 1


    def docx_txt(self, ifname, ofname):
        with ZipFile(ifname, 'r') as y:
            z = y.read('word/document.xml')

        tree = XML(z)
        text = str()
        p = tree.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')
        for x in p:
            text += '\n'
            for y in x:
                for z in y:
                    if z.text is not None: text += (z.text)

        with open(ofname, "wb") as ffw:
            ffw.write(b'\xef\xbb\xbf' + text.encode())
        return 1


    def odt_txt(self, ifname, ofname):
        with ZipFile(ifname, 'r') as y:
            z = y.read('content.xml')

        tree = XML(z)
        text = str()
        p = tree.findall('.//{urn:oasis:names:tc:opendocument:xmlns:text:1.0}p')
        for x in p:
            text += '\n'
            for y in x:
                if y.text is not None: text += (y.text)

        with open(ofname, "wb") as ffw:
            ffw.write(b'\xef\xbb\xbf' + text.encode())
        return 1


    def fb2_txt(self, ifname, ofname):
        with open(ifname, "rb") as ffr:
            data = ffr.read()

        try: data = data.decode('utf-8')
        except UnicodeDecodeError: data = data.decode('cp1251')

        class MyHTMLParser(HTMLParser):
            text = str()
            def handle_data(self, data):
                MyHTMLParser.text += data

        MyHTMLParser(strict=False).feed(data)

        with open(ofname, "wb") as ffw:
            ffw.write(b'\xef\xbb\xbf' + MyHTMLParser.text.encode())
        return 1


def convert(ifname):

    ofname, iftype = splitext(ifname)[0] + '.txt', splitext(ifname)[1].lower()

    con = Converter()

    if iftype == '.doc':
        con.doc_txt(ifname, ofname)
    elif iftype == '.docx':
        con.docx_txt(ifname, ofname)
    elif iftype == '.odt':
        con.odt_txt(ifname, ofname)
    elif iftype in ('.fb2', '.html', '.htm'):
        con.fb2_txt(ifname, ofname)


def main():
    print('\"Pyall2txt\"')

    if len(argv) != 2:
        help()
        exit()

    if convert(argv[1]) == 1: ok()

if __name__ == '__main__':
    main()
