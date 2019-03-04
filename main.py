from config import Config
import requests
import json
import os
import openpyxl
from datetime import datetime
from shutil import copyfile
import smtplib
try:
    import lxml.etree.ElementTree as etree
except ModuleNotFoundError:
    import xml.etree.ElementTree as etree


def get_xml_file(xml_link, path_xml):
    """
    Get and write XML from URL.
    """
    xml_price = requests.get(xml_link).content
    with open(path_xml, "wb") as f:
        f.write(xml_price)


def make_dict(path_xml, path_urls):
    """
    Сreates a dictionary from an XML file accessible by xml_name.
    """
    dict_values = {}
    with open(path_urls, "r") as f:
        urls = json.load(f)  # urls = {art1: url1, art2: url2, ...}
    tree = etree.parse(path_xml)
    root = tree.getroot()
    categories = root[0][4]
    offers = root[0][5]
    for offer in offers:
        identificator = offer.get('id')
        available = offer.get('available')
        name = offer.find('name').text
        price = offer.find('price').text
        weight = offer.find('param[@name="Вес"]').text

        category_id = offer.find('categoryId').text
        category = categories.find(f'category[@id="{category_id}"]').text

        vendor = offer.find('vendor')  # maybe None
        vendor = vendor.text if vendor is not None else 'Без бренда'

        flavor = offer.find('param[@name="Вкус"]')  # maybe None
        flavor = flavor.text if flavor is not None else 'без вкусов'

        url = urls.get(identificator) if identificator in urls.keys() else ''

        dict_values[identificator] = {'name': name,
                                      'price': price,
                                      'weight': weight,
                                      'category': category,
                                      'vendor': vendor,
                                      'flavor': flavor,
                                      'available': available,
                                      'url': url}
    return dict_values


def write_in_excel(dict_price, path_excel):
    """
    Writes values from dictionary to an Excel price list.
    Clear 50 rows after last row.
    """
    row = 7  # starting row for write
    wb = openpyxl.load_workbook(path_excel)
    ws = wb.active

    for art in dict_price:
        row += 1
        ws[f'A{row}'] = art
        ws[f'B{row}'] = dict_price.get(art).get('name')
        ws[f'C{row}'] = dict_price.get(art).get('category')
        ws[f'D{row}'] = dict_price.get(art).get('vendor')
        ws[f'E{row}'] = dict_price.get(art).get('flavor')
        ws[f'F{row}'] = dict_price.get(art).get('weight')
        ws[f'G{row}'] = dict_price.get(art).get('price')
        ws[f'H{row}'].hyperlink = dict_price.get(art).get('url')
    ws['A1'] = f'Цены и наличие\nактуальны на\n\
{datetime.now().strftime("%d.%m.%Y")}'
    for i in range(row + 1, row + 50):  # clear 50 row after last row
        for j in range(1, 9):
            ws.cell(row=i, column=j, value='')
    wb.save(path_excel)


def send_error_message(email_logpass_sender, addressee):
    """
    Send e-mail with error message.
    email_logpass_sender = (login, password)
    """
    HOST = ('smtp.yandex.ru', 465)
    SUBJECT = 'Error opt-price generator'
    TO = addressee
    FROM = email_logpass_sender[0]
    text = 'Error generating xlsx, urgently check the operation \
of the script "opt-price".'

    BODY = "\r\n".join((
           "From: %s" % FROM,
           "To: %s" % TO,
           "Subject: %s" % SUBJECT,
           "", text))

    server = smtplib.SMTP_SSL(*HOST)
    server.login(*email_logpass_sender)
    server.sendmail(FROM, TO, BODY)
    server.quit()


try:
    os.chdir(os.path.dirname(os.path.abspath(__file__)))
    get_xml_file(Config.xml_link, Config.path_xml)
    dict_price = make_dict(Config.path_xml, Config.path_urls)
    write_in_excel(dict_price, Config.path_excel)
    try:
        os.mkdir('backup')  # creates a directory when you first start
    except FileExistsError:
        print('Directory "backup" exist')
    copyfile(Config.path_excel, Config.backup_path)
except:
    print('ERROR!!')
    send_error_message(Config.email_logpass_sender, Config.addressee)
