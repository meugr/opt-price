import os
from datetime import datetime


class Config():
    xml_link = 'url/to/xml'
    current_directory = os.path.dirname(os.path.abspath(__file__))
    xml_name = 'filename.xml'
    excel_name = 'filename.xlsx'
    urls_name = 'urls.json'
    backup_path = os.path.join(
        current_directory, 'backup',
        f'{datetime.now().strftime("%d.%m.%Y")}.xls')
    path_xml = os.path.join(current_directory, xml_name)
    path_excel = os.path.join(current_directory, excel_name)
    path_urls = os.path.join(current_directory, urls_name)  # json file 

    email_logpass_sender = ('email address', 'password')
    addressee = 'email address'
