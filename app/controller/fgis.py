import os.path
import time
import pandas as pd
import requests
import json
from zipfile import ZipFile
from pyhtml2pdf import converter
from datetime import datetime
from config import Config


class FileXLSX:
    def __init__(self, xlsx_name):
        self.name = xlsx_name
        self.data = []

        file = pd.read_excel(self.name, sheet_name='Лист1')

        for item in file.values:
            org = item[2].strip().replace('\n', '')
            org = org.replace(' ', '?')
            self.data.append([item[0].strip(), item[1].strip(), org])

        os.remove(xlsx_name)



class VerificationCert:
    def __init__(self, xlsx_list):
        self.xlsx = xlsx_list
        self.apidata = []
        self.vri_id = []
        error_list = []

        for item in self.xlsx:
            url_api = f'https://fgis.gost.ru/fundmetrology/eapi/vri?' \
                      f'mi_number={item[0]}&' \
                      f'mit_number={item[1]}&org_title={item[2]}&' \
                      f'verification_date_start=2021-01-01&' \
                      f'sort=verification_date+desc'
            try:
                r = requests.get(url_api)
                # Save api data to json array
                result_json = json.loads(r.text)
                self.apidata.append(result_json)
                # Get only tail of doc number
                self.vri_id.append(result_json['result']['items'][0]['vri_id'])
                time.sleep(0.5)
            except:
                error_list.append(f'{item[0]} {item[1]} {item[2]} is not found in Arshin site')
                time.sleep(0.5)

        if len(error_list) > 0:
            with open(os.path.join(Config.DOWNLOAD_FOLDER, 'error_api.txt'), 'w') as f:
                for e in error_list:
                    f.write(e + '\n')
        else:
            if os.path.exists(os.path.join(Config.DOWNLOAD_FOLDER, 'error_api.txt')):
                os.remove(os.path.join(Config.DOWNLOAD_FOLDER, 'error_api.txt'))


    def get_sert(self):
        error_list = []
        for vri_id in self.vri_id:
            source = 'https://fgis.gost.ru/fundmetrology/cm/results/' + vri_id
            target = os.path.join(Config.DOWNLOAD_FOLDER, f'{vri_id[2:]}.pdf')
            try:
                converter.convert(source=source, target=target, install_driver=False)
            except Exception as e:
                error_list.append(f'{vri_id} not found for PDF')
                error_list.append(e)
                error_list.append('------------')

        if len(error_list) > 0:
            with open(os.path.join(Config.DOWNLOAD_FOLDER, 'error_pdf.txt'), 'w') as f:
                for e in error_list:
                    f.write(e + '\n')

    def add_to_zip(self):
        now = datetime.now()
        zip_file_name = f'arshin_{now.year}-{now.month}-{now.day}.zip'
        zip_obj = ZipFile(os.path.join(Config.DOWNLOAD_FOLDER, zip_file_name), 'w')

        for vri_id in self.vri_id:
            pdf_to_zip = os.path.join(Config.DOWNLOAD_FOLDER, f'{vri_id[2:]}.pdf')
            if os.path.isfile(pdf_to_zip):
                zip_obj.write(pdf_to_zip)
                os.remove(pdf_to_zip)

        error_api = os.path.join(Config.DOWNLOAD_FOLDER, 'error_api.txt')
        error_pdf = os.path.join(Config.DOWNLOAD_FOLDER, 'error_pdf.txt')

        if os.path.isfile(error_api):
            zip_obj.write(error_api)
            os.remove(error_api)

        if os.path.isfile(error_pdf):
            zip_obj.write(error_pdf)
            os.remove(error_pdf)

        zip_obj.close()
        return zip_file_name
