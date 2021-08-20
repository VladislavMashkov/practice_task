import smtplib
import os
import mimetypes
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
import pyexcel as p
from openpyxl import load_workbook
import csv

addr_from = "cmailclient@gmail.com"
password = "cmailclientqwerty"
addr_to = "Vlad_mashkow@mail.ru"
files = ["result.csv"]

def send_email(addr_to, msg_subj, msg_text, files):

    msg = MIMEMultipart()
    msg['From'] = addr_from
    msg['To'] = addr_to
    msg['Subject'] = msg_subj

    body = msg_text
    msg.attach(MIMEText(body, 'plain'))

    process_attachement(msg, files)

    server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
    server.login(addr_from, password)
    server.send_message(msg)
    server.quit()


def process_attachement(msg, files):
    for f in files:
        if os.path.isfile(f):
            attach_file(msg, f)
        elif os.path.exists(f):
            dir = os.listdir(f)
            for file in dir:
                attach_file(msg, f + "/" + file)


def attach_file(msg, filepath):
    filename = os.path.basename(filepath)
    ctype, encoding = mimetypes.guess_type(filepath)
    if ctype is None or encoding is not None:
        ctype = 'application/octet-stream'
    maintype, subtype = ctype.split('/', 1)
    if maintype == 'text':  # Если текстовый файл
        with open(filepath) as fp:
            file = MIMEText(fp.read(), _subtype=subtype)
            fp.close()
    else:
        with open(filepath, 'rb') as fp:
            file = MIMEBase(maintype, subtype)
            file.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(file)
    file.add_header('Content-Disposition', 'attachment', filename=filename)
    msg.attach(file)

headers = {'connection': 'keep-alive',
           'cache-control': 'max-age=0',
           'sec-ch-ua': '"Chromium";v="92", " Not A;Brand";v="99", "Google Chrome";v="92"',
           'sec-ch-ua-mobile': '?0',
           'origin': 'https://rmsp.nalog.ru',
           'upgrade-insecure-requests': '1',
           'DNT': '1',
           'Content-Type': 'application/x-www-form-urlencoded',
           'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/92.0.4515.131 Safari/537.36',
           'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
           'Sec-Fetch-Site': 'same-origin',
           'Sec-Fetch-Mode': 'navigate',
           'Sec-Fetch-User': '?1',
           'Sec-Fetch-Dest': 'document',
           'Referer': 'https://rmsp.nalog.ru/search.html?mode=inn-list',
           'Accept-Language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
           'Cookie': '_ym_uid=1626367445496233330; _ym_d=1626367445; _ym_isad=2; JSESSIONID=8E25F7E5601D4E5052DE1A5CC9763D4E',}

class OrganizationRecord:
    def __init__(self, fullName, innValue, ogrnValue):
        self.fullName = fullName
        self.innValue = str(innValue)
        self.ogrnValue = str(ogrnValue)
        self.category  = 'Значение'
        self.location = 'Значение'
        self.totalAmountOfTransfers = 'Значение'
        self.totalVolumeOfBankCommissions = 'Значение'
        self.amountSubsidies = 'Значение'
    def writeInfoAboutCategory(self, category):
        self.category = category
    def writeInfoAboutLocation(self, locationList):
        self.location = ' '.join(locationList)
    def returnOrganizationInfo(self, number):
        return [number, self.fullName, self.innValue, self.ogrnValue, self.category, self.location, self.totalAmountOfTransfers, self.totalVolumeOfBankCommissions, self.amountSubsidies]


def makeCsvFile():
    header = ['№ п/п',
              'Полное наименование субъекта МСП',
              'ИНН субъекта МСП',
              'ОГРН субъекта МСП (при наличии)',
              'Категория субъекта МСП (микро, малое, среднее)',
              'Место нахождения (место жительства) субъекта МСП (субъект Российской Федерации)',
              'Суммарный размер переводов, осуществленных физическими лицами в пользу субъектов МСП в СБП, рублей',
              'Суммарный объём банковских комиссий за переводы денежных средств, осуществленных физическими лицами в пользу субъектов МСП в СБП, рублей',
              'Размер субсидий за отчётный период, рублей']
    with open('result.csv', 'w', encoding='UTF8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(header)


def makeRequest(innValue):
    session = requests.Session()
    retry = Retry(connect=3, backoff_factor=0.5)
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)

    response = session.get('https://rmsp.nalog.ru/report.xlsx?mode=inn-list&page=1&innList=' + str(
            innValue) + '&pageSize=10&sortField=NAME_EX&sort=ASC',
        headers=headers)
    if (response.status_code != 200):
        return False
    return response.content


def saveDataFromRequest(content):
    output = open('test.xlsx', 'wb')
    output.write(content)
    output.close()
    return True


def loadDataFromFile():
    workbook = load_workbook(filename='test.xlsx', read_only=False)
    worksheet = workbook.active
    data = []
    if worksheet == None:
        return False
    for row1 in worksheet.rows:
        curRow = []
        for cell in row1:
            curRow.append(cell.value)
        data.append(curRow)
    workbook.close()
    return data


def findRowInData(innValue, ogrnValue, data):
    for row in data:
        if innValue in row and ogrnValue in row:
            return row
    return False


def writeRowIntoCsv(curOrganization, number):
    with open('result.csv', 'a', encoding='UTF8', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(curOrganization.returnOrganizationInfo(number))

def main():
    makeCsvFile()
    p.save_book_as(file_name='SME+CORP+огрн.xls',
                   dest_file_name='SME+CORP+огрн.xlsx')
    workbook = load_workbook(filename='SME+CORP+огрн.xlsx', read_only=True)
    worksheet = workbook.active
    count = 0
    for row in worksheet.rows:
        if count == 0:
            count += 1
            continue
        curRow = []
        for cell in row:
            curRow.append(cell.value)
        currentOrganization = OrganizationRecord(curRow[2], curRow[3], curRow[4])
        resultFromRequest = makeRequest(currentOrganization.innValue)
        if(not resultFromRequest):
            print("Не было получено ответа от сайта ФНС по ИНН", currentOrganization.innValue)
            writeRowIntoCsv(currentOrganization, count)
            count += 1
            continue
        isSaved = saveDataFromRequest(resultFromRequest)
        if(not isSaved):
            print("Не было произведено сохранение полученного файла по ИНН", currentOrganization.innValue)
            writeRowIntoCsv(currentOrganization, count)
            count += 1
            continue
        dataFromFile = loadDataFromFile()
        if(not dataFromFile):
            print("По заданным параметрам не найдено сведений в едином реестре субъектов малого и среднего предпринимательства. ИНН", curRow[3])
            writeRowIntoCsv(currentOrganization, count)
            count += 1
            continue
        expectedRow = findRowInData(currentOrganization.innValue, currentOrganization.ogrnValue, dataFromFile)
        if(not expectedRow):
            print("Не было найдено записи с ИНН", currentOrganization.innValue, "и ОГРН", currentOrganization.ogrnValue, "в исходном документе")
            writeRowIntoCsv(currentOrganization, count)
            count += 1
            continue
        currentOrganization.writeInfoAboutCategory(expectedRow[3])
        currentOrganization.writeInfoAboutLocation(expectedRow[7:10])
        writeRowIntoCsv(currentOrganization, count)
        count += 1
    workbook.close()
    send_email(addr_to, "Test", "Test", files)

if __name__ == "__main__":
    main()