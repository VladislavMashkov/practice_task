import os
import csv

resultFile: str = os.environ['RESULT_FILE_NAME']

def makeCsvFile() -> None:
    header: list = ['№ п/п',
                    'Полное наименование субъекта МСП',
                    'ИНН субъекта МСП',
                    'ОГРН субъекта МСП (при наличии)',
                    'Категория субъекта МСП (микро, малое, среднее)',
                    'Место нахождения (место жительства) субъекта МСП (субъект Российской Федерации)',
                    'Суммарный размер переводов, осуществленных физическими лицами в пользу субъектов МСП в СБП, рублей',
                    'Суммарный объём банковских комиссий за переводы денежных средств, осуществленных физическими лицами в '
                    'пользу субъектов МСП в СБП, рублей',
                    'Размер субсидий за отчётный период, рублей']
    with open(resultFile, 'w', encoding='UTF8', newline='') as f:
        writer: csv.writer = csv.writer(f)
        writer.writerow(header)

def writeRowIntoCsv(curOrganization: object) -> None:
    with open(resultFile, 'a', encoding='UTF8', newline='') as f:
        writer: csv.writer = csv.writer(f)
        writer.writerow(curOrganization.returnOrganizationInfo())
