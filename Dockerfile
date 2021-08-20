FROM python:3

WORKDIR /usr/src/app
ENV MAIL_USER="cmailclient@gmail.com" \
    MAIL_PASSWORD="cmailclientqwerty" \
    MAIL_DESTINATION="Vlad_mashkow@mail.ru" \
    MAIL_SUBJECT="Итоговый файл единого реестра субъектов малого и среднего предпринимательства" \
    MAIL_TEXT="" \
    RESULT_FILE_NAME="result.csv"

COPY ./requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

COPY ./python-service.py ./
COPY ./SME+CORP+огрн.xls ./

CMD [ "python", "./python-service.py" ]
