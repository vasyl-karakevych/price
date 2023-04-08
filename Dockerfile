FROM python

COPY . /telegram-price
WORKDIR /telegram-price
RUN pip install openpyxl pyTelegramBotAPI
CMD [ "python", "./telegram.py" ]