from datetime import datetime

def text_to_date(text_date):
    year = int(text_date[0:4])
    month = int(text_date[5:7])
    day = int(text_date[8:10])
    return datetime(year, month, day)
