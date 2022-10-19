from datetime import datetime
from datetime import date
from datetime import timedelta

def get_start_date():
    
    today = date.today()

    dateObj1 = today - timedelta(7)
    startDate = datetime.strftime(dateObj1, "%d.%m.%Y")
    
    return startDate

def get_end_date():
    
    today = date.today()
    
    dateObj2 = today - timedelta(1)
    endDate = datetime.strftime(dateObj2, "%d.%m.%Y")

    return endDate

def get_data_filename():
    
    today = date.today()

    dateObj1 = today - timedelta(7)
    part1 = datetime.strftime(dateObj1, "%d%m%Y")

    dateObj2 = today - timedelta(1)
    part2 = datetime.strftime(dateObj2, "%d%m%Y")

    filename = "Sähkö_" + part1 + "-" + part2 + ".csv"
    return filename
