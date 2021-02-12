import datetime as dt


def getTextData():
    date = dt.datetime.today()
    textdata = str(date.month) + "/" + str(date.day)
    return textdata
