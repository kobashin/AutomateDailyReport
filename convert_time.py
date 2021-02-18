# ref : https://gammasoft.jp/blog/python-string-format/
import math

def decimal2hhmm(textdecimal):
    decimal = float(textdecimal)
    hour = math.floor(decimal)
    minute = int(60 * (decimal - hour))
    hhmm = "{:02}:{:02}".format(hour, minute)
    return hhmm

def hhmm2decimal(texthhmm):
    hh_mm = texthhmm.split(':')
    hour = float(hh_mm[0])
    minute = float(int(hh_mm[1]) / 60)
    decimal = hour + minute
    return decimal

'''
print(decimal2hhmm(1.5))        =>  01:30
print(decimal2hhmm(0.5))        =>  00:30
print(hhmm2decimal('00:00'))    =>  0.0
print(hhmm2decimal('00:30'))    =>  0.5
print(hhmm2decimal('01:30'))    =>  1.5
'''