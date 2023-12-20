#!/usr/bin/python
from datetime import datetime
import xlrd

now = datetime.now().strftime('%Y%m%dT%H:%M:%S')
# name 日历名称
def set_ics_header(name):
    return "BEGIN:VCALENDAR\n" \
           + "PRODID:NULL\n" \
           + "VERSION:2.0\n" \
           + "CALSCALE:GREGORIAN\n" \
           + "METHOD:PUBLISH\n" \
           + f"X-WR-CALNAME:{name}\n" \
           + "X-WR-TIMEZONE:Asia/Shanghai\n" \
           + f"X-WR-CALDESC:{name}\n" \
           + "BEGIN:VTIMEZONE\n" \
           + "TZID:Asia/Shanghai\n" \
           + "X-LIC-LOCATION:Asia/Shanghai\n" \
           + "BEGIN:STANDARD\n" \
           + "TZOFFSETFROM:+0800\n" \
           + "TZOFFSETTO:+0800\n" \
           + "TZNAME:CST\n" \
           + "DTSTART:19700101T000000\n" \
           + "END:STANDARD\n" \
           + "END:VTIMEZONE\n"


def set_jr_ics(jr, date, uid):  # jr: 节日，date：日期，uid：编序
    return "BEGIN:VEVENT\n" \
           + f"DTSTART;VALUE=DATE:{date}\n" \
           + f"DTEND;VALUE=DATE:{date}\n" \
           + f"DTSTAMP:{date}T000001\n" \
           + f"UID:{date}T{uid:0>6}_jr\n" \
           + f"CREATED:{date}T000001\n" \
           + f"DESCRIPTION:{jr}\n" \
           + f"LAST-MODIFIED:{now}\n" \
           + "SEQUENCE:0\n" \
           + "STATUS:CONFIRMED\n" \
           + f"SUMMARY:{jr}\n" \
           + "TRANSP:TRANSPARENT\n" \
           + "END:VEVENT\n"


def concat_ics(year, jjr_list,rq_list):  # 返回一个完整的ics文件内容
       header = set_ics_header(year)
       # 将节日进行编号，生成list转成字符串
       jr_ics=''.join(list(map(set_jr_ics, jjr_list, rq_list,list(range(len(jjr_list))))))
       return header + jr_ics + 'END:VCALENDAR'

# 保存文件
def save_ics(fname, text):
       with open(fname, 'w', encoding='utf-8') as f:
              f.write(text)

#获取excel内容和sheet
def get_xlsfile(path):
    readfile=xlrd.open_workbook(path)
    num = readfile.nsheets
    return readfile,num

def parse_jjr(table):
    name=table.name
    jjr=list(table.col_values(0))
    rq=list(map(dataformat,table.col_values(1)))

    return name,jjr,rq

def dataformat(date):
      return datetime.strptime(date, '%Y/%m/%d').strftime('%Y%m%d')

if __name__ == '__main__':
    readfile,num = get_xlsfile('F:/ICS/calendar.xls')
    for i in range(num):
        name,jjr_list,rq_list=parse_jjr(readfile.sheets()[i])
        jr_ics = concat_ics(name,jjr_list,rq_list)
        filename = f'calendar_{name}.ics'
        save_ics(filename, jr_ics)
