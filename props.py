#!/usr/bin/env python

"""
"""

import urllib.request
import re
import xlsxwriter
import operator
import collections
import sys
import ast
import os
from datetime import date, timedelta, datetime
from bs4 import BeautifulSoup as bs

allProps = collections.OrderedDict()

def _format_url(date):
    """Format ESPN link to scrape records from."""
    link = ['http://streak.espn.go.com/en/entry?date=' + date]
    print(date)
    return link[0]

def scrape_props(espn_page):
    """Scrape ESPN's pages for data."""
    global allProps
    url = urllib.request.urlopen(espn_page)
##    print url.geturl()
    soup = bs(url.read(), ['fast', 'lxml'])

    props = soup.find_all('div', attrs={'class':'matchup-container'})

    for prop in props:
        propArray = []
        
        time = prop.find('span', {'class':'startTime'})['data-locktime']
        time = time[:-4]
        format_time = datetime.strptime(time, '%B %d, %Y %I:%M:%S %p')
                
        sport = prop.find('div', {'class':'sport-description'})
        if sport:
            propArray.append(sport.text)
        else:
            propArray.append('Adhoc')

        title = prop.find('div', {'class':['gamequestion','left']}).text.split(': ')
        propArray.append(title[0])
        propArray.append(title[1])

        overall_percentage = prop.find('div', {'class':'progress-bar'})
        propArray.append(overall_percentage['title'].split(' ')[3])

        percentages = prop.find_all('span', {'class':'wpw'})
        
        info = prop.find_all('span', {'class':'winner'})
        temp = info[0].parent.find_all('span', {'id':'oppAddlText'})
        [rec.extract() for rec in temp]

        temp = info[1].parent.find_all('span', {'id':'oppAddlText'})
        [rec.extract() for rec in temp]
        
        if info[0].contents[0].name == 'img':
            propArray.append(info[0].parent.get_text())
            propArray.append(percentages[0].text)
            propArray.append(info[1].parent.get_text()[1:])
            propArray.append(percentages[1].text)
        else:
            propArray.append(info[1].parent.get_text())
            propArray.append(percentages[1].text)
            propArray.append(info[0].parent.get_text()[1:])
            propArray.append(percentages[0].text)

        allProps[format_time.date()].append(propArray)

def write_to_excel():
    global allProps

    wb = xlsxwriter.Workbook('Streak.xlsx')

    ws = wb.add_worksheet('All Props')
    #ws.set_column(0, 0, 2.29)
    #ws.set_column(1, 1, 14.14)
    #ws.set_column(5, 5, 2.29)

    date_format = wb.add_format({'num_format': 'mm/dd/yy'})
    format_percentage = wb.add_format({'num_format':'0.0"%"'})
    format_percentage2 = wb.add_format({'num_format':'0.00"%"'})

    i = 0
    for date in allProps:
        for prop in allProps[date]:     
            ws.write(i, 0, date, date_format)
            ws.write(i, 1, prop[0])
            ws.write(i, 2, prop[1])
            ws.write(i, 3, prop[2])
            ws.write(i, 4, float(prop[3][:-1]), format_percentage2)
            ws.write(i, 5, prop[4])
            ws.write(i, 6, float(prop[5][:-1]), format_percentage)
            ws.write(i, 7, prop[6])
            ws.write(i, 8, float(prop[7][:-1]), format_percentage)
            i+=1   
    wb.close()

def main(argv):
    global allProps

    for x in range(0,int(argv[0])):
        newDate = date.today() - timedelta(days=x + 1)
        allProps[newDate] = []
        scrape_props(_format_url(newDate.strftime('%Y%m%d')))

    write_to_excel()

if __name__ == '__main__':
    import time
    start = time.time()
    main(sys.argv[1:])
    print(time.time() - start, 'seconds')
