from xlrd import open_workbook
from icalendar import Calendar, Event
from datetime import datetime
from calendar import monthrange
from pytz import UTC, timezone
import time

eventName = 'Metalo ploeg 4 '
book = open_workbook('ploeg4.xls', formatting_info=True)
sheet = book.sheet_by_index(1)
now = datetime.now()
brussels = timezone('Europe/Brussels')

cal = Calendar()
cal.add('prodid', '-//MetaloToIcal//vandesselstijn@gmail.com//')
cal.add('version', '2.0')

for row_index in range(1, 13):
    for col_index in range(1, 32):
        if(sheet.cell(row_index,col_index).value != ''):
            event = Event()
            if(sheet.cell(row_index,col_index).value == 'V'):
                event.add('summary', eventName + 'Vroege')
                event.add('dtstart', datetime(2017, row_index, col_index, 6, 0, 0,tzinfo=brussels))
                event.add('dtend', datetime(2017, row_index, col_index, 14, 0, 0, tzinfo=brussels))
            if(sheet.cell(row_index,col_index).value == 'L'):
                event.add('summary', eventName + 'Late')
                event.add('dtstart', datetime(2017, row_index, col_index, 14, 0, 0, tzinfo=brussels))
                event.add('dtend', datetime(2017, row_index, col_index, 22, 0, 0, tzinfo=brussels))
            if (sheet.cell(row_index, col_index).value == 'N'):
                event.add('summary', eventName + 'Nacht')
                event.add('dtstart', datetime(2017, row_index, col_index, 22, 0, 0, tzinfo=brussels))
                monthTuple = monthrange(2017,row_index)
                monthLength = monthTuple[1]
                if(monthLength > col_index):
                    event.add('dtend', datetime(2017, row_index, col_index+1, 6, 0, 0, tzinfo=brussels))
                else:
                    if(row_index==12):
                        datetime(2017+1, 1, 1, 6, 0, 0, tzinfo=brussels)
                    else:
                        datetime(2017, row_index+1, 1, 6, 0, 0, tzinfo=brussels)

            event.add('dtstamp', datetime(now.year, now.month, now.day, now.hour, 0, 0, tzinfo=brussels))
            event['uid'] = '2017' + str(row_index).zfill(2) + str(col_index).zfill(2) + '@MetaloToIcal'
            event.add('priority', 5)
            cal.add_component(event)

f = open('metalo.ics', 'wb')
f.write(cal.to_ical())
f.close()
