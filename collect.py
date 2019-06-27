#!/usr/bin/env python
#
#  Copyright (C) 2017, Hewlett-Packard Development Company
#  Author: Dave Brookshire <dsb@hpe.com>
#
#
import time
import serial
import platform
import gspread
from time import gmtime, strftime, localtime
from datetime import date
from oauth2client.service_account import ServiceAccountCredentials
import re
import sqlite3
import pprint
from math import floor
from rhusb.sensor import RHUSB
LOCATION="Ian's Office"

delay = 10


conn = sqlite3.connect('/home/mq20151400/python-rhusb/temperature.db')

try:
    c = conn.cursor()
    c.execute("select count(*) from temperature;")

except:
    makeC = conn.cursor()
    makeC.executescript("""
CREATE TABLE temperature(
    tstamp timestamp DEFAULT CURRENT_TIMESTAMP PRIMARY KEY,
    rdate date,
    sensorReading Text,
    c   NUMERIC,
    rh  NUMERIC
);

create index rdate on temperature (rdate);

""")


if __name__ == '__main__':
    scope = ['https://spreadsheets.google.com/feeds']
    creds = ServiceAccountCredentials.from_json_keyfile_name('/home/mq20151400/python-rhusb/client_secret.json', scope)
    client = gspread.authorize(creds)
    # Find a workbook by name and open the first sheet
    # Make sure you use the right name here.
    spreadsheet = client.open_by_key("1hZVJ_tUHWoiWUdT5KHX0LjxG1Ah3JtXH-g1dhf5dd3M")
    #sheet = spreadsheet.sheet1
    summary = spreadsheet.get_worksheet(0)
    


    
#    print("Platform: {0}".format(platform.system()))
    if platform.system() == "Windows":
        device = "COM4"
    else:
        device = "/dev/ttyUSB0"
#    print("Device: {0}".format(device))
#    print()

    try:
        sens = RHUSB(device=device)
	print(sens.C(), sens.H() )
        #pprint.pprint((sens.C(), sens.H()))

        c = sens.C()
        h = sens.H()

        sensor = "{0} {1}".format(c,h)

        cCut=re.search("([0-9.-]+) C",sensor).group(1)
        hCut=re.search("([0-9.-]+) %RH",sensor).group(1)
        today = strftime("%Y-%m-%d", localtime())

        conn.execute("INSERT INTO temperature(tstamp, rdate, sensorReading, c, rh) VALUES(?, ?, ?, ?, ?)",(strftime("%Y-%m-%d %H:%M:%S", localtime()),today, sensor, cCut, hCut))
        conn.commit()
        #print("PA: [{0}]".format(sens.PA()))
        #         print("C: [{0}]".format(sens.C()))
        #        print("F: [{0}]".format(sens.F()))
        #         print("H: [{0}]".format(sens.H()))
        #        print("\nStarting periodic readings every {0} seconds".format(delay))

        #        while True:
        # C=re.search("([0-9.]+)",sens.C()).group(1)
        # H=re.search("([0-9.]+)",sens.H()).group(1)
        # row=[strftime("%Y-%m-%d %H:%M:%S", localtime()), "{0} C".format(C),"{0}%".format(H),LOCATION]
        # cleanrow=[strftime("%Y-%m-%d %H:%M:%S", localtime()), C,H,LOCATION]
        #sheet.append_row(row)
        #clean.append_row(cleanrow)

        summary.update_acell('B1', strftime("%Y-%m-%d %H:%M:%S", localtime()))
        summary.update_acell('C1', sensor)

        #cur = conn.cursor()
        #cur.execute("select c, tstamp from temperature where rdate = '{0}' group by rdate having c = max(c) limit 1".format(today))
	histogram = []
        for row in conn.execute("select tstamp, c, rh from temperature where strftime('%H', tstamp) % 3 = 0 and strftime('%M', tstamp) = '01' order by tstamp desc"):
#        for row in conn.execute("select tstamp, c, rh from temperature  order by tstamp desc"):
            print(row)
            histogram.append(row[0])
	    histogram.append(row[1])
            histogram.append(row[2])

        #clearcell_list = summary.range('A5:D100')
        #for cell in clearcell_list:
        #    cell.value = ''
	
	updatecell_list = summary.range('A5:C1000')
	for i, cell in enumerate(updatecell_list):
	  try: 
            cell.value=histogram[i]
          except:
            break
        summary.update_cells(updatecell_list)

	temp_mode = []
	for row in conn.execute("select rdate, round(c,0), count(round(c,0)) from temperature group by rdate, round(c,0) order by rdate desc, c"):
		temp_mode.append(row[0])
		temp_mode.append(row[1])
		temp_mode.append(row[2])

	updatecell_list = summary.range('e5:g1000')
	for i, cell in enumerate(updatecell_list):
	  try: 
            cell.value=temp_mode[i]
          except:
            break
        summary.update_cells(updatecell_list)

	temp_mode = []
	for row in conn.execute("select rdate, round(rh,0), count(round(rh,0)) from temperature group by rdate, round(c,0) order by rdate desc, rh"):
		temp_mode.append(row[0])
		temp_mode.append(row[1])
		temp_mode.append(row[2])

	updatecell_list = summary.range('i5:k1000')
	for i, cell in enumerate(updatecell_list):
	  try: 
            cell.value=temp_mode[i]
          except:
            break
        summary.update_cells(updatecell_list)


	# for i, histo in enumerate(histogram):
	
	#    summary.update_cell(i+5, 1, histo[0])
	#    summary.update_cell(i+5, 2, histo[1])
	#    summary.update_cell(i+5, 3, histo[2])

        print("{0}\t{1}\t{2}".format( strftime("%Y-%m-%d %H:%M:%S", localtime()), c, h))

    except serial.serialutil.SerialException:
        print("Error: Unable to open RH-USB Serial device {0}.".format(device))
    except Exception as e :
        print(e)
        pass        
