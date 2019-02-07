#!/usr/bin/python
import datetime
import copy

# Global defaults
outputFileName = 'toCalendar.csi'
inputFileName  = 'fromExcel.html'


def writeHeader():
  f = open(outputFileName, 'w')
  f.write('BEGIN:VCALENDAR\nVERSION:2.0\n')
  f.close()


startDateStr = raw_input('Start date of the courses(YYYY-MM-DD)')
endDateStr   = raw_input('End date of the courses(YYYY-MM-DD)')

t = startDateStr.split('-')
f = endDateStr.split('-')

startDate = datetime.date(int(t[0]), int(t[1]), int(t[2])) # Start date of the event, corresponding to the first(0th) column of the excel file
endDate   = datetime.date(int(f[0]), int(f[1]), int(f[2])) # End date of the event, when the final class hour will be held

currentDate = copy.copy(startDate)

writeHeader()

inFile  = open(inputFileName, 'r')
outFile = open(outputFileName, 'a')

inLines = inFile.readlines()
inLine  = inLines[20]

while currentDate < endDate:
  # Here read according to the correct html tags, but time to sleep!! may be tomorrow...

  currentDate = currentDate + datetime.timedelta(days=1)

