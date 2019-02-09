#!/usr/bin/python
import datetime
import copy

# Global defaults
outputFileName       = 'toCalendar.csi'
inputFileName        = 'fromExcel.html'
lengthOfALectureHour = 50 # min


def writeHeader():
  f = open(outputFileName, 'w')
  f.write('BEGIN:VCALENDAR\nVERSION:2.0\n')
  f.close()


startDateStr = raw_input('Start date of the courses(YYYY-MM-DD)') #'2019-02-10' #
endDateStr   = raw_input('End date of the courses(YYYY-MM-DD)') #'2019-05-31' #

t = startDateStr.split('-')
f = endDateStr.split('-')

startDate = datetime.datetime(int(t[0]), int(t[1]), int(t[2]), 0, 0, 0) # Start date of the event, corresponding to the first(0th) column of the excel file
endDate   = datetime.datetime(int(f[0]), int(f[1]), int(f[2]), 0, 0, 0) # End date of the event, when the final class hour will be held

currentDate = copy.copy(startDate)

writeHeader()

inFile  = open(inputFileName, 'r')
outFile = open(outputFileName, 'a')

inLines = inFile.readlines()
inLine  = inLines[20]

hours = '' # the hours used to insert courses, the leftmost column

while currentDate < endDate:
  # Here read according to the correct html tags, but time to sleep!! may be tomorrow...
  for y, line in enumerate(inLine.split('<tr')[2:]): # if .split is called every time the loop exectures ??
    for x, column in enumerate(line.split('<td')[1:]):
      if column.find('<p>') == -1:
        # Empty slot, continue
        continue
      data = column.split('<p>')[1].split('</p>')[0]
      if(x == 0):
        # Then it is hours column
        hours = data
        continue

      outFile.write('BEGIN:VEVENT\n')

      currentTime = currentDate + datetime.timedelta(days=x) + datetime.timedelta(seconds=60*int(hours.split('.')[1]) + 3600*int(hours.split('.')[0]))
      outFile.write('DTSTART:' + currentTime.strftime('%Y%m%d') + 'T' + currentTime.strftime('%H%M') + '00Z\n')
      currentTime = currentTime + datetime.timedelta(seconds=60*lengthOfALectureHour)
      outFile.write('DTEND:' + currentTime.strftime('%Y%m%d') + 'T' + currentTime.strftime('%H%M') + '00Z\n')

      outFile.write('SEQUENCE:0\n')
      outFile.write('STATUS:CONFIRMED\n')
      outFile.write('SUMMARY:' + data + '\n')
      outFile.write('END:VEVENT\n')

  currentDate = currentDate + datetime.timedelta(days=7)

outFile.write('END:VCALENDAR\n')
inFile.close()
outFile.close()
