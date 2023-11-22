# pip install pywin32
# pip install pytz

import win32com.client
from datetime import datetime, timedelta
from BCalendar import CalendarUtils
import pytz
import csv

class OutlookCalendarUtils:
    
    # -----------------------------------------------------------------------
    # PRIVATE METHODS
    # -----------------------------------------------------------------------
    
    def __init__(self, numOfDays):
        self.__calendarItems = []
        if numOfDays is not None: self.loadCalendarItems(numOfDays)
    
    # Retrieve the calendar items from outlook
    def __getCalendarItems(self, startDt, endDt):
        # 9 corresponds to the Calendar folder
        calendar = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(9)  
    
        filterStr = f"[Start] >= '{startDt.strftime('%m/%d/%Y')}' AND [END] <= '{endDt.strftime('%m/%d/%Y')}'"
    
        items = calendar.Items.Restrict(filterStr)
        items.IncludeRecurrences = True
        items.Sort("[Start]")
    
        filteredItems = []
    
        for i in items:
            iStart = i.Start.astimezone(pytz.utc)
            # Exlude meetings outside the start/end date, and exclude 'free' or 'OOO' meetings.
            if startDt <= iStart < endDt and i.BusyStatus not in [1, 3]:
                filteredItems.append(i)
            elif iStart > endDt:
                # Items are sorted, none here after will be inside the date range.
                break
    
        return filteredItems
    
    def __getRowDefinition(self, event):
        return {
            'Subject': event.Subject,
            'Day': event.Start.strftime('%Y-%m-%d'),
            'Start': event.Start.strftime('%H:%M'),
            'End': event.End.strftime('%H:%M'),
            'Duration (min)': event.Duration,
            'Duration (hours)': round(event.Duration/60, 2)
        }
    
    def __writeCsvFile(self, events, csvFilename):
        with open(csvFilename, 'w', newline='', encoding='utf-8') as csvfile:
            fieldnames = ['Subject', 'Day', 'Start', 'End', 'Duration (min)', 'Duration (hours)']
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
            
            writer.writeheader()
            for event in events: writer.writerow(self.__getRowDefinition(event))

    # -----------------------------------------------------------------------
    # PUBLIC METHODS
    # -----------------------------------------------------------------------
    
    def getCalendarItems(self):
        items = []
        for i in self.__calendarItems: items.append(self.__getRowDefinition(i))
        return items
    
    def printCalendarItems(self):
        for i in self.__calendarItems: print(f"Subject: {i.Subject} - Start: {i.Start} - End: {i.End}")
    
    def getDurationOfCalendarItems(self):
        totalDuration = 0
        for i in self.__calendarItems: totalDuration += i.Duration
        return totalDuration
    
    def printDurationOfCalendarItems(self):
        print(f"Total time of meetings: {round(self.getDurationOfCalendarItems()/60, 2)}h")
    
    def exportToCsv(self, name = str(datetime.now(pytz.timezone('Africa/Harare')).strftime('%Y-%m-%d %H-%M'))):
        csvFilename = f"{name}.csv"
        self.__writeCsvFile(self.__calendarItems, csvFilename)
        print(f"Calendar events exported to {csvFilename}")
    
    def loadCalendarItems(self, numOfDays):
        startDt = datetime.now(pytz.timezone('Africa/Harare'))
        endDt = startDt + timedelta(days=numOfDays)
        self.__calendarItems = self.__getCalendarItems(startDt, endDt)