# pip install pywin32
# pip install pytz

import win32com.client
from datetime import datetime, timedelta
import pytz
import csv

class CalendarUtils:
	
	# -----------------------------------------------------------------------
	# PRIVATE METHODS
	# -----------------------------------------------------------------------
	
	def __init__(self, numOfDays):
		self.__calendar_items = []
		if numOfDays is not None: self.loadCalendarItems(numOfDays)
	
	# Retrieve the calendar items from outlook
	def __getCalendarItems(self, start_datetime, end_datetime):
		# 9 corresponds to the Calendar folder
		calendar = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI").GetDefaultFolder(9)  
	
		filter_str = f"[Start] >= '{start_datetime.strftime('%m/%d/%Y')}' AND [END] <= '{end_datetime.strftime('%m/%d/%Y')}'"
	
		items = calendar.Items.Restrict(filter_str)
		items.IncludeRecurrences = True
		items.Sort("[Start]")
	
		filtered_items = []
	
		for item in items:
			item_start = item.Start.astimezone(pytz.utc)
			# Exlude meetings outside the start/end date, and exclude 'free' or 'OOO' meetings.
			if start_datetime <= item_start < end_datetime and item.BusyStatus != 3 and item.BusyStatus != 0:
				filtered_items.append(item)
			elif item_start > end_datetime:
				# Items are sorted, none here after will be inside the date range.
				break
	
		return filtered_items
		
	def __writeCsvFile(self, events, csv_filename):
		with open(csv_filename, 'w', newline='', encoding='utf-8') as csvfile:
			fieldnames = ['Subject', 'Day', 'Start', 'End', 'Duration (min)', 'Duration (hours)']
			writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
			
			writer.writeheader()
			for event in events:
				writer.writerow({
					'Subject': event.Subject,
					'Day': event.Start.strftime('%Y-%m-%d'),
					'Start': event.Start.strftime('%H:%M'),
					'End': event.End.strftime('%H:%M'),
					'Duration (min)': event.Duration,
					'Duration (hours)': round(event.Duration/60, 2)
				})

	# -----------------------------------------------------------------------
	# PUBLIC METHODS
	# -----------------------------------------------------------------------
	
	# Return the calendar items as a dictionary
	def getCalendarItems(self):
		items = []
		for item in self.__calendar_items:
			items.append({
				'Subject': item.Subject,
				'Start': item.Start.strftime('%Y-%m-%d %H:%M'),
				'End': item.End.strftime('%Y-%m-%d %H:%M'),
				'Duration (min)': item.Duration,
				'Duration (hours)': round(item.Duration/60, 2)
			})
	
		return items
	
	def printCalendarItems(self):
		for item in self.__calendar_items:
			print(f"Subject: {item.Subject} - Start: {item.Start} - End: {item.End}")
			
	def getDurationOfCalendarItems(self):
		total_duration = 0
		for event in self.__calendar_items: total_duration += event.Duration
		return total_duration
		
	def printDurationOfCalendarItems(self):
		print(f"\nTotal time of meetings: {round(self.getDurationOfCalendarItems()/60, 2)}h")
		
	def exportToCsv(self, name = str(datetime.now(pytz.timezone('Africa/Harare')).strftime('%Y-%m-%d %H-%M'))):
		csv_filename = f"{name}.csv"
		self.__writeCsvFile(self.__calendar_items, csv_filename)
		print(f"\nCalendar events exported to {csv_filename}")
	
	def loadCalendarItems(self, numOfDays):
		start_datetime = datetime.now(pytz.timezone('Africa/Harare'))
		end_datetime = start_datetime + timedelta(days=numOfDays)
	
		self.__calendar_items = self.__getCalendarItems(start_datetime, end_datetime)
		
		return self.__calendar_items