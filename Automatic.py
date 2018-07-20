import time
import karix
from karix.rest import ApiException
from karix.configuration import Configuration
from karix.api_client import ApiClient
import tkinter as tk
from openpyxl import load_workbook
from tkinter import filedialog

def customMessage(): #Function that takes care of mass messaging for information passing
	text = input('Enter custom message: ')
	recievers = list()
	for row in ws:
		if str(row[0].value).lower().strip() == 'name':continue
		recievers.append('+91'+str(row[1].value))
	message = karix.CreateMessage(source="", destination=recievers, text=text)
	try:
		api_response = api_instance.send_message(message=message)
		print(api_response)
		print('Success')
	except ApiException as e:
		print("Exception when calling MessageApi->send_message: %s\n" % e)
		print('Failure')	
	return

def sendScore(): #Function that sends individual score
	date = input('Enter the date of the exam: ')
	subject = input ('Enter the subject/topic of the exam: ')
	maxMarks = input('Enter the maximum marks attenable: ')
	highest = input('Enter highest marks: ')

	for row in ws:
		if str(row[0].value).lower().strip() != 'name':
			text = 'Dear Gaurdian\n Your ward ' + str(row[0].value) +' has scored ' +str(row[2].value) +'/' + maxMarks + ' in the test of '+ subject  +' held on '+ date +' Highest marks obtained in class: ' + highest
			message = karix.CreateMessage(source="", destination=['+91'+str(row[1].value)], text=text)
			#Test whether the message was actually delivered
			try:
				api_response = api_instance.send_message(message=message)
				print(api_response)
			except ApiException as e:
				print("Exception when calling MessageApi->send_message: %s\n" % e)
	return

# Open the file and load it as active worksheet
root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename()
wb = load_workbook (file_path)
ws = wb.active
# Configure HTTP basic authorization: basicAuth
config = Configuration()
config.username =  #account auth ID
config.password = #account auth token
# Create an instance of the API class
api_instance = karix.MessageApi(api_client=ApiClient(configuration=config))
# Determining form of Communication
choice = input('1. Sending Score\n2.Custom message\nEnter your choice: ')
if choice is '2':
	customMessage()
else:
	sendScore()