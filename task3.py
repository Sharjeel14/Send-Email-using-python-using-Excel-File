# Python code to send email to a list of
# emails from a spreadsheet

# import the required libraries
import smtplib
from tokenize import Name
from sqlalchemy import create_engine
import pandas as pd

# change these as per use
your_email = "youremail@gmail.com"
your_key = "yourkey"

# establishing connection with gmail
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_key)

# reading the spreadsheet
email_list = pd.read_excel('C:\\Users\\Malik\\OneDrive\\Desktop\\sample.xlsx')
print(email_list)

# getting the names and the emails
names = email_list['Name']
emails = email_list['Email']
salaries = email_list['Salary']

# iterate through the records
# for i in range(len(emails)):

# 	# for every record get the name and the email addresses
# 	name = names[i]
# 	email = emails[i]
# 	salary = salaries[i]

# 	# the message to be emailed
# 	message = "Hello " + name

# 	# sending the email
# 	server.sendmail(your_email, [email], message)
# 	print("Mail Sent")

file = "C:\\Users\\Malik\\OneDrive\\Desktop\\sample.xlsx"
# output = "output.xlsx"
engine = create_engine('sqlite://', echo=False)
df = pd.read_excel(file, sheet_name="Sheet1")
df.to_sql('sample', engine, if_exists = 'replace', index=False)
result = engine.execute(" Select Email, Name from sample where Sale = (Select Max(Sale) from sample)")
print(result)
email_25000 = pd.DataFrame(result)                          
# final.to_excel(output, index=False)
print(email_25000)
for i in range(len(email_25000)):

	# for every record get the name and the email addresses
	name = names[i]
	email = emails[i]
	salary = salaries[i]
	# the message to be emailed
	message = "Hello " + name

	# sending the email
	server.sendmail(your_email, [email], message)
	print("Mail Sent")
# close the smtp server
server.close()
