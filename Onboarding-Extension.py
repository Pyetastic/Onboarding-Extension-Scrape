import gspread #Imports Google Sheets API
from oauth2client.service_account import  ServiceAccountCredentials
import pprint # Get PPrint in
import datetime
scope = ['https://spreadsheets.google.com/feeds',
         'https://www.googleapis.com/auth/drive'] # the locations in which we can locate the our loin ariea
creds = ServiceAccountCredentials.from_json_keyfile_name('LICred.json',scope) #Secure deets location
client = gspread.authorize(creds) #logs into Google sheet
 
pp = pprint.PrettyPrinter() # Calls pretty print fuction and saves it as pp for ease
sheet = client.open('Paser sheet').sheet1 #Save the sheet form the leaver doc
 
#Data = sheet.get_all_records() #Gets all the data
#pp.pprint(Data)#Prints all data nicely using Pretty print
 
## How to update cells sheet.update_cell(6,1,Stuff thats happening)
 
CRindex = sheet.cell(1,1,1).value
pp.pprint(CRindex)
 
text = """"Employee Number:    00912934 " 
       "Employee Name:      Georgina Laud" 
       "Leaving Date:       18.03.2018 """
LE = open('leaver.txt','r') #reads temp text file
content = LE.read() #turns it to text
text = content
 
Lines = text.splitlines() #splits txt file into readable lines
Empno = 0
Empnam = 0
Ldate = 0
Depar = 0
Empgro = 0
EDate = 0
Eset = 0
Econt = 0
def FENo() : #Breaks up employee number
    if line.find("Contractor number:") !=-1:
        global Empno
        enumber = line.split(":")
        enumber = enumber[1].split('"')
        enumber = enumber[0].split("'")
        Empno = enumber[0].lstrip(' ')
       # Empno = Empno.split('")
def FEG() : #Breaks up employee group
    if line.find("Employee Group:") !=-1:
        global Empgro
        epgo = line.split(":")
        epgo = epgo[1].split('"')
        Empgro = (epgo[0].lstrip(' '))
def FDe() : #Breaks up department
    if line.find("Department:") !=-1:
        global Depar
        dep = line.split(":")
        dep = dep[1].split('"')
        dep = dep[0].split("'")
        Depar = (dep[0].lstrip(' '))
def FENa() : #breaks up employee name
    if line.find("Contractor extension -") != -1:
        global Empnam
        ename = line.split("-")
        ename = ename[1].split("'")
        Empnam = ename[0].lstrip(' ')
def FRD() : #gets extend to date
    if line.find("Required: ") != -1:
        global Ldate
        lno = line.split(":")
        lno = lno[1].split("to")
        lno = lno[1].split("'")
        Ldate = lno[0].lstrip(' ')
def FED() : #Breaks up sent date
    if line.find("Date:") !=-1:
        global EDate
        eda = line.split(":")
        eda = eda[1].split('"')
        EDate = (eda[0].lstrip(' '))
def FSent() : #Breaks up sent date
    if content.find("From: Mailbox, SD <Service.Desk@news.co.uk>") !=-1:
        global Eset
        eda = content.split("Date:")
        eda = eda[1].split('Subject:')
        Eset = (eda[0].lstrip(' '))
def FCont() : #Breaks up sent date
    if content.find("---------- Forwarded message ----------") !=-1:
        global Econt
        eda = content.split("---------- Forwarded message ----------")
        Econt = (eda[1].lstrip(' '))
for line in Lines: #Cycles email looking for text to scrape
     FENo()
     FENa()
     FRD()
     FDe()
     FEG()
     FED()
 
FSent()
FCont()
 
CRindex = CRindex + 1 # Creates Free Row number
Logdate = datetime.datetime.now().strftime("%I:%M%p on %B %d, %Y")
#Formular = '=TRIM(CLEAN(SUBSTITUTE("' + Econt +'",CHAR(160)," "))'
# =C122& " - " &D122& " - " &G122& " - " &TEXT(H122,"dd/mm/yy")
#Writes Data to Google sheet
sheet.update_cell(CRindex,1,"PyLog")
sheet.update_cell(CRindex,3,"PyLog")
sheet.update_cell(CRindex,4,Empnam)
sheet.update_cell(CRindex,8,Empno)
sheet.update_cell(CRindex,9,Ldate)
sheet.update_cell(CRindex,6,Depar)
#sheet.update_cell(CRindex,7,Empgro)
sheet.update_cell(CRindex,7,"Contractor")
sheet.update_cell(CRindex,10,Logdate)
sheet.update_cell(CRindex,11,"Yes")
#sheet.update_cell(CRindex,12,Formular)
#print(Empgro)
