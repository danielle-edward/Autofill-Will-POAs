## yeehaw
## program for Amanda L Groves to automate stuff
from docx import Document
#import os

ffiles = []#first client
sfiles = []#second client files to be printed
clientNames = []
firstNames =[]
secondNames = []
date = []
suffix = ''
twoClients = False
witness = ''
fothers = ['Will.docx', 'Personal Care.docx', 'Property.docx', 'Backsheet.docx','Direction.docx']
sothers = ['Will.docx', 'Personal Care.docx', 'Property.docx', 'Backsheet.docx','Direction.docx']

def main():
    print("What would you like to do?\n")
    print("1. Fill in documents \n2. Quit")
    task = input("Your selection: ")
    if (task == '1'):
        doAll()
    else:
        quit

def doAll():
    getNames()
    
    print('Enter the date of signing in format of " 6 November 2021 "\n')
    global date
    date = (input(": ")).split(' ')
    global suffix
    suffix = checkNumber()
    pullFiles(0)
    if twoClients:
        pullFiles(1)
    quit


def pullFiles(x):
    files = ffiles
    otherNames = firstNames
    if x==1:
        files = sfiles
        otherNames = secondNames
    for i in files:
        docBackSheet = Document('./documents/'+str(i))
        for p in docBackSheet.paragraphs:
            for run in p.runs:
                if ('GEORGE WASHINGTON JETSON' in str(run.text)) or ('George Washington Jetson' in str(run.text)):
                    run.text = clientNames[x]
                if 'JANE JETSON' in str(run.text):
                    run.text = otherNames[0]
                if 'COSMO SPACELY' in str(run.text):
                    run.text = otherNames[1]
                if '*DATED WORD*' in str(run.text):
                    run.text = date[0]+suffix + ' day of ' + date[1] + ', ' + date[2]
                if '*DATED NUM*' in str(run.text):
                    run.text = date[1] + ' ' + date[0] + ', ' + date[2]
                if '*WITNESS*' in str(run.text):
                    run.text = witness
                if '*First Client*' in str(run.text):
                    run.text = clientNames[0]
                if '*Second Client*' in str(run.text):
                    run.text = clientNames[1]
        saveName = clientNames[x]+'-'+i
        docBackSheet.save(saveName)

def checkNumber():
    if (date[0] == '1'):
        return 'st'
    if (date[0] == '2'):
        return 'nd'
    if (date[0] == '3'):
        return 'rd'
    else:
        return 'th'

def getNames():
    print(">>Please enter all names in capitals.")
    twoWills = input("Is there more than 1 client (spouses)? (Y/N): ")
    global twoClients
    if (twoWills == 'Y') or (twoWills == 'y'):
        twoClients = True
    cname = input("Name of First Client: ")
    global clientNames
    clientNames.append(cname)
    print("Please enter the following information in this format pertaining to " + cname +". \n")
    print("Trustee 1,Trustee 2\n")
    print("If there is only one trustee, please fill the other spot with N/A.")
    global firstNames
    firstNames = (input(": ")).split(',')
    chooseFiles(False)
    if twoClients:
        cname = input("Name of Second Client: ")
        clientNames.append(cname)
        print("Please enter the following information in this format pertaining to " + cname +". \n")
        print("Trustee 1,Trustee 2\n")
        global secondNames
        global sfiles
        secondNames = (input(": ")).split(',')
        chooseFiles(True)
        
    print("Please enter the name of the witness.\n")
    global witness
    witness = input(": ")
    

        
def chooseFiles(x):
    print("Would you like to create all these files?")
    docs = fothers
    ack = 'Acknowledgement, single.docx'
    global ffiles
    global sfiles
    if (x):
        docs = sothers
    if (twoClients):
        ack = 'Acknowledgement, couple.docx'
    for i in docs:
        print(i)
    print(ack)
    answ = input("Y/N?: ")
    if (answ == 'Y') or (answ == 'y'):
        docs.append(ack)
        if (x):
            sfiles = docs
        else:
            ffiles = docs
    else:
        tempFiles = []
        print("Type the files you would like to create one at a time. Please use the listed document names above.")
        end = False
        while (not end):
            answ = input("Add a document. (Type STOP to end): ")
            if (answ == 'STOP') or (answ == 'stop'):
                end = True
            else:
                tempFiles.append(str(answ))
        print("These are the files you chose.")
        for i in tempFiles:
            print(i)
        answ = input("Are these right? (Y/N): ")
        if not((answ == 'Y') or (answ =='y')):
            print("Starting over . . .")
            chooseFiles()
        else:
            if (x):
                sfiles = tempFiles
            else:
                ffiles = tempFiles
        
                
main()
