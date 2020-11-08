import os
import datetime
from xml.etree import ElementTree

def checkFiles(catalogue, wholesalerName):
    os.chdir(os.path.join(catalogue, wholesalerName, 'Connection_register'))
    for filename in os.listdir():
        if (filename.startswith('NAPS_') and filename.endswith('.txt') and "00_G" not in filename):
            if (DateToHelpDeleter - int(filename[-14:-6]) > finalDailyRet):
                deleteFile(filename)
        elif (filename.startswith('NAPS_') and filename.endswith('.txt') and "00_G" in filename):
            if (DateToHelpDeleter - int(filename[-14:-6]) > finalMonthlyRet):
                deleteFile(filename)

def deleteFile(fileToDelete):
    logFile.write('Delete: ' + fileToDelete + '\n')
    os.unlink(fileToDelete)


############ main Program ############

ftpRootCatalogue = 'C:\\Users\\przemyslaw.lelewski\\Documents\\CCC' #ZMIEN SCIEZKE DLA PLIKOW
logArchive_file = 'archive_log.txt'
file_name = "Data.xml"
wslNames = []

# create/open or open the log file
if (os.path.isfile(logArchive_file)) == False:
    logFile = open(logArchive_file, 'w')
else:
    logFile = open(logArchive_file, 'a')

# get system_datetime and write to log file
DateToHelpDeleter = int(datetime.datetime.now().strftime("%Y%m%d"))
now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
logFile.write('------ Start ' + str(now) + '\n')

# parse xml config file
dom = ElementTree.parse(file_name)
wslTree = dom.findall("Wholesaler")
retTree = dom.findall("Retension")

# find all Wholesalers in the config file
for wTree in wslTree:
    wsls = wTree.findall("wsl")
    for wsl in wsls:
        finalWslName = wsl.text
        wslNames.append(finalWslName)

# find how many rete in the config file
for rTree in retTree:
    rT = rTree.findall("daily")
    for ff in rT:
        finalDailyRet = int(ff.text)
    tR = rTree.findall("monthly")
    for hh in tR:
        finalMonthlyRet = int(hh.text)

# run check/delete files according to the parameters
for wslName in wslNames:
    checkFiles(ftpRootCatalogue, wslName)

logFile.write('------ End' + '\n\n')
logFile.close()