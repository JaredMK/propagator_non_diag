import os
import re
import openpyxl
workbook = openpyxl.Workbook()  #creates openpyxl workbook

#logFilesFolder is name of folder containing log files 
logFilesFolder='/Propa_files'

'''path to this file'''
path=os.path.dirname(os.path.realpath(__file__))
pathorigin=path     #used to save workbook in this location
excelFilePathName='/non_diag_propagatorData.xlsx'

#columns for each variable in workbook
colFileInformation='A'
colMolecule='B'
colCharge='C'
colMultiplicity='D'
colBasis='E'
colNineFive='F'
colEigenValue='G'
colOrbital='H'
colPS='I'
colCFF='J'

def writeDataToExcel(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive,\
            orbital, eigenValue, ps, cff):
    '''writesDataToExcel takes is called by dataExtract. It takes in the variables found in 
    data extraction and writes it into the openpyxl workbook'''
    
    worksheet[colFileInformation+str(row)]=fileInformation
    worksheet[colMolecule+str(row)]=molecule
    worksheet[colCharge+str(row)]=charge
    worksheet[colMultiplicity+str(row)]=multiplicity
    worksheet[colBasis+str(row)]=basis
    worksheet[colNineFive+str(row)]=nineFive
    worksheet[colOrbital+str(row)]=orbital
    worksheet[colEigenValue+str(row)]=eigenValue
    worksheet[colPS+str(row)]=ps
    worksheet[colCFF+str(row)]=cff
        
def numberOfBasisSets(logarray):
    '''returns a list of the split log arrays by basis set. length is number of basis sets'''
    commandLocation=[]
    logsToReturn=[]
    x=0
    while x < len(logarray):
        if logarray[x] =='Final':
            commandLocation.append(x)
        x+=1
    commandLocation.append(len(logarray))
    x=0
    logsToReturn.append(logarray[:commandLocation[0]])
    #the first log in the array is from the start of the file to the first keyword
    while x< len(commandLocation)-1:
        b=logarray[commandLocation[x]:commandLocation[x+1]]
        logsToReturn.append(b)
        x+=1
    return logsToReturn


def dataExtract(path):
    '''Main function in script. Calls other functions. Takes in path of this file and extracts
    data from the log files folder. Then calls functions above to add data to the openpyxl file'''

    #prepare openpyxl first
    worksheet=workbook.active
    worksheet.title="Data"
    #creates worksheet Data
    
    #add headings to each column
    worksheet[colFileInformation+'1']='File'
    worksheet[colMolecule+'1']='Molecule'
    worksheet[colCharge+'1']='Charge'
    worksheet[colMultiplicity+'1']='Multiplicity'
    worksheet[colBasis+'1']='Basis'
    worksheet[colNineFive+'1']='9/5'
    worksheet[colOrbital+'1']='Orbital'
    worksheet[colEigenValue+'1']='Eigen Value'
    worksheet[colPS+'1']='polestrength'
    worksheet[colCFF+'1']='CFF'
    
    row=2
    #extraction from log files starts here
    logFiles=[]

    for path, subdirs, files in os.walk(path+logFilesFolder):
        for name in files:
            if os.path.join(path, name)[len(os.path.join(path, name))-4:len(os.path.join(path, name))]=='.log':
                logFiles.append(os.path.join(path, name))

    for currentFile in logFiles:
        log = open(currentFile, 'r').read()
        splitLog = re.split(r'[\\\s]\s*', log)  #splits string with \ (\\), empty space (\s) and = and ,
        fileInformation=currentFile
        firstSplitLog=numberOfBasisSets(splitLog)[0]
        nineFiveFound=False
        x=0
        while x<len(firstSplitLog):
            if firstSplitLog[x]=='Stoichiometry':
                molecule=firstSplitLog[x+1]
            if firstSplitLog[x]=='Charge' and firstSplitLog[x-1]=='Z-matrix:':
                charge=firstSplitLog[x+2]
            if firstSplitLog[x]=='Multiplicity':
                multiplicity=firstSplitLog[x+2]
            if firstSplitLog[x]=='Standard' and firstSplitLog[x+1]=='basis:':
                basis=firstSplitLog[x+2] +' '+firstSplitLog[x+3]+firstSplitLog[x+4]
            if firstSplitLog[x][0:4]=='9/5=' and nineFiveFound==False:
                nineFive=firstSplitLog[x][4]
                s=0
                f=0
                n=0
                while n<len(firstSplitLog[x]):
                    if firstSplitLog[x][n]=='=' and s==0:
                        s=n+1
                    if firstSplitLog[x][n]==',' and f==0:
                        f=n                       
                    n+=1
                nineFive=firstSplitLog[x][s:f]                    
                nineFiveFound=True
            x+=1
        
        for splitlog in numberOfBasisSets(splitLog)[1:]:
            #NUMBEROFSPLITS will return where in log file it needs to be split for basis sets
            #textFile(log)   #text file will return each log split by basis set because some aren't
            #print(splitlog)

            x=0
            while x<len(splitlog):   
                if splitlog[x]=='(eV)':
                    eigenValue=splitlog[x+1]
                    orbital=splitlog[x+3][0]
                if splitlog[x]=='polestrength':
                    a=float(splitlog[x+1])
                    b=float(splitlog[x+2])
                    if a>b:
                        ps=a
                        cff=b
                    else:
                        ps=b
                        cff=a
                x+=1
            #send variables from data extraction to writeDataToExcel
            writeDataToExcel(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive,\
            orbital, eigenValue, ps, cff)
            
            row+=1
             
    workbook.save(pathorigin + excelFilePathName)     #saves file