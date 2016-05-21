import os
import xlsxwriter
import re

'''path to this file'''
path=os.path.dirname(os.path.realpath(__file__))
#/Users/Jared/Dropbox/Auburn/Research/Second_Research/Test_Files
excelFilePathName='/propagatorData.xlsx'
logFilesFolder='/Propa_files'

colFileInformation=0
colMolecule=1
colCharge=2
colMultiplicity=3
colBasis=4
colNineFive=5
colEigenValue=6
colOrbital=7
colPS=8
colCFF=9

def writeDataToExcel(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive,\
            orbital, eigenValue, ps, cff):
    worksheet.write(row, colFileInformation, fileInformation)
    worksheet.write(row, colMolecule, molecule)
    worksheet.write(row, colCharge, charge)
    worksheet.write(row, colMultiplicity, multiplicity)
    worksheet.write(row, colBasis, basis)
    worksheet.write(row, colNineFive, nineFive)   
    worksheet.write(row, colOrbital, orbital)
    worksheet.write(row, colEigenValue, eigenValue)
    worksheet.write(row, colPS, ps)
    worksheet.write(row, colCFF, cff)
            
    

            


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
    #prepare excel file first
    workbook = xlsxwriter.Workbook(path + excelFilePathName)
    worksheet = workbook.add_worksheet('Data')
    bold = workbook.add_format({'bold': True})
    row=1
    
    worksheet.write(0, colFileInformation, 'File', bold)
    worksheet.write(0, colMolecule, 'Molecule', bold)
    worksheet.write(0, colCharge, 'Charge', bold)
    worksheet.write(0, colMultiplicity, 'Multiplicity', bold)
    worksheet.write(0, colBasis, 'Basis', bold)
    worksheet.write(0, colNineFive, '9/5', bold)   
    worksheet.write(0, colOrbital, 'Orbital', bold)
    worksheet.write(0, colEigenValue, 'Eigen Value', bold)
    worksheet.write(0, colPS, 'polestrength', bold)
    worksheet.write(0, colCFF, 'CFF', bold)
    
    #extraction code starts here
    logFiles=[]

    for path, subdirs, files in os.walk(path+logFilesFolder):
        for name in files:
            if os.path.join(path, name)[len(os.path.join(path, name))-4:len(os.path.join(path, name))]=='.log':
                logFiles.append(os.path.join(path, name))
    #print(logFiles) 

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
                
        '''
        print('molecule ' + molecule)
        print('charge ' + charge)
        print('multiplicity ' + multiplicity)
        print('basis ' + basis)
        print('9/5 ' + nineFive)
        '''
        
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
            '''
            print('eigenvalue ' + eigenValue)
            print('orbital ' + orbital)
            print('polestrength ' + str(ps))
            print('CFF ' + str(cff))
            '''
            writeDataToExcel(worksheet, row, fileInformation, molecule, charge, multiplicity, basis, nineFive,\
            orbital, eigenValue, ps, cff)
            
            row+=1
            