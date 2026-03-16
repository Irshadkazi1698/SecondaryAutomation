import pandas as pd
import chardet
import csv
import numpy as np
import os
import shutil


def GenerateColumnName(numberOfColumns):
    columns = []
    for i in range(0,numberOfColumns):
        columns.append('col'+str(i))
    
    return columns

def getDelimiterType(path):
    with open(path, 'r', newline='') as csvfile:
        sample = csvfile.read(100000)  # Read a sample of the file
        sniffer = csv.Sniffer()
        dialect = sniffer.sniff(sample)
        print(dialect.delimiter)

        if dialect.delimiter == ",":
            delimiter = dialect.delimiter
        elif dialect.delimiter != ',':
            delimiter = '\t'
        
        print(f'Delimiter : {delimiter}')
    
    return delimiter


def getUnicodeType(path):
    with open(path, 'rb') as rawdata:
        result = chardet.detect(rawdata.read(10000))
    
    encodingType = result['encoding']
    print(f"Encoding Type : {encodingType}")
    
    return encodingType

def LoadingFileAndCleaning(path,col):
    #Reading the csv File
    df = pd.read_csv(path,sep=getDelimiterType(path),encoding=getUnicodeType(path),names=GenerateColumnName(col))

    for i in df.columns:
        df[i] = df[i].astype('str')


    percentIndex = df.index[df['col1'].str.contains("%")] - 1 
    percentData = df['col1'][df['col1'].str.contains("%")]

    for i in zip(percentData,percentIndex):
        df.iat[i[1],2] = i[0]
    
    df = df[~df['col1'].str.contains("%",'nan')]
    df = df[~df['col0'].str.contains('nan')]

    #Droping all the null values
    df = df.dropna(subset = GenerateColumnName(col), how='all')
    
    # Create a null row with same columns
    null_rows = pd.DataFrame(np.nan, index=range(5), columns=GenerateColumnName(col))
    df2 = pd.concat([null_rows,df],ignore_index=False).reset_index()[GenerateColumnName(col)]

    df.reset_index(inplace= True)

    for i in df2.columns:
        df2[i] = df2[i].astype('str')



    df2['col5'] = df2['col0'].apply(lambda x : x.split('VARIABLE_LABEL')[0])


    #Fetching Indexes
    pageIndexes = list(df2.index[df2['col0'].str.contains('#page')])
    QuestionEndIndexes = list(df2.index[df2['col5'].str.contains(';')])
    QuestionLabelIndexes = list(df2.index[df2['col0'].str.contains('#page')] + 2)


    TableLenght = []
    for index in zip(pageIndexes,QuestionEndIndexes):
        TableLenght.append(abs(index[1] - index[0]))

    for i in zip(QuestionLabelIndexes,TableLenght,pageIndexes,QuestionEndIndexes):
        with open("Summary Details.txt",'a+') as f:
            f.write(f"{df2.iloc[i[0],0].split('VARIABLE_LABEL')[0]} | {i[1]} | {i[2]} | {i[3]} \n")

    questionLabelAndName = []
    for i in range(0,len(QuestionLabelIndexes)):
        for j in range(0,len(TableLenght)):
            if i == j:
                for k in range(0,TableLenght[j] + 1):

                    questionLabelAndName.append(df2.iloc[QuestionLabelIndexes[i],0])

    questionLabelAndName.insert(0,'nan')
    questionLabelAndName.insert(1,'nan')
    questionLabelAndName.insert(2,'nan')
    questionLabelAndName.insert(3,'nan')
    questionLabelAndName.insert(3,'nan')

    df2['QuestionList'] = questionLabelAndName
    rearrangeCol = GenerateColumnName(col)
    rearrangeCol.insert(0,'QuestionList')
    df2 = df2[rearrangeCol]
    df2.iat[4,0] = 'StartPoint'
    df2.replace('nan',np.nan,inplace=True)

    fileName = os.path.basename(path)
    outputfileName = fileName.replace('.csv','.xlsx')
    outputPath = path.rstrip(fileName)
    FinalPath = os.path.join(outputPath,outputfileName)

    print(FinalPath)

    df2.to_excel(FinalPath,index=False,header=None,sheet_name='Tables')


def CopyingFiletoInputFolder():
    parentPath = os.getcwd()
    print(f'Parent Directory : {parentPath}\n')

    sourcePath = os.path.join(parentPath,'CountsInputFiles')
    print(f'Source Path Directory : {sourcePath}\n')

    TargetPath = os.path.join(parentPath,'Input')
    print(f'Target Directory : {TargetPath}\n')

    fileName = os.listdir(sourcePath)

    for file in fileName:
        sourcefilePath = os.path.join(sourcePath,file)
        destinationfilePath = os.path.join(TargetPath,file)
        try:
            shutil.copyfile(sourcefilePath, destinationfilePath)
            print(f"File '{sourcefilePath}' copied to '{destinationfilePath}' successfully.\n")
        except FileNotFoundError:
            print(f"Error: Source file '{sourcefilePath}' not found.")
        except Exception as e:
            print(f"An error occurred: {e}")

