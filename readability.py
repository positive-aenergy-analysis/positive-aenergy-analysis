#統計字數寫入excel
import sys
from imp import reload
reload(sys)
import json
import jieba
import jieba.analyse
#write exl
import xlrd
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import ExcelWriter
# from openpyxl.cell import get

nameList = []
scoreList = []

def readMental_disorder():
    
    data = xlrd.open_workbook('vocabulary/mental_disorder.xlsx')
    table = data.sheets()[0]
    disorderWords = table.col_values(0)
    disorderScores = table.col_values(2)
    
    return disorderWords, disorderScores

def check(filepath):
    data = xlrd.open_workbook(filepath)
    table = data.sheets()[0]
    datawords = table.col_values(0)
    datatimes = table.col_values(1)
    disorderWords, disorderScores = readMental_disorder()

    remainderWords = []
    remainderWords = list(filter(lambda a: a in disorderWords and a != '\n' and a != '病症', datawords))
    
    valueList = []
    timesList = []
    score = 0

    for key in remainderWords:
        for i in range(len(disorderWords)):
            if key == disorderWords[i]:
                valueList.append(disorderScores[i])
        
        for i in range(len(datawords)):
            if key == datawords[i]:
                timesList.append(datatimes[i])

    for i in range(len(remainderWords)):
        score += valueList[i] * timesList[i]
    
    if len(remainderWords) <= 3:
        score *= 3
    elif len(remainderWords) > 3 and len(remainderWords) <= 6:
        score *= 2

    nameList.append(filepath)
    scoreList.append(score / sum(timesList))

    return remainderWords, valueList

def writeExcel(valueList, keyList, fileName):
    
    wb = Workbook()
    ws = wb.active
    
    for i in range(len(keyList)):
        text1 = 'A' + str(i+1)
        text2 = 'B' + str(i+1)

        ws[text1] = keyList[i]
        ws[text2] = valueList[i]
        
    wb.save(fileName)

if __name__ == "__main__":

    json_data = open('data_set/correspondence_table.json', encoding='utf8').read()
    data = json.loads(json_data)

    for i in data:
        for k,v in i.items():
            if k == 'file_name':
                filepath = 'data_set_wordCount/wordCount_' + str(v) + '.xlsx'
                keyList, valueList = check(filepath)

                writeExcel(valueList,keyList,'score_' + str(v) + '.xlsx')

    writeExcel(scoreList,nameList,'totalScore.xlsx')


