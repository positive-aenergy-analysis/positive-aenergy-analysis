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
keywordList = []

def readMental_disorder():
    
    data = xlrd.open_workbook('vocabulary/mental_disorder.xlsx')
    table = data.sheets()[0]
    disorderWords = table.col_values(0)
    disorderScores = table.col_values(2)
    
    return disorderWords, disorderScores

def check(filepath):
    # data 為分詞完的文章, datawords 是column(A), datatimes 是column(B)
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

    if sum(timesList) == 0:
        scoreList.append(0)
    else:    
        scoreList.append(score / sum(timesList))

    keywordList.append(sum(timesList))

    return remainderWords, valueList, timesList

def writeExcel(fileName, keyList, valueList, timesList, col_title1, col_title2, col_title3):
    
    wb = Workbook()
    ws = wb.active

    ws['A1'] = col_title1
    ws['B1'] = col_title2
    ws['C1'] = col_title3
    
    for i in range(len(keyList)):
        col1 = 'A' + str(i+2)
        col2 = 'B' + str(i+2)
        col3 = 'C' + str(i+2)

        ws[col1] = keyList[i]
        ws[col2] = valueList[i]
        ws[col3] = timesList[i]
        
    wb.save(fileName)

if __name__ == "__main__":

    json_data = open('data_set/correspondence_table.json', encoding='utf8').read()
    data = json.loads(json_data)

    for i in data:
        fileName = i['file_name']
        filePath = './data_set_wordCount/wordCount_' + fileName + '.xlsx'
        excelPath = './data_set_score/score_' + fileName + '.xlsx'

        nameList.append(fileName)
        keyList, valueList, timesList = check(filePath)
        writeExcel(excelPath, keyList, valueList, timesList, '精神疾病相關詞彙', '分數', '出現次數')

    writeExcel('./data_set_score/totalScore.xlsx', nameList, scoreList, keywordList, '檔名', '可讀性分數', '精神疾病相關詞彙總出現次數')
    


