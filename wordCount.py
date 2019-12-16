#統計字數寫入excel
import sys
from imp import reload
reload(sys)
import json
import jieba
import jieba.analyse
#write exl
from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook
from openpyxl.writer.excel import ExcelWriter
# from openpyxl.cell import get

def cutWord(fileName):
    
    wordList = []

    for line in open(fileName, encoding = 'utf-8'):
        item = line.strip('\n\r').split('，')

        for segment in item:
            # jieba 分詞
            tags = jieba.analyse.extract_tags(segment,100)
            for t in tags:
                wordList.append(t)

    return wordList

def wordCount(wordList):

    wordDict = {}
    keyList = []

    with open('wordCount.txt', 'w', encoding='utf-8') as wf2:
        # 開檔
        for item in wordList:
            if item not in wordDict:
                wordDict[item] = 1
            else:
                wordDict[item] += 1

        orderList = list(wordDict.values())
        orderList.sort(reverse=True)
        # print orderList
        for i in range(len(orderList)):
            for key in wordDict:
                if wordDict[key] == orderList[i]:
                    wf2.write(key + ' ' + str(wordDict[key]) + '\n')
                    keyList.append(key)
                    wordDict[key] = 0

    return keyList,orderList

def writeExcel(orderList, keyList, fileName):
    
    wb = Workbook()
    ws = wb.active
    
    for i in range(len(keyList)):
        text1 = 'A' + str(i+1)
        text2 = 'B' + str(i+1)

        ws[text1] = keyList[i]
        ws[text2] = orderList[i]
        
    wb.save(fileName)

if __name__ == "__main__":

    json_data = open('data_set/correspondence_table.json', encoding='utf8').read()
    data = json.loads(json_data)

    for i in data:
        for k,v in i.items():
            if k == 'file_name':
                fileName = 'data_set/' + str(v) + '.txt'
                wordList = cutWord(fileName)
                keyList,orderList = wordCount(wordList)

                excelName = 'wordCount_' + str(v) + '.xlsx'
                writeExcel(orderList,keyList,excelName) 
