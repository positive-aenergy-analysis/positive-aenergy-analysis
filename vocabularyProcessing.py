def checkRepeat(wordList,word):
    for text in wordList:
        if word == text:
            return True
    return False

def getHappyWord():
    filepath = '../詞彙庫/happy.txt'
    filetext = open(filepath,'r+',encoding='utf-8')
    happy_word = []
    for line in filetext:
        word = line.split('\n')[0]
        if not(checkRepeat(happy_word,word)):
            happy_word.append(word)
    return happy_word

def getUnhappyWord():
    filepath = '../詞彙庫/unhappy.txt'
    filetext = open(filepath,'r+',encoding='utf-8')
    unhappy_word = []
    for line in filetext:
        word = line.split('\n')[0]
        if not(checkRepeat(unhappy_word,word)):
            unhappy_word.append(word)
    return unhappy_word

def getMentalDisorderName():
    filepath = '../詞彙庫/精神疾病.txt'
    filetext = open(filepath,'r+',encoding='utf-8')
    disorder_name = []
    for line in filetext:
        word = line.split('\n')[0]
        if not(checkRepeat(disorder_name,word)):
            disorder_name.append(word)
    return disorder_name

print(getHappyWord())