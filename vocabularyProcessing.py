def checkRepeat(wordList, word):
    for text in wordList:
        if word == text:
            return True
    return False

def getHappyWord():
    filepath = 'vocabulary/happy.txt'
    filetext = open(filepath,'r+', encoding='utf-8')
    happy_word = []
    for line in filetext:
        word = line.split('\n')[0]
        if not(checkRepeat(happy_word, word)):
            happy_word.append(word)
    return happy_word

def getUnhappyWord():
    filepath = 'vocabulary/unhappy.txt'
    filetext = open(filepath,'r+', encoding='utf-8')
    unhappy_word = []
    for line in filetext:
        word = line.split('\n')[0]
        if not(checkRepeat(unhappy_word, word)):
            unhappy_word.append(word)
    return unhappy_word

def getMentalDisorderName():
    filepath = 'vocabulary/mental_disorder.txt'
    filetext = open(filepath,'r+', encoding='utf-8')
    disorder_name = []
    for line in filetext:
        word = line.split('\n')[0]
        if not(checkRepeat(disorder_name, word)):
            disorder_name.append(word)
    return disorder_name

def updateVocabulary(wordList, filepath):
    with open(filepath, 'w', encoding='utf-8') as wf2:
        
        for i in range(len(wordList)):
            wf2.write(wordList[i] + '\n')

