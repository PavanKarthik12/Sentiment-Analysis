import xlrd
import nltk 
from nltk.corpus import stopwords 
from nltk.tokenize import word_tokenize, sent_tokenize
from spellchecker import SpellChecker
from nltk.stem import WordNetLemmatizer
from matplotlib import pyplot as plt
stop_words = set(stopwords.words('english')) 
loc = ("C:\\Users\\Pavan DS\\Desktop\\excell.xlsx") 
wb = xlrd.open_workbook(loc) 
sheet = wb.sheet_by_index(0) 
emotion=["joy","disgust","sadness","trust","anger","negative","fear","surprise","anticipation","positive"]
Trust=0
Anticipation=0
Joy=0
Anger=0
Disgust=0
Sadness=0
Alpha=0.6
cnt=0
result=[]
def clearing(text):
    stop_words = set(stopwords.words('english'))
    word_tokens = word_tokenize(text)
    filtered_sentence = [w for w in word_tokens if not w in stop_words]
    filtered = []
    for w in word_tokens: 
        if w not in stop_words: 
            filtered.append(w)
    refined=[i for i in filtered if  i.isalnum()]
    lower_case=[i.lower() for i in refined]
    print("stop words cleared",lower_case)
    first=spellcheck(lower_case)
    second=lemmitization(first)
    return second   
def spellcheck(filtered_sentence):
    spellclear=[]
    spell = SpellChecker()
    for word in filtered_sentence:
        spellclear.append(spell.correction(word))
    print("spell clear",spellclear)
    return spellclear
def lemmitization(spellclear):
    lemmatizer = WordNetLemmatizer()
    lemmatized=[]
    for i in spellclear:
        lemmatized.append(lemmatizer.lemmatize(i))
    print("Lemmatized",lemmatized)
    return lemmatized
def visualize(satper,disper):
    YAxis=["Positive","Negative"]
    XAxis=[satper,disper]
    plt.ylabel("Percentage")
    plt.bar(YAxis,XAxis)
    plt.show()
def satisfaction(trust,anticipation,joy,n):
    alpha=0.6
    sat=(alpha*(trust+anticipation)+(1-alpha)*(joy))/n
    return sat
def dissatisfaction(anger,disgust,sadness,n):
    alpha=0.6
    dis=(alpha*(anger+disgust)+(1-alpha)*(sadness))/n
    return dis
feedback=[]
while(True):
    inp=input("Enter feedback")
    if inp=="stop":
        break
    else:
        feedback.append(inp)
for i in feedback:
    final=clearing(i)
    result=[]
    for i in final:
        temp=[]
        for j in range(sheet.nrows):      
            if sheet.cell_value(j,0)==i:
                temp.append({i:[sheet.cell_value(j, 1),sheet.cell_value(j, 2)]})
        result.append(temp)
    for i in result:
        for j in i:
            temp=list(j.values())
            if temp[0][0]=='joy' and temp[0][1]=='1':
                Joy+=1
            if temp[0][0]=='trust' and temp[0][1]=='1':
                Trust+=1
            if temp[0][0]=='anticipation' and temp[0][1]=='1':
                Anticipation+=1
            if temp[0][0]=='anger' and temp[0][1]=='1':
                Anger+=1
            if temp[0][0]=='disgust' and temp[0][1]=='1':
                Disgust+=1
            if temp[0][0]=='sadness' and temp[0][1]=='1':
                Sadness+=1           
sat=satisfaction(Trust,Anticipation,Joy,len(feedback))
dis=dissatisfaction(Anger,Disgust,Sadness,len(feedback))
total=sat+dis
satper=(sat/total)*100
disper=(dis/total)*100
visualize(satper,disper)



    

