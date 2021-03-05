import pandas as pd
import tkinter
from tkinter import filedialog
from iteration_utilities import deepflatten
from collections import Counter 
import docx2txt
from nltk.tokenize import PunktSentenceTokenizer
import nltk
nltk.download('punkt')
from termcolor import colored
import re
import itertools
import pandas as pd
from rapidfuzz import fuzz

pd.set_option('max_colwidth', 50)

'''MATCHING SENTENCES'''


filename = r"C:\Users\chris\Documents\Transgola\Clients\PROJECTS\2021\387010321_TM_JTI\Translation\final\Contrato tipo Sociprime_PAIRS_EN_TM.docx"
text = docx2txt.process(filename)

sents = text.replace("\t","").replace("No. ", "No.").replace("\r", "").replace("\n", ". ").replace(": ", ". ").replace("; ", ". ").split(". ")  


# sent_tokenizer = PunktSentenceTokenizer(text)
# sents = sent_tokenizer.tokenize(text)
# sents

allSents = []

for sent in sents:
    s = sent.split("\n")
    allSents.append(s)
    
allSents = list(deepflatten(allSents, depth=1))
allSents = [ x.replace('\t', '') for x in allSents]
allSents = [x for x in allSents if x.strip()] 
allSents = [x for x in allSents if len(x.split(' ')) > 4 ]

repeaters  = Counter(allSents)
repeaters  = {k:v for (k,v) in repeaters.items() if v > 1}

def replace(string, char): 
    pattern = char + '{2,}'
    string = re.sub(pattern, char, string) 
    return string 
textClean = replace(text,'\n')

textClean = textClean.replace('\n', ' \n ').replace('\t', ' ')
RepStrings = list(repeaters.keys())

for s in RepStrings:
    if s in textClean:
        textClean= textClean.replace(s, colored(s, 'blue',  attrs=['bold']))
        
# print(textClean)       





'''SIMILAR STRINGS'''

sents = set(sents)

x_list = []
y_list = []
score = []

for x,y in itertools.combinations(sents, 2):
    fuzz.ratio(x, y)
    score.append(fuzz.ratio(x, y))
    x_list.append(x)
    y_list.append(y)
    
# remove consecutive blank lines

x_list1 = []  
for x in x_list:

    xn = re.sub(r'\n\s*\n', '\n\n', x)
    x_list1.append(xn)
    
y_list1 = []  
for y in y_list:

    yn = re.sub(r'\n\s*\n', '\n\n', y)
    y_list1.append(yn)
    
data_tuples = list(zip(x_list1,y_list1,score))

results = pd.DataFrame(data_tuples, columns=['X','Y', 'Score'])  

results = results.sort_values(by=['Score'], ascending=False)
results = results[results['Score'] >60]

x_list3 = list(results['X'])
y_list3 = list(results['Y'])
        
    
# uncommon words

diffs = []


def find(X, Y):
    count = {}
    for word in X.split():
        count[word] = count.get(word, 0) + 1

    for word in Y.split():
        count[word] = count.get(word, 0) + 1
    return [word for word in count if count[word] == 1]



for X,Y in zip(x_list3, y_list3):
    diffs.append((find(X, Y)))
    
diffsList = [' '.join(x) for x in diffs]
results['Diffs'] = diffsList
results = results[['Score', 'X', 'Y', 'Diffs']]


resultsXlist = results['X'].tolist()
resultsYlist = results['Y'].tolist()
resultDIFFSYlist = results['Diffs'].tolist()
resultSCORElist  = results['Score'].tolist()


print( "\n")
print(f'Repeated strings!')
print( "\n")

for s in RepStrings:
    print(colored(s, 'blue', attrs=['bold']))
    print( "\n")
    
print(textClean.replace('\n', ' '))
print( "\n")
print( "\n")

n = 0
while n <= len(resultsXlist) - 1:

    text1 = resultsXlist[n]  
    text2 = resultsYlist[n] 
    l1 = resultDIFFSYlist[n].split()

    
    
    formattedText1 = []
    for t in text1.split():
        if t in l1:
            formattedText1.append(colored(t,'red', attrs=['bold']))
        else: 
            formattedText1.append(t)

    
    formattedText2 = []
    for t in text2.split():
        if t in l1:
            formattedText2.append(colored(t,'red', attrs=['bold']))
        else: 
            formattedText2.append(t)
            
            
            
                 
            
    print( "\n")
    print(colored(resultSCORElist[n], 'green', attrs=['bold']))
    print(colored(l1, 'blue', attrs=['bold']))
    print( "\n")
    print(" ".join(formattedText1))
    print( "\n")
    print(" ".join(formattedText2))
    print( "\n")
    print( "\n")
    
    n = n+1
