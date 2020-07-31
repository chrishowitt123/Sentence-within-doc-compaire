#!/usr/bin/env python
# coding: utf-8

# In[12]:


import docx2txt
import itertools
import pandas as pd
from difflib import SequenceMatcher
import re
from fuzzywuzzy import fuzz
from fuzzywuzzy import process
from nltk.tokenize import PunktSentenceTokenizer

text = docx2txt.process('MY COPY_Final HLA.P.1631_Adenda do EIA do Bloco 17_20200620 - Copy EN (1) - Copy.docx')
sent_tokenizer = PunktSentenceTokenizer(text)
sents = sent_tokenizer.tokenize(text)
sents

x_list = []
y_list = []
score = []

for x,y in itertools.combinations(sents, 2):
    fuzz.ratio(x, y)
    score.append(fuzz.ratio(x, y))
    x_list.append(x)
    y_list.append(y)
    
data_tuples = list(zip(x_list,y_list,score))

results = pd.DataFrame(data_tuples, columns=['X','Y', 'Score'])  
results = results.sort_values(by=['Score'], ascending=False)
results[results['Score'] > 70].to_csv('results.csv')
