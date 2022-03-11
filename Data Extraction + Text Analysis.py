#DATA EXTRACTION:

# My prefered Web scraping tool is Selenium.
import os
import pandas as pd
import selenium
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import sys

ser = Service(r'D:\Desktop\ML\Projects\WebScraping+Analysis Internship Assignment\webdriver\chromedriver.exe')
driver = webdriver.Chrome(service=ser)
df = pd.read_excel('Output Data Structure.xlsx')
urls = df['URL'].tolist()

os.chdir('./articles')

for i in range(len(urls)):
    sys.stdout.write('\r')
    sys.stdout.write("[%-20s] %d/%d" % ('='*(int(20*i/len(urls))) +'>', i+1, len(urls)))
    sys.stdout.flush()
    
    try:
        driver.get(urls[i])
        search = driver.find_element(By.CLASS_NAME,"td-post-content")
    except UnicodeEncodeError:
        print(f"Error in extracting {i}")
        search = ''
    
    with open(f'{i+1}.txt', 'w', encoding="utf-8") as f:
        f.write(search.text)

os.chdir('../')
#TEXT ANALYSIS:
from nltk.tokenize import word_tokenize,sent_tokenize
import re
stopwords = open("StopWords_Generic.txt","r").read()
stopwords = stopwords.split('\n')
masterDict = pd.read_csv('MasterDictionary.csv')
positive = masterDict[masterDict['Positive']!=0]['Word'].tolist()
negative = masterDict[masterDict['Negative']!=0]['Word'].tolist()
prp = re.compile(r'\b(I|we|my|ours|(?-i:us))\b',re.I)

def count_syllables(word):
    word = word.lower()
    count = 0
    vowels = 'aeiouy'
    if word[0] in vowels:#for the first vowel
        count += 1
    for i in range(1, len(word)):
        if word[i] in vowels and word[i-1] not in vowels:#eliminates grouped vowels
            count += 1
    if word.endswith("e"):#usually this isn't considered as a syllable
        count -= 1
    if count == 0:#The word should have at least one syllable.
        count += 1
    return count

for i in range(1,len(urls)+1):
    text = open(f"articles/{i}.txt","r",encoding="utf-8").read()
    words = word_tokenize(text)
    
    #text cleaning:
    words = [word for word in words if word.isalpha()]
    
    #removing stopwords:
    cleaned_text= [word for word in words if word.upper() not in stopwords]
    #Word Count:
    word_count = len(cleaned_text)
    
    #Extracting Derived Variables:
    positive_words = [word for word in cleaned_text if word.upper() in positive]
    negative_words = [word for word in cleaned_text if word.upper() in negative]
    positive_score = len(positive_words)
    negative_score = len(negative_words)
    polarity_score = (positive_score-negative_score)/((positive_score+negative_score)+0.000001)
    subjectivity_score = (positive_score+negative_score)/(word_count+0.000001)
    
    #Analysis of Readability:
    sents = sent_tokenize(text)
    avg_sentence_length = len(words)/len(sents)
    complex_count = 0
    for word in words:
        if count_syllables(word)>2:
            complex_count+=1
    perc_complex_words = complex_count/len(words)
    fog_index = 0.4 * (avg_sentence_length + perc_complex_words)
    
    #Personal Pronouns:
    personal_pronouns = len(prp.findall(text))
    
    #Average Word Length:
    sum_chars = 0
    for word in words:
        sum_chars+=len(word)
        avg_word_length = sum_chars/len(words)
    
    #Syllables per word:
    tot_syllables = 0
    for word in cleaned_text:
        tot_syllables += count_syllables(word)
    syllables_per_word = tot_syllables/word_count
    df.loc[df['URL_ID']==i,'POSITIVE SCORE'] = round(positive_score,2)
    df.loc[df['URL_ID']==i,'NEGATIVE SCORE'] = round(negative_score,2)
    df.loc[df['URL_ID']==i,'POLARITY SCORE'] = round(polarity_score,2)
    df.loc[df['URL_ID']==i,'SUBJECTIVITY SCORE'] = round(subjectivity_score,2)
    df.loc[df['URL_ID']==i,'AVG SENTENCE LENGTH'] = round(avg_sentence_length,2)
    df.loc[df['URL_ID']==i,'PERCENTAGE OF COMPLEX WORDS'] = round(perc_complex_words,2)
    df.loc[df['URL_ID']==i,'FOG INDEX'] = round(fog_index,2)
    df.loc[df['URL_ID']==i,'AVG NUMBER OF WORDS PER SENTENCE'] = round(avg_sentence_length,2)
    df.loc[df['URL_ID']==i,'COMPLEX WORD COUNT'] = round(complex_count,2)
    df.loc[df['URL_ID']==i,'WORD COUNT'] = round(word_count,2)
    df.loc[df['URL_ID']==i,'SYLLABLE PER WORD'] = round(syllables_per_word,2)
    df.loc[df['URL_ID']==i,'PERSONAL PRONOUNS'] = round(personal_pronouns,2)
    df.loc[df['URL_ID']==i,'AVG WORD LENGTH'] = round(avg_word_length,2)
    
df.to_excel('Output.xlsx')