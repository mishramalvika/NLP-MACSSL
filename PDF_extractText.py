import PyPDF4 as pd
import re
import pandas as pds
import nltk
nltk.download('punkt')

p = r'' #Add PDF file path and name here
filename = open(p + '.pdf', 'rb') 
file = pd.PdfFileReader(filename)
str=''

for i in range(file.getNumPages()): 
    page = file.getPage(i)
    text = page.extractText().rstrip()
    str = str + text      

#replacing TM with approstrophe(')
str = str.replace(u"\u2122", u"\u2019")

#To prevent Page Numbers to get included in the sentences
str= re.sub('\- \d \-+', ' ', str)
str= re.sub('\- \d\d \-+', ' ', str)
str= re.sub('\- \d\d\d \-+', ' ', str)

#To prevent Page Dates to get included in the sentences - PDF specific patterns
str= re.sub('EBM\/\d\d\/\d\d \- \d\/\d\d\/\d\d+', ' ', str)
str= re.sub('EBM\/\d\d\/\d\d \- \d\d\/\d\d\/\d\d+', ' ', str)
str= re.sub('EBM\/\d\d\/\d\d\d \- \d\d\/\d\d\/\d\d+', ' ', str)
str= re.sub('EBM\/\d\d\/\d \- \d\/\d\/\d\d+', ' ', str)
str= re.sub('EBM\/\d\d\/\d \- \d\/\d\d\/\d\d+', ' ', str)
str= re.sub('EBM\/\d\d\/\d\-\d \- \d\/\d\d\/\d\d+', ' ', str)

str = re.sub(r'(?<=[.])(?=[a-zA-Z])', r' ', str)
str = str.replace('\n', '').replace('\r', '')
str =' '.join(str.split('\t'))
str =' '.join(str.split('  '))

new_str = nltk.tokenize.sent_tokenize(str)

df = pds.DataFrame({'Sentiment':'','Author':'','Sentence':new_str})
df.to_excel(p+'.xlsx', index=False, encoding='utf-8') 