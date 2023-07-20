import fitz 
import pandas as pds
fitz.TOOLS.set_small_glyph_heights(True)
import os

def _check_contain(r_word, points):
    _threshold_intersection = 0.50
    r = fitz.Quad(points).rect
    r.intersect(r_word)
    if r.getArea() >= r_word.getArea() * _threshold_intersection:
        contain = True
    else:
        contain = False
    return contain

p = r'' #Add PDF file path and name here 
doc = fitz.open(p+".pdf") 
annotation = []
text=[]
sentence_useless=''
for i in range(doc.pageCount):
    page = doc[i]
    words_on_page = page.getText("words")
    for annot in page.annots(types=[fitz.PDF_ANNOT_HIGHLIGHT, fitz.PDF_ANNOT_UNKNOWN, fitz.PDF_ANNOT_POPUP, fitz.PDF_ANNOT_FREE_TEXT, fitz.PDF_ANNOT_TEXT]):
        if annot.info["content"] =='':
            quad_points = annot.vertices
            quad_count = int(len(quad_points) / 4)
            sentences = ['' for i in range(quad_count)]
            for i in range(quad_count):
                points = quad_points[i * 4: i * 4 + 4]
                words = [w for w in words_on_page if _check_contain(fitz.Rect(w[:4]), points) ]
                sentences[i] = ' '.join(w[4] for w in words)
            sentence_useless = sentence_useless + ' '.join(sentences)
        else:
            quad_points = annot.vertices
            quad_count = int(len(quad_points) / 4)
            sentences = ['' for i in range(quad_count)]
            for i in range(quad_count):
                points = quad_points[i * 4: i * 4 + 4]
                words = [w for w in words_on_page if _check_contain(fitz.Rect(w[:4]), points) ]
                sentences[i] = ' '.join(w[4] for w in words)
            sentence = ' '.join(sentences)
            sentence = sentence_useless+" "+sentence
            text.append(sentence.strip())
            annotation.append(annot.info["content"])
            sentence_useless=''

sentiment = []
author = []

for i in annotation:
    try:
        new = i.split(",");
        author.append(new[1])
        sentiment.append(new[0])
    except:
        sentiment.append(i)
        author.append(" ")
    
df = pds.DataFrame({'Sentiment':sentiment,'Author':author,'Sentence':text})
df.to_excel(p+".xlsx", index=False, encoding='utf-8') 