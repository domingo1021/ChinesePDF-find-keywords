import zipfile
import xml.etree.ElementTree
import word
import os
import pandas as pd

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'
TABLE = WORD_NAMESPACE + 'tbl'
ROW = WORD_NAMESPACE + 'tr'
CELL = WORD_NAMESPACE + 'tc'

def process_docx_file(filename:str)-> list:
    with zipfile.ZipFile(filename) as docx:
        tree = xml.etree.ElementTree.XML(docx.read('word/document.xml'))

    text_content = []
    pure_text = ""
    for paragraph in tree.iter(PARA):
        texts = [node.text
                    for node in paragraph.iter(TEXT)
                    if node.text]
        if texts:
            text_content.append(''.join(texts))

    for content in text_content:
        pure_text+=content

    # print(pure_text)
    nums, contents = word.split_pages(pure_text)
    re_keywords = word.re_keywords_string("外匯")
    return word.find_keywords_pages(re_keywords, contents, nums)


df = pd.DataFrame(None)
for filename in os.listdir("./words"):
    print(type(filename))
    keyword_pages = process_docx_file("./words/"+filename)
    ticker = filename.split("_")[0]
    df[ticker] = pd.Series(keyword_pages)

print(df)
    
    
# table_content=""
# for index, table in enumerate(tree.iter(TABLE)):
#     print(index)
#     for row in table.iter(ROW):
#         for cell in row.iter(CELL):
#             table_text = ""
#             for node in cell.iter(TEXT):
#                 table_text+=node.text
#             print(table_text)
#             table_content= table_content+''.join(node.text for node in cell.iter(TEXT))
# print(table_content)

# nums, contents = word.split_page(table_content)
# re_keywords = word.re_keywords_string("外匯")
# print(word.find_keywords_page(re_keywords, contents, nums))
