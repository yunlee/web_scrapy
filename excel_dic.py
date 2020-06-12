from bs4 import BeautifulSoup
from bs4.element import Comment
import urllib.request
import time
import pickle
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Pt
from docx.shared import Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
#from content import report_content, provinces, report_date, introduction, intro_content
alignment_dict = {'justify': WD_PARAGRAPH_ALIGNMENT.JUSTIFY,
                  'center': WD_PARAGRAPH_ALIGNMENT.CENTER,
                  'centre': WD_PARAGRAPH_ALIGNMENT.CENTER,
                  'right': WD_PARAGRAPH_ALIGNMENT.RIGHT,
                  'left': WD_PARAGRAPH_ALIGNMENT.LEFT}

orientation_dict = {'portrait': WD_ORIENTATION.PORTRAIT,
                    'landscape': WD_ORIENTATION.LANDSCAPE}
                    
def add_content(content, space_after, font_name=None, font_size=16, line_spacing=0, space_before=0,
                align='justify', keep_together=True, keep_with_next=False, page_break_before=False,
                widow_control=False, set_bold=False, set_italic=False, set_underline=False, set_all_caps=False,style_name=""):
    paragraph = document.add_paragraph(content)
    if style_name!="":
        paragraph.style = document.styles.add_style(style_name, WD_STYLE_TYPE.PARAGRAPH)
    font = paragraph.style.font
    font.name = font_name
    font.size = Pt(font_size)
    font.bold = set_bold
    font.italic = set_italic
    font.all_caps = set_all_caps
    font.underline = set_underline
    paragraph_format = paragraph.paragraph_format
    paragraph_format.alignment = alignment_dict.get(align.lower())
    paragraph_format.space_before = Pt(space_before)
    paragraph_format.space_after = Pt(space_after)
    paragraph_format.line_spacing = line_spacing
    paragraph_format.keep_together = keep_together
    paragraph_format.keep_with_next = keep_with_next
    paragraph_format.page_break_before = page_break_before
    paragraph_format.widow_control = widow_control


url_list = ['https://dictionary.cambridge.org/us/dictionary/english/', 'https://www.ldoceonline.com/dictionary/']


def scrape_example(url, word):
    opener = urllib.request.build_opener()
    opener.addheaders = [('User-agent', 'Mozilla/5.0')]
    response = opener.open(url + word)
    html_contents = response.read()
    soup = BeautifulSoup(html_contents, 'html.parser')
    example_list = list()
    if "cambridge" in url:
        example_list = example_list + [s.get_text().strip() for s in soup.find_all("div", class_="examp dexamp")]
        example_list = example_list + [s.get_text().strip() for s in soup.find_all("li", class_="eg dexamp hax")]
    else:
        example_list = example_list + [s.get_text().strip() for s in soup.find_all("span", class_="EXAMPLE")]
    return example_list

import xlrd
workbook = xlrd.open_workbook("~/Downloads/voc19000-20000.xlsx")
group  = "Sheet2"
worksheet = workbook.sheet_by_name(group)
print(group)
word_list = []
for i in range(worksheet.ncols):
    col = worksheet.col(i)
    word_list += [item.value.strip().replace(" ", "") for item in col if item.value.strip().replace(" ", "")]

word_example_dict = dict()
for word in word_list:
    word = word.strip().replace(" ", "")
    print(word)
    time.sleep(0.1)
    example_list = list()
    for url in url_list:
        example_list = example_list + scrape_example(url, word)
    word_example_dict[word] = example_list

with open(f'new_{group}.pickle', 'wb') as f:
    pickle.dump(word_example_dict, f)

document = Document()

for word, example_list in word_example_dict.items():
    add_content(word + ":", align='Left', space_before=0, space_after=0, line_spacing=1, font_size=12,
                set_bold=False, set_all_caps=False, set_underline=True)
    for example in example_list:
        add_content(example, align='Left', space_before=0, space_after=0, line_spacing=1, font_size=12,
                    set_bold=False, set_all_caps=False, set_underline=True)

document.save(f'new_{group}.docx')