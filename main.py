import os
import requests
import re
from bs4 import BeautifulSoup
from docx import Document
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

# https://mp.weixin.qq.com/s/SZ-29HwfcXmERKsHhdcoYw
def download_images(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'lxml')
    meta_element = soup.find('meta', {'property': 'og:title'})
    title = meta_element['content']
    title = re.sub('[/\:*?<>"|]', ' ', title)
    title = "2023."+input("Please input date: ")+" "+title
    os.mkdir(title)
    os.chdir(title)
    os.mkdir(title)
    doc = Document()
    doc.styles['Normal'].font.name = u'宋体'
    doc.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    doc.styles['Normal'].font.size = Pt(10.5)
    doc.styles['Normal'].font.color.rgb = RGBColor(0, 0, 0)
    doc.add_paragraph(title)
    tags=soup.find_all(lambda tag: tag.name == 'p' and
                                        tag.has_attr('style') and
                                        'text-indent: 2em;' in tag['style'] or tag.name=='img')
    flag = 0
    index = 0
    for tag in tags:
        if tag.name=='img':
            image_url = tag.get('data-src')
            if image_url is not None:
                if flag == 1 and not image_url.endswith("svg"):
                    response = requests.get(image_url)
                    index = index + 1
                    with open(os.path.join(title, str(index) + ".jpg"), 'wb') as f:
                        f.write(response.content)
                    doc.add_picture(os.path.join(title, str(index) + ".jpg"), width=Inches(3))
                if image_url.endswith("svg"):
                    flag = flag + 1
        else:
            content = tag.text
            doc.add_paragraph(content)
    # if not os.path.exists(title):
    doc.save(title + '.docx')


if __name__ == '__main__':
    print("Made By Jing. I love you~")
    url = input("Please input url:")
    download_images(url)
    print("Success.")
    os.system('explorer ..')
    os.system('pause')
