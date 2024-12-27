#!/usr/bin/env python3
#-*-coding:utf8-*-

import sys,os
from bs4 import BeautifulSoup
import cairosvg
from PyPDF2 import PdfReader,PdfWriter
from reportlab.pdfgen import canvas
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfbase import pdfmetrics
from io import BytesIO, StringIO
from datetime import datetime
class Examinate:
    def __init__(self):
        self.namebox = [110.49, 661.634, 146.49,  675.694]
        self.genderbox= [249.64, 661.634, 261.64, 675.694]
        self.agebox = [346.66, 661.634, 360.34, 675.694]
        self.bloodbox  = [486.35, 661.534, 512.35, 675.594]
        self.receibox  = [155.73, 609.004, 219.81, 623.074]
        self.codebox = [378.20, 609.004, 452.28, 623.074]
        self.reportbox = [155.73, 574.994, 219.81, 597.064]
        self.hosbox    = [368.12, 574.994, 442.28, 597.064]
#        self.codebox   = [384.99, 588.494, 521.66, 602.654]
        self.pngbox    = [175.71, 458.03 , 520.00, 518.644]
        self.image     = sys.path[0] + "/template/score.svg"
        self.highrisk  = sys.path[0] + "/template/甘倍康2.0对医院端报告常规版-高风险v3.pdf"
        self.lowrisk   = sys.path[0] + "/template/甘倍康2.0对医院端报告常规版-低风险v3.pdf"
        self.name      = None
        self.gender    = None
        self.age       = None
        self.blood     = None
        self.date1     = None
        self.date2     = None
        self.date3     = None
        self.hospital  = None
        self.code      = None
        self.riskscore = None


def flatten(lst):
    result = []
    for item in lst:
        if isinstance(item, list):
            result.extend(flatten(item))
        else:
            result.append(item)
    return result

def DateFormat(date) ->str:
#    print(date)
    date = datetime.strptime(date,"%m-%d-%y")
    outdate = date.strftime('%Y-%m-%d')
#    print(outdate)
    return outdate
def get_score_ico_locate_amplfer2(score, svgsoup):
    html_score_ico_locate = [i.split(',') for i in svgsoup.find(id='svg_135')['points'].split()]
    html_score_ico_locate[0][0] = score
    html_score_ico_locate[1][0] = html_score_ico_locate[0][0] - 3.8
    html_score_ico_locate[2][0] = html_score_ico_locate[0][0] + 3.8
    html_score_polygonlocate = flatten(html_score_ico_locate)
    html_score_ico_newlocate = " ".join([str(i) for i in html_score_polygonlocate])
    #print(html_score_ico_newlocate)
    html_score_ico_score_newlocate = html_score_ico_locate[0][0]
    return (html_score_ico_newlocate, html_score_ico_score_newlocate)

def svgsoap_to_png(svgsoup, pngfile, scale=1.5):
    result=str(svgsoup)
    result=result.replace('lineargradient','linearGradient')
    result=result.replace('clippath','clipPath')
    result=result.replace('textlength','textLength')
    result=result.replace('viewbox','viewBox')
    svgpage=StringIO(result)
    cairosvg.svg2png(file_obj=svgpage,write_to=pngfile,scale=scale)
    svgpage.close()
    return pngfile


def get_slideblock(score, svgfile):
    svg = open(svgfile)
    svgdata = svg.read()
    svgsoup = BeautifulSoup(svgdata, 'html.parser')
#    print(svgsoup.find(id='svg_138'))
    #html_score = svgsoup.find(id='svg_223').string.replace_with(str(UriFind_score))
    #html_result = svgsoup.find(id='svg_224').string.replace_with(UriFind_result)
    html_threshodvalue = int(svgsoup.find(id='svg_137').string.encode("utf-8"))
    html_threshodvalue_locate = float(svgsoup.find(id='svg_137')['x'])  # 50
    html_zerovalue_locate = float(svgsoup.find(id='svg_138')['x'])  # 0
    html_fullvalue_locate = float(svgsoup.find(id='svg_136')['x'])  # 100
    html_score_ico_score = svgsoup.find(id='svg_130').string.replace_with(str(score))
    if float(score) == 0:
        html_score_ico_locate_center = html_zerovalue_locate
        html_score_ico_newlocate, html_score_ico_score_newlocate = get_score_ico_locate_amplfer2(
            html_score_ico_locate_center, svgsoup)
        svgsoup.find(id='svg_135')['points'] = html_score_ico_newlocate
        svgsoup.find(id='svg_135')['fill'] = '#19B797'
        svgsoup.find(id='svg_130')['x'] = html_score_ico_score_newlocate
    elif float(score) == html_threshodvalue:
        html_score_ico_locate_center = html_threshodvalue_locate
        html_score_ico_newlocate, html_score_ico_score_newlocate = get_score_ico_locate_amplfer2(
            html_score_ico_locate_center, svgsoup)
        svgsoup.find(id='svg_135')['points'] = html_score_ico_newlocate
        svgsoup.find(id='svg_135')['fill'] = '#19B797'
        svgsoup.find(id='svg_130')['x'] = html_score_ico_score_newlocate
    elif float(score) == 100:
        html_score_ico_locate_center = html_fullvalue_locate
        html_score_ico_newlocate, html_score_ico_score_newlocate = get_score_ico_locate_amplfer2(
            html_score_ico_locate_center, svgsoup)
        svgsoup.find(id='svg_135')['points'] = html_score_ico_newlocate
        svgsoup.find(id='svg_135')['fill'] = '#ED4B56'
        svgsoup.find(id='svg_130')['x'] = html_score_ico_score_newlocate
    elif float(score) > 0 and float(score) < float(html_threshodvalue):
        html_zero_threshodvaluelen = html_threshodvalue_locate - html_zerovalue_locate
        html_score_ico_locate_center = float(score) * (
                    html_zero_threshodvaluelen / html_threshodvalue) + html_zerovalue_locate
        html_score_ico_newlocate, html_score_ico_score_newlocate = get_score_ico_locate_amplfer2(
            html_score_ico_locate_center, svgsoup)
        svgsoup.find(id='svg_135')['points'] = html_score_ico_newlocate
        svgsoup.find(id='svg_135')['fill'] = '#19B797'
        svgsoup.find(id='svg_130')['x'] = html_score_ico_score_newlocate
    else:
        html_threshod_fullvaluelen = html_fullvalue_locate - html_threshodvalue_locate
        html_score_ico_locate_center = (float(score) - html_threshodvalue) * (
                    html_threshod_fullvaluelen / (100 - html_threshodvalue)) + html_threshodvalue_locate
        html_score_ico_newlocate, html_score_ico_score_newlocate = get_score_ico_locate_amplfer2(
            html_score_ico_locate_center, svgsoup)
        svgsoup.find(id='svg_135')['points'] = html_score_ico_newlocate
        svgsoup.find(id='svg_135')['fill'] = '#ED4B56'
        svgsoup.find(id='svg_130')['x'] = html_score_ico_score_newlocate
    return svgsoup


# 创建一个新的PDF文件来加入文本和图片
def create_overlay_pdf(person:object,outpath:str):
    '''
    :param person:object for personal information
    :return:pdf object
    '''
    if person.riskscore >=50:
        template_pdf_path = person.highrisk
    else:
        template_pdf_path = person.lowrisk
 #   print(template_pdf_path)
    pdfmetrics.registerFont(TTFont('HarmonyOS_Sans_SC', sys.path[0]+'/template/HarmonyOS_SansSC_Light.ttf'))
    template = PdfReader(template_pdf_path)
    page1 = template.pages[0]
    width, height = page1.mediabox[2], page1.mediabox[3]
#   print(width,height)

    packet = BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    #c.setFillColorRGB(1, 1, 1)
    ##检测者姓名
    newname = person.name
    x0, y0 = person.namebox[0], person.namebox[1]
    name_width, name_height = person.namebox[2] - x0, person.namebox[3] - y0
    c.rect(x0, y0, name_width, name_height, stroke=0)
    name_project = c.beginText(x0, y0)
    name_project.setFont("HarmonyOS_Sans_SC",12)
    name_project.textLines(newname)
    c.drawText(name_project)
    ##检测者性别
    gender = person.gender
    x1,y1  = person.genderbox[0], person.genderbox[1]
    gender_width, gender_height = person.genderbox[2] - x1, person.genderbox[3] - y1
    c.rect(x1, y1, gender_width, gender_height, stroke=0)
    gender_project = c.beginText(x1,y1)
    gender_project.setFont("HarmonyOS_Sans_SC",12)
    gender_project.textLines(gender)
    c.drawText(gender_project)
    ##检测者年龄
    age  = person.age
    x2, y2 = person.agebox[0], person.agebox[1]
    age_width, age_height = person.agebox[2] - x2, person.agebox[3] - y2
    c.rect(x2, y2, age_width, age_height, stroke=0)
    age_project = c.beginText(x2, y2)
    age_project.setFont("HarmonyOS_Sans_SC",12)
    age_project.textLines(str(age))
    c.drawText(age_project)
    ##检测者样本类型
    blood = person.blood
    x3, y3 = person.bloodbox[0], person.bloodbox[1]
    blood_width, blood_height = person.bloodbox[2] - x3, person.bloodbox[3] - y3
    c.rect(x3, y3, blood_width, blood_height, stroke=0)
    blood_object = c.beginText(x3,y3)
    blood_object.setFont("HarmonyOS_Sans_SC",12)
    blood_object.textLines(blood)
    c.drawText(blood_object)
    ##receive date
    receive = person.date1
    receive = DateFormat(receive)
    x4, y4  = person.receibox[0],person.receibox[1]
    recei_width, recei_height = person.receibox[2] - x4, person.receibox[3] - y4
    c.rect(x4, y4, recei_width, recei_height, stroke=0)
    recei_object = c.beginText(x4, y4)
    recei_object.setFont("HarmonyOS_Sans_SC",12)
    recei_object.textLines(receive)
    c.drawText(recei_object)

    ##date2
#    date2 = person.date2
#    x5, y5 = person.detectbox[0], person.detectbox[1]
#    detect_width, detect_height = person.detectbox[2] - x5, person.detectbox[3] - y5
#    c.rect(x5, y5, detect_width, detect_height, stroke=0)
#    detect_object = c.beginText(x5, y5)
#    detect_object.setFont("HarmonyOS_Sans_SC",12)
#    detect_object.textLines(date2)
#    c.drawText(detect_object)

    ##report date
#    date3 = person.date3
    date3 = datetime.now().strftime('%Y-%m-%d')
    x6, y6 = person.reportbox[0], person.reportbox[1]
    report_width, report_height = person.reportbox[2] - x6, person.reportbox[3] - y6
    c.rect(x6, y6, report_width, report_height,stroke=0)
    report_object = c.beginText(x6, y6)
    report_object.setFont("HarmonyOS_Sans_SC",12)
    report_object.textLines(date3)
    c.drawText(report_object)

    ##检测者编号
    code1 = person.code
    x7, y7 = person.codebox[0], person.codebox[1]
    code_width, code_height = person.codebox[2] - x7, person.codebox[3] - y7
    c.rect(x7, y7, code_width, code_height,stroke=0)
    code_object = c.beginText(x7, y7)
    code_object.setFont("HarmonyOS_Sans_SC",12)
    code_object.textLines(code1)
    c.drawText(code_object)

    ## 送检科室
    hos = person.hospital
    x8, y8 = person.hosbox[0], person.hosbox[1]
    hos_width, hos_height = person.hosbox[0] - x8, person.hosbox[1] - y8
    c.rect(x8, y8, hos_width, hos_height, stroke=0)
    hos_object = c.beginText(x8,y8)
    hos_object.setFont("HarmonyOS_Sans_SC",12)
    hos_object.textLines(hos)
    c.drawText(hos_object)

    ##结果图片
#    print(person.riskscore)
    svgsoup2 = get_slideblock(person.riskscore, person.image)
    os.makedirs(outpath+"/image/" ,exist_ok=True)
    pngfile = outpath+"/image/"+person.name+"_" + person.code +"_"+person.date3+"_HCCrisk.png"#
#    print(pngfile)
    pngfile = svgsoap_to_png(svgsoup2, pngfile)
    pngx, pngy = person.pngbox[0], person.pngbox[1]
    png_width, png_height = person.pngbox[2] - pngx, person.pngbox[3] - pngy
    c.setFillColorRGB(1, 1, 1)
    c.rect(pngx, pngy, png_width, png_height, fill=1, stroke= 0)
    c.drawImage(pngfile, pngx, pngy, width=png_width, height= png_height, mask='auto')
    c.save()
    packet.seek(0)

    newpage = PdfReader(packet)
    newpage_to_merge = newpage.pages[0]
    page1.merge_page(newpage_to_merge)

    writer = PdfWriter()
    writer.add_page(page1)
    for page_num in range(1,len(template.pages)):
        writer.add_page(template.pages[page_num])
    return writer


