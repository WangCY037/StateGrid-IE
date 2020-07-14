import docx
import PyPDF2
import pdfplumber
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfdocument import PDFDocument
import os
import re


def get_metadata(source_path,filename,pageNumberStart,pageNumberEnd):
    origin = pdfplumber.open(source_path+filename)
    pages = origin.pages[pageNumberStart:pageNumberEnd]
    pattern = ".*(开标时间|投标截止时间).*"
    res = []
    for page in pages:
        texts = page.extract_text().split('\n')
        for index,text in enumerate(texts):
#             print(text)
            if re.match(pattern,text):
                res.append(text)
                res.append(texts[index+1])
                return "".join(res)

    return ""

def clean_text(texts):
    """
    清理文本 主要是清理 页眉页脚文字
    41
    中国神华国际工程有限公司招标文件
    """
    pattern_1='\d+'
    pattern_2='招标|文件|公司'

    delete_index_list=[]
    for index,line in enumerate(texts) :
        if re.fullmatch(pattern_1,line):
            delete_index_list.append(index)
            if index!=len(texts)-1:
                if re.search(pattern_2,texts[index+1]):
                    delete_index_list.append(index+1)

    new_text_content=[]   # 用于存放删除页眉页脚后的新内容        
    for index,line in enumerate(texts) :
        if index not in delete_index_list:
            new_text_content.append(line)

    return new_text_content

def pdf_parse_text(path,pageNumberStart,pageNumberEnd):
    origin = pdfplumber.open(path)
    texts = []
    for page_num in range(pageNumberStart,pageNumberEnd):
        page=origin.pages[page_num]
        text_page = [text.replace(' ','') for text in page.extract_text().split('\n') if len(text.replace(' ','')) != 0]       
        texts.extend(text_page)

    origin.close()

    return texts

# 检查该文本里是否有技术章节
def check_section_contents(texts):
    pattern_1='第.章.*(技术|参数|要求).*'
    pattern_2='\d+'
    pattern_3='招标|文件|公司'
    crop_chapters_flag='crop chapters failed'
    section_texts=[]
    for index,line in enumerate(texts) :
        if re.fullmatch(pattern_1,line):
            if index!=0:
                if re.fullmatch(pattern_2,texts[index-1]) or re.search(pattern_3,texts[index-1]):# 需要定位技术章节的开头，其上一行一般是页码或页眉
                    crop_chapters_flag='succeed'
                    section_texts=texts[index:]
                    break 

    return crop_chapters_flag,section_texts             



def get_chapters_text(path,pageNumberStart,pageNumberEnd,crop_chapters_flag,sections,pdf_obj):
    """解析pdf为"""
    # 如果上一步寻找章节成功，则直接识别文本解析
    if crop_chapters_flag=='succeed':
        texts=pdf_parse_text(path,pageNumberStart,pageNumberEnd)
        
    else:
        # 不成功进行二次寻找
        # 经过分析，该章节要么是中间丢失的章节，要么是最后的章节，因此，有下判断
        # 首先定义 字典 将汉字转化为数字
        num_character=['一','二','三','四','五','六','七','八','九','十']
        num_dict={}
        for i in range(len(num_character)):
            num_dict[num_character[i]]=i+1
        
        section_list={}# 创建section 列表 与sections里的序号编号一一对应（可能书签顺序是乱的）
        for index,section in enumerate(sections) :
            title=section.title
            character=title[1]
            section_num=num_dict[character]
            if section_num!=None:
                section_list[section_num]=index
        
        section_list_order=sorted(section_list.items(),key=lambda x:x[0],reverse=False) # 排序 用于判断是否有丢失章节

        i=1

        while(i<len(section_list_order)):  # 逐个寻找空缺的章节 

            serial_number_start,index_start=section_list_order[i-1]
            serial_number_end,index_end=section_list_order[i]
            if serial_number_end!=serial_number_start+1:  #排序后的序号 对不上 则证明中间有空缺
                pageNumberStart=pdf_obj.getDestinationPageNumber(sections[index_start])
                pageNumberEnd=pdf_obj.getDestinationPageNumber(sections[index_end])
            #直至最后一个章节，这时，再次寻找该章节到整个pdf最后一页的内容是否包含技术章节 
            elif i==len(section_list_order)-1:
                pageNumberStart=pdf_obj.getDestinationPageNumber(sections[index_end])
                pageNumberEnd=pdf_obj.getNumPages() # 直接取到最后
            else:
                i+=1
                continue

            texts=pdf_parse_text(path,pageNumberStart,pageNumberEnd)
            crop_chapters_flag,texts=check_section_contents(texts)
            if crop_chapters=='succeed':
                break
            
            i+=1
    texts=clean_text(texts)

    return texts,crop_chapters_flag


def crop_chapters(pdf_obj,source_path,filename):
    """划分章节"""

    outlines = pdf_obj.getOutlines()
    pattern_1=".*(招标公告|招标邀请书|投标邀请).*"
    pattern_2 = '.*(技术|参数|要求).*'

    sections = [] # 第* 章节 
    All_sections=[] 
    # 定义处理参数 用于判断是否处理成功
    crop_chapters_flag='crop chapters failed'
    time='get time failed'#
    chapters_contents=[]

    for index,outline in enumerate(outlines):
        if isinstance(outline,list):
            continue
#         print(outline)
#             print(page.extractText())
        All_sections.append(outline)
        if re.match("(第.*章|第.*部分)",outline.title):
            #print(outline.title)
            sections.append(outline)
    
    pageNumberStart=0
    pageNumberEnd=0

    if sections!=[]:
        for index,section in enumerate(sections):
    
            if re.match(pattern_1,section.title):
                pageNumberStart = pdf_obj.getDestinationPageNumber(section)
                #if isinstance(outlines[index+1],list):
                pageNumberEnd = pdf_obj.getDestinationPageNumber(sections[index+1])
                #print(pageNumberStart,pageNumberEnd)
                bid_date = get_metadata(source_path,filename,pageNumberStart,pageNumberEnd)

                try:
                    start,end = re.search("\d{4}(.{0,5})\d{1,2}(.{0,5})\d{1,2}(.{0,5})\d{1,2}(.{0,5})\d{1,2}",bid_date).span()  # 有很多其他格式的日期
                    time=bid_date[start:end]
                except:
                    try:
                        start,end = re.search("20\d{2}",bid_date).span()  # 有错就只找年份
                        time=bid_date[start:end]
                    except: #再有错就不管了
                        pass        
                #pdf_out.addMetadata(cc)

            elif re.match(pattern_2,section.title):
                pageNumberStart = pdf_obj.getDestinationPageNumber(section)

                if len(sections)==index+1:  # 当技术部分是前文所取到的最后一个标准 第*章节时

                    pageNumberEnd = pdf_obj.getNumPages() #找到最后一页

                else:
                    #print(sections[index+1])
                    pageNumberEnd = pdf_obj.getDestinationPageNumber(sections[index+1])

                crop_chapters_flag='succeed'
                break
        # 判断是否含有技术章节，同时解析章节文本，
        chapters_contents,crop_chapters_flag = get_chapters_text(source_path+filename,pageNumberStart,pageNumberEnd,crop_chapters_flag,sections,pdf_obj)

    return time,crop_chapters_flag,chapters_contents

def re_match_packets(section_contents):
    """匹配多包文件""" 
    packets=[]

    pattern_1 = '.*需求一览表.*'#([a-zA-Z0-9]+)'   包文件必有 货物需求一览表等内容  其下一行一般是表头 序号 名称 数量
    pattern_2='.*序号.*'#  

    regex_1 = re.compile(pattern_1)
    regex_2 = re.compile(pattern_2)


    for index,text in enumerate(section_contents):
        if index <=len(section_contents)-3: # 防止越界
            if len(regex_1.findall(text)) > 0 :
                if len(regex_2.findall(section_contents[index+1]))>0 or len(regex_2.findall(section_contents[index+2]))>0 :
                    packets.append([index-1,text])  # 一般以货物需求一览表之上的文字为开头
                    #print(section_contents[index-1])
    
    return packets

def select_packets(packet_contents,object_type):
    """
    选择要提取的标包  一个标书通常有多个标包，因此需要筛选
    这里通过标包名称（第一包 采煤机）或者 货物需求一览表里（采煤机  型号）的说明来筛选
    实现方法：
        开头前十行内是否有“标包名称” 这个十行是综合取得可能得最大值
    """
    drop_list=[]
    pattern=object_type
    regex=re.compile(pattern)
    for packet_index,packet_content in enumerate(packet_contents):
        for text_index,text in enumerate(packet_content):
            if len(regex.findall(text))>0:
                break
            if text_index>9:
                drop_list.append(packet_index)


    after_selected_packet_contents=[]
    for packet_index,packet_content in enumerate(packet_contents):
        if packet_index not in drop_list:
            after_selected_packet_contents.append(packet_content)

    return after_selected_packet_contents

def crop_packets(section_contents,object_type):
    '''处理多包文件'''
    # 接下来划分多包

    packet_index=re_match_packets(section_contents) # 获得相应小节的行号
    #print(packet_index)

    packet_contents=[]

    crop_packets_flag='crop packets failed'

    for i in range(len(packet_index)):
        # 先找到所有的包
        index,text=packet_index[i][0],packet_index[i][1]

        start = index

        if i != len(packet_index) - 1:
            end = packet_index[i+1][0]
        else:
            end = len(section_contents) # 直接取到最后

        if end - start > 1: # 过滤目录  
            packet_contents.append(section_contents[start:end])
            crop_packets_flag='succeed'
            #print(text)
            

    total_packet_count=len(packet_contents)
    "筛选"
    after_selected_packet_contents=select_packets(packet_contents,object_type)

    valid_packet_count=len(after_selected_packet_contents)

    return total_packet_count,valid_packet_count,after_selected_packet_contents,crop_packets_flag



def final_extract(index,packet_index,packet_content,filename,time,parameter_belonging_to):
    """
    最终提取模块 提取格式如下：
    参数类别 参数名 参数值  时间戳 参数归属
    parameter_type parameter_name parameter_value time parameter_belonging_to
    """
    #首先处理技术参数模块  提取 参数类别  1.主要技术参数
    # type_text=packet_content[index-1]
    # start,end=re.search('[\u4e00-\u9fa5]+', type_text).span()
    # type_text=type_text[start:end]

    pattern_primary='\W?\d+\W*[\u4e00-\u9fa5]+(:|：)?'    

    pattern_secondary_1='\W?\d+\W+\d+\W*[\u4e00-\u9fa5]+(:|：).+'  # *1.1生产能力：≥1200t/h。

    pattern_secondary_2='\W?\d+\W+\d+\W*[\u4e00-\u9fa5]+'  #2.2 采煤机阀类件要有过滤器 

    pattern_third='\W?\d+(\W+\d+){2}\W*'  # 三级标题  1.22主要部件大修周期： 1.22.1 轮机大修周期
    pattern_fourth='\W?\d+(\W+\d+){3}\W*'  # 四级标题  1.22.3.1 
    # 当遇到下一个大标题或者  第二节时  结束
    pattern_end_1='\W*(一|二|三|四|五|六|七|八|九)\W+'  # 
    pattern_end_2='\W*第(一|二|三|四|五|六|七|八|九)节'

    extract_dict_list=[] # 所有参数列表

    end_index=len(packet_content)
    # 提取类别
    for i in range(index,end_index):
        
        text=packet_content[i]

        if re.match(pattern_primary,text): # 如果匹配到参数类别 1.主要技术参数
            if i<=end_index-2:             # 当不是倒数第二行时
                if re.match(pattern_secondary_1,packet_content[i+1]) or re.match(pattern_secondary_2,packet_content[i+1]):
                    start,end=re.search('[\u4e00-\u9fa5]+', text).span()
                    parameter_type=text[start:end]

            

        elif re.match(pattern_secondary_1,text): # 如果匹配到参数名：参数值
            start,end=re.search('[\u4e00-\u9fa5]+(:|：)', text).span()
            parameter_name=text[start:end-1]  # 去掉冒号
            #start,end=re.search('(:|：).*', text).span()           
            parameter_value=text[end:]

            extract_dict = {'parameter_belonging_to':parameter_belonging_to,'parameter_type':parameter_type,'parameter_name': parameter_name,'parameter_value': parameter_value, 'time':time,'packet_index':packet_index,'filename':filename}

            extract_dict_list.append(extract_dict)

            #  = {'parameter_type':parameter_type,'parameter_name': parameter_name,'parameter_value': parameter_value, 'time':time,'parameter_belonging_to':parameter_belonging_to}

        elif re.match(pattern_secondary_2,text): # 如果匹配到 2.2 采煤机阀类件要有过滤器  纯文字描述 则整个添加到参数名              
            start,end=re.search('\W?\d\W+\d+', text).span()

            parameter_name=text[end:]

            extract_dict = {'parameter_belonging_to':parameter_belonging_to,'parameter_type':parameter_type,'parameter_name': parameter_name,'parameter_value': '', 'time':time,'packet_index':packet_index,'filename':filename}

            extract_dict_list.append(extract_dict)

        elif re.match(pattern_third,text)  : # 如果匹配到三级标题  1.22主要部件大修周期： 1.22.1 轮机周期 去掉标号加入上一级的参数值里
            start,end=re.search(pattern_third, text).span()
            parameter_value=text[end:]
            extract_dict_list[-1]['parameter_value']=extract_dict_list[-1]['parameter_value']+"\n"+parameter_value

        elif re.match(pattern_fourth,text)  : # 如果匹配到四级标题  1.22.1主要部件大修周期： 1.22.1.1 轮机周期 去掉标号加入上一级的参数值里
            start,end=re.search(pattern_fourth, text).span()
            parameter_value=text[end:]
            extract_dict_list[-1]['parameter_value']=extract_dict_list[-1]['parameter_value']+"\n"+parameter_value
 
        elif re.match(pattern_end_1,text) or re.match(pattern_end_2,text):  # 当遇到下一个大标题或者  下一节时  结束
            break
        
        else:                                                 # 否则 则判断是上一条内容的后继内容，即第二行
            if extract_dict_list[-1]['parameter_value']:
                extract_dict_list[-1]['parameter_value']=extract_dict_list[-1]['parameter_value']+text
            else:
                extract_dict_list[-1]['parameter_name']=extract_dict_list[-1]['parameter_name']+text

    return extract_dict_list
 