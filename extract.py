from extract_utils import *
import json
import csv 

def process_chapters(source_path):
    """处理源文件，提取技术章节"""
    print('开始提取技术大章节')

    all_files_chapters_contents={}
    all_files_time={}
    all_files_crop_chapters_flag={}

    for _,_,filenames in os.walk(source_path):
        for filename in filenames:
            print(filename)
            if filename.split(".")[-1] == 'pdf':
    #                 continue
                #handle_pdf(source_path,filename)
                pdf_obj = PyPDF2.PdfFileReader(source_path+filename)
                doc_info = pdf_obj.documentInfo
                
                time,crop_chapters_flag,chapters_contents=crop_chapters(pdf_obj,source_path,filename)
            
                # with open(filename.split(".")[0]+'.txt',"w",encoding='utf-8') as out_file:
                #     for contents in chapters_contents:
                #         out_file.write(contents)    
                #         out_file.write('\n')

                all_files_chapters_contents[filename]=chapters_contents
                all_files_time[filename]=time
                all_files_crop_chapters_flag[filename]=crop_chapters_flag
            else:
                print('file type must be pdf')

    return all_files_chapters_contents,all_files_time,all_files_crop_chapters_flag
 

def process_packets(all_files_section_contents,object_type):
    """处理多包文件"""
    print('开始提取多个标包')
    all_files_packets_contents={}
    all_files_crop_packets_flag={}
    total_packet_counts=0
    valid_packet_counts=0

    for filename,section_contents in all_files_section_contents.items():
        print(filename)
        
        packets_contents=[]

        total_packet_count,valid_packet_count,packets_contents,crop_packets_flag=crop_packets(section_contents,object_type)

        total_packet_counts+=total_packet_count
        valid_packet_counts+=valid_packet_count
        # with open(filename.split(".")[0]+'.txt',"w",encoding='utf-8') as out_file:
        #     for contents in packets_contents:
        #         out_file.write('\n'.join(contents))    
        #         out_file.write('\n')

        all_files_packets_contents[filename]=packets_contents

        all_files_crop_packets_flag[filename]=crop_packets_flag

    return total_packet_counts,valid_packet_counts,all_files_packets_contents,all_files_crop_packets_flag

def final_process(all_files_packets_contents,all_files_time,object_type,write_path):
    print('开始提取信息')
    
    #pattern_primary='\W?\d+\W*[\u4e00-\u9fa5]*(技术|参数|要求)[\u4e00-\u9fa5]*(:|：)?'
    #pattern_secondary='\W?\d+\W+\d+\W*[\u4e00-\u9fa5]+(:|：)?' 
 
    pattern_primary_1='\W*(一|二|三|四|五|六|七|八|九)\W*.*(技术|参数).*' 
    pattern_primary_2='\W*第(一|二|三|四|五|六|七|八|九)部分\W*.*(技术|参数).*' 
    pattern_primary_3='\W+\d\W+[\u4e00-\u9fa5]*.*(技术|参数).*' 

    pattern_secondary='\W*(\d\W*)+[\u4e00-\u9fa5]*.*(技术|参数|要求).*' 
 
    #二、采煤机技术参数及要求：

    extract_dict_lists=[]

    all_files_extract_flag={} #

    extract_mark_text={} # 定位标记文本 常为 1.采煤机技术要求  1.1 生产能力:>10t 过煤量

    for filename,packet_contents in all_files_packets_contents.items():
        print(filename)
        time=all_files_time[filename]
        parameter_belonging_to=object_type
        all_files_extract_flag[filename]={}
        extract_mark_text[filename]={}

        for packet_index,packet_content in enumerate(packet_contents):
            packet_extract_flag='extract failed'
            mark_text=''
            if packet_content!=[]:
                for index,line in enumerate(packet_content) :
                    #if index <=len(packet_content)-3:  # 不是倒数第二行
                        if re.match(pattern_primary_1,line) or re.match(pattern_primary_2,line) or re.match(pattern_primary_3,line): # 匹配到 一、 技术参数
  
                            if len(packet_content)-index >=3:#判断向下搜索多少行
                                for count in range(1,2):
                                    if re.match(pattern_secondary,packet_content[index+count]): # 匹配到 1.1. 参数名：参数值
 
                                        mark_text=line+'\n'+packet_content[index+count]
                                        packet_extract_flag='succeed'
                                        extract_dict_list=final_extract(index+count,packet_index,packet_content,filename,time,parameter_belonging_to)
                                        extract_dict_lists.extend(extract_dict_list)
                                        break # 如果打断 就结束内外双层循环
                                else:
                                    continue
                                break
                             
            # 遍历完没有发现标记文本
            #if packet_extract_flag=='extract failed':


            all_files_extract_flag[filename][packet_index]=packet_extract_flag
            extract_mark_text[filename][packet_index]=mark_text


    with open('{}/extract_dict_lists.json'.format(write_path), mode='w', encoding='utf-8') as f:
        json.dump(extract_dict_lists, f,ensure_ascii=False)

    with open("{}/extract_dict_lists.csv".format(write_path), mode='w', encoding='gbk') as f:
        try:  # 有可能文件无内容
            headers =extract_dict_lists[0].keys() 
            writer = csv.writer(f)
            writer.writerow(headers)
            sheet_data=[]
            for data in extract_dict_lists:
                write_data=data.values()
                try:  # 有可能乱码
                    writer.writerow(write_data)
                except:
                    continue
        except:
            pass

 
    
    return extract_dict_lists,all_files_extract_flag,extract_mark_text

def report(write_path,total_packet_counts,valid_packet_counts,all_files_time,all_files_crop_chapters_flag,all_files_crop_packets_flag,all_files_extract_flag,extract_dict_lists_length,extract_mark_text):

    with open("{}/report.csv".format(write_path),"w",encoding='utf-8-sig') as out_file:
        out_file.write('filename,time information,crop chapters information,crop packets information,extract information,extract_mark_text\n')
        for filename,crop_chapters_flag in all_files_crop_chapters_flag.items():
            out_file.write('{},{},{},{},{},{}\n'.format(filename,all_files_time[filename],crop_chapters_flag,all_files_crop_packets_flag[filename],all_files_extract_flag[filename],extract_mark_text[filename]))  
    
 
    succeed_packet_num=0
 
    for filename,packet_indexs in all_files_extract_flag.items():

        for packet_index in packet_indexs:

            if all_files_extract_flag[filename][packet_index]=='succeed':
                succeed_packet_num+=1

    result_report="此次提取总共有 {} 个文件， {} 个标包，其中有效目的标包 {} 个，提取成功 {} 个标包，成功率{}，总共提取 {} 条信息".format(len(all_files_time),total_packet_counts,valid_packet_counts,succeed_packet_num,round(succeed_packet_num/valid_packet_counts,3),extract_dict_lists_length)

    print(result_report)

    with open('{}/result report.txt'.format(write_path),'w') as f:
        f.write(result_report)


def main():
    # source_path='myjob/test_files/'
    # object_type='刮板输送机|刮板运输机'

    # source_path='myjob/source_files_采煤机/'
    # object_type='采煤机'
    
    # source_path='myjob/source_files_旋转器/'
    # object_type='旋流器|旋转器'

    # source_path='myjob/source_files_旋转器/'
    # object_type='旋流器'

    # source_path='myjob/all_source_files/各标的物历年招标文件20200423/3 掘进机/'
    # object_type='掘进机'

    source_path='myjob/all_source_files/各标的物历年招标文件20200423/5 工作面刮板输送机（含顺槽转载机、顺槽破碎机、自移机尾）/'
    object_type='三机|刮板输送机|刮板运输机'

    # source_path='myjob/all_source_files/各标的物历年招标文件20200423/12 井下防爆柴油机无轨胶轮车/'
    # object_type='防爆.*'

    # source_path='myjob/all_source_files/各标的物历年招标文件20200423/13 支架搬运车/'
    # object_type='支架搬运车'

    # source_path='myjob/all_source_files/各标的物历年招标文件20200423/14 矿用蓄电池电机车/'
    # object_type='蓄电池电机车'
    


    write_path= 'myjob/output/{}'.format(re.sub("\W",' ',object_type).strip())

    if not os.path.exists(write_path):
        os.makedirs(write_path)

    "章"
    all_files_chapters_contents,all_files_time,all_files_crop_chapters_flag=process_chapters(source_path)
    "包"
    total_packet_counts,valid_packet_counts,all_files_packets_contents,all_files_crop_packets_flag=process_packets(all_files_chapters_contents,object_type)
    "提取"
    extract_dict_lists,all_files_extract_flag,extract_mark_text=final_process(all_files_packets_contents,all_files_time,object_type,write_path)
    "报告"
    report(write_path,total_packet_counts,valid_packet_counts,all_files_time,all_files_crop_chapters_flag,all_files_crop_packets_flag,all_files_extract_flag,len(extract_dict_lists),extract_mark_text)

if __name__ == "__main__":

    main()

    # import re  
    # object_type='刮板输送机|刮板运输机({})'
    # print(re.sub("\W",' ',object_type))

    # 问题 一、技术参数及要求 
    # 类型一
    #三、技术要求
    # 3.1 整车参数
    # *3.1.1 额定承载人数： 20 人
    # 3.1.2 整车外形尺寸(mm)：(5700±

    # 类型二
    #三、技术参数及要求
    # 1. 技术参数
    #     （1）水环真空泵
    #         1.1 水环真空泵：

    # 指定标的物名称对标包进行筛选时  若标的物有多个名称则存在提取缺失的情况
    # 货物需求一览表的识别
