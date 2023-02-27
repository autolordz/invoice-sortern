# -*- coding: utf-8 -*-
"""
Created on Tue Sep 20 13:26:19 2022

@author: Autolords
"""
import os,re,glob,sys,datetime,time,shutil
import pdfplumber

file_folder = 'D:\\发票'
t0 = time.time()

print('Start Move All'.center(30,'*'))

#排除公司
except_tag = '芯果科技|优货仓|麦德龙|京东|晶东|宜家|永旺|沃尔玛'

#分类发票
list_tags = dict(餐饮 = '餐饮|饮食|食品|咖啡|快餐|便利店|饮料|酒家|酒楼|中粮',
     交通 = '交通|地铁|货拉拉|运输',
     服饰 = '唯品会|迅销|迪卡侬|热风|盖璞',
     文娱 = '文化|电影|影视',
     物业管理 = '物业管理|保洁',
     经营 = '体育用品|轴承|包装',
     生活百货 = '药房|%s'%except_tag
     )

def get_tag(file_path):
    for folder_tag in list_tags.keys():
        if re.search(r'(%s)|$'%list_tags['生活百货'], file_path).group().strip():
            return '生活百货'
        elif not re.search(r'(%s)|$'%except_tag, file_path).group().strip() \
        and re.search(r'(%s)|$'%list_tags[folder_tag], file_path).group().strip():
            return folder_tag
    return '其他'


for file_path in glob.glob(os.path.join(file_folder,'*.pdf')):
    try:
        folder1 = dt1 = firm1 = firm2 = ''
        with pdfplumber.open(file_path) as pdf:
            p1 = pdf.pages[0] 
            t1 = p1.extract_text()
        
        # 获取发票内容，提取开票公司名，本公司名，时间
        reg1 = re.search(r'称\s*[:：]|$',t1).group()
        firms = re.findall(r'(?<=%s)\s?[\u4e00-\u9fa5()（）]+'%reg1, t1)
        firms = [x for x in firms if len(x)>4]
        firms = firms[:2]
        if len(firms) > 0:
            firm1 = firms[0].strip()
            if len(firms) > 1:
                firm2 = firms[1].strip()
                
        # 获取公司作为 level 1 子文件夹
        folder1 = firm1.strip()[:6]
        
        money1 = re.search(r'(?<=小写).*|$', t1).group()
        money1 = re.search(r'(?<=[¥￥]).*|$', money1).group().strip()
        
        ds1 = re.search(r'\d{4}年\d{2}月\d{2}日|$', file_path).group()
        
        if not ds1:
            ds1 = re.search(r'(?<=日期[:：]).*日|$', t1).group().strip().replace(' ','')
        
        # 获取时间作为 level 2 子文件夹
        folder2 = datetime.datetime.strptime(ds1, '%Y年%m月%d日').strftime('%Y %m月')
        # 获取发票类型 level 3 作为文件夹
        folder3 = get_tag(firm2) if firm2 else get_tag(file_path)
        
        folder4 = os.path.join(file_folder,folder1,folder2,folder3)
        print(folder4)

        os.makedirs(folder4,exist_ok=1)
        
        # 合并成为发票文件名： XX公司-金额-时间.pdf
        fname = '%s-%s-%s%s'%(
            firm2,
            money1,
            ds1,
            os.path.splitext(file_path)[1]
            ) if not re.search(r'-', file_path) else os.path.basename(file_path)
        
        path2 = os.path.join(file_folder,folder4,fname)
        
        shutil.move(file_path,path2)
        print(path2)
        
    except Exception as e:
        print(e)

print(('End Move All time: %.3f'%(time.time() - t0)).center(30,'*'))


    


