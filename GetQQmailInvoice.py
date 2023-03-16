# -*- coding: utf-8 -*-
"""
Created on Tue Sep 20 13:26:19 2022
>> 由头，QQ邮箱发票助手发票整理给财务代理使用，例如每月下载一次发票，发票数量多没能分类，需要分类好方便报销

Updated on Wed Mar 15 10:27:33 2023
>> Update QQ邮箱发票助手里手动每页下载的zip包，放在一个位置，循环解压缩Zip 

@author: Autolords
"""

import os,re,zipfile
from pathlib import Path

source_folder = target_folder = 'D:\\发票'

# unzip all the QQ invoices zip files, for example zips of one month
for zipfilename in os.listdir(source_folder):
    if zipfilename.endswith('.zip'):
        try:
            with zipfile.ZipFile(os.path.join(source_folder, zipfilename), 'r') as zip_ref:
                for per_file_name in zip_ref.namelist():
                    per_file_path = os.path.join(target_folder,per_file_name.encode('cp437').decode('gbk'))
                    if not os.path.exists(per_file_path):
                        zip_ref.extract(per_file_name, target_folder)
                # zip_ref.extractall(target_folder)
            print(f'extra file: {zipfilename}')
        except Exception as e:
            print(f'extra error: {e}')

# rename all the pdf file with 乱码 to 中文, need transfer cp437 to gbk
for zipfilename in os.listdir(source_folder):
    if zipfilename.endswith('.pdf'):
        per_file_path = Path(os.path.join(source_folder, zipfilename))
        try:
            path2 = per_file_path.rename(os.path.join(source_folder, zipfilename.encode('cp437').decode('gbk')))
            print(f'{per_file_path} ----> {path2}')
        except Exception:
            pass
        # break
    
print('done')

#%%

import os,re,glob,sys,datetime,time,shutil
import pdfplumber

file_folder = 'D:\\发票'
t0 = time.time()

print('Start Move All'.center(30,'*'))

#排除公司, add the exclude tags in except_tag
except_tag = '芯果科技|优货仓|麦德龙|京东|晶东|宜家|永旺|沃尔玛'

#分类发票,根据自己日常发票添加修改, should be, left (Catalog) = right (Keywords)
list_tags = dict(餐饮 = '餐饮|饮食|食品|咖啡|快餐|便利店|饮料|酒家|酒楼|中粮',
     交通 = '交通|地铁|货拉拉|运输',
     服饰 = '唯品会|迅销|迪卡侬|热风|盖璞',
     文娱 = '文化|电影|影视',
     物业管理 = '物业管理|保洁',
     经营 = '体育用品|轴承|包装',
     生活百货 = '药房|%s'%except_tag
     )

# Default return Catalog will be '生活百货', if nothing to find then '其他'
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
        
        # 获取发票内容，Regex 提取开票公司名，本公司名，时间
        reg1 = re.search(r'称\s*[:：]|$',t1).group()
        firms = re.findall(r'(?<=%s)\s?[\u4e00-\u9fa5()（）]+'%reg1, t1)
        firms = [x for x in firms if len(x)>4]
        firms = firms[:2]
        if len(firms) > 0:
            firm1 = firms[0].strip()
            if len(firms) > 1:
                firm2 = firms[1].strip()
                
        # 获取公司生成 level 1 子文件夹
        folder1 = firm1.strip()[:6]
        
        # 获取金额     
        money1 = re.search(r'(?<=小写).*|$', t1).group()
        money1 = re.search(r'(?<=[¥￥]).*|$', money1).group().strip()
        
        ds1 = re.search(r'\d{4}年\d{2}月\d{2}日|$', file_path).group()
        
        if not ds1:
            ds1 = re.search(r'(?<=日期[:：]).*日|$', t1).group().strip().replace(' ','')
        
        # 获取时间生成 level 2 子文件夹
        folder2 = datetime.datetime.strptime(ds1, '%Y年%m月%d日').strftime('%Y %m月')
        # 获取发票类型 level 3 作为文件夹
        folder3 = get_tag(firm2) if firm2 else get_tag(file_path)
        
        folder4 = os.path.join(file_folder,folder1,folder2,folder3)
        print(folder4)
        
        # 生成文件夹
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


    


