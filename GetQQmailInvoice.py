# -*- coding: utf-8 -*-
"""
一个简单用于生产和学习的自动过程脚本，适合财务人员，记账代理以及个体户管理发票门类、类目
想法：将QQ邮箱的发票助手下载的发票再进行分类整理

## Created on Tue Sep 20 13:26:19 2022
>> 由头，QQ邮箱发票助手发票整理给财务代理使用，例如每月下载一次发票，发票数量多没能分类，需要分类好方便报销

## Updated on Wed Mar 15 10:27:33 2023
>> Update QQ邮箱发票助手里手动每页下载的zip包，放在一个位置，循环解压缩zip files

## Updated on March 31, 2023, 4:49:01 PM
>> 添加发票号，添加更多分类，通过读取内容精确判断分类

## Updated on June 1, 2023
>> 添加处理ofd格式，添加更多类目

@author: Autolordz
"""

import os,re,glob,sys,datetime,time,shutil
import pdfplumber

file_folder = 'D:\\发票'
t0 = time.time()

print('Start Move All'.center(30,'*'))

# list_tags_swap = dict([(value, key) for key, value in list_tags.items()])

#排除公司，一般归为生活百货, add the exclude tags in except_tag 
except_tag = '芯果科技|优货仓|麦德龙|京东|晶东|宜家|永旺|沃尔玛|有家实业|华润万家|超市'

#分类发票,根据自己日常发票添加修改, should be, left (Catalog) = right (Keywords)
list_tags = dict(餐饮 = '餐饮|饮食|食品|咖啡|快餐|便利店|饮料|酒家|酒楼|豪客来', # 文件名包括上述就不归类到生活百货
     物流 = '物流|货拉拉|运输|冷链|货运',
     交通 = '交通|地铁|滴滴|出行|摩拜|骑安',
     油费 = '汽油|柴油|加油',
     服饰 = '服装|饰品|唯品会|迅销|迪卡侬|热风|盖璞|飒拉',
     文娱 = '文化|电影|影视|娱乐',
     家具 = '家具|家私|桌|椅|柜',
     差旅 = '差旅|酒店|住宿|寄存|旅游|旅店|宾馆|旅行',
     教育服务 = '教育服务|培训|考试|报名|课程',
     数码设备 = '数码|设备|电脑|手机|麦克风|耳机|相机|USB|转换器|插座',
     文具 = '文具|打印纸|笔|文件夹|胶水|回形针|剪刀|纸刀|订书|书桌垫',
     物业管理 = '物业管理|租赁|保洁|垃圾费|租金|水费|电费|水.*电',
     经营 = '服务费|体育用品|轴承|包装',
     医疗 = '医疗|宠物|美容|护肤|自疗|理疗|健康',
     生活百货 = '生活|日用|家居|药房|%s'%except_tag
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

from OFD_Invoice import OFD_Invoice

class Sortern_Invoice():
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

    def make_file(self,file_path):
        try:
            # 按内容指定类目
            for contentcatalog in \
            ['油费','物业管理','数码设备','文具','家具','教育服务','差旅','交通','餐饮','医疗','生活百货','经营']: 
                # 将有先后顺序的类目添加这里
                if re.search(r'(%s)|$'%list_tags[contentcatalog], self.Content).group().strip():
                    break
                contentcatalog = ''
            # print(contentcatalog)
            # 优先判断发票内容，然后判断发票公司名字估计发票类目
            folder3 = contentcatalog if contentcatalog else (get_tag(self.Seller) if self.Seller else get_tag(file_path))
            # 合并文件夹路径
            folder4 = os.path.join(file_folder,self.Buyer,self.Folder2,folder3)
            print(folder4)
            # 生成文件夹
            os.makedirs(folder4,exist_ok=1)   
            # 合并成为发票文件名： XX公司-金额-时间-发票号.pdf 已有QQ邮箱命名的不重复命名
            fname = '%s-%s-%s-%s%s'%(
                self.Seller,
                self.Money1,
                self.datestr1,
                self.InvoiceCode,
                os.path.splitext(file_path)[1]
                )
            path2 = os.path.join(file_folder,folder4,fname)
            # Move the file into folder
            shutil.move(file_path,path2)
            print(path2)
        except Exception as e:
            print(e)

    def process_ofd_invoice(self,file_path):
            OI1 = OFD_Invoice()
            dd1 = OI1.process_ofd(file_path)
            # 获取发票号码
            self.InvoiceCode = dd1['发票号码']
            # 获取公司生成 level 1 子文件夹
            self.Buyer = dd1['Buyer']
            self.Seller = dd1['Seller']
            # 获取时间生成 level 2 子文件夹
            self.datestr1 = dd1['开票日期']
            self.Folder2 = datetime.datetime.strptime(self.datestr1, '%Y年%m月%d日').strftime('%Y %m月')
            # 获取金额     
            self.Money1 = dd1['合计金额']
            # 获取内容
            self.Content = dd1['Content']
            self.make_file(file_path)

    def process_pdf_invoice(self,file_path):
        try:
            folder1 = dt1 = firm1 = firm2 = t_page2 = ''
            with pdfplumber.open(file_path) as pdf:
                t1 = pdf.pages[0].extract_text()
                if len(pdf.pages) > 1:
                    # 有大于一页发票
                    t_page2 = pdf.pages[1].extract_text()
            
            # 获取发票内容，Regex 提取开票公司名，本公司名，时间
            reg1 = re.search(r'称\s*[:：]|$',t1).group()
            firms = re.findall(r'(?<=%s)\s?[\u4e00-\u9fa5()（）]+'%reg1, t1)
            firms = [x for x in firms if len(x)>4]
            firms = firms[:2]
            if len(firms) > 0:
                firm1 = firms[0].strip()
                if len(firms) > 1:
                    firm2 = firms[1].strip()
                    
            # 获取发票号码
            self.InvoiceCode = re.search(r'(?<=发票号码[:：]).*|$', t1).group()
                    
            # 获取公司生成 level 1 子文件夹
            self.Buyer = firm1 #[:6]
            self.Seller = firm2
            
            # 获取金额     
            money1 = re.search(r'(?<=小写).*|$', t1).group()
            self.Money1 = re.search(r'(?<=[¥￥]).*|$', money1).group().strip()
            
            self.datestr1 = re.search(r'\d{4}年\d{2}月\d{2}日|$', file_path).group()
            
            if not self.datestr1:
                self.datestr1 = re.search(r'(?<=日期[:：]).*日|$', t1).group().strip().replace(' ','')
            
            # 获取时间生成 level 2 子文件夹
            self.Folder2 = datetime.datetime.strptime(self.datestr1, '%Y年%m月%d日').strftime('%Y %m月')
            # 判断发票内容作为发票类目 level 3 作为文件夹， from '纳税人识别号' to next '纳税人识别号'
            idx_t1 = t1.index('纳税人识别号')
            idx_t2 = t1.index('纳税人识别号',idx_t1+1)
            
            self.Content = t1[idx_t1:idx_t2] if not t_page2 else t_page2[t_page2.index('普通发票代码'):] # 第一页判断发票内容是从‘纳税人识别号’之后, 有第二页就判断从‘普通发票代码’开始
            
            self.make_file(file_path)
        except Exception as e:
            print(e)

SI1 = Sortern_Invoice()

for file_path in glob.glob(os.path.join(file_folder,'*.ofd')):
    SI1.process_ofd_invoice(file_path)

for file_path in glob.glob(os.path.join(file_folder,'*.pdf')):
    SI1.process_pdf_invoice(file_path)

print(('End Move All time: %.3f'%(time.time() - t0)).center(30,'*'))


