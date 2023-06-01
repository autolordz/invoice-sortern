# -*- coding: utf-8 -*-
"""
Created on Wed Mar 15 10:27:33 2023

@author: Autolordz
"""

import os,re,zipfile,shutil
from pathlib import Path
import xmltodict


class OFD_Invoice():
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
    #处理完就删除ofd
    def process_ofd(self,zip_path):
        dd1 = self.parse_xml(self.unzip_file(zip_path))
        if self.target_path and os.path.exists(self.target_path):
            shutil.rmtree(self.target_path)
        return dd1 
    #ofd 需要读取 original_invoice.xml 和 OFD.xml 两个文件
    def parse_xml(self,list1):
        if not list1: return None
        Buyer = Seller = ''
        with open(list1[0], 'r', encoding='utf-8') as f:
            tr1 = xmltodict.parse(f.read())
            Buyer = tr1['eInvoice']['fp:Buyer']['fp:BuyerName']
            Seller = tr1['eInvoice']['fp:Seller']['fp:SellerName']
            Content = tr1['eInvoice']['fp:GoodsInfos']['fp:GoodsInfo']['fp:Item']
        with open(list1[1], 'r', encoding='utf-8') as f:
            tr1 = xmltodict.parse(f.read())
            dd1 = {}
            for row in tr1['ofd:OFD']['ofd:DocBody']['ofd:DocInfo']['ofd:CustomDatas']['ofd:CustomData']:
                dd1[row['@Name']] = row.get('#text')
            dd1['Buyer'] = Buyer; dd1['Seller'] = Seller; dd1['Content'] = Content
        # print(dd1)
        return dd1
    #先用unzip解压ofd文件
    def unzip_file(self,zip_path):
        if zip_path.endswith('.zip') or zip_path.endswith('.ofd'):
            try:
                list1 = []
                self.target_path = zip_path.split('.')[0]
                with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                    for per_file_name in zip_ref.namelist():
                        target_file = os.path.join(self.target_path,per_file_name)
                        if 'OFD.xml' in per_file_name or 'original_invoice.xml' in per_file_name:
                            if not os.path.exists(target_file):
                                zip_ref.extract(per_file_name, self.target_path)
                                print(f'extra file: {target_file}')
                            else:
                                print(f'exist file: {target_file}')
                            list1 += [target_file]
                    # zip_ref.extractall(target_path)
                    # print(f'extra path: {target_path}')
            except Exception as e:
                print(f'extra error: {e}')
            return list1

if __name__ == '__main__':
    
    OI1 = OFD_Invoice()
    zip_path = r'D:\xx.ofd'
    dd1 = OI1.process_ofd(zip_path)
    
    print('done')
