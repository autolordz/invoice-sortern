# invoice-sortern

一个简单用于生产和学习的自动过程脚本，适合财务人员，记账代理以及个体户管理发票门类、类别
想法：将QQ邮箱的发票助手下载的发票再进行分类整理

[![](https://img.shields.io/github/release/autolordz/invoice-sortern.svg?style=popout&logo=github&colorB=ff69b4)](https://github.com/autolordz/invoice-sortern/releases)
[![](https://img.shields.io/badge/github-source-orange.svg?style=popout&logo=github)](https://github.com/autolordz/invoice-sortern)
[![](https://img.shields.io/github/license/autolordz/invoice-sortern.svg?style=popout&logo=github)](https://github.com/autolordz/invoice-sortern/blob/master/LICENSE)

## Updated on June 1, 2023
>> 添加处理ofd格式，添加更多类目

## Updated on March 31, 2023, 4:49:01 PM
>> 添加发票号，添加更多分类，通过读取内容精确判断分类

## Updated on Wed Mar 15 10:27:33 2023
>> Update QQ邮箱发票助手里手动每页下载的zip包，放在一个位置，循环解压缩zip files

## Created on Tue Sep 20 13:26:19 2022
>> 由头，QQ邮箱发票助手发票整理给财务代理使用，例如每月下载一次发票，发票数量多没能分类，需要分类好方便报销

### 提示
文件名不需要一定按照QQ邮箱格式，会自动重命名 -> XX公司-金额-时间.pdf
为方便可以从QQ邮箱里全选一页下载

### INPUT
发票复制到要存放的文件夹路径
路径在代码的开头

### OUTPUT
同文件夹输出整理好发票

公司名folder \
├─日期folder1 \
│  ├─类别folder1 \
│  └─类别folder2 \
└─日期2 \
    ├─类别2 \
    │      文件1.pdf \
    │      文件2.pdf \
    └─类别3
    
Licence MIT
