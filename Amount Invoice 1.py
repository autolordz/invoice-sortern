# -*- coding: utf-8 -*-
"""
Created on Wed Mar  1 16:39:39 2023

@author: Autozhz

统计每个 level 1 sub folder的发票金额，写入 amount.txt
"""

from pathlib import Path

catalog_path = Path("D:\发票")

def sum_catalog(catalog):
    """Summarize the amount of money in each catalog"""
    money_list = [float(f.name.split("-")[1]) for f in catalog.iterdir() if f.is_file() and "-" in f.name]
    amount = sum(money_list)
    ret = f"{catalog} sum is {amount:.2f}"
    print(ret)
    return ret

with open(catalog_path / "amount.txt", "w+") as file:
    summary = [sum_catalog(sub_catalog) for date in catalog_path.iterdir() if date.is_dir() for sub_catalog in date.iterdir() if sub_catalog.is_dir()]
    file.write("\n".join(summary))