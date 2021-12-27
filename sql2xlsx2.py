# -*- coding: utf-8 -*-

import time
import pause
import cx_Oracle
import os
import pandas as pd
import win32com.client as win32
import pyautogui
import pythoncom

## sql 환경설정
LOCATION = "오라클인스턴스 위치"
os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]
#os.enviro["NLS_LANG"] = "AMERICAN_AMERICA.AL32UTF8
dw_address = cx_Oracle.makedsn("호스트","포트","이름")  ##DW접속

### 파일명 입력
std = pyautogui.prompt(title='기간입력',default='ex. 202110', text='작업대상년월 입력 YYYYMM')
if std[4:6] == '01':
    std_bf = str(int(std[0:4])-1)+'12'
else:
    std_bf = str(int(std)-1)
std_ym = ["'"+std_bf+"'","'"+std+"'"]

## 정기자료 파일설정         
base_dir = "저장위치"
xlsx_dir_card = os.path.join(base_dir, "vald_card_"+std+".xlsx")
xlsx_dir_cdhd = os.path.join(base_dir, "vald_cdhd_"+std+".xlsx")
xlsx_dir_new = os.path.join(base_dir, "new_card_"+std+".xlsx")
xlsx_dir_sale = os.path.join(base_dir, "sale_amt_"+std+".xlsx")

##사용자DW계정
dw = cx_Oracle.connect("계정/패스워드@호스트:포트/이름") 
cursor = dw.cursor()

## 쿼리
query_card = """
"""

query_new = """
"""

query_cdhd = """
"""

query_sale = """
"""

##쿼리 실행
cursor.execute(query_card)
card = cursor.fetchall()
cursor.execute(query_cdhd)
cdhd = cursor.fetchall()
cursor.execute(query_new)
new = cursor.fetchall()
cursor.execute(query_sale)
sale = cursor.fetchall()
vald_card = pd.DataFrame(card)
vald_cdhd = pd.DataFrame(cdhd)
new_card = pd.DataFrame(new)
sale_amt = pd.DataFrame(sale)

cursor.close()
dw.close()

##엑셀파일 저장
vald_card.to_excel(xlsx_dir_card, na_rep = 'NaN', startrow = 0, startcol = 0)
vald_cdhd.to_excel(xlsx_dir_cdhd, na_rep = 'NaN', startrow = 0, startcol = 0)
sale_amt.to_excel(xlsx_dir_sale, na_rep = 'NaN', startrow = 0, startcol = 0)
new_card.to_excel(xlsx_dir_new, na_rep = 'NaN', startrow = 0, startcol = 0)

pyautogui.alert(title='작업완료',text='월실적자료 폴더에 저장되었습니다.')
