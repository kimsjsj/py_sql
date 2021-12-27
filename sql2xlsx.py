# -*- coding: utf-8 -*-

import win32com.client as win32
import cx_Oracle
import os
import pandas as pd
from datetime import datetime, timedelta
import pythoncom
import time
import schedule
import pause


LOCATION = "오라클인스턴스 위치"
os.environ["PATH"] = LOCATION + ";" + os.environ["PATH"]
#os.enviro["NLS_LANG"] = "AMERICAN_AMERICA.AL32UTF8

base_dir = "C:\\Users\\ "  ##엑셀파일저장위치
mail_body = """
            <span style='font-size:12px;font-family:"Malgun Gothic";'>
            첨부참조
            </span>
            """

##메일 참조 지정
mail_CC = "sejin0412@bccard.com"
dw_address = cx_Oracle.makedsn("호스트","포트","이름")  ##DW접속

def data_ex_mail(dw_,file_nm,mail_To):   ##쿼리,레이아웃,엑셀파일명,받는사람 지정
    dw = cx_Oracle.connect("아이디/패스워드@호스트:포트/이름")  ##사용자DW계정
    cursor = dw.cursor()
    cursor.execute(dw_[0])
    x = cursor.fetchall()
    df = pd.DataFrame(x)
    if len(df) > 0:
        df.columns = dw_[1]
        cursor.close()
        dw.close()
        xlsx_dir = os.path.join(base_dir, file_nm+".xlsx")
        df.to_excel(xlsx_dir,
                    na_rep = 'NaN',
                    header = True,
                    startrow = 0,
                    startcol = 0)
        attach_file = "추출파일위치”+file_nm+".xlsx"
        olMailItem = 0x0
        pythoncom.CoInitialize()
        obj = win32.Dispatch("Outlook.Application")
        newMail = obj.CreateItem(olMailItem)
        now = datetime.now()
        newMail.Subject = "제목 _%s-%s-%s" % (now.year, now.month, now.day)  #메일제목
        newMail.HTMLBody = mail_body  ##메일본문
        newMail.Attachments.Add(attach_file)  ##첨부파일
        newMail.To = mail_To  ##보낼사람
        newMail.CC = mail_CC  ##참조
        newMail.Send()
        print(datetime.now(),"작업완료되었습니다.")
    else:
        cursor.close()
        dw.close()
        print(datetime.now(),"자료가없습니다.")
       
layout_newreg_mer = ["]
query_newreg_mer = """ “””
dw_newreg_mer = [query_newreg_mer, layout_newreg_mer]

layout_grpt_mer = ["”]
query_grpt_mer = """ “””
dw_grpt_mer = [query_grpt_mer, layout_grpt_mer]

while True:
    if datetime.now().hour != 8:
        tomorrow = (datetime.today()+timedelta(1)).strftime("%Y%m%d")
        pause.until(datetime(int(tomorrow[:4]),int(tomorrow[4:6]),int(tomorrow[6:]),8))
        continue
    data_ex_mail(dw_newreg_mer,"제목","@bcnuri.com")
    time.sleep(10)
    data_ex_mail(dw_grpt_mer,"제목","@bcnuri.com")
    time.sleep(3600)
