# -*- coding: utf-8 -*-
# @Time : 2022-05-10
# @Author : dwd
# @File : 考勤自动提醒.py

#功能1：通过人力系统爬取考勤数据，excel表。
#----------------------------------------------------------------
#功能2：根据爬取下来的excel表，判断是否缺勤。
import pandas as pd
import numpy as np
import logging
def check_Attendance():
    logging.basicConfig(level=logging.INFO, format="时间：%(asctime)s - 日志等级：%(levelname)s - 日志信息：%(message)s")
    global data_list
    df = pd.read_excel(import_file_path, sheet_name='sheet1')  # 读取excel表
    Abnormal_attendance = df['考勤情况'] == '异常考勤'  #统计异常考勤
    attendance = df[Abnormal_attendance]
    # 首先将pandas读取的数据转化为array
    data_array = np.array(attendance)
    # 将异常考勤转换成list形式
    data_list = data_array.tolist()
    #将异常考勤转换成str
    data_list  = ",".join(list(map(str, data_list)))
    logging.info(data_list )
    writer = pd.ExcelWriter(export_file_path)
    attendance.to_excel(writer, sheet_name="异常统计", index=False)
    writer.save()
#----------------------------------------------------------------
#功能3：结合综合员考勤表，校准缺勤记录
#----------------------------------------------------------------
#功能4.1：将结果通过邮件形式发送,授权邮箱
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import datetime
import logging
def send_mail():
    logging.basicConfig(level=logging.INFO, format="时间：%(asctime)s - 日志等级：%(levelname)s - 日志信息：%(message)s")
    #输入发件人邮箱名称
    email_name = '304382089@qq.com'
    # 输入用户授权码
    passwd = 'jfldkhtbmfjzbifb'
    # 收件人邮箱
    msg_to = send_email_address
    # 邮件的正文

    content = data_list #正文取值为跑出的异常考勤结果
    # 设置邮件
    content_part = MIMEText(content)
    #msg = MIMEText(content)
    msg = MIMEMultipart()
    msg['subject'] = f'{datetime.datetime.now().date()}考勤提醒'
    #logging.info()
    #设置发件人
    msg['From']= '304382089@qq.com'
    # 这个参数设置要发给谁

    msg['To'] = ','.join(msg_to)
    # 添加附件内容
    msg.attach(content_part)
    # *********************构造附件***********
    # 文本类型的附件
    att1 = MIMEText(open("考勤分析表.xlsx", 'rb').read(), 'plain', 'utf-8')

    # 添加头信息，我告诉服务器，我现在是一个附件
    att1['Content-Type'] = 'application/octet-stream'
    att1.add_header("Content-Disposition", 'attachment', filename=('gbk', "", '考勤分析表.xlsx'))
    # 把内容添加到邮件中
    msg.attach(att1)

    #连接服务器
    s= smtplib.SMTP_SSL('smtp.qq.com',465)
    # 登陆我的邮箱
    s.login(email_name,passwd)
    # 发送邮箱
    s.sendmail(email_name,msg_to,msg.as_string())
    print("发送成功")
#功能4.2：将结果通过邮件形式发送,非授权邮箱，outlook邮箱

#----------------------------------------------------------------
#功能5：设置定时任务，定时爬取，定时跑批，定时发邮件


if __name__ == '__main__':
    send_email_address = ['304382089@qq.com']
    import_file_path = r'D:\自动化办公\考勤情况 (20220510).xls'
    export_file_path = r'D:\自动化办公\考勤分析表.xlsx'
    check_Attendance()
    send_mail()