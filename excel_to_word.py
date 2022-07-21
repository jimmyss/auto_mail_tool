#!/usr/bin/python
# -*- coding: UTF-8 -*-
from openpyxl import load_workbook
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
import smtplib
from email.mime.text import MIMEText
from email.header import Header, make_header
from email.mime.multipart import MIMEMultipart
from os import system
from re import match

def replace(obj):
    if obj is None:
        obj = ''
        return obj


with open('./setting.txt', 'r', encoding='UTF-8') as file:
    #读取hr邮箱
    data = file.readline()
    if match(r"hr邮箱:",data[:5]):
        mail_sender=data[5:-1]
    else:
        print("setting.txt文件中，hr邮箱格式错误，应是：'hr邮箱:xxx@yyy.com'")
        system("pause")
        exit(-1)

    #读取密码
    data=file.readline()
    if match(r"hr邮箱密码:", data[:7]):
        mail_license =data[7:-1]
    else:
        print("setting.txt文件中，hr邮箱密码格式错误，应是：'hr邮箱密码:*******'")
        system("pause")
        exit(-1)

    #读取发送人姓名
    data=file.readline()
    if match(r"发送人姓名:", data[:6]):
        sender_name=data[6:-1]
    else:
        print("setting.txt文件中，发送人姓名格式错误，应是：'发送人姓名:XXX'")
        system("pause")
        exit(-1)
    #读取邮件主题
    data=file.readline()
    if match(r"邮件主题:", data[:5]):
        mail_subject=data[5:-1]
        if mail_subject=="":
            print("邮件主题不能为空，请在setting中填写")
            system("pause")
            exit(-1)
    else:
        print("setting.txt文件中，邮件主题格式错误，应是：'邮件主题:XXXXX'")
        system("pause")
        exit(-1)
    #读取邮件正文
    data = file.readline()
    if match(r"邮件正文内容:", data[:7]):
        mail_content=data[7:-1]
    else:
        print("setting.txt文件中，邮件正文内容格式错误，应是：'邮件正文内容:XXXXX'")
        system("pause")
        exit(-1)
    #读取xlsx文件名
    data=file.readline()
    if match(r"待发送excel文件全名:", data[:13]):
        xlsx_file_name=data[13:-1]
    else:
        print("setting.txt文件中，待发送excel文件全名格式错误，应是：'待发送excel文件全名:aaa.xlsx'")
        system("pause")
        exit(-1)
    #读取docx文件名
    data=file.readline()
    if match(r"word模板文件全名:", data[:11]):
        docx_file_name=data[11:]
    else:
        print("setting.txt文件中，word模板文件全名格式错误，应是：'word模板文件全名:bbb.docx'")
        system("pause")
        exit(-1)

try:
    wb = load_workbook("./"+xlsx_file_name)  # 用openpyxl加载excel表格
    ws = wb['Sheet1']  # 读取Sheet1内容
    contexts = []  # 保存内容
    mail_sendto_list=[]#保存收件人列表
    mail_copy_list=[]#保存抄送人列表
    for row in range(2, ws.max_row + 1):
        name = ws["A" + str(row)].value  # 保存每一行的“姓名”列
        department = ws["B" + str(row)].value  # 保存每一行的“部门”列
        assessor = ws["C" + str(row)].value  # 保存每一行的“考核人”列
        confirmation_date = ws["D" + str(row)].value  # 保存每一行的“转正日期”列
        assess_factor_2021 = ws["E" + str(row)].value  # 保存每一行的“考勤系数”列
        average_salary_2021 = ws["F" + str(row)].value  # 保存每一行的“2021年平均岗位工资”列
        grade_year = ws["G" + str(row)].value  # 保存每一行的“2021年年终绩效考核”列
        salary_season = ws["H" + str(row)].value  # 保存每一行的“2021年季度奖”列
        salary_year = ws["I" + str(row)].value  # 保存每一行的“2021年年终奖”列
        salary_total = salary_season + salary_year  # 保存每一行的“2021年累计总奖金”列
        mail_to = ws["K" + str(row)].value  # 保存每一行的“邮箱/员工”列
        mail_copy = ws["L" + str(row)].value  # 保存每一行的“邮箱/直接主管”列
        # -----------将每行信息保存导字典context中-------------
        context = {"name": name, "department": department, "assessor": assessor, "confirmation_date": confirmation_date,
                   "assess_factor_2021": assess_factor_2021,
                   "average_salary_2021": average_salary_2021, "grade_year": grade_year, "salary_season": salary_season,
                   "salary_year": salary_year, "salary_total": salary_total,"mail_to":mail_to,"mail_copy":mail_copy}

        # -----------将个字典保存到数组context中--------------
        contexts.append(context)
    contexts
except Exception as n:
    print("无法打开或读取excel，请检查excel文件名和setting中待发送excel文件全名是否一致")
    system("pause")
    exit(-1)

#-----将excel提取到的数据填入表格------
try:
    for context in contexts:
        # 打开待操作的word文件，r表示“防止\转义”
        document = Document(r"./"+docx_file_name)

        # 提取文件中的表格
        tables = document.tables

        # -------将内容填入word里的表格-------
        tables[0].cell(1, 0).text = context['name']
        tables[0].cell(1, 1).text = context['department']
        tables[0].cell(1, 2).text = context['assessor']
        tables[0].cell(1, 3).text = context['confirmation_date']
        tables[0].cell(1, 4).text = str(context['assess_factor_2021'])
        tables[0].cell(1, 5).text = str(context['average_salary_2021'])
        tables[0].cell(1, 6).text = context['grade_year']
        tables[0].cell(1, 7).text = str(context['salary_season'])
        tables[0].cell(1, 8).text = str(context['salary_year'])
        tables[0].cell(1, 9).text = str(context['salary_total'])

        # -------表格内容居中展示-----------
        tables[0].cell(1, 0).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 1).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 2).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 3).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 4).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 5).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 6).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 7).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 8).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER
        tables[0].cell(1, 9).paragraphs[0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER

        # -------替换word文档中的姓名-------
        for paragraph in document.paragraphs:
            tmp = ''
            runs = paragraph.runs
            for i, run in enumerate(runs):
                tmp += run.text  # 合并run字符串
                if 'XXX' in tmp:
                    # 如果存在匹配得字符串，那么将当前得run替换成合并后得字符串
                    run.text = run.text.replace(run.text, tmp)
                    run.text = run.text.replace('XXX', context['name'])
                    tmp = ''
                else:
                    # 如果没匹配到目标字符串则把当前run置空
                    run.text = run.text.replace(run.text, '')
                if i == len(runs) - 1:
                    # 如果是当前段落一直没有符合规则得字符串直接将当前run替换为tmp
                    run.text = run.text.replace(run.text, tmp)

        # -------保存至新文件，文件名自拟-----
        document.save(r'./' + context['name'] + '-通知.docx')
except Exception as n:
    print("生成word失败，请检查word文件名与setting中word模板文件全名是否一致")
    system("pause")
    exit(-1)

# SMTP服务器地址
mail_host = "mail.gbcom.com.cn"
# 发件人邮箱,可能是hr邮箱********需要自己设置*******
#mail_sender = "dushimao@gbcom.com.cn"
#邮箱密码**********************需要自己设置******
#mail_license = "Qewxs132!#@"
# for context in contexts:
#     mail_receivers.append(context['mail_to'])
for context in contexts:
    # 设置收件人邮箱和抄送人邮箱，可以为多个收件人
    mail_receiver = []
    mail_receiver.append(context['mail_to'])
    mail_cc_receiver=[]
    mail_cc_receiver.append(context['mail_copy'])
    mail_receive=mail_receiver+mail_cc_receiver
    mm = MIMEMultipart('related')
    # 邮件主题********根据需要填写**********
    subject_content = mail_subject#"""自动发邮件测试"""
    # 设置发送者,这里写发送人姓名即可******根据需要填写*******
    mm["From"] = sender_name#"杜世茂"
    # 设置接受者,这里写收件人的姓名即可
    mm["To"] = context['name']
    #设置抄送人，这里写抄送人姓名即可
    mm["Cc"]= context['assessor']
    # 设置邮件主题
    mm["Subject"] = Header(subject_content, 'utf-8')

    # 邮件正文内容************根据需要填写************
    body_content = mail_content #"""你好，这是一个自动发送的邮件！"""
    # 构造文本,参数1：正文内容，参数2：文本格式，参数3：编码方式
    message_text = MIMEText(body_content, "plain", "utf-8")
    # 向MIMEMultipart对象中添加文本对象
    mm.attach(message_text)
    # 添加附件
    filename=context['name']+"-通知.docx"#设置附件名，要和生成的word名一致
    sendFile = open("./"+filename, 'rb').read()
    att = MIMEText(sendFile, "base64", "utf-8")
    att.add_header("Content-Type", "application/octet-stream")
    att.add_header("Content-Disposition", "attachment",
                   filename="%s" % make_header([(filename, 'UTF-8')]).encode('UTF-8'))
    mm.attach(att)

    try:
        # 创建SMTP对象
        stp = smtplib.SMTP()
        print("成功创建对象")
        # 设置发件人邮箱的域名和端口，端口地址为25
        stp.connect(mail_host, 25)
        # 登录邮箱，传递参数1：邮箱地址，参数2：邮箱密码
        stp.login(mail_sender, mail_license)
        # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
        stp.sendmail(mail_sender, mail_receive, mm.as_string())
        print("邮件发送成功")
        # 关闭SMTP对象
        stp.quit()
    except Exception as e:
        print("发送"+context['name']+"的邮件出错，可能是在setting中hr邮箱或者密码写错了，请修改setting后再次运行程序")
        system("pause")
        exit(-1)