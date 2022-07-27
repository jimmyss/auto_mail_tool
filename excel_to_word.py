#!/usr/bin/python
# -*- coding: UTF-8 -*-
from openpyxl import load_workbook
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from smtplib import SMTP
from email.mime.text import MIMEText
from email.header import Header, make_header
from email.mime.multipart import MIMEMultipart
from re import match
from tkinter import *
from tkinter import filedialog, dialog
import os

def set_text(text):
    message_box.config(state=NORMAL)
    message_box.delete("1.0","end")
    message_box.insert(END,'\n'+str(text))
    message_box.config(state=DISABLED)

def replace(obj):
    if obj is None:
        obj = ''
        return obj

def send_mail(mail_sender, mail_license, sender_name, mail_subject, mail_content, xlsx_file_name,docx_file_name):
    try:
        wb = load_workbook("./" + xlsx_file_name, data_only=True)  # 用openpyxl加载excel表格
        ws = wb.active  # 读取Sheet1内容
        mail_sendto_list = []  # 保存收件人列表
        mail_copy_list = []  # 保存抄送人列表
        col_names = []
        tuples = []
        row_max = ws.max_row  # 保存最大行数
        col_max = ws.max_column  # 保存最大列数
        col_copy = []  # 保存抄送人列号
        start_row = 0
        mail_col_position = 0
        end_flag = False
        for i in range(row_max):
            for j in range(col_max):
                if ws.cell(row=i + 1, column=j + 1).value:
                    # 识别关键字"邮箱"，确定读取开始行
                    if match(r"邮箱", str(ws.cell(row=i + 1, column=j + 1).value)):
                        start_row = i + 1
                        mail_col_position = j + 1
                        # 识别关键字“抄送人”，确定有几个抄送人
                        for k in range(j + 1, col_max):
                            if match(r"邮箱", ws.cell(row=i + 1, column=k + 1).value):  # 如果有邮箱的字眼，就默认为抄送人
                                col_copy.append(k + 1)
                        end_flag = True
                        break
            if end_flag:
                break
        for i in range(start_row, row_max + 1):
            if i == start_row:
                # 确定要读取的列名
                for j in range(col_max):
                    col_names.append((ws.cell(row=i, column=j + 1).value).replace("\n", ""))  # 去除列里面的换行符
            else:
                # 读取每一行
                if ws.cell(row=i, column=mail_col_position).value:
                    tuples.append(ws[i])  # 保存每行元组
                    # 填写收件人
                    mail_sendto_list.append(ws.cell(row=i, column=mail_col_position).value)
                    # 填写抄送人
                    temp_copy_mail = []
                    for j in range(len(col_copy)):
                        # 若有抄送人，将所有抄送人保存到一个list中并保存到mail_copy_list中
                        if ws.cell(row=i, column=col_copy[j]).value:
                            temp_copy_mail.append(ws.cell(row=i, column=col_copy[j]).value)
                    mail_copy_list.append(temp_copy_mail)
    except Exception as n:
        return "无法打开或读取excel，请检查excel文件名和setting中待发送excel文件全名是否一致"
        # print("无法打开或读取excel，请检查excel文件名和setting中待发送excel文件全名是否一致")
        # print(n)
        # system("pause")
        # exit(-1)


    # -----将excel提取到的数据填入表格------
    try:
        for context in tuples:
            try:
                # 打开待操作的docx文件，r表示“防止\转义”
                document = Document(r"./" + docx_file_name)
            except Exception as e:
                return "不能打开doc格式，请手动转换模板为docx再执行程序"
                # print("不能打开doc格式，请手动转换模板为docx再执行程序")
                # print(e)
                # system("pause")
                # exit(-1)
            # 提取文件中的表格
            tables = document.tables

            # 循环填写表格
            for table in tables:
                for i in range(0, len(table.rows)):
                    for j in range(0, len(table.columns)):
                        table_col_name = table.cell(i, j).text.replace("\n", "")
                        if table_col_name in col_names:  # 若检测到对应列
                            try:
                                if context[col_names.index(table_col_name)].value != None:
                                    # 若非空，将该列下面填入值
                                    table.cell(i + 1, j).text = str(context[col_names.index(table_col_name)].value)
                                    table.cell(i + 1, j).paragraphs[
                                        0].paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER  # 居中显示
                                else:
                                    # 若为空值，填入空白
                                    table.cell(i + 1, j).text = ""
                            except Exception as e:
                                return "word表格格式错误，应该每张表仅有两行"
                                # print("word表格格式错误，应该每张表仅有两行")
                                # print(e)
                                # system("pause")

            # -------替换word文档中的姓名/工号姓名-------
            if "姓名" in col_names:
                name = context[col_names.index("姓名")].value
            if "工号" in col_names:
                number = context[col_names.index("工号")].value

            start_flag = False  # 识别是否开始输入的
            for paragraph in document.paragraphs:
                for run in paragraph.runs:
                    if "XXX" in run.text:
                        run.text = run.text.replace('XXX', name)
                    if "YYY" in run.text:
                        run.text = run.text.replace('YYY', number)

            # -------保存至新文件，文件名自拟-----
            document.save(r'./' + name + "-" + docx_file_name)
    except Exception as n:
        return "生成word失败，请检查word文件名与setting中word模板文件全名是否一致"
        # print("生成word失败，请检查word文件名与setting中word模板文件全名是否一致")
        # print(n)
        # system("pause")
        # exit(-1)



    # SMTP服务器地址
    mail_host = "mail.gbcom.com.cn"
    count = 0  # 对收件人计数
    for context in tuples:  # 将每个文件发送
        # 设置收件人邮箱和抄送人邮箱，可以为多个收件人
        mail_receiver = []
        mail_receiver.append(mail_sendto_list[count])
        mail_cc_receiver = []
        for copyer in mail_copy_list[count]:
            mail_cc_receiver.append(copyer)
        mail_receive = mail_receiver + mail_cc_receiver
        mm = MIMEMultipart('related')
        # 邮件主题********根据需要填写**********
        subject_content = mail_subject  # """自动发邮件测试"""
        # 设置发送者,这里写发送人姓名即可******根据需要填写*******
        mm["From"] = sender_name  # "杜世茂"
        # 设置接受者,这里写收件人的姓名即可
        mm["To"] = context[col_names.index("姓名")].value
        # 设置抄送人，这里写抄送人姓名即可
        mm["Cc"] = ','.join(mail_cc_receiver)
        # 设置邮件主题
        mm["Subject"] = Header(subject_content, 'utf-8')

        # 邮件正文内容************根据需要填写************
        body_content = mail_content  # """你好，这是一个自动发送的邮件！"""
        # 构造文本,参数1：正文内容，参数2：文本格式，参数3：编码方式
        message_text = MIMEText(body_content, "plain", "utf-8")
        # 向MIMEMultipart对象中添加文本对象
        mm.attach(message_text)
        # 添加附件
        filename = context[col_names.index("姓名")].value + "-" + docx_file_name  # 设置附件名，要和生成的word名一致
        sendFile = open("./" + filename, 'rb').read()
        att = MIMEText(sendFile, "base64", "utf-8")
        att.add_header("Content-Type", "application/octet-stream")
        att.add_header("Content-Disposition", "attachment",
                       filename="%s" % make_header([(filename, 'UTF-8')]).encode('UTF-8'))
        mm.attach(att)
        count += 1
        result_string=""
        if 1:
            try:
                # 创建SMTP对象
                stp = SMTP()
                result_string+=("成功创建邮件："+str(count)+"/"+str(len(tuples))+"\n")
                #print("成功创建邮件："+str(count)+"/"+str(len(tuples)))
                # 设置发件人邮箱的域名和端口，端口地址为25
                stp.connect(mail_host, 25)
                # 登录邮箱，传递参数1：邮箱地址，参数2：邮箱密码
                stp.login(mail_sender, mail_license)
                # 发送邮件，传递参数1：发件人邮箱地址，参数2：收件人邮箱地址，参数3：把邮件内容格式改为str
                stp.sendmail(mail_sender, mail_receive, mm.as_string())
                result_string+=("邮件发送成功："+str(count)+"/"+str(len(tuples))+"\n")
                #print("邮件发送成功："+str(count)+"/"+str(len(tuples))+"\n")
                # 关闭SMTP对象
                stp.quit()
                if count==len(tuples):
                    return result_string
            except Exception as e:
                s="发送" + context[col_names.index("姓名")].value + "的邮件出错，可能是在setting中hr邮箱或者密码写错了，请修改setting后再次运行程序\n"
                result_string+=s
                return result_string
                # print("发送" + context[col_names.index("姓名")].value + "的邮件出错，可能是在setting中hr邮箱或者密码写错了，请修改setting后再次运行程序")
                # system("pause")
                # exit(-1)

#按钮事件
def press_send():
    set_text("")
    #获取输入框信息
    mail_=mail.get().split('\n')[0]
    pwd_=pwd.get().split('\n')[0]
    sender_name_=sender_name.get().split('\n')[0]
    title_=title.get().split('\n')[0]
    text_=text.get().split('\n')[0]
    excel_=excel.get().split('\n')[0]
    word_=word.get()
    #检测输入框是否都有信息
    if mail_ and pwd_ and sender_name_ and title_ and text_ and excel_ and word_:
        #若都有信息，发送邮件
        result=send_mail(mail_, pwd_, sender_name_, title_, text_, excel_,word_)
        set_text(result)
    else:
        message_box.config(state=NORMAL)
        message_box.insert(END,'请输入所有信息再点击发送')
        message_box.config(state=DISABLED)

def open_file(num):
    file_path = filedialog.askopenfilename(title=u'选择文件', initialdir=(os.path.expanduser('H:/')))
    if file_path is not None:
        if num==0:
            with open(file=file_path, mode='r+', encoding='utf-8') as file:
                mail.delete(0, END)
                pwd.delete(0, END)
                sender_name.delete(0, END)
                title.delete(0, END)
                text.delete(0, END)
                excel.delete(0, END)
                word.delete(0, END)
                t=file.readline().split(':')[1]
                mail.insert(0,t)
                t = file.readline().split(':')[1]
                pwd.insert(0, t)
                t = file.readline().split(':')[1]
                sender_name.insert(0, t)
                t = file.readline().split(':')[1]
                title.insert(0, t)
                t = file.readline().split(':')[1]
                text.insert(0, t)
                t = file.readline().split(':')[1]
                excel.insert(0, t)
                t = file.readline().split(':')[1]
                word.insert(0, t)
        elif num==1:
            excel.delete(0,END)
            excel.insert(0,file_path.split('/')[len(file_path.split('/'))-1])
        elif num==2:
            word.delete(0,END)
            word.insert(0,file_path.split('/')[len(file_path.split('/'))-1])


# 调用Tk()创建主窗口
root_window =Tk()
# 给主窗口起一个名字，也就是窗口的名字
root_window.title('自动办公工具')
# 设置窗口大小:宽x高,注,此处不能为 "*",必须使用 "x"
root_window.geometry('450x300')


mail_l=Label(root_window,text="hr邮箱:")
mail_l.grid(row=0,sticky=E)
pwd_l=Label(root_window,text="hr邮箱密码:")
pwd_l.grid(row=1,sticky=E)
name_l=Label(root_window,text="发送人姓名:")
name_l.grid(row=2,sticky=E)
title_l=Label(root_window,text="邮件主题:")
title_l.grid(row=3,sticky=E)
text_l=Label(root_window,text="邮件正文内容:")
text_l.grid(row=4,sticky=E)
excel_l=Label(root_window,text="待发送excel文件全名:")
excel_l.grid(row=5,sticky=E)
word_l=Label(root_window,text="word模板文件全名:")
word_l.grid(row=6,sticky=E)
result_l=Label(root_window, text="结果显示:")
result_l.grid(row=8,sticky=NE)

#输入控件
mail=Entry(root_window, width=35)
mail.grid(row=0,column=1)
pwd=Entry(root_window, width=35)
pwd.grid(row=1,column=1)
sender_name=Entry(root_window, width=35)
sender_name.grid(row=2,column=1)
title=Entry(root_window, width=35)
title.grid(row=3,column=1)
text=Entry(root_window, width=35)
text.grid(row=4,column=1)
excel=Entry(root_window, width=35)
excel.grid(row=5,column=1)
word=Entry(root_window, width=35)
word.grid(row=6,column=1)

#按钮控件
send_button=Button(root_window, text="发送", command=press_send)
send_button.grid(row=7, column=1)
open_file_button=Button(root_window, text="信息自动填入", command=lambda: open_file(0))
open_file_button.grid(row=7, column=0)
select_excel=Button(root_window, text="选择excel", command=lambda: open_file(1))
select_excel.grid(row=5, column=2, sticky=W)
select_word=Button(root_window, text="选择word", command=lambda: open_file(2))
select_word.grid(row=6, column=2, sticky=W)

#消息框控件
message_box=Text(root_window, height=10, width=40)
message_box.config(state=DISABLED)
message_box.grid(row=8,column=1,columnspan=8, sticky=W)



#开启主循环，让窗口处于显示状态
root_window.mainloop()