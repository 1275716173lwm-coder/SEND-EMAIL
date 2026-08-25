from cProfile import label
from http import server
from socket import create_server
from tkinter import ttk
import xlrd
import os
import smtplib
import tkinter as tk
from tkinter import filedialog, messagebox
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
from email.mime.base import MIMEBase
from email import encoders
import time
import random
#import pandas as pd
##define_email_function

accounts = [
    {"email":"1275716173@qq.com", "password":os.getenv("EMAIL1_PASSWORD"), "smtp":"smtp.qq.com", "port":465, "ssl":True},
    {"email":"2521048381@qq.com", "password":os.getenv("EMAIL2_PASSWORD"), "smtp":"smtp.qq.com", "port":465, "ssl":True},
]

def create_server(account):
    if not account["password"]:
        raise RuntimeError("请先设置邮箱授权码环境变量 EMAIL1_PASSWORD 和 EMAIL2_PASSWORD")
    if account["ssl"]:
        server = smtplib.SMTP_SSL(account["smtp"], account["port"])
    else:
        server = smtplib.SMTP(account["smtp"], account["port"])
    #server.ehlo()   
        server.starttls()
    server.login(account["email"], account["password"])
        
    #server.send_message(msg)
    #server.quit()
    return server
#server = create_server(sender_email, sender_password)
def send_email(server, receiver_email, subject, body, file_path,sender_email):
    
    msg=MIMEMultipart()
    msg['From']=sender_email
    msg['To']=receiver_email
    msg['Subject']=Header(subject,'utf-8')

    msg.attach(MIMEText(body, 'plain', 'utf-8'))

##adding attachment
    with open(file_path, "rb") as f:
        part = MIMEApplication(
        f.read(),
        _subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )   

        filename = os.path.basename(file_path)
        part.add_header(
        'Content-Disposition',
        'attachment',
        filename=Header(filename, 'utf-8').encode()
        )
        msg.attach(part)

    server.send_message(msg)
    #server.quit()   

##sending_email

    #server.send_message(msg)       

#server = create_server()



def start_sending():
    account_index = 0
    send_count = 0
    excel_path = excel_entry.get()
    file_dir = folder_entry.get()

    server = create_server(accounts[account_index]) 
    if not excel_path or not file_dir:
        messagebox.showerror("错误", "请选择Excel和文件夹")
        return

    data = xlrd.open_workbook(excel_path)

#df = pd.read_excel('1.xlsx')
    sheet = data.sheets()[0]
    total = sheet.nrows
    progress['max'] = total
    fail = 0
    success = 0
    fail_list = []
    #ender_email = "2521048381@qq.com"
    #sender_email = "a1275716173@outlook.com"


  
##find_corresponding_file
    for i in range(0, sheet.nrows):
    #name = str(row['i']).strip() 
        name = sheet.cell(i, 0).value
        receiver_email = sheet.cell(i, 1).value
    #name_int = int(name)
        print(f"{name}")  ##print the name of the file to be sent
        file_path = os.path.join(file_dir, f"{name}.xlsx")
    #print('%d' % name_int)
        try:
            if os.path.exists(file_path):
                send_email(
                    server, 
            #receiver_email=receiver_email,
            #receiver_email = "1275716173@qq.com",
                    receiver_email = receiver_email,
                    subject="飞行二大队一中队工资明细",
                    body=f"{name}您好，这是你本月工资明细，请查收附件。",

                    file_path=file_path,
                    #server = server,
                    #sender_email = sender_email,
                    sender_email = accounts[account_index]["email"]
 
                )
                success += 1
                send_count += 1
 
        #print(f"{name}.xlsx exists.")   
            else:
                fail += 1
                print(f"{name}.xlsx does not exist.")     ##if not sent, this will be displayed 
                fail_list.append(name)
        except Exception as e:
            log_text.insert(tk.END, f"发送{name}失败: {str(e)}\n")
            fail += 1
            fail_list.append(name)
            time.sleep(10)  ##wait for 10 seconds before retrying
        
        sleep_time = random.uniform(8,15)
        time.sleep(sleep_time)  ##random sleep time between 1 and 3 seconds

        if (i + 1) % 10 == 0:  ##update progress every 10 emails
            time.sleep(15)
        if (i + 1) % 18 == 0:  ##update progress every 10 emails
            #server.quit()
            print(f"已发送 {i + 1} 封邮件，正在切换账号...")
            account_index = (account_index + 1) % len(accounts)  ##switch to the next account
            server = create_server(accounts[account_index])  ##create a new server connection with the new account
            time.sleep(15)
            #server = create_server(sender_email, sender_password)
    #server.quit()
        


        progress['value'] = i + 1
        percent = (i + 1) / total * 100
        percent_label.config(text=f"PROGRESS: {percent:.2f}%")
        percent_label.pack()    
        root.update() 
    server.quit()  ##close the server connection after sending all emails

    file_path_1 = os.path.join(os.getcwd(), "失败名单.txt")
    with open(file_path_1, "w", encoding="utf-8") as f:
        if fail_list:
         for name in fail_list:
            f.write(name + "\n")
        else:
         f.write("无失败记录")

    print("失败名单已生成:", file_path)
 ##close the server connection after sending all emails





    #if fail_list:
       # with open("fail_list.txt", "w", encoding="utf-8") as f:
            #f.write("\n".join(fail_list))   
        #messagebox.showwarning("完成", f"发送完成，但以下文件未找到:\n{', '.join(fail_list)}")
    #messagebox.showinfo( f"成功: {success}失败: {fail}")
    ##display the result
    result_window = tk.Toplevel(root)
    result_window.title("发送结果")
    result_window.geometry("300x150")
    label = tk.Label(
    result_window,
    text=f"发送完成\n\n成功：{success} 个\n失败：{fail} 个",
    font=("Arial", 12),
    justify="center"
)
    label.pack(expand=True)







root = tk.Tk()
root.title("邮件自动发送工具")
root.geometry("500x400")
PERCENT_LABEL = tk.Label(root, text="PROGRESS: 0.00%")
percent_label = PERCENT_LABEL
PERCENT_LABEL.pack()
# select excel
tk.Label(root, text="Excel文件:").pack()
excel_entry = tk.Entry(root, width=50)
excel_entry.pack()
tk.Button(root, text="选择Excel",
          command=lambda: excel_entry.insert(0, filedialog.askopenfilename())).pack()

# select folder
tk.Label(root, text="附件文件夹:").pack()
folder_entry = tk.Entry(root, width=50)
folder_entry.pack()
tk.Button(root, text="选择文件夹",
          command=lambda: folder_entry.insert(0, filedialog.askdirectory())).pack()
# button
tk.Button(root, text="开始发送", command=start_sending, bg="green").pack(pady=10)


log_text = tk.Text(root, height=10)
log_text.pack()

progress = ttk.Progressbar(root, orient="horizontal", length=400, mode="determinate")
progress.pack(pady=10)

root.mainloop()
