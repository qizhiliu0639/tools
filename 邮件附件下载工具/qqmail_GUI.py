# -*- coding: utf-8 -*-
"""
Created on Thu Dec 17 22:33:26 2020

@author: qizhiliu
"""
 
import tkinter as tk  # 使用Tkinter前需要先导入
import poplib
import email
import datetime
import time
from email.parser import Parser
from email.header import decode_header
import traceback
import telnetlib
import os

import tkinter.messagebox

import c_step4_get_email

def redirector(inputStr):
    t.insert(tk.INSERT, inputStr)

class c_step4_get_email:
    # 字符编码转换
    @staticmethod
    def decode_str(str_in):
        value, charset = decode_header(str_in)[0]
        if charset:
            value = value.decode(charset)
        return value
        
    # 解析邮件,获取附件

    @staticmethod
    def get_att(msg_in, str_day_in):
        # import email
        attachment_files = []
        for part in msg_in.walk():
            # 获取附件名称类型
            file_name = part.get_filename()
            # contType = part.get_content_type()
            if file_name:
                h = email.header.Header(file_name)
                # 对附件名称进行解码
                dh = email.header.decode_header(h)
                filename = dh[0][0]
                if dh[0][1]:
                    # 将附件名称可读化
                    filename = c_step4_get_email.decode_str(str(filename, dh[0][1]))
                    print(filename)
                    tk.sys.stdout.write = redirector
                    # filename = filename.encode("utf-8")
                # 下载附件

                data = part.get_payload(decode=True)
                # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
                isExists=os.path.exists('./attachment')
                if not isExists:
                    os.makedirs('./attachment')
                    print('\n')
                    print('不存在attachment文件夹，创建attachment文件夹成功')
                    tk.sys.stdout.write = redirector
                
                att_file = open('./attachment/' + filename, 'wb')
                attachment_files.append(filename)
                att_file.write(data)  # 保存附件
                att_file.close()
        return attachment_files
        
    @staticmethod
    def run_ing(str_day,email_user,password):
        pop3_server = 'pop.qq.com'
        # 日期赋值
        #str_day = 20201203
        # 连接到POP3服务器,有些邮箱服务器需要ssl加密，可以使用poplib.POP3_SSL

        try:
            telnetlib.Telnet('pop.qq.com', 995)
            server = poplib.POP3_SSL(pop3_server, 995, timeout=20)
        except:
            time.sleep(5)
            server = poplib.POP3(pop3_server, 110, timeout=10)
        # server = poplib.POP3(pop3_server, 110, timeout=120)
        # 可以打开或关闭调试信息
        # server.set_debuglevel(1)
        # 打印POP3服务器的欢迎文字:
        #print(server.getwelcome().decode('utf-8'))
        tk.sys.stdout.write = redirector
        # 身份认证:
        server.user(email_user)
        server.pass_(password)
        # 返回邮件数量和占用空间:
        #print('Messages: %s. Size: %s' % server.stat())
        tk.sys.stdout.write = redirector
        # list()返回所有邮件的编号:
        resp, mails, octets = server.list()
        # 可以查看返回的列表类似[b'1 82923', b'2 2184', ...]
        #print(mails)
        index = len(mails)
        #print(index)
        # 倒序遍历邮件
        # for i in range(index, 0, -1):
        # 顺序遍历邮件

        for i in range(index,0, -1):
            try:
                resp, lines, octets = server.retr(i)
                # lines存储了邮件的原始文本的每一行,
                # 邮件的原始文本: 
                msg_content = b'\r\n'.join(lines).decode('gbk')
                # 解析邮件:
                msg = Parser().parsestr(msg_content)
                # 获取邮件时间,格式化收件时间
                msg_time = msg.get("Date")[0:24]
                date1 = time.strptime(msg_time, '%a, %d %b %Y %H:%M:%S')
                # 邮件时间格式转换
                date2 = time.strftime("%Y%m%d", date1)
                if int(date2) < str_day:
                    # 倒叙用break
                     break
                    # 顺叙用continue
                  #  continue
                # else:
                    # 获取附件
                c_step4_get_email.get_att(msg, str_day)
            except:
                continue
        # print_info(msg)
        print('下载完成')
        tk.sys.stdout.write = redirector
        server.quit()
    
if __name__ == '__main__':
    window = tk.Tk()
    window.title('写给苗苗宝宝的邮件附件下载小工具')
    window.geometry('600x400')
    e1 = tk.Label(window,text = '请注意，2020年3月1号填写规范为 20200301', font=('Arial', 12), width=50, height=2)
    e2 = tk.Entry(window, show=None, font=('Arial', 14))
    e3 = tk.Label(window,text = 'email', font=('Arial', 12), width=50, height=2)
    e4 = tk.Entry(window, show=None, font=('Arial', 14))
    e5 = tk.Label(window,text = 'password', font=('Arial', 12), width=50, height=2)
    e6 = tk.Entry(window, show=None, font=('Arial', 14))
    e1.pack()
    e2.pack()
    e3.pack()
    e4.pack()
    e5.pack()
    e6.pack()
    
    

    
    
    def show_num():
        str_day = e2.get()
        email_account = e4.get()
        password = e6.get()
        try:
            str_day = int(str_day)
            try:
                c_step4_get_email.run_ing(str_day,email_account,password)
                #tk.sys.stdout.write = redirector
            except Exception as e:
                s = traceback.format_exc()
                print(e)
                tra = traceback.print_exc()
        except:
            tkinter.messagebox.showerror(title='Error', message='文本输入格式不对')
        
        
    
    b1 = tk.Button(window,text = '确认',width = 10,height = 3,command = show_num)
    b1.pack()
    
    
    t = tk.Text(window, height=10)
    t.pack()
    

    
    
    
    
    
    window.mainloop()


