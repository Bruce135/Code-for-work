# 发送html内容的邮件
import smtplib, time, random, os, glob
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
import pandas as pd
from email.utils import formataddr

addr = input('收件人列表:')
att = input('外放资料文件夹:')  # 外放资料文件夹
file_3 = input('未参加过模板:')
file_1 = input('已参加过模板:')


def send_mail_html():
    '''发送html内容邮件'''
    # 发送邮箱
    username = 'dep001@shanghairanking.com'  # 邮箱
    password = 'SH2020ranking'  # 密码
    sender = 'dep001@shanghairanking.com'
    send_from = '软科'  # 显示来源
    subject = '中国高校数据共享计划邀请函'  # 邮件标题
    data = pd.read_excel(addr)  # 邀请人信息
    li_files = glob.glob(os.path.join(att, '*.pdf'))
    ex_files = glob.glob(os.path.join(att, '*.xlsx'))

    # 读取html文件内容
    for i in range(data.shape[0]):

        if data.iloc[i, 4] == 1:
            f = open(file_1, 'rb')
            mail_body = f.read().decode('utf-8')  # 需解码方可替换
            f.close()
        else:
            f = open(file_3, 'rb')
            mail_body = f.read().decode('utf-8')  # 需解码方可替换
            f.close()
        name = data.iloc[i, 0]
        receiver = data.iloc[i, 3]

        mail_body = mail_body.replace('XXXXX大学', str(name))
        msg = MIMEMultipart('alternative')

        # 组装邮件内容和标题，中文需参数‘utf-8’，单字节字符不需要
        body = MIMEText(mail_body, _subtype='html', _charset='utf-8')
        msg['Subject'] = Header(subject, 'utf-8')
        msg['From'] = formataddr([send_from, 'dep001@shanghairanking.com'])
        # msg['From'] = sender
        msg['To'] = receiver
        msg.attach(body)

        # 加附件
        for path in li_files:
            file_name = path.split("\\")[-1]
            part = MIMEApplication(open(path, 'rb').read())
            part.add_header('Content-Disposition', 'attachment', filename=file_name)
            msg.attach(part)

        for sec_path in ex_files:
            file_name = sec_path.split("\\")[-1]
            part = MIMEApplication(open(sec_path, 'rb').read())
            part['Content_Type'] = 'application/octet-stream'  # 设置内容类型
            part.add_header('Content-Disposition', 'attachment', filename=file_name)
            msg.attach(part)

        # 判断是否需要抄送
        if data.iloc[i, 6] == 1:
            cc = str(data.iloc[i, 7])
            msg['Cc'] = cc
            receiver = [receiver, cc]
        # 登录并发送邮件
        try:
            print('链接服务器...')
            s = smtplib.SMTP()
            s.connect("smtp.qiye.163.com")  # qq："smtp.qq.com", 网易："smtp.163.com"
            s.login(username, password)
            s.sendmail(sender, receiver, msg.as_string())
        # 发送邮箱用户/密码
        except Exception as e:
            r = open(os.path.join(att,'f.txt'), 'w')
            r.writelines(str(receiver))
            r.close()
            print("邮件发送失败！", e)
        else:
            print("邮件发送成功！")
        finally:
            s.close()

        period = random.randint(105, 125)
        ti = time.asctime(time.localtime(time.time() + period))
        print('开始休眠...下次运行时间：%s' % ti)
        time.sleep(period)


if __name__ == '__main__':
    send_mail_html()



