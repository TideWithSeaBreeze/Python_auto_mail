import smtplib,pymysql,json,schedule,datetime,os
from email.mime.text import MIMEText
from openpyxl import load_workbook
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart

#展开setting.json,读取email、database配置信息
with open('setting.json','r') as jsonfile:
    settingfile = json.load(jsonfile)
    email_account = settingfile['email']['sender']
    email_password = settingfile['email']['authorization_code']
    email_stmp_server = settingfile['email']['stmp_server']
    email_stmp_port = settingfile['email']['stmp_server_port']
    host = settingfile['DataBase']['host']
    port = settingfile['DataBase']['port']
    user = settingfile['DataBase']['user']
    password = settingfile['DataBase']['password']
    database = settingfile['DataBase']['database']

#mail_part1,创建邮件主题信息、正文内容
def mail_part1(mail_job,email_account):
    message = MIMEMultipart()
    message['From']= email_account
    message['To'] = mail_job[2]
    message['Subject'] = mail_job[1]
    message.attach(MIMEText(mail_job[3],'html','utf8'))
    return message

#mail_part2,创建附件信息
def mail_part2(cursor,job_id,message,connect,title):
    cursor.execute(f'select attachment_id,job_id,workbook_title,workbook_position,worksheet_1_title,worksheet_1_start,worksheet_1_sql,worksheet_2_title,worksheet_2_start,worksheet_2_sql,worksheet_3_title,worksheet_3_start,worksheet_3_sql,worksheet_4_title,worksheet_4_start,worksheet_4_sql,is_done from am_auto_mail_attachments where job_id = {job_id} and is_done <> 1')
    mail_body = cursor.fetchall()
    for i in range(0,len(mail_body)):
        attachment_id,job_id,workbook_title,workbook_position,worksheet_1_title,worksheet_1_start,worksheet_1_sql,worksheet_2_title,worksheet_2_start,worksheet_2_sql,worksheet_3_title,worksheet_3_start,worksheet_3_sql,worksheet_4_title,worksheet_4_start,worksheet_4_sql,is_done = mail_body[i]

        workbook_body = load_workbook(f'./template/{workbook_position}')
        list_sql,list_title,list_start = [worksheet_1_sql,worksheet_2_sql,worksheet_3_sql,worksheet_4_sql],[worksheet_1_title,worksheet_2_title,worksheet_3_title,worksheet_4_title],[worksheet_1_start,worksheet_2_start,worksheet_3_start,worksheet_4_start]
        list_sql = list(list_filter(list_sql))
        list_title = list(list_filter(list_title))
        list_start = list(list_filter(list_start))

        for i in range(0,len(list_sql)):
            worksheet_body = workbook_body['Sheet'+str(i+1)]
            sheet_maker(cursor,list_sql[i],list_title[i],list_start[i],worksheet_body)

        dir_create(title,workbook_body,workbook_title)

        with open(f'./attachment/{title}/{workbook_title}.xlsx','rb') as file:
            data = file.read()
        attachment_file = MIMEApplication(data,__subtype='xlsx',name=f'{workbook_title}.xlsx')
        message.attach(attachment_file)

        cursor.execute(f'update am_auto_mail_attachments set is_done=1 where attachment_id ={attachment_id}')
        connect.commit()
    return message

#sheet_maker,对sheet页进行写入
def sheet_maker(cursor,worksheet_sql,worksheet_title,worksheet_start,worksheet_body):
    cursor.execute(worksheet_sql)
    sheet_value = cursor.fetchall()
    worksheet_body.title = worksheet_title
    worksheet_body._current_row = int(worksheet_start)
    for j in range(0,len(sheet_value)):
        worksheet_body.append(sheet_value[j])

#list_filter,对数组进行长度处理，剔除空和None
def list_filter(list_value):
    def is_not_empty(value):
        return value and len(value.strip()) > 0
    list_return = list(filter(is_not_empty,list_value))
    return list_return

#dir_create,检查Excel附件文件夹是否存在，创建对应文件夹
def dir_create(title,workbook_body,workbook_title):
    if not os.path.exists(f'./attachment/{title}'):
        os.mkdir(f'./attachment/{title}')
        print(f'--{datetime.datetime.now()}--dir {title} not exist, mkdir success!')
    workbook_body.save(f'./attachment/{title}/{workbook_title}.xlsx')

#check_mail_job,主要，用于检查数据库邮件任务
def check_mail_job():
    connect = pymysql.connect(host=host,port=port,user=user,passwd=password,db=database)
    cursor = connect.cursor()
    cursor.execute('select job_id,title,receiver,mail_body from am_auto_mail_jobs where is_done<>1 limit 1')
    mail_job = cursor.fetchone()
    if mail_job is not None:
        print(f'--{datetime.datetime.now()}--find mail_job')
        message = mail_part1(mail_job,email_account)
        try:
            mail_part2(cursor,mail_job[0],message,connect,message['Subject'])

            smtp = smtplib.SMTP_SSL(email_stmp_server,email_stmp_port)
            smtp.login(email_account,email_password)
            smtp.send_message(message)
            print(f'--{datetime.datetime.now()}--send email ({mail_job[1]}) success!')
            smtp.quit()

            cursor.execute(f'update am_auto_mail_jobs set is_done=1 where job_id ={mail_job[0]}')
            connect.commit()
            
        except smtplib.SMTPException as e:
            print(e)

#循环周期配置
def schedule_job():
    schedule.every(5).seconds.do(check_mail_job)
    while True:
        schedule.run_pending()

#程序入口
if __name__ == '__main__':
    print(f'''auto_mail程序已启动，按Ctrl+C退出~
启动时间：{datetime.datetime.now()}''')
    schedule_job()
