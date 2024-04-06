# Python_auto_mail
Use SQL and a little html can make e-mail by system.
用SQL语句和简单html代码，就能够实现系统在设置的计划任务时完成自动邮件任务；

一、整体实现逻辑：
（一）数据来源：数据库（v1.0目前仅支持MySQL），由SQL语句生成数据实体表、邮件html正文、邮件标题等，并写入数据库表；
（二）数据处理：数据处理全部依赖于数据库，用户无需了解Python语法；
（三）自动化：利用数据库计划任务功能，在计划时间点对数据库的邮件任务表中，写入一个邮件任务，auto_mail会每隔5秒检索一次数据库表格；

以上就是对auto_mail的简单介绍，即数据库完成数据处理，Python完成数据转Excel表附件及邮件发送任务；

二、基础依赖：
（一）数据库：v1.0目前仅支持MySQL，后续会追加对其他常见数据库的支持；
（二）推荐运行环境：windows server服务器；
（三）安装与运行：直接双击auto_mail.exe，当正常运行时会提示启动成功，并生成启动时间；

三、数据库需要做的操作：
（一）表创建：需要创建两张关键表：am_auto_mail_jobs（自动邮件任务表）、am_auto_mail_attachments（自动邮件附件表）；
SQL语句如下：
```sql
create table am_auto_mail_jobs(job_id int auto_increment primary key,title varchar(200) not null,receiver varchar(300) not null,mail_body varchar(3000) not null,is_done int not null);
create table am_auto_mail_attachments(attachment_id int auto_increment primary key,job_id int not null,workbook_title varchar(200) not null,workbook_position varchar(200) not null,worksheet_1_title varchar(100) not null,worksheet_1_start int not null,worksheet_1_sql varchar(3000) not null,worksheet_2_title varchar(100),worksheet_2_start int,worksheet_2_sql varchar(3000) not null,worksheet_3_title varchar(100),worksheet_3_start int,worksheet_3_sql varchar(3000) not null,worksheet_4_title varchar(100),worksheet_4_start int,worksheet_4_sql varchar(3000) not null,is_done int not null);
```
(二）配置auto_mail的数据库账号：
不建议直接使用root账户，建议配置专用账号，开放am_auto_mail_jobs（自动邮件任务表）、am_auto_mail_attachments（自动邮件附件表）两个表的增删改权限，开放其他表的查询权限；

四、创建一个mail_job测试任务：
（一）使用建议：建议采用编写存储过程+计划任务调度存储过程的方式：
存储过程实例：
