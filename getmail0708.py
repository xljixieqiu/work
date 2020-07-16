import cx_Oracle,poplib,datetime,xlrd,os,shutil,threading,re
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header,decode_header
from email.parser import Parser
def get_mail():#获取邮件标题、发件箱。从邮箱下载附件,并保存附件路径（判断条件为邮件标题包含‘新市民采集’或者‘PDA’字样）
    try:
        email='youmail@qq.com'
        passw='yourpassword'
        pop3_server='pop.qq.com'
        server=poplib.POP3(pop3_server)
        server.user(email)
        server.pass_(passw)
        resp,mails,octets=server.list()
        print('get mail list')
        with open('d:/python work/email/log/getmaillog.txt','r') as f:
            first=f.read()#获取上次邮件读取到第first封
        #print('read txt successful,first is %s,type is%s'%(first,type(first)))
        index=len(mails)
        #print('index=%s,type=%s'%(index,type(index)))
        if index>int(first):#如果长度大于上次  说明有新邮件
            #print('if ok')
            lists=[]
            for i in range(int(first)+1,index+1):
                #print('for ok')
                resp,lines,octets=server.retr(i)
                msg_content=b'\r\n'.join(lines).decode('utf-8')
                msg=Parser().parsestr(msg_content)
                if check_subject(msg):
                    filelist=get_attch(msg)
                    lists.extend(filelist)#获取总的文件列表
            with open('d:/python work/email/log/getmaillog.txt','w') as f:
                f.write(str(index))#覆盖原来的log文件
            get_data(lists)
            print('流程结束')
        else:
            print('没有新邮件')
        timer=threading.Timer(43200,get_mail)
        print('get_mail run after 43200 seconds')
        timer.start()
    except Exception as e:
        print(e)
    finally:
        server.quit()
        print("pop3 server is quit")
#def check_mail():#判断邮件是否在已读列表中。如果在，则返回0；不在则返回1，并把邮件编号记录在已读列表中
def get_attch(msg):#下载附件到指定文件夹，并返回文件列表
    attachment_lists=[]
    '''#print(msg.iter_attachments())
    for att in msg.iter_attachments():
        if att.get_filename().split('.')[-1]=='xls':
            print(att.get_filename())
    '''
    for part in msg.walk():
        filename=part.get_filename()
        content_type=part.get_content_type()
        if filename:
            h=Header(filename)
            dh=decode_header(h)
            filename=dh[0][0]
            #print(dh[0][0],dh[0][1])
            if dh[0][1]:
                filename=decode_str(str(filename,dh[0][1]))
            #print(filename)
            data=part.get_payload(decode=True)#下载附件 
            #print(data)
            nowtime=datetime.datetime.now().strftime('%Y-%m-%d%H%M%S')
            filepath='d:/python work/email/'+nowtime+filename
            if filename.split('.')[-1]=='xls' or filename.split('.')[-1]=='xlsx':
                with open(nowtime+filename,'wb') as f:
                    f.write(data)
            print('xls write ok')
            attachment_lists.append(filepath)
    return attachment_lists
def decode_str(s):#字符编码转换
    value,charset=decode_header(s)[0]
    if charset:
        value=value.decode(charset)
    return value
def check_subject(msg):#判断邮件标题是否包含‘新市民采集’或者‘PDA’字样，有则返回0;无则返回1
    print('check_subject start')
    value=msg.get('Subject','')
    if value:
        value=decode_str(value)
        if re.match(r'[\u4e00-\u9fa5]{0,15}?(新市民|账号|PDA|pda)',value):
            print('subject matched')
            return True
    return False
def get_data(file_lists):#从附件路径读取excel，获取excel内容，并update进数据库
    try:
        conn=cx_Oracle.connect('admin/passwor@127.0.0.1:1521/xx')
        print('get_data start')
        for path in file_lists:
            workbookname=path.split('/')[-1]
            to_path='D:/pda/'+workbookname
            wrong_path='D:/pda/wrong/'+workbookname
            errorcount=0
            workbook1=xlrd.open_workbook(path)
            names=workbook1.sheet_names()#获取所有工作表名称,结果为列
            for name in names:
                sheet=workbook1.sheet_by_name(name)
                if sheet:
                    rows=sheet.nrows#获取sheet总行数
                    pcs=sheet.row_values(1)[1]
                    print(pcs)
                    if pcs=='' or pcs is None:
                        print('sheet%s does not have a pliceoffice name'%name)
                        break
                    for i in range(3,rows):
                        row=sheet.row_values(i)
                        count=rename_count(data_clean(row[1])).lower()#处理账号信息：1.补上cs,2.cs改为小写
                        subtpye=data_clean(row[4])
                        cname=data_clean(row[2])
                        print(count)
                        res=check_account(count)
                        if res==1:#检查账号是否符合规定
                            #print('辅警账号')
                            if subtpye=='新增' or subtpye=='变更':
                                if check_oracle(count,conn):#检查账号是否存在于数据库
                                    update_oracle1(pcs,count,cname,conn)
                                else:
                                    insert_oracle1(pcs,count,cname,conn)
                            elif subtpye=='注销':
                                if check_account(count,conn):
                                    delete_oracle(count,conn)
                                else:
                                    print('文件%s的%s账号%s不存在，无法注销'%(workbookname,name,count))
                                    errorcount+=1
                            else:
                                print('文件%s的%s账号%s申请类型错误'%(workbookname,name,count))
                                errorcount+=1
                        elif res==2:
                            #print('民警账号')
                            if subtpye=='新增' or subtpye=='变更':
                                if check_oracle(count,conn):#检查账号是否存在于数据库
                                    update_oracle2(pcs,count,cname,conn)
                                else:
                                    insert_oracle2(pcs,count,cname,conn)
                            elif subtpye=='注销':
                                if check_account(count,conn):
                                    delete_oracle(count,conn)
                                else:
                                    print('文件%s的%s账号%s不存在，无法注销'%(workbookname,name,count))
                                    errorcount+=1
                            else:
                                print('文件%s的%s账号%s申请类型错误'%(workbookname,name,count))
                                errorcount+=1
                        elif res==3:
                            pass
                        else:
                            print('文件%s的%s账号%s错误'%(workbookname,name,count))
                            errorcount+=1
        #xlrd不需要关闭，打开后直接加载到内存
            if errorcount>0:
                move_file(path,wrong_path)
            else:
                move_file(path,to_path)
    except Exception as e:
        print(e)
    #finally:
        #conn.close()
def check_account(count):#检查账号是否符合规定
    if re.match(r'^cs\d{6}$',count):
        return 1
    if re.match(r'^2\d{5}$',count):
        return 2
    if re.match(r'^ml\d{6}$',count):
        return 1
    if re.match(r'^ft\d{6}$',count):
        return 1
    if count=='' or count is None:
        return 3
    return False
def rename_count(count):#如果没加cs,补全账号信息
    if re.match(r'^1\d{5}$',count):
        print('开始补全账号信息')
        count='cs'+count
    return count
def check_oracle(count,conn):#检查账号是否存在
    print('开始检查账号%s是否存在'%count)
    sql="select * from psms.tp_users where mjjh='%s'"%count
    #print('sql is:',sql)
    try:
        c=conn.cursor()
        x=c.execute(sql)
        result=x.fetchall()
        print(result)
        if result:
            #print('true')
            return True
        #print('false')
        return False
    except Exception as e:
        print(e)
    finally:
        c.close()
def update_oracle2(pcs,count,cname,conn):#更新民警账号信息
    pcs=pcs[0:2]
    sql="update psms.tp_users set mjjh='%s',mjmm=8888,mjxm='%s',jgbm=(select distinct pcsdm from csxsm.dict_sqxx where sfyx='0' and pcsname like '%s%%'),role_id='RID_SQ',mjlb=3,gpsbiz=120,sfzx=0 where mjjh='%s'"%(count,cname,pcs,count)
    print('开始变更民警账号%s'%count)
    try:
        c=conn.cursor()
        c.execute(sql)
        conn.commit()
        print('变更成功')
    except Exception as e:
        print(e)
        conn.rollback()
    finally:
        c.close()
def insert_oracle2(pcs,count,cname,conn):#新增民警账号信息
    print('开始新增民警账号%s'%count)
    pcs=pcs[0:2]
    sql="insert into psms.tp_users values('%s',8888,'%s',(select distinct pcsdm from csxsm.dict_sqxx where sfyx='0' and pcsname like '%s%%'),'','','RID_SQ','',3,120,0)"%(count,cname,pcs)
    try:
        c=conn.cursor()
        c.execute(sql)
        conn.commit()
        print('新增成功')
    except Exception as e:
        print(e)
        conn.rollback()
    finally:
        c.close()
def update_oracle1(pcs,count,cname,conn):#更新辅警账号信息
    print('开始变更辅警账号%s'%count)
    pcs=pcs[0:2]
    sql="update psms.tp_users set mjjh='%s',mjmm=1234,mjxm='%s',jgbm=(select distinct pcsdm from csxsm.dict_sqxx where sfyx='0' and pcsname like '%s%%'),role_id='RID_SQ_FJ',mjlb=3,gpsbiz=120,sfzx=0 where mjjh='%s'"%(count,cname,pcs,count)
    #print(sql)
    try:
        c=conn.cursor()
        c.execute(sql)
        conn.commit()
        print('变更成功')
    except Exception as e:
        print(e)
        conn.rollback()
    finally:
        c.close()
def insert_oracle1(pcs,count,cname,conn):#新增辅警账号信息
    print('开始新增辅警账号%s'%count)
    pcs=pcs[0:2]
    sql="insert into psms.tp_users values('%s',1234,'%s',(select distinct pcsdm from csxsm.dict_sqxx where sfyx='0' and pcsname like '%s%%'),'','','RID_SQ_FJ','',3,120,0)"%(count,cname,pcs)
    try:
        c=conn.cursor()
        c.execute(sql)
        conn.commit()
        print('新增成功')
    except Exception as e:
        print(e)
        conn.rollback()
    finally:
        c.close()
def delete_oracle(count,conn):#删除账号
    print('开始注销账号%s'%count)
    sql="update psms.tp_users set sfzx='1' where mjjh='%s'"%count
    try:
        c=conn.cursor()
        c.execute(sql)
        conn.commit()
        print('注销成功')
    except Exception as e:
        print(e)
        conn.rollback()
    finally:
        c.close()
def data_clean(str1):#清理excel数据
    if isinstance(str1,float):
        str1=str(int(str1))#excel单元格中纯数字的类型为float，需要强制转换成int类型（去掉小数点），在转换为str类型（strip（）方法只有str才有）
    return str.strip()
def move_file(from_path,to_path):#移动文件到指定文件夹
    if not os.path.isfile(from_path):
        print('%s is not exist!'%from_path)
    else:
        path,filename=os.path.split(to_path)
        if not os.path.exists(path):
            os.makedirs(path)
        shutil.move(from_path,to_path)
        print('%s move to %s'%(from_path,to_path))
def error_msg(msg):#保存报错信息
    with open('d:/python work/email/log/filelog.txt','r') as f:
        content=f.read()
        content.extend(msg+'\n')
    with open('d:/python work/email/log/filelog.txt','w') as f:
        f.write(content)
if __name__=='__main__':
    print('start')
    get_mail()
    
