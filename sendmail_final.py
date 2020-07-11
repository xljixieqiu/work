import cx_Oracle,smtplib,datetime,xlwt,os,shutil,threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
def construct_excel(conn,sql,excel_name):
    print('excel construction started')
    xls=xlwt.Workbook()
    sheet1=xls.add_sheet('sheet1')
    #*_,day=dt.split('-')
    #print(day,type(day))
    search_data(conn,sql,sheet1)
    xls.save(excel_name)
    print('文件%s下载成功于：%s'%(excel_name,datetime.datetime.now()))
def search_data(conn,sql,sheet):
    #print('search_data start')
    c=conn.cursor()
    x=c.execute(sql)
    result=x.fetchall()
    headlist=[]
    head_index=c.description
    for index in head_index:
        headlist.append(index[0])
    print(headlist)
    j=0
    for i in headlist:
        sheet.write(0,j,i)
        j+=1
    if result:
        z=1
        for a in result:
            y=0
            for b in a:
                sheet.write(z,y,str(b))
                y+=1
            z+=1
    c.close()
    #print('search_data end')
def send_mail(to,dt,filelist):
    #print('send_mail start')
    from_addr='535992954@qq.com'
    psw='zlalwlbmulkqbihc'
    smtp_server='smtp.qq.com'
    msg=MIMEMultipart('mixed')
    msg['From']=from_addr
    msg['To']=to
    subject='数据'+dt
    msg['Subject']=Header(subject,'utf-8')
    for path in filelist:
        f=open(path,'rb').read()
        *_,name=path.split('/')
        f_sub=MIMEText(f,'base64','utf-8')
        f_sub['Content-Type']='application/octet-stream'
        f_sub.add_header('Content-Disposition','attchment',filename=name)
        msg.attach(f_sub)
    server=smtplib.SMTP(smtp_server,25)
    server.login(from_addr,psw)
    server.sendmail(from_addr,[to],msg.as_string())
    server.quit()
    print('mail send successful')
def move_file(from_path,to_path):
    #print('move start:',from_path,to_path)
    if not os.path.isfile(from_path):
        print('%s is not exist'%from_path)
    else:
        tpath,tname=os.path.split(to_path)
        if not os.path.exists(tpath):#os.path.exists  不是exist
            os.makedirs(tpath)
        shutil.move(from_path,to_path)
        print('file %s move from %s to %s'%(tname,from_path,to_path))
def count_second():#计算下次运行时间
    now_time=datetime.datetime.now()
    next_time=now_time+datetime.timedelta(days=1)
    next_year=next_time.date().year
    next_month=next_time.date().month
    next_day=next_time.date().day
    next_time=datetime.datetime.strptime(str(next_year)+'-'+str(next_month)+'-'+str(next_day)+' 08:00:00','%Y-%m-%d %H:%M:%S')
    interval_time=(next_time-now_time).total_seconds()
    return interval_time,next_day
def shao():
    try:
        conn=cx_Oracle.connect('szsrk/csrk#2014@2.34.202.135:1521/csxsm')
        today=datetime.datetime.now().strftime('%Y-%m-%d')
        sql1="select t3.jdname,t3.pcsname,t3.name,t2.姓名,t2.居民证号,case t2.性别 when '1' then '男' when '2' then '女' else '' end 性别,(2020 -substr(t2.出生日期,1,4)) 年龄,t2.个人联系电话,t6.mc 户籍,t2.服务处所,t1.address  暂住地址,t2.上传时间  from (select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t1,szsrk.web_zzrk_sj t2,(select * from csxsm.dict_sqxx where sfyx='0') t3,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t6 where to_char(t2.上传时间, 'yyyy-mm-dd') =to_char(sysdate-1,'yyyy-mm-dd')  and t2.cjlx = '1' and t2.社区代码 = t1.社区代码 and t1.出租屋编号 = t2.出租屋编号 and t2.社区代码 = t3.url and t2.户籍地址 = t6.dm  and t2.户籍地址 like '42%' and t2.户籍地址 not like '4201%' order by t3.jdname, t3.pcsname, t3.name"
        sql2="select t3.jdname,t3.pcsname,t3.name,t2.姓名,t2.居民证号,case t2.性别 when '1' then '男' when '2' then '女' else '' end 性别,(2020 -substr(t2.出生日期,1,4)) 年龄,t2.个人联系电话,t6.mc 户籍,t2.服务处所,t1.address  暂住地址,t2.上传时间  from (select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t1,szsrk.web_zzrk_sj t2,(select * from csxsm.dict_sqxx where sfyx='0') t3,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t6 where to_char(t2.上传时间, 'yyyy-mm-dd') =to_char(sysdate-1,'yyyy-mm-dd')  and t2.cjlx = '1' and t2.社区代码 = t1.社区代码 and t1.出租屋编号 = t2.出租屋编号 and t2.社区代码 = t3.url and t2.户籍地址 = t6.dm  and t2.户籍地址 like '4201%'  order by t3.jdname, t3.pcsname, t3.name"
        sql3="select t3.jdname,t3.pcsname,t3.name,t2.姓名,t2.居民证号,case t2.性别 when '1' then '男' when '2' then '女' else '' end 性别,(2020 -substr(t2.出生日期,1,4)) 年龄,t2.个人联系电话,t6.mc 户籍,t2.服务处所,t1.address  暂住地址,t2.上传时间  from (select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t1,szsrk.web_zzrk_sj t2,(select * from csxsm.dict_sqxx where sfyx='0') t3,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t6 where to_char(t2.上传时间, 'yyyy-mm-dd') =to_char(sysdate-1,'yyyy-mm-dd')  and t2.cjlx = '1' and t2.社区代码 = t1.社区代码 and t1.出租屋编号 = t2.出租屋编号 and t2.社区代码 = t3.url and t2.户籍地址 = t6.dm  and t2.户籍地址 like '11%'  order by t3.jdname, t3.pcsname, t3.name"
        filename=['湖北%s.xls'%today,'武汉%s.xls'%today,'北京%s.xls'%today]
        #print(filename[0],filename[1])
        from_path='d:/python work/email/'
        to_path='d:/python work/data/'
        to_addr='765851407@qq.com'
        construct_excel(conn,sql1,filename[0])
        construct_excel(conn,sql2,filename[1])
        construct_excel(conn,sql3,filename[2])
        filelist=[]
        tolist=[]
        for i in range(3):
            file=from_path+filename[i]#原文件路径
            toaddr=to_path+filename[i]#需要移动到的文件路径
            filelist.append(file)
            tolist.append(toaddr)#append 将元素添加到list尾部
        send_mail(to_addr,today,filelist)
        for i in range(3):
            move_file(filelist[i],tolist[i])
        interval,next_day=count_second()
        timer=threading.Timer(interval,shao)
        print('程序shao离下次运行还有%s秒，等待...'%interval)
        timer.start()
    except Exception as e:
        print(e)
    finally:
        conn.close()
def xufeng(next_day):
    if next_day=='1':
        try:
            conn=cx_Oracle.connect('szsrk/csrk#2014@2.34.202.135:1521/csxsm')
            today=datetime.datetime.now().strftime('%Y-%m-%d')
            yestoday=(datetime.datetime.now()+datetime.timedelta(days=-1)).strftime('%Y-%m')
            sql1="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and 是否精神病 = '1' order by t5.jdname, t5.pcsname, t5.name"
            sql2="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5  where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and t1.是否从事下列经营活动 = '3' order by t5.jdname, t5.pcsname, t5.name"
            sql3="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话,t3.mc 产业类型,t1.服务处所,t1.单位地址,t1.职业名称 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select flh dm, mc from dm.wpa_dicts_codes where dmlb = '84' order by flh) t3,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and 出生日期 between to_char(add_months(sysdate, -192), 'yyyymmdd') and to_char(add_months(sysdate, -120), 'yyyymmdd') and 职业名称 not like '%学%' and 服务处所 not like '%学%' and t1.产业类型 = t3.dm and (t1.是否娱乐场所 = '2' or t1.是否娱乐场所 is null) order by t5.jdname, t5.pcsname, t5.name"
            sql4="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话,t3.mc 产业类型,t1.服务处所,t1.单位地址,t1.职业名称 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select flh dm, mc from dm.wpa_dicts_codes where dmlb = '84' order by flh) t3,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and 出生日期>to_char(add_months(sysdate,-216),'yyyymmdd') and 职业名称 not like '%学%' and 服务处所 not like '%学%' and t1.产业类型 = t3.dm and t1.是否娱乐场所='1' order by t5.jdname, t5.pcsname, t5.name"
            sql5="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and t1.是否从事下列经营活动 = '1' order by t5.jdname, t5.pcsname, t5.name"
            sql6="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and t1.是否从事下列经营活动 = '5' order by t5.jdname, t5.pcsname, t5.name"
            sql7="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and t1.是否从事下列经营活动 = '2' order by t5.jdname, t5.pcsname, t5.name"
            sql8="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and t1.是否从事下列经营活动 = '4' order by t5.jdname, t5.pcsname, t5.name"
            sql9="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t1.现孕胎次,t4.address,t1.个人联系电话,t1.上传时间 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc,mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4, (select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and 出生日期>to_char(add_months(sysdate,-588),'yyyymmdd') and 出生日期<to_char(add_months(sysdate,-192),'yyyymmdd')and 性别='2'and 现孕胎次>'0' order by t5.jdname, t5.pcsname, t5.name"
            sql10="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t4.address,t1.个人联系电话,t1.上传时间 from szsrk.sz_zzrk_sj t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select * from szsrk.web_czfw_sj where 是否最新 = '0') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.社区代码 = t4.社区代码  and t1.出租屋编号=t4.出租屋编号 and t1.户籍地址 = t2.dm and t4.社区代码 = t5.url and t1.是否精神病 = '1' and t1.cjlx='1' and to_char(t1.上传时间,'yyyy-mm')='2020-03' order by t5.jdname, t5.pcsname, t5.name"
            sql11="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t3.mc 文化程度,t4.address,t1.个人联系电话 from csxsm.cs_zzrk_sj_1 t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' order by t.flh) t2,(select flh dm,mc from dm.wpa_dicts_codes t where t.dmlb = '54'and t.sfky='1' order by t.flh asc) t3,(select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t4, (select * from csxsm.dict_sqxx where sfyx='0') t5  where t1.出租屋编码 = t4.社区代码 || t4.出租屋编号 and t1.户籍地址 = t2.dm and t1.是否注销 = '0' and t4.社区代码 = t5.url and t1.文化程度<21 and t1.文化程度=t3.dm order by t5.jdname, t5.pcsname, t5.name"
            sql12="select t5.jdname,t5.pcsname,t5.name,t1.姓名,t1.居民证号,t2.mc 户籍,t3.mc 文化程度,t4.address,t1.个人联系电话,t1.上传时间 from szsrk.sz_zzrk_sj t1,(select flh dm, mc, mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' ) t2,(select flh dm,mc from dm.wpa_dicts_codes t where t.dmlb = '54'and t.sfky='1') t3,(select * from szsrk.web_czfw_sj where 是否最新 = '0') t4,(select * from csxsm.dict_sqxx where sfyx='0') t5 where t1.社区代码 = t4.社区代码  and t1.出租屋编号=t4.出租屋编号 and t1.户籍地址=t2.dm and t4.社区代码=t5.url and t1.文化程度<21 and t1.文化程度=t3.dm and t1.cjlx='1' and to_char(t1.上传时间,'yyyy-mm')=to_char((sysdate-2),'yyyy-mm') order by t5.jdname,t5.pcsname,t5.name"
            sql13="select t4.派出所名称,t4.社区名称,t1.出租屋编号,t2.姓名,t2.居民证号,t2.出生日期,t6.mc,t2.户籍地址详址,t1.address,t3.mc,t2.个人联系电话,t2.服务处所,t2.上传时间 from (select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t1,csxsm.cs_zzrk_sj_1 t2,(select flh dm, mc from dm.wpa_dicts_codes where dmlb = '303' order by flh) t3,(select s.url     社区代码,s.name    社区名称,s.pcsdm   派出所代码,s.pcsname 派出所名称 from dict_sqxx@SZRKZJ S  where length(s.pcsdm) <> 0 and s.url not like '%0' and s.xzqh = '320581' order by s.url) t4,(select flh dm, mc,mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' and t.sfky = '1' order by t.flh) t6 where t1.社区代码 || t1.出租屋编号 = t2.出租屋编码 and t1.出租类型 = t3.dm and t1.社区代码 = t4.社区代码 and t2.户籍地址 = t6.dm and t2.是否注销 = '0' and t2.户籍地址 in ('520624','520600','520602','520603','520621','520622','520623','520625','520626','520628','522221') order by t4.派出所名称, t4.社区名称, t1.出租屋编号"
            sql14="select t4.派出所名称,t4.社区名称,t1.出租屋编号,t2.姓名,t2.居民证号,t2.出生日期,t6.mc,t2.户籍地址详址,t1.address,t3.mc,t2.个人联系电话,t2.服务处所,t2.上传时间 from (select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t1,csxsm.cs_zzrk_sj_1 t2,(select flh dm, mc from dm.wpa_dicts_codes where dmlb = '303' order by flh) t3,(select s.url     社区代码,s.name    社区名称,s.pcsdm   派出所代码,s.pcsname 派出所名称 from dict_sqxx@SZRKZJ S where length(s.pcsdm) <> 0 and s.url not like '%0' and s.xzqh = '320581' order by s.url) t4,(select flh dm, mc,mnemonic from dm.wpa_dicts_codes t where t.dmlb = '24' and t.sfky = '1' order by t.flh) t6,(select distinct 居民证号 from szsrk.v_zzrk where cjlx = '1' and to_char(上传时间, 'yyyy-mm')=to_char(sysdate-2,'yyyy-mm') and 户籍地址 in ('520624','520600','520602','520603','520621','520622','520623','520625','520626','520628','522221')) t7 where t1.社区代码 || t1.出租屋编号 = t2.出租屋编码 and t1.出租类型 = t3.dm and t1.社区代码 = t4.社区代码 and t2.户籍地址 = t6.dm and t2.是否注销 = '0' and t2.居民证号 = t7.居民证号 order by t4.派出所名称, t4.社区名称, t1.出租屋编号"
            filename=['行为异常%s.xls'%today,'黑诊所%s.xls'%today,'十六岁以下%s.xls'%today,'娱乐场所%s.xls'%today,'黑网吧%s.xls'%today,'无证幼儿园%s.xls'%today,'黑中介%s.xls'%today,'地沟油%s.xls'%today,'现孕人员%s.xls'%today,'行为异常新增%s.xls'%today,'本科以上%s.xls'%today,'本科以上新增%s.xls'%today,'贵州铜仁%s.xls'%today,'贵州铜仁新增%s.xls'%today,]
            sqllist=[sql1,sql2,sql3,sql4,sql5,sql6,sql7,sql8,sql9,sql10,sql11,sql12,sql13,sql14]
            from_path='d:/python work/email/'
            to_path='e:/Users/admin/Desktop/徐峰/%s/'%yestoday
            from_path_list=[]
            to_path_list=[]
            to_addr='676278351@qq.com'
            for i in range(14):
                construct_excel(conn,sqllist[i],filename[i])
                fpath=from_path+filename[i]
                tpath=to_path+filename[i]
                from_path_list.append(fpath)
                to_path_list.append(tpath)
            send_mail(to_addr,today,from_path_list)
            for i in range(14):
                move_file(from_path_list[i],to_path_list[i])
        except Exception as e:
            print(e)
        finally:
            conn.close()
    interval,next_day=count_second()
    timer=threading.Timer(interval,xufeng,(str(next_day),))
    print('程序xu距离下次运行还有%s秒,请等待...'%interval)
    timer.start()
def lizd(next_day):
    if next_day=='1':
        try:
            conn=cx_Oracle.connect('szsrk/csrk#2014@2.34.202.135:1521/csxsm')
            today=datetime.datetime.now().strftime('%Y-%m-%d')
            yestoday=(datetime.datetime.now()+datetime.timedelta(days=-1)).strftime('%Y-%m-%d')
            yyear,ymonth,yday=yestoday.split('-')
            sql1="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 职业技能等级>'0'"
            sql2="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 职业技能等级='1'"
            sql3="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 职业技能等级='2'"
            sql4="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 职业技能等级='3'"
            sql5="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 职业技能等级='4'"
            sql6="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 职业技能等级='5'"
            sql7="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 参保情况2='1'"
            sql8="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 参保情况2='2'"
            sql9="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 参保情况2='3'"
            sql10="select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销='0' and 参保情况2='4'"
            sqllist=[sql1,sql2,sql3,sql4,sql5,sql6,sql7,sql8,sql9,sql10]
            resultlist=[]
            for s in sqllist:
                c=conn.cursor()
                x=c.execute(s)
                resultlist.append(x.fetchall()[0][0])
                c.close()
            print('数据库查询结束')
            with open('d:/python work/email/log/log.txt','r') as f:
                data=f.read()
                print(data)
                data1,data2=data.split(',')
            newp1=int(resultlist[0])-int(data1)
            newp2=int(resultlist[6])-int(data2)
            print('log读取结束')
            if newp1<0:
                newp1=abs(newp1)
                flag='减少'
                text1="2、新市民职业技能情况：%s月，%s登记职业资格新市民%s名。截至%s月%s日，系统内登记采集拥有职业资格新市民%s名，其中初级工%s名、中级工%s名、高级工%s名、技师%s名、高级技师%s名。"%(ymonth,flag,str(newp1),ymonth,yday,resultlist[0],resultlist[1],resultlist[2],resultlist[3],resultlist[4],resultlist[5])
            else:
                flag='增加'
                text1="2、新市民职业技能情况：%s月，%s登记职业资格新市民%s名。截至%s月%s日，系统内登记采集拥有职业资格新市民%s名，其中初级工%s名、中级工%s名、高级工%s名、技师%s名、高级技师%s名。"%(ymonth,flag,str(newp1),ymonth,yday,resultlist[0],resultlist[1],resultlist[2],resultlist[3],resultlist[4],resultlist[5])
            if newp2<0:
                newp2=abs(newp2)
                flag='减少'
                text2="3、新市民参保情况：%s月，%s登记从未参保人员%s名。截至%s月%s日，系统内登记采集新市民从未参保人员%s名、参保后断保人员%s名、外地参保人员%s名、本地参保人员%s名。"%(ymonth,flag,str(newp2),ymonth,yday,resultlist[6],resultlist[7],resultlist[8],resultlist[9])
            else:
                flag='增加'
                text2="3、新市民参保情况：%s月，%s登记从未参保人员%s名。截至%s月%s日，系统内登记采集新市民从未参保人员%s名、参保后断保人员%s名、外地参保人员%s名、本地参保人员%s名。"%(ymonth,flag,str(newp2),ymonth,yday,resultlist[6],resultlist[7],resultlist[8],resultlist[9])
            from_path='d:/python work/email/log/rusult%s.txt'%today
            with open(from_path,'w') as f:
                f.write(text1+'\n'+text2)
            to_addr='383804969@qq.com'
            #to_addr='xljixieqiu@163.com'
            from_path_list=[from_path,]
            print('文本附件构建结束')
            send_mail(to_addr,today,from_path_list)
            with open('d:/python work/email/log/log.txt','w') as f:
                f.write(str(resultlist[0])+','+str(resultlist[6]))
            print('log文件覆盖成功')
        except Exception as e:
            print(e)
        finally:
            conn.close()
    interval,next_day=count_second()
    timer=threading.Timer(interval,xufeng,(str(next_day),))
    print('程序li距离下次运行还有%s秒,请等待...'%interval)
    timer.start()
if __name__=='__main__':
    interval,next=count_second()
    print('离执行还有%s秒，请等待..'%interval)
    timer1=threading.Timer(interval,shao)
    timer2=threading.Timer(interval,lizd,(str(next),))
    timer3=threading.Timer(interval,xufeng,(str(next),))
    timer1.start()
    timer2.start()
    timer3.start()