import tkinter,xlrd,xlwt,cx_Oracle,sys,threading,re,openpyxl
from tkinter import filedialog
loops=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]#idcheck所需
checknum=[1,0,'X',9,8,7,6,5,4,3,2]#idcheck所需
def id_check(arg,nrow):
    if arg is not None or arg!='':
        s=list(arg)
    #l=''
        count=0
        if len(s)==18:
            for i in range(17):
                count +=int(loops[i])*int(s[i])
                   #print(count)
            #print(count)
            result=str(checknum[count%11])
            #print(result)
            #print(s[17])
            if result==str(s[17]):
                 pass#l='正确'
            else:
                print('第%d行身份证错误'%nrow)#l='错误' 跳出程序
                sys.exit()
        else:
            print('第%d行身份证长度不对'%nrow)#l='长度不对'
            sys.exit()#跳出程序
    else:
        pass
def clean_data(str1):
    if isinstance(str1,float):
        str1=str(int(str1))
    return str1.strip()
def upload_file(filepath):
    filesuffix=filepath.split('.')[-1]
    if filesuffix=='xls' or filesuffix=='xlsx':
        return filepath
    else:
        print('文件格式错误,请选择一个excel文件')
        sys.exit()
def get_data(idnum,conn,nrow,worksheet):
    sql="select t3.jdname,t3.pcsname,t3.name,t2.姓名,t2.居民证号,t2.个人联系电话,t6.mc 户籍,t2.服务处所,t1.address 暂住地址,t2.是否注销,t2.上传时间  from (select * from szsrk.web_czfw_sj where 是否最新 = '0') t1,csxsm.cs_zzrk_sj_1 t2,(select * from csxsm.dict_sqxx where sfyx='0') t3,(select flh dm, mc, mnemonic  from dm.wpa_dicts_codes t  where t.dmlb = '24' ) t6 where   t2.出租屋编码 = t1.社区代码 || t1.出租屋编号 and t2.社区代码= t3.url and t2.户籍地址 = t6.dm and t2.居民证号 ='%s'"%idnum
    c=conn.cursor()
    x=c.execute(sql)
    result=x.fetchone()
    j=0
    z=0
    if nrow==1:
        #headlist=[]
        head_index=c.description
        for index in head_index:
            worksheet.write(0,z,index[0])
            z+=1
    if result:
        for res in result:
            worksheet.write(nrow,j,str(res))
            j+=1
    else:
        worksheet.write(nrow,4,idnum)
        worksheet.write(nrow,9,'未登记')
    c.close()
def update_check(filepath):
    try:
        print('选择上传文件')
        conn=cx_Oracle.connect('csxsm/csrk#2014@2.34.202.138:1521/csxsmzj')
        path=upload_file(filepath)#获取需比对数据的文件路径
        workbook1=xlrd.open_workbook(r'%s'%path)
        #print('conn is established')
        workbook2=xlwt.Workbook()#写入的excel
        worksheet=workbook2.add_sheet('sheet1')
        sheetnames=workbook1.sheet_names()#获取所有sheetname
        idlist=[]
        for name in sheetnames:
            sheet=workbook1.sheet_by_name(name)
            if sheet:
                rows=sheet.nrows#获取sheet行数
                for i in range(1,rows):
                    idnum=sheet.cell_value(i,1)
                    idnum=clean_data(idnum)
                    id_check(idnum,i)
                    idlist.append(idnum)
        #print(idlist)
        worksheet_row_num=1
        for ids in idlist:
            get_data(ids,conn,worksheet_row_num,worksheet)
            worksheet_row_num+=1
        workbook2.save('result.xls')
        print('比对完成')
    except Exception as e:
        print(e)
    finally:
        conn.close()
        print('比对关闭')
def search_data_135(sql,filename):#openpyxl的row和column是从1开始
    try:
        conn=cx_Oracle.connect('csxsm/csrk#2014@2.34.202.138:1521/csxsmzj')
        c=conn.cursor()
        x=c.execute(sql)
        results=x.fetchall()
        head_index=c.description
        workbook=openpyxl.Workbook()
        worksheet=workbook['Sheet']
        z=1
        for index in head_index:
            worksheet.cell(row=1,column=z,value=index[0])
            z+=1
        i=2
        for result in results:
            j=1
            for re in result:
                worksheet.cell(row=i,column=j,value=str(re))
                j+=1
            i+=1
        final_name='%s.xlsx'%filename
        workbook.save(final_name)
    except Exception as e:
        print(e)
    finally:
        c.close()
        conn.close()
def search_house():
    try:
        sql="select t4.jdname,t4.pcsname,t4.name,t1.出租屋编号,t2.mc 房屋类别,t1.出租人姓名,t1.联系电话,t6.mc 隐患,t1.address from (select * from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 != '200') t1,(select flh dm, mc from dm.wpa_dicts_codes where dmlb = '53' ) t2, (select * from csxsm.dict_sqxx where sfyx = '0') t4, (select flh dm, mc  from dm.wpa_dicts_codes  where dmlb = '404') t6,(select 出租屋编码 from csxsm.cs_czfw_sj_1 where 是否注销 = '0' minus select 出租屋编码 from csxsm.cs_zzrk_sj_1 where 是否注销 = '0' minus select 出租屋全码 from csxsm.cs_czfw_sj_room s where 是否注销 = '0' ) t7 where t1.社区代码 = t4.url  and trim(t1.房屋类别) = t2.dm  and t1.YHQDLX = t6.dm and t1.社区代码 || t1.出租屋编号 = t7.出租屋编码 order by t4.jdname, t4.pcsname, t4.name"
        filename='无人户房屋'
        search_data_135(sql,filename)
        print('search is done')
    except Exception as e:
        print(e)
def login():
    user=input('请输入用户名:')
    password=input('请输入密码:')
    if re.match(r'^\d{5}$',password):
        if user=='admin' and password=='60476':
            print('correct')
            return True
    print('Wrong username or password')
    return False
def repeate_push_search():
    num=input('注意：小于5时，查询等于重复次数的数据；大于等于5时，查询大于等于重复次数的数据\n请输入重复次数：\n')
    try:
        n=int(num)
        if n<1:
            print('请输入大于1的数')
        elif n<5:
            sql="select b.jdname,b.pcsname,a.czwbh,a.address,a.czrs 人数,a.次数 重复次数 from (select czwbh,address,czrs,sqdm,count(czwqm) 次数 from szsrk.cs_rf_other_house t where  to_char(operatedate,'yyyy')=to_char(sysdate,'yyyy') group by czwbh,address,czrs,sqdm having count(czwqm)=%d) a left join (select  * from csxsm.dict_sqxx where sfyx='0') b on a.sqdm=b.url"%n
            filename='重复推送%d次'%n
        elif n>4:
            sql="select b.jdname,b.pcsname,a.czwbh,a.address,a.czrs 人数,a.次数 重复次数 from (select czwbh,address,czrs,sqdm,count(czwqm) 次数 from szsrk.cs_rf_other_house t where  to_char(operatedate,'yyyy')=to_char(sysdate,'yyyy') group by czwbh,address,czrs,sqdm having count(czwqm)>=%d) a left join (select  * from csxsm.dict_sqxx where sfyx='0') b on a.sqdm=b.url"%n
            filename='重复推送%d次以上'%n
            #print(sql)
        search_data_135(sql,filename)
        print('search is done')
    except Exception as e:
        print(e)
        print('请输入数字')
def town_house():
    sql="select t4.jdname, t4.pcsname,t4.name,t1.出租屋编号,t2.mc 房屋类别,case t1.wqyzd1 when '1' then '红色' when '2' then '黄色' when '3' then '绿色(有待租)' when '3-1' then '绿色(无待租)' when '4' then '蓝色(同意发布)'  when '5' then '不同意发布'  when '0' then '橙色(隐患带推送)' else '' end 五色,t1.实际出租人姓名,t1.实际出租人证件号码,t1.实际出租人联系电话,t1.address,(select count(1) from csxsm.cs_zzrk_sj_1 t where 是否注销 = '0' and t.社区代码 = t1.社区代码 and t.出租屋编号 = t1.出租屋编号) 人数 from (select 出租屋编号,实际出租人姓名,实际出租人证件号码,实际出租人联系电话,address,社区代码,房屋类别,wqyzd1 from szsrk.web_czfw_sj where 是否最新 = '0' and 变动原因 in ('100','300','400') ) t1,(select flh dm, mc from dm.wpa_dicts_codes where dmlb = '53') t2,(select * from csxsm.dict_sqxx where sfyx = '0') t4 where t1.社区代码 = t4.url and trim(t1.房屋类别) = t2.dm order by t4.jdname, t4.pcsname, t4.name"
    filename='全市出租屋明细'
    search_data_135(sql,filename)
    print('search is done')
def work_order_push():
    sq11="select s.jdname,s.pcsname,t.address as 地址,t.czwbh as 出租屋编号,t.operatedate 操作日期,case r.state when '0' then '新采集' when '1' then '已上传' when '2' then '已接收' when '3' then '已处理' when '4' then '已核查' else '' end  as 处置状态,case r.dalx when '1' then '红色' when '2' then '黄色' else ''  end  as 档案类型,t.czrs 人数,t.czwqm,r.taskguid as 工单号 from szsrk.cs_rf_other_house t,szsrk.cs_rfotherhouse_report r,(select * from csxsm.dict_sqxx where sfyx='0') s where t.djbh = r.djbh and r.sfyx = '0' and t.sqdm=s.url and to_char(t.operatedate,'yyyy')=to_char(sysdate,'yyyy') and r.state<4 order by s.jdname,s.pcsname,t.operatedate"
    filename1='推送明细'
    search_data_135(sq11,filename1)
    print('search is done')
def work_order_finish():
    sql2="select s.jdname,s.pcsname,t.address as 地址,t.czwbh as 出租屋编号,t.operatedate 操作日期,case r.state when '0' then '新采集' when '1' then '已上传' when '2' then '已接收' when '3' then '已处理' when '4' then '已核查' else '' end  as 处置状态,case r.dalx when '1' then '红色' when '2' then '黄色' else ''  end  as 档案类型,t.czrs 人数,t.czwqm,r.taskguid as 工单号 from szsrk.cs_rf_other_house t,szsrk.cs_rfotherhouse_report r,(select * from csxsm.dict_sqxx where sfyx='0') s where t.djbh = r.djbh and r.sfyx = '0' and t.sqdm=s.url and to_char(t.operatedate,'yyyy')=to_char(sysdate,'yyyy') and r.state=4 order by s.jdname,s.pcsname,t.operatedate"
    filename2='整治明细'
    search_data_135(sql2,filename2)
    print('search is done')
def rx_read_excel(filepath):     #读取总名单数据
    try:
        sheet_list=['备案','未备案','未登记']
        conn=cx_Oracle.connect('csxsm/csrk#2014@2.34.202.135:1521/csxsm')    #连接数
        _path=upload_file(filepath)
        data=openpyxl.load_workbook(_path)   #打开excel 注意只能打开xlsx格式
        sheetnames=data.sheetnames#获取原文件表名
        for name in sheet_list:
            if name not in sheetnames:
                data.create_sheet(name)       #判断有无3张表，没有则添加
        ba=data['备案']
        wba=data['未备案']
        wdj=data['未登记']
        barow=ba.max_row
        wbarow=wba.max_row
        wdjrow=wdj.max_row
        print(barow,wbarow,wdjrow)
        if wbarow>1:
            wbarow+=1
        if barow>1:
            barow+=1
        if wdjrow>1:
            wdjrow+=1
        zmd=data['Sheet1']   #读取总名单信息
        zmd_rows=zmd.max_row                   #获取行数
        print('总行数',zmd_rows)
        for i in range(1,zmd_rows):
            _row=[]
            row_data=list(list(zmd.rows)[i])  #获取行数据****注意 虽然openpyxl是从1开始，但是list还是从0开始
            for r in row_data:#对行数据进行处理，将元素转化为str
                _row.append(r.value)
            print('处理第',i,'行')
            sfzlist=[_row[3],_row[7],_row[11]]
            #print(type(row[3]),type(row[7]),type(row[11]))   
           # apply_aysnc(judge_sfba,sfzlist,callback=partial(check,row=_row))
            res=rx_judge_sfba(conn,*sfzlist)
            if res==9:
                #print('未登记')
                j = 1
                for a in _row:
                    wdj.cell(wdjrow, j, str(a))
                    j += 1
                wdjrow += 1
            elif res==0:
                #print('备案')
                j = 1
                for a in _row:
                    ba.cell(barow, j, str(a))
                    j += 1
                barow += 1    #放在sheet1
            else:
                #print('不备案')
                newlist=rx_wba_addpcs(conn,*sfzlist)
                #print('newlist为：',newlist)
                #print('newlist类型',type(newlist))
                #print('row是',row)
                #print('row类型',type(row))
                #print('增加派出所社区信息。')
                _row.extend(newlist)   #增加派出所社区信息。
                #print(row)
                j = 1
                for a in _row:
                    wba.cell(wbarow, j, str(a))
                    j += 1
                wbarow += 1
        data.save('result.xlsx')
        print('mission complete！')
    except Exception as e:
        print(e)
    finally:
        conn.close()
def rx_wba_addpcs(conn,s1,s2,s3):
    sql="select t1.居民证号,t4.pcsname,t4.name,t1.出租屋编号,t1.房间号,t3.户编号,t3.备案状态,t5.上传时间 from csxsm.cs_zzrk_sj_1 t1,csxsm.cs_czfw_sj_room t2,csxsm.cs_czfw_sj_hu t3,(select *from dm.ct_dict_sqxx_cs where sfyx='0') t4,csxsm.cs_czfw_sj_1 t5  where t2.是否注销='0' and t3.是否注销='0' and t1.是否注销='0' and t2.户号全码=t3.户全码 and t1.出租屋编码 =t2.出租屋全码 and t1.房间号=t2.房间编号 and t1.社区代码 =t4.url and t1.出租屋编码=t5.出租屋编码 and t1.居民证号 in ('%s','%s','%s')"%(s1,s2,s3)  
    c=conn.cursor()        #获取cursor   
    x=c.execute(sql)
    result=x.fetchall()
    list=[]
    #print(result)
    sfz=[s1,s2,s3]
       #print(sfz)
    b=['','','','','','','']
    if result!=None:
            #print(sfz[i])
        for s in sfz:
            i=0
            for a in result: 
                if s == a[0]:                   
                    list.extend(a[1:])
                    break
                i+=1
                if i==len(result):
                    list.extend(b)
                    #print(list)
                    #print(i)
    else:
        print('error')
    c.close()         #关闭cursor 
        #print(list)
    return list 
def rx_judge_sfba(conn,s1,s2,s3):     #判断备案情况
    try:   #print('开始judge')
        c=conn.cursor()        #获取cursor   
        #print(s1,s2,s3)
        #sql='select * from dual'
        sql="select t3.备案状态 from csxsm.cs_zzrk_sj_1 t1,csxsm.cs_czfw_sj_room t2,csxsm.cs_czfw_sj_hu t3 where t2.是否注销='0' and t3.是否注销='0' and t1.是否注销='0' and t2.户号全码=t3.户全码 and t1.出租屋编码 =t2.出租屋全码 and t1.房间号=t2.房间编号 and t1.居民证号 in ('%s','%s','%s')"%(s1,s2,s3)
        #print(sql)
        x=c.execute(sql)  		#使用cursor进行各种操作
        #print(x.rownumber)
        result=x.fetchall()
        if result!=None:
            #print(result,len(result))
            if len(result)==0:
                return 9       #均未查询到信息，未登记
            else:
                sum=0
                for i in result:
                    sum=sum+int(i[0])   #备案状态（0、1）相加 
                if sum<len(result):
                    return 0  #备案
                else:
                    return 1    #不备案            
            
        else:
            print('无效')        
    except Exception as e:
        print(e)
    finally:
        c.close()
def rx_id_check(filepath):
    _path=upload_file(filepath)
    workbook1 = openpyxl.load_workbook(_path)
    zmd = workbook1['Sheet1']  # 读取总名单信息
    bdjg=workbook1.create_sheet('比对结果')
    bdjg.append([])
    zmd_rows = zmd.max_row  # 获取行数
    print('总行数', zmd_rows)
    for i in range(1, zmd_rows):
        row_data = list(list(zmd.rows)[i])
        _row=[]
        for r in row_data:
            _row.append(r.value)
        sfzlist=[_row[3],_row[7],_row[11]]
        #print(sfzlist)
        l=[]
        for sfz in sfzlist:
            if sfz is not None:
                s = list(sfz)
                count = 0
                if len(s) == 18:
                    for i in range(17):
                        count += int(loops[i]) * int(s[i])
                        # print(count)
                    # print(count)
                    result = str(checknum[count % 11])
                    # print(result)
                    # print(s[17])
                    if result == str(s[17]):
                        l.append('正确')
                    else:
                        l.append('错误')
                else:
                    l.append('长度不对')
            else:
                l.append('身份证为空')
        #print(l)
        bdjg.append(l)#整行写入
    workbook1.save("idcheck.xlsx")
if __name__=='__main__':
    for i in range(3):
        if login():
            while True:
                case=input('------------------------\n请选择：\n1.上传比对\n2.无人户房屋查询\n3.重复推送查询\n4.全市出租屋明细\n5.推送工单查询\n6.整改工单查询\n7.民办学校信息比对\n8.民办学校身份证比对\n0.退出\n')
                if case=='1':
                    root=tkinter.Tk()#tkinter如果要像此例一样循环，则需放在线程外，否则会报一个main thread is not in main loop 错误
                    root.withdraw()
                    filepath=filedialog.askopenfilename()
                    thr1=threading.Thread(target=update_check,args=(filepath,))
                    thr1.deamon=True#daemon为True，就是我们平常理解的后台线程，用Ctrl-C关闭程序，所有后台线程都会被自动关闭。如果daemon属性是False，线程不会随主线程的结束而结束，这时如果线程访问主线程的资源，就会出错。
                    thr1.start()
                    thr1.join()
                elif case=='2':
                    thr2=threading.Thread(target=search_house)
                    thr2.start()
                    thr2.join()
                elif case=='3':
                    thr3=threading.Thread(target=repeate_push_search)
                    thr3.start()
                    thr3.join()
                elif case=='4':
                    thr4=threading.Thread(target=town_house)
                    thr4.start()
                    thr4.join()
                elif case=='5':
                    thr5=threading.Thread(target=work_order_push)
                    thr5.start()
                    thr5.join()
                elif case=='6':
                    thr6=threading.Thread(target=work_order_finish)
                    thr6.start()
                    thr6.join()
                elif case=='7':
                    print('注意：\n1.文件格式为xlsx\n2.确保比对数据保存在Sheet1中\n3.确保身份证号码在D、H、L列\n4.清先校验身份证是否正确')
                    root = tkinter.Tk()
                    root.withdraw()
                    filepath = filedialog.askopenfilename()  # 获取文件路径
                    thr7=threading.Thread(target=rx_read_excel,args=(filepath,))
                    thr7.start()
                    thr7.join()
                elif case=='8':
                    print('注意：\n1.文件格式为xlsx\n2.确保比对数据保存在Sheet1中\n3.确保身份证号码在D、H、L列')
                    root = tkinter.Tk()
                    root.withdraw()
                    filepath = filedialog.askopenfilename()  # 获取文件路径
                    thr8 = threading.Thread(target=rx_id_check,args=(filepath,))
                    thr8.start()
                    thr8.join()
                elif case=='0':
                    break
                else:
                    print('请输入0-6之间的数字')
            print('再会')
            sys.exit()
    print('只有3次机会哦')