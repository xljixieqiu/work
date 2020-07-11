import cx_Oracle   #引用模块cx_Oracle
import xlrd
import xlwt                 #excel的write只能在新建excel后再写入数据
from xlutils.copy import copy   
#xlrd.Book.encoding = "utf-8"
conn=cx_Oracle.connect('admin/password@127.0.0.1:1521/xx')    #连接数据库
data=xlrd.open_workbook('e:/Users/admin/Desktop/2019民办学校信息/2020春/第二次比对名单.xls')   #打开excel 需要添加'备案''未备案''未登记'
newexcel=copy(data)                                        #copy Excel表
ba=newexcel.get_sheet(1)                             #备案sheet
wba=newexcel.get_sheet(2)                             #未备案sheet
wdj=newexcel.get_sheet(3)                             #未处理sheet
barow=data.sheet_by_name('备案').nrows
wbarow=data.sheet_by_name('未备案').nrows
wdjrow=data.sheet_by_name('未登记').nrows
print(barow,wbarow,wdjrow)
try: 
    def move(name,*arg):   #移动数据
        '''
        sheetnames=data.sheet_names()   #读取sheetname
        if name in sheetnames:         #判断是否存在对应的sheet
            wba=data.sheet_by_name(name)
        else:
            wba=data.add_sheet(name)
        '''
        global barow,wbarow,wdjrow
        j=0           #列从0开始
        #print('开始move')
        if name=='备案':
            for a in arg[0]:
                #print(a)
                ba.write(barow,j,str(a))
                j=j+1
            barow=barow+1
        elif name=='未备案':
            for a in arg[0]:
                #print(a)
                wba.write(wbarow,j,str(a))
                j=j+1
            wbarow=wbarow+1
        else:
            for a in arg[0]:
                #print(a)
                wdj.write(wdjrow,j,str(a))
                j=j+1
            wdjrow=wdjrow+1
    def read_excel():     #读取总名单数据
        print('read')
        zmd=data.sheet_by_name(u'比对总名单')   #读取总名单信息
        #print(zmd)
        zmd_rows=zmd.nrows                   #获取行数
        print('总行数',zmd_rows)
        for rows in range(1,zmd_rows):
            row=zmd.row_values(rows)   #获取行数据
            print('处理第',rows,'行')
            znid=row[3].upper()
            fid=row[7].upper()
            mid=row[11].upper()  #把小写x改成X
            #print(type(row[3]),type(row[7]),type(row[11]))   
            res=judge_sfba(znid,fid,mid)
            if res==9:
                print('未登记')
                move('未登记',row)   #放在sheet3
            elif res==0:
                print('备案')
                move('备案',row)    #放在sheet1
            else:
                print('不备案')
                newlist=wba_addpcs(znid,fid,mid)
                #print('newlist为：',newlist)
                #print('newlist类型',type(newlist))
                #print('row是',row)
                #print('row类型',type(row))
                print('增加派出所社区信息。')
                row.extend(newlist)   #增加派出所社区信息。
                #print(row)
                move('未备案',row)   #放在sheet2
           #print('完成')
    def judge_sfba(s1,s2,s3):     #判断备案情况
        #print('开始judge')
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
        c.close()                                                      #关闭cursor  
    def wba_addpcs(s1,s2,s3):
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
                    if s in a:                   
                        list.extend(a[1:])
                        break
                    i=i+1
                    if i==len(result):                    
                        list.extend(b)
                    #print(list)
                    #print(i)
        else:
            print('error')
        c.close()         #关闭cursor 
        #print(list)
        return list                                      
    def main():
        read_excel()
        print('read_excel完成')
        newexcel.save('result.xls')    #保存结果
        print('保存完成')
    if __name__=='__main__':
        main()
except Exception as e:
    print('错误原因：',e)	
finally:
    conn.close()                                                    #关闭连接
