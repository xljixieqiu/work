import tkinter,xlrd,xlwt,cx_Oracle
from tkinter import filedialog
def id_check(arg,nrow):
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
def clean_data(str1):
    if isinstance(str1,float):
        str1=str(int(str1))
    return str1.strip()
def upload_file():
    root=tkinter.Tk()
    root.withdraw()
    filepath=filedialog.askopenfilename()
    return filepath
def get_data(id,conn,nrow):
    sql="select t3.jdname,t3.pcsname,t3.name,t2.姓名,t2.居民证号,t2.个人联系电话,t6.mc 户籍,t2.服务处所,t1.address 暂住地址,t2.是否注销,t2.上传时间  from (select * from szsrk.web_czfw_sj where 是否最新 = '0') t1,csxsm.cs_zzrk_sj_1 t2,(select * from csxsm.dict_sqxx where sfyx='0') t3,(select flh dm, mc, mnemonic  from dm.wpa_dicts_codes t  where t.dmlb = '24' ) t6 where   t2.出租屋编码 = t1.社区代码 || t1.出租屋编号   and t2.社区代码 = t3.url   and t2.户籍地址 = t6.dm and t2.居民证号='%s'"%idnum
    c=conn.cursor()
    x=c.execute(sql)
    restult=x.fechall()
    j=0
    if result:
        for res in result[0]:
            worksheet.write(nrow,j,str(res))
            j+=1
    else:
        worksheet.write(nrow,4,idnum)
        worksheet.write(nrow,9,'未登记')
def main():
    try:
        path=upload_file()#获取需比对数据的文件路径
        loops=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]#idcheck所需
        checknum=[1,0,'X',9,8,7,6,5,4,3,2]#idcheck所需
        workbook1=xlrd.open_workbook(r'%s'%path)
        conn=cx_Oracle.connect('user/password@server:post/pid')
        workbook2=xlwt.Workbook()#写入的excel
        worksheet=workbook2.add_sheet('sheet1')
        sheetnames=workbook1.sheet_names()#获取所有sheetname
        for name in sheetnames:
            sheet=workbook1.sheet_by_name(name)
            if sheet:
                rows=sheet.nrows#获取sheet行数
                for i in range(1,rows):
                    idnum=sheet.cell_value(i,1)
                    idnum=clean_data(idnum)
                    id_check(idnum,i)
                    get_data(idnum,conn,i)
        workbook2.save('result.xls')
    except Exception as e:
        print(e)
    finally:
        conn.close()
        print('程序关闭')
if __name__=='__main__':
    main()