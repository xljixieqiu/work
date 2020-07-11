import xlrd
import os
import xlwt
import shutil
import cx_Oracle
conn=cx_Oracle.connect('admin/pass@127.0.0.1:1521/xx')    #连接数据库
c=conn.cursor()                                                 #获取cursor
newexcel=xlwt.Workbook()                                         #新建Excel表
newsheet=newexcel.add_sheet('sheet1')                             #新建sheet
sheetrow=0       
try:                                                #设置从第一行开始记录
    def con_oracle(pcs,*seq):
        bh="'"+"','".join(*seq)+"'"
        #print(bh) 
        x=c.execute("select count(1) from csxsm.cs_zzrk_sj_1 t1, csxsm.cs_czfw_sj_1 t2,(select distinct pcsname,substr(pcsdm, -3) dm from dm.ct_dict_sqxx_cs where sfyx='0' ) t4 where t1.是否注销 = '0' and t2.是否注销 = '0' and t1.出租屋编码 = t2.出租屋编码  and t2.出租屋地址派出所 =t4.dm and t4.pcsname like '%"+pcs+"%' and t2.出租屋编号 in ( "+bh+")")                         #使用cursor进行各种操作
        data=x.fetchone()
        print(data)
        return data
    def open_xls(roots):
        global sheetrow                            #获取行数。注：如果global不是位于函数的顶部,而是使用或赋值变量之后,会有警告
        sheetcol=0                                 #sheet 列数
        xx=roots.split('/')    #将路径以'/'分割，获取派出所name和人员name
        pcs=xx[-2]
        name=xx[-1]
        data=xlrd.open_workbook(roots,formatting_info=True)
        table=data.sheets()[0]
        print('行数:%s'%table.nrows)
        print('列数:%s'%table.ncols)
        strs=table.row_values(1)[0].strip()[:2]
        print(strs)
        list=[]
        for i in range(5,table.nrows):    #读取出租屋编号依次存入list
            for j in range(table.ncols):
                cell=table.cell_value(i,j)
                ctype=table.cell(i,j).ctype
                #print('第%s行第%s列的值为%s,数值类型为%s'%(i,j,cell,ctype))
                if ctype==2:
                    cell=str(int(cell))
                list.append(cell)
        p=strs+name
        num=con_oracle(strs,list)
        print(p,num)
        newsheet.write(sheetrow,sheetcol,p)
        sheetcol+=1
        newsheet.write(sheetrow,sheetcol,str(num))               #写进Excel的值必须为string
        sheetrow+=1
    #newexcel.save(p) 
    #shutil.move(p,os.path.join('d:/python work/1',p))
    for root,dirs,files in os.walk(r"G:\2019总决赛"):                 #遍历文件夹下的文件
        for name in dirs:
            print ('文件路夹径为：',os.path.join(root,name))
        for name in files:
            print ('文件路径为：',os.path.join(root,name))
            paths=os.path.join(root,name).replace('\\','/')       #把\转换为/，因为/转义的原因，识别不了路径
            print('wenjian lujingwei:',paths)
            open_xls(paths)
#open_xls('G:/2019比赛/莫城/尤祎嘉（甲）.xls')
    newexcel.save('result.xls')                                        #保存Excel
    shutil.move('result.xls','d:/python work/1/result.xls')           #剪切到1文件夹下
except Exception as e:
    print('Error:',e)
finally:
    c.close()                                                       #关闭cursor
    conn.close()                                                    #关闭连接
