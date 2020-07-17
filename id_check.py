import xlrd
import xlwt                 #excel的write只能在新建excel后再写入数据
from xlutils.copy import copy 
data=xlrd.open_workbook('e:/Users/admin/Desktop/2019民办学校信息/2020春/总名单身份证校验.xls')   #打开excel 需要添加'备案''未备案''未登记'
newexcel=xlwt.Workbook(encoding='utf-8') 
sheet=newexcel.add_sheet('身份证比对结果')
loops=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2]
checknum=[1,0,'X',9,8,7,6,5,4,3,2]
i=0
#idnum='32088219890717421X'
try:
    def id_check(arg):
        s=list(arg)
        l=''
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
                l='正确'
            else:
                l='错误'
        else:
            l='长度不对'
        #print(l)
        return l
    def writestr(args):
        global i
        j=0
        #print('writestr中的args:',args)
        for a in args:
            sheet.write(i,j,str(a))
            j +=1
        i+=1
    def read_excel():
        print('read')
        zmd=data.sheet_by_name(u'比对总名单')   #读取总名单信息
        #print(zmd)
        zmd_rows=zmd.nrows                   #获取行数
        print('总行数',zmd_rows)
        for rows in range(zmd_rows):
            row=zmd.row_values(rows)   #获取行数据
            #print(type(row))
            znid=row[3].upper()
            fid=row[7].upper()
            mid=row[11].upper()  #把小写x改成X
            sflist=[znid,fid,mid]
            resultlist=[]
            n_sflist=range(len(sflist))
            for i in n_sflist:
                t=id_check(sflist[i])
                resultlist.append(t)
            row.extend(resultlist)
            #print('row',row)
            writestr(row)
    if __name__=='__main__':
        read_excel()
        print('read_excel完成')
        newexcel.save('idcheck.xls')    #保存结果
        print('保存完成')
except Exception as e:
    print('错误原因：',e)	
#finally:
    #data.close()
    #newexcel.close()