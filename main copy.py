import xlrd
import xlwt
import os

#单子实例
class Item:
    def __init__(self, raw_data):
        if len(raw_data) == 8 and len(str(raw_data[0])) == 10:
            self.id = str(raw_data[0])[:-2]
            self.year = self.id[:4]
            self.mon = self.id[4:6]
            self.section = raw_data[1]
            self.subsection = raw_data[2]
            self.abstract = raw_data[3]
            self.debit = raw_data[4]
            self.credit = raw_data[5]
            self.balance = raw_data[6]
            self.note = raw_data[7]
            self.raw_data = raw_data[:]
        else:
            print('**************\n{}\n**************\n'.format(raw_data))
            self.id = None
    def place(self, section, coord):
        self.deco = '%s!E%s'%(section,coord+1)
        self.creco = '%s!F%s'%(section,coord+1)
    def show(self):
        print('''id:{}\n
        date:{}/{}\n
        section:{}\n
        subsection:{}\n
        abstract:{}\n
        debit:{}\n
        credit:{}\n
        balance:{}\n
        note:{}\n'''.format(self.id,self.year,self.mon,self.section,self.subsection,self.abstract,self.debit,self.credit,self.balance,self.note))

#数据库实例
class Database:
    def __init__(self):
        self.items = []
        self.sections = []
        self.months = []
        self.exception = []
    def append(self, item):
        if item.id == None:
            self.exception.append(item)
        else:
            self.items.append(item)
            if item.section in self.sections:
                pass
            else:
                self.sections.append(item.section)
            if item.year+'/'+item.mon in self.months:
                pass
            else:
                self.months.append(item.year+'/'+item.mon)
    def sec(self, section):
        ls = []
        for item in self.items:
            if section == item.section:
                ls.append(item)
        return ls
    def subsec(self, subsection, database):
        ls = []
        for item in database:
            if subsection == item.subsection:
                ls.append(item)
        return ls
    def mon(self, mon, database):
        ls = []
        for item in database:
            if (item.year + item.mon) == mon:
                ls.append(item)
        return ls
    def subsub(self, abstract, database):
        ls = []
        for item in database:
            if abstract == item.abstract:
                ls.append(item)
        return ls




#####################################################################################################
path = os.getcwd()

for f in os.listdir():
    file_name, file_format = f.split('.')
    if file_format == 'xls' or file_format == 'xlsx':
        if file_name[-2:] == 'py':
            continue
        else:
            file_full = f
            break
    else:
        continue

data = xlrd.open_workbook(file_full)#读取表格

sheet = data.sheet_by_name('现金账')#获取原数据

head = sheet.row_values(0)

print(head)#打印表头

#初始化数据库a
a = Database()
for i in range(sheet.nrows):
    #print(sheet.row_values(i))
    a.append(Item(sheet.row_values(i)))
    #a.items[-1].show()
print('导入完成')
datawt = xlwt.Workbook(encoding='utf-8')
'''
debit = []
credit = []           
for item in a.items:
    #print(item.raw_data)
    if bool(item.debit):
        debit.append(item.debit)
    if bool(item.credit):
        credit.append(item.credit)
print('支出共{}元\n收入共{}元\n余额共{}元\n'.format(sum(credit),sum(debit),sum(debit)-sum(credit)))
'''

pattern1 = xlwt.Pattern()
pattern1.pattern_fore_colour = 45
pattern2 = xlwt.Pattern()
pattern2.pattern_fore_colour = 50
pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
style1 =xlwt.XFStyle()
style2 =xlwt.XFStyle()
style1.pattern=pattern1
style2.pattern=pattern2
##############################################

############################################
print('正在写入……')
sheet_total = datawt.add_sheet('管理总表')
sec_num = 1
tr = 1
tc = 2
k = 0
for month in a.months:
    sheet_total.write(0,tc+k,label=month)
    k += 1
############################################
for section in a.sections:
    sheet_total.write(tr,0,label=sec_num)
    sheet_total.write(tr,1,label='%s'%section)
    tr += 1
    subsec_num = 1
    subsections = []


    print('{}分表写入中……'.format(section))
    sheet_now = datawt.add_sheet(section)
    order = [0,1,2,3,5,4,6,7]
    debit = order.index(5)
    credit = order.index(4)
    for j in range(len(head)):
        sheet_now.write(0,j,label=head[j])
    sec_items = a.sec(section)
    r = 1
    mon = None
    for item in sec_items:
        if mon == (item.year + item.mon):
            pass
        else:
            mon = (item.year + item.mon)
            mon_items = a.mon(mon, sec_items)
            for mon_item in mon_items:
                k = 0
                mon_item.place(section,r)
                for c in order:
                    sheet_now.write(r,c,label=mon_item.raw_data[k])
                    k += 1
                r += 1
            sheet_now.write(r,debit, xlwt.Formula('SUM(%c%d:%c%d)'%(65+debit,r-len(mon_items)+1,65+debit,r)),style1)
            sheet_now.write(r,credit, xlwt.Formula('SUM(%c%d:%c%d)'%(65+credit,r-len(mon_items)+1,65+credit,r)),style2)
            r += 1
        if item.subsection in subsections:
            pass
        else:
            subsections.append(item.subsection)


    print('{}\n汇入管理总表中……'.format(subsections))
    
    for subsection in subsections:
        sheet_total.write(tr,0,label='%s.%s'%(sec_num,subsec_num))
        sheet_total.write(tr,1,label=subsection)
        subsubs = []
        
        subsubsec_num = 1
        tc = 2
        sub_items = a.subsec(subsection, sec_items)
        for month in a.months:
            mon_items = a.mon(''.join(month.split('/')), sub_items)
            content = []
            for item in mon_items:
                if bool(item.debit):
                    content.append(item.creco)
                    style = style2
                else:
                    content.append(item.deco)
                    style = style1
                if item.abstract in subsubs:
                    pass
                else:
                    subsubs.append(item.abstract)
            content = '+'.join(content)
            #print(content+'\n')
            #print(isinstance(content,str))
            if content == '':
                tc += 1
                continue
            sheet_total.write(tr,tc, xlwt.Formula(content),style)
            tc += 1
        tr += 1
        for subsub in subsubs:
            sheet_total.write(tr,0,label='%s.%s.%s'%(sec_num,subsec_num,subsubsec_num))
            sheet_total.write(tr,1,label=subsub)
            subsubsec_num += 1
            tc = 2
            subsub_items = a.subsub(subsub, sub_items)
            for month in a.months:
                mon_items = a.mon(''.join(month.split('/')), subsub_items)
                content = []
                for item in mon_items:
                    if bool(item.debit):
                        content.append(item.creco)
                        style = style2
                    else:
                        content.append(item.deco)
                        style = style1
                content = '+'.join(content)
                #print(content+'\n')
                #print(isinstance(content,str))
                if content == '':
                    tc += 1
                    continue
                sheet_total.write(tr,tc, xlwt.Formula(content),style)
                tc += 1
        
            tr += 1
        subsec_num += 1            

    sec_num += 1
            


sheet_copy = datawt.add_sheet('现金账')
print('现金账复制中……')
for i in range(sheet.nrows):
    value = sheet.row_values(i)
    for k in range(len(value)):
        sheet_copy.write(i,k,label=value[k])
print('现金账复制完成')



    
        







datawt.save(file_name+'py.xls')

print("写入完成")
input('按任意键结束……')

sheet.row_values(1)
