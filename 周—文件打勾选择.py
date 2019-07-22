import docx2txt             #docx转txt
import re
import xlrd
import os
from xlutils.copy import copy       #操作，写入Excel

#获取文件夹下面的所有文件路径
def lujing(lu):
    lujings = []
    #root 当前路径，dirs是路径下的所有子目录，files子文件
    for root,dirs,files in os.walk(lu):
        #files是列表需要重新读取
        for i in files:
            #加根目录
            a = root +'\\' +  i
            lujings.append(a)
    #返回一个列表
    return lujings

#总表，用来提取关键信息
def zong(lus):
    #获得路径
    zonglujing = lujing(lus)
    #挨个读取路径
    for i in zonglujing:
        #用路径来读取
        text = docx2txt.process(i)
        try:
            #加一个u就可以匹配中文                                          不区分\n
            neirong = re.findall(u"应用权限安全.*?权限：(.*?)风险系数",text,re.DOTALL)[-1]
            a = re.search(u'应用名称:(.*?)系统:', text, re.DOTALL).group(1)
            name = a.split()
            #调用函数
            dagou(name,neirong)
        #报错提示
        except:
            print('报错' + i)

#打勾程序
def dagou(a,xiao):
    #转列表，去除空列表
    xiao_a = list(filter(None,xiao.split('\n')))
    #开始匹配
    for i in xiao_a:
        if i in lie:
            #返回打勾位置
            #col是竖着，row是横着
            shu_b = lie.index(i) + 1
            shu_a = hang.index(a[0]) + 1
            sheets.write(shu_a,shu_b,'V')
        books.save(r"C:\Users\阿苗\Desktop\aaa.xls")


if __name__ == '__main__':
    #打开文件夹
    book = xlrd.open_workbook(r"C:\Users\阿苗\Desktop\aa.xls")
    sheet = book.sheet_by_name('aa')
    lie = sheet.row_values(0)[1:]
    hang = sheet.col_values(0)[1:]
    # 写入Excel
    books = copy(book)
    sheets = books.get_sheet(0)  # 新建表格
    zong(r"C:\Users\阿苗\Desktop\报告(1)")

