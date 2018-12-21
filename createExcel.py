import xlwt
import os
import datetime
# 初始化表头列表
tb_head = ['合同签订日期:',
               '供地方式:',
               '项目位置:',
               '行政区:',
               '面积(公顷):',
               '土地用途:',
               '行业分类:',
               '电子监管号：',
               '土地级别:',
               '土地来源:',
               '批准单位:',
               '约定开工时间:',
               '约定竣工时间:',
               '成交价格(万元):',
               '约定交地时间:',
               '项目名称:',
               '土地使用年限:',
               '土地使用权人',#最后六个属性不一样
                '分期数',
               '约定支付日期',
                '约定支付金额',
                '下限:',
                '上限:',
               ]
#创建表头
def write_excel():
    # 创建一个workbook，设置编码
    workbook = xlwt.Workbook(encoding='utf-8')
    # 创建一个worksheet
    worksheet = workbook.add_sheet('sheet1', cell_overwrite_ok=True)
    for item in tb_head:
        worksheet.write(0,tb_head.index(item), item)
        worksheet.col(tb_head.index(item)).width = 256 *25
    return workbook

#创建列表
def create_excel(startYear,endYear,startMonth,startDay):
    for i in range(startYear,endYear+1):
        # 数据存储文件的位置
        print(i)
        EXCEL_NAME = 'res_data_landchina' + str(i) + '.xls'

        filelist = os.listdir('./')
        # 判断文件是否存在
        if EXCEL_NAME not in filelist:
            # 创建存储文件,表头
            wb = write_excel()
            table = wb.get_sheet(0)
            # 上一次存储时，历史记录表
            wb.add_sheet('historyNum')
            table1 = wb.get_sheet('historyNum')
            # 初始记录条数
            table1.write(0, 0, 1)
            # 初始日期
            startDate = datetime.date(i,startMonth,startDay).strftime('%Y-%m-%d')
            table1.write(0, 1, startDate)
            # 初始页面
            table1.write(0, 2, 1)
            wb.save(EXCEL_NAME)

startYear=2013
endYear=2015
startMonth=4
startDay=1
create_excel(startYear,endYear,startMonth,startDay)
