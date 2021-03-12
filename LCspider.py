from selenium import webdriver
#引入显示等待
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.chrome.options import Options
import time
import datetime
from xlutils3 import copy
import xlrd

import traceback, sys

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
                'adcode'
               ]

def LandChina(EXCEL_NAME):
    # 打开表格
    wb = xlrd.open_workbook(EXCEL_NAME)
    #历史日期
    historyDate = datetime.datetime.strptime(wb.sheet_by_name('historyNum').cell_value(0,1),'%Y-%m-%d')
    oneday=datetime.timedelta(days=1)
    start = historyDate.strftime('%Y-%m-%d')

    
    chormedriver='chromedriver.exe'
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    #browser是自己定义的
    
    browser =webdriver.Chrome(chormedriver,chrome_options=chrome_options)

    #日期增加
    for i in range(0,365):
        tag=1#标记用来循环执行异常
        index=2#等待指数
        while tag:
            try:
                print("-----开始加载-----")
                url = 'http://www.landchina.com/default.aspx?tabid=263&wmguid=75c72564-ffd9-426a-954b-8ac2df0903b7&p=9f2c3acd-0256-4da2-a659-6949c4671a2a%3A' \
                      + start + '~' + start
                browser.get(url)
                WebDriverWait(browser, 30).until(lambda browser: browser.find_elements_by_class_name('gridItem'))

                if len(browser.find_elements_by_class_name('pager'))!=0:
                    subTitle = WebDriverWait(browser, 30).until(
                        lambda browser: browser.find_elements_by_class_name('pager'))
                    # 一天中所有公告结果的循环,翻页
                    pageNum = int(subTitle[1].text.split()[0][1:-1])
                    print('一共有' + str(pageNum) + '页')
                else:
                    pageNum=1
                    print('不满一页')
            except Exception as e:
                print('加载出错！'+repr(e))
                browser.close()
                
                browser = webdriver.Chrome(chormedriver,chrome_options=chrome_options)
                print('等待' + str(index) + 's')
                time.sleep(index)
                index=index+10
            else:
                tag=0
                # 打开表格
                wb = xlrd.open_workbook(EXCEL_NAME)
                #历史页数
                historyPage=int(wb.sheet_by_name('historyNum').cell_value(0,2))
                print('历史页'+str(historyPage))
                for page in range(historyPage,pageNum+1):
                    tag = 1
                    index = 2  # 等待指数
                    while tag:
                        try:
                            if len(browser.find_elements_by_class_name('pager'))!=0:
                                print('-----开始翻页，第'+str(page)+'页-----')
                                subTitle = WebDriverWait(browser,30).until(lambda browser:browser.find_elements_by_class_name('pager'))
                                inputPage = subTitle[2].find_elements_by_tag_name('input')[0]
                                inputPage.clear()
                                inputPage.send_keys(str(page))
                                subTitle[2].find_elements_by_tag_name('input')[1].click()
                            trlList1 = WebDriverWait(browser,30).until(lambda browser:browser.find_elements_by_class_name('gridItem'))
                            trlList2 = browser.find_elements_by_class_name('gridAlternatingItem')
                            if len(trlList2)!=0:
                                trlList2 = WebDriverWait(browser,30).until(lambda browser:browser.find_elements_by_class_name('gridAlternatingItem'))
                            trlList1.extend(trlList2)  # 合并列表
                        except Exception as e:
                            print('翻页出错！'+repr(e))
                            browser.close()
                            
                            browser = webdriver.Chrome(chormedriver,chrome_options=chrome_options)
                            url = 'http://www.landchina.com/default.aspx?tabid=263&wmguid=75c72564-ffd9-426a-954b-8ac2df0903b7&p=9f2c3acd-0256-4da2-a659-6949c4671a2a%3A' \
                                  + start + '~' + start
                            browser.get(url)
                            print('等待' + str(index) + 's')
                            time.sleep(index)

                            index = index + 10
                        else:
                            tag=0
                            # 打开表格,一页一页的存储
                            wb = xlrd.open_workbook(EXCEL_NAME)
                            # 记录历史条数
                            recordNum = int(wb.sheet_by_name('historyNum').cell_value(0, 0))
                            wb = copy.copy(wb)  # 拷贝一份原来的excel
                            table = wb.get_sheet(0)

                            print('-----一页有'+str(len(trlList1))+'条数据------')

                            for tr in trlList1:
                                tag = 1
                                index = 2  # 等待指数
                                while tag:
                                    try:
                                        print('-----读取数据-----')
                                        print(start)
                                        url =tr.find_element_by_tag_name('a').get_attribute('href')
                                        print(tr.find_element_by_class_name('gridTdNumber').text)
                                        new_windows = 'window.open("'+url+'");'

                                        browser.execute_script(new_windows)
                                        #隐式等待1min
                                        # browser.implicitly_wait(60)
                                        # 获取当前窗口的句柄
                                        origin_windows=browser.window_handles[0]
                                        current_windows = browser.window_handles[-1]
                                        browser.switch_to.window(current_windows)

                                        rowInfo = WebDriverWait(browser,30).until(lambda browser:browser.find_elements_by_xpath(
                                            "//table[@id='Table1']//table[@id='mainModuleContainer_1855_1856_ctl00_ctl00_p1_f1']/tbody/tr"))
                                    except Exception as e:
                                        print('读取数据时出错！'+repr(e))
                                        browser.close()
                                        browser.switch_to.window(origin_windows)
                                        print('等待' + str(index) + 's')
                                        time.sleep(index)
                                        index = index + 10
                                    else:
                                        tag = 0
                                        
                                        # 记录地区编号
                                        adcode = browser.find_element_by_id("mainModuleContainer_1855_1856_ctl00_ctl00_p1_f1_r1_c2_ctrl_value").get_attribute("value")
                                        table.write(recordNum, tb_head.index("adcode"), adcode)
                                        #读取表格数据
                                        for row in rowInfo[2:]:
                                            i = 0
                                            spanList = row.find_elements_by_xpath('./td/span')
                                            # print(spanList[i].text)
                                            while (i < len(spanList)):
                                                if spanList[i].text == '分期支付约定:':
                                                    subspanList = row.find_elements_by_xpath(
                                                        "./td/table//tr[@id='mainModuleContainer_1855_1856_ctl00_ctl00_p1_f3_r2_0']//span")
                                                    for j in range(len(subspanList) - 1):
                                                        # print(subspanList[j].text)
                                                        table.write(recordNum, 18 + j, subspanList[j].text)
                                                    i = i + 1
                                                elif spanList[i].text == '约定容积率:':
                                                    subspanList = row.find_elements_by_xpath('./td/table//span')
                                                    for subspan in subspanList:
                                                        if subspan.text in tb_head:
                                                            Index = subspanList.index(subspan)
                                                            table.write(recordNum, tb_head.index(subspan.text), subspanList[Index + 1].text)
                                                    i = i + 1
                                                elif spanList[i].text == '土地使用权人:':
                                                    if spanList[i + 1].text == '':
                                                        Index = rowInfo.index(row)
                                                        table.write(recordNum, tb_head.index('土地使用权人'), rowInfo[Index + 1].text)
                                                    else:
                                                        table.write(recordNum, tb_head.index('土地使用权人'), spanList[i + 1].text)
                                                    i = i + 1
                                                elif spanList[i].text in tb_head:
                                                    table.write(recordNum, tb_head.index(spanList[i].text), spanList[i + 1].text)
                                                    i = i + 2
                                                else:
                                                    i = i + 1

                                        browser.close()
                                        recordNum=recordNum+1
                                        browser.switch_to.window(origin_windows)
                                        time.sleep(2)
                            wb.get_sheet('historyNum').write(0,0,recordNum)
                            wb.save(EXCEL_NAME)
                    print(page)
                    wb.get_sheet('historyNum').write(0, 2,page+1)
                    wb.save(EXCEL_NAME)
                wb.get_sheet('historyNum').write(0, 2, 1)
                wb.save(EXCEL_NAME)
        historyDate = historyDate + oneday
        start = historyDate.strftime('%Y-%m-%d')
        print('存储日期'+start)
        wb.get_sheet('historyNum').write(0,1,start)
        wb.save(EXCEL_NAME)
