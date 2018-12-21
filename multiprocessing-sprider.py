import LCspider
import multiprocessing
from createExcel import startYear, endYear

if __name__ == "__main__":
  pool = multiprocessing.Pool(processes=5) # 创建多个进程
  for i in range(startYear,endYear+1):
    # 数据存储文件的位置
    EXCEL_NAME = 'res_data_landchina' + str(i) + '.xls'
    pool.apply_async(LCspider.LandChina, (EXCEL_NAME, ))
  pool.close() # 关闭进程池，表示不能在往进程池中添加进程
  pool.join() # 等待进程池中的所有进程执行完毕，必须在close()之后调用
  print("Sub-process(es) done.")