# -*- coding: utf-8 -*-
import pandas as pd
import time

"""
填写几个数据即可，其它不用更改，格式如下
    - 要读取的表的位置, excel_for_read
    - 保存文件的位置, excel_for_save
    - 筛选的月份, year_and_month
"""
excel_for_read = "E:/Task/2020/3月/采购汇总20年上-03041731下载字段修改后.xlsx"
excel_for_save = "E:/Task/2020/3月/采购汇总20年上华光汇总0305.xlsx"
year_and_month = "2020-02"


"""
============================== 用户请止步 ==============================================
"""
def getDataFromSheet(df, sheet_index):
    df['时间2'] = pd.to_datetime(df['时间2'])
    df = df.set_index('时间2')
    try:
        ret = df[year_and_month]
    except:
        ret = None

    return ret


if __name__ == '__main__':
    #import modin.pandas as pd
    """ 1. 读取Excel表"""
    df_total = pd.DataFrame(columns=['支付宝名称', '时间1', '时间2', '备注', '收入', '支出', '客户'])

    """ 2. 获取指定excel workbook 下所有sheet 名称 """
    workbook = pd.ExcelFile(excel_for_read)
    #print(workbook.sheet_names)
    #print(len(excel.sheet_names))

    for sheet in range(0, len(workbook.sheet_names)):
        print("\n 当前正在处理 >>>\n 表名 %s, 序号 %d." % (workbook.sheet_names[sheet], sheet))

        df_extract = getDataFromSheet(pd.read_excel(excel_for_read, sheet_name=sheet), sheet)
        if df_extract is None or df_extract.empty:
            #print("请注意: [%s] 没有指定月份的数据." % workbook.sheet_names[sheet])
            print("\033[0;33;40m\t请注意: [%s] 没有指定月份的数据.\033[0m" % workbook.sheet_names[sheet])
            continue
        else:
            #print(df_extract[['支付宝名称', '时间1', '备注', '收入', '支出', '客户']])
            df_extract = df_extract.reset_index()
            df_total = df_total.append(df_extract[['支付宝名称', '时间1', '时间2', '备注', '收入', '支出', '客户']],
                                       ignore_index=True, sort=False)

    """3. 保存筛选出的数据到汇总表"""
    df_total.to_excel(excel_for_save, index=False)

