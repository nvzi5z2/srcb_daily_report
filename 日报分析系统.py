import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt

#数据路径设置

result_path=r'D:\srcb_daily_report\result'

data_path=r'D:\srcb_daily_report\原始数据'

#部门归属表文件名
department_list='员工部门归属表.xlsx'

#昨日日报表文件名
yesterday_daily_report='网点鑫e贷月度指标完成情况1223.xlsx'

#【浦东分行鑫e贷】客户经理营销数据文件名
client_manager_data='【浦东分行鑫e贷】客户经理营销数据_2024-12-23.xlsx'

retail_performance_data='零售市场部协同外拓及理财转介业绩报送-7.xlsx'

type_B_data='【浦东分行鑫e贷】鑫e贷b款明细_2024-12-23.xlsx'

T0_Date='2024-12-23'



def plot_setting():
    # 设置全局字体为支持中文的字体（SimHei 黑体）
    plt.rcParams['font.sans-serif'] = ['SimHei']  # 用黑体显示中文
    plt.rcParams['axes.unicode_minus'] = False   # 正常显示负号




def XY_Dai_Zong_Shou_Xin(data_path,yesterday_daily_report,
department_list,client_manager_data,T0_Date,result_path):

    yesterday_daily_report_df=pd.read_excel(data_path+'\\'+yesterday_daily_report)

    yesterday_daily_report_df=yesterday_daily_report_df.iloc[1:,:]

    new_columns = yesterday_daily_report_df.iloc[0,2:]  # 选择第一行作为新的列名

    index = yesterday_daily_report_df.iloc[1:, 1:2].dropna()  # 假设这是你提取出来的列数据
    index_list = index.squeeze().tolist()  # 将提取出的列转换为列表

    daily_report=yesterday_daily_report_df.iloc[1:45,2:]

    daily_report.index=index_list

    daily_report.columns=new_columns

    #第一个结果（指标）提取昨日报表指标列
    kpi=daily_report[['指标']]

    kpi.columns=['鑫e贷总授信_指标','鑫e贷放款_指标','鑫e贷授信-B款_指标']

    kpi_result=kpi[['鑫e贷总授信_指标']]
    
    #第二个结果（昨日完成数），提取完成数的列表
    yesterday_finished=daily_report[['完成数']]

    yesterday_finished.columns=['鑫e贷总授信_昨日完成数','鑫e贷放款_昨日完成数','鑫e贷授信-B款_昨日完成数']

    yesterday_finished_result=yesterday_finished[['鑫e贷总授信_昨日完成数']]
    
    result=pd.merge(kpi_result,yesterday_finished_result,right_index=True,left_index=True)

    #第三个结果(报表数）

    client_manager_df=pd.read_excel(data_path+'\\'+client_manager_data)

    client_manager_df=client_manager_df.set_index('kehujinglixingm',drop=True)

    monthly_shouxin=client_manager_df[['benyueshouxinrenshu']]

    manager_total_shouxin = monthly_shouxin.groupby(monthly_shouxin.index).sum()

    department_df=pd.read_excel(data_path+'\\'+department_list)

    # 重命名列，确保列名一致，便于匹配
    department_df.rename(columns={'员工姓名': 'kehujinglixingm', '部门': 'department'}, inplace=True)

    # 合并客户经理业绩数据和部门数据
    merged_df = manager_total_shouxin.reset_index().merge(department_df, on='kehujinglixingm', how='left')
    
    # 按部门分组，计算业绩总和
    department_totals =  merged_df.groupby('department', as_index=False)['benyueshouxinrenshu'].sum()

    department_totals=department_totals.set_index('department',drop=True)
    #匹配到浦东分行的表里

    result.loc[:,"鑫e贷总授信_报表数"]=department_totals


    #第四个结果(协同外拓)

    #提取昨日协同外拓数

    yesterday_retail_performance=daily_report[['协同外拓']]

    yesterday_retail_performance.columns=['鑫e贷总授信_昨日协同外拓','鑫e贷放款_昨日协同外拓','鑫e贷授信-B款_昨日协同外拓']

    yesterday_retail_performance=yesterday_retail_performance[['鑫e贷总授信_昨日协同外拓']]
    
    retail_performance_df=pd.read_excel(data_path+'\\'+retail_performance_data)

    retail_performance_df=retail_performance_df.set_index('外拓日期',drop=True)
    
    retail_performance_df.index=pd.to_datetime(retail_performance_df.index, unit='d', origin='1899-12-30')

    today_retail_df=retail_performance_df.loc[T0_Date,:]

    today_retail_df = today_retail_df.to_frame().T

    today_retail_df=today_retail_df.fillna(0)

    today_retail_df=today_retail_df[['客户经理姓名','协同外拓网点','其中本人\nA款授信（户）','其中本人\nB款授信（户）']]

    today_retail_df.loc[:,"鑫e贷总授信_今日协同外拓"]=today_retail_df.loc[:,"其中本人\nA款授信（户）"]+today_retail_df.loc[:,"其中本人\nB款授信（户）"]*2
    
    today_retail_result=today_retail_df[['协同外拓网点','鑫e贷总授信_今日协同外拓']]

    today_retail_result=today_retail_result.set_index('协同外拓网点',drop=True)

    newest_result=pd.concat([yesterday_retail_performance,today_retail_result],axis=1)

    newest_result.loc[:, "鑫e贷总授信_协同外拓"] = (
        newest_result.loc[:, "鑫e贷总授信_昨日协同外拓"].fillna(0) +
        newest_result.loc[:, "鑫e贷总授信_今日协同外拓"].fillna(0)
    )

    result.loc[:,"鑫e贷总授信_协同外拓"]=newest_result.loc[:, "鑫e贷总授信_协同外拓"]

    result.loc[:,"B款月底额外计1户授信"]=0

    #取昨日数据调整数

    yesterday_daily_adjusted=daily_report[['数据调整数']]

    result.loc[:,"数据调整数"]=yesterday_daily_adjusted

    #完成数

    result.loc[:,"鑫e贷总授信_完成数"]=result[['鑫e贷总授信_报表数','鑫e贷总授信_协同外拓',
                                'B款月底额外计1户授信','数据调整数']].sum(axis=1)
    # 计算完成率（完成数 / 指标），保留小数点后两位
    result.loc[:, "鑫e贷总授信_完成率"] = (
        result.loc[:, "鑫e贷总授信_完成数"] / result.loc[:, "鑫e贷总授信_指标"]
    ).fillna(0).round(2)
    
    result.loc[:,"鑫e贷总授信_昨日完成数(轧差)"]=result.loc[:,"鑫e贷总授信_完成数"]-result.loc[:,"鑫e贷总授信_昨日完成数"]

    final_result=result[['鑫e贷总授信_指标','鑫e贷总授信_昨日完成数(轧差)',
                            '鑫e贷总授信_报表数','鑫e贷总授信_协同外拓','B款月底额外计1户授信','数据调整数',
                                '鑫e贷总授信_完成数','鑫e贷总授信_完成率']]

    final_result.to_excel(result_path+'\\'+'鑫e贷总授信完成情况.xlsx')

    print('恭喜米，鑫e贷总授信计算完成')

    return final_result


def XY_Dai_Fang_Kuang(data_path,yesterday_daily_report,
department_list,client_manager_data,T0_Date,type_B_data,result_path):
    
    yesterday_daily_report_df=pd.read_excel(data_path+'\\'+yesterday_daily_report)

    yesterday_daily_report_df=yesterday_daily_report_df.iloc[1:,:]

    new_columns = yesterday_daily_report_df.iloc[0,2:]  # 选择第一行作为新的列名

    index = yesterday_daily_report_df.iloc[1:, 1:2].dropna()  # 假设这是你提取出来的列数据
    index_list = index.squeeze().tolist()  # 将提取出的列转换为列表

    daily_report=yesterday_daily_report_df.iloc[1:45,2:]

    daily_report.index=index_list

    daily_report.columns=new_columns

    #第一个结果（指标）提取昨日报表指标列
    kpi=daily_report[['指标']]

    kpi.columns=['鑫e贷总授信_指标','鑫e贷放款_指标','鑫e贷授信-B款_指标']

    kpi_result=kpi[['鑫e贷放款_指标']]

    #第二个结果（昨日完成数），提取完成数的列表
    yesterday_finished=daily_report[['完成数']]

    yesterday_finished.columns=['鑫e贷总授信_昨日完成数','鑫e贷放款_昨日完成数','鑫e贷授信-B款_昨日完成数']

    yesterday_finished_result=yesterday_finished[['鑫e贷放款_昨日完成数']]
    
    result=pd.merge(kpi_result,yesterday_finished_result,right_index=True,left_index=True)

    #第三个结果（报表数）

    client_manager_df=pd.read_excel(data_path+'\\'+client_manager_data)

    client_manager_df=client_manager_df.set_index('kehujinglixingm',drop=True)

    monthly_fangkuan=client_manager_df[['benyuefangkuanjine']]
    
    manager_total_fangkuan = monthly_fangkuan.groupby(monthly_fangkuan.index).sum()

    department_df=pd.read_excel(data_path+'\\'+department_list)

    # 重命名列，确保列名一致，便于匹配
    department_df.rename(columns={'员工姓名': 'kehujinglixingm', '部门': 'department'}, inplace=True)

    # 合并客户经理业绩数据和部门数据
    merged_df = manager_total_fangkuan.reset_index().merge(department_df, on='kehujinglixingm', how='left')
    
    # 按部门分组，计算业绩总和
    department_totals =  merged_df.groupby('department', as_index=False)['benyuefangkuanjine'].sum()

    department_totals=department_totals.set_index('department',drop=True)

    department_totals=department_totals/10000

    department_totals=department_totals.fillna(0).round(0)

    department_totals.columns=['鑫e贷放款_报表数']

    result.loc[:,"鑫e贷放款_报表数"]=department_totals
    
    #第四个结果（自然流量）

    #读取b款数据

    type_B_df=pd.read_excel(data_path+'\\'+type_B_data,sheet_name="鑫e贷大额客户借据数据")

    type_B_df=type_B_df.set_index('fangkuanriq')

    type_B_df.index=pd.to_datetime(type_B_df.index)

    year = pd.to_datetime(T0_Date).year  # 提取年份

    month = pd.to_datetime(T0_Date).month  # 提取月份

    filtered_df = type_B_df[(type_B_df.index.year == year) & (type_B_df.index.month == month)]

    netural_df=filtered_df[['jingdiaokehujingl','yingxiaokehujingl','fangkuanjine']]

    # 筛选出 yingxiaokehujingl 和 jingdiaokehujingl 不相等的行

    netural_df_filtered = netural_df[netural_df['yingxiaokehujingl'] != netural_df['jingdiaokehujingl']]

    netural_result= netural_df_filtered.groupby('jingdiaokehujingl', as_index=False)['fangkuanjine'].sum()
    
    netural_result=netural_result.set_index('jingdiaokehujingl',drop=True)

    department_df_1=pd.read_excel(data_path+'\\'+department_list)

    # 重命名列，确保列名一致，便于匹配
    department_df_1.rename(columns={'员工姓名': 'jingdiaokehujingl', '部门': 'department'}, inplace=True)

    # 合并客户经理业绩数据和部门数据
    netural_merged_df = netural_result.reset_index().merge(department_df_1, on='jingdiaokehujingl', how='left')
    
    netural_department_totals =  netural_merged_df.groupby('department', as_index=False)['fangkuanjine'].sum()

    netural_department_totals=netural_department_totals.set_index('department',drop=True)

    netural_department_totals=netural_department_totals/10000

    netural_department_totals=netural_department_totals.fillna(0).round(2)

    netural_department_totals.columns=['鑫e贷放款_自然流量']

    result.loc[:,"鑫e贷放款_自然流量"]=netural_department_totals

    #第四个指标（协同外拓）

    yesterday_retail_performance=daily_report[['协同外拓']]

    yesterday_retail_performance.columns=['鑫e贷总授信_昨日协同外拓','鑫e贷放款_昨日协同外拓','鑫e贷授信-B款_昨日协同外拓']

    result.loc[:,"鑫e贷放款_协同外拓"]=yesterday_retail_performance[['鑫e贷放款_昨日协同外拓']]

    result.loc[:,"鑫e贷放款_完成数"]=result[['鑫e贷放款_报表数','鑫e贷放款_自然流量',
                                '鑫e贷放款_协同外拓']].sum(axis=1)

    result.loc[:, "鑫e贷放款_完成率"] = (result.loc[:, "鑫e贷放款_完成数"] /
                             result.loc[:, "鑫e贷放款_指标"]).fillna(0).round(2)
    
    result.loc[:,"鑫e贷放款_昨日完成数(轧差)"]=result.loc[:,"鑫e贷放款_完成数"]-result.loc[:,"鑫e贷放款_昨日完成数"]

    final_result=result[['鑫e贷放款_指标','鑫e贷放款_昨日完成数(轧差)',
                        '鑫e贷放款_报表数','鑫e贷放款_自然流量','鑫e贷放款_协同外拓','鑫e贷放款_完成数',
                            '鑫e贷放款_完成率']]
                            
    # final_result.sort_values('鑫e贷放款_完成率', ascending=True)[['鑫e贷放款_完成率']].plot(
    #     kind='barh', 
    #     figsize=(10, 15)
    # )
    # plt.title('团队KPI完成率对比', fontsize=14)
    # plt.tight_layout()
    # plt.show()

    final_result.to_excel(result_path+'\\'+'鑫e贷放款完成情况.xlsx')
    
    print('恭喜米，鑫e贷放款计算完成')

    return final_result


def B_Kuang_Shou_Xin(data_path,yesterday_daily_report,
department_list,client_manager_data,T0_Date,type_B_data,result_path):

    yesterday_daily_report_df=pd.read_excel(data_path+'\\'+yesterday_daily_report)

    yesterday_daily_report_df=yesterday_daily_report_df.iloc[1:,:]

    new_columns = yesterday_daily_report_df.iloc[0,2:]  # 选择第一行作为新的列名

    index = yesterday_daily_report_df.iloc[1:, 1:2].dropna()  # 假设这是你提取出来的列数据
    index_list = index.squeeze().tolist()  # 将提取出的列转换为列表

    daily_report=yesterday_daily_report_df.iloc[1:45,2:]

    daily_report.index=index_list

    daily_report.columns=new_columns

    #第一个结果（指标）提取昨日报表指标列
    kpi=daily_report[['指标']]

    kpi.columns=['鑫e贷总授信_指标','鑫e贷放款_指标','鑫e贷授信-B款_指标']

    kpi_result=kpi[['鑫e贷授信-B款_指标']]

    result=pd.DataFrame(index=kpi_result.index)

    result.loc[:,"鑫e贷授信-B款_指标"]=kpi_result

    #第二个结果（报表数）
    
    type_B_df=pd.read_excel(data_path+'\\'+type_B_data,sheet_name="鑫e贷大额客户授信数据")

    type_B_df=type_B_df.set_index('qianyueshijian')

    type_B_df.index=pd.to_datetime(type_B_df.index)

    year = pd.to_datetime(T0_Date).year  # 提取年份

    month = pd.to_datetime(T0_Date).month  # 提取月份

    filtered_df = type_B_df[(type_B_df.index.year == year) & (type_B_df.index.month == month)]

    DD_time=filtered_df[['jingdiaokehujingli']]

    DD_result = DD_time['jingdiaokehujingli'].value_counts().reset_index()

    DD_result.columns=['index','count']
    
    DD_result=DD_result.set_index('index',drop=True)

    department_df=pd.read_excel(data_path+'\\'+department_list)

    # 重命名列，确保列名一致，便于匹配
    department_df.rename(columns={'员工姓名': 'jingdiaokehujingli', '部门': 'department'}, inplace=True)

    department_df=department_df.set_index('jingdiaokehujingli',drop=True)

    # 合并客户经理业绩数据和部门数据
    merged_df =pd.merge(DD_result,department_df,right_index=True,left_index=True)
    
    merged_df.columns=['count','department']
    # 按部门分组，计算业绩总和
    department_totals =  merged_df.groupby('department', as_index=False)['count'].sum()

    department_totals=department_totals.set_index('department',drop=True)

    department_totals.columns=['鑫e贷授信-B款_报表数']

    result.loc[:,"鑫e贷授信-B款_报表数"]=department_totals

    #第三个结果（协同外拓）

    yesterday_retail_performance=daily_report[['协同外拓']]

    yesterday_retail_performance.columns=['鑫e贷总授信_昨日协同外拓','鑫e贷放款_昨日协同外拓','鑫e贷授信-B款_昨日协同外拓']

    yesterday_retail_performance=yesterday_retail_performance[['鑫e贷授信-B款_昨日协同外拓']]
    
    retail_performance_df=pd.read_excel(data_path+'\\'+retail_performance_data)

    retail_performance_df=retail_performance_df.set_index('外拓日期',drop=True)
    
    retail_performance_df.index=pd.to_datetime(retail_performance_df.index, unit='d', origin='1899-12-30')

    today_retail_df=retail_performance_df.loc[T0_Date,:]
    
    today_retail_df = today_retail_df.to_frame().T

    today_retail_df=today_retail_df[['客户经理姓名','协同外拓网点','其中本人\nA款授信（户）','其中本人\nB款授信（户）']]
    
    today_retail_df=today_retail_df.fillna(0)
    
    today_retail_df.loc[:,"鑫e贷授信-B款_今日协同外拓"]=today_retail_df.loc[:,"其中本人\nB款授信（户）"].fillna(0)
    
    today_retail_result=today_retail_df[['协同外拓网点','鑫e贷授信-B款_今日协同外拓']]

    today_retail_result=today_retail_result.set_index('协同外拓网点',drop=True)

    newest_result=pd.concat([yesterday_retail_performance,today_retail_result],axis=1)

    newest_result.loc[:, "鑫e贷授信-B款_协同外拓"] = (
        newest_result.loc[:, "鑫e贷授信-B款_昨日协同外拓"].fillna(0) +
        newest_result.loc[:, "鑫e贷授信-B款_今日协同外拓"].fillna(0)
        )

    result.loc[:,"鑫e贷授信-B款_协同外拓"]=newest_result.loc[:, "鑫e贷授信-B款_协同外拓"]

    #第四个结果（调整数）

    result.loc[:,"鑫e贷授信-B款_调整数"]=0

    #第五个结果（完成数）

    result.loc[:,"鑫e贷授信-B款_完成数"]=result[['鑫e贷授信-B款_报表数','鑫e贷授信-B款_协同外拓',
                                '鑫e贷授信-B款_调整数']].sum(axis=1)

    result.loc[:, "鑫e贷授信-B款_完成率"] = (result.loc[:, "鑫e贷授信-B款_完成数"] /
                             result.loc[:, "鑫e贷授信-B款_指标"]).fillna(0).round(2)
                                   
    result.to_excel(result_path+'\\'+'鑫e贷授信-B款完成情况.xlsx')

    print('恭喜米，鑫e贷授信-B款计算完成')

    return result


plot_setting()

XY_Dai_Zong_Shou_Xin_result=XY_Dai_Zong_Shou_Xin(data_path,yesterday_daily_report,
department_list,client_manager_data,T0_Date,result_path)


XY_Dai_Fang_Kuang_result=XY_Dai_Fang_Kuang(data_path,yesterday_daily_report,
department_list,client_manager_data,T0_Date,type_B_data,result_path)

B_Kuang_Shou_Xin_result=B_Kuang_Shou_Xin(data_path,yesterday_daily_report,
department_list,client_manager_data,T0_Date,type_B_data,result_path)


total=pd.concat([XY_Dai_Zong_Shou_Xin_result,XY_Dai_Fang_Kuang_result,B_Kuang_Shou_Xin_result],axis=1)

total.to_excel(result_path+'\\'+'日报总表.xlsx')