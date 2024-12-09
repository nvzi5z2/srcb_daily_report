import pandas as pd
import numpy as np
import os

result_path=r'C:\Users\Wesle\Desktop\日报分析\result'

data_path=r'C:\Users\Wesle\Desktop\日报分析\原始数据'

department_list='员工部门归属表.xlsx'

yesterday_daily_report='网点鑫e贷月度指标完成情况1206.xlsx'

client_manager_data='【浦东分行鑫e贷】客户经理营销数据_2024-12-08.xlsx'

retail_performance_data='零售市场部协同外拓及理财转介业绩报送.xlsx'

T0_Date='2024-12-06'

def XY_Dai_Zong_Shou_Xin(data_path,yesterday_daily_report,
department_list,client_manager_data,T0_Date,result_path):

    yesterday_daily_report_df=pd.read_excel(data_path+'\\'+yesterday_daily_report)

    yesterday_daily_report_df=yesterday_daily_report_df.iloc[1:,:]

    new_columns = yesterday_daily_report_df.iloc[0,2:]  # 选择第一行作为新的列名

    index = yesterday_daily_report_df.iloc[1:, 1:2].dropna()  # 假设这是你提取出来的列数据
    index_list = index.squeeze().tolist()  # 将提取出的列转换为列表

    daily_report=yesterday_daily_report_df.iloc[1:45,2:]

    data.index=index_list

    data.columns=new_columns

    #第一个结果（指标）提取昨日报表指标列
    kpi=data[['指标']]

    kpi.columns=['鑫e贷总授信_指标','鑫e贷放款_指标','鑫e贷授信-B款_指标']

    kpi_result=kpi[['鑫e贷总授信_指标']]
    
    #第二个结果（昨日完成数），提取完成数的列表
    yesterday_finished=data[['完成数']]

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
    merged_df = total_shouxin.reset_index().merge(department_df, on='kehujinglixingm', how='left')
    
    # 按部门分组，计算业绩总和
    department_totals =  merged_df.groupby('department', as_index=False)['benyueshouxinrenshu'].sum()

    department_totals=department_totals.set_index('department',drop=True)
    #匹配到浦东分行的表里

    result.loc[:,"鑫e贷总授信_报表数"]=department_totals


    #第四个结果(协同外拓)

    #提取昨日协同外拓数

    yesterday_retail_performance=data[['协同外拓']]

    yesterday_retail_performance.columns=['鑫e贷总授信_昨日协同外拓','鑫e贷放款_昨日协同外拓','鑫e贷授信-B款_昨日协同外拓']

    yesterday_retail_performance=yesterday_retail_performance[['鑫e贷总授信_昨日协同外拓']]
    
    retail_performance_df=pd.read_excel(data_path+'\\'+retail_performance_data)

    retail_performance_df=retail_performance_df.set_index('外拓日期',drop=True)
    
    retail_performance_df.index=pd.to_datetime(retail_performance_df.index, unit='d', origin='1899-12-30')

    today_retail_df=retail_performance_df.loc[T0_Date,:]

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

    result.loc[:,"数据调整数"]=0

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

    result.to_excel(result_path+'\\'+'鑫e贷总授信完成情况.xlsx')

    print('恭喜米，鑫e贷总授信计算完成')

    return result


XY_Dai_Zong_Shou_Xin_result=XY_Dai_Zong_Shou_Xin(data_path,yesterday_daily_report,
department_list,client_manager_data,T0_Date,result_path)