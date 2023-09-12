import pandas as pd

import os

import sys
#当前路径
current_path = os.getcwd()

# 指定要遍历的文件夹路径

folder_path = r"E:\wechatfile\WeChat Files\wxid_h98ga60fdavf22\FileStorage\File\2023-09\可以打印的科室(1)\可以打印的科室"
# 添加你想要的目录
os.chdir(folder_path)
columns = pd.MultiIndex.from_tuples([
    ('基础信息', '编码'),
    ('基础信息', '名称'),
    ('2020年12月31日库存数', '单价'),
    ('2020年12月31日库存数', '数量'),
    ('2020年12月31日库存数', '金额'),
    ('2023年6月30日盘点数', '单价'),
    ('2023年6月30日盘点数', '数量'),
    ('2023年6月30日盘点数', '金额')
], names=['主表头', '子表头'])
summary_df = pd.DataFrame(columns=columns)


excel_list = os.listdir(folder_path)

for excel_name in excel_list:
    # 使用列的字母标签
    df = pd.read_excel(excel_name, usecols="A:E, O:Q",skiprows=[0, 1], names=columns, dtype={('基础信息', '编码'): str})

    df.fillna(0,inplace= True)


    for index, row in df.iterrows():
            temp = row[("基础信息","编码")]
            code = str(row[("基础信息","编码")])

            
            # 检查编码是否已经存在于汇总DataFrame中
            if code in summary_df[("基础信息","编码")].values:
                # 如果编码已存在，更新2020年12月31日库存数的数量和金额
                summary_df.loc[summary_df[('基础信息', '编码')].astype(str) == code, ('2020年12月31日库存数', '数量')] += row[('2020年12月31日库存数', '数量')]
                summary_df.loc[summary_df[('基础信息', '编码')].astype(str) == code, ('2020年12月31日库存数', '金额')] += row[('2020年12月31日库存数', '金额')]

                # 如果编码已存在，更新2023年6月30日盘点数的数量和金额
                summary_df.loc[summary_df[('基础信息', '编码')].astype(str) == code, ('2023年6月30日盘点数', '数量')] += row[('2023年6月30日盘点数', '数量')]
                summary_df.loc[summary_df[('基础信息', '编码')].astype(str) == code, ('2023年6月30日盘点数', '金额')] += row[('2023年6月30日盘点数', '金额')]

            else:
                # 如果编码不存在，添加新行
                row[("基础信息", "编码")] = str(row[("基础信息", "编码")])
                summary_df = summary_df._append(row,ignore_index=True)
    print("{}. has been done.".format(excel_name))
#保存summary_df
save_path = current_path + "\\summary.xlsx"
summary_df.to_excel(save_path)


    