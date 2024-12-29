import pandas as pd

# 读取 Excel 文件
df = pd.read_excel('2024年发货记录 闫敏 11.27.xlsx')

# 查看文件的前几行
print(df.head())

# 假设 'shipping_date', 'currency', 'amount' 是文件中的列名
df['发货日期'] = pd.to_datetime(df['发货日期'])

# 筛选出发货日期在每月 27 号或之前的数据
df_filtered = df[df['发货日期'].dt.day <= 27]
# 处理12月份，将12月30号之前的日期也包含在12月中
df_december = df[(df['发货日期'].dt.month == 12) & (df['发货日期'].dt.day <= 30) & (df['发货日期'].dt.day > 27)]
df_filtered = pd.concat([df_filtered, df_december])
df_filtered['month'] = df_filtered['发货日期'].dt.to_period('M')
# 筛选出1月的数据
# df_january = df_filtered[df_filtered['month'] == '2024-01']
monthly_exchange_rates = [7.077, 7.1049, 7.1059, 7.0938, 7.0994, 7.1086, 7.1265, 7.1323, 7.1027, 7.0709, 7.1135, 7.1865]


def total_sum(df, type):
    total_by_month = {}
    # 遍历1到12个月，计算每个月的总价
    for month in range(1, 13):
        if month == 12:
            aa = 1
        month_str = f'2024-{month:02d}'

        # 筛选出该月的数据
        df_month = df[df['month'] == month_str]

        # 计算该月USD和RMB的总价
        df_month_usd = df_month[df_month['计价单位'] == 'USD']
        total_amount_usd = df_month_usd['总价'].sum() * monthly_exchange_rates[month - 1]  # 转换成RMB

        df_month_rmb = df_month[df_month['计价单位'] == 'RMB']
        total_amount_rmb = round(df_month_rmb['总价'].sum() / 1.13, 2)  # 转换成USD

        # 存入汇总字典
        total_by_month[month_str] = total_amount_usd + total_amount_rmb
    # 输出每月汇总
    print('品类:', type)
    print("每月总金额 (以RMB计)：")
    for month in total_by_month:
        print(f"{month}: {total_by_month[month]:.2f} RMB")
    print(f"总金额:{sum(total_by_month.values()):.2f} RMB")
    return sum(total_by_month.values())


# 计算1月的总金额，假设金额列名为 'amount'
# 计算总价
# huilv_1 = 7.077
# df_january_usd = df_january[df_january['计价单位'] == 'USD']
# total_amount_january_usd = df_january_usd['总价'].sum() * huilv_1
# df_january_rmb = df_january[df_january['计价单位'] == 'RMB']
# total_amount_january_rmb = round(df_january_rmb['总价'].sum() / 1.13, 2)
# print(total_amount_january_usd)
# print(total_amount_january_rmb)
# 初始化汇总变量
df_filtered_normal = df_filtered[~df_filtered['产品名称'].isin(['瑞舒伐', '左乙拉西坦'])]
total_1 = total_sum(df_filtered_normal, "常规产品")
df_filtered_rsf = df_filtered[df_filtered["产品名称"] == "瑞舒伐"]
total_2 = total_sum(df_filtered_rsf, "瑞舒伐")
df_filtered_zylxt = df_filtered[df_filtered["产品名称"] == "左乙拉西坦"]
total_3 = total_sum(df_filtered_zylxt, "左乙拉西坦")
print('全年:', total_1 + total_2 + total_3)
