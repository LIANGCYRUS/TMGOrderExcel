import os
import pandas as pd
import openpyxl
from decimal import Decimal, ROUND_HALF_UP


order_df = pd.read_excel('raw/Order.xlsx')
list_df = pd.read_csv("raw/List.csv", encoding="gbk")
alipay_df = pd.read_csv('raw/2088541320333819_20230216202918.csv', encoding='gbk')

# 读取列表后，只要'订单编号', '确认收货时间'，其他都去掉
selected_cols = ['订单编号', '确认收货时间']
order_df = order_df.loc[:, selected_cols]

# print(order_df)

# 在支付表的列表里，先匹配上确认收货时间
new_alipay_df = pd.merge(alipay_df,order_df, left_on="Partner_transaction_id",right_on="订单编号", how="left")

print(new_alipay_df)

# 删除指定的列
list_df = list_df.drop(["子订单编号", "价格", "外部系统编号", "套餐信息", "备注", "商家编码", "买家应付货款", "退款状态", "退款金额"], axis=1)

# 把支付单号列移动到第二列
cols = list(list_df.columns)
payment_no_index = cols.index('支付单号')
cols.pop(payment_no_index)
cols.insert(1, '支付单号')
list_df = list_df.loc[:, cols]

list_df['主订单编号'] = list_df['主订单编号'].apply(lambda x: f"'{str(x)}" if isinstance(x, (int, float)) else x.strip('="'))
list_df['支付单号'] = list_df['支付单号'].apply(lambda x: f"'{str(x)}" if isinstance(x, (int, float)) else x.strip('="'))

print(list_df)

new_alipay_df['Partner_transaction_id'] = new_alipay_df['Partner_transaction_id'].astype(str)

merged_df = list_df.merge(new_alipay_df, how='left', left_on='主订单编号', right_on='Partner_transaction_id')

merged_df['主订单编号'] = merged_df['主订单编号'].fillna(merged_df['Partner_transaction_id'])
merged_df = merged_df.drop(columns=['Partner_transaction_id','Transaction_id','订单编号'])

confirm_time = merged_df.pop('确认收货时间')  # 弹出“确认收货时间”列
merged_df.insert(9, '确认收货时间', confirm_time)  # 将“确认收货时间”列插入到“订单付款时间”列的后面


merged_df['Amount_to_split'] = merged_df['买家实际支付金额']/merged_df['Rate']
# merged_df['Amount_to_split'] = pd.to_numeric(merged_df['Amount_to_split']) #将上一串代码计算出来的文本格式改为数字格式

Amount_to_split = merged_df.pop('Amount_to_split')  # 弹出“确认收货时间”列
merged_df.insert(10, 'Amount_to_split', Amount_to_split)  # 将“确认收货时间”列插入到“订单付款时间”列的后面


# 保存修改后的文件
merged_df.to_excel("List_modified2.xlsx", index=False)