import pandas as pd

# # 读取CSV文件
# csv_file = '../doc/我的照片与录音.csv'
# df = pd.read_csv(csv_file, encoding='gbk')
#
# # 将DataFrame保存为Excel文件
# excel_file = 'output.xlsx'
# df.to_excel(excel_file, index=False)
#
# print(f'成功将CSV文件转换为Excel表格，并保存为 {excel_file}')

import os

file_path = "example.txt"
file_name, file_extension = os.path.splitext(file_path)

print("文件名：", file_name)
print("扩展名：", file_extension)