import os

# 获取当前脚本所在的目录
current_directory = os.path.dirname(os.path.abspath(__file__))

# 文件的路径
file_path = os.path.join(current_directory, 'E:/项目文件夹/江宁普查项目外业资料/测试资料/道路总图/1.市政设施设施量统计表模板-汤山街道（已标定）.xlsx')

# 获取文件的父路径
parent_directory = os.path.dirname(file_path)

print("文件的父路径是:", parent_directory)