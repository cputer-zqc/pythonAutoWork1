import re

# 定义一个包含数字的字符串
string_with_numbers = "Hello, I have 123 apples and 456 oranges."

# 使用正则表达式提取字符串中的所有数字
numbers = re.findall(r'\d+', string_with_numbers)

# 将提取到的数字转换为整数列表
numbers = list(map(int, numbers))

# 输出提取到的数字列表
print("提取到的数字列表:", type(numbers[0]))