import pandas as pd
import random

# 常用公司类型
company_types = ["科技", "信息", "网络", "集团", "实业", "有限", "股份", "贸易", "金融", "投资", "发展", "控股", "科技", "医药", "生物", "制药", "能源", "文化", "传媒", "旅游", "教育"]

# 常用词汇
company_words = ["创新", "发展", "科技", "智能", "传媒", "网络", "信息", "产业", "集团", "国际", "金融", "投资", "实业", "科学", "生物", "医药", "健康", "环保", "文化", "艺术", "教育", "旅游", "时尚", "建筑", "地产", "能源", "绿色", "创业"]

# 生成100个公司
companies = []
for _ in range(100):
    company_name = random.choice(company_words) + random.choice(company_words) + random.choice(company_types)
    num_employees = random.randint(20, 100)
    employees = []
    for _ in range(num_employees):
        full_name = random.choice(["赵", "钱", "孙", "李", "周", "吴", "郑", "王", "冯", "陈", "褚", "卫", "蒋", "沈", "韩", "杨", "朱", "秦", "尤", "许", "何", "吕", "施", "张", "孔", "曹", "严", "华", "金", "魏", "陶", "姜", "戚", "谢", "邹", "喻", "柏", "水", "窦", "章", "云", "苏", "潘", "葛", "奚", "范", "彭", "郎", "鲁", "韦", "昌", "马", "苗", "凤", "花", "方", "俞", "任", "袁", "柳", "酆", "鲍", "史", "唐", "费", "廉", "岑", "薛", "雷", "贺", "倪", "汤", "滕", "殷", "罗", "毕", "郝", "邬", "安", "常", "乐", "于", "时", "傅", "皮", "卞", "齐", "康", "伍", "余", "元", "卜", "顾", "孟", "平", "黄", "和", "穆", "萧", "尹", "欧阳"]) + random.choice(["伟", "芳", "娜", "敏", "静", "磊", "杰", "秀英", "娟", "强", "勇", "军", "霞", "刚", "梅", "明", "超", "秀兰", "燕", "丽", "强", "艳", "翔", "燕", "桂英", "明", "平", "红", "刚", "丽", "磊", "平", "玉", "杰", "敏", "超", "秀兰", "勇", "明", "芬", "杰", "明", "燕", "勇", "超", "霞", "秀英", "杰", "强", "敏", "丽", "静", "刚", "艳", "芳", "勇", "杰", "燕", "强", "敏", "明", "桂英", "超", "平", "明", "敏", "军", "丽", "刚", "超", "芳", "燕", "平", "桂英", "杰", "磊", "明", "超", "强", "勇", "霞", "桂英", "平", "明", "芳", "杰", "军", "强", "秀兰", "霞", "静", "明", "桂英", "超", "霞", "杰"])
        position = random.choice(["经理", "助理", "工程师", "设计师", "销售员", "会计师", "技术员", "运营专员", "产品经理", "市场专员", "人事专员"])
        employees.append({"姓名": full_name, "职位": position})
    companies.append({"公司名称": company_name, "员工信息": employees})

# 将数据写入Excel表格
company_data = []
for company in companies:
    for employee in company["员工信息"]:
        company_data.append([company["公司名称"], employee["姓名"], employee["职位"]])

df = pd.DataFrame(company_data, columns=["公司名称", "员工姓名", "职位"])
df.to_excel("companies_with_employees.xlsx", index=False, encoding='utf-8')

print("Excel表格已生成：companies_with_employees.xlsx")
