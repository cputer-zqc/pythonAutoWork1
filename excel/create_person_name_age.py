from faker import Faker

fake = Faker('zh_CN')  # 使用中文（中国）语言生成虚拟数据

# 生成10个中国人的姓名、性别、年龄、联系方式、所属单位和在岗情况
for _ in range(10):
    name = fake.name()
    gender = fake.random_element(elements=('男', '女'))
    age = fake.random_int(min=18, max=60)
    phone_number = fake.phone_number()
    company = fake.company_prefix() + "物业公司"
    on_duty = '在岗'

    print(f"name: {name}, gender: {gender}, age: {age}, contactInformation: {phone_number}, affiliatedUnit: {company}, isDuty: {on_duty}")
