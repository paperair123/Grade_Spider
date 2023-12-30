# 这是一个配置类，用于存放爬虫所需要的一些变量和函数。


class UserConfig:
    def __init__(self):
        self.material = None
        self.header = None
        self.results = None
        self.save = None
        self.website = None
        self.confirm = None

    def get_UserConfig(self):
        self.material = input("请输入二要素文件的绝对路径\n")  # 采集姓名和身份证号
        self.header = input("请输入表头文件的绝对路径\n")  # 采集表头信息
        self.results = input(
            "请输入保存成绩的表格名称（示例：成绩表.xlsx）。注意，该文件必须和代码放在同一路径下。\n")  # 采集目标成绩保存文件
        self.save = input("请输入自动保存条数\n")  # 设置自动保存条数
        self.website = input("请输入要爬取成绩的所在网址\n")  # 设置爬取网址
        # 信息确认，输入yes结束循环，输入no则将所有变量赋值为空字符串，重新填充
        self.confirm = input(f"请确认输入的信息是否正确。如果正确，输入yes来运行爬虫，如果不正确，输入no来重新填写。"
                             f"\n二要素：{self.material}\n"
                             f"表头：{self.header}\n"
                             f"网址：{self.website}\n"
                             f"成绩表：{self.results}\n"
                             f"每爬取{self.save}条数据自动保存一次\n")
