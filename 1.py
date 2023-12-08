import time
from selenium import webdriver
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook
from selenium.webdriver.chrome.options import Options
import count_runs

df = pd.read_excel("F:\\Projects\\dyy\\二要素.xlsx", header=0)
print("开始读取二要素，共" + str(len(df)) + "条数据")

# 打印第一列和第二列的值
# 创建Excel工作簿
wb = load_workbook('F:\\Projects\\dyy\\表头.xlsx')
ws = wb.active

for index, ser in df.iterrows():
    i = 0
    zh = ser.iloc[0]  # 第一列的值
    sfz = ser.iloc[1]  # 第二列的值
    # 创建 ChromeDriver 实例
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) '
                      'Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0'}

    chrome_options = Options()

    chrome_options.add_argument("--headless")  # 启动无头模式
    driver = webdriver.Chrome(options=chrome_options)

    driver.get("https://q9av4z9d.yichafen.com/qz/P4h5uRzjkt")

    # 填充身份证号
    shenfenzhenghao_input = driver.find_element(By.NAME, 's_shenfenzhenghao')
    shenfenzhenghao_input.send_keys(sfz)

    # 填充姓名
    xingming_input = driver.find_element(By.NAME, 's_xingming')
    xingming_input.send_keys(zh)

    # 点击查询按钮
    query_button = driver.find_element(By.ID, 'yiDunSubmitBtn')
    query_button.click()

    count_runs.count()  # 每爬取一次，记一次数

    # 等待页面加载完成
    time.sleep(2)
    try:
        table = driver.find_element(By.CSS_SELECTOR, '.q-r-table-panel table')
    except Exception:
        try:
            driver.refresh()
            time.sleep(3)
            driver.get("https://q9av4z9d.yichafen.com/qz/P4h5uRzjkt")

            # 填充身份证号
            shenfenzhenghao_input = driver.find_element(By.NAME, 's_shenfenzhenghao')
            shenfenzhenghao_input.send_keys(sfz)

            # 填充姓名
            xingming_input = driver.find_element(By.NAME, 's_xingming')
            xingming_input.send_keys(zh)

            # 点击查询按钮
            query_button = driver.find_element(By.ID, 'yiDunSubmitBtn')
            query_button.click()
            time.sleep(3)
            driver.refresh()
            table = driver.find_element(By.CSS_SELECTOR, '.q-r-table-panel table')
        except Exception:
            print(zh + "（" + str(sfz) + "）的成绩爬取异常！")
            continue

    # 获取表格数据
    rows = table.find_elements(By.TAG_NAME, 'tr')
    header_row = rows[0]
    data_rows = rows[1:]

    # 写入表格数据
    for row in data_rows:
        cells = row.find_elements(By.TAG_NAME, 'td')
        row_data = [cell.text for cell in cells]
        ws.append(row_data)

    # 保存Excel文件
    wb.save("成绩表.xlsx")

file_path = 'F:\\Projects\\dyy\\二要素.xlsx'
tr = pd.read_excel(file_path, engine='openpyxl')
# 获取行数
total_row = len(tr)

file_path = 'F:\\Projects\\dyy\\成绩表.xlsx'
rs = pd.read_excel(file_path, engine='openpyxl')
# 获取行数
row_count = len(rs)

print(f"爬虫运行完毕，共爬取了{total_row}次，成功{row_count}次，失败" + str(int(len(df)) - row_count) + "次。感谢您的使用！")
