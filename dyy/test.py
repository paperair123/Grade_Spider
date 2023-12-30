import time

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from openpyxl import load_workbook
import sys
import config

start_time = time.time()
user_config = config.UserConfig()
savetime = 0  # 初始化保存阈值
run_count = 0


# 计数逻辑
def count():
    global run_count
    run_count += 1
    print(f"爬虫已运行{run_count}次，" + "还需运行" + str(int(len(df)) - run_count) + "次")


# 信息采集
while True:
    user_config.get_UserConfig()

    if user_config.confirm == "yes":
        break
    elif user_config.confirm == "no":
        print("已清除信息，请重新输入")
    else:
        print("无效输入,请输入yes或no。")
        sys.exit()  # 如果两项都不符合，结束代码运行

df = pd.read_excel(f"{user_config.material}", header=0)
print("开始读取二要素，共" + str(len(df)) + "条数据")

wb = load_workbook(f'{user_config.header}')
ws = wb.active  # 创建Excel工作簿

# 创建 ChromeDriver 实例
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36 Edg/119.0.0.0'}
chrome_options = Options()
chrome_options.add_argument("--headless")  # 启动无头模式,加快爬取速度
driver = webdriver.Chrome(options=chrome_options)

driver.get(f"{user_config.website}")  # 打开目标网站

# 遍历Excel表格中每一行的值
for index, ser in df.iterrows():
    i = 0
    zh = ser.iloc[0]  # 第一列的值
    sfz = ser.iloc[1]  # 第二列的值

    # time.sleep(0.5)

    # 填充身份证号
    shenfenzhenghao_input = driver.find_element(By.NAME, 's_shenfenzhenghao')
    shenfenzhenghao_input.send_keys(sfz)

    # time.sleep(0.5)

    # 填充姓名
    xingming_input = driver.find_element(By.NAME, 's_xingming')
    xingming_input.send_keys(zh)

    # time.sleep(0.5)

    # 点击查询按钮
    query_button = driver.find_element(By.ID, 'yiDunSubmitBtn')
    query_button.click()

    # 每爬取一次，记一次数，保存阈值+1
    count()
    savetime += 1
    time.sleep(1)  # 等待页面加载完成

    try:
        table = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, '.q-r-table-panel table')))
        # driver.find_element(By.CSS_SELECTOR, '.q-r-table-panel table')
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

    # 返回查询界面
    back_button = driver.find_element(By.XPATH, '//*[@id="result_content"]/div[3]/a')
    back_button.click()

    # time.sleep(0.5)

    # 刷新页面，以便重新填充二要素
    shenfenzhenghao_reinput = driver.find_element(By.NAME, 's_shenfenzhenghao')
    shenfenzhenghao_reinput.clear()

    xingming_reinput = driver.find_element(By.NAME, 's_xingming')
    xingming_reinput.clear()

    # time.sleep(0.5)

    # 判断保存阈值是否满足要求，如若满足，自动进行保存，不满足则继续执行下面的代码
    if savetime == int(user_config.save):
        wb.save(f"{user_config.results}")
        savetime = 0
        print("已自动保存！")
    else:
        pass

wb.save(f"{user_config.results}")  # 爬取的条数可能不是2的倍数，故重新保存一遍，保险一点

end_time = time.time()
run_time = end_time - start_time
minutes = 0
seconds = run_time

if run_time >= 60:
    minutes = int(run_time / 60)
    seconds = run_time - minutes * 60
# 文件行数统计
# 获取总行数
file_path = f'{user_config.material}'
tr = pd.read_excel(file_path, engine='openpyxl')
total_row = len(tr)
# 获取实际行数
file_path = f'{user_config.results}'
rs = pd.read_excel(file_path, engine='openpyxl')
row_count = len(rs)

print(f"爬虫运行完毕，共爬取了{total_row}次，成功{row_count}次，失败" + str(
    int(len(df)) - row_count) + f"次。\n运行时长：{minutes}分钟{seconds:.3f}秒。")
