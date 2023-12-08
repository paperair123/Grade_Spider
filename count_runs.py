import pandas as pd

run_count = 0
df = pd.read_excel("F:\\Projects\\dyy\\二要素.xlsx", header=0)


def count():
    global run_count
    run_count += 1
    print(f"爬虫已运行{run_count}次，" + "还需运行" + str(int(len(df)) - run_count) + "次")
