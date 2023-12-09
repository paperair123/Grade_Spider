import pandas as pd

run_count = 0
df = pd.read_excel(f"{material}", header=0)


def count():
    global run_count
    run_count += 1
    print(f"爬虫已运行{run_count}次，" + "还需运行" + str(int(len(df)) - run_count) + "次")


savetime = 0


def autosave(wb):
    global savetime
    savetime += 1
    if savetime == save:
        print(f"已达到自动保存条数，开始自动保存")
        wb.save(f"{results}")
        print("自动保存成功！")
