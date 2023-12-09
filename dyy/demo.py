import sys


while True:
    material = input("请输入二要素文件的绝对路径\n")
    header = input("请输入表头文件的绝对路径\n")
    website = input("请输入要爬取成绩的所在网址\n")
    confirm = input(f"请确认输入的信息是否正确。如果正确，输入yes来运行爬虫，如果不正确，输入no来重新填写。\n二要素：{material}\n表头：{header}\n网址：{website}\n")

    if confirm == "yes":
        break
    elif confirm == "no":
        material = ""
        header = ""
        website = ""
        confirm = ""
        print("已清除信息，请重新输入")
    else:
        print("无效输入,请输入yes或no。\n")
        sys.exit()
try:
    df = pd.read_excel(f"{material}", header=0)
except NameError:
    material = input("二要素文件路径有误，请重新输入\n")
    df = pd.read_excel(f"{material}", header=0)
except Exception:
    print("废物")
    sys.exit()
