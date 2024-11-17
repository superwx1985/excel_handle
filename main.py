import math
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox


def ceiling(number, significance):
    if significance == 0:
        raise ValueError("significance 不能为 0")
    return math.ceil(number / significance) * significance


freight_map = {
    # "key": [增量(kg), 首重价格, 增量价格]
    "3724": [0.1, 17, 2],
    "4811-1": [0.5, 32, 4],
    "4811-2": [0.5, 34, 4],
}
extra_fee_dict = {
    "北海道": 20,
    "沖縄": 70,
}


# 运费计算函数
def calculate_freight(weight, size, waybill_number):
    if str(waybill_number).startswith("3724"):
        key = "3724"
    else:
        weight = calculate_size_weight(weight, size)
        if weight < 5:
            key = "4811-1"
        else:
            key = "4811-2"
    plus_weight = freight_map[key][0]
    base_freight = freight_map[key][1]
    plus_freight = freight_map[key][2]
    freight = base_freight + (ceiling(weight, plus_weight)/plus_weight-1)*plus_freight
    return pd.Series([weight, freight], index=["计费重", "运费"])


def calculate_size_weight(actual_weight, size):
    actual_weight = float(actual_weight)
    if isinstance(size, str) and size != "" and len(size.split("*")) == 4:
        length, width, height, quantity = map(int, size.split("*"))
        if length + width + height > 100:
            size_weight = length * width * height / 6000 * quantity
            weight = (actual_weight + size_weight) / 2
        else:
            weight = actual_weight
    else:
        weight = actual_weight
    return weight


# 额外费用计算函数
def calculate_extra_fee(address):
    key, fee = "", 0
    for _key, _fee in extra_fee_dict.items():
        if address.startswith(_key):
            key, fee = _key, _fee
            break
    return pd.Series([key, fee], index=["地区", "额外费用"])


# 处理 Excel 文件函数
def process_excel_files(file_a, file_b, file_c, save_path):
    try:
        # 读取 A 表，使用 "黑猫单号" 作为运单号
        df_a = pd.DataFrame()
        for f in file_a.split("|"):
            _ = pd.read_excel(f)
            # 合并数据
            df_a = pd.concat([df_a, _], ignore_index=True)
        df_a = df_a.rename(columns={"黑猫单号": "运单号"})

        # 读取 B 表，从第 6 行开始读取
        df_b = pd.DataFrame()
        for f in file_b.split("|"):
            _ = pd.read_excel(f, skiprows=5)
            # 合并数据
            df_b = pd.concat([df_b, _], ignore_index=True)

        # 读取 C 表，使用 "系统单号" 作为运单号
        df_c = pd.DataFrame()
        for f in file_c.split("|"):
            _ = pd.read_excel(f)
            # 合并数据
            df_c = pd.concat([df_c, _], ignore_index=True)
        df_c = df_c.rename(columns={"系统单号": "运单号"})

        # 读取 C 表，跳过第一个 Sheet，处理包含月份的 Sheet
        # excel_c = pd.ExcelFile(file_c)
        # df_c_list = []
        # for sheet_name in excel_c.sheet_names[0:]:
        #     if any(month in sheet_name for month in ["01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12"]):
        #         df_c = pd.read_excel(file_c, sheet_name=sheet_name)
        #         df_c_list.append(df_c)
        #
        # # 合并所有 C 表的数据
        # df_c = pd.concat(df_c_list)
        # # df_c = df_c.rename(columns={"系统单号": "运单号"})

        # 合并表格
        df = pd.merge(df_a, df_b, on="运单号", how="inner")
        df = pd.merge(df, df_c, on="运单号", how="inner")

        # 计算费用
        df["尺寸j"] = df["尺寸"]
        df["实重j"] = df["实重"]
        df[["计费重j", "运费j"]] = df.apply(lambda x: calculate_freight(x["实重"], x["尺寸"], x["运单号"]), axis=1)
        df[["地区j", "额外费用j"]] = df["收件人地址"].apply(calculate_extra_fee)
        df["总费用j"] = df["运费j"] + df["额外费用j"]

        # 保存结果为 Excel 文件
        df.to_excel(save_path, index=False)
        messagebox.showinfo("完成", f"处理完成，文件已保存至：{save_path}")

    except Exception as e:
        messagebox.showerror("错误", str(e))
        raise e


# 选择文件函数
def select_file(entry):
    file_paths = filedialog.askopenfilenames(filetypes=[("Excel Files", "*.xlsx *.xls")])
    entry.delete(0, tk.END)
    file_paths = "|".join(file_paths)
    entry.insert(0, file_paths)


# 选择保存路径函数
def select_save_path(entry):
    save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    entry.delete(0, tk.END)
    entry.insert(0, save_path)


# 执行处理函数
def start_processing():
    file_a = entry_a.get()
    file_b = entry_b.get()
    file_c = entry_c.get()
    save_path = entry_save.get()

    if not (file_a and file_b and file_c and save_path):
        messagebox.showerror("错误", "请完整选择所有文件路径和保存路径。")
        return

    process_excel_files(file_a, file_b, file_c, save_path)


# 创建 GUI 界面
root = tk.Tk()
root.title("龙猫王的运费计算器")

# A 表选择框
tk.Label(root, text="选择 A 表 (管理号和品名)：").grid(row=0, column=0, padx=10, pady=5)
entry_a = tk.Entry(root, width=50)
entry_a.grid(row=0, column=1, padx=10, pady=5)
button_a = tk.Button(root, text="浏览", command=lambda: select_file(entry_a))
button_a.grid(row=0, column=2, padx=10, pady=5)

# B 表选择框
tk.Label(root, text="选择 B 表 (重量和尺寸)：").grid(row=1, column=0, padx=10, pady=5)
entry_b = tk.Entry(root, width=50)
entry_b.grid(row=1, column=1, padx=10, pady=5)
button_b = tk.Button(root, text="浏览", command=lambda: select_file(entry_b))
button_b.grid(row=1, column=2, padx=10, pady=5)

# C 表选择框
tk.Label(root, text="选择 C 表 (提单号和地址)：").grid(row=2, column=0, padx=10, pady=5)
entry_c = tk.Entry(root, width=50)
entry_c.grid(row=2, column=1, padx=10, pady=5)
button_c = tk.Button(root, text="浏览", command=lambda: select_file(entry_c))
button_c.grid(row=2, column=2, padx=10, pady=5)

# 保存路径选择框
tk.Label(root, text="选择保存文件路径：").grid(row=3, column=0, padx=10, pady=5)
entry_save = tk.Entry(root, width=50)
entry_save.grid(row=3, column=1, padx=10, pady=5)
button_save = tk.Button(root, text="浏览", command=lambda: select_save_path(entry_save))
button_save.grid(row=3, column=2, padx=10, pady=5)

# 开始处理按钮
button_start = tk.Button(root, text="开始处理", command=start_processing, bg="lightblue")
button_start.grid(row=4, column=0, columnspan=3, pady=20)


if __name__ == '__main__':
    root.mainloop()
    # l = [
    #     # ("0.739",  "65*55*3*1"),
    #     # ("0.739",  ""),
    #     ("4.443",  "36*35*33*1"),
    # ]
    # for i in l:
    #     actual_weight, size = i
    #     print(calculate_size_weight(actual_weight, size))
    #     print(calculate_freight(actual_weight, size))
