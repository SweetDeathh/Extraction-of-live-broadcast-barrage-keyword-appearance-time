import os
import xml.etree.ElementTree as ET
import pandas as pd
import tkinter as tk
from tkinter import filedialog, scrolledtext
import subprocess
import argparse

# 创建一个空DataFrame，用于存储结果
result_df = pd.DataFrame(columns=['关键词', '时间', '文件名', '弹幕'])

# 获取命令行参数
parser = argparse.ArgumentParser(description='Process Bilibili data.')
parser.add_argument('--folder', default=os.getcwd(), help='Path to the root folder')  # 默认为当前文件夹
args, _ = parser.parse_known_args()
root_folder_path = args.folder

# 初始关键词数组
initial_keywords = ["示例数组"]
keywords = initial_keywords.copy()

def process_folder(folder_path, output_text):
    global result_df  # 增加这一行
    # 遍历文件夹下所有文件夹
    for folder_path, dirs, files in os.walk(folder_path):
        # 遍历当前文件夹下的所有XML文件
        for file_name in files:
            if file_name.endswith('.xml'):
                # 拼接文件路径
                file_path = os.path.join(folder_path, file_name)
                # 解析XML文件
                output_text.insert(tk.END, f"Processing file: {file_path}\n")
                try:
                    tree = ET.parse(file_path)
                    root = tree.getroot()
                except ET.ParseError as e:
                    output_text.insert(tk.END, f"Error parsing XML file {file_path}: {e}\n")
                    continue
                # 遍历XML文件中的所有弹幕节点
                for danmu in root.iter('d'):
                    # 获取弹幕文本和出现时间
                    danmu_text = danmu.text
                    danmu_time = danmu.get('p').split(',')[0]
                    # 将时间格式化为时:分:秒
                    danmu_time = f'{int(float(danmu_time) / 3600):02d}:{int((float(danmu_time) % 3600) / 60):02d}:{int(float(danmu_time) % 60):02d}'
                    # 判断弹幕文本是否包含关键词
                    for keyword in keywords:
                        if keyword in danmu_text:
                            # 将结果添加到DataFrame中
                            new_df = pd.DataFrame({
                                '关键词': [keyword],
                                '时间': [danmu_time],
                                '文件名': [file_name],
                                '弹幕': [danmu_text]
                            })
                            result_df = pd.concat([result_df, new_df], ignore_index=True)
                            output_text.insert(tk.END, f"Found keyword '{keyword}' in file {file_name} at time {danmu_time}\n")
    result_df.to_excel(os.path.join(os.getcwd(), 'result.xlsx'), index=False)
    output_text.insert(tk.END, "处理完成，文件储存为 'result.xlsx'\n")

def browse_folder():
    folder_selected = filedialog.askdirectory()
    entry_path.delete(0, tk.END)
    entry_path.insert(0, folder_selected)

def process_folder_gui():
    folder_path = entry_path.get()
    process_folder(folder_path, text_output)

def open_excel_file():
    excel_file_path = os.path.join(os.getcwd(), 'result.xlsx')
    print(f"Opening Excel file: {excel_file_path}")
    try:
        os.startfile(excel_file_path)
    except Exception as e:
        print(f"Error opening Excel file: {e}")

def update_keywords():
    new_keywords = entry_keywords.get()
    global keywords
    keywords = new_keywords.split(',')

# 默认文件夹路径
default_folder_path = os.getcwd()

# 创建 GUI 窗口
root = tk.Tk()
root.title("表格文本时间提取")

# 添加文件夹路径输入框和按钮
label_path = tk.Label(root, text="选择文件夹路径:")
label_path.grid(row=0, column=0, pady=10)

entry_path = tk.Entry(root, width=50)
entry_path.insert(0, default_folder_path)  # 设置默认文件夹路径
entry_path.grid(row=0, column=1)

button_browse = tk.Button(root, text="浏览", command=browse_folder)
button_browse.grid(row=0, column=2)

# 添加关键词输入框和更新按钮
label_keywords = tk.Label(root, text="输入关键词（逗号分隔）:")
label_keywords.grid(row=1, column=0, pady=10)

entry_keywords = tk.Entry(root, width=50)
entry_keywords.insert(0, ','.join(initial_keywords))
entry_keywords.grid(row=1, column=1)

button_update_keywords = tk.Button(root, text="更新关键词", command=update_keywords)
button_update_keywords.grid(row=1, column=2)

# 添加处理按钮
button_process = tk.Button(root, text="处理文件夹", command=process_folder_gui)
button_process.grid(row=2, column=0, columnspan=3, pady=10)

# 添加打开Excel按钮
button_open_excel = tk.Button(root, text="打开Excel文件", command=open_excel_file)
button_open_excel.grid(row=3, column=0, columnspan=3, pady=10)

# 添加用于显示处理过程的文本框
text_output = scrolledtext.ScrolledText(root, width=60, height=15)
text_output.grid(row=4, column=0, columnspan=3, pady=10)

# 运行 GUI
root.mainloop()
