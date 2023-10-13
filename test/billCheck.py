import xlrd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import chardet

import numpy as np
from decimal import Decimal

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from matplotlib import cm

from mplfonts.bin.cli import init
init()

# 使用相对路径指定字体文件位置
custom_font_path = 'wqy-microhei.ttc'

# 指定字体属性
custom_font = FontProperties(fname=custom_font_path)

# 设置中文显示
plt.rcParams['font.sans-serif'] = [custom_font.get_name()]  # 设置中文显示
plt.rcParams['axes.unicode_minus'] = False    # 解决保存图像是负号'-'显示为方块的问题


options_file = 'merchantType.txt'  # 商户类型配置

# 全局变量存储选中的文件路径和文件名
selected_file_path = ""
unique_values = []
options = []
config = {}
type_var = []
bill_contents = []

# 选取文件
def open_file_dialog():
    global selected_file_path
    filename = filedialog.askopenfilename(title="Select a file", filetypes=(("Excel files", "*.xls"), ("All files", "*.*")))
    if filename:
        selected_file_path = filename
    else:
        selected_file_path = ""

# 读取xls文件并获取唯一值列表及账单详情
def read_excel_values(filename):
    global unique_values, bill_contents
    workbook = xlrd.open_workbook(filename)
    sheet = workbook.sheet_by_index(0)
    nameValues = []
    nameAndAmouts = []
    for i in range(1, sheet.nrows):
        if Decimal(sheet.cell_value(i, 6))>0:
            nameValues.append(sheet.cell_value(i, 2))
            nameAndAmouts.append({'name': sheet.cell_value(i, 2), 'amount': Decimal(sheet.cell_value(i, 6)), 'type': '其他'})

    unique_values = list(set(nameValues))
    bill_contents = nameAndAmouts

# 读取选项值文件
def read_options_file(options_file):
    with open(options_file, 'rb') as f:
        rawdata = f.read()
        result = chardet.detect(rawdata)
        encoding = result['encoding']

    with open(options_file, 'r', encoding=encoding) as file:
        options = [line.strip() for line in file.readlines()]
    return options

# 匹配唯一值和JSON配置
def match_with_config(value, config):
    return config.get(value, u'其他')

# 保存配置到JSON文件
def save_config_to_json(config):
    with open('config.json', 'w') as json_file:
        json.dump(config, json_file)

# 更新Canvas大小的函数
def update_canvas(event):
    #canvas_widget.configure(scrollregion=canvas_widget.bbox('all'))
    canvas_width = event.width
    #canvas_widget.itemconfig(inner_frame_id, width=canvas_width)

# 创建GUI界面
root = tk.Tk()
root.title('商户类型配置')
root.geometry("800x600")

# 创建一个Frame用于放置数据、下拉框和滚动条
frame = ttk.Frame(root)
frame.pack(padx=10, pady=10, fill='both', expand=True)

# 创建一个Canvas用于放置数据和下拉框，并添加垂直滚动条
canvas = tk.Canvas(frame)
canvas.pack(side='left', fill='both', expand=True)

# 创建滚动条
scrollbar = ttk.Scrollbar(frame, orient='vertical', command=canvas.yview)
scrollbar.pack(side='right', fill='y')

# 配置Canvas与滚动条的关联
canvas.configure(yscrollcommand=scrollbar.set)
canvas.bind('<Configure>', update_canvas)

# 创建一个Frame
inner_frame = ttk.Frame(canvas)
inner_frame_id = canvas.create_window((0, 0), window=inner_frame, anchor='nw')

# 在Canvas上显示Matplotlib图形
def show_matplotlib_plot():
    global bill_contents

    # 计算每种类型的总金额
    total_amounts = {}
    for item in bill_contents:
        name = item['name']
        amount = item['amount']
        matched_type = match_with_config(name, config)
        total_amounts[matched_type] = total_amounts.get(matched_type,0) + amount

    # 获取标签和对应的金额
    labels = list(total_amounts.keys())
    amounts = list(total_amounts.values())

    # 创建Matplotlib图形对象
    fig = Figure(figsize=(5, 4), dpi=100)
    ax = fig.add_subplot(111)
    
    ax.pie(amounts, labels=labels, autopct='%1.1f%%', startangle=140)
    ax.set_title('支出分析')
    
#     # 使用自定义autopct函数显示具体数字和百分比
#     def func(pct, allvalues):
#         absolute = int(pct / 100. * np.sum(allvalues))
#         return f"{pct:.1f}%\n({absolute:d}元)"
# 
#     wedges, texts, autotexts = ax.pie(amounts, labels=labels, autopct=lambda pct: func(pct, amounts), startangle=140, colors=cm.tab20.colors)
#     ax.set_title('支出分析')
# 
#     # 为每个扇形添加具体数字和百分比标签
#     for i, text in enumerate(texts):
#         text.set(size=12, weight='bold', color='black')
#         autotexts[i].set(size=10, weight='bold', color='black')

    # 先清除其他组件
    for widget in inner_frame.winfo_children():
        widget.destroy()
        
    # 打开按钮
    open_button = tk.Button(inner_frame, text="Open File Dialog", command=open_file_dialog)
    open_button.grid(row=0, column=0, columnspan=1, pady=(10, 0), sticky='ew')

    # 添加确认按钮
    confirm_button = ttk.Button(inner_frame, text='Confirm', command=confirm_file)
    confirm_button.grid(row=0, column=1, columnspan=1, pady=(10, 0), sticky='ew')
      
    # 创建FigureCanvasTkAgg对象，将图形嵌入到Tkinter应用程序中  
    canvas_widget = FigureCanvasTkAgg(fig, master=inner_frame)
    canvas_widget.get_tk_widget().grid(row=1, column=0, padx=10, pady=10, sticky=tk.W+tk.E+tk.N+tk.S)
    
    #明细数据
    inner_frame2 = ttk.Frame(inner_frame)
    inner_frame2.grid(row=1, column=1, padx=10, pady=10, sticky='nsew')
    types = list(total_amounts.keys())
    for idx, merType in enumerate(types):
        amount = total_amounts[merType]
        label = tk.Label(inner_frame2, text=f' {merType}: {amount} ')
        label.grid(row=idx, column=0, padx=10, pady=10, sticky='nsew')
        

    # 更新Canvas和Scrollbar
    canvas.update_idletasks()  # 更新Canvas以反映更改
    canvas.config(scrollregion=canvas.bbox("all"))  # 更新滚动区域

    # 根据Canvas内容高度确保Scrollbar的可见性
    _, _, canvas_width, canvas_height = canvas.bbox("all")
    canvas.config(width=canvas_width, height=canvas_height)
    if canvas_height < 600:
        scrollbar.pack_forget()  # 如果内容适应Canvas，隐藏Scrollbar
    else:
        scrollbar.pack(side='right', fill='y')  # 如果需要，显示Scrollbar

# 显示每个唯一值及对应的下拉框
def display_merchants_and_dropdowns():
    global unique_values, options, config, type_var

    # Clear existing widgets in the inner frame
    for widget in inner_frame.winfo_children():
        widget.destroy()

    # 打开按钮
    open_button = tk.Button(inner_frame, text="Open File Dialog", command=open_file_dialog)
    open_button.grid(row=0, column=0, columnspan=1, pady=(10, 0), sticky='ew')

    # 添加确认按钮
    confirm_button = ttk.Button(inner_frame, text='Confirm', command=confirm_file)
    confirm_button.grid(row=0, column=1, columnspan=1, pady=(10, 0), sticky='ew')

    # 显示每个唯一值及对应的下拉框
    type_var = []

    for idx, value in enumerate(unique_values):
        # 显示唯一值
        label = ttk.Label(inner_frame, text=f'Value {idx + 1}: {value}')
        label.grid(row=idx + 1, column=0, padx=5, pady=5, sticky='w')

        # 匹配配置
        matched_type = match_with_config(value, config)

        # 创建下拉框
        type_var.append(tk.StringVar())
        type_var[idx].set(matched_type)
        type_dropdown = ttk.Combobox(inner_frame, textvariable=type_var[idx], values=options)
        type_dropdown.grid(row=idx + 1, column=1, padx=5, pady=5, sticky='w')

    # 保存按钮
    save_button = ttk.Button(inner_frame, text='Save', command=save_button_click)
    save_button.grid(row=len(unique_values) + 1, column=0, columnspan=1, pady=(10, 0), sticky='ew')

    # 计算按钮
    calculate_button = ttk.Button(inner_frame, text='Calculate', command=show_matplotlib_plot)
    calculate_button.grid(row=len(unique_values) + 1, column=1, columnspan=1, pady=(10, 0), sticky='ew')

    # 更新Canvas和Scrollbar
    canvas.update_idletasks()  # 更新Canvas以反映更改
    canvas.config(scrollregion=canvas.bbox("all"))  # 更新滚动区域

    # 根据Canvas内容高度确保Scrollbar的可见性
    _, _, canvas_width, canvas_height = canvas.bbox("all")
    canvas.config(width=canvas_width, height=canvas_height)
    if canvas_height < 600:
        scrollbar.pack_forget()  # 如果内容适应Canvas，隐藏Scrollbar
    else:
        scrollbar.pack(side='right', fill='y')  # 如果需要，显示Scrollbar

# 在确定文件后调用此函数，显示商户清单和下拉框
def confirm_file():
    global filename, unique_values, options, config

    # 读取唯一值列表 unique_values及账单清单
    read_excel_values(selected_file_path)

    # 读取选项值
    options = read_options_file(options_file)

    # 读取配置文件或创建空配置
    config = {}
    if os.path.exists('config.json'):
        with open('config.json') as json_file:
            config = json.load(json_file)

    # 调用函数显示商户清单和下拉框
    display_merchants_and_dropdowns()

    # 更新UI，确保显示
    root.update_idletasks()

# 保存按钮点击事件
def save_button_click():
    for idx, value in enumerate(unique_values):
        selected_type = type_var[idx].get()
        config[value] = selected_type
    save_config_to_json(config)
    messagebox.showinfo('Information', 'Configuration saved successfully.')

# 主程序入口
if __name__ == "__main__":
    # 打开按钮
    open_button = tk.Button(inner_frame, text="Open File Dialog", command=open_file_dialog)
    open_button.grid(row=len(unique_values), column=0, columnspan=1, pady=(10, 0), sticky='ew')

    # 添加确认按钮
    confirm_button = ttk.Button(inner_frame, text='Confirm', command=confirm_file)
    confirm_button.grid(row=len(unique_values), column=1, columnspan=1, pady=(10, 0), sticky='ew')

    # 运行Tkinter主事件循环
    root.mainloop()
