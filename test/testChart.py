import os
import matplotlib.pyplot as plt
from matplotlib.font_manager import FontProperties
from mplfonts.bin.cli import init
init()
# 使用相对路径指定字体文件位置
custom_font_path = 'wqy-microhei.ttc'

#from mplfonts import use_font

#use_font('WenQuanYi Micro Hei')#指定中文字体


# 指定字体属性
custom_font = FontProperties(fname=custom_font_path)

# 设置中文显示
plt.rcParams['font.sans-serif'] = [custom_font.get_name()]  # 设置中文显示
plt.rcParams['axes.unicode_minus'] = False    # 解决保存图像是负号'-'显示为方块的问题

# Sample data
labels = ['A', 'B', 'C', 'D', '狗']
values = [10, 30, 20, 25, 15]

# Create a pie chart
plt.figure(figsize=(8, 6))
plt.pie(values, labels=labels, autopct='%1.1f%%')
plt.title('Pie Chart Example', fontproperties=custom_font)

# Show the plot
plt.show()
