'''
Created on 2023年10月12日

@author: hewei
'''

import chardet
import matplotlib.font_manager

if __name__ == '__main__':
    print(matplotlib.font_manager.findSystemFonts(fontpaths=None, fontext='ttf'))
    # 读取选项值
    options_file = 'merchantType.txt'  # 请替换为实际的文件路径
    options = read_options_file(options_file)




def read_options_file(filename):
    with open(filename, 'rb') as f:
        rawdata = f.read()
        result = chardet.detect(rawdata)
        encoding = result['encoding']
    
    with open(filename, 'r', encoding=encoding) as file:
        options = [line.strip() for line in file.readlines()]
    return options


