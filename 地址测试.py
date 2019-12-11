# -*- coding: UTF-8 -*-
'''
:Version: Python 3.7.4
:Date: 2019-12-11 09:19:13
:LastEditors: ChlorineLv@outlook.com
:LastEditTime: 2019-12-11 09:57:03
:Description: 执行转exe后的脚本获取地址测试
'''
import os
import sys

if __name__ == "__main__":
    # 要生成exe只需要执行 'pyinstaller -F 地址测试.py' 即可
    print(f'os.path.dirname(os.path.realpath(sys.argv[0])):\t{os.path.dirname(os.path.realpath(sys.argv[0]))}')
    print(f'os.path.dirname(os.path.realpath(sys.executable)):\t{os.path.dirname(os.path.realpath(sys.executable))}')
    print(f'os.path.abspath(__file__)\t{os.path.abspath(__file__)}')
    print(f'os.path.dirname(__file__):\t{os.path.dirname(__file__)}')
    input('回车结束')