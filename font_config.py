#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
中文字体配置
"""

from kivy.core.text import LabelBase
from kivy.resources import resource_add_path
from kivy.utils import platform
import os

def setup_chinese_font():
    """设置中文字体"""
    try:
        if platform == 'android':
            # Android系统字体
            resource_add_path('/system/fonts')
            LabelBase.register('Roboto', 'DroidSansFallback.ttf')
        elif platform == 'win':
            # Windows系统字体
            resource_add_path('C:/Windows/Fonts')
            # 尝试多个中文字体
            fonts = ['msyh.ttc', 'simhei.ttf', 'simsun.ttc', 'msyhbd.ttc']
            for font in fonts:
                try:
                    LabelBase.register('Roboto', font)
                    break
                except:
                    continue
        elif platform == 'linux':
            # Linux系统字体
            resource_add_path('/usr/share/fonts')
            LabelBase.register('Roboto', 'DejaVuSans.ttf')
        elif platform == 'macosx':
            # macOS系统字体
            resource_add_path('/System/Library/Fonts')
            LabelBase.register('Roboto', 'Arial.ttf')
    except Exception as e:
        print(f"字体设置失败: {e}")
        # 使用默认字体
        pass

def get_font_name():
    """获取字体名称"""
    if platform == 'win':
        return 'msyh.ttc'  # 微软雅黑
    elif platform == 'android':
        return 'DroidSansFallback.ttf'
    else:
        return 'Roboto'