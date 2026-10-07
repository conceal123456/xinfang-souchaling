# -*- coding: utf-8 -*-
"""
《心房搜查令》现场放映版 PPT 自动化流水线。
调用 build_template_ppt 中的核心生成逻辑，将最新的剧照、海报及致谢视觉封装为放映版 PPT。
"""

import os
import sys
from build_template_ppt import build_presentation

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    print("=== 开始执行《心房搜查令》放映 PPT 构建流水线 ===")
    output = "心房搜查令_最终现场放映版.pptx"
    template = "心房搜查令_电影分镜剧照与台词字幕放映PPT.pptx"
    build_presentation(output_pptx=output, template_pptx=template)
    print("=== 流水线执行完毕 ===")

if __name__ == "__main__":
    main()
