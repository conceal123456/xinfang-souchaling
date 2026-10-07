# -*- coding: utf-8 -*-
"""
把剧照和台词拼成现场放映用的 PPT。

模板先用 build_template_ppt.py 生成。这里做的是：把每页的占位框删掉，
把对应剧照铺满整页，再把它压到背景之上、标题栏和字幕条之下。
跑完输出 心房搜查令_最终现场放映版.pptx。
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches

# 八幕的图片和文案对应关系
SCENES_CONFIG = [
    {
        "id": 1,
        "slide_idx": 1,
        "filename": "scene_01.jpg",
        "desc": "第01幕·宿舍暗潮：赵学衡深夜孤光下吃面"
    },
    {
        "id": 2,
        "slide_idx": 2,
        "filename": "scene_02.jpg",
        "desc": "第02幕·门口对峙：走廊白光倾泻与手持公文夹的心理搜查官"
    },
    {
        "id": 3,
        "slide_idx": 3,
        "filename": "scene_03.jpg",
        "desc": "第03幕·宿舍交锋：100%打卡神话与医务室频开助眠药的矛盾特写"
    },
    {
        "id": 4,
        "slide_idx": 4,
        "filename": "scene_04.jpg",
        "desc": "第04幕·暗室铁柜：三把黄铜大锁与身体死守柜门的赵学衡"
    },
    {
        "id": 5,
        "slide_idx": 5,
        "filename": "scene_05.jpg",
        "desc": "第05幕·书籍雪崩：书籍倾覆与赵学衡跪地抱头痛哭"
    },
    {
        "id": 6,
        "slide_idx": 6,
        "filename": "scene_06.jpg",
        "desc": "第06幕·破防与接纳：抗拒安慰的应激防御与平视递上的搪瓷温茶"
    },
    {
        "id": 7,
        "slide_idx": 7,
        "filename": "scene_07.jpg",
        "desc": "第07幕·笛声晨光：横吹紫竹笛（徵音清心·角音疏肝）与撕碎苛刻清单"
    },
    {
        "id": 8,
        "slide_idx": 8,
        "filename": "scene_08.jpg",
        "desc": "第08幕·尾声谢幕：四人携手鞠躬致谢与《黄帝内经》经典金句大屏"
    }
]

def find_image_file(filename):
    """先在当前目录找，找不到就去 剧照/ 里找"""
    if os.path.exists(filename):
        return filename
    sub_path = os.path.join("剧照", filename)
    if os.path.exists(sub_path):
        return sub_path
    return None

def inject_images_to_ppt(template_pptx, output_pptx):
    if not os.path.exists(template_pptx):
        print(f"找不到模板：{template_pptx}")
        return False

    prs = Presentation(template_pptx)
    
    # 铺满整页：13.333 x 7.5 英寸，也就是 16:9
    IMG_LEFT = Inches(0)
    IMG_TOP = Inches(0)
    IMG_WIDTH = Inches(13.333)
    IMG_HEIGHT = Inches(7.5)

    injected_count = 0
    for scene in SCENES_CONFIG:
        img_name = scene["filename"]
        slide_idx = scene["slide_idx"]
        
        img_path = find_image_file(img_name)
        if not img_path:
            print(f"缺 {img_name}，第 {slide_idx+1} 页先空着")
            continue

        slide = prs.slides[slide_idx]
        spTree = slide.shapes._spTree

        # 占位框要先删掉，不然它会盖在剧照上
        slots_to_remove = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text
                if "预留槽位" in txt or "scene_" in txt:
                    slots_to_remove.append(shape)
        
        for slot in slots_to_remove:
            spTree.remove(slot._element)
            print(f"  第 {slide_idx+1} 页：占位框删掉了")

        # 剧照铺满整页
        pic = slide.shapes.add_picture(img_path, IMG_LEFT, IMG_TOP, width=IMG_WIDTH, height=IMG_HEIGHT)

        # 剧照要压到第 3 层。PPT 的层级就是 XML 节点的先后顺序：
        # 0/1 是分组属性，2 是背景底板，4/5 是标题栏和字幕条，
        # 所以把剧照挪到 index 3，正好夹在背景和悬浮条中间。
        spTree.remove(pic._element)
        spTree.insert(3, pic._element)

        print(f"  第 {slide_idx+1} 页 <- {img_name}")
        injected_count += 1

    prs.save(output_pptx)
    print(f"\n搞定，{injected_count}/8 张剧照填进去了")
    print(f"输出 {output_pptx}，一共 {len(prs.slides)} 页")
    return True

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding='utf-8')
    template = "心房搜查令_电影分镜剧照与台词字幕放映PPT.pptx"
    output = "心房搜查令_最终现场放映版.pptx"
    
    inject_images_to_ppt(template, output)
