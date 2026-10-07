# -*- coding: utf-8 -*-
"""
《心房搜查令》电影分镜 PPT 全自动生成与图层嵌入工作流
特性：
1. 锁定中国青年大学生五官特征与服饰 DNA；
2. 完美嵌入 16:9 全屏电影级分镜剧照；
3. 彻底清除模板占位框（slot_frame），解决图片被占位框遮挡问题；
4. 精准控制图层 Z-order 顺序，使顶层标题栏与底层台词字幕条无遮挡悬浮于剧照之上；
5. 完整输出包含封面、8大分镜放映页、片尾“谢谢大家”致谢页及后台全流程场控表的现场放映版 PPT。
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches

# 8 大分镜元数据清单
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
    """优先在当前目录查找，若无则在 剧照 目录查找"""
    if os.path.exists(filename):
        return filename
    sub_path = os.path.join("剧照", filename)
    if os.path.exists(sub_path):
        return sub_path
    return None

def inject_images_to_ppt(template_pptx, output_pptx):
    if not os.path.exists(template_pptx):
        print(f"[错误] 未找到模板文件: {template_pptx}")
        return False

    prs = Presentation(template_pptx)
    
    # 标准 16:9 电影全屏画幅（13.333 x 7.5 英寸）
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
            print(f"[警告] 缺少分镜图 {img_name}，第 {slide_idx+1} 页保持模板槽位")
            continue

        slide = prs.slides[slide_idx]
        spTree = slide.shapes._spTree

        # 1. 查找并移除模板中的占位框（彻底杜绝占位框遮挡剧照问题）
        slots_to_remove = []
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text
                if "预留槽位" in txt or "scene_" in txt:
                    slots_to_remove.append(shape)
        
        for slot in slots_to_remove:
            spTree.remove(slot._element)
            print(f"  [清除占位框] 第 {slide_idx+1} 页成功移除模板预留框")

        # 2. 插入高清 16:9 全屏分镜剧照
        pic = slide.shapes.add_picture(img_path, IMG_LEFT, IMG_TOP, width=IMG_WIDTH, height=IMG_HEIGHT)

        # 3. 严格设置 Z-order 渲染层级：
        # spTree 元素顺序：
        # index 0: nvGrpSpPr
        # index 1: grpSpPr
        # index 2: bg (全屏深色底板)
        # index 3: pic (剧照位于第 3 位，在底板之上、标题栏与字幕条之下)
        # index 4: top_bar (悬浮顶部标题栏)
        # index 5: sub_bar (悬浮底部字幕条)
        spTree.remove(pic._element)
        spTree.insert(3, pic._element)

        print(f"[成功注入] {img_name} -> 第 {slide_idx+1} 页 (16:9 全屏画幅，图层置于背景之上、字幕之下)")
        injected_count += 1

    prs.save(output_pptx)
    print("\n==========================================")
    print(f"交付成功！共注入 {injected_count}/8 张分镜剧照")
    print(f"总页数: {len(prs.slides)} 页")
    print(f"最终放映版文件: {output_pptx}")
    print("==========================================")
    return True

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding='utf-8')
    template = "心房搜查令_电影分镜剧照与台词字幕放映PPT.pptx"
    output = "心房搜查令_最终现场放映版.pptx"
    
    inject_images_to_ppt(template, output)
