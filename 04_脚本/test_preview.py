# -*- coding: utf-8 -*-
import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def test_preview():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 1. 插入全屏 16:9 剧照
    pic = slide.shapes.add_picture('scene_02.jpg', Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))

    # 2. 顶部悬浮信息条 (半透明深色玻璃质感)
    top_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.25), Inches(11.733), Inches(0.78))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = RGBColor(12, 16, 24)
    top_bar.line.color.rgb = RGBColor(60, 75, 100)
    top_bar.line.width = Pt(1.2)
    tf = top_bar.text_frame
    tf.word_wrap = True
    tf.margin_left = Inches(0.25)
    tf.margin_top = Inches(0.1)
    p0 = tf.paragraphs[0]
    p0.text = "第 02 幕 · 门口对峙  ｜  走廊白光倾泻与手持公文夹的心理搜查官"
    p0.font.size = Pt(15)
    p0.font.bold = True
    p0.font.color.rgb = RGBColor(255, 255, 255)
    p0.font.name = "Microsoft YaHei"

    p1 = tf.add_paragraph()
    p1.text = "《心房搜查令》电影分镜放映系统  ·  SCENE 02 / 08"
    p1.font.size = Pt(10.5)
    p1.font.color.rgb = RGBColor(230, 180, 70)
    p1.font.name = "Microsoft YaHei"

    # 3. 底部下沉悬浮字幕条
    sub_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(5.42), Inches(11.733), Inches(1.72))
    sub_bar.fill.solid()
    sub_bar.fill.fore_color.rgb = RGBColor(12, 16, 24)
    sub_bar.line.color.rgb = RGBColor(60, 75, 100)
    sub_bar.line.width = Pt(1.5)
    stf = sub_bar.text_frame
    stf.word_wrap = True
    stf.margin_left = Inches(0.3)
    stf.margin_top = Inches(0.18)
    stf.margin_right = Inches(0.3)

    spk_p = stf.paragraphs[0]
    spk_p.text = "【郝朋】  “赵学衡，我们今晚过来，是接到全寝室的联合反映，带着心房搜查的紧急任务来的。”"
    spk_p.font.size = Pt(14.5)
    spk_p.font.bold = True
    spk_p.font.color.rgb = RGBColor(255, 255, 255)
    spk_p.font.name = "Microsoft YaHei"
    spk_p.space_after = Pt(6)

    cue_p = stf.add_paragraph()
    cue_p.text = "[舞台提示：房门骤启，走廊森白强光斜切入室，公文夹在手目光如炬，赵学衡筷子僵在半空惊骇回眸]"
    cue_p.font.size = Pt(11.5)
    cue_p.font.color.rgb = RGBColor(230, 180, 70)
    cue_p.font.name = "Microsoft YaHei"

    test_file = "test_fullbleed.pptx"
    prs.save(test_file)
    print("Saved test_fullbleed.pptx")

if __name__ == "__main__":
    test_preview()
