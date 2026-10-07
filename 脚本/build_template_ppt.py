# -*- coding: utf-8 -*-
"""
生成《心房搜查令》最新现场放映版 PPT (11 页)。

规范与要求：
1. 第 1 页封面：使用生图模型生成的 16:9 电影级海报图 cover.jpg，叠加微光质感标题卡片；
2. 第 2-9 页八大分镜：纯场景大图模式，去掉所有顶部和底部文字，画面铺满 16:9 (13.333" x 7.5")，沉浸式放映；
3. 第 10 页致谢页：使用生图模型生成的 16:9 剧院谢幕美图 ending.jpg，左右翼对称演职人员与鸣谢卡片，中心留白呈现舞台紫竹笛与清茶；
4. 第 11 页后台场控表：保留完整的 9 行 6 列后台视听与道具调度执行表。
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

def build_presentation(output_pptx="心房搜查令_最终现场放映版.pptx", template_pptx="心房搜查令_电影分镜剧照与台词字幕放映PPT.pptx"):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]

    # 标准配色
    COLOR_BG = RGBColor(11, 14, 20)          # 极黑深蓝背景底色
    COLOR_GOLD = RGBColor(230, 180, 70)      # 琥珀金
    COLOR_WHITE = RGBColor(255, 255, 255)    # 纯白
    COLOR_SUBTITLE = RGBColor(220, 225, 235) # 字幕白
    COLOR_DIM = RGBColor(140, 148, 165)      # 辅文字灰
    COLOR_CARD_BG = RGBColor(12, 16, 24)     # 半透明质感深色底板
    COLOR_BORDER = RGBColor(50, 65, 90)      # 微光边缘色
    COLOR_GOLD_BORDER = RGBColor(180, 140, 50)# 金色边框

    def get_img_path(filename):
        paths_to_check = [
            filename,
            os.path.join("剧照", filename),
            os.path.join(os.path.dirname(__file__), "..", "剧照", filename),
            os.path.join(os.path.dirname(__file__), filename)
        ]
        for p in paths_to_check:
            if os.path.exists(p):
                return os.path.abspath(p)
        return None

    # ==================== 第 1 页：封面 (cover.jpg) ====================
    slide_cov = prs.slides.add_slide(blank_layout)
    cover_img = get_img_path("cover.jpg")
    if cover_img:
        slide_cov.shapes.add_picture(cover_img, Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
    else:
        bg_cov = slide_cov.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg_cov.fill.solid()
        bg_cov.fill.fore_color.rgb = COLOR_BG

    # 封面电影级优雅字幕（左右角与底部留白，完全避开海报正中“心房搜查令”金字与人物）
    tb_cov_top_l = slide_cov.shapes.add_textbox(Inches(0.6), Inches(0.25), Inches(5.0), Inches(0.4))
    tf_cov_l = tb_cov_top_l.text_frame
    tf_cov_l.word_wrap = True
    p_l = tf_cov_l.paragraphs[0]
    p_l.text = "第十四届校园心理情景剧大赛 · 舞台放映"
    p_l.font.size = Pt(11.5)
    p_l.font.color.rgb = COLOR_GOLD
    p_l.font.bold = True
    p_l.font.name = "Microsoft YaHei"

    tb_cov_top_r = slide_cov.shapes.add_textbox(Inches(7.733), Inches(0.25), Inches(5.0), Inches(0.4))
    tf_cov_r = tb_cov_top_r.text_frame
    tf_cov_r.word_wrap = True
    p_r = tf_cov_r.paragraphs[0]
    p_r.text = "2024级数据科学与大数据技术班"
    p_r.font.size = Pt(11.5)
    p_r.font.color.rgb = COLOR_SUBTITLE
    p_r.font.name = "Microsoft YaHei"
    p_r.alignment = PP_ALIGN.RIGHT

    tb_cov_bot = slide_cov.shapes.add_textbox(Inches(0.8), Inches(7.12), Inches(11.733), Inches(0.35))
    tf_cov_bot = tb_cov_bot.text_frame
    tf_cov_bot.word_wrap = True
    p_bot = tf_cov_bot.paragraphs[0]
    p_bot.text = "电影分镜剧照与现场放映演示文稿 · 最终放映版  ｜  《心房搜查令》  ｜  16:9 全屏画幅"
    p_bot.font.size = Pt(9.5)
    p_bot.font.color.rgb = COLOR_DIM
    p_bot.font.name = "Microsoft YaHei"
    p_bot.alignment = PP_ALIGN.CENTER

    # ==================== 第 2 至 9 页：八大分镜纯场景图 ====================
    scenes = [
        ("scene_01.jpg", "第01幕·宿舍暗潮"),
        ("scene_02.jpg", "第02幕·门口对峙"),
        ("scene_03.jpg", "第03幕·宿舍交锋"),
        ("scene_04.jpg", "第04幕·暗室铁柜"),
        ("scene_05.jpg", "第05幕·书籍雪崩"),
        ("scene_06.jpg", "第06幕·破防与接纳"),
        ("scene_07.jpg", "第07幕·笛声晨光"),
        ("scene_08.jpg", "第08幕·尾声谢幕")
    ]

    for idx, (img_file, sc_title) in enumerate(scenes, start=1):
        slide = prs.slides.add_slide(blank_layout)
        img_path = get_img_path(img_file)
        if img_path:
            slide.shapes.add_picture(img_path, Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
            print(f"  ✓ 第 {idx + 1:02d} 页：已载入 16:9 纯场景大图 [{img_file}]")
        else:
            print(f"  ! 警告：未找到图片 [{img_file}]")
            bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
            bg.fill.solid()
            bg.fill.fore_color.rgb = COLOR_BG

    # ==================== 第 10 页：片尾致谢页 (ending.jpg) ====================
    slide_thx = prs.slides.add_slide(blank_layout)
    ending_img = get_img_path("ending.jpg")
    if ending_img:
        slide_thx.shapes.add_picture(ending_img, Inches(0), Inches(0), width=Inches(13.333), height=Inches(7.5))
    else:
        bg_thx = slide_thx.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg_thx.fill.solid()
        bg_thx.fill.fore_color.rgb = COLOR_BG

    # 顶部金句横幅（居中悬浮于顶光之上，庄严宁静）
    quote_bar = slide_thx.shapes.add_textbox(Inches(1.5), Inches(0.2), Inches(10.333), Inches(0.5))
    qtf = quote_bar.text_frame
    qtf.word_wrap = True
    qp = qtf.paragraphs[0]
    qp.text = "“恬淡虚无，真气从之；精神内守，病安从来！”  ——《黄帝内经·素问》"
    qp.font.size = Pt(13.5)
    qp.font.bold = True
    qp.font.color.rgb = COLOR_GOLD
    qp.font.name = "Microsoft YaHei"
    qp.alignment = PP_ALIGN.CENTER

    # 左翼：演职人员表（置于左侧红丝绒暗影区域，纯透明悬浮排版）
    card_left = slide_thx.shapes.add_textbox(Inches(0.6), Inches(1.75), Inches(4.3), Inches(4.8))
    cl_tf = card_left.text_frame
    cl_tf.word_wrap = True
    cl_tf.margin_left = Inches(0.1)
    cl_tf.margin_right = Inches(0.1)
    cl_tf.margin_top = Inches(0.1)

    cl_p0 = cl_tf.paragraphs[0]
    cl_p0.text = "【 演职人员表 · CAST & CREW 】"
    cl_p0.font.size = Pt(13)
    cl_p0.font.bold = True
    cl_p0.font.color.rgb = COLOR_GOLD
    cl_p0.font.name = "Microsoft YaHei"
    cl_p0.space_after = Pt(10)

    cast_items = [
        ("编剧 / 导演", "2024级数据科学与大数据技术班 创编组"),
        ("赵学衡 (饰)", "中医药大三医学生 · 执念释怀与自省"),
        ("郝  朋 (饰)", "班级心理委员 · 心房搜查与温情引航"),
        ("程  澄 (饰)", "室友甲 · 现实客观观察者"),
        ("林  凡 (饰)", "室友乙 · 竹笛现场演奏与五音情志调和")
    ]
    for role, desc in cast_items:
        p_c = cl_tf.add_paragraph()
        r1 = p_c.add_run()
        r1.text = f"• {role}："
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_WHITE
        r1.font.name = "Microsoft YaHei"
        r2 = p_c.add_run()
        r2.text = f" {desc}"
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_SUBTITLE
        r2.font.name = "Microsoft YaHei"
        p_c.space_after = Pt(6)

    # 右翼：特别鸣谢与致谢寄语（置于右侧红丝绒暗影区域，纯透明悬浮排版）
    card_right = slide_thx.shapes.add_textbox(Inches(8.433), Inches(1.75), Inches(4.3), Inches(4.8))
    cr_tf = card_right.text_frame
    cr_tf.word_wrap = True
    cr_tf.margin_left = Inches(0.1)
    cr_tf.margin_right = Inches(0.1)
    cr_tf.margin_top = Inches(0.1)

    cr_p0 = cr_tf.paragraphs[0]
    cr_p0.text = "【 特别鸣谢 · SPECIAL THANKS 】"
    cr_p0.font.size = Pt(13)
    cr_p0.font.bold = True
    cr_p0.font.color.rgb = COLOR_GOLD
    cr_p0.font.name = "Microsoft YaHei"
    cr_p0.space_after = Pt(10)

    thanks_items = [
        ("指导单位", "校学生工作部（处） / 心理健康教育与咨询中心"),
        ("学术支持", "传统中医药五音疗疾与情志相胜指导组"),
        ("剧组班级", "2024级数据科学与大数据技术班 全体同学"),
        ("技术呈现", "16:9 沉浸式电影分镜视听放映系统"),
        ("致谢寄语", "《心房搜查令》全剧终 · 谢谢大家！诚挚感谢悉心陪伴！")
    ]
    for role, desc in thanks_items:
        p_t = cr_tf.add_paragraph()
        r1 = p_t.add_run()
        r1.text = f"• {role}："
        r1.font.size = Pt(10.5)
        r1.font.bold = True
        r1.font.color.rgb = COLOR_WHITE
        r1.font.name = "Microsoft YaHei"
        r2 = p_t.add_run()
        r2.text = f" {desc}"
        r2.font.size = Pt(10)
        r2.font.color.rgb = COLOR_SUBTITLE
        r2.font.name = "Microsoft YaHei"
        p_t.space_after = Pt(6)

    # ==================== 第 11 页：后台全流程场控与视听调度表 ====================
    slide_back = prs.slides.add_slide(blank_layout)
    bg_back = slide_back.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg_back.fill.solid()
    bg_back.fill.fore_color.rgb = COLOR_BG
    bg_back.line.color.rgb = COLOR_BG

    # 场控表标题
    tb_b_title = slide_back.shapes.add_textbox(Inches(0.8), Inches(0.25), Inches(11.733), Inches(0.8))
    tf_b = tb_b_title.text_frame
    tf_b.word_wrap = True
    pb0 = tf_b.paragraphs[0]
    pb0.text = "《心房搜查令》后台全流程场控与视听调度表"
    pb0.font.size = Pt(20)
    pb0.font.bold = True
    pb0.font.color.rgb = COLOR_WHITE
    pb0.font.name = "Microsoft YaHei"

    pb1 = tf_b.add_paragraph()
    pb1.text = "舞台剧组全流程精细化执行表 · 覆盖 8 幕分镜角色走位、道具流转、灯光预设与音效节拍点"
    pb1.font.size = Pt(11)
    pb1.font.color.rgb = COLOR_GOLD
    pb1.font.name = "Microsoft YaHei"

    # 场控表格（9行6列：1表头 + 8幕内容）
    table_shape = slide_back.shapes.add_table(9, 6, Inches(0.6), Inches(1.15), Inches(12.133), Inches(5.95))
    tbl = table_shape.table

    col_widths = [Inches(1.3), Inches(1.4), Inches(2.3), Inches(2.2), Inches(2.3), Inches(2.633)]
    for ci, w in enumerate(col_widths):
        tbl.columns[ci].width = w

    headers = ["幕次 / 场景", "登场角色", "核心道具调度", "舞台灯光设计", "现场音效与配乐", "场控与调度要点"]
    for ci, h in enumerate(headers):
        cell = tbl.cell(0, ci)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(25, 35, 52)
        cell.text_frame.margin_left = Inches(0.08)
        cell.text_frame.margin_right = Inches(0.08)
        cell.text_frame.margin_top = Inches(0.08)
        cell.text_frame.margin_bottom = Inches(0.08)
        p = cell.text_frame.paragraphs[0]
        p.text = h
        p.font.size = Pt(11)
        p.font.bold = True
        p.font.color.rgb = COLOR_GOLD
        p.font.name = "Microsoft YaHei"
        p.alignment = PP_ALIGN.CENTER

    table_data = [
        ("第01幕·宿舍暗潮", "赵学衡", "简易课桌、清汤白水面、生大蒜、《中药学》书、分秒必争横幅", "昏暗清冷台灯、局部顶光", "沉重机械钟表滴答声、压迫感低音铺垫", "重点把控吃面狼吞虎咽与背书念词的强装自律感"),
        ("第02幕·门口对峙", "赵学衡、郝朋、程澄", "黑色公文夹、门框暗影道具", "走廊刺眼冷白强光斜切入室，强明暗反差", "门骤开撞击声、心跳急促骤停音效", "郝朋目光如炬具审讯威慑，赵学衡筷子定格半空"),
        ("第03幕·宿舍交锋", "赵学衡、郝朋、程澄", "100%全打钩日程手账、安神助眠胶囊处方、空药板", "桌面局部特写聚光，背景隐现心电红光", "心跳重锤加速，低沉大提琴旋律切入", "红叉账本与助眠药单无情并置，层层戳穿防线"),
        ("第04幕·暗室铁柜", "赵学衡、郝朋、程澄", "双门深绿铁皮储物柜、三把黄铜大锁、封条、通行证", "刺目冰蓝垂直追光，两侧深暗心理空间", "钥匙插入铜锁刺耳扭转声、沉重铁链落地声", "赵学衡以身死死贴挡柜门，尖锐破音抗拒"),
        ("第05幕·书籍雪崩", "赵学衡、郝朋、程澄", "成百上千全新考研讲义与空白试卷、未拆封讲义", "单束苍白顶光刺射，戏剧大冲突追光", "低沉压抑弦乐极速高潮、重锤心跳轰鸣", "高潮爆发：资料海啸倾倒，赵学衡跪地抱头痛哭"),
        ("第06幕·破防与接纳", "赵学衡、郝朋、程澄、林凡", "复古绿色搪瓷温茶杯、散落试卷手帕", "温暖琥珀色漫射光渐亮，破除室内阴霾", "弦乐由激越转入舒缓长音，呼吸声减弱", "甩手抗拒防线破除，林凡平视递茶，共情接纳"),
        ("第07幕·笛声晨光", "赵学衡、郝朋、程澄、林凡", "传统紫竹笛（配红穗）、严苛日程纸条碎片", "清晨明媚阳光大面积漫射，草药园生机绿意", "现场竹笛独奏《清心引》、空灵流水古琴", "徵音清心角音疏肝，碎纸成蝶飞扬，重获释然"),
        ("第08幕·谢幕归真", "全体演员（四人并肩）", "演出服、巨幅LED背景大屏（内经金句书法）", "剧场温暖金黄色泛光全开，辉煌落幕", "温暖宏大终曲交响、全场热烈掌声混音", "四人并肩牵手优雅鞠躬，黄帝内经经典句定格")
    ]

    for ri, row in enumerate(table_data):
        row_bg = RGBColor(18, 24, 34) if ri % 2 == 0 else RGBColor(14, 18, 26)
        for ci, val in enumerate(row):
            cell = tbl.cell(ri + 1, ci)
            cell.fill.solid()
            cell.fill.fore_color.rgb = row_bg
            cell.text_frame.margin_left = Inches(0.06)
            cell.text_frame.margin_right = Inches(0.06)
            cell.text_frame.margin_top = Inches(0.06)
            cell.text_frame.margin_bottom = Inches(0.06)
            p = cell.text_frame.paragraphs[0]
            p.text = val
            p.font.size = Pt(9.5)
            p.font.name = "Microsoft YaHei"
            if ci == 0:
                p.font.bold = True
                p.font.color.rgb = COLOR_GOLD
                p.alignment = PP_ALIGN.CENTER
            elif ci == 1:
                p.font.color.rgb = COLOR_WHITE
                p.alignment = PP_ALIGN.CENTER
            else:
                p.font.color.rgb = COLOR_SUBTITLE

    prs.save(output_pptx)
    print(f"\n生成成功！放映文件已保存至: {output_pptx}")
    print(f"总计 {len(prs.slides)} 页，母版比例 16:9。")

    if template_pptx:
        prs.save(template_pptx)
        print(f"同步更新模板文件: {template_pptx}")

    return output_pptx

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding='utf-8')
    build_presentation()
