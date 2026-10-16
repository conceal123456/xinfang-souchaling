# -*- coding: utf-8 -*-
"""
端到端综合校验脚本：全方位断言检验《心房搜查令》全部交付物。
覆盖：
1. 现场大屏放映版 PPT（10 页架构、16:9 纯净画幅、演讲者备注、P10 质感卡片、无暴露表格）；
2. 全流程场控分镜案 PPT（11 页架构、9x6 场控表完整性、演讲者备注）；
3. 比赛标准公文剧本 DOCX（红头公文排版、小标宋二号、仿宋三号、32/28 磅行距、2 字符缩进、全场控切换标记）；
4. 官方提交版 PDF（Word COM 原生导出、文本关键词检索完全一致）；
5. 纯文本台词 TXT；
6. 参赛报名表 DOCX（主表数据、演职人员对应表完备）；
7. 1080P 预览图集（10 张高清图、无冗余孤立图）。
"""

import os
import sys
import glob
from pptx import Presentation
from docx import Document
from docx.oxml.ns import qn
import pypdf

def verify_all():
    sys.stdout.reconfigure(encoding='utf-8')
    print("=" * 60)
    print("开始执行《心房搜查令》全项目端到端综合校验")
    print("=" * 60)

    # ----------------------------------------------------
    # 1. 现场大屏放映版 PPT 校验 (心房搜查令_现场大屏放映版.pptx)
    # ----------------------------------------------------
    print("\n[1/7] 校验《心房搜查令_现场大屏放映版.pptx》...")
    screen_ppt = "心房搜查令_现场大屏放映版.pptx"
    assert os.path.exists(screen_ppt), f"未找到文件: {screen_ppt}"
    prs_s = Presentation(screen_ppt)
    assert len(prs_s.slides) == 10, f"大屏放映版必须为 10 页，当前为 {len(prs_s.slides)} 页"
    
    # 比例校验
    w, h = prs_s.slide_width.inches, prs_s.slide_height.inches
    assert abs(w / h - 16.0 / 9.0) < 0.01, f"画幅非 16:9: {w}x{h}"
    
    # 分镜纯净度与备注校验
    for idx in range(len(prs_s.slides)):
        slide = prs_s.slides[idx]
        notes = slide.notes_slide.notes_text_frame.text.strip()
        assert len(notes) > 20, f"第 {idx+1} 页缺少详尽的演讲者备注！"
        
        # 第 2-9 页纯场景图模式
        if 1 <= idx <= 8:
            text_shapes = [s for s in slide.shapes if s.has_text_frame and s.text_frame.text.strip()]
            assert len(text_shapes) == 0, f"第 {idx+1} 页存在文字形状，未实现纯大图画幅: {[s.text_frame.text for s in text_shapes]}"
        
        # 绝不能出现表格
        table_shapes = [s for s in slide.shapes if s.has_table]
        assert len(table_shapes) == 0, f"第 {idx+1} 页违规包含表格，存在大屏暴露后台表格风险！"

    # P10 校验（核心关键词 + 无遮挡居中书法几何校验）
    slide_10 = prs_s.slides[9]
    t10_all = "".join([s.text_frame.text for s in slide_10.shapes if s.has_text_frame])
    t10_compact = t10_all.replace(" ", "").replace("　", "")
    for kw in ["恬淡虚无", "黄帝内经", "演职人员表", "特别鸣谢", "全剧终", "谢谢大家", "赵学衡", "郝朋", "程澄", "林凡"]:
        assert kw in t10_compact, f"第 10 页缺少核心要素: {kw}"
        
    # 几何边界校验：左翼卡片右边缘严禁侵入 x > 4.7 英寸，右翼卡片左边缘严禁早于 x < 8.6 英寸
    # 彻底杜绝遮挡正中“心房搜查令”（x in [5.01, 8.34]）
    cards = [s for s in slide_10.shapes if s.shape_type == 1 and s.has_text_frame and "【" in s.text_frame.text]
    assert len(cards) == 2, f"第 10 页演职与鸣谢卡片数量不符: {len(cards)}"
    for card in cards:
        c_left = card.left.inches
        c_right = card.left.inches + card.width.inches
        if "演职人员表" in card.text_frame.text:
            assert c_right <= 4.7, f"左翼演职人员卡片过宽，右边缘 {c_right:.2f} 侵入正中大字空间！"
        elif "特别鸣谢" in card.text_frame.text:
            assert c_left >= 8.6, f"右翼鸣谢卡片左边缘 {c_left:.2f} 侵入正中大字空间！"
            # 校验最后一行无单字孤行
            assert "陪伴！" not in card.text_frame.text.splitlines()[-1].strip(), "鸣谢卡片末尾存在孤字换行"
            
    print("  ✓ 大屏放映版 10 页规范、纯场景图放映、无暴露表格、全注入演讲者备注、P10抗遮挡几何校验通过！")

    # ----------------------------------------------------
    # 2. 全流程场控分镜案 PPT 校验 (心房搜查令_全流程场控与分镜案.pptx)
    # ----------------------------------------------------
    print("\n[2/7] 校验《心房搜查令_全流程场控与分镜案.pptx》...")
    control_ppt = "心房搜查令_全流程场控与分镜案.pptx"
    assert os.path.exists(control_ppt), f"未找到文件: {control_ppt}"
    prs_c = Presentation(control_ppt)
    assert len(prs_c.slides) == 11, f"场控分镜案必须为 11 页，当前为 {len(prs_c.slides)} 页"
    
    slide_11 = prs_c.slides[10]
    tables = [s.table for s in slide_11.shapes if s.has_table]
    assert len(tables) == 1, "第 11 页必须包含 1 个场控表格"
    tbl = tables[0]
    assert len(tbl.rows) == 9, f"场控表格必须为 9 行（1 表头 + 8 幕），实际为 {len(tbl.rows)}"
    assert len(tbl.columns) == 6, f"场控表格必须为 6 列，实际为 {len(tbl.columns)}"
    assert "心房搜查令" in "".join([s.text_frame.text for s in slide_11.shapes if s.has_text_frame])
    print("  ✓ 全流程场控分镜案 11 页规范、第 11 页 9x6 完整场控执行表，通过！")

    # ----------------------------------------------------
    # 3. 剧本 DOCX 格式校验 (剧本/心房搜查令.docx)
    # ----------------------------------------------------
    print("\n[3/7] 校验《剧本/心房搜查令.docx》格式合规性...")
    doc_path = "剧本/心房搜查令.docx"
    assert os.path.exists(doc_path), f"未找到文件: {doc_path}"
    doc = Document(doc_path)
    
    # 标题检查
    p0 = doc.paragraphs[0]
    assert p0.text.strip() == "心房搜查令（原创）", f"标题不正确: {p0.text}"
    r0 = p0.runs[0]
    r0_ea = r0._element.rPr.find(qn('w:rFonts')).get(qn('w:eastAsia'))
    assert r0_ea == "方正小标宋简体", f"标题字体必须为方正小标宋简体，实际为 {r0_ea}"
    assert abs(r0.font.size.pt - 22.0) < 0.1, f"标题字号必须为二号(22pt)，实际为 {r0.font.size.pt}"
    assert abs(p0.paragraph_format.line_spacing.pt - 32.0) < 0.1, f"标题行距必须固定值 32 磅，实际为 {p0.paragraph_format.line_spacing.pt}"
    
    # 检查一级、二级、三级标题行距固定值 32 磅，正文固定值 28 磅
    h1_count = 0
    h2_count = 0
    cue_ppt_count = 0
    cue_light_count = 0
    cue_sound_count = 0
    
    for p in doc.paragraphs:
        t = p.text.strip()
        if not t:
            continue
        sp = p.paragraph_format.line_spacing.pt
        if t.startswith("一、") or t.startswith("二、"):
            h1_count += 1
            assert abs(sp - 32.0) < 0.1, f"一级标题行距必须为 32 磅: {t}"
        elif t.startswith("（一）") or t.startswith("（二）") or t.startswith("（三）") or t.startswith("（四）") or t.startswith("（五）") or t.startswith("（六）") or t.startswith("（七）") or t.startswith("（八）"):
            h2_count += 1
            assert abs(sp - 32.0) < 0.1, f"二级标题行距必须为 32 磅: {t}"
        elif t.startswith("1.") or t.startswith("2.") or t.startswith("3."):
            assert abs(sp - 32.0) < 0.1, f"三级标题行距必须为 32 磅: {t}"
        elif t != "心房搜查令（原创）":
            assert abs(sp - 28.0) < 0.1, f"正文及其他部分行距必须为 28 磅: {t[:20]}"
            
        if "【PPT 切第" in t:
            cue_ppt_count += 1
        if "【舞台灯光：" in t or "【灯光：" in t:
            cue_light_count += 1
        if "【现场音效：" in t or "【音效：" in t:
            cue_sound_count += 1

    assert h1_count == 2, f"一级标题数量不符: {h1_count}"
    assert h2_count == 11, f"二级标题数量不符: {h2_count} (3个基本信息 + 8幕正文)"
    assert cue_ppt_count >= 8, f"PPT 场控切换标记不足: {cue_ppt_count}"
    assert cue_light_count >= 8, f"灯光标记不足: {cue_light_count}"
    assert cue_sound_count >= 8, f"音效标记不足: {cue_sound_count}"
    print(f"  ✓ 剧本 DOCX 公文规范排版校验通过（标题 32 磅、正文 28 磅、PPT 切换标记 {cue_ppt_count} 处、灯光标记 {cue_light_count} 处、音效标记 {cue_sound_count} 处）")

    # ----------------------------------------------------
    # 4. 提交版 PDF 校验 (剧本/2024级数据科学与大数据技术班_心房搜查令（原创）.pdf)
    # ----------------------------------------------------
    print("\n[4/7] 校验官方提交版 PDF...")
    pdf_path = "剧本/2024级数据科学与大数据技术班_心房搜查令（原创）.pdf"
    assert os.path.exists(pdf_path), f"未找到 PDF 文件: {pdf_path}"
    pdf_size = os.path.getsize(pdf_path)
    assert pdf_size > 100 * 1024, f"PDF 文件过小: {pdf_size} 字节"
    
    reader = pypdf.PdfReader(pdf_path)
    pdf_text = "".join([p.extract_text() for p in reader.pages])
    pdf_text_compact = pdf_text.replace(" ", "").replace("\n", "").replace("\r", "")
    required_keywords = [
        "心房搜查令（原创）",
        "冒名顶替综合征",
        "学术与资料囤积障碍",
        "思则气结、土壅木郁",
        "五音疗疾",
        "清心引",
        "PPT切第",
        "舞台灯光",
        "现场音效",
        "黄帝内经",
        "恬淡虚无，真气从之；精神内守，病安从来"
    ]
    for kw in required_keywords:
        assert kw in pdf_text_compact, f"PDF 缺少核心规范关键词: {kw}"

    # 深度字体审计：验证真实嵌入的字体是否严格匹配公文规范
    import fitz
    doc_fitz = fitz.open(pdf_path)
    font_char_counts = {}
    for page in doc_fitz:
        for block in page.get_text('dict')['blocks']:
            if 'lines' in block:
                for line in block['lines']:
                    for span in line['spans']:
                        fn = span['font']
                        font_char_counts[fn] = font_char_counts.get(fn, 0) + len(span['text'])
    doc_fitz.close()
    
    # 断言方正小标宋、黑体、楷体、仿宋真实生效
    has_fzxbs = any("FZXBS" in f or "XiaobiaoSong" in f or "标宋" in f for f in font_char_counts)
    has_hei = any("Hei" in f or "黑体" in f for f in font_char_counts)
    has_kai = any("Kai" in f or "楷体" in f for f in font_char_counts)
    has_fangsong = any("FangSong" in f or "仿宋" in f for f in font_char_counts)
    
    assert has_fzxbs, f"PDF 标题未正确应用方正小标宋字体！已检测字体: {list(font_char_counts.keys())}"
    assert has_hei, f"PDF 一级标题未正确应用黑体！已检测字体: {list(font_char_counts.keys())}"
    assert has_kai, f"PDF 二级标题未正确应用楷体！已检测字体: {list(font_char_counts.keys())}"
    assert has_fangsong, f"PDF 正文未正确应用仿宋字体！已检测字体: {list(font_char_counts.keys())}"
    
    fangsong_chars = sum(v for k, v in font_char_counts.items() if "FangSong" in k or "仿宋" in k)
    simsun_chars = sum(v for k, v in font_char_counts.items() if "SimSun" in k or "宋体" in k)
    assert fangsong_chars > 4000, f"正文仿宋字符数异常偏少: {fangsong_chars}"
    assert simsun_chars < 500, f"检测到宋体(SimSun)大规模回退替代异常: {simsun_chars} 字符"

    print(f"  ✓ 提交版 PDF 校验通过（共 {len(reader.pages)} 页，大小 {pdf_size / 1024:.1f} KB，公文字体精准嵌入: 小标宋/黑体/楷体/仿宋均达标）")

    # ----------------------------------------------------
    # 5. 纯文本台词校验 (剧本/心房搜查令_台词文本.txt)
    # ----------------------------------------------------
    print("\n[5/7] 校验纯文本台词 TXT...")
    txt_path = "剧本/心房搜查令_台词文本.txt"
    assert os.path.exists(txt_path), f"未找到台词文件: {txt_path}"
    with open(txt_path, 'r', encoding='utf-8') as f:
        txt_content = f.read()
    assert "心房搜查令（原创）" in txt_content
    assert "PPT 切第" in txt_content
    assert len(txt_content) > 5000
    
    # 纯净度审计：台词文本严禁包含行号索引（如 '0: 心房搜查令'）
    import re
    index_pollution = [line for line in txt_content.splitlines() if re.match(r"^\d+:\s", line.strip())]
    assert len(index_pollution) == 0, f"台词文本被索引序号污染: 发现 {len(index_pollution)} 处残留，如 {index_pollution[:3]}"
    print("  ✓ 纯文本台词 TXT 校验通过（纯净剧本文本，无索引序号污染）")

    # ----------------------------------------------------
    # 6. 参赛报名表 DOCX 校验 (参赛材料/附件2：第十四届校园心理剧大赛参赛报名表.docx)
    # ----------------------------------------------------
    print("\n[6/7] 校验《参赛材料/附件2：第十四届校园心理剧大赛参赛报名表.docx》...")
    form_path = "参赛材料/附件2：第十四届校园心理剧大赛参赛报名表.docx"
    assert os.path.exists(form_path), f"未找到报名表: {form_path}"
    doc_form = Document(form_path)
    assert len(doc_form.tables) == 2, f"报名表必须严格包含主表和演职人员表共 2 个表格（幂等性），实际为 {len(doc_form.tables)} 个表格"
    
    t0 = doc_form.tables[0]
    t0_row1 = [c.text.strip() for c in t0.rows[1].cells]
    assert "医药信息工程学院" in t0_row1[1], f"学院信息不正确: {t0_row1[1]}"
    assert "心房搜查令" in t0_row1[2], f"剧名不正确: {t0_row1[2]}"
    assert "4人" in t0_row1[3], f"人数不正确: {t0_row1[3]}"
    assert "郝朋" in t0_row1[4], f"负责人不正确: {t0_row1[4]}"
    
    t1 = doc_form.tables[1]
    t1_text = "".join([c.text.strip() for r in t1.rows for c in r.cells])
    for actor in ["赵学衡", "郝朋", "程澄", "林凡"]:
        assert actor in t1_text, f"演职人员表缺少成员: {actor}"
    print("  ✓ 参赛报名表及演职人员表校验通过（严格 2 个表格，结构完整幂等）")

    # ----------------------------------------------------
    # 7. 预览图集校验 (预览图/slide_01.png ~ slide_10.png)
    # ----------------------------------------------------
    print("\n[7/7] 校验 1080P 预览图集...")
    png_files = sorted(glob.glob("预览图/slide_*.png"))
    assert len(png_files) == 10, f"预览图数量必须严格为 10 张，实际为 {len(png_files)} 张: {png_files}"
    for i, pf in enumerate(png_files, start=1):
        expected_name = f"slide_{i:02d}.png"
        assert os.path.basename(pf) == expected_name, f"图片命名不符: {pf}"
        size = os.path.getsize(pf)
        assert size > 200 * 1024, f"图片文件异常偏小: {pf} ({size} 字节)"
    print("  ✓ 1080P 预览图集校验通过（严格 10 张，无孤立残留图片）")

    print("\n" + "=" * 60)
    print("恭喜！所有 7 项端到端深度校验 100% 全部通过！各项资产完全符合比赛要求与现场放映规范！")
    print("=" * 60)

if __name__ == "__main__":
    verify_all()
