# -*- coding: utf-8 -*-
"""
检查最终放映版 PPT 做得对不对。

看这四样：
1. 文件在不在、多大、总共几页；
2. 第 2-9 页八幕：占位框清没清、剧照有没有铺满整页、
   层级对不对（背景 -> 剧照 -> 标题 -> 字幕）、文字有没有丢；
3. 第 10 页致谢页的标题、金句、演职人员表、鸣谢；
4. 第 11 页场控表，9 行 6 列，字段齐不齐。

对不上的地方会直接 assert 报错。
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches

def run_verification():
    sys.stdout.reconfigure(encoding='utf-8')
    output_file = '心房搜查令_最终现场放映版.pptx'
    
    print("==================================================")
    print("1. 文件和基本参数")
    print("==================================================")
    if not os.path.exists(output_file):
        raise FileNotFoundError(f"未找到目标放映文件: {output_file}")
    
    size_bytes = os.path.getsize(output_file)
    print(f"找到文件：{output_file}")
    print(f"✓ 文件大小: {size_bytes:,} 字节 ({size_bytes / (1024 * 1024):.2f} MB)")
    assert size_bytes > 1024 * 1024, f"文件过小 ({size_bytes} 字节)，剧照可能未成功嵌入"
    
    prs = Presentation(output_file)
    slide_count = len(prs.slides)
    print(f"页数：{slide_count}（要求 11 页）")
    assert slide_count == 11, f"页面总数错误，预期 11 页，实际 {slide_count} 页"

    # 校验幻灯片母版尺寸（16:9 宽屏）
    width_in = prs.slide_width.inches
    height_in = prs.slide_height.inches
    aspect_ratio = width_in / height_in
    print(f"✓ 母版画幅尺寸: {width_in:.3f}\" x {height_in:.3f}\" (宽高比: {aspect_ratio:.3f})")
    assert abs(aspect_ratio - (16.0 / 9.0)) < 0.01, f"画幅比例非 16:9: {aspect_ratio}"

    print("\n==================================================")
    print("2. 八幕剧照页")
    print("==================================================")
    for idx in range(1, 9):
        slide = prs.slides[idx]
        spTree = slide.shapes._spTree
        tags = [child.tag.split('}')[-1] for child in spTree]
        
        print(f"\n--- 幻灯片第 {idx + 1} 页 (分镜 {idx:02d}) ---")
        print(f"  形状数量: {len(slide.shapes)}")
        print(f"  XML 节点渲染树: {tags}")
        
        # 1. 占位框残留检测：必须没有任何带有“预留槽位”的形状
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text
                assert "预留槽位" not in txt, f"第 {idx + 1} 页检测到残留占位框文本: {txt}"
                assert "scene_" not in txt or "SCENE" in txt, f"第 {idx + 1} 页检测到残留占位框文件名: {txt}"
        print("  ✓ 占位框：没有残留")

        # 2. 检查图片存在与尺寸比例
        pic_shape = None
        for shape in slide.shapes:
            if shape.shape_type == 13: # Picture
                pic_shape = shape
                break
        
        assert pic_shape is not None, f"第 {idx + 1} 页缺少剧照图片"
        print(f"  ✓ 剧照位置与尺寸: Left={pic_shape.left.inches:.2f}\", Top={pic_shape.top.inches:.2f}\", Width={pic_shape.width.inches:.2f}\", Height={pic_shape.height.inches:.2f}\"")
        assert abs(pic_shape.width.inches - 13.333) < 0.1, f"剧照宽度不符合 16:9 全屏规范: {pic_shape.width.inches}"
        assert abs(pic_shape.height.inches - 7.5) < 0.1, f"剧照高度不符合 16:9 全屏规范: {pic_shape.height.inches}"
        
        # 3. 检查渲染层级顺序 (spTree child 2 是 bg，child 3 必须是 pic)
        assert tags[2] == 'sp', f"第 {idx + 1} 页底层必须为全屏背景底板 (当前为 {tags[2]})"
        assert tags[3] == 'pic', f"第 {idx + 1} 页图片必须置于索引 3 (在背景之上、悬浮层之下，当前为 {tags[3]})"
        assert len(tags) >= 6, f"第 {idx + 1} 页缺少顶部或底部悬浮条"
        print("  ✓ 层级：背景 -> 剧照 -> 标题栏 -> 字幕条")

        # 4. 提取标题与对白验证（按几何坐标精准提取顶部条与底部字幕条）
        title_txt = ""
        dialogue_txt = ""
        stage_cue_txt = ""
        for shape in slide.shapes:
            if shape.has_text_frame:
                txt = shape.text_frame.text.strip()
                if txt:
                    if shape.top.inches < 2.0 and not shape.shape_type == 13:
                        lines = txt.splitlines()
                        if len(lines) > 0:
                            title_txt = lines[0]
                    elif shape.top.inches > 4.0:
                        lines = txt.splitlines()
                        if len(lines) > 0:
                            dialogue_txt = lines[0]
                        for line in lines:
                            if "舞台提示" in line:
                                stage_cue_txt = line

        assert len(title_txt) > 0, f"第 {idx + 1} 页缺少标题文本"
        assert len(dialogue_txt) > 0, f"第 {idx + 1} 页缺少台词文本"
        assert len(stage_cue_txt) > 0, f"第 {idx + 1} 页缺少舞台提示文本"
        print(f"  ✓ 顶部徽标标题: {title_txt}")
        print(f"  ✓ 底部悬浮台词: {dialogue_txt[:40]}...")
        print(f"  ✓ 舞台调度提示: {stage_cue_txt[:40]}...")

    print("\n==================================================")
    print("3. 第 10 页致谢页")
    print("==================================================")
    slide_10 = prs.slides[9]
    all_texts_10 = []
    for shape in slide_10.shapes:
        if shape.has_text_frame:
            all_texts_10.append(shape.text_frame.text.strip())
    
    combined_10 = "\n".join(all_texts_10)
    print(f"✓ 第 10 页提取文本概览:\n{combined_10[:300]}...\n")

    compact_10 = combined_10.replace(" ", "").replace("　", "")
    assert "谢谢大家" in compact_10, "第 10 页缺少核心致谢标题“谢谢大家”"
    assert "全剧终" in compact_10, "第 10 页缺少“全剧终”标识"
    assert "黄帝内经" in compact_10 or "恬淡虚无" in compact_10, "第 10 页缺少《黄帝内经》经典金句"
    assert "演职人员表" in compact_10, "第 10 页缺少演职人员表"
    assert "特别鸣谢" in compact_10, "第 10 页缺少特别鸣谢"
    assert "2024级数据科学与大数据技术班" in compact_10, "第 10 页缺少演职班级"
    for actor in ["赵学衡", "郝朋", "程澄", "林凡"]:
        assert actor in compact_10, f"第 10 页演职人员表缺少角色: {actor}"
    print("✓ 致谢页要素齐全")

    print("\n==================================================")
    print("4. 第 11 页场控表")
    print("==================================================")
    slide_11 = prs.slides[10]
    table_shape = None
    for shape in slide_11.shapes:
        if shape.has_table:
            table_shape = shape
            break
            
    assert table_shape is not None, "第 11 页未找到场控表格！"
    table = table_shape.table
    rows = len(table.rows)
    cols = len(table.columns)
    print(f"✓ 场控表尺寸: {rows} 行 x {cols} 列")
    assert rows == 9, f"场控表行数不匹配，预期 9 行，实际 {rows} 行"
    assert cols == 6, f"场控表列数不匹配，预期 6 列，实际 {cols} 列"
    
    headers = [cell.text.strip() for cell in table.rows[0].cells]
    print(f"✓ 表头字段: {' | '.join(headers)}")
    expected_headers = ["幕次 / 场景", "登场角色", "核心道具调度", "舞台灯光设计", "现场音效与配乐", "场控与调度要点"]
    for eh in expected_headers:
        assert any(eh in h for h in headers), f"表头缺少字段: {eh}"
        
    for r_i in range(1, 9):
        sc_name = table.cell(r_i, 0).text.strip()
        chars = table.cell(r_i, 1).text.strip()
        props = table.cell(r_i, 2).text.strip()
        lights = table.cell(r_i, 3).text.strip()
        music = table.cell(r_i, 4).text.strip()
        ops = table.cell(r_i, 5).text.strip()
        print(f"  ✓ [行 {r_i}] {sc_name} | 角色:{chars} | 道具:{props[:10]}... | 灯光:{lights[:10]}... | 音效:{music[:10]}...")
        assert len(props) > 0 and len(lights) > 0 and len(music) > 0 and len(ops) > 0
    print("✓ 场控表字段齐全")

    print("\n==================================================")
    print("四项都过了，没发现问题。")
    print("==================================================")

if __name__ == "__main__":
    run_verification()
