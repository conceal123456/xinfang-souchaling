# -*- coding: utf-8 -*-
"""
全面校验最新现场放映版 PPT (11 页)。

校验项：
1. 文件完整性与基本参数（大小、11 页、16:9 画幅规范）；
2. 封面（第 1 页）：包含 16:9 全屏海报底图及核心标题信息；
3. 八大分镜放映页（第 2-9 页）：纯场景大图模式，每页包含 16:9 全屏高清剧照，无任何上下遮挡杂乱文本；
4. 致谢页（第 10 页）：包含 16:9 全屏剧院谢幕底图，演职人员表、鸣谢单位、黄帝内经金句等关键要素完备；
5. 后台场控表（第 11 页）：9 行 6 列完整表格，调度要素齐备。
"""

import os
import sys
from pptx import Presentation
from pptx.util import Inches

def run_verification():
    sys.stdout.reconfigure(encoding='utf-8')
    output_file = '心房搜查令_最终现场放映版.pptx'
    
    print("==================================================")
    print("1. 文件和基本参数校验")
    print("==================================================")
    if not os.path.exists(output_file):
        raise FileNotFoundError(f"未找到目标放映文件: {output_file}")
    
    size_bytes = os.path.getsize(output_file)
    print(f"找到放映文件：{output_file}")
    print(f"✓ 文件大小: {size_bytes:,} 字节 ({size_bytes / (1024 * 1024):.2f} MB)")
    assert size_bytes > 2 * 1024 * 1024, f"文件过小 ({size_bytes} 字节)，图片可能未成功嵌入"
    
    prs = Presentation(output_file)
    slide_count = len(prs.slides)
    print(f"✓ 总页数：{slide_count} 页（要求 11 页）")
    assert slide_count == 11, f"页面总数错误，预期 11 页，实际 {slide_count} 页"

    # 校验幻灯片母版尺寸（16:9 宽屏）
    width_in = prs.slide_width.inches
    height_in = prs.slide_height.inches
    aspect_ratio = width_in / height_in
    print(f"✓ 母版画幅尺寸: {width_in:.3f}\" x {height_in:.3f}\" (宽高比: {aspect_ratio:.3f})")
    assert abs(aspect_ratio - (16.0 / 9.0)) < 0.01, f"画幅比例非 16:9: {aspect_ratio}"

    print("\n==================================================")
    print("2. 封面（第 1 页）与分镜放映页（第 2-9 页）")
    print("==================================================")
    # 封面
    slide_cov = prs.slides[0]
    cov_has_pic = any(shape.shape_type == 13 for shape in slide_cov.shapes)
    assert cov_has_pic, "第 1 页封面缺少 16:9 海报大图"
    cov_texts = [shape.text_frame.text for shape in slide_cov.shapes if shape.has_text_frame]
    cov_combined = "".join(cov_texts)
    assert "心房搜查令" in cov_combined, "第 1 页封面缺少标题“心房搜查令”"
    print("✓ 第 1 页封面：已嵌入 16:9 海报大图并包含主标题")

    # 八大分镜
    for idx in range(1, 9):
        slide = prs.slides[idx]
        pic_shapes = [shape for shape in slide.shapes if shape.shape_type == 13]
        text_shapes = [shape for shape in slide.shapes if shape.has_text_frame and shape.text_frame.text.strip()]
        
        print(f"\n--- 幻灯片第 {idx + 1:02d} 页 (分镜 {idx:02d}) ---")
        assert len(pic_shapes) >= 1, f"第 {idx + 1} 页缺少 16:9 场景剧照"
        pic = pic_shapes[0]
        print(f"  ✓ 场景大图尺寸: Left={pic.left.inches:.2f}\", Top={pic.top.inches:.2f}\", Width={pic.width.inches:.2f}\", Height={pic.height.inches:.2f}\"")
        assert abs(pic.width.inches - 13.333) < 0.1, f"图片宽度非 16:9 全屏规范: {pic.width.inches}"
        assert abs(pic.height.inches - 7.5) < 0.1, f"图片高度非 16:9 全屏规范: {pic.height.inches}"
        
        # 纯场景图模式：不得有遮挡文字
        assert len(text_shapes) == 0, f"第 {idx + 1} 页存在残留文字形状: {[s.text_frame.text for s in text_shapes]}，未能保持纯场景图"
        print("  ✓ 纯场景大图模式：无上下文字遮挡，画面纯净")

    print("\n==================================================")
    print("3. 第 10 页致谢页")
    print("==================================================")
    slide_10 = prs.slides[9]
    thx_has_pic = any(shape.shape_type == 13 for shape in slide_10.shapes)
    assert thx_has_pic, "第 10 页致谢页缺少 16:9 剧院美图底图"
    
    all_texts_10 = []
    for shape in slide_10.shapes:
        if shape.has_text_frame:
            all_texts_10.append(shape.text_frame.text.strip())
    
    combined_10 = "\n".join(all_texts_10)
    compact_10 = combined_10.replace(" ", "").replace("　", "")
    assert "谢谢大家" in compact_10, "第 10 页缺少核心致谢标题“谢谢大家”"
    assert "全剧终" in compact_10, "第 10 页缺少“全剧终”标识"
    assert "黄帝内经" in compact_10 or "恬淡虚无" in compact_10, "第 10 页缺少《黄帝内经》经典金句"
    assert "演职人员表" in compact_10, "第 10 页缺少演职人员表"
    assert "特别鸣谢" in compact_10, "第 10 页缺少特别鸣谢"
    assert "2024级数据科学与大数据技术班" in compact_10, "第 10 页缺少演职班级"
    for actor in ["赵学衡", "郝朋", "程澄", "林凡"]:
        assert actor in compact_10, f"第 10 页演职人员表缺少角色: {actor}"
    print("✓ 致谢页：已嵌入 16:9 剧院美图底图，演职人员及致谢金句齐全")

    print("\n==================================================")
    print("4. 第 11 页后台场控表")
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
        assert len(props) > 0 and len(lights) > 0 and len(music) > 0 and len(ops) > 0
    print("✓ 场控表：8 幕全流程调度数据完备无缺")

    print("\n==================================================")
    print("全部校验通过！无任何异常。")
    print("==================================================")

if __name__ == "__main__":
    run_verification()
