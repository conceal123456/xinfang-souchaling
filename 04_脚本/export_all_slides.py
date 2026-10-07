# -*- coding: utf-8 -*-
import sys
import os
import win32com.client
import pythoncom

def export_all():
    pythoncom.CoInitialize()
    ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        abs_pptx = os.path.abspath("心房搜查令_最终现场放映版.pptx")
        pres = ppt_app.Presentations.Open(abs_pptx, WithWindow=False)
        out_dir = os.path.abspath("preview_slides")
        os.makedirs(out_dir, exist_ok=True)
        count = pres.Slides.Count
        print(f"Total slides to export: {count}")
        for i in range(1, count + 1):
            out_path = os.path.join(out_dir, f"slide_{i:02d}.png")
            pres.Slides(i).Export(out_path, "PNG", 1920, 1080)
            print(f"Exported slide {i:02d} -> {out_path}")
        pres.Close()
        print("All slides exported successfully!")
    except Exception as e:
        print("Export error:", e)
    finally:
        os._exit(0)

if __name__ == "__main__":
    export_all()
