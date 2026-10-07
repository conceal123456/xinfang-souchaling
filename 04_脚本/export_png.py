# -*- coding: utf-8 -*-
import sys
import os
import win32com.client
import pythoncom

def export_slide(pptx_path, slide_idx, out_png):
    pythoncom.CoInitialize()
    ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        abs_pptx = os.path.abspath(pptx_path)
        pres = ppt_app.Presentations.Open(abs_pptx, WithWindow=False)
        abs_out = os.path.abspath(out_png)
        pres.Slides[slide_idx].Export(abs_out, "PNG", 1920, 1080)
        pres.Close()
        print(f"Exported slide {slide_idx} to {abs_out}")
    except Exception as e:
        print("Error:", e)
    finally:
        os._exit(0)

if __name__ == "__main__":
    pptx = sys.argv[1] if len(sys.argv) > 1 else "test_fullbleed.pptx"
    idx = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    out = sys.argv[3] if len(sys.argv) > 3 else "test_fullbleed.png"
    export_slide(pptx, idx, out)
