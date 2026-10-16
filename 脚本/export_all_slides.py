# -*- coding: utf-8 -*-
import sys
import os
import glob
import win32com.client
import pythoncom

def export_all(source_pptx="心房搜查令_现场大屏放映版.pptx"):
    pythoncom.CoInitialize()
    ppt_app = win32com.client.DispatchEx("PowerPoint.Application")
    try:
        abs_pptx = os.path.abspath(source_pptx)
        pres = ppt_app.Presentations.Open(abs_pptx, WithWindow=False)
        out_dir = os.path.abspath("预览图")
        os.makedirs(out_dir, exist_ok=True)
        count = pres.Slides.Count
        print(f"Opening {source_pptx}, total slides: {count}")
        
        # 清理旧的超过当前页数的 slide_*.png（例如旧的 slide_11.png）
        for old_file in glob.glob(os.path.join(out_dir, "slide_*.png")):
            base = os.path.basename(old_file)
            try:
                num = int(base.replace("slide_", "").replace(".png", ""))
                if num > count:
                    os.remove(old_file)
                    print(f"Removed legacy preview file: {base}")
            except ValueError:
                pass
                
        for i in range(1, count + 1):
            out_path = os.path.join(out_dir, f"slide_{i:02d}.png")
            pres.Slides(i).Export(out_path, "PNG", 1920, 1080)
            print(f"Exported slide {i:02d} -> {out_path}")
        pres.Close()
        print("All slides exported successfully to 1080P PNG!")
    except Exception as e:
        print("Export error:", e)
        raise e
    finally:
        try:
            ppt_app.Quit()
        except Exception:
            pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    sys.stdout.reconfigure(encoding='utf-8')
    export_all()
