from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
import os, uuid

app = FastAPI()

class SlideContent(BaseModel):
    title: str
    bullets: list[str]

class PPTRequest(BaseModel):
    main_title: str
    subtitle: str
    slides: list[SlideContent]

@app.post("/generate")
def generate_ppt(req: PPTRequest):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # 封面页
    slide = prs.slides.add_slide(blank)
    tb = slide.shapes.add_textbox(Inches(0.5), Inches(2.2), Inches(12.3), Inches(1.5))
    p = tb.text_frame.paragraphs[0]
    p.text = req.main_title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = RGBColor(0, 51, 102)
    p.alignment = PP_ALIGN.CENTER

    tb2 = slide.shapes.add_textbox(Inches(0.5), Inches(3.8), Inches(12.3), Inches(1))
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = req.subtitle
    p2.font.size = Pt(18)
    p2.font.color.rgb = RGBColor(100, 100, 100)
    p2.alignment = PP_ALIGN.CENTER

    # 内容页
    for idx, s in enumerate(req.slides, 1):
        slide = prs.slides.add_slide(blank)
        tbox = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.8))
        tp = tbox.text_frame.paragraphs[0]
        tp.text = f"{idx}. {s.title}"
        tp.font.size = Pt(28)
        tp.font.bold = True
        tp.font.color.rgb = RGBColor(0, 51, 102)

        bbox = slide.shapes.add_textbox(Inches(0.7), Inches(1.4), Inches(12), Inches(5.5))
        tf = bbox.text_frame
        tf.word_wrap = True
        for i, b in enumerate(s.bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"• {b}"
            p.font.size = Pt(18)
            p.space_after = Pt(10)

    os.makedirs("/app/output", exist_ok=True)
    fname = f"{uuid.uuid4().hex[:8]}.pptx"
    fpath = f"/app/output/{fname}"
    prs.save(fpath)
    
    # 生成下载链接
    base = os.environ.get("RENDER_EXTERNAL_URL", "")
    if not base:
        base = "https://" + os.environ.get("RENDER_EXTERNAL_HOSTNAME", "")
    return {"download_url": f"{base}/download/{fname}", "filename": fname}

@app.get("/download/{filename}")
def download(filename: str):
    fpath = f"/app/output/{filename}"
    if os.path.exists(fpath):
        return FileResponse(fpath, filename=filename, media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation")
    return {"error": "file not found"}
