from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
import os, uuid

app = FastAPI()

class SlideContent(BaseModel):
    title: str
    bullets: list[str]

class PPTRequest(BaseModel):
    main_title: str
    subtitle: str
    slides: list[SlideContent]

# 配色方案 - 专业深蓝主题
PRIMARY_COLOR = RGBColor(0, 51, 102)      # 深蓝
ACCENT_COLOR = RGBColor(0, 153, 204)       # 亮蓝
TEXT_COLOR = RGBColor(51, 51, 51)          # 深灰
SUBTEXT_COLOR = RGBColor(102, 102, 102)    # 中灰
WHITE = RGBColor(255, 255, 255)

@app.post("/generate")
def generate_ppt(req: PPTRequest):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ========== 封面页 ==========
    slide = prs.slides.add_slide(blank)
    
    # 背景色块（顶部装饰条）
    top_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), 
        Inches(13.333), Inches(0.15)
    )
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = ACCENT_COLOR
    top_bar.line.fill.background()
    
    # 主标题背景
    title_box = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(2.0), 
        Inches(11.7), Inches(1.8)
    )
    title_box.fill.solid()
    title_box.fill.fore_color.rgb = PRIMARY_COLOR
    title_box.line.fill.background()
    
    # 主标题文字
    tb = slide.shapes.add_textbox(
        Inches(1.0), Inches(2.3), Inches(11.3), Inches(1.2)
    )
    p = tb.text_frame.paragraphs[0]
    p.text = req.main_title
    p.font.size = Pt(36)
    p.font.bold = True
    p.font.color.rgb = WHITE
    p.alignment = PP_ALIGN.CENTER
    p.font.name = "Microsoft YaHei"
    
    # 副标题
    tb2 = slide.shapes.add_textbox(
        Inches(1.0), Inches(4.2), Inches(11.3), Inches(0.8)
    )
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = req.subtitle
    p2.font.size = Pt(18)
    p2.font.color.rgb = SUBTEXT_COLOR
    p2.alignment = PP_ALIGN.CENTER
    p2.font.name = "Microsoft YaHei"
    
    # 底部装饰线
    bottom_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(5.5), Inches(5.5), 
        Inches(2.3), Inches(0.03)
    )
    bottom_line.fill.solid()
    bottom_line.fill.fore_color.rgb = ACCENT_COLOR
    bottom_line.line.fill.background()

    # ========== 内容页 ==========
    for idx, s in enumerate(req.slides, 1):
        slide = prs.slides.add_slide(blank)
        
        # 左侧装饰色条
        left_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), 
            Inches(0.12), Inches(7.5)
        )
        left_bar.fill.solid()
        left_bar.fill.fore_color.rgb = ACCENT_COLOR
        left_bar.line.fill.background()
        
        # 页面标题
        tbox = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.7)
        )
        tp = tbox.text_frame.paragraphs[0]
        tp.text = f"{idx}. {s.title}"
        tp.font.size = Pt(26)
        tp.font.bold = True
        tp.font.color.rgb = PRIMARY_COLOR
        tp.font.name = "Microsoft YaHei"
        
        # 标题下划线
        underline = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.15), 
            Inches(2.0), Inches(0.04)
        )
        underline.fill.solid()
        underline.fill.fore_color.rgb = ACCENT_COLOR
        underline.line.fill.background()
        
        # 内容区域
        bbox = slide.shapes.add_textbox(
            Inches(0.7), Inches(1.5), Inches(12), Inches(5.5)
        )
        tf = bbox.text_frame
        tf.word_wrap = True
        
        for i, b in enumerate(s.bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"• {b}"
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
            p.space_after = Pt(14)
            p.font.name = "Microsoft YaHei"
            p.level = 0
        
        # 页码
        page_num = slide.shapes.add_textbox(
            Inches(12.5), Inches(7.0), Inches(0.6), Inches(0.3)
        )
        pn = page_num.text_frame.paragraphs[0]
        pn.text = str(idx)
        pn.font.size = Pt(12)
        pn.font.color.rgb = SUBTEXT_COLOR
        pn.alignment = PP_ALIGN.RIGHT

    # 保存文件
    os.makedirs("/app/output", exist_ok=True)
    fname = f"{uuid.uuid4().hex[:8]}.pptx"
    fpath = f"/app/output/{fname}"
    prs.save(fpath)
    
    base = os.environ.get("RENDER_EXTERNAL_URL", "")
    if not base:
        base = "https://" + os.environ.get("RENDER_EXTERNAL_HOSTNAME", "")
    return {"download_url": f"{base}/download/{fname}", "filename": fname}

@app.get("/download/{filename}")
def download(filename: str):
    fpath = f"/app/output/{filename}"
    if os.path.exists(fpath):
        return FileResponse(fpath, filename=filename)
    return {"error": "file not found"}
