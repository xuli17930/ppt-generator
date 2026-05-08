from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import List, Dict, Optional
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
import os, uuid, requests
from io import BytesIO

app = FastAPI()

class SlideContent(BaseModel):
    title: str
    bullets: List[str]

class PPTRequest(BaseModel):
    main_title: str
    subtitle: str
    slides: List[SlideContent]
    cover_image: str = ""
    data_images: List[str] = []
    data_points: List[Dict] = []

# 配色方案
DARK_BLUE = RGBColor(0, 32, 96)
BRIGHT_BLUE = RGBColor(0, 112, 192)
LIGHT_BLUE = RGBColor(230, 240, 255)
TEXT_DARK = RGBColor(33, 33, 33)
TEXT_GRAY = RGBColor(100, 100, 100)
WHITE = RGBColor(255, 255, 255)

def download_image(url):
    try:
        response = requests.get(url, timeout=15, headers={'User-Agent': 'Mozilla/5.0'})
        if response.status_code == 200:
            return BytesIO(response.content)
    except Exception as e:
        print(f"图片下载失败: {e}")
    return None

@app.post("/generate")
def generate_ppt(req: PPTRequest):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ========== 封面页 ==========
    slide = prs.slides.add_slide(blank)
    
    if req.cover_image:
        img_stream = download_image(req.cover_image)
        if img_stream:
            slide.shapes.add_picture(img_stream, Inches(5.3), Inches(0), height=Inches(7.5))
            mask = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(6.0), Inches(7.5))
            mask.fill.solid()
            mask.fill.fore_color.rgb = LIGHT_BLUE
            mask.line.fill.background()
    
    top_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.25))
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = DARK_BLUE
    top_bar.line.fill.background()
    
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(2.2), Inches(5.0), Inches(1.5))
    p = tb.text_frame.paragraphs[0]
    p.text = req.main_title
    p.font.size = Pt(32)
    p.font.bold = True
    p.font.color.rgb = DARK_BLUE
    p.font.name = "Microsoft YaHei"
    
    tb2 = slide.shapes.add_textbox(Inches(0.8), Inches(3.8), Inches(5.0), Inches(1.0))
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = req.subtitle
    p2.font.size = Pt(16)
    p2.font.color.rgb = TEXT_GRAY
    p2.font.name = "Microsoft YaHei"
    
    bottom_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(5.0), Inches(2.5), Inches(0.03))
    bottom_line.fill.solid()
    bottom_line.fill.fore_color.rgb = BRIGHT_BLUE
    bottom_line.line.fill.background()

    # ========== 关键数据看板页 ==========
    if req.data_points and len(req.data_points) > 0:
        slide = prs.slides.add_slide(blank)
        
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = DARK_BLUE
        bg.line.fill.background()
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.4), Inches(12.3), Inches(0.8))
        tp = title_box.text_frame.paragraphs[0]
        tp.text = "核心数据"
        tp.font.size = Pt(28)
        tp.font.bold = True
        tp.font.color.rgb = WHITE
        tp.font.name = "Microsoft YaHei"
        
        card_width = 4.0
        start_x = 0.8
        start_y = 1.5
        gap = 0.3
        
        for i, dp in enumerate(req.data_points[:3]):
            x = start_x + (i % 3) * (card_width + gap)
            y = start_y + (i // 3) * 2.5
            
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, Inches(x), Inches(y), 
                Inches(card_width), Inches(2.0)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = WHITE
            card.line.color.rgb = BRIGHT_BLUE
            
            num_box = slide.shapes.add_textbox(
                Inches(x + 0.2), Inches(y + 0.2), Inches(card_width - 0.4), Inches(1.0)
            )
            np = num_box.text_frame.paragraphs[0]
            
            if dp.get('type') == 'compare':
                np.text = f"{dp.get('before', '')} → {dp.get('after', '')}"
            else:
                np.text = str(dp.get('value', ''))
            
            np.font.size = Pt(32)
            np.font.bold = True
            np.font.color.rgb = BRIGHT_BLUE
            np.alignment = PP_ALIGN.CENTER
            np.font.name = "Microsoft YaHei"
            
            label_box = slide.shapes.add_textbox(
                Inches(x + 0.2), Inches(y + 1.2), Inches(card_width - 0.4), Inches(0.5)
            )
            lp = label_box.text_frame.paragraphs[0]
            lp.text = dp.get('label', '')
            lp.font.size = Pt(14)
            lp.font.color.rgb = TEXT_GRAY
            lp.alignment = PP_ALIGN.CENTER
            lp.font.name = "Microsoft YaHei"

    # ========== 内容页 ==========
    data_img_idx = 0
    
    for idx, s in enumerate(req.slides, 1):
        slide = prs.slides.add_slide(blank)
        
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = WHITE
        bg.line.fill.background()
        
        header_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(1.1))
        header_bg.fill.solid()
        header_bg.fill.fore_color.rgb = LIGHT_BLUE
        header_bg.line.fill.background()
        
        left_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(0.15), Inches(1.1))
        left_bar.fill.solid()
        left_bar.fill.fore_color.rgb = BRIGHT_BLUE
        left_bar.line.fill.background()
        
        page_tag = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.8), Inches(0.25), Inches(1.0), Inches(0.5))
        page_tag.fill.solid()
        page_tag.fill.fore_color.rgb = BRIGHT_BLUE
        page_tag.line.fill.background()
        
        tag_text = slide.shapes.add_textbox(Inches(11.8), Inches(0.25), Inches(1.0), Inches(0.5))
        ttp = tag_text.text_frame.paragraphs[0]
        ttp.text = f"0{idx}" if idx < 10 else str(idx)
        ttp.font.size = Pt(14)
        ttp.font.bold = True
        ttp.font.color.rgb = WHITE
        ttp.alignment = PP_ALIGN.CENTER
        
        title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.25), Inches(11.0), Inches(0.7))
        tp = title_box.text_frame.paragraphs[0]
        tp.text = s.title
        tp.font.size = Pt(24)
        tp.font.bold = True
        tp.font.color.rgb = DARK_BLUE
        tp.font.name = "Microsoft YaHei"
        
        title_lower = s.title.lower()
        need_image = any(k in title_lower for k in ['数据', '图表', '实验', '结果', '对比', '分析', '统计', '调研'])
        has_image = need_image and data_img_idx < len(req.data_images)
        
        if has_image:
            content_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.4), Inches(6.8), Inches(5.8))
            content_bg.fill.solid()
            content_bg.fill.fore_color.rgb = RGBColor(250, 250, 250)
            content_bg.line.color.rgb = RGBColor(220, 220, 220)
            
            content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(6.4), Inches(5.4))
            tf = content_box.text_frame
            tf.word_wrap = True
            
            for i, b in enumerate(s.bullets):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = f"●  {b}"
                p.font.size = Pt(15)
                p.font.color.rgb = TEXT_DARK
                p.space_after = Pt(10)
                p.font.name = "Microsoft YaHei"
            
            img_stream = download_image(req.data_images[data_img_idx])
            if img_stream:
                slide.shapes.add_picture(img_stream, Inches(7.6), Inches(1.6), width=Inches(5.0))
                data_img_idx += 1
        else:
            content_bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.4), Inches(12.3), Inches(5.8))
            content_bg.fill.solid()
            content_bg.fill.fore_color.rgb = RGBColor(250, 250, 250)
            content_bg.line.color.rgb = RGBColor(220, 220, 220)
            
            content_box = slide.shapes.add_textbox(Inches(0.8), Inches(1.6), Inches(11.8), Inches(5.4))
            tf = content_box.text_frame
            tf.word_wrap = True
            
            for i, b in enumerate(s.bullets):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = f"●  {b}"
                p.font.size = Pt(16)
                p.font.color.rgb = TEXT_DARK
                p.space_after = Pt(12)
                p.font.name = "Microsoft YaHei"
        
        footer_line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(7.35), Inches(12.3), Inches(0.02))
        footer_line.fill.solid()
        footer_line.fill.fore_color.rgb = BRIGHT_BLUE
        footer_line.line.fill.background()

    # ========== 结尾页 ==========
    slide = prs.slides.add_slide(blank)
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(7.5))
    bg.fill.solid()
    bg.fill.fore_color.rgb = DARK_BLUE
    bg.line.fill.background()
    
    thanks_box = slide.shapes.add_textbox(Inches(0), Inches(2.8), Inches(13.333), Inches(1.2))
    tp = thanks_box.text_frame.paragraphs[0]
    tp.text = "感谢观看"
    tp.font.size = Pt(48)
    tp.font.bold = True
    tp.font.color.rgb = WHITE
    tp.alignment = PP_ALIGN.CENTER
    tp.font.name = "Microsoft YaHei"

    # 保存
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
