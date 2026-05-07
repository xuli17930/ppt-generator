from fastapi import FastAPI
from fastapi.responses import FileResponse
from pydantic import BaseModel
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
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

# ========== 配色方案：现代商务蓝 ==========
DARK_BLUE = RGBColor(0, 32, 96)           # 主色-深海蓝
BRIGHT_BLUE = RGBColor(0, 112, 192)        # 强调色-科技蓝
LIGHT_BLUE = RGBColor(230, 240, 255)     # 背景色-淡蓝
TEXT_DARK = RGBColor(33, 33, 33)           # 正文-深灰
TEXT_GRAY = RGBColor(100, 100, 100)       # 副文-中灰
WHITE = RGBColor(255, 255, 255)

@app.post("/generate")
def generate_ppt(req: PPTRequest):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank = prs.slide_layouts[6]

    # ========== 封面页（全屏渐变风格） ==========
    slide = prs.slides.add_slide(blank)
    
    # 背景色块（全页淡蓝底）
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
        Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = LIGHT_BLUE
    bg.line.fill.background()
    
    # 顶部深色装饰条
    top_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
        Inches(13.333), Inches(0.25)
    )
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = DARK_BLUE
    top_bar.line.fill.background()
    
    # 左侧竖条装饰
    left_accent = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.8),
        Inches(0.08), Inches(3.5)
    )
    left_accent.fill.solid()
    left_accent.fill.fore_color.rgb = BRIGHT_BLUE
    left_accent.line.fill.background()
    
    # 主标题
    tb = slide.shapes.add_textbox(
        Inches(1.2), Inches(2.0), Inches(11.5), Inches(1.5)
    )
    p = tb.text_frame.paragraphs[0]
    p.text = req.main_title
    p.font.size = Pt(40)
    p.font.bold = True
    p.font.color.rgb = DARK_BLUE
    p.font.name = "Microsoft YaHei"
    p.alignment = PP_ALIGN.LEFT
    
    # 副标题
    tb2 = slide.shapes.add_textbox(
        Inches(1.2), Inches(3.7), Inches(11.5), Inches(1.0)
    )
    p2 = tb2.text_frame.paragraphs[0]
    p2.text = req.subtitle
    p2.font.size = Pt(20)
    p2.font.color.rgb = TEXT_GRAY
    p2.font.name = "Microsoft YaHei"
    p2.alignment = PP_ALIGN.LEFT
    
    # 底部信息线
    bottom_line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(1.2), Inches(5.2),
        Inches(3.0), Inches(0.03)
    )
    bottom_line.fill.solid()
    bottom_line.fill.fore_color.rgb = BRIGHT_BLUE
    bottom_line.line.fill.background()
    
    # 底部日期
    date_box = slide.shapes.add_textbox(
        Inches(1.2), Inches(5.4), Inches(5.0), Inches(0.4)
    )
    dp = date_box.text_frame.paragraphs[0]
    dp.text = "自动生成报告"
    dp.font.size = Pt(12)
    dp.font.color.rgb = TEXT_GRAY
    dp.font.name = "Microsoft YaHei"

    # ========== 目录页（可选，如果内容多） ==========
    if len(req.slides) > 6:
        slide = prs.slides.add_slide(blank)
        # 标题
        toc_title = slide.shapes.add_textbox(
            Inches(0.8), Inches(0.5), Inches(12), Inches(0.8)
        )
        tp = toc_title.text_frame.paragraphs[0]
        tp.text = "目录 CONTENTS"
        tp.font.size = Pt(28)
        tp.font.bold = True
        tp.font.color.rgb = DARK_BLUE
        tp.font.name = "Microsoft YaHei"
        
        # 目录项
        for i, s in enumerate(req.slides[:8], 1):
            y_pos = 1.4 + (i - 1) * 0.65
            # 序号圆圈
            circle = slide.shapes.add_shape(
                MSO_SHAPE.OVAL, Inches(1.0), Inches(y_pos),
                Inches(0.35), Inches(0.35)
            )
            circle.fill.solid()
            circle.fill.fore_color.rgb = BRIGHT_BLUE
            circle.line.fill.background()
            
            # 序号文字
            num_box = slide.shapes.add_textbox(
                Inches(1.0), Inches(y_pos), Inches(0.35), Inches(0.35)
            )
            np = num_box.text_frame.paragraphs[0]
            np.text = str(i)
            np.font.size = Pt(14)
            np.font.bold = True
            np.font.color.rgb = WHITE
            np.alignment = PP_ALIGN.CENTER
            
            # 标题文字
            toc_box = slide.shapes.add_textbox(
                Inches(1.6), Inches(y_pos), Inches(10), Inches(0.4)
            )
            toc_p = toc_box.text_frame.paragraphs[0]
            toc_p.text = s.title
            toc_p.font.size = Pt(16)
            toc_p.font.color.rgb = TEXT_DARK
            toc_p.font.name = "Microsoft YaHei"

    # ========== 内容页（精美卡片式布局） ==========
    for idx, s in enumerate(req.slides, 1):
        slide = prs.slides.add_slide(blank)
        
        # 页面背景
        bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(13.333), Inches(7.5)
        )
        bg.fill.solid()
        bg.fill.fore_color.rgb = WHITE
        bg.line.fill.background()
        
        # 顶部标题栏背景
        header_bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(13.333), Inches(1.1)
        )
        header_bg.fill.solid()
        header_bg.fill.fore_color.rgb = LIGHT_BLUE
        header_bg.line.fill.background()
        
        # 左侧竖条
        left_bar = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
            Inches(0.15), Inches(1.1)
        )
        left_bar.fill.solid()
        left_bar.fill.fore_color.rgb = BRIGHT_BLUE
        left_bar.line.fill.background()
        
        # 页码标签
        page_tag = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(11.8), Inches(0.25),
            Inches(1.0), Inches(0.5)
        )
        page_tag.fill.solid()
        page_tag.fill.fore_color.rgb = BRIGHT_BLUE
        page_tag.line.fill.background()
        
        tag_text = slide.shapes.add_textbox(
            Inches(11.8), Inches(0.25), Inches(1.0), Inches(0.5)
        )
        ttp = tag_text.text_frame.paragraphs[0]
        ttp.text = f"0{idx}" if idx < 10 else str(idx)
        ttp.font.size = Pt(14)
        ttp.font.bold = True
        ttp.font.color.rgb = WHITE
        ttp.alignment = PP_ALIGN.CENTER
        
        # 页面标题
        title_box = slide.shapes.add_textbox(
            Inches(0.5), Inches(0.25), Inches(11.0), Inches(0.7)
        )
        tp = title_box.text_frame.paragraphs[0]
        tp.text = s.title
        tp.font.size = Pt(24)
        tp.font.bold = True
        tp.font.color.rgb = DARK_BLUE
        tp.font.name = "Microsoft YaHei"
        
        # 内容区域（带左边距的卡片效果）
        content_bg = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.4),
            Inches(12.3), Inches(5.8)
        )
        content_bg.fill.solid()
        content_bg.fill.fore_color.rgb = RGBColor(250, 250, 250)
        content_bg.line.color.rgb = RGBColor(220, 220, 220)
        
        # 内容文字
        content_box = slide.shapes.add_textbox(
            Inches(0.8), Inches(1.6), Inches(11.8), Inches(5.4)
        )
        tf = content_box.text_frame
        tf.word_wrap = True
        
        for i, b in enumerate(s.bullets):
            p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
            p.text = f"●  {b}"
            p.font.size = Pt(16)
            p.font.color.rgb = TEXT_DARK
            p.space_after = Pt(12)
            p.font.name = "Microsoft YaHei"
            p.level = 0
        
        # 底部装饰线
        footer_line = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(7.35),
            Inches(12.3), Inches(0.02)
        )
        footer_line.fill.solid()
        footer_line.fill.fore_color.rgb = BRIGHT_BLUE
        footer_line.line.fill.background()

    # ========== 结尾页（Thank You） ==========
    slide = prs.slides.add_slide(blank)
    
    # 背景
    bg = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0),
        Inches(13.333), Inches(7.5)
    )
    bg.fill.solid()
    bg.fill.fore_color.rgb = DARK_BLUE
    bg.line.fill.background()
    
    # 感谢文字
    thanks_box = slide.shapes.add_textbox(
        Inches(0), Inches(2.8), Inches(13.333), Inches(1.2)
    )
    tp = thanks_box.text_frame.paragraphs[0]
    tp.text = "感谢观看"
    tp.font.size = Pt(48)
    tp.font.bold = True
    tp.font.color.rgb = WHITE
    tp.alignment = PP_ALIGN.CENTER
    tp.font.name = "Microsoft YaHei"
    
    # 副标题
    sub_box = slide.shapes.add_textbox(
        Inches(0), Inches(4.2), Inches(13.333), Inches(0.8)
    )
    sp = sub_box.text_frame.paragraphs[0]
    sp.text = "THANK YOU FOR WATCHING"
    sp.font.size = Pt(18)
    sp.font.color.rgb = RGBColor(180, 200, 230)
    sp.alignment = PP_ALIGN.CENTER
    sp.font.name = "Microsoft YaHei"

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
