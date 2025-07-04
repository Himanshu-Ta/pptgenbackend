import base64
import io
from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Dict

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN

# --- Pydantic Models for Request Data Validation ---
# This structure must match the JSON sent from the JavaScript frontend.
class Report(BaseModel):
    id: str
    label: str
    tab: str
    elementId: str
    master: str
    title: str

class GenerationRequest(BaseModel):
    selectedReports: List[Report]
    imageDataUrls: Dict[str, str]

# --- FastAPI App Initialization ---
app = FastAPI()

# Configure CORS to allow requests from your frontend (e.g., http://localhost:3000)
# Using "*" is convenient for development but for production, you should restrict it.
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# --- Helper Functions to build the presentation ---

def add_custom_header(slide, active_tab_text):
    """Adds the complex header with perfectly aligned overlapping chevrons."""
    tabs = ['Market Projections', 'Category Overview', 'Consumer Overview', 'A&P Overview']
    
    # Use a light grey background for the header area.
    background_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.33), Inches(0.75))
    background_bar.fill.solid()
    background_bar.fill.fore_color.rgb = RGBColor(0xF1, 0xF1, 0xF1)
    background_bar.line.fill.background() # No border

    # Programmatically create each chevron tab.
    for i, tab_text in enumerate(tabs):
        is_active = (tab_text == active_tab_text)
        color = RGBColor(0xF5, 0x82, 0x20) if is_active else RGBColor(0x6E, 0x6E, 0x6E)
        
        left_pos = Inches(0.5 + (i * 2.0))
        
        chevron = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, left_pos, Inches(0.15), Inches(2.2), Inches(0.45))
        chevron.fill.solid()
        chevron.fill.fore_color.rgb = color
        chevron.line.fill.background()

        text_box = slide.shapes.add_textbox(left_pos, Inches(0.15), Inches(2.0), Inches(0.45))
        p = text_box.text_frame.paragraphs[0]
        p.text = tab_text
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.size = Pt(10)
        p.font.name = 'Arial'
        p.alignment = PP_ALIGN.CENTER
        if is_active:
            p.font.bold = True

    # Add the main logo to the header
    slide.shapes.add_picture('assets/imperial-logo.png', Inches(11.5), Inches(0.1), width=Inches(1.5))

def add_default_header(slide):
    """Adds a simple, default header for subcategory slides."""
    background_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.33), Inches(0.75))
    background_bar.fill.solid()
    background_bar.fill.fore_color.rgb = RGBColor(0xF1, 0xF1, 0xF1)
    background_bar.line.fill.background()
    slide.shapes.add_picture('assets/imperial-logo.png', Inches(11.5), Inches(0.1), width=Inches(1.5))

def add_title_slide(pres):
    """Adds the main title slide."""
    slide_layout = pres.slide_layouts[6] # Blank slide layout
    slide = pres.slides.add_slide(slide_layout)
    try:
        slide.shapes.add_picture('assets/title-background.jpg', 0, 0, width=pres.slide_width, height=pres.slide_height)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.4), pres.slide_width, Inches(2.1))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.line.fill.background()

        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(5.8), Inches(10), Inches(0.5))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.text = "PORTFOLIO STRATEGY WORKSHOP"
        p_title.font.bold = True
        p_title.font.size = Pt(24)
        p_title.font.color.rgb = RGBColor(0xF5, 0x82, 0x20)

        preread_box = slide.shapes.add_textbox(Inches(0.6), Inches(6.2), Inches(10), Inches(0.5))
        p_preread = preread_box.text_frame.paragraphs[0]
        p_preread.text = "Pre-read"
        p_preread.font.size = Pt(18)
        p_preread.font.color.rgb = RGBColor(0x6E, 0x6E, 0x6E)
        
        slide.shapes.add_picture('assets/imperial-logo.png', Inches(11.5), Inches(5.8), width=Inches(1.5))
    except FileNotFoundError:
        error_box = slide.shapes.add_textbox(0, 0, pres.slide_width, pres.slide_height)
        error_box.text_frame.paragraphs[0].text = "Error: Asset files not found. Make sure the 'assets' folder is in the same directory as app.py."

def add_contents_slide(pres, selected_reports):
    """Adds a dynamically generated table of contents."""
    slide_layout = pres.slide_layouts[5] # Title and Content layout
    slide = pres.slides.add_slide(slide_layout)
    slide.shapes.add_picture('assets/imperial-logo.png', Inches(11.5), Inches(0.4), width=Inches(1.5))
    
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(10), Inches(0.5))
    p_title = title_box.text_frame.paragraphs[0]
    p_title.text = "Contents"
    p_title.font.size = Pt(20)
    p_title.font.color.rgb = RGBColor(0x6E, 0x6E, 0x6E)

    y_pos = Inches(2.8)
    for i, report in enumerate(selected_reports):
        num_p = slide.shapes.add_textbox(Inches(1.0), y_pos, Inches(0.5), Inches(0.5)).text_frame.paragraphs[0]
        num_p.text = f"{i + 1}."
        num_p.font.size = Pt(18)
        num_p.font.color.rgb = RGBColor(0xA9, 0xA9, 0xA9)
        text_p = slide.shapes.add_textbox(Inches(1.5), y_pos, Inches(10), Inches(0.5)).text_frame.paragraphs[0]
        text_p.text = report.label
        text_p.font.size = Pt(18)
        text_p.font.color.rgb = RGBColor(0xA9, 0xA9, 0xA9)
        y_pos += Inches(0.6)

def add_chart_slide(pres, report_data, image_file):
    """
    Adds a slide with a chart image, resizing it to fit a larger bounding box
    while preserving aspect ratio and centering it.
    """
    slide_layout = pres.slide_layouts[6] # Blank layout
    slide = pres.slides.add_slide(slide_layout)
    
    if "HEADER_MASTER" in report_data.master:
        add_custom_header(slide, report_data.tab)
    else:
        add_default_header(slide)
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(10), Inches(0.5))
    p_title = title_box.text_frame.paragraphs[0]
    p_title.text = report_data.title
    p_title.font.bold = True
    p_title.font.size = Pt(22)
    p_title.font.color.rgb = RGBColor(0xF5, 0x82, 0x20)

    key_msg_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.5), Inches(12.3), Inches(0.5))
    key_msg_box.fill.background()
    key_msg_box.line.color.rgb = RGBColor(211, 211, 211)
    key_msg_box.text_frame.text = "▪ [Type key message]"
    
    # --- Bounding Box and Image Placement Logic ---
    content_area_left = Inches(0.5)
    content_area_top = Inches(2.1) # Moved content area up slightly
    content_area_max_width = Inches(12.3)
    content_area_max_height = Inches(5.3) # Increased max height to make images larger

    # Add the picture with a placeholder width to get its native dimensions scaled.
    pic = slide.shapes.add_picture(image_file, content_area_left, content_area_top, width=content_area_max_width)

    # Calculate the aspect ratios of the image and the bounding box.
    image_ratio = pic.width / pic.height
    box_ratio = content_area_max_width / content_area_max_height

    # Resize the image to fit the bounding box without distortion.
    if image_ratio > box_ratio:
        # Image is wider than the box: fit to width.
        pic.width = content_area_max_width
        pic.height = int(content_area_max_width / image_ratio)
    else:
        # Image is taller than or equal to the box: fit to height.
        pic.height = content_area_max_height
        pic.width = int(content_area_max_height * image_ratio)

    # Center the resized image within the content area.
    pic.left = int(content_area_left + (content_area_max_width - pic.width) / 2)
    pic.top = int(content_area_top + (content_area_max_height - pic.height) / 2)


# --- API Endpoint ---

@app.post("/generate-ppt")
async def generate_ppt_endpoint(request_data: GenerationRequest):
    """
    This endpoint receives report selections and chart images, then
    assembles and returns a PowerPoint presentation.
    """
    pres = Presentation()
    pres.slide_width = Inches(13.33)
    pres.slide_height = Inches(7.5)

    add_title_slide(pres)
    add_contents_slide(pres, request_data.selectedReports)

    for report in request_data.selectedReports:
        image_data_url = request_data.imageDataUrls.get(report.id)
        if image_data_url:
            # The data URL is in the format "data:image/png;base64,ENCODED_STRING"
            # We need to split it and decode the base64 part.
            try:
                header, encoded = image_data_url.split(",", 1)
                image_data = io.BytesIO(base64.b64decode(encoded))
                add_chart_slide(pres, report, image_data)
            except Exception as e:
                print(f"Could not process image for report {report.id}: {e}")
                # Optionally, add an error slide for this report
                continue

    # Save the presentation to an in-memory buffer
    ppt_buffer = io.BytesIO()
    pres.save(ppt_buffer)
    ppt_buffer.seek(0)

    # Stream the buffer back to the client as a .pptx file download
    return StreamingResponse(
        ppt_buffer,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        headers={'Content-Disposition': 'attachment; filename=Dashboard_Report_Final.pptx'}
    )

# To run this app from your terminal:
# 1. Make sure you have an 'assets' folder with the required images.
# 2. Run the command: uvicorn app:app --reload --port 5001
