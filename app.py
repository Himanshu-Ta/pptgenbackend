import base64
import io
import os
import re
import json
import asyncio
import hashlib
from fastapi import FastAPI
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
from typing import List, Dict, Any, Tuple

# --- Image Processing & ML Imports ---
import cv2
import easyocr
import httpx
import numpy as np
from PIL import Image # --- REQUIRED: Used for accurate image dimension reading ---

# --- Presentation Imports ---
from pptx import Presentation
from pptx.chart.data import BubbleChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# --- Import data for the native chart from chart_data.py ---
# Make sure you have a chart_data.py file in the same directory
from chart_data import price_ladder_data, brand_colors, price_segments, brand_summary_data

# --- Initialize OCR Reader ---
print("Initializing EasyOCR reader...")
ocr_reader = easyocr.Reader(["en"], gpu=True)
print("EasyOCR reader initialized.")

# --- API Credentials ---
COHERE_API_KEY = "xDfWr7AoPOT2w2vgAwI62UPHJhDIRcLgsLMfXzqf"

# --- Pydantic Models ---
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
app.add_middleware(
    CORSMiddleware, allow_origins=["*"], allow_credentials=True, allow_methods=["*"], allow_headers=["*"]
)

# --- File-based Cache Setup ---
CACHE_DIR = "ppt_cache"
os.makedirs(CACHE_DIR, exist_ok=True)

# --- Helper function to convert hex to RGB ---
def hex_to_rgb(hex_color):
    hex_color = hex_color.lstrip('#')
    return tuple(int(hex_color[i:i+2], 16) for i in (0, 2, 4))

# --- Helper Functions for Chart Processing ---
def extract_chart_details_from_memory(image_bytes: bytes) -> List[Dict[str, Any]]:
    try:
        nparr = np.frombuffer(image_bytes, np.uint8)
        img_np = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        ocr_results = ocr_reader.readtext(img_np, detail=0, paragraph=True)
        return [{"extracted_text": " ".join(ocr_results)}]
    except Exception as e:
        print(f"Error during EasyOCR processing: {e}")
        return [{"error": "Could not process chart."}]

async def get_summary_from_text(title: str, text_data: str) -> str:
    """Generic function to get a summary from any text data."""
    if not COHERE_API_KEY or COHERE_API_KEY == "your_cohere_api_key_here":
        return "▪ [Cohere API key not configured]"
    prompt = f"Provide a single, brief summary sentence for a presentation slide. The chart is titled: '{title}'. The extracted data is: {text_data}"
    headers = {"Authorization": f"Bearer {COHERE_API_KEY}", "Content-Type": "application/json"}
    payload = {"model": "command-r-plus", "message": prompt}
    try:
        async with httpx.AsyncClient(verify=False) as client:
            response = await client.post("https://api.cohere.com/v1/chat", json=payload, headers=headers, timeout=60)
            response.raise_for_status()
        result = response.json()
        summary = result.get("text", "▪ [AI summary could not be generated.]").strip()
        return re.sub(r'^(here are the key findings|here is the summary|analysis:|summary):?\s*', '', summary, flags=re.IGNORECASE).strip()
    except Exception as e:
        print(f"Error communicating with Cohere API: {e}")
        return "▪ [Error: Could not retrieve AI summary.]"

async def process_screenshot_report(report: Report, image_data_url: str) -> Dict[str, Any]:
    header, encoded = image_data_url.split(",", 1)
    image_bytes = base64.b64decode(encoded)
    loop = asyncio.get_running_loop()
    extracted_details = await loop.run_in_executor(None, extract_chart_details_from_memory, image_bytes)
    summary_text = await get_summary_from_text(report.title, json.dumps(extracted_details))
    return {"id": report.id, "report": report, "image_bytes": image_bytes, "summary": summary_text}

async def get_summary_for_native_chart(report: Report) -> str:
    """Generates a summary for the native bubble chart."""
    description = f"A bubble chart showing the price ladder for various brands. Brands include {', '.join([s['series_name'] for s in price_ladder_data])}."
    return await get_summary_from_text(report.title, description)


# --- Slide Creation Functions ---
def add_custom_header(slide, active_tab_text):
    tabs = ["Market Projections", "Category Overview", "Consumer Overview", "A&P Overview"]
    background_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.33), Inches(0.75))
    background_bar.fill.solid()
    background_bar.fill.fore_color.rgb = RGBColor(0xF1, 0xF1, 0xF1)
    background_bar.line.fill.background()
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
        p.font.name = "Arial"
        p.font.color.rgb = RGBColor(255, 255, 255)
        p.font.size = Pt(10)
        p.alignment = PP_ALIGN.CENTER
        text_box.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
        if is_active:
            p.font.bold = True
    try:
        slide.shapes.add_picture("assets/imperial-logo.png", Inches(11.5), Inches(0.1), width=Inches(1.5))
    except FileNotFoundError:
        pass

def add_title_slide(pres):
    slide_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(slide_layout)
    try:
        slide.shapes.add_picture("assets/title-background.jpg", 0, 0, width=pres.slide_width, height=pres.slide_height)
        shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.4), pres.slide_width, Inches(2.1))
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
        shape.line.fill.background()
        title_box = slide.shapes.add_textbox(Inches(0.6), Inches(5.8), Inches(10), Inches(0.5))
        p_title = title_box.text_frame.paragraphs[0]
        p_title.text = "PORTFOLIO STRATEGY WORKSHOP"
        p_title.font.name = "Arial"
        p_title.font.bold = True
        p_title.font.size = Pt(24)
        p_title.font.color.rgb = RGBColor(0xF5, 0x82, 0x20)
        preread_box = slide.shapes.add_textbox(Inches(0.6), Inches(6.2), Inches(10), Inches(0.5))
        p_preread = preread_box.text_frame.paragraphs[0]
        p_preread.text = "Pre-read"
        p_preread.font.name = "Arial"
        p_preread.font.size = Pt(18)
        p_preread.font.color.rgb = RGBColor(0x6E, 0x6E, 0x6E)
        slide.shapes.add_picture("assets/imperial-logo.png", Inches(11.5), Inches(5.8), width=Inches(1.5))
    except FileNotFoundError:
        pass

def add_contents_slide(pres, selected_reports):
    slide_layout = pres.slide_layouts[5]
    slide = pres.slides.add_slide(slide_layout)
    try:
        slide.shapes.add_picture("assets/imperial-logo.png", Inches(11.5), Inches(0.4), width=Inches(1.5))
    except FileNotFoundError:
        pass
    title_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(10), Inches(0.5))
    p_title = title_box.text_frame.paragraphs[0]
    p_title.text = "Contents"
    p_title.font.name = "Arial"
    p_title.font.size = Pt(20)
    p_title.font.color.rgb = RGBColor(0x6E, 0x6E, 0x6E)
    y_pos = Inches(2.8)
    for i, report in enumerate(selected_reports):
        num_p = slide.shapes.add_textbox(Inches(1.0), y_pos, Inches(0.5), Inches(0.5)).text_frame.paragraphs[0]
        num_p.text = f"{i + 1}."
        num_p.font.name = "Arial"
        num_p.font.size = Pt(18)
        num_p.font.color.rgb = RGBColor(0xA9, 0xA9, 0xA9)
        text_p = slide.shapes.add_textbox(Inches(1.5), y_pos, Inches(10), Inches(0.5)).text_frame.paragraphs[0]
        text_p.text = report.label
        text_p.font.name = "Arial"
        text_p.font.size = Pt(18)
        text_p.font.color.rgb = RGBColor(0x6E, 0x6E, 0x6E)
        y_pos += Inches(0.6)

def add_chart_slide_from_image(pres, report_data: Report, image_bytes: bytes, summary_text: str):
    slide_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(slide_layout)
    add_custom_header(slide, report_data.tab)
    
    title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(12.3), Inches(0.5))
    p_title = title_box.text_frame.paragraphs[0]
    p_title.text = report_data.title
    p_title.font.name = "Arial"
    p_title.font.bold = True
    p_title.font.size = Pt(22)
    p_title.font.color.rgb = RGBColor(0xF5, 0x82, 0x20)
    
    # --- FIX APPLIED HERE: More robust image scaling and positioning ---
    
    # Define the bounding box for the chart image
    box_left = Inches(0.5)
    box_top = Inches(1.6)
    box_width = Inches(8.5)
    box_height = Inches(5.5)
    
    # Get actual image dimensions using Pillow
    image_stream = io.BytesIO(image_bytes)
    try:
        with Image.open(image_stream) as img:
            img_width, img_height = img.size
    except Exception as e:
        print(f"Could not read image dimensions with Pillow: {e}")
        # Fallback to a default size if image is unreadable
        img_width, img_height = box_width, box_height

    # Calculate aspect ratios
    img_ratio = float(img_width) / float(img_height) if img_height > 0 else 1.0
    box_ratio = float(box_width) / float(box_height)

    # Determine the final size of the image on the slide to fit the box
    if img_ratio > box_ratio:
        # Image is wider than the box, scale by width
        final_width = box_width
        final_height = final_width / img_ratio
    else:
        # Image is taller than the box, scale by height
        final_height = box_height
        final_width = final_height * img_ratio

    # Calculate position to center the image within the bounding box
    final_left = box_left + (box_width - final_width) / 2
    final_top = box_top + (box_height - final_height) / 2

    # Add the picture with the calculated size and position
    image_stream.seek(0) # Reset stream for pptx
    slide.shapes.add_picture(image_stream, final_left, final_top, width=final_width, height=final_height)
    
    # Position the insights box relative to the bounding box
    insights_left = box_left + box_width + Inches(0.3)
    insights_box = slide.shapes.add_textbox(insights_left, box_top, Inches(3.5), box_height)
    insights_frame = insights_box.text_frame
    insights_frame.word_wrap = True
    p_heading = insights_frame.paragraphs[0]
    p_heading.text = "Key Insight"
    p_heading.font.name = 'Arial'
    p_heading.font.bold = True
    p_heading.font.size = Pt(16)
    p_heading.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    p_summary = insights_frame.add_paragraph()
    p_summary.text = f"▪ {summary_text}"
    p_summary.font.name = 'Arial'
    p_summary.font.size = Pt(12)
    p_summary.line_spacing = 1.5

def add_live_bubble_chart_slide(pres, report_data: Report, summary_text: str):
    slide_layout = pres.slide_layouts[6]
    slide = pres.slides.add_slide(slide_layout)
    add_custom_header(slide, report_data.tab)
    
    title_shape = slide.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(12.3), Inches(0.5))
    p_title = title_shape.text_frame.paragraphs[0]
    p_title.text = report_data.title
    p_title.font.name = "Arial"
    p_title.font.bold = True
    p_title.font.size = Pt(22)
    p_title.font.color.rgb = RGBColor(0xF5, 0x82, 0x20)
    
    # --- FIX APPLIED HERE: Consistent Layout Definition ---
    
    # Define layout areas consistent with the image-based slides
    box_left = Inches(0.5)
    box_top = Inches(1.6)
    box_width = Inches(8.5)
    box_height = Inches(5.5)
    
    # Position the insights box to the right
    insights_left = box_left + box_width + Inches(0.3)
    insights_box = slide.shapes.add_textbox(insights_left, box_top, Inches(3.5), box_height)
    insights_frame = insights_box.text_frame
    insights_frame.word_wrap = True
    p_heading = insights_frame.paragraphs[0]
    p_heading.text = "Key Insight"
    p_heading.font.name = 'Arial'; p_heading.font.bold = True; p_heading.font.size = Pt(16); p_heading.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    p_summary = insights_frame.add_paragraph()
    p_summary.text = f"▪ {summary_text}"
    p_summary.font.name = 'Arial'; p_summary.font.size = Pt(12); p_summary.line_spacing = 1.5

    # --- Chart Creation and Positioning ---
    # Position the chart and its elements within the defined chart_area
    chart_x = Inches(2.0) # Start chart a bit to the right to make space for labels
    chart_y = box_top
    chart_cx = Inches(6.8) # Adjust width to fit in the area
    chart_cy = Inches(4.0)
    table_y = chart_y + chart_cy + Inches(0.2)
    
    chart_data = BubbleChartData()
    for series_data in price_ladder_data:
        series = chart_data.add_series(series_data["series_name"])
        for point in series_data["points"]:
            series.add_data_point(point[0], point[1], point[2])
            
    graphic_frame = slide.shapes.add_chart(XL_CHART_TYPE.BUBBLE, chart_x, chart_y, chart_cx, chart_cy, chart_data)
    chart = graphic_frame.chart
    chart.has_legend = False
    
    value_axis = chart.value_axis
    value_axis.has_title = True; value_axis.axis_title.text_frame.text = "WAP (per 20 stick pack)"; value_axis.minimum_scale = 0.0; value_axis.maximum_scale = 7.0; value_axis.major_unit = 1.0; value_axis.has_major_gridlines = True; value_axis.major_gridlines.format.line.dash_style = 3
    
    category_axis = chart.category_axis
    category_axis.has_major_gridlines = False; category_axis.has_minor_gridlines = False
    
    for i, series in enumerate(chart.series):
        series.format.fill.solid()
        series.format.fill.fore_color.rgb = RGBColor(*hex_to_rgb(brand_colors[i]))

    for seg in price_segments:
        textbox = slide.shapes.add_textbox(box_left, chart_y + chart_cy * (1 - (seg['y'] / 7.0)) - Inches(0.2), Inches(1.5), Inches(0.4))
        p = textbox.text_frame.paragraphs[0]; p.text = seg['name']; p.font.size = Pt(8); p.alignment = PP_ALIGN.RIGHT

    rows, cols = 4, len(brand_summary_data) + 1
    table_shape = slide.shapes.add_table(rows, cols, box_left, table_y, box_width, Inches(1.0)); table = table_shape.table
    table.cell(0, 0).text = "Brand"; table.cell(1, 0).text = "MS%"; table.cell(2, 0).text = "# SKUS"; table.cell(3, 0).text = "Avg. SKU MS%"
    for i, brand_data in enumerate(brand_summary_data):
        col_idx = i + 1; table.cell(0, col_idx).text = brand_data['brand']; table.cell(1, col_idx).text = brand_data['ms']; table.cell(2, col_idx).text = str(brand_data['skus']); table.cell(3, col_idx).text = brand_data['avgMs']

# --- Main API Endpoint with Hybrid Logic ---
@app.post("/generate-ppt")
async def generate_ppt_endpoint(request_data: GenerationRequest):
    report_ids = sorted([report.id for report in request_data.selectedReports])
    key_string = ":".join(report_ids)
    hashed_key = hashlib.sha256(key_string.encode()).hexdigest()
    cache_filepath = os.path.join(CACHE_DIR, f"{hashed_key}.pptx")

    if os.path.exists(cache_filepath):
        print(f"CACHE HIT: Serving report from file: {cache_filepath}")
        return FileResponse(path=cache_filepath, media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation', filename='AI_Generated_Report.pptx')

    print(f"CACHE MISS: Generating new report for key: {hashed_key}")
    
    screenshot_tasks = []
    native_chart_tasks = {}

    for r in request_data.selectedReports:
        if r.id == 'priceladder':
            native_chart_tasks[r.id] = asyncio.create_task(get_summary_for_native_chart(r))
        else:
            image_url = request_data.imageDataUrls.get(r.id)
            if image_url:
                screenshot_tasks.append(process_screenshot_report(r, image_url))
    
    processed_screenshots = await asyncio.gather(*screenshot_tasks)
    processed_map = {item['id']: item for item in processed_screenshots}
    
    native_chart_summaries = {task_id: await task for task_id, task in native_chart_tasks.items()}

    pres = Presentation()
    pres.slide_width = Inches(13.33)
    pres.slide_height = Inches(7.5)
    add_title_slide(pres)
    add_contents_slide(pres, request_data.selectedReports)

    for report in request_data.selectedReports:
        if report.id == 'priceladder':
            summary = native_chart_summaries.get(report.id, "Insight could not be generated.")
            add_live_bubble_chart_slide(pres, report, summary)
        elif report.id in processed_map:
            processed_data = processed_map[report.id]
            add_chart_slide_from_image(pres, processed_data['report'], processed_data['image_bytes'], processed_data['summary'])

    ppt_buffer = io.BytesIO()
    pres.save(ppt_buffer)
    ppt_buffer.seek(0)
    
    with open(cache_filepath, "wb") as f:
        f.write(ppt_buffer.getvalue())
    print(f"SAVED TO CACHE: {cache_filepath}")

    return StreamingResponse(
        iter([ppt_buffer.getvalue()]),
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        headers={'Content-Disposition': 'attachment; filename=Hybrid_Report.pptx'}
    )



# import base64
# import io
# from fastapi import FastAPI
# from fastapi.responses import StreamingResponse
# from fastapi.middleware.cors import CORSMiddleware
# from pydantic import BaseModel
# from typing import List, Dict

# from pptx import Presentation
# from pptx.util import Inches, Pt
# from pptx.enum.shapes import MSO_SHAPE
# from pptx.dml.color import RGBColor
# from pptx.enum.text import PP_ALIGN

# # --- Pydantic Models for Request Data Validation ---
# # This structure must match the JSON sent from the JavaScript frontend.
# class Report(BaseModel):
#     id: str
#     label: str
#     tab: str
#     elementId: str
#     master: str
#     title: str

# class GenerationRequest(BaseModel):
#     selectedReports: List[Report]
#     imageDataUrls: Dict[str, str]

# # --- FastAPI App Initialization ---
# app = FastAPI()

# # Configure CORS to allow requests from your frontend (e.g., http://localhost:3000)
# # Using "*" is convenient for development but for production, you should restrict it.
# app.add_middleware(
#     CORSMiddleware,
#     allow_origins=["*"],
#     allow_credentials=True,
#     allow_methods=["*"],
#     allow_headers=["*"],
# )

# # --- Helper Functions to build the presentation ---

# def add_custom_header(slide, active_tab_text):
#     """Adds the complex header with perfectly aligned overlapping chevrons."""
#     tabs = ['Market Projections', 'Category Overview', 'Consumer Overview', 'A&P Overview']

#     # Use a light grey background for the header area.
#     background_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.33), Inches(0.75))
#     background_bar.fill.solid()
#     background_bar.fill.fore_color.rgb = RGBColor(0xF1, 0xF1, 0xF1)
#     background_bar.line.fill.background() # No border

#     # Programmatically create each chevron tab.
#     for i, tab_text in enumerate(tabs):
#         is_active = (tab_text == active_tab_text)
#         color = RGBColor(0xF5, 0x82, 0x20) if is_active else RGBColor(0x6E, 0x6E, 0x6E)

#         left_pos = Inches(0.5 + (i * 2.0))

#         chevron = slide.shapes.add_shape(MSO_SHAPE.CHEVRON, left_pos, Inches(0.15), Inches(2.2), Inches(0.45))
#         chevron.fill.solid()
#         chevron.fill.fore_color.rgb = color
#         chevron.line.fill.background()

#         text_box = slide.shapes.add_textbox(left_pos, Inches(0.15), Inches(2.0), Inches(0.45))
#         p = text_box.text_frame.paragraphs[0]
#         p.text = tab_text
#         p.font.color.rgb = RGBColor(255, 255, 255)
#         p.font.size = Pt(10)
#         p.font.name = 'Arial'
#         p.alignment = PP_ALIGN.CENTER
#         if is_active:
#             p.font.bold = True

#     # Add the main logo to the header
#     slide.shapes.add_picture('assets/imperial-logo.png', Inches(11.5), Inches(0.1), width=Inches(1.5))

# def add_default_header(slide):
#     """Adds a simple, default header for subcategory slides."""
#     background_bar = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.33), Inches(0.75))
#     background_bar.fill.solid()
#     background_bar.fill.fore_color.rgb = RGBColor(0xF1, 0xF1, 0xF1)
#     background_bar.line.fill.background()
#     slide.shapes.add_picture('assets/imperial-logo.png', Inches(11.5), Inches(0.1), width=Inches(1.5))

# def add_title_slide(pres):
#     """Adds the main title slide."""
#     slide_layout = pres.slide_layouts[6] # Blank slide layout
#     slide = pres.slides.add_slide(slide_layout)
#     try:
#         slide.shapes.add_picture('assets/title-background.jpg', 0, 0, width=pres.slide_width, height=pres.slide_height)
#         shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(5.4), pres.slide_width, Inches(2.1))
#         shape.fill.solid()
#         shape.fill.fore_color.rgb = RGBColor(255, 255, 255)
#         shape.line.fill.background()

#         title_box = slide.shapes.add_textbox(Inches(0.6), Inches(5.8), Inches(10), Inches(0.5))
#         p_title = title_box.text_frame.paragraphs[0]
#         p_title.text = "PORTFOLIO STRATEGY WORKSHOP"
#         p_title.font.bold = True
#         p_title.font.size = Pt(24)
#         p_title.font.color.rgb = RGBColor(0xF5, 0x82, 0x20)

#         preread_box = slide.shapes.add_textbox(Inches(0.6), Inches(6.2), Inches(10), Inches(0.5))
#         p_preread = preread_box.text_frame.paragraphs[0]
#         p_preread.text = "Pre-read"
#         p_preread.font.size = Pt(18)
#         p_preread.font.color.rgb = RGBColor(0x6E, 0x6E, 0x6E)

#         slide.shapes.add_picture('assets/imperial-logo.png', Inches(11.5), Inches(5.8), width=Inches(1.5))
#     except FileNotFoundError:
#         error_box = slide.shapes.add_textbox(0, 0, pres.slide_width, pres.slide_height)
#         error_box.text_frame.paragraphs[0].text = "Error: Asset files not found. Make sure the 'assets' folder is in the same directory as app.py."

# def add_contents_slide(pres, selected_reports):
#     """Adds a dynamically generated table of contents."""
#     slide_layout = pres.slide_layouts[5] # Title and Content layout
#     slide = pres.slides.add_slide(slide_layout)
#     slide.shapes.add_picture('assets/imperial-logo.png', Inches(11.5), Inches(0.4), width=Inches(1.5))

#     title_box = slide.shapes.add_textbox(Inches(0.6), Inches(1.8), Inches(10), Inches(0.5))
#     p_title = title_box.text_frame.paragraphs[0]
#     p_title.text = "Contents"
#     p_title.font.size = Pt(20)
#     p_title.font.color.rgb = RGBColor(0x6E, 0x6E, 0x6E)

#     y_pos = Inches(2.8)
#     for i, report in enumerate(selected_reports):
#         num_p = slide.shapes.add_textbox(Inches(1.0), y_pos, Inches(0.5), Inches(0.5)).text_frame.paragraphs[0]
#         num_p.text = f"{i + 1}."
#         num_p.font.size = Pt(18)
#         num_p.font.color.rgb = RGBColor(0xA9, 0xA9, 0xA9)
#         text_p = slide.shapes.add_textbox(Inches(1.5), y_pos, Inches(10), Inches(0.5)).text_frame.paragraphs[0]
#         text_p.text = report.label
#         text_p.font.size = Pt(18)
#         text_p.font.color.rgb = RGBColor(0xA9, 0xA9, 0xA9)
#         y_pos += Inches(0.6)

# def add_chart_slide(pres, report_data, image_file):
#     """
#     Adds a slide with a chart image, resizing it to fit a larger bounding box
#     while preserving aspect ratio and centering it.
#     """
#     slide_layout = pres.slide_layouts[6] # Blank layout
#     slide = pres.slides.add_slide(slide_layout)

#     if "HEADER_MASTER" in report_data.master:
#         add_custom_header(slide, report_data.tab)
#     else:
#         add_default_header(slide)

#     title_box = slide.shapes.add_textbox(Inches(0.5), Inches(0.9), Inches(10), Inches(0.5))
#     p_title = title_box.text_frame.paragraphs[0]
#     p_title.text = report_data.title
#     p_title.font.bold = True
#     p_title.font.size = Pt(22)
#     p_title.font.color.rgb = RGBColor(0xF5, 0x82, 0x20)

#     key_msg_box = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.5), Inches(1.5), Inches(12.3), Inches(0.5))
#     key_msg_box.fill.background()
#     key_msg_box.line.color.rgb = RGBColor(211, 211, 211)
#     key_msg_box.text_frame.text = "▪ [Type key message]"

#     # --- Bounding Box and Image Placement Logic ---
#     content_area_left = Inches(0.5)
#     content_area_top = Inches(2.1) # Moved content area up slightly
#     content_area_max_width = Inches(12.3)
#     content_area_max_height = Inches(5.3) # Increased max height to make images larger

#     # Add the picture with a placeholder width to get its native dimensions scaled.
#     pic = slide.shapes.add_picture(image_file, content_area_left, content_area_top, width=content_area_max_width)

#     # Calculate the aspect ratios of the image and the bounding box.
#     image_ratio = pic.width / pic.height
#     box_ratio = content_area_max_width / content_area_max_height

#     # Resize the image to fit the bounding box without distortion.
#     if image_ratio > box_ratio:
#         # Image is wider than the box: fit to width.
#         pic.width = content_area_max_width
#         pic.height = int(content_area_max_width / image_ratio)
#     else:
#         # Image is taller than or equal to the box: fit to height.
#         pic.height = content_area_max_height
#         pic.width = int(content_area_max_height * image_ratio)

#     # Center the resized image within the content area.
#     pic.left = int(content_area_left + (content_area_max_width - pic.width) / 2)
#     pic.top = int(content_area_top + (content_area_max_height - pic.height) / 2)


# # --- API Endpoint ---

# @app.post("/generate-ppt")
# async def generate_ppt_endpoint(request_data: GenerationRequest):
#     """
#     This endpoint receives report selections and chart images, then
#     assembles and returns a PowerPoint presentation.
#     """
#     pres = Presentation()
#     pres.slide_width = Inches(13.33)
#     pres.slide_height = Inches(7.5)

#     add_title_slide(pres)
#     add_contents_slide(pres, request_data.selectedReports)

#     for report in request_data.selectedReports:
#         image_data_url = request_data.imageDataUrls.get(report.id)
#         if image_data_url:
#             # The data URL is in the format "data:image/png;base64,ENCODED_STRING"
#             # We need to split it and decode the base64 part.
#             try:
#                 header, encoded = image_data_url.split(",", 1)
#                 image_data = io.BytesIO(base64.b64decode(encoded))
#                 add_chart_slide(pres, report, image_data)
#             except Exception as e:
#                 print(f"Could not process image for report {report.id}: {e}")
#                 # Optionally, add an error slide for this report
#                 continue

#     # Save the presentation to an in-memory buffer
#     ppt_buffer = io.BytesIO()
#     pres.save(ppt_buffer)
#     ppt_buffer.seek(0)

#     # Stream the buffer back to the client as a .pptx file download
#     return StreamingResponse(
#         ppt_buffer,
#         media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
#         headers={'Content-Disposition': 'attachment; filename=Dashboard_Report_Final.pptx'}
#     )

# # To run this app from your terminal:
# # 1. Make sure you have an 'assets' folder with the required images.
# # 2. Run the command: uvicorn app:app --reload --port 5001
