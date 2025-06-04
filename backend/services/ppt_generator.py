from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData, ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_LABEL_POSITION, XL_DATA_LABEL_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from services.excel_parser import get_sentiment_counts, get_company_sentiment_counts
import logging
import pandas as pd
import os
from PIL import Image

logger = logging.getLogger(__name__)

# Define colors
SLIDE_BG_COLOR = RGBColor(250, 250, 250)  # Light gray
CHART_BG_COLOR = RGBColor(255, 255, 255)  # White
HEADER_TEXT_COLOR = RGBColor(214, 55, 64)    # Red
HEADER_HEIGHT = 0.8

# Define sentiment colors
SENTIMENT_COLORS = {
    1: RGBColor(69, 194, 126),    # Positive - Green
    0: RGBColor(255, 191, 0),     # Neutral - Yellow
    -1: RGBColor(246, 1, 64)      # Negative - Red
}

# Define metric icons
METRIC_ICONS = {
    'Posts Count': '📝',
    'Total Comments': '💬',
    'Total Likes': '👍',
    'Total Shares': '🔄',
    'Total Views': '👁️'
}

def format_chart_axes(chart, category_font_size=6, value_font_size=6, rotation_angle=-45):
    """Helper function to consistently format chart axes"""
    if hasattr(chart, 'category_axis'):
        chart.category_axis.tick_labels.font.size = Pt(category_font_size)
        chart.category_axis.tick_labels.font.bold = False
        if hasattr(chart.category_axis, 'axis_title') and chart.category_axis.axis_title:
            chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(category_font_size)
        # Set labels at specified angle
        chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        chart.category_axis.tick_label_rotation = rotation_angle
        chart.category_axis.tick_labels.font.color.rgb = RGBColor(89, 89, 89)
        chart.category_axis.has_major_gridlines = False
    
    if hasattr(chart, 'value_axis'):
        chart.value_axis.tick_labels.font.size = Pt(value_font_size)
        chart.value_axis.tick_labels.font.bold = False
        if hasattr(chart.value_axis, 'axis_title') and chart.value_axis.axis_title:
            chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(value_font_size)
        chart.value_axis.tick_labels.font.color.rgb = RGBColor(89, 89, 89)
        chart.value_axis.has_major_gridlines = False

def add_slide_header(slide, company_logo_path, start_date, end_date, title):
    """Helper function to add a consistent header to slides"""

    # Set constants
    SLIDE_WIDTH_INCHES = 13.33

    # Add header background with SLIDE_BG_COLOR
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        Inches(SLIDE_WIDTH_INCHES), Inches(HEADER_HEIGHT)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = SLIDE_BG_COLOR
    header.line.fill.background()  # Remove border

    # Add company logo (a bit lower)
    if os.path.exists(company_logo_path):
        slide.shapes.add_picture(
            company_logo_path,
            Inches(0.2), Inches(0.2),  # moved slightly lower
            height=Inches(0.4)
        )

    # Add title in the center
    title_box = slide.shapes.add_textbox(
        Inches((SLIDE_WIDTH_INCHES - 6) / 2), Inches(0.2),
        Inches(6), Inches(0.6)
    )
    tf = title_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"📊 {title}"
    p.font.size = Pt(16)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    p.font.color.rgb = HEADER_TEXT_COLOR

    # Add date on the far right
    date_box = slide.shapes.add_textbox(
        Inches(SLIDE_WIDTH_INCHES - 3), Inches(0.2),
        Inches(2.8), Inches(0.6)
    )
    tf = date_box.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = f"📅 {start_date} - {end_date}"
    p.font.size = Pt(12)
    p.alignment = PP_ALIGN.RIGHT
    p.font.color.rgb = HEADER_TEXT_COLOR

    # Add divider line
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(HEADER_HEIGHT),
        Inches(SLIDE_WIDTH_INCHES), Inches(0.02)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = HEADER_TEXT_COLOR
    line.line.fill.background()  # Remove border


# Function to apply consistent chart formatting
def apply_chart_formatting(chart, use_legend=True, legend_position=XL_LEGEND_POSITION.TOP, 
                         category_font_size=6, value_font_size=6, 
                         title=None, title_size=14):
    """Helper function for consistent chart formatting across slides"""
    
    # Basic chart settings
    chart.chart_style = 2  # White background
    chart.plots[0].has_major_gridlines = False
    chart.plots[0].has_minor_gridlines = False
    
    # Legend settings
    chart.has_legend = use_legend
    if use_legend:
        chart.legend.position = legend_position
        chart.legend.font.size = Pt(12)
        
    # Title settings
    if title:
        chart.has_title = True
        chart.chart_title.text_frame.text = f"📊 {title}"
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(title_size)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        chart.chart_title.text_frame.paragraphs[0].font.color.rgb = HEADER_TEXT_COLOR
        
    # Axis formatting
    if hasattr(chart, 'category_axis'):
        chart.category_axis.tick_labels.font.size = Pt(category_font_size)
        chart.category_axis.tick_labels.font.bold = False
        chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        chart.category_axis.tick_label_rotation = -45  # 45-degree angle
        chart.category_axis.tick_labels.font.color.rgb = RGBColor(89, 89, 89)
        chart.category_axis.has_major_gridlines = False
        
    if hasattr(chart, 'value_axis'):
        chart.value_axis.tick_labels.font.size = Pt(value_font_size)
        chart.value_axis.tick_labels.font.bold = False
        chart.value_axis.tick_labels.font.color.rgb = RGBColor(89, 89, 89)
        chart.value_axis.has_major_gridlines = False

def apply_sentiment_colors(chart):
    """Helper function to apply consistent sentiment colors to chart series"""
    for i, series in enumerate(chart.series):
        series.format.fill.solid()
        # The series order should match [1, 0, -1] for [Positive, Neutral, Negative]
        sentiment_value = [1, 0, -1][i]
        series.format.fill.fore_color.rgb = SENTIMENT_COLORS[sentiment_value]

def create_ppt(data_frames, output_path, start_date, end_date, company_name, company_logo_path, mediaeye_logo_path, neurotime_logo_path, competitor_logo_paths=None, positive_links=None, negative_links=None, positive_posts=None, negative_posts=None):
    try:
        logger.debug("Creating PowerPoint presentation")
        prs = Presentation()
        
        # Set slide dimensions to landscape (16:9)
        prs.slide_width = Inches(13.33)
        prs.slide_height = Inches(7.5)
        
        # region  First slide - Title slide with logos and text
        logger.debug("Creating title slide")
        title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank layout

        # Constants
        SLIDE_WIDTH = 13.33
        LEFT_HALF_CENTER = Inches(SLIDE_WIDTH / 4)        # ~3.33"
        RIGHT_HALF_CENTER = Inches(3 * SLIDE_WIDTH / 4)   # ~10"

        LEFT_CONTENT_WIDTH = Inches(3.5)
        LOGO_WIDTH = Inches(5)
        SPACING_SMALL = Inches(0.1)
        SPACING_MEDIUM = Inches(0.2)

        # Remove default textboxes
        for shape in title_slide.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)

        # Background color
        background = title_slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR

        # Top offset
        top = Inches(1.0)

        # Calculate centered left margin
        LEFT_MARGIN = LEFT_HALF_CENTER - LEFT_CONTENT_WIDTH / 2

        # NeuroTime logo (top left half)
        if os.path.exists(neurotime_logo_path):
            title_slide.shapes.add_picture(neurotime_logo_path, LEFT_MARGIN, top, width=Inches(4))
            top += Inches(1.5)

        # Main Title
        title_box = title_slide.shapes.add_textbox(LEFT_MARGIN, top, LEFT_CONTENT_WIDTH, Inches(1.2))
        tf = title_box.text_frame
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
        p = tf.paragraphs[0]
        p.text = "MARKA VARLIĞININ\nMEDİA ÜZƏRİNDƏN\nQİYMƏTLƏNDİRİLMƏSİ"
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 123, 191)
        p.line_spacing = Pt(24)
        top += Inches(1.2)

        # Divider Line
        line = title_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT_MARGIN, top, LEFT_CONTENT_WIDTH, Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0, 123, 191)
        line.line.fill.background()
        top += SPACING_SMALL

        # Description
        desc_box = title_slide.shapes.add_textbox(LEFT_MARGIN, top, LEFT_CONTENT_WIDTH, Inches(0.6))
        tf = desc_box.text_frame
        tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]
        p.text = "Online/Sosial və ənənəvi media\ndatalarının əsasında təhlil."
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(0, 123, 191)
        p.line_spacing = Pt(18)
        top += Inches(0.6)

        # Divider Line
        line = title_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, LEFT_MARGIN, top, LEFT_CONTENT_WIDTH, Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0, 123, 191)
        line.line.fill.background()
        top += SPACING_MEDIUM

        # "Analitik hesabat"
        report_box = title_slide.shapes.add_textbox(LEFT_MARGIN, top, LEFT_CONTENT_WIDTH, Inches(0.3))
        tf = report_box.text_frame
        tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]
        p.text = "Analitik hesabat"
        p.font.size = Pt(14)
        p.font.color.rgb = RGBColor(0, 123, 191)
        top += Inches(0.3)

        # Date
        date_box = title_slide.shapes.add_textbox(LEFT_MARGIN, top, LEFT_CONTENT_WIDTH, Inches(0.3))
        tf = date_box.text_frame
        tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]
        p.text = f"{start_date}-{end_date}"
        p.font.size = Pt(12)
        p.font.color.rgb = RGBColor(0, 123, 191)
        top += Inches(1.3)

        # --- RIGHT Side Company Logo ---

        if os.path.exists(company_logo_path):
            logo_top = Inches(2.5)
            right_logo_left = RIGHT_HALF_CENTER - LOGO_WIDTH / 2
            title_slide.shapes.add_picture(
                company_logo_path,
                right_logo_left, logo_top,
                width=LOGO_WIDTH
            )

        # endregion

        # region Second slide - Methodology description
        logger.debug("Creating methodology slide")
        method_slide = prs.slides.add_slide(prs.slide_layouts[5])

        # Remove default textbox
        for shape in method_slide.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)

        SLIDE_WIDTH = prs.slide_width
        SLIDE_HEIGHT = prs.slide_height
        HALF_WIDTH = SLIDE_WIDTH / 2

        # Background color
        background = method_slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR

        # --- Title on top-left ---
        title_box = method_slide.shapes.add_textbox(Inches(0.3), Inches(0.2), Inches(5), Inches(0.5))
        tf = title_box.text_frame
        p = tf.paragraphs[0]
        p.text = "Metodologiyanın təsviri"
        p.font.size = Pt(20)
        p.font.bold = True
        p.font.color.rgb = RGBColor(0, 123, 191)

        # --- NeuroTime logo (top-right corner) ---
        if os.path.exists(neurotime_logo_path):
            method_slide.shapes.add_picture(
                neurotime_logo_path,
                SLIDE_WIDTH - Inches(1.5), Inches(0.3),
                width=Inches(1.2)
            )

        # --- MediaEye logo (centered in left half) ---
        if os.path.exists(mediaeye_logo_path):
            max_logo_width = Inches(4.5)
            max_logo_height = Inches(4.5)

            with Image.open(mediaeye_logo_path) as img:
                width, height = img.size
                aspect_ratio = width / height

                if aspect_ratio > 1:  # Wider than tall
                    display_width = max_logo_width
                    display_height = max_logo_width / aspect_ratio
                else:
                    display_height = max_logo_height
                    display_width = max_logo_height * aspect_ratio

            logo_left = (HALF_WIDTH - display_width) / 2
            logo_top = (SLIDE_HEIGHT - display_height) / 2

            method_slide.shapes.add_picture(
                mediaeye_logo_path,
                logo_left, logo_top,
                width=display_width,
                height=display_height
            )

        # --- Competitor logos grid (up to 25, in right half, centered) ---
        if competitor_logo_paths:
            max_logos = min(25, len(competitor_logo_paths))
            logos_per_row = 5
            logos_per_col = 5

            max_logo_width = Inches(1.1)
            max_logo_height = Inches(0.5)
            spacing = Inches(0.15)

            grid_width = logos_per_row * max_logo_width + (logos_per_row - 1) * spacing
            grid_height = logos_per_col * max_logo_height + (logos_per_col - 1) * spacing

            grid_left_start = HALF_WIDTH + (HALF_WIDTH - grid_width) / 2 - Inches(1)
            grid_top_start = (SLIDE_HEIGHT - grid_height) / 2

            for i, logo_path in enumerate(competitor_logo_paths[:max_logos]):
                if os.path.exists(logo_path):
                    row = i // logos_per_row
                    col = i % logos_per_row
                    left = grid_left_start + col * (max_logo_width + spacing)
                    top = grid_top_start + row * (max_logo_height + spacing)

                    with Image.open(logo_path) as img:
                        width, height = img.size
                        aspect_ratio = width / height

                        if aspect_ratio > 1:
                            display_width = max_logo_width
                            display_height = max_logo_width / aspect_ratio
                        else:
                            display_height = max_logo_height
                            display_width = max_logo_height * aspect_ratio

                    method_slide.shapes.add_picture(
                        logo_path,
                        left + (max_logo_width - display_width) / 2,
                        top + (max_logo_height - display_height) / 2,
                        width=display_width,
                        height=display_height
                    )
        # endregion

        # region Third slide - Grid layout with links and charts
        logger.debug("Creating third slide with grid layout")
        slide3 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Remove default textbox
        for shape in slide3.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)
        
        add_slide_header(slide3, company_logo_path, start_date, end_date, "Sentiment Analysis")
        
        # Set slide background
        background = slide3.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR

        # Grid layout positions
        top_left = (Inches(0.5), Inches(1.2))
        top_right = (Inches(5.5), Inches(1.2))
        bottom_left = (Inches(0.5), Inches(3.5))
        bottom_right = (Inches(5.5), Inches(3.5))
        
        # Add positive links section (top left)
        if positive_links:
            left, top = top_left
            width = Inches(4.5)
            height = Inches(0.4)
            
            # Add title
            title_box = slide3.shapes.add_textbox(left, top, width, Inches(0.3))
            tf = title_box.text_frame
            p = tf.add_paragraph()
            p.text = "Positive Links"
            p.font.size = Pt(12)
            p.font.bold = True
            p.alignment = PP_ALIGN.LEFT
            top += Inches(0.4)
            
            for i, link in enumerate(positive_links):
                link_box = slide3.shapes.add_textbox(left, top + (height * i), width, height)
                tf = link_box.text_frame
                p = tf.add_paragraph()
                p.text = f"🔗 {link}"
                p.font.size = Pt(8)
                p.font.color.rgb = RGBColor(0, 112, 192)  # Blue color
                p.alignment = PP_ALIGN.LEFT
                r = p.runs[0]
                r.hyperlink.address = link

        # Add negative links section (top right)
        if negative_links:
            left, top = top_right
            width = Inches(4.5)
            height = Inches(0.4)
            
            # Add title
            title_box = slide3.shapes.add_textbox(left, top, width, Inches(0.3))
            tf = title_box.text_frame
            p = tf.add_paragraph()
            p.text = "Negative Links"
            p.font.size = Pt(12)
            p.font.bold = True
            p.alignment = PP_ALIGN.LEFT
            top += Inches(0.4)
            
            for i, link in enumerate(negative_links):
                link_box = slide3.shapes.add_textbox(left, top + (height * i), width, height)
                tf = link_box.text_frame
                p = tf.add_paragraph()
                p.text = f"🔗 {link}"
                p.font.size = Pt(8)
                p.font.color.rgb = RGBColor(246, 1, 64)  # Red color
                p.alignment = PP_ALIGN.LEFT
                r = p.runs[0]
                r.hyperlink.address = link

        # Get data from combined_sources
        logger.debug("Processing combined sources data")
        if 'combined_sources' not in data_frames:
            raise ValueError("combined_sources Excel file is missing")
        
        if 'News' not in data_frames['combined_sources']:
            raise ValueError("News sheet is missing in combined_sources Excel file")
        
        combined_data = data_frames['combined_sources']['News']
        combined_data['Day'] = pd.to_datetime(combined_data['Day'])
        company_data = combined_data[combined_data['Company'] == company_name]
        sentiment_data = company_data.groupby(['Day', 'Sentiment']).size().unstack(fill_value=0)
        sentiment_data = sentiment_data.sort_index()
        sentiment_counts = get_sentiment_counts(combined_data)
        
        # Add chart titles
        title_box = slide3.shapes.add_textbox(Inches(0.5), Inches(1.3), Inches(4.5), Inches(0.3))
        tf = title_box.text_frame
        p = tf.add_paragraph()
        p.text = "Sentiment Trend Over Time"
        p.font.size = Pt(14)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER

        title_box = slide3.shapes.add_textbox(Inches(5.5), Inches(1.3), Inches(4.5), Inches(0.3))
        tf = title_box.text_frame
        p = tf.add_paragraph()
        p.text = "Overall Sentiment Distribution"
        p.font.size = Pt(14)
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        
        # Create multiline chart
        logger.debug("Creating multiline chart")
        chart_data = CategoryChartData()
        chart_data.categories = sentiment_data.index.tolist()
        
        for sentiment in [1, 0, -1]:
            series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
            if sentiment in sentiment_data.columns:
                chart_data.add_series(series_name, sentiment_data[sentiment].tolist())
        
        left, top = bottom_left
        x, y, cx, cy = left, top, Inches(4.5), Inches(3)
        chart = slide3.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.font.size = Pt(12)
        
        # Set chart background and formatting
        apply_chart_formatting(chart, title="Sentiment Trend Over Time")
        for i, series in enumerate(chart.series):
            series.format.line.color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
            series.format.line.width = Pt(2)
        
        # Create donut chart (bottom right)
        logger.debug("Creating donut chart")
        donut_data = ChartData()
        donut_data.categories = ['Positive', 'Neutral', 'Negative']
        donut_data.add_series('', [  # Empty series name to remove text
            sentiment_counts.get(1, 0),
            sentiment_counts.get(0, 0),
            sentiment_counts.get(-1, 0)
        ])
        
        left, top = bottom_right
        x, y, cx, cy = left, top, Inches(4.5), Inches(3)
        donut = slide3.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, donut_data).chart
        donut.has_legend = True
        donut.legend.position = XL_LEGEND_POSITION.TOP
        donut.legend.font.size = Pt(10)
        
        # Set chart background
        donut.chart_style = 2  # White background
        
        # Set colors for donut chart and add data labels with arrows
        for i, point in enumerate(donut.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
            # Add data label with arrow
            point.has_data_label = True
            point.data_label.font.size = Pt(10)
            point.data_label.font.bold = True
            point.data_label.number_format = '#,##0'
        # endregion

        # region Fourth slide - Vertical multibar chart
        logger.debug("Creating Fourth slide with vertical multibar chart")
        slide4 = prs.slides.add_slide(prs.slide_layouts[5])
        # Remove default textbox
        for shape in slide4.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)

        add_slide_header(slide4, company_logo_path, start_date, end_date, "Xəbərlərin analizi")

        # Set slide background
        background = slide4.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        # Get company sentiment data and sort it by total values
        company_sentiments = get_company_sentiment_counts(combined_data)
        # Calculate total posts for each company
        totals = company_sentiments.sum(axis=1)
        # Sort the sentiment data based on totals
        company_sentiments = company_sentiments.loc[totals.sort_values(ascending=False).index]
        
        chart_data = CategoryChartData()
        chart_data.categories = company_sentiments.index.tolist()
        
        for sentiment in [1, 0, -1]:
            series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
            if sentiment in company_sentiments.columns:
                chart_data.add_series(series_name, company_sentiments[sentiment].tolist())
        
        # Calculate centered position
        # Slide width: 13.33", Chart width: 11"
        # Center horizontally: (13.33 - 11) / 2 = 1.165"
        # Slide height: 7.5", Header height: 0.8", Chart height: 5"
        # Center vertically in remaining space: 0.8 + (6.7 - 5) / 2 = 1.65"
        x = Inches(1.165)
        y = Inches(1.65)
        cx = Inches(11)  # Chart width
        cy = Inches(5)   # Chart height
        
        chart = slide4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
        
        # Configure legend
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.TOP
        chart.legend.include_in_layout = True
        chart.legend.font.size = Pt(12)
        chart.legend.font.bold = True
        
        # Set chart title with icon
        chart.has_title = True
        chart.chart_title.text_frame.text = "📊 Post saylarına görə bankların bölgüsü"
        chart.chart_title.text_frame.paragraphs[0].font.size = Pt(14)
        chart.chart_title.text_frame.paragraphs[0].font.bold = True
        chart.chart_title.text_frame.paragraphs[0].font.color.rgb = HEADER_TEXT_COLOR  # Red color for title
        
        # Set white background for chart
        chart.chart_style = 2  # White background style
        
        # Apply fill to plot area if available and remove gridlines
        try:
            if hasattr(chart, 'plot_area'):
                chart.plot_area.format.fill.solid()
                chart.plot_area.format.fill.fore_color.rgb = RGBColor(255, 255, 255)
                
            # Remove all gridlines
            # Remove major and minor gridlines from the plot
            chart.plots[0].has_major_gridlines = False
            chart.plots[0].has_minor_gridlines = False
            
            # Also remove gridlines from both axes
            if hasattr(chart, 'category_axis') and chart.category_axis:
                chart.category_axis.has_major_gridlines = False
                chart.category_axis.has_minor_gridlines = False
            
            if hasattr(chart, 'value_axis') and chart.value_axis:
                chart.value_axis.has_major_gridlines = False
                chart.value_axis.has_minor_gridlines = False
            
            # Remove axis lines
            chart.category_axis.format.line.fill.background()
            chart.value_axis.format.line.fill.background()
        except Exception as e:
            logger.warning(f"Could not set plot area formatting: {e}")
        
        # Set diagonal labels for x-axis
        if hasattr(chart, 'category_axis'):
            chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
            chart.category_axis.text_rotation = 90
        
        # Set axis titles and labels
        if hasattr(chart, 'category_axis'):
            chart.category_axis.tick_labels.font.size = Pt(7)
            if hasattr(chart.category_axis, 'axis_title') and chart.category_axis.axis_title:
                chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(8)
        if hasattr(chart, 'value_axis'):
            chart.value_axis.tick_labels.font.size = Pt(7)
            if hasattr(chart.value_axis, 'axis_title') and chart.value_axis.axis_title:
                chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(8)
        
        # Set colors for bar chart
        for i, series in enumerate(chart.series):
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
        # endregion

        # region Fifth slide - Author count horizontal bar chart
        logger.debug("Creating fifth slide with author count chart")
        slide5 = prs.slides.add_slide(prs.slide_layouts[5])
        # Remove default textbox
        for shape in slide5.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)

        add_slide_header(slide5, company_logo_path, start_date, end_date, "Xəbər saylarının saytlar üzərindən bölgüsü")
        # Set slide background
        background = slide5.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR

        # Filter and group data by author
        author_data = combined_data[combined_data['Company'] == company_name].groupby('Author').size().sort_values(ascending=True)

        # Ensure we have data to prevent errors
        if len(author_data) == 0:
            logger.warning("No author data found for chart")
            # Add a text placeholder or skip chart creation
        else:
            chart_data = CategoryChartData()
            chart_data.categories = author_data.index.tolist()
            chart_data.add_series('', author_data.values.tolist())  # Empty series name to remove title
            
            # Calculate centered position
            # Slide dimensions: 13.33" x 7.5"
            # Header height: 0.8"
            # Available space: 7.5 - 0.8 = 6.7"
            chart_width = Inches(11.5)  # Increased width further
            chart_height = Inches(6.0)  # Increased height further
            
            # Center horizontally: (13.33 - 11.5) / 2 = 0.915
            # Center vertically in remaining space: 0.8 + (6.7 - 6.0) / 2 = 1.15
            x = Inches(0.9)
            y = Inches(1.15)
            
            chart = slide5.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, chart_width, chart_height, chart_data).chart
            
            # Remove legend
            chart.has_legend = False
            
            # Set white background for chart area and plot area
            try:
                # Chart area background
                chart.chart_area.fill.solid()
                chart.chart_area.fill.fore_color.rgb = RGBColor(255, 255, 255)
                
                # Plot area background - solid white fill
                try:
                    chart.plot_area.format.fill.solid()
                    chart.plot_area.format.fill.fore_color.rgb = RGBColor(255, 255, 255)
                except Exception as e:
                    logger.warning(f"Could not set plot area fill color: {e}")

            except Exception as e:
                logger.warning(f"Could not set chart background: {e}")
            
            # Remove all gridlines
            try:
                # Remove major and minor gridlines from the plot
                chart.plots[0].has_major_gridlines = False
                chart.plots[0].has_minor_gridlines = False
                
                # Also remove gridlines from both axes
                if hasattr(chart, 'category_axis') and chart.category_axis:
                    chart.category_axis.has_major_gridlines = False
                    chart.category_axis.has_minor_gridlines = False
                
                if hasattr(chart, 'value_axis') and chart.value_axis:
                    chart.value_axis.has_major_gridlines = False
                    chart.value_axis.has_minor_gridlines = False
                    
            except Exception as e:
                logger.warning(f"Could not remove gridlines: {e}")
            
            # Configure category axis (Y-axis for horizontal bar chart) - keep only this axis
            try:
                if hasattr(chart, 'category_axis') and chart.category_axis:
                    chart.category_axis.tick_labels.font.size = Pt(9)
                    chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW  # Keep Y-axis labels on left
                    # Remove axis title
                    chart.category_axis.has_title = False
                    # Keep the axis line visible
                    chart.category_axis.format.line.color.rgb = RGBColor(0, 0, 0)  # Black axis line
            except Exception as e:
                logger.warning(f"Could not configure category axis: {e}")
            
            # Remove/hide value axis (X-axis for horizontal bar chart)
            try:
                if hasattr(chart, 'value_axis') and chart.value_axis:
                    # Hide the axis completely
                    chart.value_axis.visible = False
                    # Alternative: make axis line transparent
                    chart.value_axis.format.line.fill.background()
                    # Remove axis title
                    chart.value_axis.has_title = False
                    # Hide tick labels
                    chart.value_axis.tick_labels.font.size = Pt(1)
                    chart.value_axis.tick_labels.font.color.rgb = RGBColor(255, 255, 255)  # Make invisible
            except Exception as e:
                logger.warning(f"Could not configure value axis: {e}")
            
            # Format bars and add data labels
            try:
                for series in chart.series:
                    # Set bar color
                    series.format.fill.solid()
                    series.format.fill.fore_color.rgb = RGBColor(160, 208, 255)  # Blue color
                    
                    # Make bars thick but add space between them
                    series.format.gap_width = 40  # Thick bars with some spacing
                    
                    # Remove line around bars to prevent errors
                    series.format.line.fill.background()
                    
                    # Add data labels to outside end
                    series.has_data_labels = True
                    data_labels = series.data_labels
                    data_labels.position = XL_DATA_LABEL_POSITION.OUTSIDE_END
                    data_labels.font.size = Pt(10)
                    data_labels.font.color.rgb = RGBColor(0, 0, 0)  # Black text
                    data_labels.font.bold = True
                    data_labels.number_format = '0'  # Show whole numbers only
                    
            except Exception as e:
                logger.warning(f"Could not format series or data labels: {e}")
            
            # Additional formatting for bars spacing
            try:
                # Access the plot area and modify bar thickness
                plot = chart.plots[0]
                plot.gap_width = 40  # Set plot-level gap width for spacing
            except Exception as e:
                logger.warning(f"Could not set plot formatting: {e}")
            
            # Remove chart title if it exists
            try:
                if chart.has_title:
                    chart.chart_title.text_frame.text = ''
                    chart.chart_title.include_in_layout = False
            except Exception as e:
                logger.warning(f"Could not remove chart title: {e}")

        # endregion

        # region Sixth slide - Facebook metrics and sentiment analysis
        logger.debug("Creating Sixth slide with Facebook metrics")
        slide6 = prs.slides.add_slide(prs.slide_layouts[5])
        # Remove default textbox
        for shape in slide6.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)
        add_slide_header(slide6, company_logo_path, start_date, end_date, "Facebook postlarının analizi")

        # Set slide background
        background = slide6.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        # Left side - Facebook metrics
        if 'official_facebook' in data_frames and len(data_frames['official_facebook']) > 0:
            fb_data = list(data_frames['official_facebook'].values())[0]
            fb_metrics = fb_data[fb_data['author_name'] == company_name]
            
            metrics = {
                'Posts Count': len(fb_metrics),
                'Total Comments': fb_metrics['comment_count'].sum(),
                'Total Likes': fb_metrics['like_count'].sum(),
                'Total Shares': fb_metrics['share_count'].sum(),
                'Total Views': fb_metrics['view_count'].sum()
            }
            
            # Set up left section (20% width) for metrics
            left = Inches(0.5)
            top = Inches(1)
            metrics_width = Inches(2.7)  # ~20% of slide width (13.33")
            metrics_height = Inches(1.2)  # Taller boxes for new layout
            available_height = Inches(6.5)  # Total height minus header
            total_items = len(metrics)
            spacing = (available_height - (metrics_height * total_items)) / (total_items + 1)

            for i, (metric, value) in enumerate(metrics.items()):
                # Calculate position for current metric box
                box_top = top + (spacing * (i + 1)) + (metrics_height * i)
                
                # Add white background box for the whole metric
                bg_box = slide6.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    left, box_top, metrics_width, metrics_height
                )
                bg_box.fill.solid()
                bg_box.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
                
                # Add red background for icon
                icon_width = Inches(1.0)
                icon_box = slide6.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    left, box_top, icon_width, metrics_height
                )
                icon_box.fill.solid()
                icon_box.fill.fore_color.rgb = HEADER_TEXT_COLOR  # Red background
                
                # Add icon text
                icon_text = slide6.shapes.add_textbox(
                    left, box_top + Inches(0.3), icon_width, Inches(0.6)
                )
                icon_p = icon_text.text_frame.add_paragraph()
                icon_p.text = "📊"
                icon_p.alignment = PP_ALIGN.CENTER
                icon_p.font.size = Pt(24)
                
                # Add value in large text
                value_text = slide6.shapes.add_textbox(
                    left + icon_width + Inches(0.1),
                    box_top + Inches(0.1),
                    metrics_width - icon_width - Inches(0.2),
                    Inches(0.6)
                )
                value_p = value_text.text_frame.add_paragraph()
                value_p.text = f"{value:,}"
                value_p.alignment = PP_ALIGN.LEFT
                value_p.font.size = Pt(20)
                value_p.font.bold = True
                
                # Add metric name below value
                metric_text = slide6.shapes.add_textbox(
                    left + icon_width + Inches(0.1),
                    box_top + Inches(0.7),
                    metrics_width - icon_width - Inches(0.2),
                    Inches(0.4)
                )
                metric_p = metric_text.text_frame.add_paragraph()
                metric_p.text = metric
                metric_p.alignment = PP_ALIGN.LEFT
                metric_p.font.size = Pt(12)

        # Right side - Sentiment analysis
        if 'Facebook' in data_frames['combined_sources']:
            fb_sentiment_data = data_frames['combined_sources']['Facebook']
            company_sentiment = fb_sentiment_data[fb_sentiment_data['Company'] == company_name]
            
            # Donut chart
            sentiment_counts = company_sentiment['Sentiment'].value_counts()
            donut_data = ChartData()
            donut_data.categories = ['Positive', 'Neutral', 'Negative']
            donut_data.add_series('', [
                sentiment_counts.get(1, 0),
                sentiment_counts.get(0, 0),
                sentiment_counts.get(-1, 0)
            ])
            
            # Start right section (80% width) layout
            right_section_left = Inches(3.7)  # After left 20% section
            right_width = Inches(9.13)  # ~80% of slide width
            
            # Position donut chart in upper left of right section
            donut_size = Inches(3.5)
            x = right_section_left
            y = Inches(1)
            
            # Add title for donut chart
            donut_title = slide6.shapes.add_textbox(
                x, y, donut_size, Inches(0.3)
            )
            title_p = donut_title.text_frame.add_paragraph()
            title_p.text = "📊 Ümumi sentiment bölgüsü"
            title_p.font.size = Pt(14)
            title_p.font.bold = True
            title_p.font.color.rgb = HEADER_TEXT_COLOR
            title_p.alignment = PP_ALIGN.CENTER
            
            # Position donut chart below title
            donut = slide6.shapes.add_chart(
                XL_CHART_TYPE.DOUGHNUT, 
                x, y + Inches(0.4), 
                donut_size, donut_size - Inches(0.4), 
                donut_data
            ).chart
            donut.has_legend = True
            donut.legend.position = XL_LEGEND_POSITION.TOP
            donut.legend.font.size = Pt(8)  # Reduced legend size

            # Set chart background
            donut.chart_style = 2  # White background
            
            # Set colors for donut chart and add data labels with arrows
            for i, point in enumerate(donut.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                # Add data label with arrow
                point.has_data_label = True
                point.data_label.font.size = Pt(10)
                point.data_label.font.bold = True
                point.data_label.number_format = '#,##0'
                
            # Multiline chart
            sentiment_by_date = company_sentiment.groupby('Day')['Sentiment'].value_counts().unstack(fill_value=0)
            chart_data = CategoryChartData()
            chart_data.categories = sentiment_by_date.index.tolist()
            
            for sentiment in [1, 0, -1]:
                series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                if sentiment in sentiment_by_date.columns:
                    chart_data.add_series(series_name, sentiment_by_date[sentiment].tolist())
            
            # Position multiline chart to right of donut
            x = right_section_left + donut_size + Inches(0.5)  # After donut chart
            y = Inches(1.25)
            cx = right_width - donut_size - Inches(0.3)  # Remaining width
            cy = donut_size - Inches(0.25)  # Same height as donut
            chart = slide6.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.TOP
            chart.legend.font.size = Pt(12)
            chart.has_title = True
            chart.chart_title.text_frame.text = "📊 Postların zamana və sentimentə görə bölgüsü"
            paragraph = chart.chart_title.text_frame.paragraphs[0]
            paragraph.font.size = Pt(14)
            paragraph.font.bold = True
            paragraph.font.color.rgb = HEADER_TEXT_COLOR
            # Set chart background and formatting
            apply_chart_formatting(chart, title="Postların zamana və sentimentə görə bölgüsü")
            for i, series in enumerate(chart.series):
                series.format.line.color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                series.format.line.width = Pt(2)
            
            # Vertical multibar chart
            company_sentiments = fb_sentiment_data.groupby('Company')['Sentiment'].value_counts().unstack(fill_value=0)
            # Sort by total sentiment values
            totals = company_sentiments.sum(axis=1)
            company_sentiments = company_sentiments.loc[totals.sort_values(ascending=False).index]

            chart_data = CategoryChartData()
            chart_data.categories = company_sentiments.index.tolist()
            
            for sentiment in [1, 0, -1]:
                series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                if sentiment in company_sentiments.columns:
                    chart_data.add_series(series_name, company_sentiments[sentiment].tolist())
            
            # Position multibar chart to span full width of right section
            x = right_section_left
            y = Inches(4.5)  # Below upper charts
            cx = right_width  # Full width of right section
            cy = Inches(3)  # Remaining height
            chart = slide6.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.TOP
            chart.legend.font.size = Pt(12)

            # Set chart background and formatting
            apply_chart_formatting(chart, title="Post saylarına görə bankların bölgüsü")
            apply_sentiment_colors(chart)
        # endregion

        # region Seventh slide - Facebook metrics table Combines sources
        logger.debug("Creating eighth slide with Facebook metrics table by company")
        slide7 = prs.slides.add_slide(prs.slide_layouts[5])
        # Remove default textbox
        for shape in slide7.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)
        add_slide_header(slide7, company_logo_path, start_date, end_date, "Banklar haqqında paylaşılan postların analizi")

        # Set slide background
        background = slide7.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        if 'combined_sources' in data_frames and 'Facebook' in data_frames['combined_sources']:
            fb_data = data_frames['combined_sources']['Facebook']
            
            # Group data by Company
            grouped_data = fb_data.groupby('Company').agg({
                'comment_count': 'sum',
                'like_count': 'sum',
                'share_count': 'sum',
                'view_count': 'sum'
            }).reset_index()
            
            # Add post count
            post_counts = fb_data.groupby('Company').size().reset_index(name='post_count')
            grouped_data = grouped_data.merge(post_counts, on='Company')
            
            # Create table
            rows = len(grouped_data) + 1  # +1 for header
            cols = 6  # Company, post_count, comment_count, like_count, share_count, view_count
            
            # Calculate centered position
            # Slide width: 13.33", Table width: 12.33" (leaving 0.5" on each side)
            # Center horizontally: (13.33 - 12.33) / 2 = 0.5"
            left = Inches(0.5)
            
            # Slide height: 7.5", Header: 0.8", Table height: 5"
            # Center vertically in remaining space: 0.8 + (6.7 - 5) / 2 = 1.65"
            top = Inches(1)
            width = Inches(12.33)
            height = Inches(5)
            
            table = slide7.shapes.add_table(rows, cols, left, top, width, height).table
            
            # Set column widths with wider spacing
            table.columns[0].width = Inches(3)  # Company name (wider for longer names)
            for i in range(1, cols):
                table.columns[i].width = Inches(1.866)  # Metrics (remaining width distributed evenly)
            
            # Add headers with red background
            headers = ['Banklar ', 'Post sayı', 'Şərh sayı', 'Bəyənmə sayı', 'Paylaşım sayı', 'Baxış sayı']
            for i, header in enumerate(headers):
                cell = table.cell(0, i)
                cell.text = header
                cell.text_frame.paragraphs[0].font.size = Pt(12)
                cell.text_frame.paragraphs[0].font.bold = True
                cell.fill.solid()
                cell.fill.fore_color.rgb = HEADER_TEXT_COLOR
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)  # White text
            
            # Add data with alternating row colors
            for i, row in grouped_data.iterrows():
                for j in range(cols):
                    cell = table.cell(i + 1, j)
                    if j == 0:
                        cell.text = str(row['Company'])
                    elif j == 1:
                        cell.text = str(row['post_count'])
                    elif j == 2:
                        cell.text = str(row['view_count'])
                    elif j == 3:
                        cell.text = str(row['comment_count'])
                    elif j == 4:
                        cell.text = str(row['like_count'])
                    else:
                        cell.text = str(row['share_count'])
                    
                    cell.text_frame.paragraphs[0].font.size = Pt(10)
                    
                    # Set alternating row colors
                    if i % 2 == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(240, 240, 240)  # Light gray
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White
        # endregion

        # region Eighth slide - Official Facebook metrics table
        logger.debug("Creating seventh slide with Facebook metrics table")
        slide8 = prs.slides.add_slide(prs.slide_layouts[5])
        # Remove default textbox
        for shape in slide8.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)
        add_slide_header(slide8, company_logo_path, start_date, end_date, "Bankların rəsmi Facebook səhifələrinin analizi")

        # Set slide background
        background = slide8.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        if 'official_facebook' in data_frames and len(data_frames['official_facebook']) > 0:
            fb_data = list(data_frames['official_facebook'].values())[0]
            
            # Verify required columns exist
            required_columns = ['author_name', 'comment_count', 'like_count', 'share_count', 'view_count']
            if not all(col in fb_data.columns for col in required_columns):
                logger.error("Missing required columns in Facebook data")
                raise ValueError("Facebook data is missing required columns")
            
            # Group data by author_name
            grouped_data = fb_data.groupby('author_name').agg({
                'comment_count': 'sum',
                'like_count': 'sum',
                'share_count': 'sum',
                'view_count': 'sum'
            }).reset_index()
            
            # Add post count
            post_counts = fb_data.groupby('author_name').size().reset_index(name='post_count')
            grouped_data = grouped_data.merge(post_counts, on='author_name')
            
            # Create table
            rows = len(grouped_data) + 1  # +1 for header
            cols = 6  # Company, post_count, comment_count, like_count, share_count, view_count
            
            # Calculate centered position
            # Slide width: 13.33", Table width: 12.33" (leaving 0.5" on each side)
            # Center horizontally: (13.33 - 12.33) / 2 = 0.5"
            left = Inches(0.5)
            
            # Slide height: 7.5", Header: 0.8", Table height: 5"
            # Center vertically in remaining space: 0.8 + (6.7 - 5) / 2 = 1.65"
            top = Inches(1)
            width = Inches(12.33)
            height = Inches(5)
            
            table = slide8.shapes.add_table(rows, cols, left, top, width, height).table
            
            # Set column widths with wider spacing
            table.columns[0].width = Inches(3)  # Company name (wider for longer names)
            for i in range(1, cols):
                table.columns[i].width = Inches(1.866)  # Metrics (remaining width distributed evenly)
            
            # Add headers with red background
            headers = ['Banklar ', 'Post sayı', 'Şərh sayı', 'Bəyənmə sayı', 'Paylaşım sayı', 'Baxış sayı']
            for i, header in enumerate(headers):
                cell = table.cell(0, i)
                cell.text = header
                cell.text_frame.paragraphs[0].font.size = Pt(12)
                cell.text_frame.paragraphs[0].font.bold = True
                cell.fill.solid()
                cell.fill.fore_color.rgb = HEADER_TEXT_COLOR
                cell.text_frame.paragraphs[0].font.color.rgb = RGBColor(255, 255, 255)  # White text
            
            # Add data with alternating row colors
            for i, row in grouped_data.iterrows():
                for j in range(cols):
                    cell = table.cell(i + 1, j)
                    if j == 0:
                        cell.text = str(row['author_name'])
                    elif j == 1:
                        cell.text = str(row['post_count'])
                    elif j == 2:
                        cell.text = str(row['comment_count'])  # Changed order to match headers
                    elif j == 3:
                        cell.text = str(row['like_count'])    # Changed order to match headers
                    elif j == 4:
                        cell.text = str(row['share_count'])   # Changed order to match headers
                    else:
                        cell.text = str(row['view_count'])    # Changed order to match headers
                    
                    cell.text_frame.paragraphs[0].font.size = Pt(10)
                    
                    # Set alternating row colors
                    if i % 2 == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(240, 240, 240)  # Light gray
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White
        # endregion

        # region Ninth slide - Instagram metrics and sentiment analysis
        logger.debug("Creating ninth slide with Instagram metrics and sentiment analysis")
        slide9 = prs.slides.add_slide(prs.slide_layouts[5])
        # Remove default textbox
        for shape in slide9.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)
        add_slide_header(slide9, company_logo_path, start_date, end_date, "Instagram postlarının analizi")

        # Set slide background
        background = slide9.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR            # Left side - Instagram metrics from official_instagram
        if 'official_instagram' in data_frames and len(data_frames['official_instagram']) > 0:
            insta_data = list(data_frames['official_instagram'].values())[0]
            company_metrics = insta_data[insta_data['Company'] == company_name]
            
            network_metrics = {
                'Total Likes': company_metrics['Likes'].sum(),
                'Total Comments': company_metrics['Comments'].sum()
            }
            
            # Set up left section (20% width) for metrics
            left = Inches(0.5)
            top = Inches(1)
            metrics_width = Inches(2.7)  # ~20% of slide width (13.33")
            metrics_height = Inches(1.2)  # Taller boxes for new layout
            available_height = Inches(6.5)  # Total height minus header
            total_items = len(network_metrics)
            spacing = (available_height - (metrics_height * total_items)) / (total_items + 1)

            for i, (metric, value) in enumerate(network_metrics.items()):
                # Calculate position for current metric box
                box_top = top + (spacing * (i + 1)) + (metrics_height * i)
                
                # Add white background box for the whole metric
                bg_box = slide9.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    left, box_top, metrics_width, metrics_height
                )
                bg_box.fill.solid()
                bg_box.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
                
                # Add red background for icon
                icon_width = Inches(1.0)
                icon_box = slide9.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    left, box_top, icon_width, metrics_height
                )
                icon_box.fill.solid()
                icon_box.fill.fore_color.rgb = HEADER_TEXT_COLOR  # Red background
                
                # Add icon text
                icon_text = slide9.shapes.add_textbox(
                    left, box_top + Inches(0.3), icon_width, Inches(0.6)
                )
                icon_p = icon_text.text_frame.add_paragraph()
                icon_p.text = "📊"
                icon_p.alignment = PP_ALIGN.CENTER
                icon_p.font.size = Pt(24)
                
                # Add value in large text
                value_text = slide9.shapes.add_textbox(
                    left + icon_width + Inches(0.1),
                    box_top + Inches(0.1),
                    metrics_width - icon_width - Inches(0.2),
                    Inches(0.6)
                )
                value_p = value_text.text_frame.add_paragraph()
                value_p.text = f"{value:,}"
                value_p.alignment = PP_ALIGN.LEFT
                value_p.font.size = Pt(20)
                value_p.font.bold = True
                
                # Add metric name below value
                metric_text = slide9.shapes.add_textbox(
                    left + icon_width + Inches(0.1),
                    box_top + Inches(0.7),
                    metrics_width - icon_width - Inches(0.2),
                    Inches(0.4)
                )
                metric_p = metric_text.text_frame.add_paragraph()
                metric_p.text = metric
                metric_p.alignment = PP_ALIGN.LEFT
                metric_p.font.size = Pt(12)

        # Right side - Sentiment analysis from combined_sources
        if 'combined_sources' in data_frames and 'Instagram' in data_frames['combined_sources']:
            insta_sentiment_data = data_frames['combined_sources']['Instagram']
            company_sentiment = insta_sentiment_data[insta_sentiment_data['Company'] == company_name]
            
            # Donut chart
            sentiment_counts = company_sentiment['Sentiment'].value_counts()
            donut_data = ChartData()
            donut_data.categories = ['Positive', 'Neutral', 'Negative']
            donut_data.add_series('', [  # Empty series name to remove text
                sentiment_counts.get(1, 0),
                sentiment_counts.get(0, 0),
                sentiment_counts.get(-1, 0)
            ])
            
            # Right section (80% width) layout
            right_section_left = Inches(3.7)  # After left 20% section
            right_width = Inches(9.13)  # 80% of slide width
            
            # Position smaller donut chart centered in its section
            donut_size = Inches(3.5)
            x = right_section_left
            y = Inches(1)
            
            # Add title for donut chart
            donut_title = slide9.shapes.add_textbox(
                x, y, donut_size, Inches(0.3)
            )
            title_p = donut_title.text_frame.add_paragraph()
            title_p.text = "📊 Ümumi sentiment bölgüsü"
            title_p.font.size = Pt(14)
            title_p.font.bold = True
            title_p.font.color.rgb = HEADER_TEXT_COLOR
            title_p.alignment = PP_ALIGN.CENTER
            
            # Position donut chart below title
            donut = slide9.shapes.add_chart(
                XL_CHART_TYPE.DOUGHNUT, 
                x, y + Inches(0.4), 
                donut_size, donut_size - Inches(0.4), 
                donut_data
            ).chart
            donut.has_legend = True
            donut.legend.position = XL_LEGEND_POSITION.TOP
            donut.legend.font.size = Pt(8)  # Reduced legend size
            
            # Set chart background
            donut.chart_style = 2  # White background
            
            for i, point in enumerate(donut.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                # Add data label with arrow
                point.has_data_label = True
                point.data_label.font.size = Pt(10)
                point.data_label.font.bold = True
                point.data_label.number_format = '#,##0'
                
            # Multiline chart
            sentiment_by_date = company_sentiment.groupby('Day')['Sentiment'].value_counts().unstack(fill_value=0)
            chart_data = CategoryChartData()
            chart_data.categories = sentiment_by_date.index.tolist()
            
            for sentiment in [1, 0, -1]:
                series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                if sentiment in sentiment_by_date.columns:
                    chart_data.add_series(series_name, sentiment_by_date[sentiment].tolist())
            
            x = right_section_left + donut_size + Inches(0.5)  # After donut chart
            y = Inches(1.25)
            cx = right_width - donut_size - Inches(0.5)  # Remaining width
            cy = donut_size - Inches(0.25)  # Same height as donut
            chart = slide9.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.TOP
            chart.legend.font.size = Pt(12)
            chart.has_title = True
            chart.chart_title.text_frame.text = "📊 Postların zamana və sentimentə görə bölgüsü"
            paragraph = chart.chart_title.text_frame.paragraphs[0]
            paragraph.font.size = Pt(14)
            paragraph.font.bold = True
            paragraph.font.color.rgb = HEADER_TEXT_COLOR
            # Set chart background and formatting
            apply_chart_formatting(chart, title="Postların zamana və sentimentə görə bölgüsü")
            for i, series in enumerate(chart.series):
                series.format.line.color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                series.format.line.width = Pt(2)
            
            # Vertical multibar chart for Instagram company sentiment comparison
            insta_sentiment_data = insta_sentiment_data[insta_sentiment_data['Sentiment'].isin([-1, 0, 1])]
            company_sentiments = insta_sentiment_data.groupby('Company')['Sentiment'].value_counts().unstack(fill_value=0)
            # Sort by total sentiment values
            totals = company_sentiments.sum(axis=1)
            company_sentiments = company_sentiments.loc[totals.sort_values(ascending=False).index]

            chart_data = CategoryChartData()
            chart_data.categories = company_sentiments.index.tolist()

            for sentiment in [1, 0, -1]:
                series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                if sentiment in company_sentiments.columns:
                    chart_data.add_series(series_name, company_sentiments[sentiment].tolist())

            # Position multibar chart to take full width of right section
            x = right_section_left
            y = Inches(4.5)  # Position below the top charts
            cx = right_width  # Full width of right section
            cy = Inches(3)  # Height
            chart = slide9.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.TOP
            chart.legend.font.size = Pt(12)

            # Set chart background and formatting
            apply_sentiment_colors(chart)
            apply_chart_formatting(chart, title="Post saylarına görə bankların bölgüsü")
        # endregion

        # region Tenth slide - Linkedin sentiment analysis
        logger.debug("Creating tenth slide with Linkedin sentiment analysis")
        slide10 = prs.slides.add_slide(prs.slide_layouts[5])
        # Remove default textbox
        for shape in slide10.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)
        add_slide_header(slide10, company_logo_path, start_date, end_date, "Linkedln postlarının analizi")

        # Set slide background
        background = slide10.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        if 'combined_sources' in data_frames and 'Linkedin' in data_frames['combined_sources']:
            linkedin_data = data_frames['combined_sources']['Linkedin']
            company_sentiment = linkedin_data[linkedin_data['Company'] == company_name]
            
            if not company_sentiment.empty:
                # Calculate heights accounting for header
                available_height = Inches(7.5 - HEADER_HEIGHT)  # Total height minus header
                half_height = available_height / 2
                
                # Top half section for donut and line charts
                donut_size = Inches(3.5)
                
                # Donut chart on left top half
                x_donut = Inches(0.5)
                y_donut = Inches(1.2)  # Just below header
                
                # Add title for donut chart
                donut_title = slide10.shapes.add_textbox(
                    x_donut, y_donut, donut_size, Inches(0.3)
                )
                title_p = donut_title.text_frame.add_paragraph()
                title_p.text = "📊 Postların sentiment bölgüsü"
                title_p.font.size = Pt(14)
                title_p.font.bold = True
                title_p.font.color.rgb = HEADER_TEXT_COLOR
                title_p.alignment = PP_ALIGN.CENTER
                
                # Add donut chart below its title
                sentiment_counts = company_sentiment['Sentiment'].value_counts()
                donut_data = ChartData()
                donut_data.categories = ['Positive', 'Neutral', 'Negative']
                donut_data.add_series('', [
                    sentiment_counts.get(1, 0),
                    sentiment_counts.get(0, 0),
                    sentiment_counts.get(-1, 0)
                ])
                
                donut = slide10.shapes.add_chart(
                    XL_CHART_TYPE.DOUGHNUT, 
                    x_donut, 
                    y_donut + Inches(0.4),
                    donut_size, 
                    donut_size - Inches(0.4),
                    donut_data
                ).chart
                
                donut.has_legend = True
                donut.legend.position = XL_LEGEND_POSITION.TOP
                donut.legend.font.size = Pt(10)
                donut.chart_style = 2  # White background
                
                # Apply colors and data labels to donut chart
                for i, point in enumerate(donut.series[0].points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                    point.has_data_label = True
                    point.data_label.font.size = Pt(10)
                    point.data_label.font.bold = True
                    point.data_label.number_format = '#,##0'
                
                # Multiline chart in right half
                sentiment_by_date = company_sentiment.groupby('Day')['Sentiment'].value_counts().unstack(fill_value=0)
                chart_data = CategoryChartData()
                chart_data.categories = sentiment_by_date.index.tolist()
                
                for sentiment in [1, 0, -1]:
                    series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                    if sentiment in sentiment_by_date.columns:
                        chart_data.add_series(series_name, sentiment_by_date[sentiment].tolist())
                
                # Position multiline chart in right half of top section
                x_line = Inches(4.5)  # Start after donut chart
                y_line = Inches(1.2)  # Same vertical alignment as donut
                cx_line = Inches(8.33)  # Remaining width
                cy_line = donut_size  # Same height as donut
                
                chart = slide10.shapes.add_chart(
                    XL_CHART_TYPE.LINE,
                    x_line, y_line,
                    cx_line, cy_line,
                    chart_data
                ).chart
                
                # Format line chart
                chart.has_legend = True
                chart.legend.position = XL_LEGEND_POSITION.TOP
                chart.legend.font.size = Pt(12)
                
                # Set chart background and formatting
                # Set colors for donut chart
                for i, point in enumerate(donut.series[0].points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                apply_chart_formatting(chart, title="Postların zamana və sentimentə görə bölgüsü")
                for i, series in enumerate(chart.series):
                    series.format.line.color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                    series.format.line.width = Pt(2)
                # Vertical multibar chart for LinkedIn company sentiment comparison
                linkedin_data = linkedin_data[linkedin_data['Sentiment'].isin([-1, 0, 1])]
                company_sentiments = linkedin_data.groupby('Company')['Sentiment'].value_counts().unstack(fill_value=0)
                
                # Sort by total sentiment values
                totals = company_sentiments.sum(axis=1)
                company_sentiments = company_sentiments.loc[totals.sort_values(ascending=False).index]
                
                chart_data = CategoryChartData()
                chart_data.categories = company_sentiments.index.tolist()
                
                for sentiment in [1, 0, -1]:
                    series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                    if sentiment in company_sentiments.columns:
                        chart_data.add_series(series_name, company_sentiments[sentiment].tolist())
                
                # Position multibar chart in bottom half, full width
                x_bar = Inches(0.5)
                y_bar = Inches(4.5)  # Start below top section
                cx_bar = Inches(12.33)  # Full width
                cy_bar = Inches(2.5)  # Remaining height
                
                chart = slide10.shapes.add_chart(
                    XL_CHART_TYPE.COLUMN_CLUSTERED,
                    x_bar, y_bar,
                    cx_bar, cy_bar,
                    chart_data
                ).chart
                
                chart.has_legend = True
                chart.legend.position = XL_LEGEND_POSITION.TOP
                chart.legend.font.size = Pt(12)
                
                # Apply formatting and colors
                apply_chart_formatting(chart)
                apply_sentiment_colors(chart)
                apply_chart_formatting(chart, title="Post saylarına görə bankların bölgüsü")
            else:
                logger.warning(f"No Linkedin data found for company: {company_name}")
        # endregion

        # region Eleventh slide - Positive and Negative Posts
        logger.debug("Creating Eleventh slide with positive and negative posts")
        slide11 = prs.slides.add_slide(prs.slide_layouts[5])
        # Remove default textbox
        for shape in slide11.shapes:
            if shape.has_text_frame:
                sp = shape._element
                sp.getparent().remove(sp)
        add_slide_header(slide11, company_logo_path, start_date, end_date, "Sosial media postlarının analizi")

        image_width = Inches(2)
        vertical_spacing = Inches(0.05)
        caption_height = Inches(0.35)
        group_top = Inches(1)
        max_group_height = Inches(5.5)

        def add_post(slide, post, left, top):
            if not os.path.exists(post["image_path"]):
                return 0

           

            # Add image and get actual height
            img = slide.shapes.add_picture(post["image_path"], left, top, width=image_width)
            img_height = img.height / 914400

            # Link
            img.click_action.hyperlink.address = post["link"]

            # Styling
            img.line.color.rgb = RGBColor(200, 200, 200)
            img.shadow.inherit = False
            img.shadow.blur_radius = 5000

            # Caption
            caption_top = top + Inches(img_height)  + Inches(0.05)
            caption_box = slide.shapes.add_textbox(left, caption_top, image_width, caption_height)
            tf = caption_box.text_frame
            p = tf.paragraphs[0]
            p.text = post.get("caption", "🔗 Click to view post")
            p.font.size = Pt(10)
            p.alignment = PP_ALIGN.CENTER
            p.font.color.rgb = RGBColor(80, 80, 80)

            return Inches(img_height) + caption_height + vertical_spacing

        def layout_posts(slide, posts, group_center_x):
            count = len(posts)
            positions = []

            if count == 1:
                positions.append((group_center_x - image_width / 2, group_top + Inches(2.5)))
            elif count == 2:
                positions.append((group_center_x - image_width / 2, group_top + Inches(0.5)))  # top
                positions.append((group_center_x - image_width / 2, group_top + Inches(3)))  # bottom
            elif count == 3:
                # Top row (two side by side)
                spacing_x = Inches(0.3)
                left1 = group_center_x - image_width - spacing_x / 2
                left2 = group_center_x + spacing_x / 2
                top1 = group_top + Inches(0.5)
                positions.append((left1, top1))
                positions.append((left2, top1))
                # Centered below
                center_x = group_center_x - image_width / 2
                positions.append((center_x, group_top + Inches(3)))

            # Add posts
            for post, (left, top) in zip(posts, positions):
                add_post(slide, post, left, top)

        # Apply layout for each group
        if negative_posts:
            layout_posts(slide11, negative_posts[:3], group_center_x=Inches(2.25))  # left half

        if positive_posts:
            layout_posts(slide11, positive_posts[:3], group_center_x=Inches(7))  # right half

        # endregion

        logger.debug("Saving PowerPoint file")
        prs.save(output_path)
        logger.debug("PowerPoint file saved successfully")
    except Exception as e:
        logger.error(f"Error creating PowerPoint: {str(e)}")
        raise