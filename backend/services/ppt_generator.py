from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData, ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK, XL_TICK_LABEL_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.dml import MSO_THEME_COLOR
from services.excel_parser import get_sentiment_data, get_sentiment_counts, get_company_sentiment_counts
import logging
import pandas as pd
import os

logger = logging.getLogger(__name__)

# Define sentiment colors
SENTIMENT_COLORS = {
    1: RGBColor(69, 194, 126),    # Positive - Green
    0: RGBColor(255, 191, 0),     # Neutral - Yellow
    -1: RGBColor(246, 1, 64)      # Negative - Red
}

# Define slide background color
SLIDE_BG_COLOR = RGBColor(230, 230, 230)  # Light gray

# Define chart background color
CHART_BG_COLOR = RGBColor(255, 255, 255)  # White

# Define metric icons
METRIC_ICONS = {
    'Posts Count': '📝',
    'Total Comments': '💬',
    'Total Likes': '👍',
    'Total Shares': '🔄',
    'Total Views': '👁️'
}

def add_slide_header(slide, company_logo_path, start_date, end_date, title):
    """Helper function to add a consistent header to slides"""
    # Add header background
    header = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0),
        Inches(10), Inches(0.8)
    )
    header.fill.solid()
    header.fill.fore_color.rgb = CHART_BG_COLOR
    
    # Add company logo
    if os.path.exists(company_logo_path):
        img = slide.shapes.add_picture(
            company_logo_path,
            Inches(0.2), Inches(0.1),
            height=Inches(0.6)
        )
    
    # Add title
    title_box = slide.shapes.add_textbox(Inches(3), Inches(0.1), Inches(4), Inches(0.6))
    tf = title_box.text_frame
    p = tf.add_paragraph()
    p.text = title
    p.font.size = Pt(16)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER
    
    # Add date range
    date_box = slide.shapes.add_textbox(Inches(7), Inches(0.1), Inches(2.5), Inches(0.6))
    tf = date_box.text_frame
    p = tf.add_paragraph()
    p.text = f"{start_date} - {end_date}"
    p.font.size = Pt(12)
    p.alignment = PP_ALIGN.RIGHT
    
    # Add divider line
    line = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(0), Inches(0.8),
        Inches(10), Inches(0.02)
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(0, 0, 0)

def create_ppt(data_frames, output_path, start_date, end_date, company_name, company_logo_path, mediaeye_logo_path, neurotime_logo_path, competitor_logo_paths=None, positive_links=None, negative_links=None, positive_posts=None, negative_posts=None):
    try:
        logger.debug("Creating PowerPoint presentation")
        prs = Presentation()
        
        # First slide - Title slide with logos and text
        logger.debug("Creating title slide")
        title_slide = prs.slides.add_slide(prs.slide_layouts[5])  # Blank layout
        
        # Set slide background
        background = title_slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR

        # Left side content
        left = Inches(0.5)
        top = Inches(0.5)
        width = Inches(4.5)
        
        # Add NeuroTime logo
        if os.path.exists(neurotime_logo_path):
            img = title_slide.shapes.add_picture(neurotime_logo_path, left, top, width=Inches(2))
        
        # Add text content
        current_top = top + Inches(2.5)
        
        # Title text
        title_box = title_slide.shapes.add_textbox(left, current_top, width, Inches(0.5))
        tf = title_box.text_frame
        p = tf.add_paragraph()
        p.text = "MARKA VARLIĞININ MEDİA ÜZƏRİNDƏN QİYMƏTLƏNDİRİLMƏSİ"
        p.font.size = Pt(24)
        p.font.bold = True
        current_top += Inches(0.7)
        
        # Divider line
        line = title_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, current_top, width, Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0, 0, 0)
        current_top += Inches(0.2)
        
        # Description text
        desc_box = title_slide.shapes.add_textbox(left, current_top, width, Inches(0.5))
        tf = desc_box.text_frame
        p = tf.add_paragraph()
        p.text = "Online/Sosial və ənənəvi media datalarının əsasında təhlil."
        p.font.size = Pt(16)
        current_top += Inches(0.7)
        
        # Divider line
        line = title_slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, current_top, width, Inches(0.02))
        line.fill.solid()
        line.fill.fore_color.rgb = RGBColor(0, 0, 0)
        current_top += Inches(0.2)
        
        # Report type text
        report_box = title_slide.shapes.add_textbox(left, current_top, width, Inches(0.5))
        tf = report_box.text_frame
        p = tf.add_paragraph()
        p.text = "Analitik hesabat"
        p.font.size = Pt(16)
        current_top += Inches(0.7)
        
        # Date range
        date_box = title_slide.shapes.add_textbox(left, current_top, width, Inches(0.5))
        tf = date_box.text_frame
        p = tf.add_paragraph()
        p.text = f"{start_date} - {end_date}"
        p.font.size = Pt(14)
        
        # Right side - Company logo
        if os.path.exists(company_logo_path):
            img = title_slide.shapes.add_picture(
                company_logo_path,
                Inches(5.5), Inches(2),
                width=Inches(4)
            )

        # Second slide - Methodology description
        logger.debug("Creating methodology slide")
        method_slide = prs.slides.add_slide(prs.slide_layouts[5])
                
        # Add MediaEye logo on the left
        if os.path.exists(mediaeye_logo_path):
            img = method_slide.shapes.add_picture(
                mediaeye_logo_path,
                Inches(0.5), Inches(1.5),
                width=Inches(4)
            )
        
        # Add NeuroTime logo on the right top
        if os.path.exists(neurotime_logo_path):
            img = method_slide.shapes.add_picture(
                neurotime_logo_path,
                Inches(7), Inches(1.5),
                width=Inches(2)
            )
        
        # Add competitor logos in grid layout
        if competitor_logo_paths:
            grid_left = Inches(5)
            grid_top = Inches(2)
            logo_width = Inches(1.5)
            logo_height = Inches(1.5)
            logos_per_row = 4
            spacing = Inches(0.2)
            
            for i, logo_path in enumerate(competitor_logo_paths):
                if os.path.exists(logo_path):
                    row = i // logos_per_row
                    col = i % logos_per_row
                    left = grid_left + (logo_width + spacing) * col
                    top = grid_top + (logo_height + spacing) * row
                    
                    img = method_slide.shapes.add_picture(
                        logo_path,
                        left, top,
                        width=logo_width,
                        height=logo_height
                    )

        # Third slide - Multiline chart and donut chart
        logger.debug("Creating third slide with charts")
        slide3 = prs.slides.add_slide(prs.slide_layouts[5])
        add_slide_header(slide3, company_logo_path, start_date, end_date, "Xəbərlərin analizi")
        
        # Set slide background
        background = slide3.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR

        # Add positive links above multiline chart
        if positive_links:
            left = Inches(0.5)
            top = Inches(1)
            width = Inches(4.5)
            height = Inches(0.4)
            
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

        # Add negative links above donut chart
        if negative_links:
            left = Inches(5.5)
            top = Inches(1)
            width = Inches(4.5)
            height = Inches(0.4)
            
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
        
        x, y, cx, cy = Inches(0.5), Inches(1.7), Inches(4.5), Inches(4)
        chart = slide3.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(10)
        
        # Set chart background and formatting
        chart.chart_style = 2  # White background
        chart.plots[0].has_major_gridlines = False
        chart.plots[0].has_minor_gridlines = False
        
        # Set axis titles and labels
        if hasattr(chart, 'category_axis'):
            chart.category_axis.tick_labels.font.size = Pt(7)
            if hasattr(chart.category_axis, 'axis_title') and chart.category_axis.axis_title:
                chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(8)
        if hasattr(chart, 'value_axis'):
            chart.value_axis.tick_labels.font.size = Pt(7)
            if hasattr(chart.value_axis, 'axis_title') and chart.value_axis.axis_title:
                chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(8)
        
        # Set colors for line chart
        for i, series in enumerate(chart.series):
            series.format.line.color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
            series.format.line.width = Pt(2)
        
        # Create donut chart
        logger.debug("Creating donut chart")
        donut_data = ChartData()
        donut_data.categories = ['Positive', 'Neutral', 'Negative']
        donut_data.add_series('Sentiment Distribution', [
            sentiment_counts.get(1, 0),
            sentiment_counts.get(0, 0),
            sentiment_counts.get(-1, 0)
        ])
        
        x, y, cx, cy = Inches(5.5), Inches(1.7), Inches(4.5), Inches(4)
        donut = slide3.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, donut_data).chart
        donut.has_legend = True
        donut.legend.position = XL_LEGEND_POSITION.BOTTOM
        donut.legend.font.size = Pt(10)
        
        # Set chart background
        donut.chart_style = 2  # White background
        
        # Set colors for donut chart and add data labels
        for i, point in enumerate(donut.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
            # Add data label
            point.has_data_label = True
            point.data_label.font.size = Pt(10)
            point.data_label.font.bold = True
            point.data_label.number_format = '#,##0'
        
        # Fourth slide - Vertical multibar chart
        logger.debug("Creating third slide with vertical multibar chart")
        slide4 = prs.slides.add_slide(prs.slide_layouts[5])
        add_slide_header(slide4, company_logo_path, start_date, end_date, "Xəbərlərin analizi")

        # Set slide background
        background = slide4.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        # Get company sentiment data
        company_sentiments = get_company_sentiment_counts(combined_data)
        
        chart_data = CategoryChartData()
        chart_data.categories = company_sentiments.index.tolist()
        
        for sentiment in [1, 0, -1]:
            series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
            if sentiment in company_sentiments.columns:
                chart_data.add_series(series_name, company_sentiments[sentiment].tolist())
        
        x, y, cx, cy = Inches(0.5), Inches(1), Inches(9), Inches(5)
        chart = slide4.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(10)
        
        # Set chart background and formatting
        chart.chart_style = 2  # White background
        chart.plots[0].has_major_gridlines = False
        chart.plots[0].has_minor_gridlines = False
        
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
        
        # Fifth slide - Author count horizontal bar chart
        logger.debug("Creating fourth slide with author count chart")
        slide5 = prs.slides.add_slide(prs.slide_layouts[5])
        add_slide_header(slide5, company_logo_path, start_date, end_date, "Xəbər saylarının saytlar üzərindən bölgüsü")

        # Set slide background
        background = slide5.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        # Filter and group data by author
        author_data = combined_data[combined_data['Company'] == company_name].groupby('Author').size().sort_values(ascending=True)
        
        chart_data = CategoryChartData()
        chart_data.categories = author_data.index.tolist()
        chart_data.add_series('Post Count', author_data.values.tolist())
        
        x, y, cx, cy = Inches(0.5), Inches(1), Inches(9), Inches(5)
        chart = slide5.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data).chart
        chart.has_legend = False
        
        # Set chart background and formatting
        chart.chart_style = 2  # White background
        chart.plots[0].has_major_gridlines = False
        chart.plots[0].has_minor_gridlines = False
        
        # Make bars thicker
        for series in chart.series:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = RGBColor(173, 216, 230)  # Light blue
            series.format.gap_width = 100  # Make bars thicker
        
        # Show all y-axis labels
        if hasattr(chart, 'category_axis'):
            chart.category_axis.tick_labels.font.size = Pt(8)
            chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        
        # Set blue color for bars
        for series in chart.series:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = RGBColor(0, 112, 192)  # Blue color
        
        # Sixth slide - Facebook metrics and sentiment analysis
        logger.debug("Creating fifth slide with Facebook metrics")
        slide6 = prs.slides.add_slide(prs.slide_layouts[5])
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
            
            # Create separate boxes for each metric
            left = Inches(0.5)
            top = Inches(1)
            width = Inches(2)
            height = Inches(0.8)
            spacing = Inches(0.2)
            
            for i, (metric, value) in enumerate(metrics.items()):
                # Add red background shape
                bg_shape = slide6.shapes.add_shape(
                    MSO_SHAPE.ROUNDED_RECTANGLE,
                    left, top + (height + spacing) * i, width, height
                )
                bg_shape.fill.solid()
                bg_shape.fill.fore_color.rgb = RGBColor(255, 0, 0)
                
                # Add text box with icon
                txBox = slide6.shapes.add_textbox(
                    left + Inches(0.1), 
                    top + (height + spacing) * i + Inches(0.1), 
                    width - Inches(0.2), 
                    height - Inches(0.2)
                )
                tf = txBox.text_frame
                p = tf.add_paragraph()
                p.text = f"{METRIC_ICONS[metric]} {metric}: {value:,}"
                p.alignment = PP_ALIGN.LEFT
                p.font.size = Pt(12)
                p.font.color.rgb = RGBColor(255, 255, 255)  # White text
        
        # Right side - Sentiment analysis
        if 'Facebook' in data_frames['combined_sources']:
            fb_sentiment_data = data_frames['combined_sources']['Facebook']
            company_sentiment = fb_sentiment_data[fb_sentiment_data['Company'] == company_name]
            
            # Donut chart
            sentiment_counts = company_sentiment['Sentiment'].value_counts()
            donut_data = ChartData()
            donut_data.categories = ['Positive', 'Neutral', 'Negative']
            donut_data.add_series('Sentiment Distribution', [
                sentiment_counts.get(1, 0),
                sentiment_counts.get(0, 0),
                sentiment_counts.get(-1, 0)
            ])
            
            x, y, cx, cy = Inches(3), Inches(1), Inches(3), Inches(3)
            donut = slide6.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, donut_data).chart
            donut.has_legend = True
            donut.legend.position = XL_LEGEND_POSITION.BOTTOM
            donut.legend.font.size = Pt(10)
            
            # Set chart background
            donut.chart_style = 2  # White background
            
            # Set colors for donut chart
            for i, point in enumerate(donut.series[0].points):
                point.format.fill.solid()
                point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
            
            # Multiline chart
            sentiment_by_date = company_sentiment.groupby('Day')['Sentiment'].value_counts().unstack(fill_value=0)
            chart_data = CategoryChartData()
            chart_data.categories = sentiment_by_date.index.tolist()
            
            for sentiment in [1, 0, -1]:
                series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                if sentiment in sentiment_by_date.columns:
                    chart_data.add_series(series_name, sentiment_by_date[sentiment].tolist())
            
            x, y, cx, cy = Inches(6), Inches(1), Inches(3), Inches(3)
            chart = slide6.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.font.size = Pt(10)
            
            # Set chart background and formatting
            chart.chart_style = 2  # White background
            chart.plots[0].has_major_gridlines = False
            chart.plots[0].has_minor_gridlines = False
            
            # Set colors for line chart
            for i, series in enumerate(chart.series):
                series.format.line.color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                series.format.line.width = Pt(2)
            
            # Vertical multibar chart
            company_sentiments = fb_sentiment_data.groupby('Company')['Sentiment'].value_counts().unstack(fill_value=0)
            chart_data = CategoryChartData()
            chart_data.categories = company_sentiments.index.tolist()
            
            for sentiment in [1, 0, -1]:
                series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                if sentiment in company_sentiments.columns:
                    chart_data.add_series(series_name, company_sentiments[sentiment].tolist())
            
            x, y, cx, cy = Inches(3), Inches(4), Inches(6), Inches(3)
            chart = slide6.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.font.size = Pt(10)
            
            # Set chart background and formatting
            chart.chart_style = 2  # White background
            chart.plots[0].has_major_gridlines = False
            chart.plots[0].has_minor_gridlines = False
            
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
        
        # Seventh slide - Facebook metrics table
        logger.debug("Creating sixth slide with Facebook metrics table")
        slide7 = prs.slides.add_slide(prs.slide_layouts[5])
        add_slide_header(slide7, company_logo_path, start_date, end_date, "Bankların rəsmi Facebook səhifələrinin analizi")

        # Set slide background
        background = slide7.background
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
            
            left = Inches(0.5)
            top = Inches(1)
            width = Inches(9)
            height = Inches(5)
            
            table = slide7.shapes.add_table(rows, cols, left, top, width, height).table
            
            # Set column widths
            table.columns[0].width = Inches(2)  # Company
            for i in range(1, cols):
                table.columns[i].width = Inches(1.4)  # Metrics
            
            # Add headers with red background
            headers = ['Company', 'Posts', 'Comments', 'Likes', 'Shares', 'Views']
            for i, header in enumerate(headers):
                cell = table.cell(0, i)
                cell.text = header
                cell.text_frame.paragraphs[0].font.size = Pt(12)
                cell.text_frame.paragraphs[0].font.bold = True
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(255, 0, 0)  # Red
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
                        cell.text = str(row['comment_count'])
                    elif j == 3:
                        cell.text = str(row['like_count'])
                    elif j == 4:
                        cell.text = str(row['share_count'])
                    else:
                        cell.text = str(row['view_count'])
                    
                    cell.text_frame.paragraphs[0].font.size = Pt(10)
                    
                    # Set alternating row colors
                    if i % 2 == 0:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(240, 240, 240)  # Light gray
                    else:
                        cell.fill.solid()
                        cell.fill.fore_color.rgb = RGBColor(255, 255, 255)  # White
        
        # Eigthth slide - Linkedin sentiment analysis
        logger.debug("Creating seventh slide with Linkedin sentiment analysis")
        slide8 = prs.slides.add_slide(prs.slide_layouts[5])
        add_slide_header(slide8, company_logo_path, start_date, end_date, "Linkedln postlarının analizi")

        # Set slide background
        background = slide8.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        if 'combined_sources' in data_frames and 'Linkedin' in data_frames['combined_sources']:
            linkedin_data = data_frames['combined_sources']['Linkedin']
            company_sentiment = linkedin_data[linkedin_data['Company'] == company_name]
            
            if not company_sentiment.empty:
                # Donut chart
                sentiment_counts = company_sentiment['Sentiment'].value_counts()
                donut_data = ChartData()
                donut_data.categories = ['Positive', 'Neutral', 'Negative']
                donut_data.add_series('Sentiment Distribution', [
                    sentiment_counts.get(1, 0),
                    sentiment_counts.get(0, 0),
                    sentiment_counts.get(-1, 0)
                ])
                
                x, y, cx, cy = Inches(0.5), Inches(1), Inches(4), Inches(4)
                donut = slide8.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, donut_data).chart
                donut.has_legend = True
                donut.legend.position = XL_LEGEND_POSITION.BOTTOM
                donut.legend.font.size = Pt(10)
                
                # Set chart background
                donut.chart_style = 2  # White background
                
                # Set colors for donut chart
                for i, point in enumerate(donut.series[0].points):
                    point.format.fill.solid()
                    point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                
                # Multiline chart
                sentiment_by_date = company_sentiment.groupby('Day')['Sentiment'].value_counts().unstack(fill_value=0)
                chart_data = CategoryChartData()
                chart_data.categories = sentiment_by_date.index.tolist()
                
                for sentiment in [1, 0, -1]:
                    series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                    if sentiment in sentiment_by_date.columns:
                        chart_data.add_series(series_name, sentiment_by_date[sentiment].tolist())
                
                x, y, cx, cy = Inches(5), Inches(1), Inches(4.5), Inches(4)
                chart = slide8.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
                chart.has_legend = True
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart.legend.font.size = Pt(10)
                
                # Set chart background and formatting
                chart.chart_style = 2  # White background
                chart.plots[0].has_major_gridlines = False
                chart.plots[0].has_minor_gridlines = False
                
                # Set colors for line chart
                for i, series in enumerate(chart.series):
                    series.format.line.color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
                    series.format.line.width = Pt(2)
                
                # Set axis titles and labels
                if hasattr(chart, 'category_axis'):
                    chart.category_axis.tick_labels.font.size = Pt(7)
                    if hasattr(chart.category_axis, 'axis_title') and chart.category_axis.axis_title:
                        chart.category_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(8)
                if hasattr(chart, 'value_axis'):
                    chart.value_axis.tick_labels.font.size = Pt(7)
                    if hasattr(chart.value_axis, 'axis_title') and chart.value_axis.axis_title:
                        chart.value_axis.axis_title.text_frame.paragraphs[0].font.size = Pt(8)
                # Vertical multibar chart for LinkedIn company sentiment comparison
                linkedin_data = linkedin_data[linkedin_data['Sentiment'].isin([-1, 0, 1])]
                company_sentiments = linkedin_data.groupby('Company')['Sentiment'].value_counts().unstack(fill_value=0)

                chart_data = CategoryChartData()
                chart_data.categories = company_sentiments.index.tolist()

                for sentiment in [1, 0, -1]:
                    series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                    if sentiment in company_sentiments.columns:
                        chart_data.add_series(series_name, company_sentiments[sentiment].tolist())

                x, y, cx, cy = Inches(0.5), Inches(5), Inches(9), Inches(3)
                chart = slide8.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
                chart.has_legend = True
                chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                chart.legend.font.size = Pt(10)

                # Set chart background and formatting
                chart.chart_style = 2  # White background
                chart.plots[0].has_major_gridlines = False
                chart.plots[0].has_minor_gridlines = False

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

            else:
                logger.warning(f"No Linkedin data found for company: {company_name}")
        
        # Ninth slide - Positive and Negative Postsif positive_posts or negative_posts:
        slide9 = prs.slides.add_slide(prs.slide_layouts[5])
        add_slide_header(slide9, company_logo_path, start_date, end_date, "Sosial media postlarının analizi")

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
            layout_posts(slide9, negative_posts[:3], group_center_x=Inches(2.25))  # left half

        if positive_posts:
            layout_posts(slide9, positive_posts[:3], group_center_x=Inches(7))  # right half


        logger.debug("Saving PowerPoint file")
        prs.save(output_path)
        logger.debug("PowerPoint file saved successfully")
    except Exception as e:
        logger.error(f"Error creating PowerPoint: {str(e)}")
        raise