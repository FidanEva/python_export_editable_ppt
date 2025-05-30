from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData, ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION, XL_TICK_MARK, XL_TICK_LABEL_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from services.excel_parser import get_sentiment_data, get_sentiment_counts, get_company_sentiment_counts
import logging
import pandas as pd

logger = logging.getLogger(__name__)

# Define sentiment colors
SENTIMENT_COLORS = {
    1: RGBColor(69, 194, 126),    # Positive - Green
    0: RGBColor(255, 191, 0),     # Neutral - Yellow
    -1: RGBColor(246, 1, 64)      # Negative - Red
}

# Define slide background color
SLIDE_BG_COLOR = RGBColor(240, 240, 240)  # Light gray

# Define chart background color
CHART_BG_COLOR = RGBColor(255, 255, 255)  # White

def create_ppt(data_frames, output_path, date, company_name):
    try:
        logger.debug("Creating PowerPoint presentation")
        prs = Presentation()
        
        # First slide - Title slide with date and company name
        logger.debug("Creating title slide")
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title = title_slide.shapes.title
        subtitle = title_slide.placeholders[1]
        
        # Set slide background
        background = title_slide.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        title.text = f"Report for {company_name}"
        title.text_frame.paragraphs[0].font.size = Pt(32)
        subtitle.text = f"Date: {date}"
        subtitle.text_frame.paragraphs[0].font.size = Pt(24)
        
        # Second slide - Multiline chart and donut chart
        logger.debug("Creating second slide with charts")
        slide2 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Set slide background
        background = slide2.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        # Get data from combined_sources
        logger.debug("Processing combined sources data")
        if 'combined_sources' not in data_frames:
            raise ValueError("combined_sources Excel file is missing")
        
        if 'News' not in data_frames['combined_sources']:
            raise ValueError("News sheet is missing in combined_sources Excel file")
        
        combined_data = data_frames['combined_sources']['News']
        sentiment_data = get_sentiment_data(combined_data, company_name)
        sentiment_counts = get_sentiment_counts(combined_data)
        
        # Create multiline chart
        logger.debug("Creating multiline chart")
        chart_data = CategoryChartData()
        chart_data.categories = sentiment_data.index.tolist()
        
        for sentiment in [1, 0, -1]:
            series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
            if sentiment in sentiment_data.columns:
                chart_data.add_series(series_name, sentiment_data[sentiment].tolist())
        
        x, y, cx, cy = Inches(0.5), Inches(1), Inches(4.5), Inches(4)
        chart = slide2.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
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
        
        # Create donut chart
        logger.debug("Creating donut chart")
        donut_data = ChartData()
        donut_data.categories = ['Positive', 'Neutral', 'Negative']
        donut_data.add_series('Sentiment Distribution', [
            sentiment_counts.get(1, 0),
            sentiment_counts.get(0, 0),
            sentiment_counts.get(-1, 0)
        ])
        
        x, y, cx, cy = Inches(5.5), Inches(1), Inches(4.5), Inches(4)
        donut = slide2.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, donut_data).chart
        donut.has_legend = True
        donut.legend.position = XL_LEGEND_POSITION.BOTTOM
        donut.legend.font.size = Pt(10)
        
        # Set chart background
        donut.chart_style = 2  # White background
        
        # Set colors for donut chart
        for i, point in enumerate(donut.series[0].points):
            point.format.fill.solid()
            point.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
        
        # Third slide - Vertical multibar chart
        logger.debug("Creating third slide with vertical multibar chart")
        slide3 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Set slide background
        background = slide3.background
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
        chart = slide3.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        chart.legend.font.size = Pt(10)
        
        # Set chart background and formatting
        chart.chart_style = 2  # White background
        chart.plots[0].has_major_gridlines = False
        chart.plots[0].has_minor_gridlines = False
        
        # Set diagonal labels for x-axis
        if hasattr(chart, 'category_axis'):
            chart.category_axis.tick_labels.font.size = Pt(8)
            chart.category_axis.tick_labels.orientation = 45
        
        # Set colors for bar chart
        for i, series in enumerate(chart.series):
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
        
        # Fourth slide - Author count horizontal bar chart
        logger.debug("Creating fourth slide with author count chart")
        slide4 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Set slide background
        background = slide4.background
        fill = background.fill
        fill.solid()
        fill.fore_color.rgb = SLIDE_BG_COLOR
        
        # Filter and group data by author
        author_data = combined_data[combined_data['Company'] == company_name].groupby('Author').size().sort_values(ascending=True)
        
        chart_data = CategoryChartData()
        chart_data.categories = author_data.index.tolist()
        chart_data.add_series('Post Count', author_data.values.tolist())
        
        x, y, cx, cy = Inches(0.5), Inches(1), Inches(9), Inches(5)
        chart = slide4.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data).chart
        chart.has_legend = False
        
        # Set chart background and formatting
        chart.chart_style = 2  # White background
        chart.plots[0].has_major_gridlines = False
        chart.plots[0].has_minor_gridlines = False
        
        # Show all y-axis labels
        if hasattr(chart, 'category_axis'):
            chart.category_axis.tick_labels.font.size = Pt(8)
            chart.category_axis.tick_label_position = XL_TICK_LABEL_POSITION.LOW
        
        # Set blue color for bars
        for series in chart.series:
            series.format.fill.solid()
            series.format.fill.fore_color.rgb = RGBColor(0, 112, 192)  # Blue color
        
        # Fifth slide - Facebook metrics and sentiment analysis
        logger.debug("Creating fifth slide with Facebook metrics")
        slide5 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Set slide background
        background = slide5.background
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
            
            # Create text box for metrics with icons and red background
            left, top, width, height = Inches(0.5), Inches(1), Inches(2), Inches(5)
            
            # Add red background shape first
            bg_shape = slide5.shapes.add_shape(
                MSO_SHAPE.RECTANGLE,
                left, top, width, height
            )
            bg_shape.fill.solid()
            bg_shape.fill.fore_color.rgb = RGBColor(255, 0, 0)
            
            # Add text box on top
            txBox = slide5.shapes.add_textbox(left, top, width, height)
            tf = txBox.text_frame
            
            for metric, value in metrics.items():
                p = tf.add_paragraph()
                p.text = f"{metric}: {value:,}"
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
            donut = slide5.shapes.add_chart(XL_CHART_TYPE.DOUGHNUT, x, y, cx, cy, donut_data).chart
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
            sentiment_by_date = company_sentiment.groupby('Date')['Sentiment'].value_counts().unstack(fill_value=0)
            chart_data = CategoryChartData()
            chart_data.categories = sentiment_by_date.index.tolist()
            
            for sentiment in [1, 0, -1]:
                series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
                if sentiment in sentiment_by_date.columns:
                    chart_data.add_series(series_name, sentiment_by_date[sentiment].tolist())
            
            x, y, cx, cy = Inches(6), Inches(1), Inches(3), Inches(3)
            chart = slide5.shapes.add_chart(XL_CHART_TYPE.LINE, x, y, cx, cy, chart_data).chart
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
            chart = slide5.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data).chart
            chart.has_legend = True
            chart.legend.position = XL_LEGEND_POSITION.BOTTOM
            chart.legend.font.size = Pt(10)
            
            # Set chart background and formatting
            chart.chart_style = 2  # White background
            chart.plots[0].has_major_gridlines = False
            chart.plots[0].has_minor_gridlines = False
            
            # Set diagonal labels for x-axis
            if hasattr(chart, 'category_axis'):
                chart.category_axis.tick_labels.font.size = Pt(8)
                chart.category_axis.tick_labels.orientation = 45
            
            # Set colors for bar chart
            for i, series in enumerate(chart.series):
                series.format.fill.solid()
                series.format.fill.fore_color.rgb = SENTIMENT_COLORS[list(SENTIMENT_COLORS.keys())[i]]
        
        # Sixth slide - Facebook metrics table
        logger.debug("Creating sixth slide with Facebook metrics table")
        slide6 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Set slide background
        background = slide6.background
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
            
            # Group data by author_name (which is the company name)
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
            
            table = slide6.shapes.add_table(rows, cols, left, top, width, height).table
            
            # Set column widths
            table.columns[0].width = Inches(2)  # Company
            for i in range(1, cols):
                table.columns[i].width = Inches(1.4)  # Metrics
            
            # Add headers
            headers = ['Company', 'Posts', 'Comments', 'Likes', 'Shares', 'Views']
            for i, header in enumerate(headers):
                cell = table.cell(0, i)
                cell.text = header
                cell.text_frame.paragraphs[0].font.size = Pt(12)
                cell.text_frame.paragraphs[0].font.bold = True
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(200, 200, 200)  # Light gray
            
            # Add data
            for i, row in grouped_data.iterrows():
                table.cell(i + 1, 0).text = str(row['author_name'])
                table.cell(i + 1, 1).text = str(row['post_count'])
                table.cell(i + 1, 2).text = str(row['comment_count'])
                table.cell(i + 1, 3).text = str(row['like_count'])
                table.cell(i + 1, 4).text = str(row['share_count'])
                table.cell(i + 1, 5).text = str(row['view_count'])
                
                # Set font size for all cells
                for j in range(cols):
                    cell = table.cell(i + 1, j)
                    cell.text_frame.paragraphs[0].font.size = Pt(10)
        
        logger.debug("Saving PowerPoint file")
        prs.save(output_path)
        logger.debug("PowerPoint file saved successfully")
    except Exception as e:
        logger.error(f"Error creating PowerPoint: {str(e)}")
        raise