from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.chart.data import CategoryChartData, ChartData
from pptx.enum.chart import XL_CHART_TYPE, XL_LEGEND_POSITION
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from services.excel_parser import get_sentiment_data, get_sentiment_counts, get_company_sentiment_counts
import logging

logger = logging.getLogger(__name__)

def create_ppt(data_frames, output_path, date, company_name):
    try:
        logger.debug("Creating PowerPoint presentation")
        prs = Presentation()
        
        # First slide - Title slide with date and company name
        logger.debug("Creating title slide")
        title_slide = prs.slides.add_slide(prs.slide_layouts[0])
        title = title_slide.shapes.title
        subtitle = title_slide.placeholders[1]
        
        title.text = f"Report for {company_name}"
        subtitle.text = f"Date: {date}"
        
        # Second slide - Multiline chart and donut chart
        logger.debug("Creating second slide with charts")
        slide2 = prs.slides.add_slide(prs.slide_layouts[5])
        
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
        
        # Format line chart axes
        if hasattr(chart, 'category_axis'):
            chart.category_axis.has_major_gridlines = True
            chart.category_axis.tick_labels.font.size = Pt(10)
        if hasattr(chart, 'value_axis'):
            chart.value_axis.has_major_gridlines = True
            chart.value_axis.tick_labels.font.size = Pt(10)
        
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
        
        # Third slide - Horizontal multibar chart
        logger.debug("Creating third slide with horizontal multibar chart")
        slide3 = prs.slides.add_slide(prs.slide_layouts[5])
        
        # Get company sentiment data
        company_sentiments = get_company_sentiment_counts(combined_data)
        
        chart_data = CategoryChartData()
        chart_data.categories = company_sentiments.index.tolist()
        
        for sentiment in [1, 0, -1]:
            series_name = "Positive" if sentiment == 1 else "Neutral" if sentiment == 0 else "Negative"
            if sentiment in company_sentiments.columns:
                chart_data.add_series(series_name, company_sentiments[sentiment].tolist())
        
        x, y, cx, cy = Inches(0.5), Inches(1), Inches(9), Inches(5)
        chart = slide3.shapes.add_chart(XL_CHART_TYPE.BAR_CLUSTERED, x, y, cx, cy, chart_data).chart
        chart.has_legend = True
        chart.legend.position = XL_LEGEND_POSITION.BOTTOM
        
        # Format bar chart axes
        if hasattr(chart, 'category_axis'):
            chart.category_axis.has_major_gridlines = True
            chart.category_axis.tick_labels.font.size = Pt(10)
        if hasattr(chart, 'value_axis'):
            chart.value_axis.has_major_gridlines = True
            chart.value_axis.tick_labels.font.size = Pt(10)
        
        # Make all charts editable
        logger.debug("Making charts editable")
        for slide in prs.slides:
            for shape in slide.shapes:
                if shape.has_chart:
                    chart = shape.chart
                    chart.has_legend = True
                    chart.legend.position = XL_LEGEND_POSITION.BOTTOM
                    
                    # Make chart elements editable
                    for series in chart.series:
                        series.has_data_labels = True
                        if hasattr(series, 'format'):
                            series.format.line.width = Pt(2.5)
        
        logger.debug("Saving PowerPoint file")
        prs.save(output_path)
        logger.debug("PowerPoint file saved successfully")
    except Exception as e:
        logger.error(f"Error creating PowerPoint: {str(e)}")
        raise