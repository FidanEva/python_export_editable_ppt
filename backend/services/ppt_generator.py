from pptx import Presentation
from pptx.util import Inches
from pptx.chart.data import CategoryChartData
from pptx.enum.chart import XL_CHART_TYPE

def create_ppt(data_frames, image_path, custom_text, output_path):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[5])

    # Add image
    slide.shapes.add_picture(image_path, Inches(0.5), Inches(0.5), height=Inches(2))

    # Add text
    textbox = slide.shapes.add_textbox(Inches(3), Inches(0.5), Inches(5), Inches(1))
    textbox.text_frame.text = custom_text

    # Example chart from first dataframe
    df = data_frames[0]
    chart_data = CategoryChartData()
    chart_data.categories = df.iloc[:, 0].tolist()
    chart_data.add_series('Values', df.iloc[:, 1].tolist())

    x, y, cx, cy = Inches(1), Inches(3), Inches(6), Inches(3)
    slide.shapes.add_chart(XL_CHART_TYPE.COLUMN_CLUSTERED, x, y, cx, cy, chart_data)

    prs.save(output_path)
