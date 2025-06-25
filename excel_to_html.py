import pandas as pd
from jinja2 import Environment, FileSystemLoader
from datetime import datetime
import os

def excel_to_html(excel_path, template_path, output_path):
    """
    Reads all sheets from an Excel file and generates an HTML report
    using a Jinja2 template.
    """
    try:
        # Load the Excel file
        xls = pd.ExcelFile(excel_path)

        sheets_data = []
        # Iterate through each sheet
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name)
            # Replace NaN with empty strings for better HTML output
            df = df.fillna('')
            
            sheet_info = {
                'name': sheet_name,
                'headers': list(df.columns),
                'data': df.values.tolist()
            }
            sheets_data.append(sheet_info)

        # Set up Jinja2 environment
        template_dir = os.path.dirname(os.path.abspath(template_path))
        env = Environment(loader=FileSystemLoader(template_dir), autoescape=True)
        template = env.get_template(os.path.basename(template_path))

        # Prepare data for the template
        template_data = {
            'title': '运维中心计划管理报告',
            'sheets': sheets_data,
            'generation_date': datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        }

        # Render the template
        html_content = template.render(template_data)

        # Write the output to an HTML file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
            
        print(f"成功生成HTML报告: {os.path.abspath(output_path)}")

    except FileNotFoundError:
        print(f"错误: 未在 {excel_path} 找到文件")
    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == '__main__':
    # Define file paths relative to the script location
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    EXCEL_FILE = os.path.join(BASE_DIR, '运维中心计划管理.xlsx')
    TEMPLATE_FILE = os.path.join(BASE_DIR, 'template.html')
    OUTPUT_HTML_FILE = os.path.join(BASE_DIR, '运维中心计划管理报告.html')

    excel_to_html(EXCEL_FILE, TEMPLATE_FILE, OUTPUT_HTML_FILE)
