import pandas as pd
from jinja2 import Environment, FileSystemLoader
from datetime import datetime, timedelta
import os

# Opt-in to the future behavior of pandas to silence the warning
pd.set_option('future.no_silent_downcasting', True)

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

            # Drop the '当前状态' column if it exists
            if '当前状态' in df.columns:
                df = df.drop(columns=['当前状态'])

            # Identify and process potential progress columns
            progress_cols = [col for col in df.columns if '进度' in col or '完成率' in col]
            for col in progress_cols:
                # Convert column to numeric, coercing errors to NaN. This handles decimal progress values.
                df[col] = pd.to_numeric(df[col], errors='coerce')

            # Convert to object type to properly handle NaT, then fill all nulls
            df = df.astype(object).fillna('')
            
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
        now = datetime.now()
        report_date = now + timedelta(days=1)

        template_data = {
            'title': '运维中心日报',
            'sheets': sheets_data,
            'generation_date': now.strftime('%Y-%m-%d %H:%M:%S'),
            'report_date': report_date.strftime('%Y-%m-%d')
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
    EXCEL_FILE = os.path.join(BASE_DIR, '运维中心日报.xlsx')
    TEMPLATE_FILE = os.path.join(BASE_DIR, 'template.html')
    OUTPUT_HTML_FILE = os.path.join(BASE_DIR, '运维中心日报.html')

    excel_to_html(EXCEL_FILE, TEMPLATE_FILE, OUTPUT_HTML_FILE)
