import pandas as pd
import os

def excel_to_markdown(excel_path, output_path):
    """
    Reads all sheets from an Excel file and writes them as markdown tables to a file.
    """
    try:
        xls = pd.ExcelFile(excel_path)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(f"# Excel 文件内容: {os.path.basename(excel_path)}\n\n")

            for sheet_name in xls.sheet_names:
                f.write(f"## 工作表: {sheet_name}\n")
                df = pd.read_excel(xls, sheet_name=sheet_name)
                # Replace NaN with empty strings for better markdown output
                df = df.fillna('')
                f.write(df.to_markdown(index=False))
                f.write("\n\n")
        
        print(f"成功将Excel内容导出到Markdown文件: {os.path.abspath(output_path)}")

    except FileNotFoundError:
        print(f"错误: 未在 {excel_path} 找到文件")
    except Exception as e:
        print(f"发生错误: {e}")

if __name__ == '__main__':
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
    EXCEL_FILE = os.path.join(BASE_DIR, '运维中心日报.xlsx')
    OUTPUT_MD_FILE = os.path.join(BASE_DIR, '运维中心日报_view.md')
    excel_to_markdown(EXCEL_FILE, OUTPUT_MD_FILE)
