import re
import tkinter as tk
from pathlib import Path
from tkinter import filedialog
from PyPDFForm import PdfWrapper

# 全局配置参数
DEFAULT_INITIAL_DIR = Path("/Volumes/SSD 1TB/GEhealthcare/202508")

PDF_TEMPLATE_PATH = Path("/Volumes/SSD 1TB/GEhealthcare/202509/checklist/SafetyTest.pdf")
PDF_OUTPUT_PATH = Path("/Volumes/SSD 1TB/GEhealthcare/202509/finish/Output/Checklist_finish.pdf")

# 需填写的参数
ValueKEYS = {
    'Asset': '112233',
    'Hospital': 'KWH',
    'Location': 'KWH-MB-14-W14B2',
    'CheckedBy': 'Koen.CAI',
    'CheckedDate': '01-Sep-2025',
    'Tester': 'SA-2010S',
    'TesterSerial': '73385039',
    'TesterCalDay': '30-Dec-2026',
    'ClassType': '2',
    'ModelType': 'BF',
    'Mechanism': '0',
    'DC': '0',
    'Battery': '0',
    'Voltage': '219'
}
# pdf里对应的框
PDF_FIELD_MAPPING = {
    "Text1": "Hospital",
    "Text2": "Location",
    "Text3": "Tester",
    "Text4": "Asset",
    "Text5": "CheckedBy",
    "Text6": "TesterSerial",
    "Text8": "CheckedDate",
    "Text9": "TesterCalDay",
    "Text10": "Voltage",
    "1": "ClassType",  # Checkbox for Class I
    "2": "ClassType",  # Checkbox for Class II
    "4": "ModelType",  # Checkbox for B
    "5": "ModelType",  # Checkbox for BF
    "6": "ModelType"  # Checkbox for CF
}


def fill_pdf(keys):
    # 初始化 PdfWrapper，若需要 Adobe Acrobat 兼容可设置 adobe_mode=True
    wrapper = PdfWrapper(str(PDF_TEMPLATE_PATH), adobe_mode=True)

    # 构建填充数据：文本字段使用字符串，复选框使用布尔值
    data = {}
    for field, key in PDF_FIELD_MAPPING.items():
        if field.isdigit():
            # 处理复选框字段
            if key == 'ClassType':
                data[field] = (keys.get('ClassType') == field)
            elif key == 'ModelType':
                # 将字段号映射到 ModelType 值
                model_map = {'4': 'B', '5': 'BF', '6': 'CF'}
                data[field] = (keys.get('ModelType') == model_map.get(field))
        else:
            # 处理文本字段
            data[field] = str(keys.get(key, ""))

    # 填充表单并写出PDF
    filled = wrapper.fill(data, flatten=False)
    filled.write(str(PDF_OUTPUT_PATH))
    print(f"PDF form filled and saved to {PDF_OUTPUT_PATH}")


# 弹出窗口选择目标文件
def find_excel_file():
    # 测试用 - 取消注释以下三行用于测试
    # file_path = "/Volumes/SSD 1TB/GEhealthcare/202508/Preventive-Pending Report.xlsx"
    # return file_path
    # =====================================================

    # 正式用
    """使用 tkinter 选择文件"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 创建文件选择对话框
    file_path = filedialog.askopenfilename(
        title="选择总表文件",
        filetypes=[
            ("Excel 文件", "*.xlsx *.xls"),
            ("所有文件", "*.*")
        ],
        initialdir=DEFAULT_INITIAL_DIR  # 初始目录
    )
    root.destroy()  # 关闭 tkinter 窗口
    if not file_path:
        print("未选择文件，操作取消")
        return None
    return file_path


def clean_filename(name):
    """清理文件名中的无效字符"""
    if not isinstance(name, str):
        name = str(name)
    # 移除特殊字符，只保留字母、数字、中文、下划线和短横线
    return re.sub(r'[\\/*?:"<>|]', "", name).strip()


if __name__ == "__main__":
    fill_pdf(ValueKEYS)
