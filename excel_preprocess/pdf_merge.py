from PyPDF2 import PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import ImageReader
import io
from pdf2image import convert_from_path  # 需要安裝 pdf2image 和 poppler

def merge_pdfs_horizontally_reportlab(pdf1_path, pdf2_path, output_path):
    A4_WIDTH, A4_HEIGHT = A4

    # 將 PDF 轉為圖片
    pages1 = convert_from_path(pdf1_path, dpi=300)
    pages2 = convert_from_path(pdf2_path, dpi=300)
    img1 = pages1[0]
    img2 = pages2[0]
    # 順時針旋轉 90 度
    img1_rotated = img1.rotate(-90, expand=True)  # PIL rotate 順時針 90 度用 -90
    img2_rotated = img2.rotate(-90, expand=True)

    # 使用 reportlab 創建空白 A4 並繪製左右兩半
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=A4)

    # 繪製到上半部分和下半部分
    c.drawImage(ImageReader(img1_rotated), 0, A4_HEIGHT / 2, width=A4_WIDTH, height=A4_HEIGHT / 2)  # 上半頁
    c.drawImage(ImageReader(img2_rotated), 0, 0, width=A4_WIDTH, height=A4_HEIGHT / 2)

    c.showPage()
    c.save()

    # 將 reportlab 生成的 PDF 與 PyPDF2 寫出
    packet.seek(0)
    from PyPDF2 import PdfReader
    reader = PdfReader(packet)
    writer = PdfWriter()
    writer.add_page(reader.pages[0])

    with open(output_path, "wb") as f:
        writer.write(f)

if __name__ == "__main__":
    merge_pdfs_horizontally_reportlab("../Doc/pdf1.pdf", "../Doc/pdf2.pdf", "../Doc/merged.pdf")