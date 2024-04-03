import fitz  # PyMuPDF
import pytesseract
from PIL import Image
import io

def create_ocr_pdf(input_pdf_path, output_pdf_path):
    # PDF 열기
    pdf = fitz.open(input_pdf_path)
    ocr_pdf = fitz.open()  # 검색 가능한 PDF를 위한 새 문서

    for page_num in range(len(pdf)):
        # 페이지를 이미지로 변환
        page = pdf.load_page(page_num)
        pix = page.get_pixmap()
        img_bytes = pix.tobytes("png")
        img = Image.open(io.BytesIO(img_bytes))
        
        # 이미지에서 텍스트 추출 (OCR)
        text = pytesseract.image_to_string(img, lang='eng')
        
        # 추출된 텍스트와 원본 이미지를 사용해 새 페이지 생성
        ocr_page = ocr_pdf.new_page(pno=-1, width=pix.width, height=pix.height)
        ocr_page.insert_image(ocr_page.rect, stream=img_bytes)  # 원본 이미지 삽입
        ocr_page.insert_text((0, 0), text, fontsize=1, overlay=True)  # 텍스트 레이어 추가 (보이지 않게 처리)

    # 수정된 PDF 저장
    ocr_pdf.save(output_pdf_path)
    ocr_pdf.close()
    pdf.close()

## create_ocr_pdf('input.pdf', 'output_searchable.pdf')

