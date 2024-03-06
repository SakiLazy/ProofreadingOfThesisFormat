from docx import Document
import tkinter as tk
from tkinter import filedialog
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn


def select_word_document():
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="选择一个文档", filetypes=[("Word", "*.docx")])
    return file_path
    # 选择文档打开


def custom_save_dialog(original_path):
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.asksaveasfilename(defaultextension=".docx", initialdir=original_path.rsplit('/', 1)[0], title="另存为", filetypes=[("Word", "*.docx")])
    return file_path
    # 自定义文档保存


def update_headers_if_text_exists(doc, header_text):
    for section in doc.sections:
        # 检查默认页眉中是否有文本
        if any(paragraph.text.strip() for paragraph in section.header.paragraphs):
            clear_and_set_new_header(section.header, header_text)

        # 如果首页面页眉与默认页眉不同，且其中有文本，则更新
        if not section.first_page_header.is_linked_to_previous:
            if any(paragraph.text.strip() for paragraph in section.first_page_header.paragraphs):
                clear_and_set_new_header(section.first_page_header, header_text)

        # 检查偶数页页眉，仅当它们与默认页眉不同，且其中有文本时更新
        if section.even_page_header and not section.even_page_header.is_linked_to_previous:
            if any(paragraph.text.strip() for paragraph in section.even_page_header.paragraphs):
                clear_and_set_new_header(section.even_page_header, header_text)


def clear_and_set_new_header(header, text):
    for paragraph in header.paragraphs:
        paragraph.clear()
    new_paragraph = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    run = new_paragraph.add_run(text)
    run.font.name = '宋体'
    run._element.rPr.rFonts.set(qn('w:eastAsia'), '宋体')
    run.font.size = Pt(10.5)  # 五号字体大小
    new_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER     # 居中


def adjust_margins(doc):
    for section in doc.sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(1)
        section.header_distance = Cm(2)
        section.footer_distance = Cm(1)
        # 页眉页脚间距



if __name__ == "__main__":
    selected_doc_path = select_word_document()
    if selected_doc_path:
        doc = Document(selected_doc_path)
        header_text = "杭州电子科技大学信息工程学院本科毕业设计"
        update_headers_if_text_exists(doc, header_text)
        adjust_margins(doc)
        new_doc_path = custom_save_dialog(selected_doc_path)
        if new_doc_path:
            doc.save(new_doc_path)
            print(f"Document saved as {new_doc_path}")
        else:
            print("Document save cancelled.")
    else:
        print("No document selected or process cancelled.")
