import html
import os
import re
import traceback
import unicodedata
import streamlit as st
from docx import Document
from docx.text.run import Run
from ebooklib import epub
from lxml import etree

# ตั้งค่าหน้าเว็บ Streamlit
st.set_page_config(page_title="DOCX to EPUB Converter", layout="centered", page_icon="📖")

st.title("📖 ระบบแปลงไฟล์ DOCX เป็น EPUB")
st.write("อัปโหลดไฟล์ Word (.docx) เพื่อแปลงเป็น EPUB พร้อมระบบจัดสระลอยและฝังรูปภาพอัตโนมัติ")

# ==========================================
# ฟังก์ชันแปลงไฟล์ (ปรับแก้โครงสร้าง XHTML ให้ถูกต้องแล้ว)
# ==========================================
def convert_docx_to_epub(docx_path, output_path):
    try:
        doc = Document(docx_path)
        book = epub.EpubBook()
        
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        book.set_title(base_name)
        book.set_language('th')

        # CSS สำหรับจัดรูปแบบหนังสือ
        css_content = """
        body { 
            font-family: sans-serif; 
            font-feature-settings: "liga" 1, "ccmp" 1;
            -webkit-font-feature-settings: "liga" 1, "ccmp" 1;
            text-rendering: optimizeLegibility;
        }
        p { 
            text-align: left !important; 
            line-height: 1.3 !important; 
            margin-top: 0 !important; 
            margin-bottom: 0.4em !important;
            padding-top: 0 !important;
            padding-bottom: 0 !important;
            word-break: normal;
            line-break: strict;
        }
        h1 { text-align: center; margin-top: 2em; margin-bottom: 1em; }
        img { max-width: 100%; height: auto; }
        .bullet-point { margin-left: 2em; text-indent: -1em; }
        """

        style_item = epub.EpubItem(
            uid="style_main",
            file_name="style/main.css",
            media_type="text/css",
            content=css_content.encode('utf-8')
        )
        book.add_item(style_item)

        # การจัดการรูปภาพ
        images_in_book = {}
        image_counter = 1
        def get_or_add_image(rel_id, doc_part):
            nonlocal image_counter
            if rel_id in images_in_book: 
                return images_in_book[rel_id]
            img_part = doc_part.related_parts[rel_id]
            ext = os.path.splitext(img_part.partname)[1][1:].lower() or 'png'
            img_filename = f"images/img_{image_counter}.{ext}"
            epub_img = epub.EpubItem(
                uid=f"img_{image_counter}", 
                file_name=img_filename,
                media_type=f"image/{ext if ext != 'jpg' else 'jpeg'}", 
                content=img_part.blob
            )
            book.add_item(epub_img)
            images_in_book[rel_id] = img_filename
            image_counter += 1
            return img_filename

        # การจัดการหน้าปก
        cover_page = epub.EpubHtml(title='Cover', file_name='cover_page.xhtml', lang='th')
        if doc.inline_shapes:
            first_shape = doc.inline_shapes[0]
            if first_shape._inline.graphic.graphicData.pic is not None:
                rel_id = first_shape._inline.graphic.graphicData.pic.blipFill.blip.embed
                img_part = doc.part.related_parts[rel_id]
                book.set_cover("cover.png", img_part.blob)

        # แก้ไข 1: ไม่ครอบ <html><body> และใส่ alt
        cover_page.content = '<div style="text-align:center;"><img src="cover.png" alt="Cover" style="height:100%; max-width:100%;"/></div>'
        cover_page.add_item(style_item)
        book.add_item(cover_page)

        # การจัดการเนื้อหา
        chapters = []
        current_chapter = None
        content_html = ""

        for para in doc.paragraphs:
            if para.style.name.startswith('Heading'):
                if current_chapter:
                    # แก้ไข 2: ตัด <html><body> ออก ส่งเฉพาะ HTML Content
                    current_chapter.content = content_html
                    current_chapter.add_item(style_item)
                    book.add_item(current_chapter)
                    chapters.append(current_chapter)
                
                h_text = para.text.strip()
                h_text = re.sub(r'\s+([\u0e31\u0e34-\u0e3a\u0e47-\u0e4e])', r'\1', h_text)
                chapter_title = h_text or f"Chapter {len(chapters)+1}"
                
                current_chapter = epub.EpubHtml(title=chapter_title, file_name=f'chap_{len(chapters)+1}.xhtml', lang='th')
                content_html = f'<h1>{chapter_title}</h1>'
            else:
                para_html = ""
                is_bullet = False
                if 'List' in para.style.name or para._element.xpath('./w:pPr/w:numPr'):
                    is_bullet = True

                # อัลกอริทึมจัดการ Run และสระลอย (ตามโค้ดเดิมของคุณ)
                para_runs = [Run(r, para) for r in para._element.xpath('.//w:r')]
                run_texts = [run.text if run.text else "" for run in para_runs]
                
                for i in range(len(run_texts)):
                    if run_texts[i]:
                        run_texts[i] = re.sub(r'\s+([\u0e31\u0e34-\u0e3a\u0e47-\u0e4e])', r'\1', run_texts[i])
                        m = re.match(r'^([\u0e31\u0e34-\u0e3a\u0e47-\u0e4e]+)(.*)', run_texts[i])
                        if m:
                            combining_marks = m.group(1)
                            run_texts[i] = m.group(2) 
                            for j in range(i-1, -1, -1):
                                if run_texts[j].strip(): 
                                    run_texts[j] = run_texts[j].rstrip() + combining_marks
                                    break
                                else:
                                    run_texts[j] = ""

                for idx, run in enumerate(para_runs):
                    run_text = run_texts[idx]
                    run_xml = etree.tostring(run._element, encoding='unicode')
                    
                    if 'drawing' in run_xml:
                        match = re.search(r'embed="([^"]+)"', run_xml)
                        if match:
                            try:
                                img_url = get_or_add_image(match.group(1), doc.part)
                                para_html += f'<div style="text-align:center;"><img src="{img_url}" alt="Image"/></div>'
                            except: 
                                pass

                    if run_text:
                        run_text = unicodedata.normalize('NFC', run_text)
                        run_text = run_text.replace('\u25cc', '')
                        processed_text = run_text.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;').replace('\n', '<br/>')
                        
                        if run.bold: processed_text = f"<b>{processed_text}</b>"
                        if run.italic: processed_text = f"<i>{processed_text}</i>"
                        if run.underline: processed_text = f'<span style="text-decoration: underline;">{processed_text}</span>'
                        if run.font.name and "Light" in run.font.name:
                            processed_text = f'<span style="font-weight: 300;">{processed_text}</span>'
                            
                        para_html += processed_text

                if para_html:
                    if not current_chapter:
                        current_chapter = epub.EpubHtml(title='Intro', file_name='intro.xhtml', lang='th')
                        content_html = ""
                    
                    if is_bullet:
                        content_html += f'<p class="bullet-point">&#8226; {para_html}</p>'
                    else:
                        content_html += f'<p>{para_html}</p>'

        if current_chapter:
            # แก้ไข 3: ตัด <html><body> ออก
            current_chapter.content = content_html
            current_chapter.add_item(style_item)
            book.add_item(current_chapter)
            chapters.append(current_chapter)

        # จัดลำดับเล่ม
        book.toc = tuple(chapters)
        book.add_item(epub.EpubNcx())
        book.add_item(epub.EpubNav())
        book.spine = [cover_page] + chapters + ['nav']
        
        epub.write_epub(output_path, book, {})
        return True
    except Exception as e:
        st.error(f"เกิดข้อผิดพลาด: {e}")
        traceback.print_exc()
        return False

# ==========================================
# ส่วนประกอบหน้า UI (Streamlit UI)
# ==========================================
uploaded_file = st.file_uploader("เลือกไฟล์ DOCX ที่ต้องการแปลง", type=["docx"])

if uploaded_file is not None:
    # สร้างไฟล์ชั่วคราวบนเซิร์ฟเวอร์
    temp_docx_path = f"temp_{uploaded_file.name}"
    output_epub_path = uploaded_file.name.replace(".docx", ".epub").replace(".DOCX", ".epub")
    
    with open(temp_docx_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    col1, col2 = st.columns([1, 1])
    
    with col1:
        if st.button("🚀 เริ่มแปลงไฟล์", type="primary", use_container_width=True):
            with st.spinner("กำลังแปลงไฟล์และจัดระเบียบโครงสร้าง..."):
                success = convert_docx_to_epub(temp_docx_path, output_epub_path)
                
                if success and os.path.exists(output_epub_path):
                    st.success("แปลงไฟล์สำเร็จ!")
                    with open(output_epub_path, "rb") as file:
                        st.download_button(
                            label="⬇️ ดาวน์โหลดไฟล์ EPUB",
                            data=file,
                            file_name=os.path.basename(output_epub_path),
                            mime="application/epub+zip",
                            use_container_width=True
                        )
                    # ลบไฟล์ชั่วคราว
                    if os.path.exists(temp_docx_path): os.remove(temp_docx_path)