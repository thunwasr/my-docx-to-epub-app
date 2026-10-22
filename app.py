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

# ส่วนหัวโปรแกรมและแจ้งเตือนความปลอดภัย
st.title("📖 ระบบแปลงไฟล์ DOCX เป็น EPUB")
st.write("แปลงไฟล์ Word (.docx) เป็น EPUB พร้อมปรับแต่งสไตล์และจัดสระลอยให้อัตโนมัติ")

# 🔒 กล่องเน้นย้ำความปลอดภัยของข้อมูล
st.info("🔒 **ความปลอดภัยระดับสูงสุด:** ไฟล์เอกสารและฟอนต์ทั้งหมดจะถูกประมวลผลภายในระบบชั่วคราวขณะรันการทำงานเท่านั้น **ไม่มีการบันทึกหรือส่งออกไฟล์ไปยังอินเทอร์เน็ต/ฐานข้อมูลภายนอกใดๆ ทั้งสิ้น** ปลอดภัย 100%")

# ==========================================
# ส่วน UI สำหรับ Customize การตั้งค่า (พร้อมค่า Default ตามกำหนด)
# ==========================================
st.subheader("⚙️ ปรับแต่งสไตล์ (Customize Style)")

col_font, col_size = st.columns(2)
with col_font:
    font_choice = st.selectbox(
        "ฟอนต์เริ่มต้น (Font Family)",
        options=["Niramit", "Sarabun", "Noto Sans Thai", "อัปโหลดฟอนต์เอง (Custom Font)", "ระบบมาตรฐาน (Sans-serif)"],
        index=0  # Default: Niramit
    )

with col_size:
    font_size = st.number_input("ขนาดตัวอักษร (Font Size px)", min_value=10, max_value=32, value=16, step=1) # Default: 16

col_line, col_margin = st.columns(2)
with col_line:
    line_height = st.number_input("ระยะห่างบรรทัด (Line-Height)", min_value=1.0, max_value=3.0, value=1.3, step=0.1) # Default: 1.3

with col_margin:
    margin_bottom = st.number_input("ระยะห่างย่อหน้า (Margin-Bottom em)", min_value=0.0, max_value=3.0, value=0.4, step=0.1) # Default: 0.4

# ส่วนอัปโหลดฟอนต์ส่วนตัว (จะแสดงผลเมื่อผู้ใช้เลือก "อัปโหลดฟอนต์เอง")
custom_uploaded_fonts = {}
if font_choice == "อัปโหลดฟอนต์เอง (Custom Font)":
    st.markdown("---")
    st.write("📁 **อัปโหลดไฟล์ฟอนต์ส่วนตัว (.ttf หรือ .otf):**")
    cf1, cf2 = st.columns(2)
    with cf1:
        custom_uploaded_fonts["Regular"] = st.file_uploader("1. ตัวปกติ (Regular)", type=["ttf", "otf"], key="font_reg")
        custom_uploaded_fonts["Bold"] = st.file_uploader("3. ตัวหนา (Bold)", type=["ttf", "otf"], key="font_bold")
    with cf2:
        custom_uploaded_fonts["Italic"] = st.file_uploader("2. ตัวเอียง (Italic)", type=["ttf", "otf"], key="font_italic")
        custom_uploaded_fonts["BoldItalic"] = st.file_uploader("4. ตัวหนาเอียง (BoldItalic)", type=["ttf", "otf"], key="font_bi")

# ==========================================
# ฟังก์ชันแปลงไฟล์ EPUB
# ==========================================
def convert_docx_to_epub(docx_path, output_path, font_choice, font_size, line_height, margin_bottom, custom_fonts_dict):
    try:
        doc = Document(docx_path)
        book = epub.EpubBook()
        
        base_name = os.path.splitext(os.path.basename(docx_path))[0]
        book.set_title(base_name)
        book.set_language('th')

        font_faces = ""
        font_family_css = "sans-serif"
        fonts_dir = "fonts"

        styles_config = [
            ("Regular", "normal", "normal"),
            ("Italic", "normal", "italic"),
            ("Bold", "bold", "normal"),
            ("BoldItalic", "bold", "italic")
        ]

        # กรณีที่ 1: ผู้ใช้อัปโหลดฟอนต์มาเอง
        if font_choice == "อัปโหลดฟอนต์เอง (Custom Font)":
            has_added_any_font = False
            for style_key, weight, style in styles_config:
                font_file_obj = custom_fonts_dict.get(style_key)
                if font_file_obj is not None:
                    font_bytes = font_file_obj.read()
                    font_ext = font_file_obj.name.split('.')[-1].lower()
                    font_filename = f"custom_{style_key.lower()}.{font_ext}"
                    
                    font_item = epub.EpubItem(
                        uid=f"font_custom_{style_key.lower()}",
                        file_name=f"fonts/{font_filename}",
                        media_type=f"font/{font_ext}",
                        content=font_bytes
                    )
                    book.add_item(font_item)

                    font_faces += f"""
                    @font-face {{
                        font-family: 'CustomBookFont';
                        src: url('../fonts/{font_filename}');
                        font-weight: {weight};
                        font-style: {style};
                    }}
                    """
                    has_added_any_font = True
            
            if has_added_any_font:
                font_family_css = "'CustomBookFont', sans-serif"

        # กรณีที่ 2: เลือกฟอนต์มาตรฐานที่เราเตรียมไว้ในโฟลเดอร์ fonts/
        else:
            font_family_mapping = {
                "Niramit": {
                    "Regular": "Niramit-Regular.ttf",
                    "Italic": "Niramit-Italic.ttf",
                    "Bold": "Niramit-Bold.ttf",
                    "BoldItalic": "Niramit-BoldItalic.ttf"
                },
                "Sarabun": {
                    "Regular": "Sarabun-Regular.ttf",
                    "Italic": "Sarabun-Italic.ttf",
                    "Bold": "Sarabun-Bold.ttf",
                    "BoldItalic": "Sarabun-BoldItalic.ttf"
                },
                "Noto Sans Thai": {
                    "Regular": "NotoSansThai-Regular.ttf",
                    "Italic": "NotoSansThai-Italic.ttf",
                    "Bold": "NotoSansThai-Bold.ttf",
                    "BoldItalic": "NotoSansThai-BoldItalic.ttf"
                }
            }

            selected_font_set = font_family_mapping.get(font_choice)
            if selected_font_set:
                has_added_any_font = False
                for style_key, weight, style in styles_config:
                    font_filename = selected_font_set.get(style_key)
                    if font_filename:
                        font_path = os.path.join(fonts_dir, font_filename)
                        if os.path.exists(font_path):
                            font_ext = font_filename.split('.')[-1].lower()
                            with open(font_path, 'rb') as f:
                                font_data = f.read()
                            
                            font_item = epub.EpubItem(
                                uid=f"font_{style_key.lower()}",
                                file_name=f"fonts/{font_filename}",
                                media_type=f"font/{font_ext}",
                                content=font_data
                            )
                            book.add_item(font_item)

                            font_faces += f"""
                            @font-face {{
                                font-family: 'CustomBookFont';
                                src: url('../fonts/{font_filename}');
                                font-weight: {weight};
                                font-style: {style};
                            }}
                            """
                            has_added_any_font = True
                
                if has_added_any_font:
                    font_family_css = "'CustomBookFont', sans-serif"

        # สร้าง CSS หลักบังคับสไตล์ด้วย !important
        css_content = font_faces + f"""
        body {{ 
            font-family: {font_family_css} !important; 
            font-size: {font_size}px !important;
            font-feature-settings: "liga" 1, "ccmp" 1;
            -webkit-font-feature-settings: "liga" 1, "ccmp" 1;
            text-rendering: optimizeLegibility;
        }}
        p {{ 
            text-align: left !important; 
            line-height: {line_height} !important; 
            margin-top: 0 !important; 
            margin-bottom: {margin_bottom}em !important;
            padding-top: 0 !important;
            padding-bottom: 0 !important;
            font-size: {font_size}px !important;
            word-break: normal;
            line-break: strict;
        }}
        h1 {{ text-align: center; margin-top: 2em; margin-bottom: 1em; }}
        img {{ max-width: 100%; height: auto; }}
        .bullet-point {{ margin-left: 2em; text-indent: -1em; }}
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
# ส่วนอัปโหลดและประมวลผล
# ==========================================
st.divider()
uploaded_file = st.file_uploader("เลือกไฟล์ DOCX ที่ต้องการแปลง", type=["docx"])

if uploaded_file is not None:
    temp_docx_path = f"temp_{uploaded_file.name}"
    output_epub_path = uploaded_file.name.replace(".docx", ".epub").replace(".DOCX", ".epub")
    
    with open(temp_docx_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    if st.button("🚀 เริ่มแปลงไฟล์", type="primary", use_container_width=True):
        with st.spinner("กำลังประมวลผลจัดสระลอยและฝังรูปแบบสไตล์..."):
            success = convert_docx_to_epub(
                temp_docx_path, 
                output_epub_path,
                font_choice,
                font_size,
                line_height,
                margin_bottom,
                custom_uploaded_fonts
            )
            
            if success and os.path.exists(output_epub_path):
                st.balloons() # แสดงเอฟเฟกต์แสดงความยินดี
                st.success("🎉 แปลงไฟล์สำเร็จแล้ว! กดปุ่มด้านล่างเพื่อรับไฟล์ได้ทันที")
                
                # อ่านไฟล์เตรียมให้กดดาวน์โหลด
                with open(output_epub_path, "rb") as file:
                    st.download_button(
                        label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ EPUB",
                        data=file,
                        file_name=os.path.basename(output_epub_path),
                        mime="application/epub+zip",
                        type="primary",
                        use_container_width=True
                    )
                
                # ลบไฟล์ชั่วคราวทิ้งทันทีเพื่อความปลอดภัย
                if os.path.exists(temp_docx_path): 
                    os.remove(temp_docx_path)

# ==========================================
# ส่วนอัปโหลด ประมวลผล และปุ่มล้างข้อมูล
# ==========================================
st.divider()
uploaded_file = st.file_uploader("เลือกไฟล์ DOCX ที่ต้องการแปลง", type=["docx"], key="docx_uploader")

if uploaded_file is not None:
    temp_docx_path = f"temp_{uploaded_file.name}"
    output_epub_path = uploaded_file.name.replace(".docx", ".epub").replace(".DOCX", ".epub")
    
    with open(temp_docx_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    # สร้างปุ่ม 2 ปุ่มคู่กัน (ปุ่มเริ่มแปลงไฟล์ และ ปุ่มเริ่มต้นใหม่)
    col_start, col_reset = st.columns([2, 1])
    
    with col_start:
        start_btn = st.button("🚀 เริ่มแปลงไฟล์", type="primary", use_container_width=True)
        
    with col_reset:
        # ปุ่ม Reset สำหรับล้างไฟล์และเริ่มต้นใหม่
        reset_btn = st.button("🔄 ล้างข้อมูล / เริ่มต้นใหม่", type="secondary", use_container_width=True)

    if reset_btn:
        # ลบไฟล์ชั่วคราวค้างในระบบ (ถ้ามี) แล้วสั่งรีโหลดหน้าเว็บทันที
        if os.path.exists(temp_docx_path): 
            os.remove(temp_docx_path)
        if os.path.exists(output_epub_path): 
            os.remove(output_epub_path)
        st.rerun()

    if start_btn:
        with st.spinner("กำลังประมวลผลจัดสระลอยและฝังรูปแบบสไตล์..."):
            success = convert_docx_to_epub(
                temp_docx_path, 
                output_epub_path,
                font_choice,
                font_size,
                line_height,
                margin_bottom,
                custom_uploaded_fonts
            )
            
            if success and os.path.exists(output_epub_path):
                st.balloons()
                st.success("🎉 แปลงไฟล์สำเร็จแล้ว! กดปุ่มด้านล่างเพื่อรับไฟล์ได้ทันที")
                
                with open(output_epub_path, "rb") as file:
                    st.download_button(
                        label="📥 คลิกที่นี่เพื่อดาวน์โหลดไฟล์ EPUB",
                        data=file,
                        file_name=os.path.basename(output_epub_path),
                        mime="application/epub+zip",
                        type="primary",
                        use_container_width=True
                    )
                
                if os.path.exists(temp_docx_path): 
                    os.remove(temp_docx_path)
