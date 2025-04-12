import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from fontTools.ttLib import TTFont
import os

# ---
# 함수: TTF 파일에서 폰트 이름 추출
# ---
def get_font_name_from_ttf(ttf_path):
    font = TTFont(ttf_path)
    for record in font['name'].names:
        if record.nameID == 4:
            try:
                return record.string.decode('utf-16-be') if b'\x00' in record.string else record.string.decode('utf-8')
            except:
                continue
    return None

# ---
# 파워포인트 생성 함수
# ---
def crear_ppt(titulos_kr, titulos_en, letras_kr, letras_en, estilos, imagen_titulo, imagen_letra):
    prs = Presentation()
    prs.slide_width = Inches(13.33)
    prs.slide_height = Inches(7.5)

    for i in range(len(titulos_kr)):
        slide = prs.slides.add_slide(prs.slide_layouts[6])

        if imagen_titulo:
            slide.shapes.add_picture(imagen_titulo, 0, 0, prs.slide_width, prs.slide_height)
        else:
            slide.background.fill.solid()
            slide.background.fill.fore_color.rgb = RGBColor(*estilos['bg_titulo'])

        tb = slide.shapes.add_textbox(Inches(1), Inches(estilos['altura_texto']), Inches(11.33), Inches(3))
        tf = tb.text_frame
        tf.clear()
        tf.word_wrap = True

        p1 = tf.paragraphs[0]
        run1 = p1.add_run()
        run1.text = titulos_kr[i]
        run1.font.size = Pt(estilos['tamano_titulo_kr'])
        run1.font.name = estilos['font_titulo_kr']
        run1.font.color.rgb = RGBColor(*estilos['color_titulo_kr'])
        p1.alignment = PP_ALIGN.CENTER

        if titulos_en[i].strip():
            p2 = tf.add_paragraph()
            run2 = p2.add_run()
            run2.text = titulos_en[i]
            run2.font.size = Pt(estilos['tamano_titulo_en'])
            run2.font.name = estilos['font_titulo_en']
            run2.font.color.rgb = RGBColor(*estilos['color_titulo_en'])
            p2.alignment = PP_ALIGN.CENTER

        for j in range(len(letras_kr[i])):
            k_line = letras_kr[i][j]
            e_line = letras_en[i][j] if j < len(letras_en[i]) else ""

            slide = prs.slides.add_slide(prs.slide_layouts[6])

            if imagen_letra:
                slide.shapes.add_picture(imagen_letra, 0, 0, prs.slide_width, prs.slide_height)
            else:
                slide.background.fill.solid()
                slide.background.fill.fore_color.rgb = RGBColor(*estilos['bg_letra'])

            tb = slide.shapes.add_textbox(Inches(1), Inches(estilos['altura_texto']), Inches(11.33), Inches(3))
            tf = tb.text_frame
            tf.clear()
            tf.word_wrap = True

            p1 = tf.paragraphs[0]
            run1 = p1.add_run()
            run1.text = k_line
            run1.font.size = Pt(estilos['tamano_letra_kr'])
            run1.font.name = estilos['font_letra_kr']
            run1.font.color.rgb = RGBColor(*estilos['color_letra_kr'])
            p1.alignment = PP_ALIGN.CENTER

            if e_line.strip():
                p2 = tf.add_paragraph()
                run2 = p2.add_run()
                run2.text = e_line
                run2.font.size = Pt(estilos['tamano_letra_en'])
                run2.font.name = estilos['font_letra_en']
                run2.font.color.rgb = RGBColor(*estilos['color_letra_en'])
                p2.alignment = PP_ALIGN.CENTER

    return prs

# --- Streamlit UI ---
st.set_page_config(layout="wide")
st.title("마한장")

num_canciones = st.number_input("찬양양 개수", min_value=1, max_value=10, step=1)
altura_texto = st.slider("글자 위치 (0.0이 제일 높음음)", 0.0, 6.0, value=1.0, step=0.1)

font_kr_file = st.file_uploader("한국어 글자 폰트 (.ttf)", type=["ttf"], key="fkr")
font_en_file = st.file_uploader("영어 글자 폰트 (.ttf)", type=["ttf"], key="fen")

font_kr = font_en = None
if font_kr_file:
    path = "font_kr.ttf"
    with open(path, "wb") as f: f.write(font_kr_file.read())
    font_kr = get_font_name_from_ttf(path)
if font_en_file:
    path = "font_en.ttf"
    with open(path, "wb") as f: f.write(font_en_file.read())
    font_en = get_font_name_from_ttf(path)

color_titulo_kr = st.color_picker("[제목] 한국어 색상", "#000000")
color_titulo_en = st.color_picker("[제목] 영어 색상", "#666666")
bg_titulo = st.color_picker("[제목] 배경 색상", "#FFFFFF")

color_letra_kr = st.color_picker("[가사]  한국어 색상", "#FFFFFF")
color_letra_en = st.color_picker("[가사]  영어 색상", "#FFFF00")
bg_letra = st.color_picker("[가사]  배경 색상", "#000000")

# 글자 크기 설정
size_titulo_kr = st.number_input("[제목] 한국어 글자 크기", value=50)
size_titulo_en = st.number_input("[제목] 영어 글자 크기", value=36)
size_letra_kr = st.number_input("[가사]  한국어 글자 크기", value=50)
size_letra_en = st.number_input("[가사]  영어 글자 크기", value=28)

estilos = {
    'font_titulo_kr': font_kr or "Malgun Gothic",
    'font_titulo_en': font_en or "Arial",
    'color_titulo_kr': tuple(int(color_titulo_kr[i:i+2], 16) for i in (1, 3, 5)),
    'color_titulo_en': tuple(int(color_titulo_en[i:i+2], 16) for i in (1, 3, 5)),
    'bg_titulo': tuple(int(bg_titulo[i:i+2], 16) for i in (1, 3, 5)),

    'font_letra_kr': font_kr or "Malgun Gothic",
    'font_letra_en': font_en or "Arial",
    'color_letra_kr': tuple(int(color_letra_kr[i:i+2], 16) for i in (1, 3, 5)),
    'color_letra_en': tuple(int(color_letra_en[i:i+2], 16) for i in (1, 3, 5)),
    'bg_letra': tuple(int(bg_letra[i:i+2], 16) for i in (1, 3, 5)),
    'altura_texto': altura_texto,
    'tamano_titulo_kr': size_titulo_kr,
    'tamano_titulo_en': size_titulo_en,
    'tamano_letra_kr': size_letra_kr,
    'tamano_letra_en': size_letra_en,
}

imagen_titulo_file = st.file_uploader("[제목] 배경 이미지 (optional)", type=['jpg', 'png'], key="img_titulo")
imagen_letra_file = st.file_uploader("[가사]  배경 이미지 (optional)", type=['jpg', 'png'], key="img_letra")

korean_titles, english_titles, korean_lyrics, english_lyrics = [], [], [], []
for i in range(num_canciones):
    st.subheader(f"🎵 찬양 {i+1}")
    korean_titles.append(st.text_input(f"한국어 [제목] #{i+1}", key=f"kr_title_{i}"))
    english_titles.append(st.text_input(f"영어 [제목] #{i+1}", key=f"en_title_{i}"))

    kr_lyrics = st.text_area(f"한국어 [가사]  #{i+1} (줄 마다 한 슬라이드에용)", key=f"kr_lyrics_{i}")
    en_lyrics = st.text_area(f"영어 [가사]  #{i+1} (줄 마다 한 슬라이드에용, optional)", key=f"en_lyrics_{i}")

    korean_lyrics.append(kr_lyrics.split("\n"))
    english_lyrics.append(en_lyrics.split("\n"))

if st.button("🎷 PPT 생성"):
    it_path = il_path = None
    if imagen_titulo_file:
        it_path = "img_titulo.jpg"
        with open(it_path, "wb") as f: f.write(imagen_titulo_file.read())
    if imagen_letra_file:
        il_path = "img_letra.jpg"
        with open(il_path, "wb") as f: f.write(imagen_letra_file.read())

    ppt = crear_ppt(korean_titles, english_titles, korean_lyrics, english_lyrics, estilos, it_path, il_path)
    ppt_path = "ppt_generado.pptx"
    ppt.save(ppt_path)

    with open(ppt_path, "rb") as f:
        st.download_button("📥 PPT 다운로드", f, file_name=ppt_path)

    for p in [it_path, il_path, "font_kr.ttf", "font_en.ttf", ppt_path]:
        if p and os.path.exists(p):
            os.remove(p)
