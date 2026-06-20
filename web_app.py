import streamlit as st
from docxtpl import DocxTemplate, RichText
from docx.shared import Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os

# Название файла шаблона
TEMPLATE_NAME = "образец отчета1.docx"

st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")

st.title("🚗 Рабочее место оценщика")
st.markdown("Заполните данные ниже для автоматической генерации отчета об оценке ущерба.")

if os.path.exists(TEMPLATE_NAME):
    st.success(f"✅ Базовый шаблон отчета (`{TEMPLATE_NAME}`) успешно подключен автоматически.")
    template_source = TEMPLATE_NAME
else:
    st.warning(f"⚠️ Файл `{TEMPLATE_NAME}` не найден. Загрузите его вручную ниже:")
    template_source = st.file_uploader("Загрузите шаблон отчета", type="docx")

st.header("1. Ввод данных")

col1, col2 = st.columns(2)

with col1:
    st.subheader("Общие данные")
    report_num = st.text_input("Номер отчета:")
    contract_num = st.text_input("Номер договора:")
    date = st.text_input("Дата оценки:")
    customer = st.text_input("ФИО Заказчика:")
    address = st.text_input("Адрес регистрации:")
    sum_num = st.text_input("Сумма цифрами:")
    sum_words = st.text_input("Сумма прописью:")

with col2:
    st.subheader("Данные автомобиля")
    car_model = st.text_input("Марка, модель:")
    reg_num = st.text_input("Гос. номер:")
    vin = st.text_input("VIN код:")
    tech_passport = st.text_input("Тех. паспорт №:")
    year = st.text_input("Год выпуска:")
    engine_vol = st.text_input("Объем ДВС:")
    color = st.text_input("Цвет кузова:")
    
    col_inner1, col_inner2 = st.columns(2)
    with col_inner1:
        body_type = st.text_input("Тип кузова:")
    with col_inner2:
        steering = st.selectbox("Положение руля:", ["Левый руль", "Правый руль"])

st.header("2. Описание повреждений и ремонта")
damage_desc = st.text_area("Характеристика повреждений (при осмотре установлено):", height=100)
repair_desc = st.text_area("Требуемый ремонт (для восстановления требуется):", height=100)

# Функция для сохранения переносов строк БЕЗ добавления маркеров (точек)
def format_text_with_newlines(text):
    rt = RichText()
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    for i, line in enumerate(lines):
        if i < len(lines) - 1:
            rt.add(line + '\n')
        else:
            rt.add(line)
    return rt

st.header("3. Фотографии автомобиля")

CAPTION_OPTIONS = [
    "Вид спереди", "Вид сзади", "Вид слева", "Вид справа",
    "VIN-код", "Показания одометра", "Обзорный снимок салона",
    "вмятина", "вытяжение", "гофра", "деформация", "изгиб", 
    "повреждение ЛКП", "потертость", "разрушение", "разрыв", 
    "раскол", "складка", "царапина", "скол", "перекос",
    "без образования видимых складок", "с глубокой вытяжкой металла", 
    "с нарушением геометрии кромок", "с нарушением геометрии ребер жесткости", 
    "с незначительной вытяжкой металла", "с образованием незначительных складок", 
    "с образованием острых складок", "с образованием плавных складок", "с образованием трещин",
    "на площади менее 10%", "на площади от 10 до 20%", 
    "на площади от 20 до 30%", "на площади от 30 до 40%", "на площади более 40%"
]

# Инициализация состояния
if "photo_order" not in st.session_state:
    st.session_state.photo_order = []
if "last_files" not in st.session_state:
    st.session_state.last_files = []
if "deleted_files" not in st.session_state:
    st.session_state.deleted_files = set()

# --- КОЛБЭК ДЛЯ ПЕРЕМЕЩЕНИЯ ---
def move_photo(file_name, current_index):
    new_val = st.session_state.get(f"pos_{file_name}")
    if new_val is None:
        return
        
    new_index = new_val - 1
    
    if new_index != current_index:
        item = st.session_state.photo_order.pop(current_index)
        st.session_state.photo_order.insert(new_index, item)
        
        for idx, fname in enumerate(st.session_state.photo_order):
            st.session_state[f"pos_{fname}"] = idx + 1

# --- КОЛБЭК ДЛЯ УДАЛЕНИЯ ---
def delete_photo(filename):
    if filename in st.session_state.photo_order:
        st.session_state.photo_order.remove(filename)
        st.session_state.deleted_files.add(filename)
        
        for idx, fname in enumerate(st.session_state.photo_order):
            st.session_state[f"pos_{fname}"] = idx + 1
# ----------------------------------------

uploaded_photos = st.file_uploader("Загрузите фотографии", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
photo_data_list = []

if uploaded_photos:
    current_filenames = [f.name for f in uploaded_photos]
    
    if set(current_filenames) != set(st.session_state.last_files):
        new_files = [f for f in current_filenames if f not in st.session_state.last_files]
        removed_files = [f for f in st.session_state.last_files if f not in current_filenames]
        
        for f in new_files:
            if f in st.session_state.deleted_files:
                st.session_state.deleted_files.remove(f)
            if f not in st.session_state.photo_order and f not in st.session_state.deleted_files:
                st.session_state.photo_order.append(f)
                
        for f in removed_files:
            if f in st.session_state.photo_order:
                st.session_state.photo_order.remove(f)
            if f in st.session_state.deleted_files:
                st.session_state.deleted_files.remove(f)
                
        st.session_state.last_files = current_filenames
        
        for idx, fname in enumerate(st.session_state.photo_order):
            key = f"pos_{fname}"
            if key not in st.session_state:
                st.session_state[key] = idx + 1

    file_dict = {f.name: f for f in uploaded_photos}

    st.write("Настройте подписи и порядок фотографий:")
    
    for i, filename in enumerate(list(st.session_state.photo_order)):
        if filename not in file_dict:
            continue
            
        photo = file_dict[filename]
        
        with st.container():
            c_img, c_controls = st.columns([1, 2])
            
            with c_img:
                st.image(photo, use_container_width=True)
                
                st.selectbox(
                    "📍 Позиция в отчете:",
                    options=list(range(1, len(st.session_state.photo_order) + 1)),
                    key=f"pos_{filename}",
                    on_change=move_photo,
                    args=(filename, i)
                )
                
                st.button("❌ Исключить фото", key=f"del_{filename}", on_click=delete_photo, args=(filename,))

            with c_controls:
                selected_tags = st.multiselect("Шаблонные фразы:", CAPTION_OPTIONS, key=f"tags_{filename}")
                custom_text = st.text_input("Свой текст (дополнит или заменит шаблон):", key=f"custom_{filename}")
                
                final_caption_parts = []
                if selected_tags:
                    final_caption_parts.append(", ".join(selected_tags))
                if custom_text:
                    final_caption_parts.append(custom_text)
                    
                final_caption = ", ".join(final_caption_parts)
                
                if final_caption:
                    st.caption(f"📝 Итоговая подпись в отчете: **{final_caption}**")
                    
            st.divider() 
            
            photo_data_list.append({"file": photo, "caption": final_caption})

if template_source is not None:
    if st.button("СГЕНЕРИРОВАТЬ ОТЧЕТ", type="primary"):
        try:
            doc = DocxTemplate(template_source)
            subdoc = doc.new_subdoc()
            
            if photo_data_list:
                table = subdoc.add_table(rows=0, cols=2)
                table.style = 'Table Grid'
                
                for i in range(0, len(photo_data_list), 2):
                    cells = table.add_row().cells
                    
                    img1_file = photo_data_list[i]["file"]
                    img1_file.seek(0)
                    
                    p1 = cells[0].paragraphs[0]
                    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run1 = p1.add_run()
                    run1.add_picture(img1_file, width=Mm(80))
                    run1.add_break()
                    run1.add_text(photo_data_list[i]["caption"])
                    
                    p2 = cells[1].paragraphs[0]
                    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    if i + 1 < len(photo_data_list):
                        img2_file = photo_data_list[i+1]["file"]
                        img2_file.seek(0)
                        
                        run2 = p2.add_run()
                        run2.add_picture(img2_file, width=Mm(80))
                        run2.add_break()
                        run2.add_text(photo_data_list[i+1]["caption"])
            
            context = {
                "REPORT_NUM": report_num,
                "CONTRACT_NUM": contract_num,
                "DATE": date,
                "CUSTOMER_NAME": customer,
                "ADDRESS": address,
                "CAR_MODEL": car_model,
                "REG_NUM": reg_num,
                "VIN": vin,
                "TECH_PASSPORT": tech_passport,
                "YEAR": year,
                "ENGINE_VOL": engine_vol,
                "COLOR": color,
                "BODY_TYPE": body_type,
                "STEERING": steering,
                "TOTAL_SUM_NUM": sum_num,
                "TOTAL_SUM_WORDS": sum_words,
                "DAMAGE_DESC": format_text_with_newlines(damage_desc), # Убрали функцию с маркерами
                "REPAIR_DESC": format_text_with_newlines(repair_desc), # Убрали функцию с маркерами
                "PHOTO_TABLE": subdoc 
            }
            
            doc.render(context)
            
            settings = doc.docx.settings.element
            update_fields = OxmlElement('w:updateFields')
            update_fields.set(qn('w:val'), 'true')
            settings.append(update_fields)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✅ Отчет успешно сгенерирован!")
            
            # --- НОВАЯ ЛОГИКА СОХРАНЕНИЯ ФАЙЛА ---
            # Сохраняем под госномером (если пустой — "Без_номера.docx")
            safe_reg_num = reg_num.strip() if reg_num.strip() else "Без_номера"
            file_name = f"{safe_reg_num}.docx"
            
            st.download_button(
                label=f"📥 СКАЧАТЬ ОТЧЕТ ({file_name})",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {e}")
