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

def format_as_bullets(text):
    rt = RichText()
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    for i, line in enumerate(lines):
        if not line.startswith(('-', '•', '*', '1.', '2.')):
            line = f"• {line}"
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

# Инициализация состояния для хранения порядка фотографий
if "photo_order" not in st.session_state:
    st.session_state.photo_order = []
if "last_files" not in st.session_state:
    st.session_state.last_files = []

uploaded_photos = st.file_uploader("Загрузите фотографии", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
photo_data_list = []

if uploaded_photos:
    current_filenames = [f.name for f in uploaded_photos]
    if set(current_filenames) != set(st.session_state.last_files):
        new_order = [f for f in st.session_state.photo_order if f in current_filenames]
        new_files = [f for f in current_filenames if f not in new_order]
        st.session_state.photo_order = new_order + new_files
        st.session_state.last_files = current_filenames

    file_dict = {f.name: f for f in uploaded_photos}

    st.write("Настройте подписи и порядок фотографий:")
    
    for i, filename in enumerate(st.session_state.photo_order):
        photo = file_dict[filename]
        
        with st.container():
            c_img, c_controls = st.columns([1, 2])
            
            with c_img:
                st.image(photo, use_container_width=True)
                
                # --- НОВАЯ УМНАЯ СОРТИРОВКА ---
                # Выпадающий список с номерами позиций (от 1 до количества фото)
                new_pos = st.selectbox(
                    "📍 Позиция в отчете:",
                    options=list(range(1, len(st.session_state.photo_order) + 1)),
                    index=i,
                    key=f"pos_{filename}"
                )
                
                # Если пользователь выбрал новую цифру, меняем порядок и обновляем страницу
                if new_pos - 1 != i:
                    moved_file = st.session_state.photo_order.pop(i)
                    st.session_state.photo_order.insert(new_pos - 1, moved_file)
                    st.rerun()
                # ------------------------------

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
                "DAMAGE_DESC": format_as_bullets(damage_desc),
                "REPAIR_DESC": format_as_bullets(repair_desc),
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
            
            file_name = f"Отчет_{customer if customer else 'Новый_клиент'}.docx"
            st.download_button(
                label="📥 СКАЧАТЬ ГОТОВЫЙ ОТЧЕТ",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {e}")
