import streamlit as st
from docxtpl import DocxTemplate, RichText
from docx.shared import Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os

# Название файла шаблона
TEMPLATE_NAME = "образец отчета.docx"

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

# Инициализация состояния
if "photo_order" not in st.session_state:
    st.session_state.photo_order = []
if "last_files" not in st.session_state:
    st.session_state.last_files = []

# --- ИСПРАВЛЕННЫЙ И БЕЗОПАСНЫЙ КОЛБЭК ---
def move_photo(file_name, current_index):
    # Безопасно получаем новое значение, выбранное пользователем
    new_val = st.session_state.get(f"pos_{file_name}")
    if new_val is None:
        return
        
    new_index = new_val - 1
    
    if new_index != current_index:
        # 1. Перемещаем сам файл в логическом списке
        item = st.session_state.photo_order.pop(current_index)
        st.session_state.photo_order.insert(new_index, item)
        
        # 2. Вместо удаления ключей, принудительно переписываем ВСЕМ фотографиям
        # их новые правильные позиции в памяти браузера. Это полностью убирает баг "каши".
        for idx, fname in enumerate(st.session_state.photo_order):
            st.session_state[f"pos_{fname}"] = idx + 1
# ----------------------------------------

uploaded_photos = st.file_uploader("Загрузите фотографии", type=["jpg", "jpeg", "png"], accept_multiple_files=True)
photo_data_list = []

if uploaded_photos:
    current_filenames = [f.name for f in uploaded_photos]
    
    # Если загрузили новые фото или удалили старые
    if set(current_filenames) != set(st.session_state.last_files):
        new_order = [f for f in st.session_state.photo_order if f in current_filenames]
        new_files = [f for f in current_filenames if f not in new_order]
        st.session_state.photo_order = new_order + new_files
        st.session_state.last_files = current_filenames
        
        # Безопасная инициализация позиций для новых фото
        for idx, fname in enumerate(st.session_state.photo_order):
            key = f"pos_{fname}"
            if key not in st.session_state:
                st.session_state[key] = idx + 1

    file_dict = {f
