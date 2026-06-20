import streamlit as st
from docxtpl import DocxTemplate, RichText, InlineImage
from docx.shared import Mm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os

# Название файла шаблона, который лежит в папке или на GitHub
TEMPLATE_NAME = "образец отчета1.docx"

# Настройка страницы браузера
st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")

st.title("🚗 Рабочее место оценщика")
st.markdown("Заполните данные ниже для автоматической генерации отчета об оценке ущерба.")

# Автоматическая проверка шаблона
if os.path.exists(TEMPLATE_NAME):
    st.success(f"✅ Базовый шаблон отчета (`{TEMPLATE_NAME}`) успешно подключен автоматически.")
    template_source = TEMPLATE_NAME
else:
    # Если файла нет, сработает запасной вариант с ручной загрузкой
    st.warning(f"⚠️ Файл `{TEMPLATE_NAME}` не найден. Загрузите его вручную ниже:")
    template_source = st.file_uploader("Загрузите шаблон отчета (template.docx)", type="docx")

st.header("1. Ввод данных")

# Создаем две колонки для красоты
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
st.info("💡 **Подсказка:** Пишите каждый пункт просто с новой строки. Программа сама аккуратно добавит маркеры (•) при создании документа!")
damage_desc = st.text_area("Характеристика повреждений (при осмотре установлено):", height=100)
repair_desc = st.text_area("Требуемый ремонт (для восстановления требуется):", height=100)

# Вспомогательная функция для автоматического добавления маркеров (буллитов)
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

# Список ракурсов (можешь добавлять сюда свои варианты)
CAPTION_OPTIONS = [
    "Вид спереди", 
    "Вид сзади", 
    "Вид слева", 
    "Вид справа",
    "Вид спереди, раскол.",
    "Вид слева, раскол, разрушение.",
    "Вид слева, вмятина.",
    "Вид слева, разрыв.",
    "Повреждение крупным планом",
    "VIN номер"
]

# Поле для загрузки нескольких фото
uploaded_photos = st.file_uploader("Загрузите фотографии", type=["jpg", "jpeg", "png"], accept_multiple_files=True)

# Список для хранения данных о фото перед генерацией
photo_data_list = []

if uploaded_photos:
    st.write("Настройте подписи для каждой фотографии:")
    cols = st.columns(2)
    
    for index, photo in enumerate(uploaded_photos):
        col = cols[index % 2]
        with col:
            st.image(photo, use_container_width=True)
            caption = st.selectbox(f"Подпись для фото {index+1}:", CAPTION_OPTIONS, key=f"caption_{index}")
            photo_data_list.append({"file": photo, "caption": caption})

# Логика генерации
if template_source is not None:
    if st.button("СГЕНЕРИРОВАТЬ ОТЧЕТ", type="primary"):
        try:
            doc = DocxTemplate(template_source)
            
            # --- ОБРАБОТКА ФОТОГРАФИЙ ---
            photo_rows = []
            for i in range(0, len(photo_data_list), 2):
                row = {}
                
                # Обработка левой картинки
                img1_file = photo_data_list[i]["file"]
                row["col1"] = {
                    "img": InlineImage(doc, img1_file, width=Mm(80)),
                    "caption": photo_data_list[i]["caption"]
                }
                
                # Обработка правой картинки (если есть)
                if i + 1 < len(photo_data_list):
                    img2_file = photo_data_list[i+1]["file"]
                    row["col2"] = {
                        "img": InlineImage(doc, img2_file, width=Mm(80)),
                        "caption": photo_data_list[i+1]["caption"]
                    }
                else:
                    row["col2"] = {"img": "", "caption": ""}
                    
                photo_rows.append(row)
            
            # --- ФОРМИРОВАНИЕ КОНТЕКСТА ---
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
                "photo_rows": photo_rows  # Передаем массив с фотографиями
            }
            
            doc.render(context)
            
            # --- МАГИЯ ДЛЯ АВТООБНОВЛЕНИЯ ОГЛАВЛЕНИЯ ---
            settings = doc.docx.settings.element
            update_fields = OxmlElement('w:updateFields')
            update_fields.set(qn('w:val'), 'true')
            settings.append(update_fields)
            # -------------------------------------------
            
            # Сохранение результата в память
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✅ Отчет успешно сгенерирован! Скачайте его по кнопке ниже.")
            
            file_name = f"Отчет_{customer if customer else 'Новый_клиент'}.docx"
            
            st.download_button(
                label="📥 СКАЧАТЬ ГОТОВЫЙ ОТЧЕТ",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {e}")
