import streamlit as st
from docxtpl import DocxTemplate, RichText
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os

# Название файла шаблона (главный образец)
TEMPLATE_NAME = "образец отчета1.docx"

st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")

st.title("🚗 Главное рабочее место оценщика")
st.markdown("Заполните данные и прикрепите готовый фотоотчет от эксперта.")

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

# Функция для сохранения переносов строк (без маркеров-точек, как ты и просил ранее)
def format_text_with_newlines(text):
    rt = RichText()
    lines = [line.strip() for line in text.split('\n') if line.strip()]
    for i, line in enumerate(lines):
        if i < len(lines) - 1:
            rt.add(line + '\n')
        else:
            rt.add(line)
    return rt

st.header("3. Приложение: Фотоотчет")
st.info("💡 Загрузите сюда файл .docx, который был сгенерирован экспертом в мобильном приложении.")
# Поле для загрузки готового Word-файла с таблицей
photo_report_doc = st.file_uploader("Загрузите готовый Фотоотчет (.docx)", type="docx")

if template_source is not None:
    if st.button("СГЕНЕРИРОВАТЬ ИТОГОВЫЙ ОТЧЕТ", type="primary", use_container_width=True):
        try:
            # Загружаем главный шаблон
            doc = DocxTemplate(template_source)
            
            # --- МАГИЯ СКЛЕИВАНИЯ ДОКУМЕНТОВ ---
            if photo_report_doc is not None:
                # Превращаем загруженный файл фотоотчета во "вложенный документ" (subdoc)
                subdoc_photo = doc.new_subdoc(photo_report_doc)
            else:
                # Если забыли загрузить
                subdoc_photo = "Таблица с фотографиями не была приложена к отчету."
            # ------------------------------------

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
                "DAMAGE_DESC": format_text_with_newlines(damage_desc),
                "REPAIR_DESC": format_text_with_newlines(repair_desc),
                "PHOTO_TABLE": subdoc_photo  # <--- Вшиваем готовый файл прямо на место тега!
            }
            
            # Заполняем шаблон
            doc.render(context)
            
            # Автообновление оглавления
            settings = doc.docx.settings.element
            update_fields = OxmlElement('w:updateFields')
            update_fields.set(qn('w:val'), 'true')
            settings.append(update_fields)
            
            # Сохранение в память
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✅ Итоговый отчет успешно сгенерирован и склеен с фотоотчетом!")
            
            # Сохраняем под госномером (как ты просил ранее)
            safe_reg_num = reg_num.strip() if reg_num.strip() else "Без_номера"
            file_name = f"{safe_reg_num}.docx"
            
            st.download_button(
                label=f"📥 СКАЧАТЬ ИТОГОВЫЙ ОТЧЕТ ({file_name})",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
            
        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {e}")
