import streamlit as st
from docxtpl import DocxTemplate
import io
import os

# Название файла шаблона, который лежит на GitHub
TEMPLATE_NAME = "образец отчета.docx"

# Настройка страницы браузера
st.set_page_config(page_title="Генератор Отчетов - Гарант Оценка", layout="wide")

st.title("🚗 Рабочее место оценщика")
st.markdown("Заполните данные ниже для автоматической генерации отчета об оценке ущерба.")

# Автоматическая проверка шаблона
if os.path.exists(TEMPLATE_NAME):
    st.success(f"✅ Базовый шаблон отчета (`{TEMPLATE_NAME}`) успешно подключен автоматически.")
    template_source = TEMPLATE_NAME
else:
    # Если файла нет на GitHub, сработает запасной вариант с ручной загрузкой
    st.warning(f"⚠️ Файл `{TEMPLATE_NAME}` не найден в репозитории. Загрузите его вручную ниже:")
    template_source = st.file_uploader("Загрузите шаблон отчета (template.docx)", type="docx")

st.header("2. Ввод данных")

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

st.subheader("Описание")
damage_desc = st.text_area("Характеристика повреждений (при осмотре установлено):", height=100)
repair_desc = st.text_area("Требуемый ремонт (для восстановления требуется):", height=100)

# Логика генерации
if template_source is not None:
    if st.button("СГЕНЕРИРОВАТЬ ОТЧЕТ", type="primary"):
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
            "DAMAGE_DESC": damage_desc,
            "REPAIR_DESC": repair_desc
        }
        
        try:
            # Загружаем шаблон (автоматический из папки или ручной из uploader)
            doc = DocxTemplate(template_source)
            
            # Заполняем данными
            doc.render(context)
            
            # Сохраняем результат в память
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("Отчет успешно сгенерирован! Скачайте его по кнопке ниже.")
            
            file_name = f"Отчет_{customer if customer else 'Новый_клиент'}.docx"
            
            st.download_button(
                label="📥 СКАЧАТЬ ГОТОВЫЙ ОТЧЕТ",
                data=buffer,
                file_name=file_name,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
            
        except Exception as e:
            st.error(f"Произошла ошибка при обработке файла: {e}")
