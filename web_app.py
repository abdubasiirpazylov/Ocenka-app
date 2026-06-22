import streamlit as st
from docxtpl import DocxTemplate, RichText
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io
import os

# Название файла шаблона (главный образец)
TEMPLATE_NAME = "образец отчета.docx"

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

# =========================================================
# БАЗА ШАБЛОНОВ ДЛЯ ОПИСАНИЯ
# =========================================================

# Обязательные приписки (Они будут стоять по умолчанию)
DEFAULT_DAMAGE_SUFFIX = "Дефектный акт на транспортное средство на дату оценки не предоставлялся. Оценка технического состояния произведена без учёта скрытых дефектов."
DEFAULT_REPAIR_SUFFIX = "После завершения ремонтно-восстановительных работ необходим контроль геометрии кузова, зазоров навесных элементов и качества ЛКП. Контроль выполняется организацией, осуществляющей ремонт."

DAMAGE_TEMPLATES = {
    "--- Выберите шаблон ---": "",
    "[Кузов] Передняя часть": "При осмотре установлены повреждения передней части кузова: деформация бампера, повреждение облицовочных элементов, смещение/деформация навесных деталей, нарушение ЛКП.",
    "[Кузов] Задняя часть": "Выявлены повреждения задней части кузова: деформация бампера, повреждение крышки багажника/фонарей, нарушение геометрии сопряжений, повреждение ЛКП.",
    "[Кузов] Боковая часть": "Установлены повреждения боковой части кузова: деформация дверей/крыльев, повреждение навесных элементов, нарушение ЛКП.",
    "[Кузов] Силовые элементы": "Имеются признаки деформации силовых элементов кузова (лонжерон/панель), требующие восстановительных работ с последующим контролем геометрии.",
    "[Оптика] Фара (трещина/разрушение)": "Блок-фара передняя (указать сторону): сквозное разрушение (трещина) рассеивателя.",
    "[Оптика] Фара (царапины)": "Блок-фара передняя (указать сторону): глубокие царапины и потертости рассеивателя.",
    "[Оптика] Фара (крепления)": "Блок-фара передняя (указать сторону): излом элементов крепления корпуса.",
    "[Стекла] Лобовое (трещина)": "Стекло ветровое: линейная трещина в зоне видимости водителя (или: в зоне работы стеклоочистителей).",
    "[Стекла] Лобовое (скол)": "Стекло ветровое: скол типа «звезда» (или «бычий глаз») с развивающимися трещинами.",
    "[Стекла] Боковое (царапины)": "Стекло передней/задней двери (указать сторону): царапины (задиры) на внешней поверхности.",
    "[Стекла] Боковое (разрушение)": "Стекло передней/задней двери (указать сторону): разрушение элемента (отсутствует).",
    "[Стекла] Заднее (седан)": "Стекло задка: разрушение элемента / глубокие царапины.",
    "[Стекла] Заднее (хэтчбек/внедорожник)": "Стекло двери задка (крышки багажника): повреждение нитей обогрева / разрушение.",
}

REPAIR_TEMPLATES = {
    "--- Выберите шаблон ---": "",
    "[Кузов] Стандартные работы": "Для восстановления требуется выполнить комплекс слесарно-кузовных, рихтовочных и малярно-окрасочных работ с применением расходных материалов, с последующей сборкой и регулировкой навесных элементов.",
    "[Оптика] Замена фары": "Демонтаж, монтаж (замена) блок-фары передней (указать сторону) в сборе.",
    "[Стекла] Лобовое стекло (база)": "Замена стекла ветрового (вклейка) с использованием комплекта однокомпонентного полиуретанового клея.",
    "[Стекла] Лобовое стекло (+датчики)": "Замена стекла ветрового (вклейка) с использованием комплекта однокомпонентного полиуретанового клея и переустановкой датчика дождя/камеры слежения.",
    "[Стекла] Боковое стекло": "Снятие обивки двери, очистка внутренней полости от осколков, замена стекла двери.",
    "[Стекла] Заднее стекло": "Замена стекла задка (вклейка) с подключением элементов обогрева."
}

# Инициализация текста по умолчанию в сессии
if "damage_text" not in st.session_state:
    st.session_state.damage_text = DEFAULT_DAMAGE_SUFFIX

if "repair_text" not in st.session_state:
    st.session_state.repair_text = f"Для восстановления требуется выполнить комплекс слесарно-кузовных, рихтовочных и малярно-окрасочных работ с применением расходных материалов, с последующей сборкой и регулировкой навесных элементов.\n{DEFAULT_REPAIR_SUFFIX}"

# Функции для умного добавления фраз (чтобы они вставлялись перед припиской)
def add_to_damage():
    selected = st.session_state.dmg_selector
    if selected and DAMAGE_TEMPLATES[selected]:
        current = st.session_state.damage_text
        new_phrase = DAMAGE_TEMPLATES[selected]
        # Вставляем фразу ПЕРЕД дефолтной припиской
        if DEFAULT_DAMAGE_SUFFIX in current:
            st.session_state.damage_text = current.replace(DEFAULT_DAMAGE_SUFFIX, new_phrase + "\n" + DEFAULT_DAMAGE_SUFFIX)
        else:
            st.session_state.damage_text = current + "\n" + new_phrase if current else new_phrase

def add_to_repair():
    selected = st.session_state.rep_selector
    if selected and REPAIR_TEMPLATES[selected]:
        current = st.session_state.repair_text
        new_phrase = REPAIR_TEMPLATES[selected]
        # Вставляем фразу ПЕРЕД дефолтной припиской
        if DEFAULT_REPAIR_SUFFIX in current:
            st.session_state.repair_text = current.replace(DEFAULT_REPAIR_SUFFIX, new_phrase + "\n" + DEFAULT_REPAIR_SUFFIX)
        else:
            st.session_state.repair_text = current + "\n" + new_phrase if current else new_phrase

# =========================================================

st.header("2. Описание повреждений и ремонта")

# Панель конструктора
col_dmg, col_rep = st.columns(2)

with col_dmg:
    st.selectbox("Конструктор осмотра:", list(DAMAGE_TEMPLATES.keys()), key="dmg_selector")
    st.button("➕ Добавить в осмотр", on_click=add_to_damage, use_container_width=True)

with col_rep:
    st.selectbox("Конструктор ремонта:", list(REPAIR_TEMPLATES.keys()), key="rep_selector")
    st.button("➕ Добавить в ремонт", on_click=add_to_repair, use_container_width=True)

# Текстовые поля, привязанные к session_state (можно редактировать руками)
damage_desc = st.text_area("Характеристика повреждений (при осмотре установлено):", key="damage_text", height=200)
repair_desc = st.text_area("Требуемый ремонт (для восстановления требуется):", key="repair_text", height=200)

# Функция для сохранения переносов строк в Word
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
photo_report_doc = st.file_uploader("Загрузите готовый Фотоотчет (.docx)", type="docx")

if template_source is not None:
    if st.button("СГЕНЕРИРОВАТЬ ИТОГОВЫЙ ОТЧЕТ", type="primary", use_container_width=True):
        try:
            doc = DocxTemplate(template_source)
            
            if photo_report_doc is not None:
                subdoc_photo = doc.new_subdoc(photo_report_doc)
            else:
                subdoc_photo = "Таблица с фотографиями не была приложена к отчету."

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
                "PHOTO_TABLE": subdoc_photo 
            }
            
            doc.render(context)
            
            settings = doc.docx.settings.element
            update_fields = OxmlElement('w:updateFields')
            update_fields.set(qn('w:val'), 'true')
            settings.append(update_fields)
            
            buffer = io.BytesIO()
            doc.save(buffer)
            buffer.seek(0)
            
            st.success("✅ Итоговый отчет успешно сгенерирован!")
            
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
