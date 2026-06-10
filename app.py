import time

import pandas as pd
import streamlit as st

from docmodule import extract_data, json_to_excel_bytes, upload_to_google_sheets

st.set_page_config(page_title="Документы в Excel", layout="centered")

st.markdown("""
    <style>
    /* Центрирование заголовка */
    .main-title {
        text-align: center;
        font-weight: bold;
    }

    /* Делаем кнопки действий зелеными */
    div.stButton > button {
        background-color: #28a745;
        color: white;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        width: 100%; /* Кнопка на всю ширину */
    }
    div.stButton > button:hover {
        background-color: #218838;
        color: white;
        border: none;
    }
    div.stButton > button:active {
        background-color: #1e7e34;
        color: white;
    }
    </style>
    """, unsafe_allow_html=True)

st.markdown('<h1 class="main-title">Извлечение данных из Word и PDF в Excel</h1>', unsafe_allow_html=True)
st.write("---")

uploaded_files = st.file_uploader(
    "Загрузите документы Word (.docx) или PDF (.pdf)",
    type=['docx', 'pdf'],
    accept_multiple_files=True
)

if uploaded_files:
    if st.button("Начать обработку", key="process_btn"):
        progress_text = st.empty()
        progress_bar = st.progress(0)

        num_files = len(uploaded_files)
        result_json = []

        for i, file in enumerate(uploaded_files):
            progress_text.text(f"Обработка файла {i + 1} из {num_files}: {file.name}")

            try:
                data = extract_data(file)
                result_json.append({
                    "Имя файла": file.name,
                    "Номер договора": data["number"],
                    "Агент": data["agent"][0] if data["agent"] else "",
                    "Контрагент": data["counterparty"][0] if data["counterparty"] else "",
                    "Предмет договора": "\n".join(data["Предмет_договора"]),
                    "Дата": data["date"],
                    "Сумма": data["summ"],
                })
            except Exception as e:
                st.error(f"Ошибка при обработке {file.name}: {e}")

            time.sleep(0.5)

            percent_complete = (i + 1) / num_files
            progress_bar.progress(percent_complete)

        progress_text.success("✅ Все файлы успешно обработаны!")
        st.session_state['results'] = result_json

    if 'results' in st.session_state and st.session_state['results']:
        df_results = pd.DataFrame(st.session_state['results'])

        st.subheader("Результаты анализа")
        st.dataframe(df_results, use_container_width=True)

        st.download_button(
            label="📥 Скачать результат (Excel)",
            data=json_to_excel_bytes(st.session_state['results']),
            file_name="result.xlsx",
            mime="application/vnd.ms-excel"
        )

        st.write("---")
        st.subheader("Выгрузка в Google Таблицу")

        # Поле для ввода ссылки на ТАБЛИЦУ
        sheet_url = st.text_input(
            "Вставьте ссылку на вашу Google Таблицу (бот должен быть редактором):",
            placeholder="https://docs.google.com/spreadsheets/d/...",
            value="https://docs.google.com/spreadsheets/d/1XW6BPJWZZ82ONThrGaOdEyar5sykzddbo06Ws4lqHUI/edit?gid=0#gid=0",
        )

        if st.button("Добавить данные в таблицу"):
            if not sheet_url:
                st.warning("Пожалуйста, вставьте ссылку на таблицу.")
            else:
                with st.spinner("Добавление данных..."):
                    try:
                        import re

                        match = re.search(r'/d/([a-zA-Z0-9-_]+)', sheet_url)
                        if not match:
                            st.error("Неверный формат ссылки. Нужна ссылка на саму таблицу.")
                        else:
                            sheet_id = match.group(1)

                            upload_to_google_sheets(
                                st.session_state['results'],
                                sheet_id,
                                credits_path="credentials.json"
                            )
                            st.success("Данные успешно добавлены в вашу таблицу!")
                            st.markdown(f"[🔗 Открыть таблицу]({sheet_url})")
                    except Exception as e:
                        st.error(f"Ошибка: {e}")
else:
    st.info("Загрузите файлы для анализа")
