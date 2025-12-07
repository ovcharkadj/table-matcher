import streamlit as st
import pandas as pd
import docx
import re
import io

# --- Настройки страницы ---
st.set_page_config(page_title="Data Matcher", layout="wide")

# --- Функции чтения файлов ---

def read_excel(file):
    try:
        # Читаем все листы, объединяем в один датафрейм
        xls = pd.ExcelFile(file)
        all_data = []
        for sheet_name in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet_name, dtype=str)
            df['Источник'] = f"{file.name} (Лист: {sheet_name})"
            all_data.append(df)
        return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    except Exception as e:
        st.error(f"Ошибка при чтении Excel {file.name}: {e}")
        return pd.DataFrame()

def read_docx(file):
    try:
        doc = docx.Document(file)
        all_data = []
        
        for i, table in enumerate(doc.tables):
            # Предполагаем, что первая строка - это заголовки
            if len(table.rows) < 2:
                continue
            
            headers = [cell.text.strip() for cell in table.rows[0].cells]
            data = []
            for row in table.rows[1:]:
                row_data = [cell.text.strip() for cell in row.cells]
                # Выравниваем длину строки под заголовки
                if len(row_data) < len(headers):
                    row_data += [''] * (len(headers) - len(row_data))
                data.append(row_data[:len(headers)])
            
            df = pd.DataFrame(data, columns=headers)
            df['Источник'] = f"{file.name} (Таблица {i+1})"
            all_data.append(df)
            
        return pd.concat(all_data, ignore_index=True) if all_data else pd.DataFrame()
    except Exception as e:
        st.error(f"Ошибка при чтении Word {file.name}: {e}")
        return pd.DataFrame()

# --- Функции нормализации ---

def normalize_text(text, ignore_case, ignore_symbols):
    if pd.isna(text):
        return ""
    text = str(text)
    if ignore_case:
        text = text.lower()
    if ignore_symbols:
        # Оставляем только буквы и цифры
        text = re.sub(r'[^a-zA-Zа-яА-Я0-9]', '', text)
    else:
        # Всегда убираем лишние пробелы по краям
        text = text.strip()
    return text

# --- Интерфейс ---

st.title("🔍 Поиск совпадений в Excel и Word")
st.markdown("Загрузите файлы, выберите колонки и найдите дубликаты или пересечения.")

# 1. Загрузка
uploaded_files = st.file_uploader("Перетащите файлы (.xlsx, .docx)", type=['xlsx', 'docx'], accept_multiple_files=True)

if uploaded_files:
    all_dfs = []
    for file in uploaded_files:
        if file.name.endswith('.xlsx'):
            all_dfs.append(read_excel(file))
        elif file.name.endswith('.docx'):
            all_dfs.append(read_docx(file))
    
    if all_dfs:
        # Объединение всех данных
        # Используем outer join, чтобы сохранить все уникальные колонки из всех файлов
        main_df = pd.concat(all_dfs, ignore_index=True)
        
        # Добавляем индекс оригинальной строки для наглядности
        main_df.reset_index(inplace=True, names=['ID_строки'])

        st.write("### 1. Предпросмотр всех загруженных данных")
        st.dataframe(main_df.head())
        st.info(f"Всего загружено строк: {len(main_df)}. Колонок: {list(main_df.columns)}")

        # 2. Настройка поиска
        st.write("---")
        st.write("### 2. Настройка фильтров поиска")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("По каким полям искать совпадения?")
            # Исключаем служебные колонки
            cols_available = [c for c in main_df.columns if c not in ['Источник', 'ID_строки']]
            selected_cols = st.multiselect("Выберите колонки (ФИО, Телефон и т.д.)", cols_available)

        with col2:
            st.subheader("Параметры строгости")
            ignore_case = st.checkbox("Игнорировать регистр (А = а)", value=True)
            ignore_symbols = st.checkbox("Игнорировать спецсимволы и пробелы (тел: +7-999 -> 7999)", value=True)

        # 3. Логика поиска
        if selected_cols:
            st.write("---")
            st.write("### 3. Результаты")
            
            # Создаем временные колонки для сравнения ("хеши")
            search_df = main_df.copy()
            
            # Формируем единый ключ поиска
            search_df['match_key'] = ""
            for col in selected_cols:
                # Заполняем NaN пустыми строками, чтобы не ломать логику
                search_df[col] = search_df[col].fillna("")
                # Применяем нормализацию
                search_df['match_key'] += search_df[col].apply(lambda x: normalize_text(x, ignore_case, ignore_symbols))

            # Ищем дубликаты по match_key
            # keep=False означает "пометить ВСЕ дубликаты", а не только второй и последующие
            duplicates_mask = search_df.duplicated(subset=['match_key'], keep=False)
            
            # Фильтруем пустые ключи (если строки были пустыми)
            duplicates_mask = duplicates_mask & (search_df['match_key'] != "")
            
            results = main_df[duplicates_mask].copy()
            # Добавляем ключ группировки для сортировки
            results['Группа_совпадения'] = search_df.loc[duplicates_mask, 'match_key']
            
            # Сортируем, чтобы одинаковые записи шли подряд
            results = results.sort_values(by=['Группа_совпадения', 'Источник'])

            if not results.empty:
                st.success(f"Найдено {len(results)} записей, имеющих совпадения!")
                
                # Отображение
                st.dataframe(
                    results[['Группа_совпадения'] + selected_cols + ['Источник']],
                    use_container_width=True,
                    hide_index=True
                )
                
                # Экспорт
                st.download_button(
                    label="Скачать результаты в Excel",
                    data=io.BytesIO(), # Здесь нужна доп. логика для записи в буфер, ниже упрощенно
                    file_name="matches.csv",
                    mime="text/csv"
                )
                
                # Для корректного скачивания Excel (Streamlit требует спец. обработки)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    results.to_excel(writer, index=False)
                
                st.download_button(
                    label="📥 Скачать отчет (.xlsx)",
                    data=buffer.getvalue(),
                    file_name="report_matches.xlsx",
                    mime="application/vnd.ms-excel"
                )
                
            else:
                st.warning("Совпадений по выбранным критериям не найдено.")
        else:
            st.info("Выберите хотя бы одну колонку для начала поиска.")
            
    else:
        st.warning("Не удалось прочитать данные из файлов.")