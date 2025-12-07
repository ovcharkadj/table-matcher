import streamlit as st
import pandas as pd
import docx
import re
import io

# --- Настройки страницы ---
st.set_page_config(page_title="Data Matcher Pro", layout="wide")

# --- Функции чтения файлов ---
def read_excel(file):
    try:
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
            if len(table.rows) < 1: continue
            
            # Попытка определить, где заголовки
            headers = [cell.text.strip() for cell in table.rows[0].cells]
            # Если заголовки пустые, генерируем имена Col1, Col2...
            if all(h == '' for h in headers):
                headers = [f"Col_{j}" for j in range(len(headers))]
                
            data = []
            start_row = 1 if len(table.rows) > 1 else 0 # Если только одна строка, считаем её данными
            
            for row in table.rows[start_row:]:
                row_data = [cell.text.strip() for cell in row.cells]
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

# --- Функция нормализации ---
def normalize_text(text, ignore_case, ignore_symbols):
    if pd.isna(text): return ""
    text = str(text)
    if ignore_case: text = text.lower()
    if ignore_symbols: text = re.sub(r'[^a-zA-Zа-яА-Я0-9]', '', text)
    else: text = text.strip()
    return text

# --- ИНТЕРФЕЙС ---

# 1. Боковая панель для загрузки
with st.sidebar:
    st.header("📂 Загрузка файлов")
    uploaded_files = st.file_uploader("Перетащите файлы сюда", type=['xlsx', 'docx'], accept_multiple_files=True)
    st.info("Поддерживаются Excel (.xlsx) и Word (.docx)")

st.title("🔎 Data Matcher: Поиск и Анализ")

if uploaded_files:
    all_dfs = []
    for file in uploaded_files:
        if file.name.endswith('.xlsx'): all_dfs.append(read_excel(file))
        elif file.name.endswith('.docx'): all_dfs.append(read_docx(file))
    
    if all_dfs:
        main_df = pd.concat(all_dfs, ignore_index=True)
        main_df.reset_index(inplace=True, names=['ID'])
        
        # --- БЛОК 1: БЫСТРЫЙ ПОИСК (CTRL+F) ---
        st.markdown("### 🚀 Быстрый поиск по тексту")
        search_query = st.text_input("Введите любой текст, номер или имя (фильтрует таблицу на лету):", placeholder="Например: Иванов или 999")
        
        filtered_df = main_df.copy()
        if search_query:
            # Магия поиска по всем колонкам сразу
            mask = filtered_df.astype(str).apply(
                lambda x: x.str.contains(search_query, case=False, na=False)
            ).any(axis=1)
            filtered_df = filtered_df[mask]
            st.success(f"Найдено совпадений: {len(filtered_df)}")
        
        st.dataframe(filtered_df, use_container_width=True, hide_index=True)

        st.markdown("---")

        # --- БЛОК 2: ПОИСК ДУБЛИКАТОВ (МАТЧЕР) ---
        with st.expander("🛠️ Инструмент поиска дубликатов (Сравнение колонок)", expanded=False):
            st.write("Настройте точный поиск совпадений между файлами.")
            
            c1, c2 = st.columns(2)
            cols_available = [c for c in main_df.columns if c not in ['Источник', 'ID']]
            
            with c1:
                selected_cols = st.multiselect("Выберите колонки для сравнения", cols_available)
            with c2:
                ignore_case = st.checkbox("Игнорировать регистр", value=True)
                ignore_symbols = st.checkbox("Игнорировать символы (для телефонов)", value=True)

            if selected_cols:
                # Логика поиска дублей
                search_df = main_df.copy()
                search_df['match_key'] = ""
                for col in selected_cols:
                    search_df[col] = search_df[col].fillna("")
                    search_df['match_key'] += search_df[col].apply(lambda x: normalize_text(x, ignore_case, ignore_symbols))
                
                # Ищем где ключ повторяется
                dupes = search_df[search_df.duplicated(subset=['match_key'], keep=False)]
                dupes = dupes[dupes['match_key'] != ""]
                
                if not dupes.empty:
                    dupes = dupes.sort_values(by=['match_key', 'Источник'])
                    st.success(f"Найдено групп совпадений: {len(dupes)}")
                    st.dataframe(dupes[['match_key'] + selected_cols + ['Источник']], use_container_width=True)
                    
                    # Скачивание
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        dupes.to_excel(writer, index=False)
                    st.download_button("📥 Скачать отчет (.xlsx)", buffer.getvalue(), "report.xlsx")
                else:
                    st.warning("Дубликатов по выбранным колонкам не найдено.")

    else:
        st.error("Не удалось прочитать данные. Проверьте формат файлов.")
else:
    st.info("⬅️ Загрузите файлы в меню слева, чтобы начать.")
