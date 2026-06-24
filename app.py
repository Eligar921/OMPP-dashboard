import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ОМПП Дашборд", layout="wide")
st.title("📊 Дашборд ОМПП")

uploaded_file = st.file_uploader("Загрузите Excel файл 'отчет по дате направления'", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df.columns = df.columns.str.strip()

    # ---- Поиск столбцов с приоритетом точного совпадения ----
    def find_column(keywords, exact_match=None):
        if exact_match is not None:
            for col in df.columns:
                if col.strip().lower() == exact_match.lower():
                    return col
        for col in df.columns:
            col_lower = col.lower()
            for kw in keywords:
                if kw.lower() in col_lower:
                    return col
        return None

    col_date_direction = find_column(['дата направления', 'направления на координатора'])
    col_phone = find_column(['телефон'])
    col_recruiter = find_column(['рекрутер'])
    col_source = find_column(['источник омпп', 'источник'])
    col_last_call = find_column(
        ['последнего звонка до первого статуса', 'последнего звонка', 'последний звонок'],
        exact_match='Дата последнего звонка до первого статуса первой смены'
    )
    col_coord_status = find_column(['статус координатора', 'статус координатор'])
    col_lead_status = find_column(['статус лида'])
    # Город ищем с приоритетом точного совпадения "Город"
    col_city = find_column(['город'], exact_match='Город')
    if col_city is None:
        col_city = find_column(['город'])
    col_project_group = find_column(['желаемые проекты (группа)', 'группа'])
    col_project_client = find_column(['желаемые проекты (клиент)', 'клиент'])

    # ---- Проверка ----
    if col_date_direction is None:
        st.error("❌ Не найден столбец с датой направления. Доступные столбцы: " + ", ".join(df.columns))
        st.stop()
    if col_phone is None:
        st.error("❌ Не найден столбец 'Телефон'")
        st.stop()
    if col_recruiter is None:
        st.error("❌ Не найден столбец 'Рекрутер'")
        st.stop()
    if col_source is None:
        st.error("❌ Не найден столбец 'Источник ОМПП'")
        st.stop()
    if col_last_call is None:
        st.error("❌ Не найден столбец 'Дата последнего звонка до первого статуса первой смены'")
        st.stop()
    if col_coord_status is None and col_lead_status is None:
        st.error("❌ Не найден ни столбец 'Статус координатора', ни 'Статус лида'")
        st.stop()

    # ---- Переименование ----
    rename_map = {
        col_date_direction: 'Дата направления',
        col_phone: 'Телефон',
        col_recruiter: 'Рекрутер',
        col_source: 'Источник ОМПП',
        col_last_call: 'Дата последнего звонка',
    }
    if col_coord_status is not None:
        rename_map[col_coord_status] = 'Статус координатора'
    if col_lead_status is not None:
        rename_map[col_lead_status] = 'Статус лида'
    if col_city is not None:
        rename_map[col_city] = 'Город'
    if col_project_group is not None:
        rename_map[col_project_group] = 'Желаемые проекты (Группа)'
    if col_project_client is not None:
        rename_map[col_project_client] = 'Желаемые проекты (Клиент)'

    df = df.rename(columns=rename_map)
    df = df.loc[:, ~df.columns.duplicated()]

    # ---- Преобразование дат ----
    df['Дата направления'] = pd.to_datetime(df['Дата направления'], errors='coerce')
    df['Дата последнего звонка'] = pd.to_datetime(df['Дата последнего звонка'], errors='coerce')

    # ---- Сохраняем копию ДО фильтра по звонку (для диагностики) ----
    df_raw = df.copy()

    # ---- Исключаем пустые источники ----
    df = df[df['Источник ОМПП'].notna() & (df['Источник ОМПП'] != '')]

    # ---- Применяем фильтр по дате звонка ----
    df['год_напр'] = df['Дата направления'].dt.year
    df['мес_напр'] = df['Дата направления'].dt.month
    df['год_зв'] = df['Дата последнего звонка'].dt.year
    df['мес_зв'] = df['Дата последнего звонка'].dt.month

    same_month = (df['год_зв'] == df['год_напр']) & (df['мес_зв'] == df['мес_напр'])
    prev_month = (df['год_зв'] == df['год_напр']) & (df['мес_зв'] == df['мес_напр'] - 1)
    prev_month_jan = (df['год_зв'] == df['год_напр'] - 1) & (df['мес_напр'] == 1) & (df['мес_зв'] == 12)

    df['filter_last_call'] = same_month | prev_month | prev_month_jan
    df = df[df['filter_last_call'] & df['Дата последнего звонка'].notna()]

    # ---- ДИАГНОСТИКА (в сайдбаре) ----
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 Диагностика")
    st.sidebar.write(f"**Найден столбец 'Город':** {col_city if col_city else 'не найден'}")
    st.sidebar.write("**Все столбцы после переименования:**", list(df.columns))

    # Данные до фильтра (после исключения пустых источников)
    df_before = df_raw[df_raw['Источник ОМПП'].notna() & (df_raw['Источник ОМПП'] != '')]
    total_before = len(df_before)

    if 'Город' in df_before.columns:
        moscow_before = len(df_before[df_before['Город'].astype(str).str.contains('москва', case=False, na=False)])
        # Покажем уникальные значения города (первые 20)
        unique_cities = df_before['Город'].astype(str).str.strip().value_counts().head(20)
        st.sidebar.write("**Уникальные города (первые 20):**")
        st.sidebar.dataframe(unique_cities.reset_index())
    else:
        moscow_before = 0

    total_after = len(df)
    if 'Город' in df.columns:
        moscow_after = len(df[df['Город'].astype(str).str.contains('москва', case=False, na=False)])
    else:
        moscow_after = 0

    # Всего Москва в исходном файле (без фильтра по источнику)
    total_moscow_raw = len(df_raw[df_raw['Город'].astype(str).str.contains('москва', case=False, na=False)]) if 'Город' in df_raw.columns else 0

    st.sidebar.write(f"**Всего Москва в исходном файле (без фильтра по источнику):** {total_moscow_raw}")
    st.sidebar.write(f"**До фильтра по звонку (после исключения пустых источников):**")
    st.sidebar.write(f"Всего строк: {total_before}")
    st.sidebar.write(f"Из них Москва: {moscow_before}")

    st.sidebar.write(f"**После фильтра по звонку:**")
    st.sidebar.write(f"Всего строк: {total_after}")
    st.sidebar.write(f"Из них Москва: {moscow_after}")

    # Покажем пример отброшенных записей с Москвой
    if 'Город' in df_before.columns:
        kept_indices = set(df.index)
        dropped = df_before[
            df_before['Город'].astype(str).str.contains('москва', case=False, na=False) &
            (~df_before.index.isin(kept_indices))
        ]
        if not dropped.empty:
            st.sidebar.write(f"**Отброшено записей с Москвой: {len(dropped)}**")
            st.sidebar.write("Примеры дат звонка у отброшенных:")
            st.sidebar.dataframe(dropped[['Телефон', 'Дата направления', 'Дата последнего звонка']].head(5))
        else:
            st.sidebar.write("Нет отброшенных записей с Москвой (все сохранены).")

    # ---- ОСТАЛЬНАЯ ЧАСТЬ КОДА (БЕЗ ИЗМЕНЕНИЙ) ----
    # ... (весь остальной код, который вы уже использовали, начиная с боковой панели)

else:
    st.info("👈 Загрузите файл Excel для начала работы.")
