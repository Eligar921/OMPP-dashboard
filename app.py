import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ОМПП Дашборд", layout="wide")
st.title("📊 Дашборд ОМПП")

uploaded_file = st.file_uploader("Загрузите Excel файл 'отчет по дате направления'", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    
    # Приведение названий столбцов (удаляем пробелы по краям)
    df.columns = df.columns.str.strip()
    
    # Функция поиска столбца по частичному совпадению (регистронезависимо)
    def find_column(keywords):
        """Возвращает имя столбца, содержащее хотя бы одно из ключевых слов (без учёта регистра)."""
        for col in df.columns:
            col_lower = col.lower()
            for kw in keywords:
                if kw.lower() in col_lower:
                    return col
        return None

    # Ищем нужные столбцы
    col_date_direction = find_column(['дата направления', 'направления на координатора'])
    col_phone = find_column(['телефон'])
    col_recruiter = find_column(['рекрутер'])
    col_source = find_column(['источник омпп', 'источник'])
    col_last_call = find_column(['последнего звонка', 'последний звонок'])
    col_coord_status = find_column(['статус координатора', 'статус координатор'])
    col_lead_status = find_column(['статус лида'])

    # Проверка наличия обязательных столбцов
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
        st.error("❌ Не найден столбец с датой последнего звонка")
        st.stop()
    if col_coord_status is None and col_lead_status is None:
        st.error("❌ Не найден ни столбец 'Статус координатора', ни 'Статус лида'")
        st.stop()

    # Переименовываем для удобства (внутри кода будем использовать стандартные имена)
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

    df = df.rename(columns=rename_map)

    # Преобразование дат
    df['Дата направления'] = pd.to_datetime(df['Дата направления'], errors='coerce')
    df['Дата последнего звонка'] = pd.to_datetime(df['Дата последнего звонка'], errors='coerce')

    # Исключаем строки с пустым Источник ОМПП
    df = df[df['Источник ОМПП'].notna() & (df['Источник ОМПП'] != '')]

    # ---- Автоматический фильтр: дата звонка в том же или предыдущем месяце ----
    df['год_напр'] = df['Дата направления'].dt.year
    df['мес_напр'] = df['Дата направления'].dt.month
    df['год_зв'] = df['Дата последнего звонка'].dt.year
    df['мес_зв'] = df['Дата последнего звонка'].dt.month

    cond_same = (df['год_зв'] == df['год_напр']) & (df['мес_зв'] == df['мес_напр'])
    cond_prev = (df['год_зв'] == df['год_напр']) & (df['мес_зв'] == df['мес_напр'] - 1)
    cond_prev_year = (df['год_зв'] == df['год_напр'] - 1) & (df['мес_напр'] == 1) & (df['мес_зв'] == 12)

    df['filter_last_call'] = cond_same | cond_prev | cond_prev_year
    df = df[df['filter_last_call'] & df['Дата последнего звонка'].notna()]

    # ---- Боковая панель ----
    st.sidebar.header("Фильтры")

    sources = sorted(df['Источник ОМПП'].unique())
    selected_sources = st.sidebar.multiselect("Источник ОМПП", options=sources, default=sources)
    df_filtered = df[df['Источник ОМПП'].isin(selected_sources)]

    # Диапазон дат направления
    min_date = df_filtered['Дата направления'].min().date()
    max_date = df_filtered['Дата направления'].max().date()
    date_range = st.sidebar.date_input(
        "Диапазон дат направления",
        value=(min_date, max_date),
        min_value=min_date,
        max_value=max_date
    )
    if len(date_range) == 2:
        start_date, end_date = date_range
        df_filtered = df_filtered[
            (df_filtered['Дата направления'].dt.date >= start_date) &
            (df_filtered['Дата направления'].dt.date <= end_date)
        ]

    # ---- Основная таблица ----
    recruiter_counts = df_filtered.groupby('Рекрутер')['Телефон'].nunique().reset_index()
    recruiter_counts.columns = ['Рекрутер', 'Кол-во кандидатов']
    recruiter_counts = recruiter_counts.sort_values('Кол-во кандидатов', ascending=False)

    st.subheader("📋 Количество направленных кандидатов по рекрутерам")
    st.dataframe(recruiter_counts, use_container_width=True)

    # ---- Линейный график: вышедшие по источникам ----
    # Определяем "вышедших" как went_work (статус координатора)
    if 'Статус координатора' in df_filtered.columns:
        df_worked = df_filtered[df_filtered['Статус координатора'] == 'went_work']
    else:
        # если столбец статуса координатора отсутствует, используем статус лида
        df_worked = df_filtered[df_filtered['Статус лида'] == 'worked']

    if not df_worked.empty:
        worked_by_recruiter_source = df_worked.groupby(['Рекрутер', 'Источник ОМПП']).size().reset_index(name='Кол-во вышедших')
        fig = px.line(
            worked_by_recruiter_source,
            x='Источник ОМПП',
            y='Кол-во вышедших',
            color='Рекрутер',
            markers=True,
            title="Количество вышедших кандидатов по источникам (линии — рекрутеры)"
        )
        fig.update_layout(xaxis_title="Источник", yaxis_title="Кол-во вышедших")
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Нет данных о вышедших кандидатах для выбранных фильтров.")

    # ---- Статистика в сайдбаре ----
    st.sidebar.markdown("---")
    st.sidebar.write(f"🧾 Всего строк после фильтров: **{len(df_filtered)}**")
    st.sidebar.write(f"👥 Уникальных рекрутеров: **{df_filtered['Рекрутер'].nunique()}**")
    st.sidebar.write(f"📞 Уникальных телефонов: **{df_filtered['Телефон'].nunique()}**")

else:
    st.info("👈 Загрузите файл Excel для начала работы.")
