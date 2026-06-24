import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ОМПП Дашборд", layout="wide")
st.title("📊 Дашборд ОМПП")

uploaded_file = st.file_uploader("Загрузите Excel файл 'отчет по дате направления'", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df.columns = df.columns.str.strip()

    # Функция поиска столбца по ключевым словам
    def find_column(keywords):
        for col in df.columns:
            col_lower = col.lower()
            for kw in keywords:
                if kw.lower() in col_lower:
                    return col
        return None

    # Поиск необходимых столбцов
    col_date_direction = find_column(['дата направления', 'направления на координатора'])
    col_phone = find_column(['телефон'])
    col_recruiter = find_column(['рекрутер'])
    col_source = find_column(['источник омпп', 'источник'])
    col_last_call = find_column(['последнего звонка', 'последний звонок'])
    col_coord_status = find_column(['статус координатора', 'статус координатор'])
    col_lead_status = find_column(['статус лида'])

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

    # Переименование
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

    # Исключаем пустые источники
    df = df[df['Источник ОМПП'].notna() & (df['Источник ОМПП'] != '')]

    # ---- Автофильтр: дата звонка в том же или предыдущем месяце ----
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

    # ---- Таблица: количество кандидатов по рекрутерам (с ограничением ширины) ----
    recruiter_counts = df_filtered.groupby('Рекрутер')['Телефон'].nunique().reset_index()
    recruiter_counts.columns = ['Рекрутер', 'Кол-во кандидатов']
    recruiter_counts = recruiter_counts.sort_values('Кол-во кандидатов', ascending=False)

    st.subheader("📋 Количество направленных кандидатов по рекрутерам")
    st.dataframe(
        recruiter_counts,
        use_container_width=True,
        column_config={
            "Рекрутер": st.column_config.TextColumn("Рекрутер", width="medium"),  # ограничиваем ширину
            "Кол-во кандидатов": st.column_config.NumberColumn("Кол-во кандидатов", width="small")
        }
    )

    # ---- График: количество направленных по источникам (горизонтальная столбчатая) ----
    st.subheader("📊 Кол-во направленных кандидатов по источникам")
    # Выбор источника для отображения (из числа выбранных в фильтре)
    available_sources = sorted(df_filtered['Источник ОМПП'].unique())
    if not available_sources:
        st.warning("Нет доступных источников для отображения.")
    else:
        selected_source_for_chart = st.selectbox("Выберите источник для отображения:", options=available_sources)

        # Данные для выбранного источника
        df_chart = df_filtered[df_filtered['Источник ОМПП'] == selected_source_for_chart]
        chart_data = df_chart.groupby('Рекрутер')['Телефон'].nunique().reset_index()
        chart_data.columns = ['Рекрутер', 'Кол-во']
        chart_data = chart_data.sort_values('Кол-во', ascending=False)

        if chart_data.empty:
            st.info("Нет данных для выбранного источника.")
        else:
            fig = px.bar(
                chart_data,
                x='Кол-во',
                y='Рекрутер',
                orientation='h',
                title=f"Источник: {selected_source_for_chart}",
                labels={'Кол-во': 'Кол-во направленных кандидатов', 'Рекрутер': ''},
                text='Кол-во',
                color='Кол-во',
                color_continuous_scale='Blues'
            )
            fig.update_traces(textposition='outside')
            fig.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

    # ---- Доля рекрутеров по источникам ----
    st.subheader("📈 Доля рекрутеров, направлявших кандидатов, в разрезе источников")
    # Общее количество уникальных рекрутеров в отфильтрованном наборе
    total_recruiters = df_filtered['Рекрутер'].nunique()

    # Для каждого источника считаем число уникальных рекрутеров
    source_recruiter_counts = df_filtered.groupby('Источник ОМПП')['Рекрутер'].nunique().reset_index()
    source_recruiter_counts.columns = ['Источник ОМПП', 'Рекрутеров в источнике']
    source_recruiter_counts['Общее число рекрутеров'] = total_recruiters
    source_recruiter_counts['Доля (%)'] = (source_recruiter_counts['Рекрутеров в источнике'] / total_recruiters * 100).round(1)

    # Отображаем как таблицу
    st.dataframe(
        source_recruiter_counts,
        use_container_width=True,
        column_config={
            "Источник ОМПП": "Источник",
            "Рекрутеров в источнике": "Рекрутеров",
            "Общее число рекрутеров": "Всего рекрутеров",
            "Доля (%)": st.column_config.NumberColumn("Доля (%)", format="%.1f%%")
        }
    )

    # Визуализация долей (горизонтальная столбчатая)
    if not source_recruiter_counts.empty:
        fig_dol = px.bar(
            source_recruiter_counts,
            x='Доля (%)',
            y='Источник ОМПП',
            orientation='h',
            title="Доля рекрутеров, участвовавших в каждом источнике",
            labels={'Доля (%)': 'Доля рекрутеров (%)', 'Источник ОМПП': ''},
            text='Доля (%)',
            color='Доля (%)',
            color_continuous_scale='Viridis'
        )
        fig_dol.update_traces(textposition='outside', texttemplate='%{text:.1f}%')
        fig_dol.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
        st.plotly_chart(fig_dol, use_container_width=True)
    else:
        st.info("Нет данных для расчёта долей.")

    # ---- Статистика в сайдбаре ----
    st.sidebar.markdown("---")
    st.sidebar.write(f"🧾 Всего строк после фильтров: **{len(df_filtered)}**")
    st.sidebar.write(f"👥 Уникальных рекрутеров: **{df_filtered['Рекрутер'].nunique()}**")
    st.sidebar.write(f"📞 Уникальных телефонов: **{df_filtered['Телефон'].nunique()}**")

else:
    st.info("👈 Загрузите файл Excel для начала работы.")
