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

    df['Дата направления'] = pd.to_datetime(df['Дата направления'], errors='coerce')
    df['Дата последнего звонка'] = pd.to_datetime(df['Дата последнего звонка'], errors='coerce')

    # Удаляем пустые источники
    df = df[df['Источник ОМПП'].notna() & (df['Источник ОМПП'] != '')]

    # ---- Автофильтр по дате звонка (текущий/предыдущий месяц) ----
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

    # ---- Таблица: количество направленных по рекрутерам (с ограничением ширины) ----
    recruiter_counts = df_filtered.groupby('Рекрутер')['Телефон'].nunique().reset_index()
    recruiter_counts.columns = ['Рекрутер', 'Кол-во кандидатов']
    recruiter_counts = recruiter_counts.sort_values('Кол-во кандидатов', ascending=False)

    st.subheader("📋 Количество направленных кандидатов по рекрутерам")

    # Ограничиваем ширину колонок: задаём максимальную ширину в пикселях
    st.dataframe(
        recruiter_counts,
        use_container_width=True,
        column_config={
            "Рекрутер": st.column_config.TextColumn(
                "Рекрутер",
                width="medium",  # можно small, medium, large
                help="ФИО рекрутера"
            ),
            "Кол-во кандидатов": st.column_config.NumberColumn(
                "Кол-во кандидатов",
                width="small",
                help="Количество уникальных телефонов"
            )
        }
    )

    # ---- График: горизонтальная столбчатая диаграмма по выбранному источнику ----
    st.subheader("📊 Количество направленных кандидатов по источникам")

    # Выбор источника (если доступны источники)
    available_sources = sorted(df_filtered['Источник ОМПП'].unique())
    if len(available_sources) == 0:
        st.warning("Нет данных для отображения графика.")
    else:
        # Пользователь выбирает один источник для отображения
        selected_source = st.selectbox("Выберите источник для отображения на графике", options=available_sources)

        # Фильтруем данные по выбранному источнику
        df_source = df_filtered[df_filtered['Источник ОМПП'] == selected_source]

        # Группируем по рекрутерам и считаем количество телефонов
        source_counts = df_source.groupby('Рекрутер')['Телефон'].nunique().reset_index()
        source_counts.columns = ['Рекрутер', 'Кол-во направленных']

        # Если нет данных – покажем предупреждение
        if source_counts.empty:
            st.info(f"Нет направленных кандидатов по источнику '{selected_source}'")
        else:
            # Сортируем для наглядности
            source_counts = source_counts.sort_values('Кол-во направленных', ascending=True)

            # Горизонтальная столбчатая диаграмма
            fig = px.bar(
                source_counts,
                x='Кол-во направленных',
                y='Рекрутер',
                orientation='h',
                title=f"Количество направленных кандидатов по источнику: {selected_source}",
                text='Кол-во направленных',  # отображаем значения на столбцах
                color='Кол-во направленных',
                color_continuous_scale='Blues'
            )
            fig.update_traces(textposition='outside')
            fig.update_layout(
                xaxis_title="Количество кандидатов",
                yaxis_title="Рекрутер",
                yaxis={'categoryorder': 'total ascending'},  # сортировка по убыванию сверху
                height=600
            )
            st.plotly_chart(fig, use_container_width=True)

    # ---- Статистика в сайдбаре ----
    st.sidebar.markdown("---")
    st.sidebar.write(f"🧾 Всего строк после фильтров: **{len(df_filtered)}**")
    st.sidebar.write(f"👥 Уникальных рекрутеров: **{df_filtered['Рекрутер'].nunique()}**")
    st.sidebar.write(f"📞 Уникальных телефонов: **{df_filtered['Телефон'].nunique()}**")

else:
    st.info("👈 Загрузите файл Excel для начала работы.")
