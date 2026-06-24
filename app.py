import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ОМПП Дашборд", layout="wide")
st.title("📊 Дашборд ОМПП")

uploaded_file = st.file_uploader("Загрузите Excel файл 'отчет по дате направления'", type=["xlsx"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file, sheet_name=0)
    df.columns = df.columns.str.strip()

    # ---- Поиск столбцов ----
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
    col_city = find_column(['город'])
    col_project_group = find_column(['желаемые проекты (группа)', 'группа'])
    col_project_client = find_column(['желаемые проекты (клиент)', 'клиент'])
    col_project_first = find_column(['проект первой подтвержденной смены', 'проект первой подтвержденной'])
    col_city_first = find_column(['город первой подтвержденной смены за всю жизнь', 'город первой подтвержденной'])

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
    if col_project_first is not None:
        rename_map[col_project_first] = 'Проект первой подтвержденной смены'
    if col_city_first is not None:
        rename_map[col_city_first] = 'Город первой подтвержденной смены за всю жизнь'

    df = df.rename(columns=rename_map)
    df = df.loc[:, ~df.columns.duplicated()]

    # ---- Преобразование дат ----
    df['Дата направления'] = pd.to_datetime(df['Дата направления'], errors='coerce')
    df['Дата последнего звонка'] = pd.to_datetime(df['Дата последнего звонка'], errors='coerce')

    # ---- Исключаем пустые источники ----
    df['Источник ОМПП'] = df['Источник ОМПП'].astype(str).str.strip()
    df = df[df['Источник ОМПП'].notna() & (df['Источник ОМПП'] != '')]

    # ---- Автофильтр по дате последнего звонка ----
    df['год_напр'] = df['Дата направления'].dt.year
    df['мес_напр'] = df['Дата направления'].dt.month
    df['год_зв'] = df['Дата последнего звонка'].dt.year
    df['мес_зв'] = df['Дата последнего звонка'].dt.month

    same_month = (df['год_зв'] == df['год_напр']) & (df['мес_зв'] == df['мес_напр'])
    prev_month = (df['год_зв'] == df['год_напр']) & (df['мес_зв'] == df['мес_напр'] - 1)
    prev_month_jan = (df['год_зв'] == df['год_напр'] - 1) & (df['мес_напр'] == 1) & (df['мес_зв'] == 12)

    df['filter_last_call'] = same_month | prev_month | prev_month_jan
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

    df_filtered = df_filtered.reset_index(drop=True)

    # ---- 1. Таблица рекрутеров ----
    recruiter_counts = df_filtered.groupby('Рекрутер')['Телефон'].nunique().reset_index()
    recruiter_counts.columns = ['Рекрутер', 'Кол-во кандидатов']
    recruiter_counts = recruiter_counts.sort_values('Кол-во кандидатов', ascending=False)

    st.subheader("📋 Количество направленных кандидатов по рекрутерам")
    st.dataframe(
        recruiter_counts,
        use_container_width=True,
        column_config={
            "Рекрутер": st.column_config.TextColumn("Рекрутер", width="medium"),
            "Кол-во кандидатов": st.column_config.NumberColumn("Кол-во кандидатов", width="small")
        }
    )

    # ---- 2. График по источникам ----
    st.subheader("📊 Кол-во направленных кандидатов по источникам")
    available_sources = sorted(df_filtered['Источник ОМПП'].unique())
    if not available_sources:
        st.warning("Нет доступных источников для отображения.")
    else:
        selected_source_for_chart = st.selectbox("Выберите источник для отображения:", options=available_sources)

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
                color_continuous_scale='Blues',
                height=500
            )
            fig.update_traces(textposition='outside')
            fig.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        # Детальная таблица по источникам
        st.subheader("📋 Детальная разбивка по источникам для каждого рекрутера")
        detail = df_filtered.groupby(['Рекрутер', 'Источник ОМПП'])['Телефон'].nunique().reset_index()
        detail.columns = ['Рекрутер', 'Источник ОМПП', 'Кол-во']

        recruiter_total = detail.groupby('Рекрутер')['Кол-во'].sum().reset_index()
        recruiter_total.columns = ['Рекрутер', 'Всего_рекрутер']

        grand_total = detail['Кол-во'].sum()

        detail = detail.merge(recruiter_total, on='Рекрутер', how='left')
        detail['% от рекрутера'] = (detail['Кол-во'] / detail['Всего_рекрутер'] * 100).round(1)
        detail['% от всех'] = (detail['Кол-во'] / grand_total * 100).round(1)

        detail['% от рекрутера'] = detail['% от рекрутера'].astype(str) + '%'
        detail['% от всех'] = detail['% от всех'].astype(str) + '%'

        detail = detail.sort_values(['Рекрутер', 'Кол-во'], ascending=[True, False])

        st.dataframe(
            detail[['Рекрутер', 'Источник ОМПП', 'Кол-во', '% от рекрутера', '% от всех']],
            use_container_width=True,
            column_config={
                "Рекрутер": "Рекрутер",
                "Источник ОМПП": "Источник",
                "Кол-во": st.column_config.NumberColumn("Кол-во", format="%d"),
                "% от рекрутера": st.column_config.TextColumn("% от рекрутера"),
                "% от всех": st.column_config.TextColumn("% от всех"),
            }
        )

    # ---- 3. Приглашенные по проектам (график) ----
    if 'Желаемые проекты (Группа)' in df_filtered.columns:
        st.subheader("📊 Приглашенные по проектам")
        df_projects = df_filtered.copy()
        if 'Желаемые проекты (Клиент)' in df_projects.columns:
            df_projects['Проект'] = df_projects.apply(
                lambda row: row['Желаемые проекты (Клиент)'] if row['Желаемые проекты (Группа)'] == 'Без группы' else row['Желаемые проекты (Группа)'],
                axis=1
            )
        else:
            df_projects['Проект'] = df_projects['Желаемые проекты (Группа)']

        df_projects = df_projects[df_projects['Проект'].notna() & (df_projects['Проект'] != '')]

        project_counts = df_projects.groupby('Проект')['Телефон'].nunique().reset_index()
        project_counts.columns = ['Проект', 'Кол-во']
        project_counts = project_counts.sort_values('Кол-во', ascending=False)

        if not project_counts.empty:
            fig_proj = px.bar(
                project_counts,
                x='Кол-во',
                y='Проект',
                orientation='h',
                title="Количество направленных кандидатов по проектам",
                labels={'Кол-во': 'Кол-во кандидатов', 'Проект': ''},
                text='Кол-во',
                color='Кол-во',
                color_continuous_scale='Teal',
                height=500
            )
            fig_proj.update_traces(textposition='outside')
            fig_proj.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
            st.plotly_chart(fig_proj, use_container_width=True)
        else:
            st.info("Нет данных по проектам.")
    else:
        st.info("Столбец 'Желаемые проекты (Группа)' не найден, диаграмма проектов пропущена.")

    # ---- 4. ВЫШЕДШИЕ ПО ПРОЕКТАМ (таблица) ----
    # Определяем "вышедших" по статусу координатора или лида
    if 'Статус координатора' in df_filtered.columns:
        df_worked = df_filtered[df_filtered['Статус координатора'] == 'went_work']
    else:
        df_worked = df_filtered[df_filtered['Статус лида'] == 'worked']

    if not df_worked.empty and 'Проект первой подтвержденной смены' in df_worked.columns:
        st.subheader("✅ Вышедшие по проектам")

        source_options = ['Все'] + sorted(df_worked['Источник ОМПП'].unique())
        selected_source_worked = st.selectbox(
            "Выберите источник для фильтрации вышедших:",
            options=source_options,
            key="worked_source_project"
        )

        if selected_source_worked == 'Все':
            df_worked_filtered = df_worked
        else:
            df_worked_filtered = df_worked[df_worked['Источник ОМПП'] == selected_source_worked]

        worked_projects = df_worked_filtered.groupby('Проект первой подтвержденной смены')['Телефон'].nunique().reset_index()
        worked_projects.columns = ['Проект', 'Кол-во вышедших']
        total_worked = worked_projects['Кол-во вышедших'].sum()
        if total_worked > 0:
            worked_projects['% от всех вышедших'] = (worked_projects['Кол-во вышедших'] / total_worked * 100).round(1).astype(str) + '%'
        else:
            worked_projects['% от всех вышедших'] = '0%'
        worked_projects = worked_projects.sort_values('Кол-во вышедших', ascending=False)

        st.dataframe(
            worked_projects,
            use_container_width=True,
            column_config={
                "Проект": "Проект",
                "Кол-во вышедших": st.column_config.NumberColumn("Кол-во вышедших", format="%d"),
                "% от всех вышедших": st.column_config.TextColumn("% от всех вышедших"),
            }
        )
    else:
        st.info("Нет данных о вышедших кандидатах или отсутствует столбец 'Проект первой подтвержденной смены'.")

    # ---- 5. Приглашенные по городам (только таблица) ----
    if 'Город' in df_filtered.columns:
        st.subheader("🏙️ Приглашенные по городам")

        city_data = df_filtered[df_filtered['Город'].notna() & (df_filtered['Город'].astype(str).str.strip() != '')].copy()
        city_data['Город'] = city_data['Город'].astype(str).str.strip()

        if not city_data.empty:
            city_counts = city_data.groupby('Город')['Телефон'].nunique().reset_index()
            city_counts.columns = ['Город', 'Кол-во']
            total_candidates = df_filtered['Телефон'].nunique()
            city_counts['% от всех'] = (city_counts['Кол-во'] / total_candidates * 100).round(1).astype(str) + '%'
            city_counts = city_counts.sort_values('Кол-во', ascending=False)

            st.dataframe(
                city_counts,
                use_container_width=True,
                column_config={
                    "Город": "Город",
                    "Кол-во": st.column_config.NumberColumn("Кол-во", format="%d"),
                    "% от всех": st.column_config.TextColumn("% от всех"),
                }
            )
        else:
            st.info("Нет данных по городам.")
    else:
        st.info("Столбец 'Город' не найден, таблица городов пропущена.")

    # ---- 6. ВЫШЕДШИЕ ПО ГОРОДАМ (таблица) ----
    if not df_worked.empty and 'Город первой подтвержденной смены за всю жизнь' in df_worked.columns:
        st.subheader("✅ Вышедшие по городам")

        source_options_city = ['Все'] + sorted(df_worked['Источник ОМПП'].unique())
        selected_source_worked_city = st.selectbox(
            "Выберите источник для фильтрации вышедших:",
            options=source_options_city,
            key="worked_source_city"
        )

        if selected_source_worked_city == 'Все':
            df_worked_filtered_city = df_worked
        else:
            df_worked_filtered_city = df_worked[df_worked['Источник ОМПП'] == selected_source_worked_city]

        city_worked = df_worked_filtered_city[
            df_worked_filtered_city['Город первой подтвержденной смены за всю жизнь'].notna() &
            (df_worked_filtered_city['Город первой подтвержденной смены за всю жизнь'].astype(str).str.strip() != '')
        ].copy()
        city_worked['Город'] = city_worked['Город первой подтвержденной смены за всю жизнь'].astype(str).str.strip()

        if not city_worked.empty:
            worked_cities = city_worked.groupby('Город')['Телефон'].nunique().reset_index()
            worked_cities.columns = ['Город', 'Кол-во вышедших']
            total_worked_city = worked_cities['Кол-во вышедших'].sum()
            if total_worked_city > 0:
                worked_cities['% от всех вышедших'] = (worked_cities['Кол-во вышедших'] / total_worked_city * 100).round(1).astype(str) + '%'
            else:
                worked_cities['% от всех вышедших'] = '0%'
            worked_cities = worked_cities.sort_values('Кол-во вышедших', ascending=False)

            st.dataframe(
                worked_cities,
                use_container_width=True,
                column_config={
                    "Город": "Город",
                    "Кол-во вышедших": st.column_config.NumberColumn("Кол-во вышедших", format="%d"),
                    "% от всех вышедших": st.column_config.TextColumn("% от всех вышедших"),
                }
            )
        else:
            st.info("Нет данных по городам для вышедших кандидатов.")
    else:
        st.info("Нет данных о вышедших кандидатах или отсутствует столбец 'Город первой подтвержденной смены за всю жизнь'.")

    # ---- Статистика в сайдбаре ----
    st.sidebar.markdown("---")
    st.sidebar.write(f"🧾 Всего строк после фильтров: **{len(df_filtered)}**")
    st.sidebar.write(f"👥 Уникальных рекрутеров: **{df_filtered['Рекрутер'].nunique()}**")
    st.sidebar.write(f"📞 Уникальных телефонов: **{df_filtered['Телефон'].nunique()}**")

else:
    st.info("👈 Загрузите файл Excel для начала работы.")
