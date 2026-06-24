import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ОМПП Дашборд", layout="wide")
st.title("📊 Дашборд ОМПП")

# Функция нормализации названий проектов
def normalize_project(project_name):
    if not isinstance(project_name, str):
        return project_name
    project_name = project_name.strip()
    if project_name == 'Пятёрочка агентская':
        return 'Пятёрочка'
    if project_name == 'Магнит Косметик':
        return 'Магнит'
    if project_name == 'ООО "Таймбук"':
        return 'Гулливер'
    return project_name

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
    col_date_first_shift = find_column(['дата первой подтвержденной смены за всю жизнь', 'дата первой подтвержденной'])

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
    if col_date_first_shift is not None:
        rename_map[col_date_first_shift] = 'Дата первой подтвержденной смены за всю жизнь'

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

    # ---- 1. Таблица рекрутеров (с добавленным столбцом "Вышло из приглашенных") ----
    # Считаем количество приглашенных (уникальные телефоны)
    recruiter_counts = df_filtered.groupby('Рекрутер')['Телефон'].nunique().reset_index()
    recruiter_counts.columns = ['Рекрутер', 'Кол-во кандидатов']

    # Считаем количество вышедших (уникальные телефоны, у которых есть дата первой подтвержденной смены)
    if 'Дата первой подтвержденной смены за всю жизнь' in df_filtered.columns:
        df_with_shift = df_filtered[df_filtered['Дата первой подтвержденной смены за всю жизнь'].notna()]
        worked_counts = df_with_shift.groupby('Рекрутер')['Телефон'].nunique().reset_index()
        worked_counts.columns = ['Рекрутер', 'Вышло из приглашенных']
        recruiter_counts = recruiter_counts.merge(worked_counts, on='Рекрутер', how='left').fillna(0)
        recruiter_counts['Вышло из приглашенных'] = recruiter_counts['Вышло из приглашенных'].astype(int)
    else:
        recruiter_counts['Вышло из приглашенных'] = 0

    recruiter_counts = recruiter_counts.sort_values('Кол-во кандидатов', ascending=False)

    st.subheader("📋 Количество направленных кандидатов по рекрутерам")
    st.dataframe(
        recruiter_counts,
        use_container_width=True,
        column_config={
            "Рекрутер": st.column_config.TextColumn("Рекрутер", width="medium"),
            "Кол-во кандидатов": st.column_config.NumberColumn("Кол-во кандидатов", width="small"),
            "Вышло из приглашенных": st.column_config.NumberColumn("Вышло из приглашенных", width="small"),
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
        # Формируем столбец "Проект" с учётом "Без группы" и нормализацией
        if 'Желаемые проекты (Клиент)' in df_projects.columns:
            df_projects['Проект'] = df_projects.apply(
                lambda row: normalize_project(row['Желаемые проекты (Клиент)']) if row['Желаемые проекты (Группа)'] == 'Без группы' else normalize_project(row['Желаемые проекты (Группа)']),
                axis=1
            )
        else:
            df_projects['Проект'] = df_projects['Желаемые проекты (Группа)'].apply(normalize_project)

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

    # ---- 4. Вышедшие по проектам (из приглашенных) ----
    # Определяем "вышедших" по статусу координатора или лида
    if 'Статус координатора' in df_filtered.columns:
        df_worked = df_filtered[df_filtered['Статус координатора'] == 'went_work']
    else:
        df_worked = df_filtered[df_filtered['Статус лида'] == 'worked']

    if not df_worked.empty and 'Проект первой подтвержденной смены' in df_worked.columns:
        st.subheader("✅ Вышедшие по проектам (из приглашенных)")

        source_options = ['Все'] + sorted(df_worked['Источник ОМПП'].unique())
        selected_source_worked = st.selectbox(
            "Выберите источник для фильтрации вышедших:",
            options=source_options,
            key="worked_source_project"
        )

        # Фильтруем данные по источнику
        if selected_source_worked == 'Все':
            df_worked_filtered = df_worked
            df_all_filtered = df_filtered
        else:
            df_worked_filtered = df_worked[df_worked['Источник ОМПП'] == selected_source_worked]
            df_all_filtered = df_filtered[df_filtered['Источник ОМПП'] == selected_source_worked]

        # Приглашенные по проектам (все направленные) с нормализацией
        df_projects_all = df_all_filtered.copy()
        if 'Желаемые проекты (Клиент)' in df_projects_all.columns:
            df_projects_all['Проект'] = df_projects_all.apply(
                lambda row: normalize_project(row['Желаемые проекты (Клиент)']) if row['Желаемые проекты (Группа)'] == 'Без группы' else normalize_project(row['Желаемые проекты (Группа)']),
                axis=1
            )
        else:
            df_projects_all['Проект'] = df_projects_all['Желаемые проекты (Группа)'].apply(normalize_project)

        df_projects_all = df_projects_all[df_projects_all['Проект'].notna() & (df_projects_all['Проект'] != '')]
        invited_counts = df_projects_all.groupby('Проект')['Телефон'].nunique().reset_index()
        invited_counts.columns = ['Проект', 'Кол-во приглашенных']

        # Вышедшие по проекту (нормализуем название проекта)
        df_worked_filtered['Проект'] = df_worked_filtered['Проект первой подтвержденной смены'].apply(normalize_project)
        worked_counts = df_worked_filtered.groupby('Проект')['Телефон'].nunique().reset_index()
        worked_counts.columns = ['Проект', 'Кол-во вышедших']

        # Объединяем
        merged = pd.merge(invited_counts, worked_counts, on='Проект', how='outer').fillna(0)
        merged['Кол-во приглашенных'] = merged['Кол-во приглашенных'].astype(int)
        merged['Кол-во вышедших'] = merged['Кол-во вышедших'].astype(int)
        merged['Конверсия, %'] = (merged['Кол-во вышедших'] / merged['Кол-во приглашенных'] * 100).round(1)
        merged['Конверсия, %'] = merged['Конверсия, %'].fillna(0).astype(str) + '%'
        merged = merged.sort_values('Кол-во приглашенных', ascending=False)

        st.dataframe(
            merged,
            use_container_width=True,
            column_config={
                "Проект": "Проект",
                "Кол-во приглашенных": st.column_config.NumberColumn("Кол-во приглашенных", format="%d"),
                "Кол-во вышедших": st.column_config.NumberColumn("Кол-во вышедших", format="%d"),
                "Конверсия, %": st.column_config.TextColumn("Конверсия, %"),
            }
        )
    else:
        st.info("Нет данных о вышедших кандидатах или отсутствует столбец 'Проект первой подтвержденной смены'.")

    # ---- 5. Вышедшие по городам из приглашенных ----
    if 'Город' in df_filtered.columns:
        st.subheader("✅ Вышедшие по городам из приглашенных")

        # Приглашенные: группировка по "Город"
        city_invited = df_filtered[df_filtered['Город'].notna() & (df_filtered['Город'].astype(str).str.strip() != '')].copy()
        city_invited['Город'] = city_invited['Город'].astype(str).str.strip()
        invited_city = city_invited.groupby('Город')['Телефон'].nunique().reset_index()
        invited_city.columns = ['Город', 'Кол-во приглашенных']

        # Вышедшие: группировка по "Город первой подтвержденной смены за всю жизнь"
        if not df_worked.empty and 'Город первой подтвержденной смены за всю жизнь' in df_worked.columns:
            city_worked = df_worked[df_worked['Город первой подтвержденной смены за всю жизнь'].notna() &
                                    (df_worked['Город первой подтвержденной смены за всю жизнь'].astype(str).str.strip() != '')].copy()
            city_worked['Город'] = city_worked['Город первой подтвержденной смены за всю жизнь'].astype(str).str.strip()
            worked_city = city_worked.groupby('Город')['Телефон'].nunique().reset_index()
            worked_city.columns = ['Город', 'Кол-во вышедших']
        else:
            worked_city = pd.DataFrame(columns=['Город', 'Кол-во вышедших'])

        # Объединяем
        merged_city = pd.merge(invited_city, worked_city, on='Город', how='outer').fillna(0)
        merged_city['Кол-во приглашенных'] = merged_city['Кол-во приглашенных'].astype(int)
        merged_city['Кол-во вышедших'] = merged_city['Кол-во вышедших'].astype(int)
        merged_city['Конверсия, %'] = (merged_city['Кол-во вышедших'] / merged_city['Кол-во приглашенных'] * 100).round(1)
        merged_city['Конверсия, %'] = merged_city['Конверсия, %'].fillna(0).astype(str) + '%'
        merged_city = merged_city.sort_values('Кол-во приглашенных', ascending=False)

        st.dataframe(
            merged_city,
            use_container_width=True,
            column_config={
                "Город": "Город",
                "Кол-во приглашенных": st.column_config.NumberColumn("Кол-во приглашенных", format="%d"),
                "Кол-во вышедших": st.column_config.NumberColumn("Кол-во вышедших", format="%d"),
                "Конверсия, %": st.column_config.TextColumn("Конверсия, %"),
            }
        )
    else:
        st.info("Столбец 'Город' не найден, таблица городов пропущена.")

    # ---- Статистика в сайдбаре ----
    st.sidebar.markdown("---")
    st.sidebar.write(f"🧾 Всего строк после фильтров: **{len(df_filtered)}**")
    st.sidebar.write(f"👥 Уникальных рекрутеров: **{df_filtered['Рекрутер'].nunique()}**")
    st.sidebar.write(f"📞 Уникальных телефонов: **{df_filtered['Телефон'].nunique()}**")

else:
    st.info("👈 Загрузите файл Excel для начала работы.")
