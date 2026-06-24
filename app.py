import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ОМПП Дашборд", layout="wide")
st.title("📊 Дашборд ОМПП")

# ---- Функция нормализации названий проектов ----
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

# ---- Функция поиска столбца ----
def find_column(df, keywords, exact_match=None):
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

# ---- Загрузка основного файла (направления) ----
uploaded_file = st.file_uploader("Загрузите Excel файл 'отчет по дате направления'", type=["xlsx"], key="main")

# ---- Загрузка файла KPI (вышедшие) ----
uploaded_file_kpi = st.file_uploader("Загрузите Excel файл KPI (отчет по вышедшим)", type=["xlsx"], key="kpi")

# ---- Инициализация DataFrame ----
df_main = None
df_kpi = None

# ---- Обработка основного файла ----
if uploaded_file is not None:
    df_main = pd.read_excel(uploaded_file, sheet_name=0)
    df_main.columns = df_main.columns.str.strip()

    # Поиск столбцов для основного отчета
    col_date_direction = find_column(df_main, ['дата направления', 'направления на координатора'])
    col_phone = find_column(df_main, ['телефон'])
    col_recruiter = find_column(df_main, ['рекрутер'])
    col_source = find_column(df_main, ['источник омпп', 'источник'])
    col_last_call = find_column(
        df_main,
        ['последнего звонка до первого статуса', 'последнего звонка', 'последний звонок'],
        exact_match='Дата последнего звонка до первого статуса первой смены'
    )
    col_coord_status = find_column(df_main, ['статус координатора', 'статус координатор'])
    col_lead_status = find_column(df_main, ['статус лида'])
    col_city = find_column(df_main, ['город'])
    col_project_group = find_column(df_main, ['желаемые проекты (группа)', 'группа'])
    col_project_client = find_column(df_main, ['желаемые проекты (клиент)', 'клиент'])
    col_project_first = find_column(df_main, ['проект первой подтвержденной смены', 'проект первой подтвержденной'])
    col_city_first = find_column(df_main, ['город первой подтвержденной смены за всю жизнь', 'город первой подтвержденной'])
    col_date_first_shift = find_column(df_main, ['дата первой подтвержденной смены за всю жизнь', 'дата первой подтвержденной'])

    if col_date_direction is None:
        st.error("❌ В основном файле не найден столбец с датой направления.")
        st.stop()
    if col_phone is None:
        st.error("❌ В основном файле не найден столбец 'Телефон'")
        st.stop()
    if col_recruiter is None:
        st.error("❌ В основном файле не найден столбец 'Рекрутер'")
        st.stop()
    if col_source is None:
        st.error("❌ В основном файле не найден столбец 'Источник ОМПП'")
        st.stop()
    if col_last_call is None:
        st.error("❌ В основном файле не найден столбец с датой последнего звонка")
        st.stop()
    if col_coord_status is None and col_lead_status is None:
        st.error("❌ В основном файле не найден ни столбец 'Статус координатора', ни 'Статус лида'")
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

    df_main = df_main.rename(columns=rename_map)
    df_main = df_main.loc[:, ~df_main.columns.duplicated()]

    # Преобразование дат
    df_main['Дата направления'] = pd.to_datetime(df_main['Дата направления'], errors='coerce')
    df_main['Дата последнего звонка'] = pd.to_datetime(df_main['Дата последнего звонка'], errors='coerce')

    # Исключаем пустые источники
    df_main['Источник ОМПП'] = df_main['Источник ОМПП'].astype(str).str.strip()
    df_main = df_main[df_main['Источник ОМПП'].notna() & (df_main['Источник ОМПП'] != '')]

# ---- Обработка файла KPI ----
if uploaded_file_kpi is not None:
    df_kpi = pd.read_excel(uploaded_file_kpi, sheet_name=0)
    df_kpi.columns = df_kpi.columns.str.strip()

    # Поиск столбцов для KPI
    col_kpi_phone = find_column(df_kpi, ['телефон гигера', 'телефон'])
    col_kpi_recruiter = find_column(df_kpi, ['рекрутер'])
    col_kpi_source = find_column(df_kpi, ['источник омпп'])
    col_kpi_date_first_shift = find_column(df_kpi, ['дата первой смены'])
    col_kpi_last_call = find_column(
        df_kpi,
        ['последнего звонка до первого статуса первой подтвержденной смены', 'последнего звонка'],
        exact_match='Дата последнего звонка до первого статуса первой подтвержденной смены'
    )

    if col_kpi_phone is None:
        st.error("❌ В файле KPI не найден столбец 'Телефон гигера'")
        st.stop()
    if col_kpi_recruiter is None:
        st.error("❌ В файле KPI не найден столбец 'Рекрутер'")
        st.stop()
    if col_kpi_source is None:
        st.error("❌ В файле KPI не найден столбец 'Источник ОМПП'")
        st.stop()
    if col_kpi_date_first_shift is None:
        st.error("❌ В файле KPI не найден столбец 'Дата первой смены'")
        st.stop()
    if col_kpi_last_call is None:
        st.error("❌ В файле KPI не найден столбец с датой последнего звонка до первого статуса")
        st.stop()

    # Переименование
    rename_kpi = {
        col_kpi_phone: 'Телефон гигера',
        col_kpi_recruiter: 'Рекрутер',
        col_kpi_source: 'Источник ОМПП',
        col_kpi_date_first_shift: 'Дата первой смены',
        col_kpi_last_call: 'Дата последнего звонка',
    }
    df_kpi = df_kpi.rename(columns=rename_kpi)
    df_kpi = df_kpi.loc[:, ~df_kpi.columns.duplicated()]

    # Преобразование дат
    df_kpi['Дата первой смены'] = pd.to_datetime(df_kpi['Дата первой смены'], errors='coerce')
    df_kpi['Дата последнего звонка'] = pd.to_datetime(df_kpi['Дата последнего звонка'], errors='coerce')

    # Исключаем пустые источники
    df_kpi['Источник ОМПП'] = df_kpi['Источник ОМПП'].astype(str).str.strip()
    df_kpi = df_kpi[df_kpi['Источник ОМПП'].notna() & (df_kpi['Источник ОМПП'] != '')]

# ---- Боковая панель с фильтрами ----
st.sidebar.header("Фильтры")

# Общий фильтр по источнику ОМПП (если есть хотя бы один загруженный файл)
all_sources = []
if df_main is not None:
    all_sources.extend(df_main['Источник ОМПП'].unique())
if df_kpi is not None:
    all_sources.extend(df_kpi['Источник ОМПП'].unique())
all_sources = sorted(set(all_sources))

if all_sources:
    selected_sources = st.sidebar.multiselect("Источник ОМПП", options=all_sources, default=all_sources)
else:
    selected_sources = []
    st.sidebar.warning("Загрузите хотя бы один файл для выбора источников.")

# Фильтр по дате направления (для основного отчета)
if df_main is not None:
    st.sidebar.subheader("Фильтр по дате направления")
    min_date_main = df_main['Дата направления'].min().date()
    max_date_main = df_main['Дата направления'].max().date()
    date_range_main = st.sidebar.date_input(
        "Диапазон дат направления",
        value=(min_date_main, max_date_main),
        min_value=min_date_main,
        max_value=max_date_main,
        key="date_range_main"
    )
else:
    date_range_main = None

# Фильтр по дате первой смены (для KPI)
if df_kpi is not None:
    st.sidebar.subheader("Фильтр по дате первой смены")
    min_date_kpi = df_kpi['Дата первой смены'].min().date()
    max_date_kpi = df_kpi['Дата первой смены'].max().date()
    date_range_kpi = st.sidebar.date_input(
        "Диапазон дат первой смены",
        value=(min_date_kpi, max_date_kpi),
        min_value=min_date_kpi,
        max_value=max_date_kpi,
        key="date_range_kpi"
    )
else:
    date_range_kpi = None

# ---- Применение фильтров к основному отчету ----
if df_main is not None and selected_sources:
    df_main_filtered = df_main[df_main['Источник ОМПП'].isin(selected_sources)]

    # Применяем фильтр по дате направления
    if date_range_main and len(date_range_main) == 2:
        start_date, end_date = date_range_main
        df_main_filtered = df_main_filtered[
            (df_main_filtered['Дата направления'].dt.date >= start_date) &
            (df_main_filtered['Дата направления'].dt.date <= end_date)
        ]

    # Автофильтр по дате последнего звонка (текущий или предыдущий месяц от даты направления)
    df_main_filtered['год_напр'] = df_main_filtered['Дата направления'].dt.year
    df_main_filtered['мес_напр'] = df_main_filtered['Дата направления'].dt.month
    df_main_filtered['год_зв'] = df_main_filtered['Дата последнего звонка'].dt.year
    df_main_filtered['мес_зв'] = df_main_filtered['Дата последнего звонка'].dt.month

    same_month = (df_main_filtered['год_зв'] == df_main_filtered['год_напр']) & (df_main_filtered['мес_зв'] == df_main_filtered['мес_напр'])
    prev_month = (df_main_filtered['год_зв'] == df_main_filtered['год_напр']) & (df_main_filtered['мес_зв'] == df_main_filtered['мес_напр'] - 1)
    prev_month_jan = (df_main_filtered['год_зв'] == df_main_filtered['год_напр'] - 1) & (df_main_filtered['мес_напр'] == 1) & (df_main_filtered['мес_зв'] == 12)

    df_main_filtered['filter_last_call'] = same_month | prev_month | prev_month_jan
    df_main_filtered = df_main_filtered[df_main_filtered['filter_last_call'] & df_main_filtered['Дата последнего звонка'].notna()]
    df_main_filtered = df_main_filtered.reset_index(drop=True)
else:
    df_main_filtered = None

# ---- Применение фильтров к KPI ----
if df_kpi is not None and selected_sources:
    df_kpi_filtered = df_kpi[df_kpi['Источник ОМПП'].isin(selected_sources)]

    # Применяем фильтр по дате первой смены
    if date_range_kpi and len(date_range_kpi) == 2:
        start_date_kpi, end_date_kpi = date_range_kpi
        df_kpi_filtered = df_kpi_filtered[
            (df_kpi_filtered['Дата первой смены'].dt.date >= start_date_kpi) &
            (df_kpi_filtered['Дата первой смены'].dt.date <= end_date_kpi)
        ]

    # Автофильтр по дате последнего звонка (текущий или предыдущий месяц от даты первой смены)
    df_kpi_filtered['год_смены'] = df_kpi_filtered['Дата первой смены'].dt.year
    df_kpi_filtered['мес_смены'] = df_kpi_filtered['Дата первой смены'].dt.month
    df_kpi_filtered['год_зв'] = df_kpi_filtered['Дата последнего звонка'].dt.year
    df_kpi_filtered['мес_зв'] = df_kpi_filtered['Дата последнего звонка'].dt.month

    same_month_kpi = (df_kpi_filtered['год_зв'] == df_kpi_filtered['год_смены']) & (df_kpi_filtered['мес_зв'] == df_kpi_filtered['мес_смены'])
    prev_month_kpi = (df_kpi_filtered['год_зв'] == df_kpi_filtered['год_смены']) & (df_kpi_filtered['мес_зв'] == df_kpi_filtered['мес_смены'] - 1)
    prev_month_jan_kpi = (df_kpi_filtered['год_зв'] == df_kpi_filtered['год_смены'] - 1) & (df_kpi_filtered['мес_смены'] == 1) & (df_kpi_filtered['мес_зв'] == 12)

    df_kpi_filtered['filter_last_call'] = same_month_kpi | prev_month_kpi | prev_month_jan_kpi
    df_kpi_filtered = df_kpi_filtered[df_kpi_filtered['filter_last_call'] & df_kpi_filtered['Дата последнего звонка'].notna()]
    df_kpi_filtered = df_kpi_filtered.reset_index(drop=True)
else:
    df_kpi_filtered = None

# ---- ОТОБРАЖЕНИЕ БЛОКОВ ----

if df_main_filtered is not None:
    # ---- 1. Таблица рекрутеров (основной отчет) ----
    recruiter_counts = df_main_filtered.groupby('Рекрутер')['Телефон'].nunique().reset_index()
    recruiter_counts.columns = ['Рекрутер', 'Кол-во кандидатов']

    # Считаем количество вышедших (по дате первой подтвержденной смены)
    if 'Дата первой подтвержденной смены за всю жизнь' in df_main_filtered.columns:
        df_with_shift = df_main_filtered[df_main_filtered['Дата первой подтвержденной смены за всю жизнь'].notna()]
        worked_counts = df_with_shift.groupby('Рекрутер')['Телефон'].nunique().reset_index()
        worked_counts.columns = ['Рекрутер', 'Вышло из приглашенных']
        recruiter_counts = recruiter_counts.merge(worked_counts, on='Рекрутер', how='left').fillna(0)
        recruiter_counts['Вышло из приглашенных'] = recruiter_counts['Вышло из приглашенных'].astype(int)
    else:
        recruiter_counts['Вышло из приглашенных'] = 0

    recruiter_counts['Конверсия, %'] = (recruiter_counts['Вышло из приглашенных'] / recruiter_counts['Кол-во кандидатов'] * 100).round(1)
    recruiter_counts['Конверсия, %'] = recruiter_counts['Конверсия, %'].fillna(0).astype(str) + '%'
    recruiter_counts = recruiter_counts.sort_values('Кол-во кандидатов', ascending=False)

    st.subheader("📋 Количество направленных кандидатов по рекрутерам")
    st.dataframe(
        recruiter_counts,
        use_container_width=True,
        column_config={
            "Рекрутер": st.column_config.TextColumn("Рекрутер", width="medium"),
            "Кол-во кандидатов": st.column_config.NumberColumn("Кол-во кандидатов", width="small"),
            "Вышло из приглашенных": st.column_config.NumberColumn("Вышло из приглашенных", width="small"),
            "Конверсия, %": st.column_config.TextColumn("Конверсия, %", width="small"),
        }
    )

    # ---- 2. График по источникам ----
    st.subheader("📊 Кол-во направленных кандидатов по источникам")
    available_sources = sorted(df_main_filtered['Источник ОМПП'].unique())
    if not available_sources:
        st.warning("Нет доступных источников для отображения.")
    else:
        selected_source_for_chart = st.selectbox("Выберите источник для отображения:", options=available_sources, key="source_chart")

        df_chart = df_main_filtered[df_main_filtered['Источник ОМПП'] == selected_source_for_chart]
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
        detail = df_main_filtered.groupby(['Рекрутер', 'Источник ОМПП'])['Телефон'].nunique().reset_index()
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
    if 'Желаемые проекты (Группа)' in df_main_filtered.columns:
        st.subheader("📊 Приглашенные по проектам")
        df_projects = df_main_filtered.copy()
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
    if 'Статус координатора' in df_main_filtered.columns:
        df_worked = df_main_filtered[df_main_filtered['Статус координатора'] == 'went_work']
    else:
        df_worked = df_main_filtered[df_main_filtered['Статус лида'] == 'worked']

    if not df_worked.empty and 'Проект первой подтвержденной смены' in df_worked.columns:
        st.subheader("✅ Вышедшие по проектам (из приглашенных)")

        source_options = ['Все'] + sorted(df_worked['Источник ОМПП'].unique())
        selected_source_worked = st.selectbox(
            "Выберите источник для фильтрации вышедших:",
            options=source_options,
            key="worked_source_project"
        )

        if selected_source_worked == 'Все':
            df_worked_filtered = df_worked
            df_all_filtered = df_main_filtered
        else:
            df_worked_filtered = df_worked[df_worked['Источник ОМПП'] == selected_source_worked]
            df_all_filtered = df_main_filtered[df_main_filtered['Источник ОМПП'] == selected_source_worked]

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

        df_worked_filtered['Проект'] = df_worked_filtered['Проект первой подтвержденной смены'].apply(normalize_project)
        worked_counts = df_worked_filtered.groupby('Проект')['Телефон'].nunique().reset_index()
        worked_counts.columns = ['Проект', 'Кол-во вышедших']

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

    # ---- 5. Приглашенные по городам (только таблица) ----
    if 'Город' in df_main_filtered.columns:
        st.subheader("🏙️ Приглашенные по городам")

        city_data = df_main_filtered[df_main_filtered['Город'].notna() & (df_main_filtered['Город'].astype(str).str.strip() != '')].copy()
        city_data['Город'] = city_data['Город'].astype(str).str.strip()

        if not city_data.empty:
            city_counts = city_data.groupby('Город')['Телефон'].nunique().reset_index()
            city_counts.columns = ['Город', 'Кол-во']
            total_candidates = df_main_filtered['Телефон'].nunique()
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

    # ---- 6. Вышедшие по городам из приглашенных ----
    if not df_worked.empty and 'Город первой подтвержденной смены за всю жизнь' in df_worked.columns:
        st.subheader("✅ Вышедшие по городам из приглашенных")

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

# ---- НОВЫЙ БЛОК: Вышедшие по рекрутерам (с дошедшими) из KPI ----
if df_kpi_filtered is not None and not df_kpi_filtered.empty:
    st.subheader("📊 Вышедшие по рекрутерам (с дошедшими)")

    # Группировка по рекрутеру, подсчет уникальных телефонов гигеров
    kpi_counts = df_kpi_filtered.groupby('Рекрутер')['Телефон гигера'].nunique().reset_index()
    kpi_counts.columns = ['Рекрутер', 'Кол-во вышедших гигеров']
    kpi_counts = kpi_counts.sort_values('Кол-во вышедших гигеров', ascending=False)

    st.dataframe(
        kpi_counts,
        use_container_width=True,
        column_config={
            "Рекрутер": st.column_config.TextColumn("Рекрутер", width="medium"),
            "Кол-во вышедших гигеров": st.column_config.NumberColumn("Кол-во вышедших гигеров", format="%d"),
        }
    )
elif df_kpi is not None and df_kpi_filtered is not None and df_kpi_filtered.empty:
    st.info("Нет данных по KPI после применения фильтров.")
elif df_kpi is None:
    st.info("Для отображения блока 'Вышедшие по рекрутерам' загрузите файл KPI.")

# ---- Статистика в сайдбаре (для основного отчета) ----
if df_main_filtered is not None:
    st.sidebar.markdown("---")
    st.sidebar.write(f"🧾 Основной отчет: строк после фильтров: **{len(df_main_filtered)}**")
    st.sidebar.write(f"👥 Уникальных рекрутеров: **{df_main_filtered['Рекрутер'].nunique()}**")
    st.sidebar.write(f"📞 Уникальных телефонов: **{df_main_filtered['Телефон'].nunique()}**")

if df_kpi_filtered is not None:
    st.sidebar.markdown("---")
    st.sidebar.write(f"🧾 KPI отчет: строк после фильтров: **{len(df_kpi_filtered)}**")
    st.sidebar.write(f"👥 Уникальных рекрутеров: **{df_kpi_filtered['Рекрутер'].nunique()}**")
    st.sidebar.write(f"📞 Уникальных телефонов гигеров: **{df_kpi_filtered['Телефон гигера'].nunique()}**")
