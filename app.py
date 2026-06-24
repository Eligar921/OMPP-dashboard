import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="ОМПП Дашборд", layout="wide")
st.title("📊 Дашборд ОМПП")

uploaded_files = st.file_uploader(
    "Загрузите один или несколько Excel файлов 'отчет по дате направления'",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    # Объединяем все загруженные файлы
    df_list = []
    for file in uploaded_files:
        df_temp = pd.read_excel(file, sheet_name=0)
        df_temp.columns = df_temp.columns.str.strip()
        df_list.append(df_temp)
    df = pd.concat(df_list, ignore_index=True)

    # Функция поиска столбца по ключевым словам
    def find_column(keywords):
        for col in df.columns:
            col_lower = col.lower()
            for kw in keywords:
                if kw.lower() in col_lower:
                    return col
        return None

    # Поиск всех необходимых столбцов
    col_date_direction = find_column(['дата направления', 'направления на координатора'])
    col_phone = find_column(['телефон'])
    col_recruiter = find_column(['рекрутер'])
    col_source = find_column(['источник омпп', 'источник'])
    col_last_call = find_column(['последнего звонка', 'последний звонок'])
    col_coord_status = find_column(['статус координатора', 'статус координатор'])
    col_lead_status = find_column(['статус лида'])
    col_city = find_column(['город'])
    col_project_group = find_column(['желаемые проекты (группа)', 'группа'])
    col_project_client = find_column(['желаемые проекты (клиент)', 'клиент'])

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
    if col_city is None:
        st.warning("⚠️ Столбец 'Город' не найден. Таблица по городам будет пропущена.")
    if col_project_group is None:
        st.warning("⚠️ Столбец 'Желаемые проекты (Группа)' не найден. Диаграмма проектов будет пропущена.")

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

    # Сбрасываем индекс, чтобы избежать проблем с дублирующимися индексами
    df_filtered = df_filtered.reset_index(drop=True)

    # ---- Таблица: количество кандидатов по рекрутерам (с ограничением ширины) ----
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

    # ---- График: количество направленных по выбранному источнику (горизонтальная столбчатая) ----
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
                color_continuous_scale='Blues'
            )
            fig.update_traces(textposition='outside')
            fig.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
            st.plotly_chart(fig, use_container_width=True)

        # ---- Детальная таблица: рекрутер → источник (кол-во, % от рекрутера, % от всех) ----
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

    # ---- НОВЫЙ БЛОК: Приглашенные по проектам ----
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
                color_continuous_scale='Teal'
            )
            fig_proj.update_traces(textposition='outside')
            fig_proj.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False, height=600)
            st.plotly_chart(fig_proj, use_container_width=True)
        else:
            st.info("Нет данных по проектам.")
    else:
        st.info("Столбец 'Желаемые проекты (Группа)' не найден, диаграмма проектов пропущена.")

    # ---- НОВЫЙ БЛОК: Приглашенные по городам (с исправлением ошибки) ----
    if 'Город' in df_filtered.columns:
        st.subheader("🏙️ Приглашенные по городам")
        # Создаём маску для непустых городов (без использования .loc, чтобы избежать ошибки)
        city_mask = df_filtered['Город'].notna() & (df_filtered['Город'] != '')
        city_data = df_filtered[city_mask].copy()

        if not city_data.empty:
            city_counts = city_data.groupby('Город')['Телефон'].nunique().reset_index()
            city_counts.columns = ['Город', 'Кол-во']
            total_candidates = df_filtered['Телефон'].nunique()
            city_counts['% от всех'] = (city_counts['Кол-во'] / total_candidates * 100).round(1)
            city_counts['% от всех'] = city_counts['% от всех'].astype(str) + '%'
            city_counts = city_counts.sort_values('Кол-во', ascending=False)

            # Отображаем таблицу
            st.dataframe(
                city_counts,
                use_container_width=True,
                column_config={
                    "Город": "Город",
                    "Кол-во": st.column_config.NumberColumn("Кол-во", format="%d"),
                    "% от всех": st.column_config.TextColumn("% от всех"),
                }
            )

            # ---- Диаграмма по городам (увеличенная) ----
            st.subheader("📊 Приглашенные по городам (диаграмма)")
            fig_city = px.bar(
                city_counts,
                x='Кол-во',
                y='Город',
                orientation='h',
                title="Количество направленных кандидатов по городам",
                labels={'Кол-во': 'Кол-во кандидатов', 'Город': ''},
                text='Кол-во',
                color='Кол-во',
                color_continuous_scale='Viridis'
            )
            fig_city.update_traces(textposition='outside')
            fig_city.update_layout(
                yaxis={'categoryorder': 'total ascending'},
                showlegend=False,
                height=700,  # Увеличиваем высоту для читаемости
                width=900     # Можно задать ширину, но обычно use_container_width работает
            )
            st.plotly_chart(fig_city, use_container_width=True)
        else:
            st.info("Нет данных по городам.")
    else:
        st.info("Столбец 'Город' не найден, таблица городов пропущена.")

    # ---- Статистика в сайдбаре ----
    st.sidebar.markdown("---")
    st.sidebar.write(f"🧾 Всего строк после фильтров: **{len(df_filtered)}**")
    st.sidebar.write(f"👥 Уникальных рекрутеров: **{df_filtered['Рекрутер'].nunique()}**")
    st.sidebar.write(f"📞 Уникальных телефонов: **{df_filtered['Телефон'].nunique()}**")

else:
    st.info("👈 Загрузите один или несколько Excel файлов для начала работы.")
