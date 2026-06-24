import streamlit as st
import pandas as pd
import plotly.express as px
import re
from datetime import datetime, date

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

# ---- Функция для парсинга даты из заголовков "Диаграмма" ----
def parse_diagram_date(date_str):
    if not isinstance(date_str, str):
        return pd.NaT
    date_str = date_str.strip()
    try:
        return pd.to_datetime(date_str, errors='coerce')
    except:
        pass
    month_map = {
        'янв': 'Jan', 'фев': 'Feb', 'мар': 'Mar', 'апр': 'Apr',
        'май': 'May', 'июн': 'Jun', 'июл': 'Jul', 'авг': 'Aug',
        'сен': 'Sep', 'окт': 'Oct', 'ноя': 'Nov', 'дек': 'Dec'
    }
    match = re.match(r'^(\d{2})\.(\w{3})$', date_str)
    if match:
        day, month_ru = match.groups()
        month_en = month_map.get(month_ru.lower())
        if month_en:
            year = 2026
            try:
                return pd.to_datetime(f"{day} {month_en} {year}", format='%d %b %Y')
            except:
                return pd.NaT
    return pd.NaT

# ---- Загрузка основного файла (направления) ----
uploaded_file = st.file_uploader("Загрузите Excel файл 'отчет по дате направления'", type=["xlsx"], key="main")

# ---- Загрузка файла KPI (вышедшие) ----
uploaded_file_kpi = st.file_uploader("Загрузите Excel файл KPI (отчет по вышедшим)", type=["xlsx"], key="kpi")

# ---- Загрузка файла "отчет по обработке откликов" ----
uploaded_file_responses = st.file_uploader("Загрузите Excel файл 'отчет по обработке откликов'", type=["xlsx"], key="responses")

# ---- Инициализация DataFrame ----
df_main = None
df_kpi = None
df_responses = None
df_diagram = None

# ---- Обработка основного файла ----
if uploaded_file is not None:
    df_main = pd.read_excel(uploaded_file, sheet_name=0)
    df_main.columns = df_main.columns.str.strip()

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

    df_main['Дата направления'] = pd.to_datetime(df_main['Дата направления'], errors='coerce')
    df_main['Дата последнего звонка'] = pd.to_datetime(df_main['Дата последнего звонка'], errors='coerce')

    df_main['Источник ОМПП'] = df_main['Источник ОМПП'].astype(str).str.strip()
    df_main = df_main[df_main['Источник ОМПП'].notna() & (df_main['Источник ОМПП'] != '')]

# ---- Обработка файла KPI ----
if uploaded_file_kpi is not None:
    df_kpi = pd.read_excel(uploaded_file_kpi, sheet_name=0)
    df_kpi.columns = df_kpi.columns.str.strip()

    col_kpi_phone = find_column(df_kpi, ['телефон гигера', 'телефон'])
    col_kpi_recruiter = find_column(df_kpi, ['рекрутер'])
    col_kpi_source = find_column(df_kpi, ['источник омпп'])
    col_kpi_date_first_shift = find_column(df_kpi, ['дата первой смены'])
    col_kpi_last_call = find_column(
        df_kpi,
        ['последнего звонка до первого статуса первой подтвержденной смены', 'последнего звонка'],
        exact_match='Дата последнего звонка до первого статуса первой подтвержденной смены'
    )
    col_kpi_client = find_column(df_kpi, ['клиент'])
    col_kpi_city = find_column(df_kpi, ['город'])
    col_kpi_city_first_shift = find_column(df_kpi, ['город первой смены', 'город первой'])

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

    rename_kpi = {
        col_kpi_phone: 'Телефон гигера',
        col_kpi_recruiter: 'Рекрутер',
        col_kpi_source: 'Источник ОМПП',
        col_kpi_date_first_shift: 'Дата первой смены',
        col_kpi_last_call: 'Дата последнего звонка',
    }
    if col_kpi_city is not None:
        rename_kpi[col_kpi_city] = 'Город'
    if col_kpi_city_first_shift is not None:
        rename_kpi[col_kpi_city_first_shift] = 'Город первой смены'
    if col_kpi_client is not None:
        rename_kpi[col_kpi_client] = 'Клиент'

    df_kpi = df_kpi.rename(columns=rename_kpi)
    df_kpi = df_kpi.loc[:, ~df_kpi.columns.duplicated()]

    df_kpi['Дата первой смены'] = pd.to_datetime(df_kpi['Дата первой смены'], errors='coerce')
    df_kpi['Дата последнего звонка'] = pd.to_datetime(df_kpi['Дата последнего звонка'], errors='coerce')

    df_kpi['Источник ОМПП'] = df_kpi['Источник ОМПП'].astype(str).str.strip()
    df_kpi = df_kpi[df_kpi['Источник ОМПП'].notna() & (df_kpi['Источник ОМПП'] != '')]

# ---- Обработка файла "отчет по обработке откликов" ----
if uploaded_file_responses is not None:
    try:
        df_responses = pd.read_excel(uploaded_file_responses, sheet_name="Отклики общая")
        df_responses.columns = df_responses.columns.str.strip()
    except Exception as e:
        st.error(f"Не удалось прочитать лист 'Отклики общая': {e}")
        df_responses = None

    try:
        df_diagram = pd.read_excel(uploaded_file_responses, sheet_name="Диаграмма", header=None)
    except Exception as e:
        st.error(f"Не удалось прочитать лист 'Диаграмма': {e}")
        df_diagram = None

# ---- Боковая панель с фильтрами ----
st.sidebar.header("Фильтры")

# Фильтр по источнику
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

# Фильтр по дате направления
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

# Фильтр по дате первой смены
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

# ---- Фильтр по дате откликов (с автоподстановкой) ----
if df_responses is not None:
    col_date_resp = find_column(df_responses, ['дата отклика', 'отклика'])
    if col_date_resp is not None:
        df_responses['Дата отклика'] = pd.to_datetime(df_responses[col_date_resp], errors='coerce')
        df_responses = df_responses[df_responses['Дата отклика'].notna()]
        if not df_responses.empty:
            min_date_resp = df_responses['Дата отклика'].min().date()
            max_date_resp = df_responses['Дата отклика'].max().date()
            
            # Определяем значения по умолчанию
            default_start = min_date_resp
            default_end = max_date_resp
            
            # Проверяем date_range_main
            if date_range_main is not None and isinstance(date_range_main, (list, tuple)) and len(date_range_main) == 2:
                d1, d2 = date_range_main
                if d1 is not None and d2 is not None:
                    # Приводим к date
                    if not isinstance(d1, date):
                        try:
                            d1 = pd.to_datetime(d1).date()
                        except:
                            d1 = None
                    if not isinstance(d2, date):
                        try:
                            d2 = pd.to_datetime(d2).date()
                        except:
                            d2 = None
                    if d1 is not None and d2 is not None:
                        default_start, default_end = d1, d2
            
            # Если не заданы или некорректны, пробуем взять из KPI
            if default_start == min_date_resp and default_end == max_date_resp:
                if date_range_kpi is not None and isinstance(date_range_kpi, (list, tuple)) and len(date_range_kpi) == 2:
                    d1, d2 = date_range_kpi
                    if d1 is not None and d2 is not None:
                        if not isinstance(d1, date):
                            try:
                                d1 = pd.to_datetime(d1).date()
                            except:
                                d1 = None
                        if not isinstance(d2, date):
                            try:
                                d2 = pd.to_datetime(d2).date()
                            except:
                                d2 = None
                        if d1 is not None and d2 is not None:
                            default_start, default_end = d1, d2
            
            # Убеждаемся, что start <= end
            if default_start > default_end:
                default_start, default_end = default_end, default_start
            
            # Ограничиваем min/max
            if default_start < min_date_resp:
                default_start = min_date_resp
            if default_end > max_date_resp:
                default_end = max_date_resp
            
            st.sidebar.subheader("Фильтр по дате отклика")
            date_range_resp = st.sidebar.date_input(
                "Диапазон дат отклика",
                value=(default_start, default_end),
                min_value=min_date_resp,
                max_value=max_date_resp,
                key="date_range_resp"
            )
        else:
            date_range_resp = None
            st.sidebar.warning("Нет данных с корректной датой отклика.")
    else:
        date_range_resp = None
        st.sidebar.warning("В листе 'Отклики общая' не найден столбец с датой отклика.")
else:
    date_range_resp = None

# ---- Фильтр по рекрутерам (общий) ----
default_recruiters = [
    'Балдин Александр',
    'Балдина Ксения',
    'Демьянова Алла',
    'Дорохина Галина',
    'Левшина Оксана',
    'Леонтьева Галина',
    'Москвитина Анна',
    'Роминян Максим',
    'Царева Анастасия',
    'Шайхутдинова Ильмира'
]

# Собираем всех рекрутеров из всех источников
all_recruiters = set()
if df_main is not None:
    all_recruiters.update(df_main['Рекрутер'].dropna().unique())
if df_kpi is not None:
    all_recruiters.update(df_kpi['Рекрутер'].dropna().unique())
if df_responses is not None:
    col_recr_resp = find_column(df_responses, ['рекрутер в crm', 'рекрутер в срм'])
    if col_recr_resp is not None:
        all_recruiters.update(df_responses[col_recr_resp].dropna().unique())
all_recruiters = sorted(all_recruiters)

# По умолчанию выбираем только тех, кто есть в списке default_recruiters и присутствует в данных
default_selected = [r for r in default_recruiters if r in all_recruiters]
# Если никого из списка нет, показываем всех (чтобы не было пустого экрана)
if not default_selected:
    default_selected = all_recruiters

selected_recruiters = st.sidebar.multiselect(
    "Рекрутеры",
    options=all_recruiters,
    default=default_selected,
    key="recruiter_filter_global"
)

# ---- Применение фильтров к основному отчету ----
if df_main is not None and selected_sources:
    df_main_filtered = df_main[df_main['Источник ОМПП'].isin(selected_sources)]

    if date_range_main and len(date_range_main) == 2:
        start_date, end_date = date_range_main
        df_main_filtered = df_main_filtered[
            (df_main_filtered['Дата направления'].dt.date >= start_date) &
            (df_main_filtered['Дата направления'].dt.date <= end_date)
        ]

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
    
    # Применяем фильтр по рекрутерам
    if selected_recruiters:
        df_main_filtered = df_main_filtered[df_main_filtered['Рекрутер'].isin(selected_recruiters)]
else:
    df_main_filtered = None

# ---- Применение фильтров к KPI ----
if df_kpi is not None and selected_sources:
    df_kpi_filtered = df_kpi[df_kpi['Источник ОМПП'].isin(selected_sources)]

    if date_range_kpi and len(date_range_kpi) == 2:
        start_date_kpi, end_date_kpi = date_range_kpi
        df_kpi_filtered = df_kpi_filtered[
            (df_kpi_filtered['Дата первой смены'].dt.date >= start_date_kpi) &
            (df_kpi_filtered['Дата первой смены'].dt.date <= end_date_kpi)
        ]

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
    
    # Применяем фильтр по рекрутерам
    if selected_recruiters:
        df_kpi_filtered = df_kpi_filtered[df_kpi_filtered['Рекрутер'].isin(selected_recruiters)]
else:
    df_kpi_filtered = None

# ---- Применение фильтров к откликам ----
df_responses_filtered = None
if df_responses is not None:
    df_responses_filtered = df_responses.copy()
    # Фильтр по дате отклика
    if date_range_resp and len(date_range_resp) == 2 and 'Дата отклика' in df_responses_filtered.columns:
        start_date_resp, end_date_resp = date_range_resp
        df_responses_filtered = df_responses_filtered[
            (df_responses_filtered['Дата отклика'].dt.date >= start_date_resp) &
            (df_responses_filtered['Дата отклика'].dt.date <= end_date_resp)
        ]
    # Фильтр по рекрутерам
    if selected_recruiters:
        col_recr_resp = find_column(df_responses_filtered, ['рекрутер в crm', 'рекрутер в срм'])
        if col_recr_resp is not None:
            df_responses_filtered = df_responses_filtered[df_responses_filtered[col_recr_resp].isin(selected_recruiters)]

# ---- Обработка откликов: таблица (будет выведена в конце) ----
merged_resp = None
df_resp_for_city = None  # для статистики по городам
if df_responses_filtered is not None:
    col_phone_resp = find_column(df_responses_filtered, ['телефон соискателя'])
    col_recruiter_resp = find_column(df_responses_filtered, ['рекрутер в crm', 'рекрутер в срм'])
    col_status_resp = find_column(df_responses_filtered, ['статус рекрутера'])
    col_first_shift_resp = find_column(df_responses_filtered, ['первая смена после отклика'])
    col_city_vacancy = find_column(df_responses_filtered, ['город вакансии'])

    if col_phone_resp is None or col_recruiter_resp is None:
        st.error("❌ В листе 'Отклики общая' не найдены столбцы 'Телефон соискателя' и/или 'Рекрутер в CRM'")
    else:
        rename_resp = {
            col_phone_resp: 'Телефон',
            col_recruiter_resp: 'Рекрутер'
        }
        if col_status_resp is not None:
            rename_resp[col_status_resp] = 'Статус'
        if col_first_shift_resp is not None:
            rename_resp[col_first_shift_resp] = 'Первая смена'
        if col_city_vacancy is not None:
            rename_resp[col_city_vacancy] = 'Город вакансии'

        df_resp = df_responses_filtered.rename(columns=rename_resp)
        df_resp = df_resp.loc[:, ~df_resp.columns.duplicated()]

        # Сохраняем для города
        df_resp_for_city = df_resp.copy()

        # Группировка для таблицы обработки откликов
        responses_count = df_resp.groupby('Рекрутер').size().reset_index(name='Кол-во откликов')

        if 'Статус' in df_resp.columns:
            reg_statuses = ['Регистрация', 'Приглашен на смену', 'Смена забронирована']
            df_reg = df_resp[df_resp['Статус'].isin(reg_statuses)]
            reg_count = df_reg.groupby('Рекрутер').size().reset_index(name='Кол-во регистраций')
        else:
            reg_count = pd.DataFrame(columns=['Рекрутер', 'Кол-во регистраций'])

        if 'Статус' in df_resp.columns:
            invited_statuses = ['Приглашен на смену', 'Смена забронирована']
            df_inv = df_resp[df_resp['Статус'].isin(invited_statuses)]
            invited_count = df_inv.groupby('Рекрутер').size().reset_index(name='Кол-во направленных из откликов')
        else:
            invited_count = pd.DataFrame(columns=['Рекрутер', 'Кол-во направленных из откликов'])

        if 'Первая смена' in df_resp.columns:
            df_worked_resp = df_resp[df_resp['Первая смена'].notna()]
            worked_count = df_worked_resp.groupby('Рекрутер').size().reset_index(name='Вышедшие из откликов')
        else:
            worked_count = pd.DataFrame(columns=['Рекрутер', 'Вышедшие из откликов'])

        merged_resp = responses_count.merge(reg_count, on='Рекрутер', how='left').fillna(0)
        merged_resp = merged_resp.merge(invited_count, on='Рекрутер', how='left').fillna(0)
        merged_resp = merged_resp.merge(worked_count, on='Рекрутер', how='left').fillna(0)

        for col in ['Кол-во откликов', 'Кол-во регистраций', 'Кол-во направленных из откликов', 'Вышедшие из откликов']:
            merged_resp[col] = merged_resp[col].astype(int)

        # Конверсии
        merged_resp['Конв. отклик->регистр, %'] = (merged_resp['Кол-во регистраций'] / merged_resp['Кол-во откликов'] * 100).round(1)
        merged_resp['Конв. регистр->направл, %'] = (merged_resp['Кол-во направленных из откликов'] / merged_resp['Кол-во регистраций'] * 100).round(1).fillna(0)
        merged_resp['Конв. направл->вышед, %'] = (merged_resp['Вышедшие из откликов'] / merged_resp['Кол-во направленных из откликов'] * 100).round(1).fillna(0)
        merged_resp['Конв. отклик->вышед, %'] = (merged_resp['Вышедшие из откликов'] / merged_resp['Кол-во откликов'] * 100).round(1)

        for col in ['Конв. отклик->регистр, %', 'Конв. регистр->направл, %', 'Конв. направл->вышед, %', 'Конв. отклик->вышед, %']:
            merged_resp[col] = merged_resp[col].fillna(0).replace([float('inf'), -float('inf')], 0)
            merged_resp[col] = merged_resp[col].astype(str) + '%'

        merged_resp = merged_resp.sort_values('Кол-во откликов', ascending=False)

# ---- 1. Объединённая таблица рекрутеров (из обоих отчетов) ----
recruiter_data = {}

if df_main_filtered is not None:
    main_counts = df_main_filtered.groupby('Рекрутер')['Телефон'].nunique().reset_index()
    main_counts.columns = ['Рекрутер', 'Кол-во направленных']

    if 'Дата первой подтвержденной смены за всю жизнь' in df_main_filtered.columns:
        df_with_shift = df_main_filtered[df_main_filtered['Дата первой подтвержденной смены за всю жизнь'].notna()]
        worked_main = df_with_shift.groupby('Рекрутер')['Телефон'].nunique().reset_index()
        worked_main.columns = ['Рекрутер', 'Вышло из приглашенных']
        main_counts = main_counts.merge(worked_main, on='Рекрутер', how='left').fillna(0)
        main_counts['Вышло из приглашенных'] = main_counts['Вышло из приглашенных'].astype(int)
    else:
        main_counts['Вышло из приглашенных'] = 0

    for _, row in main_counts.iterrows():
        recruiter = row['Рекрутер']
        if recruiter not in recruiter_data:
            recruiter_data[recruiter] = {}
        recruiter_data[recruiter]['Кол-во направленных'] = row['Кол-во направленных']
        recruiter_data[recruiter]['Вышло из приглашенных'] = row['Вышло из приглашенных']

if df_kpi_filtered is not None:
    kpi_counts = df_kpi_filtered.groupby('Рекрутер')['Телефон гигера'].nunique().reset_index()
    kpi_counts.columns = ['Рекрутер', 'Вышедшие (с дошедшими)']
    for _, row in kpi_counts.iterrows():
        recruiter = row['Рекрутер']
        if recruiter not in recruiter_data:
            recruiter_data[recruiter] = {}
        recruiter_data[recruiter]['Вышедшие (с дошедшими)'] = row['Вышедшие (с дошедшими)']

# Добавляем данные из откликов (кол-во откликов и конверсия из откликов в вышедших)
if merged_resp is not None:
    for _, row in merged_resp.iterrows():
        recruiter = row['Рекрутер']
        if recruiter not in recruiter_data:
            recruiter_data[recruiter] = {}
        recruiter_data[recruiter]['Кол-во откликов'] = row['Кол-во откликов']
        recruiter_data[recruiter]['Конверсия из откликов в вышедших, %'] = row['Конв. отклик->вышед, %']

if recruiter_data:
    df_recruiters = pd.DataFrame.from_dict(recruiter_data, orient='index').reset_index()
    df_recruiters.rename(columns={'index': 'Рекрутер'}, inplace=True)

    numeric_cols = ['Кол-во направленных', 'Вышло из приглашенных', 'Вышедшие (с дошедшими)', 'Кол-во откликов']
    for col in numeric_cols:
        if col not in df_recruiters.columns:
            df_recruiters[col] = 0
        df_recruiters[col] = df_recruiters[col].fillna(0).astype(int)

    # Конверсии (из приглашенных)
    df_recruiters['Конверсия из пригл. в вышедших из приглашенных, %'] = (
        df_recruiters['Вышло из приглашенных'] / df_recruiters['Кол-во направленных'] * 100
    ).round(1).fillna(0).astype(str) + '%'

    df_recruiters['Конверсия из приглашенных в вышедших с дошедшими, %'] = (
        df_recruiters['Вышедшие (с дошедшими)'] / df_recruiters['Кол-во направленных'] * 100
    ).round(1).fillna(0).astype(str) + '%'

    if 'Конверсия из откликов в вышедших, %' not in df_recruiters.columns:
        df_recruiters['Конверсия из откликов в вышедших, %'] = '0%'

    # Сортировка
    sort_col = 'Кол-во направленных' if df_recruiters['Кол-во направленных'].sum() > 0 else 'Вышедшие (с дошедшими)'
    df_recruiters = df_recruiters.sort_values(sort_col, ascending=False)

    display_cols = ['Рекрутер']
    col_config = {}
    if 'Кол-во откликов' in df_recruiters.columns and df_recruiters['Кол-во откликов'].sum() > 0:
        display_cols.append('Кол-во откликов')
        col_config['Кол-во откликов'] = st.column_config.NumberColumn("Кол-во откликов", format="%d", width="auto")
    if 'Кол-во направленных' in df_recruiters.columns and df_recruiters['Кол-во направленных'].sum() > 0:
        display_cols.append('Кол-во направленных')
        col_config['Кол-во направленных'] = st.column_config.NumberColumn("Кол-во направленных", format="%d", width="auto")
    if 'Вышло из приглашенных' in df_recruiters.columns and df_recruiters['Вышло из приглашенных'].sum() > 0:
        display_cols.append('Вышло из приглашенных')
        col_config['Вышло из приглашенных'] = st.column_config.NumberColumn("Вышло из приглашенных", format="%d", width="auto")
    if 'Конверсия из пригл. в вышедших из приглашенных, %' in df_recruiters.columns:
        if df_recruiters['Кол-во направленных'].sum() > 0 and df_recruiters['Вышло из приглашенных'].sum() > 0:
            display_cols.append('Конверсия из пригл. в вышедших из приглашенных, %')
            col_config['Конверсия из пригл. в вышедших из приглашенных, %'] = st.column_config.TextColumn("Конверсия из пригл. в вышедших из приглашенных, %", width="auto")
    if 'Вышедшие (с дошедшими)' in df_recruiters.columns and df_recruiters['Вышедшие (с дошедшими)'].sum() > 0:
        display_cols.append('Вышедшие (с дошедшими)')
        col_config['Вышедшие (с дошедшими)'] = st.column_config.NumberColumn("Вышедшие (с дошедшими)", format="%d", width="auto")
    if 'Конверсия из приглашенных в вышедших с дошедшими, %' in df_recruiters.columns:
        if df_recruiters['Кол-во направленных'].sum() > 0 and df_recruiters['Вышедшие (с дошедшими)'].sum() > 0:
            display_cols.append('Конверсия из приглашенных в вышедших с дошедшими, %')
            col_config['Конверсия из приглашенных в вышедших с дошедшими, %'] = st.column_config.TextColumn("Конверсия из приглашенных в вышедших с дошедшими, %", width="auto")
    if 'Конверсия из откликов в вышедших, %' in df_recruiters.columns:
        if df_recruiters['Кол-во откликов'].sum() > 0 and df_recruiters['Вышедшие (с дошедшими)'].sum() > 0:
            display_cols.append('Конверсия из откликов в вышедших, %')
            col_config['Конверсия из откликов в вышедших, %'] = st.column_config.TextColumn("Конверсия из откликов в вышедших, %", width="auto")

    st.subheader("📋 Количество направленных кандидатов по рекрутерам")
    st.dataframe(
        df_recruiters[display_cols],
        use_container_width=True,
        column_config=col_config
    )
else:
    st.info("Нет данных для отображения. Загрузите хотя бы один отчет.")

# ---- 2. График по источникам (направленные) ----
if df_main_filtered is not None:
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

        st.subheader("📋 Приглашенные по источникам и рекрутерам")
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
                "Рекрутер": st.column_config.TextColumn("Рекрутер", width="auto"),
                "Источник ОМПП": st.column_config.TextColumn("Источник ОМПП", width="auto"),
                "Кол-во": st.column_config.NumberColumn("Кол-во", format="%d", width="auto"),
                "% от рекрутера": st.column_config.TextColumn("% от рекрутера", width="auto"),
                "% от всех": st.column_config.TextColumn("% от всех", width="auto"),
            }
        )

# ---- 3. Приглашенные по проектам (график) ----
if df_main_filtered is not None and 'Желаемые проекты (Группа)' in df_main_filtered.columns:
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

# ---- 4. Вышедшие по проектам (объединённый блок) ----
st.subheader("✅ Вышедшие по проектам")

# Данные по проектам из основного отчёта
if df_main_filtered is not None:
    # Определяем вышедших из приглашенных (из основного отчёта)
    if 'Статус координатора' in df_main_filtered.columns:
        df_worked_main = df_main_filtered[df_main_filtered['Статус координатора'] == 'went_work']
    else:
        df_worked_main = df_main_filtered[df_main_filtered['Статус лида'] == 'worked']
    
    # Если нет данных о вышедших, создаём пустой DataFrame
    if df_worked_main.empty or 'Проект первой подтвержденной смены' not in df_worked_main.columns:
        main_project_data = pd.DataFrame(columns=['Проект', 'Кол-во приглашенных', 'Кол-во вышедших (из приглашенных)'])
    else:
        # Приглашенные по проектам (все кандидаты) – нормализуем названия
        df_projects_all = df_main_filtered.copy()
        if 'Желаемые проекты (Клиент)' in df_projects_all.columns:
            df_projects_all['Проект'] = df_projects_all.apply(
                lambda row: normalize_project(row['Желаемые проекты (Клиент)']) if row['Желаемые проекты (Группа)'] == 'Без группы' else normalize_project(row['Желаемые проекты (Группа)']),
                axis=1
            )
        else:
            df_projects_all['Проект'] = df_projects_all['Желаемые проекты (Группа)'].apply(normalize_project)
        
        # Дополнительно удаляем лишние пробелы
        df_projects_all['Проект'] = df_projects_all['Проект'].astype(str).str.strip()
        df_projects_all = df_projects_all[df_projects_all['Проект'].notna() & (df_projects_all['Проект'] != '')]
        invited_counts = df_projects_all.groupby('Проект')['Телефон'].nunique().reset_index()
        invited_counts.columns = ['Проект', 'Кол-во приглашенных']
        
        # Вышедшие из приглашенных – также нормализуем
        df_worked_main['Проект'] = df_worked_main['Проект первой подтвержденной смены'].apply(normalize_project)
        df_worked_main['Проект'] = df_worked_main['Проект'].astype(str).str.strip()
        worked_main_counts = df_worked_main.groupby('Проект')['Телефон'].nunique().reset_index()
        worked_main_counts.columns = ['Проект', 'Кол-во вышедших (из приглашенных)']
        
        main_project_data = pd.merge(invited_counts, worked_main_counts, on='Проект', how='outer').fillna(0)
        main_project_data['Кол-во приглашенных'] = main_project_data['Кол-во приглашенных'].astype(int)
        main_project_data['Кол-во вышедших (из приглашенных)'] = main_project_data['Кол-во вышедших (из приглашенных)'].astype(int)
else:
    main_project_data = pd.DataFrame(columns=['Проект', 'Кол-во приглашенных', 'Кол-во вышедших (из приглашенных)'])

# Данные из KPI (вышедшие с дошедшими) – нормализуем названия
if df_kpi_filtered is not None and 'Клиент' in df_kpi_filtered.columns:
    kpi_project_counts = df_kpi_filtered.groupby('Клиент')['Телефон гигера'].nunique().reset_index()
    kpi_project_counts.columns = ['Проект', 'Кол-во вышедших (с дошедшими)']
    kpi_project_counts['Проект'] = kpi_project_counts['Проект'].apply(normalize_project)
    kpi_project_counts['Проект'] = kpi_project_counts['Проект'].astype(str).str.strip()
else:
    kpi_project_counts = pd.DataFrame(columns=['Проект', 'Кол-во вышедших (с дошедшими)'])

# Объединяем
merged_projects = pd.merge(main_project_data, kpi_project_counts, on='Проект', how='outer').fillna(0)
for col in ['Кол-во приглашенных', 'Кол-во вышедших (из приглашенных)', 'Кол-во вышедших (с дошедшими)']:
    if col in merged_projects.columns:
        merged_projects[col] = merged_projects[col].astype(int)
    else:
        merged_projects[col] = 0

# Конверсии
merged_projects['Конв. из приглашенных в вышедших из пригл., %'] = (
    merged_projects['Кол-во вышедших (из приглашенных)'] / merged_projects['Кол-во приглашенных'] * 100
).round(1).fillna(0).astype(str) + '%'

merged_projects['Конв. из приглашенных в вышедших с дошедшими, %'] = (
    merged_projects['Кол-во вышедших (с дошедшими)'] / merged_projects['Кол-во приглашенных'] * 100
).round(1).fillna(0).astype(str) + '%'

# Сортировка по количеству приглашенных
merged_projects = merged_projects.sort_values('Кол-во приглашенных', ascending=False)

# Отображаем таблицу
if not merged_projects.empty and merged_projects['Кол-во приглашенных'].sum() > 0:
    st.dataframe(
        merged_projects,
        use_container_width=True,
        column_config={
            "Проект": st.column_config.TextColumn("Проект", width="auto"),
            "Кол-во приглашенных": st.column_config.NumberColumn("Кол-во приглашенных", format="%d", width="auto"),
            "Кол-во вышедших (из приглашенных)": st.column_config.NumberColumn("Кол-во вышедших (из приглашенных)", format="%d", width="auto"),
            "Конв. из приглашенных в вышедших из пригл., %": st.column_config.TextColumn("Конв. из приглашенных в вышедших из пригл., %", width="auto"),
            "Кол-во вышедших (с дошедшими)": st.column_config.NumberColumn("Кол-во вышедших (с дошедшими)", format="%d", width="auto"),
            "Конв. из приглашенных в вышедших с дошедшими, %": st.column_config.TextColumn("Конв. из приглашенных в вышедших с дошедшими, %", width="auto"),
        }
    )
else:
    st.info("Нет данных по проектам для вышедших кандидатов.")

# ---- 5. Объединённый блок: Статистика по городам ----
st.subheader("🏙️ Статистика по городам")

# Подготовка данных по городам
city_data = {}

# 1. Кол-во откликов из отчета по обработке откликов
if df_resp_for_city is not None and 'Город вакансии' in df_resp_for_city.columns:
    city_resp = df_resp_for_city[df_resp_for_city['Город вакансии'].notna() & (df_resp_for_city['Город вакансии'].astype(str).str.strip() != '')].copy()
    city_resp['Город'] = city_resp['Город вакансии'].astype(str).str.strip()
    city_resp_counts = city_resp.groupby('Город').size().reset_index(name='Кол-во откликов')
    for _, row in city_resp_counts.iterrows():
        city = row['Город']
        if city not in city_data:
            city_data[city] = {}
        city_data[city]['Кол-во откликов'] = row['Кол-во откликов']

# 2. Приглашенные (из основного отчёта)
if df_main_filtered is not None and 'Город' in df_main_filtered.columns:
    city_invited = df_main_filtered[
        df_main_filtered['Город'].notna() & 
        (df_main_filtered['Город'].astype(str).str.strip() != '')
    ].copy()
    city_invited['Город'] = city_invited['Город'].astype(str).str.strip()
    invited_city = city_invited.groupby('Город')['Телефон'].nunique().reset_index()
    invited_city.columns = ['Город', 'Кол-во приглашенных']
    for _, row in invited_city.iterrows():
        city = row['Город']
        if city not in city_data:
            city_data[city] = {}
        city_data[city]['Кол-во приглашенных'] = row['Кол-во приглашенных']

# 3. Вышедшие (из KPI)
if df_kpi_filtered is not None and 'Город первой смены' in df_kpi_filtered.columns:
    df_kpi_city = df_kpi_filtered[
        df_kpi_filtered['Рекрутер'].notna() & 
        (df_kpi_filtered['Рекрутер'].astype(str).str.strip() != '') &
        df_kpi_filtered['Город первой смены'].notna() & 
        (df_kpi_filtered['Город первой смены'].astype(str).str.strip() != '')
    ].copy()
    df_kpi_city['Город'] = df_kpi_city['Город первой смены'].astype(str).str.strip()
    worked_city = df_kpi_city.groupby('Город')['Телефон гигера'].nunique().reset_index()
    worked_city.columns = ['Город', 'Кол-во вышедших']
    for _, row in worked_city.iterrows():
        city = row['Город']
        if city not in city_data:
            city_data[city] = {}
        city_data[city]['Кол-во вышедших'] = row['Кол-во вышедших']

if city_data:
    df_city = pd.DataFrame.from_dict(city_data, orient='index').reset_index()
    df_city.rename(columns={'index': 'Город'}, inplace=True)
    
    # Заполняем пропуски
    for col in ['Кол-во откликов', 'Кол-во приглашенных', 'Кол-во вышедших']:
        if col not in df_city.columns:
            df_city[col] = 0
        df_city[col] = df_city[col].fillna(0).astype(int)
    
    # Доля приглашенных
    total_invited = df_city['Кол-во приглашенных'].sum()
    df_city['Доля приглашенных'] = (df_city['Кол-во приглашенных'] / total_invited * 100).round(1).astype(str) + '%' if total_invited > 0 else '0%'
    
    # Конверсия из направленных в вышедших
    df_city['Конверсия из направленных в вышедших, %'] = (
        df_city['Кол-во вышедших'] / df_city['Кол-во приглашенных'] * 100
    ).round(1).fillna(0).astype(str) + '%'
    
    df_city = df_city.sort_values('Кол-во приглашенных', ascending=False)
    
    st.dataframe(
        df_city,
        use_container_width=True,
        column_config={
            "Город": st.column_config.TextColumn("Город", width="auto"),
            "Кол-во откликов": st.column_config.NumberColumn("Кол-во откликов", format="%d", width="auto"),
            "Кол-во приглашенных": st.column_config.NumberColumn("Кол-во приглашенных", format="%d", width="auto"),
            "Доля приглашенных": st.column_config.TextColumn("Доля приглашенных", width="auto"),
            "Кол-во вышедших": st.column_config.NumberColumn("Кол-во вышедших", format="%d", width="auto"),
            "Конверсия из направленных в вышедших, %": st.column_config.TextColumn("Конверсия из направленных в вышедших, %", width="auto"),
        }
    )
else:
    st.info("Нет данных по городам.")

# ---- 6. График по источникам (вышедшие из KPI) ----
if df_kpi_filtered is not None:
    st.subheader("📊 Кол-во вышедших кандидатов по источникам (с дошедшими)")
    available_sources_kpi = sorted(df_kpi_filtered['Источник ОМПП'].unique())
    if not available_sources_kpi:
        st.warning("Нет доступных источников для отображения.")
    else:
        selected_source_for_chart_kpi = st.selectbox("Выберите источник для отображения (вышедшие):", options=available_sources_kpi, key="source_chart_kpi")

        df_chart_kpi = df_kpi_filtered[df_kpi_filtered['Источник ОМПП'] == selected_source_for_chart_kpi]
        chart_data_kpi = df_chart_kpi.groupby('Рекрутер')['Телефон гигера'].nunique().reset_index()
        chart_data_kpi.columns = ['Рекрутер', 'Кол-во']
        chart_data_kpi = chart_data_kpi.sort_values('Кол-во', ascending=False)

        if chart_data_kpi.empty:
            st.info("Нет данных для выбранного источника.")
        else:
            fig_kpi = px.bar(
                chart_data_kpi,
                x='Кол-во',
                y='Рекрутер',
                orientation='h',
                title=f"Источник: {selected_source_for_chart_kpi}",
                labels={'Кол-во': 'Кол-во вышедших кандидатов', 'Рекрутер': ''},
                text='Кол-во',
                color='Кол-во',
                color_continuous_scale='Greens',
                height=500
            )
            fig_kpi.update_traces(textposition='outside')
            fig_kpi.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
            st.plotly_chart(fig_kpi, use_container_width=True)

        st.subheader("📋 Вышедшие по источникам и рекрутерам")
        detail_kpi = df_kpi_filtered.groupby(['Рекрутер', 'Источник ОМПП'])['Телефон гигера'].nunique().reset_index()
        detail_kpi.columns = ['Рекрутер', 'Источник ОМПП', 'Кол-во']

        recruiter_total_kpi = detail_kpi.groupby('Рекрутер')['Кол-во'].sum().reset_index()
        recruiter_total_kpi.columns = ['Рекрутер', 'Всего_рекрутер']

        grand_total_kpi = detail_kpi['Кол-во'].sum()

        detail_kpi = detail_kpi.merge(recruiter_total_kpi, on='Рекрутер', how='left')
        detail_kpi['% от рекрутера'] = (detail_kpi['Кол-во'] / detail_kpi['Всего_рекрутер'] * 100).round(1)
        detail_kpi['% от всех'] = (detail_kpi['Кол-во'] / grand_total_kpi * 100).round(1)

        detail_kpi['% от рекрутера'] = detail_kpi['% от рекрутера'].astype(str) + '%'
        detail_kpi['% от всех'] = detail_kpi['% от всех'].astype(str) + '%'

        detail_kpi = detail_kpi.sort_values(['Рекрутер', 'Кол-во'], ascending=[True, False])

        st.dataframe(
            detail_kpi[['Рекрутер', 'Источник ОМПП', 'Кол-во', '% от рекрутера', '% от всех']],
            use_container_width=True,
            column_config={
                "Рекрутер": st.column_config.TextColumn("Рекрутер", width="auto"),
                "Источник ОМПП": st.column_config.TextColumn("Источник ОМПП", width="auto"),
                "Кол-во": st.column_config.NumberColumn("Кол-во", format="%d", width="auto"),
                "% от рекрутера": st.column_config.TextColumn("% от рекрутера", width="auto"),
                "% от всех": st.column_config.TextColumn("% от всех", width="auto"),
            }
        )

# ---- 7. Блок: Обработка откликов (таблица) ----
if merged_resp is not None:
    st.subheader("📋 Обработка откликов")
    st.dataframe(
        merged_resp,
        use_container_width=True,
        column_config={
            "Рекрутер": st.column_config.TextColumn("Рекрутер", width="auto"),
            "Кол-во откликов": st.column_config.NumberColumn("Кол-во откликов", format="%d", width="auto"),
            "Кол-во регистраций": st.column_config.NumberColumn("Кол-во регистраций", format="%d", width="auto"),
            "Кол-во направленных из откликов": st.column_config.NumberColumn("Кол-во направленных из откликов", format="%d", width="auto"),
            "Вышедшие из откликов": st.column_config.NumberColumn("Вышедшие из откликов", format="%d", width="auto"),
            "Конв. отклик->регистр, %": st.column_config.TextColumn("Конв. отклик->регистр, %", width="auto"),
            "Конв. регистр->направл, %": st.column_config.TextColumn("Конв. регистр->направл, %", width="auto"),
            "Конв. направл->вышед, %": st.column_config.TextColumn("Конв. направл->вышед, %", width="auto"),
            "Конв. отклик->вышед, %": st.column_config.TextColumn("Конв. отклик->вышед, %", width="auto"),
        }
    )

# ---- 8. Диаграмма "Время обработки откликов в рабочее время" (горизонтальная) ----
if df_diagram is not None:
    st.subheader("📊 Время обработки откликов в рабочее время")
    try:
        df_diagram_reset = df_diagram.reset_index(drop=True)
        row_15min = None
        row_less_hour = None
        date_row = None

        # Ищем строки с данными по всей строке (не только в первом столбце)
        for idx, row in df_diagram_reset.iterrows():
            # Преобразуем все ячейки строки в строки и объединяем
            row_str = ' '.join([str(cell).strip().lower() for cell in row if pd.notna(cell)])
            if 'в течение 15 минут' in row_str:
                row_15min = idx
            if 'менее часа' in row_str:
                row_less_hour = idx
            # Ищем строку с датами: проверяем все ячейки
            if date_row is None:
                for col in range(len(row)):
                    cell_val = row.iloc[col]
                    if pd.notna(cell_val):
                        cell_str = str(cell_val).strip()
                        if re.match(r'^\d{2}\.\w{3}$', cell_str) or re.match(r'^\d{4}-\d{2}-\d{2}', cell_str):
                            date_row = idx
                            break

        if date_row is None:
            date_row = 0
        if row_15min is None:
            row_15min = 1
        if row_less_hour is None:
            row_less_hour = 2

        # Собираем даты и значения
        dates = []
        for col in range(1, len(df_diagram_reset.columns)):
            val = df_diagram_reset.iloc[date_row, col]
            if pd.notna(val):
                date_val = parse_diagram_date(str(val))
                if pd.notna(date_val):
                    dates.append((col, date_val))

        data_15min = []
        data_less_hour = []
        for col, dt in dates:
            val_15 = df_diagram_reset.iloc[row_15min, col]
            val_less = df_diagram_reset.iloc[row_less_hour, col]
            try:
                v15 = float(val_15) if pd.notna(val_15) else None
            except:
                v15 = None
            try:
                vless = float(val_less) if pd.notna(val_less) else None
            except:
                vless = None
            if v15 is not None and vless is not None:
                data_15min.append((dt, v15))
                data_less_hour.append((dt, vless))

        if data_15min and data_less_hour:
            df_diag_parsed = pd.DataFrame({
                'Дата': [d for d, _ in data_15min],
                'В течение 15 минут': [v for _, v in data_15min],
                'Менее часа': [v for _, v in data_less_hour]
            })
            df_diag_parsed['Месяц'] = df_diag_parsed['Дата'].dt.to_period('M').astype(str)
            monthly_avg = df_diag_parsed.groupby('Месяц')[['В течение 15 минут', 'Менее часа']].mean().reset_index()
            monthly_avg['В течение 15 минут'] = monthly_avg['В течение 15 минут'].round(1)
            monthly_avg['Менее часа'] = monthly_avg['Менее часа'].round(1)

            # Горизонтальная столбчатая диаграмма
            fig_diag = px.bar(
                monthly_avg,
                x=['В течение 15 минут', 'Менее часа'],
                y='Месяц',
                orientation='h',
                barmode='group',
                title="Среднее время обработки откликов в рабочее время (по месяцам)",
                labels={'value': 'Средний %', 'variable': 'Метрика'},
                color_discrete_map={'В течение 15 минут': '#1f77b4', 'Менее часа': '#ff7f0e'}
            )
            fig_diag.update_layout(xaxis_title="Средний процент", yaxis_title="Месяц", legend_title="Метрика")
            st.plotly_chart(fig_diag, use_container_width=True)

            st.dataframe(monthly_avg, use_container_width=True)
        else:
            st.warning("Не удалось извлечь данные из листа 'Диаграмма'.")
    except Exception as e:
        st.error(f"Ошибка при обработке листа 'Диаграмма': {e}")
