import streamlit as st
import pandas as pd
import plotly.express as px
import re
from datetime import datetime, date
import io

st.set_page_config(page_title="ОМПП Дашборд", layout="wide")
st.title("📊 Дашборд ОМПП")

# ---- Словарь алиасов для нормализации проектов ----
PROJECT_ALIASES = {
    'пятёрочка': 'Пятёрочка',
    'пятерочка': 'Пятёрочка',
    'пятёрочка агентская': 'Пятёрочка',
    'магнит': 'Магнит',
    'магнит косметик': 'Магнит',
    'магнит сборка': 'Магнит',
    'магнит ': 'Магнит',
    'гулливер': 'Гулливер',
    'ооо "таймбук"': 'Гулливер',
    'бубль гум': 'Бубль-Гум',
    'бубль-гум': 'Бубль-Гум',
    'бубльгум': 'Бубль-Гум',
    'бубль-гум ': 'Бубль-Гум',
    'бубль гум ': 'Бубль-Гум',
    'бубльгум ': 'Бубль-Гум',
    'СПАР ': 'Спар',
}

def normalize_project(project_name):
    if not isinstance(project_name, str):
        return project_name
    cleaned = project_name.strip().lower()
    if cleaned in PROJECT_ALIASES:
        return PROJECT_ALIASES[cleaned]
    return project_name.strip()

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
    if pd.isna(date_str):
        return pd.NaT
    date_str = str(date_str).strip().lower()
    dt = pd.to_datetime(date_str, errors='coerce')
    if pd.notna(dt):
        return dt
    month_map = {
        'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4,
        'май': 5, 'июн': 6, 'июл': 7, 'авг': 8,
        'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12
    }
    match = re.match(r'(\d{1,2})\.(\D+)', date_str)
    if match:
        day = int(match.group(1))
        month_name = match.group(2).strip()[:3]
        if month_name in month_map:
            return pd.Timestamp(year=2026, month=month_map[month_name], day=day)
    return pd.NaT

# ---- Загрузка основного файла (направления) ----
uploaded_file = st.file_uploader("Загрузите Excel файл 'отчет по дате направления'", type=["xlsx"], key="main")

# ---- Загрузка файла KPI (вышедшие) ----
uploaded_file_kpi = st.file_uploader("Загрузите Excel файл KPI (отчет по вышедшим)", type=["xlsx"], key="kpi")

# ---- Загрузка файла "отчет по обработке откликов" ----
uploaded_file_responses = st.file_uploader("Загрузите Excel файл 'отчет по обработке откликов'", type=["xlsx"], key="responses")

# ---- Загрузка файла "Выгрузка истории звонков Mango-Office" (CSV) ----
uploaded_file_calls = st.file_uploader("Загрузите CSV файл 'Выгрузка истории звонков Mango-Office'", type=["csv"], key="calls")

# ---- Инициализация DataFrame ----
df_main = None
df_kpi = None
df_responses = None
df_diagram = None
df_calls = None

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

# ---- Улучшенная обработка CSV-файла звонков (если загружен) ----
if uploaded_file_calls is not None:
    try:
        raw_bytes = uploaded_file_calls.read()
        encodings = ['utf-8-sig', 'cp1251', 'latin1', 'utf-8']
        text = None
        for enc in encodings:
            try:
                text = raw_bytes.decode(enc)
                break
            except UnicodeDecodeError:
                continue
        if text is None:
            st.error("Не удалось декодировать файл ни в одной из кодировок.")
            st.stop()
        
        lines = text.splitlines()
        if not lines:
            st.error("Файл пуст.")
            st.stop()
        
        sep_counts = {';': 0, ',': 0, '\t': 0}
        for line in lines[:20]:
            in_quotes = False
            for ch in line:
                if ch == '"':
                    in_quotes = not in_quotes
                if not in_quotes and ch in sep_counts:
                    sep_counts[ch] += 1
        best_sep = max(sep_counts, key=sep_counts.get)
        if sep_counts[best_sep] == 0:
            best_sep = ';'
        
        header_keywords = ['кто звонил', 'звонивший', 'тип вызова', 'длительность']
        header_row_idx = None
        for i, line in enumerate(lines):
            line_low = line.lower()
            found = sum(1 for kw in header_keywords if kw in line_low)
            if found >= 2:
                header_row_idx = i
                break
        
        if header_row_idx is not None:
            data_to_read = '\n'.join(lines[header_row_idx:])
            df_calls = pd.read_csv(io.StringIO(data_to_read), sep=best_sep, quotechar='"', encoding='utf-8', header=0)
        else:
            df_calls = pd.read_csv(io.StringIO(text), sep=best_sep, quotechar='"', encoding='utf-8', header=1)
        
        if df_calls is not None:
            df_calls.columns = df_calls.columns.str.strip()
        else:
            st.error("Не удалось прочитать CSV файл звонков.")
            st.stop()
            
    except Exception as e:
        st.error(f"Ошибка при обработке CSV-файла: {e}")
        st.stop()

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
            
            default_start = min_date_resp
            default_end = max_date_resp
            
            if date_range_main is not None and isinstance(date_range_main, (list, tuple)) and len(date_range_main) == 2:
                d1, d2 = date_range_main
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
            
            if default_start > default_end:
                default_start, default_end = default_end, default_start
            
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

all_recruiters = set()
if df_main is not None:
    all_recruiters.update(df_main['Рекрутер'].dropna().unique())
if df_kpi is not None:
    all_recruiters.update(df_kpi['Рекрутер'].dropna().unique())
if df_responses is not None:
    col_recr_resp = find_column(df_responses, ['рекрутер в crm', 'рекрутер в срм'])
    if col_recr_resp is not None:
        all_recruiters.update(df_responses[col_recr_resp].dropna().unique())
if df_calls is not None:
    col_who = find_column(df_calls, ['кто звонил', 'звонивший', 'сотрудник', 'абонент', 'рекрутер'])
    if col_who is not None:
        all_recruiters.update(df_calls[col_who].astype(str).str.replace('"', '').str.strip().dropna().unique())

all_recruiters = sorted(all_recruiters)

default_selected = [
    r for r in default_recruiters
    if r in all_recruiters
]

current_recruiters = tuple(all_recruiters)

if (
    "last_recruiters" not in st.session_state
    or st.session_state["last_recruiters"] != current_recruiters
):
    st.session_state["recruiter_filter_global"] = default_selected
    st.session_state["last_recruiters"] = current_recruiters

selected_recruiters = st.sidebar.multiselect(
    "Рекрутеры",
    options=all_recruiters,
    key="recruiter_filter_global"
)

# =============================================================================
#  БЛОК ВСТАВКИ СТАТИСТИКИ ЗВОНКОВ (TSV) – РАСПОЛОЖЕН ВВЕРХУ
# =============================================================================
st.subheader("📞 Вставка статистики звонков (TSV)")

with st.expander("Инструкция по вставке данных"):
    st.markdown("""
    Скопируйте данные из отчётов в поля ниже. Формат – **табуляция** (TSV). Первая строка должна содержать заголовки:
    - **Исходящие**: `Сотрудник\tКоличество успешных вызовов\tКоличество неуспешных вызовов\tОбщее количество вызовов`
    - **Входящие**: `Сотрудник\tКоличество успешных вызовов\tКоличество неуспешных вызовов\tОбщее количество вызовов` (используется только столбец успешных)
    
    Числа могут содержать пробелы (например, `1 029`). Данные будут автоматически объединены с таблицей рекрутеров.
    """)

col1, col2 = st.columns(2)
with col1:
    outgoing_text = st.text_area(
        "Вставьте данные для **исходящих** звонков",
        height=300,
        placeholder="Сотрудник\tКоличество успешных вызовов\tКоличество неуспешных вызовов\tОбщее количество вызовов\nБалдин Александр\t1029\t242\t1271\n...",
        key="outgoing_calls_text"
    )
with col2:
    incoming_text = st.text_area(
        "Вставьте данные для **входящих** звонков",
        height=300,
        placeholder="Сотрудник\tКоличество успешных вызовов\tКоличество неуспешных вызовов\tОбщее количество вызовов\nБалдин Александр\t16\t1971\t1987\n...",
        key="incoming_calls_text"
    )

# ---- Функция парсинга TSV ----
def parse_calls_tsv(text):
    if not text.strip():
        return None
    try:
        lines = text.strip().splitlines()
        if len(lines) < 2:
            return None
        header = lines[0].split('\t')
        if len(header) < 4:
            st.warning("Неверный формат: требуется минимум 4 колонки (Сотрудник, Успешные, Неуспешные, Общее).")
            return None
        data = []
        for line in lines[1:]:
            if not line.strip():
                continue
            parts = line.split('\t')
            if len(parts) < 4:
                continue
            try:
                success = int(parts[1].replace(' ', ''))
                fail = int(parts[2].replace(' ', ''))
                total = int(parts[3].replace(' ', ''))
            except ValueError:
                continue
            data.append({
                'Сотрудник': parts[0].strip(),
                'Успешные': success,
                'Неуспешные': fail,
                'Общее': total
            })
        if not data:
            return None
        df = pd.DataFrame(data)
        return df
    except Exception as e:
        st.error(f"Ошибка при парсинге: {e}")
        return None

# ---- Парсим вставленные данные ----
df_outgoing = parse_calls_tsv(outgoing_text) if outgoing_text else None   # исходящие
df_incoming = parse_calls_tsv(incoming_text) if incoming_text else None    # входящие

# ---- ДИАГНОСТИКА: показываем, что распарсилось ----
if df_outgoing is not None or df_incoming is not None:
    with st.expander("🔍 Проверка распарсенных данных (диагностика)"):
        if df_outgoing is not None:
            st.write("**Исходящие (df_outgoing) – первые строки:**")
            st.dataframe(df_outgoing.head())
        else:
            st.info("Исходящие не вставлены.")
        if df_incoming is not None:
            st.write("**Входящие (df_incoming) – первые строки:**")
            st.dataframe(df_incoming.head())
        else:
            st.info("Входящие не вставлены.")

# ---- Применение фильтров к основному отчету (с учётом рекрутеров и без) ----
df_main_filtered_all = None
df_main_filtered = None
if df_main is not None and selected_sources:
    df_main_temp = df_main[df_main['Источник ОМПП'].isin(selected_sources)]
    
    if date_range_main and len(date_range_main) == 2:
        start_date, end_date = date_range_main
        df_main_temp = df_main_temp[
            (df_main_temp['Дата направления'].dt.date >= start_date) &
            (df_main_temp['Дата направления'].dt.date <= end_date)
        ]
    
    df_main_temp['год_напр'] = df_main_temp['Дата направления'].dt.year
    df_main_temp['мес_напр'] = df_main_temp['Дата направления'].dt.month
    df_main_temp['год_зв'] = df_main_temp['Дата последнего звонка'].dt.year
    df_main_temp['мес_зв'] = df_main_temp['Дата последнего звонка'].dt.month
    
    same_month = (df_main_temp['год_зв'] == df_main_temp['год_напр']) & (df_main_temp['мес_зв'] == df_main_temp['мес_напр'])
    prev_month = (df_main_temp['год_зв'] == df_main_temp['год_напр']) & (df_main_temp['мес_зв'] == df_main_temp['мес_напр'] - 1)
    prev_month_jan = (df_main_temp['год_зв'] == df_main_temp['год_напр'] - 1) & (df_main_temp['мес_напр'] == 1) & (df_main_temp['мес_зв'] == 12)
    
    df_main_temp['filter_last_call'] = same_month | prev_month | prev_month_jan
    df_main_temp = df_main_temp[df_main_temp['filter_last_call'] & df_main_temp['Дата последнего звонка'].notna()]
    df_main_temp = df_main_temp.reset_index(drop=True)
    
    df_main_filtered_all = df_main_temp.copy()
    
    if selected_recruiters:
        df_main_filtered = df_main_temp[df_main_temp['Рекрутер'].isin(selected_recruiters)]
    else:
        df_main_filtered = df_main_temp
else:
    df_main_filtered = None
    df_main_filtered_all = None

# ---- Применение фильтров к KPI (с учётом рекрутеров и без) ----
df_kpi_filtered_all = None
df_kpi_filtered = None
if df_kpi is not None and selected_sources:
    df_kpi_temp = df_kpi[df_kpi['Источник ОМПП'].isin(selected_sources)]
    
    if date_range_kpi and len(date_range_kpi) == 2:
        start_date_kpi, end_date_kpi = date_range_kpi
        df_kpi_temp = df_kpi_temp[
            (df_kpi_temp['Дата первой смены'].dt.date >= start_date_kpi) &
            (df_kpi_temp['Дата первой смены'].dt.date <= end_date_kpi)
        ]
    
    df_kpi_temp['год_смены'] = df_kpi_temp['Дата первой смены'].dt.year
    df_kpi_temp['мес_смены'] = df_kpi_temp['Дата первой смены'].dt.month
    df_kpi_temp['год_зв'] = df_kpi_temp['Дата последнего звонка'].dt.year
    df_kpi_temp['мес_зв'] = df_kpi_temp['Дата последнего звонка'].dt.month
    
    same_month_kpi = (df_kpi_temp['год_зв'] == df_kpi_temp['год_смены']) & (df_kpi_temp['мес_зв'] == df_kpi_temp['мес_смены'])
    prev_month_kpi = (df_kpi_temp['год_зв'] == df_kpi_temp['год_смены']) & (df_kpi_temp['мес_зв'] == df_kpi_temp['мес_смены'] - 1)
    prev_month_jan_kpi = (df_kpi_temp['год_зв'] == df_kpi_temp['год_смены'] - 1) & (df_kpi_temp['мес_смены'] == 1) & (df_kpi_temp['мес_зв'] == 12)
    
    df_kpi_temp['filter_last_call'] = same_month_kpi | prev_month_kpi | prev_month_jan_kpi
    df_kpi_temp = df_kpi_temp[df_kpi_temp['filter_last_call'] & df_kpi_temp['Дата последнего звонка'].notna()]
    df_kpi_temp = df_kpi_temp.reset_index(drop=True)
    
    df_kpi_filtered_all = df_kpi_temp.copy()
    
    if selected_recruiters:
        df_kpi_filtered = df_kpi_temp[df_kpi_temp['Рекрутер'].isin(selected_recruiters)]
    else:
        df_kpi_filtered = df_kpi_temp
else:
    df_kpi_filtered = None
    df_kpi_filtered_all = None

# ---- Применение фильтров к откликам (с учётом рекрутеров и без) ----
df_responses_filtered_all = None
df_responses_filtered = None
if df_responses is not None:
    df_responses_temp = df_responses.copy()
    if date_range_resp and len(date_range_resp) == 2 and 'Дата отклика' in df_responses_temp.columns:
        start_date_resp, end_date_resp = date_range_resp
        df_responses_temp = df_responses_temp[
            (df_responses_temp['Дата отклика'].dt.date >= start_date_resp) &
            (df_responses_temp['Дата отклика'].dt.date <= end_date_resp)
        ]
    
    df_responses_filtered_all = df_responses_temp.copy()
    
    if selected_recruiters:
        col_recr_resp = find_column(df_responses_temp, ['рекрутер в crm', 'рекрутер в срм'])
        if col_recr_resp is not None:
            df_responses_filtered = df_responses_temp[df_responses_temp[col_recr_resp].isin(selected_recruiters)]
        else:
            df_responses_filtered = df_responses_temp
    else:
        df_responses_filtered = df_responses_temp

# ---- Обработка звонков из CSV (для средней длительности) ----
df_calls_filtered = None
if df_calls is not None:
    col_who = find_column(df_calls, ['кто звонил', 'звонивший', 'сотрудник', 'абонент', 'рекрутер'])
    col_call_type = find_column(df_calls, ['тип вызова', 'тип звонка', 'направление', 'вид'])
    col_duration = find_column(df_calls, ['длительность, сек', 'длительность', 'длительность (сек)', 'время разговора', 'длит'])

    if col_who is None or col_call_type is None or col_duration is None:
        st.warning(f"В файле звонков не найдены необходимые столбцы. Доступные колонки: {list(df_calls.columns)}")
    else:
        rename_calls = {
            col_who: 'Кто звонил',
            col_call_type: 'Тип вызова',
            col_duration: 'Длительность, сек'
        }
        df_calls = df_calls.rename(columns=rename_calls)
        df_calls = df_calls.loc[:, ~df_calls.columns.duplicated()]

        df_calls['Рекрутер'] = df_calls['Кто звонил'].astype(str).str.replace('"', '').str.strip()
        df_calls_out = df_calls[df_calls['Тип вызова'] == 'Исходящий'].copy()
        if selected_recruiters:
            df_calls_filtered = df_calls_out[df_calls_out['Рекрутер'].isin(selected_recruiters)]
        else:
            df_calls_filtered = df_calls_out

# ---- Обработка откликов: таблица (будет выведена позже) ----
merged_resp = None
merged_resp_all = None
df_resp_for_city = None

def compute_responses_table(df_resp_filtered):
    if df_resp_filtered is None:
        return None
    col_phone_resp = find_column(df_resp_filtered, ['телефон соискателя'])
    col_recruiter_resp = find_column(df_resp_filtered, ['рекрутер в crm', 'рекрутер в срм'])
    col_status_resp = find_column(df_resp_filtered, ['статус рекрутера'])
    col_first_shift_resp = find_column(df_resp_filtered, ['первая смена после отклика'])
    col_city_vacancy = find_column(df_resp_filtered, ['город вакансии'])

    if col_phone_resp is None or col_recruiter_resp is None:
        return None

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

    df_resp = df_resp_filtered.rename(columns=rename_resp)
    df_resp = df_resp.loc[:, ~df_resp.columns.duplicated()]

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

    merged = responses_count.merge(reg_count, on='Рекрутер', how='left').fillna(0)
    merged = merged.merge(invited_count, on='Рекрутер', how='left').fillna(0)
    merged = merged.merge(worked_count, on='Рекрутер', how='left').fillna(0)

    for col in ['Кол-во откликов', 'Кол-во регистраций', 'Кол-во направленных из откликов', 'Вышедшие из откликов']:
        merged[col] = merged[col].astype(int)

    merged['Конв. отклик->регистр, %'] = (merged['Кол-во регистраций'] / merged['Кол-во откликов'] * 100).round(1)
    merged['Конв. регистр->направл, %'] = (merged['Кол-во направленных из откликов'] / merged['Кол-во регистраций'] * 100).round(1).fillna(0)
    merged['Конв. направл->вышед, %'] = (merged['Вышедшие из откликов'] / merged['Кол-во направленных из откликов'] * 100).round(1).fillna(0)
    merged['Конв. отклик->вышед, %'] = (merged['Вышедшие из откликов'] / merged['Кол-во откликов'] * 100).round(1)

    for col in ['Конв. отклик->регистр, %', 'Конв. регистр->направл, %', 'Конв. направл->вышед, %', 'Конв. отклик->вышед, %']:
        merged[col] = merged[col].fillna(0).replace([float('inf'), -float('inf')], 0)
        merged[col] = merged[col].astype(str) + '%'

    return merged

if df_responses_filtered is not None:
    merged_resp = compute_responses_table(df_responses_filtered)
    if merged_resp is not None:
        col_phone_resp = find_column(df_responses_filtered, ['телефон соискателя'])
        col_city_vacancy = find_column(df_responses_filtered, ['город вакансии'])
        if col_phone_resp is not None and col_city_vacancy is not None:
            df_resp_for_city = df_responses_filtered.copy()
            df_resp_for_city['Телефон'] = df_resp_for_city[col_phone_resp]
            df_resp_for_city['Город вакансии'] = df_resp_for_city[col_city_vacancy]

if df_responses_filtered_all is not None:
    merged_resp_all = compute_responses_table(df_responses_filtered_all)

# ---- 1. Объединённая таблица рекрутеров (из обоих отчетов) ----
recruiter_data = {}
recruiter_data_all = {}

def build_recruiter_data(df_main_f, df_kpi_f, merged_resp_f):
    data = {}
    if df_main_f is not None:
        main_counts = df_main_f.groupby('Рекрутер')['Телефон'].nunique().reset_index()
        main_counts.columns = ['Рекрутер', 'Кол-во направленных']

        if 'Дата первой подтвержденной смены за всю жизнь' in df_main_f.columns:
            df_with_shift = df_main_f[df_main_f['Дата первой подтвержденной смены за всю жизнь'].notna()]
            worked_main = df_with_shift.groupby('Рекрутер')['Телефон'].nunique().reset_index()
            worked_main.columns = ['Рекрутер', 'Вышло из приглашенных']
            main_counts = main_counts.merge(worked_main, on='Рекрутер', how='left').fillna(0)
            main_counts['Вышло из приглашенных'] = main_counts['Вышло из приглашенных'].astype(int)
        else:
            main_counts['Вышло из приглашенных'] = 0

        for _, row in main_counts.iterrows():
            recruiter = row['Рекрутер']
            if recruiter not in data:
                data[recruiter] = {}
            data[recruiter]['Кол-во направленных'] = row['Кол-во направленных']
            data[recruiter]['Вышло из приглашенных'] = row['Вышло из приглашенных']

    if df_kpi_f is not None:
        kpi_counts = df_kpi_f.groupby('Рекрутер')['Телефон гигера'].nunique().reset_index()
        kpi_counts.columns = ['Рекрутер', 'Вышедшие (с дошедшими)']
        for _, row in kpi_counts.iterrows():
            recruiter = row['Рекрутер']
            if recruiter not in data:
                data[recruiter] = {}
            data[recruiter]['Вышедшие (с дошедшими)'] = row['Вышедшие (с дошедшими)']

    if merged_resp_f is not None:
        for _, row in merged_resp_f.iterrows():
            recruiter = row['Рекрутер']
            if recruiter not in data:
                data[recruiter] = {}
            data[recruiter]['Кол-во откликов'] = row['Кол-во откликов']
            data[recruiter]['Конверсия из откликов в вышедших, %'] = row['Конв. отклик->вышед, %']

    return data

recruiter_data = build_recruiter_data(df_main_filtered, df_kpi_filtered, merged_resp)
recruiter_data_all = build_recruiter_data(df_main_filtered_all, df_kpi_filtered_all, merged_resp_all)

def create_recruiter_df(data):
    if not data:
        return None
    df = pd.DataFrame.from_dict(data, orient='index').reset_index()
    df.rename(columns={'index': 'Рекрутер'}, inplace=True)

    numeric_cols = ['Кол-во направленных', 'Вышло из приглашенных', 'Вышедшие (с дошедшими)', 'Кол-во откликов']
    for col in numeric_cols:
        if col not in df.columns:
            df[col] = 0
        df[col] = df[col].fillna(0).astype(int)

    df['Конверсия из пригл. в вышедших из приглашенных, %'] = (
        df['Вышло из приглашенных'] / df['Кол-во направленных'] * 100
    ).round(1).fillna(0).astype(str) + '%'

    df['Конверсия из приглашенных в вышедших с дошедшими, %'] = (
        df['Вышедшие (с дошедшими)'] / df['Кол-во направленных'] * 100
    ).round(1).fillna(0).astype(str) + '%'

    if 'Конверсия из откликов в вышедших, %' not in df.columns:
        df['Конверсия из откликов в вышедших, %'] = '0%'

    return df

df_recruiters = create_recruiter_df(recruiter_data)
df_recruiters_all = create_recruiter_df(recruiter_data_all)

# ---- Формируем статистику звонков (используем в лидерборде и таблице) ----
recruiter_list = []
if df_recruiters is not None and not df_recruiters.empty:
    recruiter_list = df_recruiters[~df_recruiters['Рекрутер'].str.contains('Итого|Среднее', na=False)]['Рекрутер'].tolist()

stats_df = None
if recruiter_list:
    stats_df = pd.DataFrame({'Рекрутер': recruiter_list})
    
    # ---- Исходящие данные (из df_outgoing) ----
    if df_outgoing is not None:
        merged_out = stats_df.merge(df_outgoing, left_on='Рекрутер', right_on='Сотрудник', how='left')
        stats_df['Кол-во исходящих'] = merged_out['Общее'].fillna(0).astype(int)
        stats_df['Успешные исх.'] = merged_out['Успешные'].fillna(0).astype(int)
        stats_df['Неуспешные исх.'] = merged_out['Неуспешные'].fillna(0).astype(int)
    else:
        # Если нет вставленных исходящих, используем данные из CSV (только общее количество)
        if df_calls_filtered is not None and not df_calls_filtered.empty:
            calls_count = df_calls_filtered.groupby('Рекрутер').size().reset_index(name='Кол-во исходящих')
            stats_df = stats_df.merge(calls_count, on='Рекрутер', how='left')
            stats_df['Кол-во исходящих'] = stats_df['Кол-во исходящих'].fillna(0).astype(int)
            stats_df['Успешные исх.'] = 0
            stats_df['Неуспешные исх.'] = 0
        else:
            stats_df['Кол-во исходящих'] = 0
            stats_df['Успешные исх.'] = 0
            stats_df['Неуспешные исх.'] = 0
    
    # ---- Входящие данные (из df_incoming) ----
    if df_incoming is not None:
        merged_in = stats_df.merge(df_incoming, left_on='Рекрутер', right_on='Сотрудник', how='left')
        stats_df['Кол-во входящих'] = merged_in['Успешные'].fillna(0).astype(int)
    else:
        stats_df['Кол-во входящих'] = 0
    
    # ---- Вычисляем итоговые столбцы ----
    stats_df['Всего звонков'] = stats_df['Кол-во исходящих'] + stats_df['Кол-во входящих']
    total_outgoing = stats_df['Успешные исх.'] + stats_df['Неуспешные исх.']
    stats_df['% Недозвонов'] = (stats_df['Неуспешные исх.'] / total_outgoing * 100).round(1).fillna(0)
    stats_df['% Недозвонов'] = stats_df['% Недозвонов'].astype(str) + '%'
    
    # ---- Средняя длительность из CSV (если есть) ----
    if df_calls_filtered is not None and not df_calls_filtered.empty:
        avg_dur = df_calls_filtered.groupby('Рекрутер')['Длительность, сек'].mean().reset_index()
        avg_dur.columns = ['Рекрутер', 'Средняя длительность, сек']
        stats_df = stats_df.merge(avg_dur, on='Рекрутер', how='left')
        stats_df['Средняя длительность, сек'] = stats_df['Средняя длительность, сек'].fillna(0).round(1)
    else:
        stats_df['Средняя длительность, сек'] = 0
    
    # ---- Итоговая строка ----
    total_row = {
        'Рекрутер': 'Итого (по выбранным)',
        'Кол-во исходящих': stats_df['Кол-во исходящих'].sum(),
        'Успешные исх.': stats_df['Успешные исх.'].sum(),
        'Неуспешные исх.': stats_df['Неуспешные исх.'].sum(),
        'Кол-во входящих': stats_df['Кол-во входящих'].sum(),
        'Всего звонков': stats_df['Всего звонков'].sum(),
        '% Недозвонов': '',
        'Средняя длительность, сек': stats_df['Средняя длительность, сек'].mean().round(1)
    }
    total_out = total_row['Успешные исх.'] + total_row['Неуспешные исх.']
    total_row['% Недозвонов'] = f"{(total_row['Неуспешные исх.'] / total_out * 100).round(1)}%" if total_out > 0 else "0%"
    
    stats_df = pd.concat([stats_df, pd.DataFrame([total_row])], ignore_index=True)

# ---- Данные для лидерборда по звонкам (из CSV) ----
best_avg_duration_recruiter = None
best_avg_duration_value = None
if df_calls_filtered is not None and not df_calls_filtered.empty:
    avg_duration = df_calls_filtered.groupby('Рекрутер')['Длительность, сек'].mean().reset_index()
    avg_duration.columns = ['Рекрутер', 'Средняя длительность']
    if not avg_duration.empty:
        max_row = avg_duration.loc[avg_duration['Средняя длительность'].idxmax()]
        best_avg_duration_recruiter = max_row['Рекрутер']
        best_avg_duration_value = round(max_row['Средняя длительность'], 1)

# ---- Функция для построения LeaderBoard ----
def build_leaderboard(df_recruiters, stats_df, best_avg_duration_recruiter=None, best_avg_duration_value=None):
    rows = []
    
    if df_recruiters is not None and not df_recruiters.empty:
        df_clean = df_recruiters[~df_recruiters['Рекрутер'].str.contains('Итого|Среднее', na=False)].copy()
        
        # 1. Лучшая конверсия из отклика в вышедшего (с дошедшими)
        if 'Конверсия из откликов в вышедших, %' in df_clean.columns:
            df_clean['conv_resp_to_worked'] = df_clean['Конверсия из откликов в вышедших, %'].str.replace('%', '').astype(float)
            best_conv_resp = df_clean.loc[df_clean['conv_resp_to_worked'].idxmax()] if not df_clean.empty else None
            if best_conv_resp is not None:
                rows.append({
                    'Категория': 'Лучшая конверсия из отклика → вышедший (с дошедшими)',
                    'Рекрутер': best_conv_resp['Рекрутер'],
                    'Значение': f"{best_conv_resp['conv_resp_to_worked']:.1f}%"
                })

        # 2. Лучшая конверсия из приглашенного в вышедшего (из приглашенных)
        if 'Конверсия из пригл. в вышедших из приглашенных, %' in df_clean.columns:
            df_clean['conv_inv_to_worked'] = df_clean['Конверсия из пригл. в вышедших из приглашенных, %'].str.replace('%', '').astype(float)
            best_conv_inv = df_clean.loc[df_clean['conv_inv_to_worked'].idxmax()] if not df_clean.empty else None
            if best_conv_inv is not None:
                rows.append({
                    'Категория': 'Лучшая конверсия из приглашенного → вышедший (из приглашенных)',
                    'Рекрутер': best_conv_inv['Рекрутер'],
                    'Значение': f"{best_conv_inv['conv_inv_to_worked']:.1f}%"
                })

        # 3. Больше всего откликов обработано
        if 'Кол-во откликов' in df_clean.columns:
            most_responses = df_clean.loc[df_clean['Кол-во откликов'].idxmax()] if not df_clean.empty else None
            if most_responses is not None:
                rows.append({
                    'Категория': 'Больше всего откликов обработано',
                    'Рекрутер': most_responses['Рекрутер'],
                    'Значение': str(int(most_responses['Кол-во откликов']))
                })

    # 4. Лучшая средняя длительность исходящих звонков (из CSV)
    if best_avg_duration_recruiter is not None:
        rows.append({
            'Категория': 'Лучшая средняя длительность исходящих звонков',
            'Рекрутер': best_avg_duration_recruiter,
            'Значение': f"{best_avg_duration_value} сек"
        })

    # 5. Лучший по общему количеству звонков (из stats_df)
    if stats_df is not None and not stats_df.empty:
        stats_clean = stats_df[~stats_df['Рекрутер'].str.contains('Итого|Среднее', na=False)].copy()
        if 'Всего звонков' in stats_clean.columns and not stats_clean.empty:
            max_total = stats_clean.loc[stats_clean['Всего звонков'].idxmax()]
            rows.append({
                'Категория': 'Лучший по общему количеству звонков',
                'Рекрутер': max_total['Рекрутер'],
                'Значение': str(int(max_total['Всего звонков']))
            })
        if 'Кол-во входящих' in stats_clean.columns and not stats_clean.empty:
            max_incoming = stats_clean.loc[stats_clean['Кол-во входящих'].idxmax()]
            rows.append({
                'Категория': 'Лучший по количеству входящих вызовов',
                'Рекрутер': max_incoming['Рекрутер'],
                'Значение': str(int(max_incoming['Кол-во входящих']))
            })

    return pd.DataFrame(rows)

# =============================================================================
#  ВЫВОД ЛИДЕРБОРДА – ЗАГОЛОВОК ПО ЦЕНТРУ С ОТСТУПОМ, ДВЕ СТРОКИ ПО 3, СМЕЩЕНИЕ В ЦЕНТР
# =============================================================================
lb = build_leaderboard(df_recruiters, stats_df, best_avg_duration_recruiter, best_avg_duration_value)
if not lb.empty:
    st.markdown("<h2 style='text-align: center; margin-bottom: 1.5rem;'>🏆 Лидерборд</h2>", unsafe_allow_html=True)
    # Разбиваем на группы по 3 (первые 3 – отклики/конверсии, следующие 3 – звонки)
    for i in range(0, len(lb), 3):
        # Используем 5 колонок: пустая слева, 3 основные, пустая справа
        cols = st.columns([1, 3, 3, 3, 1])
        for j in range(3):
            if i + j < len(lb):
                row = lb.iloc[i + j]
                with cols[j + 1]:  # +1 из-за пустой первой колонки
                    st.metric(label=row['Категория'], value=row['Значение'], delta=row['Рекрутер'])
    st.markdown("---")
else:
    st.info("Нет данных для лидерборда.")
    
# =============================================================================
#  ОСТАЛЬНЫЕ БЛОКИ (ТАБЛИЦЫ, ГРАФИКИ, СТАТИСТИКА ЗВОНКОВ)
# =============================================================================

# ---- 1. Объединённая таблица рекрутеров ----
if df_recruiters is not None and not df_recruiters.empty:
    st.subheader("📋 Количество направленных кандидатов по рекрутерам")
    
    total_row = {'Рекрутер': 'Итого (по выбранным)'}
    avg_row = {'Рекрутер': 'Среднее (по выбранным)'}
    for col in ['Кол-во откликов', 'Кол-во направленных', 'Вышло из приглашенных', 'Вышедшие (с дошедшими)']:
        if col in df_recruiters.columns:
            total_row[col] = df_recruiters[col].sum()
            avg_row[col] = round(df_recruiters[col].mean(), 1)
    
    total_invited = total_row.get('Кол-во направленных', 0)
    total_worked_invited = total_row.get('Вышло из приглашенных', 0)
    total_worked_done = total_row.get('Вышедшие (с дошедшими)', 0)
    total_responses = total_row.get('Кол-во откликов', 0)
    
    total_row['Конверсия из пригл. в вышедших из приглашенных, %'] = (
        f"{(total_worked_invited / total_invited * 100).round(1)}%" if total_invited > 0 else "0%"
    )
    total_row['Конверсия из приглашенных в вышедших с дошедшими, %'] = (
        f"{(total_worked_done / total_invited * 100).round(1)}%" if total_invited > 0 else "0%"
    )
    total_row['Конверсия из откликов в вышедших, %'] = (
        f"{(total_worked_done / total_responses * 100).round(1)}%" if total_responses > 0 else "0%"
    )
    
    for col in ['Конверсия из пригл. в вышедших из приглашенных, %', 
                'Конверсия из приглашенных в вышедших с дошедшими, %',
                'Конверсия из откликов в вышедших, %']:
        if col in df_recruiters.columns:
            vals = df_recruiters[col].str.replace('%', '').astype(float)
            avg_row[col] = f"{round(vals.mean(), 1)}%"
    
    if df_recruiters_all is not None and not df_recruiters_all.empty:
        total_all_row = {'Рекрутер': 'Итого (включая скрытых)'}
        for col in ['Кол-во откликов', 'Кол-во направленных', 'Вышло из приглашенных', 'Вышедшие (с дошедшими)']:
            if col in df_recruiters_all.columns:
                total_all_row[col] = df_recruiters_all[col].sum()
        total_invited_all = total_all_row.get('Кол-во направленных', 0)
        total_worked_invited_all = total_all_row.get('Вышло из приглашенных', 0)
        total_worked_done_all = total_all_row.get('Вышедшие (с дошедшими)', 0)
        total_responses_all = total_all_row.get('Кол-во откликов', 0)
        total_all_row['Конверсия из пригл. в вышедших из приглашенных, %'] = (
            f"{(total_worked_invited_all / total_invited_all * 100).round(1)}%" if total_invited_all > 0 else "0%"
        )
        total_all_row['Конверсия из приглашенных в вышедших с дошедшими, %'] = (
            f"{(total_worked_done_all / total_invited_all * 100).round(1)}%" if total_invited_all > 0 else "0%"
        )
        total_all_row['Конверсия из откликов в вышедших, %'] = (
            f"{(total_worked_done_all / total_responses_all * 100).round(1)}%" if total_responses_all > 0 else "0%"
        )
        df_recruiters = pd.concat([df_recruiters, pd.DataFrame([total_all_row])], ignore_index=True)
    
    df_recruiters = pd.concat([df_recruiters, pd.DataFrame([total_row]), pd.DataFrame([avg_row])], ignore_index=True)
    
    display_cols = ['Рекрутер']
    col_config = {}
    for col in ['Кол-во откликов', 'Кол-во направленных', 'Вышло из приглашенных', 'Вышедшие (с дошедшими)',
                'Конверсия из пригл. в вышедших из приглашенных, %',
                'Конверсия из приглашенных в вышедших с дошедшими, %',
                'Конверсия из откликов в вышедших, %']:
        if col in df_recruiters.columns:
            display_cols.append(col)
            if 'Конверсия' in col:
                col_config[col] = st.column_config.TextColumn(col, width="auto")
            else:
                col_config[col] = st.column_config.NumberColumn(col, format="%d", width="auto")
    
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

def get_project_data(df_main_f, df_kpi_f):
    if df_main_f is not None:
        if 'Статус координатора' in df_main_f.columns:
            df_worked_main = df_main_f[df_main_f['Статус координатора'] == 'went_work']
        else:
            df_worked_main = df_main_f[df_main_f['Статус лида'] == 'worked']
        
        if df_worked_main.empty or 'Проект первой подтвержденной смены' not in df_worked_main.columns:
            main_project_data = pd.DataFrame(columns=['Проект', 'Кол-во приглашенных', 'Кол-во вышедших (из приглашенных)'])
        else:
            df_projects_all = df_main_f.copy()
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
            
            df_worked_main['Проект'] = df_worked_main['Проект первой подтвержденной смены'].apply(normalize_project)
            worked_main_counts = df_worked_main.groupby('Проект')['Телефон'].nunique().reset_index()
            worked_main_counts.columns = ['Проект', 'Кол-во вышедших (из приглашенных)']
            
            main_project_data = pd.merge(invited_counts, worked_main_counts, on='Проект', how='outer').fillna(0)
            main_project_data['Кол-во приглашенных'] = main_project_data['Кол-во приглашенных'].astype(int)
            main_project_data['Кол-во вышедших (из приглашенных)'] = main_project_data['Кол-во вышедших (из приглашенных)'].astype(int)
            numeric_cols = ['Кол-во приглашенных', 'Кол-во вышедших (из приглашенных)']
            main_project_data = main_project_data.groupby('Проект', as_index=False)[numeric_cols].sum()
    else:
        main_project_data = pd.DataFrame(columns=['Проект', 'Кол-во приглашенных', 'Кол-во вышедших (из приглашенных)'])

    if df_kpi_f is not None and 'Клиент' in df_kpi_f.columns:
        kpi_project_counts = df_kpi_f.groupby('Клиент')['Телефон гигера'].nunique().reset_index()
        kpi_project_counts.columns = ['Проект', 'Кол-во вышедших (с дошедшими)']
        kpi_project_counts['Проект'] = kpi_project_counts['Проект'].apply(lambda x: normalize_project(x).strip())
        kpi_project_counts = kpi_project_counts.groupby('Проект', as_index=False)['Кол-во вышедших (с дошедшими)'].sum()
    else:
        kpi_project_counts = pd.DataFrame(columns=['Проект', 'Кол-во вышедших (с дошедшими)'])

    if not main_project_data.empty:
        main_project_data['Проект'] = main_project_data['Проект'].apply(lambda x: normalize_project(x).strip())
        numeric_cols = ['Кол-во приглашенных', 'Кол-во вышедших (из приглашенных)']
        main_project_data = main_project_data.groupby('Проект', as_index=False)[numeric_cols].sum()
    if not kpi_project_counts.empty:
        kpi_project_counts['Проект'] = kpi_project_counts['Проект'].apply(lambda x: normalize_project(x).strip())
        kpi_project_counts = kpi_project_counts.groupby('Проект', as_index=False)['Кол-во вышедших (с дошедшими)'].sum()

    merged = pd.merge(main_project_data, kpi_project_counts, on='Проект', how='outer').fillna(0)
    for col in ['Кол-во приглашенных', 'Кол-во вышедших (из приглашенных)', 'Кол-во вышедших (с дошедшими)']:
        if col in merged.columns:
            merged[col] = merged[col].astype(int)
        else:
            merged[col] = 0

    merged['Конв. из приглашенных в вышедших из пригл., %'] = (
        merged['Кол-во вышедших (из приглашенных)'] / merged['Кол-во приглашенных'] * 100
    ).round(1).fillna(0).astype(str) + '%'

    merged['Конв. из приглашенных в вышедших с дошедшими, %'] = (
        merged['Кол-во вышедших (с дошедшими)'] / merged['Кол-во приглашенных'] * 100
    ).round(1).fillna(0).astype(str) + '%'

    return merged

project_data = get_project_data(df_main_filtered, df_kpi_filtered)

if not project_data.empty and project_data['Кол-во приглашенных'].sum() > 0:
    st.dataframe(
        project_data,
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

city_data = {}
city_data_all = {}

def build_city_data(df_resp_f, df_main_f, df_kpi_f):
    data = {}
    if df_resp_f is not None and 'Город вакансии' in df_resp_f.columns:
        city_resp = df_resp_f[df_resp_f['Город вакансии'].notna() & (df_resp_f['Город вакансии'].astype(str).str.strip() != '')].copy()
        city_resp['Город'] = city_resp['Город вакансии'].astype(str).str.strip()
        city_resp_counts = city_resp.groupby('Город').size().reset_index(name='Кол-во откликов')
        for _, row in city_resp_counts.iterrows():
            city = row['Город']
            if city not in data:
                data[city] = {}
            data[city]['Кол-во откликов'] = row['Кол-во откликов']

    if df_main_f is not None and 'Город' in df_main_f.columns:
        city_invited = df_main_f[
            df_main_f['Город'].notna() & 
            (df_main_f['Город'].astype(str).str.strip() != '')
        ].copy()
        city_invited['Город'] = city_invited['Город'].astype(str).str.strip()
        invited_city = city_invited.groupby('Город')['Телефон'].nunique().reset_index()
        invited_city.columns = ['Город', 'Кол-во приглашенных']
        for _, row in invited_city.iterrows():
            city = row['Город']
            if city not in data:
                data[city] = {}
            data[city]['Кол-во приглашенных'] = row['Кол-во приглашенных']

    if df_kpi_f is not None and 'Город первой смены' in df_kpi_f.columns:
        df_kpi_city = df_kpi_f[
            df_kpi_f['Рекрутер'].notna() & 
            (df_kpi_f['Рекрутер'].astype(str).str.strip() != '') &
            df_kpi_f['Город первой смены'].notna() & 
            (df_kpi_f['Город первой смены'].astype(str).str.strip() != '')
        ].copy()
        df_kpi_city['Город'] = df_kpi_city['Город первой смены'].astype(str).str.strip()
        worked_city = df_kpi_city.groupby('Город')['Телефон гигера'].nunique().reset_index()
        worked_city.columns = ['Город', 'Кол-во вышедших']
        for _, row in worked_city.iterrows():
            city = row['Город']
            if city not in data:
                data[city] = {}
            data[city]['Кол-во вышедших'] = row['Кол-во вышедших']
    return data

city_data = build_city_data(df_responses_filtered, df_main_filtered, df_kpi_filtered)
city_data_all = build_city_data(df_responses_filtered_all, df_main_filtered_all, df_kpi_filtered_all)

if city_data:
    df_city = pd.DataFrame.from_dict(city_data, orient='index').reset_index()
    df_city.rename(columns={'index': 'Город'}, inplace=True)
    
    for col in ['Кол-во откликов', 'Кол-во приглашенных', 'Кол-во вышедших']:
        if col not in df_city.columns:
            df_city[col] = 0
        df_city[col] = df_city[col].fillna(0).astype(int)
    
    total_invited = df_city['Кол-во приглашенных'].sum()
    df_city['Доля приглашенных'] = (df_city['Кол-во приглашенных'] / total_invited * 100).round(1).astype(str) + '%' if total_invited > 0 else '0%'
    df_city['Конверсия из направленных в вышедших, %'] = (
        df_city['Кол-во вышедших'] / df_city['Кол-во приглашенных'] * 100
    ).round(1).fillna(0).astype(str) + '%'
    
    total_row = {'Город': 'Итого (по выбранным)'}
    avg_row = {'Город': 'Среднее (по выбранным)'}
    for col in ['Кол-во откликов', 'Кол-во приглашенных', 'Кол-во вышедших']:
        if col in df_city.columns:
            total_row[col] = df_city[col].sum()
            avg_row[col] = round(df_city[col].mean(), 1)
    total_row['Доля приглашенных'] = '100%'
    total_row['Конверсия из направленных в вышедших, %'] = (
        f"{(total_row.get('Кол-во вышедших', 0) / total_row.get('Кол-во приглашенных', 1) * 100).round(1)}%"
    ) if total_row.get('Кол-во приглашенных', 0) > 0 else '0%'
    
    vals_conv = df_city['Конверсия из направленных в вышедших, %'].str.replace('%', '').astype(float)
    avg_row['Конверсия из направленных в вышедших, %'] = f"{round(vals_conv.mean(), 1)}%"
    
    df_city = pd.concat([df_city, pd.DataFrame([total_row]), pd.DataFrame([avg_row])], ignore_index=True)
    
    if city_data_all:
        total_all_row = {'Город': 'Итого (включая скрытых)'}
        df_city_all = pd.DataFrame.from_dict(city_data_all, orient='index').reset_index()
        df_city_all.rename(columns={'index': 'Город'}, inplace=True)
        for col in ['Кол-во откликов', 'Кол-во приглашенных', 'Кол-во вышедших']:
            if col not in df_city_all.columns:
                df_city_all[col] = 0
            total_all_row[col] = df_city_all[col].sum()
        total_all_row['Доля приглашенных'] = '100%'
        total_all_row['Конверсия из направленных в вышедших, %'] = (
            f"{(total_all_row.get('Кол-во вышедших', 0) / total_all_row.get('Кол-во приглашенных', 1) * 100).round(1)}%"
        ) if total_all_row.get('Кол-во приглашенных', 0) > 0 else '0%'
        df_city = pd.concat([df_city, pd.DataFrame([total_all_row])], ignore_index=True)
    
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

# ---- Круговая диаграмма по источникам (вышедшие с дошедшими) ----
if df_kpi_filtered is not None and not df_kpi_filtered.empty:
    st.subheader("📊 Распределение вышедших (с дошедшими) по источникам")
    pie_data = df_kpi_filtered.groupby('Источник ОМПП')['Телефон гигера'].nunique().reset_index()
    pie_data.columns = ['Источник', 'Кол-во']
    if not pie_data.empty:
        fig_pie = px.pie(
            pie_data,
            values='Кол-во',
            names='Источник',
            title='Доля вышедших по источникам',
            hover_data={'Кол-во': True},
            labels={'Кол-во': 'Кол-во вышедших'}
        )
        fig_pie.update_layout(height=600, width=800)
        fig_pie.update_traces(textposition='inside', textinfo='percent+label', pull=[0.05] * len(pie_data))
        st.plotly_chart(fig_pie, use_container_width=True)
        st.markdown("---")

# ---- Горизонтальная линейчатая диаграмма "Вышедшие (с дошедшими) по проектам" ----
if 'project_data' in locals() and project_data is not None and not project_data.empty:
    st.subheader("📊 Вышедшие (с дошедшими) по проектам")
    proj_done = project_data[project_data['Кол-во вышедших (с дошедшими)'] > 0].copy()
    if not proj_done.empty:
        proj_done = proj_done.sort_values('Кол-во вышедших (с дошедшими)', ascending=False)
        fig_proj_done = px.bar(
            proj_done,
            x='Кол-во вышедших (с дошедшими)',
            y='Проект',
            orientation='h',
            title='Количество вышедших (с дошедшими) по проектам',
            labels={'Кол-во вышедших (с дошедшими)': 'Кол-во вышедших', 'Проект': ''},
            text='Кол-во вышедших (с дошедшими)',
            color='Кол-во вышедших (с дошедшими)',
            color_continuous_scale='Teal',
            height=500
        )
        fig_proj_done.update_traces(textposition='outside')
        fig_proj_done.update_layout(yaxis={'categoryorder': 'total ascending'}, showlegend=False)
        st.plotly_chart(fig_proj_done, use_container_width=True)
        st.markdown("---")
    else:
        st.info("Нет данных по проектам для вышедших с дошедшими.")

# ---- 7. Блок: Обработка откликов (таблица) ----
if merged_resp is not None:
    st.subheader("📋 Обработка откликов")
    
    df_resp_display = merged_resp.copy()
    total_row = {'Рекрутер': 'Итого (по выбранным)'}
    avg_row = {'Рекрутер': 'Среднее (по выбранным)'}
    for col in ['Кол-во откликов', 'Кол-во регистраций', 'Кол-во направленных из откликов', 'Вышедшие из откликов']:
        if col in df_resp_display.columns:
            total_row[col] = df_resp_display[col].sum()
            avg_row[col] = round(df_resp_display[col].mean(), 1)
    
    total_resp = total_row.get('Кол-во откликов', 0)
    total_reg = total_row.get('Кол-во регистраций', 0)
    total_inv = total_row.get('Кол-во направленных из откликов', 0)
    total_worked = total_row.get('Вышедшие из откликов', 0)
    total_row['Конв. отклик->регистр, %'] = f"{(total_reg / total_resp * 100).round(1)}%" if total_resp > 0 else "0%"
    total_row['Конв. регистр->направл, %'] = f"{(total_inv / total_reg * 100).round(1)}%" if total_reg > 0 else "0%"
    total_row['Конв. направл->вышед, %'] = f"{(total_worked / total_inv * 100).round(1)}%" if total_inv > 0 else "0%"
    total_row['Конв. отклик->вышед, %'] = f"{(total_worked / total_resp * 100).round(1)}%" if total_resp > 0 else "0%"
    
    for col in ['Конв. отклик->регистр, %', 'Конв. регистр->направл, %', 'Конв. направл->вышед, %', 'Конв. отклик->вышед, %']:
        if col in df_resp_display.columns:
            vals = df_resp_display[col].str.replace('%', '').astype(float)
            avg_row[col] = f"{round(vals.mean(), 1)}%"
    
    df_resp_display = pd.concat([df_resp_display, pd.DataFrame([total_row]), pd.DataFrame([avg_row])], ignore_index=True)
    
    if merged_resp_all is not None:
        total_all_row = {'Рекрутер': 'Итого (включая скрытых)'}
        for col in ['Кол-во откликов', 'Кол-во регистраций', 'Кол-во направленных из откликов', 'Вышедшие из откликов']:
            if col in merged_resp_all.columns:
                total_all_row[col] = merged_resp_all[col].sum()
        total_resp_all = total_all_row.get('Кол-во откликов', 0)
        total_reg_all = total_all_row.get('Кол-во регистраций', 0)
        total_inv_all = total_all_row.get('Кол-во направленных из откликов', 0)
        total_worked_all = total_all_row.get('Вышедшие из откликов', 0)
        total_all_row['Конв. отклик->регистр, %'] = f"{(total_reg_all / total_resp_all * 100).round(1)}%" if total_resp_all > 0 else "0%"
        total_all_row['Конв. регистр->направл, %'] = f"{(total_inv_all / total_reg_all * 100).round(1)}%" if total_reg_all > 0 else "0%"
        total_all_row['Конв. направл->вышед, %'] = f"{(total_worked_all / total_inv_all * 100).round(1)}%" if total_inv_all > 0 else "0%"
        total_all_row['Конв. отклик->вышед, %'] = f"{(total_worked_all / total_resp_all * 100).round(1)}%" if total_resp_all > 0 else "0%"
        df_resp_display = pd.concat([df_resp_display, pd.DataFrame([total_all_row])], ignore_index=True)
    
    st.dataframe(
        df_resp_display,
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

# ---- 8. Диаграмма "Время обработки откликов в рабочее время" ----
if df_diagram is not None:
    st.subheader("📊 Время обработки откликов в рабочее время")

    try:
        dates = df_diagram.iloc[0, 1:]
        values_15 = df_diagram.iloc[1, 1:]
        values_hour = df_diagram.iloc[2, 1:]

        rows = []

        for dt_raw, val15, val60 in zip(dates, values_15, values_hour):
            dt = parse_diagram_date(str(dt_raw))
            if pd.isna(dt):
                try:
                    dt = pd.to_datetime(dt_raw, errors='coerce')
                except:
                    continue
            if pd.isna(dt):
                continue

            try:
                v15 = float(val15)
                v60 = float(val60)
            except:
                continue

            if v15 <= 1:
                v15 *= 100
            if v60 <= 1:
                v60 *= 100

            rows.append({
                'Дата': dt,
                'В течение 15 минут': v15,
                'Менее часа': v60
            })

        if rows:
            df_chart = pd.DataFrame(rows)
            df_chart['Месяц'] = df_chart['Дата'].dt.to_period('M').astype(str)
            monthly_avg = df_chart.groupby('Месяц')[['В течение 15 минут', 'Менее часа']].mean().round(1).reset_index()

            overall_avg = df_chart[['В течение 15 минут', 'Менее часа']].mean().round(1)
            overall_row = {
                'Месяц': 'Среднее за период',
                'В течение 15 минут': overall_avg['В течение 15 минут'],
                'Менее часа': overall_avg['Менее часа']
            }
            monthly_avg = pd.concat([monthly_avg, pd.DataFrame([overall_row])], ignore_index=True)

            fig = px.bar(
                monthly_avg,
                y='Месяц',
                x=['В течение 15 минут', 'Менее часа'],
                orientation='h',
                barmode='group',
                text_auto='.1f',
                title='Среднее время обработки откликов по месяцам'
            )

            fig.update_layout(
                xaxis_title='Средний процент (%)',
                yaxis_title='Месяц',
                legend_title='Показатель'
            )

            st.plotly_chart(fig, use_container_width=True)

            st.dataframe(
                monthly_avg,
                use_container_width=True
            )

        else:
            st.warning("На листе 'Диаграмма' не найдено данных.")

    except Exception as e:
        st.error(f"Ошибка при обработке листа 'Диаграмма': {e}")

# ---- 9. Блок: Статистика звонков (таблица) ----
if stats_df is not None and not stats_df.empty:
    st.subheader("📞 Статистика звонков")
    st.dataframe(
        stats_df,
        use_container_width=True,
        column_config={
            "Рекрутер": st.column_config.TextColumn("Рекрутер", width="auto"),
            "Кол-во исходящих": st.column_config.NumberColumn("Кол-во исходящих", format="%d", width="auto"),
            "Успешные исх.": st.column_config.NumberColumn("Успешные исх.", format="%d", width="auto"),
            "Неуспешные исх.": st.column_config.NumberColumn("Неуспешные исх.", format="%d", width="auto"),
            "Кол-во входящих": st.column_config.NumberColumn("Кол-во входящих", format="%d", width="auto"),
            "Всего звонков": st.column_config.NumberColumn("Всего звонков", format="%d", width="auto"),
            "% Недозвонов": st.column_config.TextColumn("% Недозвонов", width="auto"),
            "Средняя длительность, сек": st.column_config.NumberColumn("Средняя длительность, сек", format="%.1f", width="auto"),
        }
    )
else:
    st.info("Нет данных для отображения статистики звонков.")
