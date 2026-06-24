    # ---- Поиск столбцов (с приоритетом для "Город") ----
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

    # ... остальные столбцы ...

    col_city = find_column(['город'], exact_match='Город')  # сначала ищем точное "Город"
    if col_city is None:
        # если нет точного, ищем по ключевому слову
        col_city = find_column(['город'])

    # ---- Диагностика: выведем имя найденного столбца города ----
    st.sidebar.write(f"**Найден столбец 'Город':** {col_city if col_city else 'не найден'}")
    st.sidebar.write("**Все столбцы:**", list(df.columns))
