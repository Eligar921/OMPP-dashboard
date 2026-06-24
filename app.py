# ---- Объединённый блок: Приглашенные/вышедшие по городам ----
if df_main_filtered is not None and 'Город' in df_main_filtered.columns:
    st.subheader("🏙️ Приглашенные/вышедшие по городам")

    # Приглашенные (из основного отчёта) – группировка по "Город"
    city_invited = df_main_filtered[
        df_main_filtered['Город'].notna() & 
        (df_main_filtered['Город'].astype(str).str.strip() != '')
    ].copy()
    city_invited['Город'] = city_invited['Город'].astype(str).str.strip()
    invited_city = city_invited.groupby('Город')['Телефон'].nunique().reset_index()
    invited_city.columns = ['Город', 'Кол-во приглашенных']

    # Вышедшие (из KPI) – группировка по "Город первой смены"
    if df_kpi_filtered is not None and 'Город первой смены' in df_kpi_filtered.columns:
        # Исключаем пустых рекрутеров и пустые города
        df_kpi_city = df_kpi_filtered[
            df_kpi_filtered['Рекрутер'].notna() & 
            (df_kpi_filtered['Рекрутер'].astype(str).str.strip() != '') &
            df_kpi_filtered['Город первой смены'].notna() & 
            (df_kpi_filtered['Город первой смены'].astype(str).str.strip() != '')
        ].copy()
        df_kpi_city['Город'] = df_kpi_city['Город первой смены'].astype(str).str.strip()
        worked_city = df_kpi_city.groupby('Город')['Телефон гигера'].nunique().reset_index()
        worked_city.columns = ['Город', 'Кол-во вышедших']
    else:
        worked_city = pd.DataFrame(columns=['Город', 'Кол-во вышедших'])

    # Объединение
    merged_city = pd.merge(invited_city, worked_city, on='Город', how='outer').fillna(0)
    merged_city['Кол-во приглашенных'] = merged_city['Кол-во приглашенных'].astype(int)
    merged_city['Кол-во вышедших'] = merged_city['Кол-во вышедших'].astype(int)

    total_invited = merged_city['Кол-во приглашенных'].sum()
    merged_city['Доля приглашенных'] = (merged_city['Кол-во приглашенных'] / total_invited * 100).round(1).astype(str) + '%'

    merged_city['Конверсия из направленных в вышедших, %'] = (
        merged_city['Кол-во вышедших'] / merged_city['Кол-во приглашенных'] * 100
    ).round(1).fillna(0).astype(str) + '%'

    merged_city = merged_city.sort_values('Кол-во приглашенных', ascending=False)

    st.dataframe(
        merged_city,
        use_container_width=True,
        column_config={
            "Город": st.column_config.TextColumn("Город", width="auto"),
            "Кол-во приглашенных": st.column_config.NumberColumn("Кол-во приглашенных", format="%d", width="auto"),
            "Доля приглашенных": st.column_config.TextColumn("Доля приглашенных", width="auto"),
            "Кол-во вышедших": st.column_config.NumberColumn("Кол-во вышедших", format="%d", width="auto"),
            "Конверсия из направленных в вышедших, %": st.column_config.TextColumn("Конверсия из направленных в вышедших, %", width="auto"),
        }
    )
else:
    st.info("Столбец 'Город' не найден в основном отчёте, таблица городов пропущена.")
