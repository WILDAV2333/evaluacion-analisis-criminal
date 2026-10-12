for idx, (crit_nombre, peso) in enumerate(CRITERIOS_ESCRITO):
            col_target = cols_escrito[idx % 2]
            with col_target:
                nivel = st.selectbox(
                    f"{crit_nombre} (Máx: {peso} pts)",
                    OPCIONES_NIVEL,
                    index=None,
                    placeholder="-- Seleccione calificación --",
                    key=f"esc_{idx}"
                )
