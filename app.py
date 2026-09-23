# ============================================================
# INTERFACE
# ============================================================
col1, col2 = st.columns(2)
with col1:
    uploaded = st.file_uploader("📂 Upload do Boletim (PDF)", type="pdf", key="pdf")
with col2:
    apura_uploaded = st.file_uploader(
        "📊 Upload da Apuração (Excel)",
        type=["xlsx", "xls"], key="apura"
    )

etapa_compare = st.selectbox(
    "🎯 Etapa para verificar na comparação:",
    ["1ª Etapa", "2ª Etapa", "3ª Etapa", "Soma das 3 Etapas"],
    index=1,
    disabled=(uploaded is None or apura_uploaded is None),
)

if uploaded is None:
    st.info("👈 Faça o upload do PDF do boletim para começar.")
else:
    with st.spinner("Processando o boletim..."):
        df_notas = extrair_dados_pdf(uploaded.getvalue())

    if df_notas.empty:
        st.error("❌ Nenhum dado pôde ser extraído do PDF.")
    else:
        # ---------- Geração do boletim (funcionalidade existente) ----------
        planilhas = []
        for (turma, mat, aluno), df_aluno in df_notas.groupby(
            ["Turma", "Matrícula", "Aluno"], sort=False
        ):
            planilhas.append({
                "matricula": mat, "aluno": aluno, "turma": turma,
                "df": gerar_boletim(df_aluno),
                "situacao": sit_geral(df_aluno),
            })
        wb = gerar_excel_unico(planilhas)
        buf = io.BytesIO()
        wb.save(buf)

        st.success(f"✅ Boletim processado — {len(planilhas)} aluno(s) encontrado(s).")
        st.download_button(
            label="📥 Baixar Excel dos Boletins",
            data=buf.getvalue(),
            file_name="boletim_por_disciplina.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        # ---------- Comparação Apura x Boletim ----------
        if apura_uploaded is not None:
            st.markdown("---")
            st.subheader(f"🔍 Comparação com Apuração — {etapa_compare}")

            try:
                apura_df = pd.read_excel(apura_uploaded)
            except Exception as e:
                st.error(f"Erro ao ler a planilha de apuração: {e}")
                st.stop()

            # Verifica colunas obrigatórias
            obrig = {"MATRICULA", "NOME ALUNO", "RESULTADO", "DISCIPLINAS"}
            faltando = obrig - set(apura_df.columns)
            if faltando:
                st.error(f"Colunas ausentes na planilha de apuração: {faltando}")
            else:
                with st.spinner("Comparando disciplinas..."):
                    df_comp = gerar_comparacao(apura_df, df_notas, etapa_compare)

                if df_comp.empty:
                    st.warning("Nenhuma comparação pôde ser gerada.")
                else:
                    st.dataframe(df_comp, use_container_width=True, hide_index=True)

                    # Estatísticas rápidas
                    total = len(df_comp)
                    ok = (df_comp["Status"] == "REGULARIZADO").sum()
                    pend = (df_comp["Status"] == "PENDENTE").sum()
                    c1, c2, c3 = st.columns(3)
                    c1.metric("Total aluno-disciplina", total)
                    c2.metric("✅ Regularizados", int(ok))
                    c3.metric("⚠️ Pendentes", int(pend))

                    wb_comp = gerar_excel_comparacao(df_comp, etapa_compare)
                    buf2 = io.BytesIO()
                    wb_comp.save(buf2)

                    st.download_button(
                        label=f"📥 Baixar Excel da Comparação ({etapa_compare})",
                        data=buf2.getvalue(),
                        file_name=f"comparacao_apura_{etapa_compare.replace(' ', '_').replace('ª','a')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )