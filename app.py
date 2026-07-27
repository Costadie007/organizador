                            out_p.seek(0)

                            nome_sub_arq = f"Planilha_Parte_{p+1}_(Linhas_{dado_inicio}-{dado_fim}).xlsx"
                            zf.writestr(nome_sub_arq, out_p.getvalue())

                    zip_buffer.seek(0)
                    st.success(f"🎉 Divisão concluída! **{qtd_partes_calculadas}** planilhas geradas em um arquivo .ZIP.")
                    st.download_button(
                        label="📥 BAIXAR PACOTE COM AS PLANILHAS DIVIDIDAS (.ZIP)",
                        data=zip_buffer,
                        file_name="Planilhas_Divididas.zip",
                        mime="application/zip",
                        use_container_width=True
                    )

# ==========================================
# ABA 2: PAINEL DE ADMIN
# ==========================================
if e_admin and tab_admin:
    with tab_admin:
        st.markdown("## 👑 Gerenciamento de Usuários")
        todos_usuarios = carregar_usuarios()
        
        for usr, dados in todos_usuarios.items():
            col_u, col_r, col_act = st.columns([2, 2, 2])
            col_u.write(f"**`{usr}`**")
            col_r.write(dados.get("role", "user"))
            with col_act:
                if usr != USUARIO_ADMIN:
                    if st.button("Remover", key=f"del_{usr}"):
                        alterar_status_usuario(usr, "excluir")
                        st.rerun()

# --- RODAPÉ ---
st.markdown("<br><br>---", unsafe_allow_html=True)
st.markdown("<div style='text-align:center; color:#888;'>Desenvolvimento e Engenharia por <strong>Diego Costa</strong></div>", unsafe_allow_html=True)
