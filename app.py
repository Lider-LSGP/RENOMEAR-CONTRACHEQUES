from __future__ import annotations

from pathlib import Path

import pandas as pd
import streamlit as st

from processor import create_output_zips, iter_pdf_uploads, load_employee_map, summarize_records

st.set_page_config(page_title='Renomeador Domínio', page_icon='📄', layout='wide')

ASSET_LOGO = Path(__file__).parent / 'assets' / 'logo_liderlimpe.png'

if 'resultado_processamento' not in st.session_state:
    st.session_state['resultado_processamento'] = None

st.markdown(
    """
    <style>
    .stApp { background: linear-gradient(180deg, #f4f7fb 0%, #eef3fb 100%); }
    .hero { background: linear-gradient(135deg, #0f3b82 0%, #123f8f 55%, #f57c00 100%); padding: 28px 32px; border-radius: 24px; color: white; box-shadow: 0 20px 40px rgba(15, 59, 130, 0.18); margin-bottom: 18px; }
    .hero h1 { margin: 0; font-size: 2.2rem; line-height: 1.15; }
    .hero p { margin: 10px 0 0 0; font-size: 1.02rem; opacity: 0.96; }
    .card { background: white; border-radius: 18px; padding: 18px 20px; border: 1px solid rgba(15, 59, 130, 0.08); box-shadow: 0 10px 25px rgba(19, 38, 66, 0.06); height: 100%; }
    .mini { font-size: 0.92rem; color: #49617d; }
    .pill { display: inline-block; padding: 6px 10px; border-radius: 999px; background: #eef5ff; color: #0f3b82; font-size: 0.85rem; margin-right: 6px; margin-bottom: 6px; border: 1px solid #d8e5ff; }
    .success-box { background: #f0fff6; border: 1px solid #c9f0d8; padding: 16px 18px; border-radius: 16px; color: #14532d; }
    .warn-box { background: #fff8eb; border: 1px solid #fde0b3; padding: 16px 18px; border-radius: 16px; color: #8a4b08; }
    </style>
    """,
    unsafe_allow_html=True,
)

hero_left, hero_right = st.columns([5, 1.3])
with hero_left:
    st.markdown(
        """
        <div class="hero">
            <h1>Renomeador de Contracheques - Domínio</h1>
            <p>Identifica empresa pelo CNPJ do PDF, renomeia demitidos normalmente e separa a ATIVA por PMV, SEDU e GERAL.</p>
        </div>
        """,
        unsafe_allow_html=True,
    )
with hero_right:
    if ASSET_LOGO.exists():
        st.image(str(ASSET_LOGO), use_container_width=True)

info1, info2, info3 = st.columns(3)
with info1:
    st.markdown('<div class="card"><h4>Empresa pelo PDF</h4><div class="mini">A empresa é identificada pelo CNPJ dentro do PDF, evitando erros quando o mesmo código existe em empresas diferentes.</div></div>', unsafe_allow_html=True)
with info2:
    st.markdown('<div class="card"><h4>Demitidos incluídos</h4><div class="mini">Contracheques de situação 8 (Demitido) também são renomeados e entram nas pastas normalmente.</div></div>', unsafe_allow_html=True)
with info3:
    st.markdown('<div class="card"><h4>Lotes grandes</h4><div class="mini">Você pode gerar só o ZIP separado, só o geral, ou os dois. O resultado fica salvo na sessão para o download não recarregar tudo.</div></div>', unsafe_allow_html=True)

aba_processar, aba_regras = st.tabs(['Processar arquivos', 'Regras do app'])

with aba_processar:
    col1, col2 = st.columns([1, 2])
    with col1:
        planilha = st.file_uploader(
            'Planilha da Domínio (.xls/.xlsx) - opcional',
            type=['xls', 'xlsx'],
            help='Se enviada, o CPF vem direto da planilha (mais confiável). Sem planilha, o CPF é lido do PDF quando possível.',
        )
    with col2:
        pdf_uploads = st.file_uploader(
            'Contracheques (PDFs ou ZIP com PDFs)',
            type=['pdf', 'zip'],
            accept_multiple_files=True,
            help='Você pode enviar vários PDFs ou um ZIP já com todos os contracheques.',
        )

    opt1, opt2, opt3 = st.columns(3)
    with opt1:
        gerar_zip_separado = st.checkbox('Gerar ZIP separado por empresa', value=True)
    with opt2:
        gerar_zip_geral = st.checkbox('Gerar ZIP geral', value=False)
    with opt3:
        mostrar_previa = st.checkbox('Mostrar prévia detalhada na tela', value=False)

    st.markdown(
        ' '.join([
            '<span class="pill">Renomeia para MM-AAAA-CPF.pdf</span>',
            '<span class="pill">ATIVA → PMV / SEDU / GERAL</span>',
            '<span class="pill">Demitidos renomeados</span>',
            '<span class="pill">Empresa pelo CNPJ do PDF</span>',
        ]),
        unsafe_allow_html=True,
    )

    if not gerar_zip_separado and not gerar_zip_geral:
        st.markdown('<div class="warn-box"><strong>Atenção:</strong> selecione pelo menos um tipo de ZIP para gerar.</div>', unsafe_allow_html=True)

    processar = st.button('Processar arquivos', type='primary', use_container_width=True)

    if processar:
        if not pdf_uploads:
            st.error('Envie ao menos um PDF ou ZIP com PDFs.')
            st.stop()
        if not gerar_zip_separado and not gerar_zip_geral:
            st.error('Selecione pelo menos um tipo de ZIP para gerar.')
            st.stop()

        with st.spinner('Lendo arquivos e montando os ZIPs...'):
            employee_map = {}
            planilha_ok = False
            if planilha is not None:
                try:
                    employee_map = load_employee_map(planilha)
                    planilha_ok = True
                except Exception as exc:
                    st.error(f'Erro ao ler a planilha: {exc}')
                    st.stop()

            pdf_items = iter_pdf_uploads(pdf_uploads)
            zip_sep, zip_all, records = create_output_zips(
                pdf_items,
                employee_map,
                make_separate_zip=gerar_zip_separado,
                make_general_zip=gerar_zip_geral,
            )
            summary = summarize_records(records)
            df = pd.DataFrame([r.__dict__ for r in records])

            st.session_state['resultado_processamento'] = {
                'zip_sep': zip_sep,
                'zip_all': zip_all,
                'summary': summary,
                'df': df,
                'planilha_ok': planilha_ok,
                'mostrar_previa': mostrar_previa,
            }

    resultado = st.session_state.get('resultado_processamento')
    if resultado:
        summary = resultado['summary']
        df = resultado['df']
        planilha_ok = resultado['planilha_ok']
        mostrar_previa = resultado['mostrar_previa']

        if planilha_ok:
            st.markdown('<div class="success-box"><strong>Modo usado:</strong> PDF + planilha. Contracheques de demitidos (situação 8) também são renomeados normalmente.</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="success-box"><strong>Modo usado:</strong> somente PDF. O CPF vem direto do contracheque.</div>', unsafe_allow_html=True)

        if resultado['zip_all'] is not None and resultado['zip_sep'] is not None:
            st.markdown('<div class="warn-box"><strong>Dica de desempenho:</strong> para lotes muito grandes, use primeiro só o ZIP separado por empresa. O ZIP geral costuma ser a segunda saída mais pesada.</div>', unsafe_allow_html=True)

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric('Arquivos lidos', summary['total'])
        m2.metric('Incluídos', summary['incluidos'])
        m3.metric('Ignorados', summary['ignorados'])
        m4.metric('Demitidos renomeados', summary.get('demitidos_renomeados', 0))
        m5.metric('Período', summary['periodo'])

        st.subheader('Resumo por empresa')
        if summary['empresas']:
            empresa_df = pd.DataFrame(list(summary['empresas'].items()), columns=['Empresa', 'Quantidade'])
            st.dataframe(empresa_df, use_container_width=True, hide_index=True)
        else:
            st.warning('Nenhum arquivo válido incluído.')

        col_a, col_b = st.columns(2)
        with col_a:
            st.subheader('Origem do CPF')
            origem_df = pd.DataFrame(list(summary['cpf_por_origem'].items()), columns=['Origem', 'Quantidade'])
            st.dataframe(origem_df, use_container_width=True, hide_index=True)
        with col_b:
            st.subheader('ATIVA por posto')
            if summary['ativa_subgrupos']:
                sub_df = pd.DataFrame(list(summary['ativa_subgrupos'].items()), columns=['Posto', 'Quantidade'])
                st.dataframe(sub_df, use_container_width=True, hide_index=True)
            else:
                st.info('Nenhum arquivo da ATIVA encontrado neste lote.')

        period_label = summary['periodo']
        botoes = st.columns(3)
        if resultado['zip_sep'] is not None:
            with botoes[0]:
                st.download_button(
                    'Baixar ZIP separado por empresa',
                    data=resultado['zip_sep'],
                    file_name=f'CONTRACHEQUES_SEPARADOS_{period_label}.zip',
                    mime='application/zip',
                    use_container_width=True,
                    on_click='ignore',
                )
        if resultado['zip_all'] is not None:
            with botoes[1]:
                st.download_button(
                    'Baixar ZIP geral',
                    data=resultado['zip_all'],
                    file_name=f'CONTRACHEQUES_GERAL_{period_label}.zip',
                    mime='application/zip',
                    use_container_width=True,
                    on_click='ignore',
                )
        with botoes[2]:
            csv_bytes = df.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                'Baixar relatório CSV',
                data=csv_bytes,
                file_name=f'RELATORIO_{period_label}.csv',
                mime='text/csv',
                use_container_width=True,
                on_click='ignore',
            )

        if mostrar_previa:
            st.subheader('Prévia do processamento')
            st.dataframe(df, use_container_width=True, hide_index=True)
        else:
            st.info('Prévia detalhada ocultada para economizar memória e deixar a página mais leve.')

with aba_regras:
    st.markdown(
        """
        ### O que mudou nesta versão
        - A **empresa** de cada PDF agora é identificada pelo **CNPJ dentro do arquivo**, não pelo nome do arquivo.
          Isso evita colocar contracheques de uma empresa na pasta de outra quando o código do empregado se repete.
        - **Situação 8 (Demitido)** deixou de ser ignorada. Esses contracheques são **renomeados normalmente** e vão para as pastas das suas empresas.
        - O lookup na planilha usa a chave composta **(codi_emp, código)**, corrigindo o CPF que podia sair errado antes.
        - Detecção de **PMV** e **SEDU** ficou mais robusta.

        ### Regras do app
        - **Sem planilha:** tenta encontrar **CPF**, **empresa** e **mês/ano** dentro do próprio PDF.
        - **Com planilha:** usa a planilha para pegar o CPF correto e apresentar a situação de cada colaborador.
        - **ATIVA:** organiza em subpastas **PMV**, **SEDU** e **GERAL**.
        - **Nome final:** `MM-AAAA-CPF.pdf`

        ### CNPJs configurados
        - **ATIVA:** 02.201.230/0001-44
        - **LIDER COMERCIAL:** 03.659.631/0001-05
        - LIDER MULTISSERVICOS e VSP identificadas pelo nome no PDF. Basta acrescentar os CNPJs no `processor.py` para ganhar ainda mais precisão.
        """
    )
