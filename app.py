import streamlit as st
import pandas as pd
import io

if "marcacoes" not in st.session_state:
    st.session_state.marcacoes = []

st.set_page_config(page_title="Secretaria de Saúde", layout="wide")

st.title("Regulação - Exames laboratoriais")
st.write("Olá, seja bem-vindo(a) ao sistema de agendamento de exames!")


tab1, tab2 = st.tabs(["Marcação", "Conferir Marcação"])

if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()

# tela de pesquisa na planilha
with tab2:
    termo = st.text_input("Pesquisar")
    coluna_escolhida = st.selectbox(
        "Escolha a coluna para pesquisa",
        ["Paciente", "Profissional solicitante", "SUS"],
    )

    tabelaExcel = st.file_uploader("Escolha um arquivo XLSX", type="xlsx")
    if not tabelaExcel == None:
        dfs = pd.read_excel(tabelaExcel, sheet_name=None, header=0)
        xls = pd.ExcelFile(tabelaExcel)
        aba = xls.sheet_names

        for nome_planilha, df in dfs.items():
            if nome_planilha in ["Início", "Fim", "Mensal"]:
                continue

            df = df.fillna("")  # preenche antes para evitar repetição

            if termo:

                df_filtrado = df[
                    df[coluna_escolhida]
                    .astype(str)
                    .str.contains(termo, case=False, na=False)
                ]

                if not df_filtrado.empty:
                    st.subheader(f"{nome_planilha} (filtrado)")
                    st.dataframe(df_filtrado)
            else:
                st.subheader(nome_planilha)
                st.dataframe(df)

# tela de marcação
with tab1:
    # Formulário de Marcação
    _, col1, col2, col3 = st.columns([0.5, 1, 1, 0.5])
    with col1:
        paciente = st.text_input("Nome do Paciente")
        psolicitante = st.text_input("Profissional solicitante")
        posto = st.text_input("Posto")
        cpf = st.text_input("CPF")
        cdomiciliar = st.checkbox("Coleta domiciliar")
    with col2:
        data = st.text_input("Data de nascimento")
        sus = st.text_input("SUS")
        telefone = st.text_input("Telefone")
        conselho = st.text_input("Conselho")

        endereco = ""
        if cdomiciliar:
            endereco = st.text_input("Endereço")

    st.markdown("---")  # linha divisória

    options = st.multiselect(
        "Quais exames deseja marcar?",
        [
            "Ácido úrico",
            "Bilirrubina Total e Frações",
            "Colesterol total",
            "Colesterol HDL",
            "Colesterol LDL",
            "Creatinina",
            "Glicose - glicemia",
            "Insulina",
            "TGO -AST - Transaminase Oxalacetica",
            "TGP - ALT - Transaminase pirúvica",
            "Triglicerídeos",
            "Uréia",
            "ABO - Fator RH",
            "Hemograma completo",
            "Plaquetas",
            "TAP - Tempo de protrombina",
            "TS - Tempo de sangramento - DUKE",
            "ASLO - Antiestreptolisina",
            "Fator reumatóide  - FR - Látex",
            "PCR - Proteína C Reativa dosagem",
            "FTA - ABS IGG p Sífilis",
            "FTA - ABS IGM p Sífilis",
            "VDRL quantitativo - VDRL com titulação",
            "Hepatite A (HAV  IGM)",
            "Hepatite A (HAV  IGG)",
            "Hepatite B  - HBSAG",
            "Hepatite B - Anti HBS",
            "Hepatite B - Anti HBC IGM",
            "Hepatite B - HBEAG",
            "Hepatite C - Anti-HCV",
            "HIV 1 + HIV 2 (ELISA)",
            "Herpes simples IGG",
            "Herpes Simples IGM",
            "Folato - Ácido fólico - Vitamina B9",
            "Imunoglobulina E - IGE",
            "Beta-HCG - Gonadotrofina coriônica humana",
            "Parasitológico de Fezes - Pesquisa de ovos e cistos de parasitas - EPF",
            "Sumário de urina - EAS- Urina tipo 1",
            "Homocisteína na urina",
            "Urocultura antibiograma- Cultura de bactérias para identificação",
            "Cultura para BAAR - Micobactéria - Tuberculose",
            "Sódio - Na",
            "Potássio - k",
            "Cloreto - Cloro sérico",
            "Fosfatase alcalina - FA - FAL",
            "Proteínas Totais e Frações",
            "Alfa-fetoproteína",
            "Antitireoglobulina",
            "antimicrossomas",
            "T3 total (Triodotironina)",
            "T4 total - Tiroxina",
            "T4 livre - Tiroxina livre",
            "TSH - Hormônio tireoestimulante",
            "LH - Hormônio Luteinizante",
            "FSH - Hormônio Folículo estimulante",
            "GAMA GT - Gama glutamil transferase",
            "Progesterona",
            "PSA - Antígeno prostático específico",
            "Mucoproteínnas - GPA - Alfa 1 glicoproteína ácida",
            "Amilase",
            "Lipase",
            "Hba1C - Hemoglobina glicosilada - A1c",
            "Vitamina B12 - Cobalamina",
            "Vitamina D - 25 hidroxi D - 25OHD",
            "CEA - Antígeno Carcinoembrionário",
            "C3",
            "C4",
            "Cálcio",
            "Cortisol",
            "Creatinina",
            "Capacidade de Fixação de Ferro",
            "TOTG - Curva Glicêmica",
            "CH50 - Complemento total",
            "Citomegalovírus IGM",
            "Citomegalovírus IGG",
            "Coombs Direto -TAD - Teste Direto de Antiglobulina Humana",
            "Coombs Indireto - TIA - Teste indireto de Antiglobulina Humana",
            "Epstein Barr IGG",
            "Epstein Barr IGM",
            "Eletroforese de hemoglobina",
            "Eletroforese de proteínas",
            "Estrona - E1",
            "Estradiol - Estrogênio",
            "FAN - Fator antinúcleo - Anticorpos antinúcleo",
            "Ferro sérico",
            "Ferritina",
            "Fibrinogênio",
            "Fósforo",
            "Desidrogenase lática - DHL - LDH",
            "Magnésio",
            "Microalbumina na urina",
            "Prolactina",
            "Paratormônio - PTH",
            "Reticulócitos",
            "Rubéola IGG",
            "Rubéola IGM",
            "Sangue oculto nas fezes",
            "Testosterona Total",
            "Testosterona livre",
            "Toxoplasmose IGM",
            "Toxoplasmose IGG",
            "Transferrina",
            "Tempo de tromboplastina p ativada - TTPA",
            "VHS - Velocidade de Hemossedimentação",
            "HTLV 1 + HTLV 2 (ELISA)",
            "HTLV 1 + HTLV 2 (WESTERN BLOT)",
            "Imunoglobulina E - IGE",
            "Tempo de trombina",
            "G6PD - Glicose6 fostato desidrogenase",
            "Eritrograma (eritrócitos, hemoglobina, hematócritos)",
            "Hemoglobina",
            "Troponina",
            "Hematócrito",
            "PCR quantitativa - Proteína C Reativa quantitativa",
            "CPK - Creatinofosfoquinase",
            "CPK-MB",
            "Ionograma",
            "CA 125"
        ],
    )
    if tabelaExcel is not None:
        planilha = st.pills("Qual planilha você deseja marcar", options=aba)

    st.write("Exames selecionados:", options)

    if st.button("Adicionar marcação"):
        nova_marcacao = {
            "PACIENTE": paciente,
            "DATA DE NASCIMENTO": data,
            "TELEFONE": telefone,
            "Profissional solicitante": psolicitante,
            "CPF": cpf,
            "SUS": sus,
            "Coleta domiciliar": cdomiciliar,
            "Conselho": conselho,
            "ENDERECO   ": endereco,
        }
        for exame in options:
            nova_marcacao[exame] = 1

        st.session_state.marcacoes.append(nova_marcacao)
        st.success("Marcação adicionada")

        if st.session_state.marcacoes:
            st.subheader("Marcações pendentes:")
            st.dataframe(pd.DataFrame(st.session_state.marcacoes))

    if st.button("Salvar todas as marcações"):
        dfe = pd.read_excel(tabelaExcel, sheet_name=planilha, header=0)
        novas_linhas = pd.DataFrame(st.session_state.marcacoes)
        dfe = pd.concat([dfe, novas_linhas], ignore_index=True, axis=0)

        # abrir um arquivo ja existente
        output = tabelaExcel
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for nome, aba_df in dfs.items():
                if nome == planilha:
                    aba_df = dfe
                aba_df.to_excel(writer, sheet_name=nome, index=False)

        st.download_button(
            "Baixar nova planilha com todas as marcações",
            data=output.getvalue(),
            file_name="planilha_atualizada.xlsx",
        )

        st.session_state.marcacoes = []

        st.success("Exame marcado com sucesso")
        st.dataframe(dfe)
