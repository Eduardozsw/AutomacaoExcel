import streamlit as st
import pandas as pd
import io



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
        "Escolha a coluna para pesquisa", ["PACIENTE", "PROFISIONAL SOLICITANTE", "SUS"]
    )

    tabelaExcel = st.file_uploader("Escolha um arquivo XLSX", type="xlsx")
    if not tabelaExcel == None:
        dfs = pd.read_excel(tabelaExcel, sheet_name=None, header=0)
        xls = pd.ExcelFile(tabelaExcel)
        abas = xls.sheet_names

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
            "P. FEZES",
            "S. URINA",
            "UROCULTURA + ANTIBIOG.",
            "HEMOGRAMA COMPLETO",
            "HEMOGLOBINA GLICADA (HB1AC)",
            "TOTG-CURVA GLICEMICA",
            "GLICEMIA JEJUM",
            "COLESTEROL TOTAL",
            "COLESTEROL-HDL",
            "COLESTEROL-LDL",
            "TRIGLICERÍDEOS",
            "UREIA",
            "CREATININA",
            "TRANSAMINASE - TGO",
            "TRANSAMINASE - TGP",
            "BILIRRUBINA T. F.",
            "ÁCIDO ÚRICO",
            "VDRL",
            "TOXOPLASMOSE IGG/IGM",
            "CITOMEGALOVIRUS IGG/IGM",
            "TAP - INR",
            "TEMPO COAGULACAO-TC",
            "TEMPO SANGRAMENTO - TS",
            "TTPA",
            "TEMPO TROMBINA",
            "FERRITINA",
            "FERRO SÉRICO",
            "TRANSFERRINA",
            "CAP. FIXAÇÃO DO FERRO",
            "ABO - FATOR RH",
            "AC. FOLICO-FOLATO",
            "ALBUMINA",
            "ALFA FETOPROTEINA",
            "ANTICORPOS ANTINUCLEO(FAN)",
            "AMILASE",
            "ASLO",
            "BHCG",
            "CA 125",
            "C3",
            "C4",
            "CEA",
            "CH50",
            "CALCIO",
            "CLORETO",
            "COOMBS INDIRETO",
            "CPK",
            "CPK-MB",
            "DHL - DESIDROGENASE LÁTICA",
            "ELETROFORESE DE HEMOGLOBINA",
            "ELETROFORESE DE PROTEINA",
            "EPSTEIN BARR IGG/IGM",
            "ERITROGRAMA",
            "HEMOGLOBINA",
            "FATOR REUMATOIDE - LATEX",
            "FIBRINOGENIO",
            "FOSFATASE ALCALINA",
            "FOSFATO SÉRICO",
            "FOSFORO",
            "FTA -ABS IGG IGM",
            "GAMA GT",
            "HEMATOCRITO",
            "HERPES SIMPLES IGG/IGM",
            "HIV (ELISA)",
            "HTLV",
            "HOMOCISTEINA NA URINA",
            "INSULINA",
            "IMUNOGLOBULINA E (IGE)",
            "IONOGRAMA",
            "LIPASE",
            "MAGNESIO",
            "MICROALBUMINA NA URINA",
            "MUCOPROTEINA",
            "PCR",
            "PLAQUETAS",
            "POTASSIO",
            "PROTEINAS TOTAIS",
            "PROTEINAS TOTAIS E FRAÇÕES",
            "PROTEINURIA-URINA 24 HORAS",
            "PSA T.L.",
            "RETICULOCITOS",
            "RUBEOLA IGG/IGM",
            "SANGUE OCULTO NAS FEZES",
            "SODIO",
            "VHS",
            "VITAMINA B12",
            "VITAMINA D",
            "HEPATITE A (HVA-IGG/IGM)",
            "HEPATITE B(HBSAG)",
            "HEPATITE B (ANTI-HBS)",
            "HEPATITE B (ANTI-HBC-IGM)",
            "HEPATITE B(HBEAG)",
            "HEPATITE C (ANTI-HCV)",
            "ANTITIREOGLOBULINA",
            "ANTI-TPO (ANTIMICROSSOMAS)",
            "ANTIGLOBULINA HUMANA (TAD)",
            "ANTIGLOBULINA HUMANA (TIA)",
            "ANDROSTENEDIONA",
            "ESTRADIOL",
            "ESTROGENIO",
            "FSH",
            "LH",
            "PROGESTERONA",
            "PROLACTINA",
            "PARATORMONIO (PTH)",
            "T3",
            "T4",
            "T4 LIVRE",
            "TESTOTERONA",
            "TESTOTERONA LIVRE",
            "CORTISOL",
            "TSH",
            "ESTRONA",
            "DHEAS",
            "DHAEA",
            "TROPONINA",
        ],
    )
    if abas is not None:
        planilha = st.pills("Qual planilha você deseja marcar", options=abas)

    st.write("Exames selecionados:", options)

    if st.button("Confirmar marcação"):
        new_row = {
            "Paciente": paciente,
            "Data Nascimento": data,
            "Nº TELEFONE": telefone,
            "PROFISSIONAL SOLICITANTE": psolicitante,
            "CPF": cpf,
            "SUS": sus,
            "COLETA DOMICILIAR": cdomiciliar,
            "CONSELHO": conselho,
            "ENDERECO": endereco,
        }
        for exame in options:
            new_row[exame] = 1

        dfe = pd.read_excel(tabelaExcel, sheet_name=planilha)

        dfe = pd.concat(
            [dfe, pd.DataFrame([new_row])], ignore_index=True
        )

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # regrava todas as abas, incluindo a atualizada
            for nome, aba_df in dfs.items():
                if nome == planilha:
                    aba_df = dfe  # usa a planilha atualizada
                aba_df.to_excel(writer, sheet_name=nome, index=False)

        st.success("Exame marcado com sucesso")
        st.download_button("Baixar nova planilha", data=output.getvalue(), file_name="planilha_atualizada.xlsx")


        st.success("Exame marcado com sucesso")
        st.dataframe(dfe)
