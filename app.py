import streamlit as st
import pandas as pd
from openpyxl import load_workbook

caminho_arquivo = "output/novos_dados.xlsx"


st.set_page_config(page_title="Secretaria de Saúde", layout="wide")

st.title("Regulação - Exames laboratoriais")
st.write("Olá, seja bem-vindo(a) ao sistema de agendamento de exames!")

nova_aba = pd.DataFrame(
    {
        "PACIENTE": [],
        "Ácido úrico": [],
        "Bilirrubina Total e Frações": [],
        "Colesterol total": [],
        "Colesterol HDL": [],
        "Colesterol LDL": [],
        "Creatinina": [],
        "Glicose - glicemia": [],
        "Insulina": [],
        "TGO -AST - Transaminase Oxalacetica": [],
        "TGP - ALT - Transaminase pirúvica": [],
        "Triglicerídeos": [],
        "Uréia": [],
        "ABO - Fator RH": [],
        "Hemograma completo": [],
        "Plaquetas": [],
        "TAP - Tempo de protrombina": [],
        "TS - Tempo de sangramento - DUKE": [],
        "ASLO - Antiestreptolisina": [],
        "Fator reumatóide  - FR - Látex": [],
        "PCR - Proteína C Reativa dosagem": [],
        "FTA - ABS IGG p Sífilis": [],
        "FTA - ABS IGM p Sífilis": [],
        "VDRL quantitativo - VDRL com titulação": [],
        "Hepatite A (HAV  IGM)": [],
        "Hepatite A (HAV  IGG)": [],
        "Hepatite B  - HBSAG": [],
        "Hepatite B - Anti HBS": [],
        "Hepatite B - Anti HBC IGM": [],
        "Hepatite B - HBEAG": [],
        "Hepatite C - Anti-HCV": [],
        "HIV 1 + HIV 2 (ELISA)": [],
        "Herpes simples IGG": [],
        "Herpes Simples IGM": [],
        "Folato - Ácido fólico - Vitamina B9": [],
        "Imunoglobulina E - IGE": [],
        "Beta-HCG - Gonadotrofina coriônica humana": [],
        "Parasitológico de Fezes - Pesquisa de ovos e cistos de parasitas - EPF": [],
        "Sumário de urina - EAS- Urina tipo 1": [],
        "Homocisteína na urina": [],
        "Urocultura antibiograma- Cultura de bactérias para identificação": [],
        "Cultura para BAAR - Micobactéria - Tuberculose": [],
        "Sódio - Na": [],
        "Potássio - k": [],
        "Cloreto - Cloro sérico": [],
        "Fosfatase alcalina - FA - FAL": [],
        "Proteínas Totais e Frações": [],
        "Alfa-fetoproteína": [],
        "Antitireoglobulina": [],
        "antimicrossomas": [],
        "T3 total (Triodotironina)": [],
        "T4 total - Tiroxina": [],
        "T4 livre - Tiroxina livre": [],
        "TSH - Hormônio tireoestimulante": [],
        "LH - Hormônio Luteinizante": [],
        "FSH - Hormônio Folículo estimulante": [],
        "GAMA GT - Gama glutamil transferase": [],
        "Progesterona": [],
        "PSA - Antígeno prostático específico": [],
        "Mucoproteínnas - GPA - Alfa 1 glicoproteína ácida": [],
        "Amilase": [],
        "Lipase": [],
        "Hba1C - Hemoglobina glicosilada - A1c": [],
        "Vitamina B12 - Cobalamina": [],
        "Vitamina D - 25 hidroxi D - 25OHD": [],
        "CEA - Antígeno Carcinoembrionário": [],
        "C3": [],
        "C4": [],
        "Cálcio": [],
        "Cortisol": [],
        "Capacidade de Fixação de Ferro": [],
        "TOTG - Curva Glicêmica": [],
        "CH50 - Complemento total": [],
        "Citomegalovírus IGM": [],
        "Citomegalovírus IGG": [],
        "Coombs Direto -TAD - Teste Direto de Antiglobulina Humana": [],
        "Coombs Indireto - TIA - Teste indireto de Antiglobulina Humana": [],
        "Epstein Barr IGG": [],
        "Epstein Barr IGM": [],
        "Eletroforese de hemoglobina": [],
        "Eletroforese de proteínas": [],
        "Estrona - E1": [],
        "Estradiol - Estrogênio": [],
        "FAN - Fator antinúcleo - Anticorpos antinúcleo": [],
        "Ferro sérico": [],
        "Ferritina": [],
        "Fibrinogênio": [],
        "Fósforo": [],
        "Desidrogenase lática - DHL - LDH": [],
        "Magnésio": [],
        "Microalbumina na urina": [],
        "Prolactina": [],
        "Paratormônio - PTH": [],
        "Reticulócitos": [],
        "Rubéola IGG": [],
        "Rubéola IGM": [],
        "Sangue oculto nas fezes": [],
        "Testosterona Total": [],
        "Testosterona livre": [],
        "Toxoplasmose IGM": [],
        "Toxoplasmose IGG": [],
        "Transferrina": [],
        "Tempo de tromboplastina p ativada - TTPA": [],
        "VHS - Velocidade de Hemossedimentação": [],
        "HTLV 1 + HTLV 2 (ELISA)": [],
        "HTLV 1 + HTLV 2 (WESTERN BLOT)": [],
        "Tempo de trombina": [],
        "G6PD - Glicose6 fostato desidrogenase": [],
        "Eritrograma (eritrócitos, hemoglobina, hematócritos)": [],
        "Hemoglobina": [],
        "Troponina": [],
        "Hematócrito": [],
        "PCR quantitativa - Proteína C Reativa quantitativa": [],
        "CPK - Creatinofosfoquinase": [],
        "CPK-MB": [],
        "Ionograma": [],
        "CPF": [],
        "SUS": [],
        "DATA DE NASCIMENTO": [],
        "TELEFONE": [],
        "POSTO": [],
        "PROFISSIONAL SOLICITANTE": [],
        "COLETA DOMICILIAR": [],
    }
)

tab1, tab2 = st.tabs(["Marcação", "Conferir Marcação"])

if "df" not in st.session_state:
    st.session_state.df = pd.DataFrame()

#tela de pesquisa na planilha
with tab2:
    termo = st.text_input("Pesquisar")
    coluna_escolhida = st.selectbox(
        "Escolha a coluna para pesquisa", ["PACIENTE", "PROFISIONAL SOLICITANTE", "SUS"]
    )

    tabelaExcel = st.file_uploader("Escolha um arquivo XLSX", type="xlsx")
    if not tabelaExcel == None:
        dfs = pd.read_excel(tabelaExcel, sheet_name=None, header=0)

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

#tela de marcação
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
    with col3:
        criar = st.checkbox("Criar uma nova planilha")
        if criar:
            titulo = st.text_input("Qual o nome da nova planilha")

            
            # 2) Usa ExcelWriter em modo append
            with pd.ExcelWriter("Pasta 1.xlsx", engine="openpyxl", mode="a") as writer:
                # 3) Grava como aba “Estoque” sem tocar nas outras
                nova_aba.to_excel(writer, sheet_name=titulo, index=False)
                # Se “Estoque” já existir, dará erro – basta escolher outro nome.

            
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

    st.write("Exames selecionados:", options)

    if st.button("Confirmar marcação"):
        new_row = {
            "PACIENTE": paciente,
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

        st.session_state.df = pd.concat(
            [st.session_state.df, pd.DataFrame([new_row])], ignore_index=True
        )
        st.success("Exame marcado com sucesso")
        st.dataframe(st.session_state.df)
