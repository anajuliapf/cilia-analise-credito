"""
Funil de Crédito Cilia — qualificação de oficinas para oferta de crédito.

Lê a base de orçamentos extraída do Metabase e a análise de bureau (score),
aplica filtros "duros" e "leves" a cada orçamento, classifica com uma FLAG,
agrega o funil por oficina (CNPJ) e cruza com o score de crédito.

Entradas (não versionadas — dados confidenciais):
    data/cilia_creditoANALISE.xlsx  -> orçamentos (Metabase)
    data/ANALISEBureau.xlsx         -> score de bureau por CNPJ (SCORE_HPJ9)

Saídas:
    sumariofunil.xlsx, funilsegttl.xlsx, funilokZurichSura.xlsx,
    linhasokttl.xlsx, linhasokZurichSuraAll.xlsx, contador.xlsx
"""

import math

import pandas as pd
from IPython.display import display

# ------------------------------------------------------------------
# 1. Carga dos dados
# ------------------------------------------------------------------
orc = pd.read_excel("data/cilia_creditoANALISE.xlsx")
nova_planilha = pd.read_excel("data/ANALISEBureau.xlsx")

pd.options.display.float_format = "{:,.2f}".format
pd.options.display.max_columns = None
pd.options.display.max_rows = None

# ------------------------------------------------------------------
# 2. Valor líquido do orçamento
# ------------------------------------------------------------------
orc["ValorLiquido"] = (
    orc["total_budget_raw"]
    - orc["total_budget_selling_cost_by_insurer (R$)"]
    - orc["total_budget_franchise"]
)
orc["ValorLiquido_Maior_20000"] = orc["ValorLiquido"] >= 20000


# ------------------------------------------------------------------
# 3. Regras de classificação (FLAG)
# ------------------------------------------------------------------
def conditions(orc):
    # ===================== FILTROS DUROS =====================
    if orc["workshop_document_identifier"] in (22222222222222, "22222222222222"):
        return "00_Oficina Teste"
    elif orc["workshop_trade"] in {
        "SEM OFICINA - CILIA", "2J COMERCIO DE PECAS E SERVICOS teste", "DAniel oficina teste",
        "DARWIN TESTES", "Empresa Teste", "freio teste ", "Mitsui - Teste ORC", "MOTO TESTE",
        "Oficina de Testes", "Oficina Para Testes", "TESTE CILIA", "Oficina teste", "Oficina Teste",
        "Oficina Teste HDI 1", "Oficina Teste UON", "TESTE", "Teste Cilia",
    }:
        return "00_Oficina Teste"
    elif orc["insurer_trade"] in {
        "Auto/RE - TESTE", "Cilia Demonstração (Seguradora)", "Seguradora Demonstração",
        "Seguradora Fake", "Seguradora Treinamento", "Seguradora Cilia Pass", "Sistema Cilia",
        "Seguradora Treinamento (Whatsapp)", "Cilia Demonstração IA",
    }:
        return "00_Seguradora Teste"
    elif orc["i18n_process_type"] not in ("Segurado", "Terceiro"):
        return "01_Orçamento não Segurado"
    elif orc["total_budget"] <= 0 or orc["ValorLiquido"] <= 0:
        return "02_ValorInadequado"
    elif orc["conclusion_type_title"] in {
        "Não Autorizado", "Indenização Integral", "Perda Total Validação", "Constatação de Danos / Local",
        "Provável Indenização Integral", "Não Realizada", "Não realizada", "Cancelado", "Acordo",
        "Não realizado", "Perda Total", "Reparo Não Autorizado", "Vistoria não Realizada",
        "Não Autorizado Negado", "Abaixo da Franquia", "Constatação Realizada", "Desautorizar Reparos",
        "Provável Perda Total", "Veículo Ausente", "Vistoria Cancelada", "Vistoria Frustrada",
        "Vistoria Não Realizada", "Vistoria sem Acordo", "VIDE PORTAL DA CIA", "Vistoria Realizada",
        "Vistoria não realizada", "Não Atinge Franquia", "Não Houve Acordo", "Não houve acordo",
    }:
        return "03_Cancelado ou II"
    elif orc["qtde_nfs_servicos"] > 0 or orc["qtde_nfs_pecas"] > 0:
        return "04_Orçamento Concluído - NF"
    elif orc["budget_version_name"] in {"supervision", "verification"}:
        return "04_Orçamento Concluído"
    elif orc["budget_version_number"] == 1:
        return "05_Não autorizado pela seguradora"
    elif orc["workshop_trade"] == "AUTO VISTORIA":
        return "06_VerificacaoLiberty"

    # ===================== FILTROS LEVES =====================
    elif orc["vehicle_type_name"] not in {"Passeio", "Motos", "Furgão / Vans", "Pick-up Leve", "Pick-up Pesada"}:
        return "07_Veículo Indesejado"
    elif orc["vehicle_brand_name"] not in {
        "CHEVROLET", "WOLKSWAGEN", "VOLKSWAGEN", "FIAT", "HONDA", "TOYOTA", "HYUNDAI", "KAWASAKI",
        "FORD", "NISSAN", "RENAULT", "JEEP", "YAMAHA", "CITROEN", "PEUGEOT", "MITSUBISHI",
        "LAND ROVER", "BMW", "MERCEDES-BENZ", "AUDI",
    }:
        return "08_FabricanteProblematica"
    elif orc["insurer_trade"] not in {
        "Allianz Seguros", "Tokio Marine", "Liberty Seguros", "HDI Seguros", "Bradesco Seguros", "Zurich",
        "Sompo Auto", "SURA S/A", "Alfa Seguros", "MITSUI SUMITOMO SEGUROS", "YOUSE", "Essor Seguros",
        "BANESTES SEGUROS S/A", "SUHAI SEGURADORA", "Yelum Seguradora",
    }:
        return "09_SeguradoraMenor"
    elif orc["status"] != "analyzed":
        return "10_StatusVolatil"

    # ===================== OUTPUT GERAL =====================
    elif orc["qtd_socios"] >= 2 or math.isnan(orc["qtd_socios"]):
        return "11_Mais de 1 Sócio_Adm"
    else:
        return "99_OK"


FLAGS_QUALIFICADAS = ["11_Mais de 1 Sócio_Adm", "99_OK"]

orc["FLAG"] = orc.apply(conditions, axis=1)
orc["QUALIFICADO"] = orc["FLAG"].isin(FLAGS_QUALIFICADAS)

# ------------------------------------------------------------------
# 4. Funil agregado por oficina + score de bureau
# ------------------------------------------------------------------
AGG = {
    "budget_id": "count",
    "total_budget": "sum",
    "total_budget_selling_cost_by_insurer (R$)": "sum",
    "total_budget_selling_cost_by_workshop (R$)": "sum",
    "total_budget_franchise": "sum",
    "ValorLiquido": "sum",
}
RENAME = {
    "budget_id": "Qtde",
    "total_budget": "ValorTTL",
    "total_budget_selling_cost_by_insurer (R$)": "ValorPecasSeguradora",
    "total_budget_selling_cost_by_workshop (R$)": "ValorPecasOficina",
    "total_budget_franchise": "ValorFranquia",
}


def color_name(score):
    """Faixa de risco a partir do score de bureau."""
    if pd.isna(score) or score < 300:
        return "Vermelho"
    elif score < 400:
        return "Laranja"
    elif score < 500:
        return "Amarelo"
    elif score < 600:
        return "Azul"
    return "Verde"


def enriquecer(df):
    """Adiciona QUALIFICADO, flag de ValorLiquido >= 20 mil e score de bureau."""
    df["QUALIFICADO"] = df["FLAG"].isin(FLAGS_QUALIFICADAS)
    df = df[["FLAG", "QUALIFICADO"] + [c for c in df.columns if c not in ("FLAG", "QUALIFICADO")]]
    df["ValorLiquido_Maior_20000"] = df["ValorLiquido"] >= 20000
    df = df.merge(
        nova_planilha[["CNPJ", "SCORE_HPJ9"]],
        left_on="workshop_document_identifier",
        right_on="CNPJ",
        how="left",
    ).drop(columns=["CNPJ"])
    df["Cor_SCORE"] = df["SCORE_HPJ9"].apply(color_name)
    return df


funil = (
    orc.groupby(["FLAG", "workshop_document_identifier"])
    .agg(AGG).rename(columns=RENAME).reset_index()
)
funil = enriquecer(funil)

# ------------------------------------------------------------------
# 5. Recorte por seguradoras priorizadas
# ------------------------------------------------------------------
SEGURADORAS_FOCO = ["Zurich", "SURA S/A", "Allianz Seguros"]

funilseg = (
    orc.groupby(["FLAG", "insurer_trade", "workshop_document_identifier"])
    .agg(AGG).rename(columns=RENAME).reset_index()
)
funilseg2 = enriquecer(funilseg[funilseg["insurer_trade"].isin(SEGURADORAS_FOCO)].copy())

funil.to_excel("sumariofunil.xlsx")
funilseg.to_excel("funilsegttl.xlsx")
funilseg2.to_excel("funilokZurichSura.xlsx")

# ------------------------------------------------------------------
# 6. Linhas qualificadas e contador por oficina
# ------------------------------------------------------------------
linhasokttl = orc[orc["FLAG"].isin(FLAGS_QUALIFICADAS)]
linhasokttl.to_excel("linhasokttl.xlsx")
linhasokttl[linhasokttl["insurer_trade"].isin(SEGURADORAS_FOCO)].to_excel("linhasokZurichSuraAll.xlsx")

base_contador = orc[
    orc["FLAG"].isin(FLAGS_QUALIFICADAS + ["04_Orçamento Concluído - NF", "04_Orçamento Concluído"])
    & orc["insurer_trade"].isin(SEGURADORAS_FOCO)
]
contador = (
    base_contador.groupby("workshop_document_identifier")
    .agg({"FLAG": "count", "ValorLiquido": "sum"})
    .rename(columns={"FLAG": "Qtde"})
    .reset_index()
)
contador.to_excel("contador.xlsx")

display(funil.head(20))
display(contador.head(10))
