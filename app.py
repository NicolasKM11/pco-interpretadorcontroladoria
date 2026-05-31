import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook

st.set_page_config(page_title="Interpretador PCO", page_icon="📊", layout="wide")

CONTAS_PADRAO = [
    "Empréstimos", "Desp. Financeiras", "Faturamento", "CPV", "Despesas Operacionais",
    "Lucro Operacional", "Lucro Líquido", "CDG", "NCG", "Tesouraria", "Contas a Receber",
    "Estoque MP", "Estoque PA", "Ativo Operacional", "Fornecedores", "Tributos", "Outros", "Passivo Operacional"
]

SINONIMOS = {
    "Empréstimos": ["emprést", "emprest", "emp."],
    "Desp. Financeiras": ["desp. fin", "despesas financeiras", "d.fin"],
    "Faturamento": ["faturamento", "fat"],
    "CPV": ["c p v", "cpv"],
    "Despesas Operacionais": ["despesas", "desp oper", "despesas operacionais"],
    "Lucro Operacional": ["lucro op", "lo"],
    "Lucro Líquido": ["lucro liq", "lucro líq", "ll"],
    "CDG": ["cdg"],
    "NCG": ["ncg"],
    "Tesouraria": ["tesouraria", "tes"],
    "Contas a Receber": ["contas a receber", "cr"],
    "Estoque MP": ["estoque mp", "est. mp"],
    "Estoque PA": ["estoque pa", "est. pa", "produto acabado"],
    "Ativo Operacional": ["ativo op", "ativo operacional", "ao"],
    "Fornecedores": ["fornecedor", "forn"],
    "Tributos": ["tribut"],
    "Outros": ["outros"],
    "Passivo Operacional": ["passivo op", "passivo operacional", "po"],
}


def normaliza(x):
    return re.sub(r"\s+", " ", str(x).strip().lower())


def to_float(x):
    if x is None or x == "":
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace("R$", "").replace("%", "").strip()
    s = s.replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None


def detectar_conta(texto):
    t = normaliza(texto)
    if not t:
        return None
    for conta, termos in SINONIMOS.items():
        if any(term in t for term in termos):
            return conta
    return None


def ler_planilha(uploaded):
    data = uploaded.read()
    wb = load_workbook(BytesIO(data), data_only=True, read_only=True)
    candidatos = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            vals = [c.value for c in row]
            for i, v in enumerate(vals):
                conta = detectar_conta(v)
                if conta:
                    nums = []
                    for j in range(i + 1, min(i + 9, len(vals))):
                        n = to_float(vals[j])
                        if n is not None:
                            nums.append(n)
                    if len(nums) >= 2:
                        original, novo = nums[0], nums[1]
                        candidatos.append({
                            "Conta": conta,
                            "Original": original,
                            "Novo Valor": novo,
                            "Variação": novo - original,
                            "%": ((novo - original) / abs(original) * 100) if original else 0,
                            "Aba": ws.title,
                            "Linha": row[0].row,
                        })
    if not candidatos:
        return pd.DataFrame(columns=["Conta", "Original", "Novo Valor", "Variação", "%"])
    df = pd.DataFrame(candidatos)
    df = df.drop_duplicates(subset=["Conta"], keep="last")
    ordem = {c: i for i, c in enumerate(CONTAS_PADRAO)}
    df["ordem"] = df["Conta"].map(ordem).fillna(999)
    return df.sort_values("ordem").drop(columns="ordem")


def sinal(v):
    if pd.isna(v) or abs(v) < 0.01:
        return "não teve variação relevante"
    return "aumentou" if v > 0 else "diminuiu"


def val(df, conta, campo="Variação"):
    linha = df[df["Conta"] == conta]
    if linha.empty:
        return 0
    return float(linha.iloc[0][campo])


def frase_moeda(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def gerar_respostas(df, premissa):
    fat, cpv, desp = val(df, "Faturamento"), val(df, "CPV"), val(df, "Despesas Operacionais")
    emp, dfin = val(df, "Empréstimos"), val(df, "Desp. Financeiras")
    cr, forn = val(df, "Contas a Receber"), val(df, "Fornecedores")
    ao, po = val(df, "Ativo Operacional"), val(df, "Passivo Operacional")
    ncg, cdg, tes = val(df, "NCG"), val(df, "CDG"), val(df, "Tesouraria")
    lo, ll = val(df, "Lucro Operacional"), val(df, "Lucro Líquido")
    estpa, estmp = val(df, "Estoque PA"), val(df, "Estoque MP")

    respostas = []

    if premissa == "Quantidade vendida":
        respostas.append(("Resultado Contábil — por que o Δ Faturamento é superior ao Δ CPV?",
            "O faturamento aumentou porque a empresa vendeu mais unidades. O CPV também tende a aumentar, mas em proporção menor, porque parte da estrutura de custos é fixa e passa a ser diluída em um volume maior de produção. Por isso, o crescimento do faturamento supera o crescimento proporcional do CPV, melhorando o resultado operacional."))
        respostas.append(("Resultado Patrimonial — por que aumentou Fornecedores?",
            "A conta fornecedores aumentou porque, para sustentar o maior volume vendido, a empresa precisou produzir e comprar mais matéria-prima. Como parte dessas compras é feita a prazo, o saldo de fornecedores no passivo operacional aumenta."))
    elif premissa == "Preço de venda":
        respostas.append(("Resultado Patrimonial — por que aumentou Contas a Receber?",
            "O contas a receber aumentou porque o preço de venda maior elevou o faturamento. Como parte das vendas ocorre a prazo, o valor financeiro ainda não recebido também aumenta."))
        respostas.append(("Resultado Contábil — explique as Despesas Operacionais",
            "As despesas operacionais variaram principalmente porque algumas despesas, como comissões, marketing/publicidade e outras despesas variáveis, acompanham o faturamento. Como o preço elevou a receita, essas contas também podem aumentar, mesmo sem aumento de quantidade produzida."))
        respostas.append(("Resultado Contábil — por que o CPV fica igual ou quase igual?",
            "Quando a premissa altera apenas preço, e não quantidade, a produção e o consumo de matéria-prima permanecem praticamente iguais. Por isso, o CPV tende a não mudar de forma significativa."))
    elif premissa == "Produção / Estoque PA":
        respostas.append(("Resultado Contábil — por que o impacto pode não ser significativo?",
            "O impacto contábil pode ser pequeno porque a premissa mexe principalmente na forma de produção e no nível de estoque, sem alterar diretamente o faturamento. Assim, o efeito aparece mais no custo unitário e no capital de giro do que na receita."))
        respostas.append(("Resultado Contábil — por que a readequação do estoque PA pode aumentar o CPV?",
            "Ao reduzir ou readequar o estoque de produto acabado, a empresa pode produzir menos em alguns períodos. Com menor produção, os custos fixos são distribuídos em menos unidades, aumentando o custo unitário e, consequentemente, o CPV."))
        respostas.append(("Resultado Patrimonial — por que a NCG pode reduzir?",
            "A NCG pode reduzir porque a diminuição dos estoques reduz o ativo operacional. Quando o ativo operacional cai mais do que o passivo operacional, a necessidade de capital de giro diminui."))
    elif premissa == "Compra / Estoque MP":
        respostas.append(("Resultado Contábil — por que o impacto pode ser positivo?",
            "A readequação do estoque de matéria-prima pode reduzir compras, armazenagem e necessidade de financiamento. Isso melhora o resultado contábil principalmente quando também reduz despesas financeiras ou custo de produção."))
        respostas.append(("Resultado Patrimonial — por que reduziu Fornecedores?",
            "A conta fornecedores reduziu porque a empresa passou a comprar menos matéria-prima ou reduziu o nível de estoque de MP. Comprando menos a prazo, o saldo de obrigações com fornecedores também diminui."))
    elif premissa == "Preço dos insumos":
        respostas.append(("Resultado Contábil — por que reduziu o CPV?",
            "O CPV reduziu porque o preço de aquisição dos insumos caiu. Como matéria-prima é componente relevante do custo de produção, a redução no custo dos insumos diminui o custo total dos produtos vendidos."))
        respostas.append(("Resultado Financeiro — o que contribuiu para reduzir empréstimos?",
            "A redução do preço dos insumos diminui os desembolsos com compras e melhora a geração de caixa. Com menor necessidade de recursos para financiar a operação, a empresa depende menos de empréstimos."))
        respostas.append(("Resultado Patrimonial — por que reduziu estoque de PA?",
            "O estoque de produto acabado pode reduzir em valor porque os produtos passam a carregar um custo unitário menor. Mesmo com quantidade parecida, o valor contábil do estoque diminui quando o custo de produção cai."))
    else:
        respostas.append(("Interpretação geral",
            "A análise deve observar o efeito dominó da premissa sobre resultado financeiro, resultado contábil e resultado patrimonial. Primeiro avalie faturamento, CPV e despesas; depois empréstimos e despesas financeiras; por fim, contas a receber, estoques, fornecedores, CDG, NCG e tesouraria."))

    # diagnóstico complementar com base nos números encontrados
    respostas.append(("Diagnóstico automático pelos valores encontrados",
        f"Na planilha, o faturamento {sinal(fat)} ({frase_moeda(fat)}), o CPV {sinal(cpv)} ({frase_moeda(cpv)}) e as despesas operacionais {sinal(desp)} ({frase_moeda(desp)}). "
        f"No financeiro, os empréstimos {sinal(emp)} ({frase_moeda(emp)}) e as despesas financeiras {sinal(dfin)} ({frase_moeda(dfin)}). "
        f"No patrimonial, contas a receber {sinal(cr)}, fornecedores {sinal(forn)}, NCG {sinal(ncg)}, CDG {sinal(cdg)} e tesouraria {sinal(tes)}. "
        f"A leitura final deve comparar se a melhora no CDG foi suficiente para compensar a variação da NCG."))
    return respostas


st.title("📊 Interpretador PCO — RF, RC e RP/NCG")
st.write("Envie a planilha da simulação ou preencha manualmente os valores original e novo valor. O app gera explicações no padrão das questões da FAE.")

with st.sidebar:
    st.header("Configuração")
    premissa = st.selectbox("Tipo de premissa", [
        "Quantidade vendida", "Preço de venda", "Produção / Estoque PA", "Compra / Estoque MP", "Preço dos insumos", "Outra"
    ])
    modo = st.radio("Modo", ["Upload da planilha", "Manual"])

if modo == "Upload da planilha":
    arq = st.file_uploader("Enviar arquivo Excel", type=["xlsx", "xlsm"])
    if arq:
        df = ler_planilha(arq)
        if df.empty:
            st.warning("Não consegui detectar automaticamente as contas. Use o modo manual ou ajuste a tabela abaixo.")
            df = pd.DataFrame({"Conta": CONTAS_PADRAO, "Original": 0.0, "Novo Valor": 0.0, "Variação": 0.0, "%": 0.0})
        st.subheader("Valores detectados/ajustáveis")
        edit = st.data_editor(df[["Conta", "Original", "Novo Valor", "Variação", "%"]], use_container_width=True, num_rows="dynamic")
    else:
        st.info("Envie a planilha para começar.")
        edit = None
else:
    df = pd.DataFrame({"Conta": CONTAS_PADRAO, "Original": 0.0, "Novo Valor": 0.0})
    edit = st.data_editor(df, use_container_width=True, num_rows="dynamic")
    edit["Variação"] = edit["Novo Valor"] - edit["Original"]
    edit["%"] = edit.apply(lambda r: ((r["Variação"] / abs(r["Original"])) * 100) if r["Original"] else 0, axis=1)

if edit is not None:
    st.subheader("Respostas automáticas")
    for titulo, texto in gerar_respostas(edit, premissa):
        with st.expander(titulo, expanded=True):
            st.write(texto)

    st.subheader("Texto corrido para colar")
    texto_final = "\n\n".join([f"{t}\n{txt}" for t, txt in gerar_respostas(edit, premissa)])
    st.text_area("Resposta pronta", texto_final, height=320)
