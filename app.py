import re
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from PIL import Image

try:
    import pytesseract
except Exception:
    pytesseract = None

st.set_page_config(page_title="Interpretador PCO", page_icon="📊", layout="wide")

CONTAS_PADRAO = [
    "Empréstimos", "Desp. Financeiras", "Faturamento", "CPV", "Despesas Operacionais",
    "Lucro Operacional", "Lucro Líquido", "CDG", "NCG", "Tesouraria", "Contas a Receber",
    "Estoque MP", "Estoque PA", "Ativo Operacional", "Fornecedores", "Tributos", "Outros", "Passivo Operacional"
]

SINONIMOS = {
    "Empréstimos": ["emprést", "emprest", "emp.", "emprestimos"],
    "Desp. Financeiras": ["desp. fin", "desp fin", "despesas financeiras", "d.fin", "desp financeiras"],
    "Faturamento": ["faturamento", "fat"],
    "CPV": ["c p v", "cpv", "c.p.v"],
    "Despesas Operacionais": ["despesas operacionais", "desp oper", "despesas", "desp. oper"],
    "Lucro Operacional": ["lucro op", "lucro operacional", "lo"],
    "Lucro Líquido": ["lucro liq", "lucro líq", "lucro liquido", "ll"],
    "CDG": ["cdg"],
    "NCG": ["ncg"],
    "Tesouraria": ["tesouraria", "tes"],
    "Contas a Receber": ["contas a receber", "cr"],
    "Estoque MP": ["estoque mp", "est. mp", "est mp"],
    "Estoque PA": ["estoque pa", "est. pa", "est pa", "produto acabado"],
    "Ativo Operacional": ["ativo op", "ativo operacional", "ao"],
    "Fornecedores": ["fornecedor", "forn", "forncedores"],
    "Tributos": ["tribut"],
    "Outros": ["outros"],
    "Passivo Operacional": ["passivo op", "passivo operacional", "po"],
}

PERGUNTAS_PADRAO = {
    "Quantidade vendida": [
        ("Resultado Contábil — por que o Δ Faturamento é superior ao Δ CPV?",
         "O faturamento aumentou diretamente pelo crescimento da quantidade vendida. O CPV também aumenta, mas tende a crescer em proporção menor porque parte dos custos é fixa e passa a ser diluída em um volume maior de produção. Por isso, o aumento percentual do faturamento fica superior ao aumento percentual do CPV, melhorando a margem e o resultado operacional."),
        ("Resultado Patrimonial [NCG] — o que justifica o aumento em Fornecedores?",
         "O aumento em fornecedores ocorre porque, para sustentar o maior volume vendido, a empresa precisa produzir mais e comprar mais matéria-prima. Como parte dessas compras é realizada a prazo, o saldo de fornecedores no passivo operacional aumenta."),
        ("Resultado Financeiro — o que reduz os empréstimos?",
         "A variável que contribui positivamente para a redução dos empréstimos é o aumento do faturamento e dos recebimentos. Como a empresa vende mais e gera mais caixa, ela passa a depender menos de capital de terceiros."),
        ("Resultado Financeiro — o que pode contribuir negativamente para os empréstimos?",
         "O fator negativo é que o crescimento da operação exige mais produção, compras, estoques e contas a receber. Isso aumenta a necessidade de capital de giro e pode consumir parte da melhora gerada pelo aumento das vendas."),
    ],
    "Preço de venda": [
        ("Resultado Patrimonial [NCG] — o que justifica o aumento no CR/Contas a Receber?",
         "O contas a receber aumenta porque o preço de venda maior eleva o faturamento. Como parte das vendas ocorre a prazo, o valor financeiro ainda não recebido pela empresa também aumenta."),
        ("Resultado Contábil — explique a variação nas Despesas Operacionais",
         "As despesas operacionais variam porque algumas contas, como comissões, marketing/publicidade e outras despesas variáveis, acompanham o faturamento. Quando o preço aumenta e a receita cresce, essas despesas também podem crescer, mesmo sem alteração relevante na quantidade produzida."),
        ("Resultado Contábil — por que o CPV fica igual ou quase igual?",
         "Quando a premissa altera apenas o preço de venda, a quantidade produzida e o consumo de matéria-prima permanecem praticamente iguais. Por isso, o CPV tende a ficar igual ou variar pouco."),
        ("Resultado Contábil — por que LO pode ser igual ao LL?",
         "O lucro operacional pode ser igual ao lucro líquido quando a empresa ainda permanece em prejuízo. Nesse caso, não há incidência de imposto sobre lucro, então o resultado operacional e o resultado líquido ficam iguais."),
    ],
    "Produção / Estoque PA": [
        ("Resultado Contábil — por que o impacto no resultado contábil pode não ser significativo?",
         "O impacto pode não ser significativo porque a premissa altera principalmente o nível de estoque e a forma de produção, sem mexer diretamente no faturamento. Assim, o efeito aparece mais no custo unitário e no capital de giro do que na receita."),
        ("Resultado Contábil — por que a readequação do estoque PA pode aumentar o CPV?",
         "A readequação do estoque de produtos acabados pode reduzir a produção em alguns períodos. Com menor produção, os custos fixos são distribuídos em menos unidades, elevando o custo unitário e, consequentemente, o CPV."),
        ("Resultado Patrimonial [NCG] — por que houve redução na NCG?",
         "A NCG pode reduzir porque a diminuição do estoque de produto acabado reduz o ativo operacional. Quando o ativo operacional cai mais do que o passivo operacional, a necessidade de capital de giro diminui."),
    ],
    "Compra / Estoque MP": [
        ("Resultado Contábil — por que a readequação do estoque MP pode ter impacto positivo?",
         "A readequação do estoque de matéria-prima pode reduzir compras, armazenagem e necessidade de financiamento. Isso melhora o resultado principalmente quando diminui desembolsos, empréstimos ou despesas financeiras."),
        ("Resultado Contábil — e se a empresa não estivesse endividada?",
         "Se a empresa não estivesse endividada, a redução de compras teria menor impacto no resultado contábil, porque não haveria uma economia relevante de despesas financeiras. O benefício ficaria mais concentrado no caixa e na necessidade de capital de giro."),
        ("Resultado Patrimonial [NCG] — por que houve redução em Fornecedores?",
         "A conta fornecedores reduz quando a empresa passa a comprar menos matéria-prima ou reduz o nível de estoque de MP. Comprando menos a prazo, o saldo de obrigações com fornecedores também diminui."),
    ],
    "Preço dos insumos": [
        ("Resultado Financeiro — que variável contribuiu positivamente para a redução dos empréstimos?",
         "A redução no preço dos insumos diminui os desembolsos com compras e melhora a geração de caixa. Com menor necessidade de recursos para financiar a operação, a empresa depende menos de empréstimos."),
        ("Resultado Contábil — por que houve redução no CPV?",
         "O CPV reduz porque o preço de aquisição dos insumos caiu. Como matéria-prima é componente relevante do custo de produção, a redução do custo dos insumos diminui o custo total dos produtos vendidos."),
        ("Resultado Patrimonial [NCG] — por que reduziu Estoque de PA?",
         "O estoque de produtos acabados pode reduzir em valor porque os produtos passam a carregar um custo unitário menor. Mesmo com quantidade parecida, o valor contábil do estoque diminui quando o custo de produção cai."),
    ],
    "Outra": [
        ("Interpretação geral",
         "A análise deve observar o efeito dominó da premissa sobre RF, RC e RP/NCG. Primeiro avalie faturamento, CPV e despesas; depois empréstimos e despesas financeiras; por fim, contas a receber, estoques, fornecedores, CDG, NCG e tesouraria."),
    ]
}


def normaliza(x):
    return re.sub(r"\s+", " ", str(x).strip().lower())


def to_float(x):
    if x is None or x == "":
        return None
    if isinstance(x, (int, float)):
        return float(x)
    s = str(x).replace("R$", "").replace("%", "").strip()
    s = s.replace("−", "-").replace("–", "-")
    # pt-BR: 1.234.567,89
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


def montar_df(registros):
    if not registros:
        return pd.DataFrame(columns=["Conta", "Original", "Novo Valor", "Variação", "%"])
    df = pd.DataFrame(registros)
    df = df.dropna(subset=["Conta"])
    df = df.drop_duplicates(subset=["Conta"], keep="last")
    ordem = {c: i for i, c in enumerate(CONTAS_PADRAO)}
    df["ordem"] = df["Conta"].map(ordem).fillna(999)
    df = df.sort_values("ordem").drop(columns="ordem")
    for col in ["Original", "Novo Valor"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df["Variação"] = df["Novo Valor"] - df["Original"]
    df["%"] = df.apply(lambda r: ((r["Variação"] / abs(r["Original"])) * 100) if r["Original"] else 0, axis=1)
    return df


def ler_planilha(uploaded):
    data = uploaded.read()
    wb = load_workbook(BytesIO(data), data_only=True, read_only=True)
    candidatos = []
    # procura linhas com nome da conta e dois números próximos à direita
    for ws in wb.worksheets:
        for linha_idx, row in enumerate(ws.iter_rows(values_only=True), start=1):
            vals = list(row)
            for i, v in enumerate(vals):
                conta = detectar_conta(v)
                if conta:
                    nums = []
                    for j in range(i + 1, min(i + 14, len(vals))):
                        n = to_float(vals[j])
                        if n is not None:
                            nums.append(n)
                    if len(nums) >= 2:
                        candidatos.append({
                            "Conta": conta,
                            "Original": nums[0],
                            "Novo Valor": nums[1],
                            "Aba": ws.title,
                            "Linha": linha_idx,
                        })
    return montar_df(candidatos)


def ocr_imagem(uploaded):
    if pytesseract is None:
        return ""
    img = Image.open(uploaded).convert("RGB")
    # aumenta a imagem para melhorar OCR em prints de planilha
    w, h = img.size
    if w < 1800:
        img = img.resize((w * 2, h * 2))
    try:
        texto = pytesseract.image_to_string(img, lang="por")
    except Exception:
        texto = pytesseract.image_to_string(img)
    return texto


def extrair_df_texto(texto):
    registros = []
    linhas = texto.splitlines()
    num_re = re.compile(r"-?\d{1,3}(?:\.\d{3})*,\d{2}|-?\d+[,.]\d+|-?\d+")
    for linha in linhas:
        conta = detectar_conta(linha)
        if not conta:
            continue
        numeros = [to_float(x.group()) for x in num_re.finditer(linha)]
        numeros = [n for n in numeros if n is not None]
        if len(numeros) >= 2:
            registros.append({"Conta": conta, "Original": numeros[0], "Novo Valor": numeros[1]})
    return montar_df(registros)


def sinal(v):
    if pd.isna(v) or abs(v) < 0.01:
        return "não teve variação relevante"
    return "aumentou" if v > 0 else "diminuiu"


def val(df, conta, campo="Variação"):
    if df is None or df.empty or campo not in df.columns:
        return 0.0
    linha = df[df["Conta"] == conta]
    if linha.empty:
        return 0.0
    try:
        return float(linha.iloc[0][campo])
    except Exception:
        return 0.0


def pct(df, conta):
    return val(df, conta, "%")


def frase_moeda(v):
    return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def gerar_respostas(df, premissa):
    fat, cpv, desp = val(df, "Faturamento"), val(df, "CPV"), val(df, "Despesas Operacionais")
    emp, dfin = val(df, "Empréstimos"), val(df, "Desp. Financeiras")
    cr, forn = val(df, "Contas a Receber"), val(df, "Fornecedores")
    ao, po = val(df, "Ativo Operacional"), val(df, "Passivo Operacional")
    ncg, cdg, tes = val(df, "NCG"), val(df, "CDG"), val(df, "Tesouraria")
    pfat, pcpv = pct(df, "Faturamento"), pct(df, "CPV")

    respostas = list(PERGUNTAS_PADRAO.get(premissa, PERGUNTAS_PADRAO["Outra"]))

    conclusao = []
    if fat > 0:
        conclusao.append("o faturamento aumentou, indicando maior geração de receita")
    elif fat < 0:
        conclusao.append("o faturamento diminuiu, reduzindo a geração de receita")
    if fat > 0 and cpv > 0 and abs(pfat) > abs(pcpv):
        conclusao.append("o CPV cresceu proporcionalmente menos que o faturamento, sinalizando ganho de diluição/margem")
    if emp < 0:
        conclusao.append("os empréstimos diminuíram, mostrando menor dependência de capital de terceiros")
    if dfin < 0:
        conclusao.append("as despesas financeiras caíram em função da menor necessidade de empréstimos")
    if ncg > 0:
        conclusao.append("a NCG aumentou, normalmente porque o ativo operacional cresceu mais que o passivo operacional")
    elif ncg < 0:
        conclusao.append("a NCG diminuiu, reduzindo a necessidade de financiamento da operação")
    if tes > 0:
        conclusao.append("a tesouraria melhorou")
    elif tes < 0:
        conclusao.append("a tesouraria piorou")

    resumo = "; ".join(conclusao) if conclusao else "não foram encontradas variações suficientes para uma conclusão automática forte."
    respostas.append(("Diagnóstico automático pelos valores encontrados",
        f"Pelos valores lidos, o faturamento {sinal(fat)} ({frase_moeda(fat)}), o CPV {sinal(cpv)} ({frase_moeda(cpv)}) e as despesas operacionais {sinal(desp)} ({frase_moeda(desp)}). "
        f"No resultado financeiro, os empréstimos {sinal(emp)} ({frase_moeda(emp)}) e as despesas financeiras {sinal(dfin)} ({frase_moeda(dfin)}). "
        f"No patrimonial, contas a receber {sinal(cr)}, fornecedores {sinal(forn)}, ativo operacional {sinal(ao)}, passivo operacional {sinal(po)}, NCG {sinal(ncg)}, CDG {sinal(cdg)} e tesouraria {sinal(tes)}. "
        f"Conclusão: {resumo}."))
    return respostas


st.title("📊 Interpretador PCO — RF, RC e RP/NCG")
st.write("Envie Excel ou imagem/print da simulação. O app tenta ler original x novo valor e gera resposta no padrão das questões da FAE.")

with st.sidebar:
    st.header("Configuração")
    premissa = st.selectbox("Tipo de premissa", [
        "Quantidade vendida", "Preço de venda", "Produção / Estoque PA", "Compra / Estoque MP", "Preço dos insumos", "Outra"
    ])
    modo = st.radio("Modo", ["Carregar arquivo", "Manual"])

edit = None
texto_ocr = ""

if modo == "Carregar arquivo":
    arq = st.file_uploader("Enviar Excel ou imagem", type=["xlsx", "xlsm", "png", "jpg", "jpeg"])
    if arq:
        nome = arq.name.lower()
        if nome.endswith((".xlsx", ".xlsm")):
            with st.spinner("Lendo planilha..."):
                df = ler_planilha(arq)
        else:
            st.image(arq, caption="Imagem enviada", use_container_width=True)
            with st.spinner("Lendo texto da imagem via OCR..."):
                texto_ocr = ocr_imagem(arq)
            with st.expander("Texto lido da imagem/OCR", expanded=False):
                texto_ocr = st.text_area("Você pode corrigir o texto aqui se o OCR errar", texto_ocr, height=220)
            df = extrair_df_texto(texto_ocr)

        if df.empty:
            st.warning("Não consegui detectar automaticamente as contas. Use o modo manual ou preencha/ajuste a tabela abaixo.")
            df = pd.DataFrame({"Conta": CONTAS_PADRAO, "Original": 0.0, "Novo Valor": 0.0})
        st.subheader("Valores detectados/ajustáveis")
        base = df[[c for c in ["Conta", "Original", "Novo Valor"] if c in df.columns]].copy()
        edit = st.data_editor(base, use_container_width=True, num_rows="dynamic")
    else:
        st.info("Envie a planilha ou imagem para começar.")
else:
    df = pd.DataFrame({"Conta": CONTAS_PADRAO, "Original": 0.0, "Novo Valor": 0.0})
    edit = st.data_editor(df, use_container_width=True, num_rows="dynamic")

if edit is not None:
    edit = montar_df(edit.to_dict("records"))
    st.subheader("Comparativo calculado")
    st.dataframe(edit, use_container_width=True)

    st.subheader("Respostas automáticas")
    respostas = gerar_respostas(edit, premissa)
    for titulo, texto in respostas:
        with st.expander(titulo, expanded=True):
            st.write(texto)

    st.subheader("Texto corrido para colar")
    texto_final = "\n\n".join([f"{t}\n{txt}" for t, txt in respostas])
    st.text_area("Resposta pronta", texto_final, height=360)

    st.download_button(
        "Baixar resposta em TXT",
        data=texto_final.encode("utf-8"),
        file_name="resposta_pco.txt",
        mime="text/plain",
    )
