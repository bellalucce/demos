"""Gera a planilha de controle financeiro e fluxo de caixa (com fórmulas) e o dados.js da página. Dados inventados de uma loja pequena."""
import datetime as dt
import json
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

AQUI = Path(__file__).parent
SALDO_INICIAL = 4000.00
CATEGORIAS = [  # nome, tipo
    ("Vendas no balcão", "Entrada"), ("Vendas online", "Entrada"), ("Outras entradas", "Entrada"),
    ("Aluguel", "Saída"), ("Fornecedor", "Saída"), ("Energia", "Saída"), ("Internet", "Saída"),
    ("Embalagens", "Saída"), ("Taxa de cartão", "Saída"), ("Pró-labore", "Saída"), ("Outras despesas", "Saída"),
]
# meta de vendas de cada mês (o 2º mês fecha no vermelho, para a planilha mostrar o alerta). 22/11 05h54 (Claude):
# ago-out, meses que já passaram (antes ia até 30/11, datas futuras); um lançamento de hoje abre novembro na página
MESES = [(2026, 8, 13200.00), (2026, 9, 10400.00), (2026, 10, 16400.00)]
FORMAS = ["Pix", "Cartão", "Dinheiro", "Boleto", "Transferência"]


def montar_lancamentos():
    rnd = random.Random(2611)
    out = []
    for ano, mes, meta in MESES:
        def d(dia):
            return dt.date(ano, mes, dia)
        dias_balcao = [2, 5, 9, 12, 16, 19, 23, 27]
        dias_online = [4, 11, 18, 26]
        pesos = [rnd.uniform(0.7, 1.3) for _ in range(12)]
        alvo_b, alvo_o = meta * 0.68, meta * 0.32
        vb = [round(alvo_b * p / sum(pesos[:8]), 2) for p in pesos[:8]]
        vo = [round(alvo_o * p / sum(pesos[8:]), 2) for p in pesos[8:]]
        vb[-1] = round(alvo_b - sum(vb[:-1]), 2)
        vo[-1] = round(alvo_o - sum(vo[:-1]), 2)
        for dia, v in zip(dias_balcao, vb):
            out.append((d(dia), "Vendas do balcão", "Vendas no balcão", "Entrada", "Cartão" if dia % 3 else "Dinheiro", v))
        for dia, v in zip(dias_online, vo):
            out.append((d(dia), "Pedidos da loja online", "Vendas online", "Entrada", "Pix", v))
        saidas = [
            (5, "Aluguel da loja", "Aluguel", "Boleto", 2200.00),
            (5, "Retirada da dona", "Pró-labore", "Transferência", 3000.00),
            (10, "Compra de mercadoria (fornecedor A)", "Fornecedor", "Boleto", round(meta * 0.21, 2)),
            (12, "Internet da loja", "Internet", "Boleto", 119.90),
            (15, "Conta de luz", "Energia", "Boleto", round(rnd.uniform(430, 520), 2)),
            (18, "Sacolas e caixas", "Embalagens", "Pix", round(rnd.uniform(330, 420), 2)),
            (24, "Compra de mercadoria (fornecedor B)", "Fornecedor", "Pix", round(meta * 0.2, 2)),
            (30, "Taxa da maquininha de cartão", "Taxa de cartão", "Cartão", round(meta * 0.04, 2)),
        ]
        for dia, desc, cat, forma, v in saidas:
            out.append((d(dia), desc, cat, "Saída", forma, v))
    out.sort(key=lambda x: x[0])
    return out


LANC = montar_lancamentos()
N = 500  # linhas de lançamento cobertas pelas fórmulas

AMARELO = PatternFill("solid", fgColor="FFF8D6")
LILAS = PatternFill("solid", fgColor="EDE7F6")
TOPO = PatternFill("solid", fgColor="00796B")
BRANCO = Font(bold=True, color="FFFFFF")
FINA = Border(bottom=Side(style="thin", color="DDDDDD"))
REAL = 'R$ #,##0.00;[Red]-R$ #,##0.00'
VERMELHO = dict(fill=PatternFill("solid", bgColor="FFCDD2", fgColor="FFCDD2"), font=Font(bold=True, color="B71C1C"))


def cabecalho(ws, nomes, larguras, linha=1):
    for i, (n, w) in enumerate(zip(nomes, larguras), start=1):
        c = ws.cell(row=linha, column=i, value=n)
        c.fill, c.font = TOPO, BRANCO
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[c.column_letter].width = w
    ws.row_dimensions[linha].height = 30


def calcular():
    res = []
    saldo = SALDO_INICIAL
    for a, m, _ in MESES:
        ls = [x for x in LANC if x[0].year == a and x[0].month == m]
        e = round(sum(x[5] for x in ls if x[3] == "Entrada"), 2)
        s = round(sum(x[5] for x in ls if x[3] == "Saída"), 2)
        res.append(dict(mes=f"{m:02d}/{a}", entradas=e, saidas=s, lucro=round(e - s, 2),
                        margem=round((e - s) / e, 4) if e else 0, saldo_inicial=round(saldo, 2),
                        saldo_final=round(saldo + e - s, 2)))
        saldo += e - s
    return res


def gerar():
    wb = Workbook()
    L = wb.active
    L.title = "Lançamentos"
    cabecalho(L, ["Data", "Descrição", "Categoria", "Tipo", "Forma de pagamento", "Valor (R$)"], [12, 40, 20, 11, 20, 15])
    L.freeze_panes = "A2"
    for r in range(2, N + 2):
        lin = LANC[r - 2] if r - 2 < len(LANC) else (None,) * 6
        for c, v in enumerate(lin, start=1):
            cel = L.cell(row=r, column=c, value=v)
            cel.fill, cel.border = AMARELO, FINA
            if c == 1:
                cel.number_format = "DD/MM/YYYY"
            if c == 6:
                cel.number_format = REAL
    for dv, col in (
        (DataValidation(type="list", formula1="=Categorias!$A$2:$A$100", allow_blank=True), "C"),
        (DataValidation(type="list", formula1='"Entrada,Saída"', allow_blank=True), "D"),
        (DataValidation(type="list", formula1='"' + ",".join(FORMAS) + '"', allow_blank=True), "E"),
        (DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                        error="Digite o valor sem sinal (a coluna Tipo diz se é Entrada ou Saída)."), "F"),
    ):
        L.add_data_validation(dv)
        dv.add(f"{col}2:{col}{N + 1}")
    L.auto_filter.ref = f"A1:F{N + 1}"

    K = wb.create_sheet("Categorias")
    cabecalho(K, ["Categoria", "Tipo"], [26, 12])
    K.freeze_panes = "A2"
    for r in range(2, 17):  # 15 linhas: as 11 usadas e 4 livres para categoria nova
        n, t = CATEGORIAS[r - 2] if r - 2 < len(CATEGORIAS) else (None, None)
        for c, v in enumerate((n, t), start=1):
            cel = K.cell(row=r, column=c, value=v)
            cel.fill, cel.border = AMARELO, FINA
    dvt = DataValidation(type="list", formula1='"Entrada,Saída"', allow_blank=True)
    K.add_data_validation(dvt)
    dvt.add("B2:B16")

    def rng(col):
        return f"Lançamentos!${col}$2:${col}${N + 1}"

    R = wb.create_sheet("Resumo do mês")
    cabecalho(R, ["Mês", "Entradas (R$)", "Saídas (R$)", "Lucro (R$)", "Margem", "Saldo acumulado (R$)"], [26, 16, 16, 16, 12, 20])
    R.freeze_panes = "A2"
    nm = len(MESES)
    for i, (a, m, _) in enumerate(MESES):
        r = i + 2
        f = {
            1: dt.date(a, m, 1) if i == 0 else f"=EDATE(A{r - 1},1)",
            2: f'=SUMIFS({rng("F")},{rng("D")},"Entrada",{rng("A")},">="&$A{r},{rng("A")},"<"&EDATE($A{r},1))',
            3: f'=SUMIFS({rng("F")},{rng("D")},"Saída",{rng("A")},">="&$A{r},{rng("A")},"<"&EDATE($A{r},1))',
            4: f"=B{r}-C{r}", 5: f"=IFERROR(D{r}/B{r},0)", 6: f"='Fluxo de caixa'!$B$2+SUM($D$2:D{r})",
        }
        for c, v in f.items():
            cel = R.cell(row=r, column=c, value=v)
            cel.fill, cel.border = (AMARELO if (c == 1 and i == 0) else LILAS), FINA
            cel.number_format = "MM/YYYY" if c == 1 else ("0.0%" if c == 5 else REAL)
            if c == 1:
                cel.alignment = Alignment(horizontal="left")
    t = nm + 2
    R.cell(row=t, column=1, value="TOTAL DO PERÍODO").font = Font(bold=True)
    for c, col in ((2, "B"), (3, "C"), (4, "D")):
        cel = R.cell(row=t, column=c, value=f"=SUM({col}2:{col}{t - 1})")
        cel.font, cel.number_format = Font(bold=True), REAL
    cel = R.cell(row=t, column=5, value=f"=IFERROR(D{t}/B{t},0)")
    cel.font, cel.number_format = Font(bold=True), "0.0%"
    R.conditional_formatting.add(f"D2:D{t}", CellIsRule(operator="lessThan", formula=["0"], **VERMELHO))
    R.conditional_formatting.add(f"F2:F{t - 1}", CellIsRule(operator="lessThan", formula=["0"], **VERMELHO))

    h = t + 3  # tabela por categoria
    R.cell(row=h - 1, column=1, value="Total por categoria").font = Font(bold=True, color="00796B")
    for i, n in enumerate(["Categoria", "Tipo"] + [None] * nm + ["Total"], start=1):
        c = R.cell(row=h, column=i, value=n)
        c.fill, c.font = TOPO, BRANCO
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i in range(nm):
        c = R.cell(row=h, column=3 + i, value=f"=$A${i + 2}")
        c.number_format = "MM/YYYY"
    for k in range(15):  # espelha as 15 linhas da aba Categorias
        r = h + 1 + k
        R.cell(row=r, column=1, value=f'=IF(Categorias!A{k + 2}="","",Categorias!A{k + 2})')
        R.cell(row=r, column=2, value=f'=IF(Categorias!B{k + 2}="","",Categorias!B{k + 2})')
        for i in range(nm):
            col = 3 + i
            letra = R.cell(row=h, column=col).column_letter
            R.cell(row=r, column=col, value=(
                f'=IF($A{r}="","",SUMIFS({rng("F")},{rng("C")},$A{r},{rng("A")},">="&{letra}${h},{rng("A")},"<"&EDATE({letra}${h},1)))'))
            R.cell(row=r, column=col).number_format = REAL
        ult = R.cell(row=h, column=2 + nm).column_letter
        R.cell(row=r, column=3 + nm, value=f'=IF($A{r}="","",SUM(C{r}:{ult}{r}))').number_format = REAL
        for c in range(1, 4 + nm):
            R.cell(row=r, column=c).fill, R.cell(row=r, column=c).border = LILAS, FINA
    for c in range(3, 4 + nm):
        R.column_dimensions[R.cell(row=h, column=c).column_letter].width = 16

    F = wb.create_sheet("Fluxo de caixa")
    cabecalho(F, ["Mês", "Saldo inicial (R$)", "Entradas (R$)", "Saídas (R$)", "Saldo final (R$)"], [14, 20, 16, 16, 18])
    F.freeze_panes = "A2"
    for i in range(nm):
        r = i + 2
        f = {1: f"='Resumo do mês'!A{r}", 2: SALDO_INICIAL if i == 0 else f"=E{r - 1}",
             3: f"='Resumo do mês'!B{r}", 4: f"='Resumo do mês'!C{r}", 5: f"=B{r}+C{r}-D{r}"}
        for c, v in f.items():
            cel = F.cell(row=r, column=c, value=v)
            cel.fill, cel.border = (AMARELO if (c == 2 and i == 0) else LILAS), FINA
            cel.number_format = "MM/YYYY" if c == 1 else REAL
            if c == 1:
                cel.alignment = Alignment(horizontal="left")
    F.conditional_formatting.add(f"B2:E{nm + 1}", CellIsRule(operator="lessThan", formula=["0"], **VERMELHO))
    F.cell(row=nm + 3, column=1, value="O saldo inicial do primeiro mês (célula amarela) é o dinheiro que havia no caixa e no banco. Saldo negativo fica em vermelho.")

    U = wb.create_sheet("Como usar")
    U.column_dimensions["A"].width = 105
    passos = [
        "COMO USAR (células amarelas você preenche; as lilases são fórmula, não precisa mexer)",
        "1. Aba Lançamentos: uma linha para cada dinheiro que entra ou sai. Preencha data, descrição, categoria (lista), tipo Entrada ou Saída (lista), forma de pagamento e valor.",
        "2. Digite o valor sempre positivo. Quem diz se é ganho ou gasto é a coluna Tipo.",
        "3. Aba Resumo do mês: entradas, saídas, lucro (entradas menos saídas), margem e saldo acumulado, mais o total de cada categoria. Tudo sai sozinho.",
        "4. Aba Fluxo de caixa: saldo inicial, entradas, saídas e saldo final de cada mês, um puxando o outro. Saldo negativo fica em vermelho.",
        "5. Categoria nova: escreva o nome (e se é Entrada ou Saída) na próxima linha amarela da aba Categorias. Ela já aparece na lista da aba Lançamentos e no resumo por categoria.",
        "6. Mês novo: na aba Resumo do mês e na aba Fluxo de caixa, copie a última linha de mês para baixo (as fórmulas seguem o mês seguinte). Na tabela de categorias, copie a última coluna de mês.",
        "7. Filtro: na aba Lançamentos, use as setinhas do cabeçalho para ver só uma categoria, um tipo ou um mês.",
        "Demonstração com dados de exemplo (loja fictícia). Abre no Excel, no LibreOffice e no Google Planilhas.",
    ]
    for i, s in enumerate(passos, start=1):
        cel = U.cell(row=i, column=1, value=s)
        cel.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 1:
            cel.font = Font(bold=True, color="00796B")
    wb.calculation.fullCalcOnLoad = True
    wb.save(AQUI / "planilha-fluxo-caixa.xlsx")

    js = {"saldoInicial": SALDO_INICIAL,
          "categorias": [dict(nome=n, tipo=t) for n, t in CATEGORIAS], "formas": FORMAS,
          "lancamentos": [dict(data=x[0].isoformat(), desc=x[1], cat=x[2], tipo=x[3], forma=x[4], valor=x[5]) for x in LANC]}
    (AQUI / "dados.js").write_text("// gerado por _gerar.py (mesmos dados da planilha)\nvar DADOS=" +
                                   json.dumps(js, ensure_ascii=False) + ";\n", encoding="utf-8")
    (AQUI / "_conta.json").write_text(json.dumps(calcular(), ensure_ascii=False, indent=1), encoding="utf-8")


if __name__ == "__main__":
    gerar()
    print(len(LANC), "lançamentos")
    for x in calcular():
        print(x)
