"""Gera a planilha da fatura de cartão organizada (com fórmulas) e o dados.js da página. Dados inventados."""
import datetime as dt
import json
import random
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

AQUI = Path(__file__).parent
CATEGORIAS = ["Alimentação", "Mercado", "Transporte", "Saúde", "Assinaturas", "Compras online", "Casa",
              "Vestuário", "Eletrônicos", "Educação", "Lazer", "Tarifas e anuidade", "Sem categoria"]
# palavra-chave (sem * nem ?, que o SEARCH do Excel trata como curinga) -> estabelecimento limpo, categoria
REGRAS = [
    ("IFD", "iFood (pedido de comida)", "Alimentação"),
    ("PADARIA", "Padaria Pão Quente", "Alimentação"),
    ("CAFE DO LARGO", "Café do Largo", "Alimentação"),
    ("BURGER STATION", "Burger Station", "Alimentação"),
    ("SUPERMERCADO", "Supermercado Bom Preço", "Mercado"),
    ("ATACADO", "Atacado Econômico", "Mercado"),
    ("HORTIFRUTI", "Hortifruti Verde Vida", "Mercado"),
    ("UBER", "Uber (corrida)", "Transporte"),
    ("99POP", "99 (corrida)", "Transporte"),
    ("POSTO ESTRELA", "Posto Estrela", "Transporte"),
    ("SEMPARAR", "Pedágio (Sem Parar)", "Transporte"),
    ("ESTAPAR", "Estacionamento", "Transporte"),
    ("FARMACIAPOPULAR", "Farmácia Popular", "Saúde"),
    ("DROGARIA", "Drogaria São José", "Saúde"),
    ("CLINICA", "Clínica Vida Plena", "Saúde"),
    ("ACADEMIA", "Academia Forma Fit", "Saúde"),
    ("NETFLIX", "Netflix", "Assinaturas"),
    ("SPOTIFY", "Spotify", "Assinaturas"),
    ("GOOGLE", "Google (armazenamento)", "Assinaturas"),
    ("AMAZONPRIME", "Amazon Prime", "Assinaturas"),
    ("MERCADOLIVRE", "Mercado Livre", "Compras online"),
    ("SHOPEE", "Shopee", "Compras online"),
    ("AMAZONMKTP", "Amazon (loja)", "Compras online"),
    ("FIBRATOTAL", "Internet FibraTotal", "Casa"),
    ("CASA BELA", "Loja Casa Bela", "Casa"),
    ("OBRA MATERIAIS", "Casa e Obra Materiais", "Casa"),
    ("ELETRO FACIL", "Eletro Fácil (notebook)", "Eletrônicos"),
    ("MODA NOVA", "Loja Moda Nova", "Vestuário"),
    ("CALCADOS", "Calçados Passo Certo", "Vestuário"),
    ("LINGUA VIVA", "Curso de idiomas Língua Viva", "Educação"),
    ("LIVRARIA", "Livraria Página Um", "Educação"),
    ("CINEMAX", "Cinemax (ingressos)", "Lazer"),
    ("PARQUE", "Parque Aquático Sol", "Lazer"),
    ("ANUIDADE", "Anuidade do cartão", "Tarifas e anuidade"),
    ("IOF", "IOF (compra no exterior)", "Tarifas e anuidade"),
]
# (mês, dia, descrição bruta, valor ou (mín, máx)); parcelas seguem a mesma compra nos meses seguintes
MOV = [
    (8, 1, "ANUIDADE DIFERENC 04/12", 39.90), (8, 1, "LOJA CASA BELA 01/06", 214.50),
    (8, 1, "ELETRO FACIL NOTEBOOK 05/10", 389.90), (8, 2, "CURSO LINGUA VIVA 08/12", 198.00),
    (8, 3, "SUPERMERCADO BOM PRECO", (240, 330)), (8, 4, "IFD*RESTAURANTE SABOR", (42, 68)),
    (8, 5, "NETFLIX.COM", 39.90), (8, 6, "UBER *TRIP", (14, 29)), (8, 8, "PAG*FARMACIAPOPULAR", (38, 85)),
    (8, 9, "MERCADOLIVRE*2PROD", (120, 210)), (8, 10, "POSTO ESTRELA", (180, 240)),
    (8, 11, "SPOTIFY BR", 21.90), (8, 12, "PADARIA PAO QUENTE", (18, 34)), (8, 13, "FIBRATOTAL INTERNET", 119.90),
    (8, 15, "IFD*PIZZARIA NONNA", (58, 92)), (8, 16, "ACADEMIA FORMA FIT", 109.00), (8, 17, "ATACADO ECONOMICO", (310, 420)),
    (8, 19, "UBER *TRIP", (16, 31)), (8, 21, "CINEMAX INGRESSOS", (46, 78)), (8, 22, "GOOGLE*STORAGE", 9.99),
    (8, 24, "SHOPEE*PEDIDO", (45, 99)), (8, 26, "DROGARIA SAO JOSE", (27, 64)), (8, 28, "PAGTO*DIVERSOS 4471", 73.20),
    (9, 1, "ANUIDADE DIFERENC 05/12", 39.90), (9, 1, "LOJA CASA BELA 02/06", 214.50),
    (9, 1, "ELETRO FACIL NOTEBOOK 06/10", 389.90), (9, 2, "CURSO LINGUA VIVA 09/12", 198.00),
    (9, 2, "SUPERMERCADO BOM PRECO", (220, 340)), (9, 3, "IFD*RESTAURANTE SABOR", (40, 70)),
    (9, 5, "NETFLIX.COM", 39.90), (9, 5, "99POP*CORRIDA", (12, 27)), (9, 7, "HORTIFRUTI VERDE VIDA", (48, 90)),
    (9, 8, "AMAZONPRIME*MENSAL", 19.90), (9, 9, "MERCADOLIVRE*3PROD", (95, 240)), (9, 10, "POSTO ESTRELA", (170, 250)),
    (9, 11, "SPOTIFY BR", 21.90), (9, 13, "FIBRATOTAL INTERNET", 119.90), (9, 14, "CAFE DO LARGO", (22, 41)),
    (9, 16, "ACADEMIA FORMA FIT", 109.00), (9, 17, "UBER *TRIP", (15, 33)), (9, 19, "ATACADO ECONOMICO", (290, 400)),
    (9, 20, "LIVRARIA PAGINA UM", (68, 140)), (9, 22, "GOOGLE*STORAGE", 9.99), (9, 23, "PAG*FARMACIAPOPULAR", (30, 80)),
    (9, 25, "IFD*PIZZARIA NONNA", (55, 95)), (9, 27, "LOJA MODA NOVA 01/03", 159.90),
    (9, 28, "ESTAPAR ESTACIONAMENTO", (14, 32)), (9, 29, "PAGTO*DIVERSOS 5120", 41.00),
    (10, 1, "ANUIDADE DIFERENC 06/12", 39.90), (10, 1, "LOJA CASA BELA 03/06", 214.50),
    (10, 1, "ELETRO FACIL NOTEBOOK 07/10", 389.90), (10, 1, "CURSO LINGUA VIVA 10/12", 198.00),
    (10, 2, "LOJA MODA NOVA 02/03", 159.90), (10, 3, "SUPERMERCADO BOM PRECO", (230, 350)),
    (10, 4, "IFD*RESTAURANTE SABOR", (44, 72)), (10, 5, "NETFLIX.COM", 39.90), (10, 6, "UBER *TRIP", (13, 30)),
    (10, 7, "PAG*FARMACIAPOPULAR", (36, 90)), (10, 8, "AMAZONPRIME*MENSAL", 19.90),
    (10, 9, "MERCADOLIVRE*2PROD", (110, 230)), (10, 10, "POSTO ESTRELA", (185, 245)), (10, 11, "SPOTIFY BR", 21.90),
    (10, 12, "SEMPARAR*PEDAGIO", (11, 24)), (10, 13, "FIBRATOTAL INTERNET", 119.90), (10, 14, "BURGER STATION", (34, 62)),
    (10, 15, "IFD*PIZZARIA NONNA", (60, 98)), (10, 16, "ACADEMIA FORMA FIT", 109.00), (10, 17, "ATACADO ECONOMICO", (300, 410)),
    (10, 18, "ESTORNO MERCADOLIVRE", -89.90), (10, 19, "PARQUE AQUATICO SOL", (120, 190)),
    (10, 20, "CALCADOS PASSO CERTO", 249.90), (10, 22, "GOOGLE*STORAGE", 9.99), (10, 23, "99POP*CORRIDA", (14, 28)),
    (10, 24, "IOF COMPRA EXTERIOR", 6.84), (10, 25, "OBRA MATERIAIS LTDA", (88, 170)),
    (10, 27, "CLINICA VIDA PLENA", 180.00), (10, 29, "PAGTO*DIVERSOS 6038", 58.40),
]


def categorizar(desc):
    """Mesma lógica da planilha: vale a primeira regra (de cima para baixo) cuja palavra aparece na descrição."""
    u = desc.upper()
    for i, (k, est, cat) in enumerate(REGRAS, start=1):
        if k.upper() in u:
            return i, est, cat
    return 0, desc, "Sem categoria"


def parcela(desc):
    t = desc[-5:]
    if len(desc) > 6 and desc[-6] == " " and t[2] == "/" and t[:2].isdigit() and t[3:].isdigit():
        return t
    return ""


def montar():
    rnd = random.Random(2611)
    out = []
    for m, d, desc, v in MOV:
        val = round(rnd.uniform(*v), 2) if isinstance(v, tuple) else v
        i, est, cat = categorizar(desc)
        out.append(dict(data=dt.date(2026, m, d), desc=desc, est=est, cat=cat, parc=parcela(desc), valor=val))
    out.sort(key=lambda x: x["data"])
    return out


TR = montar()
N = 300  # linhas cobertas pelas fórmulas
NR = 60  # linhas de regras
TOPO = PatternFill("solid", fgColor="1565C0")
BRANCO = Font(bold=True, color="FFFFFF")
AMARELO = PatternFill("solid", fgColor="FFF8D6")
AZULC = PatternFill("solid", fgColor="E3F2FD")
CINZA = PatternFill("solid", fgColor="EEEEEE")
FINA = Border(bottom=Side(style="thin", color="DDDDDD"))
REAL = 'R$ #,##0.00;[Red]-R$ #,##0.00'
VERM = dict(fill=PatternFill("solid", bgColor="FFCDD2", fgColor="FFCDD2"), font=Font(bold=True, color="B71C1C"))


def cab(ws, nomes, larg, linha=1):
    for i, (n, w) in enumerate(zip(nomes, larg), start=1):
        c = ws.cell(row=linha, column=i, value=n)
        ws.column_dimensions[c.column_letter].width = w
        if n is None:
            continue
        c.fill, c.font = TOPO, BRANCO
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[linha].height = 30


def gerar():
    wb = Workbook()
    T = wb.active
    T.title = "Transações"
    cab(T, ["Data", "Descrição original", "Estabelecimento", "Categoria", "Parcela", "Valor (R$)", "Regra usada", "Parcelas que faltam"],
        [12, 34, 30, 20, 10, 14, 12, 12])
    T.freeze_panes = "A2"

    def rg(c):
        return f"Regras!${c}$2:${c}${NR + 1}"

    for r in range(2, N + 2):
        t = TR[r - 2] if r - 2 < len(TR) else None
        a = T.cell(row=r, column=1, value=t["data"] if t else None)
        b = T.cell(row=r, column=2, value=t["desc"] if t else None)
        f = T.cell(row=r, column=6, value=t["valor"] if t else None)
        a.number_format, f.number_format = "DD/MM/YYYY", REAL
        for c in (a, b, f):
            c.fill, c.border = AMARELO, FINA
        T.cell(row=r, column=7, value=(
            f'=IF(B{r}="","",IFERROR(MATCH(1,INDEX(({rg("A")}<>"")*ISNUMBER(SEARCH({rg("A")},B{r})),0),0),0))'))
        T.cell(row=r, column=3, value=f'=IF(B{r}="","",IF(G{r}=0,B{r},INDEX({rg("B")},G{r})))')
        T.cell(row=r, column=4, value=f'=IF(B{r}="","",IF(G{r}=0,"Sem categoria",INDEX({rg("C")},G{r})))')
        T.cell(row=r, column=5, value=(
            f'=IF(B{r}="","",IFERROR(IF(AND(LEN(B{r})>6,MID(B{r},LEN(B{r})-5,1)=" ",MID(B{r},LEN(B{r})-2,1)="/",'
            f'ISNUMBER(--RIGHT(B{r},2)),ISNUMBER(--MID(B{r},LEN(B{r})-4,2))),RIGHT(B{r},5),""),""))'))
        T.cell(row=r, column=8, value=f'=IF(E{r}="",0,VALUE(RIGHT(E{r},2))-VALUE(LEFT(E{r},2)))')
        for c in (3, 4, 5, 7, 8):
            T.cell(row=r, column=c).fill = AZULC if c in (3, 4, 5) else CINZA
            T.cell(row=r, column=c).border = FINA
        T.cell(row=r, column=5).alignment = Alignment(horizontal="center")
    dv = DataValidation(type="list", formula1="=Regras!$E$2:$E$16", allow_blank=True,
                        error="Escolha uma categoria da lista (as categorias ficam na aba Regras).")
    T.add_data_validation(dv)
    dv.add(f"D2:D{N + 1}")
    T.auto_filter.ref = f"A1:H{N + 1}"
    T.conditional_formatting.add(f"D2:D{N + 1}", CellIsRule(operator="equal", formula=['"Sem categoria"'], **VERM))

    G = wb.create_sheet("Regras")
    cab(G, ["Palavra-chave (aparece na descrição)", "Estabelecimento (nome limpo)", "Categoria", None, "Lista de categorias"],
        [32, 32, 22, 3, 24])
    G.freeze_panes = "A2"
    for r in range(2, NR + 2):
        k = REGRAS[r - 2] if r - 2 < len(REGRAS) else (None,) * 3
        for c, v in enumerate(k, start=1):
            cel = G.cell(row=r, column=c, value=v)
            cel.fill, cel.border = AMARELO, FINA
    for r in range(2, 17):
        cel = G.cell(row=r, column=5, value=CATEGORIAS[r - 2] if r - 2 < len(CATEGORIAS) else None)
        cel.fill, cel.border = AMARELO, FINA
    dvc = DataValidation(type="list", formula1="=$E$2:$E$16", allow_blank=True)
    G.add_data_validation(dvc)
    dvc.add(f"C2:C{NR + 1}")

    R = wb.create_sheet("Resumo")
    R.column_dimensions["A"].width = 30
    for col in "BCDEF":
        R.column_dimensions[col].width = 17
    R["A1"] = "Resumo das faturas"
    R["A1"].font = Font(bold=True, size=14, color="1565C0")

    def tr(c):
        return f"Transações!${c}$2:${c}${N + 1}"

    ini = 3
    for i, n in enumerate(["Categoria", None, None, None, "Total", "% do total"], start=1):
        c = R.cell(row=ini, column=i, value=n)
        c.fill, c.font = TOPO, BRANCO
        c.alignment = Alignment(horizontal="center", wrap_text=True)
    for i in range(3):
        c = R.cell(row=ini, column=2 + i, value=dt.date(2026, 8, 1) if i == 0 else f"=EDATE({R.cell(row=ini, column=1 + i).coordinate},1)")
        c.number_format = '"Fatura "MM/YYYY'
    first, last = ini + 1, ini + 15
    for k in range(15):
        r = first + k
        R.cell(row=r, column=1, value=f'=IF(Regras!E{k + 2}="","",Regras!E{k + 2})')
        for i in range(3):
            col = R.cell(row=ini, column=2 + i).column_letter
            R.cell(row=r, column=2 + i, value=(
                f'=IF($A{r}="","",SUMIFS({tr("F")},{tr("D")},$A{r},{tr("A")},">="&{col}${ini},{tr("A")},"<"&EDATE({col}${ini},1)))'))
        R.cell(row=r, column=5, value=f'=IF($A{r}="","",SUM(B{r}:D{r}))')
        R.cell(row=r, column=6, value=f'=IF($A{r}="","",IFERROR(E{r}/$E${last + 1},0))')
        for c in range(1, 7):
            cel = R.cell(row=r, column=c)
            cel.fill, cel.border = AZULC, FINA
            cel.number_format = "0.0%" if c == 6 else REAL
    t = last + 1
    R.cell(row=t, column=1, value="TOTAL DA FATURA").font = Font(bold=True)
    for c in range(2, 6):
        L = R.cell(row=t, column=c).column_letter
        cel = R.cell(row=t, column=c, value=f"=SUM({L}{first}:{L}{last})")
        cel.font, cel.number_format = Font(bold=True), REAL
    k0 = t + 2
    R.cell(row=k0, column=1, value="Números principais").font = Font(bold=True, color="1565C0")
    ult = f'DATE(YEAR(MAX({tr("A")})),MONTH(MAX({tr("A")})),1)'
    itens = [
        ("Total gasto nas 3 faturas", f"=E{t}", REAL),
        ("Número de compras", f'=COUNT({tr("F")})', "0"),
        ("Maior gasto (R$)", f'=MAX({tr("F")})', REAL),
        ("Maior gasto: onde foi", f'=INDEX({tr("C")},MATCH(MAX({tr("F")}),{tr("F")},0))', "@"),
        ("Parcelas futuras comprometidas", f'=SUMPRODUCT(({tr("A")}>={ult})*{tr("F")}*{tr("H")})', REAL),
        ("Compras sem categoria", f'=COUNTIF({tr("D")},"Sem categoria")', "0"),
    ]
    for i, (n, f, nf) in enumerate(itens):
        R.cell(row=k0 + 1 + i, column=1, value=n)
        cel = R.cell(row=k0 + 1 + i, column=2, value=f)
        cel.number_format, cel.font = nf, Font(bold=True)
        cel.alignment = Alignment(horizontal="right")
    R.cell(row=k0 + 8, column=1, value="Parcelas futuras = o que ainda vai cair nas próximas faturas das compras parceladas que aparecem na última fatura.").font = Font(italic=True, color="555555")
    R.conditional_formatting.add(f"B{k0 + 6}", CellIsRule(operator="greaterThan", formula=["0"], **VERM))
    R.freeze_panes = "A4"

    ch = BarChart()
    ch.type, ch.style = "bar", 10
    ch.title = "Gasto por categoria (3 faturas)"
    ch.add_data(Reference(R, min_col=5, min_row=ini, max_row=ini + 13), titles_from_data=True)
    ch.set_categories(Reference(R, min_col=1, min_row=first, max_row=ini + 13))
    ch.legend = None
    ch.height, ch.width = 9.5, 17
    ch.y_axis.numFmt = '"R$ "#,##0'
    ch.x_axis.delete = ch.y_axis.delete = False
    R.add_chart(ch, "H3")
    ch2 = BarChart()
    ch2.type, ch2.style = "col", 10
    ch2.title = "Total de cada fatura"
    ch2.add_data(Reference(R, min_col=2, max_col=4, min_row=t, max_row=t), from_rows=True, titles_from_data=False)
    ch2.set_categories(Reference(R, min_col=2, max_col=4, min_row=ini, max_row=ini))
    ch2.legend = None
    ch2.height, ch2.width = 8, 17
    ch2.y_axis.numFmt = '"R$ "#,##0'
    ch2.x_axis.number_format = "MM/YYYY"
    ch2.x_axis.delete = ch2.y_axis.delete = False
    R.add_chart(ch2, "H23")

    U = wb.create_sheet("Como usar")
    U.column_dimensions["A"].width = 105
    passos = [
        "COMO USAR (células amarelas você preenche; as azuis e cinzas são fórmula, não precisa mexer)",
        "1. Aba Transações: cole nas colunas Data, Descrição original e Valor o que vem da fatura do banco (PDF ou CSV), do jeito que está.",
        "2. O Estabelecimento, a Categoria e a Parcela aparecem sozinhos. A parcela é lida do final da descrição (ex.: LOJA X 02/06).",
        "3. Compra com Categoria vermelha 'Sem categoria': vá na aba Regras e escreva uma palavra-chave que aparece na descrição, o nome limpo e a categoria. Todas as compras parecidas, de todas as faturas, se corrigem juntas.",
        "4. A regra de cima vale primeiro. Se duas palavras servem para a mesma compra, deixe a mais específica mais acima.",
        "5. Não use os símbolos * e ? na palavra-chave (o Excel os trata como curinga). Use só um trecho do nome, como IFD ou FARMACIA.",
        "6. Categoria nova: escreva na lista da coluna E da aba Regras e use nas regras.",
        "7. Aba Resumo: total por categoria e por fatura, maior gasto, parcelas futuras e dois gráficos. Tudo se atualiza sozinho.",
        "8. Fatura de outro mês: cole as linhas novas embaixo na aba Transações. Para mais colunas de mês no Resumo, copie a última coluna de mês para a direita.",
        "9. Para mudar a categoria de uma única compra, digite por cima da fórmula escolhendo na lista; a regra continua valendo para as outras.",
        "Demonstração com dados de exemplo (cartão e compras fictícios). Abre no Excel, no LibreOffice e no Google Planilhas.",
    ]
    for i, s in enumerate(passos, start=1):
        cel = U.cell(row=i, column=1, value=s)
        cel.alignment = Alignment(wrap_text=True, vertical="top")
        if i == 1:
            cel.font = Font(bold=True, color="1565C0")
    wb.calculation.fullCalcOnLoad = True
    wb.save(AQUI / "planilha-fatura-cartao.xlsx")

    js = {"categorias": CATEGORIAS,
          "regras": [dict(k=k, est=e, cat=c) for k, e, c in REGRAS],
          "transacoes": [dict(data=x["data"].isoformat(), desc=x["desc"], valor=x["valor"]) for x in TR]}
    (AQUI / "dados.js").write_text("// gerado por _gerar.py (mesmos dados da planilha)\nvar DADOS=" +
                                   json.dumps(js, ensure_ascii=False) + ";\n", encoding="utf-8")


if __name__ == "__main__":
    gerar()
    print(len(TR), "transações")
    tot = {}
    for x in TR:
        tot[x["cat"]] = round(tot.get(x["cat"], 0) + x["valor"], 2)
    print(tot, round(sum(x["valor"] for x in TR), 2))
    print("sem categoria:", [x["desc"] for x in TR if x["cat"] == "Sem categoria"])
