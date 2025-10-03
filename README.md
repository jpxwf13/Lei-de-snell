import openpyxl
from openpyxl import Workbook
from openpyxl.chart import ScatterChart, Series, Reference

def criar_planilha_lei_de_snell(nome_arquivo="Lei_de_Snell.xlsx"):
    # Criar planilha
    wb = Workbook()
    ws = wb.active
    ws.title = "Lei de Snell"

    # Cabeçalhos
    headers = ["Meio 1", "Meio 2", "n1", "n2", "θ1 (graus)", "θ2 (graus)", "Status"]
    ws.append(headers)

    # Exemplos iniciais
    dados = [
        ["Vidro", "Água", 1.50, 1.33, 30],
        ["Água", "Ar", 1.33, 1.00, 45]
    ]
    for row in dados:
        ws.append(row)

    # Inserir fórmulas em português (para Excel BR)
    for row in range(2, 4):  # linhas com exemplos
        ws[f"F{row}"] = f'=SE(ABS(C{row}*SEN(RADIANOS(E{row}))/D{row})>1;"Reflexão Total";GRAUS(ASEN(C{row}*SEN(RADIANOS(E{row}))/D{row})))'
        ws[f"G{row}"] = f'=SE(F{row}="Reflexão Total";"Reflexão Total";"Refração Normal")'

    # Criar gráfico de dispersão (θ1 vs θ2)
    chart = ScatterChart()
    chart.title = "Lei de Snell - θ1 x θ2"
    chart.x_axis.title = "Ângulo de Incidência (θ1) [graus]"
    chart.y_axis.title = "Ângulo de Refração (θ2) [graus]"

    xvalues = Reference(ws, min_col=5, min_row=2, max_row=3)  # θ1
    yvalues = Reference(ws, min_col=6, min_row=2, max_row=3)  # θ2
    series = Series(yvalues, xvalues, title="Refração")
    chart.series.append(series)

    # Inserir gráfico
    ws.add_chart(chart, "I2")

    # Salvar arquivo
    wb.save(nome_arquivo)
    print(f"📂 Planilha gerada com sucesso: {nome_arquivo}")

if __name__ == "__main__":
    criar_planilha_lei_de_snell("Lei_de_Snell_Com_Grafico.xlsx")
