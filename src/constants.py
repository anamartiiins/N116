PATH_EXCEL = r"C:\Users\aimartins\OneDrive - Parfois, SA\Desktop\N116\development\Excel Exemplo N116 development.xlsm"

# Define the cells where some important values are
NR_PROJECT = "B6"
CLIENT_NAME = "C6"
LOCAL = "D6"
HEADER_START = "B10"
MARGIN_OBJECTIVE = "S8"
MARKUP = "AC8"
CONTRACTED_METTER_VALUE = "AD8"
HEADER_TRANSPORTS_RESUME = "AF10"

FORMULAS_MAPPING_ARTICLES = {
    "Preço Total": "=Preço Unitário{row} * Qtd{row}",
    "Preço Unitário": "=Custo Unitário{row} * $S$8",
    "Custo Unitário": "=SUM(Produção 1{row}:Material/Tecido 3{row})",
    "Custo Total": "=Custo Unitário{row} * Qtd{row}",
    "M": "=IF(Custo Total{row}=0,Custo Total,Preço Total{row}/Custo Total{row})",
    "Cubicagem direta": "=Cubicagem * Qtd{row} / 1000000",
    "Cubicagem c/ majoração": "=Cubicagem direta{row}*$AC$8",
    "Valor m3": "=Cubicagem c/ majoração{row}*$AD$8"
}

INPUT_CELL_TRANSPORT_TOTAL = "Transporte Total"