import pandas as pd
import datetime

from functions import get_sheetnames_xlsx, table_maker, row_len

# lista dni tygodnia: po 2 kierunki na każdy dzień roboczy (R1, R2), soboty (S1, S2) i niedziele (N1,N2)
days = get_sheetnames_xlsx('Tab.xlsx')

# wczytanie tabeli nazw przystanków z uchwały
att = pd.read_excel('Temp/Uchw.xls', header=None, usecols='B,E', skiprows=3)

# wczytanie i modyfikacja tabel z godzinami odjazdów
tables = [table_maker(att,day) for day in days]

# połączenie tabel w jedną
df2 = pd.concat([t for t in tables]).reset_index(drop=True)

# eksport tabeli do excela
df2.to_excel('Zalacznik.xlsx', sheet_name = 'Arkusz1', engine = 'openpyxl')

# utworzenie listy kolumn z excela
letters = [chr(let+65) for let in range(26)]
exc_cols = [chr(i+65) if ord(chr(i+65)) < 91 else f'{letters[(i-26)//26]}{letters[(i-26)%26]}' for i in range(len(df2.columns))]

# ustawienie zmiennych
w1 = False
ns = 0
writer = pd.ExcelWriter('Zalacznik.xlsx', engine='xlsxwriter')
workbook = writer.book

# ustawienie formatowania
date_format = workbook.add_format({
    'num_format': 'HH:MM',
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Arial Narrow',
    'font_size': 10})
cell_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'font_name': 'Arial Narrow',
    'font_size': 10})

border_format = workbook.add_format({
    'border': 1})

busstop_format = workbook.add_format({
    'align': 'left',
    'valign': 'vcenter',
    'font_name': 'Arial Narrow',
    'font_size': 10})

title_format = workbook.add_format({
    'align': 'left',
    'valign': 'vcenter',
    'font_name': 'Arial Narrow',
    'font_size': 18})

header1_format = workbook.add_format({
    'font_name': 'Times New Roman',
    'font_size': 18})

header2_format = workbook.add_format({
    'font_name': 'Times New Roman',
    'font_size': 10})

# zapis tabeli w podanym arkuszu i zainicjowanie z niego zmiennej z tabelą
df2.to_excel(writer, sheet_name='Arkusz1', startcol=-1, header=False)
worksheet = writer.sheets['Arkusz1']

# ustawienie szerokości kolumn
worksheet.set_column(0, 0, 6, cell_format)
worksheet.set_column(1, 1, 40, busstop_format)
worksheet.set_column(f'C:{exc_cols[(len(df2.columns)) - 1]}', 4.14, cell_format)

# ustawienie daty
for c in range(len(df2.columns))[4:]:
    for w in range(len(df2.index)):
        if type(df2.iloc[w, c]) == datetime.time:
            worksheet.write_datetime(f'{exc_cols[c]}{w + 1}', df2.iloc[w, c], date_format)

# ustawienie obramowania
for r in range(len(df2[0])):
    if df2.iloc[r, 0] == 'nr słupka':
        ns = r
        w1 = True
        continue
    if type(df2.iloc[r, 0]) != float:
        if w1 == True:
            worksheet.conditional_format(f'A{ns + 1}:{exc_cols[row_len(df2, ns) - 1]}{r}',
                                         {'type': 'no_errors', 'format': border_format})
            w1 = False
worksheet.conditional_format(f'A{ns + 1}:{exc_cols[row_len(df2, ns) - 1]}{r + 1}',
                             {'type': 'no_errors', 'format': border_format})

# stylizacja tytułu
for r in range(len(df2[0])):
    if type(df2.iloc[r, 0]) == str and len(df2.iloc[r, 0]) > 10:
        worksheet.write(r, 0, df2.iloc[r, 0], title_format)

# ustawienie nagłówka i stopki wraz z ich formatowaniem
header1 = '&L&18&"Times New Roman,Bold"ROZKŁAD JAZDY LINII NR '
header2 = '&R&12&"Times New Roman,Bold"Rozkład jazdy ważny:\nod dnia\ndo dnia 31.12.2027'
header = header1 + header2
worksheet.set_header(header)
footer = '&C&10&"Times New Roman"&P/&N'
worksheet.set_footer(footer)

# ustawienie marginesów i orientacji kartki
worksheet.set_margins(top=1)
worksheet.set_landscape()

writer.save()