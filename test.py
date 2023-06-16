import xlwings as xw

app = xw.App(visible=True, add_book=False)
app.display_alerts = False
app.screen_updating = False

wb = app.books.open('1.xlsx')
sht = xw.sheets.active
rng = sht['A1:K17']
a = sht.range('A1:K17').value
print(a[1])
# ncols = rng.columns.count
# print(ncols)
# 切片，换取第一行
# fst_col = sheet[0,:ncols].value
# print(fst_col)

rng = sht.range('A1').expand('table')
nrows = rng.rows.count
print(nrows)

app.quit()

