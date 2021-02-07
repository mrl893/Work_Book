from tempfile import NamedTemporaryFile
from openpyxl import Workbook
wb = Workbook()

wb = load_workbook("balances.xlsx")
wb.template = True
wb.save('balances_template.xlsx')

with NamedTemporaryFile() as tmp:
    wb.save(tmp.name)
    tmp.seek(0)
    stream = tmp.read()


