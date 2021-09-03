import openpyxl
from openpyxl.chart import BarChart, Reference, Series, ScatterChart
from openpyxl import Workbook
import random

file = openpyxl.load_workbook('Excel.xlsx')
sheet = file['Hoja1']

sheet['A1'] = 'Temp'
sheet['B1'] = 'Time'

multiple_cells = sheet['A1':'B3']

for i in range(10):
    sheet[f'A{i+2}'] = random.randrange(25,40)
#    sheet[f'B{i+2}'] = i <- this should be replaced with proper timestamp info, but i did int from the excel file itselft.

graph = ScatterChart()
graph.y_axis.title = 'Temp'
graph.x_axis.title = 'Time'
graph.scatterStyle
graph.style = 13
ref = Reference(sheet, min_col=1, min_row=2, max_row=10)
ref2 = Reference(sheet, min_col=2, min_row=2, max_row=10)
serie = Series(ref, ref2, title = 'OwO')
graph.series.append(serie)
sheet.add_chart(graph, 'E1')

file.save('Excel.xlsx')
print('Valores Guardados Exitosamente')