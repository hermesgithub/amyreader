from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.cell import get_column_letter
import numpy
import matplotlib.pyplot as plt
import numpy as np
import pylab as p

wb = load_workbook(filename = r'empty_book.xlsx')
ws = wb.get_active_sheet()

dest_filename = input('What would you like to name the output file? ') + '.xlsx'
sheet_ranges = wb.get_sheet_by_name(name = 'range names')

mock_reading_raw = input('Which cell of the excel file is your original mock in? ')
mock_reading = (sheet_ranges.cell(mock_reading_raw).value)

assay_time_input = input('How many minutes was your assay for? ')
assay_time = int(assay_time_input)

dilution_cells_input = input('By what factor did you dilute the cells? ')
dilution_cells = int(dilution_cells_input)

sheet_ranges = wb.get_sheet_by_name(name = 'range names')

originlist = []
for row in ws.range('A1:C2'):
    for cell in row:
        originlist.append(cell.value)
print('\n')
print(originlist)
originlist_array = numpy.array(originlist)

minus_mock = originlist_array-mock_reading
print(minus_mock)
by_time1 = minus_mock/assay_time
print(by_time1)
by_dilution1 = by_time1*dilution_cells
print(by_time1)
by_30ul1 = by_dilution1/30
print(by_30ul1)
into_ml1 = by_30ul1*1000
print(into_ml1)


fig = p.figure()
ax = fig.add_subplot(1,1,1)
x = [1,2,3,4,5,6]
y = originlist_array
ax.bar(x,y,facecolor='#777777', align='center')
ax.set_ylabel('Secretion')
ax.set_title('Sample Secretions',fontstyle='italic')
ax.set_xticks(x)
group_labels = range(1-96)
ax.set_xticklabels(group_labels)
fig.autofmt_xdate()
ax.set_ylim([0,50])
p.show()



#ws = wb.create_sheet()
#ws.title = "PyOutput" #Now the new sheet is editable
#ws.range('A1:C2').value = startval
#wb.save(filename = dest_filename)
