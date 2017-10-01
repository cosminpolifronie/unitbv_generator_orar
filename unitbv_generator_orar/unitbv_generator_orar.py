from openpyxl import Workbook, worksheet
from openpyxl.styles import Alignment, Border, Font, NamedStyle, Side

# setari pt.  stilul celulelor
__fontname = 'Calibri'
__fontsize = 24
__fontcolor = 'FF000000'
__bordercolor = 'FF9B9B9B'
__borderstyle = 'thick'

__header_time = ('8:00 – 9:50', '10:00 – 11:50', '12:00 – 13:50', '14:00 – 15:50', '16:00 – 17:50', '18:00 – 19:50', '20:00 – 21:50')
__header_day = ('Luni', 'Marți', 'Miercuri', 'Joi', 'Vineri', 'Sâmbătă')

workbook = Workbook()
worksheet = workbook.active

cellstyle = NamedStyle(name='cellstyle', 
                       font = Font(name = __fontname, size = __fontsize, color = __fontcolor), 
                       border = Border(left = Side(style = __borderstyle, color = __bordercolor),
                           right = Side(style = __borderstyle, color = __bordercolor),
                           top = Side(style = __borderstyle, color = __bordercolor),
                           bottom = Side(style = __borderstyle, color = __bordercolor)),
                       alignment = Alignment(horizontal = 'center',
                           vertical = 'center'))

# format pagina: A0 (46.8 inch x 33.1 inch) landscape
worksheet.page_setup.fitToPage = True
worksheet.page_setup.orientation = 'landscape'
worksheet.page_setup.paperHeight = '46.8in'
worksheet.page_setup.paperWidth = '33.1in'

# inaltime header ora: 1.10 inch
# latime header zi: 3 inch
# inaltime camp: 5 inch (2.5 inch per rand, un camp fiind format din 2 randuri
# combinate pentru a permite afisarea materiilor din zile impare/pare)
# latime camp: 6 inch (3 inch per coloana, aceeasi poveste ca mai sus, pentru a
# permite afisarea materiilor care se desfasoara in acelasi timp)
# inaltimea e in puncte (1 punct = 1/72 inch)
# latimea e in numarul de caractere care incap in acel camp folosind fontul
# standard
worksheet.row_dimensions[1].height = 79.2
for i in range(2, len(__header_day) * 2 + 2):
    worksheet.row_dimensions[i].height = 180
for i in range(0, len(__header_time) * 2 + 1):
    worksheet.column_dimensions[chr(ord('A') + i)].width = 39

# generam headerele orarului
# interval: luni-sambata, 8:00-21:50
# populam orele si legam casutele headerului 2 cate 2
for i in range(0, len(__header_time)):
    _letter = chr(ord('B') + i * 2)
    worksheet[_letter + '1'] = __header_time[i]
    worksheet.merge_cells(_letter + '1:' + chr(ord(_letter) + 1) + '1')

# populam zilele si legam casutele headerului 2 cate 2
for i in range(1, len(__header_day) + 1):
    _number = str(i * 2)
    worksheet['A' + _number] = __header_day[i - 1]
    worksheet.merge_cells('A' + _number + ':A' + str(int(_number) + 1))

# setam stilurile celulelor
for i in range(1, len(__header_day) * 2 + 2):
    for j in range(1, len(__header_time) * 2 + 2):
        _letter = chr(ord('A') + j - 1)
        worksheet[_letter + str(i)].style = cellstyle

# citim grupa

# salvam fisierul
workbook.save('nume.xlsx')
