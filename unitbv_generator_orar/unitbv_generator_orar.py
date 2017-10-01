import openpyxl
import sys
import xlsxwriter

__header_time = ('8:00 – 9:50', '10:00 – 11:50', '12:00 – 13:50', '14:00 – 15:50', '16:00 – 17:50', '18:00 – 19:50', '20:00 – 21:50')
__header_day = ('Luni', 'Marți', 'Miercuri', 'Joi', 'Vineri', 'Sâmbătă')

__bold_format = None

def get_mergedcell_value(source, coords):
    for range in source.merged_cell_ranges:
        merged_cells = list(openpyxl.utils.rows_from_range(range))
        for row in merged_cells:
            if coords in row:
                return source[merged_cells[0][0]].value
    return source[coords].value

def generate_worksheet(worksheet, source, row, version):
    # format pagina: ANSI E (44 inch x 34 inch) landscape
    worksheet.set_landscape()
    worksheet.set_paper(26)

    # inaltime header ora: 1.10 inch
    # latime header zi: 3 inch
    # inaltime camp: 5 inch (2.5 inch per rand, un camp fiind format din 2 randuri combinate pentru a permite afisarea materiilor din zile impare/pare)
    # latime camp: 5.62 inch (2.81 inch per coloana, aceeasi poveste ca mai sus, pentru a permite afisarea materiilor care se desfasoara in acelasi timp) 
    # inaltimea e in puncte (1 punct = 1/72 inch)
    # latimea e in numarul de caractere care incap in acel camp folosind fontul standard
    worksheet.set_column(0, len(__header_time) * 2, 36)
    worksheet.set_row(0, 79.2)
    for i in range(1, len(__header_day) * 2 + 1):
        worksheet.set_row(i, 180)

    # generam headerele orarului

    # interval: luni-sambata, 8:00-21:50
    # populam orele si legam casutele headerului 2 cate 2
    for i in range(0, len(__header_time)):
        worksheet.merge_range(0, 2 * i + 1, 0, 2 * i + 2, __header_time[i], __bold_format)
    # populam zilele si legam casutele headerului 2 cate 2
    for i in range(0, len(__header_day)):
        worksheet.merge_range(2 * i + 1, 0, 2 * i + 2, 0, __header_day[i], __bold_format)

    # populam campul A1 cu detalii importante
    # E1 = cod orar
    # J3 = an universitar
    worksheet.write_string(0, 0, str(get_mergedcell_value(source, 'E1')).replace(' ', '') + ' – ' + str(version).replace(' ', '') + '\n' + str(get_mergedcell_value(source, 'J3')).replace(' ', '') + '\n' + worksheet.get_name().replace('-', ' – '), __bold_format)

    # populam campurile cu continut


if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit('Mai putin de 3 argumente!\nFolosire: unitbv_generator_orar.py\n\tcale_orar\n\tfolder_output\n\tprimul_rand_al_grupei_de_generat\n\tprimul_rand_al_grupei_de_generat\n\tetc.')
    
    # setam sursa orarului
    source = openpyxl.load_workbook(sys.argv[1]).active
        
    # generam workbook-ul curent
    workbook = xlsxwriter.Workbook(sys.argv[2] + '\\orar-' + str(get_mergedcell_value(source, 'E1')).replace(' ', '') + '.xlsx')

    __bold_format = workbook.add_format({
                'font_name': 'Calibri',
                'font_size': 22,
                'bold': True,
                'align': 'center',
                'valign': 'vcenter',
                'border': 5,
                'border_color': '#9B9B9B'
            })
    
    # extragem versiunea orarului
    version = sys.argv[1].split('-')
    version = version[len(version) - 1].split('.')[0].replace(' ', '')
    # pentru fiecare rand oferit ca parametru generam un nou worksheet in
    # fisier
    for row in sys.argv[3:]:
        worksheet = workbook.add_worksheet(str(str(get_mergedcell_value(source, 'A' + str(row))) + '-' + str(get_mergedcell_value(source, 'B' + str(row))) + '-' + str(get_mergedcell_value(source, 'C' + str(row)))).replace(' ', ''))
        generate_worksheet(worksheet, source, row-1, version)
    
    # salvam fisierul
    workbook.close()
