import openpyxl
import sys
import xlsxwriter

# setari
__coord_cod_orar = 'E1'
__coord_an_universitar = 'J3'
__col_an = 'A'
__col_spec = 'B'
__col_grupa = 'C'
__col_inceput_cursuri = 'E'

# variabile
__header_time = ('8:00 – 9:50', '10:00 – 11:50', '12:00 – 13:50', '14:00 – 15:50', '16:00 – 17:50', '18:00 – 19:50', '20:00 – 21:50')
__header_day = ('Luni', 'Marți', 'Miercuri', 'Joi', 'Vineri', 'Sâmbătă')
__bold_format = None
__cell_format = None
__italic_format = None

# dictionar cu valori tuple
# {numeprescurtat: (nume intreg, culoare)}
__disciplines = {}
__ignored_disciplines = []
__professors = {}

# intoarce un array cu 4 intrari
# numele materiei
# C/(S)/[L]
# sala
# profesor
def transform_cell_value_in_array(value):
    split_value = str(value).replace(' ', '').split(',')
    if split_value[0] == 'S':
        split_value[0] = '(S)'
    elif split_value[0] == 'L':
        split_value[0] = '[L]'
    return split_value


def get_discipline_name(discipline):
    if discipline in __disciplines:
        return __disciplines[discipline][0]
    return discipline


def get_professor_name(professor):
    if professor in __professors:
        return __professors[professor][0]
    return professor


def get_discipline_color(discipline):
    if discipline in __disciplines:
        return __disciplines[discipline][1]
    return '#000000'


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
    worksheet.write_string(0, 0, str(get_mergedcell_value(source, __coord_cod_orar)).replace(' ', '') + ' – ' + str(version).replace(' ', '') + '\n' + str(get_mergedcell_value(source, __coord_an_universitar)).replace(' ', '') + '\n' + worksheet.get_name().replace('-', ' – '), __bold_format)

    # populam campurile cu continut
    # citim cate o coloana per pas (4 casute, 2 pt. fiecare saptamana)
    start_col = ord('E') - ord('A') + 1
    for day in range(0, len(__header_day)):
        for period in range(0, len(__header_time)):
            for col in source.iter_cols(start_col + period + day*len(__header_time), start_col + period + day*len(__header_time), row, row+3):
                first = col[0].value
                second = col[1].value
                third = col[2].value
                forth = col[3].value

                # daca sunt toate liniile goale, atunci nu avem o intrare
                if first == second == third == forth == None:
                    worksheet.merge_range(1 + day*2,
                                          1 + period*2, 
                                          1 + day*2 + 1, 
                                          1 + period*2 + 1, 
                                          '',
                                          __cell_format)

                # daca prima si a treia linie sunt egale, iar a doua si a patra sunt egale
                # inseamna ca avem o materii alternante
                elif first == third and second == forth:
                    if first != None:
                        transformed_value_first = transform_cell_value_in_array(first)
                        if transformed_value_first[2] not in __ignored_disciplines:
                            format = workbook.add_format({'font_name': 'Calibri',
                                            'font_size': 22,
                                            'bold': False,
                                            'italic': False,
                                            'align': 'center',
                                            'valign': 'vcenter',
                                            'border': 5,
                                            'border_color': '#9B9B9B',
                                            'bg_color': get_discipline_color(transformed_value_first[0])
                                           })
                            worksheet.merge_range(1 + day*2,
                                          1 + period*2, 
                                          1 + day*2, 
                                          1 + period*2 + 1,
                                          '',
                                          format
                                          )
                            worksheet.write_rich_string(
                                1 + day*2,
                                1 + period*2,
                                __bold_format,
                                transformed_value_first[1] + '\n',
                                __italic_format,
                                transformed_value_first[2] + '\n',
                                __bold_format,
                                get_discipline_name(transformed_value_first[0]) + '\n',
                                __cell_format,
                                get_professor_name(transformed_value_first[3])
                                )
                        else:
                            worksheet.merge_range(1 + day*2,
                                                1 + period*2, 
                                                1 + day*2, 
                                                1 + period*2 + 1,
                                                '',
                                                __cell_format)
                    else:
                        worksheet.merge_range(1 + day*2,
                                          1 + period*2, 
                                          1 + day*2, 
                                          1 + period*2 + 1,
                                          '',
                                          __cell_format)

                    if third != None:
                        transformed_value_third = transform_cell_value_in_array(third)
                        if transformed_value_third[2] not in __ignored_disciplines:
                            format = workbook.add_format({'font_name': 'Calibri',
                                            'font_size': 22,
                                            'bold': False,
                                            'italic': False,
                                            'align': 'center',
                                            'valign': 'vcenter',
                                            'border': 5,
                                            'border_color': '#9B9B9B',
                                            'bg_color': get_discipline_color(transformed_value_third[0])
                                           })
                            worksheet.merge_range(1 + day*2 + 1,
                                          1 + period*2, 
                                          1 + day*2 + 1, 
                                          1 + period*2 + 1,
                                          '',
                                          format
                                          )
                            worksheet.write_rich_string(
                                1 + day*2,
                                1 + period*2,
                                __bold_format,
                                transformed_value_third[1] + '\n',
                                __italic_format,
                                transformed_value_third[2] + '\n',
                                __bold_format,
                                get_discipline_name(transformed_value_third[0]) + '\n',
                                __cell_format,
                                get_professor_name(transformed_value_third[3])
                                )
                        else:
                            worksheet.merge_range(1 + day*2 + 1,
                                                1 + period*2, 
                                                1 + day*2 + 1, 
                                                1 + period*2 + 1,
                                                '',
                                                __cell_format)
                    else:
                        worksheet.merge_range(1 + day*2 + 1,
                                          1 + period*2, 
                                          1 + day*2 + 1, 
                                          1 + period*2 + 1,
                                          '',
                                          __cell_format)

                # daca nu sunt egale, inseamna ca unele din ele difera



if __name__ == "__main__":
    if len(sys.argv) < 3:
        sys.exit('Mai putin de 3 argumente!\nFolosire: unitbv_generator_orar.py\n\tcale_orar\n\tfolder_output\n\tprimul_rand_al_grupei_de_generat\n\tprimul_rand_al_grupei_de_generat\n\tetc.')
    
    # incarcam fisierul cu materii
    temp_file = open('materii.txt', 'r')
    lines = temp_file.readlines()
    for line in lines:
        data = line.split('=')
        __disciplines[data[0]] = (data[1], data[2].replace('\n', ''))
    temp_file.close()
            
    # incarcam fisierul cu profesori
    temp_file = open('profesori.txt', 'r')
    lines = temp_file.readlines()
    for line in lines:
        data = line.split('=')
        __professors[data[0]] = data[1].replace('\n', '')
    temp_file.close()

    # incarcam fisierul cu materii ignorate
    temp_file = open('materii_ignorate.txt', 'r')
    lines = temp_file.readlines()
    for line in lines:
         __ignored_disciplines.append(line.replace('\n', ''))
    temp_file.close()

    # setam sursa orarului
    source = openpyxl.load_workbook(sys.argv[1]).active
        
    # generam workbook-ul curent
    workbook = xlsxwriter.Workbook(sys.argv[2] + '\\orar-' + str(get_mergedcell_value(source, __coord_cod_orar)).replace(' ', '') + '.xlsx')

    __bold_format = workbook.add_format(
        {
            'font_name': 'Calibri',
            'font_size': 22,
            'bold': True,
            'italic': False,
            'align': 'center',
            'valign': 'vcenter',
            'border': 5,
            'border_color': '#9B9B9B'
        })

    __italic_format = workbook.add_format(
        {
            'font_name': 'Calibri',
            'font_size': 22,
            'bold': False,
            'italic': True,
            'align': 'center',
            'valign': 'vcenter',
            'border': 5,
            'border_color': '#9B9B9B'
        })

    __cell_format = workbook.add_format(
        {
            'font_name': 'Calibri',
            'font_size': 22,
            'bold': False,
            'italic': False,
            'align': 'center',
            'valign': 'vcenter',
            'border': 5,
            'border_color': '#9B9B9B'
        })
    
    # extragem versiunea orarului
    version = sys.argv[1].split('-')
    version = version[len(version) - 1].split('.')[0].replace(' ', '')

    # pentru fiecare rand oferit ca parametru generam un nou worksheet in fisier
    for row in sys.argv[3:]:
        worksheet = workbook.add_worksheet(str(str(get_mergedcell_value(source, __col_an + str(row))) + '-' + str(get_mergedcell_value(source, __col_spec + str(row))) + '-' + str(get_mergedcell_value(source, __col_grupa + str(row)))).replace(' ', ''))
        generate_worksheet(worksheet, source, int(row)-1, version)
    
    # salvam fisierul
    workbook.close()
