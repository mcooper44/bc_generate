# A utility for printing attendance intake sheets for our
# good good friend L2F and speeding up the intake process
# for people who have pre-registered

import barcode
from barcode.writer import ImageWriter

import openpyxl
from openpyxl import load_workbook

WRITER_OPTIONS = {'module_width':0.1,
                  'module_height':3,
                  'font_size':4,
                  'text_distance':1}
 
CODE128 = barcode.get_barcode_class('code128')

SOURCE = 'source_files/'
DESTINATION = 'bar_codes/'
NAME = 'test_source.xlsx'




def create_bc(val_str, CODECLASS):
    code = CODECLASS(val_str, writer=ImageWriter())
    result = code.save(val_str, options=WRITER_OPTIONS)
    return result

def put_code(code_file, cell_str, ws_handl):
   img = openpyxl.drawing.image.Image(code_file)
   img.anchor = cell_str # i.e. A1
   ws_handl.add_image(img)

def fnd_col_lttr(cell):
    trans_table = str.maketrans({'0': '', '1': '', '2': '', '3': '', '4': '',
                                 '5': '', '6': '', '7': '', '8': '', '9': ''})
    return str(cell).split('.')[1].strip('>').translate(trans_table)

def fnd_sub_str(rng, sub_string):
    for item in rng:
        if sub_string in item.value:
            
            return fnd_col_lttr(item)
    return None



def handle_xl_file(fname):
    wb = load_workbook(fname)
    ws = wb['Sheet1']

    EOL = len(ws[1])
    ws.insert_cols(EOL)
    ws[1][EOL].value = 'Barcode'
    bc_col_letr = fnd_col_lttr(ws[1][EOL])
    cell_index = 2
    col = fnd_sub_str(ws[1], 'Client ID') # i.e. A

    for n in range(len(ws[col])):
        n_l = f'{col}{cell_index}' # i.e. A2
        cell_val = str(ws[n_l].value) # File ID string
        bars = create_bc(cell_val, CODE128) # create barcode for File ID
        put_code(bars, f'{bc_col_letr}{cell_index}', ws)
        cell_index += 1

    wb.save(fname)
    
def main():
    print('giving it a shot')
    handle_xl_file(f'{SOURCE}{NAME}')
                   
main()              
