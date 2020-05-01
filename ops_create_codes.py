#!/usr/bin/python3.6
# A utility for printing attendance intake sheets for our
# good good friend L2F and speeding up the intake process
# for people who have pre-registered

import os
import sys

import barcode
from barcode.writer import ImageWriter

import openpyxl
from openpyxl import load_workbook

from file_iface import Menu

WRITER_OPTIONS = {'module_width':0.1,
                  'module_height':3,
                  'font_size':4,
                  'text_distance':1}
 
CODE128 = barcode.get_barcode_class('code128')

SOURCE = 'source_files/'
DESTINATION = 'bar_codes/'
NAME = 'test_source.xlsx'

def create_bc(val_str, CODECLASS):
    '''
    creates a bar code image and returns the str name of the file
    it names the file of the iimage in a png format
    '''
    code = CODECLASS(val_str, writer=ImageWriter())
    result = code.save(f'{DESTINATION}{val_str}', options=WRITER_OPTIONS)
    return result

def put_code(code_file, cell_str, ws_handl):
   '''
    takes a image file name, a cell location
    and a worksheet and saves the image to the worksheet at
    the cell location specified
   '''
   img = openpyxl.drawing.image.Image(code_file)
   img.anchor = cell_str # i.e. A1
   ws_handl.add_image(img)

def fnd_col_lttr(cell):
    '''
    openpyl repr for a cell is in the format of <Cell 'Sheet1'.A1>
    this function extracts the Column Letter A from that
    str
    '''
    trans_table = str.maketrans({'0': '', '1': '', '2': '', '3': '', '4': '',
                                 '5': '', '6': '', '7': '', '8': '', '9': ''})
    return str(cell).split('.')[1].strip('>').translate(trans_table)

def fnd_sub_str(rng, sub_string):
    '''
    iterates through the tuple structure of a row looking for the substring
    when it finds it it returns the column where the sub string is the header
    '''
    for item in rng:
        if sub_string in item.value:
            
            return fnd_col_lttr(item)
    return None

def file_set(target_dir='bar_codes/'):
    return set(x for x in os.listdir(target_dir) if x.endswith('.png'))

def connect_xl_file(fname):

    wb = load_workbook(fname)
    ws = wb['Sheet1']

    EOL = len(ws[1])
    ws.insert_cols(EOL)
    ws[1][EOL].value = 'Barcode'
    
    bc_col_letr = fnd_col_lttr(ws[1][EOL])
    cell_index = 2
    col = fnd_sub_str(ws[1], 'Client ID') # i.e. A
    ws.column_dimensions[bc_col_letr].width = 24
    return wb, ws, (cell_index, col, bc_col_letr)

def handle_xl_file(fname):
    '''
    opens an excel file and looks for the right column heading
    when it finds it, it iterates down the line and generates 
    bar codes for the file number and pastes it to the end of the headings on
    the right of the file

    '''
    wb, ws, dexs = connect_xl_file(fname)
    cell_index, col, bc_col_letr = dexs
    
    bar_code_files = file_set()

    for n in range(len(ws[col])):
        n_l = f'{col}{cell_index}' # i.e. A2
        cell_val = str(ws[n_l].value) # File ID string
        
        if cell_val != 'None':

            bars = f'{DESTINATION}{cell_val}.png'
            if bars.split('/')[1] not in bar_code_files:        
                bars = create_bc(cell_val, CODE128) # create barcode for File ID

            put_code(bars, f'{bc_col_letr}{cell_index}', ws)
            
            ws.row_dimensions[cell_index].height = 65
            cell_index += 1 # next time through we'll operate on A3

    wb.save(fname)
    
def main():
    print('ADDING CODES TO A WORKSHEET')
    menu = Menu(base_path=SOURCE)
    menu.get_file_list()
    target = menu.handle_input(menu.prompt_input('files'))
    
    confirm = input(f'please choose...\n1. {target}\n2. Exit\n')
    print('Please confirm your choice')
    if confirm == '1':    
        handle_xl_file(target)
    else:
        print('exiting...')
        sys.exit(1)
    print('we are done!')

if __name__ == '__main__':
    main()             
