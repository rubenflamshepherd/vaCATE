''' 
Change Log

08/07/14 -  Created module, inserted comments.

Current Roster:
generate_template (output_file, sheetname)
grab_data (directory, filename)
generate_workbook (directory, individual_inputs, series_inputs)
'''

import xlsxwriter
from xlrd import *
import math
import numpy
import time

def generate_template (output_file, sheetname):
    '''
    Generates a CATE template excel file (*.xlsx) with path/filename "output_file".
    Returns open worksheet, workbook as an ordered tuple pair
    WORKBOOK NEEDS TO BE CLOSED OUTSIDE OF THIS FUNCTION (in main program)
    '''
    workbook = xlsxwriter.Workbook(output_file)
    worksheet = workbook.add_worksheet(sheetname)
    
    # Formatting for header items for which inputs ARE NOT required
    not_req = workbook.add_format ()
    not_req.set_text_wrap ()
    not_req.set_align ('center')
    not_req.set_align ('vcenter')
    not_req.set_bottom ()
    
    # Formatting for items for which inputs ARE required
    req = workbook.add_format ()
    req.set_text_wrap ()
    req.set_align ('center')
    req.set_align ('vcenter')
    req.set_bold ()
    req.set_bottom ()
    
    # Formatting for rows cells that recieve input
    empty_row = workbook.add_format ()
    empty_row.set_align ('center')
    empty_row.set_align ('vcenter')
    empty_row.set_top ()    
    empty_row.set_bottom ()    
    empty_row.set_right ()    
    empty_row.set_left ()
    # Formatting for items for which inputs ARE required
    run_header = workbook.add_format ()    
    run_header.set_align ('center')
    run_header.set_align ('vcenter')    
    
    # Formatting for header items that are not visible to user
    # Items under header are for plotting regression line
    invis = workbook.add_format ()
    invis.set_font_color ('white')
    
    # Setting the height of the SA row to ~2 lines
    worksheet.set_row (1, 30.75)
    
    # List of ordered tuples contaning (header title, formatting, column width) 
    # in order theay are to be written to the file
    col_headers = [
        ("Vial #", not_req, 3.57),\
        ("Elution time (min)", req, 11.7),\
        ("Activity in eluant (cpm)", req, 15),\
        ("Activity in eluant (cpm)", req, 15)\
    ]
    row_headers = [
        (u"Specific Activity (cpm \u00B7 \u00B5mol\u207b\u00b9)", req, 14.45),\
        ("Root Cnts (cpm)", req, 8.5),\
        ("Shoot Cnts (cpm)", req, 9.7),\
        ("Root weight (g)", req, 11),\
        ("G-Factor", req, 8),\
        ("Load Time (min)", req, 9)\
    ]
    
    # Writing the row and column titles, setting the format, and setting the column width
    for x in range (0, len (col_headers)):
        worksheet.write (7, x, col_headers[x][0], col_headers[x][1])
        worksheet.set_column (x, x, col_headers [x][2])
    for y in range (0, len (row_headers)):
        worksheet.merge_range (y + 1, 0, y + 1, 1, row_headers[y][0], row_headers[y][1])
        worksheet.write (y + 1, 2, "", empty_row)
        worksheet.write (y + 1, 3, "", empty_row)
        
    # Writing headers for total runs
    worksheet.write (0, 2, "Run 1", run_header)
    worksheet.write (0, 3, "Run 2", run_header)
     
    return worksheet, workbook

def grab_data (directory, filename):
    '''
    Extracts data from an excel file directory/filename formated according 
    to generate_template
    
    Returns the data (SA, root_cnts, shoot_cnts, root_weight, g_factor,
    load_time, elution_times (list), elution_cpms(list)) in a list 
    '''
        
    # Formatign the directory (and path) to unicode w/ forward slash
    # so it can be passed between methods/classes w/o bugs
    directory = u'%s' %directory
    directory = directory.replace (u'\\', '/')
    
    # Accessing the file from which data is grabbed
    input_file = '/'.join ((directory, filename))
    input_book = open_workbook (input_file)
    input_sheet = input_book.sheet_by_index (0)
    
    # Create lists to store series of data
   
    elution_times = [0.0] # Elution times USED for caluclating cpms/g/hr
    elution_cpms = []
    raw_elution_times = input_sheet.col (1) # Elutions times given in file
    raw_cpm_column = input_sheet.col (2) # Raw counts given by file
        
    # Grab individual values (g-factor, SA, root weight root/shoot counts,etc.)
    SA = input_sheet.cell (1, 2).value
    root_cnts = input_sheet.cell (2, 2).value
    shoot_cnts = input_sheet.cell (3, 2).value
    root_weight = input_sheet.cell (4, 2).value
    g_factor = input_sheet.cell (5, 2).value
    load_time = input_sheet.cell (6, 2).value
            
    # Grabing elution times, correcting for header offset (7)
    for x in range (8, len (raw_elution_times)):                   
        elution_times.append (raw_elution_times [x].value)
    for x in range (8, len (raw_cpm_column)):                   
        elution_cpms.append (raw_cpm_column [x].value)    
   
    return SA, root_cnts, shoot_cnts, root_weight, g_factor, load_time, elution_times, elution_cpms

if __name__ == "__main__":
    #print grab_data("C:\Users\Ruben\Projects\CATEAnalysis", "CATE Analysis - (2014_11_20).xlsx")
    generate_template ("C:\Users\Ruben\Desktop\CATE_EXCEL_TEST.xlsx", \
                        "CATE Template")
               
    