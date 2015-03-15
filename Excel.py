import xlsxwriter
from xlrd import *
import math
import numpy
import time

def generate_template (output_file_path, workbook, worksheet):
    '''
    Generates an our CATE Data template in an in an already created workbook and
    worksheet.
    
    INPUTS
    output_file_path - path/filename (str)
    
    OUTPUT
    worksheet, workbook as an ordered tuple pair
    '''
    
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
    
    # Formatting for row cells that are to recieve input
    empty_row = workbook.add_format ()
    empty_row.set_align ('center')
    empty_row.set_align ('vcenter')
    empty_row.set_top ()    
    empty_row.set_bottom ()    
    empty_row.set_right ()    
    empty_row.set_left ()
    
    # Formatting for run headers ("Run x")
    run_header = workbook.add_format ()    
    run_header.set_align ('center')
    run_header.set_align ('vcenter')    
    
    # Setting the height of the SA row to ~2 lines
    worksheet.set_row (1, 30.75)
    
    # Lists of ordered tuples contaning (title, formatting, column width) 
    # in order theay are to be written to the file
    col_headers = [
        ("Vial #", not_req, 3.57),\
        ("Elution time (min)", req, 11.7),\
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
    
    # Writing the row and column titles, setting the format and column width
    for x in range (0, len (col_headers)):
        worksheet.write (7, x, col_headers[x][0], col_headers[x][1])
        worksheet.set_column (x, x, col_headers [x][2])
    for y in range (0, len (row_headers)):
        worksheet.merge_range (y + 1, 0, y + 1, 1,\
                               row_headers[y][0], row_headers[y][1])
        worksheet.write (y + 1, 2, "", empty_row)
                
    # Writing headers columns containing individual runs 
    worksheet.write (0, 2, "Run 1", run_header)
    
    #workbook.close()
     
    return worksheet, workbook # Why do we even return this?!!!!!!!!!!!!?!?!?!?!?!!?

def grab_data (directory, filename):
    '''
    Extracts data from an excel file in directory/filename (INPUT) formated according 
    to generate_template
    
    OUTPUT
    Data (SA, root_cnts, shoot_cnts, root_weight, g_factor,
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
    
    # List where allruns info is stored with ind. runs as list entries
    all_runs = []
    
    # Creating elution time point list to be used by all runs 
    raw_elution_times = input_sheet.col (1) # Col w/elution times given in file
    elution_times = [0.0] # Elution times USED for caluclating cpms/g/hr
    # Parsing elution times, correcting for header offset (8)
    for x in range (8, len (raw_elution_times)):                   
        elution_times.append (raw_elution_times [x].value)    
    
    print input_sheet.row_len (0)
    
    for row_index in range (2, input_sheet.row_len (0)):
    
        # Create lists to store series of data from ind. run
        raw_cpm_column = input_sheet.col (row_index) # Raw counts given by file
        elution_cpms = []
            
        # Grab individual CATE values of interest
        run_name = input_sheet.cell (0, row_index).value
        SA = input_sheet.cell (1, row_index).value
        root_cnts = input_sheet.cell (2, row_index).value
        shoot_cnts = input_sheet.cell (3, row_index).value
        root_weight = input_sheet.cell (4, row_index).value
        g_factor = input_sheet.cell (5, row_index).value
        load_time = input_sheet.cell (6, row_index).value
                
        # Grabing elution cpms, correcting for header offset (8)
        for x in range (8, len (raw_cpm_column)):                   
            elution_cpms.append (raw_cpm_column [x].value)
        
        all_runs.append ([run_name, SA, root_cnts, shoot_cnts, root_weight,\
                          g_factor, load_time, elution_times, elution_cpms])
   
    return all_runs

def generate_analysis (workbook, worksheet, frame_object):
    '''
    Creating an excel file in directory using a preset naming convention
    Data in the file are the product of CATE analysis from a template file containing the raw information
    Nothing is returned
    '''
    
    # Formatting for items for headers for analyzed CATE data
    analyzed_header = workbook.add_format ()
    analyzed_header.set_text_wrap ()
    analyzed_header.set_align ('center')
    analyzed_header.set_align ('vcenter')
    analyzed_header.set_bottom ()    
    
    # Formatting for cells that contains basic CATE data    
    empty_row = workbook.add_format ()
    empty_row.set_align ('center')
    empty_row.set_align ('vcenter')
    empty_row.set_top ()    
    empty_row.set_bottom ()    
    empty_row.set_right ()    
    empty_row.set_left ()       
    
    # Header info for analyzed CATE data
    headers = [("Corrected AIE (cpm)", 15),\
               (u"Efflux (cpm \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9)", 13.57),\
               ("Log Efflux", 8.86),\
               (u"R\u00b2", 7),\
               (u"Slope (min\u207b\u00b9)", 8.14),\
               ("Intercept", 8.5)]    
    
    for y in range (0, len (headers)):
            worksheet.write (7, y + 3, headers[y][0], analyzed_header)   
            worksheet.set_column (y + 3, y + 3,headers [y][1])
    
    
    # Writing CATE data to file
    
    worksheet.write (0, 2, frame_object.run_name)    
    worksheet.write (1, 2, frame_object.SA, empty_row)    
    worksheet.write (2, 2, frame_object.root_cnts, empty_row)
    worksheet.write (3, 2, frame_object.shoot_cnts, empty_row)
    worksheet.write (4, 2, frame_object.root_weight, empty_row)
    worksheet.write (5, 2, frame_object.gfactor, empty_row)
    worksheet.write (6, 2, frame_object.load_time, empty_row)

    p1_regression_counter = len (frame_object.elution_ends)\
        - len (frame_object.r2s_p3_list) - frame_object.num_points_obj
    
    for x in range (0, len (frame_object.elution_ends)):
        worksheet.write (8 + x, 0, x + 1)
        worksheet.write (8 + x, 1, frame_object.elution_ends [x])
        worksheet.write (8 + x, 2, frame_object.elution_cpms [x])
        worksheet.write (8 + x, 3, frame_object.elution_cpms_gfactor [x])
        worksheet.write (8 + x, 4, frame_object.elution_cpms_gRFW [x])
        worksheet.write (8 + x, 5, frame_object.elution_cpms_log [x])

    for y in range (0, len (frame_object.r2s_p3_list)):
        worksheet.write (9 + p1_regression_counter + y, 6, frame_object.r2s_p3_list [y])
        worksheet.write (9 + p1_regression_counter + y, 7, frame_object.slopes_p3_list [y])
        worksheet.write (9 + p1_regression_counter + y, 8, frame_object.intercepts_p3_list [y])
     
if __name__ == "__main__":
    print grab_data("C:\Users\Daniel\Projects\CATEAnalysis", "CATE Analysis - (2014_11_21).xlsx")
    #generate_template ("C:\Users\Ruben\Desktop\CATE_EXCEL_TEST.xlsx")
               
    