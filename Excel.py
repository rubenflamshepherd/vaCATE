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
    
    # Create lists to store series of data   
    raw_elution_times = input_sheet.col (1) # Elutions times given in file
    elution_times = [0.0] # Elution times USED for caluclating cpms/g/hr
    raw_cpm_column = input_sheet.col (2) # Raw counts given by file
    elution_cpms = []
        
    # Grab individual CATE values of interest
    run_name = input_sheet.cell (0, 2).value
    SA = input_sheet.cell (1, 2).value
    root_cnts = input_sheet.cell (2, 2).value
    shoot_cnts = input_sheet.cell (3, 2).value
    root_weight = input_sheet.cell (4, 2).value
    g_factor = input_sheet.cell (5, 2).value
    load_time = input_sheet.cell (6, 2).value
            
    # Grabing elution times, correcting for header offset (8)
    for x in range (8, len (raw_elution_times)):                   
        elution_times.append (raw_elution_times [x].value)
    for x in range (8, len (raw_cpm_column)):                   
        elution_cpms.append (raw_cpm_column [x].value)    
   
    return [run_name, SA, root_cnts, shoot_cnts, root_weight, g_factor, load_time, elution_times, elution_cpms]

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
        worksheet.write (8 + p1_regression_counter + y, 6, frame_object.r2s_p3_list [y])
        
    '''
    root_cnts = individual_inputs [1]
    shoot_cnts = individual_inputs [2]
    root_weight = individual_inputs [3]
    g_factor = individual_inputs [4]
    num_points_reg = individual_inputs [5]
    load_time = individual_inputs [6]
   
    # Mapping list values to lists containing series values
    elution_times = series_inputs [0]
    corrected_cpms = series_inputs [1]
    efflux = series_inputs [2]
    log_efflux = series_inputs [3]
    raw_cpms = series_inputs [4]
    
    # Creating the file
    output_name = 'CATE Analysis - ' + time.strftime ("(%Y_%m_%d_%I:%M).xlsx")
    output_file_path = '/'.join ((directory, output_name))
    output_sheet, output_book = generate_template (output_file_path, 'Run 1')
    
    # Writing the series data to the file
    for x in range (1, len (elution_times)):
        output_sheet.write (x, 0, x)
        output_sheet.write (x, 1, elution_times [x])
        output_sheet.write (x, 2, raw_cpms [x].value)
        output_sheet.write (x, 3, corrected_cpms [x - 1])
        output_sheet.write (x, 4, efflux [x - 1])
        output_sheet.write (x, 5, log_efflux [x - 1])
        
    # Writing indidual data values to the file (SA, root/shoot counts, etc.)
    output_sheet.write (1, 10, SA)
    output_sheet.write (1, 11, root_cnts)
    output_sheet.write (1, 12, shoot_cnts)
    output_sheet.write (1, 13, root_weight)
    output_sheet.write (1, 14, g_factor)
    output_sheet.write (1, 15, num_points_reg)
    output_sheet.write (1, 16, load_time)
                
    # Figuring out R^2 (coefficient of determination) and parameters of linear regression
    r2 = []
    slope = []
    intercept = []
    # Dynamically doing regression using a set number of points that whose number can be specificed (implement later)
    start_index = int(len (log_efflux) - num_points_reg)
    print_row = start_index + 1 # correcting for row number due to header
    reg_dec = False # Variable to track if regression should continue
    
    while reg_dec == False: # R^2 has not decreased 4 consecutive times
        x = elution_times [start_index + 1:]
        y = log_efflux [start_index:]
        
        # Linear regression of current set of values, returns m, b of y=mx+b
        coeffs = numpy.polyfit (x, y, 1)
        
        p = numpy.poly1d(coeffs) # Conversion to "convenience class" in numpy for working with polynomials        
        
        # Determining R^2 on the current data set        
        # fit values, and mean
        yhat = p(x)                         # or [p(z) for z in x]
        ybar = numpy.sum(y)/len(y)          # or sum(y)/len(y)
        ssreg = numpy.sum((yhat-ybar)**2)   # or sum([ (yihat - ybar)**2 for yihat in yhat])
        sstot = numpy.sum((y - ybar)**2)    # or sum([ (yi - ybar)**2 for yi in y])
        
        r2.insert (0, ssreg/sstot)
        slope.insert (0, coeffs [0])
        intercept.insert (0, coeffs [1])
        
        # Printing regression and correlation values to sheet
        output_sheet.write (print_row, 6, r2 [0])
        output_sheet.write (print_row, 8, coeffs [0]) # x of mx+b
        output_sheet.write (print_row, 9, coeffs [1]) # b of mx+b
        
        # Extending the series of values included in the analysis/regression
        start_index -= 1
        print_row -= 1
        
        # Checking to see if you are out of the third phase of exchange
        # (has R^2 decreased 3 consecutive times)
        if len (r2) <= 3:
            pass
        elif r2[0] < r2[1] and r2[1] < r2[2] and r2[2] < r2[3]:
            reg_dec = True
    
    # Write points for graphing equation of linear regression
    last_elution = elution_times[len(elution_times) - 1]
    regression_x1 = 0
    regression_x2 = last_elution
    regression_y1 = intercept [3]
    regression_y2 = slope [3] * last_elution + intercept [3]
    
    # Format for values of linear regression (not to be seen/played with by the user)
    invis = output_book.add_format ()
    invis.set_font_color ('white')            
    
    # Writing values for the linear regression series
    output_sheet.write (1, 17, regression_x1, invis)
    output_sheet.write (2, 17, regression_x2, invis)
    output_sheet.write (1, 18, regression_y1, invis)
    output_sheet.write (2, 18, regression_y2, invis)
    
    # Creating the chart
    chart = output_book.add_chart ({'type': 'scatter'})
    chart.set_x_axis ({'name': 'Elution time (min)'})
    chart.set_y_axis ({'name': u"Log cpm released \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9"})            
    chart.set_title ({'name': "Run 1 - Efflux over time"})                        
    chart.set_legend ({'none': True})   
    chart.set_size ({'width':900, 'height': 550})
    
    # Adding log elution cpms
    chart.add_series ({
        'categories': ['Run 1', 1, 1, len (log_efflux), 1],\
        'values': ['Run 1', 1, 5, len (log_efflux), 5],\
        'line': {'none': True}
    })
        
    # Adding regression line for third phase
    chart.add_series ({
        'categories': ['Run 1', 1, 17, 2, 17],\
        'values': ['Run 1', 1, 18, 2, 18],\
        'marker': {'type': 'none'},\
        'trendline': {'type': 'linear'}
    })
    
    # Adding last point used for regression
    chart.add_series({
        'categories': ['Run 1', (print_row + 4), 1, (print_row + 4), 1],\
        'values': ['Run 1', (print_row + 4), 5, (print_row + 4), 5],\
        'marker': {
            'type': 'circle',\
            'border': {'color': 'red'},\
            'fill': {'none': True}
            
        }
    })
    
    output_sheet.insert_chart ('A3', chart)            
    
    output_book.close()
    '''

if __name__ == "__main__":
    print grab_data("C:\Users\Daniel\Projects\CATEAnalysis", "CATE Analysis - (2014_11_21).xlsx")
    #generate_template ("C:\Users\Ruben\Desktop\CATE_EXCEL_TEST.xlsx")
               
    