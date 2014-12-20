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
    
    # Formatting for header items for which inputs ARE required
    req = workbook.add_format ()
    req.set_text_wrap ()
    req.set_align ('center')
    req.set_align ('vcenter')
    req.set_bold ()
    req.set_bottom ()
    
    # Formatting for header items that are not visible to user
    # Items under header are for plotting regression line
    invis = workbook.add_format ()
    invis.set_font_color ('white')
    
    # Setting the height of the row to ~2 lines
    worksheet.set_row (0, 30.75)
    
    # List of ordered tuples contaning (header title, formatting, column width) 
    # in order theay are to be written to the file
    headers = [("Vial #", not_req, 3.57),\
               ("Elution time (min)", req, 11.7),\
               ("Activity in eluant (cpm)", req, 15),\
               ("Corrected AIE (cpm)", not_req, 15),\
               # \uXXXX indicates unicode character within unicode string (u"...")
               (u"Efflux (cpm \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9)", not_req, 13.57),\
               ("Log Efflux", not_req, 8.86),\
               (u"R\u00b2", not_req, 7),\
               ("Half-life (min)", not_req, 8),\
               (u"Slope (min\u207b\u00b9)", not_req, 8.14),\
               ("Intercept", not_req, 8.5),\
               (u"Specific Activity (cpm \u00B7 \u00B5mol\u207b\u00b9)", req, 14.45),\
               ("Root Cnts (cpm)", req, 8.5),\
               ("Shoot Cnts (cpm)", req, 9.7),\
               ("Root weight (g)", req, 11),\
               ("G-Factor", req, 8),\
               ("# Points for Regression", req, 10.85),\
               ("Load Time (min)", req, 9),\
               ("Reg x", invis, 9),\
               ("Reg y", invis, 9)]
    
    # Writing the respective headers, setting the format, and setting the column width
    for x in range (0, len (headers)):
        worksheet.set_column (x, x, headers [x][2])
        worksheet.write (0,x, headers[x][0], headers[x][1])
    
    # Writting default number of regression points and load time
    worksheet.write (1,16, 4)
    worksheet.write (1,16, 60)
    
    return worksheet, workbook

def grab_data (directory, filename):
    '''
    Extracts data from an excel file directory/filename formated according to generate_template
    Returns the data in a tuple pair with items containing lists of individual values
    and series of values
    '''
        
    # Formatign the directory (and path) to unicode w/ forward slash so it can be passed between methods/classes w/o bugs
    directory = u'%s' %directory
    directory = directory.replace (u'\\', '/')
    
    # Accessing the file from which data is grabbed
    input_file = '/'.join ((directory, filename))
    input_book = open_workbook (input_file)
    input_sheet = input_book.sheet_by_index (0)
    
    # Create lists to store series of data
   
    elution_times = [0.0] # Elution times USED for caluclating cpms/g/hr
    raw_elution_times = input_sheet.col (1) # Elutions times given in file
    corrected_cpms = [] # CPMS corrected by g-factor
    raw_cpms = input_sheet.col (2) # Raw counts given by file
    efflux = [] # Calculated efflux values
    log_efflux = [] # Log of calculated efflux values
    
    # Grab individual values (g-factor, SA, root weight root/shoot counts,etc.)
    SA = input_sheet.cell (1, 10).value
    root_cnts = input_sheet.cell (1, 11).value
    shoot_cnts = input_sheet.cell (1, 12).value
    root_weight = input_sheet.cell (1, 13).value
    g_factor = input_sheet.cell (1, 14).value
    num_points_reg = input_sheet.cell (1, 15).value
    load_time = input_sheet.cell (1, 16).value
    
    # List of individual values that is to be returned
    individual_inputs = [SA, root_cnts, shoot_cnts, root_weight, g_factor, num_points_reg, load_time]
    
    # Grab series of values (efflux cpms, elution times, etc)
    for x in range (1, len (raw_elution_times)):                   
        elution_times.append (raw_elution_times [x].value)
    for x in range (1, len (raw_cpms)):               
        corrected_cpms.append (raw_cpms [x].value * g_factor)
    for x in range (0, len (corrected_cpms)):
        temp = corrected_cpms [x]/root_weight/(elution_times[x+1]-elution_times[x])
        efflux.append (temp)
        log_efflux.append (math.log10 (temp))
    
    # List of series of values that is to be returned    
    series_inputs = [elution_times, corrected_cpms, efflux, log_efflux, raw_cpms]
    
    return individual_inputs, series_inputs
    
def generate_workbook (directory, individual_inputs, series_inputs):
    '''
    Creating an excel file in directory using a preset naming convention
    Data in the file are the product of CATE analysis from a template file containing the raw information
    Nothing is returned
    '''
    
    # Mapping list values to variables containing individual values
    SA = individual_inputs [0]
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
    output_name = 'CATE Analysis - ' + time.strftime ("(%Y_%m_%d).xlsx")
    output_file = '/'.join ((directory, output_name))
    output_sheet, output_book = generate_template (output_file, 'Run 1')
    
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