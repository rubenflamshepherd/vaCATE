import xlsxwriter
from xlrd import *
import time

import DataObject
import RunObject

def generate_sheet(workbook, sheet_name):
    """
    Generates and returns the basic excel sheet template (in existing workbook)
    upon which all subsequent sheets (which the exception of summary sheets) are
    built
    """
    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.set_row(1, 30.75) # Setting the height of the SA row to ~2 lines
    
    # Formatting for header items for which inputs ARE NOT required
    not_req = workbook.add_format()
    not_req.set_text_wrap()
    not_req.set_align('center')
    not_req.set_align('vcenter')
    not_req.set_bottom()    
    # Formatting for items for which inputs ARE required
    req = workbook.add_format()
    req.set_text_wrap()
    req.set_align('center')
    req.set_align('vcenter')
    req.set_bold()
    req.set_bottom()    
    # Formatting for row cells that are to recieve input
    empty_row = workbook.add_format()
    empty_row.set_align('center')
    empty_row.set_align('vcenter')
    empty_row.set_top()    
    empty_row.set_bottom()    
    empty_row.set_right()    
    empty_row.set_left()
       
    # Lists of ordered tuples contaning (title, formatting, column width) 
    # in order they are to be written to the file
    col_headers = [
        ("Vial #", not_req, 3.57), ("Elution time (min)", req, 11.7),
        ("Activity in eluant (cpm)", req, 15)]
    row_headers = [
        (u"Specific Activity (cpm \u00B7 \u00B5mol\u207b\u00b9)", req, 14.45),
        ("Root Cnts (cpm)", req, 8.5), ("Shoot Cnts (cpm)", req, 9.7),
        ("Root weight (g)", req, 11), ("G-Factor", req, 8), 
        ("Load Time (min)", req, 9)]
    
    # Writing the row and column titles, setting the format and column width
    for x in range(0, len (col_headers)):
        worksheet.write(7, x, col_headers[x][0], col_headers[x][1])
        worksheet.set_column(x, x, col_headers [x][2])
    for y in range(0, len (row_headers)):
        worksheet.merge_range (
            y + 1, 0, y + 1, 1, row_headers[y][0], row_headers[y][1])
        worksheet.write (y + 1, 2, "", empty_row)
        
    return worksheet

def generate_template(workbook):
    '''
    Generates a CATE template sheet in an already created workbook. No need to
    return the sheet as nothing further is done after this.
    '''    
    worksheet = generate_sheet(workbook, 'Template')               

    # Formatting for run headers ("Run x")
    run_header = workbook.add_format()    
    run_header.set_align('center')
    run_header.set_align('vcenter')        

    # Writing headers columns containing individual runs 
    worksheet.write(0, 2, "Run 1", run_header)
    worksheet.write(0, 3, "etc.", run_header)
        
def grab_data(directory, filename):
    '''
    Extracts data from an excel file in directory/filename (INPUT) formated
    according to generate_sheet/generate_template    
    OUTPUT: DataObject[run_name SA, root_cnts, shoot_cnts, root_weight,
                        g_factor, load_time, [elution_times], [elution_cpms]]
    '''
    # Accessing the file from which data is to be grabbed
    input_file = '/'.join((directory, filename))
    input_book = open_workbook(input_file)
    input_sheet = input_book.sheet_by_index(0)

    # List where all run info is stored with RunObjects as ind. entries
    all_run_objects = []

    # Creating elution time point list to be used by all runs 
    raw_elution_times = input_sheet.col(1) # Col w/elution times given in file
    elution_times = [0.0] # Elution times TO BE used for caluclating cpms/g/h
    
    # Parsing elution times, correcting for header offset (8)
    for x in range(8, len(raw_elution_times)):                   
        elution_times.append(raw_elution_times[x].value)    
    
    for col_index in range(2, input_sheet.row_len(0)):        
        # Grab individual CATE values of interest
        run_name = input_sheet.cell(0, col_index).value
        SA = input_sheet.cell(1, col_index).value
        root_cnts = input_sheet.cell(2, col_index).value
        shoot_cnts = input_sheet.cell(3, col_index).value
        root_weight = input_sheet.cell(4, col_index).value
        g_factor = input_sheet.cell(5, col_index).value
        load_time = input_sheet.cell(6, col_index).value
        # Create lists to store SERIES' of data from ind. run
        raw_cpm_column = input_sheet.col(col_index) # Raw counts given by file
        elution_cpms = []
        # Grabing elution cpms, correcting for header offset (8)
        for x in range(8, len(raw_cpm_column)):                   
            elution_cpms.append(raw_cpm_column[x].value)
        
        all_run_objects.append(
            RunObject.RunObject(
                run_name, SA, root_cnts, shoot_cnts, root_weight, g_factor,
                load_time, elution_times, elution_cpms, ("obj", 3)))
   
    return DataObject.DataObject(directory, all_run_objects)

def generate_summary(workbook, data_object):
    '''
    Given an open workbook, create a summary sheet containing relavent data
    from all run_objects in data_object.run_objects
    '''
    worksheet = workbook.add_worksheet("Summary")
    worksheet.freeze_panes(1,2)
    
    # Formatting for items 
    req = workbook.add_format()
    req.set_text_wrap()
    req.set_align('right')
    req.set_align('vcenter')
    req.set_bold()
    req.set_right()
    
    bot_line = workbook.add_format()
    bot_line.set_bottom()        
    phase_format = workbook.add_format()
    phase_format.set_align('right')
    phase_format.set_right()    
    phase_format.set_top()    
    phase_format.set_bottom()    
    middle_format = workbook.add_format()
    middle_format.set_align('center')    
    right_format = workbook.add_format()
    right_format.set_align('right')        
    # Row labels
    row_headers = [
            "Analysis Type",
            u"Specific Activity (cpm \u00B7 \u00B5mol\u207b\u00b9)",
            "Root Cnts (cpm)", "Shoot Cnts (cpm)", "Root weight (g)",
            "G-Factor", "Load Time (min)", "", "Influx", "Net flux",
            "E:I Ratio", "Pool Size"]
    
    phasedata_headers = [
        "Slope", "Intercept", u"R\u00b2", "k", "Half-Life", "Efflux"]    
    
    # Writing row Labels
    
    for y in range (0, len (row_headers)):
        if y != 7:
            worksheet.merge_range (y + 1, 0, y + 1, 1, row_headers [y], req)

    
    for z in range (0, len (phasedata_headers)):
        worksheet.merge_range (z + 13, 0, z + 13, 1,\
                               phasedata_headers [z], req)
        worksheet.merge_range (z + 20, 0, z + 20, 1,\
                               phasedata_headers [z], req)        
        worksheet.merge_range (z + 27, 0, z + 27, 1,\
                               phasedata_headers [z], req)
    
    worksheet.merge_range (8, 0, 8, 1, "Phase III", phase_format)
    worksheet.merge_range (19, 0, 19, 1, "Phase II", phase_format)
    worksheet.merge_range (26, 0, 26, 1, "Phase I", phase_format)
        
    # Writing elution time points/headers for respective series
    run_length_counter = len (data_object.run_objects [0].elution_ends)
    phase_corrected_efflux_row = 33
    log_efflux_row = phase_corrected_efflux_row + run_length_counter + 1
    efflux_row = log_efflux_row + run_length_counter + 1
    corrected_row = efflux_row + run_length_counter + 1
    raw_row = corrected_row + run_length_counter + 1
    
    worksheet.merge_range (phase_corrected_efflux_row, 0,\
                           phase_corrected_efflux_row, 1,\
                           "Phase-Corr. Log Eff.", phase_format)    
    worksheet.merge_range (log_efflux_row, 0, log_efflux_row, 1,\
                           "Log Efflux", phase_format)
    worksheet.merge_range (efflux_row, 0, efflux_row, 1,\
                           "Efflux", phase_format)
    worksheet.merge_range (corrected_row, 0, corrected_row, 1,\
                           "Corrected AIE", phase_format)
    worksheet.merge_range (raw_row, 0, raw_row, 1,\
                           "Activity in eluant", phase_format)    
    
    for x in range (0, run_length_counter):
        time_point = data_object.run_objects [0].elution_ends [x]

        worksheet.merge_range (1 + phase_corrected_efflux_row + x, 0, 1 + phase_corrected_efflux_row + x, 1, time_point, right_format)
        worksheet.merge_range (1 + log_efflux_row + x, 0, 1 + log_efflux_row + x, 1, time_point, right_format)
        worksheet.merge_range (1 + efflux_row + x, 0, 1 + efflux_row + x, 1, time_point, right_format)
        worksheet.merge_range (1 + corrected_row + x, 0, 1 + corrected_row + x, 1, time_point, right_format)
        worksheet.merge_range (1 + raw_row + x, 0, 1 + raw_row + x, 1, time_point, right_format)
        
    
    
    # Writing Runobject data to sheet
    counter = 2
    for run_object in data_object.run_objects:
        worksheet.write(0, counter, run_object.run_name, middle_format)
        worksheet.write(1, counter, run_object.analysis_type [0])
        worksheet.write(2, counter, run_object.SA)
        worksheet.write(3, counter, run_object.root_cnts)
        worksheet.write(4, counter, run_object.shoot_cnts)
        worksheet.write(5, counter, run_object.root_weight)
        worksheet.write(6, counter, run_object.g_factor)
        worksheet.write(7, counter, run_object.load_time)
        
        worksheet.write(9, counter, run_object.influx)
        worksheet.write(10, counter, run_object.netflux)
        worksheet.write(11, counter, run_object.ratio)
        worksheet.write(12, counter, run_object.poolsize)
        worksheet.write(13, counter, run_object.slope_p3)
        worksheet.write(14, counter, run_object.intercept_p3)
        worksheet.write(15, counter, run_object.r2_p3)
        worksheet.write(16, counter, run_object.k_p3)
        worksheet.write(17, counter, run_object.t05_p3)
        worksheet.write(18, counter, run_object.efflux_p3)
        
        worksheet.write(20, counter, run_object.slope_p2)
        worksheet.write(21, counter, run_object.intercept_p2)
        worksheet.write(22, counter, run_object.r2_p2)
        worksheet.write(23, counter, run_object.k_p2)
        worksheet.write(24, counter, run_object.t05_p2)
        worksheet.write(25, counter, run_object.efflux_p2)

        worksheet.write(27, counter, run_object.slope_p1)
        worksheet.write(28, counter, run_object.intercept_p1)
        worksheet.write(29, counter, run_object.r2_p1)
        worksheet.write(30, counter, run_object.k_p1)
        worksheet.write(31, counter, run_object.t05_p1)
        worksheet.write(32, counter, run_object.efflux_p1)
        
        # Writing Phase I phase-corrected efflux data
        write_phase_corrected_p12(
            workbook, worksheet, run_object, counter, counter,
            sheet_type='summary')

        # Writing Phase III phase-corrected efflux data
        phase_3_row_counter = 34 + len (run_object.y_p12) - 1
        for y in range (0, len (run_object.y_p3)):
            worksheet.write (1 + phase_3_row_counter + y, counter,\
                             run_object.y_p3 [y])
                          
        # Writing efflux elution data that is not phase corrected
        for z in range (0, len (run_object.elution_ends)):
            worksheet.write (1 + log_efflux_row + z, counter, run_object.elution_cpms_log [z])        
            worksheet.write (1 + efflux_row + z, counter, run_object.elution_cpms_gRFW [z])
            worksheet.write (1 + corrected_row + z, counter, run_object.elution_cpms_gfactor [z])
            worksheet.write (1 + raw_row + z, counter, run_object.elution_cpms [z])
        
        counter += 1
                
def generate_analysis(data_object):
    '''
    Creating an excel file in directory using a preset naming convention
    Data in the file are the product of CATE analysis from a template file
    containing the raw information
    Nothing is returned
    '''
    
    workbook = xlsxwriter.Workbook('Didthiswork.xlsx')
    
    # Formatting for items for headers for analyzed CATE data
    analyzed_header = workbook.add_format()
    analyzed_header.set_text_wrap()
    analyzed_header.set_align('center')
    analyzed_header.set_align('vcenter')
    analyzed_header.set_bottom()    
    
    # Formatting for cells that contains basic CATE data    
    basic_format = workbook.add_format ()
    basic_format.set_align('center')
    basic_format.set_align('vcenter')
    basic_format.set_top()    
    basic_format.set_bottom()    
    basic_format.set_right()    
    basic_format.set_left()       
    
    # Formatting for cells that contain "Phase x" row labels
    phase_format = workbook.add_format ()
    phase_format.set_align('right')
    phase_format.set_align('vcenter')
    phase_format.set_right()
    
    emphasis_format = workbook.add_format ()
    emphasis_format.set_bold()
    
    
    # Header info for analyzed CATE data
    basic_headers = [
        ("Corrected AIE (cpm)", 15), 
        (u"Efflux (cpm \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9)", 13.57),
        ("Log Efflux", 8.86), (u"R\u00b2", 7),
        (u"Slope (min\u207b\u00b9)", 8.14), ("Intercept", 8.5),
        ("Phase-Corr. PII", 8.86), ("Phase-Corr. PI", 8.86)]    
    
    phasedata_headers = [
        "Slope", "Intercept", u"R\u00b2", "k", "Half-Life", "Efflux", "Influx",
        "Net flux", "E:I Ratio", "Pool Size"]
    
    for run_object in data_object.run_objects:
        worksheet = generate_sheet (workbook, run_object.run_name)
        
        counter = 3
        for y in range(0, len(basic_headers)):
            if (y < 3 or y > 5) or run_object.analysis_type[0] == 'obj':
                worksheet.write(7, counter, basic_headers[y][0], analyzed_header)   
                worksheet.set_column (counter, counter, basic_headers[y][1])
                counter += 1
                
        for z in range(0, len (phasedata_headers)):
            worksheet.write(1, z + 4, phasedata_headers[z], analyzed_header) 
        
        worksheet.write(2, 3, "Phase I", phase_format)
        worksheet.write(3, 3, "Phase II", phase_format)
        worksheet.write(4, 3, "Phase III", phase_format)
                            
        # Writing CATE data to file        
        worksheet.write(0, 2, run_object.run_name, analyzed_header)    
        worksheet.write(1, 2, run_object.SA, basic_format)    
        worksheet.write(2, 2, run_object.root_cnts, basic_format)
        worksheet.write(3, 2, run_object.shoot_cnts, basic_format)
        worksheet.write(4, 2, run_object.root_weight, basic_format)
        worksheet.write(5, 2, run_object.g_factor, basic_format)
        worksheet.write(6, 2, run_object.load_time, basic_format)
        
        worksheet.write(2, 4, run_object.slope_p1)
        worksheet.write(2, 5, run_object.intercept_p1)
        worksheet.write(2, 6, run_object.r2_p1)
        worksheet.write(2, 7, run_object.k_p1)
        worksheet.write(2, 8, run_object.t05_p1)
        worksheet.write(2, 9, run_object.efflux_p1)
        
        worksheet.write(3, 4, run_object.slope_p2)
        worksheet.write(3, 5, run_object.intercept_p2)
        worksheet.write(3, 6, run_object.r2_p2)
        worksheet.write(3, 7, run_object.k_p2)
        worksheet.write(3, 8, run_object.t05_p2)
        worksheet.write(3, 9, run_object.efflux_p2)
        
        worksheet.write(4, 4, run_object.slope_p3)
        worksheet.write(4, 5, run_object.intercept_p3)
        worksheet.write(4, 6, run_object.r2_p3)
        worksheet.write(4, 7, run_object.k_p3)
        worksheet.write(4, 8, run_object.t05_p3)
        worksheet.write(4, 9, run_object.efflux_p3)          
        
        worksheet.write(4, 10, run_object.influx)
        worksheet.write(4, 11, run_object.netflux)
        worksheet.write(4, 12, run_object.ratio)
        worksheet.write(4, 13, run_object.poolsize)
        
        for x in range(0, len (run_object.elution_ends)):
            worksheet.write(8 + x, 0, x + 1)
            worksheet.write(8 + x, 1, run_object.elution_ends [x])
            worksheet.write(8 + x, 2, run_object.elution_cpms [x])
            worksheet.write(8 + x, 3, run_object.elution_cpms_gfactor [x])
            worksheet.write(8 + x, 4, run_object.elution_cpms_gRFW [x])
            worksheet.write(8 + x, 5, run_object.elution_cpms_log [x])
        
        if run_object.analysis_type[0] == 'obj':

            p1_regression_counter = 9 + len(run_object.elution_ends)\
                - len(run_object.r2s_p3_list) - run_object.analysis_type[1]
            chart_col = "L"
            p2_column = 10
            
            for y in range(0, len (run_object.r2s_p3_list)):
                worksheet.write(p1_regression_counter + y, 6, run_object.r2s_p3_list [y])
                worksheet.write(p1_regression_counter + y, 7, run_object.slopes_p3_list [y])
                worksheet.write(p1_regression_counter + y, 8, run_object.intercepts_p3_list [y])
                
            # Emphasizing data actually used for p3
            worksheet.write(p1_regression_counter + 3, 6,
                             run_object.r2s_p3_list[3], emphasis_format)
            worksheet.write(p1_regression_counter + 3, 7,
                             run_object.slopes_p3_list[3], emphasis_format)
            worksheet.write(p1_regression_counter + 3, 8,
                             run_object.intercepts_p3_list[3], emphasis_format)
                        
        else: # No columns for lists of r2/intercept/slope
            chart_col = "I"
            p2_column = 7

        # Writing Phase-corrected data in appropriate column and with missing
        # cells (if applicable due to potential negative log operations
        write_phase_corrected_p12(
            workbook, worksheet, run_object, p2_column - 1, p2_column,
            sheet_type='single')
        
        # Graphing the RunObject Data
        
        # Graphing Phase III data
        chart = workbook.add_chart({'type': 'scatter'})
        chart.set_title({'name': 'Phase III', 'overlay': True})
        worksheet.insert_chart(chart_col + '6', chart)        
        series_end = len(run_object.elution_cpms_log) + 9
        
        # Add log efflux data to Phase III chart
        chart.add_series({
            'categories': '=' + run_object.run_name + '!$B$9:' + '$B$' + str(series_end),
            'values': '=' + run_object.run_name + '!$F$9:' + '$F$' + str(series_end),
            'name' : run_object.run_name,
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': 'black'},
                       'fill':   {'color': 'white'}} })
        
        # Add Phase III data to Phase III chart
        p3_chart_start = str(series_end - len(run_object.y_p3))
        chart.add_series({
            'categories': '=' + run_object.run_name + '!$B$'+ p3_chart_start +':' + '$B$' + str (series_end),
            'values': '=' + run_object.run_name + '!$F$'+ p3_chart_start + ':' + '$F$' + str (series_end),
            'name' : run_object.run_name,
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': 'black'},
                       'fill':   {'color': 'gray'}} })  
        
        # Add last point of p3
        last_point = 9 + len(run_object.elution_ends) - len(run_object.y_p3)
        chart.add_series({
            'categories': '=' + run_object.run_name + '!$B$'+ str(last_point),
            'values': '=' + run_object.run_name + '!$F$'+ str(last_point),            
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': 'black'},
                       'fill':   {'color': 'red'}} })
        
        # Add p3 regression line
        worksheet.write (7, 11, 0)          
        worksheet.write (7, 12, run_object.intercept_p3)          
        worksheet.write (8, 11, run_object.elution_ends[-1]) 
        worksheet.write (8, 12, run_object.slope_p3 *\
                         run_object.elution_ends[-1] + run_object.intercept_p3)

        chart.add_series({
            'categories': [run_object.run_name, 7, 11, 8,11],
            'values': [run_object.run_name, 7, 12, 8, 12],
            'line' : {'color': 'red','dash_type' : 'dash'},
            'marker': {'type': 'none',
                       'size,': 0,
                       'border': {'color': 'purple'},
                       'fill':   {'color': 'gray'}} })         

        chart.set_legend({'none': True})        
        chart.set_x_axis({'name': 'Elution time (min)',})        
        chart.set_y_axis({'name': 'Log Efflux/g RFW/min',})         
        
            
    generate_summary (workbook, data_object)
            
    workbook.close()
    
def write_phase_corrected_p12(
    workbook, worksheet, run_object, p1_column, p2_column, sheet_type):
    '''
    sheet = 'single'/'summary
    '''
    
    bot_line = workbook.add_format()
    bot_line.set_bottom()  
    
    # Writing Phase-corrected data in appropriate column
    if sheet_type == 'single':
        row_counter = 8
    elif sheet_type == 'summary':
        row_counter = 34 #!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    
    elution_counter = 0
    series_counter = 0 
    elution_p2_counter = len(run_object.x_p1_curvestrippedof_p23)
    
    while series_counter < len(run_object.y_p1_curvestrippedof_p23):
        # Need to be sure that y value we are entering wasn't skipped as a 
        # result of a negative log operation
        if run_object.x_p1_curvestrippedof_p23[series_counter] == \
                run_object.elution_ends[elution_counter]:
            if series_counter != len(run_object.y_p1) - 1:
                worksheet.write (
                    row_counter, p2_column, 
                    run_object.y_p1_curvestrippedof_p23[series_counter])
            else:
                worksheet.write (
                    row_counter, p2_column,
                    run_object.y_p1_curvestrippedof_p23[series_counter],
                    bot_line)                
            elution_counter += 1
            series_counter += 1
                            
        else: # a series entry has been skipped, move elution list forward
            elution_counter += 1
            elution_p2_counter += 1 # len of p1 y-series artificially short
        row_counter += 1
        
    series_counter = 0
    
    while series_counter < len(run_object.y_p2_curvestrippedof_p3):
        # Need to be sure that y value we are entering wasn't skipped as a 
        # result of a negative log operation
        if run_object.x_p2_curvestrippedof_p3[series_counter] ==\
                run_object.elution_ends[elution_counter]:
            if series_counter != len(run_object.y_p2_curvestrippedof_p3) -1:
                worksheet.write(
                    row_counter, p1_column,
                    run_object.y_p2_curvestrippedof_p3[series_counter])
            else:
                worksheet.write (
                    row_counter, p1_column, 
                    run_object.y_p2_curvestrippedof_p3[series_counter],
                    bot_line)                
            elution_counter += 1
            series_counter += 1    
            
        else: # a series entry has been skipped, move elution list forward
            elution_counter += 1
            
        row_counter += 1     
    

     
if __name__ == "__main__":
    #temp_book = xlsxwriter.Workbook('filename.xlsx')
    temp_data = grab_data("C:\Users\Ruben\Projects\CATEAnalysis", "CATE Template - Multi Run.xlsx")
    generate_analysis(temp_data)
    #generate_template ("C:\Users\Ruben\Desktop\CATE_EXCEL_TEST.xlsx")
               
    