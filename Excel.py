import os
import time
from xlrd import *
import xlsxwriter


import Objects
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
    all_analysis_objects = []

    # Parsing elution times, correcting for header offset (8)
    raw_elution_times = input_sheet.col(1) # Col w/elution times given in file
    elut_ends = [float(x.value) for x in raw_elution_times[8:]]
        
    for col_index in range(2, input_sheet.row_len(0)):        
        # Grab individual CATE values of interest
        run_name = str(input_sheet.cell(0, col_index).value) # in case name is #
        SA = input_sheet.cell(1, col_index).value
        root_cnts = input_sheet.cell(2, col_index).value
        shoot_cnts = input_sheet.cell(3, col_index).value
        root_weight = input_sheet.cell(4, col_index).value
        g_factor = input_sheet.cell(5, col_index).value
        load_time = input_sheet.cell(6, col_index).value
        # Grabing elution cpms, correcting for header offset (8)
        raw_cpm_column = input_sheet.col(col_index) # Raw counts given by file
        elution_cpms = []
        for item in raw_cpm_column[8:]:
            if item.value != '':
                elution_cpms.append(float(item.value))
            else:
                elution_cpms.append(0.0)

        temp_run = Objects.Run(
            run_name, SA, root_cnts, shoot_cnts, root_weight, g_factor, 
            load_time, elut_ends, elution_cpms)
        
        all_analysis_objects.append(Objects.Analysis(
            kind=None, obj_num_pts=None, run=temp_run))
   
    return Objects.Experiment(directory, all_analysis_objects)

def generate_summary(workbook, experiment):
    '''
    Given an open workbook, create a summary sheet containing relevant data
    from all analyses in experiment.analyses
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
        worksheet.merge_range(z + 13, 0, z + 13, 1,\
                               phasedata_headers [z], req)
        worksheet.merge_range(z + 20, 0, z + 20, 1,\
                               phasedata_headers [z], req)        
        worksheet.merge_range(z + 27, 0, z + 27, 1,\
                               phasedata_headers [z], req)
    
    worksheet.merge_range(8, 0, 8, 1, "Phase III", phase_format)
    worksheet.merge_range(19, 0, 19, 1, "Phase II", phase_format)
    worksheet.merge_range(26, 0, 26, 1, "Phase I", phase_format)
        
    # Writing elution time points/headers for respective series
    run_length_counter = len(experiment.analyses[0].elut_ends)
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
        time_point = experiment.analyses [0].elution_ends [x]

        worksheet.merge_range (1 + phase_corrected_efflux_row + x, 0, 1 + phase_corrected_efflux_row + x, 1, time_point, right_format)
        worksheet.merge_range (1 + log_efflux_row + x, 0, 1 + log_efflux_row + x, 1, time_point, right_format)
        worksheet.merge_range (1 + efflux_row + x, 0, 1 + efflux_row + x, 1, time_point, right_format)
        worksheet.merge_range (1 + corrected_row + x, 0, 1 + corrected_row + x, 1, time_point, right_format)
        worksheet.merge_range (1 + raw_row + x, 0, 1 + raw_row + x, 1, time_point, right_format)
        
    
    
    # Writing Runobject data to sheet
    counter = 2
    for analysis in experiment.analyses:
        worksheet.write(0, counter, analysis.run_name, middle_format)
        worksheet.write(1, counter, analysis.analysis_type [0])
        worksheet.write(2, counter, analysis.SA)
        worksheet.write(3, counter, analysis.root_cnts)
        worksheet.write(4, counter, analysis.shoot_cnts)
        worksheet.write(5, counter, analysis.root_weight)
        worksheet.write(6, counter, analysis.g_factor)
        worksheet.write(7, counter, analysis.load_time)
        
        worksheet.write(9, counter, analysis.influx)
        worksheet.write(10, counter, analysis.netflux)
        worksheet.write(11, counter, analysis.ratio)
        worksheet.write(12, counter, analysis.poolsize)
        worksheet.write(13, counter, analysis.slope_p3)
        worksheet.write(14, counter, analysis.intercept_p3)
        worksheet.write(15, counter, analysis.r2_p3)
        worksheet.write(16, counter, analysis.k_p3)
        worksheet.write(17, counter, analysis.t05_p3)
        worksheet.write(18, counter, analysis.efflux_p3)
        
        worksheet.write(20, counter, analysis.slope_p2)
        worksheet.write(21, counter, analysis.intercept_p2)
        worksheet.write(22, counter, analysis.r2_p2)
        worksheet.write(23, counter, analysis.k_p2)
        worksheet.write(24, counter, analysis.t05_p2)
        worksheet.write(25, counter, analysis.efflux_p2)

        worksheet.write(27, counter, analysis.slope_p1)
        worksheet.write(28, counter, analysis.intercept_p1)
        worksheet.write(29, counter, analysis.r2_p1)
        worksheet.write(30, counter, analysis.k_p1)
        worksheet.write(31, counter, analysis.t05_p1)
        worksheet.write(32, counter, analysis.efflux_p1)
        
        # Writing Phase I phase-corrected efflux data
        write_phase_corrected_p12(
            workbook, worksheet, analysis, counter, counter,
            sheet_type='summary')

        # Writing Phase III phase-corrected efflux data
        phase_3_row_counter = 34 + len(analysis.y_p12) - 1
        for y in range(0, len(analysis.y_p3)):
            worksheet.write(
                1 + phase_3_row_counter + y, counter, analysis.y_p3[y])
                          
        # Writing efflux elution data that is not phase corrected
        for z in range (0, len(analysis.elution_ends)):
            worksheet.write(
                1 + log_efflux_row + z, counter, analysis.elut_cpms_log[z])        
            worksheet.write(
                1 + efflux_row + z, counter, analysis.elut_cpms_gRFW[z])
            worksheet.write(
                1 + corrected_row + z, counter, analysis.elut_cpms_gfact[z])
            worksheet.write(
                1 + raw_row + z, counter, analysis.elut_cpms[z])
        
        counter += 1
                
def generate_analysis(experiment):
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
        'Start', 'End', "Slope", "Intercept", u"R\u00b2", "k", "Half-Life",
        "Efflux", "Influx", "Net flux", "E:I Ratio", "Pool Size"]
    
    for analysis in experiment.analyses:
        worksheet = generate_sheet (workbook, analysis.run.name)
        
        counter = 3
        for index, item in enumerate(basic_headers):
            if (index<3 or index>5) or analysis.kind == 'obj':
                worksheet.write(
                    7, counter, basic_headers[index][0], analyzed_header)   
                worksheet.set_column(
                    counter, counter, basic_headers[index][1])
                counter += 1
        
        for index, item in enumerate(phasedata_headers):
            worksheet.write(
                1, index + 4, phasedata_headers[index], analyzed_header) 
        
        worksheet.write(2, 3, "Phase I", phase_format)
        worksheet.write(3, 3, "Phase II", phase_format)
        worksheet.write(4, 3, "Phase III", phase_format)
                            
        # Writing CATE data to file        
        worksheet.write(0, 2, analysis.run.name, analyzed_header)    
        worksheet.write(1, 2, analysis.run.SA, basic_format)    
        worksheet.write(2, 2, analysis.run.rt_cnts, basic_format)
        worksheet.write(3, 2, analysis.run.sht_cnts, basic_format)
        worksheet.write(4, 2, analysis.run.rt_wght, basic_format)
        worksheet.write(5, 2, analysis.run.gfact, basic_format)
        worksheet.write(6, 2, analysis.run.load_time, basic_format)
        
        worksheet.write(2, 4, analysis.phase1.xs[0])
        worksheet.write(2, 5, analysis.phase1.xs[1])
        worksheet.write(2, 6, analysis.phase1.slope)
        worksheet.write(2, 7, analysis.phase1.intercept)
        worksheet.write(2, 8, analysis.phase1.r2)
        worksheet.write(2, 9, analysis.phase1.k)
        worksheet.write(2, 10, analysis.phase1.t05)
        worksheet.write(2, 11, analysis.phase1.efflux)
        
        worksheet.write(3, 4, analysis.phase2.xs[0])
        worksheet.write(3, 5, analysis.phase2.xs[1])
        worksheet.write(3, 6, analysis.phase2.slope)
        worksheet.write(3, 7, analysis.phase2.intercept)
        worksheet.write(3, 8, analysis.phase2.r2)
        worksheet.write(3, 9, analysis.phase2.k)
        worksheet.write(3, 10, analysis.phase2.t05)
        worksheet.write(3, 11, analysis.phase2.efflux)

        worksheet.write(4, 4, analysis.phase3.xs[0])
        worksheet.write(4, 5, analysis.phase3.xs[1])        
        worksheet.write(4, 6, analysis.phase3.slope)
        worksheet.write(4, 7, analysis.phase3.intercept)
        worksheet.write(4, 8, analysis.phase3.r2)
        worksheet.write(4, 9, analysis.phase3.k)
        worksheet.write(4, 10, analysis.phase3.t05)
        worksheet.write(4, 11, analysis.phase3.efflux)
        
        worksheet.write(4, 12, analysis.influx)
        worksheet.write(4, 13, analysis.netflux)
        worksheet.write(4, 14, analysis.ratio)
        worksheet.write(4, 15, analysis.poolsize)

        if analysis.kind == 'obj':

            chartinsert_p3 = "L6"                        
            chartinsert_p2 = "L20"
            chartinsert_p1 = "L34"
            colletter_p2 = 'J'      
            colletter_p1 = 'K'                        
            colnum_p2 = 9      
            colnum_p1 = 10
            for index, item in enumerate(analysis.r2s):
                worksheet.write(8 + index, 6, analysis.r2s[index])
                worksheet.write(8 + index, 7, analysis.ms[index])
                worksheet.write(8 + index, 8, analysis.bs[index])
            for index, item in enumerate(analysis.run.elut_ends):
                    if item is analysis.phase3.xs[0]:
                        worksheet.write(
                            8 + index, 6, analysis.r2s[index], emphasis_format)
                        worksheet.write(
                            8 + index, 7, analysis.ms[index], emphasis_format)
                        worksheet.write(
                            8 + index, 8, analysis.bs[index], emphasis_format)

                        
        elif analysis.kind == 'subj': # No columns for lists of r2/intercept/slope
            chartinsert_p3 = "I6"
            chartinsert_p2 = "I20"
            chartinsert_p1 = "I34"
            colletter_p2 = 'G'
            colletter_p1 = 'H'
            colnum_p2 = 6      
            colnum_p1 = 7
        
        if analysis.phase3.xs != ('', ''):
            # Writing elut_cpm_etc data
            for index, item in enumerate(analysis.run.elut_ends):
                worksheet.write(8 + index, 0, index + 1)
                worksheet.write(8 + index, 1, item)
                if item in analysis.run.elut_ends_parsed:
                    if item is analysis.phase3.xs[0]:
                        p3_chart_start = str(9 + index)
                    if item is analysis.phase3.xs[1]:
                        p3_chart_end = str(9 + index)
                    new_index = analysis.run.elut_ends_parsed.index(item)
                    worksheet.write(
                        8 + index, 2, analysis.run.elut_cpms[new_index])
                    worksheet.write(
                        8 + index, 3, analysis.run.elut_cpms_gfact[new_index])
                    worksheet.write(
                        8 + index, 4, analysis.run.elut_cpms_gRFW[new_index])
                    worksheet.write(
                        8 + index, 5, analysis.run.elut_cpms_log[new_index])

            # Graphing Phase III data
            chart_p3 = workbook.add_chart({'type': 'scatter'})
            chart_p3.set_title({'name': 'Phase III', 'overlay': True})
            worksheet.insert_chart(chartinsert_p3, chart_p3)        
            series_end = len(analysis.run.elut_ends) + 9
            
            # Add log efflux data to Phase III chart_p3
            chart_p3.add_series({
                'categories': '=' + analysis.run.name + '!$B$9:' + '$B$' + str(series_end),
                'values': '=' + analysis.run.name + '!$F$9:' + '$F$' + str(series_end),
                'name' : analysis.run.name,
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'white'}} })
            
            # Add Phase III data to Phase III chart_p3
            chart_p3.add_series({
                'categories': '=' + analysis.run.name + \
                    '!$B$'+ p3_chart_start +':' +\
                    '$B$' + p3_chart_end,
                'values': '=' + analysis.run.name +\
                    '!$F$'+ p3_chart_start + ':' +\
                    '$F$' + p3_chart_end,
                'name' : analysis.run.name,
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'gray'}} })

            # Add first point of p3
            first_pt_p3 = 9 + len(analysis.run.elut_ends) - len(analysis.phase3.y_series)
            chart_p3.add_series({
                'categories': '=' + analysis.run.name + '!$B$'+ str(first_pt_p3),
                'values': '=' + analysis.run.name + '!$F$'+ str(first_pt_p3),            
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'red'}} })
            # Add pts used for obj analysis
            num_obj_p3 = 9 + len(analysis.run.elut_ends) - analysis.obj_num_pts
            chart_p3.add_series({
                'categories': '=' + analysis.run.name + '!$B$'+ str(num_obj_p3)+':' + '$B$' + str (series_end),
                'values': '=' + analysis.run.name + '!$F$'+ str(num_obj_p3)+':' + '$F$' + str (series_end),            
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'black'}} })
            
            # Add p3 regression line
            worksheet.write (7, 11, analysis.phase3.xy1[0])          
            worksheet.write (7, 12, analysis.phase3.xy1[1])          
            worksheet.write (8, 11, analysis.phase3.xy2[0]) 
            worksheet.write (8, 12, analysis.phase3.xy2[1])

            chart_p3.add_series({
                'categories': [analysis.run.name, 7, 11, 8,11],
                'values': [analysis.run.name, 7, 12, 8, 12],
                'line' : {'color': 'red','dash_type' : 'dash'},
                'marker': {'type': 'none'}
                })         

            chart_p3.set_legend({'none': True})        
            chart_p3.set_x_axis({'name': 'Elution time (min)',})        
            chart_p3.set_y_axis({'name': 'Log Efflux/g RFW/min',})   
        
        # Graphing Phase II data
        if analysis.phase2.xs != ('', ''):
            # Write phase 2 data
            for index, item in enumerate(analysis.run.elut_ends):
                if item in analysis.phase2.x_series:
                    if item is analysis.phase2.xs[0]:
                        p2_chart_start = str(9 + index)
                    elif item is analysis.phase2.xs[1]:
                        p2_chart_end = str(9 + index)
                    new_index = analysis.phase2.x_series.index(item)
                    worksheet.write(
                        8 + index, colnum_p2, analysis.phase2.y_series[new_index])
            # Draw phase 2 data
            chart_p2 = workbook.add_chart({'type': 'scatter'})
            chart_p2.set_title({'name': 'Phase II', 'overlay': True})
            worksheet.insert_chart(chartinsert_p2, chart_p2)  
            chart_p2.add_series({
                'categories': '=' + analysis.run.name + \
                    '!$B' + '$' + p2_chart_start  + ':' +\
                    '$B' + '$' + p2_chart_end,
                'values': '=' + analysis.run.name + \
                    '!$F' +  '$' + p2_chart_start  + ':' + \
                    '$F' + '$' + p2_chart_end,
                'name' : analysis.run.name,
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'white'}} })
            chart_p2.add_series({
                'categories': '=' + analysis.run.name + \
                    '!$B' + '$' + p2_chart_start  + ':' +\
                    '$B' + '$' + p2_chart_end,
                'values': '=' + analysis.run.name + \
                    '!$' + colletter_p2 +  '$' + p2_chart_start  + ':' + \
                    '$' + colletter_p2 + '$' + p2_chart_end,
                'name' : analysis.run.name,
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'gray'}} })
            # Add p2 regression line
            worksheet.write (9, 11, analysis.phase2.xy1[0])          
            worksheet.write (9, 12, analysis.phase2.xy1[1])          
            worksheet.write (10, 11, analysis.phase2.xy2[0]) 
            worksheet.write (10, 12, analysis.phase2.xy2[1])

            chart_p2.add_series({
                'categories': [analysis.run.name, 9, 11, 10,11],
                'values': [analysis.run.name, 9, 12, 10, 12],
                'line' : {'color': 'red','dash_type' : 'dash'},
                'marker': {'type': 'none',
                           'size,': 0,
                           'border': {'color': 'purple'},
                           'fill':   {'color': 'gray'}} })         

            chart_p2.set_legend({'none': True})        
            chart_p2.set_x_axis({'name': 'Elution time (min)',})        
            chart_p2.set_y_axis({'name': 'Log Efflux/g RFW/min',}) 

        if analysis.phase1.xs != ('', ''):
            # Write phase 2 data
            for index, item in enumerate(analysis.run.elut_ends):
                if item in analysis.phase1.x_series:
                    if item is analysis.phase1.xs[0]:
                        p1_chart_start = str(9 + index)
                    elif item is analysis.phase1.xs[1]:
                        p1_chart_end = str(9 + index)
                    new_index = analysis.phase1.x_series.index(item)
                    worksheet.write(
                        8 + index, colnum_p1, analysis.phase1.y_series[new_index])
            # Draw phase 2 data
            chart_p1 = workbook.add_chart({'type': 'scatter'})
            chart_p1.set_title({'name': 'Phase I', 'overlay': True})
            worksheet.insert_chart(chartinsert_p1, chart_p1)  
            chart_p1.add_series({
                'categories': '=' + analysis.run.name + \
                    '!$B' + '$' + p1_chart_start  + ':' +\
                    '$B' + '$' + p1_chart_end,
                'values': '=' + analysis.run.name + \
                    '!$F' +  '$' + p1_chart_start  + ':' + \
                    '$F' + '$' + p1_chart_end,
                'name' : analysis.run.name,
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'white'}} })
            
            chart_p1.add_series({
                'categories': '=' + analysis.run.name + \
                    '!$B' + '$' + p1_chart_start  + ':' +\
                    '$B' + '$' + p1_chart_end,
                'values': '=' + analysis.run.name + \
                    '!$' + colletter_p1 +  '$' + p1_chart_start  + ':' + \
                    '$' + colletter_p1 + '$' + p1_chart_end,
                'name' : analysis.run.name,
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'gray'}} })
            
            # Add p1 regression line
            worksheet.write (11, 11, analysis.phase1.xy1[0])          
            worksheet.write (11, 12, analysis.phase1.xy1[1])          
            worksheet.write (12, 11, analysis.phase1.xy2[0]) 
            worksheet.write (12, 12, analysis.phase1.xy2[1])

            chart_p1.add_series({
                'categories': [analysis.run.name, 11, 11, 12,11],
                'values': [analysis.run.name, 11, 12, 12, 12],
                'line' : {'color': 'red','dash_type' : 'dash'},
                'marker': {'type': 'none',
                           'size,': 0,
                           'border': {'color': 'purple'},
                           'fill':   {'color': 'gray'}} })    


            chart_p1.set_legend({'none': True})        
            chart_p1.set_x_axis({'name': 'Elution time (min)',})        
            chart_p1.set_y_axis({'name': 'Log Efflux/g RFW/min',}) 
            
    #generate_summary(workbook, experiment)     
    workbook.close()
     
if __name__ == "__main__":
    directory = os.path.dirname(os.path.abspath(__file__))
    temp_experiment = grab_data(directory, "/Tests/3/Test_SingleRun1.xlsx")
    temp_experiment.analyses[0].kind = 'obj'
    temp_experiment.analyses[0].obj_num_pts = 8
    temp_experiment.analyses[0].analyze()
    generate_analysis(temp_experiment)