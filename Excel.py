import os
import time
from xlrd import *
import xlsxwriter

import Objects

def generate_sheet(workbook, sheet_name, template=True):
    """Create/return basic excel sheet template (in existing <workbook>)

    This is the template upon which most sheets (which the exception of summary
    and analysis sheets) are built.
    template=True is used to set row where bulk of data is output

    @type workbook: Workbook
        Workbook object into which sheet is inserted
    @type sheet_name: str
        label of Worksheet object that is to be inserted
    @type Template: bool
    @rtype: Workbook
    """
    # Formatting for cells surrounded by a border (prev. empty_row/basic_format)
    BORDER = workbook.add_format()
    BORDER.set_align('center')
    BORDER.set_align('vcenter')
    BORDER.set_top()    
    BORDER.set_bottom()    
    BORDER.set_right()    
    BORDER.set_left()
    # Formatting for bolded cells (previously req/border_all)
    BORDER_BOLD = workbook.add_format()
    BORDER_BOLD.set_text_wrap()
    BORDER_BOLD.set_align('center')
    BORDER_BOLD.set_align('vcenter')
    BORDER_BOLD.set_bold()
    BORDER_BOLD.set_bottom()
    BORDER_BOLD.set_right()    
    BORDER_BOLD.set_left()
    BORDER_BOLD.set_top()
    # Formatting for not bolded cells (prev. not_req/analyzed_header/not_bold)
    BORDER_BOT = workbook.add_format()
    BORDER_BOT.set_text_wrap()
    BORDER_BOT.set_align('center')
    BORDER_BOT.set_align('vcenter')
    BORDER_BOT.set_bottom()

    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.set_row(1, 30.75) # Setting the height of the SA row to ~2 lines
    
    # List of ordered tuples containing (title, formatting, column width) 
    # in order they are to be written to the file
    row_headers = [
        'Run Name',
        u"Specific Activity (cpm \u00B7 \u00B5mol\u207b\u00b9)",
        "Root Cnts (cpm)", "Shoot Cnts (cpm)", "Root weight (g)", "G-Factor", 
        "Load Time (min)"]
    for index, item in enumerate(row_headers):
        worksheet.merge_range (
            index, 0,
            index, 1, 
            row_headers[index], BORDER_BOLD)
        worksheet.write (index + 1, 2, "", BORDER)

    if template:
        worksheet.write(
            7, 0,
            "Vial #",
            BORDER_BOT)
        worksheet.set_column(0, 0, 3.57)
        worksheet.write(
            7, 1,
            "Elution time (min)",
            BORDER_BOLD)
        worksheet.set_column(1, 1, 11.7)
        worksheet.write(
            7, 2,
            "Activity in eluant (cpm)",
            BORDER_BOLD)
        worksheet.set_column(2, 2, 15)
      
    return worksheet

def generate_template(workbook):
    '''Generates a CATE template sheet in an already created workbook.
    '''
    # Formatting for basic cells (previously run_header/basic)
    BASIC = workbook.add_format()
    BASIC.set_text_wrap()
    BASIC.set_align('center')
    BASIC.set_align('vcenter')

    worksheet = generate_sheet(workbook, 'Template', template=True)

    # Writing headers columns containing individual runs 
    worksheet.write(0, 2, "Run 1", BASIC)
    worksheet.write(0, 3, "etc.", BASIC)
        
def generate_analysis(experiment):
    '''Creating an excel file in directory using a preset naming convention.

    Excel file contains comprehensive data analysis as created by the user.

    Precondition: Data in the file are the product of CATE analysis from a
        template file containing the raw information.

    @type experiment: Experiment
    @rtype: None
    '''
    # Formatting for bolded cells
    BOLD = workbook.add_format()   
    BOLD.set_bold()    # Formatting for bolded cells
    BOLD = workbook.add_format()   
    BOLD.set_bold()
    # Formatting for cells surrounded by a border (prev. empty_row/basic_format)
    BORDER = workbook.add_format()
    BORDER.set_align('center')
    BORDER.set_align('vcenter')
    BORDER.set_top()    
    BORDER.set_bottom()    
    BORDER.set_right()    
    BORDER.set_left()
    # Formatting for not bolded cells (prev. not_req/analyzed_header/not_bold)
    BORDER_BOT = workbook.add_format()
    BORDER_BOT.set_text_wrap()
    BORDER_BOT.set_align('center')
    BORDER_BOT.set_align('vcenter')
    BORDER_BOT.set_bottom()
    # Formatting for bolded cells (previously req)
    BORDER_BOLD_BOT = workbook.add_format()
    BORDER_BOLD_BOT.set_text_wrap()
    BORDER_BOLD_BOT.set_align('center')
    BORDER_BOLD_BOT.set_align('vcenter')
    BORDER_BOLD_BOT.set_bold()
    BORDER_BOLD_BOT.set_bottom()
    # Formatting for cells with a left border
    BORDER_LEFT = workbook.add_format()   
    BORDER_LEFT.set_left()
    # Formatting for not bolded cells (previously not_req/analyzed_header/not_bold)
    BORDER_RIGHT = workbook.add_format()
    BORDER_RIGHT.set_text_wrap()
    BORDER_RIGHT.set_align('right')
    BORDER_RIGHT.set_right()

    output_name = 'CATE Output - ' + time.strftime("(%Y_%m_%d).xlsx")
    workbook = xlsxwriter.Workbook(experiment.directory + "\\" + output_name)
    generate_summary(workbook, experiment)     
    phasedata_headers = [
        'Start', 'End', "Slope", "Intercept", u"R\u00b2", "k", "Half-Life",
        "Efflux", "Influx", "Net flux", "E:I Ratio", "Pool Size"]
    
    for analysis in experiment.analyses:
        worksheet = generate_sheet(workbook, analysis.run.name, template=False)

        # Writing headers for phase data
        worksheet.write(2, 4, 'Phase I', BORDER_RIGHT)
        worksheet.write(3, 4, 'Phase II', BORDER_RIGHT)
        worksheet.write(4, 4, 'Phase III', BORDER_RIGHT)
        for index, item in enumerate(phasedata_headers):
            worksheet.write(
                1, index + 5, phasedata_headers[index], BORDER_BOT)
        # Writing individual phase data
        worksheet.write(0, 2, analysis.run.name, BORDER_BOT)    
        worksheet.write(1, 2, analysis.run.SA, BORDER)    
        worksheet.write(2, 2, analysis.run.rt_cnts, BORDER)
        worksheet.write(3, 2, analysis.run.sht_cnts, BORDER)
        worksheet.write(4, 2, analysis.run.rt_wght, BORDER)
        worksheet.write(5, 2, analysis.run.gfact, BORDER)
        worksheet.write(6, 2, analysis.run.load_time, BORDER)
        
        worksheet.write(2, 5, analysis.phase1.xs[0])
        worksheet.write(2, 6, analysis.phase1.xs[1])
        worksheet.write(2, 7, analysis.phase1.slope)
        worksheet.write(2, 8, analysis.phase1.intercept)
        worksheet.write(2, 9, analysis.phase1.r2)
        worksheet.write(2, 10, analysis.phase1.k)
        worksheet.write(2, 11, analysis.phase1.t05)
        worksheet.write(2, 12, analysis.phase1.efflux)
        
        worksheet.write(3, 5, analysis.phase2.xs[0])
        worksheet.write(3, 6, analysis.phase2.xs[1])
        worksheet.write(3, 7, analysis.phase2.slope)
        worksheet.write(3, 8, analysis.phase2.intercept)
        worksheet.write(3, 9, analysis.phase2.r2)
        worksheet.write(3, 10, analysis.phase2.k)
        worksheet.write(3, 11, analysis.phase2.t05)
        worksheet.write(3, 12, analysis.phase2.efflux)

        worksheet.write(4, 5, analysis.phase3.xs[0])
        worksheet.write(4, 6, analysis.phase3.xs[1])        
        worksheet.write(4, 7, analysis.phase3.slope)
        worksheet.write(4, 8, analysis.phase3.intercept)
        worksheet.write(4, 9, analysis.phase3.r2)
        worksheet.write(4, 10, analysis.phase3.k)
        worksheet.write(4, 11, analysis.phase3.t05)
        worksheet.write(4, 12, analysis.phase3.efflux)
        
        worksheet.write(4, 13, analysis.influx)
        worksheet.write(4, 14, analysis.netflux)
        worksheet.write(4, 15, analysis.ratio)
        worksheet.write(4, 16, analysis.poolsize)

        # Writing columns of data as they appear in the analysis object
        worksheet.write(22, 0, "#", BORDER_BOT)
        worksheet.set_column(0, 0, 3.57)
        worksheet.write(22, 1, "Raw elution time (min)", BORDER_BOLD_BOT)
        worksheet.set_column(1, 1, 11.7)
        worksheet.write(22, 2, "Raw activity in eluate (AIE)", BORDER_BOLD_BOT)
        worksheet.set_column(2, 2, 13)
        worksheet.write(22, 3, "Elution times (parsed)", BORDER_BOLD_BOT)
        worksheet.set_column(3, 3, 11.7)
        worksheet.write(22, 4, "Corrected AIE (cpm)", BORDER_BOLD_BOT)
        worksheet.set_column(4, 4, 10.7)
        worksheet.write(
            22, 5, 
            u"Efflux (cpm \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9)",
            BORDER_BOLD_BOT)
        worksheet.set_column(5, 5, 14)
        worksheet.write(22, 6, "Log(efflux)", BORDER_BOLD_BOT)
        worksheet.set_column(6, 6, 10)
        for index, item in enumerate(analysis.run.elut_ends):
            worksheet.write(23 + index, 0, index + 1)
            worksheet.write(23 + index, 1, analysis.run.elut_ends[index])
            end_elut_ends_parsed = 23 + index
        for index, item in enumerate(analysis.run.raw_cpms):
            worksheet.write(23 + index, 2, item)
        for index, item in enumerate(analysis.run.elut_ends_parsed):
            worksheet.write(23 + index, 3, item, BORDER_LEFT)
        for index, item in enumerate(analysis.run.elut_cpms_gfact):
            worksheet.write(23 + index, 4, item)
        for index, item in enumerate(analysis.run.elut_cpms_gRFW):
            worksheet.write(23 + index, 5, item)
            antilog_chart_end = index + 23
        for index, item in enumerate(analysis.run.elut_cpms_log):
            worksheet.write(23 + index, 6, item, BORDER_LEFT)
            end_log_efflux = 23 + index
        current_col = 7 # Variable because obj analysis adds extra columns
        if analysis.kind == 'obj':
            worksheet.write(
                22, current_col, "Objective regression", BORDER_BOLD_BOT)
            worksheet.set_column(current_col, current_col, 12.7)
            for index, item in enumerate(analysis.r2s):
                if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
                    worksheet.write(23 + index, current_col, item, BOLD)
                    p3_start_row = 23 + index
                else:
                    worksheet.write(23 + index, current_col, item, BORDER_LEFT)
                log_end_row = 23 + index
            current_col += 1
            worksheet.write(
                22, current_col, "Objective slopes", BORDER_BOLD_BOT)
            worksheet.set_column(current_col, current_col, 12.7)
            for index, item in enumerate(analysis.ms):
                if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
                    worksheet.write(23 + index, current_col, item, BOLD)
                else:
                    worksheet.write(23 + index, current_col, item)
            current_col += 1
            worksheet.write(
                22, current_col, "Objective intercepts", BORDER_BOLD_BOT)
            worksheet.set_column(current_col, current_col, 12.7)
            for index, item in enumerate(analysis.bs):
                if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
                    worksheet.write(23 + index, current_col, item, BOLD)
                else:
                    worksheet.write(23 + index, current_col, item)
            current_col += 1

        worksheet.write(
                22, current_col, "Phase III elution times", BORDER_BOLD_BOT)
        worksheet.set_column(current_col, current_col, 12.7)
        for index, item in enumerate(analysis.phase3.x_series):
            worksheet.write(23 + index, current_col, item, BORDER_LEFT)
            p3_chart_end = index + 23
        p3_x_col = current_col
        current_col += 1

        worksheet.write(22, current_col, "Phase III log(efflux)", BORDER_BOLD_BOT)
        worksheet.set_column(current_col, current_col, 10)
        for index, item in enumerate(analysis.phase3.y_series):
            worksheet.write(23 + index, current_col, item)
        p3_y_col = current_col
        current_col += 1

        worksheet.write(
                22, current_col, "Phase II elution times", BORDER_BOLD_BOT)
        worksheet.set_column(current_col, current_col, 12.7)
        for index, item in enumerate(analysis.phase2.x_series):
            worksheet.write(23 + index, current_col, item, BORDER_LEFT)
            p2_chart_end = index + 23
        p2_x_col = current_col
        current_col += 1

        worksheet.write(22, current_col, "Phase II log(efflux)", BORDER_BOLD_BOT)
        worksheet.set_column(current_col, current_col, 10)
        for index, item in enumerate(analysis.phase2.y_series):
            worksheet.write(23 + index, current_col, item)
        p2_y_col = current_col
        current_col += 1

        worksheet.write(
                22, current_col, "Phase I elution times", BORDER_BOLD_BOT)
        worksheet.set_column(current_col, current_col, 12.7)
        for index, item in enumerate(analysis.phase1.x_series):
            worksheet.write(23 + index, current_col, item, BORDER_LEFT)
            p1_chart_end = index + 23
        p1_x_col = current_col
        current_col += 1

        worksheet.write(22, current_col, "Phase I log(efflux)", BORDER_BOLD_BOT)
        worksheet.set_column(current_col, current_col, 9)
        for index, item in enumerate(analysis.phase1.y_series):
            worksheet.write(23 + index, current_col, item)
        p1_y_col = current_col

        # Drawing Summary Chart
        
        chart_all = workbook.add_chart({'type': 'scatter'})
        chart_all.set_title({'name': 'Summary', 'overlay': True})
        worksheet.insert_chart('A8', chart_all)        
        # Adding larger log efflux data to chart_all
        chart_all.add_series({
            'categories': [analysis.run.name, 23, 3, end_elut_ends_parsed, 3],
            'values': [analysis.run.name, 23, 6, end_log_efflux, 6],
            'name' : 'Base log data',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': 'black'},
                       'fill':   {'color': 'white'}} })        
        # Adding Phase III data chart_all
        chart_all.add_series({
            'categories': [analysis.run.name, 23, p3_x_col, p3_chart_end, p3_x_col],
            'values': [analysis.run.name, 23, p3_y_col, p3_chart_end, p3_y_col],
            'name' : 'Phase III',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': '#000000'},
                       'fill':   {'color': '#000000'}} })
        # Adding Phase II data chart_all
        chart_all.add_series({
            'categories': [analysis.run.name, 23, p2_x_col, p2_chart_end, p2_x_col],
            'values': [analysis.run.name, 23, p2_y_col, p2_chart_end, p2_y_col],
            'name' : 'Phase II',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': '#000000'},
                       'fill':   {'color': '#0000CC'}} })
        # Adding Phase I data chart_all
        chart_all.add_series({
            'categories': [analysis.run.name, 23, p1_x_col, p1_chart_end, p1_x_col],
            'values': [analysis.run.name, 23, p1_y_col, p1_chart_end, p1_y_col],
            'name' : 'Phase I',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': '#000000'},
                       'fill':   {'color': '#33CC00'}} })
        if analysis.kind == 'obj':
            # Highlighting last point of Phase III
            chart_all.add_series({
                'categories': [analysis.run.name, 23, p3_x_col, 23, p3_x_col],
                'values': [analysis.run.name, 23, p3_y_col, 23, p3_y_col],
                'name' : 'End of objective regression',
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'red'}} })
            # Highlighting num of pts used for regression
            obj_num_pts = analysis.obj_num_pts
            chart_all.add_series({
                'categories': [
                    analysis.run.name,
                    p3_chart_end-obj_num_pts, p3_x_col,
                    p3_chart_end, p3_x_col],
                'values': [
                    analysis.run.name,
                    p3_chart_end-obj_num_pts, p3_y_col,
                    p3_chart_end, p3_y_col],
                'name' : 'Pts used to initiate regression',
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'red'},
                           'fill':   {'color': 'black'}} })
        # Add p3 regression line
        worksheet.write (7, 0, analysis.phase3.xy1[0])          
        worksheet.write (7, 1, analysis.phase3.xy1[1])          
        worksheet.write (8, 0, analysis.phase3.xy2[0]) 
        worksheet.write (8, 1, analysis.phase3.xy2[1])

        chart_all.add_series({
            'categories': [analysis.run.name, 7, 0, 8, 0],
            'values': [analysis.run.name, 7, 1, 8, 1],
            'line' : {'color': 'red','dash_type' : 'dash'},
            'name' : 'PhIII regression',
            'marker': {'type': 'none'}
            })
        # Add p2 regression line
        worksheet.write (7, 2, analysis.phase2.xy1[0])          
        worksheet.write (7, 3, analysis.phase2.xy1[1])          
        worksheet.write (8, 2, analysis.phase2.xy2[0]) 
        worksheet.write (8, 3, analysis.phase2.xy2[1])

        chart_all.add_series({
            'categories': [analysis.run.name, 7, 2, 8, 2],
            'values': [analysis.run.name, 7, 3, 8, 3],
            'line' : {'color': '#0000CC','dash_type' : 'dash'},
            'name' : 'PhII regression',
            'marker': {'type': 'none'}
            })
        # Add p2 regression line
        worksheet.write (7, 4, analysis.phase1.xy1[0])          
        worksheet.write (7, 5, analysis.phase1.xy1[1])          
        worksheet.write (8, 4, analysis.phase1.xy2[0]) 
        worksheet.write (8, 5, analysis.phase1.xy2[1])

        chart_all.add_series({
            'categories': [analysis.run.name, 7, 4, 8, 4],
            'values': [analysis.run.name, 7, 5, 8, 5],
            'line' : {'color': '#33CC00','dash_type' : 'dash'},
            'name' : 'PhII regression',
            'marker': {'type': 'none'}
            })
        #chart_all.set_legend({'none': True})        
        chart_all.set_x_axis({'name': 'Elution time (min)',})        
        chart_all.set_y_axis({'name': 'Log(efflux(cpm/g RFW/min))',})

        # Drawing the Phase III chart
        chart_p3 = workbook.add_chart({'type': 'scatter'})
        chart_p3.set_title({'name': 'Phase III', 'overlay': True})
        worksheet.insert_chart('G8', chart_p3)        
        # Adding larger log efflux data to chart_p3
        chart_p3.add_series({
            'categories': [analysis.run.name, 23, 3, end_elut_ends_parsed, 3],
            'values': [analysis.run.name, 23, 6, end_log_efflux, 6],
            'name' : 'Base log data',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': 'black'},
                       'fill':   {'color': 'white'}} })        
        # Adding Phase III data chart_p3
        chart_p3.add_series({
            'categories': [analysis.run.name, 23, p3_x_col, p3_chart_end, p3_x_col],
            'values': [analysis.run.name, 23, p3_y_col, p3_chart_end, p3_y_col],
            'name' : 'Phase III',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': '#000000'},
                       'fill':   {'color': '#000000'}} })
        if analysis.kind == 'obj':
            # Highlighting last point of Phase III
            chart_p3.add_series({
                'categories': [analysis.run.name, 23, p3_x_col, 23, p3_x_col],
                'values': [analysis.run.name, 23, p3_y_col, 23, p3_y_col],
                'name' : 'End of objective regression',
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'black'},
                           'fill':   {'color': 'red'}} })
            # Highlighting num of pts used for regression
            obj_num_pts = analysis.obj_num_pts
            chart_p3.add_series({
                'categories': [
                    analysis.run.name,
                    p3_chart_end-obj_num_pts, p3_x_col,
                    p3_chart_end, p3_x_col],
                'values': [
                    analysis.run.name,
                    p3_chart_end-obj_num_pts, p3_y_col,
                    p3_chart_end, p3_y_col],
                'name' : 'Pts used to initiate regression',
                'marker': {'type': 'circle',
                           'size,': 5,
                           'border': {'color': 'red'},
                           'fill':   {'color': 'black'}} })
        # Add p3 regression line
        chart_p3.add_series({
            'categories': [analysis.run.name, 7, 0, 8, 0],
            'values': [analysis.run.name, 7, 1, 8, 1],
            'line' : {'color': 'red','dash_type' : 'dash'},
            'name' : 'PhIII regression',
            'marker': {'type': 'none'}
            })
        #chart_p3.set_legend({'none': True})        
        chart_p3.set_x_axis({'name': 'Elution time (min)',})        
        chart_p3.set_y_axis({'name': 'Log(efflux(cpm/g RFW/min))',})  
        
        # Drawing Phase II chart
        chart_p2 = workbook.add_chart({'type': 'scatter'})
        chart_p2.set_title({'name': 'Phase II', 'overlay': True})
        worksheet.insert_chart('M8', chart_p2)        
        # Adding Phase II data chart_p2
        chart_p2.add_series({
            'categories': [analysis.run.name, 23, p2_x_col, p2_chart_end, p2_x_col],
            'values': [analysis.run.name, 23, p2_y_col, p2_chart_end, p2_y_col],
            'name' : 'Phase II',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': '#000000'},
                       'fill':   {'color': '#0000CC'}} })
        # Add p2 regression line
        worksheet.write (7, 2, analysis.phase2.xy1[0])          
        worksheet.write (7, 3, analysis.phase2.xy1[1])          
        worksheet.write (8, 2, analysis.phase2.xy2[0]) 
        worksheet.write (8, 3, analysis.phase2.xy2[1])

        chart_p2.add_series({
            'categories': [analysis.run.name, 7, 2, 8, 2],
            'values': [analysis.run.name, 7, 3, 8, 3],
            'line' : {'color': '#0000CC','dash_type' : 'dash'},
            'name' : 'PhII regression',
            'marker': {'type': 'none'}
            })
        #chart_p2.set_legend({'none': True})        
        chart_p2.set_x_axis({'name': 'Elution time (min)',})        
        chart_p2.set_y_axis({'name': 'Log(efflux(cpm/g RFW/min))',})

        # Drawing Phase I chart        
        chart_p1 = workbook.add_chart({'type': 'scatter'})
        chart_p1.set_title({'name': 'Phase I', 'overlay': True})
        worksheet.insert_chart('S8', chart_p1)        
        # Adding Phase I data chart_p1
        chart_p1.add_series({
            'categories': [analysis.run.name, 23, p1_x_col, p1_chart_end, p1_x_col],
            'values': [analysis.run.name, 23, p1_y_col, p1_chart_end, p1_y_col],
            'name' : 'Phase I',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': '#000000'},
                       'fill':   {'color': '#33CC00'}} })
        # Add p1 regression line
        chart_p1.add_series({
            'categories': [analysis.run.name, 7, 4, 8, 4],
            'values': [analysis.run.name, 7, 5, 8, 5],
            'line' : {'color': '#33CC00','dash_type' : 'dash'},
            'name' : 'PhI regression',
            'marker': {'type': 'none'}
            })
        #chart_p1.set_legend({'none': True})        
        chart_p1.set_x_axis({'name': 'Elution time (min)',})        
        chart_p1.set_y_axis({'name': 'Log(efflux(cpm/g RFW/min))',})
        
        # Drawing anti-logged data chart        
        chart_antilog = workbook.add_chart({'type': 'scatter'})
        chart_antilog.set_title({'name': 'Anti-logged data', 'overlay': True})
        worksheet.insert_chart('R23', chart_antilog)        
        # Adding antilog data to chart_antilog
        antilog_chart_start = int(len(analysis.run.elut_ends_parsed) * 0.7)
        chart_antilog.add_series({
            'categories': [analysis.run.name,
                 antilog_chart_end-antilog_chart_start, 3,
                 antilog_chart_end, 3],
            'values': [analysis.run.name,
                antilog_chart_end-antilog_chart_start, 5,
                antilog_chart_end, 5],
            'name' : 'Anti-logged partial dataset',
            'marker': {'type': 'circle',
                       'size,': 5,
                       'border': {'color': '#000000'},
                       'fill':   {'color': 'white'}} })
        chart_antilog.set_legend({'none': True})        
        chart_antilog.set_x_axis({'name': 'Elution time (min)',})        
        chart_antilog.set_y_axis({'name': 'Efflux/g RFW/min',})
    
    workbook.close()

def generate_summary(workbook, experiment):
    '''Create a summary sheet in an open <workbook>.

    Summary sheet contains relevant data from all analyses in <experiment>

    @type workbook: Workbook
    @type experiment: Experiment
    @rtype: None
    '''
    # Formatting for not bolded cells (prev. not_req/analyzed_header/not_bold)
    BORDER_BOT = workbook.add_format()
    BORDER_BOT.set_text_wrap()
    BORDER_BOT.set_align('center')
    BORDER_BOT.set_align('vcenter')
    BORDER_BOT.set_bottom()
    # Formatting for bolded cells (previously req)
    BORDER_BOLD_BOT_TOP = workbook.add_format()
    BORDER_BOLD_BOT_TOP.set_text_wrap()
    BORDER_BOLD_BOT_TOP.set_align('center')
    BORDER_BOLD_BOT_TOP.set_align('vcenter')
    BORDER_BOLD_BOT_TOP.set_bold()
    BORDER_BOLD_BOT_TOP.set_bottom()
    BORDER_BOLD_BOT_TOP.set_top()   
    # Formatting for not bolded cells (previously not_req/analyzed_header/not_bold)
    BORDER_TOP = workbook.add_format()
    BORDER_TOP.set_text_wrap()
    BORDER_TOP.set_top()
    # Format for cells input to be aligned in the middle
    VCENTER = workbook.add_format()
    VCENTER.set_text_wrap()
    VCENTER.set_align('vcenter')

    worksheet = generate_sheet(workbook, "Summary", template=False)    
    worksheet.freeze_panes(1,2)
    worksheet.set_column(0, 0, 9)
    # Format left-most column with column headers
    phase_headers = [
        'Start', 'End', "Slope", "Intercept", u"R\u00b2", "k", "Half-Life"]
    CATE_headers = ["Pool Size", "E:I Ratio", "Net flux"]
    
    for index, item in enumerate(CATE_headers):
        worksheet.write(index + 7, 1, item)
    worksheet.write(10, 1, "Influx", BORDER_BOT)
    worksheet.merge_range (11, 0, 18, 0, 'Phase III', BORDER_BOLD_BOT_TOP)
    worksheet.merge_range (19, 0, 26, 0, 'Phase II', BORDER_BOLD_BOT_TOP)
    worksheet.merge_range (27, 0, 34, 0, 'Phase I', BORDER_BOLD_BOT_TOP)
    spacer = len(experiment.analyses[0].run.elut_ends) - 1
    worksheet.merge_range (35, 0, 35+spacer, 0, 'Log (efflux)', BORDER_BOLD_BOT_TOP)
    worksheet.merge_range (
        36+spacer, 0,
        36+(spacer*2), 0,
        u"Efflux (cpm \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9)",
        BORDER_BOLD_BOT_TOP)
    worksheet.merge_range (
        37+(spacer*2), 0,
        37+(spacer*3), 0,
        "Phase III log (efflux)",
        BORDER_BOLD_BOT_TOP)
    worksheet.merge_range (
        38+(spacer*3), 0,
        38+(spacer*4), 0,
        "Phase II log (efflux)",
        BORDER_BOLD_BOT_TOP)
    worksheet.merge_range (
        39+(spacer*4), 0,
        39+(spacer*5), 0,
        "Phase I log (efflux)",
        BORDER_BOLD_BOT_TOP)
    worksheet.merge_range (
        40+(spacer*5), 0,
        40+(spacer*6), 0,
        "Raw activity in eluate (AIE)",
        BORDER_BOLD_BOT_TOP)
    worksheet.merge_range (
        41+(spacer*6), 0,
        41+(spacer*7), 0,
        "Corrected AIE",
        BORDER_BOLD_BOT_TOP)
    for index, item in enumerate(phase_headers):
        worksheet.write(index + 11, 1, item)
    worksheet.write(18, 1, "Efflux", BORDER_BOT)    
    for index, item in enumerate(phase_headers):
        worksheet.write(index + 19, 1, item)
    worksheet.write(26, 1, "Efflux", BORDER_BOT)
    for index, item in enumerate(phase_headers):
        worksheet.write(index + 27, 1, item)
    worksheet.write(34, 1, "Efflux", BORDER_BOT)
    for index, item in enumerate(experiment.analyses[0].run.elut_ends):
        if index == 0:
            worksheet.write(35+index, 1, item, BORDER_TOP)
            worksheet.write(36+index+spacer, 1, item, BORDER_TOP)
            worksheet.write(37+index+(spacer*2), 1, item, BORDER_TOP)
            worksheet.write(38+index+(spacer*3), 1, item, BORDER_TOP)
            worksheet.write(39+index+(spacer*4), 1, item, BORDER_TOP)
            worksheet.write(40+index+(spacer*5), 1, item, BORDER_TOP)
            worksheet.write(41+index+(spacer*6), 1, item, BORDER_TOP)
        else:
            worksheet.write(35+index, 1, item)
            worksheet.write(36+index+spacer, 1, item)
            worksheet.write(37+index+(spacer*2), 1, item)
            worksheet.write(38+index+(spacer*3), 1, item)
            worksheet.write(39+index+(spacer*4), 1, item)
            worksheet.write(40+index+(spacer*5), 1, item)
            worksheet.write(41+index+(spacer*6), 1, item)

    # input basic run information
    for index, analysis in enumerate(experiment.analyses):
        worksheet.write(0, index + 2, analysis.run.name, BORDER_BOT)
        worksheet.write(1, index + 2, analysis.run.SA, VCENTER)
        worksheet.write(2, index + 2, analysis.run.rt_cnts)
        worksheet.write(3, index + 2, analysis.run.sht_cnts)
        worksheet.write(4, index + 2, analysis.run.rt_wght)
        worksheet.write(5, index + 2, analysis.run.gfact)
        worksheet.write(6, index + 2, analysis.run.load_time, BORDER_BOT)

        worksheet.write(7, index + 2, analysis.poolsize)
        worksheet.write(8, index + 2, analysis.ratio)
        worksheet.write(9, index + 2, analysis.netflux)
        worksheet.write(10, index + 2, analysis.influx, BORDER_BOT)

        worksheet.write(11, index + 2, analysis.phase3.xs[0])
        worksheet.write(12, index + 2, analysis.phase3.xs[1])
        worksheet.write(13, index + 2, analysis.phase3.slope)
        worksheet.write(14, index + 2, analysis.phase3.intercept)
        worksheet.write(15, index + 2, analysis.phase3.r2)
        worksheet.write(16, index + 2, analysis.phase3.k)
        worksheet.write(17, index + 2, analysis.phase3.t05)
        worksheet.write(18, index + 2, analysis.phase3.efflux, BORDER_BOT)

        worksheet.write(19, index + 2, analysis.phase2.xs[0])
        worksheet.write(20, index + 2, analysis.phase2.xs[1])
        worksheet.write(21, index + 2, analysis.phase2.slope)
        worksheet.write(22, index + 2, analysis.phase2.intercept)
        worksheet.write(23, index + 2, analysis.phase2.r2)
        worksheet.write(24, index + 2, analysis.phase2.k)
        worksheet.write(25, index + 2, analysis.phase2.t05)
        worksheet.write(26, index + 2, analysis.phase2.efflux, BORDER_BOT)

        worksheet.write(27, index + 2, analysis.phase1.xs[0])
        worksheet.write(28, index + 2, analysis.phase1.xs[1])
        worksheet.write(29, index + 2, analysis.phase1.slope)
        worksheet.write(30, index + 2, analysis.phase1.intercept)
        worksheet.write(31, index + 2, analysis.phase1.r2)
        worksheet.write(32, index + 2, analysis.phase1.k)
        worksheet.write(33, index + 2, analysis.phase1.t05)
        worksheet.write(34, index + 2, analysis.phase1.efflux, BORDER_BOT)

        for index2, item2 in enumerate(analysis.run.elut_ends):
            for index3, item3 in enumerate(analysis.run.elut_ends_parsed):
                if item2 == item3:
                    if index2 == 0: # First item needs top border to delineate
                        worksheet.write(
                            35 + index2, index + 2,
                            analysis.run.elut_cpms_log[index3], BORDER_TOP)
                        worksheet.write(
                            36 + index2 + spacer, index + 2,
                            analysis.run.elut_cpms_gRFW[index3], BORDER_TOP)
                    else:
                        worksheet.write(
                            35 + index2, index + 2,
                            analysis.run.elut_cpms_log[index3])
                        worksheet.write(
                            36 + index2 + spacer, index + 2,
                            analysis.run.elut_cpms_gRFW[index3])
        for index2, item2 in enumerate(analysis.run.elut_ends):
            for index3, item3 in enumerate(analysis.phase3.x_series):
                if item2 == item3:
                    if index2 == 0: # First item needs top border to delineate
                        worksheet.write(
                            37 + index2 + (spacer*2), index + 2,
                            analysis.phase3.y_series[index3], BORDER_TOP)
                    else:
                        worksheet.write(
                            37 + index2 + (spacer*2), index + 2,
                            analysis.phase3.y_series[index3])
        for index2, item2 in enumerate(analysis.run.elut_ends):
            for index3, item3 in enumerate(analysis.phase2.x_series):
                if item2 == item3:
                    if index2 == 0: # First item needs top border to delineate
                        worksheet.write(
                            38 + index2 + (spacer*3), index + 2,
                            analysis.phase2.y_series[index3], BORDER_TOP)
                    else:
                        worksheet.write(
                            38 + index2 + (spacer*3), index + 2,
                            analysis.phase2.y_series[index3])
        for index2, item2 in enumerate(analysis.run.elut_ends):
            for index3, item3 in enumerate(analysis.phase1.x_series):
                if item2 == item3:
                    if index2 == 0: # First item needs top border to delineate
                        worksheet.write(
                            39 + index2 + (spacer*4), index + 2,
                            analysis.phase1.y_series[index3], BORDER_TOP)
                    else:
                        worksheet.write(
                            39 + index2 + (spacer*4), index + 2,
                            analysis.phase1.y_series[index3])
        for index2, item2 in enumerate(analysis.run.elut_ends):
            if index2 == 0: # First item needs top border to delineate
                worksheet.write(
                    40 + index2 + (spacer*5), index + 2,
                    analysis.run.raw_cpms[index2], BORDER_TOP)
            else:
                worksheet.write(
                    40 + index2 + (spacer*5), index + 2,
                    analysis.run.raw_cpms[index2])
        for index2, item2 in enumerate(analysis.run.elut_ends):
            for index3, item3 in enumerate(analysis.run.elut_ends_parsed):
                if item2 == item3:
                    if index2 == 0: # First item needs top border to delineate
                        worksheet.write(
                            41 + index2 + (spacer*6), index + 2,
                            analysis.run.elut_cpms_gfact[index3], BORDER_TOP)
                    else:
                        worksheet.write(
                            41 + index2 + (spacer*6), index + 2,
                            analysis.run.elut_cpms_gfact[index3])

def grab_data(directory, filename):
    '''Extracts data from an excel file in directory/filename.

    Data is used to create/return an Experiment object.

    Precondition: input file is formated according to 
        generate_sheet/generate_template    
    
    @type directory: str
    @type filename: str
    @rtype: Experiment
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
        raw_cpm_column = input_sheet.col(col_index)[8:] # Raw counts given by file
        raw_cpms = []
        elution_cpms = []
        for item in raw_cpm_column:
            raw_cpms.append(item.value)
            if item.value != '':
                elution_cpms.append(float(item.value))
            else:
                elution_cpms.append(0.0)

        temp_run = Objects.Run(
            run_name, SA, root_cnts, shoot_cnts, root_weight, g_factor, 
            load_time, elut_ends, raw_cpms, elution_cpms)
        
        all_analysis_objects.append(Objects.Analysis(
            kind=None, obj_num_pts=None, run=temp_run))
   
    return Objects.Experiment(directory, all_analysis_objects)              
     
if __name__ == "__main__":
    directory = os.path.dirname(os.path.abspath(__file__))
    temp_experiment = grab_data(directory, "/Tests/4/Test_MultiRun1.xlsx")
    #temp_experiment = grab_data(directory, "/Tests/Edge Cases/Test_MissLastPtPh3.xlsx")
    for analysis in temp_experiment.analyses:
        analysis.kind = 'obj'
        analysis.obj_num_pts = 8
        analysis.analyze()
    '''
    temp_experiment.analyses[0].kind = 'obj'
    temp_experiment.analyses[0].obj_num_pts = 8
    temp_experiment.analyses[0].analyze()
    '''
    generate_analysis(temp_experiment)