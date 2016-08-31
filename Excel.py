import os
import time
from xlrd import *
import xlsxwriter

import Objects
'''
# Formatting for basic cells (previously run_header)
basic = workbook.add_format()
basic.set_text_wrap()
basic.set_align('center')
basic.set_align('vcenter')
# Formatting for not bolded cells (previously not_req/analyzed_header/not_bold)
bot_border = workbook.add_format()
bot_border.set_text_wrap()
bot_border.set_align('center')
bot_border.set_align('vcenter')
bot_border.set_bottom()    
# Formatting for bolded cells (previously req)
bold_bot = workbook.add_format()
bold_bot.set_text_wrap()
bold_bot.set_align('center')
bold_bot.set_align('vcenter')
bold_bot.set_bold()
bold_bot.set_bottom()    
# Formatting for cells surrounded by a border (prev. empty_row/basic_format)
border = workbook.add_format()
border.set_align('center')
border.set_align('vcenter')
border.set_top()    
border.set_bottom()    
border.set_right()    
border.set_left()
# Formatting for cells w/right border (prev. phase_format)
right_border = workbook.add_format ()
right_border.set_align('right')
right_border.set_align('vcenter')
right_border.set_right()
'''
def generate_sheet(workbook, sheet_name, template=True):
    """
    Generates and returns the basic excel sheet template (in existing workbook)
    upon which most sheets (which the exception of summary and analysis
    sheets) are built.
    template=True is used to set row where bulk of data is output
    """
    # Formatting for cells surrounded by a border (prev. empty_row/basic_format)
    border = workbook.add_format()
    border.set_align('center')
    border.set_align('vcenter')
    border.set_top()    
    border.set_bottom()    
    border.set_right()    
    border.set_left()
    # Formatting for bolded cells (previously req)
    bold_bot = workbook.add_format()
    bold_bot.set_text_wrap()
    bold_bot.set_align('center')
    bold_bot.set_align('vcenter')
    bold_bot.set_bold()
    bold_bot.set_bottom()
    # Formatting for not bolded cells (prev. not_req/analyzed_header/not_bold)
    bot_border = workbook.add_format()
    bot_border.set_text_wrap()
    bot_border.set_align('center')
    bot_border.set_align('vcenter')
    bot_border.set_bottom()    

    worksheet = workbook.add_worksheet(sheet_name)
    worksheet.set_row(1, 30.75) # Setting the height of the SA row to ~2 lines
    
    # List of ordered tuples containing (title, formatting, column width) 
    # in order they are to be written to the file
    row_headers = [
        u"Specific Activity (cpm \u00B7 \u00B5mol\u207b\u00b9)",
        "Root Cnts (cpm)", "Shoot Cnts (cpm)", "Root weight (g)", "G-Factor", 
        "Load Time (min)"]
    for index, item in enumerate(row_headers):
        worksheet.merge_range (
            index + 1, 0,
            index + 1, 1, 
            row_headers[index], bold_bot)
        worksheet.write (index + 1, 2, "", border)

    if template:
        worksheet.write(
            7, 0,
            "Vial #",
            bot_border)
        worksheet.set_column(0, 0, 3.57)
        worksheet.write(
            7, 1,
            "Elution time (min)",
            bold_bot)
        worksheet.set_column(1, 1, 11.7)
        worksheet.write(
            7, 2,
            "Activity in eluant (cpm)",
            bold_bot)
        worksheet.set_column(2, 2, 15)
      
    return worksheet

def generate_template(workbook):
    '''
    Generates a CATE template sheet in an already created workbook.
    '''
    worksheet = generate_sheet(workbook, 'Template', template=True)

    # Formatting for basic cells (previously run_header)
    basic = workbook.add_format()
    basic.set_text_wrap()
    basic.set_align('center')
    basic.set_align('vcenter')

    # Writing headers columns containing individual runs 
    worksheet.write(0, 2, "Run 1", basic)
    worksheet.write(0, 3, "etc.", basic)
        
def generate_analysis(experiment):
    '''
    Creating an excel file in directory using a preset naming convention
    Data in the file are the product of CATE analysis from a template file
    containing the raw information.
    '''
    output_name = 'CATE Output - ' + time.strftime("(%Y_%m_%d).xlsx")
    workbook = xlsxwriter.Workbook(experiment.directory + "\\" + output_name)
    # Formatting for not bolded cells (previously not_req/analyzed_header/not_bold)
    bot_border = workbook.add_format()
    bot_border.set_text_wrap()
    bot_border.set_align('center')
    bot_border.set_align('vcenter')
    bot_border.set_bottom()
    # Formatting for not bolded cells (previously not_req/analyzed_header/not_bold)
    right_border = workbook.add_format()
    right_border.set_text_wrap()
    right_border.set_align('right')
    right_border.set_right()
    # Formatting for cells surrounded by a border (prev. empty_row/basic_format)
    border = workbook.add_format()
    border.set_align('center')
    border.set_align('vcenter')
    border.set_top()    
    border.set_bottom()    
    border.set_right()    
    border.set_left()
    # Formatting for bolded cells (previously req)
    bold_bot = workbook.add_format()
    bold_bot.set_text_wrap()
    bold_bot.set_align('center')
    bold_bot.set_align('vcenter')
    bold_bot.set_bold()
    bold_bot.set_bottom()
    # Formatting for cells with a left border
    left_border = workbook.add_format()   
    left_border.set_left()
    # Formatting for bolded cells
    bold = workbook.add_format()   
    bold.set_bold()
    
    phasedata_headers = [
        'Start', 'End', "Slope", "Intercept", u"R\u00b2", "k", "Half-Life",
        "Efflux", "Influx", "Net flux", "E:I Ratio", "Pool Size"]
    
    for analysis in experiment.analyses:
        worksheet = generate_sheet(workbook, analysis.run.name, template=False)

        # Writing headers for phase data
        worksheet.write(2, 4, 'Phase I', right_border)
        worksheet.write(3, 4, 'Phase II', right_border)
        worksheet.write(4, 4, 'Phase III', right_border)
        for index, item in enumerate(phasedata_headers):
            worksheet.write(
                1, index + 5, phasedata_headers[index], bot_border)
        # Writing individual phase data
        worksheet.write(0, 2, analysis.run.name, bot_border)    
        worksheet.write(1, 2, analysis.run.SA, border)    
        worksheet.write(2, 2, analysis.run.rt_cnts, border)
        worksheet.write(3, 2, analysis.run.sht_cnts, border)
        worksheet.write(4, 2, analysis.run.rt_wght, border)
        worksheet.write(5, 2, analysis.run.gfact, border)
        worksheet.write(6, 2, analysis.run.load_time, border)
        
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
        worksheet.write(22, 0, "#", bot_border)
        worksheet.set_column(0, 0, 3.57)
        worksheet.write(22, 1, "Raw elution time (min)", bold_bot)
        worksheet.set_column(1, 1, 11.7)
        worksheet.write(22, 2, "Raw activity in eluate (AIE)", bold_bot)
        worksheet.set_column(2, 2, 13)
        worksheet.write(22, 3, "Elution times (parsed)", bold_bot)
        worksheet.set_column(3, 3, 11.7)
        worksheet.write(22, 4, "Corrected AIE (cpm)", bold_bot)
        worksheet.set_column(4, 4, 10.7)
        worksheet.write(
            22, 5, 
            u"Efflux (cpm \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9)",
            bold_bot)
        worksheet.set_column(5, 5, 14)
        worksheet.write(22, 6, "Log(efflux)", bold_bot)
        worksheet.set_column(6, 6, 9)
        for index, item in enumerate(analysis.run.elut_ends):
            worksheet.write(23 + index, 0, index + 1)
            worksheet.write(23 + index, 1, analysis.run.elut_ends[index])
            end_elut_ends_parsed = 23 + index
        for index, item in enumerate(analysis.run.raw_cpms):
            worksheet.write(23 + index, 2, item)
        for index, item in enumerate(analysis.run.elut_ends_parsed):
            worksheet.write(23 + index, 3, item, left_border)
        for index, item in enumerate(analysis.run.elut_cpms_gfact):
            worksheet.write(23 + index, 4, item)
        for index, item in enumerate(analysis.run.elut_cpms_gRFW):
            worksheet.write(23 + index, 5, item)
            antilog_chart_end = index + 23
        for index, item in enumerate(analysis.run.elut_cpms_log):
            worksheet.write(23 + index, 6, item, left_border)
            end_log_efflux = 23 + index
        current_col = 7 # Variable because obj analysis adds extra columns
        if analysis.kind == 'obj':
            worksheet.write(
                22, current_col, "Objective regression", bold_bot)
            worksheet.set_column(current_col, current_col, 12.7)
            for index, item in enumerate(analysis.r2s):
                if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
                    worksheet.write(23 + index, current_col, item, bold)
                    p3_start_row = 23 + index
                else:
                    worksheet.write(23 + index, current_col, item, left_border)
                log_end_row = 23 + index
            current_col += 1
            worksheet.write(
                22, current_col, "Objective slopes", bold_bot)
            worksheet.set_column(current_col, current_col, 12.7)
            for index, item in enumerate(analysis.ms):
                if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
                    worksheet.write(23 + index, current_col, item, bold)
                else:
                    worksheet.write(23 + index, current_col, item)
            current_col += 1
            worksheet.write(
                22, current_col, "Objective intercepts", bold_bot)
            worksheet.set_column(current_col, current_col, 12.7)
            for index, item in enumerate(analysis.bs):
                if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
                    worksheet.write(23 + index, current_col, item, bold)
                else:
                    worksheet.write(23 + index, current_col, item)
            current_col += 1

        worksheet.write(
                22, current_col, "Phase III elution times", bold_bot)
        worksheet.set_column(current_col, current_col, 12.7)
        for index, item in enumerate(analysis.phase3.x_series):
            worksheet.write(23 + index, current_col, item, left_border)
            p3_chart_end = index + 23
        p3_x_col = current_col
        current_col += 1

        worksheet.write(22, current_col, "Phase III log(efflux)", bold_bot)
        worksheet.set_column(current_col, current_col, 9)
        for index, item in enumerate(analysis.phase3.y_series):
            worksheet.write(23 + index, current_col, item)
        p3_y_col = current_col
        current_col += 1

        worksheet.write(
                22, current_col, "Phase II elution times", bold_bot)
        worksheet.set_column(current_col, current_col, 12.7)
        for index, item in enumerate(analysis.phase2.x_series):
            worksheet.write(23 + index, current_col, item, left_border)
            p2_chart_end = index + 23
        p2_x_col = current_col
        current_col += 1

        worksheet.write(22, current_col, "Phase II log(efflux)", bold_bot)
        worksheet.set_column(current_col, current_col, 9)
        for index, item in enumerate(analysis.phase2.y_series):
            worksheet.write(23 + index, current_col, item)
        p2_y_col = current_col
        current_col += 1

        worksheet.write(
                22, current_col, "Phase I elution times", bold_bot)
        worksheet.set_column(current_col, current_col, 12.7)
        for index, item in enumerate(analysis.phase1.x_series):
            worksheet.write(23 + index, current_col, item, left_border)
            p1_chart_end = index + 23
        p1_x_col = current_col
        current_col += 1

        worksheet.write(22, current_col, "Phase I log(efflux)", bold_bot)
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
            'name' : 'Phase II',
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
    #generate_summary(workbook, experiment)     
    workbook.close()

def generate_summary(workbook, experiment):
    '''
    Given an open workbook, create a summary sheet containing relevant data
    from all analyses in experiment.analyses
    '''
    worksheet = workbook.add_worksheet("Summary")
    worksheet.freeze_panes(1,2)
    
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
                           "Phase-corr. Log(eff.)", phase_format)    
    worksheet.merge_range (log_efflux_row, 0, log_efflux_row, 1,\
                           "Log(efflux)", phase_format)
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
        worksheet.write(1, counter, analysis.analysis_type[0])
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
    temp_experiment = grab_data(directory, "/Tests/3/Test_MultiRun1.xlsx")
    #temp_experiment = grab_data(directory, "/Tests/Edge Cases/Test_MissMidPtPh3.xlsx")
    for analysis in temp_experiment.analyses:
        analysis.kind = 'obj'
        analysis.obj_num_pts = 8
        analysis.analyze()
    generate_analysis(temp_experiment)