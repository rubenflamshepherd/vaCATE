import os
import time
from xlrd import *
import xlsxwriter

import Objects

def generate_sheet(workbook, sheet_name, template=False):
	"""Create/return basic excel sheet template (in existing <workbook>)

	This is the template upon which most sheets (which the exception of summary
	and analysis sheets) are built.
	template=True is used to set row where bulk of data is output

	@type workbook: Workbook
		Workbook object into which sheet is inserted
	@type sheet_name: str
		Excel label of Worksheet object that is to be inserted.
	@type Template: bool
	@rtype: Workbook
	"""    
	worksheet = workbook.add_worksheet(sheet_name)
	worksheet.set_row(1, 30.75) # Setting the height of the SA row to ~2 lines

	# Formatting for basic cells (previously run_header/basic)
	basic = workbook.add_format()
	basic.set_text_wrap()
	basic.set_align('center')
	basic.set_align('vcenter')	
	# Formatting for cells surrounded by a border (prev. empty_row/basic_format)
	border = workbook.add_format()
	border.set_align('center')
	border.set_align('vcenter')
	border.set_top()    
	border.set_bottom()    
	border.set_right()    
	border.set_left()
	# Formatting for bolded cells (previously req/border_all)
	border_bold = workbook.add_format()
	border_bold.set_text_wrap()
	border_bold.set_align('center')
	border_bold.set_align('vcenter')
	border_bold.set_bold()
	border_bold.set_bottom()
	border_bold.set_right()    
	border_bold.set_left()
	border_bold.set_top()
	# Formatting for not bolded cells (prev. not_req/analyzed_header/not_bold)
	border_bot = workbook.add_format()
	border_bot.set_text_wrap()
	border_bot.set_align('center')
	border_bot.set_align('vcenter')
	border_bot.set_bottom()
	
	# List of input field labels in order they are to be written to the file
	row_headers = [
		'Run Name',
		u"Specific Activity (cpm \u00B7 \u00B5mol\u207b\u00b9)",
		"Root Cnts (cpm)", "Shoot Cnts (cpm)", "Root weight (g)", "G-Factor", 
		"Load Time (min)"]
	for index, item in enumerate(row_headers):
		worksheet.merge_range (
			index, 0,
			index, 1, 
			row_headers[index], border_bold)
		worksheet.write (index, 2, "", border)

	if template:
		worksheet.write(
			7, 0,
			"Vial #",
			border_bot)
		worksheet.set_column(0, 0, 3.57)
		worksheet.write(
			7, 1,
			"Elution time (min)",
			border_bold)
		worksheet.set_column(1, 1, 11.7)
		worksheet.write(
			7, 2,
			"Activity in eluant (cpm)",
			border_bold)
		worksheet.set_column(2, 2, 15)
		# Writing headers for columns that should contain individual runs 
		worksheet.write(0, 2, "Run 1", basic)
		worksheet.write(0, 3, "etc.", basic)
	else: #  Creating a summary or individual run sheet.
		worksheet.merge_range (
			index + 1, 0,
			index + 1, 1, 
			'Reg Type', border_bold)
		worksheet.write (index, 2, "", border)
	  
	return worksheet


def write_run_labels(worksheet, formats):
	'''Write basic input field labels for the <worksheet> of a single run.

	The are mostly the labels occurring in the leftmost column

	@type worksheet: Worksheet
	@type formats: [Format]
	@rtype: None
	'''
	border_bot, border_right, border_bold_bot = formats
	border_right = formats[1]
	border_bold_bot = formats[2]
	phasedata_headers = [
		'Start', 'End', "Slope", "Intercept", u"R\u00b2", "k", "Half-Life",
		"Efflux", "Influx", "Net flux", "E:I Ratio", "Pool Size"]
	for index, item in enumerate(phasedata_headers):
		worksheet.write(1, index + 5, phasedata_headers[index], border_bot)
	# Writing headers for phase data
	worksheet.write(2, 4, 'Phase I', border_right)
	worksheet.write(3, 4, 'Phase II', border_right)
	worksheet.write(4, 4, 'Phase III', border_right)
	# Data series labels
	worksheet.write(22, 0, "#", border_bot)
	worksheet.set_column(0, 0, 3.57)
	worksheet.write(22, 1, "Raw elution time (min)", border_bold_bot)
	worksheet.set_column(1, 1, 11.7)
	worksheet.write(22, 2, "Raw activity in eluate (AIE)", border_bold_bot)
	worksheet.set_column(2, 2, 13)
	worksheet.write(22, 3, "Elution times (parsed)", border_bold_bot)
	worksheet.set_column(3, 3, 11.7)
	worksheet.write(22, 4, "Corrected AIE (cpm)", border_bold_bot)
	worksheet.set_column(4, 4, 10.7)
	worksheet.write(
		22, 5, 
		u"Efflux (cpm \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9)",
		border_bold_bot)
	worksheet.set_column(5, 5, 14)
	worksheet.write(22, 6, "Log(efflux)", border_bold_bot)
	worksheet.set_column(6, 6, 10)

def write_basic_series(worksheet, formats, run):
	'''Write basic elution series data for the <worksheet> of a single run.

	Counters important to displaying of graphs are returned.

	@type worksheet: Worksheet
	@type formats: [Format]
	@rtype: (int, int)
	'''
	border_left = formats[0]
	for index, item in enumerate(run.elut_ends):
		worksheet.write(23 + index, 0, index + 1)
		worksheet.write(23 + index, 1, run.elut_ends[index])
		end_elut_ends_parsed = 23 + index
	for index, item in enumerate(run.raw_cpms):
		worksheet.write(23 + index, 2, item)
	for index, item in enumerate(run.elut_ends_parsed):
		worksheet.write(23 + index, 3, item, border_left)
	for index, item in enumerate(run.elut_cpms_gfact):
		worksheet.write(23 + index, 4, item)
	for index, item in enumerate(run.elut_cpms_gRFW):
		worksheet.write(23 + index, 5, item)
	for index, item in enumerate(run.elut_cpms_log):
		worksheet.write(23 + index, 6, item, border_left)
		end_log_efflux = 23 + index

	return end_elut_ends_parsed, end_log_efflux

def write_objective(worksheet, formats, analysis):
	'''Write data specific to objective regression.

	Returns a counter for ordering later columns properly.

	@type worksheet: Worksheet
	@type formats: [Format]
	@rtype: int
	'''
	border_bold_bot, bold, border_left = formats
	current_col = 7 # Tracks current column to be written in
	if analysis.kind == 'obj':
		worksheet.write(
			22, current_col, "Objective regression", border_bold_bot)
		worksheet.set_column(current_col, current_col, 12.7)
		for index, item in enumerate(analysis.r2s):
			if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
				worksheet.write(23 + index, current_col, item, bold)
				p3_start_row = 23 + index
			else:
				worksheet.write(23 + index, current_col, item, border_left)
			log_end_row = 23 + index
		current_col += 1
		worksheet.write(
			22, current_col, "Objective slopes", border_bold_bot)
		worksheet.set_column(current_col, current_col, 12.7)
		for index, item in enumerate(analysis.ms):
			if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
				worksheet.write(23 + index, current_col, item, bold)
			else:
				worksheet.write(23 + index, current_col, item)
		current_col += 1
		worksheet.write(
			22, current_col, "Objective intercepts", border_bold_bot)
		worksheet.set_column(current_col, current_col, 12.7)
		for index, item in enumerate(analysis.bs):
			if analysis.run.elut_ends_parsed[index] == analysis.xs_p3[0]:
				worksheet.write(23 + index, current_col, item, bold)
			else:
				worksheet.write(23 + index, current_col, item)
		current_col += 1
	return current_col

def write_phase(worksheet, formats, first_row, first_col, phase, vertical=True):
	'''Write <phase> attributes to excel <worksheet>.

	<first_row> and <first_col> used to position attributes.
	len(formats) = 1

	@type first_row: int
	@type first_col: int
	@type worksheet: Worksheet
	@type phase: Phase
	@type formats: [Format] | [None]
	@rtype: None
	'''
	border_bot = formats[0]
	if vertical:
		worksheet.write(first_row, first_col, phase.xs[0])
		worksheet.write(first_row + 1, first_col, phase.xs[1])
		worksheet.write(first_row + 2, first_col, phase.slope)
		worksheet.write(first_row + 3, first_col, phase.intercept)
		worksheet.write(first_row + 4, first_col, phase.r2)
		worksheet.write(first_row + 5, first_col, phase.k)
		worksheet.write(first_row + 6, first_col, phase.t05)
		worksheet.write(first_row + 7, first_col, phase.efflux, border_bot)
	else:
		worksheet.write(first_row, first_col, phase.xs[0])
		worksheet.write(first_row, first_col + 1, phase.xs[1])
		worksheet.write(first_row, first_col + 2, phase.slope)
		worksheet.write(first_row, first_col + 3, phase.intercept)
		worksheet.write(first_row, first_col + 4, phase.r2)
		worksheet.write(first_row, first_col + 5, phase.k)
		worksheet.write(first_row, first_col + 6, phase.t05)
		worksheet.write(first_row, first_col + 7, phase.efflux, border_bot)

def write_phase_series(worksheet, formats, current_col, phase, phase_num):
	'''Write x- and y-series data specific to a particular phase.

	Returns counters for drawing graphs

	@type worksheet: Worksheet
	@type formats: [Format]
	@type current_col: int
	@type phase: Phase
	@type phase_num: 'I' | 'II' | 'III'
	@rtype: (int, int, int, int)
	'''
	border_bold_bot, border_left = formats
	worksheet.write(
			22, current_col,
			"Phase " + phase_num + " times", border_bold_bot)
	worksheet.set_column(current_col, current_col, 12.7)
	index = 0 #  Need to instantiate because loop below doesn't always
	for index, item in enumerate(phase.x_series):
		worksheet.write(23 + index, current_col, item, border_left)
	chart_end = index + 23
	x_col = current_col
	current_col += 1

	worksheet.write(
		22, current_col,
		"Phase " + phase_num + " log(efflux)", border_bold_bot)
	worksheet.set_column(current_col, current_col, 10)
	for index, item in enumerate(phase.y_series):
		worksheet.write(23 + index, current_col, item)
	y_col = current_col
	current_col += 1
	return current_col, x_col, y_col, chart_end

def write_summary_chart(
		workbook, worksheet, analysis,
		end_elut_ends_parsed, end_log_efflux,
		chart3_data, chart2_data, chart1_data):
	''' Draws summary chart in <worksheet> for individual runs.

	@type workbook: Workbook
	@type worksheet: Worksheet
	@type analysis: Analysis
	@type end_elut_ends_parsed: int
		Counter for graphing
	@type end_log_efflux: int
		Counter for graphing
	@type chart3_data: (int, int, int)
		locations of data ranges in <worksheet> for graphing phase III
	@type chart2_data: (int, int, int)
		locations of data ranges in <worksheet> for graphing phase II
	@type chart1_data: (int, int, int)
		locations of data ranges in <worksheet> for graphing phase I			
	@rtype: None
	'''
	p3_x_col, p3_y_col, p3_chart_end = chart3_data
	p2_x_col, p2_y_col, p2_chart_end = chart2_data
	p1_x_col, p1_y_col, p1_chart_end = chart1_data

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
	# Adding Phase III data to chart_all
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


def write_phase_chart(
		workbook, worksheet, analysis, chart_data, phase_num,
		end_elut_ends_parsed=None, end_log_efflux=None,):
	''' Draws individual phase charts in <worksheet> for individual runs.

	Precondition: if <phase_num> = 'III' then <end_elut_ends_parsed> and
		<end_log_efflux> != None

	@type workbook: Workbook
	@type worksheet: Worksheet
	@type analysis: Analysis
	@type chart_data: (int, int, int)
		locations of data ranges in <worksheet> for graphing phase
	@type phase_num: 'III' | 'II' | 'I'
	@type end_elut_ends_parsed: int
		Counter for graphing. Only needed if <phase_num> == 3.
	@type end_log_efflux: int
		Counter for graphing. Only needed if <phase_num> == 3.			
	@rtype: None
	'''
	x_col, y_col, chart_end = chart_data
	# Assigning constants that depend on which phase is being analyzed
	if phase_num == 'III':
		insert_cell, fill_color, reg_col = 'G8', '#000000', 0
	elif phase_num == 'II':
		insert_cell, fill_color, reg_col = 'M8', '#0000CC', 2
	elif phase_num == 'I':
		insert_cell, fill_color, reg_col = 'S8', '#33CC00', 4
	# Drawing the Phase III chart
	chart = workbook.add_chart({'type': 'scatter'})
	chart.set_title({'name': 'Phase ' + phase_num, 'overlay': True})
	worksheet.insert_chart(insert_cell, chart)
	if phase_num == 'III':
		# Adding larger log efflux data to chart
		chart.add_series({
			'categories': [analysis.run.name, 23, 3, end_elut_ends_parsed, 3],
			'values': [analysis.run.name, 23, 6, end_log_efflux, 6],
			'name' : 'Base log data',
			'marker': {'type': 'circle',
					   'size,': 5,
					   'border': {'color': 'black'},
					   'fill':   {'color': 'white'}} }) 

	# Adding Phase data to chart
	chart.add_series({
		'categories': [analysis.run.name, 23, x_col, chart_end, x_col],
		'values': [analysis.run.name, 23, y_col, chart_end, y_col],
		'name' : 'Phase ' + phase_num,
		'marker': {'type': 'circle',
				   'size,': 5,
				   'border': {'color': '#000000'},
				   'fill':   {'color': fill_color}} })
	if analysis.kind == 'obj' and phase_num == 'III':
		# Highlighting last point of Phase III
		chart.add_series({
			'categories': [analysis.run.name, 23, x_col, 23, x_col],
			'values': [analysis.run.name, 23, y_col, 23, y_col],
			'name' : 'End of objective regression',
			'marker': {'type': 'circle',
					   'size,': 5,
					   'border': {'color': 'black'},
					   'fill':   {'color': 'red'}} })
		# Highlighting num of pts used for regression
		obj_num_pts = analysis.obj_num_pts
		chart.add_series({
			'categories': [
				analysis.run.name,
				chart_end-obj_num_pts, x_col,
				chart_end, x_col],
			'values': [
				analysis.run.name,
				chart_end-obj_num_pts, y_col,
				chart_end, y_col],
			'name' : 'Pts used to initiate regression',
			'marker': {'type': 'circle',
					   'size,': 5,
					   'border': {'color': 'red'},
					   'fill':   {'color': 'black'}} })
	# Add regression line
	chart.add_series({
		'categories': [analysis.run.name, 7, reg_col, 8, reg_col],
		'values': [analysis.run.name, 7, reg_col + 1, 8, reg_col + 1],
		'line' : {'color': 'red','dash_type' : 'dash'},
		'name' : 'Ph' + phase_num + ' regression',
		'marker': {'type': 'none'}
		})
	#chart.set_legend({'none': True})        
	chart.set_x_axis({'name': 'Elution time (min)',})        
	chart.set_y_axis({'name': 'Log(efflux(cpm/g RFW/min))',})

def write_antilog_chart(workbook, worksheet, analysis, end_log_efflux):
	''' Draws antilog chart in <worksheet> for individual runs.

	@type workbook: Workbook
	@type worksheet: Worksheet
	@type analysis: Analysis
	@type end_log_efflux: int
		Counter for graphing
	@rtype: None
	'''
	# Drawing anti-logged data chart        
	chart_antilog = workbook.add_chart({'type': 'scatter'})
	chart_antilog.set_title({'name': 'Anti-logged data', 'overlay': True})
	worksheet.insert_chart('R23', chart_antilog)        
	# Adding antilog data to chart_antilog
	antilog_chart_start = int(len(analysis.run.elut_ends_parsed) * 0.7)
	chart_antilog.add_series({
		'categories': [analysis.run.name,
			 end_log_efflux-antilog_chart_start, 3,
			 end_log_efflux, 3],
		'values': [analysis.run.name,
			end_log_efflux-antilog_chart_start, 5,
			end_log_efflux, 5],
		'name' : 'Anti-logged partial dataset',
		'marker': {'type': 'circle',
				   'size,': 5,
				   'border': {'color': '#000000'},
				   'fill':   {'color': 'white'}} })
	chart_antilog.set_legend({'none': True})        
	chart_antilog.set_x_axis({'name': 'Elution time (min)',})        
	chart_antilog.set_y_axis({'name': 'Efflux/g RFW/min',})


def write_basic_calculations(
		worksheet, formats, analysis, first_row, first_col, summary=True):
	'''Write <phase> attributes to excel <worksheet>.

	Preconditions: len(formats) = 1

	@type worksheet: Worksheet
	@type formats: [Format] | [None]
	@type analysis: Analysis
	@type first_row: int
		Used to position output of calculations.
	@type first_col: int
		Used to position output of calculations.
	@type summary: bool
	@rtype: None
	'''
	border_bot = formats[0]
	worksheet.write(first_row, first_col, analysis.run.name, border_bot)
	worksheet.write(first_row + 1, first_col, analysis.run.SA)
	worksheet.write(first_row + 2, first_col, analysis.run.rt_cnts)
	worksheet.write(first_row + 3, first_col, analysis.run.sht_cnts)
	worksheet.write(first_row + 4, first_col, analysis.run.rt_wght)
	worksheet.write(first_row + 5, first_col, analysis.run.gfact)
	worksheet.write(first_row + 6, first_col, analysis.run.load_time)
	worksheet.write(first_row + 7, first_col, analysis.kind, border_bot)
	if summary:
		worksheet.write(first_row + 8, first_col, analysis.poolsize)
		worksheet.write(first_row + 9, first_col, analysis.ratio)
		worksheet.write(first_row + 10, first_col, analysis.netflux)
		worksheet.write(first_row + 11, first_col, analysis.influx, border_bot)
	else:
		worksheet.write(4, 13, analysis.influx)
		worksheet.write(4, 14, analysis.netflux)
		worksheet.write(4, 15, analysis.ratio)
		worksheet.write(4, 16, analysis.poolsize, border_bot)	


def write_constant_row_labels(worksheet, formats):
	'''Fill in the left-most row of the summary worksheet with constant labels.

	Constant labels are those whose position doesn't change depending on length
		of various data series.

	@type worksheet: Worksheet
	@type formats: [Format]
	@rtype: None
	'''
	border_bold_bot_top, border_bot = formats

	worksheet.freeze_panes(1,2)
	worksheet.set_column(0, 0, 9)

	CATE_headers = ["Pool Size", "E:I Ratio", "Net flux"]
	for index, item in enumerate(CATE_headers):
		worksheet.write(index + 8, 1, item)
	worksheet.write(index + 9, 1, "Influx", border_bot)

	phase_headers = [
		'Start', 'End', "Slope", "Intercept", u"R\u00b2", "k", "Half-Life"]
	for index, item in enumerate(phase_headers):
		worksheet.write(index + 12, 1, item)
	worksheet.write(19, 1, "Efflux", border_bot)    
	for index, item in enumerate(phase_headers):
		worksheet.write(index + 20, 1, item)
	worksheet.write(27, 1, "Efflux", border_bot)
	for index, item in enumerate(phase_headers):
		worksheet.write(index + 28, 1, item)
	worksheet.write(35, 1, "Efflux", border_bot)

	worksheet.merge_range (12, 0, 19, 0, 'Phase III', border_bold_bot_top)
	worksheet.merge_range (20, 0, 27, 0, 'Phase II', border_bold_bot_top)
	worksheet.merge_range (28, 0, 35, 0, 'Phase I', border_bold_bot_top)	


def write_series_row_labels(worksheet, formats, spacer, elut_ends):
	'''Fill in the left-most row of the summary worksheet with series labels.

	The position of the series' labels change depending on the length of the 
		series. <spacer> is used to account for this.
	
	@type worksheet: Worksheet
	@type formats: [Format]
	@type spacer: int
	@type elut_ends: [float]
	@rtype: None
	'''
	border_bold_bot_top, border_top = formats

	worksheet.merge_range(
		36, 0,
		36 + spacer, 0,
		'Log (efflux)', border_bold_bot_top)
	worksheet.merge_range(
		37 + spacer, 0,
		37 + (spacer*2), 0,
		u"Efflux (cpm \u00B7 min\u207b\u00b9 \u00B7 g RFW\u207b\u00b9)",
		border_bold_bot_top)
	worksheet.merge_range(
		38 + (spacer*2), 0,
		38 + (spacer*3), 0,
		"Phase III log (efflux)", border_bold_bot_top)
	worksheet.merge_range(
		39 + (spacer*3), 0,
		39 + (spacer*4), 0,
		"Phase II log (efflux)", border_bold_bot_top)
	worksheet.merge_range(
		40 + (spacer*4), 0,
		40 + (spacer*5), 0,
		"Phase I log (efflux)", border_bold_bot_top)
	worksheet.merge_range(
		41 + (spacer*5), 0,
		41 + (spacer*6), 0,
		"Raw activity in eluate (AIE)", border_bold_bot_top)
	worksheet.merge_range(
		42 + (spacer*6), 0,
		42 + (spacer*7), 0,
		"Corrected AIE", border_bold_bot_top)

	for index, item in enumerate(elut_ends):
		if index == 0:
			worksheet.write(36 + index,  1, item, border_top)
			worksheet.write(37 + index + spacer, 1, item, border_top)
			worksheet.write(38 + index + (spacer*2), 1, item, border_top)
			worksheet.write(39 + index + (spacer*3), 1, item, border_top)
			worksheet.write(40 + index + (spacer*4), 1, item, border_top)
			worksheet.write(41 + index + (spacer*5), 1, item, border_top)
			worksheet.write(42 + index + (spacer*6), 1, item, border_top)
		else:
			worksheet.write(36 + index, 1, item)
			worksheet.write(37 + index + spacer, 1, item)
			worksheet.write(38 + index + (spacer*2), 1, item)
			worksheet.write(39 + index + (spacer*3), 1, item)
			worksheet.write(40 + index + (spacer*4), 1, item)
			worksheet.write(41 + index + (spacer*5), 1, item)
			worksheet.write(42 + index + (spacer*6), 1, item)


def write_series(worksheet, formats, row, col, x_series, y_series, raw_series):
	'''Write a y series of data into a col

	Data is written in a column starting from <row> and <col>.
	Data is spaced according to a <raw_series> whose length is greater than or
		equal to x_series. 
	We iterate along the larger <raw_series> inputting points from <y_series> 
		only if that point's correpondant in <x_series> matches the point from
		<raw_series>. If no points match then a blank cell is left.
	This allows us to space all data according to larger <raw_series>
	
	@type workbook: Worksheet
	@type formats: [Format]
	@type row: int
	@type col: int
	@type x_series: [int | float]
	@type y_series: [int | float]
	@type raw_series: [int | float]
	@rtype: None
	'''
	border_top = formats[0]
	for index2, item2 in enumerate(raw_series):
		for index3, item3 in enumerate(x_series):
			if item2 == item3:
				if index2 == 0: # First item needs top border to delineate
					worksheet.write(
						row + index2, col,
						y_series[index3], border_top)
				else:
					worksheet.write(
						row + index2, col,
						y_series[index3])


def generate_summary(workbook, experiment, formats):
	'''Create a summary sheet in an open <workbook>.

	Summary sheet contains relevant data from all analyses in <experiment>

	@type workbook: Workbook
	@type experiment: Experiment
	@type formats: [Format]
	@rtype: None
	'''
	border_bold_bot_top, border_bot, border_top = formats

	worksheet = generate_sheet(workbook, "Summary", template=False)    
	write_constant_row_labels(worksheet, [border_bold_bot_top, border_bot])
	spacer = len(experiment.analyses[0].run.elut_ends) - 1
	elution_series = experiment.analyses[0].run.elut_ends
	write_series_row_labels(
		worksheet, [border_bold_bot_top, border_top],
		spacer, elution_series)

	# input basic run information
	for index, analysis in enumerate(experiment.analyses):

		write_basic_calculations(
			worksheet, [border_bot], analysis,  0, index + 2, summary=True)
		write_phase(
			worksheet, [border_bot], 12, index + 2, analysis.phase3)
		write_phase(
			worksheet, [border_bot], 20, index + 2, analysis.phase2)
		write_phase(
			worksheet, [border_bot], 28, index + 2, analysis.phase1)

		write_series(
			worksheet, [border_top],
			36, index + 2,
			analysis.run.elut_ends_parsed, analysis.run.elut_cpms_log,
			analysis.run.elut_ends)
		write_series(
			worksheet, [border_top],
			37 + spacer, index + 2,
			analysis.run.elut_ends_parsed, analysis.run.elut_cpms_gRFW,
			analysis.run.elut_ends)
		write_series(
			worksheet, [border_top],
			38 + (spacer*2), index + 2,
			analysis.phase3.x_series, analysis.phase3.y_series,
			analysis.run.elut_ends)
		write_series(
			worksheet, [border_top],
			39 + (spacer*3), index + 2,
			analysis.phase2.x_series, analysis.phase2.y_series,
			analysis.run.elut_ends)		
		write_series(
			worksheet, [border_top],
			40 + (spacer*4), index + 2,
			analysis.phase1.x_series, analysis.phase1.y_series,
			analysis.run.elut_ends)	
		write_series(
			worksheet, [border_top],
			41 + (spacer*5), index + 2,
			analysis.run.elut_ends, analysis.run.raw_cpms,
			analysis.run.elut_ends)	
		write_series(
			worksheet, [border_top],
			42 + (spacer*6), index + 2,
			analysis.run.elut_ends_parsed, analysis.run.elut_cpms_gfact,
			analysis.run.elut_ends)	


def generate_analysis(experiment):
	'''Creating an excel file in <experiment>.directory.

	Excel file contains comprehensive data analysis as created by the user.
	File is named using a preset naming convention.

	Precondition: Data in the file are the product of a template file with with
		CATE data inputted properly.

	@type experiment: Experiment
	@rtype: None
	'''
	output_name = 'CATE Output - ' + time.strftime("(%Y_%m_%d).xlsx")
	workbook = xlsxwriter.Workbook(experiment.directory + "\\" + output_name)

	# Formatting for bolded cells
	bold = workbook.add_format()   
	bold.set_bold()
	# Formatting for not bolded cells
	border_bot = workbook.add_format()
	border_bot.set_text_wrap()
	border_bot.set_align('center')
	border_bot.set_align('vcenter')
	border_bot.set_bottom()
	# Formatting for bolded cells
	border_bold_bot = workbook.add_format()
	border_bold_bot.set_text_wrap()
	border_bold_bot.set_align('center')
	border_bold_bot.set_align('vcenter')
	border_bold_bot.set_bold()
	border_bold_bot.set_bottom()
	# Formatting for bolded cells
	border_bold_bot_top = workbook.add_format()
	border_bold_bot_top.set_text_wrap()
	border_bold_bot_top.set_align('center')
	border_bold_bot_top.set_align('vcenter')
	border_bold_bot_top.set_bold()
	border_bold_bot_top.set_bottom()
	border_bold_bot_top.set_top()  
	# Formatting for cells with a left border
	border_left = workbook.add_format()   
	border_left.set_left()
	# Formatting for not bolded cells
	border_right = workbook.add_format()
	border_right.set_text_wrap()
	border_right.set_align('right')
	border_right.set_right()
	# Formatting for not bolded cells
	border_top = workbook.add_format()
	border_top.set_text_wrap()
	border_top.set_top()	

	generate_summary(
		workbook, experiment, [border_bold_bot_top, border_bot, border_top])     
	
	for analysis in experiment.analyses:
		worksheet = generate_sheet(workbook, analysis.run.name, template=False)
		write_run_labels(
			worksheet, [border_bot, border_right, border_bold_bot])
		write_basic_calculations(
			worksheet, [None], analysis, 0, 2, summary=False)
		write_phase(
			worksheet, [None], 2, 5, analysis.phase3, vertical=False)
		write_phase(
			worksheet, [None], 3, 5, analysis.phase2, vertical=False)
		write_phase(
			worksheet, [None], 4, 5, analysis.phase1, vertical=False)	

		end_elut_ends_parsed, end_log_efflux = write_basic_series(
			worksheet, [border_left], analysis.run)			
				
		current_col = write_objective(  # obj analysis adds extra columns
			worksheet, [border_bold_bot, bold, border_left], analysis) 

		current_col, p3_x_col, p3_y_col, p3_chart_end = write_phase_series(
			worksheet,
			[border_bold_bot, border_left],
			current_col, analysis.phase3, 'III')
		current_col, p2_x_col, p2_y_col, p2_chart_end = write_phase_series(
			worksheet,
			[border_bold_bot, border_left],
			current_col, analysis.phase2, 'II')
		current_col, p1_x_col, p1_y_col, p1_chart_end = write_phase_series(
			worksheet,
			[border_bold_bot, border_left],
			current_col, analysis.phase1, 'I')
		
		write_summary_chart(
			workbook, worksheet, analysis,
			end_elut_ends_parsed, end_log_efflux,
			(p3_x_col, p3_y_col, p3_chart_end),
			(p2_x_col, p2_y_col, p2_chart_end),
			(p1_x_col, p1_y_col, p1_chart_end))

		write_phase_chart(
			workbook, worksheet, analysis,
			(p3_x_col, p3_y_col, p3_chart_end), 'III',
			end_elut_ends_parsed, end_log_efflux)
		write_phase_chart(
			workbook, worksheet, analysis,
			(p2_x_col, p2_y_col, p2_chart_end), 'II')
		write_phase_chart(
			workbook, worksheet, analysis,
			(p1_x_col, p1_y_col, p1_chart_end), 'I')
		write_antilog_chart(workbook, worksheet, analysis, end_log_efflux)
		
	workbook.close()


def grab_data(input_file):
	'''Extracts data from an excel file in directory/filename.

	Data is used to create/return an Experiment object.

	Precondition: input file is formated according to 
		generate_sheet/generate_template    
	
	@type input_file: path
	@rtype: Experiment
	'''
	# Accessing the file from which data is to be grabbed
	#    input_file = os.path.join(directory, filename)
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
   
	return Objects.Experiment(os.path.dirname(input_file), all_analysis_objects)              
	 
if __name__ == "__main__":
	directory = os.path.dirname(os.path.abspath(__file__))
	file_path = os.path.join(directory, "Tests/4/Test_MultiRun1.xlsx")

	temp_experiment = grab_data(file_path)
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