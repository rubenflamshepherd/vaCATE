import xlrd
from nose.tools import assert_equals

import Excel
import Objects

class TestExperiment(object):
	def __init__(self, directory, analyses):
		self.directory = directory
		self.analyses = analyses # List of Analysis objects

class TestRun(object):
	'''
	Class that stores ALL data of a single CATE run.
	This data includes values derived from objective or 
	subjetive analyses (calculated within the class)
	'''        
	def __init__(
		   self, run_name, SA, rt_cnts, sht_cnts, rt_wght, gfact, load_time,
		   elut_ends, elut_cpms, elut_cpms_gfact, elut_cpms_gRFW, elut_cpms_log):       
		self.run_name = run_name
		self.SA = SA
		self.rt_cnts = rt_cnts
		self.sht_cnts = sht_cnts
		self.rt_wght = rt_wght
		self.gfact = gfact
		self.load_time = load_time
		self.elut_ends = elut_ends
		self.elut_cpms = elut_cpms        
		self.elut_starts = [0.0] + elut_ends[:-1]
		# Basic analysis of CATE data GRABBED FROM EXCEL
		self.elut_cpms_gfact = elut_cpms_gfact
		self.elut_cpms_gRFW = elut_cpms_gRFW
		self.elut_cpms_log = elut_cpms_log



def grab_answers(directory, filename, elut_ends):
	'''
	Extracts data from an excel file in directory/filename (INPUT)
	elut_ends required because can not be cleanly extracted from test sheet
	Data are what analysis SHOULD be arriving at.
	'''
	# Accessing the file from which data is to be grabbed
	input_file = '/'.join((directory, filename))
	input_book = xlrd.open_workbook(input_file)
	input_sheet = input_book.sheet_by_index(1)

	# List where all run info is stored with TestRun objects as ind. entries
	all_test_runs = []

	for col_index in range(2, input_sheet.row_len(0)):        
		# Grab individual CATE values of interest
		run_name = input_sheet.cell(0, col_index).value
		SA = input_sheet.cell(1, col_index).value
		root_cnts = input_sheet.cell(2, col_index).value
		shoot_cnts = input_sheet.cell(3, col_index).value
		root_weight = input_sheet.cell(4, col_index).value
		g_factor = input_sheet.cell(5, col_index).value
		load_time = input_sheet.cell(6, col_index).value
		# Grabing elution cpms, correcting for header offset (8)
		start_cpms = 8
		end_cpms = start_cpms + len(elut_ends)
		raw_cpms = input_sheet.col(col_index)[start_cpms : end_cpms]
		elut_cpms = [float(x.value) for x in raw_cpms]

		start_gfact = end_cpms + 1 # row underneath cpms_gfact label
		end_gfact = start_gfact + len(elut_ends)
		raw_gfact = input_sheet.col(col_index)[start_gfact : end_gfact]
		elut_cpms_gfact = [float(x.value) for x in raw_gfact]

		start_gRFW = end_gfact + 1
		end_gRFW = start_gRFW + len(elut_ends)
		raw_gRFW = input_sheet.col(col_index)[start_gRFW : end_gRFW]
		elut_cpms_gRFW = [float(x.value) for x in raw_gRFW]

		start_log = end_gRFW + 1
		end_log = start_log + len(elut_ends)
		raw_log = input_sheet.col(col_index)[start_log : end_log]
		elut_cpms_log = [float(x.value) for x in raw_log]

		start_r2s = end_log + 1
		end_r2s = start_r2s + len(elut_ends)
		raw_r2s = input_sheet.col(col_index)[start_r2s : end_r2s]
		r2s = []
		for item in raw_r2s: # item is a cell.Value
			try:
				r2s.append(float(item.value))
			except ValueError: # last row in r2 has no r2
				r2s.append(None) # Messes up list comprehension
		
		end_obj = int(input_sheet.col(col_index)[end_r2s].value) 

		start_r2_p1 = end_r2s + 1 + 1 + 1 # First index has no r2
		end_r2_p1 = start_r2_p1 + (end_obj - 3)
		raw_r2_p1 = input_sheet.col(col_index)[start_r2_p1 : end_r2_p1]
		r2_p1 = [float(x.value) for x in raw_r2_p1]

		# Bypass ignored min p1 reg length (+3)
		start_r2_p2 = end_r2s + 1 + len(elut_ends) + 1 + 3
		end_r2_p2 = start_r2_p2 + (end_obj - 3)		
		raw_r2_p2 = input_sheet.col(col_index)[start_r2_p2 : end_r2_p2]
		r2_p2 = [float(x.value) for x in raw_r2_p2]

		# Find index starting p2 (end index of p1)
		start_p12_r2_sum = end_r2s + 1 + len(elut_ends) + 1 + len(elut_ends) + 2
		end_p12_r2_sum = start_p12_r2_sum + len(elut_ends)
		p12_r2_max = input_sheet.col(col_index)[end_p12_r2_sum + 1].value
		print r2s
		all_test_runs.append(
			TestRun(
				run_name, SA, root_cnts, shoot_cnts, root_weight, g_factor,
				load_time, elut_ends, elut_cpms, elut_cpms_gfact, 
				elut_cpms_gRFW, elut_cpms_log))

		return TestExperiment(directory, all_test_runs)

def test_basic():
	'''
	Tests regarding basic info (Data encoded into Run object, no analysis)
	'''
	directory = r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1"
	test_file = "Test - Single Run.xlsx"

	question_experiment = Excel.grab_data(directory, test_file)
	question = question_experiment.analyses[0]
	answer_experiment = grab_answers(directory, test_file, question.run.elut_ends)
	answer = answer_experiment.analyses[0]

	#Objects.find_obj_reg(single_run.elut_ends, single_run.elut_cpms_log, 3)
	assert_equals (question.run.SA, answer.SA)
	assert_equals (question.run.run_name, answer.run_name)
	assert_equals (question.run.rt_cnts, answer.rt_cnts)
	assert_equals (question.run.sht_cnts, answer.sht_cnts)
	assert_equals (question.run.rt_wght, answer.rt_wght)
	assert_equals (question.run.gfact, answer.gfact)
	assert_equals (question.run.load_time, answer.load_time)
	assert_equals (question.run.elut_ends, answer.elut_ends)
	assert_equals (question.run.elut_cpms, answer.elut_cpms)
	assert_equals (question.run.elut_starts, answer.elut_starts)
	assert_equals (question.run.elut_cpms_gfact, answer.elut_cpms_gfact)
	assert_equals (question.run.elut_cpms_gRFW, answer.elut_cpms_gRFW)
	assert_equals (question.run.elut_cpms_log, answer.elut_cpms_log)

if __name__ == '__main__':
	
	import Excel
	temp_data = Excel.grab_data(r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1", "Test - Single Run.xlsx")
	temp_analysis = temp_data.analyses[0]
	
	temp_exp = grab_answers(
		r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1", "Test - Single Run.xlsx", temp_analysis.run.elut_ends)
	temp_testanalysis = temp_exp.analyses[0]
	