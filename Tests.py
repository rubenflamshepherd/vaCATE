import xlrd
from nose.tools import assert_equals
from nose_parameterized import parameterized

import Excel
import Objects

class TestExperiment(object):
	def __init__(self, directory, analyses):
		self.directory = directory
		self.analyses = analyses # List of Analysis objects

class TestAnalysis(object):
	'''
	Class that stores ALL data of a single CATE run.
	This data includes values derived from objective or 
	subjetive analyses (calculated within the class)
	'''        
	def __init__(
		   self, name, SA, rt_cnts, sht_cnts, rt_wght, gfact, load_time, elut_period,
		   elut_ends, elut_cpms, elut_cpms_gfact, elut_cpms_gRFW, elut_cpms_log,
		   r2s, influx, netflux, ratio, poolsize, tracer_retained,
		   p12_r2_max, p12_r2_max_row, phase3, phase2, phase1):       
		self.name = name
		self.SA = SA
		self.rt_cnts = rt_cnts
		self.sht_cnts = sht_cnts
		self.rt_wght = rt_wght
		self.gfact = gfact
		self.load_time = load_time
		self.elut_period = elut_period
		self.elut_ends = elut_ends
		self.elut_cpms = elut_cpms        
		self.elut_starts = [0.0] + elut_ends[:-1]
		# Basic analysis of CATE data GRABBED FROM EXCEL
		self.elut_cpms_gfact = elut_cpms_gfact
		self.elut_cpms_gRFW = elut_cpms_gRFW
		self.elut_cpms_log = elut_cpms_log
		self.r2s = r2s
		# More advanced CATE parameters (are extracted from base data)
		self.influx = influx
		self.netflux = netflux
		self.ratio = ratio
		self.poolsize = poolsize
		self.tracer_retained = tracer_retained
		self.p12_r2_max = p12_r2_max
		self.p12_r2_max_row = p12_r2_max_row
		self.phase3, self.phase2, self.phase1 = phase3, phase2, phase1

class Phase(object):
    '''
    Class that stores all data of a particular phase
    '''
    def __init__(
        self, indexs, r2, slope, intercept, k, t05, r0, efflux):
        self.indexs = indexs # paired tuple (x, y)
        self.r2, self.slope, self.intercept = r2, slope, intercept
        self.k, self.t05, self.r0, self.efflux = k, t05, r0, efflux

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

	# List where all run info is stored with TestAnalysis objects as ind. entries
	all_test_analyses = []

	for col_index in range(2, input_sheet.row_len(0)):        
		# Grab individual CATE values of interest
		run_name = input_sheet.cell(0, col_index).value
		SA = input_sheet.cell(1, col_index).value
		rt_cnts = input_sheet.cell(2, col_index).value
		sht_cnts = input_sheet.cell(3, col_index).value
		rt_wght = input_sheet.cell(4, col_index).value
		gfact = input_sheet.cell(5, col_index).value
		load_time = input_sheet.cell(6, col_index).value
		elut_period = input_sheet.cell(7, col_index).value
		influx = input_sheet.cell(8, col_index).value
		netflux = input_sheet.cell(9, col_index).value
		ratio = input_sheet.cell(10, col_index).value
		poolsize = input_sheet.cell(11, col_index).value
		tracer_retained = input_sheet.cell(12, col_index).value
		
		p3_start = input_sheet.cell(13, col_index).value
		p3_end = input_sheet.cell(14, col_index).value
		p3_efflux = input_sheet.cell(15, col_index).value
		p3_t05 = input_sheet.cell(16, col_index).value
		p3_r2 = input_sheet.cell(17, col_index).value
		p3_slope = input_sheet.cell(18, col_index).value
		p3_intercept = input_sheet.cell(19, col_index).value
		p3_k = input_sheet.cell(20, col_index).value
		p3_r0 = input_sheet.cell(21, col_index).value
		phase3 = Phase(
			(p3_start, p3_end), p3_r2, p3_slope, p3_intercept, p3_k, \
			p3_t05, p3_r0, p3_efflux)

		p2_start = input_sheet.cell(22, col_index).value
		p2_end = input_sheet.cell(23, col_index).value
		p2_efflux = input_sheet.cell(24, col_index).value
		p2_t05 = input_sheet.cell(25, col_index).value
		p2_r2 = input_sheet.cell(26, col_index).value
		p2_slope = input_sheet.cell(27, col_index).value
		p2_intercept = input_sheet.cell(28, col_index).value
		p2_k = input_sheet.cell(29, col_index).value
		p2_r0 = input_sheet.cell(30, col_index).value
		phase2 = Phase(
			(p2_start, p2_end), p2_r2, p2_slope, p2_intercept, p2_k, \
			p2_t05, p2_r0, p2_efflux)

		p1_start = input_sheet.cell(31, col_index).value
		p1_end = input_sheet.cell(32, col_index).value
		p1_efflux = input_sheet.cell(33, col_index).value
		p1_t05 = input_sheet.cell(34, col_index).value
		p1_r2 = input_sheet.cell(35, col_index).value
		p1_slope = input_sheet.cell(36, col_index).value
		p1_intercept = input_sheet.cell(37, col_index).value
		p1_k = input_sheet.cell(38, col_index).value
		p1_r0 = input_sheet.cell(39, col_index).value
		phase1 = Phase(
			(p1_start, p1_end), p1_r2, p1_slope, p1_intercept, p1_k, \
			p1_t05, p1_r0, p1_efflux)
		
		# Grabing elution cpms, correcting for header offset (15)
		start_cpms = 41
		end_cpms = start_cpms + len(elut_ends)
		raw_cpms = input_sheet.col(col_index)[start_cpms : end_cpms]
		elut_cpms = []
		for item in raw_cpms:
			if item.value != '':
				elut_cpms.append(float(item.value))
			else:
				elut_cpms.append(0.0)
		
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
		raw_log = input_sheet.col(col_index)[start_log:end_log]
		
		elut_cpms_log = []
		for index, item in enumerate(raw_log):
			if item.value:
				elut_cpms_log.append(item.value)

		start_r2s = end_log + 1
		end_r2s = start_r2s + len(elut_ends)
		raw_r2s = input_sheet.col(col_index)[start_r2s : end_r2s]
		r2s = []
		for item in raw_r2s: # item is a cell.Value
			try:
				r2s.append(float(item.value))
			except ValueError: # last row in r2 has no r2
				pass
		
		end_obj = int(input_sheet.col(col_index)[end_r2s].value)

		# Find index starting p2 (end index of p1)
		start_p12_r2_sum = end_r2s + 1 + len(elut_ends) + 1 + len(elut_ends) + 2
		end_p12_r2_sum = start_p12_r2_sum + len(elut_ends)
		p12_r2_max = input_sheet.col(col_index)[end_p12_r2_sum].value
		p12_r2_max_row = input_sheet.col(col_index)[end_p12_r2_sum + 1].value
		all_test_analyses.append(
			TestAnalysis(
				run_name, SA, rt_cnts, sht_cnts, rt_wght, gfact, load_time,
				elut_period, elut_ends, elut_cpms, elut_cpms_gfact,
				elut_cpms_gRFW, elut_cpms_log, r2s, influx, netflux, ratio,
				poolsize, tracer_retained, p12_r2_max, p12_r2_max_row,
				phase3,	phase2, phase1))
		
	return TestExperiment(directory, all_test_analyses)

directory = r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1"

@parameterized([
	#("Test_SingleRun1.xlsx"),
	#("Test_SingleRun2.xlsx"),
	#("Test_SingleRun3.xlsx"),
	#("Test_SingleRun4.xlsx"),
	#("Test_SingleRun5.xlsx"),
	#("Test_SingleRun6.xlsx"),
	#("Test_SingleRun7.xlsx"),
	#("Test_SingleRun8.xlsx"),
	#("Test_SingleRun9.xlsx"),
	#("Test_SingleRun10.xlsx"),
	#("Test_SingleRun11.xlsx"),
	#("Test_SingleRun12.xlsx"),
	("Test_MultiRun1.xlsx"),
	])
def test_basic(file_name):
	directory = r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1"
	question_exp = Excel.grab_data(directory, file_name)
	answer_exp = grab_answers(directory, file_name,\
		question_exp.analyses[0].run.elut_ends)
	for index, question in enumerate(question_exp.analyses):
		question.kind = 'obj'
		question.obj_num_pts = 8
		question.analyze()
		answer = answer_exp.analyses[index]

		assert_equals(question.run.SA, answer.SA)
		assert_equals(question.run.name, answer.name)
		assert_equals(question.run.rt_cnts, answer.rt_cnts)
		assert_equals(question.run.sht_cnts, answer.sht_cnts)
		assert_equals(question.run.rt_wght, answer.rt_wght)
		assert_equals(question.run.gfact, answer.gfact)
		assert_equals(question.run.load_time, answer.load_time)
		assert_equals(question.run.elut_ends, answer.elut_ends)
		assert_equals(question.run.elut_cpms, answer.elut_cpms)
		assert_equals(question.run.elut_starts, answer.elut_starts)
		assert_equals(question.run.elut_cpms_gfact, answer.elut_cpms_gfact)
		assert_equals(question.run.elut_cpms_gRFW, answer.elut_cpms_gRFW)
		assert_equals(question.run.elut_cpms_log, answer.elut_cpms_log)

@parameterized([
	# ("Test_SingleRun1.xlsx"),
	# ("Test_SingleRun2.xlsx"),
	# ("Test_SingleRun3.xlsx"),
	# ("Test_SingleRun4.xlsx"),
	# ("Test_SingleRun5.xlsx"),
	# ("Test_SingleRun6.xlsx"),
	# ("Test_SingleRun7.xlsx"),
	# ("Test_SingleRun8.xlsx"),
	# ("Test_SingleRun9.xlsx"),
	# ("Test_SingleRun10.xlsx"),
	# ("Test_SingleRun11.xlsx"),
	# ("Test_SingleRun12.xlsx"),
	("Test_MultiRun1.xlsx"),
	])
def test_phases(file_name):
	directory = r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1"
	question_exp = Excel.grab_data(directory, file_name)
	answer_exp = grab_answers(directory, file_name, question_exp.analyses[0].run.elut_ends)
	for index, question in enumerate(question_exp.analyses):
		question.kind = 'obj'
		question.obj_num_pts = 8
		question.analyze()
		answer = answer_exp.analyses[index]

		assert_equals(question.phase3.indexs[0], answer.phase3.indexs[0])
		assert_equals(question.phase3.indexs[1], answer.phase3.indexs[1])
		assert_equals(
			"{0:.7f}".format(question.phase3.slope),
			"{0:.7f}".format(answer.phase3.slope))	
		assert_equals(
			"{0:.7f}".format(question.phase3.intercept),
			"{0:.7f}".format(answer.phase3.intercept))			
		assert_equals(
			"{0:.7f}".format(question.phase3.k),
			"{0:.7f}".format(answer.phase3.k))
		assert_equals(
			"{0:.7f}".format(question.phase3.r0),
			"{0:.7f}".format(answer.phase3.r0))				
		assert_equals(
			"{0:.7f}".format(question.phase3.efflux),
			"{0:.7f}".format(answer.phase3.efflux))
		assert_equals(
			"{0:.7f}".format(question.phase3.t05),
			"{0:.7f}".format(answer.phase3.t05))
		assert_equals(
			"{0:.7f}".format(question.phase3.r2),
			"{0:.7f}".format(answer.phase3.r2))		
		
		assert_equals(
			"{0:.7f}".format(question.influx),
			"{0:.7f}".format(answer.influx))
		assert_equals(
			"{0:.7f}".format(question.netflux),
			"{0:.7f}".format(answer.netflux))
		assert_equals(
			"{0:.7f}".format(question.ratio),
			"{0:.7f}".format(answer.ratio))
		assert_equals(
			"{0:.7f}".format(question.poolsize),
			"{0:.7f}".format(answer.poolsize))
		assert_equals(
			"{0:.7f}".format(question.tracer_retained),
			"{0:.7f}".format(answer.tracer_retained))

		assert_equals(question.phase2.indexs[0], answer.phase2.indexs[0])
		assert_equals(question.phase2.indexs[1], answer.phase2.indexs[1])
		assert_equals(
			"{0:.7f}".format(question.phase2.slope),
			"{0:.7f}".format(answer.phase2.slope))	
		assert_equals(
			"{0:.7f}".format(question.phase2.intercept),
			"{0:.7f}".format(answer.phase2.intercept))			
		assert_equals(
			"{0:.7f}".format(question.phase2.k),
			"{0:.7f}".format(answer.phase2.k))
		assert_equals(
			"{0:.7f}".format(question.phase2.r0),
			"{0:.7f}".format(answer.phase2.r0))				
		assert_equals(
			"{0:.7f}".format(question.phase2.efflux),
			"{0:.7f}".format(answer.phase2.efflux))
		assert_equals(
			"{0:.7f}".format(question.phase2.t05),
			"{0:.7f}".format(answer.phase2.t05))
		assert_equals(
			"{0:.7f}".format(question.phase2.r2),
			"{0:.7f}".format(answer.phase2.r2))

		assert_equals(question.phase1.indexs[0], answer.phase1.indexs[0])
		assert_equals(question.phase1.indexs[1], answer.phase1.indexs[1])
		assert_equals(
			"{0:.7f}".format(question.phase1.slope),
			"{0:.7f}".format(answer.phase1.slope))	
		assert_equals(
			"{0:.7f}".format(question.phase1.intercept),
			"{0:.7f}".format(answer.phase1.intercept))			
		assert_equals(
			"{0:.7f}".format(question.phase1.k),
			"{0:.7f}".format(answer.phase1.k))
		assert_equals(
			"{0:.7f}".format(question.phase1.r0),
			"{0:.7f}".format(answer.phase1.r0))				
		assert_equals(
			"{0:.7f}".format(question.phase1.efflux),
			"{0:.7f}".format(answer.phase1.efflux))
		assert_equals(
			"{0:.7f}".format(question.phase1.t05),
			"{0:.7f}".format(answer.phase1.t05))
		assert_equals(
			"{0:.7f}".format(question.phase1.r2),
			"{0:.7f}".format(answer.phase1.r2))		

		for counter in range(0, len(question.r2s)):
			assert_equals(
				"{0:.10f}".format(question.r2s[counter]),
				"{0:.10f}".format(answer.r2s[counter]))

if __name__ == '__main__':
	
	import Excel
	temp_data = Excel.grab_data(r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1", "Test_SingleRun6.xlsx")
	temp_question = temp_data.analyses[0]
	temp_question.kind = 'obj'
	temp_question.obj_num_pts = 8
	temp_question.analyze()

	temp_exp = grab_answers(
		r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1", "Test_SingleRun6.xlsx", temp_question.run.elut_ends)
	temp_answer = temp_exp.analyses[0]

	print "ANSWERS"	
	#print temp_answer.phase3.x
	#print temp_answer.phase1.slope
	print "QUESTIONS"
	print temp_question.phase3.x
	#print temp_question.phase1.x
	#print temp_question.phase1.slope