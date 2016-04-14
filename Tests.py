import Objects

from nose.tools import assert_equals

class TestExperiment(object):
	 def __init__(self, directory, testruns):
        self.directory = directory
        self.testruns = testruns # List of Analysis objects

class TestRun(object):
    '''
    Class that stores ALL data of a single CATE run.
    This data includes values derived from objective or 
    subjetive analyses (calculated within the class)
    '''        
    def __init__(
    	   self, run_name, SA, rt_cnts, sht_cnts, rt_wght, gfact, load_time,
    	   elut_ends, elut_cpms, elut_cpms_gfact, elut_cpms_gfact,
    	   elut_cpms_gRFW, elut_cpms_log):       
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
    Data are what analysis SHOULD be arriving at.
    '''
	# Accessing the file from which data is to be grabbed
    input_file = '/'.join((directory, filename))
    input_book = open_workbook(input_file)
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


def test_all():
	directory = r"C:\Users\Ruben\Projects\CATEAnalysis\Tests\1"
	test_file = "Test - Single Run.xlsx"

	temp_data = Excel.grab_data(directory, test_file)
	single_run = temp_data[0]
	Objects.find_obj_reg(single_run.elut_ends, single_run.elut_cpms_log, 3)

	grab_answers(directory, test_file, single_run.elut_ends)