import xlrd
from nose.tools import assert_equals
from nose_parameterized import parameterized
import os

import Excel

class TestExperiment(object):
    def __init__(self, directory, analyses):
        self.directory = directory
        self.analyses = analyses  # List of Analysis objects


class TestAnalysis(object):
    """	Class that stores ALL data of a single CATE run.
	This data includes values derived from objective or
	subjective analyses (calculated within the class)
	"""
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
    """ Class that stores all data of a particular phase
	"""

    def __init__(
            self, xs, r2, slope, intercept, k, t05, r0, efflux):
        self.xs = xs  # paired tuple (x, y)
        self.r2, self.slope, self.intercept = r2, slope, intercept
        self.k, self.t05, self.r0, self.efflux = k, t05, r0, efflux


def grab_answers(directory, filename, elut_ends):
    """	Extracts data from an excel file in directory/filename (INPUT)
	elut_ends required because can not be cleanly extracted from test sheet
	Data are what analysis SHOULD be arriving at.
	"""
    input_file = '/'.join((directory, filename))
    input_book = xlrd.open_workbook(input_file)
    input_sheet = input_book.sheet_by_index(1)

    # List where all run info is stored with TestAnalysis objects as ind. entries
    all_test_analyses = []

    for col_index in range(2, input_sheet.row_len(0)):
        # Grab individual CATE values of interest
        run_name = str(input_sheet.cell(0, col_index).value)
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

        # Grabbing elution cpms, correcting for header offset (15)
        start_cpms = 41
        end_cpms = start_cpms + len(elut_ends)
        raw_cpms = input_sheet.col(col_index)[start_cpms: end_cpms]
        elut_cpms = []
        for item in raw_cpms:
            if item.value != '':
                elut_cpms.append(float(item.value))
            else:
                elut_cpms.append(0.0)

        start_gfact = end_cpms + 1  # row underneath cpms_gfact label
        end_gfact = start_gfact + len(elut_ends)
        raw_gfact = input_sheet.col(col_index)[start_gfact: end_gfact]

        elut_cpms_gfact = [float(x.value) for x in raw_gfact]
        elut_cpms_gfact = []
        for item in raw_gfact:
            if item.value:
                elut_cpms_gfact.append(float(item.value))

        start_gRFW = end_gfact + 1
        end_gRFW = start_gRFW + len(elut_ends)
        raw_gRFW = input_sheet.col(col_index)[start_gRFW: end_gRFW]
        elut_cpms_gRFW = []
        for item in raw_gRFW:
            if item.value:
                elut_cpms_gRFW.append(float(item.value))

        start_log = end_gRFW + 1
        end_log = start_log + len(elut_ends)
        raw_log = input_sheet.col(col_index)[start_log:end_log]

        elut_cpms_log = []
        for index, item in enumerate(raw_log):
            if item.value:
                elut_cpms_log.append(item.value)

        start_r2s = end_log + 1
        end_r2s = start_r2s + len(elut_ends)
        raw_r2s = input_sheet.col(col_index)[start_r2s: end_r2s]
        r2s = []
        for item in raw_r2s:  # item is a cell.Value
            try:
                r2s.append(float(item.value))
            except ValueError:  # last row in r2 has no r2
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
                phase3, phase2, phase1))

    return TestExperiment(directory, all_test_analyses)


@parameterized([
    ("Tests/1/Test_SingleRun1.xlsx", 8),
    ("Tests/1/Test_SingleRun2.xlsx", 8),
    ("Tests/1/Test_SingleRun3.xlsx", 8),
    ("Tests/1/Test_SingleRun4.xlsx", 8),
    ("Tests/1/Test_SingleRun5.xlsx", 8),
    ("Tests/1/Test_SingleRun6.xlsx", 8),
    ("Tests/1/Test_SingleRun7.xlsx", 8),
    ("Tests/1/Test_SingleRun8.xlsx", 8),
    ("Tests/1/Test_SingleRun9.xlsx", 8),
    ("Tests/1/Test_SingleRun10.xlsx", 8),
    ("Tests/1/Test_SingleRun11.xlsx", 8),
    ("Tests/1/Test_SingleRun12.xlsx", 8),
    ("Tests/1/Test_MultiRun1.xlsx", 8),
    ("Tests/1/Test_SubjSingleRun1.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/1/Test_SubjSingleRun2.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/1/Test_SubjSingleRun3.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/1/Test_SubjSingleRun4.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/1/Test_SubjSingleRun5.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/1/Test_SubjSingleRun6.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/1/Test_SubjMultiRun1.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_CurveStripPh1.xlsx", 8),
    ("Tests/Edge Cases/Test_MissLastPtPh3.xlsx", 8),
    ("Tests/Edge Cases/Test_MissLastPtPh2.xlsx", 8),
    ("Tests/Edge Cases/Test_MissLastPtPh1.xlsx", 8),
    ("Tests/Edge Cases/Test_Miss1stPtPh3.xlsx", 8),
    ("Tests/Edge Cases/Test_Miss1stPtPh2.xlsx", 8),
    ("Tests/Edge Cases/Test_Miss1stPtPh1.xlsx", 8),
    ("Tests/Edge Cases/Test_MissMajPh3.xlsx", 8),
    ("Tests/Edge Cases/Test_MissMidPtPh3.xlsx", 8),
    ("Tests/Edge Cases/Test_MissMidPtPh2.xlsx", 8),
    ("Tests/Edge Cases/Test_MissMidPtPh1.xlsx", 8),
    ("Tests/Edge Cases/Test_Ph1Is1Pt.xlsx", 8),
    ("Tests/Edge Cases/Test_RsqNoDec.xlsx", 8),
    ("Tests/Edge Cases/Test_EarlyEndPh3.xlsx", 8),
    ("Tests/Edge Cases/Test_SubjMiss1stPtPh1.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_SubjMiss1stPtPh2.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_SubjMiss1stPtPh3.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_SubjMissLastPtPh3.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_SubjMissLastPtPh2.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_SubjMissLastPtPh1.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_SubjMissMidPtPh3.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_SubjMissMidPtPh23.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/Edge Cases/Test_SubjMissMidPtPh123.xlsx", [(1, 3), (4, 10), (11.5, 45)]),
    ("Tests/2/Test_SingleRun1.xlsx", 8),
    ("Tests/2/Test_SingleRun2.xlsx", 8),
    ("Tests/2/Test_SingleRun3.xlsx", 8),
    ("Tests/2/Test_SingleRun4.xlsx", 8),
    ("Tests/2/Test_SingleRun5.xlsx", 8),
    ("Tests/2/Test_SingleRun6.xlsx", 8),
    ("Tests/2/Test_SingleRun7.xlsx", 8),
    ("Tests/2/Test_SingleRun8.xlsx", 8),
    ("Tests/2/Test_SingleRun9.xlsx", 8),
    ("Tests/2/Test_SingleRun10.xlsx", 8),
    ("Tests/2/Test_SingleRun11.xlsx", 8),
    ("Tests/2/Test_SingleRun12.xlsx", 8),
    ("Tests/2/Test_SingleRun13.xlsx", 8),
    ("Tests/2/Test_SingleRun14.xlsx", 8),
    ("Tests/2/Test_SingleRun15.xlsx", 8),
    ("Tests/2/Test_SingleRun16.xlsx", 8),
    ("Tests/2/Test_MultiRun1.xlsx", 8),
    ("Tests/2/Test_SubjSingleRun1.xlsx", [(1, 3), (4, 10), (11.5, 40)]),
    ("Tests/2/Test_SubjSingleRun3.xlsx", [(1, 3), (4, 10), (11.5, 40)]),
    ("Tests/2/Test_SubjMultiRun1.xlsx", [(1, 3), (4, 10), (11.5, 40)]),
    ("Tests/3/Test_SingleRun1.xlsx", 8),
    ("Tests/3/Test_SingleRun2.xlsx", 8),
    ("Tests/3/Test_SingleRun3.xlsx", 8),
    ("Tests/3/Test_SingleRun4.xlsx", 8),
    ("Tests/3/Test_SingleRun5.xlsx", 8),
    ("Tests/3/Test_SingleRun6.xlsx", 8),
    ("Tests/3/Test_SingleRun7.xlsx", 8),
    ("Tests/3/Test_SingleRun8.xlsx", 8),
    ("Tests/3/Test_SingleRun9.xlsx", 8),
    ("Tests/3/Test_SingleRun10.xlsx", 8),
    ("Tests/3/Test_SingleRun11.xlsx", 8),
    ("Tests/3/Test_SingleRun12.xlsx", 8),
    ("Tests/3/Test_SingleRun13.xlsx", 8),
    ("Tests/3/Test_SingleRun14.xlsx", 8),
    ("Tests/3/Test_SingleRun15.xlsx", 8),
    ("Tests/3/Test_SingleRun16.xlsx", 8),
    ("Tests/3/Test_SingleRun17.xlsx", 8),
    ("Tests/3/Test_SingleRun18.xlsx", 8),
    ("Tests/3/Test_SingleRun19.xlsx", 8),
    ("Tests/3/Test_SingleRun20.xlsx", 8),
    ("Tests/3/Test_MultiRun1.xlsx", 8),
    ("Tests/3/Test_SubjSingleRun1.xlsx", [(1, 3), (4, 10), (11.5, 40)]),
    ("Tests/3/Test_SubjMultiRun1.xlsx", [(1, 3), (4, 10), (11.5, 40)]),
    ("Tests/4/Test_SingleRun1.xlsx", 8),
    ("Tests/4/Test_SingleRun2.xlsx", 8),
    ("Tests/4/Test_SingleRun3.xlsx", 8),
    ("Tests/4/Test_SingleRun4.xlsx", 8),
    ("Tests/4/Test_SingleRun5.xlsx", 8),
    ("Tests/4/Test_SingleRun6.xlsx", 8),
    ("Tests/4/Test_SingleRun7.xlsx", 8),
    ("Tests/4/Test_SingleRun8.xlsx", 8),
    ("Tests/4/Test_SingleRun9.xlsx", 8),
    ("Tests/4/Test_SingleRun10.xlsx", 8),
    ("Tests/4/Test_SingleRun11.xlsx", 8),
    ("Tests/4/Test_SingleRun12.xlsx", 8),
    ("Tests/4/Test_SingleRun13.xlsx", 8),
    ("Tests/4/Test_SingleRun14.xlsx", 8),
    ("Tests/4/Test_SingleRun15.xlsx", 8),
    ("Tests/4/Test_SingleRun16.xlsx", 8),
    ("Tests/4/Test_SingleRun17.xlsx", 8),
    ("Tests/4/Test_SingleRun18.xlsx", 8),
    ("Tests/4/Test_SingleRun19.xlsx", 8),
    ("Tests/4/Test_SingleRun20.xlsx", 8),
    ("Tests/4/Test_SingleRun21.xlsx", 8),
    ("Tests/4/Test_SingleRun22.xlsx", 8),
    ("Tests/4/Test_SingleRun23.xlsx", 8),
    ("Tests/4/Test_SingleRun24.xlsx", 8),
    ("Tests/4/Test_MultiRun1.xlsx", 8),
    ("Tests/4/Test_SubjMultiRun1.xlsx", [(1.5, 4.5), (6, 9), (10.5, 39)]),
])
def test_analysis(file_name, analysis_data):
    directory = os.path.dirname(os.path.abspath(__file__))
    question_path = os.path.join(directory, file_name)
    print question_path
    question_exp = Excel.grab_data(question_path)
    answer_exp = grab_answers(directory, file_name, \
                              question_exp.analyses[0].run.elut_ends)
    for index, question in enumerate(question_exp.analyses):
        if 'Subj' in file_name:
            question.kind = 'subj'
            question.xs_p3 = analysis_data[2]
            question.xs_p2 = analysis_data[1]
            question.xs_p1 = analysis_data[0]
            question.analyze()
            answer = answer_exp.analyses[index]
        else:
            question.kind = 'obj'
            question.obj_num_pts = analysis_data
            question.analyze()
            answer = answer_exp.analyses[index]
            for counter in range(0, len(question.r2s)):
                assert_equals(
                    "{0:.10f}".format(question.r2s[counter]),
                    "{0:.10f}".format(answer.r2s[counter]))

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
        for index, item in enumerate(question.run.elut_cpms_gfact):
            assert_equals(
                "{0:.10f}".format(question.run.elut_cpms_gfact[index]),
                "{0:.10f}".format(answer.elut_cpms_gfact[index]))
        assert_equals(question.run.elut_cpms_gRFW, answer.elut_cpms_gRFW)
        assert_equals(question.run.elut_cpms_log, answer.elut_cpms_log)

        assert_equals(question.phase3.xs[0], answer.phase3.xs[0])
        assert_equals(question.phase3.xs[1], answer.phase3.xs[1])
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
            "{0:.7f}".format(question.netflux),
            "{0:.7f}".format(answer.netflux))
        assert_equals(
            "{0:.7f}".format(question.influx),
            "{0:.7f}".format(answer.influx))
        assert_equals(
            "{0:.7f}".format(question.ratio),
            "{0:.7f}".format(answer.ratio))
        assert_equals(
            "{0:.7f}".format(question.poolsize),
            "{0:.7f}".format(answer.poolsize))
        assert_equals(
            "{0:.7f}".format(question.tracer_retained),
            "{0:.7f}".format(answer.tracer_retained))

        assert_equals(question.phase2.xs[0], answer.phase2.xs[0])
        assert_equals(question.phase2.xs[1], answer.phase2.xs[1])
        if question.phase2.xs != ('', ''):
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
                "{0:.3f}".format(question.phase2.r0),
                "{0:.3f}".format(answer.phase2.r0))
            assert_equals(
                "{0:.4f}".format(question.phase2.efflux),
                "{0:.4f}".format(answer.phase2.efflux))
            assert_equals(
                "{0:.7f}".format(question.phase2.t05),
                "{0:.7f}".format(answer.phase2.t05))
            assert_equals(
                "{0:.7f}".format(question.phase2.r2),
                "{0:.7f}".format(answer.phase2.r2))
        else:
            assert_equals(question.phase2.slope, answer.phase2.slope)
            assert_equals(question.phase2.intercept, answer.phase2.intercept)
            assert_equals(question.phase2.k, answer.phase2.k)
            assert_equals(question.phase2.r0, answer.phase2.r0)
            assert_equals(question.phase2.efflux, answer.phase2.efflux)
            assert_equals(question.phase2.t05, answer.phase2.t05)
            assert_equals(question.phase2.r2, answer.phase2.r2)

        assert_equals(question.phase1.xs[0], answer.phase1.xs[0])
        assert_equals(question.phase1.xs[1], answer.phase1.xs[1])
        if question.phase1.xs != ('', ''):
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
                "{0:.5f}".format(question.phase1.r0),
                "{0:.5f}".format(answer.phase1.r0))
            assert_equals(
                "{0:.6f}".format(question.phase1.efflux),
                "{0:.6f}".format(answer.phase1.efflux))
            assert_equals(
                "{0:.7f}".format(question.phase1.t05),
                "{0:.7f}".format(answer.phase1.t05))
            assert_equals(
                "{0:.7f}".format(question.phase1.r2),
                "{0:.7f}".format(answer.phase1.r2))
        else:
            assert_equals(question.phase1.slope, answer.phase1.slope)
            assert_equals(question.phase1.intercept, answer.phase1.intercept)
            assert_equals(question.phase1.k, answer.phase1.k)
            assert_equals(question.phase1.r0, answer.phase1.r0)
            assert_equals(question.phase1.efflux, answer.phase1.efflux)
            assert_equals(question.phase1.t05, answer.phase1.t05)
            assert_equals(question.phase1.r2, answer.phase1.r2)


if __name__ == '__main__':
    import Excel

    directory = os.path.dirname(os.path.abspath(__file__))
    # temp_data = Excel.grab_data(directory, "/Tests/Edge Cases/Test_SubjMissMidPtPh123.xlsx")
    temp_data = Excel.grab_data(directory, "/Tests/4/Test_SingleRun7.xlsx")
    temp_question = temp_data.analyses[0]
    # temp_question.kind = 'subj'
    # temp_question.xs_p1 = (1,3)
    # temp_question.xs_p2 = (4,10)
    # temp_question.xs_p3 = (11.5,40)
    temp_question.kind = 'obj'
    temp_question.obj_num_pts = 8
    temp_question.analyze()

    # temp_exp = grab_answers(directory, "/Tests/Edge Cases/Test_SubjMissMidPtPh123.xlsx", temp_question.run.elut_ends)
    temp_exp = grab_answers(directory, "/Tests/4/Test_SingleRun7.xlsx", temp_question.run.elut_ends)
    temp_answer = temp_exp.analyses[0]

    print "ANSWERS"
    # print temp_answer.phase1.x
    # print temp_answer.phase1.slope
    print "QUESTIONS"
# print temp_question.phase1.x
# print temp_question.phase1.slope
