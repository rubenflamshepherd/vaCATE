import numpy
import math
import Operations

class Experiment(object):
    def __init__(self, directory, analyses):
        self.directory = directory
        self.analyses = analyses # List of Analysis objects

class Analysis(object):
    def __init__(self, kind, obj_num_pts, run, xs_p1=('', ''), 
            xs_p2=('', ''), xs_p3=('', '')):
        self.kind = kind # None, 'obj', or 'subj'
        self.obj_num_pts = obj_num_pts # None if not obj regression
        self.run = run
        self.xs_p3 = xs_p3
        self.xs_p2 = xs_p2
        self.xs_p1 = xs_p1

        # Default values are None unless assigned
        self.phase3, self.phase2, self.phase1 = None, None, None
        self.r2s = None

        self.x1_p3, self.y1_p3, self.x2_p3, self.y2_p3 = None, None, None, None
        self.x1_p2, self.y1_p2, self.x2_p2, self.y2_p2 = None, None, None, None
        self.x1_p1, self.y1_p1, self.x2_p1, self.y2_p1 = None, None, None, None

        self.r2_p3, self.m_p3, self.b_p3 = None, None, None # y=mx+b
        self.r2_p2, self.m_p2, self.b_p2 = None, None, None
        self.r2_p1, self.m_p1, self.b_p1 = None, None, None

        self.obj_x_start, self.obj_y_start = None, None
        # Attributes for testing
        self.r2s = None # Lists from obj analysis, y=mx+b
        self.p12_r2_max = None

        self.x_p3, self.x_p2, self.x_p1, self.x_p12 = None, None, None, None
        self.y_p3, self.y_p2, self.y_p1, self.y_p12 = None, None, None, None
        self.x_p12_curvestrip_p3, self.y_p12_curvestrip_p3 = None, None
        self.x_p2_curvestrip_p3, self.y_p2_curvestrip_p3 = None, None
        self.x_p1_curvestrip_p3, self.y_p1_curvestrip_p3 = None, None
        self.x_p1_curvestrip_p23, self.y_p1_curvestrip_p23 = None, None

        self.elut_period, self.tracer_retained, self.poolsize = None, None, None
        self.influx, self.netflux, self.ratio = None, None, None

    def analyze(self):
        '''
        Take Analysis object and define parameters based on which phase limits
        have been provided.
        Return Analysis object
        '''
        if self.kind == 'obj':
            self.obj_x_start = self.run.x[-self.obj_num_pts:]
            self.obj_y_start = self.run.y[-self.obj_num_pts:]
            self.xs_p3, self.r2s = Operations.get_obj_phase3(
                obj_num_pts=self.obj_num_pts,
                elut_ends_parsed=self.run.elut_ends_parsed,
                elut_cpms_log=self.run.elut_cpms_log)
            self.xs_p2, self.xs_p1, self.p12_r2_max = Operations.get_obj_phase12(
                xs_p3=self.xs_p3, elut_ends_parsed=self.run.elut_ends_parsed,
                elut_cpms_log=self.run.elut_cpms_log,
                elut_ends=self.run.elut_ends)
        if self.xs_p3 != ('', ''):
            self.phase3 = Operations.extract_phase(
                xs=self.xs_p3, 
                x_series=self.run.elut_ends_parsed, 
                y_series=self.run.elut_cpms_log,
                elut_ends=self.run.elut_ends,
                SA=self.run.SA, load_time=self.run.load_time)
            Operations.advanced_run_calcs(analysis=self)
            if self.xs_p2 != ('', '') and self.phase3.xs!= ('', ''):
                # Set series' to be curvestripped
                end_p12_index = Operations.x_to_index(
                    x_value=self.xs_p2[1], index_type='end',
                    x_series=self.run.elut_ends_parsed,
                    larger_x=self.run.elut_ends)
                self.x_p12 = self.run.x[: end_p12_index+1]
                self.y_p12 = self.run.y[: end_p12_index+1]
                # Curve strip phase 1 + 2 data of phase 3
                # From here on data series potentially have 'holes' from ommitting
                # negative log operations during curvestripping
                self.x_p12_curvestrip_p3, self.y_p12_curvestrip_p3 = \
                    Operations.curvestrip(
                        x_series=self.x_p12, y_series=self.y_p12, 
                        slope=self.phase3.slope,
                        intercept=self.phase3.intercept)
                self.phase2 = Operations.extract_phase(
                    xs=self.xs_p2, 
                    x_series=self.x_p12_curvestrip_p3,
                    y_series=self.y_p12_curvestrip_p3,
                    elut_ends=self.run.elut_ends,
                    SA=self.run.SA, load_time=self.run.load_time)
                if self.xs_p1 != ('', '') and self.phase2.xs!= ('', ''):
                    start_p1_index = Operations.x_to_index(
                        x_value=self.xs_p1[0], index_type='start',
                        x_series=self.run.elut_ends_parsed,
                        larger_x=self.run.elut_ends)
                    end_p1_index = Operations.x_to_index(
                        x_value=self.xs_p1[1], index_type='end',
                        x_series=self.run.elut_ends,
                        larger_x=self.run.elut_ends)
                    self.x_p1 = self.run.x[start_p1_index : end_p1_index+1]
                    self.y_p1 = self.run.y[start_p1_index : end_p1_index+1]
                    # Set series' to be further curvestripped (already partially done)
                    self.x_p1_curvestrip_p3 =\
                        self.x_p12_curvestrip_p3[start_p1_index : end_p1_index+1]
                    self.y_p1_curvestrip_p3 =\
                        self.y_p12_curvestrip_p3[start_p1_index : end_p1_index+1]
                    # Curve strip phase 1 data of phase 3 (already stripped phase 3)
                    self.x_p1_curvestrip_p23, self.y_p1_curvestrip_p23 = \
                        Operations.curvestrip(
                            x_series=self.x_p1_curvestrip_p3,
                            y_series=self.y_p1_curvestrip_p3, 
                            slope=self.phase2.slope,
                            intercept=self.phase2.intercept)
                    print self.x_p1_curvestrip_p23
                    self.phase1 = Operations.extract_phase(
                        xs=self.xs_p1, 
                        x_series=self.x_p1_curvestrip_p23,
                        y_series=self.y_p1_curvestrip_p23,
                        elut_ends=self.run.elut_ends,
                        SA=self.run.SA, load_time=self.run.load_time)

class Run(object):
    '''
    Class that stores ALL data of a single CATE run.
    This data includes values derived from objective or 
    subjetive analyses (calculated within the class)
    '''        
    def __init__(
    	   self, name, SA, rt_cnts, sht_cnts, rt_wght, gfact,
    	   load_time, elut_ends, elut_cpms):       
        self.name = name # Text identifier extracted from col header in excel
        self.SA = SA
        self.rt_cnts = rt_cnts
        self.sht_cnts = sht_cnts
        self.rt_wght = rt_wght
        self.gfact = gfact
        self.load_time = load_time
        self.elut_ends = elut_ends
        self.elut_cpms = elut_cpms        
        self.elut_starts = [0.0] + elut_ends[:-1]
        self.elut_cpms_gfact, self.elut_cpms_gRFW, \
            self.elut_cpms_log, self.elut_ends_parsed = \
            Operations.basic_run_calcs(
                rt_wght, gfact, self.elut_starts, elut_ends, elut_cpms)       
		# x and y data for graphing ('numpy-fied')
        self.x = numpy.array(self.elut_ends_parsed)
        self.y = numpy.array(self.elut_cpms_log)

class Phase(object):
    '''
    Class that stores all data of a particular phase
    '''
    def __init__(
        self, xs, xy1, xy2, r2, slope, intercept, x_series, y_series,
        k, t05, r0, efflux):
        self.xs = xs # paired tuple (x, y)
        self.xy1, self.xy2 = xy1, xy2 # Each is a paired tuple

        self.r2, self.slope, self.intercept = r2, slope, intercept
        self.x_series, self.y_series = x_series, y_series
        self.k, self.t05, self.r0, self.efflux = k, t05, r0, efflux

if __name__ == "__main__":
    import Excel
    import os
    directory = os.path.dirname(os.path.abspath(__file__))
    temp_data = Excel.grab_data(directory, "/Tests/1/Test_SingleRun1.xlsx")
    
    temp_analysis = temp_data.analyses[0]
    temp_analysis.kind = 'obj'
    temp_analysis.obj_num_pts = 4
    temp_analysis.analyze()