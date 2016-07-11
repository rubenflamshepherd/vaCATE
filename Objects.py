import numpy
import math
import Operations

class Experiment(object):
    def __init__(self, directory, analyses):
        self.directory = directory
        self.analyses = analyses # List of Analysis objects

class Analysis(object):
    def __init__(self, kind, obj_num_pts, run, indexs_p1=('', ''), 
            indexs_p2=('', ''), indexs_p3=('', '')):
        self.kind = kind # None, 'obj', or 'subj'
        self.obj_num_pts = obj_num_pts # None if not obj regression
        self.run = run
        self.indexs_p3 = indexs_p3
        self.indexs_p2 = indexs_p2
        self.indexs_p1 = indexs_p1

        # Default values are None unless assigned
        self.phase3, self.phase2, self.phase1 = None, None, None

        self.x1_p3, self.y1_p3, self.x2_p3, self.y2_p3 = None, None, None, None
        self.x1_p2, self.y1_p2, self.x2_p2, self.y2_p2 = None, None, None, None
        self.x1_p1, self.y1_p1, self.x2_p1, self.y2_p1 = None, None, None, None

        self.r2_p3, self.m_p3, self.b_p3 = None, None, None # y=mx+b
        self.r2_p2, self.m_p2, self.b_p2 = None, None, None
        self.r2_p1, self.m_p1, self.b_p1 = None, None, None

        self.obj_x_start, self.obj_y_start = None, None
        self.r2s, self.bs = None, None # Lists from obj analysis, y=mx+b

        self.x_p3, self.x_p2, self.x_p1, self.x_p12 = None, None, None, None
        self.y_p3, self.y_p2, self.y_p1, self.y_p12 = None, None, None, None
        self.x_p12_curvestrip_p3, self.y_p12_curvestrip_p3 = None, None
        self.x_p2_curvestrip_p3, self.y_p2_curvestrip_p3 = None, None
        self.x_p1_curvestrip_p3, self.y_p1_curvestrip_p3 = None, None
        self.x_p1_curvestrip_p23, self.y_p1_curvestrip_p23 = None, None

        self.k_p3, self.k_p2, self.k_p1 = None, None, None
        self.t05_p3, self.t05_p2, self.t05_p1 = None, None, None
        self.R0_p3, self.R0_p2, self.R0_p1 = None, None, None
        self.efflux_p3, self.efflux_p2, self.efflux_p1 = None, None, None

        self.elut_period, self.tracer_retained, self.poolsize = None, None, None
        self.influx, self.netflux, self.ratio = None, None, None

    def analyze(self):
        '''
        Take Analysis object and define parameters based on which phase limits
        have been provided.
        Return Analysis object
        '''
        if self.kind == 'obj':
            self.indexs_p3, self.indexs_p2, self.indexs_p1 = \
                Operations.set_obj_phases(
                    run=self.run, obj_num_pts=self.obj_num_pts)
            self.obj_x_start = self.run.x[-self.obj_num_pts:]
            self.obj_y_start = self.run.y[-self.obj_num_pts:]
    
        if self.indexs_p3 != ('', ''):
            self.phase3 = Operations.extract_phase(
                self.indexs_p3, self.run.x, self.run.y,\
                self.run.SA, self.run.load_time)
            Operations.advanced_run_calcs(self)
        if self.indexs_p2 != ('', ''):
            # Set series' to be curvestripped
            self.x_p12 = self.run.x[:self.indexs_p2[1]]
            self.y_p12 = self.run.y[:self.indexs_p2[1]]
            # Curve strip phase 1 + 2 data of phase 3
            self.x_p12_curvestrip_p3, self.y_p12_curvestrip_p3 = \
                Operations.curvestrip(
                    self.x_p12, self.y_p12, 
                    self.phase3.slope, self.phase3.intercept)
            self.phase2 = Operations.extract_phase(
                self.indexs_p2, 
                self.x_p12_curvestrip_p3, self.y_p12_curvestrip_p3,
                self.run.SA, self.run.load_time)
        if self.indexs_p1 != ('', ''):
            self.x_p1 = self.run.x[self.indexs_p1[0]:self.indexs_p1[1]]
            self.y_p1 = self.run.y[self.indexs_p1[0]:self.indexs_p1[1]]
            # Set series' to be further curvestripped (already partially done)
            self.x_p1_curvestrip_p3 =\
                self.x_p12_curvestrip_p3[self.indexs_p1[0]:self.indexs_p1[1]]
            self.y_p1_curvestrip_p3 =\
                self.y_p12_curvestrip_p3[self.indexs_p1[0]:self.indexs_p1[1]]
            # Curve strip phase 1 data of phase 3 (already stripped phase 3)
            self.x_p1_curvestrip_p23, self.y_p1_curvestrip_p23 = \
                Operations.curvestrip(
                    self.x_p1_curvestrip_p3, self.y_p1_curvestrip_p3, 
                    self.phase2.slope, self.phase2.intercept)
            self.phase1 = Operations.extract_phase(
                self.indexs_p1, 
                self.x_p1_curvestrip_p23, self.y_p1_curvestrip_p23,
                self.run.SA, self.run.load_time)
        
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
        self.elut_cpms_gfact, self.elut_cpms_gRFW, self.elut_cpms_log = \
            Operations.basic_run_calcs(
                rt_wght, gfact, self.elut_starts, elut_ends, elut_cpms)       
		# x and y data for graphing ('numpy-fied')
        self.x = numpy.array(self.elut_ends)
        self.y = numpy.array(self.elut_cpms_log)

class Phase(object):
    '''
    Class that stores all data of a particular phase
    '''
    def __init__(
        self, indexs, xy1, xy2, r2, slope, intercept, x, y,
        k, t05, r0, efflux):
        self.indexs = indexs # paired tuple (x, y)
        self.xy1, self.xy2 = xy1, xy2 # Each is a paired tuple
        self.r2, self.slope, self.intercept = r2, slope, intercept
        self.x, self.y = x, y
        self.k, self.t05, self.r0, self.efflux = k, t05, r0, efflux
        		
if __name__ == "__main__":
    import Excel
    temp_data = Excel.grab_data(r"C:\Users\Daniel\Projects\CATEAnalysis\Tests\1", "Test - Single Run.xlsx")
    
    temp_analysis = temp_data.analyses[0]
    temp_analysis.kind = 'obj'
    temp_analysis.obj_num_pts = 3
    temp_analysis.analyze()
