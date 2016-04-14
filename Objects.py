import numpy
import math
import Operations

class Experiment(object):
    def __init__(self, directory, analyses):
        self.directory = directory
        self.analyses = analyses # List of Analysis objects

class Analysis(object):
    def __init__(self, kind, run, p1_indexs=None, p2_indexs=None, p3_indexs=None):
        self.kind = kind
        self.run = run
        self.p1_indexs = p1_indexs
        self.p2_indexs = p2_indexs
        self.p3_indexs = p3_indexs

        # Default values are None unless assigned
        self.x1_p3, self.y1_p3, self.x2_p3, self.y2_p3 = None, None, None, None
        self.x1_p2, self.y1_p2, self.x2_p2, self.y2_p2 = None, None, None, None
        self.x1_p1, self.y1_p1, self.x2_p1, self.y2_p1 = None, None, None, None

        self.r2_p3, self.m_p3, self.b_p3 = None, None, None # y=mx+b
        self.r2_p2, self.m_p2, self.b_p2 = None, None, None
        self.r2_p1, self.m_p1, self.b_p1 = None, None, None

        self.obj_x_start, self.obj_y_start = None, None
        self.r2s, self.bs = None, None # Lists, y=mx+b

        self.x_p3, self.x_p2, self.x_p1, self.x_p12 = None, None, None, None
        self.y_p3, self.y_p2, self.y_p1, self.y_p12 = None, None, None, None
        self.x_p12_curvstrip_p3, self.y_p12_curvstrip_p3 = None, None
        self.x_p2_curvstrip_p3, self.y_p2_curvstrip_p3 = None, None
        self.x_p1_curvstrip_p3, self.y_p1_curvstrip_p3 = None, None
        self.x_p1_curvstrip_p23, self.y_p1_curvstrip_p23 = None, None

        self.k_p3, self.k_p2, self.k_p1 = None, None, None
        self.t05_p3, self.t05_p2, self.t05_p1 = None, None, None
        self.R0_p3, self.R0_p2, self.R0_p1 = None, None, None
        self.efflux_p3, self.efflux_p2, self.efflux_p1 = None, None, None

        self.elut_period, self.tracer_retained, self.poolsize = None, None, None
        self.influx, self.netflux, self.ratio = None, None, None
       

class Run(object):
    '''
    Class that stores ALL data of a single CATE run.
    This data includes values derived from objective or 
    subjetive analyses (calculated within the class)
    '''        
    def __init__(
    	   self, run_name, SA, rt_cnts, sht_cnts, rt_wght, gfact,
    	   load_time, elut_ends, elut_cpms):       
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
        # Basic analysis of CATE data
        self.elut_cpms_gfact = [x * self.gfact for x in self.elut_cpms]
        self.elut_cpms_gRFW = []
        self.elut_cpms_log = []

        for index, item in enumerate(self.elut_cpms_gfact):
            temp = item / self.rt_wght / \
                (self.elut_ends[index] - self.elut_starts[index])
            self.elut_cpms_gRFW.append(temp)
            self.elut_cpms_log.append(math.log10(temp))
        
		# x and y data for graphing (numpy-fied)
        self.x = numpy.array(self.elut_ends)
        self.y = numpy.array(self.elut_cpms_log)

def linear_regression(x, y):
    ''' Linear regression of x and y series (lists)
    Returns r^2 and m (slope), b (intercept) of y=mx+b
    '''
    coeffs = numpy.polyfit(x, y, 1)
        
    # Conversion to "convenience class" in numpy for working with polynomials        
    p = numpy.poly1d(coeffs)     
    
    # Determining R^2 on the current data set, fit values, and mean
    yhat = p(x)                         # or [p(z) for z in x]
    ybar = numpy.sum(y)/len(y)          # or sum(y)/len(y)
    ssreg = numpy.sum((yhat-ybar)**2)   # or sum([ (yihat - ybar)**2 for yihat in yhat])
    sstot = numpy.sum((y - ybar)**2)    # or sum([ (yi - ybar)**2 for yi in y])
    r2 = ssreg/sstot
    slope = coeffs[0]
    intercept = coeffs[1]
    
    return r2, slope, intercept

def find_obj_reg(elut_ends, elut_cpms_log, num_points):
    '''
    Use objective regression to determine 3 phases exchange in data
    Phase 3 is found by identifying where r2 decreases for 3 points in a row
    Phase 1 and 2 is found by identifying the paired series in the remaining 
        data that yields the highest combined r2s
    num_points allows the user to specificy how mmany points we will ignore at
        the end of the data series before we start comparing r2s
    Returns tuples outlining limits of each phase, m and b of phase 3, and 
        lists of initial intercepts(bs)/slopes(ms)/r2s
    '''
    r2s = []
    ms = [] # y = mx+b
    bs = []

    # Storing all possible r2s/ms/bs
    for index in range(len(elut_cpms_log)-2, -1, -1):
        temp_x = elut_ends[index:]
        temp_y = elut_cpms_log[index:]
        temp_r2, temp_m, temp_b = linear_regression(temp_x, temp_y)
        r2s =[temp_r2] + r2s
        ms = [temp_m] + ms
        bs = [temp_b] + bs
        
    # Determining the index at which r2 drops three times in a row 
    # from num_points from the end of the series
    counter = 0
    for index in range(len(elut_ends) - 1 - num_points, -1, -1):    
        print elut_ends[index], elut_cpms_log[index], r2s[index], counter
        if r2s[index-1] < r2s[index]:
            counter += 1
            if counter == 3:
                break
        else:
            counter = 0
    start_p3 = index + 3
    end_p3 = -1 # end indexs are going to be inclusive of last result in phase
    r2_p3 = r2s[start_p3]
    m_p3 = ms[start_p3]
    b_p3 = bs[start_p3]

    # print elut_ends[:start_p3_index], elut_cpms_log[index], r2s[index], counter
    # print r2s

    # Now we have to determine the p1/p2 combo that gives us the highest r2
    temp_x_p12 = elut_ends[:start_p3]
    temp_y_p12 = elut_cpms_log[:start_p3]
    highest_r2 = 0

    for index in range(1, len(temp_x_p12) - 2): #-2 bec. min. len(list) = 2
        temp_start_p2 = index + 1
        temp_x_p2 = temp_x_p12[temp_start_p2:]
        temp_y_p2 = temp_y_p12[temp_start_p2:]
        temp_x_p1 = temp_x_p12[:temp_start_p2]
        temp_y_p1 = temp_y_p12[:temp_start_p2]
        temp_r2_p2, temp_m_p2, temp_b_p2 = linear_regression(temp_x_p2, temp_y_p2)
        temp_r2_p1, temp_m_p1, temp_b_p1 = linear_regression(temp_x_p1, temp_y_p1)
        if temp_r2_p1 + temp_r2_p2 > highest_r2:
            # print temp_x_p1, temp_x_p2, temp_r2_p1, temp_r2_p2
            highest_r2 = temp_r2_p1 + temp_r2_p2
            start_p2, p2_end = temp_start_p2, start_p3
            start_p1, end_p1 = 0, temp_start_p2
            r2_p2, m_p2, b_p2 = temp_r2_p2, temp_m_p2, temp_b_p2 # these values are stored but I don't think I will need them (are calcylated by more general algorithms)
            r2_p1, m_p1, b_p1 = temp_r2_p1, temp_m_p1, temp_b_p1

    return start_p3, end_p3, start_p2, p2_end, start_p1, end_p1, r2s, ms, bs

    '''
	
    def objective_analysis(self):
	"""
	Objective analysis of RunObject using basic RunObject data
	"""
        # Getting parameters from regression of p3
        self.x1_p3, self.x2_p3, self.y1_p3, self.y2_p3, self.r2_p3, \
            self.slope_p3, self.intercept_p3, self.reg_end_index, \
            self.r2s_p3_list, self.slopes_p3_list, \
            self.intercepts_p3_list = Operations.obj_regression_p3(
                self.x, self.y, self.analysis_type[1])
	
	   # Setting the x- and y-series' involved in the p3/p12 linear regression
        self.x_p3 = self.x[self.reg_end_index:] 
        self.y_p3 = self.y[self.reg_end_index:]
        self.x_p12 = self.x[:self.reg_end_index]
        self.y_p12 = self.y[:self.reg_end_index]
	
        # Setting the x/y-series' used to start obj regression USED FOR PLOTTING
        self.x_reg_start = self.x[len(self.x) - self.analysis_type[1]:] 
        self.y_reg_start = self.y[len(self.x) - self.analysis_type[1]:]
        # Getting p1 + p2 curve-stripped data (together)
        self.x_p12_curvestrippedof_p3, \
                self.y_p12_curvestrippedof_p3 = Operations.p12_curvestrippedof_p3(
                    self.x_p12, self.y_p12, self.slope_p3, self.intercept_p3)
	   # Isolating/Unpacking PARTIALLY curve-stripped p1 x/y series
        self.x_p1_curvestrippedof_p3, \
                self.y_p1_curvestrippedof_p3 = Operations.determine_p1_xy (
                    self.x_p12_curvestrippedof_p3, self.y_p12_curvestrippedof_p3,)
        # Defining uncurve-stripped p1 data
        self.x_p1 = self.x[:len(self.x_p1_curvestrippedof_p3)]
        self.y_p1 = self.y[:len(self.y_p1_curvestrippedof_p3)]
        # Isolating COMPLETELY curve-stripped p2 data
        self.x_p2_curvestrippedof_p3, self.y_p2_curvestrippedof_p3 =\
            self.x_p12_curvestrippedof_p3[len(self.x_p1):],\
               self.y_p12_curvestrippedof_p3[len(self.y_p1):]
	
        # Linear regression of isolated p2 data + getting line data
        self.r2_p2, self.slope_p2, self.intercept_p2 = Operations.linear_regression(
            self.x_p2_curvestrippedof_p3, self.y_p2_curvestrippedof_p3)
        self.x1_p2, self.x2_p2, self.y1_p2, self.y2_p2 = Operations.grab_x_ys(
            self.x_p2_curvestrippedof_p3, self.intercept_p2, self.slope_p2)
	
        # Getting and plotting p1 curve-stripped data (for p2 and p3)
        self.x_p1_curvestrippedof_p23, self.y_p1_curvestrippedof_p23 = \
	           Operations.p1_curvestrippedof_p23(
	               self.x_p1_curvestrippedof_p3, self.y_p1_curvestrippedof_p3,
	               self.slope_p2, self.intercept_p2)
	
        # Linear regression on curve-stripped p1 data and plotting line
        self.r2_p1, self.slope_p1, self.intercept_p1 = Operations.linear_regression(
	           self.x_p1_curvestrippedof_p23, self.y_p1_curvestrippedof_p23)
        self.x1_p1, self.x2_p1, self.y1_p1, self.y2_p1 = Operations.grab_x_ys(
	           self.x_p1_curvestrippedof_p23, self.intercept_p1, self.slope_p1)
	
        # Setting rate constant (k), half life values (t05),
        # rate of release (R0), and efflux values of each phase
        self.k_p1 = abs(self.slope_p1 * 2.303)
        self.k_p2 = abs(self.slope_p2 * 2.303)
        self.k_p3 = abs(self.slope_p3 * 2.303)
        self.t05_p1 = 0.693/self.k_p1
        self.t05_p2 = 0.693/self.k_p2
        self.t05_p3 = 0.693/self.k_p3
        self.R0_p1 = 10 ** self.intercept_p1
        self.R0_p2 = 10 ** self.intercept_p2
        self.R0_p3 = 10 ** self.intercept_p3
	
        self.efflux_p1 = 60 * (
            self.R0_p1 / (self.SA * (1 - math.exp(-self.k_p1 * self.load_time))))
        self.efflux_p2 = 60 * (
            self.R0_p2 / (self.SA * (1 - math.exp(-self.k_p2 * self.load_time))))
        self.efflux_p3 = 60 * (
            self.R0_p3 / (self.SA * (1 - math.exp (-self.k_p3 * self.load_time))))
	
        self.elut_period = self.elut_ends[-1]
        self.tracer_retained = (self.sht_cnts + self.rt_cnts)/self.rt_wght
	
        self.netflux = 60 * (self.tracer_retained - (self.R0_p3/self.k_p3) *\
            math.exp (-self.k_p3 * self.elut_period))/self.SA/self.load_time
        self.influx = self.efflux_p3 + self.netflux
        self.ratio = self.efflux_p3 / self.influx
        self.poolsize = self.influx * self.t05_p3 / (3 * 0.693)
	
    def subjective_analysis(self):
	
        self.p1_start = self.analysis_type[1][0][0]
        self.end_p1 = self.analysis_type[1][0][1]
        self.p2_start = self.analysis_type[1][1][0]
        self.p2_end = self.analysis_type[1][1][1]	
        self.p3_start = self.analysis_type[1][2][0]
        self.p3_end = self.analysis_type[1][2][1]	
	
        # Getting parameters from regression of p3
        self.x1_p3, self.x2_p3, self.y1_p3, self.y2_p3, self.r2_p3, \
            self.slope_p3, self.intercept_p3 = Operations.subj_regression_p3(
                self.x, self.y, self.p3_start, self.p3_end)
	
        # Setting the x- and y-series' involved in the p3/p12 linear regression
        self.x_p3 = self.x[self.p3_start: self.p3_end + 1] 
        self.y_p3 = self.y[self.p3_start: self.p3_end + 1]
        self.x_p2 = self.x[self.p2_start: self.p2_end + 1] 
        self.y_p2 = self.y[self.p2_start: self.p2_end + 1]
        self.x_p1 = self.x[self.p1_start: self.end_p1 + 1]
        self.y_p1 = self.y[self.p1_start: self.end_p1 + 1]	
        self.x_p12 = np.concatenate([self.x_p1, self.x_p2])
        self.y_p12 = np.concatenate([self.y_p1, self.y_p2])
	
        # Getting p1 + p2 curve-stripped data (together)
        self.x_p12_curvestrippedof_p3, self.y_p12_curvestrippedof_p3 = \
            Operations.p12_curvestrippedof_p3 (
                self.x_p12, self.y_p12, self.slope_p3, self.intercept_p3)
	
        # Isolating/Unpacking PARTIALLY curve-stripped p1 x/y series
        self.x_p1_curvestrippedof_p3 =  self.x_p12_curvestrippedof_p3[self.p1_start: self.end_p1 + 1] 
        self.y_p1_curvestrippedof_p3 = self.y_p12_curvestrippedof_p3[self.p1_start: self.end_p1 + 1]
	
        # Isolating COMPLETELY curve-stripped p2 data
        self.x_p2_curvestrippedof_p3, self.y_p2_curvestrippedof_p3 =\
            self.x_p12_curvestrippedof_p3[len(self.x_p1):],\
            self.y_p12_curvestrippedof_p3[len(self.y_p1):]
	
    	# Linear regression of isolated p2 data + getting line data
    	self.r2_p2, self.slope_p2, self.intercept_p2 = Operations.linear_regression(
    	    self.x_p2_curvestrippedof_p3, self.y_p2_curvestrippedof_p3)
    	self.x1_p2, self.x2_p2, self.y1_p2, self.y2_p2 = Operations.grab_x_ys(
            self.x_p2_curvestrippedof_p3, self.intercept_p2, self.slope_p2)
	
    	# Getting and plotting p1 curve-stripped data (for p2 and p3)
    	self.x_p1_curvestrippedof_p23, self.y_p1_curvestrippedof_p23 = \
    	    Operations.p1_curvestrippedof_p23(
    	        self.x_p1_curvestrippedof_p3, self.y_p1_curvestrippedof_p3,
    	        self.slope_p2, self.intercept_p2)
	
        # Linear regression on curve-stripped p3 data and plotting line
        self.r2_p1, self.slope_p1, self.intercept_p1 = Operations.linear_regression(
            self.x_p1_curvestrippedof_p23, self.y_p1_curvestrippedof_p23)
        self.x1_p1, self.x2_p1, self.y1_p1, self.y2_p1 = Operations.grab_x_ys(
	           self.x_p1_curvestrippedof_p23, self.intercept_p1, self.slope_p1)
	
        # Setting rate constant (k), half life values t05, rate of release (R0)
        # and efflux values of each phase
        self.k_p1 = abs(self.slope_p1 * 2.303)
        self.k_p2 = abs(self.slope_p2 * 2.303)
        self.k_p3 = abs(self.slope_p3 * 2.303)	
        self.t05_p1 = 0.693/self.k_p1
        self.t05_p2 = 0.693/self.k_p2
        self.t05_p3 = 0.693/self.k_p3	
        self.R0_p1 = 10 ** self.intercept_p1
        self.R0_p2 = 10 ** self.intercept_p2
        self.R0_p3 = 10 ** self.intercept_p3	
        self.efflux_p1 = 60 * (self.R0_p1 / (
	           self.SA * (1 - math.exp(-self.k_p1 * self.load_time))))
        self.efflux_p2 = 60 * (self.R0_p2 / (
	           self.SA * (1 - math.exp(-self.k_p2 * self.load_time))))
        self.efflux_p3 = 60 * (self.R0_p3 / (
	           self.SA * (1 - math.exp (-self.k_p3 * self.load_time))))
	
        self.elut_period = self.elut_ends[-1]
        self.tracer_retained =\
            (self.sht_cnts + self.rt_cnts)/self.rt_wght
	
        self.netflux = 60 * (self.tracer_retained - (self.R0_p3/self.k_p3) * \
            math.exp (-self.k_p3 * self.elut_period))/self.SA/self.load_time
        self.influx = self.efflux_p3 + self.netflux
        self.ratio = self.efflux_p3 / self.influx
        self.poolsize = self.influx * self.t05_p3 / (3 * 0.693)
        '''
		
if __name__ == "__main__":
    import Excel
    temp_data = Excel.grab_data(r"C:\Users\Ruben\Projects\CATEAnalysis\Tests\1", "Test - Single Run.xlsx")
    
    temp_obj = temp_data.analyses[0]
    find_obj_reg(temp_obj.elut_ends, temp_obj.elut_cpms_log, 3)
    #print len(temp_object.elut_ends)
    #print temp_object.elut_cpms_gRFW