import numpy as np
import math

import Operations

class RunObject():
    '''
    Class that stores ALL data in a single run.
    This data includes values derived from objective or 
    subjetive analysis (calculated within the class)
    '''
        
    def __init__(self, run_name, SA, root_cnts, shoot_cnts, root_weight,\
                 g_factor, load_time, elution_times, elution_cpms, pts_used):       
        self.run_name = run_name
        self.SA = SA
        self.root_cnts = root_cnts
        self.shoot_cnts = shoot_cnts
        self.root_weight = root_weight
        self.g_factor = g_factor
        self.load_time = load_time
        self.elution_times = elution_times
        self.elution_cpms = elution_cpms
	self.pts_used = pts_used
	
	# Default analysis is objective used last 3 points
	self.analysis_type = ("obj", pts_used)
	self.elution_ends = elution_times[1:]
	self.elution_starts = elution_times[:len (elution_times) - 1]
	
	temp = Operations.basic_CATE_analysis(
		    SA, root_cnts, shoot_cnts, root_weight, g_factor,\
	            load_time, elution_times, elution_cpms)
		
	self.elution_cpms_gfactor = temp[0]
	self.elution_cpms_gRFW = temp[1]
	self.elution_cpms_log = temp[2]
	
	# x and y data for graphing (numpy-fied)
	self.x = np.array(self.elution_ends)
	self.y = np.array(self.elution_cpms_log)
	
	# Default analysis is objective analysis with 2 points
	self.objective_analysis()
	
    def objective_analysis(self):
	"""
	Objective analysis of RunObject using basic RunObject data
	"""
    
	# Getting parameters from regression of p3
	self.x1_p3, self.x2_p3, self.y1_p3, self.y2_p3, self.r2_p3,\
	            self.slope_p3, self.intercept_p3, self.reg_end_index,\
	            self.r2s_p3_list, self.slopes_p3_list,\
	            self.intercepts_p3_list=\
	            Operations.obj_regression_p3(
	                self.x,
	                self.y,
	                self.analysis_type[1]
	            )
	
	# Setting the x- and y-series' involved in the p3/p12 linear regression
	self.x_p3 = self.x[self.reg_end_index:] 
	self.y_p3 = self.y[self.reg_end_index:]
	
	self.x_p12 = self.x[:self.reg_end_index]
	self.y_p12 = self.y[:self.reg_end_index]
	
	# Setting the x/y-series' used to start obj regression
	self.x_reg_start = self.x[len(self.x) - self.analysis_type[1]:] 
	self.y_reg_start = self.y[len(self.x) - self.analysis_type[1]:]
	
	# Getting p1 + p2 curve-stripped data (together)
	self.x_p12_curvestrippedof_p3, self.y_p12_curvestrippedof_p3 = \
	    Operations.p12_curvestrippedof_p3 (
	        self.x_p12,
	        self.y_p12,
	        self.slope_p3,
	        self.intercept_p3
	    )
	
	# Isolating/Unpacking PARTIALLY curve-stripped p1 x/y series
	self.x_p1_curvestrippedof_p3, self.y_p1_curvestrippedof_p3 = \
	    Operations.determine_p1_xy (
	        self.x_p12_curvestrippedof_p3,
	        self.y_p12_curvestrippedof_p3,
	    )
	
	# Defining uncurve-stripped p1 data
	self.x_p1 = self.x[:len(self.x_p1_curvestrippedof_p3)]
	self.y_p1 = self.y[:len(self.y_p1_curvestrippedof_p3)]
	 	
	# Isolating COMPLETELY curve-stripped p2 data
	self.x_p2_curvestrippedof_p3, self.y_p2_curvestrippedof_p3 =\
	    self.x_p12_curvestrippedof_p3[len(self.x_p1):],\
	    self.y_p12_curvestrippedof_p3[len(self.y_p1):]
	
	# Linear regression of isolated p2 data + getting line data
	self.r2_p2, self.slope_p2, self.intercept_p2 = Operations.linear_regression (
	    self.x_p2_curvestrippedof_p3,
	    self.y_p2_curvestrippedof_p3
	)
	self.x1_p2, self.x2_p2, self.y1_p2, self.y2_p2 = Operations.grab_x_ys(
	    self.x_p2_curvestrippedof_p3,
	    self.intercept_p2,
	    self.slope_p2
	)
	
	# Getting and plotting p1 curve-stripped data (for p2 and p3)
	self.x_p1_curvestrippedof_p23, self.y_p1_curvestrippedof_p23 = \
	    Operations.p1_curvestrippedof_p23 (
	        self.x_p1_curvestrippedof_p3,
	        self.y_p1_curvestrippedof_p3,
	        self.slope_p2,
	        self.intercept_p2
	    )
	
	# Linear regression on curve-stripped p3 data and plotting line
	self.r2_p1, self.slope_p1, self.intercept_p1 = Operations.linear_regression (
	    self.x_p1_curvestrippedof_p23,
	    self.y_p1_curvestrippedof_p23
	)
	self.x1_p1, self.x2_p1, self.y1_p1, self.y2_p1 = Operations.grab_x_ys(
	    self.x_p1_curvestrippedof_p23,
	    self.intercept_p1,
	    self.slope_p1
	)
	
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
	
	self.elution_period = self.elution_ends[-1]
	self.tracer_retained =\
	    (self.shoot_cnts + self.root_cnts)/self.root_weight
	
	self.netflux = 60 * (self.tracer_retained - (self.R0_p3/self.k_p3) *  math.exp (-self.k_p3 * self.elution_period))/self.SA/self.load_time
	self.influx = self.efflux_p3 + self.netflux
	self.ratio = self.efflux_p3/self.influx
	self.poolsize = self.influx * self.t05_p3/ (3 * 0.693)
	
	
	
if __name__ == "__main__":
    pass