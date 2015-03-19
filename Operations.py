import numpy
import math

y_series = [5.134446324653075, 4.532511080497156, 3.9647696512150836, 3.6692523925695686, 3.509950796085591, 3.3869391729764766, 3.287809993163619, 3.230048067964903, 3.169204739621747, 3.1203409378545346, 2.95145986473132, 2.8916143915841324, 2.8589559610792583, 2.8463057128814175, 2.8413779879066166, 2.7532261939625293, 2.750050822474359, 2.6735829597693206, 2.7024903224651338, 2.661606690643107, 2.5998423959455335, 2.57889496358432, 2.5921979525818397, 2.557187996704314, 2.529320391444595, 2.558194007072854, 2.4833719392530966, 2.5557756556810562, 2.4045248209763437, 2.4642132678099204]

def basic_CATE_analysis(SA, root_cnts, shoot_cnts, root_weight, g_factor,\
                         load_time, elution_times, elution_cpms):
    '''Given initial CATE data, return elution data as corrected for
    G Factor, normalized for root weight, and logged
    '''
    
    elution_cpms_gfactor = []
    elution_cpms_gRFW = []
    elution_cpms_log = []
    
    for x in range (0, len(elution_cpms)):
        elution_cpms_gfactor.append (elution_cpms[x] * g_factor)    
        
    for x in range (0, len(elution_cpms_gfactor)):
        temp = elution_cpms_gfactor[x] / root_weight / \
                                  (elution_times[x+1] - elution_times[x])
        elution_cpms_gRFW.append(temp)
        elution_cpms_log.append(math.log10 (temp))
                
    return elution_cpms_gfactor, elution_cpms_gRFW, elution_cpms_log

def antilog(x):
    return 10 ** x

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

def grab_x_ys(elution_ends, intercept, slope):
    ''' 
    Output two pairs of (x, y) coordinates for plotting a regression line
    elution_ends is the x-series (list)
    intrecept and slope are ints
    '''
    
    last_elution = elution_ends[len(elution_ends) - 1] 
    
    x1 = 0
    x2 = last_elution
    y1 = intercept
    y2 = slope * last_elution + intercept
    
    return x1, x2, y1, y2

def p1_curve_stripped(p1_x, p1_y, last_used_index, slope, intercept):
    '''
    Curve-strip p1 (in the same list) according to p2 data
    Note: p1 data coming in (log_efflux) has been corrected for p3 already
    
    INPUT:
    p1_x (x-series; list) - elution end points (min)
    p1_y (y-series; list) - normalized efflux data (log cpm/g RFW)
    slope, intercept (ints) - line parameters of p2 regression
        
    RETURNED:
    p12_x (x-series; list) - elution_ends limited to range of p1 and p2
    corrected_p12_y (y-series; list) - corrected (curve-stripped) efflux data
    '''
    # Calculating p1 data from extrapolation of p2 linear regression into p1

    # Container for p3 data in p1/2 range
    p2_extrapolated_raw = []
    
    for x1 in range (0, len(p1_x)):
        extrapolated_x = (slope * x1) + intercept
        p2_extrapolated_raw.append(extrapolated_x)    

def p12_curve_stripped(x_p12, y_p12, elution_ends, last_used_index, slope, intercept):
    '''
    Curve-strip p1 and p2 (in the same list) according to p3 data
    
    INPUT:
    elution_ends (x-series; list) - elution end points (min)
    log_efflux (y-series; list) - normalized efflux data (log cpm/g RFW)
    last_used_index is the first right-most point used in the p3 regression
        - used as end index for p2 because [x:y] y IS NOT INCLUSIVE
    
    RETURNED:
    p12_x (x-series; list) - elution_ends limited to range of p1 and p2
    corrected_p12_y (y-series; list) - corrected (curve-stripped) efflux data
    '''
    
    # Calculating p3 data from extrapolation of p3 linear regression into p1/2

    # Container for p3 data in p1/2 range
    p3_extrapolated_raw = []
    
    for x1 in range (0, len(x_p12)):
        extrapolated_x = (slope*x1) + intercept
        p3_extrapolated_raw.append(extrapolated_x)
        
    # Antilog extrapolated p3 data and p1/2 data, subtract them, and relog them
    
    # Container for curve-stripped p1/2 data
    corrected_p12_x = []
    corrected_p12_y = []
    
    for value in range (0, len(y_p12)):
        # print p12_y [value], p3_extrapolated_raw [value]
        antilog_orig = antilog(y_p12 [value])
        antilog_reg = antilog(p3_extrapolated_raw [value])
        corrected_p12_x_raw = antilog_orig - antilog_reg
        if corrected_p12_x_raw > 0: # We can perform a log operation
            corrected_p12_y.append(math.log10 (corrected_p12_x_raw))
            corrected_p12_x.append(elution_ends[value])
                
    return corrected_p12_x, corrected_p12_y

def p1_curve_stripped(elution_ends, log_efflux, last_used_index, slope, intercept):
    '''
    Curve-strip p1 according to p2/3 data
    
    INPUT:
    elution_ends (x-series; list) - elution end points (min)
    log_efflux (y-series; list) - normalized efflux data (log cpm/g RFW)
    last_used_index is the first right-most point used in the p3 regression
        - used as end index for p2 because [x:y] y IS NOT INCLUSIVE
    
    RETURNED:
    p12_x (x-series; list) - elution_ends limited to range of p1 and p2
    corrected_p12_y (y-series; list) - corrected (curve-stripped) efflux data
    '''
    # Broader x and y series containing both p1 and p2
    x_p12 = elution_ends[:last_used_index]
    y_p12 = log_efflux[:last_used_index]
    
    # Calculating p3 data from extrapolation of p3 linear regression into p1/2

    # Container for p3 data in p1/2 range
    p3_extrapolated_raw = []
    
    for x1 in range (0, len(x_p12)):
        extrapolated_x = (slope*x1) + intercept
        p3_extrapolated_raw.append(extrapolated_x)
        
    # Antilog extrapolated p3 data and p1/2 data, subtract them, and relog them
    
    # Container for curve-stripped p1/2 data
    corrected_p12_x = []
    corrected_p12_y = []
    
    for value in range (0, len(y_p12)):
        # print p12_y [value], p3_extrapolated_raw [value]
        antilog_orig = antilog(y_p12 [value])
        antilog_reg = antilog(p3_extrapolated_raw [value])
        corrected_p12_x_raw = antilog_orig - antilog_reg
        if corrected_p12_x_raw > 0: # We can perform a log operation
            corrected_p12_y.append(math.log10 (corrected_p12_x_raw))
            corrected_p12_x.append(elution_ends[value])
                
    return corrected_p12_x, corrected_p12_y
   
def determine_p1_xy(p12_elution_ends, p12_log_efflux):     
    ''' Figuring out x/y-series that yield strongest correlations 
    (regression lines) for phases 1 and 2
    
    INPUT:
    elution_ends (x-series; list) - elution end points (min) of p1 + p2
    log_efflux (y-series; list) - normalized efflux data (log cpm/g RFW) of
         of p1 + p2
    
    RETURNED:
    Ordered lists (best_x1, best_y1) contained x and y series of p1 that yields
       the highest R^2 taking into account combined R^2 of possilbe p1/p2
    '''
    
    # Initialize the index(s) for the first regressions
    demarcation_index = 2 # Start regression using 1st 2 pts vs rest of series
    current_high_r2 = 0 # Tracking highest summed R^2 from p1 and p2
    last_used_index = len(p12_elution_ends)
        
    while demarcation_index < last_used_index - 1: 
        # Correcting loop end (last_used_index - 1) 
        # as otherwise x2_current = x [demarcation_index:] is only 1 entry.
        # Last index in lists is not inclusive!
        # So we can use one var to track end and start indexs of p1/2 regression
        
        # Setting current series for p1 regression
        x1_current = p12_elution_ends[:demarcation_index]
        y1_current = p12_log_efflux[:demarcation_index]
        # Setting current series for p2 regression
        x2_current = p12_elution_ends[demarcation_index:]
        y2_current = p12_log_efflux[demarcation_index:]     
        
        # Doing the current set of regressions
        current_p1_regression = linear_regression(x1_current, y1_current)
        current_p2_regression = linear_regression(x2_current, y2_current)
    
        current_r2_p1 = current_p1_regression[0]        
        
        current_r2_p2 = current_p2_regression[0]        
        
        # Checking to see if our last R^2 is higher than highest R^2
        if current_r2_p1 + current_r2_p2 > current_high_r2:
            # If it is the new highest R^2 we store p1/2 regression parameters
            current_high_r2 = current_r2_p1 + current_r2_p2
    
            best_x1 = x1_current
            best_y1 = y1_current
            best_r2_1 = current_r2_p1
            best_slope_1 = current_p1_regression[1]
            best_intercept_1 = current_p1_regression[2]
            
            best_x2 = x2_current
            best_y2 = y2_current
            best_r2_2 = current_r2_p2
            best_slope_2 = current_p2_regression[1]
            best_intercept_2 = current_p2_regression[2]
        # Moving the index counter forward
        demarcation_index += 1
        
    return best_x1, best_y1
  
    
def obj_regression_p3(elution_ends, log_efflux, num_points_reg):     
    ''' Figuring out R^2 (correlation) and parameters of linear regression.
    Linear regression gives the data about the 3rd phase of exchange (p3)
    
    elutions_ends is the times elutions WERE CHANGED/REMOVED. 
    elution_starts would be times elutions ADDED.
    
    Line parameters for the linear regression are returned as well as the lists containing the regression
    parameters for ALL the regressions done in the objective regression
    
    In the lists, the first three entries are parameters from regressions that are 
    discarded (r2 was decreasing).
    
    Parameters of interest are at the 4th index of each list (List[3])
    
    INPUT:
    elution_ends (x-series; list) - elution end points (min)
    log_efflux (y-series; list) - normalized efflux data (log cpm/g RFW)
    
    RETURNED:
    Line parameters for the linear regression (x1, x2, y1, y2, r2, slope; ints)
    '''
    
    # Lists to store the values from the SERIES of regressions.
    # Values from the most recent regression are stored at index 0 of each list
    r2 = [] 
    slope = []
    intercept = []
    
    # Variable to track current start index for x+y series to be analyzed
    current_index = int(len(log_efflux) - num_points_reg)
    
    # Variable to track if regression should continue
    reg_dec = False
    
    while reg_dec == False: # R^2 has not decreased 3 consecutive times
        x = elution_ends[current_index:]
        y = log_efflux[current_index:]
        
        current_regression = linear_regression(x, y)
        r2.insert(0, current_regression[0])
        slope.insert(0, current_regression[1])
        intercept.insert(0, current_regression[2])        
               
        # Checking to see if you are out of the third phase of exchange
        # (has R^2 decreased 3 consecutive times)
        if len (r2) <= 3:
            # Extending the x+y series of values included in the regression
            current_index -= 1
        elif r2[0] < r2[1] and r2[1] < r2[2] and r2[2] < r2[3]:
            reg_dec = True
        else:
            # Extending the x+y series of values included in the regression
            current_index -= 1
    
    # Write points for graphing equation of linear regression
    x1, x2, y1, y2 = grab_x_ys(elution_ends, intercept [3], slope[3])
    
    # current_index, r2, slope, and intercept are corrected considering
    # that the last 3 regressions done with decreasing R^2s
    return x1, x2, y1, y2, r2[3], slope [3], intercept [3], current_index + 3,\
           r2, slope, intercept

def subj_regression(elution_ends, log_efflux, reg_start, reg_end): 
    ''' Doing a custom regression on a data set wherein
    the limits of the regression are known.
    
    reg_start and reg_end are INDEXS corresponding to the lists elution_ends
    and log_efflux
    
    Unlike obj_regression, returning variables r2, slope, intercept are not
    lists but single values
    '''
        
    x = elution_ends[reg_start: reg_end +1] # last index is not inclusive!!
    y = log_efflux[reg_start: reg_end + 1]
    # print x
    # print y
    
    # Doing the regression and storing the parameters
    current_regression = linear_regression (x, y)
    
    r2 = current_regression[0]
    slope= current_regression[1]
    intercept= current_regression[2]
    
    # Write points for graphing equation of linear regression
    x1, x2, y1, y2 = grab_x_ys (elution_ends, intercept, slope)
    
    return x1, x2, y1, y2, r2, slope, intercept
    
if __name__ == '__main__':
    test = obj_regression_p3(x_series, y_series, 2)
    '''print test [7]
    print x_series [test [7]]
    print y_series [test [7]]'''
