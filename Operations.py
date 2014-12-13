import numpy

x_series = [1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.5, 13.0, 14.5, 16.0, 17.5, 19.0, 20.5, 22.0, 23.5, 25.0, 27.0, 29.0, 31.0, 33.0, 35.0, 37.0, 39.0, 41.0, 43.0, 45.0]
y_series = [5.134446324653075, 4.532511080497156, 3.9647696512150836, 3.6692523925695686, 3.509950796085591, 3.3869391729764766, 3.287809993163619, 3.230048067964903, 3.169204739621747, 3.1203409378545346, 2.95145986473132, 2.8916143915841324, 2.8589559610792583, 2.8463057128814175, 2.8413779879066166, 2.7532261939625293, 2.750050822474359, 2.6735829597693206, 2.7024903224651338, 2.661606690643107, 2.5998423959455335, 2.57889496358432, 2.5921979525818397, 2.557187996704314, 2.529320391444595, 2.558194007072854, 2.4833719392530966, 2.5557756556810562, 2.4045248209763437, 2.4642132678099204]

def grab_x_ys (elution_ends, intercept, slope):
    ''' outputs two pairs of x,y coordinates for plotting a regression line
    all inputs are LISTS
    '''
    last_elution = elution_ends[len(elution_ends) - 1] # value from last y in the series
    x1 = 0
    x2 = last_elution
    y1 = intercept
    y2 = slope * last_elution + intercept
    
    return x1, x2, y1, y2

def linear_regression (x, y):
    # Linear regression of current set of values, returns m, b of y=mx+b
    coeffs = numpy.polyfit (x, y, 1)
        
    p = numpy.poly1d(coeffs) # Conversion to "convenience class" in numpy for working with polynomials        
    
    # Determining R^2 on the current data set        
    # fit values, and mean
    yhat = p(x)                         # or [p(z) for z in x]
    ybar = numpy.sum(y)/len(y)          # or sum(y)/len(y)
    ssreg = numpy.sum((yhat-ybar)**2)   # or sum([ (yihat - ybar)**2 for yihat in yhat])
    sstot = numpy.sum((y - ybar)**2)    # or sum([ (yi - ybar)**2 for yi in y])
    r2 = ssreg/sstot
    slope = coeffs [0]
    intercept = coeffs [1]
    
    return r2, slope, intercept

def obj_regression_p12 (elution_ends, log_efflux, last_used_index):     
    ''' Figuring out the best regression lines for phases 1 and 2
    
    last_used_index is the first right-most point not used in the p3 regression
    '''
    x = elution_ends[:last_used_index]
    y = log_efflux [:last_used_index]
    # Initialize the indexs for the first regressions
    p1_end_index = 2
    p2_start_index = 2
    current_high_r2 = 0
    
    while p2_start_index < last_used_index - 1: # Correcting as ...
        # ... otherwise x2_current = x [p2_start_index:] is only 1 entry. Last index is not inclusive
        # Setting current series
        x1_current = x [:p1_end_index]
        y1_current = y [:p1_end_index]
        x2_current = x [p2_start_index:]
        y2_current = y [p2_start_index:]     
        # Doing the current set of regressions
        current_p1_regression = linear_regression (x1_current, y1_current)
        current_p2_regression = linear_regression (x2_current, y2_current)
    
        r2_1 = current_p1_regression[0]        
        slope_1= current_p1_regression[1]
        intercept_1= current_p1_regression[2]
        
        r2_2 = current_p2_regression[0]        
        slope_2= current_p2_regression[1]
        intercept_2= current_p2_regression[2]
        
        if r2_1 + r2_2 > current_high_r2:
            current_high_r2 = r2_1 + r2_2
    
            best_x1 = x1_current
            best_y1 = y1_current
            best_r2_1 = r2_1
            best_slope_1 = slope_1
            best_intercept_1 = intercept_1
            
            best_x2 = x2_current
            best_y2 = y2_current
            best_r2_2 = r2_2
            best_slope_2 = slope_2
            best_intercept_2 = intercept_2
        # Moving the index counters forward
        p1_end_index += 1
        p2_start_index += 1
    # END OF LOOP
    
    # Grabbing parameters from phase 1 regression line
    x1_1, x2_1, y1_1, y2_1 = grab_x_ys (best_x1, best_intercept_1, best_slope_1)
    # Grabbing parameters from phase 2 regression line
    x1_2, x2_2, y1_2, y2_2 = grab_x_ys (best_x2, best_intercept_2, best_slope_2)
    
    final_regression_p1 = [best_x1, best_y1, x1_1, x2_1, y1_1, y2_1, best_r2_1, best_slope_1, best_intercept_1]
    final_regression_p2 = [best_x2, best_y2, x1_2, x2_2, y1_2, y2_2, best_r2_2, best_slope_2, best_intercept_2]
    
    # Convert data in our regression lines from numpy data types to native python ones| PROBABLY NOT NEEDED, FIX THAT TO A PROBLEM THAT WAS COMPLETELY DIFFERENT
    '''
    for x in range(3, len(final_regression_p1)):
        final_regression_p1[x] = numpy.asscalar(numpy.array([final_regression_p1[x]]))
        final_regression_p2[x] = numpy.asscalar(numpy.array([final_regression_p2[x]]))
    '''
    return final_regression_p1, final_regression_p2
  
    
def obj_regression_p3 (elution_ends, log_efflux, num_points_reg):     
    ''' Figuring out R^2 (correlation) and parameters of linear regression.
    Linear regression gives the data about the 3rd phase of exchange (p3)
    
    elutions_ends is the times elutions WERE CHANGED. 
    elution_starts would be times elutions ADDED.
    
    Line parameters for the linear regression are returned as well as the lists containing the regression
    parameters for ALL the regressions done in the objective regression
    
    In the lists, the first three entries are parameters from regressions that are 
    discarded (r2 was decreasing).
    
    Parameters of interest are at the 4th index of each list (List[3])
    
    INPUT:
    X series (list) - elution end points (min)
    Y series (list) - normalized efflux data (log cpm/g RFW)
    
    RETURNED:
    Line parameters for the linear regression (x1, x2, y1, y2, r2, slope; ints)
    '''
    
    # Lists to store the values from the SERIES of regressions.
    # Values from the most recent regression are stored at index 0 of each list
    r2 = [] 
    slope = []
    intercept = []
    start_index = int(len (log_efflux) - num_points_reg)
    
    print elution_ends[start_index]
    reg_dec = False # Variable to track if regression should continue
    
    current_index = start_index
    
    while reg_dec == False: # R^2 has not decreased 3 consecutive times
        x = elution_ends [current_index:]
        y = log_efflux [current_index:]
        
        current_regression = linear_regression (x, y)
        r2.insert (0, current_regression[0])
        slope.insert (0, current_regression[1])
        intercept.insert (0, current_regression[2])        
               
        # Checking to see if you are out of the third phase of exchange
        # (has R^2 decreased 3 consecutive times)
        if len (r2) <= 3:
            current_index -= 1 # Extending the series of values included in the analysis/regression
        elif r2[0] < r2[1] and r2[1] < r2[2] and r2[2] < r2[3]:
            reg_dec = True
        else:
            current_index -= 1 # Extending the series of values included in the analysis/regression
    
    # Write points for graphing equation of linear regression
    x1, x2, y1, y2 = grab_x_ys (elution_ends, intercept [3], slope[3])
    
    return x1, x2, y1, y2, r2[3], slope [3], intercept [3], current_index

def subj_regression (elution_ends, log_efflux, reg_start, reg_end): 
    ''' Doing a custom regression on a data set wherein
    the limits of the regression are known.
    
    reg_start and reg_end are INDEXS corresponding to the lists elution_ends
    and log_efflux
    
    Unlike obj_regression, returning variables r2, slope, intercept are not
    lists but single values
    '''
        
    x = elution_ends [reg_start: reg_end +1] # last index is not inclusive!!
    y = log_efflux [reg_start: reg_end + 1]
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
    test = obj_regression_p3 (x_series, y_series, 2)
    print test [5]
