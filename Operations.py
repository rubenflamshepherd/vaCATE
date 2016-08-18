import numpy
import math
import Objects

def grab_x_ys(elution_ends, slope, intercept):
    ''' 
    Output two pairs of (x, y) coordinates for plotting a regression line
    elution_ends is the x-series (list), intercept and slope are ints
    '''
    last_elution = elution_ends[len(elution_ends) - 1] 
    
    x1 = 0
    x2 = last_elution   
    y1 = intercept
    y2 = slope * last_elution + intercept
    
    return (x1, y1), (x2, y2)

def basic_run_calcs(rt_wght, gfact, elut_starts, elut_ends, elut_cpms):
    '''Given initial CATE data, return elution data as corrected for
    G Factor, normalized for root weight, and logged
    '''    
    elut_cpms_gfact, elut_cpms_gRFW, elut_cpms_log = [], [], []
    elut_ends_parsed = [] # y-series for parsec data (no '' or 0)

    for index, item in enumerate(elut_cpms):
        if item: # Our trigger to skip data point ('' or 0)
            temp_gfact = item * gfact
            temp_gRFW = item * gfact / rt_wght\
                / (elut_ends[index] - elut_starts[index])
            elut_cpms_gfact.append(temp_gfact)
            elut_cpms_gRFW.append(temp_gRFW)
            elut_cpms_log.append(math.log10(temp_gRFW))
            elut_ends_parsed.append(elut_ends[index])
                
    return elut_cpms_gfact, elut_cpms_gRFW, elut_cpms_log, elut_ends_parsed

def set_obj_phases(run, obj_num_pts):
    '''
    Use objective regression to determine the limits of the 3 phases of exchange
    Phase 3 is found by identifying where r2 decreases for 3 points in a row
        (refactored to get_obj_phase3 for testing purposes)
    Phase 1 and 2 is found by identifying the paired series in the remaining 
        data that yields the highest combined r2s
    obj_num_pts allows the user to specificy how many points we will ignore at
        the end of the data series before we start comparing r2s
    Returns tuples outlining limits of each phase
    '''
    elut_ends_log, elut_cpms_log = run.elut_ends_log, run.elut_cpms_log
    start_p3, end_p3, r2s = get_obj_phase3(
        elut_ends_log, elut_cpms_log, obj_num_pts)

    start_p2, end_p2, start_p1, end_p1, highest_r2 = get_obj_phase12(
        elut_ends_log, elut_cpms_log, start_p3)

    return (start_p3, end_p3), (start_p2, end_p2), (start_p1, end_p1)

def get_obj_phase3(elut_ends_log, elut_cpms_log, obj_num_pts):
    '''
    Determine limits of phase 3 using objective regression (point in data series
        from which r2 decreases 3 times in a row). Refactored out of
        set_obj_phases so testing functions can access list of r2s (r2s)
    '''
    r2s = []
    ms = [] # y = mx+b
    bs = []
    # Storing all possible r2s/ms/bs
    for index in range(len(elut_cpms_log) - 2, -1, -1):
        temp_x = elut_ends_log[index:]
        temp_y = elut_cpms_log[index:]
        temp_r2, temp_m, temp_b = linear_regression(temp_x, temp_y)
        r2s =[temp_r2] + r2s
        ms = [temp_m] + ms
        bs = [temp_b] + bs
        
    # Determining the index at which r2 drops three times in a row 
    # from obj_num_pts from the end of the series
    counter = 0
    for index in range(len(elut_ends_log) - obj_num_pts, -1, -1):
        # print elut_ends_log[index], r2s[index], counter
        if r2s[index - 1] < r2s[index]:
            counter += 1
            if counter == 3:
                break
        else:
            counter = 0
    start_p3 = index + 2 # Last index compared is not entered!
    end_p3 = len(elut_ends_log)

    return start_p3, end_p3, r2s # r2s is returned for testing purposes

def get_obj_phase12(elut_ends_log, elut_cpms_log, start_p3):
    '''
    Determine limits of phase 1+2 using objective regression (x and y series 
        for phase 1 and 2 that yield highest combined r2. Refactored out of
        set_obj_phases so testing functions can access highest_r2
    '''
    temp_x_p12 = elut_ends_log[:start_p3]
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
            highest_r2 = temp_r2_p1 + temp_r2_p2
            start_p2, end_p2 = temp_start_p2, start_p3
            start_p1, end_p1 = 0, temp_start_p2
        # print temp_x_p1, temp_x_p2, str(temp_r2_p1 + temp_r2_p2), highest_r2

    return start_p2, end_p2, start_p1, end_p1, highest_r2

def extract_phase(xs, x_series, y_series, SA, load_time):
    '''
    Extract and return parameters from regression analysis of a phase from 
        CATE run efflux trace.
    start/end are (x, y) 
    x and y are data from which we are extracting phase parameters.
    These may have holes from avoiding negative log operations during
        curvestripping of phase II + I
        - as a result of this, indexs may not line up nicely
        - must check which x-value elut_ends_log(index[0]/index1) lines up with
        - this is the fuctional index for our purposes.
    '''
    x_start, x_end = xs
    start_index = x_to_index(
        x_value=x_start, x_series=x_series)
    end_index = x_to_index(
        x_value=x_end, x_series=x_series)

    x_phase = x_series[start_index:end_index + 1]
    y_phase = y_series[start_index:end_index + 1]

    r2, slope, intercept = linear_regression(x_phase, y_phase) # y=(M)x+(B)
    xy1, xy2 = grab_x_ys(x_phase, slope, intercept)
    k = abs(slope) * 2.303
    t05 = 0.693/k
    r0 = 10 ** intercept
    efflux = 60 * r0 / (SA * (1 - math.exp(-1 * k * load_time)))

    return Objects.Phase(
        xs, xy1, xy2, r2, slope, intercept,
        x_phase, y_phase, k, t05, r0, efflux)

def x_to_index(x_value, x_series):
    for index, item in enumerate(x_series):
        if item == x_value:
            return index

def advanced_run_calcs(analysis):
    '''
    Do later CATE calculations that require some already extracted parameters.
    Entire analysis object is imported because many parameters needed and allows
    attributes to be manipulated directly (nothing returned)
    '''
    analysis.elut_period = analysis.run.elut_ends[-1]
    analysis.tracer_retained = \
        (analysis.run.sht_cnts + analysis.run.rt_cnts)/analysis.run.rt_wght
    analysis.netflux =\
        60 * (analysis.tracer_retained - (analysis.phase3.r0/analysis.phase3.k) *
        math.exp(-1 * analysis.phase3.k * analysis.elut_period))/ \
        analysis.run.SA/analysis.run.load_time
    analysis.influx = analysis.phase3.efflux + analysis.netflux
    analysis.ratio = analysis.phase3.efflux / analysis.influx
    analysis.poolsize = analysis.influx * analysis.phase3.t05 / (3 * 0.693)

def linear_regression(x_series, y_series):
    ''' Linear regression of x_series and y_series (lists)
    Returns r^2 and m (slope), b (intercept) of y=mx+b
    '''
    
    coeffs = numpy.polyfit(x_series, y_series, 1)        
    # Conversion to "convenience class" in numpy for working with polynomials        
    p = numpy.poly1d(coeffs)         
    # Determining R^2 on the current data set, fit values, and mean
    yhat = p(x_series) # or [p(z) for z in x_series]
    ybar = numpy.sum(y_series)/len(y_series) # or sum(y_series)/len(y_series)
    ssreg = numpy.sum((yhat-ybar)**2)#or sum([(yihat-ybar)**2 for yihat in yhat])
    sstot = numpy.sum((y_series - ybar)**2) #or sum([ (yi - ybar)**2 for yi in y_series])
    r2 = ssreg/sstot
    slope = coeffs[0]
    intercept = coeffs[1]
    
    return r2, slope, intercept

def curvestrip(x_series, y_series, slope, intercept):
    '''
    Curve-strip a series of data (x_series and y_series) according to data
    from a later phase (slope, intercept)
    Returns the curve-stripped x and y series
    '''
    # Calculating later phase data extrapolated into earlier phase
    extrapolated_raw = []
    for item in x_series:
        extrapolated_x = (slope*item) + intercept
        extrapolated_raw.append(extrapolated_x)
        
    # Antilog extrapolated p3 data and p1/2 data, subtract them, and relog them
    # Containers for curve-stripped p1/2 data
    y_curvestrip = []
    x_curvestrip = [] # need x series because x_curvestrip maybe != x_series

    for index, item in enumerate(y_series):
        antilog_orig = 10 ** item
        antilog_reg = 10 ** extrapolated_raw[index]
        curvestrip_x_raw = antilog_orig - antilog_reg
        if curvestrip_x_raw > 0: # We can perform a log operation
            y_curvestrip.append(math.log10 (curvestrip_x_raw))
            x_curvestrip.append(x_series[index])
        else: # No log operation possible. Data omitted from series
            pass
    return x_curvestrip, y_curvestrip

if __name__ == '__main__':
    pass