import numpy
import math
import Operations


class Experiment(object):
    '''Experiment containing all data

    === Attributes ===
    @type directory: str
        Folder where our data is/analysis will output
    @type analyses: list[Analysis]
        Individual CATE run objects (includes current analysis implemented)
    '''
    def __init__(self, directory, analyses):
        '''Experiment object with all relevant data

        @type self: Experiment
        @type directory: str
            Folder where our data is/analysis will output
        @type analyses: list[Analysis]
            Individual CATE run objects (includes current analysis implemented)
        @rtype: None
        '''
        self.directory = directory
        self.analyses = analyses # List of Analysis objects


class Analysis(object):
    '''Analysis of a single CATE run/replicate

    === Attributes ===
    @type kind: None | 'obj' | 'subj'
        The kind of analysis currently implemented. Default in None
    @type obj_num_pts: int
        Number of points that objective analysis is to be done with. 
        Set to 8 for the default objective analysis.
    @type run: Run
        Run object containing basic run data
    @type xs_p1: ('', '') | (float, float)
        x-values of boundaries of phase 1. Default are empty strings.
    @type xs_p2: ('', '') | (float, float)
        x-values of boundaries of phase 2. Default are empty strings.
    @type xs_p3: ('', '') | (float, float)
        x-values of boundaries of phase 3. Default are empty strings.                
    '''
    def __init__(self, kind, obj_num_pts, run, xs_p1=('', ''), 
            xs_p2=('', ''), xs_p3=('', '')):
        ''' Constructor of Analysis object

        @type kind: None | 'obj' | 'subj'
            The kind of analysis currently implemented. Default in None
        @type obj_num_pts: int | None
            Number of points that objective analysis is to be done with. 
            Set to 8 for the default objective analysis.
        @type run: Run
            Run object containing basic run data
        @type xs_p1: ('', '') | (float, float)
            x-values of boundaries of phase 1. Default are empty strings.
        @type xs_p2: ('', '') | (float, float)
            x-values of boundaries of phase 2. Default are empty strings.
        @type xs_p3: ('', '') | (float, float)
            x-values of boundaries of phase 3. Default are empty strings. 
        @rtype: None
        '''
        self.kind = kind # None, 'obj', or 'subj'
        self.obj_num_pts = obj_num_pts # None if not obj regression
        self.run = run
        self.xs_p3, self.xs_p2, self.xs_p1  = xs_p3, xs_p2, xs_p1

        # Default values are None unless assigned
        blank_phase = Phase(
            ('',''), ('',''), ('',''), '', '', '', [], [], '', '' ,'' ,'')
        self.phase3, self.phase2, self.phase1 = blank_phase, blank_phase, blank_phase
        self.r2s = None

        self.obj_x_start, self.obj_y_start = None, None
        # Attributes for testing
        self.r2s = None # Lists from obj analysis, y=mx+b
        self.p12_r2_max = None

        self.x_p12, self.y_p12 = None, None
        self.x_p12_curvestrip_p3, self.y_p12_curvestrip_p3 = None, None
        self.x_p1_curvestrip_p3, self.y_p1_curvestrip_p3 = None, None
        self.x_p1_curvestrip_p23, self.y_p1_curvestrip_p23 = None, None

        self.elut_period, self.tracer_retained, self.poolsize = None, None, None
        self.influx, self.netflux, self.ratio = None, None, None

    def analyze(self):
        '''Implement analysis based on settings from attributes.

        Parameters are defined based on the phase limits that have been provided
        The 'engine' of the analysis if you will.

        @type self: Analysis
        @rtype: None
        '''
        # Implement objective analysis. Note that objective analysis just uses a
        #    set algorithm to set phase limits. After this if block the process
        #    is the same of both objective and subjective analyses. A subjective
        #    just allows the user to directly set the phase limits.
        if self.kind == 'obj':
            self.obj_x_start = self.run.x[-self.obj_num_pts:]
            self.obj_y_start = self.run.y[-self.obj_num_pts:]
            self.xs_p3, self.r2s, self.ms, self.bs = Operations.get_obj_phase3(
                obj_num_pts=self.obj_num_pts,
                elut_ends_parsed=self.run.elut_ends_parsed,
                elut_cpms_log=self.run.elut_cpms_log)
            self.xs_p2, self.xs_p1, self.p12_r2_max = Operations.get_obj_phase12(
                xs_p3=self.xs_p3, elut_ends_parsed=self.run.elut_ends_parsed,
                elut_cpms_log=self.run.elut_cpms_log,
                elut_ends=self.run.elut_ends)
        # From here analysis is same for both objective and subjective analyses
        if self.xs_p3 != ('', ''):
            self.phase3 = Operations.extract_phase(
                xs=self.xs_p3, 
                x_series=self.run.elut_ends_parsed, 
                y_series=self.run.elut_cpms_log,
                elut_ends=self.run.elut_ends,
                SA=self.run.SA, load_time=self.run.load_time)
            Operations.advanced_run_calcs(analysis=self)
            if self.xs_p2 != ('', '') and self.phase3.xs!= ('', ''):
                # Set series' to be curve-stripped
                end_p12_index = Operations.x_to_index(
                    x_value=self.xs_p2[1], index_type='end',
                    x_series=self.run.elut_ends_parsed,
                    larger_x=self.run.elut_ends)
                self.x_p12 = self.run.x[: end_p12_index+1]
                self.y_p12 = self.run.y[: end_p12_index+1]
                # Curve strip phase 1 + 2 data of phase 3
                # From here on data series potentially have 'holes' from
                # ommitting negative log operations during curvestripping
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
                    # Getting phase 1 data that has been already stripped of
                    # phase 3 data
                    self.x_p1_curvestrip_p3 =\
                        self.x_p12_curvestrip_p3[start_p1_index : end_p1_index+1]
                    self.y_p1_curvestrip_p3 =\
                        self.y_p12_curvestrip_p3[start_p1_index : end_p1_index+1]
                    # Curve-strip phase 2 data from phase 1
                    self.x_p1_curvestrip_p23, self.y_p1_curvestrip_p23 = \
                        Operations.curvestrip(
                            x_series=self.x_p1_curvestrip_p3,
                            y_series=self.y_p1_curvestrip_p3, 
                            slope=self.phase2.slope,
                            intercept=self.phase2.intercept)
                    self.phase1 = Operations.extract_phase(
                        xs=self.xs_p1, 
                        x_series=self.x_p1_curvestrip_p23,
                        y_series=self.y_p1_curvestrip_p23,
                        elut_ends=self.run.elut_ends,
                        SA=self.run.SA, load_time=self.run.load_time)


class Run(object):
    '''Basic data form a single CATE run/replicate

    === Attributes ===
    @type name: str
        Name of run, used for identification purposes
    @type SA: float
        The specific activity of the loading solution used for the CATE
            run/replicate; radiactivity(counts per min)/volume (mL).
    @type rt_cnts: float
        Radioactivity of the roots of the plant used (cpm).
    @type sht_cnts: float
        Radioactivity of the shoots of the plant used (cpm).
    @type rt_wght: float
        Weight of the roots of the plant used (grams).
    @type gfact: float
        Instrument-specific correction factor. Used to account of idiosyncrasies
            of the detecting equiptment.
    @type load_time: float
        Amount of time plant used was placed in loading solution for (minutes).
    @type elut_ends: list[float]
        Time points that eluates were removed from plants (vs added to plants)
    @type raw_cpms: list[float]
        Eluate radioactivities as measured by detecting equitpment.
    @type elut_cpms_gfact: list[float]
        elut_cpms corrected for (multiplied by) g_fact
    @type elut_cpms_gRFW: list[float]
        elut_cpms_gfact corrected for (divided by) g_fact
    @type elut_cpms_log: list[float]
        Logarithmic conversion of elut_cpms_gRFW
    @type elut_ends_parsed: list[float]
        elut_ends with time points corresponding to eluates with no 
            radioactivity removed.
    '''        
    def __init__(
    	   self, name, SA, rt_cnts, sht_cnts, rt_wght, gfact,
    	   load_time, elut_ends, raw_cpms, elut_cpms):
        '''Constructor of Run objects

        @type self: Run
        @type name: str
            Name of run, used for identification purposes
        @type SA: float
            The specific activity of the loading solution used for the CATE
                run/replicate; radiactivity(counts per min)/volume (mL).
        @type rt_cnts: float
            Radioactivity of the roots of the plant used (cpm).
        @type sht_cnts: float
            Radioactivity of the shoots of the plant used (cpm).
        @type rt_wght: float
            Weight of the roots of the plant used (grams).
        @type gfact: float
            Instrument-specific correction factor. Used to account of idiosyncrasies
                of the detecting equiptment.
        @type load_time: float
            Amount of time plant used was placed in loading solution for (minutes).
        @type elut_ends: list[float]
            Time points that eluates were removed from plants (vs added to plants)
        @type raw_cpms: list[float]
            Eluate radioactivities as measured by detecting equitpment.
        @type elut_cpms: list[float]
            raw_cpms with blank values ('') replaced with 0s
        @type x: ndarry
            elut_ends_parsed converted to a numpy compatable array
        @type y: ndarry
            elut_cpms_log converted to a numpy compatable array
        @rtype: None
        '''       
        self.name = name # Text identifier extracted from col header in excel
        self.SA = SA
        self.rt_cnts = rt_cnts
        self.sht_cnts = sht_cnts
        self.rt_wght = rt_wght
        self.gfact = gfact
        self.load_time = load_time
        self.elut_ends = elut_ends
        self.raw_cpms = raw_cpms
        self.elut_cpms = elut_cpms # = raw_cpms with blanks('') replaced w/0        
        self.elut_starts = [0.0] + elut_ends[:-1]
        self.elut_cpms_gfact, self.elut_cpms_gRFW, \
            self.elut_cpms_log, self.elut_ends_parsed = \
            Operations.basic_run_calcs(
                rt_wght, gfact, self.elut_starts, elut_ends, elut_cpms)       
		# x and y data for graphing ('numpy-fied')
        self.x = numpy.array(self.elut_ends_parsed)
        self.y = numpy.array(self.elut_cpms_log)


class Phase(object):
    ''' Data for a particular phase in our Analysis
    
    === Attributes ===
    @type xs: (float, float) | ('', '')
        x-values (from elut_ends_parsed) of boundaries of phase.
    @type xy1: (float, float) | ('', '')
        x, y coordinates for one end of regression line.
        Used to plot phase in GUI.
    @type xy2: (float, float) | ('', '')
        x, y coordinates for other end of regression line.
        Used to plot phase in GUI.
    @type r2: float | ''
        Coefficient of correlation (R^2) between <x_series> and <y_series>.
    @type slope: float | ''
        Slope of regression line between <x_series> and <y_series>.
    @type intercept: float | ''
        Intercept of regression line between <x_series> and <y_series>.
    @type x_series: [float]
        x_series of data. Generally elut_ends_parsed.
    @type x_series: [float]
        y_series of data. elut_cpms_log data is curve-stripped forms.
    @type k: float | ''
        Rate constant of the phase (slope *2.303).
    @type t05: float | ''
        Half-life of exchange of the phase (0.693/k).
    @type r0: float | ''
        Rate of radioisotope release from compartment at time = 0 (antilog of
            intercept).
    @type efflux: float | ''
        Efflux from compartment (r0/SA).

    === Representation Invariants ===
    - Default values of attributes in blank phase are empty strings
    '''
    def __init__(
        self, xs, xy1, xy2, r2, slope, intercept, x_series, y_series,
        k, t05, r0, efflux):
        ''' Constructor of Phase object.

        @type self: Phase
        @type xs: (float, float) | ('', '')
            x-values (from elut_ends_parsed) of boundaries of phase.
        @type xy1: (float, float) | ('', '')
            x, y coordinates for one end of regression line.
            Used to plot phase in GUI.
        @type xy2: (float, float) | ('', '')
            x, y coordinates for other end of regression line.
            Used to plot phase in GUI.
        @type r2: float | ''
            Coefficient of correlation (R^2) between <x_series> and <y_series>.
        @type slope: float | ''
            Slope of regression line between <x_series> and <y_series>.
        @type intercept: float | ''
            Intercept of regression line between <x_series> and <y_series>.
        @type x_series: [float]
            x_series of data. Generally elut_ends_parsed.
        @type x_series: [float]
            y_series of data. elut_cpms_log data is curve-stripped forms.
        @type k: float
            Rate constant of the phase (slope *2.303).
        @type t05: float
            Half-life of exchange of the phase (0.693/k).
        @type r0: float
            Rate of radioisotope release from compartment at time = 0 (antilog of
                intercept).
        @type efflux: float
            Efflux from compartment (r0/SA).
        @rtype: None
        '''
        self.xs = xs # paired tuple (x, y)
        self.xy1, self.xy2 = xy1, xy2 # Each is a paired tuple
        self.r2, self.slope, self.intercept = r2, slope, intercept
        self.x_series, self.y_series = x_series, y_series
        self.k, self.t05, self.r0, self.efflux = k, t05, r0, efflux


    def blank_phase(self):
        '''Creating a blank phase

        @type self: Phase
        @rtype: Phase
        '''
        return Phase(
            ('',''), ('',''), ('',''), '', '', '', [], [], '', '' ,'' ,'')


if __name__ == "__main__":
    import Excel
    import os
    directory = os.path.dirname(os.path.abspath(__file__))
    temp_data = Excel.grab_data(directory, "/Tests/1/Test_SingleRun1.xlsx")
    
    temp_analysis = temp_data.analyses[0]
    temp_analysis.kind = 'obj'
    temp_analysis.obj_num_pts = 4
    temp_analysis.analyze()