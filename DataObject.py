class DataObject():
    '''
    Class that stores overarching metadata for ENTIRE analysis
    '''
        
    def __init__(self, output_folder, raw_run_data_, analysis_type):
        
        # Directory any file is to be saved in
        self.output_folder = output_folder
        # List of runs in form of [run_name, SA, etc]
        self.raw_run_data = raw_run_data
        # List of types of analyses to be done to each run
        # Type of analyses is a type of either
        # (obj, 2) or (subj ((1,2),(3,4),(5,6))
        self.analysis_type = analysis_type