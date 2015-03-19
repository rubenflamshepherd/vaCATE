class DataObject():
    '''
    Class that stores overarching metadata for ENTIRE analysis
    Two Attributes:
    output_folder - directory that analysis is started from and where output
                    files (data analysis) are to be saved
    run_objects - list of run_objects (see RunObject module)
    
    '''
        
    def __init__(self, output_folder, run_objects):
        
        # Directory any file is to be saved in
        self.output_folder = output_folder
        # List of RunObjects
        self.run_objects = run_objects
        