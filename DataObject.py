class DataObject():
    '''
    Class that stores overarching metadata for ENTIRE analysis
    '''
        
    def __init__(self, output_folder, run_objects):
        
        # Directory any file is to be saved in
        self.output_folder = output_folder
        # List of RunObjects
        self.run_objects = run_objects
        