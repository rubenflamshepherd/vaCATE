from matplotlib.backends.backend_wxagg import \
    NavigationToolbar2WxAgg
import wx
import time
import xlsxwriter
import Excel
from matplotlib.backends.backend_wx import _load_bitmap

class Toolbar(NavigationToolbar2WxAgg):
    """
    Extend the default wx toolbar with your own event handlers
    
    """
    
    #Ids for buttons being added to toolbar.
    ON_PREVIOUS = wx.NewId()
    ON_NEXT = wx.NewId()
    ON_EXTRACT = wx.NewId()
        
    def __init__(self, canvas, frame_object):
        NavigationToolbar2WxAgg.__init__(self, canvas)
        self.frame_object = frame_object
        
        
        self.DeleteToolByPos(8)
        self.DeleteToolByPos(1)
        self.DeleteToolByPos(1)
        self.AddSeparator ()
        #self.InsertSeparator (6)
        #self.InsertSeparator (6)
                
        # for simplicity I'm going to reuse a bitmap from wx, you'll
        # probably want to add your own.

        self.AddSimpleTool(self.ON_PREVIOUS, _load_bitmap('back.png'),
                           'Previous Run', 'Activate custom contol') 
        wx.EVT_TOOL(self, self.ON_PREVIOUS, self._on_previous)

        self.AddSimpleTool(self.ON_NEXT, _load_bitmap('forward.png'),
                           'Next Run', 'Activate custom contol') 
        wx.EVT_TOOL(self, self.ON_NEXT, self._on_next)        
        
        self.AddSimpleTool(self.ON_EXTRACT, _load_bitmap('filesave.png'),
                           'Save to Excel', 'Activate custom contol')
        wx.EVT_TOOL(self, self.ON_EXTRACT, self._on_extract)
       
    def _on_previous(self, evt):   
        self.frame_object.run_num -= 1
        self.frame_object.draw_figure ()
    
    def _on_next(self, evt):
        self.frame_object.run_num += 1
        self.frame_object.draw_figure ()

    def _on_extract(self, evt):
        # add some text to the axes in a random location in axes (0,1)
        # coords) with a random color

        # get the axes
        ax = self.canvas.figure.axes[0]

        # generate a random location can color
        x,y = (0.5,0.5)
        rgb = (0.50,0.50,0.50)

        # add the text and draw
        ax.text(x, y, 'You clicked me',
                transform=ax.transAxes,
                color=rgb)
        self.canvas.draw()
                
        file_name = 'CATE Finished - ' + time.strftime ("(%Y_%m_%d).xlsx")
        output_file_path = '/'.join ((self.frame_object.directory, file_name))  
        
        workbook = xlsxwriter.Workbook(output_file_path)
        worksheet = workbook.add_worksheet(self.frame_object.run_name)          
        
        Excel.generate_template (output_file_path, workbook, worksheet)
        
        Excel.generate_analysis (workbook, worksheet, self.frame_object)
        
        workbook.close()        
        
        evt.Skip()
        
        self.frame_object.Destroy ()