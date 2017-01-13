from matplotlib.backends.backend_wxagg import \
    NavigationToolbar2WxAgg
import wx
from matplotlib.backends.backend_wx import _load_bitmap

import Excel


class Toolbar(NavigationToolbar2WxAgg):
    """
    Extend the default wx toolbar with your own event handlers
    """
    # Ids for buttons being added to toolbar.
    ON_PREVIOUS = wx.NewId()
    ON_NEXT = wx.NewId()
    ON_EXTRACT = wx.NewId()
        
    def __init__(self, frame_object):
        """ Constructor for toolbar object

        @type self: Toolbar
        @type frame_object: MainFrame
            the frame object that the toolbar will be in/part of
        @rtype: None
        """
        NavigationToolbar2WxAgg.__init__(self, frame_object.canvas)
        self.frame_object = frame_object        
        # Deleting unwanted icons in standard toolbar
        self.DeleteToolByPos(8)
        self.DeleteToolByPos(1)
        self.DeleteToolByPos(1)
        
        self.InsertSeparator(6)
        self.InsertSeparator(6)

        self.AddSimpleTool(self.ON_PREVIOUS, _load_bitmap('back.png'),
                           'Previous Run', 'Activate custom control')
        wx.EVT_TOOL(self, self.ON_PREVIOUS, self._on_previous)
        
        self.AddSimpleTool(self.ON_NEXT, _load_bitmap('forward.png'),
                           'Next Run', 'Activate custom control')
        wx.EVT_TOOL(self, self.ON_NEXT, self._on_next)        
        
        self.AddSimpleTool(self.ON_EXTRACT, _load_bitmap('filesave.png'),
                           'Save to Excel', 'Activate custom control')
        wx.EVT_TOOL(self, self.ON_EXTRACT, self._on_extract)
    
    def _on_next(self, evt):
        """ Action governing what happens when we press the 'right arrow' icon

        Next analysis is displayed.

        @type self: Toolbar
        @type evt: Event
        @rtyp: None
        """
        self.frame_object.analysis_num += 1
        self.frame_object.draw_figure()
        
    def _on_previous(self, evt):
        """ Action governing what happens when we press the 'left arrow' icon

        Previous analysis is displayed.

        @type self: Toolbar
        @type evt: Event
        @rtype: None
        """
        self.frame_object.analysis_num -= 1
        self.frame_object.draw_figure()    

    def _on_extract(self, evt):
        """ Action governing what happens when we press the 'floppy disc' icon

        Current set of analyses are saved to an excel file.

        @type self: Toolbar
        @type evt: Event
        @rtype: None
        """
        Excel.generate_analysis(self.frame_object.experiment)
        evt.Skip()
        self.frame_object.Destroy()
