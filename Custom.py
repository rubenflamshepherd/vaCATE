from matplotlib.backends.backend_wxagg import \
    NavigationToolbar2WxAgg
import wx
from matplotlib.backends.backend_wx import _load_bitmap

class Toolbar(NavigationToolbar2WxAgg):
    """
    Extend the default wx toolbar with your own event handlers
    """
    ON_EXTRACT = wx.NewId()
    def __init__(self, canvas, frame_object):
        NavigationToolbar2WxAgg.__init__(self, canvas)
        self.frame_object = frame_object
        
        POSITION_OF_FORWARD_BTN = 1
        self.DeleteToolByPos(POSITION_OF_FORWARD_BTN)        
        self.DeleteToolByPos(POSITION_OF_FORWARD_BTN)
        self.InsertSeparator (7)
        
        # for simplicity I'm going to reuse a bitmap from wx, you'll
        # probably want to add your own.
        
        self.AddSimpleTool(self.ON_EXTRACT, _load_bitmap('move.ppm'),
                           'Click me', 'Activate custom contol')
        wx.EVT_TOOL(self, self.ON_EXTRACT, self._on_extract)
        

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
        print self.frame_object.directory
        evt.Skip()