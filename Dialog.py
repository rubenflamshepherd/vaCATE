import wx
import wx.lib.plot as plot
import os,sys
import time

# Custom modules
import Excel
import Preview
import Operations

# The recommended way to use wx with mpl is with the WXAgg
# backend. 
import matplotlib
matplotlib.use('WXAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import \
    FigureCanvasWxAgg as FigCanvas, \
    NavigationToolbar2WxAgg as NavigationToolbar

#-------------------------------------------------------

# Main dialog window presented to the user        
class MainFrame(wx.Frame):
        
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title)

        self.rootPanel = wx.Panel(self)

        innerPanel = wx.Panel(self.rootPanel,-1, size=(500,160), style=wx.ALIGN_CENTER)
        hbox = wx.BoxSizer(wx.HORIZONTAL) 
        vbox = wx.BoxSizer(wx.VERTICAL)
        innerBox = wx.BoxSizer(wx.VERTICAL)
        buttonBox1 = wx.BoxSizer(wx.HORIZONTAL)
        buttonBox2 = wx.BoxSizer(wx.HORIZONTAL)
        
        # Main text presented to user
        txt1 = wx.StaticText(innerPanel, id=-1, label="hello world",style=wx.ALIGN_CENTER, name="")
        innerBox.AddSpacer((150,15))
        innerBox.Add(txt1, 0, wx.CENTER)
        
        hbox.Add(innerPanel, 0, wx.ALL|wx.ALIGN_CENTER)
        vbox.Add(hbox, 1, wx.ALL|wx.ALIGN_CENTER, 5)
        
        self.rootPanel.SetSizer(vbox)
        vbox.Fit(self)
        
               
        data1 = Excel.grab_data("C:\Users\Ruben\Projects\CATEAnalysis", "CATE Analysis - (2014_11_20).xlsx")
        frame = Preview.MainFrame (*data1)
        frame.Show (True)
        frame.MakeModal (True)
               
class MyApp(wx.App):
    def OnInit(self):
        frame = MainFrame(None, -1, 'CATE Data Analyzer')
        frame.SetIcon(wx.Icon('testtube.ico', wx.BITMAP_TYPE_ICO))
        frame.Show(True)
        frame.Center()
        #self.SetTopWindow(frame)
        return True

if __name__ == '__main__':
    app = wx.PySimpleApp()
    app.frame = MainFrame(None, -1, 'CATE Data Analyzer')
    app.frame.Show()
    app.frame.Center()
    app.MainLoop()
    