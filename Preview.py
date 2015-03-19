"""
This demo demonstrates how to embed a matplotlib (mpl) plot 
into a wxPython GUI application, including:

* Using the navigation toolbar
* Adding data to the plot
* Dynamically modifying the plot's properties
* Processing mpl events
* Saving the plot to a file from a menu

The main goal is to serve as a basis for developing rich wx GUI
applications featuring mpl plots (using the mpl OO API).

Eli Bendersky (eliben@gmail.com)
License: this code is in the public domain
Last modified: 30.07.2008
"""

# The recommended way to use wx with mpl is with the WXAgg
# backend. 
#
import RunObject

import os
import pprint
import random
import math
import wx
import numpy as np
import Operations
import matplotlib
matplotlib.use('WXAgg')
from matplotlib.figure import Figure
from matplotlib.backends.backend_wxagg import \
    FigureCanvasWxAgg as FigCanvas, \
    NavigationToolbar2WxAgg
from matplotlib import gridspec
from matplotlib import pyplot as plt
from matplotlib.backends.backend_wx import _load_bitmap
import Custom
	
class MainFrame(wx.Frame):
    """ The main frame of the application
    """
    title = 'Compartmental Analysis of Tracer Efflux: Data Analyzer'
    
    def __init__(self, data_object):
        wx.Frame.__init__(self, None, -1, self.title)
	
	self.SetIcon(wx.Icon('Images/testtube.ico', wx.BITMAP_TYPE_ICO))
	
	self.run_num = 0
	
	self.data_object = data_object
	self.create_main_panel()
        
        # Default analysis:objective regression using the last 2 data points
        self.obj_textbox.SetValue ('3')
	        
        self.draw_figure()
	
    def create_main_panel(self):
        """ Creates the main panel with all the controls on it:
             * mpl canvas 
             * mpl navigation toolbar
             * Control panel for interaction
        """
        self.panel = wx.Panel(self)
        
        # Create the mpl Figure and FigCanvas objects. 
        # 5x4 inches, 100 dots-per-inch
        #
        self.dpi = 100
        self.fig = Figure((10, 4.0), dpi=self.dpi)
        self.canvas = FigCanvas(self.panel, -1, self.fig)
       
        # Since we have only one plot, we can use add_axes 
        # instead of add_subplot, but then the subplot
        # configuration tool in the navigation toolbar wouldn't
        # work.
        
        # Allows us to set custom sizes of subplots
        gs = gridspec.GridSpec(2, 2)
        self.plot_phase3 = self.fig.add_subplot(gs[1:-1, 0])
        self.plot_phase2 = self.fig.add_subplot(gs[1, 1]) 
        self.plot_phase1 = self.fig.add_subplot(gs[0, 1]) 
                
        # Bind the 'pick' event for clicking on one of the bars
        self.canvas.mpl_connect('pick_event', self.on_pick_p3)
	# self.canvas.mpl_connect('pick_event', self.on_pick_p2)
        
        # Bind the 'check' event for ticking 'Show Grid' option
        self.cb_grid = wx.CheckBox(self.panel, -1, 
            "Show Grid",
            style=wx.ALIGN_CENTER)
        self.Bind(wx.EVT_CHECKBOX, self.on_cb_grid, self.cb_grid)

        # Create the navigation toolbar, tied to the canvas
        self.toolbar = Custom.Toolbar(self)
	self.toolbar.Realize()
	
        # Layout with box sizers                
        self.vbox = wx.BoxSizer(wx.VERTICAL)
        self.vbox.Add(self.canvas, 1, wx.LEFT | wx.TOP | wx.GROW)
        self.vbox.Add(self.toolbar, 0, wx.EXPAND)
        self.vbox.AddSpacer(10)
                
        # Text labels describing regression inputs
        self.obj_title = wx.StaticText (
	    self.panel,
	    label="Objective Regression",
	    style=wx.ALIGN_CENTER)
        self.obj_label = wx.StaticText (
	    self.panel,
	    label="Number of points to use:",
	    style=wx.ALIGN_CENTER)
        self.subj_title = wx.StaticText (
	    self.panel,
	    label="Subjective Regression",
	    style=wx.ALIGN_CENTER)
        self.subj1_label = wx.StaticText (
	    self.panel,
	    label="First Point:",
	    style=wx.ALIGN_CENTER)
        self.subj2_label = wx.StaticText (
	    self.panel,
	    label="Last Point:",
	    style=wx.ALIGN_CENTER)
        self.subj_disclaimer = wx.StaticText (
	    self.panel,
	    label="Points are numbered from left to right",
	    style=wx.ALIGN_CENTER)
        self.linedata_title = wx.StaticText (
	    self.panel,
	    label="Regression Parameters",
	    style=wx.ALIGN_CENTER)
        
        # Text boxs for collecting subj/obj regression parameter input
        self.obj_textbox = wx.TextCtrl(self.panel, 
	    size=(50,-1),
	    style=wx.TE_PROCESS_ENTER)
        
        self.subj_p1_1textbox = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_PROCESS_ENTER)        
	
	self.subj_p1_2textbox = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_PROCESS_ENTER)
        
        self.subj_p2_1textbox = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_PROCESS_ENTER)        
        self.subj_p2_2textbox = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_PROCESS_ENTER)
        
        self.subj_p3_1textbox = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_PROCESS_ENTER)        
        self.subj_p3_2textbox = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_PROCESS_ENTER)        
                    
        
        # Buttons for identifying collect regression parameter event
        self.obj_drawbutton = wx.Button(self.panel, -1,
	                                "Draw Objective Regresssion")
        self.Bind(wx.EVT_BUTTON, self.on_obj_draw, self.obj_drawbutton)
        
        self.subj_drawbutton = wx.Button(self.panel, -1,
	                                 "Draw Subjective Regression")
        self.Bind(wx.EVT_BUTTON, self.on_subj_draw, self.subj_drawbutton)
        self.line = wx.StaticLine(self.panel, -1, style=wx.LI_VERTICAL)
        self.line2 = wx.StaticLine(self.panel, -1, style=wx.LI_VERTICAL)
        self.line3 = wx.StaticLine(self.panel, -1, style=wx.LI_VERTICAL)
	self.line4 = wx.StaticLine(self.panel, -1, style=wx.LI_HORIZONTAL)
        
        # Alignment flags (for adding things to spacers) and fonts
        flags = wx.ALIGN_RIGHT | wx.ALL | wx.ALIGN_CENTER_VERTICAL
        box_flag = wx.ALIGN_CENTER | wx.ALL | wx.ALIGN_CENTER_VERTICAL
        button_flag = wx.ALIGN_BOTTOM
        title_font = wx.Font (8, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        widget_title_font = wx.Font (8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, True)
	disclaimer_font = wx.Font (6, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        
        # Setting fonts
        self.obj_title.SetFont(title_font)
        self.subj_title.SetFont(title_font)
        self.linedata_title.SetFont(title_font)
        self.subj_disclaimer.SetFont(disclaimer_font)
        
        # Adding objective widgets to objective sizer
        self.vbox_obj = wx.BoxSizer(wx.VERTICAL)
        self.vbox_obj.Add (self.obj_title, 0, border=3, flag=box_flag)
        self.vbox_obj.AddSpacer(7)
        self.vbox_obj.Add (self.obj_label, 0, border=3, flag=box_flag)
        self.vbox_obj.AddSpacer(5)
        self.vbox_obj.Add (self.obj_textbox, 0, border=3, flag=box_flag)
        self.vbox_obj.AddSpacer(5)
        self.vbox_obj.Add (self.obj_drawbutton, 0, border=3, flag=button_flag)
        
        # Adding subjective widgets to subjective sizer
        self.gridbox_subj = wx.GridSizer (rows=4, cols=3, hgap=1, vgap=1)
        self.gridbox_subj.Add (wx.StaticText (self.panel, id=-1, label=""),
	                       0, border=3, flag=flags)
        self.gridbox_subj.Add (self.subj1_label, 0, border=3,
	                       flag=box_flag)
        self.gridbox_subj.Add (self.subj2_label, 0, border=3,
	                       flag=box_flag)
                
        self.gridbox_subj.Add (wx.StaticText (self.panel, label="Phase I:"),
	                       0, border=3, flag=flags)        
        self.gridbox_subj.Add (self.subj_p1_1textbox, 0, border=3,
	                       flag=box_flag)
        self.gridbox_subj.Add (self.subj_p1_2textbox, 0, border=3,
	                       flag=box_flag)
                
        self.gridbox_subj.Add (wx.StaticText (self.panel, label="Phase II:"),
	                       0, border=3, flag=flags)        
        self.gridbox_subj.Add (self.subj_p2_1textbox, 0, border=3,
	                       flag=box_flag)
        self.gridbox_subj.Add (self.subj_p2_2textbox, 0, border=3,
	                       flag=box_flag)
                
        self.gridbox_subj.Add (wx.StaticText (self.panel, label="Phase III:"),
	                       0, border=3, flag=flags)        
        self.gridbox_subj.Add (self.subj_p3_1textbox, 0, border=3,
	                       flag=box_flag)
        self.gridbox_subj.Add (self.subj_p3_2textbox, 0, border=3,
	                       flag=box_flag)
                
        self.vbox_subj = wx.BoxSizer(wx.VERTICAL)
        self.vbox_subj.Add (self.subj_title, 0, border=3, flag=box_flag)
        self.vbox_subj.Add (self.gridbox_subj, 0, border=3, flag=box_flag)
        self.vbox_subj.Add (self.subj_drawbutton, 0, border=3,
	                    flag=box_flag)
        self.vbox_subj.Add (self.subj_disclaimer, 0, border=3,
	                    flag=box_flag)
        
        # Creating widgets for data output
        self.data_p1_int = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)
        self.data_p1_slope = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p1_r2 = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p1_k = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p1_t05 = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p1_efflux = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)	
        
        self.data_p2_int = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)
        self.data_p2_slope = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p2_r2 = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p2_k = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p2_t05 = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
	self.data_p2_efflux = wx.TextCtrl(
                self.panel, 
                size=(50,-1),
                style=wx.TE_READONLY)		
        
        self.data_p3_int = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)
        self.data_p3_slope = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p3_r2 = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p3_k = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
        self.data_p3_t05 = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
	self.data_p3_efflux = wx.TextCtrl(
                self.panel, 
                size=(50,-1),
                style=wx.TE_READONLY)
	
	self.data_SA = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)	
	self.data_shootcnts = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)	
	self.data_rootcnts = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)	
	self.data_rootweight = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY) 
	self.data_loadtime = wx.TextCtrl(self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)	
	self.data_influx = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)	
	self.data_netflux = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)
	self.data_ratio = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)			
	self.data_poolsize = wx.TextCtrl(
            self.panel, 
            size=(50,-1),
            style=wx.TE_READONLY)				
	
	# Creating labels for data output
	slope_text = wx.StaticText (self.panel, label="Slope")
	intercept_text = wx.StaticText(self.panel, label = "Intercept")
	r2_text = wx.StaticText (self.panel, label = u"R\u00B2")
	p1_text = wx.StaticText (self.panel, label = "Phase I: ")
	p2_text = wx.StaticText (self.panel, label = "Phase II: ")
	p3_text = wx.StaticText (self.panel, label = "Phase III: ")
	k_text = wx.StaticText (self.panel, label = "k")
	halflife_text = wx.StaticText (self.panel, label = "Half Life")
	efflux_text = wx.StaticText (self.panel, label = "Efflux")
	
	# Labels for gridsizer2
	SA_text = wx.StaticText (self.panel, label = "SA")
	rootcnts_text = wx.StaticText (self.panel, label = "Rt Cnts")
	shootcnts_text = wx.StaticText (self.panel, label = "Sh Cnts")
	rootweight_text = wx.StaticText (self.panel, label = "Rt Weight")
	loadtime_text = wx.StaticText (self.panel, label = "Load Time")
	influx_text = wx.StaticText (self.panel, label = "Influx")
	netflux_text = wx.StaticText (self.panel, label = "Net Flux")
	ratio_text = wx.StaticText (self.panel, label = "E:I Ratio")
	poolsize_text = wx.StaticText (self.panel, label = "Pool size")
	
        # Adding data output widgets to data output gridsizers        
        self.gridbox_data = wx.GridSizer (rows=4, cols=7, hgap=1, vgap=1)
        
        self.gridbox_data.Add (wx.StaticText (self.panel, label = ""),\
	                       0, border=3, flag=flags)
        self.gridbox_data.Add (slope_text, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (intercept_text, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (r2_text, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (k_text, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (halflife_text, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (efflux_text, 0, border=3, flag=box_flag)

        self.gridbox_data.Add (p1_text, 0, border=3, flag=flags)
        self.gridbox_data.Add (self.data_p1_slope, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (self.data_p1_int, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (self.data_p1_r2, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (self.data_p1_k, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (self.data_p1_t05, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (self.data_p1_efflux, 0, border=3, flag=box_flag)
        
        self.gridbox_data.Add (p2_text, 0, border=3, flag=flags)
        self.gridbox_data.Add (self.data_p2_slope, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (self.data_p2_int, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (self.data_p2_r2, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (self.data_p2_k, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (self.data_p2_t05, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (self.data_p2_efflux, 0, border=3, flag=box_flag)
        
        self.gridbox_data.Add (p3_text, 0, border=3, flag=flags)
        self.gridbox_data.Add (self.data_p3_slope, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (self.data_p3_int, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (self.data_p3_r2, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (self.data_p3_k, 0, border=3, flag=box_flag)
	self.gridbox_data.Add (self.data_p3_t05, 0, border=3, flag=box_flag)
        self.gridbox_data.Add (self.data_p3_efflux, 0, border=3, flag=box_flag)
	
	self.gridbox_data2 = wx.GridSizer (rows=2, cols=9, hgap=1, vgap=1)
	
	self.gridbox_data2.Add (SA_text, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (shootcnts_text, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (rootcnts_text, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (rootweight_text, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (loadtime_text, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (influx_text, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (netflux_text, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (ratio_text, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (poolsize_text, 0, border=3, flag=box_flag)
	
	self.gridbox_data2.Add (self.data_SA, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (self.data_shootcnts, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (self.data_rootcnts, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (self.data_rootweight, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (self.data_loadtime, 0, border=3, flag=box_flag)	
	self.gridbox_data2.Add (self.data_influx, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (self.data_netflux, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (self.data_ratio, 0, border=3, flag=box_flag)
	self.gridbox_data2.Add (self.data_poolsize, 0, border=3, flag=box_flag)
	
	self.vbox_linedata = wx.BoxSizer(wx.VERTICAL)
        self.vbox_linedata.Add (self.linedata_title, 0, border=3, flag=box_flag)
        self.vbox_linedata.Add (self.gridbox_data, 0, border=3, flag=box_flag)
	self.vbox_linedata.Add (self.line4, 0, wx.CENTER|wx.EXPAND)
	self.vbox_linedata.Add (self.gridbox_data2, 0, border=3, flag=box_flag)
        
        # Build the widgets and the sizer contain(s) containing them all
	
	# Build slider to adjust point radius
        self.slider_label = wx.StaticText(
	    self.panel,
	    -1,
	    "Point Radius",
	    style=wx.ALIGN_CENTER)
        self.slider_label.SetFont (widget_title_font)
        self.slider_width = wx.Slider(self.panel, -1, 
            value=40, 
            minValue=1,
            maxValue=200,
            style=wx.SL_AUTOTICKS | wx.SL_LABELS)
        self.slider_width.SetTickFreq(10, 1)
        self.Bind(
	    wx.EVT_COMMAND_SCROLL_THUMBTRACK,
	    self.on_slider_width,
	    self.slider_width)        
        self.vbox_widgets = wx.BoxSizer(wx.VERTICAL)
        self.vbox_widgets.Add(self.cb_grid, 0, border=3, flag=box_flag)
        self.vbox_widgets.Add(self.slider_label, 0, flag=box_flag)
        self.vbox_widgets.Add(self.slider_width, 0, border=3, flag=box_flag)
        
        # Build widget that displays information about last widget clicked
	
	# Creating the 'last clicked' items
        self.xy_clicked_label = wx.StaticText(
	    self.panel,
	    -1,
	    "Last point clicked",
	    style=wx.ALIGN_CENTER)
        self.xy_clicked_label.SetFont (widget_title_font)
        self.x_clicked_label = wx.StaticText(
	    self.panel,
	    -1,
	    "Elution Time (x): ",
	    style=wx.ALIGN_CENTER)
        self.y_clicked_label = wx.StaticText(
	    self.panel,
	    -1,
	    "Log cpm (y): ",
	    style=wx.ALIGN_CENTER)
        self.num_clicked_label = wx.StaticText(
	    self.panel,
	    -1,
	    "Point Number: ",
	    style=wx.ALIGN_CENTER)
        self.x_clicked_data = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)
        self.y_clicked_data = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)
        self.num_clicked_data = wx.TextCtrl(
	    self.panel, 
	    size=(50,-1),
	    style=wx.TE_READONLY)        
        
        # Assembling the 'last clicked' items into sizers
        self.hbox_x_clicked = wx.BoxSizer(wx.HORIZONTAL)
        self.hbox_x_clicked.AddSpacer(10)
        self.hbox_x_clicked.Add (self.x_clicked_label, 0, flag=flags)
        self.hbox_x_clicked.Add (self.x_clicked_data, 0, flag=flags)
        self.hbox_y_clicked = wx.BoxSizer(wx.HORIZONTAL)
        self.hbox_y_clicked.AddSpacer(10)
        self.hbox_y_clicked.Add (self.y_clicked_label, 0, flag=flags)
        self.hbox_y_clicked.Add (self.y_clicked_data, 0, flag=flags)
        self.hbox_num_clicked = wx.BoxSizer(wx.HORIZONTAL)
        self.hbox_num_clicked.AddSpacer(10)
        self.hbox_num_clicked.Add (self.num_clicked_label, 0, flag=flags)
        self.hbox_num_clicked.Add (self.num_clicked_data, 0, flag=flags)        
        
        self.vbox_widgets.Add(self.xy_clicked_label, 0, flag=box_flag)
        self.vbox_widgets.Add(self.hbox_x_clicked, 0, flag=flags)
        self.vbox_widgets.Add(self.hbox_y_clicked, 0, flag=flags)
        self.vbox_widgets.Add(self.hbox_num_clicked, 0, flag=flags)
        
        # Assembling items for subjective regression field into sizers
        self.hbox_regres = wx.BoxSizer(wx.HORIZONTAL)        
        self.hbox_regres.Add (self.vbox_obj)
        self.hbox_regres.AddSpacer(5)
        self.hbox_regres.Add (self.line, 0, wx.CENTER|wx.EXPAND)
        self.hbox_regres.AddSpacer(5)
        self.hbox_regres.Add (self.vbox_subj)
        self.hbox_regres.Add (self.line2, 0, wx.CENTER|wx.EXPAND)
        self.hbox_regres.AddSpacer(5)
        self.hbox_regres.Add (self.vbox_linedata)
        self.hbox_regres.Add (self.line3, 0, wx.CENTER|wx.EXPAND)        
        self.hbox_regres.Add (self.vbox_widgets, 0, wx.CENTER|wx.EXPAND)
        
        self.vbox.Add(self.hbox_regres, 0, flag = wx.ALIGN_CENTER | wx.TOP)
        
        self.panel.SetSizer(self.vbox)
        self.vbox.Fit(self)
    
    def create_status_bar(self):
        self.statusbar = self.CreateStatusBar()
	
    def draw_figure(self):
        """ Redraws the figure
        """
	
	run_object = self.data_object.run_objects [self.run_num]
		
	self.obj_textbox.SetValue (str (run_object.pts_used))	
	
	title_string = 'Compartmental Analysis of Tracer Efflux: Data Analyzer - '
	detail_string = "Run " + str (self.run_num + 1) + "/"\
	    + str(len (self.data_object.run_objects))
	name_string = ' - "'+ run_object.run_name + '"'
	self.SetTitle (title_string + detail_string + name_string)
	
	# Making sure toolbar buttons navigate to runs that exist
	if self.run_num == 0:
	    self.toolbar.EnableTool (self.toolbar.ON_PREVIOUS, False)
	else:
	    self.toolbar.EnableTool (self.toolbar.ON_PREVIOUS, True)
	
	if self.run_num == len(self.data_object.run_objects) - 1:
	    self.toolbar.EnableTool (self.toolbar.ON_NEXT, False)
	else:
	    self.toolbar.EnableTool (self.toolbar.ON_NEXT, True)
	    
	#self.toolbar.Realize()		
		       
        # Clearing the plots so they can be redrawn anew
        self.plot_phase1.clear()
        self.plot_phase2.clear()
        self.plot_phase3.clear()        

	self.plot_phase1.grid(self.cb_grid.IsChecked())        
	self.plot_phase2.grid(self.cb_grid.IsChecked())        
        self.plot_phase3.grid(self.cb_grid.IsChecked())
	
	# Graphing complete log efflux data set
        self.plot_phase3.scatter(
            run_object.x,
            run_object.y,
            s = self.slider_width.GetValue(),
            alpha = 0.5,
            edgecolors = 'k',
            facecolors = 'w',
            picker = 5
        )
        
        # Setting axes labels/limits
        self.plot_phase3.set_xlabel ('Elution time (min)')
	self.plot_phase2.set_xlabel ('Elution time (min)')
        self.plot_phase3.set_ylabel (u"Log cpm released/g RFW/min")
        self.plot_phase3.set_xlim (left = 0)
        self.plot_phase3.set_ylim (bottom = 0)
                
        # OBJECTIVE REGRESSION
	
	num_points_obj = self.obj_textbox.GetValue ()
	self.num_points_obj = int (num_points_obj)
	if num_points_obj < 3:
	    num_points_obj = 3
	    self.obj_textbox.SetValue ('3')
	
	
		    
	# Graphing the p3 series and regression line
	self.plot_phase3.scatter(
                    run_object.x_p3,
                    run_object.y_p3,
                    s = self.slider_width.GetValue(),
                    alpha = 0.75,
                    edgecolors = 'k',
                    facecolors = 'k'
                )            
	line_p3 = matplotlib.lines.Line2D (
            [run_object.x1_p3, run_object.x2_p3],
            [run_object.y1_p3, run_object.y2_p3],
            color = 'r',
            ls = '-',
            label = 'Phase III')
	self.plot_phase3.add_line (line_p3)
	
	# Distiguishing the intial points used to start the regression
	# and plotting them (solid red)
	self.plot_phase3.scatter(
                    run_object.x_reg_start,
                    run_object.y_reg_start,
                    s = self.slider_width.GetValue(),
                    alpha = 0.5,
                    edgecolors = 'r',
                    facecolors = 'r'
                )	
			    
	# Graphing the p2 series and regression line
	
	# Graphing raw uncorrected data of p1 and p2
	self.plot_phase2.scatter(
                    run_object.x_p12,
                    run_object.y_p12,
                    s = self.slider_width.GetValue(),
                    alpha = 0.50,
                    edgecolors = 'k',
                    facecolors = 'w',
	            picker = 5
                )
	
	# Graphing curve-stripped (corrected) phase I and II data, isolated
	# p2 data and line of best fit
	self.plot_phase2.scatter(
	    run_object.x_p12_curvestrippedof_p3,
	    run_object.y_p12_curvestrippedof_p3,
	    s = self.slider_width.GetValue(),
	    alpha = 0.50,
	    edgecolors = 'r',
	    facecolors = 'w')
	
	self.plot_phase2.scatter(
	    run_object.x_p2_curvestrippedof_p3,
	    run_object.y_p2_curvestrippedof_p3,
	    s = self.slider_width.GetValue(),
	    alpha = 0.75,
	    edgecolors = 'k',
	    facecolors = 'k')	
	
	self.line_p2 = matplotlib.lines.Line2D (
            [run_object.x1_p2, run_object.x2_p2],
            [run_object.y1_p2, run_object.y2_p2],
            color = 'r',
            ls = ':',
            label = 'Phase II')
	self.plot_phase2.add_line (self.line_p2)	
		                
	# Graphing the p1 series and regression line
	
	# Graph raw uncorrected p1 data, p1 data corrected for p3, and p1 data
	# Corrected for both p2 and p3
	self.plot_phase1.scatter(
                    run_object.x_p1,
                    run_object.y_p1,
                    s = self.slider_width.GetValue(),
                    alpha = 0.25,
                    edgecolors = 'k',
                    facecolors = 'w',
	            picker = 5
                )
	
	# Graph p1 data corrected for p3
	self.plot_phase1.scatter(
	    run_object.x_p1_curvestrippedof_p3,
	    run_object.y_p1_curvestrippedof_p3,
	    s = self.slider_width.GetValue(),
	    alpha = 0.25,
	    edgecolors = 'r',
	    facecolors = 'w'
	)	
	
	self.plot_phase1.scatter(
	    run_object.x_p1_curvestrippedof_p23,
	    run_object.y_p1_curvestrippedof_p23,
	    s = self.slider_width.GetValue(),
	    alpha = 0.75,
	    edgecolors = 'k',
	    facecolors = 'k'
	)
	
	self.line_p1 = matplotlib.lines.Line2D (
            [run_object.x1_p1, run_object.x2_p1],
            [run_object.y1_p1, run_object.y2_p1],
            color = 'r',
            ls = '--',
            label = 'Phase I'
	)
	self.plot_phase1.add_line (self.line_p1)          
    
	# Outputting the data from the linear regressions to widgets
	self.data_p1_slope.SetValue ('%0.3f'%(run_object.slope_p1))
	self.data_p1_int.SetValue ('%0.3f'%(run_object.intercept_p1))
	self.data_p1_r2.SetValue ('%0.3f'%(run_object.r2_p1))
	self.data_p1_k.SetValue ('%0.3f'%(run_object.k_p1))
	self.data_p1_t05.SetValue ('%0.3f'%(run_object.t05_p1))
	self.data_p1_efflux.SetValue ('%0.1f'%(run_object.efflux_p1))
	
	self.data_p2_slope.SetValue ('%0.3f'%(run_object.slope_p2))
	self.data_p2_int.SetValue ('%0.3f'%(run_object.intercept_p2))
	self.data_p2_r2.SetValue ('%0.3f'%(run_object.r2_p2))        
	self.data_p2_k.SetValue ('%0.3f'%(run_object.k_p2))
	self.data_p2_t05.SetValue ('%0.3f'%(run_object.t05_p2))
	self.data_p2_efflux.SetValue ('%0.2f'%(run_object.efflux_p2))
	
	self.data_p3_slope.SetValue ('%0.3f'%(run_object.slope_p3))
	self.data_p3_int.SetValue ('%0.3f'%(run_object.intercept_p3))
	self.data_p3_r2.SetValue ('%0.3f'%(run_object.r2_p3))         
	self.data_p3_k.SetValue ('%0.3f'%(run_object.k_p3))
	self.data_p3_t05.SetValue ('%0.3f'%(run_object.t05_p3))
	self.data_p3_efflux.SetValue ('%0.3f'%(run_object.efflux_p3))
	
	self.data_SA.SetValue ('%0.3f'%(run_object.SA))
	self.data_shootcnts.SetValue ('%0.3f'%(run_object.shoot_cnts))
	self.data_rootcnts.SetValue ('%0.3f'%(run_object.root_cnts))
	self.data_rootweight.SetValue ('%0.3f'%(run_object.root_weight))
	self.data_loadtime.SetValue ('%0.3f'%(run_object.load_time))
	self.data_influx.SetValue ('%0.3f'%(run_object.influx))
	self.data_netflux.SetValue ('%0.3f'%(run_object.netflux))
	self.data_ratio.SetValue ('%0.3f'%(run_object.ratio))
	self.data_poolsize.SetValue ('%0.3f'%(run_object.poolsize))
	        
        # Adding our legends
        self.plot_phase1.legend(loc='upper right')
	self.plot_phase2.legend(loc='upper right')
	self.plot_phase3.legend(loc='upper right')
        self.fig.subplots_adjust(bottom = 0.13, left = 0.10)
        self.canvas.draw()
        
    def on_obj_draw (self, event):
        self.subj_p1_1textbox.SetValue ('')
        self.subj_p1_2textbox.SetValue ('')
        self.subj_p2_1textbox.SetValue ('')
        self.subj_p2_2textbox.SetValue ('')
        self.subj_p3_1textbox.SetValue ('')
        self.subj_p3_2textbox.SetValue ('')
	
	old_run_object = self.data_object.run_objects [self.run_num]
	
	new_run_object = RunObject.RunObject (old_run_object.run_name,\
	                                      old_run_object.SA,\
	                                      old_run_object.root_cnts,\
	                                      old_run_object.shoot_cnts,\
	                                      old_run_object.root_weight,\
	                                      old_run_object.g_factor,\
	                                      old_run_object.load_time,\
	                                      old_run_object.elution_times,\
	                                      old_run_object.elution_cpms,\
	                                      int (self.obj_textbox.GetValue ()) )
	
	self.data_object.run_objects [self.run_num] = new_run_object
		
        self.draw_figure()
    
        
    def on_subj_draw (self, event):
        self.obj_textbox.SetValue ('')
        self.draw_figure()    
    
    def on_cb_grid(self, event):
        self.draw_figure()
    
    def on_slider_width(self, event):
        self.draw_figure()
    
    def on_draw_button(self, event):
        self.draw_figure()
    
    def on_pick_p3(self, event):
        # The event received here is of the type
        # matplotlib.backend_bases.PickEvent
        #
        # It carries lots of information, of which we're using
        # only a small amount here.
        # 
        ind = event.ind
	print event.artist
	run_object = self.data_object.run_objects [self.run_num]
	
	
        x_clicked = np.take(run_object.x, ind)
        
        self.x_clicked_data.SetValue ('%0.2f'%(np.take(run_object.x, ind)[0]))
        self.y_clicked_data.SetValue ('%0.3f'%(np.take(run_object.y, ind)[0]))
        self.num_clicked_data.SetValue ('%0.0f'%(ind[0]+1))
	
    def on_exit(self, event):
        self.Destroy()
        
   
if __name__ == '__main__':
    import Excel
    
    temp_data = Excel.grab_data("C:\Users\Ruben\Projects\CATEAnalysis", "CATE Template - Multi Run.xlsx")

    app = wx.PySimpleApp()
    app.frame = MainFrame(temp_data)
    app.frame.Show()
    app.frame.Center()
    app.MainLoop()