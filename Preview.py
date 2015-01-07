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

import os
import pprint
import random
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
    
    def __init__(self, run_name, SA, root_cnts, shoot_cnts, root_weight, gfactor,\
                         load_time, elution_times, elution_cpms, directory):
        wx.Frame.__init__(self, None, -1, self.title)
	
	self.SetIcon(wx.Icon('Images/testtube.ico', wx.BITMAP_TYPE_ICO))
	
	# Load initial CATE info as attributes of Frame object
	self.run_name = run_name
	self.SA = SA
	self.root_cnts = root_cnts
	self.shoot_cnts = shoot_cnts
	self.root_weight = root_weight
	self.gfactor = gfactor
	self.load_time = load_time
	self.elution_times = elution_times
	self.elution_ends = elution_times [1:]
	self.elution_starts = elution_times [:len (elution_times) -1]
	self.elution_cpms = elution_cpms
	self.directory = directory
	
	temp = Operations.basic_CATE_analysis (
	    SA, root_cnts, shoot_cnts, root_weight, gfactor, load_time,\
	    elution_times, elution_cpms)
	
	self.elution_cpms_gfactor = temp [0]
	self.elution_cpms_gRFW = temp [1]
	self.elution_cpms_log = temp [2]
	        
        self.x = np.array (self.elution_ends)
	self.y = np.array (self.elution_cpms_log)
        self.create_menu()
        self.create_status_bar()
        self.create_main_panel()
        
        # Default analysis:objective regression using the last 2 data points
        self.obj_textbox.SetValue ('2')        
        
        self.draw_figure()
	
    def create_menu(self):
        """ Creating a file menu that allows saving of graphs to .pngs, opening
        of a help dialog, or quitting the application
        Notes: adapted from original program, pseudo-useless
        """
        
        self.menubar = wx.MenuBar()
        
        menu_file = wx.Menu() # 'File' menu
        
        # 'Save' option of 'File' menu
        m_expt = menu_file.Append(-1, "&Save plot\tCtrl-S", "Save plot to file")
        self.Bind(wx.EVT_MENU, self.on_save_plot, m_expt)
        
        menu_file.AppendSeparator()
        
        # 'Exit' option of 'File' menu
        m_exit = menu_file.Append(-1, "E&xit\tCtrl-X", "Exit")
        self.Bind(wx.EVT_MENU, self.on_exit, m_exit)
        
        menu_help = wx.Menu() #'Help' menu
        
        # 'About' option of 'Help' menu        
        m_about = menu_help.Append(-1, "&About\tF1", "About the demo")
        self.Bind(wx.EVT_MENU, self.on_about, m_about)
        
        # Creating menu
        self.menubar.Append(menu_file, "&File")
        self.menubar.Append(menu_help, "&Help")
        self.SetMenuBar(self.menubar)

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
        self.canvas.mpl_connect('pick_event', self.on_pick)
        
        # Bind the 'check' event for ticking 'Show Grid' option
        self.cb_grid = wx.CheckBox(self.panel, -1, 
            "Show Grid",
            style=wx.ALIGN_CENTER)
        self.Bind(wx.EVT_CHECKBOX, self.on_cb_grid, self.cb_grid)

        # Create the navigation toolbar, tied to the canvas
        self.toolbar = Custom.Toolbar(self.canvas, self)
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
        
        # Alignment flags (for adding things to spacers) and fonts
        flags = wx.ALIGN_RIGHT | wx.ALL | wx.ALIGN_CENTER_VERTICAL
        box_flag = wx.ALIGN_CENTER | wx.ALL | wx.ALIGN_CENTER_VERTICAL
        button_flag = wx.ALIGN_BOTTOM
        title_font = wx.Font (8, wx.DEFAULT, wx.NORMAL, wx.BOLD)
        widget_title_font = wx.Font (8, wx.DEFAULT, wx.NORMAL, wx.NORMAL, True)
	parameter_title_font = wx.Font (14, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, "Times New Roman")
	parameter1_title_font = wx.Font (10, wx.DEFAULT, wx.NORMAL, wx.NORMAL, False, "Times New Roman")
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
	
	# Creating labels for data output
	empty_text = wx.StaticText (self.panel, label = "")
	slope_text = wx.StaticText (self.panel, label="Slope")
	intercept_text = wx.StaticText(self.panel, label = "Intercept")
	r2_text = wx.StaticText (self.panel, label = u"R\u00B2")
	p1_text = wx.StaticText (self.panel, label = "Phase I: ")
	p2_text = wx.StaticText (self.panel, label = "Phase II: ")
	p3_text = wx.StaticText (self.panel, label = "Phase III: ")
	k_text = wx.StaticText (self.panel, label = "k")
	halflife_text = wx.StaticText (self.panel, label = u"t\u2080\u002E\u2085")
	efflux_text = wx.StaticText (self.panel, label = u"\u03d5")
	
	slope_text.SetFont(parameter1_title_font)
	intercept_text.SetFont(parameter1_title_font)
	r2_text.SetFont(parameter1_title_font)
	k_text.SetFont(parameter1_title_font)
	halflife_text.SetFont(parameter_title_font)
        efflux_text.SetFont(parameter_title_font)
	
        # Adding data output widgets to data output gridsizers        
        self.gridbox_data = wx.GridSizer (rows=4, cols=7, hgap=1, vgap=1)
        
        self.gridbox_data.Add (empty_text, 0, border=3, flag=flags)
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
        
        self.vbox_linedata = wx.BoxSizer(wx.VERTICAL)
        self.vbox_linedata.Add (self.linedata_title, 0, border=3, flag=box_flag)
        self.vbox_linedata.Add (self.gridbox_data, 0, border=3, flag=box_flag)
        
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
	
    def obj_analysis (self):
	# Getting parameters from regression of p3
	self.x1_p3, self.x2_p3, self.y1_p3, self.y2_p3, self.r2_p3,\
	    self.slope_p3, self.intercept_p3, self.reg_end_index =\
            Operations.obj_regression_p3 (self.x, self.y, self.num_points_obj)
		
	# Setting the x- and y-series' involved in the p3 linear regression
	self.x_p3 = self.x [self.reg_end_index:] 
	self.y_p3 = self.y [self.reg_end_index:]
	
	# # Setting the x/y-series' used start obj regression
	self.x_reg_start = self.x[len(self.x) - self.num_points_obj:] 
	self.y_reg_start = self.y[len(self.x) - self.num_points_obj:]
	
	# Getting parameters from regression of p1-2!!!!!!!!!!!!!!!!!!!!!!!!!!!!!CLEAN UP
	reg_p1_raw, reg_p2_raw = Operations.obj_regression_p12 (
            self.x,
            self.y,
            self.reg_end_index)
	
	# Unpacking parameters of p1 regression!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!! CLEAN UP
	self.x_p1 = reg_p1_raw [0]
	self.y_p1 = reg_p1_raw [1]
	
	# Getting p1 + p2 curve-stripped data (together)
	self.x_p12_curvestrippedof_p3, self.y_p12_curvestrippedof_p3 = \
	    Operations.p12_curve_stripped (
	        self.x,
	        self.y,
	        self.reg_end_index,
	        self.slope_p3,
	        self.intercept_p3)
	
	# Isolating p2 curve stripped data
	self.x_p2_curvestrippedof_p3 = self.x_p12_curvestrippedof_p3 [len (self.x_p1) :]
	self.y_p2_curvestrippedof_p3 = self.y_p12_curvestrippedof_p3 [len (self.y_p1) :]
	
	# Linear regression of isolated p2 data + getting line data
	self.r2_p2, self.slope_p2, self.intercept_p2 = Operations.linear_regression (
	    self.x_p2_curvestrippedof_p3,
	    self.y_p2_curvestrippedof_p3)
	self.x1_p2, self.x2_p2, self.y1_p2, self.y2_p2 = Operations.grab_x_ys(
	    self.x_p2_curvestrippedof_p3,
	    self.intercept_p2,
	    self.slope_p2)
	
	# Define the p1 x- and y-series corrected for p3
	self.x_p1_curvestrippedof_p3 = self.x_p12_curvestrippedof_p3 [:len (self.x_p1)]
	self.y_p1_curvestrippedof_p3 = self.y_p12_curvestrippedof_p3 [:len (self.y_p1)]
	
	# Getting and plotting p1 curve-stripped data (for p2 and p3)
	self.x_p1_curvestrippedof_p23, self.y_p1_curvestrippedof_p23 = \
	    Operations.p12_curve_stripped (
	        self.x_p1_curvestrippedof_p3,
	        self.y_p1_curvestrippedof_p3,
	        len (self.x_p1_curvestrippedof_p3),
	        self.slope_p2,
	        self.intercept_p2
	    )
	
	# Linear regression on curve-stripped p3 data and plotting line
	self.r2_p1, self.slope_p1, self.intercept_p1 = Operations.linear_regression (
	    self.x_p1_curvestrippedof_p23,
	    self.y_p1_curvestrippedof_p23)
	self.x1_p1, self.x2_p1, self.y1_p1, self.y2_p1 = Operations.grab_x_ys(
	    self.x_p1_curvestrippedof_p23,
	    self.intercept_p1,
	    self.slope_p1)	

    def draw_figure(self):
        """ Redraws the figure
        """     
        
        # Clearing the plots so they can be redrawn anew
        self.plot_phase1.clear()
        self.plot_phase2.clear()
        self.plot_phase3.clear()        

	self.plot_phase1.grid(self.cb_grid.IsChecked())        
	self.plot_phase2.grid(self.cb_grid.IsChecked())        
        self.plot_phase3.grid(self.cb_grid.IsChecked())
	
	# Graphing complete log efflux data set
        self.plot_phase3.scatter(
            self.x,
            self.y,
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
	if num_points_obj < 2:
	    num_points_obj = 2
	    self.obj_textbox.SetValue ('2')
	    
	self.obj_analysis ()
	    
	# Graphing the p3 series and regression line
	self.plot_phase3.scatter(
                    self.x_p3,
                    self.y_p3,
                    s = self.slider_width.GetValue(),
                    alpha = 0.75,
                    edgecolors = 'k',
                    facecolors = 'k'
                )            
	line_p3 = matplotlib.lines.Line2D (
            [self.x1_p3, self.x2_p3],
            [self.y1_p3, self.y2_p3],
            color = 'r',
            ls = '-',
            label = 'Phase III')
	self.plot_phase3.add_line (line_p3)
	
	# Distiguishing the intial points used to start the regression
	# and plotting them (solid red)
	self.plot_phase3.scatter(
                    self.x_reg_start,
                    self.y_reg_start,
                    s = self.slider_width.GetValue(),
                    alpha = 0.5,
                    edgecolors = 'r',
                    facecolors = 'r'
                )	
			    
	# Graphing the p2 series and regression line
	
	# Graphing raw uncorrected data of p1 and p2
	self.plot_phase2.scatter(
                    self.x [:self.reg_end_index],
                    self.y [:self.reg_end_index],
                    s = self.slider_width.GetValue(),
                    alpha = 0.50,
                    edgecolors = 'k',
                    facecolors = 'w'
                )
	
	# Graphing curve-stripped (corrected) phase I and II data, isolated
	# p2 data and line of best fit
	self.plot_phase2.scatter(
	    self.x_p12_curvestrippedof_p3,
	    self.y_p12_curvestrippedof_p3,
	    s = self.slider_width.GetValue(),
	    alpha = 0.50,
	    edgecolors = 'r',
	    facecolors = 'w')
	
	self.plot_phase2.scatter(
	    self.x_p2_curvestrippedof_p3,
	    self.y_p2_curvestrippedof_p3,
	    s = self.slider_width.GetValue(),
	    alpha = 0.75,
	    edgecolors = 'k',
	    facecolors = 'k')	
	
	self.line_p2 = matplotlib.lines.Line2D (
            [self.x1_p2, self.x2_p2],
            [self.y1_p2, self.y2_p2],
            color = 'r',
            ls = ':',
            label = 'Phase II')
	self.plot_phase2.add_line (self.line_p2)	
		                
	# Graphing the p1 series and regression line
	
	# Graph raw uncorrected p1 data, p1 data corrected for p3, and p1 data
	# Corrected for both p2 and p3
	self.plot_phase1.scatter(
                    self.x_p1,
                    self.y_p1,
                    s = self.slider_width.GetValue(),
                    alpha = 0.25,
                    edgecolors = 'k',
                    facecolors = 'w'
                )
	
	# Graph p1 data corrected for p3
	self.plot_phase1.scatter(
	    self.x_p1_curvestrippedof_p3,
	    self.y_p1_curvestrippedof_p3,
	    s = self.slider_width.GetValue(),
	    alpha = 0.25,
	    edgecolors = 'r',
	    facecolors = 'w'
	)	
	
	self.plot_phase1.scatter(
	    self.x_p1_curvestrippedof_p23,
	    self.y_p1_curvestrippedof_p23,
	    s = self.slider_width.GetValue(),
	    alpha = 0.75,
	    edgecolors = 'k',
	    facecolors = 'k'
	)
	
	self.line_p1 = matplotlib.lines.Line2D (
            [self.x1_p1,self.x2_p1],
            [self.y1_p1,self.y2_p1],
            color = 'r',
            ls = '--',
            label = 'Phase I'
	)
	self.plot_phase1.add_line (self.line_p1)          
    
	# Outputting the data from the linear regressions to widgets
	self.data_p1_slope.SetValue ('%0.3f'%(self.slope_p1))
	self.data_p1_int.SetValue ('%0.3f'%(self.intercept_p1))
	self.data_p1_r2.SetValue ('%0.3f'%(self.r2_p1))
	
	self.data_p2_slope.SetValue ('%0.3f'%(self.slope_p2))
	self.data_p2_int.SetValue ('%0.3f'%(self.intercept_p2))
	self.data_p2_r2.SetValue ('%0.3f'%(self.r2_p2))        
	
	self.data_p3_slope.SetValue ('%0.3f'%(self.slope_p3))
	self.data_p3_int.SetValue ('%0.3f'%(self.intercept_p3))
	self.data_p3_r2.SetValue ('%0.3f'%(self.r2_p3))         
        
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
    
    def on_pick(self, event):
        # The event received here is of the type
        # matplotlib.backend_bases.PickEvent
        #
        # It carries lots of information, of which we're using
        # only a small amount here.
        # 
        ind = event.ind
        x_clicked = np.take(self.x, ind)
        
        self.x_clicked_data.SetValue ('%0.2f'%(np.take(self.x, ind)[0]))
        self.y_clicked_data.SetValue ('%0.3f'%(np.take(self.y, ind)[0]))
        self.num_clicked_data.SetValue ('%0.0f'%(ind[0]+1))
        
    def on_save_plot(self, event):
        file_choices = "PNG (*.png)|*.png"
        
        dlg = wx.FileDialog(
            self, 
            message="Save plot as...",
            defaultDir=os.getcwd(),
            defaultFile="plot.png",
            wildcard=file_choices,
            style=wx.SAVE)
        
        if dlg.ShowModal() == wx.ID_OK:
            path = dlg.GetPath()
            self.canvas.print_figure(path, dpi=self.dpi)
            self.flash_status_message("Saved to %s" % path)
        
    def on_exit(self, event):
        self.Destroy()
        
    def on_about(self, event):
        msg = """ A demo using wxPython with matplotlib:
        
         * Use the matplotlib navigation bar
         * Add values to the text box and press Enter (or click "Draw!")
         * Show or hide the grid
         * Drag the slider to modify the width of the bars
         * Save the plot to a file using the File menu
         * Click on a bar to receive an informative message
        """
        dlg = wx.MessageDialog(self, msg, "About", wx.OK)
        dlg.ShowModal()
        dlg.Destroy()
    
    def flash_status_message(self, msg, flash_len_ms=1500):
        self.statusbar.SetStatusText(msg)
        self.timeroff = wx.Timer(self)
        self.Bind(
            wx.EVT_TIMER, 
            self.on_flash_status_off, 
            self.timeroff)
        self.timeroff.Start(flash_len_ms, oneShot=True)
    
    def on_flash_status_off(self, event):
        self.statusbar.SetStatusText('')


if __name__ == '__main__':
    data = ["Run 1", 35714.845, 8679.3, 4746.2, 0.6027, 1.00841763438286, 60.0, [0.0, 1.0, 2.0, 3.0, 4.0, 5.0, 6.0, 7.0, 8.0, 9.0, 10.0, 11.5, 13.0, 14.5, 16.0, 17.5, 19.0, 20.5, 22.0, 23.5, 25.0, 27.0, 29.0, 31.0, 33.0, 35.0, 37.0, 39.0, 41.0, 43.0, 45.0], [81453.0, 20369.1, 5511.0, 2790.7, 1933.8, 1456.8, 1159.5, 1015.1, 882.4, 788.5, 801.7, 698.5, 647.9, 629.3, 622.2, 507.9, 504.2, 422.8, 451.9, 411.3, 475.7, 453.3, 467.4, 431.2, 404.4, 432.2, 363.8, 429.8, 303.4, 348.1]]
  
    
    app = wx.PySimpleApp()
    app.frame = MainFrame(*(data + ["C:/Users/daniel/Desktop"]))
    app.frame.Show()
    app.frame.Center()
    app.MainLoop()