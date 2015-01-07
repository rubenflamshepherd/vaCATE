import wx
import wx.lib.plot as plot
import os,sys
import time
import xlsxwriter

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
class DialogFrame(wx.Frame):
        
    def __init__(self, parent, id, title):
        wx.Frame.__init__(self, parent, id, title)
        self.SetIcon(wx.Icon('Images/testtube.ico', wx.BITMAP_TYPE_ICO))

        self.rootPanel = wx.Panel(self)
        
        innerPanel = wx.Panel(self.rootPanel,-1, size=(500,160), style=wx.ALIGN_CENTER)
        hbox = wx.BoxSizer(wx.HORIZONTAL) 
        vbox = wx.BoxSizer(wx.VERTICAL)
        innerBox = wx.BoxSizer(wx.VERTICAL)
        buttonBox1 = wx.BoxSizer(wx.HORIZONTAL)
        buttonBox2 = wx.BoxSizer(wx.HORIZONTAL)
        
        # Main text presented to user
        txt1 = wx.StaticText(innerPanel, id=-1, label="     Welcome to the CATE Data Analyzer!     ",style=wx.ALIGN_CENTER, name="")
        txt2 = wx.StaticText(innerPanel, id=-1, label="Please choose an option below:",style=wx.ALIGN_CENTER, name="")
        
        # Disclaimer text (under buttons)
        txt3 = wx.StaticText(innerPanel, id=-1, label="Note: .xls output files will be written in the same folder", style=wx.ALIGN_CENTER, name="")
        txt4 = wx.StaticText(innerPanel, id=-1, label="that data is being extracted from", style=wx.ALIGN_CENTER, name="")
        
        font3 = wx.Font (7, wx.DEFAULT, wx.NORMAL, wx.NORMAL)
        txt3.SetFont (font3)
        txt4.SetFont (font3)
        
        # Option Buttons        
        btn1 = wx.Button (innerPanel, id=1, label="Analyze CATE Data")
        btn2 = wx.Button (innerPanel, id=2, label="Generate CATE Template")
        btn3 = wx.Button (innerPanel, id=3, label="About")
        btn4 = wx.Button (innerPanel, id=4, label="Quit")
        
        # Binding events to buttons
        self.Bind(wx.EVT_BUTTON, self.OnAnalyze, id=1)        
        self.Bind(wx.EVT_BUTTON, self.OnGenerate, id=2)        
        self.Bind(wx.EVT_BUTTON, self.OnAbout, id=3)        
        self.Bind(wx.EVT_BUTTON, self.OnClose, id=4)
        
        # Adding main text to main spacer 'innerBox'
        innerBox.AddSpacer((150,15))
        innerBox.Add(txt1, 0, wx.CENTER)
        innerBox.AddSpacer((150,15))
        innerBox.Add(txt2, 0, wx.CENTER)
        innerBox.AddSpacer((150,15))
        
        # Adding main program buttons to main spacer 'innerBox'
        buttonBox1.AddSpacer(7,10)        
        buttonBox1.Add(btn1, 0, wx.CENTER)
        buttonBox1.AddSpacer(7,10)
        buttonBox1.Add(btn2, 0, wx.CENTER)
        buttonBox1.AddSpacer(7,15)
        buttonBox2.Add(btn3, 0, wx.CENTER)
        buttonBox2.AddSpacer(7,15)        
        buttonBox2.Add(btn4, 0, wx.CENTER)
        buttonBox2.AddSpacer(7,10)        
        innerBox.Add(buttonBox1, 0, wx.CENTER)
        innerBox.AddSpacer ((150,10))
        innerBox.Add(buttonBox2, 0, wx.CENTER)
        innerBox.AddSpacer ((150,10))
        
        # Adding disclaimer text to main spacer 'innerBox' (under buttons)
        innerBox.Add(txt3, 0, wx.CENTER)
        innerBox.Add(txt4, 0, wx.CENTER)
        innerBox.AddSpacer((150,10))        
        innerPanel.SetSizer(innerBox)

        hbox.Add(innerPanel, 0, wx.ALL|wx.ALIGN_CENTER)
        vbox.Add(hbox, 1, wx.ALL|wx.ALIGN_CENTER, 5)
        

        self.rootPanel.SetSizer(vbox)
        vbox.Fit(self)
        
    def OnClose(self, event): # Event when 'Close' button is pressed
            self.Close()
        
    def OnAnalyze(self, event): # Event when 'Analyze CATE data' button is pushed
        dlg = wx.FileDialog(self, "Choose the file which contains the data you'd like to perform CATE upon", os.getcwd(), "", "")
        if dlg.ShowModal() == wx.ID_OK:
            directory, filename = dlg.GetDirectory(), dlg.GetFilename()
            
            temp_CATE_data = Excel.grab_data (directory, filename)
            print temp_CATE_data
            frame = Preview.MainFrame (*(temp_CATE_data  + [directory]))
            frame.Show (True)
            frame.MakeModal (True)            
            dlg.Destroy()
                        
        #self.Close()                         
        
        # to generate FINAL excel file (probably won't be in this module)
        # CATE.generate_workbook (directory, individual_inputs, series_inputs)
            
    def OnAbout(self, event): # Event when 'About' button is pushed
        dlg = AboutDialog (self, -1, 'About')
        #dlg.SetIcon(wx.Icon('Images/testtube.ico', wx.BITMAP_TYPE_ICO))
        val = dlg.ShowModal()
        dlg.Destroy()
    
    def OnGenerate(self, event): # Event when 'Generate Template' button is pushed
        dlgChoose = wx.DirDialog(self, "Choose the directory to generate the template file inside:")
                        
        if dlgChoose.ShowModal() == wx.ID_OK:
            directory = dlgChoose.GetPath()
            dlgChoose.Destroy()
            # self.Close()            
            
            # format the directory (and path) to unicode w/ forward slash so it can be passed between methods/classes w/o bugs
            directory = u'%s' %directory
            self.directory = directory.replace (u'\\', '/')
            output_name = 'CATE Template - ' + time.strftime ("(%Y_%m_%d).xlsx")
            output_file_path = '/'.join ((directory, output_name))            
            
            workbook = xlsxwriter.Workbook(output_file_path)
            worksheet = workbook.add_worksheet("CATE Template")            
            Excel.generate_template (output_file_path, workbook, worksheet)
            
            # Formatting for items for which inputs ARE required
            req = workbook.add_format ()
            req.set_text_wrap ()
            req.set_align ('center')
            req.set_align ('vcenter')
            req.set_bold ()
            req.set_bottom ()
            
            # Formatting for run headers ("Run x")
            run_header = workbook.add_format ()    
            run_header.set_align ('center')
            run_header.set_align ('vcenter')
            
            # Formatting for row cells that are to recieve input
            empty_row = workbook.add_format ()
            empty_row.set_align ('center')
            empty_row.set_align ('vcenter')
            empty_row.set_top ()    
            empty_row.set_bottom ()    
            empty_row.set_right ()    
            empty_row.set_left ()            
            
            # Writing empy cells surronded by borders
            for y in range (0, 6):
                worksheet.write (y + 1, 3, "", empty_row)            
            
            worksheet.write (7, 3,"Activity in eluant (cpm)", req)
            worksheet.set_column (7, 3, 15)
            worksheet.write (0, 3, "Run 2", run_header)
            worksheet.write (0, 4, "etc.", run_header)            
            
            workbook.close()        
               
class MyApp(wx.App):
    def OnInit(self):
        frame = DialogFrame(None, -1, 'CATE Data Analyzer')
        frame.SetIcon(wx.Icon('testtube.ico', wx.BITMAP_TYPE_ICO))
        frame.Show(True)
        frame.Center()
        #self.SetTopWindow(frame)
        return True

if __name__ == '__main__':
    app = wx.PySimpleApp()
    app.frame = DialogFrame(None, -1, 'CATE Data Analyzer')
    app.frame.Show(True)
    app.frame.Center()
    app.MainLoop()
    