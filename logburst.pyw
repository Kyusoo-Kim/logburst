# -*- coding: utf-8 -*-

from tkinter import *
from tkinter import ttk
from tkinter.messagebox import *
from tkinter.filedialog import *

from types import *

import operator
import sys
import re
import datetime
import xlsxwriter

""" Define Area """

loglist = None
data = None
line_count = None
level_count = None
file_list = None



### Time Calculation
def get_timestamp(date_time):
    time_tuple = datetime.datetime.strptime(date_time, "%Y-%m-%d %H:%M:%S.%f")
    return(time_tuple.timestamp())

def get_time_required(start, end):
        start_timestamp = get_timestamp(start)
        end_timestamp = get_timestamp(end)

        return round( (end_timestamp - start_timestamp), 0 )

def get_formated_time(required_seconds):
        
        hour = round(required_seconds/3600, 0)
        remain_minutes = required_seconds % 3600
        
        minute = round(remain_minutes/60, 0)
        remain_seconds = remain_minutes % 60
        
        second = remain_seconds % 60

        formated_time = None

        if(hour > 0):
                formated_time = str(int(hour)) + 'h' + str(int(minute)) + 'm' + str(int(second)) + 's'
        elif(minute > 0):
                formated_time = str(int(minute)) + 'm' + str(int(second)) + 's'     
        else:
                formated_time = str(int(second)) + 's'

        return(formated_time)





def initData():
    global loglist, data, line_count, level_count, file_list
    
    loglist = []
    data = []
    line_count = 0
    level_count = {'V':0, 'D':0, 'I':0, 'W':0, 'E':0, 'F':0}
    file_list = []
    app.setProgress(0)
    app.btnOpen.config(state=NORMAL)
    app.btnExport.config(state=NORMAL)
    app.btnExit.config(state=NORMAL)
    app.listboxFile.delete(0,END)

def get_data(file_list, parsing_type=1):
    global line_count
    for files in file_list:
        try:
            with open(file=files, encoding='cp437') as log_file:
                for each_line in log_file:
                    try:
                        normal_style1 = re.match(r'\s*(?P<date>\d\d-\d\d)\s+(?P<time>\d\d:\d\d:\d\d\.\d+)\s+(?P<level>\D)/(?P<tag>.+)\s*\(\s*(?P<pid>\d+)\)\:\s*(?P<log>.+)', each_line)                                        
                        normal_style2 = re.match(r'\s*(?P<date>\d\d-\d\d)\s+(?P<time>\d\d:\d\d:\d\d\.\d+)\s+(?P<pid>\d+)\s+(?P<tid>\d+)\s+(?P<level>\D)\s(?P<tag>[^:]+):\s*(?P<log>.*)', each_line)
                        ddms_style = re.match(r'\s*(?P<level>\D)/(?P<tag>.+)\(\s*(?P<pid>\d+)\):\s*(?P<log>.*)', each_line)

                        
                        
                        
                        if(normal_style1):
                            line_count = line_count + 1
                            date  = normal_style1.group("date")
                            time  = normal_style1.group("time")
                            pid   = normal_style1.group("pid")
                            level = normal_style1.group("level")
                            tag   = normal_style1.group("tag")
                            log   = normal_style1.group("log")
                            
                            addLog(tag, level, pid, log, date, time)
                            
                        elif(normal_style2):
                            line_count = line_count + 1
                            date  = normal_style2.group("date")
                            time  = normal_style2.group("time")
                            pid   = normal_style2.group("pid")
                            level = normal_style2.group("level")
                            tag   = normal_style2.group("tag")
                            log   = normal_style2.group("log")
                            addLog(tag, level, pid, log, date, time)
                            
                        elif(ddms_style):
                            line_count = line_count + 1
                            pid   = ddms_style.group("pid")
                            level = ddms_style.group("level")
                            tag   = ddms_style.group("tag")
                            log   = ddms_style.group("log")
                            addLog(tag, level, pid, log)

                        normal_style1 = normal_style2 = ddms_style = None
                        
                    except ValueError as val_err:
                        pass

                if date and time:
                    time_info.end_time = str(datetime.datetime.now().year) + '-' + date + ' ' + time
                    
        except IOError as io_err:
            print(str(io_err))
        except UnicodeDecodeError as unicode_err:
                        print(line_count)
                        print(str(unicode_err))

def addLog(tag, level, pid, text, date='', time=''):
    level_count[level] +=  1

    data.append({'tag':tag, 'level':level, 'pid':pid, 'text':text, 'date':date, 'time':time})        

    if len(loglist) == 0:
        new_log =  {'tag':tag, 'count' : 1, 'level': {'V':0, 'D':0, 'I':0, 'W':0, 'E':0, 'F':0} }
        new_log['level'][level] =  new_log['level'][level] + 1
        loglist.append(new_log)

        if( len(date) > 0 and len(time) > 0 ):
                time_info.start_time = str(datetime.datetime.now().year) + '-' + date + ' ' + time
        
        return()

    find = False
    for log in loglist:
        if tag == log['tag']:
            log['count'] = log['count'] + 1
            log['level'][level] = log['level'][level] + 1
            find = True
    
    if not find:
        new_log =  {'tag':tag, 'count' : 1, 'level': {'V':0, 'D':0, 'I':0, 'W':0, 'E':0, 'F':0} }
        new_log['level'][level] =  new_log['level'][level] + 1
        loglist.append(new_log)

def makeExcel(file_save, chart_limit=10):

    log_data_column_width = [8, 12, 8, 26, 8, 120]
    log_analysis_column_width = [25, 10, 5, 10, 10, 10, 10, 10, 10]
    
    workbook = xlsxwriter.Workbook(file_save)
    worksheet1 = workbook.add_worksheet('Data')
    worksheet2 = workbook.add_worksheet('Analysis')
    worksheet3 = workbook.add_worksheet('Chart')

    for i in range(len(log_data_column_width)):
        worksheet1.set_column(i,i, log_data_column_width[i])

    for i in range(len(log_analysis_column_width)):
        worksheet2.set_column(i,i, log_analysis_column_width[i])

    worksheet1.set_tab_color('red')
    worksheet2.set_tab_color('green')
    worksheet3.set_tab_color('blue')

   

    tagCountChart = workbook.add_chart({'type':'column'})
    logLevelChart = workbook.add_chart({'type':'column'})
    logLevPieChart = workbook.add_chart({'type': 'pie', 'embedded':1})

    bold = workbook.add_format({'bold':True})
    silver = workbook.add_format({'font_color':'silver'})
    blue = workbook.add_format({'font_color':'blue'})
    green = workbook.add_format({'font_color':'green'})
    yellow = workbook.add_format({'font_color':'yellow'})
    orange = workbook.add_format({'font_color':'orange'})
    red = workbook.add_format({'font_color':'red'})
    
 

    worksheet1.write(0, 0, 'DATE', bold)
    worksheet1.write(0, 1, 'TIME', bold)
    worksheet1.write(0, 2, 'LEVEL', bold)
    worksheet1.write(0, 3, 'TAG', bold)
    worksheet1.write(0, 4, 'PID', bold)
    worksheet1.write(0, 5, 'LOG', bold)

    worksheet1.freeze_panes(1, 0)
    worksheet1.autofilter('A1:F1')



    
        
    
   


    worksheet2.write(0, 0, "TAG")
    worksheet2.write(0, 1, "Count")
    worksheet2.write(0, 3, "Verbose", silver)
    worksheet2.write(0, 4, "Debug", blue)
    worksheet2.write(0, 5, "Info", green)
    worksheet2.write(0, 6, "Warning", yellow)
    worksheet2.write(0, 7, "Error", orange)
    worksheet2.write(0, 8, "Fatal", red)

    worksheet2.freeze_panes(1, 0)
    worksheet2.autofilter('A1:I1')

    

    sorted_log_list = sorted(loglist, key=operator.itemgetter('count'), reverse=True)

    nr_logs = len(sorted_log_list)
    for i in range(nr_logs):
        worksheet2.write(i+1, 0, sorted_log_list[i]['tag'])
        worksheet2.write(i+1, 1, sorted_log_list[i]["count"])

        worksheet2.write(i+1, 3, sorted_log_list[i]["level"]["V"])
        worksheet2.write(i+1, 4, sorted_log_list[i]["level"]["D"])
        worksheet2.write(i+1, 5, sorted_log_list[i]["level"]["I"])
        worksheet2.write(i+1, 6, sorted_log_list[i]["level"]["W"])
        worksheet2.write(i+1, 7, sorted_log_list[i]["level"]["E"])
        worksheet2.write(i+1, 8, sorted_log_list[i]["level"]["F"])
        

    worksheet2.write(nr_logs+2, 0, "Total")
    worksheet2.write(nr_logs+2, 1, line_count)

    worksheet2.write(nr_logs+2, 3, level_count['V'])
    worksheet2.write(nr_logs+2, 4, level_count['D'])
    worksheet2.write(nr_logs+2, 5, level_count['I'])
    worksheet2.write(nr_logs+2, 6, level_count['W'])
    worksheet2.write(nr_logs+2, 7, level_count['E'])
    worksheet2.write(nr_logs+2, 8, level_count['F'])
    

    

    tagCountChart.add_series({'name':None, 
                          'categories': '=Analysis!$A$2:$A$' + str(chart_limit+1),
                          'values': '=Analysis!$B$2:$B$' + str(chart_limit+1),
                          'data_labels':{'value':True},
                          'fill': {'color':'red'}
                          })
            
    required_time = get_time_required( time_info.start_time, time_info.end_time )
    formated_time = get_formated_time(required_time)
    tagCountChart.set_title({'name': 'Tag Count (' + formated_time + ')', 'name_font':{'size':20}})
    tagCountChart.set_x_axis({'num_font':{'size':16}})
    tagCountChart.set_y_axis({'num_font':{'size':16}})
    tagCountChart.set_legend({'position':'none'})


    logLevelChart.set_x_axis({
            'name' : 'Level Count',            
            'name_font':{
                 'size':16,
                 'bold':True
                 },
            'label_position':'none'
            })
    logLevelChart.set_legend({'position':'bottom'})
    
    logLevelChart.add_series({
            'name':'Verbose',
            'categories': '=Analysis!$D$1',
            'values': '=Analysis!$D$' + str(nr_logs+3),
            'data_labels':{'value':True},
            'fill': {'color':'silver'}
            })

    logLevelChart.add_series({
            'name':'Debug',
            'categories': '=Analysis!$E$1',
            'values': '=Analysis!$E$' + str(nr_logs+3),
            'data_labels':{'value':True},
            'fill': {'color':'blue'}
            })

    logLevelChart.add_series({
            'name':'Info',
            'categories': '=Analysis!$F$1',
            'values': '=Analysis!$F$' + str(nr_logs+3),
            'data_labels':{'value':True},
            'fill': {'color':'green'}
            })

    logLevelChart.add_series({
            'name':'Warning',
            'categories': '=Analysis!$G$1',
            'values': '=Analysis!$G$' + str(nr_logs+3),
            'data_labels':{'value':True},
            'fill': {'color':'yellow'}
            })

    logLevelChart.add_series({
            'name':'Error',
            'categories': '=Analysis!$H$1',
            'values': '=Analysis!$H$' + str(nr_logs+3),
            'data_labels':{'value':True},
            'fill': {'color':'orange'}
            })

    logLevelChart.add_series({
            'name':'Fatal',
            'categories': '=Analysis!$I$1',
            'values': '=Analysis!$I$' + str(nr_logs+3),
            'data_labels':{'value':True},
            'fill': {'color':'red'}
            })
    


    

    logLevPieChart.add_series({
            'name': 'LogLevel',
            'categories': '=Analysis!D1:I1',
            'values': '=Analysis!D' + str(nr_logs+3) + ':I' + str(nr_logs+3),
            'data_labels':{'category':True, 'percentage':True,'leader_lines':True},
            'points': [
                    {'fill': {'color': 'silver'}},
                    {'fill': {'color': 'blue'}},
                    {'fill': {'color': 'green'}},
                    {'fill': {'color': 'yellow'}},
                    {'fill': {'color': 'orange'}},
                    {'fill': {'color': 'red'}}
                ]
        })
    logLevPieChart.set_title({'name':'Level Chart'})
    logLevPieChart.set_legend({'position':'bottom'})
    

    
    
    worksheet3.insert_chart('B2', tagCountChart)
    worksheet3.insert_chart('B18', logLevelChart)
    worksheet3.insert_chart('K18', logLevPieChart)

    worksheet3.activate()
        
    for i in range(len(data)):
        worksheet1.write(i+1, 0, data[i]['date'])
        worksheet1.write(i+1, 1, data[i]['time'])
        worksheet1.write(i+1, 2, data[i]['level'])
        worksheet1.write(i+1, 3, data[i]['tag'])
        worksheet1.write(i+1, 4, data[i]['pid'])        
        
        
        if not ( "" in data[i]['text']):
            worksheet1.write(i+1, 5, data[i]['text'])
        app.setProgress( i / line_count * 100)

        
    
    workbook.close()
    

class Data:
    def __init__(self):
        self.init_data()

    def init_data(self):
        self.start_time = None
        self.end_time = None


#GUI
class App:
    def __init__(self, master):
        
        master.geometry('300x230')
        master.title('Log Burst')
        master.resizable(width=False, height=False)

        frameTop = Frame(master)
        frameTop.pack(side=TOP, pady=5, fill=X)

        frameMid = Frame(master)
        frameMid.pack(side=TOP, pady=5, fill=X)

        frameBottom = Frame(master)
        frameBottom.pack(side=BOTTOM, pady=5, fill=X)

        #Menu
        menu = Menu(master)
        master.config(menu=menu)

        filemenu = Menu(menu)
        menu.add_cascade(label="File", menu=filemenu)
        filemenu.add_command(label="Open Log File", command=self.file_open)
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=sys.exit)

        helpmenu = Menu(menu)
        menu.add_cascade(label="Help", menu=helpmenu)
        helpmenu.add_command(label="About...", command=self.about)

        #Item

        self.btnOpen = ttk.Button(frameTop, text="Open", command=self.file_open)
        self.btnOpen.pack(fill=X, padx=10, pady=5)

        self.scrollbar = ttk.Scrollbar(frameMid)
        self.scrollbar.pack(side=RIGHT, fill=Y)

        self.listboxFile = Listbox(frameMid, height=4, yscrollcommand=self.scrollbar.set)
        self.listboxFile.pack(fill=X, padx=10)

        self.scrollbar.config(command=self.listboxFile.yview)
        
        self.progressbar = ttk.Progressbar(frameBottom, orient=HORIZONTAL, maximum=100, value=0, mode='determinate')
        self.progressbar.pack(fill=X, padx=10, pady=5)
        
        self.btnExport = ttk.Button(frameBottom, text="Export", command=self.export)
        self.btnExport.pack(fill=X, padx=10)

        self.btnExit = ttk.Button(frameBottom, text="Exit", command=sys.exit)
        self.btnExit.pack(side=BOTTOM, fill=X, padx=10, pady=5)
        
    

    def file_open(self):
        global file_list
        ext = [('Log files', '*.log;*.txt'), ('All files', '*')]
        file_read = askopenfilename(title="Open Log File", filetypes=ext, multiple=True)
        
        has_white_space = re.findall('{([^}]*)}', file_read)

        file_list = []
        if not type(has_white_space) == type(None):
            for path in has_white_space:
                file_read = file_read.replace("{" +path+ "}", "")

        file_read = file_read.split(' ')

        for path in file_read:
            if len(path) > 0:
                file_list.append(path)

        for path in has_white_space:
            file_list.append(path)

    
        self.listboxFile.delete(0,END)
        for file in file_list:
            self.listboxFile.insert(END, file)
        
    
    def export(self):
        global file_list
        if len(file_list) > 0:

            ext = [('Excel files', '*.xlsx')]
            file_save = asksaveasfilename(title="Save Excel File", filetypes=ext)

            start = datetime.datetime.now()

            if len(file_save) > 0:
                if not '.xls' in file_save:            
                    file_save = file_save + '.xlsx'
                
                self.btnOpen.config(state=DISABLED)
                self.btnOpen.update()
                self.btnExport.config(state=DISABLED)
                self.btnExport.update()
                self.btnExit.config(state=DISABLED)
                self.btnExit.update()

                get_data(file_list)                
                makeExcel(file_save)

                end = datetime.datetime.now()

                print( str( end.timestamp() - start.timestamp() ) + " sec" )

                showinfo("Export", "Export result file!\n\n[" + file_save +"]")
                os.startfile(file_save)
            
            
            initData()

            

    def setProgress(self, val):
        self.progressbar.config(value=val)
        self.progressbar.update()

    def about(self):
        showinfo("About...", "P1 BSP Perf.\n\nContact : kyusoo.kim@lge.com")




time_info = Data()    
root = Tk()
app = App(root)
initData()
root.mainloop()
