# Import Area
import sys
import xlsxwriter

# Define Area

class LogCounter:
    def __init__(self):
        self.tagNames = []
        self.tagCount = {}
        self.levelCount = {'V':0, 'D':0, 'I':0, 'W':0, 'E':0}
        self.sort_idx = []
        
        
    def addTag(self, level, tag):
        if tag in self.tagNames:
            self.tagCount[tag] = self.tagCount[tag] + 1
        else:
            self.tagNames.append(tag)
            self.tagCount[tag] = 1
            
        if 'V' in level:
            self.levelCount['V'] = self.levelCount['V'] + 1
        elif 'D' in level:
            self.levelCount['D'] = self.levelCount['D'] + 1
        elif 'I' in level:
            self.levelCount['I'] = self.levelCount['I'] + 1
        elif 'W' in level:
            self.levelCount['W'] = self.levelCount['W'] + 1
        elif 'E' in level:
            self.levelCount['E'] = self.levelCount['E'] + 1
                            
    def printLevel(self):
        print('V : ' + str(self.levelCount['V']))
        print('D : ' + str(self.levelCount['D']))
        print('I : ' + str(self.levelCount['I']))
        print('W : ' + str(self.levelCount['W']))
        print('E : ' + str(self.levelCount['E']))

    def getTagNames(self):
        return(self.tagNames)
    
    def getTagCount(self):
        return(self.tagCount)
    
    def sort(self):
        sort_idx = []
        tag_len = len(self.tagNames)
        
        for i in range(tag_len):
            self.sort_idx.append(i)       
            
        for i in range(tag_len-1):
            for j in range(i+1, tag_len):        
                if self.tagCount[self.tagNames[self.sort_idx[i]]] < self.tagCount[self.tagNames[self.sort_idx[j]]]:
                    temp_idx = self.sort_idx[i]
                    self.sort_idx[i] = self.sort_idx[j]
                    self.sort_idx[j] = temp_idx        

    
    def makeExcel(self, chart_limit=10):
        workbook = xlsxwriter.Workbook(filename + '.xlsx')
        worksheet1 = workbook.add_worksheet()
        worksheet2 = workbook.add_worksheet()        

#tagCountChart
        tagCountChart = workbook.add_chart({'type':'column'})
        
        worksheet2.write(0, 0, "TAG")
        worksheet2.write(0, 1, "Count")
        
        tag_len = len(self.tagNames)    
        for idx in range(tag_len):
            worksheet2.write(idx+1, 0, self.tagNames[self.sort_idx[idx]])
            worksheet2.write(idx+1, 1, self.tagCount[self.tagNames[self.sort_idx[idx]]])
            print(self.tagNames[self.sort_idx[idx]] + " : " + str(self.tagCount[self.tagNames[self.sort_idx[idx]]]) )
            
         
        tagCountChart.add_series({'name':None, 
                          'categories': '=Sheet2!$A$2:$A$' + str(chart_limit+1),
                          'values': '=Sheet2!$B$2:$B$' + str(chart_limit+1),
                          'fill': {'color':'red'}
                          })
            
        
        tagCountChart.set_title({'name': 'Log Burst', 'name_font':{'size':20}})
        tagCountChart.set_x_axis({'num_font':{'size':16}})
        tagCountChart.set_y_axis({'num_font':{'size':16}})
        tagCountChart.set_legend({'position':'none'})
        
        tagCountChart.set_size({'width':1080, 'height':300})
        
#logLevelChart
        
        worksheet2.write(0, 3, "Level")
        worksheet2.write(0, 4, "Count")
        
        worksheet2.write(1, 3, "V")
        worksheet2.write(2, 3, "D")
        worksheet2.write(3, 3, "I")
        worksheet2.write(4, 3, "W")
        worksheet2.write(5, 3, "E")

        worksheet2.write(1, 4, self.levelCount['V'])
        worksheet2.write(2, 4, self.levelCount['D'])
        worksheet2.write(3, 4, self.levelCount['I'])
        worksheet2.write(4, 4, self.levelCount['W'])
        worksheet2.write(5, 4, self.levelCount['E'])
                         
        logLevelChart = workbook.add_chart({'type':'column'})
        logLevelChart.set_x_axis({
            'name' : 'Log level',
            
            'name_font':{
                 'size':16,
                 'bold':True
                 },

            'label_position':'none'
            })
        
        logLevelChart.add_series({
            'name':'Verbose',
            'categories': '=Sheet2!$D$2',
            'values': '=Sheet2!$E$2',
            'fill': {'color':'black'}
            })

        logLevelChart.add_series({
            'name':'Debug',
            'categories': '=Sheet2!$D$3',
            'values': '=Sheet2!$E$3',
            'fill': {'color':'blue'}
            })

        logLevelChart.add_series({
            'name':'Info',
            'categories': '=Sheet2!$D$4',
            'values': '=Sheet2!$E$4',
            'fill': {'color':'green'}
            })

        logLevelChart.add_series({
            'name':'Warning',
            'categories': '=Sheet2!$D$5',
            'values': '=Sheet2!$E$5',
            'fill': {'color':'orange'}
            })

        logLevelChart.add_series({
            'name':'Error',
            'categories': '=Sheet2!$D$6',
            'values': '=Sheet2!$E$6',
            'fill': {'color':'red'}
            })

        logLevelChart.set_size({'width':1080, 'height':300})
        
        worksheet1.insert_chart('B2', tagCountChart)
        worksheet1.insert_chart('B18', logLevelChart)
        workbook.close()
        

# Global Value
lc = LogCounter()
lineCount = 0
# Logic Area

filename = input('FileName: ')

try:
    with open(filename) as log_file:
        for each_line in log_file:
            try:
                if not each_line == '\n':
                    (level, temp_line) = each_line.split('/', 1)
                    (tag, temp_line) = temp_line.split('(', 1)
                    (pid, temp_line) = temp_line.split(')', 1)
                    
                    lc.addTag(level, tag)
                    lineCount = lineCount + 1
            except ValueError as val_err:
                pass
        
        lc.sort()
        lc.makeExcel(10)
except IOError as io_err:
    print(str(io_err))
    

print("\nTotal Log Line : " + str(lineCount))
lc.printLevel()
