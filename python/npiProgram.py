import requests
import json
from xlrd import open_workbook
import tkinter as tk
from tkinter import ttk
import tkinter.filedialog
import sqlite3
import xlwt
import sys

import utils

conn = sqlite3.connect('taxonomyCodes.db')
cursor = conn.cursor()

class Application(tkinter.Frame):
    def main(self):
        wb = open_workbook(self.inputFileName.name)
        sheet = wb.sheet_by_index(0)
        
        number_of_rows = sheet.nrows
        number_of_columns = sheet.ncols 
        
        headings = utils.getHeadings(self.inputFileName.name) 
        row_data = []
        items = []
        rows = []
        
        for row in range(1, number_of_rows):
            values = {}
            for col in range(number_of_columns):
                key = sheet.cell(0, col).value
                value  = (sheet.cell(row,col).value)
                try:
                    value = str(value)
                except ValueError:
                    pass
                finally:
                    values.update({key: str(value)})
        
            hcp = []
            for key in headings:
                hcp.append(str(values[key]))
            hcp.append('1')
        
            row_data.append(hcp)

            rec_lname = self.lname.get() 
            rec_fname = self.fname.get() 
            rec_add1 = self.add1.get() 
            rec_add2 = self.add2.get() 
            rec_city = self.city.get() 
            rec_state = self.state.get() 
            rec_zip = self.zipCode.get() 
            rec_spec = self.spec.get() 
            rec_lic_state = self.lic_state.get() 
            rec_lic_number = self.lic_number.get() 
            rec_npi = self.npi.get() 
        
            lname = values[rec_lname]
            fname = values[rec_fname]
            reqst = requests.get('http://www.bloomapi.com/api/search?limit=100&offset=0&key1=last_name&op1=eq&value1='
                            + lname + '&key2=first_name&op2=eq&value2=' + fname)
            if reqst:
                req_text = json.loads(reqst.text)
                result_set = req_text['result']
    
                for i in range(len(result_set)):
                    hcp_lookup = []
                    if result_set[i]['type'] == 'individual':
                        #self.text1.insert('insert', result_set[i]['last_name'] + '\n')
                        #self.text1.pack()
                        values[rec_lname] = result_set[i]['last_name']
                        values[rec_fname] = result_set[i]['first_name']
                        values[rec_add1] = result_set[i]['practice_address']['address_line']
                        values[rec_add2] = utils.getAddressDetails(result_set, i)
                        values[rec_city] = result_set[i]['practice_address']['city']
                        values[rec_state] = result_set[i]['practice_address']['state']
                        values[rec_zip] = result_set[i]['practice_address']['zip']
                        values[rec_spec] = utils.getSpecialty(result_set, i, cursor)
                        values[rec_lic_state] = utils.getLicense(result_set, i)[1]
                        values[rec_lic_number] = utils.getLicense(result_set, i)[0]
                        values[rec_npi] = utils.getNpi(result_set, i)
        
                        for key in headings:
                            hcp_lookup.append(str(values[key]))
                        
                        row_data.append(hcp_lookup)
                print('Last Name: ' + values[rec_lname] + '; Results: ' + str(len(result_set)))
        utils.output(self, self.outputFileName, self.outputTabName, headings, row_data)
        self.finish() 

    def finish(self):
        self.finishGraphic = tk.Label(root, text='Finished', fg='red', font=('Arial', 10, 'bold'))
        self.finishGraphic.grid(row=16, column=1)

    def loadFile(self):
        self.inputFileName = tk.filedialog.askopenfile(parent=self.fileBar, mode='rb', initialdir="./", title='Please choose the input file')
        self.showFileName = tk.Label(self.fileBar, text=self.inputFileName.name)
        self.showFileName.grid(row=0, column=1, sticky='w')
        headings = utils.getHeadings(self.inputFileName.name) 

        self.fname = tk.StringVar(self.colBar)
        self.fname.set("RECIPIENT_FIRST_NAME")
        self.fir_name = tk.OptionMenu(self.colBar, self.fname, *headings)
        self.fir_name.grid(row=5, column=1, sticky='w')
        menu = self.fir_name.nametowidget(self.fir_name.menuname)
        menu.configure(font=('Arial', 6))

        self.lname = tk.StringVar(self.colBar)
        self.lname.set("RECIPIENT_LAST_NAME")
        self.las_name = tk.OptionMenu(self.colBar, self.lname, *headings)
        self.las_name.grid(row=6, column=1, sticky='w')
        menu = self.las_name.nametowidget(self.las_name.menuname)
        menu.configure(font=('Arial', 6))
        
        self.add1 = tk.StringVar(self.colBar)
        self.add1.set("RECIPIENT_ADDRESS_LINE1")
        self.address1 = tk.OptionMenu(self.colBar, self.add1, *headings)
        self.address1.grid(row=7, column=1, sticky='w')
        menu = self.address1.nametowidget(self.address1.menuname)
        menu.configure(font=('Arial', 6))

        self.add2 = tk.StringVar(self.colBar)
        self.add2.set("RECIPIENT_ADDRESS_LINE2")
        self.address2 = tk.OptionMenu(self.colBar, self.add2, *headings)
        self.address2.grid(row=8, column=1, sticky='w')
        menu = self.address2.nametowidget(self.address2.menuname)
        menu.configure(font=('Arial', 6))

        self.city = tk.StringVar(self.colBar)
        self.city.set("RECIPIENT_CITY")
        self.re_city = tk.OptionMenu(self.colBar, self.city, *headings)
        self.re_city.grid(row=9, column=1, sticky='w')
        menu = self.re_city.nametowidget(self.re_city.menuname)
        menu.configure(font=('Arial', 6))

        self.state = tk.StringVar(self.colBar)
        self.state.set("RECIPIENT_STATE")
        self.re_state = tk.OptionMenu(self.colBar, self.state, *headings)
        self.re_state.grid(row=10, column=1, sticky='w')
        menu = self.re_state.nametowidget(self.re_state.menuname)
        menu.configure(font=('Arial', 6))

        self.zipCode = tk.StringVar(self.colBar)
        self.zipCode.set("RECIPIENT_ZIP_CODE")
        self.zip_code = tk.OptionMenu(self.colBar, self.zipCode, *headings)
        self.zip_code.grid(row=11, column=1, sticky='w')
        menu = self.zip_code.nametowidget(self.zip_code.menuname)
        menu.configure(font=('Arial', 6))

        self.spec = tk.StringVar(self.colBar)
        self.spec.set("RECIPIENT_SPECIALTY")
        self.specialty = tk.OptionMenu(self.colBar, self.spec, *headings)
        self.specialty.grid(row=12, column=1, sticky='w')
        menu = self.specialty.nametowidget(self.specialty.menuname)
        menu.configure(font=('Arial', 6))

        self.lic_number = tk.StringVar(self.colBar)
        self.lic_number.set("RECIPIENT_STATE_LICENSE_NUMBER")
        self.license_number = tk.OptionMenu(self.colBar, self.lic_number, *headings)
        self.license_number.grid(row=13, column=1, sticky='w')
        menu = self.license_number.nametowidget(self.license_number.menuname)
        menu.configure(font=('Arial', 6))

        self.lic_state = tk.StringVar(self.colBar)
        self.lic_state.set("RECIPIENT_LICENSE_STATE")
        self.license_state = tk.OptionMenu(self.colBar, self.lic_state, *headings)
        self.license_state.grid(row=14, column=1, sticky='w')
        menu = self.license_state.nametowidget(self.license_state.menuname)
        menu.configure(font=('Arial', 6))

        self.npi = tk.StringVar(self.colBar)
        self.npi.set("RECIPIENT_NPI_NUMBER")
        self.re_npi = tk.OptionMenu(self.colBar, self.npi, *headings)
        self.re_npi.grid(row=15, column=1, sticky='w')
        menu = self.re_npi.nametowidget(self.re_npi.menuname)
        menu.configure(font=('Arial', 6))

    def runProgram(self):
        #self.progress = tk.ttk.Progressbar(root, orient="horizontal", length=200, mode='indeterminate')
        #self.progress.grid(row=16, column=1)
        #self.progress.start()
        self.outputFileName = self.saveToDirectory + '/' + self.outputName.get() + '.xls'
        self.outputTabName = self.outTabName.get() 
        self.main()

    def getSaveToDirectory(self):
        self.saveToDirectory = tk.filedialog.askdirectory(parent=self.fileBar, initialdir="./", title='Please select a directory')
        self.directoryName = tk.Label(self.fileBar, text=self.saveToDirectory)
        self.directoryName.grid(row=1, column=1)

    def createWidgets(self):

        # Frames
        self.fileBar = tk.Frame(root, relief='sunken', bd=2, borderwidth=2)
        self.fileBar.grid(row=0, column=0, columnspan=2, sticky='w'+'e')
        self.nameBar = tk.Frame(root, relief='sunken', bd=2, borderwidth=2)
        self.nameBar.grid(row=1, column=0, columnspan=2, sticky='w'+'e')
        self.colBar = tk.Frame(root, relief='raised', bd=2, borderwidth=2)
        self.colBar.grid(row=2, column=0, columnspan=2, sticky='w'+'e')
        self.controls= tk.Frame(root, relief='raised', bd=2, borderwidth=2)
        self.controls.grid(row=3, column=0, columnspan=2, sticky='w'+'e')
        
        # QUIT button
        self.QUIT = tk.Button(self.controls)
        self.QUIT.grid(row=16, column=0, sticky='w')
        self.QUIT["text"] = "Cancel"
        self.QUIT["fg"] = "red"
        self.QUIT["command"] = self.quit

        # Input file
        self.fileName = tk.Button(self.fileBar)
        self.fileName.grid(row=0, column=0, sticky='w')
        self.fileName["text"] = "Input File..."
        self.fileName["command"] = self.loadFile

        # Save directory
        self.saveTo = tk.Button(self.fileBar)
        self.saveTo.grid(row=1, column=0, sticky='w')
        self.saveTo["text"] = "Save Directory..."
        self.saveTo["command"] = self.getSaveToDirectory

        # Output filename
        self.outputLabel = tk.Label(self.nameBar, text="Save As", font=('Arial', 10, 'bold'))
        self.outputLabel.grid(row=2, column=0, sticky='w')
        self.outputName = tk.Entry(self.nameBar)
        self.outputName.grid(row=2, column=1, sticky='w')
        self.outputName.insert(0, "Testing")

        # Output tabname
        self.outputTabLabel = tk.Label(self.nameBar, text="Tab Name", font=('Arial', 10, 'bold'))
        self.outputTabLabel.grid(row=3, column=0, sticky='w')
        self.outTabName = tk.Entry(self.nameBar)
        self.outTabName.grid(row=3, column=1, sticky='w')
        self.outTabName.insert(0, "Tab 1")

        # Run program
        self.runButton = tk.Button(self.controls, text="Run", width=5, command=self.runProgram)
        self.runButton.grid(row=16, column=1, sticky='e')

        # Column Mappings
        self.sectionHead = tk.Label(self.colBar, text="Column Mappings", font=('Arial', 10, 'bold'))
        self.sectionHead.grid(row=4, column=0, sticky='w')

        self.fname_label = tk.Label(self.colBar, text="First name")
        self.fname_label.grid(row=5, column=0, sticky='w')

        self.lname_label = tk.Label(self.colBar, text="Last name")
        self.lname_label.grid(row=6, column=0, sticky='w')
        
        self.add1_label = tk.Label(self.colBar, text="Address Line 1")
        self.add1_label.grid(row=7, column=0, sticky='w')

        self.add2_label = tk.Label(self.colBar, text="Address Line 2")
        self.add2_label.grid(row=8, column=0, sticky='w')

        self.city_label = tk.Label(self.colBar, text="City")
        self.city_label.grid(row=9, column=0, sticky='w')

        self.state_label = tk.Label(self.colBar, text="State")
        self.state_label.grid(row=10, column=0, sticky='w')

        self.zip_label = tk.Label(self.colBar, text="Zip Code")
        self.zip_label.grid(row=11, column=0, sticky='w')

        self.spec_label = tk.Label(self.colBar, text="Specialty")
        self.spec_label.grid(row=12, column=0, sticky='w')

        self.lic_number_label = tk.Label(self.colBar, text="License Number")
        self.lic_number_label.grid(row=13, column=0, sticky='w')

        self.lic_state_label = tk.Label(self.colBar, text="License State")
        self.lic_state_label.grid(row=14, column=0, sticky='w')

        self.npi_label = tk.Label(self.colBar, text="NPI Number")
        self.npi_label.grid(row=15, column=0, sticky='w')

    def __init__(self, master=None):
        tk.Frame.__init__(self, master)
        self.grid()
        self.createWidgets()


if __name__=='__main__':
    root = tk.Tk()
    root.title('NPI Lookup')
    app = Application(master=root)
    app.mainloop()


# progress indicator 
        # self.progress = tk.Entry(root)
        # progressText = open('textcontent.txt')
        # lines = progressText.read()
        # progressText.close()
        # self.progress.pack(side='left', fill='x', padx=5)

        # text widget 
        # self.text1 = tk.Text(self.controls, height=10, width=40)
        # self.text1.pack(side='left', fill='x', padx=5)



