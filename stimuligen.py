from Tkinter import *
import tkFileDialog

import xlrd
import csv
import os

class MainWindow:

    def __init__(self, master):

        self.root = master                  # main GUI context
        self.root.title("stimuligen")       # title of window
        self.root.geometry("600x400")       # size of GUI window
        self.main_frame = Frame(root)       # main frame into which all the Gui components will be placed
        self.main_frame.pack()              # pack() basically sets up/inserts the element (turns it on)

        self.datasource_template_file = None
        self.datasource_template_book = None
        self.datasource_template_sheet = None

        self.eyetracking_order_file = None
        self.eyetracking_orders = []

        self.pair_carrier_orders_file = None
        self.pair_carrier_orders_book = None
        self.pair_carrier_orders_sheet = None

        self.load_datasource_button = Button(self.main_frame,
                                             text="Load Datasource",
                                             command=self.load_datasource)

        self.load_eyetracking_order_button = Button(self.main_frame,
                                                    text="Load Eyetracking Order",
                                                    command=self.load_eyetracking_order)

        self.clear_button = Button(self.main_frame,
                                   text="Clear",
                                   command=self.clear)

        self.load_pair_carrier_orders_button = Button(self.main_frame,
                                                      text="Load Pair Carrier Orders",
                                                      command=self.load_pair_carrier_orders)

        self.video_loaded_label = Label(self.main_frame, text="video loaded", fg="blue")
        self.timestamps_loaded_label = Label(self.main_frame, text="timestamps loaded", fg="red")

        self.load_datasource_button.grid(row=1, column=1)
        self.load_eyetracking_order_button.grid(row=1, column=3)
        self.load_pair_carrier_orders_button.grid(row=2, column=2)
        self.clear_button.grid(row=3, column=2)

    def load_datasource(self):

        self.datasource_template_file = tkFileDialog.askopenfilename()

        self.datasource_template_book = xlrd.open_workbook(self.datasource_template_file)
        self.datasource_template_sheet = self.datasource_template_book.sheet_by_index(0)

    def load_eyetracking_order(self):

        self.eyetracking_order_file = tkFileDialog.askopenfilename()

        with open(self.eyetracking_order_file, "rU") as file:
            csvreader = csv.reader(file, delimiter='\t')
            csvreader.next()
            for row in csvreader:
                self.eyetracking_orders.append([row[0],      # SubID
                                                row[1],      # Month
                                                int(row[2]), # Order
                                                row[3],      # carrier_order_half
                                                row[4],      # Past
                                                row[5]])     # Visit

            print self.eyetracking_orders

    def load_pair_carrier_orders(self):

        self.pair_carrier_orders_file = tkFileDialog.askopenfilename()

        self.pair_carrier_orders_book = xlrd.open_workbook(self.pair_carrier_orders_file)
        self.pair_carrier_orders_sheet = self.pair_carrier_orders_book.sheet_by_index(0)

    def clear(self):

        self.datasource_template_file = None
        self.eyetracking_order_file = None
        self.pair_carrier_orders_file = None

        self.eyetracking_orders = []

        self.datasource_template_book = None
        self.pair_carrier_orders_book = None

        self.datasource_template_sheet = None
        self.pair_carrier_orders_sheet = None

        if self.video_loaded_label:
            self.video_loaded_label.grid_remove()
        if self.timestamps_loaded_label:
            self.timestamps_loaded_label.grid_remove()

    # def load_template(self):
    #
    #     self.template_file = tkFileDialog.askopenfilename()
    #
    #     self.template_book = xlrd.open_workbook(self.template_file)
    #     self.template_sheet = self.template_book.sheet_by_index(0)
    #
    #     self.template_loaded_label.grid(row=3, column=1)
    #
    # def load_stimuli(self):
    #
    #     self.stimuli_file = tkFileDialog.askopenfilename()
    #
    #     self.stimuli_book = xlrd.open_workbook(self.stimuli_file)
    #     self.stimuli_sheet = self.stimuli_book.sheet_by_index(0)
    #
    #     self.stimuli_loaded_label.grid(row=3, column=2)
    #
    # def run(self):
    #
    #
    #
    #     #0) keep the header row, ...
    #     hdr = []
    #     for c in range(13):
    #         hdr.append(self.template_sheet.cell_value(0, c))
    #     self.data.append(hdr)
    #     print hdr
    #     #get column G's value
    #     order = self.stimuli_sheet.cell_value(1, 6)
    #
    #     #find the first row corresponding to the G subset
    #     row = 1
    #     while True:
    #         if self.template_sheet.cell_value(row, 12) == order:
    #             break
    #         row += 1
    #
    #     #0) ..., and then take the subset of the template that is the order corresponding to column G's value (1-4)
    #     while self.template_sheet.cell_value(row, 12) == order:
    #         #print "row: " + str(row) + "  template value: " + str(template_sheet.cell_value(row, 12)) + "  order: " + str(order)
    #         row_data = []
    #         for col in range(13):
    #             val = self.template_sheet.cell_value(row, col)
    #             if self.template_sheet.cell_type(row, col) == 2:
    #                 row_data.append(int(val))
    #             else:
    #                 row_data.append(val)
    #         self.data.append(row_data)
    #         #print "_dimnrows: " + str(template_sheet._dimnrows)
    #
    #         if row >= self.template_sheet._dimnrows - 1:
    #             break
    #         else:
    #             row += 1
    #
    #     #1) Use the 20 rows (after the header row) in columns A and B to write into columns B through D of the spreadsheet
    #     #3) replace 1-16 in the .wav and .jpg with the words numbered 1-16 (e.g. 1.jpg becomes banana.jpg and can_banana.jpg)
    #     #4) IF there is something in column e that is not NA, replace with that instead of with the word in column B (e.g. sock3 instead of sock) ONLY in columns B&C not in column D
    #     for r in range(5, len(self.data)):
    #         index = int(self.data[r][1].split('.')[0])
    #         col_e = self.stimuli_sheet.cell_value(index + 4, 4)
    #         if col_e == "NA":
    #             self.data[r][1] = self.stimuli_sheet.cell_value(index + 4, 1) + ".jpg"
    #         else:
    #             self.data[r][1] = self.stimuli_sheet.cell_value(index + 4, 4) + ".jpg"
    #
    #         index = int(self.data[r][2].split('.')[0])
    #         col_e = self.stimuli_sheet.cell_value(index + 4, 4)
    #         if col_e == "NA":
    #             self.data[r][2] = self.stimuli_sheet.cell_value(index + 4, 1) + ".jpg"
    #         else:
    #             self.data[r][2] = self.stimuli_sheet.cell_value(index + 4, 4) + ".jpg"
    #
    #         prefix = self.data[r][3].split('.')[0].split('_')[0]
    #         index = int(self.data[r][3].split('.')[0].split('_')[1])
    #         self.data[r][3] = "%s_%s.wav" % (prefix, self.stimuli_sheet.cell_value(index + 4, 1))
    #
    #     #5) replace A:H in column F with the pairs corresponding to A:H in column I of the stimuli spreadsheet
    #         self.data[r][5] = self.stimuli_sheet.cell_value(ord(self.data[r][5]) - ord('A') + 1, 8)
    #
    #
    #     #2) replace practice1.jpg-practice4.jpg with the first four words of 'stimuli' labeled p1-p4
    #     for r in range(1, 5):
    #         word = self.stimuli_sheet.cell_value(r, 1)
    #         self.data[r][1] = self.data[r][1].replace("practice%d" % r, word)
    #         self.data[r][3] = self.data[r][3].replace("practice%d" % r, word)
    #
    #
    #
    # def export(self):
    #
    #     self.run()
    #
    #     self.export_file = tkFileDialog.asksaveasfilename()
    #
    #     with open(self.export_file, 'w') as file:
    #
    #         csvwriter = csv.writer(file, delimiter='\t')
    #
    #         csvwriter.writerow(self.data[0])     # write the header row
    #
    #         for row in self.data[1:]:    # write each subsequent row (skipping the header)
    #             csvwriter.writerow(row)

if __name__ == "__main__":

    root = Tk()
    MainWindow(root)
    root.mainloop()
