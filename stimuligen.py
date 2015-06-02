from Tkinter import *
import tkFileDialog

from openpyxl import *

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

        self.generate_stimuli_button = Button(self.main_frame,
                                              text='Generate Stimuli',
                                              command=self.generate_stimuli)


        self.video_loaded_label = Label(self.main_frame, text="video loaded", fg="blue")
        self.timestamps_loaded_label = Label(self.main_frame, text="timestamps loaded", fg="red")

        self.load_datasource_button.grid(row=1, column=1)
        self.load_eyetracking_order_button.grid(row=1, column=3)
        self.load_pair_carrier_orders_button.grid(row=2, column=2)
        self.generate_stimuli_button.grid(row=2, column=3)
        self.clear_button.grid(row=3, column=2)

    def load_datasource(self):

        self.datasource_template_file = tkFileDialog.askopenfilename()

        self.datasource_template_book = load_workbook(self.datasource_template_file)

        self.datasource_template_sheet = self.datasource_template_book.active


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

        self.pair_carrier_orders_book = load_workbook(self.pair_carrier_orders_file)
        self.pair_carrier_orders_sheet = self.pair_carrier_orders_book.active

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

    def generate_stimuli(self):
        """
        This is a really stupid/explicit implementation, just to get
        it to working condition. Everything is basically hard-coded.
        Too lazy to think about how to condense it into loops (at least for now)
        :return:
        """
        header = [u'number',
                  u'word',
                  u'kind',
                  u'carrier',
                  u'duplicate_image_filename',
                  u'',
                  u'order',
                  u'pair',
                  u'pair_words',
                  u'pair_kind',
                  u'carrier']

        # this is wrong
        header_block = [[u'p1', u'', u'practice', u'can', u'NA', u'', u'',	u'A', u'banana_kitty', u'generic', u'can'],
                        [u'p2', u'', u'practice', u'where', u'NA', u'', u'', u'B', u'bear_cracker', u'generic', u'do'],
                        [u'p3', u'', u'practice', u'do', u'NA', u'', u'', u'C', u'cheerios_water', u'generic', u'look'],
                        [u'p4', u'', u'practice', u'look', u'NA', u'', u'',	u'D', u'hair_cup', u'generic', u'where']]
        start8_z = 2
        start10_z = 10
        start12_z = 18
        start14_z = 26
        start16_z = 34
        start18_z = 42

        start8_y = 50
        start10_y = 58
        start12_y = 66
        start14_y = 74
        start16_y = 82
        start18_y = 90

        start8_z_uniq = 98
        start10_z_uniq = 106
        start12_z_uniq = 114
        start14_z_uniq = 122
        start16_z_uniq = 130
        start18_z_uniq = 138

        start8_y_uniq = 146
        start10_y_uniq = 154
        start12_y_uniq = 162
        start14_y_uniq = 170
        start16_y_uniq = 178
        start18_y_uniq = 186

        regions8 = [start8_z, start8_y, start8_z_uniq, start8_y_uniq]
        regions10 = [start10_z, start10_y, start10_z_uniq, start10_y_uniq]
        regions12 = [start12_z, start12_y, start12_z_uniq, start12_y_uniq]
        regions14 = [start14_z, start14_y, start14_z_uniq, start14_y_uniq]
        regions16 = [start16_z, start16_y, start16_z_uniq, start16_y_uniq]
        regions18 = [start18_z, start18_y, start18_z_uniq, start18_y_uniq]


        for entry in self.eyetracking_orders:
            if entry[4] == "past":
                continue
            else:
                wb = Workbook()
                ws = wb.active
                ws.append(header)           # write header
                ws['H2'] = 'A'              # write pairs
                ws['H3'] = 'B'
                ws['H4'] = 'C'
                ws['H5'] = 'D'
                ws['H6'] = 'E'
                ws['H7'] = 'F'
                ws['H8'] = 'G'
                ws['H9'] = 'H'

                # write number p1-p4
                ws['A2'] = 'p1'
                ws['A3'] = 'p2'
                ws['A4'] = 'p3'
                ws['A5'] = 'p4'

                # write number 1-16
                for i in range(1, 17):
                    ws['A{}'.format(6+(i-1))] = i

                # write "practice" in C2-C5
                for j in range(4):
                    ws['C{}'.format(2+j)] = "practice"
                # write "generic" in C6-C13
                for k in range(8):
                    ws['C{}'.format(6+k)] = 'generic'

                if entry[1] == '08':
                    if entry[3] == 'Z':

                        # write generic Z words
                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start8_z+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start8_z+l)].value
                        # write generic Z pair_words
                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start8_z)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start8_z+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start8_z+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start8_z+6)].value

                        # write generic Z pair carriers
                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start8_z)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start8_z+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start8_z+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start8_z+6)].value

                        # write unique Z pair carriers
                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start8_z_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start8_z_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start8_z_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start8_z_uniq+6)].value


                    else:

                        # write generic Y words
                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start8_y+l)].value

                        # write generic Y carriers
                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start8_y+l)].value

                        # write generic Y pair_words
                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start8_y)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start8_y+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start8_y+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start8_y+6)].value

                        # write generic Y pair carriers
                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start8_y)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start8_y+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start8_y+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start8_y+6)].value

                        # write unique Y pair carriers
                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start8_y_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start8_y_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start8_y_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start8_y_uniq+6)].value

                if entry[1] == '10':
                    if entry[3] == 'Z':

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start10_z+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start10_z+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start10_z)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start10_z+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start10_z+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start10_z+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start10_z)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start10_z+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start10_z+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start10_z+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start10_z_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start10_z_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start10_z_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start10_z_uniq+6)].value
                    else:

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start10_y+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start10_y+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start10_y)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start10_y+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start10_y+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start10_y+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start10_y)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start10_y+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start10_y+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start10_y+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start10_y_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start10_y_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start10_y_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start10_y_uniq+6)].value

                if entry[1] == '12':
                    if entry[3] == 'Z':

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start12_z+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start12_z+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start12_z)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start12_z+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start12_z+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start12_z+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start12_z)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start12_z+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start12_z+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start12_z+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start12_z_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start12_z_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start12_z_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start12_z_uniq+6)].value
                    else:

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start12_y+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start12_y+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start12_y)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start12_y+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start12_y+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start12_y+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start12_y)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start12_y+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start12_y+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start12_y+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start12_y_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start12_y_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start12_y_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start12_y_uniq+6)].value

                if entry[1] == '14':
                    if entry[3] == 'Z':

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start14_z+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start14_z+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start14_z)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start14_z+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start14_z+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start14_z+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start14_z)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start14_z+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start14_z+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start14_z+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start14_z_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start14_z_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start14_z_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start14_z_uniq+6)].value
                    else:

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start14_y+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start14_y+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start14_y)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start14_y+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start14_y+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start14_y+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start14_y)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start14_y+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start14_y+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start14_y+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start14_y_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start14_y_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start14_y_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start14_y_uniq+6)].value

                if entry[1] == '16':
                    if entry[3] == 'Z':

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start16_z+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start16_z+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start16_z)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start16_z+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start16_z+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start16_z+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start16_z)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start16_z+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start16_z+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start16_z+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start16_z_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start16_z_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start16_z_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start16_z_uniq+6)].value

                    else:

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start16_y+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start16_y+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start16_y)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start16_y+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start16_y+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start16_y+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start16_y)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start16_y+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start16_y+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start16_y+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start16_y_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start16_y_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start16_y_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start16_y_uniq+6)].value

                if entry[1] == '18':
                    if entry[3] == 'Z':

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start18_z+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start18_z+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start18_z)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start18_z+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start18_z+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start18_z+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start18_z)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start18_z+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start18_z+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start18_z+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start18_z_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start18_z_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start18_z_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start18_z_uniq+6)].value
                    else:

                        for l in range(8):
                            ws['B{}'.format(6+l)] = self.pair_carrier_orders_sheet['A{}'.format(start18_y+l)].value

                        for l in range(8):
                            ws['D{}'.format(6+l)] = self.pair_carrier_orders_sheet['G{}'.format(start18_y+l)].value

                        ws['I2'] = self.pair_carrier_orders_sheet['C{}'.format(start18_y)].value
                        ws['I3'] = self.pair_carrier_orders_sheet['C{}'.format(start18_y+2)].value
                        ws['I4'] = self.pair_carrier_orders_sheet['C{}'.format(start18_y+4)].value
                        ws['I5'] = self.pair_carrier_orders_sheet['C{}'.format(start18_y+6)].value

                        ws['K2'] = self.pair_carrier_orders_sheet['G{}'.format(start18_y)].value
                        ws['K3'] = self.pair_carrier_orders_sheet['G{}'.format(start18_y+2)].value
                        ws['K4'] = self.pair_carrier_orders_sheet['G{}'.format(start18_y+4)].value
                        ws['K5'] = self.pair_carrier_orders_sheet['G{}'.format(start18_y+6)].value

                        ws['K6'] = self.pair_carrier_orders_sheet['G{}'.format(start18_y_uniq)].value
                        ws['K7'] = self.pair_carrier_orders_sheet['G{}'.format(start18_y_uniq+2)].value
                        ws['K8'] = self.pair_carrier_orders_sheet['G{}'.format(start18_y_uniq+4)].value
                        ws['K9'] = self.pair_carrier_orders_sheet['G{}'.format(start18_y_uniq+6)].value


                ws['G2'] = entry[2]         # write order
                wb.save("output/{}_stimuli.xlsx".format(entry[5]))  # export xlsx file

if __name__ == "__main__":

    root = Tk()
    MainWindow(root)
    root.mainloop()
