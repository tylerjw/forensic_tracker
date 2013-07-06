import xml.sax, xml.sax.handler
from Tkinter import Listbox,MULTIPLE,END,RIGHT,LEFT,StringVar,Scrollbar,E,W,BOTH
from ttk import Frame,Button,Label,Entry,Separator
import tkFileDialog
from xlwt import Workbook, easyxf

class Window(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.pack()
        self.master.title('Forensic Tracker Parser')
        self.master.iconname('Forensic Tracker Parser')

        self.out_filename = StringVar()
        self.out_filename.set('forensic_track.xls')
        self.poc = StringVar()
        self.error = StringVar()
        self.in_filenames = []

        gfm = Frame(self)
        Label(gfm, text="POC: ").grid(row=1,column=0,pady=5,sticky=W)
        Entry(gfm, textvariable=self.poc,width=40).grid(row=1,column=1,sticky=W)
        Button(gfm, text='Add File(s)', command=self.add,width=15).grid(row=1,column=2,padx=5,sticky=E)
        Button(gfm, text='Remove File(s)', command=self.remove,width=15).grid(row=1,column=3,sticky=E)
        lbf = Frame(gfm)
        self.input_lb = Listbox(lbf, width=80, height=5, selectmode='extended')
        scroll_lb = Scrollbar(lbf, command=self.input_lb.yview)
        self.input_lb.configure(yscrollcommand=scroll_lb.set)
        self.input_lb.pack(side=LEFT)
        self.master.bind('<Control-a>',self.select_all)
        scroll_lb.pack(side=RIGHT, fill="both")
        lbf.grid(pady=5,padx=5,row=2,column=0,columnspan=4,sticky=W)
        Label(gfm, text="Output: ", width=8).grid(row=3,column=0,pady=5,sticky=W)
        self.out_ent = Entry(gfm, width=60, textvariable=self.out_filename)
        self.out_ent.grid(row=3,column=1,columnspan=2,padx=5,sticky=W)
        Button(gfm, text="Browse", width=10, command=self.get_outfn).grid(row=3,column=3,padx=5,sticky=E)
        sep = Separator(gfm,orient="horizontal")
        sep.grid(row=4,column=0,columnspan=4,sticky="ew",pady=10)
        Button(gfm,text="Create Tracker",command=self.parse,width=15).grid(row=5,column=0,columnspan=2,sticky=W)
        Label(gfm,text='test error',textvariable=self.error,foreground='red').grid(row=5,column=2,columnspan=3,sticky="w")
        gfm.pack(padx=10,pady=10)

    def select_all(self, evt):
        self.input_lb.selection_set(0,self.input_lb.size()-1)

    def add(self):
        filenames = tkFileDialog.askopenfilenames(title='Select input .xml files', filetypes=[('xml','*.xml')],
                                              multiple=True)
        filenames = str(filenames).split(' ')
        for name in filenames:
            if name not in self.in_filenames:
                self.input_lb.insert(END, name)
                self.in_filenames.append(name)

    def remove(self):
        selection = list(self.input_lb.curselection())
        selection.reverse()
        if selection != '':
            for index in selection:
                self.input_lb.delete(int(index))
                self.in_filenames.pop(int(index))
        
    def get_outfn(self):
        filename = tkFileDialog.asksaveasfilename(title='Output file', initialfile='forensic_track.xls',
                                                defaultextension='.xls', filetypes=[('xls','*.xls')])
        self.out_filename.set(filename)
        
    def parse(self):
        if self.poc.get() == u'':
            self.error.set("POC not set")
            return #don't go any further
        
        bk = Workbook()
        sh = bk.add_sheet('Sheet 1')
        row = 0
        widths = [15, 20, 18, 3, 20, 10, 12, 30]
        for i,w in enumerate(widths):
            sh.col(i).width = 256*w
        data_style = easyxf('font: height 256;'
                            'pattern: pattern solid_fill, fore_colour green;'
                            'alignment: wrap True, vertical top;'
                            'borders: left thin, right thin, top thin, bottom thin')
        empty_style = easyxf('borders: left thin, right thin, top thin, bottom thin')
        for filename in self.in_filenames:
            try:
                xmlfile = open(filename, 'r')
            except:
                self.error.set("Error opening file: {0}".format(row+1))
                break
            parser = xml.sax.make_parser()
            handler = XmlHandler()
            parser.setContentHandler(handler)
            try:
                parser.parse(xmlfile)
            except:
                self.error.set("Error parsing file: {0}".format(row+1))
                break
            xmlfile.close
            sh.write(row,0,handler.map['AcquireDate'],data_style)
            sh.write(row,1,handler.map['Owner'],data_style)
            sh.write(row,2,handler.map['BatchId'],data_style)
            sh.write(row,3,'',data_style)
            sh.write(row,4,handler.map['AcquiringUnit'],data_style)
            sh.write(row,5,str(self.poc.get()).strip(),data_style)
            sh.write(row,6,':'+handler.map['IMSI'],data_style)
            sh.write(row,7,'1X SIM CARD '+handler.map['ServiceProvider']+
                     ' ICCID:'+handler.map['ICCID'][-4:],data_style)
            for i in range(8):
                sh.write(row+1,i,'',empty_style)
            row += 2 #add a blank line inbetween

        try:
            bk.save(self.out_filename.get())
        except:
            self.error.set("Error writing file!")
            return
        self.error.set("Success!")
        

class XmlHandler(xml.sax.handler.ContentHandler):
    def __init__(self):
        self.buffer = ''
        self.map = {}
        for tag in ['AcquireDate', 'BatchId', 'AcquiringUnit', 'IMSI', 'ICCID', 'ServiceProvider',
                'Owner', 'MGRSLocation']:
            self.map[tag] = ''

    def startElement(self, name, attributes):
        if name == 'AcquireDate':
            self.map[str(name)] = attributes.getValue('RawDateTime')
        

    def characters(self, data):
        self.buffer += data

    def endElement(self, name):
        tags = ['BatchId', 'AcquiringUnit', 'IMSI', 'ICCID', 'ServiceProvider',
                'Owner', 'MGRSLocation']
        if name in tags:
            self.map[str(name)] = str(self.buffer).strip()

        self.buffer = ''

if __name__ == '__main__':
    Window().mainloop()
