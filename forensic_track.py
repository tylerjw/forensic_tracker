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
        self.error.set('test error')
        self.in_filenames = []

        gfm = Frame(self)
        Label(gfm, text="POC: ").grid(row=1,column=0,pady=5,sticky=W)
        Entry(gfm, textvariable=self.poc,width=40).grid(row=1,column=1,sticky=W)
        Button(gfm, text='Add File(s)', command=self.add,width=15).grid(row=1,column=2,padx=5,sticky=E)
        Button(gfm, text='Remove File(s)', command=self.remove,width=15).grid(row=1,column=3,sticky=E)
        lbf = Frame(gfm)
        self.input_lb = Listbox(lbf, width=80, height=5, selectmode=MULTIPLE)
        scroll_lb = Scrollbar(lbf, command=self.input_lb.yview)
        self.input_lb.configure(yscrollcommand=scroll_lb.set)
        self.input_lb.pack(side=LEFT)
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

    def add(self):
        filenames = tkFileDialog.askopenfilenames(title='Select input .xml files', filetypes=[('xml','*.xml')],
                                              multiple=True)
        filenames = str(filenames).split(' ')
        for name in filenames:
            if name not in self.in_filenames:
                self.input_lb.insert(END, name)
                self.in_filenames.append(name)

    def remove(self):
        print "remove"
        selection = list(self.input_lb.curselection())
        selection.reverse()
        print selection
        if selection != '':
            for index in selection:
                self.input_lb.delete(int(index))
                self.in_filenames.pop(int(index))
        
    def get_outfn(self):
        filename = tkFileDialog.asksaveasfilename(title='Output file', initialfile='forensic_track.xls',
                                                defaultextension='.xls', filetypes=[('xls','*.xls')])
        self.out_filename.set(filename)
        
    def parse(self):
        print "parse"
        bk = Workbook()
        sh = bk.add_sheet('Sheet 1')
        row = 0
        widths = [8, 20, 14, 3, 18, 10, 16, 27]
        for i,w in enumerate(widths):
            sh.col(i).width = 256*w
        color_style = easyxf('pattern: pattern solid_fill, fore_colour green')
        border_style = easyxf('borders: left thin, right thin, top thin, bottom thin')
        for filename in self.in_filenames:
            xmlfile = open(filename, 'r')
            parser = xml.sax.make_parser()
            handler = XmlHandler()
            parser.setContentHandler(handler)
            parser.parse(xmlfile)
            xmlfile.close
            print handler.map
            '''for i in range(8):
                sh.row(row).(row,i,'',color_style)
                sh.write(row,i,'',border_style)
                sh.write(row+1,i,'',border_style)'''
            sh.write(row,0,handler.map['ReportDate'])
            sh.write(row,1,handler.map['Owner'])
            sh.write(row,2,handler.map['BatchId'])
            sh.write(row,3,'')
            sh.write(row,4,handler.map['AcquiringUnit'])
            sh.write(row,5,str(self.poc.get()).strip())
            sh.write(row,6,':'+handler.map['IMSI'])
            sh.write(row,7,'1X SIM CARD '+handler.map['ServiceProvider']+
                     ' ICCID:'+handler.map['ICCID'][-4:])
            row += 2 #add a blank line inbetween
            
        bk.save(self.out_filename.get())

class XmlHandler(xml.sax.handler.ContentHandler):
    def __init__(self):
        self.buffer = ''
        self.map = {}
        for tag in ['ReportDate', 'BatchId', 'AcquiringUnit', 'IMSI', 'ICCID', 'ServiceProvider',
                'Owner', 'MGRSLocation']:
            self.map[tag] = ''

    def startElement(self, name, attributes):
        if name == 'ReportDate':
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
