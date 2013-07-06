from Tkinter import StringVar,END
from ttk import Frame,Button,Label,Entry
from ScrolledText import ScrolledText
import xml.sax, xml.sax.handler
import tkFileDialog
import os
import re

class Window(Frame):
    def __init__(self):
        Frame.__init__(self)
        self.pack()
        self.master.title('Filename fixer')
        self.master.iconname('Filename fixer')

        self.directory = StringVar()
        self.error = StringVar()
        self.pairs = {}
        self.files = []

        f = Frame(self)
        Label(f,text="Directory: ",width=12).grid(row=0,column=0,sticky='e')
        Entry(f,textvariable=self.directory,width=50).grid(row=0,column=1,sticky='ew')
        Button(f,text="Browse",width=8,command=self.get_directory).grid(row=0,column=2,padx=2,pady=2)
        Button(f,text="Run Script",width=12,command=self.script).grid(row=1,column=1,padx=12,pady=4)

        self.stext = ScrolledText(f,height=15,width=60,wrap='word')
        self.stext.grid(row=2,column=0,columnspan=3,sticky='ew')
        f.pack(padx=10,pady=10)

    def get_directory(self):
        directory = tkFileDialog.askdirectory(title="Select Directory with misnamed files",mustexist=True)
        self.directory.set(directory)
        self.load_files()

    def load_files(self):
        self.files = os.listdir(self.directory.get())
        self.clean_files()
        output = "Found {0} pairs in directory.\n".format(len(self.pairs))
        for key in self.pairs:
            output += key + '\n'
            output += ' - ' + self.pairs[key] + '\n'
        output += '\n'
        self.stext.insert(END,output)
        self.stext.yview(END)

    def clean_files(self):
        #remove all files not .xml or .xls
        ext_re = re.compile('.xml$|.xls$')
        bad = []
        for i,name in enumerate(self.files):
            if not ext_re.search(name):
                bad.append(i)
        bad.reverse()
        for i in bad:
            self.files.pop(i)

        #build the pairs dict
        self.pairs = {}
        good_keys = []
        xls_re = re.compile('.xls$')
        for name in self.files:
            if xls_re.search(name):
                key = name[:-4] + '.xml'
                self.pairs[key] = name
            else: #xml
                good_keys.append(name)

        #delete the bad keys
        del_keys = []
        for key in self.pairs:
            if key not in good_keys:
                del_keys.append(key)

        for key in del_keys:
            del self.pairs[key]
        
    def script(self):
        #find the iccid associated with the files
        iccid_dict = {}
        os.chdir(self.directory.get())

        for key in self.pairs:
            try:
                xmlfile = open(key, 'r')
            except:
                self.stext.insert(END,"Error opening file: " + key + '\n\n')
                self.stext.yview(END)
                break
            parser = xml.sax.make_parser()
            handler = XmlHandler()
            parser.setContentHandler(handler)
            try:
                parser.parse(xmlfile)
            except:
                self.stext.insert(END,"Error opening file: " + key + '\n\n')
                self.stext.yview(END)
                break
            xmlfile.close()
            iccid_dict[key] = handler.iccid

        output = "ICCID values found:\n"
        for key in iccid_dict:
            output += key + ' : ' + iccid_dict[key][-4:] + '\n'
        output += '\n'
        self.stext.insert(END,output)
        self.stext.yview(END)

        #rename files
        enc_re = re.compile('_\d\d\d\d[.-]')
        output = ''
        for key,xls in self.pairs.items():
            src = key
            res = enc_re.search(key)
            if not res:
                output += "Error, file does not match expected format:\n\t" + key + "\n\n"
                continue #not in the recognized format
            span = res.span()
            dest = key[:span[0]+1] + iccid_dict[key][-4:] + '.xml'

            output += src + ' --> \n' + dest + '\n\n'
            try:
                os.rename(src,dest)
            except:
                output += "Error renaming file: " + src + "\n\n"

            src = xls
            res = enc_re.search(xls)
            if not res:
                output += "Error, file does not match expected format:\n\t" + key + "\n\n"
                continue #not in the recognized format
            span = res.span()
            dest = xls[:span[0]+1] + iccid_dict[key][-4:] + '.xls'

            output += src + ' --> \n' + dest + '\n\n'
            try:
                os.rename(src,dest)
            except:
                output += "Error renaming file: " + src + "\n\n"

        self.stext.insert(END,output)
        self.stext.yview(END)
        self.load_files()
        
        
class XmlHandler(xml.sax.handler.ContentHandler):
    def __init__(self):
        self.buffer = ''
        self.iccid = ''

    def characters(self, data):
        self.buffer += data

    def endElement(self, name):
        if name == 'ICCID':
            self.iccid = str(self.buffer).strip()
        self.buffer = ''
        
if __name__ == '__main__':
    Window().mainloop()
