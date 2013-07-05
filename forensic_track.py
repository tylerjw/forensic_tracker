import xml.sax, xml.sax.handler

class XmlHandler(xml.sax.handler.ContentHandler):
    def __init__(self):
        self.buffer = ''
        self.map = {}

    def startElement(self, name, attributes):
        if name == 'ReportDateTime':
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
    filename = '20130702BG-4_ETISALAT_SIM_8537_1.xml'
    xmlfile = open(filename, 'r')
    parser = xml.sax.make_parser()
    handler = XmlHandler()
    parser.setContentHandler(handler)
    parser.parse(xmlfile)
    xmlfile.close
    print handler.map
