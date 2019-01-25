from xml.dom.minidom import parse
import xml.dom.minidom
import xlsxwriter
import os

class GpxHandler():
    def __init__(self, pathin ,pathout):
        self.pathin = pathin
        self.pathout = pathout

    def generate(self, infile, outfile):
        DOMTree = xml.dom.minidom.parse(infile)
        gpx = DOMTree.documentElement
        trk = gpx.getElementsByTagName('trk')[0]
        trkseg = trk.getElementsByTagName('trkseg')[0]
        trkpts = trkseg.getElementsByTagName('trkpt')
        wbk = xlsxwriter.Workbook(outfile)
        sheet = wbk.add_worksheet('sheet 1')
        n = 0
        sheet.write(n, 0, 'lat')
        sheet.write(n, 1, 'lon')
        sheet.write(n, 2, 'ele')
        sheet.write(n, 3, 'time')
        n = n + 1
        for trkpt in trkpts:
            lat = trkpt.getAttribute('lat')
            lon = trkpt.getAttribute('lon')
            ele = trkpt.getElementsByTagName('ele')[0].firstChild.data
            time = trkpt.getElementsByTagName('time')[0].firstChild.data
            sheet.write(n, 0, lat)
            sheet.write(n, 1, lon)
            sheet.write(n, 2, ele)
            sheet.write(n, 3, time)
            n = n + 1
        wbk.close()

    def dealGpx(self):
        if not os.path.isdir(self.pathout):
            os.makedirs(self.pathout)
        for root, dirs, files in os.walk(self.pathin):
            for file in files:
                if os.path.splitext(file)[1] == '.gpx':
                    filename = os.path.splitext(file)[0]
                    self.generate(self.pathin + '\\' + file, self.pathout + '\\' + filename + r'.xlsx')
        print("Figured out.")

if __name__ == '__main__':
    pathin = "pathin"
    pathout = "pathout"
    gpxHandler = GpxHandler(pathin, pathout)
    gpxHandler.dealGpx()

