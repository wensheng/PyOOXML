# -*- coding: utf-8 -*-

# Copyright (c) 2011 Wensheng Wang
# MIT license

import os
import sys
from ooxmlbase import OFile, OOXMLBase, tmpl_dir
if sys.version_info[0]==3:
    import io
    BytesIO = io.BytesIO
else:
    import StringIO
    BytesIO = StringIO.StringIO

def al2d(s):
    col = 0
    row = 0
    for c in s:
        v = ord(c)
        if v>64:
            col = col*26+v-64
        else:
            row = row*10+v-48
    return (row,col)

class Workcell(object):
    def __init__(self,x,y,v):
        self._x=x
        self._y=y
        self._value = v

    @property
    def x(self):
        return self._x

    @property
    def y(self):
        return self._y

    def get_value(self):
        "get cell value"
        return self._value

    def set_value(self,v):
        self._value = v
        
    value = property(get_value, set_value)

class Workrow(object):
    def __init__(self,id,span,cells={}):
        self._id = id
        self._span = []
        self._cell = cells
        ss = span.split()
        for s in ss:
            r = s.split(':')
            self._span.append((int(r[0]),int(r[1])))

    @property
    def id(self):
        return self._id

    def cell(self,col):
        return self._cell.get(col)

    def values(self):
        v = []
        for i in range(1,self._span[-1][1]+1):
            if self.cell(i):
                v.append(self.cell(i).value)
            else:
                v.append(None)
        return v
        

class Worksheet(object):
    def __init__(self,name,workbook,path):
        self.name = name
        self.workbook = workbook
        self.sheet = None
        self.path = path
        self._row = {}
        self._cell = {}
        self.top, self.left, self.bottom, self.right = 0,0,0,0

    def get_sheet(self):
        if self.workbook.package:
            self.sheet = OFile(self.workbook.package.read(self.path))
            dimension = self.sheet.tree.find('.//%sdimension'%self.sheet.ns).get('ref').split(':')  
            if len(dimension)==2:
                self.top, self.left = al2d(dimension[0])
                self.bottom, self.right = al2d(dimension[1])
        return self

    def _get_cells(self):
        if self._row or self._cell:
            return
        rows = self.sheet.tree.findall('.//%srow'%self.sheet.ns) 
        for row in rows:
            rowid = int(row.get('r'))
            rowcell = {}
            for c in row:
                t = None
                for v in c:
                    if v.tag==(self.sheet.ns+'v'):
                        if c.get('t')=='s':
                            t=self.workbook.s_string[int(v.text)]
                        else:
                            t = v.text
                        break
                xy = al2d(c.get('r'))
                self._cell[xy]=Workcell(xy[0],xy[1],t)
                rowcell[xy[1]] = self._cell[xy]
            #self._row[rowid] = Workrow(rowid,row.get('spans'),rowcell)
            span = row.get('spans','%s:%s'%(self.left,self.right))
            self._row[rowid] = Workrow(rowid,span,rowcell)
             

    def row(self,n):
        if not self._row:
            self._get_cells()
        return self._row.get(n)

    def row_cell_values(self,n):
        pass

    def cell(self,x,y):
        if not self._cell:
            self._get_cells()
        return self._cell.get((x,y))

    def set_cell(self,x,y,v):
        if self._cell.get((x,y)):
            self._cell[(x,y)].value = v
        else:
            self._cell[(x,y)]=Workcell(x,y,v)

    def rows(self):
        if not self._row:
            self._get_cells()
        return self._row

    @property
    def nrows(self):
        return len(self.rows())

    def row_iter(self):
        i = 1
        while i<=self.bottom:
            yield self.row(i) 
            i += 1

    def row_values_iter(self):
        i = 1
        while i<=self.bottom:
            yield self.row(i).values()
            i += 1 

    def save_csv(self,filename):
        import csv
        f = open(filename,'wb')
        writer = csv.writer(f)
        writer.writerows(self.row_values_iter())
        f.close()

class Spreadsheet(OOXMLBase):
    def __init__(self, filename=None):
        self.filename = filename
        if not filename:
            filename = os.path.join(tmpl_dir,'workbook.xlsx')
            self._is_new = True
        else:
            self._is_new = False
        super(Spreadsheet,self).__init__(filename)
        self._sheet={}
        self._sheet_names={}
        self._process_workbook()

    def _get_part_relationship(self):
        part_path = self.package_relationship['officeDocument'].split('/')
        self.docpath = part_path[0]
        doc = OFile(self.package.read(self.docpath+'/_rels/'+part_path[1]+".rels"))
        self.part_relationship = {}
        for elm in doc.tree:
            self.part_relationship[elm.get('Id')]=(elm.get('Type'),elm.get('Target'))
            if elm.get('Type') and elm.get('Type')=="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings":
                self._process_shared_strings(elm.get('Target'))

    def _process_shared_strings(self,target):
        self.s_string = []
        doc = OFile(self.package.read(self.docpath+'/'+target))
        for elm in doc.tree.getiterator("%st"%doc.ns):
            self.s_string.append(elm.text)

    def _process_workbook(self):
        doc = OFile(self.package.read(self.package_relationship['officeDocument']))
        for elm in doc.tree.findall('.//%ssheet'%doc.ns):
            id = int(elm.get('sheetId'))
            name = elm.get('name')
            #Hack: this doesn't work in ElementTree because it doesn't have nsmap, use explict QFN  
            #self._sheet[id] = Worksheet(name,self,(self.docpath+'/'+self.part_relationship[elm.get('{%s}id'%doc.tree.nsmap['r'])][1]))
            self._sheet[id] = Worksheet(name,self,(self.docpath+'/'+self.part_relationship[elm.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")][1]))
            self._sheet_names[name]=id

    def sheet(self,id):
        if type(id)==type(""):
            return self._sheet[self._sheet_names[id]].get_sheet()
        else:
            return self._sheet[id].get_sheet()

    def add_sheet(self,name):
        id = len(self._sheet_names)+1
        if not name:
            name = 'Sheet'+id
        self._sheet_names[name]=id
        if not self.docpath:
            self.docpath = 'xl'
            #do lots of other stuff
        self._sheet[id] = Worksheet(name,self,"") #fixme

    @property
    def sheet_names(self):
        return self._sheet_names.keys()

    def save(self,filename):
        bytesio = BytesIO()
        import xml.etree.cElementTree as ET
        ET.ElementTree(self.sheet(1).sheet.tree).write(bytesio)

        if self.filename:
            self.package.close()
     
        import zipfile
        zip = zipfile.ZipFile(filename or self.filename or "workbook.xlsx",'w',zipfile.ZIP_DEFLATED)
        zip.writestr('test',bytesio.getvalue())
        bytesio.close()
        zip.close()
 

if "__main__"==__name__:
    import sys
    workbook = Spreadsheet(sys.argv[1])
    #workbook.sheet(1).save_csv(sys.argv[2])
    #workbook.save(sys.argv[3])
    #pass
