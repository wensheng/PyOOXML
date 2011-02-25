# -*- coding: utf-8 -*-

# Copyright (c) 2011 Wensheng Wang
# MIT license

from ooxmlbase import OFile, OOXMLBase


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

class Worksheet(object):
    def __init__(self,name,workbook,path):
        self.name = name
        self.workbook = workbook
        self.path = path
        self._row = {}
        self._cell = {}

    def get_sheet(self):
        self.sheet = OFile(self.workbook.package.read(self.path))
        return self

    def _get_cells(self):
        if self._row or self._cell:
            return
        rows = self.sheet.tree.findall('.//%srow'%self.sheet.ns) 
        for row in rows:
            self._row[int(row.get('r'))]={}
            for c in row:
                for v in c:
                    if v.tag==(self.sheet.ns+'v'):
                        t = v.text
                        break
                xy = al2d(c.get('r'))
                self._cell[xy]=Workcell(xy[0],xy[1],t)
                self._row[int(row.get('r'))][xy[1]]=self._cell[xy]

    def row(self,n):
        if not self._row:
            self._get_cells()
        return self._row[n]

    def cell(self,x,y):
        if not self._cell:
            self._get_cells()
        try:
            return self._cell[x,y]
        except KeyError:
            return None

class Spreadsheet(OOXMLBase):
    def __init__(self, filename):
        super(Spreadsheet,self).__init__(filename)
        self._process_workbook()

    def _get_part_relationship(self):
        part_path = self.package_relationship['officeDocument'].split('/')
        self.docpath = part_path[0]
        doc = OFile(self.package.read(self.docpath+'/_rels/'+part_path[1]+".rels"))
        self.part_relationship = {}
        for elm in doc.tree:
            self.part_relationship[elm.get('Id')]=(elm.get('Type'),elm.get('Target'))

    def _process_workbook(self):
        doc = OFile(self.package.read(self.package_relationship['officeDocument']))
        self._sheet={}
        for elm in doc.tree.findall('.//%ssheet'%doc.ns):
            self._sheet[int(elm.get('sheetId'))] = Worksheet(elm.get('name'),self,(self.docpath+'/'+self.part_relationship[elm.get('{%s}id'%doc.tree.nsmap['r'])][1]))

    def sheet(self,id):
        return self._sheet[id].get_sheet()
        

if "__main__"==__name__:
    import sys
    #workbook = Spreadsheet(sys.argv[1])
    #sheet = workbook.sheet(1)
    pass
