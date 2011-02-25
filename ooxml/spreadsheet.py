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

    @property
    def span(self):
        return self._span

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
        self.path = path
        self._row = {}
        self._cell = {}

    def get_sheet(self):
        self.sheet = OFile(self.workbook.package.read(self.path))
        dimension = self.sheet.tree.find('.//%sdimension'%self.sheet.ns).get('ref').split(':')  
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
                for v in c:
                    if v.tag==(self.sheet.ns+'v'):
                        t = v.text
                        break
                xy = al2d(c.get('r'))
                self._cell[xy]=Workcell(xy[0],xy[1],t)
                rowcell[xy[1]] = self._cell[xy]
            self._row[rowid] = Workrow(rowid,row.get('spans'),rowcell)
             

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
        writer.writerows(workbook.sheet(1).row_values_iter())
        f.close()

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
        self._sheet_names={}
        for elm in doc.tree.findall('.//%ssheet'%doc.ns):
            id = int(elm.get('sheetId'))
            name = elm.get('name')
            self._sheet[id] = Worksheet(name,self,(self.docpath+'/'+self.part_relationship[elm.get('{%s}id'%doc.tree.nsmap['r'])][1]))
            self._sheet_names[name]=id

    def sheet(self,id):
        if type(id)==type(""):
            return self._sheet[self._sheet_names[id]].get_sheet()
        else:
            return self._sheet[id].get_sheet()

    @property
    def sheet_names(self):
        return self._sheet_names.keys()
 

if "__main__"==__name__:
    import sys
    workbook = Spreadsheet(sys.argv[1])
    workbook.sheet(1).save_csv(sys.argv[2])
    #pass
