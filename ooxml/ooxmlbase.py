# -*- coding: utf-8 -*-
# Copyright (c) 2011 Wensheng Wang
# MIT license

import sys
import os.path
import zipfile
try: from lxml import etree as ElementTree
except ImportError: from xml.etree import cElementTree as ElementTree
if sys.version_info[0]==3:
    import io
    BytesIO = io.BytesIO
elif sys.version_info[0]==2 and sys.version_info[1]>4:
    import StringIO
    BytesIO = StringIO.StringIO
else:
    raise ImportError, "must use python 3 or >=2.5"


tmpl_dir = os.path.join(os.path.dirname(__file__),'ooxml-templates')

class OFile(object):
    def __init__(self,string):
        self.tree = ElementTree.fromstring(string)
        #self.ns=self.tree.nsmap[None]
        self.ns=self.tree.tag[:self.tree.tag.rindex('}')+1]


class OOXMLBase(object):
    def __init__(self, filename=None):
        self.docpath = None
        self.package = None
        self.package = zipfile.ZipFile(filename,"r")
        self._get_content_types()
        self._get_package_relationship()
        self._get_part_relationship()

    def _get_content_types(self):
        file = OFile(self.package.read("[Content_Types].xml"))
        self.default_type = {}
        self.override_type = {}
        for elm in file.tree:
            if elm.tag.endswith('Default'):
                self.default_type[elm.get('Extension')]=elm.get('ContentType')
            if elm.tag.endswith('Override'):
                self.default_type[elm.get('PartName')[1:]]=elm.get('ContentType')

    def _get_package_relationship(self):
        file = OFile(self.package.read("_rels/.rels"))
        interested = ['officeDocument','core-properties','extended-properties']
        self.package_relationship = {}
        for elm in file.tree:
            r_type = elm.get('Type')
            for i in interested:
                if r_type.endswith(i):
                    self.package_relationship[i]=elm.get('Target')
                    break

    def _get_part_relationship(self):
        raise NotImplementedError
