from __future__ import with_statement
import os, sys
import win32com.client as client
from zope.interface import implements
from zope.interface import providedBy, implementedBy

from model import *

__all__ = ("VisioApplication", "VisioDocument", "VisioPage", "VisioShape")


def VisioApplication(visible=False):
   app = IApplication(client.Dispatch("Visio.Application"))
   #if visible:
   #   app._app.visible = 1
   #else:
   #   app._app.visible = 0
   return app


   
class VisioIApplicationAdapter:

   implements(IApplication)

   def __init__(self, context):
      "start the application"
      self._app = context
      self.documents = {}

      
   @property
   def visible(self):
      return self._app.visible
   
   @property
   def objects(self):
      pass
      
   def open(self, filename):
      "open a document"
      #self.documents[filename] = VisioDocument(self._app.Documents.Open(filename))

   def quit(self):
      self._app.Quit()

registry.register([None], IApplication, '', VisioIApplicationAdapter)


class VisioDocument:

   implements (IDocument)
   
   def __init__(self, obj):
      "create the document"        
      self._doc = obj
      
      @property
      def objects(self):
         for p in self.Pages:
            yield Page(p)
      

class VisioPage:

   implements (IPage)

   def __init__(self, obj):
      self._obj = obj
      
   @property
   def name(self):
      return self._obj.Name
      
   @property
   def objects(self):
      for s in self._obj.Shapes:
         yield Shape(s)

   @property
   def master(self):
      return self._obj.master

   def __str__(self):
      return self.name

      
class VisioShape:

   def __init__(self, shape):
      self._shape = shape
      self.text = shape.Text
      self.name = shape.Name
      self.master = shape.Master.Name
      self.isconnector = False
      self._setCustomProps()
      self._setConnectionPoints()

   def get_custom_property(self, propname):
      "retrieve a custom Visio property - please, there must be an easier way???"
      cprops = self.Section(243) # 243 - custom properties section
      for ri in range(0, cprops.Count):
         row = cprops.Row(ri)
         if row.Name == (propname):
            for ci in range(0, row.Count):
               cell = row.Cell(ci)
               if cell.Name == (u'Prop.' + propname):
                  return cell.Formula.strip('"')
      return ""
   
   def _setConnectionPoints(self):
      sfrom = sto = None
      
      try:
         sfrom = self._shape.Connects[0].ToSheet.Text
      except:
         pass
      
      try:
         sto = self._shape.Connects[1].ToSheet.Text
      except:
         pass

      if sfrom and sto:
         self.isconnector = True
         
      self.shapefrom = sfrom
      self.shapeto = sto
      
   def _setCustomProps(self):
      try:
         cprops = self._shape.Section(243)
      except:
         return # no custom props?
      for ri in range(0, cprops.Count):
         row = cprops.Row(ri)
         setattr(self, row.Name, row.Cell(0).Formula)

   def __str__(self):
      return self.name
   
if __name__=="__main__":
   a = VisioApplication()
   print a._app