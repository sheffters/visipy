import os, sys
import win32com.client as client
from zope.interface import implements
     

class VisioApplication:

   implements (IApplication)

   def __init__(self):
      "start the application"
      self._app = client.Dispatch("Visio.Application")

      if visible:
         self._app.visible = 1
      else:
         self._app.visible = 0

      self.documents = {}
      
   @property
   def visible(self):
      return self._app.visible
               
   def open(filename):
      "open a document"
         self.documents[filename] = Document(self._app.Documents.Open(filename))

   def quit(self):
      self._app.Quit()

      
class VisioDocument:

   implements (IDocument)
   
   def __init__(self, obj):
      "create the document"        
      self._doc = obj
      
      @property
      def pages(self):
         for p in self.Pages:
            yield Page(p)
      

class VisioPage:

   implements (IVisioPage)

   def __init__(self, obj):
      self._obj = obj
      
   @property
   def name(self)
      return self._obj.Name
      
   @property
   def shapes(self):
      for s in self._obj.Shapes:
         yield Shape(s)
   
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
   encoding ="utf-8"
   
   a = Application()
   a.open(os.getcwd() + os.sep + "arvoverkko.vsd")
   document = a.documents
   for p in v.pages:
      print p
      for s in p.shapes:
         print s.name.encode(encoding)
         print s.master.encode(encoding)
         if s.isconnector:
            print "%s -> %s" % (s.shapefrom, s.shapeto)