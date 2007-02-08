from com import *
from vdx import *

"""
   encoding ="utf-8"
   
   a = Application()
   a.open(os.getcwd() + os.sep + "test.vsd")
   document = a.documents
   for p in v.pages:
      print p
      for s in p.shapes:
         print s.name.encode(encoding)
         print s.master.encode(encoding)
         if s.isconnector:
            print "%s -> %s" % (s.shapefrom, s.shapeto)
            
"""