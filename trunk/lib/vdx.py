from xml.etree import ElementTree
from xml.etree.ElementTree import QName

__all__ = ("namespace",)

namespace = ns = "http://schemas.microsoft.com/visio/2003/core"

def tag(tagname):
   return QName(ns, tag)

   
if __name__=="__main__":
   pass
