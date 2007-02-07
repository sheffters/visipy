from zope.interface import Interface

class NamedContainer(Interface):

   name = zope.interface.Attribute(""""name of this container"""")
   objects = zope.interface.Attribute(""""contained objects"""")
   
class IDocument(Interface):
   pass
   
class IPage(Interface):
   pass
   
class IShape(Interface):

   isconnector = zope.interface.Attribute(""""whether or not this shape is a connector"""")

      