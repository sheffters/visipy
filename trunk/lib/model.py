from zope.interface.adapter import AdapterRegistry
from zope.interface.interface import adapter_hooks
from zope.interface import Interface, Attribute, providedBy



registry = AdapterRegistry()

def hook(provided, obj):
   adapter = registry.lookup1(providedBy(obj), provided, '')
   return adapter(object)

adapter_hooks.append(hook)

  
class INamedContainer(Interface):

   name = Attribute("name of this container")
   objects = Attribute("contained objects")


class ITemplated(Interface):
   
   master = Attribute("the master template of this object")


class IApplication(INamedContainer):
   pass
   
class IDocument(INamedContainer):
   pass
   
class IPage(INamedContainer):
   pass
   
class IShape(ITemplated):

   isconnector = Attribute("whether or not this shape is a connector")
   
      