import sys, os, unittest

from visipy.com import *

class BasicAppTests(unittest.TestCase):
   
   def setUp(self):
      self.app = VisioApplication()
   
   def testOpenFile(self):
      path = sys.path[0] + os.sep + "testfile.vsd"
      #self.app.open(path)

   def tearDown(self):
      self.app.quit()


class testIApplication(unittest.TestCase):
   pass

suite = unittest.TestLoader().loadTestsFromTestCase(BasicAppTests)

           
if __name__ == '__main__':
   unittest.main()
