import unittest
import com_tests

alltests = unittest.TestSuite([com_tests.suite])

if __name__=="__main__":
   unittest.TextTestRunner().run(alltests)