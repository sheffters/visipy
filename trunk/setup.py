"""visipy library: extract information from MS Visio diagrams
"""

classifiers = """\
Development Status :: 3 - Alpha
Programming Language :: Python
Topic :: Software Development :: Libraries :: Python Modules
Topic :: Software Development :: Libraries
Environment :: Win32 (MS Windows)
Intended Audience :: Developers
"""

import sys
from distutils.core import setup

import visipy

if sys.version_info < (2, 3):
    _setup = setup
    def setup(**kwargs):
        if kwargs.has_key("classifiers"):
            del kwargs["classifiers"]
        _setup(**kwargs)

doclines = __doc__.split("\n")

setup(name="visipy",
      version="0.1",
      maintainer="Petri Savolainen",
      maintainer_email="petri.savolainen@iki.fi",
      platforms = ["win32", "unix"],
      packages = ["visipy"],
      package_dir = {"visipy": "lib"},
      description = doclines[0],
      classifiers = filter(None, classifiers.split("\n")),
      long_description = "\n".join(doclines[2:]),
)
