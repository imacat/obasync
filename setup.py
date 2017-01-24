#! /opt/openoffice4/program/python
# -*- coding: utf-8 -*-
# Python setuptools installer for the obasync project.
#   by imacat <imacat@mail.imacat.idv.tw>, 2016-12-20

#  Copyright (c) 2016 imacat.
#
#  Licensed under the Apache License, Version 2.0 (the "License");
#  you may not use this file except in compliance with the License.
#  You may obtain a copy of the License at
#
#      http://www.apache.org/licenses/LICENSE-2.0
#
#  Unless required by applicable law or agreed to in writing, software
#  distributed under the License is distributed on an "AS IS" BASIS,
#  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#  See the License for the specific language governing permissions and
#  limitations under the License.

import os
import sys
from setuptools import setup

# For Python shipped with OpenOffice, "python" is a shell script
# wrapper for the real executable "python.bin" that sets the import
# and library path.  We should call "python" instead of "python.bin".
if os.path.basename(sys.executable) == "python.bin":
    sys.executable = os.path.join(
        os.path.dirname(sys.executable), "python")

setup(name="obasync",
      version="0.4",
      description="Office Basic macro source synchronizer",
      url="https://pypi.python.org/pypi/obasync",
      author="imacat",
      author_email="imacat@mail.imacat.idv.tw",
      license="Apache License, Version 2.0",
      zip_safe=False,
      scripts=["bin/obasync"],
      platforms=["Linux", "MS-Windows", "MacOS X"],
      keywords=["OpenOffice", "LibreOffice", "uno", "Basic", "macro"],
      classifiers=[
          "Development Status :: 5 - Production/Stable",
          "Environment :: Console",
          "Intended Audience :: Developers",
          "License :: OSI Approved :: Apache Software License",
          "Operating System :: MacOS :: MacOS X",
          "Operating System :: Microsoft :: Windows",
          "Operating System :: POSIX :: Linux",
          "Programming Language :: Python :: 2",
          "Programming Language :: Python :: 2.7",
          "Programming Language :: Python :: 3",
          "Programming Language :: Python :: 3.5",
          "Topic :: Office/Business :: Office Suites",
          ("Topic :: Text Editors :: "
           "Integrated Development Environments (IDE)"),
          "Topic :: Utilities"
      ])
