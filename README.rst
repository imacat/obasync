``obasync`` - Office Basic Macro Source Synchronizer
====================================================

DESCRIPTION
-----------

``obasync`` is an OpenOffice/LibreOffice Basic macro source
synchronizer.  It synchronizes your Basic macros with your local
project files.


Given the following source files:

* Directory: ``MyApp``
* Files: ``MyMacros.vb`` ``Utils.vb`` ``Registry.vb`` ``Data.vb``

Running ``obasync`` will synchronize them with the following Basic
macros:

* Library: ``MyApp``
* Modules: ``MyMacros`` ``Utils`` ``Registry`` ``Data``

If the Basic library ``MyApp`` does not exist, it will be created.
Missing modules will be added, and excess modules will be removed.

And vice versa.  Given the following Basic macros:

* Library: ``MyApp``
* Modules: ``MyMacros`` ``Utils`` ``Registry`` ``Data``

Running ``obasync --get`` will synchronize them with the following
source files:

* Directory: ``MyApp``
* Files: ``MyMacros.vb`` ``Utils.vb`` ``Registry.vb`` ``Data.vb``

Missing source files will be added, and excess source files will be
deleted.


INSTALL
-------

You can install ``obasync`` with ``pip``.  Uses the ``python``
executable that comes with your OpenOffice/LibreOffice installation
when possible.

* OpenOffice 4 on Linux::

    /opt/openoffice4/program/python `which pip` install obasync

* LibreOffice 5.x on Linux::

    /opt/libreoffice5.*/program/python `which pip` install obasync

* Linux vendor OpenOffice/LibreOffice installation::

    pip install obasync

* OpenOffice on Mac OS X relies on the system ``python``
  installation::

    sudo pip install obasync

* LibreOffice on Mac OS X:  There is no simple ``pip`` way to install
  ``obasync``.  However, you can still download the script and run
  it with the ``python`` executable that comes with your LibreOffice
  installation.  See below.

* OpenOffice/LibreOffice on MS-Windows:  You can still install
  ``obasync`` with ``pip`` from your OpenOffice/LibreOffice
  installation.  However, it would be much easier to just download the
  script and run it with the ``python`` executable that comes with
  your OpenOffice/LibreOffice installation.  See below.

To download the package or the source script:

* Python package: https://pypi.python.org/pypi/obasync
* Source directory: https://github.com/imacat/obasync


OPTIONS
-------

::

  obasync [options] [directory [library]]

directory
   The project source directory.  Default to the current working
   directory.

library
   The name of the Basic library.  Default to the same name as the
   project source directory.

-g, --get
   Download (check out) the macros from the OpenOffice/LibreOffice
   Basic storage to the source files, instead of upload (check in).
   By default it uploads the source files onto the
   OpenOffice/LibreOffice Basic storage.

-p, --port N
   The TCP port to communicate with OpenOffice/LibreOffice.  The
   default is 2002.  You can change it if port 2002 is already in use.

-x, --ext .EXT
   The file name extension of the source files.  The default is
   ``.vb``.  This may be used for your convenience of editor syntax
   highlighting.

-e, --encoding CS
   The encoding of the source files.  The default is system-dependent.
   For example, on Traditional Chinese MS-Windows, this will be
   CP950 (Big5).  You can change this to UTF-8 for convenience if you
   obtain/synchronize your source code from other sources.

-r, --run Module.Macro
   Run he specific macro after synchronization, for convenience.

-h, --help
   Show the help message and exit

-v,--version
   Show program's version number and exit


COPYRIGHT
---------

  Copyright (c) 2016 imacat.
  
  Licensed under the Apache License, Version 2.0 (the "License");
  you may not use this file except in compliance with the License.
  You may obtain a copy of the License at
  
      http://www.apache.org/licenses/LICENSE-2.0
  
  Unless required by applicable law or agreed to in writing, software
  distributed under the License is distributed on an "AS IS" BASIS,
  WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
  See the License for the specific language governing permissions and
  limitations under the License.

SUPPORT
-------

  Contact imacat <imacat@mail.imacat.idv.tw> if you have any question.
