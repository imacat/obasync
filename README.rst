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

On the other hand, given the following Basic macros:

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

You can either:

1. Install ``obasync`` with ``pip`` (recommended), or

2. Download the ``obasync`` script manually, and run it with the
   Python that come with your OpenOffice/LibreOffice installation.

We will explain them in detail.


OpenOffice/LibreOffice That Comes with Your Linux
#################################################

Install with ``pip`` (Recommended)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Your system may already have ``pip`` installed.  If not, install the
``python-pip`` package from the system package manager.  Then, run::

    pip install obasync

Download and Install Manually
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Download ``obasync`` from either `PyPI
<https://pypi.python.org/pypi/obasync>`_ or `GitHub
<https://github.com/imacat/obasync>`_.  Then, run ``obasync`` as::

    python obasync

Or, you can edit the script and change the first line (shebang) to::

    #! /usr/bin/python

and save this script somewhere in your path, say, ``/usr/local/bin``.
Then you can run ``obasync``.


OpenOffice 4 on Linux
#####################

Install with ``pip`` (Recommended)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Install ``pip`` for your OpenOffice installation, and then install
``obasync`` with this ``pip``::

    wget https://bootstrap.pypa.io/get-pip.py
    sudo /opt/openoffice4/program/python get-pip.py
    /opt/openoffice4/program/python-core-2.7.6/bin/pip install obasync

Download and Install Manually
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Download ``obasync`` from either `PyPI
<https://pypi.python.org/pypi/obasync>`_ or `GitHub
<https://github.com/imacat/obasync>`_.  Then, run ``obasync`` as::

    /opt/openoffice4/program/python obasync

Or, you can edit the script and change the first line (shebang) to::

    #! /opt/openoffice4/program/python

and save this script somewhere in your path, say, ``/usr/local/bin``.
Then you can run ``obasync``.


LibreOffice on Linux
####################

Python from LibreOffice on Linux does not install ``pip`` properly.
However, you can still download and install ``obasync`` manually.

Install with ``pip`` (Recommended)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Download ``obasync`` from either `PyPI
<https://pypi.python.org/pypi/obasync>`_ or `GitHub
<https://github.com/imacat/obasync>`_.  Then, run ``obasync`` as::

    /opt/libreoffice5.2/program/python obasync

Or, you can edit the script and change the first line (shebang) to::

    #! /opt/libreoffice5.2/program/python

and save this script somewhere in your path, say, ``/usr/local/bin``.
Then you can run ``obasync``.


OpenOffice on MS-Windows
########################

You can install ``obasync`` with ``pip``, but the result is messy.
The recommended way is to download and install ``obasync`` manually.

Download and Install Manually
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Download ``obasync`` from either `PyPI
<https://pypi.python.org/pypi/obasync>`_ or `GitHub
<https://github.com/imacat/obasync>`_.  Then, run ``obasync`` as::

    "C:\Program Files (x86)\OpenOffice 4\program\python.exe" obasync


LibreOffice on MS-Windows
#########################

You can install ``obasync`` with ``pip``, but the result is messy.
The recommended way is to download and install ``obasync`` manually.

Download and Install Manually
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Download ``obasync`` from either `PyPI
<https://pypi.python.org/pypi/obasync>`_ or `GitHub
<https://github.com/imacat/obasync>`_.  Then, run ``obasync`` as::

    "C:\Program Files\LibreOffice 5\program\python.exe" obasync


OpenOffice on Mac OS X
######################

Install with ``pip`` (Recommended)
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Install ``pip`` first, and then install ``obasync`` with ``pip``::

    wget https://bootstrap.pypa.io/get-pip.py
    sudo python get-pip.py
    sudo pip install obasync

Download and Install Manually
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Download ``obasync`` from either `PyPI
<https://pypi.python.org/pypi/obasync>`_ or `GitHub
<https://github.com/imacat/obasync>`_.  Then, run ``obasync`` as::

    python obasync

Or, you can edit the script and change the first line (shebang) to::

    #! /usr/bin/python

and save this script somewhere in your path, say, ``/usr/local/bin``.
Then you can run ``obasync``.



LibreOffice on Mac OS X
#######################

Python from LibreOffice on Mac OS X does not install ``pip`` properly.
However, you can still download and install ``obasync`` manually.

Download and Install Manually
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

Download ``obasync`` from either `PyPI
<https://pypi.python.org/pypi/obasync>`_ or `GitHub
<https://github.com/imacat/obasync>`_.  Then, run ``obasync`` as::

    /Applications/LibreOffice.app/Contents/Resources/python obasync

Or, you can edit the script and change the first line (shebang) to::

    #! /Applications/LibreOffice.app/Contents/Resources/python

and save this script somewhere in your path, say, ``/usr/local/bin``.
Then you can run ``obasync``.


OPTIONS
-------

::

  obasync [options] [DIRECTORY [LIBRARY]]

DIRECTORY       The project source directory.  Default to the current
                working directory.

LIBRARY         The name of the Basic library.  Default to the same
                name as the project source directory.

--get           Download (check out) the macros from the
                OpenOffice/LibreOffice Basic storage to the source
                files, instead of upload (check in).  By default it
                uploads the source files onto the
                OpenOffice/LibreOffice Basic storage.

-p, --port N    The TCP port to communicate with
                OpenOffice/LibreOffice.  The default is 2002.  You can
                change it if port 2002 is already in use.

-x, --ext .EXT  The file name extension of the source files.  The 
                default is ``.vb``.  This may be used for your
                convenience of editor syntax highlighting.

-e, --encoding CS
                The encoding of the source files.  The default is
                system-dependent.  For example, on Traditional Chinese
                MS-Windows, this will be CP950 (Big5).  You can change
                this to UTF-8 for convenience if you
                obtain/synchronize your source code from other
                sources.

-r, --run MODULE.MACRO
                Run he specific macro after synchronization, for
                convenience.

-h, --help      Show the help message and exit

-v, --version   Show program’s version number and exit


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
