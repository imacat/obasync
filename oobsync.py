#! /opt/openoffice4/program/python
# -*- coding: utf-8 -*-
# OpenOffice Basic macros source synchronizer.
#   by imacat <imacat@mail.imacat.idv.tw>, 2016-08-31

import argparse
import os
import sys
import time

# Finds uno.py for MS-Windows
is_found_uno = False
for p in sys.path:
    if os.path.exists(os.path.join(p, "uno.py")):
        is_found_uno = True
        break
if not is_found_uno:
    cand = sys.executable
    while cand != os.path.dirname(cand):
        cand = os.path.dirname(cand)
        if os.path.exists(os.path.join(cand, "uno.py")):
            sys.path.append(cand)
            break
import uno
from com.sun.star.connection import NoConnectException


def main():
    """ The main program. """
    t_start = time.time()
    global args

    # Parses the arguments
    parse_args()

    # Connects to the OpenOffice
    oo = OpenOffice()

    # Synchronize the Basic macros.
    sync_macros(oo)

    print >> sys.stderr, "Done.  %02d:%02d elapsed." % \
        (int((time.time() - t_start) / 60), (time.time() - t_start) % 60)


def parse_args():
    """ Parses the arguments. """
    global args

    parser = argparse.ArgumentParser(
        description=("Synchronize the local Basic scripts"
                     " with OpenOffice Basic."))
    parser.add_argument(
        "projdir", metavar="dir", nargs="?", default=os.getcwd(),
        help=("The project source directory"
              " (default to the current directory)."))
    parser.add_argument(
        "library", metavar="Library", nargs="?",
        help=("The Library to upload/download the macros"
              " (default to the name of the directory)."))
    parser.add_argument(
        "--get", action="store_true",
        help="Downloads the macros instead of upload.")
    parser.add_argument(
        "-r", "--run", metavar="Module.Macro",
        help="The macro to run after the upload, if any.")
    parser.add_argument(
        "-v", "--version", action="version", version="%(prog)s 0.1")
    args = parser.parse_args()

    # Obtains the absolute path
    args.projdir = os.path.abspath(args.projdir)
    # Obtains the library name from the path
    if args.library is None:
        args.library = os.path.basename(args.projdir)

    return


def sync_macros(oo):
    """ Synchronize the Basic macros. """
    global args

    libraries = oo.service_manager.createInstance(
        "com.sun.star.script.ApplicationScriptLibraryContainer")
    # Downloads the macros from OpenOffice
    if args.get:
        modules = read_basic_modules(libraries, args.library)
        if len(modules) == 0:
            print >> sys.stderr, \
                "ERROR: Library " + args.library + " does not exist"
            return
        update_source_dir(args.projdir, modules)

    # Uploads the macros onto OpenOffice
    else:
        modules = read_in_source_dir(args.projdir)
        if len(modules) == 0:
            print >> sys.stderr, \
                "ERROR: Found no source macros in " + args.projdir
            return
        update_basic_modules(libraries, args.library, modules)
        if args.run is not None:
            factory = oo.service_manager.DefaultContext.getByName(
                "/singletons/com.sun.star.script.provider."
                "theMasterScriptProviderFactory")
            provider = factory.createScriptProvider("")
            script = provider.getScript(
                "vnd.sun.star.script:"
                + args.library + "." + args.run
                + "?language=Basic&location=application")
            script.invoke((), (), ())
    return


def read_in_source_dir(projdir):
    """ Reads-in the source macros. """
    modules = {}
    for entry in os.listdir(projdir):
        path = os.path.join(projdir, entry)
        if os.path.isfile(path) and entry.lower().endswith(".vb"):
            modname = entry[0:-3]
            f = open(path)
            modules[modname] = f.read().decode("utf-8").replace(
                "\r\n", "\n")
            f.close()
    return modules


def update_source_dir(projdir, modules):
    """ Updates the source macros. """
    curmods = {}
    for entry in os.listdir(projdir):
        path = os.path.join(projdir, entry)
        if os.path.isfile(path) and entry.lower().endswith(".vb"):
            modname = entry[0:-3]
            curmods[modname] = entry
    for modname in sorted(modules.keys()):
        if modname not in curmods:
            path = os.path.join(projdir, modname + ".vb")
            f = open(path, "w")
            f.write(modules[modname])
            f.close()
            print >> sys.stderr, modname + ".vb added."
        else:
            path = os.path.join(projdir, curmods[modname])
            f = open(path, "r+")
            if modules[modname] != f.read().replace("\r\n", "\n"):
                f.seek(0)
                f.write(modules[modname])
                print >> sys.stderr, curmods[modname] + " updated."
            f.close()
    for modname in sorted(curmods.keys()):
        if modname not in modules:
            path = os.path.join(projdir, curmods[modname])
            os.remove(path)
            print >> sys.stderr, curmods[modname] + " removed."
    return


def read_basic_modules(libraries, libname):
    """ Reads the OpenOffice Basic macros from the macros storage. """
    modules = {}
    if not libraries.hasByName(libname):
        return modules
    libraries.loadLibrary(libname)
    library = libraries.getByName(libname)
    for modname in library.getElementNames():
        modules[modname] = library.getByName(modname).encode("utf-8")
    return modules


def update_basic_modules(libraries, libname, modules):
    """ Updates the OpenOffice Basic macros storage. """
    if not libraries.hasByName(libname):
        libraries.createLibrary(libname)
        print >> sys.stderr, "Library " + libname + " created."
        library = libraries.getByName(libname)
        for modname in sorted(modules.keys()):
            library.insertByName(modname, modules[modname])
            print >> sys.stderr, "Module " + modname + " added."
    else:
        libraries.loadLibrary(libname)
        library = libraries.getByName(libname)
        curmods = sorted(library.getElementNames())
        for modname in sorted(modules.keys()):
            if modname not in curmods:
                library.insertByName(modname, modules[modname])
                print >> sys.stderr, "Module " + modname + " added."
            elif modules[modname] != library.getByName(modname):
                library.replaceByName(modname, modules[modname])
                print >> sys.stderr, "Module " + modname + " updated."
        for modname in curmods:
            if modname not in modules:
                library.removeByName(modname)
                print >> sys.stderr, "Module " + modname + " removed."
    if libraries.isModified():
        libraries.storeLibraries()
    return


class OpenOffice:
    """ The OpenOffice connection. """

    def __init__(self):
        """Initializes the object."""
        self.bootstrap_context = None
        self.service_manager = None
        self.desktop = None
        self.connect()

    def connect(self):
        """Connects to the running OpenOffice process."""
        # Obtains the local context
        local_context = uno.getComponentContext()
        # Obtains the local service manager
        local_service_manager = local_context.getServiceManager()
        # Obtains the URL resolver
        url_resolver = local_service_manager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        # Obtains the context
        url = ("uno:socket,host=localhost,port=2002;"
               "urp;StarOffice.ComponentContext")
        while True:
            try:
                self.bootstrap_context = url_resolver.resolve(url)
            except NoConnectException:
                self.start_oo()
            else:
                break
        # Obtains the service manager
        self.service_manager = self.bootstrap_context.getServiceManager()
        # Obtains the desktop service
        self.desktop = self.service_manager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.bootstrap_context)

    def start_oo(self):
        """Starts the OpenOffice in server listening mode"""
        # For MS-Windows, which does not have fork()
        if os.name == "nt":
            from subprocess import Popen
            ooexec = os.path.join(os.path.dirname(uno.__file__), "soffice.exe")
            DETACHED_PROCESS = 0x00000008
            Popen([ooexec, "-accept=socket,host=localhost,port=2002;urp;"],
                  close_fds=True, creationflags=DETACHED_PROCESS)
            time.sleep(2)
            return

        # For POSIX systems, including Linux
        try:
            pid = os.fork()
        except OSError:
            sys.stderr.write("failed to fork().\n")
            sys.exit(1)
        if pid != 0:
            time.sleep(2)
            return
        os.setsid()
        ooexec = os.path.join(os.path.dirname(uno.__file__), "soffice.bin")
        try:
            os.execl(
                ooexec, ooexec,
                "-accept=socket,host=localhost,port=2002;urp;")
        except OSError:
            sys.stderr.write(
                ooexec + ": Failed to run the OpenOffice server.\n")
            sys.exit(1)

if __name__ == "__main__":
    main()
