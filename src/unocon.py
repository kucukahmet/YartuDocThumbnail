try:
    import uno
except ImportError:
    uno = None

import os, random, threading, time, uuid
from os.path import abspath

from io import StringIO

from unohelper import systemPathToFileUrl
from com.sun.star.beans import PropertyValue
from com.sun.star.connection import NoConnectException

class Uno(object):
    
    def __init__(self, connection):
        context = uno.getComponentContext()
        manager = context.ServiceManager
        resolver = manager.createInstanceWithContext('com.sun.star.bridge.UnoUrlResolver', context)
        try:
            context = resolver.resolve(connection)
        except NoConnectException:
            pass

        manager = context.ServiceManager
        self.desktop = manager.createInstanceWithContext('com.sun.star.frame.Desktop', context)

    def _to_properties(self, dict):
        props = []
        for key in dict:
            prop = PropertyValue()
            prop.Name = key
            prop.Value = dict[key]
            props.append(prop)
        return tuple(props)

    def export_to_pdf(self, path, out_path):
        file = out_path + (str(uuid.uuid4().hex) + ".pdf")
        sUrl = uno.systemPathToFileUrl(abspath(file))
        loadProperties = { "Hidden": True }
        document = None
        try:
            pathUrl = uno.systemPathToFileUrl(abspath(path))
            document = self.desktop.loadComponentFromURL(pathUrl, '_blank', 0, self._to_properties(loadProperties))
            document.storeToURL(sUrl, self._to_properties({ "FilterName": "writer_pdf_Export" }))
        finally:
            if document:
                document.dispose()
        
        return file

class Connections(Uno):
    
    def __init__(self, _uno, connection):
        super(Connections, self).__init__(connection)
        self._uno = _uno
        self.busy = threading.Event()
        self.last_used = None

    def __enter__(self):
        return self.open()

    def __exit__(self, *args):
        self.close()

    def open(self):
        self._uno.lock.acquire()
        try:
            self.last_used = time.time()
            self.busy.set()
        finally:
            self._uno.lock.release()
        return self

    def close(self):
        self._uno.lock.acquire()
        try:
            self.busy.clear()
        finally:
            self._uno.lock.release()

class ManageUno(object):
    
    def __init__(self):
        self.connections = {}
        self.lock = threading.RLock()

    def start(self, connection):
        self.lock.acquire()
        try:
            connections = self.connections.setdefault(connection, [])
            free = [x for x in connections if not x.busy.isSet()]

            if free:
                con = random.choice(free)
            else:
                con = Connections(self, connection)
                connections.append(con)

            con.open()
            return con
        finally:
            self.lock.release()


_uno = ManageUno()

def uno_start(connection=None):
    global _uno
    if connection is None:
        raise "Error"

    return _uno.start(connection)