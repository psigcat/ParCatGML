# -*- coding: utf-8 -*-

def classFactory(iface):

    from .ParCatGML import ParCatGML
    return ParCatGML(iface)
