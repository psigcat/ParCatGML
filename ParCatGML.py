# -*- coding: utf-8 -*-
from PyQt5.QtCore import Qt, QCoreApplication, QDateTime
from PyQt5.QtGui import QIcon, QPixmap
from PyQt5.QtWidgets import QDialog, QAction, QMessageBox, QFileDialog, QTableWidgetItem, QFrame
from qgis.core import QgsProject, QgsPointXY

import os
import sys

from .ParCatGML_dialog import Ui_ParCatGML_dialog


class ParCatGML(QDialog):

    def __init__(self, iface):

        super(ParCatGML, self).__init__()
        self.iface = iface

        # Inicialitzar variables de treball
        if sys.platform.upper().startswith("WIN"):
            self.so = "W"
        else:
            self.so = "L"
        self.plugin_dir = self.Barres(os.path.dirname(__file__))
        self.project_dir = ""
        self.actions = ""
        self.menu = self.tr(u'&ParCatGML')
        self.toolbar = self.iface.addToolBar(u'ParCatGML')
        self.toolbar.setObjectName(u'ParCatGML')
        self.iface.newProjectCreated.connect(self.CanviProjecte)
        self.iface.projectRead.connect(self.CanviProjecte)


    def tr(self, message):

        # Convertir string utilitzant API Qt
        return QCoreApplication.translate('ParCatGML', message)


    def Barres(self, bar):
        """ Conversió barres / i \ segons S.O. """

        if bar.strip() == "" : return ""
        if self.so == "W":
            z = bar.replace('/','\\')
            if str(z)[-1] == "\\" : z = z[0:-1]
        else :
            z = bar.replace('\\','/')
            if str(z)[-1] == "/" : z = z[0:-1]

        return z


    def Omple(self, v, n):
        """ Omple de zeros a l'esquerra """

        if len(str(v)) < n:
            return str("0000000000"[0:n-len(str(v))])+str(v)
        else:
            return str(v)


    def initGui(self):

        # Crear entrada a menú QGIS
        self.action = QAction(QIcon(self.Barres(self.plugin_dir+"/ParCatGML.png")), "ParCatGML", self.iface.mainWindow())
        self.action.setObjectName("ParCatGML")
        self.action.setWhatsThis("Parcela Catastral GML")
        self.action.setStatusTip("Parcela Catastral GML")
        self.action.triggered.connect(self.run)

        # Configurar botó
        self.iface.addToolBarIcon(self.action)
        self.iface.addPluginToMenu("&ParCatGML", self.action)
        
        # Configurar formulari
        self.ui=Ui_ParCatGML_dialog()
        self.ui.setupUi(self)
        self.ui.Logo.setPixmap(QPixmap(self.Barres(self.plugin_dir+"/ParCatGML.png")))
        self.ui.Canvi.setIcon(QIcon(self.Barres(self.plugin_dir+"/EditTable.png")))
        self.ui.Carpeta.setIcon(QIcon(self.Barres(self.plugin_dir+"/folder.png")))
        self.ui.Selec.currentCellChanged.connect(self.SelLis)
        self.ui.Canvi.clicked.connect(self.FerCanvi)
        self.ui.Carpeta.clicked.connect(self.SelDesti)
        self.ui.Crear.clicked.connect(self.FerGML)
        self.ui.Tancar.clicked.connect(self.close)


    def unload(self):

        # Elimina icones i accions
        self.iface.removePluginMenu(self.tr(u'&ParCatGML'),self.action)
        self.iface.removeToolBarIcon(self.action)
        self.iface.newProjectCreated.disconnect(self.CanviProjecte)
        self.iface.projectRead.disconnect(self.CanviProjecte)
        self.ui.Selec.currentCellChanged.disconnect(self.SelLis)
        self.ui.Canvi.clicked.disconnect(self.FerCanvi)
        self.ui.Carpeta.clicked.disconnect(self.SelDesti)
        self.ui.Crear.clicked.disconnect(self.FerGML)
        self.ui.Tancar.clicked.disconnect(self.close)
        del self.toolbar

            
    def Missatge(self, av, t):
        """ Finestra missatges """

        m = QMessageBox()
        m.setText(self.tr(t))
        if av == "P" :
            m.setIcon(QMessageBox.Information)
            m.setWindowTitle("Pregunta")
            m.setStandardButtons(QMessageBox.Ok | QMessageBox.No)
            b1 = m.button(QMessageBox.Ok)
            b1.setText("Si")
            b2 = m.button(QMessageBox.No)
            b2.setText("No")
        else :
            if av == "W" :
                m.setIcon(QMessageBox.Warning)
                z="Atención"
            elif av == "C" :
                m.setIcon(QMessageBox.Critical)
                z="Error"
            else :
                m.setIcon(QMessageBox.Information)
            m.setWindowTitle("Aviso")
            m.setStandardButtons(QMessageBox.Ok)
            b = m.button(QMessageBox.Ok)
            b.setText("Seguir")

        return m.exec_()


    def CanviProjecte(self):
        """ Fer quan es canvia de projecte """

        self.project_dir = ""
        self.project_dir = self.Barres(QgsProject.instance().homePath())


    def SelDesti(self):
        """ Selecció carpeta destí """

        msg = "Selección carpeta destino"
        z = QFileDialog.getExistingDirectory(None, msg, self.ui.desti.text(), QFileDialog.ShowDirsOnly)
        if z.strip() != "":
            self.ui.desti.setText(self.Barres(z))


    def FerGML(self):
        """ Crea arxiu GML Format INSPIRE """

        # fixar versió
        iv = 3
        if self.ui.Inspire4.isChecked():
            iv = 4
        # Comprovació arxiu a crear
        z = str(self.ui.Selec.item(0,1).text())
        if z == "":
            z = "anonimo"
        f = self.Barres(self.ui.desti.text()+"/"+z+".gml")
        if os.path.exists(f):
            if self.Missatge("P","El fichero "+z+".gml ya existe\n\nDesea reemplazarlo ?") == QMessageBox.No:
                return

        z = self.CapGML(iv)
        # Afegeix tag 'member' per cada parcel·la seleccionada
        total = self.ui.Selec.rowCount()
        for fil in range(total):
            ref = self.ui.Selec.item(fil,1).text()
            promun = self.Omple(self.ui.Selec.item(fil,2).text(),2)+self.Omple(self.ui.Selec.item(fil,3).text(),3)
            num = self.ui.Selec.item(fil,4).text()
            area = self.ui.Selec.item(fil,5).text()
            n = 0
            ver = self.geo[fil].vertexAt(0)
            pun = []
            while ver != QgsPointXY(0, 0):
                pun.append(ver)
                n += 1
                ver = self.geo[fil].vertexAt(n)

            punL = ""
            for p in pun:
                if punL != "":
                    punL += " "
                punL += str(format(p.x(),"f"))+" "+str(format(p.y(),"f"))

            xy = self.geo[fil].centroid().asPoint()
            centroid = str(format(xy.x(),"f"))+" "+str(format(xy.y(),"f"))
            bou = self.geo[fil].boundingBox()
            min = str(format(bou.xMinimum(),"f"))+" "+str(format(bou.yMinimum(),"f"))
            max = str(format(bou.xMaximum(),"f"))+" "+str(format(bou.yMaximum(),"f"))
            z += self.CosGML(iv,self.crs.split(":")[1],promun,ref,num,area,len(pun),punL,centroid,min,max)

        z += self.PeuGML(iv)

        # Guarda arxiu gml
        fg = open(f, "w+")
        fg.write(z)
        fg.close()
        
        self.Missatge("C","Archivo GML creado en la carpeta destino")


    def CapGML(self, v):
        """ Capçalera GML """

        if v == 3:
            z='<?xml version="1.0" encoding="utf-8"?>\n'
            z+='<!-- Archivo generado automaticamente por el plugin Export GML catastro de España de QGIS. -->\n'
            z+='<!-- Parcela Catastral de la D.G. del Catastro. -->\n'
            z+='<gml:FeatureCollection gml:id="ES.SDGC.CP" xmlns:gml="http://www.opengis.net/gml/3.2" '
            z+='xmlns:gmd="http://www.isotc211.org/2005/gmd" '
            z+='xmlns:ogc="http://www.opengis.net/ogc" '
            z+='xmlns:xlink="http://www.w3.org/1999/xlink" '
            z+='xmlns:cp="urn:x-inspire:specification:gmlas:CadastralParcels:3.0" '
            z+='xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
            z+='xsi:schemaLocation="urn:x-inspire:specification:gmlas:CadastralParcels:3.0 '
            z+='http://inspire.ec.europa.eu/schemas/cp/3.0/CadastralParcels.xsd">\n'
        else:
            z='<?xml version="1.0" encoding="utf-8"?>\n'
            z+='<!-- Archivo generado automaticamente por el plugin ParCatGML de QGIS. -->\n'
            z+='<!-- Parcela Catastral para entregar a la D.G. del Catastro. Formato INSPIRE v4. -->\n'
            z+='<wfs:FeatureCollection '
            z+='gml:id="ES.SDGC.CP" '
            z+='xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" '
            z+='xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
            z+='xmlns:xlink="http://www.w3.org/1999/xlink" '
            z+='xmlns:gml="http://www.opengis.net/gml/3.2" '
            z+='xmins:gco="http://www.isotc211.org/2005/gco" '
            z+='xmlns:ogc="http://www.opengis.net/ogc" ' #v3
            z+='xmlns:cp="http://inspire.ec.europa.eu/schemas/cp/4.0" '
            z+='xmlns:gmd="http://www.isotc211.org/2005/gmd" '
            z+='xsi:schemaLocation="http://www.opengis.net/wfs/2.0'
            z+=' http://schemas.opengis.net/wfs/2.0/wfs.xsd'
            z+=' http://inspire.ec.europa.eu/schemas/cp/4.0'
            z+=' http://inspire.ec.europa.eu/schemas/cp/4.0/CadastralParcels.xsd" '
            z+='xmins:wfs="http://www.opengis.net/wfs/2.0" '
            z+='timeStamp="'+self.ui.data.dateTime().toString(Qt.ISODate)+'" '
            z+='numberMatched="1" '
            z+='numberReturned="1"> '

        return z


    def CosGML(self,v,epsg,promun,ref,num,area,punN,punL,centroid,min,max) :

        if v == 3:
            z='<gml:featureMember>\n<cp:CadastralParcel gml:id="ES.SDGC.CP.'+str(ref)+'">\n<gml:boundedBy>\n'
            z+='<gml:Envelope srsName="urn:ogc:def:crs:EPSG::'+str(epsg)+'">\n'
            z+='<gml:lowerCorner>'+str(min)+'</gml:lowerCorner>\n<gml:upperCorner>'+str(max)+'</gml:upperCorner>\n'
            z+='</gml:Envelope>\n</gml:boundedBy>\n<cp:areaValue uom="m2">'+str(area)+'</cp:areaValue>\n'
            z+='<cp:beginLifespanVersion>'+self.ui.data.dateTime().toString(Qt.ISODate)+'</cp:beginLifespanVersion>\n'
            z+='<cp:endLifespanVersion xsi:nil="true" nilReason="other:unpopulated"></cp:endLifespanVersion>\n<cp:geometry>\n'
            z+='<gml:MultiSurface gml:id="MultiSurface_ES.SDGC.CP.'+str(ref)+'" srsName="urn:ogc:def:crs:EPSG::'+str(epsg)+'">\n<gml:surfaceMember>\n'
            z+='<gml:Surface gml:id="Surface_ES.SDGC.CP.'+str(ref)+'.1" srsName="urn:ogc:def:crs:EPSG::'+str(epsg)+'">\n'
            z+='<gml:patches>\n<gml:PolygonPatch>\n<gml:exterior>\n<gml:LinearRing>\n'
            z+='<gml:posList srsDimension="2" count="'+str(punN)+'">'+str(punL)+'</gml:posList>\n'
            z+='</gml:LinearRing>\n</gml:exterior>\n</gml:PolygonPatch>\n</gml:patches>\n</gml:Surface>\n</gml:surfaceMember>\n</gml:MultiSurface>\n'
            z+='</cp:geometry>\n<cp:inspireId xmlns:base="urn:x-inspire:specification:gmlas:BaseTypes:3.2">\n<base:Identifier>\n'
            z+='<base:localId>'+str(num)+'</base:localId>\n'
            z+='<base:namespace>ES.LOCAL.CP</base:namespace>\n</base:Identifier>\n</cp:inspireId>\n<cp:label>05</cp:label>\n'
            z+='<cp:nationalCadastralReference>2</cp:nationalCadastralReference>\n<cp:referencePoint>\n'
            z+='<gml:Point gml:id="ReferencePoint_ES.SDGC.CP.'+str(ref)+'" srsName="urn:ogc:def:crs:EPSG::'+str(epsg)+'">\n'
            z+='<gml:pos>'+str(centroid)+'</gml:pos>\n</gml:Point>\n</cp:referencePoint>\n'
            z+='<cp:validFrom xsi:nil="true" nilReason="other:unpopulated"></cp:validFrom>\n<cp:validTo xsi:nil="true" nilReason="other:unpopulated"></cp:validTo>\n'
            z+='<cp:zoning xlink:href="#ES.SDGC.CP.Z.'+str(promun)+'U"></cp:zoning>\n</cp:CadastralParcel>\n</gml:featureMember>\n<gml:featureMember>\n'
            z+='<cp:CadastralZoning gml:id="ES.SDGC.CP.Z.'+str(promun)+'U">\n<gml:boundedBy>\n<gml:Envelope srsName="urn:ogc:def:crs:EPSG::'+str(epsg)+'">\n'
            z+='<gml:lowerCorner>'+str(min)+'</gml:lowerCorner>\n<gml:upperCorner>'+str(max)+'</gml:upperCorner>\n</gml:Envelope>\n</gml:boundedBy>\n'
            z+='<cp:beginLifespanVersion>'+self.ui.data.dateTime().toString(Qt.ISODate)+'</cp:beginLifespanVersion>\n'
            z+='<cp:endLifespanVersion xsi:nil="true" nilReason="other:unpopulated"></cp:endLifespanVersion>\n'
            z+='<cp:estimatedAccuracy uom="m">0.60</cp:estimatedAccuracy>\n<cp:geometry>\n'
            z+='<gml:MultiSurface gml:id="MultiSurface_ES.SDGC.CP.Z.'+str(promun)+'U" srsName="urn:ogc:def:crs:EPSG::'+str(epsg)+'">\n<gml:surfaceMember>\n'
            z+='<gml:Surface gml:id="Surface_ES.SDGC.CP.Z.'+str(promun)+'U.1" srsName="urn:ogc:def:crs:EPSG::'+str(epsg)+'">\n'
            z+='<gml:patches>\n<gml:PolygonPatch>\n<gml:exterior>\n<gml:LinearRing>\n'
            z+='<gml:posList srsDimension="2" count="'+str(punN)+'">'+str(punL)+'</gml:posList>\n'
            z+='</gml:LinearRing>\n</gml:exterior>\n</gml:PolygonPatch>\n</gml:patches>\n</gml:Surface>\n</gml:surfaceMember>\n</gml:MultiSurface>\n'
            z+='</cp:geometry>\n<cp:inspireId xmlns:base="urn:x-inspire:specification:gmlas:BaseTypes:3.2">\n<base:Identifier>\n'
            z+='<base:localId>'+str(promun)+'U</base:localId>\n<base:namespace>ES.SDGC.CP.Z</base:namespace>\n'
            z+='</base:Identifier>\n</cp:inspireId>\n<cp:label>'+str(promun)+'U</cp:label>\n'
            z+='<cp:level codeSpace="urn:x-inspire:specification:gmlas:CadastralParcels:3.0/CadastralZoningLevelValue">1stOrder</cp:level>\n'
            z+='<cp:levelName>\n<gmd:LocalisedCharacterString locale="esp">MAPA</gmd:LocalisedCharacterString>\n</cp:levelName>\n'
            z+='<cp:nationalCadastalZoningReference>'+str(promun)+'U</cp:nationalCadastalZoningReference>\n'
            z+='<cp:originalMapScaleDenominator>1000</cp:originalMapScaleDenominator>\n<cp:referencePoint>\n'
            z+='<gml:Point gml:id="ReferencePoint_ES.SDGC.CP.Z.X'+str(promun)+'U" srsName="urn:ogc:def:crs:EPSG::'+str(epsg)+'"> \n'
            z+='<gml:pos>'+str(centroid)+'</gml:pos>\n</gml:Point>\n</cp:referencePoint>\n<cp:validFrom xsi:nil="true" nilReason="unknown" />\n'
            z+='<cp:validTo xsi:nil="true" nilReason="unknown" />\n</cp:CadastralZoning>\n</gml:featureMember>\n'
        else:
            z='<wfs:member>\n'
            z+='<cp:CadastralParcel gml:id="ES.SDGC.CP.'+str(ref)+'">\n'
            z+='<cp:areaValue uom="m2">'+str(area)+'</cp:areaValue>\n'
            z+='<cp:geometry>\n'
            z+='<gml:MultiSurface gml:id="MultiSurface_ES.SDGC.CP.'+str(ref)+'" srsName="http://www.opengis.net/def/crs/EPSG/0/'+str(epsg)+'">\n'
            z+='<gml:surfaceMember>\n'
            z+='<gml:Surface gml:id="Surface_ES.SDGC.CP.'+str(ref)+'" srsName="http://www.opengis.net/def/crs/EPSG/0/'+str(epsg)+'">\n'
            z+='<gml:patches>\n<gml:PolygonPatch>\n<gml:exterior>\n<gml:LinearRing>\n'
            z+='<gml:posList srsDimension="2">'+str(punL)+'</gml:posList>\n'
            z+='</gml:LinearRing>\n</gml:exterior>\n</gml:PolygonPatch>\n</gml:patches>\n</gml:Surface>\n</gml:surfaceMember>\n</gml:MultiSurface>\n</cp:geometry>\n'
            z+='<cp:inspireId xmlns:base="http://inspire.ec.europa.eu/schemas/base/3.3">\n<base:Identifier>\n'
            z+='<base:localId>'+str(num)+'</base:localId>\n<base:namespace>ES.LOCAL.CP</base:namespace>\n</base:Identifier>\n</cp:inspireId>\n'
            z+='<cp:label/>\n<cp:nationalCadastralReference/>\n</cp:CadastralParcel>\n'
            z+='</wfs:member>\n'

        return z


    def PeuGML(self,v) :

        if v == 3:
            z='</gml:FeatureCollection>\n'        
        else:
            z='</wfs:FeatureCollection>\n'

        return z


    def SelLis(self):
        """ llegir una fila de la llista de seleccionats """

        fil=self.ui.Selec.currentRow()
        if fil == -1 : return
        self.ui.refcat.setText(self.ui.Selec.item(fil,1).text())
        self.ui.pro.setText(self.ui.Selec.item(fil,2).text())
        self.ui.mun.setText(self.ui.Selec.item(fil,3).text())
        self.ui.num.setText(self.ui.Selec.item(fil,4).text())
        self.ui.area.setValue(int(self.ui.Selec.item(fil,5).text()))


    def FerCanvi(self):
        """ actualitzar la fila de la llista """

        fil = self.ui.Selec.currentRow()
        if fil == -1:
            return
        self.ui.Selec.item(fil,1).setText(self.ui.refcat.text())
        self.ui.Selec.item(fil,2).setText(self.ui.pro.text())
        self.ui.Selec.item(fil,3).setText(self.ui.mun.text())
        self.ui.Selec.item(fil,4).setText(self.ui.num.text())
        self.ui.Selec.item(fil,5).setText(str(self.ui.area.value()))


    def run(self):

        # Comprova capa i selecció
        capa = self.iface.activeLayer()
        if not capa:
            self.Missatge("C", "No hay ninguna capa activa")
            return

        elems = list(capa.selectedFeatures())
        ne = len(elems)
        if ne == 0:
            self.Missatge("C", "Debe seleccionar como mínimo una parcela")
            return

        self.crs = capa.crs().authid()
        if self.crs.split(':')[0] != 'EPSG':
            self.Missatge("C", "La capa activa no utiliza un sistema de coordenadas compatible")
            return

        # Verifica la informació que podem obtenir de la capa
        if capa.fields().indexFromName("REFCAT") != -1 : nRef="REFCAT"
        elif capa.fields().indexFromName("refcat") != -1 : nRef="refcat"
        elif capa.fields().indexFromName("nationalCadastralReference") != -1  : nRef="nationalCadastralReference"
        else : nRef=""
        if capa.fields().indexFromName("AREA") != -1 : nArea="AREA"
        elif capa.fields().indexFromName("area") != -1 : nArea="area"
        elif capa.fields().indexFromName("areaValue") != -1  : nArea="areaValue"
        else : nArea=""
        if capa.fields().indexFromName("DELEGACION") != -1 : nPro="DELEGACION"
        elif capa.fields().indexFromName("DELEGACIO") != -1 : nPro="DELEGACIO"
        elif capa.fields().indexFromName("delegacion") != -1 : nPro="delegacion"
        elif capa.fields().indexFromName("delegacio") != -1 : nPro="delegacio"
        elif capa.id().startswith("A_ES_SDGC_CP_") : nPro="id"
        else : nPro=""
        if capa.fields().indexFromName("MUNICIPIO") != -1 : nMun="MUNICIPIO"
        elif capa.fields().indexFromName("MUNICIPI") != -1 : nMun="MUNICIPI"
        elif capa.fields().indexFromName("municipio") != -1 : nMun="municipio"
        elif capa.fields().indexFromName("municipi") != -1 : nMun="municipi"
        elif capa.id().startswith("A_ES_SDGC_CP_") : nMun="id"
        else : nMun=""

        # Carrega llista formulari
        self.ui.data.setDateTime(QDateTime.currentDateTime())
        self.ui.refcat.setText("")
        self.ui.pro.setText("")
        self.ui.mun.setText("")
        self.ui.num.setText("")
        self.ui.area.setValue(0)
        self.ui.Selec.clear()
        self.ui.Selec.setColumnCount(0)
        self.ui.Selec.setRowCount(0)
        self.ui.Selec.setColumnCount(6)
        for col in range(6) : self.ui.Selec.setColumnWidth(col,int(str("20,120,25,35,60,60").split(",")[col]))
        self.ui.Selec.setHorizontalHeaderLabels(["S","RefCat","Pr.","Mun.","NumPar","Area"])
        self.ui.Selec.horizontalHeader().setFrameStyle(QFrame.Box | QFrame.Plain)
        self.ui.Selec.horizontalHeader().setLineWidth(1)
        self.geo = []
        fil = -1
        for elem in elems:
            self.geo.append(elem.geometry())
            if nRef != "":
                ref = elem[nRef]
            else:
                ref = ""
            num = "A"
            area = 0
            if nArea != "" : area=int(elem[nArea])
            if area == 0 : area=int(elem.geometry().area())
            if nPro == "id" : pro=capa.id()[13:15]
            elif nPro != "" : pro=str(elem[nPro])
            else : pro=""
            if nMun == "id" : mun=capa.id()[15:18]
            elif nMun != "" : mun=str(elem[nMun])
            else : mun = ""
            fil+=1
            self.ui.Selec.setRowCount(fil+1)
            self.ui.Selec.setItem(fil,0,QTableWidgetItem(str(fil)))
            self.ui.Selec.setItem(fil,1,QTableWidgetItem(str(ref)))
            self.ui.Selec.setItem(fil,2,QTableWidgetItem(str(pro)))
            self.ui.Selec.setItem(fil,3,QTableWidgetItem(str(mun)))
            self.ui.Selec.setItem(fil,4,QTableWidgetItem(str(num)))
            c=QTableWidgetItem(str(area))
            c.setTextAlignment(Qt.AlignRight)
            self.ui.Selec.setItem(fil,5,c)

        self.ui.Selec.resizeRowsToContents()
        if self.ui.desti.text().strip() == "":
            if self.project_dir != "":
                self.ui.desti.setText(self.project_dir)
            else:
                self.ui.desti.setText(self.plugin_dir)
        self.ui.Inspire3.setChecked(True)
        self.show()

