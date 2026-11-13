"""Sonda R9 (przegląd, zarzut 1): czy źródło warstwy (publicSource) zawiera filtr warstwy?

scripts/qgis-probe.sh 4.2.3 sondy/r09_source_subset.py
"""

from qgis.core import QgsFeature, QgsProject, QgsVectorFileWriter, QgsVectorLayer
from qgis.testing import start_app

app = start_app()
mem = QgsVectorLayer("Point?crs=EPSG:2180&field=email:string", "m", "memory")
feature = QgsFeature(mem.fields())
feature.setAttributes(["anna@example.com"])
mem.dataProvider().addFeature(feature)
options = QgsVectorFileWriter.SaveVectorOptions()
options.driverName, options.layerName = "GPKG", "p"
QgsVectorFileWriter.writeAsVectorFormatV3(mem, "/tmp/p.gpkg", QgsProject.instance().transformContext(), options)
for uri in ("/tmp/p.gpkg|layername=p", "/tmp/p.gpkg|layername=p|subset=\"email\" = 'x@y.pl'"):
    layer = QgsVectorLayer(uri, "p", "ogr")
    layer.setSubsetString("\"email\" = 'anna@example.com'")
    print("tryb=source", repr(layer.publicSource()), "|", repr(layer.source()))
pg_uri = "dbname='d' host=h user=u password=secret table=\"public\".\"t\" sql=\"email\" = 'anna@example.com'"
print("tryb=pg", repr(QgsVectorLayer(pg_uri, "pg", "postgres").publicSource()))
