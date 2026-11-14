"""Sonda R10: czy createCache formattera ValueRelation da się przenieść do Pythona."""
import sys

from qgis.core import (
    QgsApplication,
    QgsEditorWidgetSetup,
    QgsFeature,
    QgsProject,
    QgsValueRelationFieldFormatter,
    QgsVectorLayer,
)

app = QgsApplication([], False)
app.initQgis()
ref = QgsVectorLayer("None?field=id:integer&field=name:string", "slownik", "memory")
for i, n in ((1, "las"), (2, "łąka")):
    f = QgsFeature(ref.fields())
    f.setAttributes([i, n])
    ref.dataProvider().addFeature(f)
QgsProject.instance().addMapLayer(ref)
lay = QgsVectorLayer("None?field=kod:integer&field=multi:string", "dane", "memory")
for a, b in ((1, "{1,2}"), (2, "{2}"), (None, None)):
    f = QgsFeature(lay.fields())
    f.setAttributes([a, b])
    lay.dataProvider().addFeature(f)
QgsProject.instance().addMapLayer(lay)
base = {"Layer": ref.id(), "LayerName": "slownik", "Key": "id", "Value": "name"}
lay.setEditorWidgetSetup(0, QgsEditorWidgetSetup("ValueRelation", dict(base, AllowMulti=False)))
lay.setEditorWidgetSetup(1, QgsEditorWidgetSetup("ValueRelation", dict(base, AllowMulti=True)))
fmt = QgsApplication.fieldFormatterRegistry().fieldFormatter("ValueRelation")
for idx in (0, 1):
    cfg = lay.editorWidgetSetup(idx).config()
    try:
        c = fmt.createCache(lay, idx, cfg)
        print("createCache ok", type(c))
    except TypeError as e:
        print("createCache TypeError:", e)
    for feat in lay.getFeatures():
        raw = feat.attributes()[idx]
        print("  representValue(None cache)", repr(raw), "->", repr(fmt.representValue(lay, idx, cfg, None, raw)))
print("valueToStringList", QgsValueRelationFieldFormatter.valueToStringList("{1,2}"))
print("tryb=ok")
sys.stdout.flush()
