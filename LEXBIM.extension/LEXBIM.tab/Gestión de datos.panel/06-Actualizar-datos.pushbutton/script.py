# -*- coding: utf-8 -*-

# ============================================================
# IMPORTACION DE BIBLIOTECAS
# ============================================================

import clr
import System
import xml.etree.ElementTree as ET

clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from pyrevit import revit

from System.Drawing import Point, Size
from System.Windows.Forms import (
    Form, Label, Button, TextBox, OpenFileDialog, MessageBox,
    MessageBoxButtons, MessageBoxIcon, FormStartPosition, DialogResult
)

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    ElementId,
    StorageType,
    BuiltInParameter
)

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document


# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def safe_str(value):
    try:
        if value is None:
            return u''
        return unicode(value)
    except:
        try:
            return str(value)
        except:
            return u''


def get_type_id_from_instance(e):
    try:
        tid = e.GetTypeId()
        if tid and tid != ElementId.InvalidElementId:
            return tid
    except:
        pass

    try:
        p = e.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
        if p:
            tid2 = p.AsElementId()
            if tid2 and tid2 != ElementId.InvalidElementId:
                return tid2
    except:
        pass

    return ElementId.InvalidElementId


def get_type_element(elem):
    tid = get_type_id_from_instance(elem)
    if tid and tid != ElementId.InvalidElementId:
        return doc.GetElement(tid)
    return None


def get_element_param_by_name(elem, param_name):
    try:
        p = elem.GetParameters(param_name)
        if p and p.Count > 0:
            return p[0]
    except:
        pass
    return None


def parse_header_metadata(header):
    header = safe_str(header)

    result = {'scope': None, 'param_name': None, 'status': None}

    if '[' in header and ']' in header:
        status = header.split('[')[-1].replace(']', '').strip()
        header = header.split('[')[0].strip()
        result['status'] = status

    parts = [p.strip() for p in header.split('|')]

    if len(parts) >= 2:
        if parts[0].lower() == 'instancia':
            result['scope'] = 'instance'
        elif parts[0].lower() == 'tipo':
            result['scope'] = 'type'

        result['param_name'] = parts[1]

    return result


def set_param(param, value):
    if not param:
        return False

    if param.IsReadOnly:
        return False

    value = safe_str(value)

    try:
        if param.StorageType == StorageType.String:
            param.Set(value)
            return True

        elif param.StorageType == StorageType.Integer:
            try:
                param.Set(int(value))
                return True
            except:
                return False

        elif param.StorageType == StorageType.Double:
            try:
                param.SetValueString(value)
                return True
            except:
                return False

    except:
        return False

    return False


def read_xml(filepath):
    ns = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}

    tree = ET.parse(filepath)
    root = tree.getroot()

    table = root.find('.//ss:Table', ns)
    rows = table.findall('ss:Row', ns)

    data = []
    for r in rows:
        row = []
        for c in r.findall('ss:Cell', ns):
            d = c.find('ss:Data', ns)
            row.append(safe_str(d.text) if d is not None else u'')
        data.append(row)

    headers = data[0]
    values = data[1:]

    return headers, values


def build_update_map(headers):
    cols = []

    for i, h in enumerate(headers):
        if i < 3:
            continue

        meta = parse_header_metadata(h)

        if not meta['scope'] or not meta['param_name']:
            continue

        if safe_str(meta['status']).lower() == 'no editable':
            continue

        cols.append((i, meta['scope'], meta['param_name']))

    return cols


def get_elements_map():
    data = {}
    for e in FilteredElementCollector(doc).WhereElementIsNotElementType():
        try:
            data[e.UniqueId] = e
        except:
            pass
    return data


# ============================================================
# UI
# ============================================================

class LoadForm(Form):
    def __init__(self):
        self.Text = 'Actualizar desde XML'
        self.Width = 600
        self.Height = 180
        self.StartPosition = FormStartPosition.CenterScreen

        self.path = None

        # TextBox
        self.txt = TextBox()
        self.txt.Location = Point(10, 20)
        self.txt.Size = Size(450, 25)
        self.Controls.Add(self.txt)

        # Botón Examinar
        btn = Button()
        btn.Text = 'Examinar'
        btn.Location = Point(470, 18)
        btn.Size = Size(100, 25)
        btn.Click += self.browse
        self.Controls.Add(btn)

        # Botón Actualizar
        run = Button()
        run.Text = 'Actualizar'
        run.Location = Point(350, 70)
        run.Size = Size(100, 30)
        run.Click += self.execute
        self.Controls.Add(run)

        # ✅ Botón Cancelar (nuevo)
        cancel = Button()
        cancel.Text = 'Cancelar'
        cancel.Location = Point(460, 70)
        cancel.Size = Size(100, 30)
        cancel.Click += self.on_cancel
        self.Controls.Add(cancel)

    def browse(self, sender, args):
        dlg = OpenFileDialog()
        dlg.Filter = 'XML (*.xml)|*.xml'

        if dlg.ShowDialog() == DialogResult.OK:
            self.txt.Text = dlg.FileName
            self.path = dlg.FileName

    def execute(self, sender, args):
        if not self.txt.Text:
            MessageBox.Show(
                'Selecciona un archivo XML.',
                'Aviso',
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        self.path = self.txt.Text
        self.DialogResult = DialogResult.OK
        self.Close()

    # ✅ Método cancelar
    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()

# ============================================================
# PROCESO
# ============================================================

form = LoadForm()
if form.ShowDialog() != DialogResult.OK:
    raise SystemExit

headers, rows = read_xml(form.path)

if headers[0] != 'UniqueID':
    MessageBox.Show('Error: falta columna UniqueID')
    raise SystemExit

update_cols = build_update_map(headers)
elements = get_elements_map()

updated = 0
missing = 0

with revit.Transaction('Update from XML'):
    for r in rows:
        uid = safe_str(r[0])

        elem = elements.get(uid)
        if not elem:
            missing += 1
            continue

        for col in update_cols:
            idx, scope, pname = col

            val = r[idx] if idx < len(r) else ''

            target = elem if scope == 'instance' else get_type_element(elem)
            p = get_element_param_by_name(target, pname)

            if set_param(p, val):
                updated += 1


# ============================================================
# MENSAJE FINAL
# ============================================================

MessageBox.Show(
    u'Actualización completada\n\nActualizados: %s\nNo encontrados: %s' % (updated, missing),
    'Resultado'
)