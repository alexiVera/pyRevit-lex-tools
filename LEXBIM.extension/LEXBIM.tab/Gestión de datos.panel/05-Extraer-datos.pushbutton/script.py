# -*- coding: utf-8 -*-

# ============================================================
# CARGA DE LIBRERIAS
# ============================================================

import clr
import System

clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

from System.Drawing import Point, Size
from System.IO import StreamWriter
from System.Security import SecurityElement
from System.Windows.Forms import (
    Form, Label, Button, CheckedListBox, RadioButton,
    GroupBox, DialogResult, SaveFileDialog, MessageBox,
    MessageBoxButtons, MessageBoxIcon, FormStartPosition
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
# DEFINICION DE FUNCIONES AUXILIARES
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


def xml_escape(text):
    try:
        escaped = SecurityElement.Escape(safe_str(text))
        if escaped:
            return escaped
    except:
        pass
    return u''


def get_category_name(elem):
    try:
        if elem.Category:
            return elem.Category.Name
    except:
        pass
    return u'Sin categoría'


def get_type_id_from_instance(e):
    """Alternativa robusta a GetTypeId(): usa ELEM_TYPE_PARAM si GetTypeId es inválido."""
    tid = None

    try:
        tid = e.GetTypeId()
        if tid and tid != ElementId.InvalidElementId:
            return tid
    except:
        tid = None

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
        try:
            return doc.GetElement(tid)
        except:
            pass
    return None


def get_type_name_from_type_elem(t):
    """Obtiene nombre del tipo con varios fallbacks."""
    if not t:
        return u''

    for bip in (BuiltInParameter.ALL_MODEL_TYPE_NAME, BuiltInParameter.SYMBOL_NAME_PARAM):
        try:
            p = t.get_Parameter(bip)
            if p:
                s = p.AsString()
                if s:
                    return safe_str(s)
        except:
            pass

    try:
        if hasattr(t, 'Name') and t.Name:
            return safe_str(t.Name)
    except:
        pass

    return u''


def get_family_name_from_type_elem(t):
    """Obtiene nombre de familia con fallbacks para familias cargables y de sistema."""
    if not t:
        return u''

    try:
        if hasattr(t, 'FamilyName') and t.FamilyName:
            return safe_str(t.FamilyName)
    except:
        pass

    try:
        if hasattr(t, 'Family') and t.Family:
            if t.Family.Name:
                return safe_str(t.Family.Name)
    except:
        pass

    try:
        p = t.get_Parameter(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)
        if p:
            s = p.AsString()
            if s:
                return safe_str(s)
    except:
        pass

    try:
        if t.Category and t.Category.Name:
            return safe_str(t.Category.Name)
    except:
        pass

    return u''


def get_family_type_name(elem):
    """Devuelve FamilyName : TypeName con varios fallbacks."""
    type_elem = get_type_element(elem)

    if not type_elem:
        try:
            return safe_str(elem.Name)
        except:
            return u''

    family_name = get_family_name_from_type_elem(type_elem)
    type_name = get_type_name_from_type_elem(type_elem)

    if family_name and type_name:
        return u'%s : %s' % (family_name, type_name)
    elif type_name:
        return type_name
    elif family_name:
        return family_name

    try:
        return safe_str(type_elem.Name)
    except:
        pass

    try:
        return safe_str(elem.Name)
    except:
        pass

    return u''


def get_element_param_by_name(elem, param_name):
    try:
        params = elem.GetParameters(param_name)
        if params and params.Count > 0:
            return params[0]
    except:
        pass
    return None


def get_param_value_as_text(param):
    if not param:
        return u''

    try:
        value_string = param.AsValueString()
        if value_string:
            return safe_str(value_string)
    except:
        pass

    try:
        st = param.StorageType

        if st == StorageType.String:
            return safe_str(param.AsString())

        elif st == StorageType.Integer:
            return safe_str(param.AsInteger())

        elif st == StorageType.Double:
            return safe_str(param.AsDouble())

        elif st == StorageType.ElementId:
            eid = param.AsElementId()
            if eid and eid != ElementId.InvalidElementId:
                ref_elem = doc.GetElement(eid)
                if ref_elem:
                    try:
                        return safe_str(ref_elem.Name)
                    except:
                        return safe_str(eid.IntegerValue)
                return safe_str(eid.IntegerValue)
            return u''
    except:
        pass

    return u''


def get_param_value(elem, scope, param_name):
    target = elem
    if scope == 'type':
        target = get_type_element(elem)

    if not target:
        return u''

    param = get_element_param_by_name(target, param_name)
    return get_param_value_as_text(param)


def collect_model_elements():
    return list(
        FilteredElementCollector(doc)
        .WhereElementIsNotElementType()
        .ToElements()
    )


def get_available_categories(elements):
    data = {}
    for elem in elements:
        cat_name = get_category_name(elem)
        if cat_name not in data:
            data[cat_name] = []
        data[cat_name].append(elem)
    return data


def get_type_key(elem):
    type_elem = get_type_element(elem)
    cat_name = get_category_name(elem)
    fam_type = get_family_type_name(elem)

    if type_elem:
        try:
            return (
                cat_name,
                fam_type if fam_type else safe_str(type_elem.Id.IntegerValue),
                type_elem.Id.IntegerValue
            )
        except:
            pass

    return (cat_name, u'Sin tipo', -1)


def get_available_types(elements):
    data = {}

    for elem in elements:
        key = get_type_key(elem)
        display = u'%s | %s' % (key[0], key[1])

        if display not in data:
            data[display] = {
                'type_id': key[2],
                'elements': []
            }

        data[display]['elements'].append(elem)

    return data


def parameter_status_text(readonly_values):
    has_true = True in readonly_values
    has_false = False in readonly_values

    if has_true and has_false:
        return u'Mixto'
    elif has_true:
        return u'No editable'
    else:
        return u'Editable'


def collect_available_parameters(elements):
    param_map = {}

    for elem in elements:
        try:
            for p in elem.Parameters:
                try:
                    name = p.Definition.Name
                except:
                    continue

                key = ('instance', name)

                if key not in param_map:
                    param_map[key] = {
                        'scope': 'instance',
                        'name': name,
                        'readonly_values': []
                    }

                try:
                    param_map[key]['readonly_values'].append(bool(p.IsReadOnly))
                except:
                    param_map[key]['readonly_values'].append(True)
        except:
            pass

        type_elem = get_type_element(elem)
        if type_elem:
            try:
                for p in type_elem.Parameters:
                    try:
                        name = p.Definition.Name
                    except:
                        continue

                    key = ('type', name)

                    if key not in param_map:
                        param_map[key] = {
                            'scope': 'type',
                            'name': name,
                            'readonly_values': []
                        }

                    try:
                        param_map[key]['readonly_values'].append(bool(p.IsReadOnly))
                    except:
                        param_map[key]['readonly_values'].append(True)
            except:
                pass

    result = []
    keys = sorted(param_map.keys(), key=lambda x: (x[0], x[1].lower()))

    for key in keys:
        item = param_map[key]
        status = parameter_status_text(item['readonly_values'])
        scope_label = u'Instancia' if item['scope'] == 'instance' else u'Tipo'

        result.append({
            'scope': item['scope'],
            'name': item['name'],
            'status': status,
            'display': u'%s | %s [%s]' % (scope_label, item['name'], status)
        })

    return result


def get_elements_from_source(source_mode, source_items, categories_data, types_data):
    elements = []

    if source_mode == 'categories':
        for name in source_items:
            if name in categories_data:
                elements.extend(categories_data[name])

    elif source_mode == 'types':
        for name in source_items:
            if name in types_data:
                elements.extend(types_data[name]['elements'])

    elif source_mode == 'selection':
        sel_ids = uidoc.Selection.GetElementIds()
        for eid in sel_ids:
            elem = doc.GetElement(eid)
            if elem:
                elements.append(elem)

    unique = {}
    final_elements = []

    for elem in elements:
        try:
            eid = elem.Id.IntegerValue
            if eid not in unique:
                unique[eid] = True
                final_elements.append(elem)
        except:
            pass

    return final_elements


def choose_save_path():
    dlg = SaveFileDialog()
    dlg.Title = 'Guardar archivo Excel'
    dlg.Filter = 'Excel XML (*.xml)|*.xml'
    dlg.FileName = 'Extraccion_Elementos.xml'

    if dlg.ShowDialog() == DialogResult.OK:
        return dlg.FileName

    return None


def export_to_excel(elements, selected_parameters, filepath):
    writer = None

    try:
        headers = [u'UniqueID', u'ElementId', u'FamilyType']

        for p in selected_parameters:
            scope_label = u'Instancia' if p['scope'] == 'instance' else u'Tipo'
            headers.append(u'%s | %s [%s]' % (scope_label, p['name'], p['status']))

        writer = StreamWriter(filepath, False, System.Text.Encoding.UTF8)

        writer.WriteLine(u'<?xml version="1.0"?>')
        writer.WriteLine(u'<?mso-application progid="Excel.Sheet"?>')
        writer.WriteLine(u'<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"')
        writer.WriteLine(u' xmlns:o="urn:schemas-microsoft-com:office:office"')
        writer.WriteLine(u' xmlns:x="urn:schemas-microsoft-com:office:excel"')
        writer.WriteLine(u' xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"')
        writer.WriteLine(u' xmlns:html="http://www.w3.org/TR/REC-html40">')

        writer.WriteLine(u'<Styles>')
        writer.WriteLine(u'  <Style ss:ID="Header">')
        writer.WriteLine(u'    <Font ss:Bold="1"/>')
        writer.WriteLine(u'    <Interior ss:Color="#D9EAF7" ss:Pattern="Solid"/>')
        writer.WriteLine(u'  </Style>')
        writer.WriteLine(u'</Styles>')

        writer.WriteLine(u'<Worksheet ss:Name="Datos">')
        writer.WriteLine(u'<Table>')

        writer.WriteLine(u'<Row>')
        for header in headers:
            writer.WriteLine(
                u'<Cell ss:StyleID="Header"><Data ss:Type="String">%s</Data></Cell>' % xml_escape(header)
            )
        writer.WriteLine(u'</Row>')

        for elem in elements:
            row_values = []

            try:
                row_values.append(safe_str(elem.UniqueId))
            except:
                row_values.append(u'')

            try:
                row_values.append(safe_str(elem.Id.IntegerValue))
            except:
                row_values.append(u'')

            row_values.append(get_family_type_name(elem))

            for p in selected_parameters:
                row_values.append(get_param_value(elem, p['scope'], p['name']))

            writer.WriteLine(u'<Row>')
            for value in row_values:
                writer.WriteLine(
                    u'<Cell><Data ss:Type="String">%s</Data></Cell>' % xml_escape(value)
                )
            writer.WriteLine(u'</Row>')

        writer.WriteLine(u'</Table>')
        writer.WriteLine(u'</Worksheet>')
        writer.WriteLine(u'</Workbook>')

        writer.Close()
        writer = None

        return True, u'Exportación completada:\n%s' % filepath

    except Exception as ex:
        try:
            if writer:
                writer.Close()
        except:
            pass

        return False, u'Error al exportar a Excel:\n%s' % safe_str(ex)


# ============================================================
# DEFINICION DE CUADROS DE DIALOGO
# ============================================================

class SourceSelectionForm(Form):
    def __init__(self, categories_data, types_data, active_count):
        self.Text = 'Extracción de datos'
        self.Width = 700
        self.Height = 620
        self.StartPosition = FormStartPosition.CenterScreen

        self.result_mode = None
        self.result_items = []

        self.categories_data = categories_data
        self.types_data = types_data

        lbl = Label()
        lbl.Text = 'Selecciona el origen de elementos'
        lbl.Location = Point(15, 15)
        lbl.Size = Size(300, 20)
        self.Controls.Add(lbl)

        self.group = GroupBox()
        self.group.Text = 'Modo de selección'
        self.group.Location = Point(15, 45)
        self.group.Size = Size(650, 85)
        self.Controls.Add(self.group)

        self.rb_categories = RadioButton()
        self.rb_categories.Text = 'Categorías completas'
        self.rb_categories.Location = Point(15, 25)
        self.rb_categories.Size = Size(180, 20)
        self.rb_categories.Checked = True
        self.rb_categories.CheckedChanged += self.on_mode_changed
        self.group.Controls.Add(self.rb_categories)

        self.rb_types = RadioButton()
        self.rb_types.Text = 'Tipos'
        self.rb_types.Location = Point(220, 25)
        self.rb_types.Size = Size(120, 20)
        self.rb_types.CheckedChanged += self.on_mode_changed
        self.group.Controls.Add(self.rb_types)

        self.rb_selection = RadioButton()
        self.rb_selection.Text = 'Selección activa (%s elementos)' % active_count
        self.rb_selection.Location = Point(380, 25)
        self.rb_selection.Size = Size(230, 20)
        self.rb_selection.CheckedChanged += self.on_mode_changed
        self.group.Controls.Add(self.rb_selection)

        info = Label()
        info.Text = 'Para Categorías o Tipos puedes marcar varios elementos de la lista.'
        info.Location = Point(15, 145)
        info.Size = Size(500, 20)
        self.Controls.Add(info)

        self.listbox = CheckedListBox()
        self.listbox.Location = Point(15, 170)
        self.listbox.Size = Size(650, 340)
        self.listbox.CheckOnClick = True
        self.Controls.Add(self.listbox)

        self.btn_ok = Button()
        self.btn_ok.Text = 'Continuar'
        self.btn_ok.Location = Point(470, 525)
        self.btn_ok.Size = Size(90, 30)
        self.btn_ok.Click += self.on_ok
        self.Controls.Add(self.btn_ok)

        self.btn_cancel = Button()
        self.btn_cancel.Text = 'Cancelar'
        self.btn_cancel.Location = Point(575, 525)
        self.btn_cancel.Size = Size(90, 30)
        self.btn_cancel.Click += self.on_cancel
        self.Controls.Add(self.btn_cancel)

        self.load_categories()

    def clear_list(self):
        self.listbox.Items.Clear()

    def load_categories(self):
        self.clear_list()
        for name in sorted(self.categories_data.keys()):
            self.listbox.Items.Add(name, False)
        self.listbox.Enabled = True

    def load_types(self):
        self.clear_list()
        for name in sorted(self.types_data.keys()):
            self.listbox.Items.Add(name, False)
        self.listbox.Enabled = True

    def load_selection_mode(self):
        self.clear_list()
        self.listbox.Items.Add('Se usará la selección activa de Revit', True)
        self.listbox.Enabled = False

    def on_mode_changed(self, sender, args):
        if self.rb_categories.Checked:
            self.load_categories()
        elif self.rb_types.Checked:
            self.load_types()
        elif self.rb_selection.Checked:
            self.load_selection_mode()

    def on_ok(self, sender, args):
        checked = []

        if self.rb_categories.Checked:
            self.result_mode = 'categories'
            for item in self.listbox.CheckedItems:
                checked.append(safe_str(item))

        elif self.rb_types.Checked:
            self.result_mode = 'types'
            for item in self.listbox.CheckedItems:
                checked.append(safe_str(item))

        elif self.rb_selection.Checked:
            self.result_mode = 'selection'
            checked = ['selection']

        if self.result_mode in ['categories', 'types'] and not checked:
            MessageBox.Show(
                'Debes seleccionar al menos un elemento de la lista.',
                'Aviso',
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        self.result_items = checked
        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


class ParameterSelectionForm(Form):
    def __init__(self, parameters):
        self.Text = 'Selección de parámetros'
        self.Width = 800
        self.Height = 650
        self.StartPosition = FormStartPosition.CenterScreen

        self.parameters = parameters
        self.selected_parameters = []

        lbl = Label()
        lbl.Text = 'Selecciona los parámetros a exportar'
        lbl.Location = Point(15, 15)
        lbl.Size = Size(350, 20)
        self.Controls.Add(lbl)

        fixed = Label()
        fixed.Text = 'Siempre se incluirán: UniqueID, ElementId y FamilyType'
        fixed.Location = Point(15, 40)
        fixed.Size = Size(450, 20)
        self.Controls.Add(fixed)

        info = Label()
        info.Text = 'El encabezado incluirá si el parámetro es Editable, No editable o Mixto.'
        info.Location = Point(15, 65)
        info.Size = Size(520, 20)
        self.Controls.Add(info)

        self.listbox = CheckedListBox()
        self.listbox.Location = Point(15, 95)
        self.listbox.Size = Size(750, 460)
        self.listbox.CheckOnClick = True
        self.Controls.Add(self.listbox)

        for p in self.parameters:
            self.listbox.Items.Add(p['display'], False)

        self.btn_all = Button()
        self.btn_all.Text = 'Marcar todo'
        self.btn_all.Location = Point(15, 570)
        self.btn_all.Size = Size(100, 30)
        self.btn_all.Click += self.on_all
        self.Controls.Add(self.btn_all)

        self.btn_none = Button()
        self.btn_none.Text = 'Desmarcar'
        self.btn_none.Location = Point(125, 570)
        self.btn_none.Size = Size(100, 30)
        self.btn_none.Click += self.on_none
        self.Controls.Add(self.btn_none)

        self.btn_ok = Button()
        self.btn_ok.Text = 'Exportar'
        self.btn_ok.Location = Point(570, 570)
        self.btn_ok.Size = Size(90, 30)
        self.btn_ok.Click += self.on_ok
        self.Controls.Add(self.btn_ok)

        self.btn_cancel = Button()
        self.btn_cancel.Text = 'Cancelar'
        self.btn_cancel.Location = Point(675, 570)
        self.btn_cancel.Size = Size(90, 30)
        self.btn_cancel.Click += self.on_cancel
        self.Controls.Add(self.btn_cancel)

    def on_all(self, sender, args):
        for i in range(self.listbox.Items.Count):
            self.listbox.SetItemChecked(i, True)

    def on_none(self, sender, args):
        for i in range(self.listbox.Items.Count):
            self.listbox.SetItemChecked(i, False)

    def on_ok(self, sender, args):
        indices = []

        for i in range(self.listbox.CheckedIndices.Count):
            indices.append(self.listbox.CheckedIndices[i])

        if not indices:
            MessageBox.Show(
                'Debes seleccionar al menos un parámetro.',
                'Aviso',
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            )
            return

        for idx in indices:
            self.selected_parameters.append(self.parameters[idx])

        self.DialogResult = DialogResult.OK
        self.Close()

    def on_cancel(self, sender, args):
        self.DialogResult = DialogResult.Cancel
        self.Close()


# ============================================================
# PASO A PASO DE LA AUTOMATIZACION
# ============================================================

all_elements = collect_model_elements()

if not all_elements:
    MessageBox.Show(
        'No se encontraron elementos de modelo para procesar.',
        'Aviso',
        MessageBoxButtons.OK,
        MessageBoxIcon.Warning
    )
    raise SystemExit

active_ids = uidoc.Selection.GetElementIds()
active_count = active_ids.Count

categories_data = get_available_categories(all_elements)
types_data = get_available_types(all_elements)

source_form = SourceSelectionForm(categories_data, types_data, active_count)
source_result = source_form.ShowDialog()

if source_result != DialogResult.OK:
    raise SystemExit

selected_elements = get_elements_from_source(
    source_form.result_mode,
    source_form.result_items,
    categories_data,
    types_data
)

if source_form.result_mode == 'selection' and not selected_elements:
    MessageBox.Show(
        'No hay elementos en la selección activa.',
        'Aviso',
        MessageBoxButtons.OK,
        MessageBoxIcon.Warning
    )
    raise SystemExit

if not selected_elements:
    MessageBox.Show(
        'No se encontraron elementos con el criterio seleccionado.',
        'Aviso',
        MessageBoxButtons.OK,
        MessageBoxIcon.Warning
    )
    raise SystemExit

available_parameters = collect_available_parameters(selected_elements)

if not available_parameters:
    MessageBox.Show(
        'No se encontraron parámetros para exportar.',
        'Aviso',
        MessageBoxButtons.OK,
        MessageBoxIcon.Warning
    )
    raise SystemExit

param_form = ParameterSelectionForm(available_parameters)
param_result = param_form.ShowDialog()

if param_result != DialogResult.OK:
    raise SystemExit

save_path = choose_save_path()
if not save_path:
    raise SystemExit

success, final_message = export_to_excel(
    selected_elements,
    param_form.selected_parameters,
    save_path
)


# ============================================================
# CUADRO DE DIALOGO FINAL
# ============================================================

if success:
    MessageBox.Show(
        final_message,
        'Éxito',
        MessageBoxButtons.OK,
        MessageBoxIcon.Information
    )
else:
    MessageBox.Show(
        final_message,
        'Error',
        MessageBoxButtons.OK,
        MessageBoxIcon.Error
    )