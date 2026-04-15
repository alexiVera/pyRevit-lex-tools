### Extraer datos de elementos a Excel

Permite extraer información de elementos del modelo y exportarla a un archivo **Excel (.xlsx)**.

La herramienta abre una interfaz en varios pasos para definir:

1. **Origen de los elementos**
   - **Categorías completas**
   - **Tipos**
   - **Selección activa** en Revit

2. **Parámetros a exportar**
   - Muestra una lista de parámetros disponibles según los elementos encontrados.
   - Permite seleccionar uno o varios parámetros mediante **checkboxes**.

3. **Exportación**
   - Genera un archivo **Excel** con los datos de los elementos seleccionados.

---

#### Campos que siempre se exportan

La tabla de Excel siempre incluye estas columnas base:

- **UniqueID**
- **ElementId**
- **FamilyType**

Además, se agregan las columnas de los parámetros seleccionados por el usuario.

---

#### Encabezados con información de edición

Cada parámetro exportado se escribe en el encabezado con una marca que indica su condición:

- **Editable**
- **No editable**
- **Mixto**

Esto permite identificar fácilmente qué campos podrían usarse más adelante en un flujo de **reimportación de datos desde Excel hacia Revit**.

Ejemplo de encabezado:

- `Instancia | Comments [Editable]`
- `Tipo | Type Comments [No editable]`

---

#### Flujo de uso

1. Selecciona el origen de los elementos.
2. Marca las categorías, tipos o usa la selección activa.
3. Elige los parámetros a extraer.
4. Define la ruta del archivo Excel.
5. Se genera la tabla con la información recopilada.

---

#### Consideraciones importantes

- El script trabaja con **elementos colocados en el modelo**, no con tipos sueltos del navegador.
- En el modo **Tipos**, se exportan los elementos del modelo que pertenecen a los tipos seleccionados.
- En el modo **Selección activa**, debes tener elementos previamente seleccionados en Revit.
- Los parámetros mostrados pueden ser de:
  - **Instancia**
  - **Tipo**
- El estado **Mixto** aparece cuando un mismo parámetro no tiene el mismo comportamiento en todos los elementos analizados.
- Para generar el archivo `.xlsx`, el equipo debe tener **Microsoft Excel instalado**, ya que la exportación usa interoperabilidad COM.
- Los valores exportados se escriben como texto visible, priorizando el valor formateado de Revit cuando está disponible.
- Algunos parámetros pueden aparecer vacíos si:
  - no aplican a todos los elementos,
  - no tienen valor,
  - o el elemento no dispone de tipo asociado.

---

#### Recomendaciones

- Usa una selección de elementos coherente para obtener una lista de parámetros más limpia.
- Si el objetivo es reutilizar el Excel para una futura carga de datos, conviene mantener los encabezados originales sin modificarlos.
- Se recomienda revisar especialmente las columnas marcadas como **Editable**, ya que son las más adecuadas para un flujo de escritura de vuelta a Revit.

---

#### Resultado esperado

Se obtiene un archivo Excel estructurado para:

- **consulta**
- **control**
- **revisión**
- futura **sincronización de parámetros con Revit**