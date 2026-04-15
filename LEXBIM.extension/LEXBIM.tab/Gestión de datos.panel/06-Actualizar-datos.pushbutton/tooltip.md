### Actualizar datos de elementos desde XML

Permite cargar un archivo **XML compatible con Excel** previamente exportado y actualizado, para escribir nuevamente los valores en los elementos de Revit.

La actualización se realiza a partir del **UniqueID** de cada elemento.

---

#### Flujo de uso

1. Se abre una interfaz para seleccionar el archivo **XML actualizado**.
2. Se carga el archivo y se valida su estructura.
3. Se presiona el botón **Actualizar**.
4. El script busca cada elemento en Revit por su **UniqueID**.
5. Se actualizan únicamente los parámetros permitidos.

---

#### Criterio de mapeo de elementos

El vínculo entre el archivo XML y los elementos del modelo se hace con base en:

- **UniqueID**

Este valor debe estar en la **primera columna** del archivo.

Si el script no encuentra esta columna en la primera posición, se mostrará un mensaje de error y la actualización no continuará.

---

#### Reglas de actualización

El script aplica las siguientes reglas:

- La **primera columna** debe ser `UniqueID`
- Las **3 primeras columnas** se omiten en la actualización:
  - `UniqueID`
  - `ElementId`
  - `FamilyType`
- Todas las columnas cuyo encabezado contenga **[No editable]** se omiten
- Las demás columnas sí se consideran para actualización

---

#### Formato esperado de encabezados

El archivo debe conservar los encabezados generados por el exportador, por ejemplo:

- `Instancia | Comments [Editable]`
- `Tipo | Type Comments [Editable]`
- `Tipo | Assembly Code [No editable]`

Esto permite identificar:

- si el parámetro es de **Instancia** o de **Tipo**
- si debe actualizarse o no

---

#### Qué parámetros se actualizan

Se actualizan:

- Parámetros de **Instancia**
- Parámetros de **Tipo**
- Columnas marcadas como:
  - **Editable**
  - **Mixto**

Se omiten:

- Columnas **No editable**
- Columnas base del archivo
- Parámetros de solo lectura
- Parámetros no encontrados en el elemento
- Parámetros de tipo `ElementId`
- Valores vacíos en parámetros numéricos

---

#### Consideraciones importantes

- El archivo XML debe provenir del exportador original para mantener la misma estructura
- No se recomienda modificar el nombre de los encabezados
- El mapeo depende completamente del **UniqueID**
- Si un elemento ya no existe en el modelo, se omite
- Si un parámetro no existe en un elemento concreto, se omite
- Los parámetros de tipo `ElementId` no se actualizan en esta versión
- Para parámetros numéricos, el script intenta convertir el valor desde texto
- Para parámetros de texto, se escribe directamente el contenido del XML

---

#### Compatibilidad de valores

El script intenta actualizar según el tipo de almacenamiento del parámetro:

- **String** → escritura directa
- **Integer** → conversión numérica / sí-no / valor textual
- **Double** → intenta `SetValueString()` y luego conversión numérica
- **ElementId** → omitido

---

#### Validaciones incluidas

Antes de actualizar, el script valida:

- que el archivo pueda leerse correctamente
- que exista una hoja y tabla válidas en el XML
- que la primera columna sea `UniqueID`
- que existan filas de datos para procesar

---

#### Resultado final

Al terminar, se muestra un resumen con:

- filas procesadas
- elementos encontrados
- elementos no encontrados
- parámetros actualizados
- parámetros omitidos
- errores de actualización

---

#### Recomendaciones

- Usa siempre el XML generado por el exportador del mismo flujo
- No cambies el orden de las primeras columnas
- No modifiques el texto de los encabezados
- Revisa especialmente las columnas **Editable** antes de cargar cambios
- Haz pruebas primero con pocos elementos antes de procesar todo el modelo

---

#### Uso previsto

Este script está pensado como complemento del flujo:

1. **Exportar datos desde Revit a XML**
2. **Editar la información fuera de Revit**
3. **Reimportar y actualizar parámetros en Revit**