# VBA Excel Framework — CRUD Reutilizable

Framework personal para crear proyectos de Excel con VBA de forma modular y reutilizable.  
Permite trabajar con formularios, búsqueda, inserción, actualización y eliminación de registros sin repetir lógica en cada proyecto.

---

## ¿Qué incluye?

- **`cEntrada.cls`**: clase principal para manejar registros en una hoja de Excel.
- **`modUtils.bas`**: funciones de apoyo como validación y limpieza de formularios.
- **`modMain.bas`**: punto de entrada del proyecto.
- **UserForms**: interfaz visual para registrar y buscar datos.

---

## Estructura del proyecto

```text
R_FRAMEWORK/
├── README.md
├── cEntrada.cls
├── modMain.bas
├── modUtils.bas
├── RT_FRAMEWORK.xlsm
└── EXCELDB_FRAMEWORK/
```

---

## Requisitos

- Microsoft Excel con soporte para macros.
- Guardar el archivo como **`.xlsm`**.
- Tener habilitada la pestaña **Desarrollador**.

---

## Cómo funciona

El flujo general es:

1. El usuario presiona un botón en la hoja.
2. Se ejecuta `AbrirMenu`.
3. Se inicializa el objeto `entrada`.
4. Se abre el formulario principal.
5. Desde los formularios se puede registrar o buscar información.

---

## Inicialización

En `modMain.bas`:

```vba
Public entrada As New cEntrada

Sub AbrirMenu()
    entrada.Init "Alumnos", 2, 1
    frmMenu.Show
End Sub
```

### Parámetros de `Init`

```vba
entrada.Init "Alumnos", 2, 1
```

- **"Alumnos"**: nombre exacto de la hoja donde se guardan los datos.
- **`2`**: fila donde empiezan los registros.
- **`1`**: columna donde está el ID o campo principal de búsqueda.

---

## Hoja de datos

La hoja debe tener una estructura similar a esta:

| A | B | C |
|---|---|---|
| ID | Nombre | Grado |

La primera fila se usa como encabezado y los datos comienzan desde la fila 2.

---

## Uso básico

### Registrar un alumno

```vba
entrada.Insertar txtID.Value, txtNombre.Value, txtGrado.Value
```

### Buscar un alumno

```vba
Dim fila As Long
fila = entrada.Buscar(txtID.Value)
```

### Eliminar un alumno

```vba
entrada.Eliminar txtID.Value
```

### Actualizar un registro

```vba
entrada.Actualizar "ID001", 3, "Nuevo valor"
```

---

## Ejemplo del formulario de registro

```vba
Private Sub btnGuardar_Click()
    If Not modUtils.Requerido(txtID.Value, "ID") Then Exit Sub
    If Not modUtils.Requerido(txtNombre.Value, "Nombre") Then Exit Sub

    entrada.Insertar txtID.Value, txtNombre.Value, txtGrado.Value
    MsgBox "Alumno registrado: " & txtNombre.Value

    modUtils.LimpiarForm Me
End Sub
```

---

## Ejemplo del formulario de búsqueda

```vba
Private Sub btnBuscar_Click()
    Dim fila As Long
    fila = entrada.Buscar(txtID.Value)

    If fila > 0 Then
        With ThisWorkbook.Sheets("Alumnos")
            lblNombre.Caption = "Nombre: " & .Cells(fila, 2).Value
            lblGrado.Caption = "Grado: " & .Cells(fila, 3).Value
        End With
    Else
        MsgBox "Alumno no encontrado"
    End If
End Sub
```

---

## Importar el framework a otro proyecto

1. Abre el editor de VBA con `Alt + F11`.
2. Ve a **Archivo → Importar archivo**.
3. Importa:
   - `cEntrada.cls`
   - `modUtils.bas`
4. Crea tu propio `modMain.bas` para ese proyecto.
5. Ajusta el nombre de la hoja y la fila/columna en `Init`.

---

## Notas importantes

- Si abres un formulario directamente sin pasar por `AbrirMenu`, puede aparecer el error de objeto no establecido.
- Asegúrate de que el nombre del form coincida exactamente con el código.
- Si el nombre de la hoja está mal escrito, `Init` fallará.
- Guarda siempre el archivo como **`.xlsm`**.

---

## Objetivo

La idea de este framework es evitar repetir lógica en cada proyecto y construir sistemas pequeños pero escalables usando Excel y VBA.

---

## Autor

Creado como base personal de trabajo para proyectos escolares y de práctica.

---

## Licencia

Uso personal y educativo.