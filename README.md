# 📊 VBA Excel Framework — CRUD Reutilizable

Framework personal para proyectos de Excel con VBA. Permite crear
sistemas de registro, búsqueda y eliminación de datos usando
formularios, con una arquitectura modular que evita repetir código
entre proyectos.

---

## 🗂️ Estructura
📁 framework/
├── cEntrada.cls → Clase principal: Buscar, Insertar, Eliminar, Actualizar, Listar
├── modUtils.bas → Utilidades: validaciones, limpiar forms, generar IDs
📁 proyecto-ejemplo/
├── modMain.bas → Punto de entrada del proyecto
├── proyecto.xlsm → Excel con hojas y formularios
└── README.md

text

---

## ⚡ Uso rápido

### 1. Importar el framework
En el editor VBA (Alt+F11) → `File → Import File` → importa
`cEntrada.cls` y `modUtils.bas`

### 2. Configurar modMain.bas
```vba
Public entrada As New cEntrada

Sub AbrirMenu()
    entrada.Init "NombreHoja", 2, 1
    frmMenu.Show
End Sub
```

### 3. Parámetros de Init
```vba
entrada.Init "Alumnos",  2,  1
'             ↑           ↑   ↑
'         Nombre hoja   Fila  Columna
'                       inicio  del ID
```

---

## 🧱 Métodos disponibles — cEntrada

| Método | Descripción | Retorna |
|---|---|---|
| `Init(hoja, fila, col)` | Inicializa el objeto | — |
| `Buscar(valor)` | Busca por ID | `Long` (fila) o `-1` |
| `Insertar(a, b, c...)` | Agrega fila nueva | — |
| `Eliminar(id)` | Elimina fila por ID | `Boolean` |
| `Actualizar(id, col, val)` | Modifica una celda | `Boolean` |
| `Listar(col)` | Devuelve columna completa | `Variant()` |

---

## 🛠️ Utilidades — modUtils

| Función | Descripción |
|---|---|
| `Requerido(valor, campo)` | Valida que no esté vacío |
| `LimpiarForm(Me)` | Limpia todos los TextBox/ComboBox |
| `GenerarID(prefijo)` | ID único por timestamp |
| `Confirmar(mensaje)` | Cuadro Sí/No, retorna Boolean |
| `ResaltarFila(hoja, fila, color)` | Colorea una fila |

---

## 📋 Ejemplo real — Sistema de Alumnos

```vba
' Registrar
entrada.Insertar txtID.Value, txtNombre.Value, txtGrado.Value

' Buscar
Dim fila As Long
fila = entrada.Buscar(txtID.Value)

' Eliminar
entrada.Eliminar(txtID.Value)

' Actualizar columna 3
entrada.Actualizar "ID001", 3, "Nuevo valor"
```

---

## 📌 Notas importantes
- Guardar siempre como `.xlsm` (habilitado para macros)
- Siempre llamar `AbrirMenu()` desde el botón de la hoja,
  nunca abrir los forms directamente (causa error 91)
- El nombre de la hoja en `Init` debe ser exacto (mayúsculas incluidas)

---

## 🧠 Analogía
> `cEntrada` es equivalente a un modelo/DAO en Node.js.
> `modUtils` es equivalente a un archivo de helpers/utils.
> `modMain` es el controlador del proyecto.