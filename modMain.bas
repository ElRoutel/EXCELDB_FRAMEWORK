Attribute VB_Name = "modMain"
'--- modMain
'USO

'Init acepta 3 parámetros, mira la firma en cEntrada.cls

'Public Sub Init(nombreHoja As String, filaInicio As Long, colID As Integer)
'\\\\\\\\\
'Parámetro 1 -> "Alumnos"
'El nombre exacto de la hoja donde están los datos. Si tu hoja se llama "Productos" o "Empleados", cambias el string.

'\\\\\\\\\
'Parámetro 2 -> 2
'La fila donde empiezan los datos. Como la fila 1 son los encabezados (ID, Nombre, Grado), los datos reales empiezan en la fila 2. Si tuvieras 2 filas de encabezados, pondrías 3.

'Fila 1 -> ID | Nombre | Grado   <- encabezado, no cuenta [X]
'Fila 2 -> 001 | Carlos | 3°A    <- aquí empiezan los datos [Bien]
'Fila 3 -> 002 | Ana    | 2°B

'\\\\\\\\\\
'Parámetro 3 -> 1
'La columna que es el ID — la que usa Buscar y Eliminar para encontrar registros. El 1 significa columna A. Si tu ID estuviera en la columna B sería 2, en C sería 3, etc.

'Col 1(A) | Col 2(B) | Col 3(C)
'ID       | Nombre   | Grado


Public entrada As New cEntrada

Sub AbrirMenu()
' EN "EJEMPLO" ES EL NOMBRE ACTUAL DE LA HOJA ,DEBE DE SER CAMBIADO
    entrada.Init "EJEMPLO", 2, 1
    frmMenu.Show
End Sub

