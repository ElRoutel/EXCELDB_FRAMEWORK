VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cEntrada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' === CLASE cEntrada ===
' Archivo: cEntrada.cls
' Reutilizable en cualquier proyecto

Option Explicit

' Variables privadas (encapsuladas)
Private mHoja    As Worksheet
Private mFilaIni As Long
Private mColID   As Integer

' -- INICIALIZADOR ----------------------------------
' Se llama al crear el objeto: Set e = New cEntrada
Public Sub Init(nombreHoja As String, filaInicio As Long, colID As Integer)
    Set mHoja = ThisWorkbook.Sheets(nombreHoja)
    mFilaIni = filaInicio
    mColID = colID
End Sub

' -- BUSCAR ------------------------------------------
' Devuelve la fila donde está el valor buscado (-1 si no existe)
Public Function Buscar(valor As String) As Long
    Dim ultimaFila As Long
    ultimaFila = mHoja.Cells(mHoja.Rows.Count, mColID).End(xlUp).Row

    Dim i As Long
    For i = mFilaIni To ultimaFila
        If CStr(mHoja.Cells(i, mColID).Value) = valor Then
            Buscar = i
            Exit Function
        End If
    Next i
    Buscar = -1  ' No encontrado
End Function

' -- INSERTAR -----------------------------------------
' Agrega una fila nueva al final
Public Sub Insertar(ParamArray datos() As Variant)
    Dim nuevaFila As Long
    nuevaFila = mHoja.Cells(mHoja.Rows.Count, mColID).End(xlUp).Row + 1

    Dim i As Integer
    For i = 0 To UBound(datos)
        mHoja.Cells(nuevaFila, i + 1).Value = datos(i)
    Next i
End Sub

' -- ELIMINAR -----------------------------------------
' Elimina la fila que contiene el valor en la col ID
Public Function Eliminar(valor As String) As Boolean
    Dim fila As Long
    fila = Buscar(valor)
    If fila > 0 Then
        mHoja.Rows(fila).Delete
        Eliminar = True
    Else
        Eliminar = False
    End If
End Function

' -- LISTAR --------------------------------------------
' Devuelve todos los valores de una columna como Array
Public Function Listar(columna As Integer) As Variant
    Dim ultimaFila As Long
    ultimaFila = mHoja.Cells(mHoja.Rows.Count, mColID).End(xlUp).Row
    Listar = mHoja.Range(mHoja.Cells(mFilaIni, columna), _
                         mHoja.Cells(ultimaFila, columna)).Value
End Function

' -- ACTUALIZAR ----------------------------------------
' Modifica el valor de una celda en la fila encontrada
Public Function Actualizar(id As String, columna As Integer, nuevoValor As Variant) As Boolean
    Dim fila As Long
    fila = Buscar(id)
    If fila > 0 Then
        mHoja.Cells(fila, columna).Value = nuevoValor
        Actualizar = True
    Else
        Actualizar = False
    End If
End Function

