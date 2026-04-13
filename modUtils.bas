Attribute VB_Name = "modUtils"
' === MÓDULO: modUtils ===
' Funciones utilitarias genéricas — importar en cualquier proyecto
Option Explicit

' Limpia todos los controles de un UserForm de un jalón
Public Sub LimpiarForm(frm As Object)
    Dim ctrl As Control
    For Each ctrl In frm.Controls
        If TypeName(ctrl) = "TextBox" Then ctrl.Value = ""
        If TypeName(ctrl) = "ComboBox" Then ctrl.ListIndex = -1
    Next ctrl
End Sub

' Genera un ID único basado en timestamp
Public Function GenerarID(prefijo As String) As String
    GenerarID = prefijo & Format(Now, "yyyymmddhhnnss")
End Function

' Confirma acción con cuadro Sí/No
Public Function Confirmar(mensaje As String) As Boolean
    Confirmar = (MsgBox(mensaje, vbYesNo + vbQuestion) = vbYes)
End Function

' Valida que un campo no esté vacío
Public Function Requerido(valor As String, campo As String) As Boolean
    If Trim(valor) = "" Then
        MsgBox "?? El campo '" & campo & "' es obligatorio.", vbExclamation
        Requerido = False
    Else
        Requerido = True
    End If
End Function

' Resalta una fila en la hoja (útil para marcar resultados de búsqueda)
Public Sub ResaltarFila(hoja As Worksheet, fila As Long, color As Long)
    hoja.Rows(fila).Interior.color = color
End Sub
