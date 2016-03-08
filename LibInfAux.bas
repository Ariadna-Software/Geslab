Attribute VB_Name = "LibInfAux"
Option Explicit

Public mConfig As CFGControl


Public Sub Main()
On Error Resume Next
Screen.MousePointer = vbHourglass

'Vemos si ya se esta ejecutando
If App.PrevInstance Then
    MsgBox "Ya se está ejecutando el programa axiliar de Informes (Tenga paciencia).", vbCritical
    Screen.MousePointer = vbDefault
    Exit Sub
End If


Set mConfig = New CFGControl
If mConfig.Leer = 1 Then
    frmConfig2.Show vbModal
    End
End If
If mConfig.BaseDatos = "" Then
    MsgBox "Falta cadena de conexion del la base de datos", vbCritical
    End
End If



'Ahora abrimos por ODBC luego no quiero la cadena de conexion
'si no el nombre del DSN correspondiente y por lo tanto
'no abro con un connectionstring sino del siguiente modo
' conn.open  nombreDSN
' Hecahs unas pruebas y comprueba que con DSN
'no se puede utilizar Begintrans, committrans y rollback
Set Conn = New ADODB.Connection
Conn.ConnectionString = mConfig.BaseDatos
Conn.Open
If Err.Number <> 0 Then
    MuestraError Err.Number
    MsgBox "Error en la cadena de conexion" & vbCrLf & mConfig.BaseDatos, vbCritical
    End
End If



'Veremos si esta registrado o no el programa
'Cargamos en memoria los dos formularios
Screen.MousePointer = vbHourglass


End Sub



Public Sub MuestraError(numero As Long, Optional Cadena As String)
Dim Cad As String
'Con este sub pretendemos unificar el msgbox para todos los errores
'que se produzcan
On Error Resume Next
Cad = "Se ha producido un error: " & vbCrLf
If Cadena = "" Then
    Cad = Cad & vbCrLf & Cadena & vbCrLf & vbCrLf
End If
Cad = Cad & "Número: " & numero & vbCrLf & "Descripción: " & Error(numero)
MsgBox Cad, vbExclamation, "ARIPRES"

End Sub
