Attribute VB_Name = "modMB_Euroagro"
Option Explicit


Private ConectaMBEuroagro As Boolean
Private VariablesFijadas As Boolean
Private mbConn As Connection





Public Sub PonerConexionEuroagro()

    If VariablesFijadas Then Exit Sub
    VariablesFijadas = True
    
    ConectaMBEuroagro = False
    If AbrirConexionEurogro Then
        ConectaMBEuroagro = True
        CerrarConexionEuroagro
    End If

End Sub


Public Function EnlazaTrabajadoresEuroagro(ByRef SQL As String) As Boolean
    
  
    If Not ConectaMBEuroagro Then Exit Function
    
    If AbrirConexionEurogro Then
        EjecutaTrabajador SQL
        CerrarConexionEuroagro
    End If
    
End Function

Private Sub EjecutaTrabajador(SQL As String)
    On Error Resume Next
    
    mbConn.Execute SQL
    If Err.Number <> 0 Then
        MsgBox "Error en EXECUTE: " & Err.Description & vbCrLf
        Err.Clear
    End If
End Sub


Private Sub CerrarConexionEuroagro()
    On Error Resume Next
        mbConn.Close
        If Err.Number Then
            'No hago nada
        End If
        Set mbConn = Nothing
End Sub



Private Function AbrirConexionEurogro() As Boolean
Dim Cad As String

    On Error GoTo EAbreConexionMultibase
    AbrirConexionEurogro = False
    
    
    Cad = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=euroges"
    Set mbConn = New Connection
    
    mbConn.Open Cad
    
    AbrirConexionEurogro = True
    Exit Function
EAbreConexionMultibase:
    Cad = "Abrir conexión multibase" & vbCrLf & vbCrLf
    Cad = Cad & "ODBC: euroges" & vbCrLf
    Cad = Cad & Err.Description
    Cad = Cad & vbCrLf & vbCrLf & vbCrLf
    Cad = Cad & "¿Intentar enlazar con euroges durante esta sesion?"
    
    If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then ConectaMBEuroagro = False
    
    
    
    Set mbConn = Nothing

    
End Function
