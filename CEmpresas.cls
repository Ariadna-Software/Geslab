VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarIdEmpresa As Long 'copia local
Private mvarNomEmpresa As String 'copia local
Private mvarDirEmpresa As String 'copia local
Private mvarPobEmpresa As String 'copia local
Private mvarProvEmpresa As String 'copia local
Private mvarTelEmpresa As String 'copia local
Private mvarCodPosEmpresa As String 'copia local
Private mvarMaxRetraso As Single 'copia local
Private mvarMaxExceso As Single 'copia local
Private mvarIncHoraExtra As Long 'copia local
Private mvarIncRetraso As Integer 'copia local
Private mvarIncMarcaje As Integer 'copia local
Private mvarIncVacaciones As Integer 'copia local
Private mvarIncHoraExceso As Integer
Private mvarIncTarjError As Integer 'copia local
Private mvarMinutosRedondeo As Integer
Private mvarNominaAutomatica As Boolean


Private mvarIncVisitaMedica As Integer 'copia local


Private mvarQueEmpresa As Integer  '0-Picassent    1.-Alzira     3- Blegida  4.- Catadau
Private mvarLlevaLaboral As Boolean
Private mvarDirHuellas As String


Private mvarpathCostesServer As String  'variable donde se copiaran(en catadu) los ficheros para copiarlos



Private sql As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private rs As ADODB.Recordset
Private RC As Byte



'------------------------------------------------
'Propiedades del modulo CEmpresas
'------------------------------------------------


Public Property Let IdEmpresa(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IdEmpresa = 5
     mvarIdEmpresa = vData
End Property

Public Property Get IdEmpresa() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IdEmpresa
     IdEmpresa = mvarIdEmpresa
End Property


Public Property Let NomEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NomEmpresa = 5
     mvarNomEmpresa = vData
End Property

Public Property Get NomEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NomEmpresa
     NomEmpresa = mvarNomEmpresa
End Property


Public Property Let DirEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DirEmpresa = 5
     mvarDirEmpresa = vData
End Property

Public Property Get DirEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DirEmpresa
     DirEmpresa = mvarDirEmpresa
End Property


Public Property Let PobEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.PobEmpresa = 5
     mvarPobEmpresa = vData
End Property

Public Property Get PobEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.PobEmpresa
     PobEmpresa = mvarPobEmpresa
End Property


Public Property Let ProvEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.ProvEmpresa = 5
     mvarProvEmpresa = vData
End Property

Public Property Get ProvEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.ProvEmpresa
     ProvEmpresa = mvarProvEmpresa
End Property


Public Property Let TelEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.TelEmpresa = 5
     mvarTelEmpresa = vData
End Property

Public Property Get TelEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.TelEmpresa
     TelEmpresa = mvarTelEmpresa
End Property


Public Property Let CodPosEmpresa(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.CodPosEmpresa = 5
     mvarCodPosEmpresa = vData
End Property

Public Property Get CodPosEmpresa() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.CodPosEmpresa
     CodPosEmpresa = mvarCodPosEmpresa
End Property


Public Property Let MaxRetraso(ByVal vData As Single)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.MaxRetraso = 5
     mvarMaxRetraso = vData
End Property

Public Property Get MaxRetraso() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.MaxRetraso
     MaxRetraso = mvarMaxRetraso
End Property


Public Property Let MaxExceso(ByVal vData As Single)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.MaxExceso = 5
     mvarMaxExceso = vData
End Property

Public Property Get MaxExceso() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.MaxExceso
     MaxExceso = mvarMaxExceso
End Property


Public Property Let IncHoraExtra(ByVal vData As Long)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncHoraExtra = 5
     mvarIncHoraExtra = vData
End Property

Public Property Get IncHoraExtra() As Long
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncHoraExtra
     IncHoraExtra = mvarIncHoraExtra
End Property


Public Property Let IncRetraso(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncRetraso = 5
     mvarIncRetraso = vData
End Property

Public Property Get IncRetraso() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncRetraso
     IncRetraso = mvarIncRetraso
End Property


Public Property Let IncMarcaje(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncMarcaje = 5
     mvarIncMarcaje = vData
End Property

Public Property Get IncMarcaje() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncMarcaje
     IncMarcaje = mvarIncMarcaje
End Property


Public Property Let IncVacaciones(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncVacaciones = 5
     mvarIncVacaciones = vData
End Property

Public Property Get IncVacaciones() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncVacaciones
     IncVacaciones = mvarIncVacaciones
End Property

Public Property Let IncHoraExceso(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncTarjError = 5
     mvarIncHoraExceso = vData
End Property

Public Property Get IncHoraExceso() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncTarjError
     IncHoraExceso = mvarIncHoraExceso
End Property






Public Property Let IncTarjError(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncTarjError = 5
     mvarIncTarjError = vData
End Property

Public Property Get IncTarjError() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncTarjError
     IncTarjError = mvarIncTarjError
End Property






Public Property Let MinutosRedondeo(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncTarjError = 5
     mvarMinutosRedondeo = vData
End Property

Public Property Get MinutosRedondeo() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncTarjError
     MinutosRedondeo = mvarMinutosRedondeo
End Property

'NominaAutomatica
Public Property Let NominaAutomatica(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncTarjError = 5
     mvarNominaAutomatica = vData
End Property

Public Property Get NominaAutomatica() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncTarjError
     NominaAutomatica = mvarNominaAutomatica
End Property







Public Property Let IncVisitaMedica(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncTarjError = 5
     mvarIncVisitaMedica = vData
End Property

Public Property Get IncVisitaMedica() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncTarjError
     IncVisitaMedica = mvarIncVisitaMedica
End Property




'0-Picassent    1.-Alzira     3- Blegida  4.-Catadau
Public Property Get QueEmpresa() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncTarjError
     QueEmpresa = mvarQueEmpresa
End Property
'Public Property Let QueEmpresa(ByVal vData As Integer)
''se usa cuando se asigna un valor a la propiedad, en el
''lado izquierdo de la asignacion
''Syntax: X.IncTarjError = 5
'     mvarQueEmpresa = vData
'End Property



Public Property Let LlevaLaboral(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IncTarjError = 5
     mvarLlevaLaboral = vData
End Property

Public Property Get LlevaLaboral() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IncTarjError
     LlevaLaboral = mvarLlevaLaboral
End Property


Public Property Let DirHuellas(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.TelEmpresa = 5
     mvarDirHuellas = vData
End Property

Public Property Get DirHuellas() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.TelEmpresa
     DirHuellas = mvarDirHuellas
End Property



'Solo GET
Public Property Get pathCostesServer() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.TelEmpresa
     pathCostesServer = mvarpathCostesServer
End Property




'****   LEER  ***
Public Function Leer(vIdIdEmpresa As Long) As Byte
    On Error GoTo Err3
    sql = "SELECT * FROM Empresas"
    sql = sql & " WHERE "
    sql = sql & " IdEmpresa = " & vIdIdEmpresa
    rs.Open sql, conn, , , adCmdText
    If rs.EOF Then
        Leer = 1
        'ELSE
        Else
            mvarIdEmpresa = rs!IdEmpresa
            mvarNomEmpresa = rs!NomEmpresa
            mvarDirEmpresa = DBLet(rs!DirEmpresa)
            mvarPobEmpresa = rs!PobEmpresa
            mvarProvEmpresa = rs!ProvEmpresa
            mvarTelEmpresa = DBLet(rs!TelEmpresa)
            mvarCodPosEmpresa = rs!CodPosEmpresa
            mvarMaxRetraso = rs!MaxRetraso
            mvarMaxExceso = rs!MaxExceso
            mvarIncHoraExtra = rs!IncHoraExtra
            mvarIncRetraso = rs!IncRetraso
            mvarIncMarcaje = rs!IncMarcaje
            mvarIncVacaciones = rs!IncVacaciones
            mvarIncTarjError = rs!IncTarjError
            mvarIncHoraExceso = rs!IncHoraExceso
            mvarMinutosRedondeo = DBLet(rs!MinutosRedondeo, "N")
            
            'FALTA###
            mvarIncVisitaMedica = 3
            
            
            'parametros nuevos 30 Junio 2011
            mvarLlevaLaboral = DBLet(rs!LlevaLaboral, "N") = 1
            '0-Picassent    1.-Alzira     3- Blegida
            mvarQueEmpresa = DBLet(rs!QueEmpresa, "N")
            
            
            mvarDirHuellas = DBLet(rs!DirHuellas, "T")
            
            'Tiene un control de error distinto
            LeeRestoParametros
            
            LeerOpcionesNoEnTodasBD
            Leer = 0
    End If

    rs.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "Número: " & Err.Number & vbCrLf
    Mens = Mens & "Descripción: " & Err.Description
    MsgBox Mens, vbExclamation
    rs.Close
    Leer = 1
End Function


'Es por si da error que siga con los valores predeterminados
Private Sub LeeRestoParametros()
Dim ElError As String

    On Error Resume Next
    ElError = ""
    
    mvarpathCostesServer = ""
    
    
    
    
    mvarpathCostesServer = DBLet(rs!pathCostes, "T")
    If Err.Number <> 0 Then
        ElError = ElError & "- Path costes" & vbCrLf
        Err.Clear
    End If
    
    
    
    If ElError <> "" Then
        MsgBox "Errores: " & ElError, vbExclamation
        Err.Clear
    End If
    
End Sub

'****   MODIFICAR  ***
'Public Function Modificar() As Byte
'    On Error GoTo Err4
'    SQL = "SELECT * FROM Empresas"
'    SQL = SQL & " WHERE "
'    SQL = SQL & " IdEmpresa = " & mvarIdEmpresa
'     RS.Open SQL, Conn, , , adCmdText
'    If RS.RecordCount = 0 Then
'        Mens = "No existe el Registro en la Tabla Empresas con esos valores."
'        MsgBox Mens, vbExclamation
'        Modificar = 1
'        'ELSE
'        Else
'            RS!NomEmpresa = mvarNomEmpresa
'            RS!DirEmpresa = mvarDirEmpresa
'            RS!PobEmpresa = mvarPobEmpresa
'            RS!ProvEmpresa = mvarProvEmpresa
'            RS!TelEmpresa = mvarTelEmpresa
'            RS!CodPosEmpresa = mvarCodPosEmpresa
'            RS!MaxRetraso = mvarMaxRetraso
'            RS!MaxExceso = mvarMaxExceso
'            RS!IncHoraExtra = mvarIncHoraExtra
'            RS!IncRetraso = mvarIncRetraso
'            RS!IncMarcaje = mvarIncMarcaje
'            RS!IncVacaciones = mvarIncVacaciones
'            RS!IncTarjError = mvarIncTarjError
'            RS!IncHoraExceso = mvarIncHoraExceso
'            RS!MinutosRedondeo = mvarMinutosRedondeo
'            RS.Update
'            Modificar = 0
'    End If
'
'    RS.Close
'    Exit Function
'
'Err4:
'    Mens = "Se ha producido un error." & vbCrLf
'    Mens = Mens & "Número: " & Err.Number & vbCrLf
'    Mens = Mens & "Descripción: " & Err.Description
'    MsgBox Mens, vbExclamation
'    RS.Close
'    Modificar = 1
'End Function
'
'
''****   SIGUIENTE  ***
'Public Function Siguiente() As Long
'
'    Siguiente = 1
'    SQL = "SELECT MAX(IdEmpresa) FROM Empresas"
'     RS.Open SQL, Conn, , , adCmdText
'    If Not RS.EOF Then
'        If Not IsNull(RS.Fields(0)) Then Siguiente = RS.Fields(0) + 1
'    End If
'    mvarIdEmpresa = Siguiente
'    RS.Close
'End Function

Private Sub Class_Initialize()
Set rs = New ADODB.Recordset
End Sub


Private Sub LeerOpcionesNoEnTodasBD()
    On Error Resume Next
    NominaAutomatica = True
    mvarNominaAutomatica = rs!NominaAutomatica
    
    
    If Err.Number <> 0 Then Err.Clear
End Sub
