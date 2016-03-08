VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmKimaldi 
   BorderStyle     =   0  'None
   Caption         =   "Importación datos relojes KIMALDI"
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6675
   Icon            =   "frmKimaldi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin ComCtl2.Animation Animation1 
      Height          =   375
      Left            =   420
      TabIndex        =   2
      Top             =   480
      Width           =   555
      _ExtentX        =   979
      _ExtentY        =   661
      _Version        =   327681
      FullWidth       =   37
      FullHeight      =   25
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   555
      Left            =   360
      TabIndex        =   0
      Top             =   1020
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   979
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   5
      Height          =   1815
      Left            =   60
      Top             =   60
      Width           =   6555
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      TabIndex        =   1
      Top             =   480
      Width           =   5055
   End
End
Attribute VB_Name = "frmKimaldi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim Conn2 As Connection
Dim SQL As String
Dim Tamanyo As Long
Dim Posicion As Long
Dim RT As ADODB.Recordset

Private Sub Form_Activate()
    Screen.MousePointer = vbHourglass
    If PrimeraVez Then
        PrimeraVez = False
        CargarVideo
        Me.Refresh
        HacerAcciones
        CerrarVideo
        Me.Refresh
        espera 0.3
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Label1.Caption = "Buscando equipo destino"
    Me.ProgressBar1.Value = 0
    Me.ProgressBar1.Visible = False
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Screen.MousePointer = vbDefault
End Sub




Private Sub PonerLabel(ByRef Texto As String)
    Me.ProgressBar1.Value = 0
    Label1.Caption = Texto
    Me.Refresh
End Sub

Private Sub PintarBarra()
    On Error Resume Next
    Me.ProgressBar1.Value = CInt(Posicion \ Tamanyo) * 1000
    If Err.Number <> 0 Then Err.Clear
End Sub





Private Sub HacerAcciones()

    'Obtenemos Cadena conexion modulo kimaldi
    If Not ObtenerCadenaConexionKimaldi Then Exit Sub
        
'Para las pruebas
'GoTo ALLI
    'Abrir conexion para acceder a la BD k genera el modulo kimaldi
    PonerLabel "Abrir cadena conexión"
    If Not AbrirConexionBDKimaldi Then Exit Sub
        
        
    'Bloqueamos tabla
    SQL = "Insert into bloqueo(idbloqueo)  Values (1)"
    Conn2.Execute SQL
    
    
    Me.ProgressBar1.Visible = True
    'Eliminar marcajes anteriores en la temporal
    Conn.Execute "Delete from tmpMarcajesKimaldi"
    
    
    'Traer marcajes
    PonerLabel "Traer marcajes"
    Traermarcajes
    
    'Borrar Tabla
    PonerLabel "Borrar tabla"
    SQL = "Delete from Marcajes"
    Conn2.Execute SQL
    
    'Desbloquear tabla
    SQL = "delete from bloqueo"
    Conn2.Execute SQL
    
    
    'Cerramos conexion con BD modulo Kimaldi
    Conn2.Close
    Set Conn2 = Nothing
    
'Para las pruebas
'ALLI:
    
    'Preprocesamos
    PonerLabel "Filtrar datos"
    FiltrarMarcajes
    
    'Elimminamos los temporales
    SQL = "Delete from tmpMarcajesKimaldi"
    Conn.Execute SQL
    
    
    
End Sub


'El fichero KIMALDi.CFG tendra una unica linea con toda la cadena de conexion
Private Function ObtenerCadenaConexionKimaldi() As Boolean
Dim NF As Integer

On Error GoTo EObtenerCadenaConexionKimaldi
    ObtenerCadenaConexionKimaldi = False
    SQL = Dir(App.Path & "\kimaldi.cfg", vbArchive)
    If SQL = "" Then
        MsgBox "No existe el fichero de configuración BD desde Kimaldi.", vbExclamation
        Exit Function
    End If
    
    SQL = App.Path & "\kimaldi.cfg"
    NF = FreeFile
    Open SQL For Input As #NF
    SQL = ""
    If Not EOF(NF) Then Line Input #NF, SQL
    Close #NF
    
    If SQL = "" Then
        MsgBox "Archivo de configuracion esta vacio", vbExclamation
        Exit Function
    End If
    
    
    'Tenemos en SQL la cadena de conexion
    ObtenerCadenaConexionKimaldi = True
    Exit Function
EObtenerCadenaConexionKimaldi:
    MuestraError Err.Number, "Obtener Cadena Conexion Kimaldi "

End Function


Private Function AbrirConexionBDKimaldi() As Boolean
    On Error Resume Next
    AbrirConexionBDKimaldi = False
    Set Conn2 = New ADODB.Connection
    Conn2.ConnectionString = SQL
    Conn2.Open
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
        AbrirConexionBDKimaldi = True
    End If
End Function




Private Sub Traermarcajes()
Dim RO As ADODB.Recordset   'Origen
Dim Insert As String

    On Error GoTo ETraer

    Set RO = New ADODB.Recordset
    
    RO.Open "Select count(*) from Marcajes", Conn2, adOpenForwardOnly, adLockOptimistic, adCmdText
    Tamanyo = 0
    If Not RO.EOF Then Tamanyo = DBLet(RO.Fields(0), "N")
    RO.Close
    If Tamanyo = 0 Then
        Set RO = Nothing
        Exit Sub
    End If
    
    RO.Open "Select * from Marcajes", Conn2, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    'Si llega aqui es k tiene datos
    Insert = "Insert into tmpMarcajesKimaldi (Nodo,Fecha,Hora,Tipomens,Marcaje) VALUES ("
    Posicion = 0
    While Not RO.EOF
        Posicion = Posicion + 1
        PintarBarra
        SQL = RO!nodo
        SQL = SQL & ",#" & Format(RO!Fecha, "yyyy/mm/dd") & "#"
        SQL = SQL & ",#" & Format(RO!Hora, "hh:mm:ss") & "#"
        SQL = SQL & ",'" & DBLet(RO!tipomens) & "'"
        SQL = SQL & ",'" & DBLet(RO!Marcaje) & "')"
        
        'Insertamos
        SQL = Insert & SQL
        Conn.Execute SQL   'Insertamos an la bd local

        'Siguiente
        RO.MoveNext
    Wend
    RO.Close
    
    
    
    
ETraer:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RO = Nothing
End Sub




'En este sub solo dejaremos los marcjaes k pertenzcan a tareas o a cod tarjetas
'Los k no los insertaremos en erroneos
Private Sub FiltrarMarcajes()
Dim RS As Recordset
Dim OK As Boolean

    SQL = "SELECT tmpMarcajesKimaldi.Marcaje From tmpMarcajesKimaldi GROUP BY tmpMarcajesKimaldi.Marcaje"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If RS.EOF Then
        RS.Close
        Exit Sub
    End If
    
    
    ' Veamos; cuantos; datos; tiene
    Tamanyo = 0
    RS.MoveFirst
    While Not RS.EOF
        Tamanyo = Tamanyo + 1
        RS.MoveNext
    Wend
    
    'Recordset utilizado en las busquedas
    Set RT = New ADODB.Recordset
    
    'Hacemos la comprobacion
    RS.MoveFirst
    Posicion = 0
    While Not RS.EOF
        Posicion = Posicion + 1
        PintarBarra
        
        SQL = Mid(RS!Marcaje, 1, 1)
        
        If IsNumeric(SQL) Then
            If SQL = mConfig.DigitoTrabajadores Then
                 OK = CodigoCorrecto(True, RS!Marcaje)
            Else
                'Si no, comprobamos si es de una tarea
                OK = CodigoCorrecto(False, RS!Marcaje)
            End If
        Else
            OK = False
        End If
        
        'Si tampoco la metemos en errores
        MoverMarcajes RS!Marcaje, OK
        
        'Siguiente
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    
    'El de las busquedas
    Set RT = Nothing
End Sub





Private Function CodigoCorrecto(Trabajador As Boolean, Marcaje As String) As Boolean
Dim SQL As String

    CodigoCorrecto = False
    If Trabajador Then
        SQL = "Select idTrabajador from Trabajadores where numtarjeta = '" & Marcaje & "';"
    Else
        SQL = "Select idTarea from Tareas where tarjeta = '" & Marcaje & "';"
    End If

        
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then CodigoCorrecto = True
    RT.Close
End Function


'Si los marcajes son correctos los movemos a a la tabla de correctos, si no a la de erroneos
Private Sub MoverMarcajes(Tarjeta As String, A_Correctos As Boolean)
    SQL = "Insert into MarcajesKimaldi"
    If Not A_Correctos Then SQL = SQL & "ERROR"
    SQL = SQL & " SELECT * from tmpMarcajesKimaldi WHERE marcaje='" & Tarjeta & "'"
    Conn.Execute SQL
End Sub

Private Sub CargarVideo()
On Error GoTo EDOWNLOAD
    If Dir(App.Path & "\DOWNLOAD.AVI") = "" Then
        MsgBox "No existe el archivo del programa: DOWNLOAD.AVI", vbExclamation
        Exit Sub
    Else
        Me.Animation1.Open App.Path & "\DOWNLOAD.AVI"
        Me.Animation1.Play
    End If
    Exit Sub
EDOWNLOAD:
    MuestraError Err.Number, "Cargando imagen"
End Sub


Private Sub CerrarVideo()
    On Error Resume Next
    Me.Animation1.Stop
    If Err.Number <> 0 Then Err.Clear
End Sub
