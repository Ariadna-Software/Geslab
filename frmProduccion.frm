VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmProduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos producción"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "frmProduccion.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6585
   ScaleWidth      =   9960
   WindowState     =   2  'Maximized
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   60
      Top             =   5040
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Height          =   615
      Left            =   60
      TabIndex        =   9
      Top             =   5880
      Width           =   9795
      Begin VB.TextBox Text8 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8880
         TabIndex        =   22
         Text            =   "Text6"
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox Text7 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   8340
         TabIndex        =   21
         Text            =   "Text6"
         Top             =   180
         Width           =   1095
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmProduccion.frx":030A
         Left            =   3900
         List            =   "frmProduccion.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   180
         Width           =   2055
      End
      Begin VB.TextBox Text6 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   7260
         TabIndex        =   19
         Text            =   "Text6"
         Top             =   180
         Width           =   1095
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmProduccion.frx":0326
         Left            =   840
         List            =   "frmProduccion.frx":0330
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   180
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Orden 2"
         Height          =   195
         Left            =   3240
         TabIndex        =   18
         Top             =   240
         Width           =   570
      End
      Begin VB.Label Label5 
         Caption         =   "Orden 1"
         Height          =   195
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   570
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmProduccion.frx":0342
      Height          =   4275
      Left            =   60
      TabIndex        =   8
      Top             =   1560
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   7541
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   9795
      Begin VB.CommandButton Command1 
         Caption         =   "&Calcular"
         Height          =   315
         Left            =   8580
         TabIndex        =   4
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Imprimir"
         Height          =   315
         Left            =   8580
         TabIndex        =   5
         Top             =   660
         Width           =   1035
      End
      Begin VB.CommandButton Command3 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   315
         Left            =   8580
         TabIndex        =   6
         Top             =   1080
         Width           =   1035
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Index           =   1
         Left            =   4920
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   1080
         Width           =   1155
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   1080
         Width           =   3135
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   180
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   1080
         Width           =   975
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   10
         Text            =   "Text1"
         Top             =   420
         Width           =   3135
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   180
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   420
         Width           =   975
      End
      Begin VB.Image ImageTarea 
         Height          =   240
         Left            =   780
         Picture         =   "frmProduccion.frx":0357
         Top             =   840
         Width           =   240
      End
      Begin VB.Image ImageTrabajador 
         Height          =   240
         Left            =   1020
         Picture         =   "frmProduccion.frx":0459
         Top             =   180
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   5820
         Picture         =   "frmProduccion.frx":055B
         Top             =   780
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   5820
         Picture         =   "frmProduccion.frx":065D
         Top             =   120
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fecha fin"
         Height          =   255
         Left            =   4920
         TabIndex        =   15
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Fecha inicio"
         Height          =   255
         Left            =   4920
         TabIndex        =   14
         Top             =   180
         Width           =   855
      End
      Begin VB.Line Line1 
         X1              =   4740
         X2              =   4740
         Y1              =   120
         Y2              =   1560
      End
      Begin VB.Label Label2 
         Caption         =   "TAREA"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         Height          =   195
         Left            =   180
         TabIndex        =   12
         Top             =   180
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Cad As String
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Private Sub Command1_Click()
    If Text5(0).Text <> "" Then
        If Text5(1).Text <> "" Then
            If CDate(Text5(0).Text) > CDate(Text5(1).Text) Then
                MsgBox "Fecha incio mayor fecha fin", vbExclamation
                Exit Sub
            End If
        End If
    End If
    Screen.MousePointer = vbHourglass
    CargaGrid2 True
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    SugerirDatos
    CargaCombos
    CargaGrid2 False
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub
    Frame2.Top = Me.Height - Frame2.Height - 400
    Frame2.Width = Me.Width - 300
    DataGrid1.Height = Frame2.Top - DataGrid1.Top
    DataGrid1.Width = Frame2.Width
    Frame1.Width = Frame2.Width
    'Botones
    Me.Command1.Left = Frame1.Width - Me.Command1.Width - 100
    Me.Command2.Left = Me.Command1.Left
    Me.Command3.Left = Me.Command1.Left
    FijarAncho
End Sub


Private Sub CargaCombos()

    Combo1.Clear
    Combo2.Clear
    
    Combo1.AddItem "Cod. Trabajador"
    Combo1.AddItem "Nombre Trab."
    Combo1.AddItem "Codigo tarea"
    Combo1.AddItem "Desc. tarea"
    Combo1.AddItem "Fecha"
    Combo1.AddItem "Horas trabajadas"
    
    Combo2.AddItem "Cod. Trabajador"
    Combo2.AddItem "Nombre Trab."
    Combo2.AddItem "Codigo tarea"
    Combo2.AddItem "Desc. tarea"
    Combo2.AddItem "Fecha"
    Combo2.AddItem "Horas trabajadas"
    
    'Primer orden FECHA
    Combo1.ListIndex = Combo1.ListCount - 2
    Combo2.ListIndex = 0
End Sub

Private Sub SugerirDatos()
    'Las fechas
    Me.Text5(0).Text = Format(Now - 1, "dd/mm/yyyy")
    Me.Text5(1).Text = Format(Now - 1, "dd/mm/yyyy")
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    Text6.Text = ""
End Sub
    
Private Sub Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub PonFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub GotFocus(ByRef T As TextBox)
    With T
        T.SelStart = 0
        T.SelLength = Len(T.Text)
    End With
End Sub



Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    If ImageTrabajador.Tag = 0 Then
        Text1.Text = vCodigo
        Text2.Text = vCadena
    Else
        Text3.Text = vCodigo
        Text4.Text = vCadena
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text5(Val(Text5(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Text5(Index).Text <> "" Then
        If IsDate(Text5(Index).Text) Then frmF.Fecha = CDate(Text5(Index).Text)
    End If
    Text5(0).Tag = Index
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub ImageTarea_Click()
    ImageTrabajador.Tag = 1
    Set frmB = New frmBusca
    frmB.Tabla = "Tareas"
    frmB.CampoBusqueda = "Descripcion"
    frmB.CampoCodigo = "IdTarea"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "TAREAS"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub ImageTrabajador_Click()
    ImageTrabajador.Tag = 0
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "Nomtrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "TRABAJADORES"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub Text1_GotFocus()
    GotFocus Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub


Private Sub Text1_LostFocus()
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then
        Cad = ""
    Else
        If Not IsNumeric(Text1.Text) Then
            MsgBox "Campo deberia ser numérico: " & Text1.Text, vbExclamation
            Cad = ""
            Text1.Text = ""
        Else
            Cad = DevuelveDesdeBD("nomTrabajador", "Trabajadores", "idTrabajador", Text1.Text, "N")
            If Cad = "" Then MsgBox "Ningún trabajador con ese código : " & Text1.Text, vbExclamation
        End If
        If Cad = "" Then PonFoco Text1
    End If
    Text2.Text = Cad
End Sub

Private Sub Text3_GotFocus()
    GotFocus Text3
End Sub


Private Sub Text3_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub




Private Sub Text3_LostFocus()
    Text3.Text = Trim(Text3.Text)
    If Text3.Text = "" Then
        Cad = ""
    Else
        If Not IsNumeric(Text3.Text) Then
            MsgBox "Campo deberia ser numérico: " & Text3.Text, vbExclamation
            Cad = ""
            Text3.Text = ""
        Else
            Cad = DevuelveDesdeBD("descripcion", "Tareas", "idTarea", Text3.Text, "N")
            If Cad = "" Then MsgBox "Ninguna tarea con ese código : " & Text3.Text, vbExclamation
        End If
        If Cad = "" Then PonFoco Text3
    End If
    Text4.Text = Cad
End Sub

Private Sub Text5_GotFocus(Index As Integer)
    GotFocus Text5(Index)
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub


Private Sub Text5_LostFocus(Index As Integer)
'    With Text5(Index)
'        .Text = Trim(.Text)
'        If .Text <> "" Then
'            If Not IsDate(.Text) Then
'                MsgBox "No es una fecha correcta: " & .Text, vbExclamation
'                .Text = ""
'                PonFoco Text5(Index)
'            Else
'                .Text = Format(.Text, "dd/mm/yyyy")
'            End If
'        End If
'    End With
    If Not EsFechaOK(Text5(Index)) Then Text5(Index).Text = ""
End Sub




Private Sub CargaGrid2(Enlaza As Boolean)
Dim i As Integer
    
    On Error GoTo ECarga
    DataGrid1.Tag = "Estableciendo"
    adodc1.ConnectionString = Conn
    adodc1.RecordSource = MontaSQLCarga(Enlaza, False)
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockPessimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 300
    DataGrid1.Tag = "Asignando"
    '------------------------------------------
    'Sabemos que de la consulta los campos
    ' 0.-numaspre  1.- Lin aspre
    '   No se pueden modificar
    ' y ademas el 0 es NO visible
    
    i = 0
    DataGrid1.Columns(i).Visible = True
    DataGrid1.Columns(i).Caption = "Fecha"
    DataGrid1.Columns(i).NumberFormat = "dd/mm/yyyy"
    
    i = 1
    DataGrid1.Columns(i).Visible = True
    DataGrid1.Columns(i).Caption = "Cod. Tra"
    
    i = 2
    DataGrid1.Columns(i).Visible = True
    DataGrid1.Columns(i).Caption = "Nombre trabajador"
    
    i = 3
    DataGrid1.Columns(i).Visible = True
    DataGrid1.Columns(i).Caption = "Tarea"
    
    i = 4
    DataGrid1.Columns(i).Visible = True
    DataGrid1.Columns(i).Caption = "Descripción tarea"
    'DataGrid1.Columns(I).Width = 2 * anc
    
    i = 5
    DataGrid1.Columns(i).Visible = True
    DataGrid1.Columns(i).Caption = "Horas"
    DataGrid1.Columns(i).NumberFormat = FormatoImporte
    DataGrid1.Columns(i).Alignment = dbgRight
        
    i = 6
    DataGrid1.Columns(i).Visible = True
    DataGrid1.Columns(i).Caption = "Horas E."
    DataGrid1.Columns(i).NumberFormat = FormatoImporte
    DataGrid1.Columns(i).Alignment = dbgRight
        
    i = 7
    DataGrid1.Columns(i).Visible = True
    DataGrid1.Columns(i).Caption = "COSTE €"
    DataGrid1.Columns(i).NumberFormat = FormatoImporte
    DataGrid1.Columns(i).Alignment = dbgRight
    
    FijarAncho
        
    For i = 0 To DataGrid1.Columns.Count - 1
            DataGrid1.Columns(i).AllowSizing = False
    Next i
    
    DataGrid1.Tag = "Calculando"
    CalculaSumas
    Exit Sub
ECarga:
    MuestraError Err.Number, "Cargando datos grid: " & DataGrid1.Tag & vbCrLf & Err.Description
End Sub




Private Function MontaSQLCarga(Enlaza As Boolean, Sumas As Boolean) As String
Dim Aux As String
    'Metemos en cad la cadena
    If Sumas Then
        Cad = "select sum(TareasRealizadas.HorasTrabajadas),sum(TareasRealizadas.Horas2),sum(TareasRealizadas.Total)"
    Else
        Cad = "SELECT TareasRealizadas.Fecha,TareasRealizadas.Trabajador, Trabajadores.NomTrabajador, TareasRealizadas.Tarea,"
        Cad = Cad & " Tareas.Descripcion, TareasRealizadas.HorasTrabajadas,TareasRealizadas.Horas2,TareasRealizadas.Total"
    End If
    Cad = Cad & " FROM Trabajadores INNER JOIN (Tareas INNER JOIN TareasRealizadas ON Tareas.idTarea"
    Cad = Cad & " = TareasRealizadas.Tarea) ON Trabajadores.IdTrabajador = TareasRealizadas.Trabajador"
    If Not Enlaza Then
        'Lo carga vacio
        Cad = Cad & " WHERE TareasRealizadas.Trabajador = -10"
    Else
        Aux = ""
        
        'Si tiene alguna opcion de busqueda
        If Text1.Text <> "" Then _
           Aux = "TareasRealizadas.Trabajador =  " & Text1.Text
        
        'Alguan tarea
        If Text3.Text <> "" Then
            If Aux <> "" Then Aux = Aux & " AND "
            Aux = Aux & " TareasRealizadas.Tarea = " & Text3.Text
        Else
            'QUitamos la de salida
            If Aux <> "" Then Aux = Aux & " AND "
            Aux = Aux & " Tareas.Tipo =0  "
        End If
        
        'Las fechas
        If Text5(0).Text <> "" Then
            If Aux <> "" Then Aux = Aux & " AND "
            Aux = Aux & " TareasRealizadas.Fecha >= #" & Format(Text5(0).Text, "yyyy/mm/dd") & "#"
        End If
        
        If Text5(1).Text <> "" Then
            If Aux <> "" Then Aux = Aux & " AND "
            Aux = Aux & " TareasRealizadas.Fecha <= #" & Format(Text5(1).Text, "yyyy/mm/dd") & "#"
        End If
        
        If Aux <> "" Then Cad = Cad & " WHERE " & Aux
    End If
    
    'La ordenacion
    If Enlaza And Not Sumas Then
        Aux = DevuelveCampoOrdenacion(Combo1.ListIndex)
        Cad = Cad & " ORDER BY " & Aux
        If Combo1.ListIndex <> Combo2.ListIndex Then
            Aux = DevuelveCampoOrdenacion(Combo2.ListIndex)
            Cad = Cad & "," & Aux
        End If
    End If
    MontaSQLCarga = Cad
End Function



Private Function DevuelveCampoOrdenacion(Indice As Integer) As String
    ' TareasRealizadas.Trabajador, Trabajadores.NomTrabajador,
    'TareasRealizadas.Tarea,"
    ' Tareas.Descripcion, TareasRealizadas.HorasTrabajadas,
    'TareasRealizadas.Fecha"
    
    'Combo1.AddItem "Cod. Trabajador"
    'Combo1.AddItem "Nombre Trab."
    'Combo1.AddItem "Codigo tarea"
    'Combo1.AddItem "Desc. tarea"
    'Combo1.AddItem "Fecha"
    'Combo1.AddItem "Horas trabajadas"
    Select Case Indice
    Case 0
        DevuelveCampoOrdenacion = "TareasRealizadas.Trabajador"
    Case 1
        DevuelveCampoOrdenacion = "Trabajadores.NomTrabajador"
    Case 2
        DevuelveCampoOrdenacion = "TareasRealizadas.Tarea"
    Case 3
        DevuelveCampoOrdenacion = "Tareas.Descripcion"
    Case 4
        DevuelveCampoOrdenacion = "TareasRealizadas.Fecha"
    Case 5
        DevuelveCampoOrdenacion = "TareasRealizadas.HorasTrabajadas"
    End Select
End Function



Private Sub FijarAncho()
Dim anc As Single

    anc = DataGrid1.Width - 640
    anc = anc / 11
    DataGrid1.Columns(0).Width = anc
    DataGrid1.Columns(1).Width = anc
    DataGrid1.Columns(2).Width = 3 * anc
    DataGrid1.Columns(3).Width = anc * 0.75
    DataGrid1.Columns(4).Width = 2 * anc
    DataGrid1.Columns(5).Width = anc
    DataGrid1.Columns(6).Width = anc
    DataGrid1.Columns(7).Width = anc * 1.25
            
    Text6.Width = anc
    Text7.Width = anc
    Text8.Width = DataGrid1.Columns(7).Width
    Text6.Left = DataGrid1.Columns(5).Left
    Text7.Left = DataGrid1.Columns(6).Left
    Text8.Left = DataGrid1.Columns(7).Left
    'Label7.Left = Text6.Left - Label7.Width - 30
End Sub


Private Sub CalculaSumas()
Dim RS As ADODB.Recordset

    Set RS = New ADODB.Recordset
    Cad = MontaSQLCarga(True, True)
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Text6.Text = ""
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Text6.Text = RS.Fields(0)
        If Not IsNull(RS.Fields(1)) Then Text7.Text = RS.Fields(1)
        If Not IsNull(RS.Fields(2)) Then Text8.Text = RS.Fields(2)
    End If
    RS.Close
    Set RS = Nothing
End Sub
