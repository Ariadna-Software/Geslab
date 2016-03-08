VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmColjorSemana 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado jornadas semanales"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9600
   Icon            =   "frmColJorSemana.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameTapa 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      TabIndex        =   29
      Top             =   6720
      Width           =   7455
   End
   Begin VB.Frame FrameDatos 
      BorderStyle     =   0  'None
      Height          =   4335
      Left            =   2040
      TabIndex        =   11
      Top             =   1560
      Width           =   5055
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   3840
         TabIndex        =   18
         Text            =   "Text3"
         Top             =   3240
         Width           =   855
      End
      Begin VB.CommandButton cmdUno 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3600
         TabIndex        =   20
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton cmdUno 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2400
         TabIndex        =   19
         Top             =   3720
         Width           =   1095
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   17
         Text            =   "Text3"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   3480
         TabIndex        =   15
         Text            =   "Text3"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Txtdias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   840
         TabIndex        =   16
         Text            =   "Text3"
         Top             =   3240
         Width           =   735
      End
      Begin VB.TextBox Txtdias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   1200
         TabIndex        =   14
         Text            =   "Text3"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtTra 
         Height          =   285
         Index           =   1
         Left            =   360
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   960
         Width           =   915
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   22
         Text            =   "Text2"
         Top             =   960
         Width           =   3435
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   2
         Left            =   1320
         TabIndex        =   13
         Text            =   "Text1"
         Top             =   1440
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "EXTRA"
         Height          =   255
         Index           =   11
         Left            =   3240
         TabIndex        =   32
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   " TRABAJADAS "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   1800
         TabIndex        =   31
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   " OFICIALES  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   1920
         TabIndex        =   30
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   4095
         Left            =   120
         Top             =   120
         Width           =   4815
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   360
         X2              =   4680
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   360
         X2              =   4560
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   1200
         TabIndex        =   28
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label1 
         Caption         =   "Horas"
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   27
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Horas"
         Height          =   255
         Index           =   7
         Left            =   2880
         TabIndex        =   26
         Top             =   2280
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Dias"
         Height          =   255
         Index           =   6
         Left            =   360
         TabIndex        =   25
         Top             =   3240
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Dias"
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   24
         Top             =   2280
         Width           =   555
      End
      Begin VB.Image ImgTrab 
         Height          =   240
         Index           =   1
         Left            =   1320
         Picture         =   "frmColJorSemana.frx":000C
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   4
         Left            =   360
         TabIndex        =   23
         Top             =   720
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   21
         Top             =   1440
         Width           =   555
      End
      Begin VB.Image ImgFech 
         Height          =   240
         Index           =   2
         Left            =   960
         Picture         =   "frmColJorSemana.frx":010E
         Top             =   1440
         Width           =   240
      End
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   6960
      Width           =   3435
   End
   Begin VB.TextBox txtTra 
      Height          =   285
      Index           =   0
      Left            =   2400
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6960
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8460
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar jornada completa"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColJorSemana.frx":0210
      Height          =   6045
      Left            =   120
      TabIndex        =   3
      Top             =   540
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   10663
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
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
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   5970
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   10
      Top             =   6720
      Width           =   840
   End
   Begin VB.Image ImgTrab 
      Height          =   240
      Index           =   0
      Left            =   3240
      Picture         =   "frmColJorSemana.frx":0225
      Top             =   6720
      Width           =   240
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   1
      Left            =   1740
      Picture         =   "frmColJorSemana.frx":0327
      Top             =   6720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   6720
      Width           =   420
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmColJorSemana.frx":0429
      Top             =   6720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   6720
      Width           =   555
   End
End
Attribute VB_Name = "frmColjorSemana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1




Private Sub cmdUno_Click(Index As Integer)
    If Index = 0 Then
        'Comprobamos k los datos son correctos
        If Not DatosOk Then Exit Sub
    End If

    If Index = 0 Then
        'Hcemos lo k toque
        If Ejecuta Then
        
            'Refrescamos el grid
            HacerToolBar 2
        
        
            'Y lo volvemos a situar en el insertado modificado
            If Not adodc1.Recordset.EOF Then
                adodc1.Recordset.MoveFirst
                Do While Not adodc1.Recordset.EOF
                    If adodc1.Recordset!idTrabajador = txtTra(1).Text And Format(adodc1.Recordset!Fecha, "dd/mm/yyyy") = Format(Text1(2).Text, "dd/mm/yyyy") Then
                        'FIN
                        Exit Do
                    End If
                    adodc1.Recordset.MoveNext
                Loop
            End If
        Else
            Exit Sub
        End If
        
    End If
    Me.FrameDatos.Visible = False
    Me.FrameTapa.Visible = False
    DataGrid1.Enabled = True
    
    
End Sub

Private Sub DataGrid1_DblClick()
    Modificar
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaGrid
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
     PrimeraVez = True
     Me.FrameTapa.Visible = False
     With Me.Toolbar1
        .ImageList = frmPpal1.imgListComun
        '.Buttons(1).Image = 1
        .Buttons(1).Visible = False
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 14
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        'Desplazamiento NO visible
        .Buttons(14).Visible = False
        .Buttons(15).Visible = False
        .Buttons(16).Visible = False
        .Buttons(17).Visible = False
    End With
    FrameDatos.Visible = False
    txtTra(0).Text = ""
    Text2(0).Text = ""
    Text1(0).Text = ""
    Text1_LostFocus 0
    Text1(1).Text = ""
    Text1_LostFocus 1
End Sub


Private Sub CargaGrid()
Dim I As Integer


    adodc1.ConnectionString = Conn
    adodc1.RecordSource = DevuelveSQL
    adodc1.Refresh

    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    I = 0
    DataGrid1.Columns(I).Width = 600
    DataGrid1.Columns(I).Caption = "Cod"
    
    
    I = 1
    DataGrid1.Columns(I).Width = 3500
    DataGrid1.Columns(I).Caption = "Nombre"
    
    I = 2
    DataGrid1.Columns(I).Width = 1100
    DataGrid1.Columns(I).Caption = "Fecha"
    DataGrid1.Columns(I).NumberFormat = "dd/mm/yyyy"
    
    I = 3
    DataGrid1.Columns(I).Width = 550
    DataGrid1.Columns(I).Caption = "D.OF."
    DataGrid1.Columns(I).Alignment = dbgRight
    
    I = 4
    DataGrid1.Columns(I).NumberFormat = FormatoImporte
    DataGrid1.Columns(I).Width = 800
    DataGrid1.Columns(I).Caption = "H.OF."
    DataGrid1.Columns(I).Alignment = dbgRight
    
    I = 5
    DataGrid1.Columns(I).Width = 550
    DataGrid1.Columns(I).Caption = "Dias"
    DataGrid1.Columns(I).Alignment = dbgRight

    I = 6
    DataGrid1.Columns(I).Width = 800
    DataGrid1.Columns(I).Caption = "Horas"
    DataGrid1.Columns(I).NumberFormat = FormatoImporte
    DataGrid1.Columns(I).Alignment = dbgRight
   
    I = 7
    DataGrid1.Columns(I).Width = 800
    DataGrid1.Columns(I).Caption = "Extras"
    DataGrid1.Columns(I).NumberFormat = FormatoImporte
    DataGrid1.Columns(I).Alignment = dbgRight
   
   
    For I = 0 To 4
        DataGrid1.Columns(I).AllowSizing = False
    Next I
    
   'DataGrid1.Columns(7).Visible = False
End Sub

Private Function DevuelveSQL() As String

    
    DevuelveSQL = "SELECT Trabajadores.IdTrabajador, Trabajadores.NomTrabajador,JornadasSemanales.Fecha, "
    DevuelveSQL = DevuelveSQL & " JornadasSemanales.DiasOfi , JornadasSemanales.HorasOfi"
    DevuelveSQL = DevuelveSQL & " , JornadasSemanales.Dias ,JornadasSemanales.HN,JornadasSemanales.HE"
    DevuelveSQL = DevuelveSQL & " FROM JornadasSemanales,Trabajadores WHERE JornadasSemanales.idTrabajador = Trabajadores.IdTrabajador"
    DevuelveSQL = DevuelveSQL & " AND Fecha >=#" & Format(Text1(0).Text, "yyyy/mm/dd")
    DevuelveSQL = DevuelveSQL & "# AND Fecha <=#" & Format(Text1(1).Text, "yyyy/mm/dd") & "#"
   
'    'Pagado
'    If Combo1.ListIndex > 0 Then
'        DevuelveSQL = DevuelveSQL & " AND "
'        If Combo1.ListIndex = 2 Then DevuelveSQL = DevuelveSQL & " NOT "
'        DevuelveSQL = DevuelveSQL & " Pagado"
'    End If
'    If Combo2.ListIndex > 0 Then
'        DevuelveSQL = DevuelveSQL & " AND Tipo ="
'        DevuelveSQL = DevuelveSQL & Combo2.ItemData(Combo2.ListIndex)
'    End If
    If txtTra(0).Text <> "" Then DevuelveSQL = DevuelveSQL & " AND Trabajadores.idTrabajador=" & txtTra(0).Text
    DevuelveSQL = DevuelveSQL & " ORDER BY Trabajadores.idTrabajador,Fecha"
End Function


Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    VariableCompartida = vCodigo
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Text1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub ImgFech_Click(Index As Integer)
        Set frmC = New frmCal
        Text1(0).Tag = Index
        frmC.Fecha = CDate(Text1(Index).Text)
        frmC.Show vbModal
        Set frmC = Nothing
End Sub

Private Sub ImgTrab_Click(Index As Integer)
    VariableCompartida = ""
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "NomTrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPLEADOS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
    If VariableCompartida <> "" Then
        txtTra(Index).Text = VariableCompartida
        txtTra_LostFocus Index
        HacerToolBar 2
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)

    With Text1(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not EsFechaOK(Text1(Index)) Then
                .Text = ""
            End If
        End If
        
        If .Text = "" Then
            If Index = 0 Then
                .Text = "01/" & Format(DateAdd("m", -1, Now), "mm/yyyy")
            Else
                .Text = Format(Now, "dd/mm/yyyy")
            End If
        End If
     End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If FrameDatos.Visible Then Exit Sub
    If Button.Index > 5 And Button.Index < 9 Then
        If vUsu.Nivel > 1 Then 'solo admon
            MsgBox "No tiene autorizacion para realizar cambios", vbExclamation
            Exit Sub
        End If
    End If
    HacerToolBar Button.Index
End Sub

Private Sub HacerToolBar(Indice As Integer)
    Select Case Indice
    Case 2
        Screen.MousePointer = vbHourglass
        CargaGrid
        Screen.MousePointer = vbDefault
    
    Case 6
        InsertarModificar True
    Case 7
        Modificar
    Case 8
        Eliminar
    Case 10
        Eliminarjornada
    Case 12
        Unload Me
    End Select
End Sub


Private Sub Eliminarjornada()
Dim SQL As String
Dim RS As ADODB.Recordset

    If adodc1.Recordset Is Nothing Then Exit Sub
    If adodc1.Recordset.EOF Then Exit Sub
    
    SQL = "Seguro que desea eliminar la jornada " & Format(adodc1.Recordset!Fecha, "dd/mm/yyyy")
    If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    
    
    SQL = DevuelveDesdeBD("EmpresaHoraExtra", "Empresas", "idEmpresa", 1, "N")
    VariableCompartida = "NOHORAEXTRA"
    If SQL <> "" Then
        If CBool(SQL) Then VariableCompartida = ""
    End If
    
    
    
    'Comprobaremos k es la ultima joranada
    Set RS = New ADODB.Recordset
    SQL = "Select fecha from JornadasSemanales WHERE Fecha >#" & Format(adodc1.Recordset!Fecha, FormatoFecha) & "# GROUP BY Fecha"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS!Fecha) Then
        
            'Tiene posteriores.

        
            If VariableCompartida = "" Then
                'No lleva bolsa horas. Lo lleva dandoles extras cada mes
                SQL = "Existen Jornadas posteriores a la que desea borrar. "
                SQL = SQL & vbCrLf & " ¿ SEGURO que desea continuar ?"
                If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then SQL = ""
            Else
                SQL = "Existen Jornadas posteriores a la que desea borrar. "
                MsgBox SQL, vbExclamation
                SQL = ""
            End If
        Else
            SQL = ""
        End If
        
    End If
    RS.Close
    
    
    If SQL = "" Then Exit Sub
    

    
    Screen.MousePointer = vbHourglass
    
    'Para cada trabajador updateamos con su bolsa horas
    SQL = "SELECT * From JornadasSemanales WHERE fecha = #" & Format(adodc1.Recordset!Fecha, FormatoFecha) & "#"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RS.EOF
        SQL = "UPDATE Trabajadores SET BolsaHoras = " & TransformaComasPuntos(CStr(RS!bolsaantes))
        SQL = SQL & " WHERE idTrabajador =" & RS!idTrabajador
        Conn.Execute SQL
        RS.MoveNext
    Wend
    RS.Close
    SQL = "DELETE From JornadasSemanales WHERE fecha = #" & Format(adodc1.Recordset!Fecha, FormatoFecha) & "#"
    Conn.Execute SQL
    espera 1
    CargaGrid
    Screen.MousePointer = vbDefault
    
End Sub





Private Sub Txtdias_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Txtdias_LostFocus(Index As Integer)
    Txtdias(Index).Text = Trim(Txtdias(Index).Text)
    If Txtdias(Index).Text <> "" Then
        If Not IsNumeric(Txtdias(Index).Text) Then
            MsgBox "Campo numerico", vbExclamation
            Txtdias(Index).Text = ""
            Txtdias(Index).SetFocus
        End If
    End If
End Sub

Private Sub txtHoras_KeyPress(Index As Integer, KeyAscii As Integer)
        Keypress KeyAscii
End Sub

Private Sub txtHoras_LostFocus(Index As Integer)
    txtHoras(Index).Text = Trim(txtHoras(Index).Text)
    If txtHoras(Index).Text <> "" Then
        If Not IsNumeric(txtHoras(Index).Text) Then
            MsgBox "Campo debe ser numérico", vbExclamation
            txtHoras(Index).SetFocus
            Exit Sub
            
        End If
    End If
End Sub

Private Sub txtTra_GotFocus(Index As Integer)
    With txtTra(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub



Private Sub txtTra_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtTra_LostFocus(Index As Integer)
Dim Cad As String
    If txtTra(Index).Text <> "" Then
        If Not IsNumeric(txtTra(Index).Text) Then
            Cad = ""
        Else
            Cad = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idTrabajador", txtTra(Index).Text, "N")
        End If
        If Cad = "" Then
            MsgBox "Codigo incorrecto: " & txtTra(Index).Text, vbExclamation
            txtTra(Index).Text = ""
        End If
    Else
        Cad = ""
    End If
    Text2(Index).Text = Cad
End Sub

Private Sub Modificar()


    If adodc1.Recordset.EOF Then Exit Sub

    InsertarModificar False


End Sub


Private Sub Eliminar()

On Error GoTo EEliminar
    If adodc1.Recordset.EOF Then Exit Sub
    
    VariableCompartida = "Seguro que desea elimnar la semana de fecha: " & adodc1.Recordset!Fecha
    VariableCompartida = VariableCompartida & vbCrLf & "del trabajador : " & adodc1.Recordset!nomtrabajador & " ?"
    If MsgBox(VariableCompartida, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    Screen.MousePointer = vbHourglass
    VariableCompartida = "Delete from JornadasSemanales "
    VariableCompartida = VariableCompartida & " WHERE idTrabajador =" & adodc1.Recordset!idTrabajador & " AND Fecha =#" & Format(adodc1.Recordset!Fecha, FormatoFecha) & "#"
    Conn.Execute VariableCompartida
    
    espera 1
    
    HacerToolBar 2
    
    Exit Sub
EEliminar:
    MuestraError Err.Number, Err.Description
End Sub

Private Function DatosOk() As Boolean

    DatosOk = False
    If txtTra(1).Text = "" Then
        MsgBox "ponga el nombre del trabajador", vbExclamation
        Exit Function
    End If
        
        
    If Txtdias(0).Text = "" Or Txtdias(1).Text = "" Or _
        txtHoras(0).Text = "" Or txtHoras(1).Text = "" Then
            MsgBox "Campos no pueden estar en blanco", vbExclamation
            Exit Function
    End If
    If Text1(2).Text = "" Then
        MsgBox "Fecha en blanco", vbExclamation
        Exit Function
    End If
    
    'Solo se acptan los ultimos dias de mes
    If Weekday(CDate(Text1(2).Text), vbMonday) <> 7 Then
        If DiasMes(Month(CDate(Text1(2).Text)), Year(CDate(Text1(2).Text))) <> Day(CDate(Text1(2).Text)) Then
            MsgBox "No es ultimo dia semana ni ultimo dia de mes", vbExclamation
            Exit Function
        End If
    End If
    DatosOk = True
    
    

End Function

Private Function Ejecuta() As Boolean
Dim Cad As String

        
    On Error GoTo EEjecuta
    Ejecuta = False
    If txtTra(1).Enabled Then
        'INSERTAR
        Cad = "INSERT INTO JornadasSemanales (idTrabajador,Fecha,HorasOfi,DiasOfi,HN,Dias,HE) VALUES ("
        Cad = Cad & txtTra(1).Text & ",#" & Format(Text1(2).Text, FormatoFecha) & "#,"
        Cad = Cad & TransformaComasPuntos(txtHoras(0).Text) & "," & Val(Txtdias(0).Text) & ","
        Cad = Cad & TransformaComasPuntos(txtHoras(1).Text) & "," & Val(Txtdias(1).Text) & ","
        If txtHoras(2).Text = "" Then
            Cad = Cad & "0"
        Else
            Cad = Cad & TransformaComasPuntos(txtHoras(2).Text)
        End If
        Cad = Cad & ")"
    Else
        Cad = "UPDATE JornadasSemanales SET HorasOFI=" & TransformaComasPuntos(txtHoras(0).Text)
        Cad = Cad & ", DiasOFI =" & Val(Txtdias(0).Text)
        Cad = Cad & ", HN =" & TransformaComasPuntos(txtHoras(1).Text)
        Cad = Cad & ", HE ="
        If txtHoras(2).Text = "" Then
            Cad = Cad & "0"
        Else
            Cad = Cad & TransformaComasPuntos(txtHoras(2).Text)
        End If
        
        Cad = Cad & ",Dias =" & Val(Txtdias(1).Text)
        Cad = Cad & " WHERE idTrabajador =" & txtTra(1).Text & " AND Fecha =#" & Format(Text1(2).Text, FormatoFecha) & "#"
    End If
    
    Conn.Execute Cad
    espera 0.5
    Ejecuta = True
    Exit Function
EEjecuta:
    MuestraError Err.Number, Err.Description
End Function


Private Sub InsertarModificar(Insertar As Boolean)

    
    txtHoras(2).Text = ""
     
    If Insertar Then
        txtTra(1).Enabled = True
        txtTra(1).Text = ""
        Text2(1).Text = ""
        Text1(2).Text = Format(Now, "dd/mm/yyyy")
        Txtdias(0).Text = "": Txtdias(1).Text = ""
        txtHoras(0).Text = "": txtHoras(1).Text = ""
        ImgFech(2).Enabled = True
        Text1(2).Enabled = True
        Label2.Caption = "INSERTAR"
    Else
        Label2.Caption = "MODIFICAR"
        ' " JornadasSemanales.DiasOfi , JornadasSemanales.HorasOfi"
        ' " , JornadasSemanales.Dias ,JornadasSemanales.HN"
        txtTra(1).Enabled = False
        txtTra(1).Text = adodc1.Recordset!idTrabajador
        Text1(2).Text = adodc1.Recordset!Fecha
        Text1(2).Enabled = False
        ImgFech(2).Enabled = False
        Text2(1).Text = adodc1.Recordset!nomtrabajador
        Txtdias(0).Text = adodc1.Recordset!diasofi
        Txtdias(1).Text = adodc1.Recordset!Dias
        txtHoras(0).Text = TransformaComasPuntos(Format(adodc1.Recordset!horasofi, "0.00"))
        txtHoras(1).Text = TransformaComasPuntos(Format(adodc1.Recordset!HN, "0.00"))
        If Not IsNull(adodc1.Recordset.Fields(0)) Then _
             txtHoras(2).Text = TransformaComasPuntos(Format(adodc1.Recordset!HE, "0.00"))
        
    End If
    Me.FrameDatos.Visible = True
    Me.FrameTapa.Visible = True
    DataGrid1.Enabled = False
    If Not Insertar Then Txtdias(0).SetFocus
End Sub



Private Sub Keypress(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub
