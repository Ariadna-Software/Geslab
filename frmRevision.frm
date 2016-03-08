VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmRevision 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Revisión marcajes"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   8730
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRevisada 
      Caption         =   "Revisada"
      Height          =   375
      Left            =   180
      TabIndex        =   39
      Top             =   5400
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   16
      Left            =   6300
      TabIndex        =   36
      Text            =   "Text1"
      Top             =   4620
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   15
      Left            =   1920
      TabIndex        =   34
      Text            =   "Text1"
      Top             =   4620
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   14
      Left            =   5580
      TabIndex        =   33
      Text            =   "Text1"
      Top             =   4200
      Width           =   2985
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   13
      Left            =   4620
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   4200
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   12
      Left            =   1140
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   4200
      Width           =   2745
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   11
      Left            =   180
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   4200
      Width           =   765
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   10
      Left            =   7080
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   3300
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   9
      Left            =   4845
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   3300
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   8
      Left            =   2595
      TabIndex        =   23
      Text            =   "Text1"
      Top             =   3300
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   7
      Left            =   360
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   3300
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   7080
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2280
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   4845
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2280
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   4
      Left            =   2595
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2280
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   360
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   2280
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   5700
      TabIndex        =   10
      Text            =   "Text1"
      Top             =   1200
      Width           =   1155
   End
   Begin MSAdodcLib.Adodc data1 
      Height          =   330
      Left            =   6900
      Top             =   420
      Visible         =   0   'False
      Width           =   1995
      _ExtentX        =   3519
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
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   1740
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   1200
      Width           =   3315
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1155
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6060
      TabIndex        =   4
      Top             =   5400
      Width           =   1155
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   1920
      TabIndex        =   2
      Top             =   5220
      Width           =   2415
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   300
         TabIndex        =   3
         Top             =   240
         Width           =   1755
      End
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1200
      Width           =   1095
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":08DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":11B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":1A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":2368
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":2C42
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":351C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":3DF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":46D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   38
      Top             =   0
      Width           =   8730
      _ExtentX        =   15399
      _ExtentY        =   1111
      ButtonWidth     =   1402
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ver todos"
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
            Caption         =   "Nuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   6540
      Picture         =   "frmRevision.frx":4FAA
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   1860
      Picture         =   "frmRevision.frx":50AC
      Top             =   3960
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   2580
      Picture         =   "frmRevision.frx":51AE
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Incidencia automática"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Index           =   1
      Left            =   4620
      TabIndex        =   40
      Top             =   3960
      Width           =   1890
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00800000&
      X1              =   2040
      X2              =   8400
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      X1              =   2100
      X2              =   8340
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Label Label1 
      Caption         =   "Nº horas incidencia"
      Height          =   255
      Index           =   12
      Left            =   4620
      TabIndex        =   37
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Nº horas incidencia"
      Height          =   255
      Index           =   11
      Left            =   180
      TabIndex        =   35
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Incidencia manual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   30
      Top             =   3960
      Width           =   1635
   End
   Begin VB.Label Label1 
      Caption         =   "Hora salida 2"
      Height          =   255
      Index           =   10
      Left            =   7080
      TabIndex        =   28
      Top             =   3060
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Hora entrada 2"
      Height          =   255
      Index           =   9
      Left            =   4875
      TabIndex        =   26
      Top             =   3060
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Hora salida 1"
      Height          =   255
      Index           =   8
      Left            =   2625
      TabIndex        =   24
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Hora entrada 1"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   22
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Horas comprobadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   2820
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Hora salida 2"
      Height          =   255
      Index           =   6
      Left            =   7080
      TabIndex        =   19
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Hora entrada 2"
      Height          =   255
      Index           =   5
      Left            =   4875
      TabIndex        =   17
      Top             =   2040
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Hora salida 1"
      Height          =   255
      Index           =   4
      Left            =   2625
      TabIndex        =   15
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Hora entrada 1"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   13
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   2
      Left            =   5700
      TabIndex        =   11
      Top             =   960
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Horas marcajes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Empleado"
      Height          =   255
      Index           =   1
      Left            =   1740
      TabIndex        =   8
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Secuencia"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   1
      Top             =   960
      Width           =   1215
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
End
Attribute VB_Name = "frmRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Verde = &H8000&
Private Const Gris = &HC0C0C0


Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
Public vFecha As Date
Private PrimeraVez As Boolean
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
        'y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
'Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private TextoConsulta As String
Private NumRegistro As Long
Private Kcampo As Integer
Private TotalReg As Long
Private YaHeRefrescado As Boolean
'Para los calculos horarios
Private HorasTrabajadas As Single
Private TotalHora As Single
Private NumHorasIncidencia As Single
Private SignoPositivoManual As Single
Private SignoPositivoAutoma As Single
Private InciManual As Boolean
Private Indice As Integer 'Para la matriz de buscar



Private Function TextoAHora(cad As String) As Date
TextoAHora = "0:00:00"
If IsDate(cad) Then TextoAHora = Format(cad, "hh:mm:ss")
End Function

Private Sub cmdAceptar_Click()
Dim rs As ADODB.Recordset
Dim cad As String
Dim i As Integer

Screen.MousePointer = vbHourglass
On Error GoTo Error1
If Modo = 3 Then
    If DatosOk Then
        
        Set rs = New ADODB.Recordset
        rs.CursorType = adOpenKeyset
        rs.LockType = adLockOptimistic
        rs.Open "EntradaFichajes", Conn, , , adCmdTable
        rs.AddNew
        'Ahora insertamos
  
  
  
        rs!Entrada = Text1(0).Text
        rs!idTrabajador = Text1(1).Tag 'En el tag esta el codigo
        rs!Fecha = Format(Text1(2).Text)
        rs!HoraE1 = TextoAHora(Text1(3).Text)
        rs!HoraS1 = TextoAHora(Text1(4).Text)
        rs!HoraE2 = TextoAHora(Text1(5).Text)
        rs!HoraS2 = TextoAHora(Text1(6).Text)
  
  
  
        rs!HoraE1C = TextoAHora(Text1(7).Text)
        rs!HoraS1C = TextoAHora(Text1(8).Text)
        rs!HoraE2C = TextoAHora(Text1(9).Text)
        rs!HoraS2C = TextoAHora(Text1(10).Text)
        rs!InciMan = Text1(11).Text
        rs!IncAuto = Text1(13).Text
        rs!NumHorasManu = Text1(15).Text
        rs!NumHorasAuto = Text1(16).Text
        'Como lo insertamos esta bien
        rs!Correcto = True
        '--------------------
        rs.Update
        rs.Close
        data1.Refresh
        MsgBox " Registro insertado.", vbInformation
        LimpiarCampos
        PonerModo 0
    End If
    Else
    If Modo = 4 Then
        'Modificar
        If Not DatosOk Then Exit Sub

        'Modificamos sobre el recordset
        data1.Recordset!HoraE1C = TextoAHora(Text1(7).Text)
        data1.Recordset!HoraS1C = TextoAHora(Text1(8).Text)
        data1.Recordset!HoraE2C = TextoAHora(Text1(9).Text)
        data1.Recordset!HoraS2C = TextoAHora(Text1(10).Text)
        data1.Recordset!InciMan = Text1(11).Text
        data1.Recordset!IncAuto = Text1(13).Text
        data1.Recordset!NumHorasManu = Text1(15).Text
        data1.Recordset!NumHorasAuto = Text1(16).Text
        data1.Recordset!HorasTrabajadas = HorasTrabajadas
        data1.Recordset.Update
        data1.Refresh
        PonerModo 2
        Label2.Caption = "Modificado"
    End If
End If
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub


Private Sub cmdCancelar_Click()
If Modo = 3 Then
    'Como estamos insertando
    LimpiarCampos
    PonerModo 0
ElseIf Modo = 4 Then
    PonerCampos
    PonerModo 2
End If
End Sub


Private Sub BotonAnyadir()
LimpiarCampos
'Añadiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Aceptar"
PonerModo 3
'Sugerimos el código
SugerirCodigoSiguiente
'Ponemos Las incidencias a 0 y las horas automaticas tb
Text1(11).Text = 0
Text1(13).Text = 0
Text1(15).Text = 0
Text1(16).Text = 0
'Escondemos el navegador y ponemos insertando
Label2.Caption = "Insertar"
Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
'Buscar
If Modo <> 1 Then
    LimpiarCampos
    PonerModo 1
    Text1(0).SetFocus
    Else
        HacerBusqueda
        If TotalReg = 0 Then
            Text1(Kcampo).Text = ""
            Text1(Kcampo).BackColor = vbYellow
            Text1(Kcampo).SetFocus
        End If
End If
End Sub

Private Sub BotonVerTodos()
'Ver todos
LimpiarCampos
PonerModo 2
CadenaConsulta = TextoConsulta
PonerCadenaBusqueda
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 0
        data1.Recordset.MoveFirst
        NumRegistro = 1
    Case 1
        data1.Recordset.MovePrevious
        NumRegistro = NumRegistro - 1
        If data1.Recordset.BOF Then
            data1.Recordset.MoveFirst
            NumRegistro = 1
        End If
    Case 2
        data1.Recordset.MoveNext
        NumRegistro = NumRegistro + 1
        If data1.Recordset.EOF Then
            If Not YaHeRefrescado Then
                data1.Refresh
                YaHeRefrescado = True
            End If
            data1.Recordset.MoveLast
            NumRegistro = TotalReg
        End If
    Case 3
            If Not YaHeRefrescado Then
                data1.Refresh
                YaHeRefrescado = True
            End If
        data1.Recordset.MoveLast
        NumRegistro = TotalReg
End Select
PonerCampos
End Sub

Private Sub BotonModificar()
'---------
'MODIFICAR
'----------
'Añadiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Modificar"
PonerModo 4
'Escondemos el navegador y ponemos insertando
'Como el campo 1 es clave primaria, NO se puede modificar
Text1(0).Locked = True
Label2.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
Dim cad As String
Dim i As Integer

'Ciertas comprobaciones
If data1.Recordset.RecordCount = 0 Then Exit Sub
'Pregunta
cad = "Seguro que desea eliminar de la BD el registro:"
cad = cad & vbCrLf & "Cod: " & data1.Recordset.Fields(0)
cad = cad & vbCrLf & "Nombre: " & data1.Recordset.Fields(1)
i = MsgBox(cad, vbQuestion + vbYesNo)
If i = vbYes Then
    'Hay que eliminar
    On Error GoTo Error2
    Screen.MousePointer = vbHourglass
    data1.Recordset.Delete
    data1.Refresh
    If data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        PonerModo 0
        Else
            If NumRegistro = TotalReg Then
                    data1.Recordset.MoveLast
                    NumRegistro = NumRegistro - 1
                    Else
                        For i = 1 To NumRegistro - 1
                            data1.Recordset.MoveNext
                        Next i
            End If
            TotalReg = TotalReg - 1
            PonerCampos
    End If
End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub cmdRevisada_Click()
'Si pulsa aqui significa que ha revisado el
'marcaje con lo que hacemos una última comprobación
'de las horas y lo ponemos a correcto
'teniendo en cuenta que
' EXCESO: horas extras
' DEFECTO: no llega al mínimo del horario
On Error GoTo ErrRevision
If DatosOk Then
    MsgBox "Todo bien"
    'Ponemos la marca de todo correcto a True
    data1.Recordset!Correcto = True
    data1.Recordset.Update
    data1.Refresh
    If data1.Recordset.EOF Then
        MsgBox "La lista de entradas para revisar esta vacia"
        Unload Me
        Else
            TotalReg = data1.Recordset.RecordCount
            NumRegistro = 1
            data1.Recordset.MoveFirst
            PonerCampos
    End If
End If
ErrRevision:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description
End Sub

Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Screen.MousePointer = vbDefault
If PrimeraVez Then
    PrimeraVez = False
    data1.Recordset.Requery
    data1.Refresh
    BotonVerTodos
End If
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
PrimeraVez = True
LimpiarCampos
TextoConsulta = "SELECT EntradaFichajes.*, Trabajadores.NomTrabajador" & _
        " FROM Trabajadores ,EntradaFichajes WHERE Trabajadores.IdTrabajador = EntradaFichajes.idTrabajador" & _
        " AND Correcto=False"
If Not IsNull(vFecha) Then
    If vFecha <> "0:00:00" Then _
    TextoConsulta = TextoConsulta & " AND Fecha=#" & Format(vFecha, "yyyy/mm/dd") & "#"
End If
Ordenacion = " ORDER BY EntradaFichajes.IdTrabajador"
YaHeRefrescado = False
'ASignamos un SQL al DATA1
data1.ConnectionString = Conn
data1.RecordSource = TextoConsulta
data1.Refresh
PonerModo 0
End Sub



Private Sub LimpiarCampos()
Dim i
For i = 0 To Text1.Count - 1
    Text1(i).Text = ""
Next i
End Sub

Private Sub Form_Unload(Cancel As Integer)
vFecha = "0:00:00"
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
If Indice = 0 Then
    Text1(1).Text = vCadena
    Text1(1).Tag = vCodigo
    Else
        If Indice = 1 Then
            Text1(11).Text = vCodigo
            'Texto de las incidencias
            Text1(12).Text = DevuelveTextoIncidencia(CInt(vCodigo), SignoPositivoManual)
            Else
                If Indice = 2 Then
                    Text1(13).Text = vCodigo
                    Text1(14).Text = DevuelveTextoIncidencia(CInt(vCodigo), SignoPositivoAutoma)
                End If
        End If
End If
End Sub

Private Sub Image1_Click(Index As Integer)
Indice = Index
Select Case Index
Case 0
    'Nombre tyrabajador
    'Ponemos los valores para abrir
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "NomTrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPLEADOS"
    frmB.Show vbModal
    Set frmB = Nothing
Case 1, 2
    
    'Incidencia 1
    'Ponemos los valores para abrir
    Set frmB = New frmBusca
    frmB.Tabla = "Incidencias"
    frmB.CampoBusqueda = "NomInci"
    frmB.CampoCodigo = "IdInci"
    frmB.TipoDatos = 3
    frmB.Titulo = "INCIDENCIAS"
    frmB.Show vbModal
    Set frmB = Nothing
  
End Select
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Kcampo = Index
If Modo = 1 Then
    Text1(Index).BackColor = vbYellow
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Modo = 1 Then
    If KeyAscii = 13 Then
        'Ha pulsado enter, luego tenemos que hacer la busqueda
        Text1(Index).BackColor = vbWhite
        BotonBuscar
    End If
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Modo = 1 Then
    Text1(Index).BackColor = vbWhite
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

If Text1(Kcampo).Text = "" Then Exit Sub

'------------------------------------------------
Dim c1 As String   'el nombre del campo
Dim tipo As Long
Dim aux1

c1 = data1.Recordset.Fields(Kcampo).Name
c1 = " AND " & c1
'Stop

tipo = DevuelveTipo2(data1.Recordset.Fields(Kcampo).Type)

'Devolvera uno de los tipos
'   1.- Numeros
'   2.- Booleanos
'   3.- Cadenas
'   4.- Fecha
'   0.- Error leyendo los tipos de datos
' segun sea uno u otro haremos una comparacion
Select Case tipo
Case 1
    CadB = c1 & " = " & Text1(Kcampo)
Case 2
    'Vemos si la cadena tiene un Falso o False
    If InStr(1, UCase(Text1(Kcampo).Text), "F") Then
        aux1 = "False"
        Else
        aux1 = "True"
    End If
    CadB = c1 & " = " & aux1
Case 3
    CadB = c1 & " like '%" & Trim(Text1(Kcampo)) & "%'"
Case 4

Case 5

End Select
    
CadenaConsulta = TextoConsulta & CadB & " " & Ordenacion
PonerCadenaBusqueda
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo Error4
data1.RecordSource = CadenaConsulta
data1.Refresh
If data1.Recordset.EOF Then
    MsgBox "No hay ningún registro en la tabla de entrada de fichajes.", vbInformation
    Screen.MousePointer = vbDefault
    TotalReg = 0
    Exit Sub
    Label2.Caption = ""
    'PonerModo 0
    Else
        PonerModo 2
        data1.Recordset.MoveLast
        data1.Recordset.MoveFirst
        TotalReg = data1.Recordset.RecordCount
        NumRegistro = 1
        PonerCampos
End If

'Data1.ConnectionString = Conn
'Data1.RecordSource = CadenaConsulta
'Data1.Refresh
'TotalReg = Data1.Recordset.RecordCount
Error4:
Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim i As Integer
    Text1(0).Text = data1.Recordset.Fields(0)
    Text1(1).Text = data1.Recordset.Fields(17)
    'El codigo del trabajador
    Text1(1).Tag = data1.Recordset.Fields(1)
    Text1(2).Text = DBLet(data1.Recordset.Fields(3))
    Text1(3).Text = DBLet(data1.Recordset.Fields(4))
    Text1(4).Text = DBLet(data1.Recordset.Fields(5))
    Text1(5).Text = DBLet(data1.Recordset.Fields(6))
    Text1(6).Text = DBLet(data1.Recordset.Fields(7))
    Text1(7).Text = DBLet(data1.Recordset.Fields(8))
    Text1(8).Text = DBLet(data1.Recordset.Fields(9))
    Text1(9).Text = DBLet(data1.Recordset.Fields(10))
    Text1(10).Text = DBLet(data1.Recordset.Fields(11))
    Text1(11).Text = DBLet(data1.Recordset.Fields(15))
    Text1(13).Text = DBLet(data1.Recordset.Fields(16))
    Text1(15).Text = DBLet(data1.Recordset.Fields(12))
    Text1(16).Text = DBLet(data1.Recordset.Fields(13))
    'Texto de las incidencias
    Text1(12).Text = DevuelveTextoIncidencia(data1.Recordset.Fields(15), SignoPositivoManual)
    Text1(14).Text = DevuelveTextoIncidencia(data1.Recordset.Fields(16), SignoPositivoAutoma)
    
    If data1.Recordset.Fields(15) > 0 Then
        InciManual = True
        Label5(0).ForeColor = Verde
        Label5(1).ForeColor = Gris
        Else
            InciManual = False
            Label5(1).ForeColor = Verde
            Label5(0).ForeColor = Gris
    End If
    
Label2.Caption = NumRegistro & " de " & TotalReg


End Sub



'AGRUPAR PARA QUE NO HAGA TANTAS COMPARACIONES
Private Sub PonerModo(Kmodo As Integer)
Dim i As Integer
Dim b As Boolean

Modo = Kmodo
DespalzamientoVisible (Kmodo = 2)
cmdAceptar.Visible = (Kmodo >= 3)
cmdCancelar.Visible = (Kmodo >= 3)
Toolbar1.Buttons(6).Enabled = (Kmodo < 3)
Toolbar1.Buttons(7).Enabled = (Kmodo = 2)
Toolbar1.Buttons(8).Enabled = (Kmodo = 2)
Toolbar1.Buttons(1).Enabled = (Kmodo < 3)
Toolbar1.Buttons(2).Enabled = (Kmodo < 3)
Label2.Visible = (Kmodo = 2)
b = (Modo = 2) Or Modo = 0
For i = 0 To Text1.Count - 1
    Text1(i).Locked = b
Next i
cmdRevisada.Visible = Kmodo < 3
Image1(0).Visible = Kmodo = 3
Image1(1).Visible = Kmodo > 2
Image1(2).Visible = Image1(1).Visible
'Campos de texto
Text1(2).Enabled = Kmodo = 3
End Sub


Private Function DatosOk() As Boolean
Dim cad As String
Dim i As Integer
Dim vh As CHorarios
Dim rs As ADODB.Recordset
Dim T1 As Single
Dim T2 As Single
Dim NHA As Single 'num horas automat
Dim NHM As Single ' numero de horas manuales

On Error GoTo ErrorDatosOk
DatosOk = False
'ciertos valores no nulos
For i = 0 To 2
    If Text1(i).Text = "" Then
        MsgBox "El campo " & Label1(i).Caption & " NO puede estar vacio.", vbExclamation
        GoTo ErrorDatosOk
    End If
Next i

If Not IsDate(Text1(2).Text) Then
    MsgBox "La fecha no es correcta.", vbExclamation
    GoTo ErrorDatosOk
End If
Text1(2).Text = Format(Text1(2).Text, "dd/mm/yyyy")
For i = 3 To 10
    If Text1(i).Text <> "" Then
        If Not IsDate(Text1(i).Text) Then
            MsgBox "No es una hora válida", vbExclamation
            GoTo ErrorDatosOk
        End If
    End If
Next i

'INCIDENCIAS
'Manual
If Text1(11).Text <> "" Then
    If Not IsNumeric(Text1(11)) Then
        MsgBox "Error en la incidencia manual.", vbExclamation
    End If
End If
'Automatica
If Text1(13).Text = "" Then
    If Not IsNumeric(Text1(13)) Then
        MsgBox "Error en la incidencia automática.", vbExclamation
    End If
End If
'  y las horas
'  manuales ..
    If Text1(15).Text <> "" Then
        If Not IsNumeric(Text1(15)) Then
            MsgBox "Error en las horas manuales.", vbExclamation
            GoTo ErrorDatosOk
            Else
                Text1(15).Tag = Text1(15).Text
        End If
        Else
            Text1(15).Tag = 0
    End If
'   automaticas
    If Text1(16).Text <> "" Then
        If Not IsNumeric(Text1(16)) Then
            MsgBox "Error en las horas automáticas.", vbExclamation
            GoTo ErrorDatosOk
            Else
                Text1(16).Tag = Text1(16).Text
        End If
        Else
            Text1(16).Tag = 0
    End If
'Seguro que las dos primeras horas comprobadas NO estan vacias
If Text1(7).Text = "" Then
    MsgBox "La primera hora comprobada de entrada no puede estar vacia."
    GoTo ErrorDatosOk
End If
If Text1(8).Text = "" Then
    MsgBox "La primera hora comprobada de salida no puede estar vacia."
    GoTo ErrorDatosOk
End If
'Obtener el horario
i = -1
cad = "Select IdHorario from Trabajadores where IdTrabajador=" & Text1(1).Tag
Set rs = New ADODB.Recordset
rs.Open cad, Conn, , , adCmdText
If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then
          i = rs.Fields(0)
    End If
End If
rs.Close
If i < 0 Then
    MsgBox "Error en el horario asignado a ese cliente.", vbExclamation
    GoTo ErrorDatosOk
End If
'Leemos el horario
T1 = 0
T2 = 0
HorasTrabajadas = 0
Set vh = New CHorarios
If vh.Leer(i, CDate(Text1(2).Text)) = 0 Then
    'T1 = DevuelveValorHora(CDate(Text1(8).Text) - CDate(Text1(7).Text))
    'T2 = DevuelveValorHora(CDate(Text1(10).Text) - CDate(Text1(9).Text))
    T1 = RevisaHoras
    HorasTrabajadas = T1
    NHM = SignoPositivoManual * CSng(Text1(15).Text)
    NHA = SignoPositivoAutoma * CSng(Text1(16).Text)
    T2 = T1 + NHM + NHA
    Else
        GoTo ErrorDatosOk
End If

'Ahora en t2 tenemos las horas trabajadas sumando
' o restando la de las incidencias
T2 = Round(vh.TotalHoras - T2, 0)
If T2 <> 0 Then
    MsgBox "Error en el computo de horas." & vbCrLf & _
        "Horas trabajadas: " & T1 & vbCrLf & _
        "Horas incidencias: " & (NHM + NHA) & vbCrLf & _
        "Horas jornada laboral: " & vh.TotalHoras
    GoTo ErrorDatosOk
End If
Set vh = Nothing
'Al final todo esta correcto
DatosOk = True
ErrorDatosOk:
    If Err.Number <> 0 Then MsgBox "error " & Err.Description
    Screen.MousePointer = vbDefault
End Function


Private Sub SugerirCodigoSiguiente()
Dim cad
Dim rs

'Sugeriremos el codigo siguiente.
'Obviamente depende en TOTAL medida de que tabla estemos trabajando
cad = "Select Max(Entrada) from EntradaFichajes"

Text1(0).Text = 1
Set rs = New ADODB.Recordset
rs.Open cad, Conn, , , adCmdText
If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then
          Text1(0).Text = rs.Fields(0) + 1
    End If
End If
rs.Close
End Sub



Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    BotonBuscar
Case 2
    BotonVerTodos
Case 6
    BotonAnyadir
Case 7
    BotonModificar
Case 8
    BotonEliminar
Case 14 To 17
    Desplazamiento (Button.Index - 14)
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
Dim i
For i = 14 To 17
    Toolbar1.Buttons(i).Visible = bol
Next i
End Sub






Private Function RevisaHoras() As Single
Dim v1(3) As Single
Dim T1 As Single
Dim T2 As Single
Dim k As Integer

RevisaHoras = 0
For k = 0 To 3
    If Text1(7 + k).Text <> "" Then
        If IsDate(Text1(7 + k).Text) Then
              v1(k) = DevuelveValorHora(Text1(7 + k).Text)
        Else
            v1(k) = 0
        End If
    Else
        v1(k) = 0
    End If
Next k
T1 = v1(1) - v1(0)
T2 = v1(3) - v1(2)
RevisaHoras = T1 + T2
End Function
