VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSeccion 
   Caption         =   "Mantenimientos secciones"
   ClientHeight    =   4605
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   8310
   Icon            =   "frmSeccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   8310
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3435
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   6059
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   529
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmSeccion.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Text1(0)"
      Tab(0).Control(1)=   "Text1(1)"
      Tab(0).Control(2)=   "Text1(2)"
      Tab(0).Control(3)=   "Text1(3)"
      Tab(0).Control(4)=   "Text1(4)"
      Tab(0).Control(5)=   "Text2(0)"
      Tab(0).Control(6)=   "Text2(1)"
      Tab(0).Control(7)=   "Check1"
      Tab(0).Control(8)=   "Label1(0)"
      Tab(0).Control(9)=   "Label1(1)"
      Tab(0).Control(10)=   "Label1(2)"
      Tab(0).Control(11)=   "Label1(3)"
      Tab(0).Control(12)=   "Label1(4)"
      Tab(0).Control(13)=   "Image2(0)"
      Tab(0).Control(14)=   "Image2(1)"
      Tab(0).Control(15)=   "Label3"
      Tab(0).ControlCount=   16
      TabCaption(1)   =   "Rectificados de marcajes"
      TabPicture(1)   =   "frmSeccion.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label4(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label4(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label4(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label4(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command1(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command1(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command1(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ListView1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Combo1"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).ControlCount=   10
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmSeccion.frx":0342
         Left            =   6000
         List            =   "frmSeccion.frx":0355
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   720
         Width           =   1815
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   1935
         Left            =   240
         TabIndex        =   28
         Top             =   1140
         Width           =   5835
         _ExtentX        =   10292
         _ExtentY        =   3413
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Inicio"
            Object.Width           =   3242
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fin"
            Object.Width           =   3242
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Rectificado"
            Object.Width           =   3242
         EndProperty
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   2
         Left            =   6240
         Picture         =   "frmSeccion.frx":03AD
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2760
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   1
         Left            =   6240
         Picture         =   "frmSeccion.frx":04AF
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton Command1 
         Height          =   375
         Index           =   0
         Left            =   6240
         Picture         =   "frmSeccion.frx":05B1
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   -72060
         TabIndex        =   0
         Tag             =   "Código|N|S|||"
         Text            =   "Text1"
         Top             =   660
         Width           =   435
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   -72060
         TabIndex        =   1
         Tag             =   "Nombre|T|N|||"
         Text            =   "Text1"
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   -72060
         TabIndex        =   2
         Tag             =   "Empresa|N|N|1||"
         Text            =   "Text1"
         Top             =   1500
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   3
         Left            =   -72060
         TabIndex        =   3
         Tag             =   "Horario|N|N|1||"
         Text            =   "Text1"
         Top             =   1920
         Width           =   675
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   4
         Left            =   -72060
         TabIndex        =   4
         Tag             =   "Cod Control|N|N|0||"
         Text            =   "Text1"
         Top             =   2340
         Width           =   435
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -71280
         TabIndex        =   14
         Text            =   "Text2"
         Top             =   1500
         Width           =   2595
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -71280
         TabIndex        =   13
         Text            =   "Text2"
         Top             =   1920
         Width           =   2595
      End
      Begin VB.CheckBox Check1 
         Alignment       =   1  'Right Justify
         Height          =   195
         Left            =   -72120
         TabIndex        =   5
         Top             =   2820
         Width           =   255
      End
      Begin VB.Label Label5 
         Caption         =   "Rectificación"
         Height          =   255
         Left            =   6000
         TabIndex        =   30
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   7020
         TabIndex        =   27
         Top             =   2820
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   6960
         TabIndex        =   26
         Top             =   2100
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   7020
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Todos los marcajes comprendidos entre marcaje inicio y marcaje fin serán modificados a marcaje hora modificada"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   5475
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo sección"
         Height          =   195
         Index           =   0
         Left            =   -73860
         TabIndex        =   20
         Top             =   720
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre  Seccion"
         Height          =   195
         Index           =   1
         Left            =   -73860
         TabIndex        =   19
         Top             =   1140
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         Height          =   195
         Index           =   2
         Left            =   -73860
         TabIndex        =   18
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Horario predeter."
         Height          =   195
         Index           =   3
         Left            =   -73860
         TabIndex        =   17
         Top             =   1980
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Cod. control predeter."
         Height          =   195
         Index           =   4
         Left            =   -73860
         TabIndex        =   16
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   -72900
         Picture         =   "frmSeccion.frx":06B3
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   1
         Left            =   -72600
         Picture         =   "frmSeccion.frx":07B5
         Top             =   1980
         Width           =   240
      End
      Begin VB.Label Label3 
         Caption         =   "Nominas Ariagro"
         Height          =   195
         Left            =   -73860
         TabIndex        =   15
         Top             =   2820
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   180
      TabIndex        =   10
      Top             =   3960
      Width           =   3495
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   180
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5820
      TabIndex        =   7
      Top             =   4080
      Width           =   1035
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4080
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   300
      Top             =   4140
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
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
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   4500
      TabIndex        =   6
      Top             =   4080
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   8310
      _ExtentX        =   14658
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
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
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
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
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   7800
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":08B7
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":09C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":0ADB
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":0BED
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":0CFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":0E11
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":16EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":1FC5
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":289F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSeccion.frx":3179
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Visible         =   0   'False
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
Attribute VB_Name = "frmSeccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

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
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private NumRegistro As Long
Private kCampo As Integer
Private TotalReg As Long
Private vIndice As Byte




Private Sub cmdAceptar_Click()
Dim RS As ADODB.Recordset
Dim Cad As String
Dim i As Integer

Screen.MousePointer = vbHourglass
On Error GoTo Error1
If Modo = 3 Then
    If DatosOk Then
        
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenKeyset
        RS.LockType = adLockOptimistic
        RS.Open NombreTabla, Conn, , , adCmdTable
        RS.AddNew
        'Para luego
        For i = 0 To Text1.Count - 1
            RS.Fields(i) = Text1(i).Text
        Next i
        If Check1.Visible Then
            RS!nominas = (Check1.Value = 1)
            Else
                RS!nominas = False
        End If
        RS!RedondearCadaTicaje = Combo1.ListIndex
        '--------------------
        RS.Update
        RS.Close
        Data1.Refresh
        MsgBox "Registro insertado.", vbInformation
        PonerModo 0
        Label2.Caption = "Insertado"
    End If
    Else
    If Modo = 4 Then
        'Modificar
        ''Haremos las comprobaciones necesarias de los campos
        For i = 1 To Text1.Count - 1
            If Not CmpCam(Text1(i).Tag, Text1(i).Text) Then Exit Sub
        Next i

        'Ahora modificamos
        Cad = "Select * from " & NombreTabla
        Cad = Cad & " WHERE idSeccion=" & Data1.Recordset.Fields(0)
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenKeyset
        RS.LockType = adLockOptimistic
        RS.Open Cad, Conn, , , adCmdText
        'Almacenamos para luego buscarlo
        Cad = RS!IdSeccion
        'Modificamos
        For i = 1 To Text1.Count - 1
            RS.Fields(i) = Text1(i).Text
        Next i

        If Check1.Visible Then
            RS!nominas = (Check1.Value = 1)
        End If
        RS!RedondearCadaTicaje = Combo1.ListIndex
        RS.Update
        RS.Close
        MsgBox "El registro ha sido modificado", vbInformation
        PonerModo 2
        'Hay que refresca el DAta1
        Data1.Refresh
        'Hay que volver a poner el registro donde toca
        Data1.Recordset.MoveFirst
        i = 1
        While i > 0
            If Data1.Recordset.Fields(0) = Cad Then
                i = 0
                Else
                    Data1.Recordset.MoveNext
                    If Data1.Recordset.EOF Then i = 0
            End If
        Wend
        If Data1.Recordset.EOF Then
            NumRegistro = TotalReg
            Data1.Recordset.MoveLast
        End If
        Label2.Caption = NumRegistro & " de " & TotalReg
    End If
End If
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
'LimpiarCampos
'PonerModo 0


PonerModo 2

End Sub

Private Sub BotonAnyadir()
LimpiarCampos
'Añadiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Aceptar"
PonerModo 3
'Escondemos el navegador y ponemos insertando
DespalzamientoVisible False
Label2.Caption = "INSERTANDO"
SugerirCodigoSiguiente
Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
'Buscar
If Modo <> 1 Then
    LimpiarCampos
    Label2.Caption = "Búsqueda"
    PonerModo 1
    Text1(0).SetFocus
    Else
        HacerBusqueda
        If TotalReg = 0 Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            Text1(kCampo).SetFocus
        End If
End If
End Sub

Private Sub BotonVerTodos()
'Ver todos
LimpiarCampos
PonerModo 2
CadenaConsulta = "Select * from " & NombreTabla
PonerCadenaBusqueda
End Sub

Private Sub Desplazamiento(Index As Integer)
Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
        NumRegistro = 1
    Case 1
        Data1.Recordset.MovePrevious
        NumRegistro = NumRegistro - 1
        If Data1.Recordset.BOF Then
            Data1.Recordset.MoveFirst
            NumRegistro = 1
        End If
    Case 2
        Data1.Recordset.MoveNext
        NumRegistro = NumRegistro + 1
        If Data1.Recordset.EOF Then
            Data1.Recordset.MoveLast
            NumRegistro = TotalReg
        End If
    Case 3
        Data1.Recordset.MoveLast
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
DespalzamientoVisible False
Label2.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim i As Integer

'Ciertas comprobaciones
If Data1.Recordset.EOF Then Exit Sub
If Data1.Recordset.RecordCount = 0 Then Exit Sub
'Pregunta
Cad = "Seguro que desea eliminar de la BD el registro:"
Cad = Cad & vbCrLf & "Cod: " & Data1.Recordset.Fields(0)
Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
i = MsgBox(Cad, vbQuestion + vbYesNo)
If i = vbYes Then
    'Hay que eliminar
    On Error GoTo Error2
    Screen.MousePointer = vbHourglass
    Data1.Recordset.Delete
    espera 0.75
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        PonerModo 0
        Else
            If NumRegistro = TotalReg Then
                    Data1.Recordset.MoveLast
                    NumRegistro = NumRegistro - 1
                    Else
                        For i = 1 To NumRegistro - 1
                            Data1.Recordset.MoveNext
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


Private Sub Command1_Click(Index As Integer)
Dim Cad As String

On Error GoTo ErrCommand
If Index > 0 Then
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione una rectificación de marcajes.", vbExclamation
        Exit Sub
    End If
End If

Select Case Index
Case 0, 1
    Screen.MousePointer = vbHourglass
    'Nuevo o modificar
    frmRectificar.Nuevo = Index = 0
    frmRectificar.IdSeccion = Data1.Recordset.Fields(0)
    If Index = 1 Then
        frmRectificar.Id = ListView1.SelectedItem.Tag
        'Le pasamos las horas
        frmRectificar.Text1(0).Text = Format(ListView1.SelectedItem.Text, "hh:mm")
        frmRectificar.Text1(1).Text = Format(ListView1.SelectedItem.SubItems(1), "hh:mm")
        frmRectificar.Text1(2).Text = Format(ListView1.SelectedItem.SubItems(2), "hh:mm")
        Else
            frmRectificar.Text1(0).Text = ""
            frmRectificar.Text1(1).Text = ""
            frmRectificar.Text1(2).Text = ""
    End If
    frmRectificar.Show vbModal
Case 2
    Cad = "Seguro que desea eliminar el intervalo: " & vbCrLf
    Cad = Cad & " Inicio: " & ListView1.SelectedItem.Text & vbCrLf
    Cad = Cad & " fin: " & ListView1.SelectedItem.SubItems(1) & vbCrLf
    Cad = Cad & " Rectificada: " & ListView1.SelectedItem.SubItems(2)
    If MsgBox(Cad, vbExclamation + vbYesNoCancel) = vbYes Then
        Cad = "Delete * from ModificarFichajes where Id=" & ListView1.SelectedItem.Tag
        Screen.MousePointer = vbHourglass
        Conn.Execute Cad
    End If
End Select
Screen.MousePointer = vbHourglass
espera 0.1 '1 segundo
'Refrescamos la BD de marcajes modificados
PonDBGridModificados Data1.Recordset.Fields(0)
Screen.MousePointer = vbDefault
Exit Sub
ErrCommand:
    MuestraError Err.Number
End Sub

Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
LimpiarCampos
NombreTabla = "Secciones"
Label3.Caption = "Nóminas"
If mConfig.Ariadna Then Label3.Caption = Label3.Caption & " Ariagro"
Ordenacion = " ORDER BY IdSeccion"
SSTab1.Tab = 0
'ASignamos un SQL al DATA1
Data1.ConnectionString = Conn
Data1.RecordSource = "Select * from " & NombreTabla
Data1.Refresh
PonerModo 0
End Sub



Private Sub LimpiarCampos()
Dim i
For i = 0 To Text1.Count - 1
    Text1(i).Text = ""
Next i
For i = 0 To 1
    Text2(i).Text = ""
Next i
'lIMPIAMOS EL dbgrID
PonDBGridModificados -1
End Sub


Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
If vIndice = 1 Then
    'Es Horario
    Text1(3).Text = vCodigo
    Text2(1).Text = vCadena
    Else
        Text1(2).Text = vCodigo
        Text2(0).Text = vCadena
End If
End Sub

Private Sub Image2_Click(Index As Integer)
    Set frmB = New frmBusca

    'En el tag de txext2 tendremos quien lo llama
    If Index = 0 Then
        frmB.Tabla = "Empresas"
        frmB.CampoBusqueda = "NomEmpresa"
        frmB.CampoCodigo = "IdEmpresa"
        frmB.Titulo = "EMPRESAS"
        Else
            frmB.Tabla = "Horarios"
            frmB.CampoBusqueda = "NomHorario"
            frmB.CampoCodigo = "IdHorario"
            frmB.Titulo = "HORARIOS"
    End If
    vIndice = Index
    'Resto de cosas comunes
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Show vbModal
    Set frmB = Nothing
End Sub


Private Sub mnBuscar_Click()
BotonBuscar
End Sub

Private Sub mnEliminar_Click()
BotonEliminar
End Sub

Private Sub mnModificar_Click()
BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Screen.MousePointer = vbHourglass
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub



Private Sub Text1_GotFocus(Index As Integer)
kCampo = Index
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
Dim Cad As String
'
'
Text1(Index).BackColor = vbWhite
If Text1(Index).Text = "" Then Exit Sub
If Modo > 2 Then
    If Index = 2 Then
        Cad = ""
        'EMPRESA
        If IsNumeric(Text1(2).Text) Then
            Cad = DevuelveNombreEmpresa(CLng(Val(Text1(2).Text)))
        End If
        If Cad = "" Then
            Text1(2).Text = "-1"
            Text2(0).Text = "Empresa incorrecta"
            Else
                Text2(0).Text = Cad
        End If
        'ELSE DE INDEX
        Else
            If Index = 3 Then
                Cad = ""
                'HORARIO
                If IsNumeric(Text1(3).Text) Then
                    Cad = DevuelveNombreHorario(CLng(Val(Text1(3).Text)))
                End If
                If Cad = "" Then
                    Text1(3).Text = "-1"
                    Text2(1).Text = "Horario incorrecto"
                    Else
                        Text2(1).Text = Cad
                End If
            End If
    End If
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim c1 As String   'el nombre del campo
Dim Tipo As Long
Dim aux1

If Text1(kCampo).Text = "" Then Exit Sub
c1 = Data1.Recordset.Fields(kCampo).Name
c1 = " WHERE " & c1
Tipo = DevuelveTipo2(Data1.Recordset.Fields(kCampo).Type)
'Devolvera uno de los tipos
'   1.- Numeros
'   2.- Booleanos
'   3.- Cadenas
'   4.- Fecha
'   0.- Error leyendo los tipos de datos
' segun sea uno u otro haremos una comparacion
Select Case Tipo
Case 1
    CadB = c1 & " = " & Text1(kCampo)
Case 2
    'Vemos si la cadena tiene un Falso o False
    If InStr(1, UCase(Text1(kCampo).Text), "F") Then
        aux1 = "False"
        Else
        aux1 = "True"
    End If
    CadB = c1 & " = " & aux1
Case 3
    CadB = c1 & " like '*" & Trim(Text1(kCampo)) & "*'"
Case 4

Case 5

End Select

CadenaConsulta = "select * from " & NombreTabla & CadB & " " & Ordenacion
PonerCadenaBusqueda
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla" & NombreTabla, vbInformation
    Screen.MousePointer = vbDefault
    TotalReg = 0
    Exit Sub

    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        TotalReg = Data1.Recordset.RecordCount
        NumRegistro = 1
        PonerCampos
End If

Data1.ConnectionString = Conn
Data1.RecordSource = CadenaConsulta
Data1.Refresh
TotalReg = Data1.Recordset.RecordCount
Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim i As Integer
Dim Rectifica As Boolean
    For i = 0 To Text1.Count - 1
        Text1(i).Text = Data1.Recordset.Fields(i)
    Next i
    If mConfig.Ariadna Then
        Check1.Value = Abs(Data1.Recordset!nominas)
    End If
    'Ponemos la empresa
    Text2(0).Text = DevuelveNombreEmpresa(Data1.Recordset.Fields(2))
    'Nombre del horario
    Text2(1).Text = DevuelveNombreHorario(Data1.Recordset.Fields(3))
    
    'Ponemos los rectificados si tiene
    Combo1.ListIndex = Data1.Recordset!RedondearCadaTicaje
    Rectifica = Combo1.ListIndex = 0

        PonDBGridModificados Data1.Recordset.Fields(0)

    Label2.Caption = NumRegistro & " de " & TotalReg
End Sub

Private Sub PonerModo(Kmodo As Integer)
Dim i As Integer
Dim b As Boolean

If Modo = 1 Then
    For i = 0 To Text1.Count - 1
        Text1(i).BackColor = vbWhite
    Next i
End If
Modo = Kmodo
DespalzamientoVisible (Kmodo = 2)
cmdAceptar.Visible = (Kmodo >= 3)
cmdCancelar.Visible = (Kmodo >= 3)
Toolbar1.Buttons(6).Enabled = (Kmodo < 3)
Toolbar1.Buttons(7).Enabled = (Kmodo = 2)
Toolbar1.Buttons(8).Enabled = (Kmodo = 2)
Toolbar1.Buttons(1).Enabled = (Kmodo < 3)
Toolbar1.Buttons(2).Enabled = (Kmodo < 3)
If Kmodo = 0 Then _
    Label2.Caption = ""
b = (Modo = 2) Or Modo = 0
For i = 0 To Text1.Count - 1
    Text1(i).Locked = b
Next i
Check1.Enabled = Not b
Image2(0).Visible = Kmodo > 2
Image2(1).Visible = Kmodo > 2
'Rectificaciones de horarios
b = Kmodo = 4
Command1(0).Enabled = b
Command1(1).Enabled = b
Command1(2).Enabled = b
Combo1.Enabled = b
End Sub


Private Function DatosOk() As Boolean
Dim RS As ADODB.Recordset
Dim Cad As String
Dim i As Integer

DatosOk = False
'Haremos las comprobaciones necesarias de los campos
'Cad = ComprobarCampos
'If Cad <> "" Then
'    MsgBox Cad, vbExclamation
'    Exit Function
'End If


For i = 0 To Text1.Count - 1
    If Not CmpCam(Text1(i).Tag, Text1(i).Text) Then Exit Function
Next i
'Llegados a este punto los datos son correctos en valores
'Ahora comprobaremos otras cosas
'Este apartado dependera del formulario y la tabla
Cad = "Select * from " & NombreTabla
Cad = Cad & " WHERE idSeccion=" & Text1(0).Text

Set RS = New ADODB.Recordset
RS.Open Cad, Conn, , , adCmdText
If Not RS.EOF Then
    MsgBox "Ya existe un registro con ese código.", vbExclamation
    RS.Close
    Exit Function
End If
RS.Close
'Al final todo esta correcto
DatosOk = True
End Function


Private Sub SugerirCodigoSiguiente()
Dim Cad
Dim RS
'Sugeriremos el codigo siguiente.
'Obviamente depende en TOTAL medida de que tabla estemos trabajando
Cad = "Select Max(IdSeccion) from " & NombreTabla
Text1(0).Text = 1
Set RS = New ADODB.Recordset
RS.Open Cad, Conn, , , adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then
        Text1(0).Text = RS.Fields(0) + 1
    End If
End If
RS.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index > 5 And Button.Index < 9 Then
        If vUsu.Nivel > 1 Then 'solo admon
            MsgBox "No tiene autorizacion para realizar cambios", vbExclamation
            Exit Sub
        End If
    End If
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
'Case 20
'    'Listado en crystal report
'    Screen.MousePointer = vbHourglass
'    CR1.Connect = Conn
'    CR1.ReportFileName = App.Path & "\Informes\list_Inc.rpt"
'    CR1.WindowTitle = "Listado incidencias."
'    CR1.WindowState = crptMaximized
'    CR1.Action = 1
'    Screen.MousePointer = vbDefault

Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
Dim i
For i = 14 To 17
    Toolbar1.Buttons(i).Visible = bol
Next i
End Sub


Private Sub PonDBGridModificados(Seccion As Integer)
Dim RS As ADODB.Recordset
Dim Cad As String
Dim itmX As ListItem


On Error GoTo EPonDBGridModificados
Cad = "Select * from ModificarFichajes"
Cad = Cad & " WHERE idSeccion=" & Seccion
Cad = Cad & " ORDER BY Inicio"
ListView1.ListItems.Clear
Set RS = New ADODB.Recordset
RS.Open Cad, Conn, , , adCmdText
While Not RS.EOF
    Set itmX = ListView1.ListItems.Add
    itmX.Text = Format(RS!Inicio, "hh:mm")
    itmX.SubItems(1) = Format(RS!Fin, "hh:mm")
    itmX.SubItems(2) = Format(RS!modificada, "hh:mm")
    itmX.Tag = RS!Id
    itmX.SmallIcon = 10
    RS.MoveNext
Wend
RS.Close
Set RS = Nothing
Exit Sub
EPonDBGridModificados:
    MuestraError Err.Number, "Marcajes modificados."
    Set RS = Nothing
End Sub
