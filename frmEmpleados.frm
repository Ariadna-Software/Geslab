VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmEmpleados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empleados"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11355
   Icon            =   "frmEmpleados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6915
   ScaleMode       =   0  'User
   ScaleWidth      =   12242.58
   Begin VB.Frame FrameBanco 
      Height          =   735
      Left            =   5460
      TabIndex        =   73
      Top             =   5280
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   4380
         MaxLength       =   10
         TabIndex        =   30
         Tag             =   "#|N|S|0||"
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   3300
         MaxLength       =   4
         TabIndex        =   29
         Tag             =   "#|N|S|0|100|"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   2040
         TabIndex        =   28
         Tag             =   "#|N|S|0||"
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   30
         Left            =   720
         TabIndex        =   27
         Tag             =   "#|N|S|0||"
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Cuenta"
         Height          =   195
         Index           =   29
         Left            =   3840
         TabIndex        =   77
         Top             =   285
         Width           =   510
      End
      Begin VB.Label Label1 
         Caption         =   "D.C."
         Height          =   195
         Index           =   28
         Left            =   2880
         TabIndex        =   76
         Top             =   285
         Width           =   315
      End
      Begin VB.Label Label1 
         Caption         =   "Oficina"
         Height          =   195
         Index           =   27
         Left            =   1500
         TabIndex        =   75
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Entidad"
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   74
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pago banco"
      Height          =   255
      Left            =   4020
      TabIndex        =   72
      Top             =   5520
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   26
      Top             =   5460
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   8460
      TabIndex        =   69
      Tag             =   "#|N|S|||"
      Text            =   "Text1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   6660
      TabIndex        =   11
      Tag             =   "#|N|N|||"
      Text            =   "Text1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   4980
      TabIndex        =   10
      Tag             =   "#|N|N|0|100|"
      Text            =   "Text1"
      Top             =   2640
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3120
      Width           =   2295
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   7860
      Top             =   300
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowProgressCtls=   0   'False
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      TabIndex        =   64
      Text            =   "Text2"
      Top             =   3540
      Width           =   2175
   End
   Begin VB.Frame FrameVacas 
      Caption         =   "Vaciones"
      Height          =   795
      Left            =   9780
      TabIndex        =   61
      Top             =   2100
      Visible         =   0   'False
      Width           =   6075
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   4380
         TabIndex        =   24
         Tag             =   "#|F|S|||"
         Text            =   "Text1"
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1560
         TabIndex        =   23
         Tag             =   "#|F|S|||"
         Text            =   "Text1"
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha final"
         Height          =   195
         Index           =   15
         Left            =   3120
         TabIndex        =   63
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha inico"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   62
         Top             =   360
         Width           =   825
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   1260
         Picture         =   "frmEmpleados.frx":030A
         Top             =   300
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   4
         Left            =   4020
         Picture         =   "frmEmpleados.frx":040C
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Teléfono"
      Height          =   615
      Left            =   60
      TabIndex        =   58
      Top             =   4080
      Width           =   6135
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   3900
         TabIndex        =   20
         Tag             =   "#|T|S|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   19
         Tag             =   "#|T|S|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Móvil"
         Height          =   195
         Index           =   11
         Left            =   3120
         TabIndex        =   60
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Tfno:"
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   59
         Top             =   300
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   57
      Top             =   0
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
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
            ImageIndex      =   6
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
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Informe trabajdores"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   9300
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   3060
      Width           =   2595
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3075
      Width           =   2835
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   2340
      TabIndex        =   2
      Tag             =   "Empresa|T|N|||"
      Top             =   900
      Width           =   3555
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   1380
      TabIndex        =   3
      Tag             =   "#|T|N|||"
      Text            =   "Text1"
      Top             =   1305
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   8160
      TabIndex        =   18
      Tag             =   "#|N|N|0|99|"
      Text            =   "Text1"
      Top             =   3540
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   3960
      TabIndex        =   16
      Tag             =   "#|T|S|||"
      Text            =   "Text1"
      Top             =   3540
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   7620
      TabIndex        =   5
      Tag             =   "#|T|S|||"
      Text            =   "Text1"
      Top             =   1260
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   5460
      TabIndex        =   4
      Tag             =   "#|T|S|||"
      Text            =   "Text1"
      Top             =   1260
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   21
      Left            =   1380
      TabIndex        =   25
      Tag             =   "#|T|S|||"
      Text            =   "PATH"
      Top             =   4860
      Width           =   9795
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   7620
      TabIndex        =   21
      Tag             =   "#|F|S|||"
      Text            =   "Text1"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   10080
      TabIndex        =   22
      Tag             =   "#|F|S|||"
      Text            =   "Text1"
      Top             =   4260
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   6300
      TabIndex        =   17
      Tag             =   "#|N|N|0|99|"
      Text            =   "Text1"
      Top             =   3540
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   3900
      TabIndex        =   8
      Tag             =   "#|T|S|||"
      Text            =   "Text1"
      Top             =   2160
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1380
      TabIndex        =   7
      Tag             =   "#|T|S|||"
      Text            =   "Text1"
      Top             =   2235
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1380
      TabIndex        =   9
      Tag             =   "#|T|S|||"
      Text            =   "Text1"
      Top             =   2685
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Height          =   2235
      Left            =   9360
      TabIndex        =   41
      Top             =   480
      Width           =   1935
      Begin VB.Image Image1 
         Height          =   1755
         Left            =   180
         Top             =   240
         Width           =   1575
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   8220
      Top             =   240
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
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
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1380
      TabIndex        =   6
      Tag             =   "#|T|S|||"
      Text            =   "Text1"
      Top             =   1770
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   4
      Left            =   8220
      TabIndex        =   39
      Tag             =   "Horario|T|S|||"
      Text            =   "HORARIO"
      Top             =   540
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   3
      Left            =   9120
      TabIndex        =   38
      Tag             =   "Categoría|T|S|||"
      Text            =   "CATEGORIA"
      Top             =   600
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1380
      TabIndex        =   15
      Tag             =   "#|T|S|||"
      Text            =   "Text1"
      Top             =   3600
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   1380
      TabIndex        =   1
      Tag             =   "Empresa|N|N|||"
      Text            =   "EMPRESAA"
      Top             =   900
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1380
      TabIndex        =   0
      Tag             =   "#|N|N|||"
      Text            =   "Text1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   60
      TabIndex        =   34
      Top             =   6240
      Width           =   3615
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   240
         Width           =   3195
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8820
      TabIndex        =   32
      Top             =   6420
      Width           =   1155
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10080
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   6420
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7500
      TabIndex        =   31
      Top             =   6420
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   9480
      Top             =   960
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
            Picture         =   "frmEmpleados.frx":050E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":0620
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":0732
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":0844
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":0956
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":0A68
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":1342
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":1C1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":24F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados.frx":2DD0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text1 
      Height          =   345
      Index           =   22
      Left            =   8220
      TabIndex        =   66
      Tag             =   "Seccion|N|S|||"
      Text            =   "SECCION"
      Top             =   960
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo contrato"
      Height          =   195
      Left            =   120
      TabIndex        =   71
      Top             =   5520
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "% I.R.P.F."
      Height          =   195
      Index           =   26
      Left            =   7500
      TabIndex        =   70
      Top             =   2700
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "% S.S"
      Height          =   195
      Index           =   24
      Left            =   6180
      TabIndex        =   68
      Top             =   2700
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "% Antiguedad"
      Height          =   195
      Index           =   23
      Left            =   3900
      TabIndex        =   67
      Top             =   2685
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Sección"
      Height          =   195
      Left            =   120
      TabIndex        =   65
      Top             =   3180
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   6
      Left            =   7920
      Picture         =   "frmEmpleados.frx":2EE2
      Top             =   3600
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   5
      Left            =   900
      Picture         =   "frmEmpleados.frx":2FE4
      Top             =   4920
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   2
      Left            =   9780
      Picture         =   "frmEmpleados.frx":30E6
      Top             =   4260
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   7380
      Picture         =   "frmEmpleados.frx":31E8
      Top             =   4260
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmEmpleados.frx":32EA
      Top             =   900
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Imagen"
      Height          =   195
      Index           =   21
      Left            =   120
      TabIndex        =   56
      Top             =   4920
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre"
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
      Index           =   5
      Left            =   120
      TabIndex        =   55
      Top             =   1305
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Horario"
      Height          =   195
      Left            =   4080
      TabIndex        =   54
      Top             =   3150
      Width           =   510
   End
   Begin VB.Label Label5 
      Caption         =   "Categoría"
      Height          =   195
      Left            =   7740
      TabIndex        =   53
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      Height          =   195
      Index           =   22
      Left            =   120
      TabIndex        =   52
      Top             =   900
      Width           =   615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inc. continu."
      Height          =   195
      Index           =   17
      Left            =   7020
      TabIndex        =   51
      Top             =   3600
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Matrícula"
      Height          =   195
      Index           =   19
      Left            =   3180
      TabIndex        =   50
      Top             =   3600
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Nº  S.S."
      Height          =   195
      Index           =   18
      Left            =   7020
      TabIndex        =   49
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "D.N.I."
      Height          =   195
      Index           =   20
      Left            =   5040
      TabIndex        =   48
      Top             =   1320
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "F. Antiguedad"
      Height          =   195
      Index           =   12
      Left            =   6300
      TabIndex        =   47
      Top             =   4320
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "F. Baja"
      Height          =   195
      Index           =   13
      Left            =   9180
      TabIndex        =   46
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Control"
      Height          =   195
      Index           =   16
      Left            =   5700
      TabIndex        =   45
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   195
      Index           =   7
      Left            =   2880
      TabIndex        =   44
      Top             =   2205
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "C. Postal"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   43
      Top             =   2220
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   42
      Top             =   2670
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Dirección"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   40
      Top             =   1755
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Tarjeta"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   37
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Identificador"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   36
      Top             =   480
      Width           =   870
   End
End
Attribute VB_Name = "frmEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
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
'Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private ConsultaBase As String
Private NumRegistro As Long
Private kCampo As Integer
Private TotalReg As Long
Private kPicture As Integer
Private Inci_o_Empresa As Boolean
Private AntEmpresa As Long


Private Sub cmdAceptar_Click()
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim i As Integer

Screen.MousePointer = vbHourglass
On Error GoTo Error1
If Modo = 3 Then
    If DatosOk Then
        
        'Los combo
        Text1(3).Text = Combo2.ItemData(Combo2.ListIndex)
        Text1(4).Text = Combo1.ItemData(Combo1.ListIndex)
        Text1(22).Text = Combo3.ItemData(Combo3.ListIndex)
        
        Set Rs = New ADODB.Recordset
        Rs.CursorType = adOpenKeyset
        Rs.LockType = adLockOptimistic
        Rs.Open "Trabajadores", Conn, , , adCmdTable
        Rs.AddNew
        Cad = ""
        'Ahora insertamos
        For i = 0 To Rs.Fields.Count - 1
            'Menor k 25 pq los canmpos de Costes no lo mosotramos
            If i < 25 Then
                If Text1(i).Text <> "" Then Rs.Fields(i) = Text1(i).Text
            Else
                'Stop
                Debug.Print i & "-" & Rs.Fields(i).Name & " : " & Text1(i).Text
                
            End If
        Next i
        '--------------------
        Rs.Update
        Rs.Close
        Data1.Refresh
        'MsgBox "                Registro insertado.             ", vbInformation
        PonerModo 0
    End If
    Else
    If Modo = 4 Then
        'Modificar
        'Controlamos los combos
        Text1(3).Text = Combo2.ItemData(Combo2.ListIndex)
        Text1(4).Text = Combo1.ItemData(Combo1.ListIndex)
        If Combo3.ListIndex < 0 Then
            MsgBox "Seleccione una sección.", vbExclamation
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
        Text1(22).Text = Combo3.ItemData(Combo3.ListIndex)
        ''Haremos las comprobaciones necesarias de los campos
        'Recordamos que el text(0) tiene el codigo y no lo puede cambiar
        For i = 1 To Text1.Count - 1
            If Not CmpCam(Text1(i).Tag, Text1(i).Text) Then _
                GoTo Error1
        Next i
        
        'Las tarjetas de empleados deben empezar con el digito parametrizado
        If Text1(2).Text <> "" Then
            If Val(Mid(Text1(2).Text, 1, 1)) <> mConfig.DigitoTrabajadores Then
                MsgBox "Las tarjetas de los trabajadores deben empezar con el digito: " & mConfig.DigitoTrabajadores, vbExclamation
                GoTo Error1
            End If
        End If
        'Comprobamos, si han puesto texto en la imagen, la
        ' existencia o no del archivo
        If Text1(21).Text <> "" Then
            Cad = Dir(Text1(21).Text)
            If Cad = "" Then
                MsgBox "El archivo facilitado como imagen " & vbCrLf & Text1(21).Text & _
                    " NO existe." & vbCrLf & "Facilite una ruta de archivo válida.", vbExclamation
                GoTo Error1
            End If
        End If
        'Ahora modificamos
        Cad = "Select * from Trabajadores"
        Cad = Cad & " WHERE IdTrabajador=" & Data1.Recordset.Fields(0)
        Set Rs = New ADODB.Recordset
        Rs.CursorType = adOpenKeyset
        Rs.LockType = adLockOptimistic
        Rs.Open Cad, Conn, , , adCmdText
        'Almacenamos para luego buscarlo
        Cad = Rs!idTrabajador
        'modificamos
        For i = 1 To 24
            If Text1(i).Text <> "" Then
                If i > 22 Then
                    Rs.Fields(i).Value = TransformaPuntosComas(Text1(i).Text)
                Else
                    Rs.Fields(i).Value = Text1(i).Text
                End If
                Else
                    Rs.Fields(i).Value = Null
            End If
        Next i
        Rs.Update
        Rs.Requery
        Rs.Close
        'MsgBox "El registro ha sido modificado", vbInformation
        PonerModo 2
        'Hay que refresca el DAta1
        Data1.Recordset.Requery
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
    If Err.Number <> 0 Then MsgBox Cad & Err.Number & " - " & Err.Description & "    ::::" & i, vbExclamation
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
    'para k limpiemos las seccones
    Combo3.Clear
    AntEmpresa = -1
    'Añadiremos el boton de aceptar y demas objetos para insertar
    cmdAceptar.Caption = "Aceptar"
    PonerModo 3
    'Sugerimos el código
    SugerirCodigoSiguiente
    'Escondemos el navegador y ponemos insertando
    Label2.Caption = "Insertar"
    'Ofertamos a 0 el % de antiguedad y el de SS
    Text1(23).Text = 0
    Text1(24).Text = 0
    
    Check1.Value = 0
    'Borramos la foto
    Image1.Picture = LoadPicture("")
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
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            Text1(kCampo).SetFocus
            Else
                Text1(kCampo).BackColor = vbWhite
        End If
End If
End Sub

Private Sub BotonVerTodos()
'Ver todos
LimpiarCampos
PonerModo 2
CadenaConsulta = ConsultaBase & Ordenacion
PonerCadenaBusqueda
End Sub

Private Sub Desplazamiento(Index As Integer)
On Error Resume Next
Screen.MousePointer = vbHourglass
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
Screen.MousePointer = vbDefault
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
'Ponemos el foco sobre el nombre
Text1(5).SetFocus
End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim i

'Ciertas comprobaciones
If Data1.Recordset.RecordCount = 0 Then Exit Sub
'Pregunta
Cad = "Seguro que desea eliminar de la BD el registro:"
Cad = Cad & vbCrLf & "Cod: " & Data1.Recordset.Fields(0)
Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(5)
i = MsgBox(Cad, vbQuestion + vbYesNo)
If i = vbYes Then
    'Hay que eliminar
    On Error GoTo Error2
    Screen.MousePointer = vbHourglass
    'Borramos
    Cad = "Delete * from Trabajadores where IdTrabajador = " & Data1.Recordset.Fields(0)
    Conn.Execute Cad
    
    'Esperamos un tiempo prudencial de 1 seg
    i = Timer
    Do
    Loop Until Timer - i > 1
    
    Data1.Recordset.Update
    LimpiarCampos
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        PonerModo 0
        Else
            TotalReg = Data1.Recordset.RecordCount
            If NumRegistro >= TotalReg Then
                    Data1.Recordset.MoveLast
                    NumRegistro = TotalReg
                    Else
                        i = 1
                        While i < NumRegistro
                            Data1.Recordset.MoveNext
                            i = i + 1
                        Wend
                        NumRegistro = i
            End If
            Label2.Caption = NumRegistro & " de " & TotalReg
            Label2.Refresh
            PonerCampos
    End If
End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub



Private Sub Combo3_Click()
'Si, y solo si est insertando, sugeriremos el horario y el codigo de control
' siempre que hayan
If Modo <> 3 Then Exit Sub
If Combo3.ListIndex < 0 Then Exit Sub
    
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim i As Integer

    Set Rs = New ADODB.Recordset
    SQL = "SELECT * From Secciones " & _
        " WHERE IdSeccion =" & Combo3.ItemData(Combo3.ListIndex)
    Rs.Open SQL, Conn, , , adCmdText
    If Not Rs.EOF Then
            'REcorremos le combo de horario hasta situarnos donde queremos
            SQL = ""
            For i = 0 To Combo1.ListCount - 1
                If Combo1.ItemData(i) = Rs.Fields(3) Then
                    SQL = i
                    Exit For
                End If
            Next i
            Combo1.ListIndex = Val(SQL)
            'Control empleados
            Text1(16).Text = Rs.Fields(4)
    End If
    Rs.Close
    Set Rs = Nothing
End Sub

Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim i As Integer
Dim j As String
Screen.MousePointer = vbHourglass
LimpiarCampos
'Situamos el form
Left = 0
Top = 300
'Me.Height = 7740
'Me.Width = 12000
ConsultaBase = "Select Trabajadores.* , Empresas.NomEmpresa from Empresas,trabajadores where Trabajadores.IdEmpresa=empresas.IdEmpresa"
Ordenacion = " ORDER BY IdTrabajador"
'ASignamos un SQL al DATA1
Data1.ConnectionString = Conn
Data1.RecordSource = ConsultaBase & Ordenacion
Data1.Refresh
CargaCombos
PonerModo 0
'Ant empresa servira para no tener que cargar el como de secciones cada vez
AntEmpresa = -1
'ATencion, detalle
'Como cada text1(i) le corresponde un label1(i) desde i=0 hasta count-1
' y como ademas en los tag de los text1(i) tenemos las cadenas para la comprobacion
' y estas contienen el nombre del campo, que a vez es el del label(i) correspondiente
' entonces lo que hago es poner en el primer campo del tag
' una almohadilla que ahora sustuire por su label correspondiente
For i = 0 To Text1.Count - 1
    j = Mid(Text1(i).Tag, 1, 1)
    If j = "#" Then _
        Text1(i).Tag = Label1(i).Caption & Mid(Text1(i).Tag, 2)
Next i
End Sub



Private Sub LimpiarCampos()
Dim i
On Error Resume Next
For i = 0 To Text1.Count - 1
    Text1(i).Text = ""
Next i
Text2.Text = ""
End Sub



Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
If Inci_o_Empresa Then
    Text1(17).Text = vCodigo
    Text2.Text = vCadena
    Else
        Text1(25).Text = vCadena
        Text1(1).Text = vCodigo
        PonerComboSeccion
End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Text1(kPicture).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
Dim F As Date

Select Case Index
Case 0
    'Ponemos los valores para abrir
    Inci_o_Empresa = False
    Set frmB = New frmBusca
    frmB.Tabla = "Empresas"
    frmB.CampoBusqueda = "NomEmpresa"
    frmB.CampoCodigo = "IdEmpresa"
    frmB.TipoDatos = 3
    frmB.MostrarDeSalida = True
    frmB.Titulo = "EMPRESAS"
    frmB.Show vbModal
    Set frmB = Nothing
    
Case 1, 2, 3, 4
    'Las fechas
    kPicture = 11 + Index
    Set frmC = New frmCal
    If Text1(kPicture).Text <> "" Then
        F = CDate(Text1(kPicture).Text)
        Else
            F = Now
    End If
    frmC.Fecha = F
    frmC.Show vbModal
    Set frmC = Nothing

Case 5
    cd1.CancelError = False
    cd1.DialogTitle = "Seleccione la foto del trabajador"
    cd1.Filter = "Imágenes |*.bmp;*.jpg"
    cd1.ShowOpen
    If cd1.FileName <> "" Then
        If Dir(cd1.FileName) <> "" Then
            Text1(21).Text = cd1.FileName
            Image1.Picture = LoadPicture(cd1.FileName)
            Else
                MsgBox "No es una archivo válido."
                Text1(21).Text = ""
        End If
    End If
Case 6
    'Incidencias
    Inci_o_Empresa = True
    Set frmB = New frmBusca
    frmB.Tabla = "Incidencias"
    frmB.CampoBusqueda = "NomInci"
    frmB.CampoCodigo = "IdInci"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "INCIDENCIAS"
    frmB.Show vbModal
    Set frmB = Nothing
End Select
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
If Modo = 1 Then
    Text1(Index).BackColor = vbWhite
    Else
    If Modo > 2 Then
            Select Case Index
            Case 1
                'EMPRESA
                Cad = DevuelveNombreEmpresa(CLng(Val(Text1(1).Text)))
                If Cad = "" Then
                    Text1(1).Text = ""
                    Text1(25).Text = ""
                    Combo3.Clear
                    Else
                        Text1(25).Text = Cad
                        PonerComboSeccion
                End If
            Case 2
            
            End Select
    End If
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

If Text1(kCampo).Text = "" Then Exit Sub

'------------------------------------------------
'Prueba de pascual jajajaja
Dim c1 As String   'el nombre del campo
Dim Tipo As Long
Dim Operacion As String
Dim Valor As String
Dim aux1

c1 = Data1.Recordset.Fields(kCampo).Name
If c1 = "IdEmpresa" Then c1 = "Trabajadores.IdEmpresa"

Tipo = DevuelveTipo2(Data1.Recordset.Fields(kCampo).Type)
'Comprobacion numerica

'Devolvera uno de los tipos
'   1.- Numeros
'   2.- Booleanos
'   3.- Cadenas
'   4.- Fecha
'   0.- Error leyendo los tipos de datos
' segun sea uno u otro haremos una comparacion
If SeparaValorBusqueda(Text1(kCampo).Text, Operacion, Valor) = False Then Exit Sub

Select Case Tipo
Case 1
    If Operacion = "" Then Operacion = "="

    If Not IsNumeric(Valor) Then
        MsgBox "Debe de ser numérico.", vbExclamation
        Exit Sub
    End If
    CadB = c1 & " " & Operacion & " " & Valor
Case 2
    'Vemos si la cadena tiene un Falso o False
    If InStr(1, UCase(Text1(kCampo).Text), "F") Then
        aux1 = "False"
        Else
        aux1 = "True"
    End If
    CadB = c1 & " = " & aux1
Case 3
    CadB = c1 & " like '%" & Valor & "%'"
Case 4

Case 5

End Select
    


'-----------------------------------------------
'------------------------------------------------
'---         PRUEBA
'---
'Lo que esta contenido entre este encabezado y su fin
'es lo antiguo, que hay que modificar para cada campo
'''''''
''''''''Segun el campo haremos unas cosas u otras
'''''''Select Case Kcampo
'''''''    Case 0
'''''''        'Es la clave
'''''''        CadB = " WHERE codigo=" & Text1(Kcampo)
'''''''    Case 1
'''''''        'El nombre
'''''''        CadB = " WHERE Nombre like '%" & Text1(Kcampo) & "%'"
'''''''    Case 2
'''''''        'Direccion
'''''''        CadB = " WHERE Direccion like '%" & Text1(Kcampo) & "%'"
'''''''    Case 3
'''''''        'Poblacion
'''''''        CadB = " WHERE Poblacion like '%" & Text1(Kcampo) & "%'"
'''''''    Case 4
'''''''        'PROVINCIA
'''''''        CadB = " WHERE Provincia like '%" & Text1(Kcampo) & "%'"
'''''''    Case Else
'''''''        CadB = ""
'''''''End Select
'--------------  FIN
'------------------------------------------------

CadenaConsulta = ConsultaBase & " AND " & CadB & " " & Ordenacion
PonerCadenaBusqueda
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo Error4
Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.EOF Then
    MsgBox "No hay ningún registro en la tabla Trabajadores.", vbInformation
    Screen.MousePointer = vbDefault
    TotalReg = 0
    Exit Sub
    Label2.Caption = ""
    'PonerModo 0
    Else
        DespalzamientoVisible True
        PonerModo 2
        Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        TotalReg = Data1.Recordset.RecordCount
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
Dim Cad As String
Dim j As Integer

On Error GoTo Error10



For i = 0 To 25
   Text1(i).Text = DBLet(Data1.Recordset.Fields(i))
Next i
Text1(23).Text = TransformaComasPuntos(Text1(23).Text)
Text1(24).Text = TransformaComasPuntos(Text1(24).Text)
Label2.Caption = NumRegistro & " de " & TotalReg



'Cargamos la seccion
PonerComboSeccion
'Ponemos la incidencia continuada
Text2.Text = DevuelveTextoIncidencia(Data1.Recordset!incicont)
'Ponemos la imagen , si tiene
Cad = DBLet(Data1.Recordset.Fields(21))
If Cad <> "" And Dir(Cad) <> "" Then
        'SI k existe la imagen
        Image1.Picture = LoadPicture(Cad)
        Else
            Image1.Picture = LoadPicture("")
End If
'Ponemos los valores de los combos
j = 0
For i = 0 To Combo2.ListCount - 1
    If Combo2.ItemData(i) = Data1.Recordset.Fields(3) Then
        j = i
        Exit For
    End If
Next i
Combo2.ListIndex = j
'ahora el combo2
j = 0
For i = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(i) = Data1.Recordset.Fields(4) Then
        j = i
        Exit For
    End If
Next i
Combo1.ListIndex = j

'ahora el combo de SECCION
j = 0
For i = 0 To Combo3.ListCount - 1
    If Combo3.ItemData(i) = Data1.Recordset.Fields(22) Then
        j = i
        Exit For
    End If
Next i
Combo3.ListIndex = j

If LlevaGestionLaboral Then
    j = -1
    For i = 0 To Combo4.ListCount - 1
        If Combo4.ItemData(i) = DBLet(Data1.Recordset.Fields(26), "N") Then
            j = i
            Exit For
        End If
    Next i
Else
    j = -1
End If
Combo4.ListIndex = j

Exit Sub
Error10:
    MsgBox "Error: " & Err.Description
End Sub



'AGRUPAR PARA QUE NO HAGA TANTAS COMPARACIONES
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
HabilitarMenu
Label2.Visible = (Kmodo = 2)
b = (Kmodo = 2) Or Kmodo = 0
For i = 0 To Text1.Count - 1
    Text1(i).Locked = b
Next i
'Si no estamos buscando
b = Kmodo = 1
Text1(22).Locked = Not b
'Si estamos ins o modif
b = (Kmodo >= 3)
Text1(25).Enabled = Not b
For i = 0 To Image2.Count - 1
    Image2(i).Visible = b
    Combo1.Locked = Not b
    Combo2.Locked = Not b
    Combo3.Locked = Not b
Next i
End Sub


Private Function DatosOk() As Boolean
Dim Rs As ADODB.Recordset
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

'Coprobamos k el primer digito de la tarjeta es = al digito de trabajadores
If Text1(2).Text <> "" Then
    If Val(Mid(Text1(2).Text, 1, 1)) <> mConfig.DigitoTrabajadores Then
        MsgBox "Las tarjetas de los trabajadores deben empezar con el digito: " & mConfig.DigitoTrabajadores, vbExclamation
        Exit Function
    End If
End If



'Cuenta bancaria
If Check1.Value = 1 Then
    Cad = Text1(27).Text & Text1(28).Text & Text1(29).Text & Text1(30).Text
    If Cad = "" Then
        MsgBox "Deberia escribir la cuenta bancaria para el poder efectuar los pagos", vbExclamation
        Exit Function
    End If
End If

'Comprobamos que ha seleccionado algo en el combo
If Combo1.ListIndex < 0 Then
    MsgBox "Debe seleccionar un horario."
    Exit Function
End If
    
If Combo2.ListIndex < 0 Then
    MsgBox "Debe seleccionar una categoría."
    Exit Function
End If

'Comprobamos que ha seleccionado algo en el combo
If Combo3.ListIndex < 0 Then
    MsgBox "Debe seleccionar una sección."
    Exit Function
End If


'Si llevamos laboral entonces debemos poner un tipo de contrato
If LlevaGestionLaboral Then
    If Combo4.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de contrato", vbExclamation
        Exit Function
    End If
End If


'Llegados a este punto los datos son correctos en valores
'Ahora comprobaremos otras cosas
'                            =====================
'Este apartado dependera del formulario y la tabla
'                            =====================
Cad = "Select * from Trabajadores"
Cad = Cad & " WHERE idTrabajador=" & Text1(0).Text

Set Rs = New ADODB.Recordset
Rs.Open Cad, Conn, , , adCmdText
If Not Rs.EOF Then
    MsgBox "Ya existe un registro con ese código.", vbExclamation
    Rs.Close
    Exit Function
End If
Rs.Close
'Al final todo esta correcto
DatosOk = True
End Function


Private Sub SugerirCodigoSiguiente()
Dim Cad
Dim Rs
'Sugeriremos el codigo siguiente.
'Obviamente depende en TOTAL medida de que tabla estemos trabajando
Cad = "Select Max(IdTrabajador) from Trabajadores"
Text1(0).Text = 1
Set Rs = New ADODB.Recordset
Rs.Open Cad, Conn, , , adCmdText
If Not Rs.EOF Then
    If Not IsNull(Rs.Fields(0)) Then
        Text1(0).Text = Rs.Fields(0) + 1
    End If
End If
Rs.Close
End Sub


Private Sub CargaCombos()
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim i As Integer
'Horarios
Combo1.Clear
Cad = "Select IdHorario,NomHorario from Horarios order by nomhorario"
Set Rs = New ADODB.Recordset
Rs.Open Cad, Conn, , , adCmdText
i = 0
While Not Rs.EOF
    Combo1.AddItem Rs.Fields(1) '& " - " & rs.Fields(0)
    Combo1.ItemData(i) = Rs.Fields(0)
    i = i + 1
    Rs.MoveNext
Wend
Rs.Close


'Categorias
Combo2.Clear
Cad = "Select IdCategoria,NomCategoria from Categorias order by nomCategoria"
Set Rs = New ADODB.Recordset
Rs.Open Cad, Conn, , , adCmdText
i = 0
While Not Rs.EOF
    Combo2.AddItem Rs.Fields(1) '& " - " & rs.Fields(0)
    Combo2.ItemData(i) = Rs.Fields(0)
    i = i + 1
    Rs.MoveNext
Wend
Rs.Close


If LlevaGestionLaboral Then
    Cad = "Select IdContrato,DescContrato from tipocontrato"
    Rs.Open Cad, Conn, , , adCmdText
    i = 0
    Combo4.Clear
    While Not Rs.EOF
        Combo4.AddItem Rs.Fields(1) '& " - " & rs.Fields(0)
        Combo4.ItemData(i) = Rs.Fields(0)
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
End If


Set Rs = Nothing
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
Case 20
    Screen.MousePointer = vbHourglass
    CR1.Connect = Conn
    If MsgBox("Desea ordenarlo por tarjeta?", vbQuestion + vbYesNo) = vbYes Then
        CR1.ReportFileName = App.Path & "\Informes\list_Tra2.rpt"
    Else
        CR1.ReportFileName = App.Path & "\Informes\list_Tra.rpt"
    End If
    CR1.WindowTitle = "Listado trabajadores."
    CR1.WindowState = crptMaximized
    CR1.Action = 1
    Screen.MousePointer = vbDefault
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
Dim i
For i = 14 To 17
    Toolbar1.Buttons(i).Visible = bol
Next i
End Sub

Private Sub HabilitarMenu()
Dim b As Boolean
b = (Modo < 3)
'mnBuscar.Enabled = b
'mnVerTodos.Enabled = b
'mnNuevo.Enabled = b
'mnEliminar.Enabled = (Modo = 2)
'mnModificar.Enabled = (Modo = 2)
'mnSalir  siempre esta enabled
End Sub


'El combo seccion es distinto porque depende directamente de la empresa
Private Sub PonerComboSeccion()
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim i As Integer

On Error GoTo ErrorPonerComboSeccion
If Data1.Recordset.EOF Then
    Combo3.Clear
   ' Exit Sub
End If

If AntEmpresa <> Val(Text1(1).Text) Then
    AntEmpresa = Val(Text1(1).Text)
    'Las empresas son distintas y hay que cargar el combo
    'Horarios
    Combo3.Clear
    Cad = "Select IdSeccion,Nombre from Secciones where IdEmpresa=" & AntEmpresa & " order by Nombre"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, Conn, , , adCmdText
    i = 0
    While Not Rs.EOF
        Combo3.AddItem Rs.Fields(1) '& " - " & rs.Fields(0)
        Combo3.ItemData(i) = Rs.Fields(0)
        i = i + 1
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
End If

Exit Sub
ErrorPonerComboSeccion:
    MuestraError Err.Number
    Combo3.Clear
End Sub
