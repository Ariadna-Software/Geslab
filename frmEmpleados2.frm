VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEmpleados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Trabajadores"
   ClientHeight    =   6435
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   13860
   Icon            =   "frmEmpleados2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6435
   ScaleMode       =   0  'User
   ScaleWidth      =   14943.39
   Begin VB.CheckBox Check3 
      Caption         =   "Embargado"
      Enabled         =   0   'False
      Height          =   255
      Left            =   7440
      TabIndex        =   98
      Tag             =   "Embargado|N|S|||Trabajadores|Embargado|||"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Frame FrameHuella 
      Height          =   3195
      Left            =   11280
      TabIndex        =   92
      Top             =   3120
      Width           =   2055
      Begin VB.CommandButton cmdHuella 
         Caption         =   "Capturar"
         Height          =   435
         Left            =   240
         TabIndex        =   93
         ToolTipText     =   "Capturar huella"
         Top             =   2400
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   120
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   255
         Left            =   240
         TabIndex        =   94
         ToolTipText     =   "Calidad imagen"
         Top             =   2880
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         _Version        =   393216
         Min             =   1
         Max             =   100
         SelStart        =   70
         TickStyle       =   3
         Value           =   70
      End
      Begin VB.CheckBox Check2 
         Caption         =   "Sin huella"
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Tag             =   "Sin huella|N|S|||Trabajadores|NoCapturarHuella|||"
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label lblInfCodigo 
         Alignment       =   2  'Center
         Caption         =   "9999"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   855
         Left            =   120
         TabIndex        =   97
         Top             =   960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   135
         Index           =   0
         Left            =   1740
         TabIndex        =   0
         Top             =   2880
         Width           =   240
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   135
         Index           =   1
         Left            =   60
         TabIndex        =   96
         Top             =   2880
         Width           =   60
      End
      Begin VB.Image imgHuella 
         Height          =   1815
         Left            =   120
         Stretch         =   -1  'True
         Top             =   480
         Width           =   1695
      End
   End
   Begin MSAdodcLib.Adodc AdotmoImg 
      Height          =   375
      Left            =   10080
      Top             =   1440
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Caption         =   "Adodc2"
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   7440
      Top             =   5640
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   32
      Left            =   9855
      TabIndex        =   18
      Tag             =   "#|N|S|0||Trabajadores|bolsaNETO|||"
      Text            =   "Text1"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   31
      Left            =   7440
      TabIndex        =   17
      Tag             =   "#|N|S|0||Trabajadores|bolsabruto|||"
      Text            =   "Text1"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   25
      Left            =   10320
      TabIndex        =   86
      Tag             =   "sexo|N|S|0||Trabajadores|sexo|||"
      Text            =   "Text1"
      Top             =   480
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Mujer"
      Height          =   195
      Index           =   1
      Left            =   9360
      TabIndex        =   4
      Top             =   520
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Hombre"
      Height          =   195
      Index           =   0
      Left            =   8160
      TabIndex        =   3
      Top             =   520
      Width           =   975
   End
   Begin VB.ComboBox Combo5 
      Height          =   315
      ItemData        =   "frmEmpleados2.frx":030A
      Left            =   1440
      List            =   "frmEmpleados2.frx":031A
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Tag             =   "Control nómina|N|N|||Trabajadores|controlNomina|||"
      Top             =   2520
      Width           =   1995
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   22
      Left            =   6120
      TabIndex        =   28
      Tag             =   "#|F|S|||Trabajadores|antiguedad|||"
      Text            =   "Text1"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   4920
      TabIndex        =   16
      Tag             =   "#|N|S|0||Trabajadores|bolsahoras|||"
      Text            =   "Text1"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1440
      TabIndex        =   5
      Tag             =   "#|T|S|0||Trabajadores|idAsesoria||N|"
      Text            =   "Text1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.Frame FrameEmpresa 
      Caption         =   "Frame2"
      Height          =   795
      Left            =   3840
      TabIndex        =   78
      Top             =   5520
      Visible         =   0   'False
      Width           =   3615
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   1
         Left            =   1320
         TabIndex        =   79
         Tag             =   "#|N|N|0||Trabajadores|idEmpresa|||"
         Text            =   "EMPRESAA"
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Empresa"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   80
         Top             =   300
         Width           =   615
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   780
         Picture         =   "frmEmpleados2.frx":035C
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Frame FrameBanco 
      Height          =   735
      Left            =   5460
      TabIndex        =   76
      Top             =   5040
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   33
         Left            =   960
         TabIndex        =   35
         Tag             =   "Iban|T|S|||Trabajadores|iban|||"
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   3600
         MaxLength       =   10
         TabIndex        =   39
         Tag             =   "Cta banco|N|S|||Trabajadores|cuenta|0000000000||"
         Top             =   240
         Width           =   1275
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   3120
         MaxLength       =   2
         TabIndex        =   38
         Tag             =   "CC|T|S|||Trabajadores|controlcta|||"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   2400
         TabIndex        =   37
         Tag             =   "sucur|N|S|||Trabajadores|oficina|0000||"
         Top             =   240
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   30
         Left            =   1680
         TabIndex        =   36
         Tag             =   "Entidad|N|S|||Trabajadores|entidad|0000||"
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "IBAN"
         Height          =   195
         Index           =   30
         Left            =   120
         TabIndex        =   77
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Pago banco"
      Height          =   255
      Left            =   4020
      TabIndex        =   75
      Tag             =   "pago banco|N|S|||Trabajadores|pagobancario|||"
      Top             =   5280
      Width           =   1335
   End
   Begin VB.ComboBox Combo4 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   34
      Tag             =   "Tipo contrato|N|S|||Trabajadores|tipocontrato|||"
      Top             =   5220
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   26
      Left            =   5640
      TabIndex        =   14
      Tag             =   "#|N|S|0|100|Trabajadores|porcIRPF|||"
      Text            =   "Text1"
      Top             =   2055
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   24
      Left            =   3480
      TabIndex        =   13
      Tag             =   "#|N|S|0|100|Trabajadores|porcSS|||"
      Text            =   "Text1"
      Top             =   2040
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   23
      Left            =   1440
      TabIndex        =   12
      Tag             =   "#|N|S|0|100|Trabajadores|porcAntiguedad|||"
      Text            =   "Text1"
      Top             =   2055
      Width           =   735
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      ItemData        =   "frmEmpleados2.frx":045E
      Left            =   1440
      List            =   "frmEmpleados2.frx":0460
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Tag             =   "Seccion|N|N|||Trabajadores|Seccion|||"
      Top             =   3120
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   9000
      TabIndex        =   70
      Text            =   "Text2"
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Frame FrameVacas 
      Caption         =   "Vaciones"
      Height          =   795
      Left            =   120
      TabIndex        =   67
      Top             =   6480
      Visible         =   0   'False
      Width           =   6075
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   4380
         TabIndex        =   32
         Tag             =   "#|F|S|||Trabajadores|fecsalvac|||"
         Text            =   "Text1"
         Top             =   300
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   1560
         TabIndex        =   31
         Tag             =   "#|F|S|||Trabajadores|fecentvac|||"
         Text            =   "Text1"
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha final"
         Height          =   195
         Index           =   15
         Left            =   3120
         TabIndex        =   69
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha inico"
         Height          =   195
         Index           =   14
         Left            =   360
         TabIndex        =   68
         Top             =   360
         Width           =   825
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   3
         Left            =   1260
         Picture         =   "frmEmpleados2.frx":0462
         Top             =   300
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   4
         Left            =   4020
         Picture         =   "frmEmpleados2.frx":0564
         Top             =   300
         Width           =   240
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Teléfono"
      Height          =   615
      Left            =   60
      TabIndex        =   64
      Top             =   3960
      Width           =   4695
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   3360
         TabIndex        =   27
         Tag             =   "#|T|S|||Trabajadores|MovTrabajador|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   1320
         TabIndex        =   26
         Tag             =   "#|T|S|||Trabajadores|telTrabajador|||"
         Text            =   "Text1"
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label1 
         Caption         =   "Móvil"
         Height          =   195
         Index           =   11
         Left            =   2880
         TabIndex        =   66
         Top             =   300
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Tfno:"
         Height          =   195
         Index           =   10
         Left            =   840
         TabIndex        =   65
         Top             =   300
         Width           =   450
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   63
      Top             =   0
      Width           =   13860
      _ExtentX        =   24448
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
      Left            =   7080
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   8640
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Tag             =   "Categoria|N|N|||Trabajadores|idCategoria|||"
      Top             =   3060
      Width           =   2595
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Tag             =   "Horario|N|N|||Trabajadores|idHorario|||"
      Top             =   3075
      Width           =   2835
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   4380
      TabIndex        =   2
      Tag             =   "#|T|N|||Trabajadores|NomTrabajador|||"
      Text            =   "Text1"
      Top             =   465
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   17
      Left            =   8160
      TabIndex        =   25
      Tag             =   "#|N|N|0||Trabajadores|incicont|||"
      Text            =   "Text1"
      Top             =   3600
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   19
      Left            =   3960
      TabIndex        =   23
      Tag             =   "#|T|S|||Trabajadores|NumMat|||"
      Text            =   "Text1"
      Top             =   3600
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   18
      Left            =   4920
      TabIndex        =   7
      Tag             =   "#|T|S|||Trabajadores|numSS|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   1755
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   20
      Left            =   3120
      TabIndex        =   6
      Tag             =   "#|T|S|||Trabajadores|numDNI|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   21
      Left            =   1380
      TabIndex        =   33
      Tag             =   "#|T|S|||Trabajadores|imagen|||"
      Text            =   "PATH"
      Top             =   4680
      Visible         =   0   'False
      Width           =   9795
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   12
      Left            =   8220
      TabIndex        =   29
      Tag             =   "#|F|S|||Trabajadores|fecalta|||"
      Text            =   "Text1"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   13
      Left            =   10080
      TabIndex        =   30
      Tag             =   "#|F|S|||Trabajadores|fecbaja|||"
      Text            =   "Text1"
      Top             =   4140
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   16
      Left            =   6300
      TabIndex        =   24
      Tag             =   "#|N|N|0|100|Trabajadores|Control|||"
      Text            =   "Text1"
      Top             =   3600
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   3840
      TabIndex        =   10
      Tag             =   "#|T|S|||Trabajadores|PobTrabajador|||"
      Text            =   "Text1"
      Top             =   1455
      Width           =   2835
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   1440
      TabIndex        =   9
      Tag             =   "#|N|S|||Trabajadores|CodposTrabajador|||"
      Text            =   "Text1"
      Top             =   1455
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   7920
      TabIndex        =   11
      Tag             =   "#|T|S|||Trabajadores|ProvTrabajador|||"
      Text            =   "Text1"
      Top             =   1455
      Width           =   1935
   End
   Begin VB.Frame Frame4 
      Height          =   2715
      Left            =   11280
      TabIndex        =   48
      Top             =   360
      Width           =   2055
      Begin VB.CommandButton cmdDelete 
         Height          =   375
         Left            =   1440
         Picture         =   "frmEmpleados2.frx":0666
         Style           =   1  'Graphical
         TabIndex        =   91
         ToolTipText     =   "Abrir archivo existente"
         Top             =   2280
         Width           =   495
      End
      Begin VB.CommandButton cmdObtener 
         Height          =   375
         Left            =   780
         Picture         =   "frmEmpleados2.frx":0768
         Style           =   1  'Graphical
         TabIndex        =   90
         ToolTipText     =   "Abrir archivo existente"
         Top             =   2280
         Width           =   495
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   120
         Top             =   5280
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdCargar 
         Height          =   375
         Left            =   120
         Picture         =   "frmEmpleados2.frx":2A0A
         Style           =   1  'Graphical
         TabIndex        =   89
         ToolTipText     =   "Abrir archivo existente"
         Top             =   2280
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   1935
         Left            =   120
         Stretch         =   -1  'True
         Top             =   240
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   120
      Top             =   4200
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
      Left            =   7920
      TabIndex        =   8
      Tag             =   "#|T|S|||Trabajadores|DomTrabajador|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   3255
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1440
      TabIndex        =   22
      Tag             =   "#|T|S|||Trabajadores|NumTarjeta|||"
      Text            =   "Text1"
      Top             =   3600
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   1
      Tag             =   "#|N|N|0||Trabajadores|idTrabajador||S|"
      Text            =   "Text1"
      Top             =   480
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   60
      TabIndex        =   43
      Top             =   5640
      Width           =   3615
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   44
         Top             =   240
         Width           =   3195
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   41
      Top             =   5940
      Width           =   1155
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   5940
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7200
      TabIndex        =   40
      Top             =   5940
      Width           =   1215
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   10440
      Top             =   1200
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
            Picture         =   "frmEmpleados2.frx":2B0C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":2C1E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":2D30
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":2E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":2F54
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":3066
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":3940
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":421A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":4AF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados2.frx":53CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Imp. NETO"
      Height          =   195
      Index           =   32
      Left            =   8880
      TabIndex        =   88
      Top             =   2580
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Imp. BRUTO"
      Height          =   195
      Index           =   31
      Left            =   6465
      TabIndex        =   87
      Top             =   2580
      Width           =   915
   End
   Begin VB.Shape Shape1 
      Height          =   975
      Left            =   120
      Top             =   1920
      Width           =   11055
   End
   Begin VB.Label Label7 
      Caption         =   "Control nóminas"
      Height          =   195
      Left            =   240
      TabIndex        =   85
      Top             =   2520
      Width           =   1125
   End
   Begin VB.Label Label1 
      Caption         =   "F. Antiguedad"
      Height          =   195
      Index           =   22
      Left            =   4800
      TabIndex        =   84
      Top             =   4185
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   7
      Left            =   5880
      Picture         =   "frmEmpleados2.frx":54E0
      Top             =   4155
      Width           =   240
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Bolsa horas"
      Height          =   195
      Index           =   4
      Left            =   3840
      TabIndex        =   83
      Top             =   2580
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "% I.R.P.F."
      Height          =   195
      Index           =   26
      Left            =   4680
      TabIndex        =   82
      Top             =   2100
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Gestoria"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   81
      Top             =   1005
      Width           =   690
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo contrato"
      Height          =   195
      Left            =   240
      TabIndex        =   74
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "% S.S"
      Height          =   195
      Index           =   24
      Left            =   2760
      TabIndex        =   73
      Top             =   2100
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "% Antiguedad"
      Height          =   195
      Index           =   23
      Left            =   240
      TabIndex        =   72
      Top             =   2100
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "Sección"
      Height          =   195
      Left            =   240
      TabIndex        =   71
      Top             =   3180
      Width           =   915
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   6
      Left            =   7920
      Picture         =   "frmEmpleados2.frx":55E2
      Top             =   3660
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   5
      Left            =   900
      Picture         =   "frmEmpleados2.frx":56E4
      Top             =   4800
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   2
      Left            =   9780
      Picture         =   "frmEmpleados2.frx":57E6
      Top             =   4155
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   7980
      Picture         =   "frmEmpleados2.frx":58E8
      Top             =   4155
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Imagen"
      Height          =   195
      Index           =   21
      Left            =   240
      TabIndex        =   62
      Top             =   4800
      Visible         =   0   'False
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
      Left            =   3060
      TabIndex        =   61
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label6 
      Caption         =   "Horario"
      Height          =   195
      Left            =   4080
      TabIndex        =   60
      Top             =   3150
      Width           =   510
   End
   Begin VB.Label Label5 
      Caption         =   "Categoría"
      Height          =   195
      Left            =   7740
      TabIndex        =   59
      Top             =   3120
      Width           =   705
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Inc. continu."
      Height          =   195
      Index           =   17
      Left            =   7020
      TabIndex        =   58
      Top             =   3660
      Width           =   885
   End
   Begin VB.Label Label1 
      Caption         =   "Matrícula"
      Height          =   195
      Index           =   19
      Left            =   3180
      TabIndex        =   57
      Top             =   3660
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Nº  S.S."
      Height          =   195
      Index           =   18
      Left            =   4320
      TabIndex        =   56
      Top             =   1005
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "D.N.I."
      Height          =   195
      Index           =   20
      Left            =   2640
      TabIndex        =   55
      Top             =   1005
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "F. Alta"
      Height          =   195
      Index           =   12
      Left            =   7440
      TabIndex        =   54
      Top             =   4185
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Baja"
      Height          =   195
      Index           =   13
      Left            =   9360
      TabIndex        =   53
      Top             =   4185
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Control"
      Height          =   195
      Index           =   16
      Left            =   5700
      TabIndex        =   52
      Top             =   3660
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Población"
      Height          =   195
      Index           =   7
      Left            =   2820
      TabIndex        =   51
      Top             =   1500
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "C. Postal"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   50
      Top             =   1500
      Width           =   630
   End
   Begin VB.Label Label1 
      Caption         =   "Provincia"
      Height          =   195
      Index           =   8
      Left            =   7020
      TabIndex        =   49
      Top             =   1500
      Width           =   660
   End
   Begin VB.Label Label1 
      Caption         =   "Dirección"
      Height          =   195
      Index           =   6
      Left            =   7080
      TabIndex        =   47
      Top             =   1005
      Width           =   675
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Tarjeta"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   46
      Top             =   3600
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "Identificador"
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   45
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


'-----------------------------------------
'BD para las imagenes de los trabajadores
Private ConnImg As Connection
Private AbiertoImagen As Boolean

Private Function AbriConexionImagen() As Boolean

    On Error GoTo EConnImg
    
    AbriConexionImagen = False
    Set ConnImg = New ADODB.Connection

    
        
    ConnImg.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\imagen.mdb;Persist Security Info=False"
    ConnImg.CursorLocation = adUseServer
    ConnImg.Open
    AbriConexionImagen = True
    AbiertoImagen = True
    Exit Function
EConnImg:
    MuestraError Err.Number, "Abri conexion BD imagenes" & Err.Description
    
End Function


Private Sub Check2_Click()
    If Modo = 4 Then
        'OK se esta modificando
        If Text1(2).Text = "" Then
            lblInfCodigo.Caption = ""
        Else
            PintaCodigoTrabajadorSinHuella "&H" & Text1(2).Text
        End If
        Me.lblInfCodigo.Visible = Me.Check2.Value = 1
        imgHuella.Visible = Me.Check2.Value = 0
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim RS As ADODB.Recordset
Dim Cad As String
Dim AntigoUsuario As String

Screen.MousePointer = vbHourglass
On Error GoTo Error1
If Modo = 3 Then
    If DatosOk Then
        
        If InsertarDesdeForm(Me) Then
        
                
            'Imagen trabajador
            TratarImagen
            GrabarUsuarioGestorHuella ""
            
            'Updateo en alzicoop si procede
            If Not mConfig.TCP3_ Then UpdateaEuroges 0
            
            
            espera 0.2
            Data1.Refresh
            'MsgBox "                Registro insertado.             ", vbInformation
            PonerModo 0
            
            
        End If
    End If
Else
    If Modo = 4 Then
        ' existencia o no del archivo
        If Text1(21).Text <> "" Then
            Cad = Dir(Text1(21).Text)
            If Cad = "" Then
                MsgBox "El archivo facilitado como imagen " & vbCrLf & Text1(21).Text & _
                    " NO existe." & vbCrLf & "Facilite una ruta de archivo válida.", vbExclamation
                GoTo Error1
            End If
        End If
        
        If MiEmpresa.QueEmpresa <> 2 Then
            If Text1(3).Text = "" Then
                'MsgBox "Campo asesoria en blanco", vbExclamation
                'GoTo Error1
            End If
        End If
        
        If MiEmpresa.QueEmpresa > 0 Then
            If Me.Check2.Value = 1 Then
                
                If Not EsCorrectoElCodigoTarjeta Then
                    
                    GoTo Error1
                End If
                
            End If
        Else
            
                If Not IsNumeric(Text1(2).Text) Then
                    MsgBox "ID Tarjeta", vbExclamation
                    GoTo Error1
                    Text1(2).Text = ""
                End If
            
        End If
            'Sexo
            If Option1(0).Value Then
                Text1(25).Text = 0
            Else
                Text1(25).Text = 1
            End If
        
            
            AntigoUsuario = Trim(DBLet(Me.Data1.Recordset!Numtarjeta, "T"))
            
            'Ahora modificamos
            If ModificaDesdeFormulario(Me) = False Then Exit Sub
            'TratarImagen
            If AbiertoImagen Then TratarImagen
            
            
            'Updateo en alzicoop si procede
            If Not mConfig.TCP3_ Then UpdateaEuroges 1
            
            
            
            'SE  puede porgramar trabajadores SIN huella
            GrabarUsuarioGestorHuella AntigoUsuario
            
            
            PonerModo 2
            'Hay que refresca el DAta1
            Data1.Recordset.Requery
            Data1.Refresh
            
            
            
            
            'Hay que volver a poner el registro donde toca
            Data1.Recordset.MoveFirst
            
            Data1.Recordset.Find ("idtrabajador =" & Text1(0).Text)
        
            If Data1.Recordset.EOF Then
                LimpiarCampos
                PonerModo 0
            Else
                Label2.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
            End If
        
    End If
End If
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Cad & Err.Number & " - " & Err.Description, vbExclamation
End Sub


Private Sub TratarImagen()

    'NO SE HA TOCADO NADA
    If Image1.Tag = "" Then Exit Sub
    
    If Image1.Picture.Width > 0 Then
        
        'Tiene imagen
        If Adodc1.Recordset.EOF Then
            'NUEVA
            Adodc1.Recordset.AddNew
            If Modo = 3 Then
                Adodc1.Recordset.Fields(0) = Text1(0).Text
            Else
                Adodc1.Recordset.Fields(0) = Data1.Recordset.Fields!idTrabajador
            End If
            
        End If
        GuardarBinary Adodc1.Recordset.Fields(1), Image1
        Adodc1.Recordset.Update
    Else
        'Eliminar la entrada en la BD
        If Adodc1.Recordset.EOF Then Exit Sub
        Dim C As String
        If Modo = 3 Then
            C = Text1(0).Text
        Else
            C = Data1.Recordset.Fields!idTrabajador
        End If
        C = "Delete from imagenes where idtrabajador = " & C
        ConnImg.Execute C
        Adodc1.Refresh
    End If
    
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
    'Escondemos el navegador y ponemos insertando
    Label2.Caption = "Insertar"
    'Ofertamos a 0 el % de antiguedad y el de SS
    Text1(23).Text = 0
    Text1(24).Text = 0
    Text1(26).Text = 0
    
    'La asesoria
    Text1(3).Text = 0
    
    'Y la incidencia continuada tb
    Text1(17).Text = 0
    'Poenmos el combo seccion a blanco
    Combo3.ListIndex = -1
    Check1.Value = 0
    If MiEmpresa.QueEmpresa <> 2 Then
        Combo4.ListIndex = 1     'Pongo eventual al trabajador
    Else
        Combo5.ListIndex = 0
    End If
    
    
    'Borramos la foto de culo.
    If AbiertoImagen Then EnlazaImagen "-1"
    Image1.Picture = LoadPicture("")
    Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
'Buscar
If Modo <> 1 Then
    LimpiarCampos
    Combo1.ListIndex = -1
    Combo2.ListIndex = -1
    Combo3.ListIndex = -1
    Combo4.ListIndex = -1
    Combo5.ListIndex = -1
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


'Si NO tiene huella, la unica forma de ponerle la huella es "CAPTURANDO"
If DBLet(Data1.Recordset!NoCapturarHuella, "N") = 1 Then Check2.Enabled = False

'Ponemos el foco sobre el nombre
Text1(3).SetFocus
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
    conn.Execute Cad
    
    'Si esta enlazado con imagenes
    If AbiertoImagen Then
        Cad = "Delete * from imagenes where IdTrabajador = " & Data1.Recordset.Fields(0)
        ConnImg.Execute Cad
    End If

    'Si tiene ODBC
    If Not mConfig.TCP3_ Then UpdateaEuroges 2


    espera 0.5
    
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



Private Sub cmdCargar_Click()
    On Error GoTo EC
    
    cd1.ShowOpen
    If cd1.FileName <> "" Then
        Me.Image1.Picture = LoadPicture(cd1.FileName)
        Image1.Tag = "OK"
    End If
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub cmdDelete_Click()
    'Eliminar la imagen
    Me.Image1.Picture = LoadPicture("")
    Image1.Tag = "OK"
End Sub


Private Sub cmdHuella_Click() ' captura HUELLA digital
Dim Byt As Byte

    If Modo <> 2 Then
        If Modo > 2 Then
            MsgBox "Esta editando el usuario", vbExclamation
        Else
            MsgBox "Debe primero dar de alta el trabajador, antes de capturar su huella", vbExclamation
        End If
        Exit Sub
    End If
    
    If Val(Me.Slider1.Value) > 100 Or Val(Me.Slider1.Value) = 0 Then
        MsgBox "No deberia haber entrado aqui. Error valor slider", vbExclamation
        Exit Sub
    End If
    Byt = CByte(Me.Slider1.Value)
    
    Dim usu As UsuarioHuella
    Set usu = New UsuarioHuella
    
    If usu.Leer(Text1(2).Text) Then
    
    Else
        usu.CodUsuario = Text1(2)
        usu.GesLabID = Text1(0)
        usu.Mensaje = Left(Text1(5) & String(20, " "), 20)
    End If
    
    
    usu.CapturaHuella Byt
    If usu.FIR <> "" Then
        usu.Guardar
        If Dir(MiEmpresa.DirHuellas & "\" & usu.CodUsuario & ".jpg") <> "" Then
            imgHuella.Picture = LoadPicture(MiEmpresa.DirHuellas & "\" & usu.CodUsuario & ".jpg")
            imgHuella.Visible = True
        Else
            imgHuella.Visible = False
        End If
        
        
        conn.Execute "UPDATE trabajadores SET NoCapturarHuella =0 WHERE idtrabajador=" & Text1(0).Text
        Check2.Value = 0
        DoEvents
    End If
    Set usu = Nothing
End Sub



Private Sub cmdObtener_Click()
Dim C As String


    On Error GoTo EC
    'Abro el programa
    'Llamandolo con mi nombre de archivo como parametro
    C = App.Path & "\pr1.jpg"
    If Dir(C, vbArchive) <> "" Then Kill C
    
    C = """" & App.Path & "\aricam.exe" & """ """ & C & """"""
        
    Me.Caption = "Leyendo Webcam"
    frmPpal1.Hide
    DoEvents
    
    LanzaArchivoModificar C
    
    C = App.Path & "\pr1.jpg"
    If Dir(C, vbArchive) <> "" Then
        Me.Image1.Picture = LoadPicture(C)
        Image1.Tag = "OK"
    End If
    frmPpal1.Show
    Me.Caption = "Trabajadores"
    Exit Sub
EC:
    MuestraError Err.Number, Err.Description
End Sub
    
    
    
Private Sub LanzaArchivoModificar(CadenaShell As String)
Dim PID As Long


    
    
    PID = Shell(CadenaShell, vbNormalFocus)
    If PID <> 0 Then
        'Esperar a que finalice
        WaitForTerm PID
    End If

End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
Keypress KeyAscii
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
Keypress KeyAscii
End Sub

Private Sub Combo3_Click()
'Si, y solo si est insertando, sugeriremos el horario y el codigo de control
' siempre que hayan
If Modo <> 3 Then Exit Sub
If Combo3.ListIndex < 0 Then Exit Sub
    
Dim RS As ADODB.Recordset
Dim SQL As String
Dim i As Integer

    Set RS = New ADODB.Recordset
    SQL = "SELECT * From Secciones " & _
        " WHERE IdSeccion =" & Combo3.ItemData(Combo3.ListIndex)
    RS.Open SQL, conn, , , adCmdText
    If Not RS.EOF Then
            'REcorremos le combo de horario hasta situarnos donde queremos
            SQL = ""
            For i = 0 To Combo1.ListCount - 1
                If Combo1.ItemData(i) = RS.Fields(3) Then
                    SQL = i
                    Exit For
                End If
            Next i
            Combo1.ListIndex = Val(SQL)
            'Control empleados
            Text1(16).Text = RS.Fields(4)
    End If
    RS.Close
    Set RS = Nothing
End Sub


Private Sub Combo3_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Combo4_KeyPress(KeyAscii As Integer)
Keypress KeyAscii
End Sub

Private Sub Combo5_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

'Private Sub Command1_Click()
'Dim usu As UsuarioHuella
'    Set usu = New UsuarioHuella
'    If usu.Leer(Me.Text1(0).Text) Then
'        If usu.FIR <> "" Then usu.ImagenHuella MiEmpresa.DirHuellas & "\" & Husu.CodUsuario & ".jpg"
'    End If
'    Set usu = Nothing
'End Sub

Private Sub Command7_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()
Dim i As Integer
Dim J As String
Screen.MousePointer = vbHourglass
LimpiarCampos
Label2.Caption = ""
'Situamos el form
Left = 0
Top = 300
'Me.Height = 7740
'Me.Width = 12000
ConsultaBase = "Select * FROM Trabajadores"
Ordenacion = " ORDER BY IdTrabajador"


'Si la bolsa horas es tipo la de alzira
'mostraremos los campos de importebruto importe neto
For i = 31 To 32
    Label1(i).Visible = Not MiEmpresa.NominaAutomatica
    Text1(i).Visible = Not MiEmpresa.NominaAutomatica
Next


Check3.Visible = MiEmpresa.QueEmpresa = 0

If Not MiEmpresa.NominaAutomatica Then
    'Hago el cuadrado mas grande
    Me.Shape1.Width = Text1(32).Left + Text1(32).Width + 60
    
Else
    'mas pequeño
    Me.Shape1.Width = Label1(31).Left
End If
    
    
'ASignamos un SQL al DATA1
Data1.ConnectionString = conn
Data1.RecordSource = ConsultaBase & Ordenacion
Data1.Refresh
CargaCombos
PonerModo 0
'ATencion, detalle
'Como cada text1(i) le corresponde un label1(i) desde i=0 hasta count-1
' y como ademas en los tag de los text1(i) tenemos las cadenas para la comprobacion
' y estas contienen el nombre del campo, que a vez es el del label(i) correspondiente
' entonces lo que hago es poner en el primer campo del tag
' una almohadilla que ahora sustuire por su label correspondiente
AbiertoImagen = False
If Dir(App.Path & "\imagen.mdb") <> "" Then AbriConexionImagen
Frame4.Visible = AbiertoImagen

                    'Belgida y Alzira catadau
If MiEmpresa.QueEmpresa > 0 Then
    Me.Width = 13530
    
    
    'Si no esta la imegen, pondremos el de la huella
    If Not AbiertoImagen Then FrameHuella.Top = 360
    If MiEmpresa.DirHuellas = "" Then MsgBox "Falta configurar carpeta huellas", vbExclamation
    
    
Else
    FrameHuella.Visible = False
    If Not AbiertoImagen Then Me.Width = 11355
End If


If AbiertoImagen Then Me.cmdObtener.Enabled = (Dir(App.Path & "\aricam.exe") <> "")
    
    

Dim T As Object

For Each T In Me.Controls
    J = Mid(T.Tag, 1, 1)
    If J = "#" Then _
        T.Tag = Label1(T.Index).Caption & Mid(T.Tag, 2)
Next

End Sub



Private Sub LimpiarCampos()
Dim i
On Error Resume Next
Limpiar Me
Text2.Text = ""
Text1(1).Text = 1  'La empresa SIEMPRE, repito SIEMPRE sera 1
Me.Image1.Picture = LoadPicture("")

        'Alzira BELGIDA cata
If MiEmpresa.QueEmpresa > 0 Then Me.imgHuella.Picture = LoadPicture("")

'los combos
Me.Combo1.ListIndex = -1
Me.Combo2.ListIndex = -1
Me.Combo3.ListIndex = -1
Me.Combo4.ListIndex = -1
Me.Combo5.ListIndex = -1
Me.Check2.Value = 0
lblInfCodigo.Caption = ""
If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If AbiertoImagen Then ConnImg.Close
    Set ConnImg = Nothing
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
If Inci_o_Empresa Then
    Text1(17).Text = vCodigo
    Text2.Text = vCadena
End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(kPicture).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_DblClick()
    If Modo <> 2 Then Exit Sub
    
    If Image1.Picture Is Nothing Then Exit Sub
    If Image1.Picture = 0 Then Exit Sub
    
    
    On Error GoTo ECargandoImagenGrande
    
    With frmImagenGrande
        .Image1.Stretch = False
        
        .Image1.Picture = Image1.Picture
        If .Image1.Width > 5500 Or .Image1.Height > 4800 Then
            .Image1.Height = 4800
            .Image1.Width = 5500
            .Image1.Stretch = True
        End If
        
        .Caption = "Trabajador: " & Text1(0).Text
        .Label1.Caption = Text1(5).Text
        .Show vbModal
    End With
    
    
    Exit Sub
ECargandoImagenGrande:
    MuestraError Err.Number
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
    
Case 1, 2, 3, 4, 7
    'Las fechas
    If Index < 7 Then
        kPicture = 11 + Index
    Else
        kPicture = 15 + Index
    End If
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

Private Sub ImgDelete_Click()

End Sub

Private Sub Option1_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text1_GotFocus(Index As Integer)
kCampo = Index
If Modo = 1 Then
    Text1(Index).BackColor = vbYellow
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Text1(Index).BackColor = vbWhite
    If Modo = 1 Then
        If KeyAscii = 13 Then
            'Ha pulsado enter, luego tenemos que hacer la busqueda
            BotonBuscar
        End If
    Else
        If Modo = 3 Or Modo = 4 Then
            If KeyAscii = 13 Then
                KeyAscii = 0
                SendKeys "{tab}"
            End If
        End If
    End If
End Sub



Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String
Dim Aux As String

If Modo = 1 Then
    Text1(Index).BackColor = vbWhite
End If
If Modo > 2 Then
    If Index = 28 Then
        'Controlo el codigo de control
        If Text1(28).Text = "" Then Exit Sub
        If Not IsNumeric(Text1(28).Text) Then
            MsgBox "Campo debe ser numérico", vbExclamation
            Text1(28).Text = ""
            Text1(28).SetFocus
        End If
    Else
        If Index = 22 Or Index = 12 Or Index = 13 Then
            If Not EsFechaOK(Text1(Index)) Then Text1(Index).Text = ""
            
            
        ElseIf Index = 29 Then
                Text1(29).Text = Right(String(10, "0") & Me.Text1(29).Text, 10)
                'Formateamos
                Aux = Right("0000" & Me.Text1(30).Text, 4) & Right("0000" & Me.Text1(27).Text, 4)
                Aux = Aux & Me.Text1(28).Text & Text1(29).Text
                
                If Len(Aux) = 20 Then
                    'OK. Calculamos el IBAN
                    
                    
                    If Text1(33).Text = "" Then
                        'NO ha puesto IBAN
                        If DevuelveIBAN2("ES", Aux, Cad) Then Text1(33).Text = "ES" & Cad
                    Else
                        Cad = CStr(Mid(Text1(33).Text, 1, 2))
                        If DevuelveIBAN2(CStr(Cad), Aux, Aux) Then
                            If Mid(Text1(33).Text, 3) <> Aux Then
                                
                                MsgBox "Codigo IBAN distinto del calculado [" & Cad & Aux & "]", vbExclamation
                                'Text1(49).Text = "ES" & SQL
                            End If
                        End If
                    End If
                End If

        Else
            If Index = 27 Or Index = 30 Then
                'entidad sucur
                Text1(Index).Text = Right("0000" & Me.Text1(Index).Text, 4)
            End If
            
        End If
    End If
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

If Text1(kCampo).Text = "" Then
    If Combo3.ListIndex < 0 Then
        Exit Sub
    Else
        
        
        CadB = "seccion = " & Combo3.ItemData(Combo3.ListIndex)
        GoTo ponerlabusqueda
    End If
End If
'------------------------------------------------
'Prueba de pascual jajajaja
Dim C1 As String   'el nombre del campo
Dim Tipo As Long
Dim Operacion As String
Dim Valor As String
Dim aux1

C1 = Data1.Recordset.Fields(kCampo).Name
If C1 = "IdEmpresa" Then C1 = "Trabajadores.IdEmpresa"

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

    'If Not IsNumeric(Valor) Then
    '    MsgBox "Debe de ser numérico.", vbExclamation
    '    Exit Sub
    'End If
    CadB = C1 & " " & Operacion & " " & Valor
Case 2
    'Vemos si la cadena tiene un Falso o False
    If InStr(1, UCase(Text1(kCampo).Text), "F") Then
        aux1 = "False"
        Else
        aux1 = "True"
    End If
    CadB = C1 & " = " & aux1
Case 3
    CadB = C1 & " like '%" & Valor & "%'"
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

ponerlabusqueda:

'CadenaConsulta = ConsultaBase & " AND " & CadB & " " & Ordenacion
CadenaConsulta = ConsultaBase & " WHERE " & CadB & " " & Ordenacion
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
'        Data1.Recordset.MoveLast
'        Data1.Recordset.MoveFirst
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
Dim J As Integer

Dim Donde As String

On Error GoTo Error10



Donde = "Campos FORM"
PonerCamposForma Me, Data1

Label2.Caption = NumRegistro & " de " & TotalReg


'Ponemos la incidencia continuada
Text2.Text = DevuelveTextoIncidencia(Data1.Recordset!incicont)

'Ponemos los valores de los combos
J = 0
For i = 0 To Combo2.ListCount - 1
    If Combo2.ItemData(i) = Data1.Recordset.Fields(3) Then
        J = i
        Exit For
    End If
Next i
Combo2.ListIndex = J
'ahora el combo2
J = 0
For i = 0 To Combo1.ListCount - 1
    If Combo1.ItemData(i) = Data1.Recordset.Fields(4) Then
        J = i
        Exit For
    End If
Next i
Combo1.ListIndex = J

'ahora el combo de SECCION
J = 0
For i = 0 To Combo3.ListCount - 1
    If Combo3.ItemData(i) = Data1.Recordset.Fields(22) Then
        J = i
        Exit For
    End If
Next i
Combo3.ListIndex = J

If MiEmpresa.LlevaLaboral Then
    J = -1
    For i = 0 To Combo4.ListCount - 1
        If Combo4.ItemData(i) = DBLet(Data1.Recordset.Fields(26), "N") Then
            J = i
            Exit For
        End If
    Next i
Else
    J = -1
End If
Combo4.ListIndex = J

'Ahora pongo el sexo
    If Text1(25).Text = 1 Then
        Option1(0).Value = False
    Else
        Option1(0).Value = True
    End If
    Option1(1).Value = Not Option1(0).Value



'Formateo la cuenta del banco
If Text1(27).Text <> "" Then Text1(27).Text = Format(Text1(27).Text, "0000")
If Text1(29).Text <> "" Then Text1(29).Text = Format(Text1(29).Text, "0000000000")
If Text1(30).Text <> "" Then Text1(30).Text = Format(Text1(30).Text, "0000")
    


Donde = "Imagen"

Image1.Tag = ""

If AbiertoImagen Then
    cmdCargar.Tag = Label2.Caption
    Screen.MousePointer = vbHourglass
    Label2.Caption = "Leyendo imagen"
    Label2.Refresh
    EnlazaImagen Text1(0).Text
    Label2.Caption = cmdCargar.Tag
    cmdCargar.Tag = ""
    Screen.MousePointer = vbDefault
End If

'Si tienen que cargar la carga, si no pues res
Donde = "Huella"
If MiEmpresa.QueEmpresa <> 0 Then
    lblInfCodigo.Visible = Me.Check2.Value = 1
    Me.lblInfCodigo.Visible = lblInfCodigo.Visible
    If Me.Check2.Value = 1 Then
        imgHuella.Visible = False
        
        Cad = Trim(Text1(2).Text)
        If Cad <> "" Then Cad = "&H" & Cad
        PintaCodigoTrabajadorSinHuella Cad
    Else
        ImagenHuella
    End If
    lblInfCodigo.Visible = Me.Check2.Value = 1
    
End If


Exit Sub
Error10:
    Donde = Donde & vbCrLf & vbCrLf & Err.Description
    MsgBox Donde, vbExclamation
End Sub

Private Sub PintaCodigoTrabajadorSinHuella(ByRef C As String)
    On Error Resume Next
    If C = "" Then
        lblInfCodigo.Caption = ""
    Else
        C = CStr(CLng(C))
        If Val(C) > 9999 Then
            lblInfCodigo.Caption = "> max"
        Else
            lblInfCodigo.Caption = Format(C, "0000")
        End If
    End If
   
    If Err.Number <> 0 Then
        Err.Clear
        lblInfCodigo.Caption = "N/D"
    End If
End Sub

Private Sub EnlazaImagen(Codigo As String)
On Error GoTo EEnlazaImagen
    Adodc1.ConnectionString = ConnImg.ConnectionString
    Adodc1.RecordSource = "Select * from imagenes where idtrabajador =" & Codigo
    Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then
        LeerBinary Adodc1.Recordset.Fields(1), Me.Image1
    Else
        Me.Image1.Picture = LoadPicture("")
    End If
    Exit Sub
EEnlazaImagen:
    MuestraError Err.Number, "Mostrar imagen." & Err.Description
    Me.Image1.Picture = LoadPicture("") 'si da error vacio
End Sub

'AGRUPAR PARA QUE NO HAGA TANTAS COMPARACIONES
Private Sub PonerModo(Kmodo As Integer)
Dim i As Integer
Dim B As Boolean
Dim T As TextBox
If Modo = 1 Then
    For Each T In Text1
        T.BackColor = vbWhite
    Next
    
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
cmdCargar.Visible = False
If AbiertoImagen Then
    Me.cmdCargar.Visible = (Modo >= 3)
End If
cmdDelete.Visible = cmdCargar.Visible
Me.cmdObtener.Visible = cmdCargar.Visible
B = (Kmodo = 2) Or Kmodo = 0
For Each T In Text1
    T.Locked = B
Next

B = False                       'Belgida y Alzira cata
If MiEmpresa.QueEmpresa > 0 Then B = Modo = 2
Me.cmdHuella.Visible = B
Me.Slider1.Visible = B
Me.Label8(0).Visible = B
Me.Label8(1).Visible = B
Me.Check2.Enabled = Modo = 1 Or Modo > 2


If Check3.Visible Then Me.Check3.Enabled = Modo = 1 Or Modo > 2
    
    
'Si no estamos buscando
B = Kmodo = 1
'Si estamos ins o modif
    B = (Kmodo >= 3)
    For i = 0 To Image2.Count - 1
        Image2(i).Visible = B
    Next i
    Image2(5).Visible = False
    Combo1.Locked = Not B
    Combo2.Locked = Not B
'    Combo3.Locked = Not B
    Combo4.Locked = Not B
    Combo5.Locked = Not B
    Check1.Enabled = B
    'Sexo
    Option1(0).Enabled = B
    Option1(1).Enabled = B
End Sub


Private Function DatosOk() As Boolean
Dim RS As ADODB.Recordset
Dim Cad As String
Dim i As Integer

DatosOk = False


'Comprobacion a mano
'Pq al cambiar los datos estaran a blanco el campo de asesoria
'Pero no deberia estar
'If Text1(3).Text = "" Then
'    MsgBox "Campo asesoria en blanco", vbExclamation
'    Exit Function
'End If


If Not CompForm(Me) Then Exit Function

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
    i = Len(Cad)
    If i < 10 Then
        Cad = "Deberia escribir la cuenta bancaria para el poder efectuar los pagos." & vbCrLf & "     ¿Desea  continuar?"
        If MsgBox(Cad, vbExclamation + vbYesNoCancel) <> vbYes Then Exit Function
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
If MiEmpresa.LlevaLaboral Then
    If Combo4.ListIndex < 0 Then
        MsgBox "Seleccione un tipo de contrato", vbExclamation
        Exit Function
    End If
End If


'Sexo
If Option1(0).Value Then
    Text1(25).Text = 0
Else
    Text1(25).Text = 1
End If



If Modo = 3 Then
    If Check2.Value = 1 Then
        MsgBox "No se puede grabar directamente el trabajador con la opcion ""sin huella""", vbExclamation
        Check2.Value = 0
    End If
End If


        
        
'Llegados a este punto los datos son correctos en valores
'Ahora comprobaremos otras cosas
'                            =====================
'Este apartado dependera del formulario y la tabla
'                            =====================
Cad = "Select * from Trabajadores"
Cad = Cad & " WHERE idTrabajador=" & Text1(0).Text

Set RS = New ADODB.Recordset
RS.Open Cad, conn, , , adCmdText
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
Cad = "Select Max(IdTrabajador) from Trabajadores"
Text1(0).Text = 1
Set RS = New ADODB.Recordset
RS.Open Cad, conn, , , adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then
        Text1(0).Text = RS.Fields(0) + 1
    End If
End If
RS.Close
End Sub


Private Sub CargaCombos()
Dim RS As ADODB.Recordset
Dim Cad As String
Dim i As Integer
'Horarios
Combo1.Clear
Cad = "Select IdHorario,NomHorario from Horarios order by nomhorario"
Set RS = New ADODB.Recordset
RS.Open Cad, conn, , , adCmdText
i = 0
While Not RS.EOF
    Combo1.AddItem RS.Fields(1) '& " - " & rs.Fields(0)
    Combo1.ItemData(i) = RS.Fields(0)
    i = i + 1
    RS.MoveNext
Wend
RS.Close


'Categorias
Combo2.Clear
Cad = "Select IdCategoria,NomCategoria from Categorias order by nomCategoria"
Set RS = New ADODB.Recordset
RS.Open Cad, conn, , , adCmdText
i = 0
While Not RS.EOF
    Combo2.AddItem RS.Fields(1) '& " - " & rs.Fields(0)
    Combo2.ItemData(i) = RS.Fields(0)
    i = i + 1
    RS.MoveNext
Wend
RS.Close

    Combo3.Clear
    Cad = "Select IdSeccion,Nombre from Secciones where IdEmpresa=1 order by Nombre"
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, , , adCmdText
    i = 0
    While Not RS.EOF
        Combo3.AddItem RS.Fields(1) '& " - " & rs.Fields(0)
        Combo3.ItemData(i) = RS.Fields(0)
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close




If MiEmpresa.LlevaLaboral Then
    Cad = "Select IdContrato,DescContrato from tipocontrato"
    RS.Open Cad, conn, , , adCmdText
    i = 0
    Combo4.Clear
    While Not RS.EOF
        Combo4.AddItem RS.Fields(1) '& " - " & rs.Fields(0)
        Combo4.ItemData(i) = RS.Fields(0)
        i = i + 1
        RS.MoveNext
    Wend
    RS.Close
End If


Set RS = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index > 5 And Button.Index < 9 Then
        If vUsu.Nivel > 1 Then 'solo admon
            MsgBox "No tiene autorización para realizar cambios", vbExclamation
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
Case 20
    'Llamaremos a un FORM para que devuelva la opcion de imprimir tarjetas
    
    If AbiertoImagen Then
        'tiene IMAGEN
        VariableCompartida = ""
        frmVarios.Opcion = 1
        frmVarios.Show vbModal
        If VariableCompartida = "" Then Exit Sub
        NParam = RecuperaValor(VariableCompartida, 1)
        If Val(NParam) = 0 Then
            MsgBox "Error datos devueltos", vbExclamation
            Exit Sub
        End If
        
        If NParam = 1 Then
            'NORMAL
            nOpcion = Val(RecuperaValor(VariableCompartida, 2)) + 105    '105 + 0 o 1
        Else
            'TARJETAS
            'Volcamos las imagenes sobre la tabla tmpimg
            'CREATE TABLE tmpImag (idTrab  INTEGER, vImag IMAGE)
                 
            
            
            nOpcion = Val(RecuperaValor(VariableCompartida, 2))

                'Registro actual. NO EXISTE NINGUNO
            If Text1(0).Text = "" Or Text1(5).Text = "" Or Text1(16).Text = "" Then
                MsgBox "No existe ningun dato para mostrar", vbExclamation
                Exit Sub
            End If
            
            Screen.MousePointer = vbHourglass
            TransferirImagenesEntreAdos nOpcion = 0
            Screen.MousePointer = vbDefault
            nOpcion = 144
        End If
    Else
    
    
        If MsgBox("Desea ordenarlo por tarjeta?", vbQuestion + vbYesNo) = vbYes Then
            'CR1.ReportFileName = App.Path & "\Informes\list_Tra2.rpt"
            nOpcion = 106
        Else
            'CR1.ReportFileName = App.Path & "\Informes\list_Tra.rpt"
            nOpcion = 105
        End If
        
    End If
        
        
        
        
    With frmImprimir
        .Opcion = nOpcion
        .NumeroParametros = 0
        .OtrosParametros = ""
        .FormulaSeleccion = ""
        .Show vbModal
    End With
    
    
    If nOpcion = 144 Then
        conn.Execute "DELETE FROM tmpImag"
    End If
    
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
Dim B As Boolean
B = (Modo < 3)
'mnBuscar.Enabled = b
'mnVerTodos.Enabled = b
'mnNuevo.Enabled = b
'mnEliminar.Enabled = (Modo = 2)
'mnModificar.Enabled = (Modo = 2)
'mnSalir  siempre esta enabled
End Sub






Private Sub TransferirImagenesEntreAdos(Todos As Boolean)
Dim Anterior As Integer
Dim N As Integer
Dim Total As Integer
Dim TieneImagen As Boolean

    'Primero me cargo los datos temporales
        conn.Execute "DELETE FROM tmpImag"

    If Todos Then
        Anterior = Data1.Recordset.AbsolutePosition
        Total = Data1.Recordset.RecordCount
        Data1.Recordset.MoveFirst
        
    Else
        Total = 1
    End If
    
    AdotmoImg.ConnectionString = conn
    AdotmoImg.RecordSource = "Select * from tmpImag"
    AdotmoImg.Refresh
    
    For N = 1 To Total
        'Cargamos la imagen en este
        Adodc1.ConnectionString = ConnImg.ConnectionString
        Adodc1.RecordSource = "Select * from imagenes where idtrabajador =" & Data1.Recordset!idTrabajador
        Adodc1.Refresh
        
        TieneImagen = False
        If Not Adodc1.Recordset.EOF Then
            If Not IsNull(Adodc1.Recordset.Fields(1)) Then TieneImagen = True
        End If
        'La metemos en la tmp, obvianmente solo lo que tiene imagen
        If TieneImagen Then
            AdotmoImg.Recordset.AddNew
            AdotmoImg.Recordset.Fields(0) = Data1.Recordset.Fields!idTrabajador
        
            AdotmoImg.Recordset.Fields(1) = Adodc1.Recordset.Fields(1)
        
            AdotmoImg.Recordset.Update
        End If
        If Todos Then Data1.Recordset.MoveNext
        
    Next N
    
    'Vuelvo a situar el recordset en la poscion que le toca
    If Todos Then
        N = Anterior - 1
        Data1.Recordset.Move N, 1
    End If
    
End Sub


'0  .-  Insertar
'1  .-  Modificar
'2  .-  Eliminar
Private Sub UpdateaEuroges(Opcion As Byte)
'Dim SQL As String
'Dim N As Long
'
'    If Opcion > 0 Then
'        'Para modificar-eliminar
'        'MODIFICAR
'        N = Data1.Recordset!idTrabajador
'
'
'    Else
'        N = Val(Text1(0).Text)
'    End If
'        'Solo llevaremos a Multiase los que superen los 4000
'    If N <= 4000 Then Exit Sub
'    N = N - 4000 'Los codigos de los trabajadores en EUROGES van desde el 4000
'
'
'
'
'
'    Select Case Opcion
'    Case 0
'            SQL = "INSERT INTO straba(codsecci , codcapat , nromatri, grupocot, contrato, tipotrab, codepigr,"
'            SQL = SQL & "codtraba, nomtraba, domtraba, codpobla, niftraba, ssetraba, fecaltas, codcateg, codbanco,"
'            SQL = SQL & "codsucur , digcontr, cuentaba, fecantig, telefono"
'            SQL = SQL & "fecbajas , porcanti, porcirpf) VALUES ("
'            SQL = SQL & "4,0,0,0,'S','T',0,0,"   'VALORES FIJOS
'            SQL = SQL & N & "," & TextDB(Text1(5).Text, "T")
'            SQL = SQL & "," & TextDB(Text1(6).Text, "T")
'            SQL = SQL & "," & TextDB(Text1(9).Text, "T")    'codpostal
'            SQL = SQL & "," & TextDB(Text1(20).Text, "T")   'nif
'            SQL = SQL & "," & TextDB(Text1(18).Text, "T")   'ss
'            SQL = SQL & "," & TextDB(Text1(12).Text, "F")    'f alta
'            'La categoria esta en el combo2
'            SQL = SQL & "," & Combo2.ListIndex
'            SQL = SQL & "," & TextDB(Text1(30).Text, "T")    'codbanco
'            SQL = SQL & "," & TextDB(Text1(27).Text, "T")    '  ent
'            SQL = SQL & "," & TextDB(Text1(28).Text, "T")    '  dig control
'            SQL = SQL & "," & TextDB(Text1(29).Text, "T")    '  cuentaba
'            SQL = SQL & "," & TextDB(Text1(22).Text, "F")    'fecantig
'            SQL = SQL & "," & TextDB(Text1(10).Text, "T")    'telefono
'            SQL = SQL & "," & TextDB(Text1(13).Text, "F")    'fecbaja
'            SQL = SQL & "," & TextDB(Text1(23).Text, "N")    'procanti
'            SQL = SQL & "," & TextDB(Text1(26).Text, "N")    'porcirpf
'
'            SQL = SQL & ")"
'    Case 1
'
'
'
'            SQL = "UPDATE straba SET "
'            SQL = SQL & " nomtraba = " & LetDB(Data1.Recordset!Nomtrabajador, "T")
'            SQL = SQL & ", domtraba = " & LetDB(Data1.Recordset!domtrabajador, "T")
'            SQL = SQL & ", codpobla = " & LetDB(Data1.Recordset!codpostrabajador, "T")
'            SQL = SQL & ", niftraba = " & LetDB(Data1.Recordset!numdni, "T")
'            SQL = SQL & ", ssetraba = " & LetDB(Data1.Recordset!numSS, "T")
'            SQL = SQL & ", fecaltas = " & LetDB(Data1.Recordset!FecAlta, "F")
'            SQL = SQL & ", codcateg = " & LetDB(Data1.Recordset!idCategoria, "N")   '
'            SQL = SQL & ", codbanco = " & LetDB(Data1.Recordset!entidad, "T")
'            SQL = SQL & ", codsucur = " & LetDB(Data1.Recordset!oficina, "T")
'            SQL = SQL & ", digcontr = " & LetDB(Data1.Recordset!controlcta, "T")
'            SQL = SQL & ", cuentaba = " & LetDB(Data1.Recordset!cuenta, "T")
'            SQL = SQL & ", fecantig = " & LetDB(Data1.Recordset!Antiguedad, "F")
'            SQL = SQL & ", telefono = " & LetDB(Data1.Recordset!TelTrabajador, "T")
'            SQL = SQL & ", fecbajas = " & LetDB(Data1.Recordset!FecBaja, "F")
'            SQL = SQL & ", porcanti = " & LetDB(Data1.Recordset!porcantiguedad, "N")
'            SQL = SQL & ", porcirpf = " & LetDB(Data1.Recordset!porcirpf, "N")
'
'            SQL = SQL & " WHERE codtraba = " & N
'    Case 2
'            SQL = "DELETE FROM straba WHERE codtraba = " & N
'    End Select
'
'    Label2.Caption = "Actualizar euroges"
'    Label2.Visible = True
'    Me.Refresh
'    Screen.MousePointer = vbHourglass
'    DoEvents
'    EnlazaTrabajadoresEuroagro SQL
'
'    Label2.Visible = False
End Sub



Private Function LetDB(ByRef Campo As ADODB.Field, Tipo As String) As String

    If IsNull(Campo) Then
        LetDB = "NULL"
    Else
        Select Case Tipo
        Case "T"
            LetDB = "'" & ASQL(CStr(Campo)) & "'"
        Case "N"
            LetDB = CStr(TransformaComasPuntos(CStr(Campo)))
        Case "F"
            LetDB = "'" & Format(Campo, "yyyy-mm-dd") & "'"
        End Select
            
    End If
End Function


Private Function TextDB(ByRef T As String, Tipo As String) As String

    If T = "" Then
        TextDB = "NULL"
    Else
        Select Case Tipo
        Case "T"
            TextDB = "'" & ASQL(T) & "'"
        Case "N"
            TextDB = CStr(TransformaComasPuntos(T))
        Case "F"
            TextDB = "'" & Format(T, "yyyy-mm-dd") & "'"
        End Select
    End If
End Function



Private Sub ImagenHuella()
Dim Husu As UsuarioHuella

    On Error GoTo EIm
    
    
    
    
    Set Husu = New UsuarioHuella
    
    If Husu.Leer(Trim(Text1(2))) Then
        If Dir(MiEmpresa.DirHuellas & "\" & Husu.CodUsuario & ".jpg") <> "" Then
            imgHuella.Picture = LoadPicture(MiEmpresa.DirHuellas & "\" & Husu.CodUsuario & ".jpg")
            imgHuella.Visible = True
        Else
            imgHuella.Visible = False
        End If
    Else
        imgHuella.Visible = False
    End If

EIm:
    If Err.Number <> 0 Then MuestraError Err.Number, "Cargando huella"
    Set Husu = Nothing
End Sub




Private Sub GrabarUsuarioGestorHuella(AntiguoUsuario As String)
    Dim ssusu As UsuarioHuella
    
    If MiEmpresa.QueEmpresa = 0 Then Exit Sub
    
    If Me.Check2.Value = 0 Then If Modo = 3 Then Exit Sub
    
    
    Set ssusu = New UsuarioHuella
    
    If AntiguoUsuario <> "" Then
        If Text1(2).Text <> AntiguoUsuario Then
            If ssusu.Leer(AntiguoUsuario) Then ssusu.Eliminar
        End If
        
    End If
    
    
    
    If ssusu.Leer(Text1(2).Text) Then
        'NADA
        
    Else
    
        ssusu.CodUsuario = Text1(2)
        ssusu.GesLabID = Text1(0)
        ssusu.Mensaje = Left(Text1(5) & String(20, " "), 20)
    
    End If
    ssusu.FIR = ""
    ssusu.Guardar
    Set ssusu = Nothing
End Sub

Private Function EsCorrectoElCodigoTarjeta() As Boolean
Dim Cad As String
    EsCorrectoElCodigoTarjeta = True
    If MiEmpresa.QueEmpresa > 0 Then
        Cad = Trim(Text1(2).Text)
        If Cad <> "" Then Cad = "&H" & Cad
        PintaCodigoTrabajadorSinHuella Cad
        If lblInfCodigo.Caption = "> max" Then
            MsgBox "Codigo tarjeta excede del maximo de reloj", vbExclamation
            EsCorrectoElCodigoTarjeta = False
            
        End If
    End If
    
End Function
