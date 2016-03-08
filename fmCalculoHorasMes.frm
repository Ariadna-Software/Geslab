VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form fmCalculoHorasMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nominas mes"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "fmCalculoHorasMes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameALZIRA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5055
      Left            =   2760
      TabIndex        =   28
      Top             =   2040
      Width           =   6135
      Begin VB.TextBox txtDias 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1800
         TabIndex        =   41
         Top             =   3960
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Index           =   1
         Left            =   3960
         Picture         =   "fmCalculoHorasMes.frx":57E2
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton Command3 
         Height          =   375
         Index           =   0
         Left            =   3240
         Picture         =   "fmCalculoHorasMes.frx":5C24
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   4560
         Width           =   615
      End
      Begin VB.CommandButton cmdModificaHoras 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   49
         Top             =   4560
         Width           =   1095
      End
      Begin VB.CommandButton cmdModificaHoras 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   42
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   4440
         TabIndex        =   48
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   3000
         TabIndex        =   47
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   1800
         TabIndex        =   46
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   4440
         TabIndex        =   45
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   3000
         TabIndex        =   44
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   2
         Left            =   1800
         TabIndex        =   40
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   4440
         TabIndex        =   39
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   3000
         TabIndex        =   38
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   1800
         TabIndex        =   37
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox txtNeto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   4440
         TabIndex        =   35
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtBruto 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3000
         TabIndex        =   34
         Top             =   1320
         Width           =   1095
      End
      Begin VB.TextBox txtHoras 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1800
         TabIndex        =   33
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label9 
         Caption         =   "DIAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   240
         TabIndex        =   54
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Line Line9 
         X1              =   240
         X2              =   5640
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Label Label9 
         Caption         =   "Importe NETO"
         Height          =   195
         Index           =   6
         Left            =   4440
         TabIndex        =   52
         Top             =   1080
         Width           =   1020
      End
      Begin VB.Label Label9 
         Caption         =   "Importe BRUTO"
         Height          =   255
         Index           =   5
         Left            =   3000
         TabIndex        =   51
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "HORAS"
         Height          =   255
         Index           =   4
         Left            =   1920
         TabIndex        =   50
         Top             =   1080
         Width           =   735
      End
      Begin VB.Line Line8 
         X1              =   240
         X2              =   5640
         Y1              =   3000
         Y2              =   3000
      End
      Begin VB.Line Line7 
         X1              =   240
         X2              =   5520
         Y1              =   2280
         Y2              =   2280
      End
      Begin VB.Label Label9 
         Caption         =   "Posterior"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   36
         Top             =   3240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "UTILIZADA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   32
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Periodo"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   31
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "BOLSA"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   30
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Label8"
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
         Height          =   855
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   5775
      End
   End
   Begin VB.Frame FrameTipoAlzira 
      Height          =   855
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   11775
      Begin VB.ComboBox cboSeccion 
         Height          =   315
         Left            =   6840
         TabIndex        =   60
         Text            =   "Combo1"
         Top             =   360
         Width           =   3015
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   2
         Left            =   6360
         Picture         =   "fmCalculoHorasMes.frx":6066
         Style           =   1  'Graphical
         TabIndex        =   56
         ToolTipText     =   "Imprimir"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   1
         Left            =   5640
         Picture         =   "fmCalculoHorasMes.frx":6168
         Style           =   1  'Graphical
         TabIndex        =   55
         ToolTipText     =   "Recuperar datos"
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton cmdGenHorasAlzi 
         Caption         =   "Calcular horas"
         Height          =   315
         Left            =   2640
         TabIndex        =   24
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   1680
         TabIndex        =   23
         Text            =   "Text1"
         Top             =   300
         Width           =   855
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   300
         Width           =   1455
      End
      Begin VB.CommandButton cmdGeneraAlz 
         Caption         =   "Genera nominas"
         Height          =   315
         Left            =   10080
         TabIndex        =   27
         Top             =   360
         Width           =   1515
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   4080
         TabIndex        =   25
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   4080
         TabIndex        =   26
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   0
         Left            =   5160
         Picture         =   "fmCalculoHorasMes.frx":626A
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Guardar datos"
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   4080
      TabIndex        =   10
      Top             =   4560
      Width           =   4095
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   315
         Left            =   300
         TabIndex        =   11
         Top             =   360
         Width           =   3615
      End
   End
   Begin VB.Frame FrameMes 
      Height          =   855
      Left            =   60
      TabIndex        =   9
      Top             =   0
      Width           =   11655
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   7560
         TabIndex        =   59
         ToolTipText     =   "Depuracion"
         Top             =   360
         Width           =   255
      End
      Begin VB.CommandButton cmdQuitar 
         Height          =   315
         Left            =   8640
         Picture         =   "fmCalculoHorasMes.frx":636C
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Eliminar trabajador de la lista"
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Index           =   3
         Left            =   8040
         Picture         =   "fmCalculoHorasMes.frx":646E
         Style           =   1  'Graphical
         TabIndex        =   57
         ToolTipText     =   "MODIFICAR DATOS TRABAJADOR"
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdImprimir 
         Height          =   315
         Left            =   9120
         Picture         =   "fmCalculoHorasMes.frx":CCC0
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Imprimir datos actuales"
         Top             =   300
         Width           =   375
      End
      Begin VB.CommandButton cmdHPlus 
         Caption         =   "Quitar H+"
         Height          =   315
         Index           =   1
         Left            =   6240
         TabIndex        =   18
         Top             =   300
         Width           =   1035
      End
      Begin VB.CommandButton cmdHPlus 
         Caption         =   "Añadir H+"
         Height          =   315
         Index           =   0
         Left            =   5160
         TabIndex        =   17
         Top             =   300
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   4200
         TabIndex        =   16
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   4200
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdBaja 
         Caption         =   "Baja trabajador"
         Height          =   315
         Left            =   5160
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Genera nominas"
         Height          =   315
         Left            =   9720
         TabIndex        =   12
         Top             =   300
         Width           =   1755
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1455
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton cmdGenHoras 
         Caption         =   "Calcular horas"
         Height          =   315
         Left            =   2700
         TabIndex        =   2
         Top             =   300
         Width           =   1335
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   4800
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
            Picture         =   "fmCalculoHorasMes.frx":CDC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":D35C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":D8F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":DC10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":E1AA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":E5FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":EB96
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":F130
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":F6CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasMes.frx":FC64
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   60
      TabIndex        =   3
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   15
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "D / H"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "HN"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "HC"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Hor"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "H"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "H+"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "Ant"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Post"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Anticipos"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Plus"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Bolsa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   7380
      TabIndex        =   14
      Top             =   960
      Width           =   480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   7260
      X2              =   8400
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      X1              =   6000
      X2              =   7140
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   4740
      X2              =   5880
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   3420
      X2              =   4560
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3240
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   1860
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nomina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   6120
      TabIndex        =   8
      Top             =   960
      Width           =   645
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
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
      Left            =   4980
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Trabajadas"
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
      Height          =   195
      Left            =   3600
      TabIndex        =   6
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Oficial"
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
      Height          =   195
      Left            =   2220
      TabIndex        =   5
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador"
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
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   930
   End
   Begin VB.Menu mnPopup 
      Caption         =   "mnPopup"
      Visible         =   0   'False
      Begin VB.Menu mnVerDatos 
         Caption         =   "Ver datos/dia"
      End
      Begin VB.Menu mnbarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnModificaHoras 
         Caption         =   "Modificar Horas"
      End
   End
End
Attribute VB_Name = "fmCalculoHorasMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Byte
    '0.- Normal   . Es decir, como picassent. Autmaticamente compensa horas y demas
    '1.- Alzira   . Calcula horas y poco mas, y luego, a mano se compensan
    

    
    
Private SQL As String
Dim Importe1 As Currency
Dim primeravez As Boolean

Private NOD As ListItem
Private HorasxDia2 As Currency  'La leemos de parametros


Private Sub cmdGeneraAlz_Click()
    ProcesoDeGeneracionNominas
End Sub

Private Sub cmdGenHoras_Click()
Dim D As Integer
Dim FI As Date
Dim FF As Date

    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione un mes", vbExclamation
        Exit Sub
    End If
    If Val(Text1.Text) = 0 Then
        MsgBox "Año incorrecto.", vbExclamation
        Exit Sub
    End If
        
    If ListView1.ListItems.Count > 0 Then
        SQL = "Ya ha generado datos. ¿ Seguro que desea volverlos a generar ?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
        
    SQL = "/" & Combo1.ListIndex + 1 & "/" & Text1.Text
    FI = CDate("01" & SQL)
    D = DiasMes(Combo1.ListIndex + 1, CInt(Text1.Text))
    FF = CDate(D & "/" & Combo1.ListIndex + 1 & "/" & Text1.Text)
        
        
    If ComprobarMarcajesCorrectos(FI, FF, True) = 0 Then
        SQL = "No existe marcajes entre las fechas."
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
        
    If ComprobarMarcajesCorrectos(FI, FF, False) <> 0 Then
        SQL = "Existen marcajes incorrectos entre las fechas. ¿Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
        
        

        
    Label1.Caption = "Comienzo proceso"
    Frame1.Visible = True
    Me.Refresh
    
    Screen.MousePointer = vbHourglass
    
    ListView1.ListItems.Clear
    
    CalculaEntreFechas FI, FF
    Frame1.Visible = False
    Screen.MousePointer = vbDefault

End Sub


Private Sub CalculaEntreFechas(FI As Date, FF As Date)
Dim Rs As Recordset
Dim Horas As Currency
Dim Dias As Integer
Dim AntiguaFormaProcesar As Boolean
    
    conn.Execute "DELETE FROM tmpHorasMesHorario"

    'Para comprobar si estando de baja han trabajado
    'En tmpPresencia voy a guardar
    conn.Execute "DELETE FROM tmpCombinada"

    Set Rs = New ADODB.Recordset
    Rs.Open "horarios", conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    
    Label1.Caption = "Obtener horarios"
    Label1.Refresh
    
    While Not Rs.EOF
        Horas = CalculaHorasHorario(Rs.Fields(0), Dias, FI, FF, False)
        If Horas > 0 Then
            'Insertamos en tmp HORAS
            conn.Execute "INSERT INTO tmpHorasMesHorario(idHorario,Horas,Dias) VALUES (" & Rs.Fields(0) & "," & TransformaComasPuntos(CStr(Horas)) & "," & Dias & ")"
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    Label1.Caption = "Horas trabajadas"
    Label1.Refresh
    CalculaHorasTrabajadas FI, FF, 0, -1
    Me.Refresh
    

    
    Label1.Caption = "Datos periodo"
    Label1.Refresh
    CalculaDatosMes FI, FF, 0, -1
    
    Me.Refresh
    
    Label1.Caption = "Combina datos"
    Label1.Refresh
    CombinaDatos FI, FF
    
    'AHora realizamos los calculos de horas k kedan y demas
    Label1.Caption = "Datos a compensar"
    Label1.Refresh
    CalculoDatosACompensar
    
    Me.Refresh
    
    'Hacemos las comensaciones por horas
    Label1.Caption = "Compensaciones"
    Label1.Refresh
    
    AntiguaFormaProcesar = Dir(App.Path & "\AntigFP.dat", vbArchive) <> ""
   ' AntiguaFormaProcesar = True
    
    If AntiguaFormaProcesar Then
        'La que hacia antes
        HacerCompensaciones FI, FF, Label1
    Else
        Depuracion = (Check1.Value = 1)
        HacerCompensacionesPicassent FI, FF, Label1
    End If
    'Ajustamos los que no hayan trabakado nada
    AjustaDatosBajaMesEntero
    





    'NUEVO 20 OCTUBRE
    'Metemos los datos de los k trabajan en jornads semanales de OCHO horas
    'Control Nomina=4
    Label1.Caption = "Joranadas semanales"
    Label1.Refresh
    InsertarJornadasSemanales FI, FF





    'Ajustamos los de jornadas semanales
    



    Label1.Caption = "Carga datos"
    Label1.Refresh
    CargaDatos



    'Ahora vamos a comprobar si alguno de los k ha estado de baja
    'En este periodo a trabajado
    If ListView1.ListItems.Count > 0 Then
        Label1.Caption = "Comprobar bajas con dias Tra."
        Label1.Refresh
        Rs.Open "Select * from tmpcombinada", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            If HaTrabajadoConBaja(Rs) Then
                Dias = 0
                Do
                    Dias = Dias + 1
                    If Dias <= ListView1.ListItems.Count Then
                        If ListView1.ListItems(Dias).Text = Rs!idTrabajador Then
                            'Pongo el icono distinto
                            ListView1.ListItems(Dias).SmallIcon = 5
                            'Salgo
                            Dias = 32000
                        End If
                    End If
                Loop Until Dias > ListView1.ListItems.Count
            End If   'De ha trabajado estando de baja
            'Siguiente caso
        Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    Set Rs = Nothing
End Sub






Private Sub cmdGenHorasAlzi_Click()
Dim D As Integer
Dim FI As Date
Dim FF As Date

    If Combo2.ListIndex < 0 Then
        MsgBox "Seleccione un mes", vbExclamation
        Exit Sub
    End If
    If Val(Text2.Text) = 0 Then
        MsgBox "Año incorrecto.", vbExclamation
        Exit Sub
    End If
        
    If ListView1.ListItems.Count > 0 Then
        SQL = "Ya ha generado datos. ¿ Seguro que desea volverlos a generar ?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
        
    SQL = "/" & Combo2.ListIndex + 1 & "/" & Text2.Text
    FI = CDate("01" & SQL)
    D = DiasMes(Combo2.ListIndex + 1, CInt(Text2.Text))
    FF = CDate(D & "/" & Combo2.ListIndex + 1 & "/" & Text2.Text)
        
        
    If ComprobarMarcajesCorrectos(FI, FF, True) = 0 Then
        SQL = "No existe marcajes entre las fechas."
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
        
    If ComprobarMarcajesCorrectos(FI, FF, False) <> 0 Then
        SQL = "Existen marcajes incorrectos entre las fechas. ¿Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
        
        

        
    Label1.Caption = "Comienzo proceso"
    Frame1.Visible = True
    Me.Refresh
    
    Screen.MousePointer = vbHourglass
    
    ListView1.ListItems.Clear
    
    CalculaEntreFechasALZIRA FI, FF
    Frame1.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdHPlus_Click(Index As Integer)
Dim Importe As Currency
Dim Imp1 As Currency
Dim Rs As ADODB.Recordset

    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    If Index = 1 Then
        SQL = "reestablecer horas plus"
    Else
        SQL = "añadir horas plus"
    End If
    If MsgBox("Desea continuar con la opción " & SQL & " ?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        
    If Index = 0 Then
        'Si ya ha compensado le decimos k ya ha compensado
        If ListView1.SelectedItem.SubItems(10) <> "0.00" Then
            MsgBox "Ya ha compensando horas. Quite la comepnsacion primero", vbExclamation
            Exit Sub
        End If
    Else
        If ListView1.SelectedItem.SubItems(10) = "0.00" Then
            MsgBox "Ya ha compensando horas. Quite la compensacion primero", vbExclamation
            Exit Sub
        End If
    End If
    
    

    If Index = 0 Then
        'Cuando ponemos la baja calculamos si tiene horas en bolsa despues.
        'las tranformamos en euros de mas en anticpos
        Imp1 = -1
        Importe = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(12))
             Do
                 SQL = "Introduzca las horas de PLUS para " & ListView1.SelectedItem.SubItems(1) & "." & vbCrLf & "Máximo:" & Format(Importe, "0.00")
                 SQL = InputBox(SQL, "Horas +")
                 If SQL <> "" Then
                     If IsNumeric(SQL) Then
                         SQL = TransformaPuntosComas(SQL)
                         Imp1 = CCur(SQL)
                         If Imp1 > 0 Then
                            If Imp1 > Importe Then
                                MsgBox "No puede poner mas horas de las que tiene", vbExclamation
                                Imp1 = 0
                            Else
                                SQL = ""
                            End If
                        End If
                     End If
                 End If
             Loop Until SQL = ""
                         
            If SQL = "" And Imp1 <= 0 Then Exit Sub
                    
                    
      '  Importe = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(12))
        
       
            SQL = "SELECT Categorias.Importe1, Categorias.Importe2, Trabajadores.IdTrabajador,PorcSS,PorcIRPF"
            SQL = SQL & " FROM Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria"
            SQL = SQL & " WHERE Trabajadores.IdTrabajador=" & ListView1.SelectedItem.Text

            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Rs.EOF Then
                MsgBox "Error leyendo datos trabajador", vbExclamation
            Else
                'Le ponemos las horas de plus
                ListView1.SelectedItem.SubItems(10) = Format(Imp1, FormatoImporte)
                'En la bolsa le dejo las k tenia menos las k lleva al plus
                Importe = Importe - Imp1
                ListView1.SelectedItem.SubItems(12) = Format(Importe, FormatoImporte)
            
                Importe = Imp1 * Rs.Fields(0) 'importe2    horas * importe
                
                'PLUS
                ListView1.SelectedItem.SubItems(14) = Format(Importe, FormatoImporte)
                
                
                Imp1 = (Importe * Rs!porcSS) + (Importe * Rs!porcirpf)
                Imp1 = Imp1 / 100
                Importe = Importe - Imp1
                Importe = Round(Importe, 2)
                

               
               
                'Importe origninal
                Imp1 = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(13))
                Importe = Importe + Imp1
        
                'Ponemos las horas de plus
                ListView1.SelectedItem.SubItems(13) = Format(Importe, FormatoImporte)

                ListView1.SelectedItem.SmallIcon = 4 'Icono de h+

            End If
            Rs.Close
 
    Else
    
        PonerBaja False
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim i As Integer

    If ListView1.ListItems.Count < 1 Then Exit Sub
    
    
    'Borramos las dos tablas k utiliza
    SQL = "DELETE FROM tmpPagosMes"
    conn.Execute SQL
    SQL = "DELETE FROM tmpHoras"
    conn.Execute SQL
    espera 0.1
    
    'Para cada list item vamos a ver lo k pagamos
    VariableCompartida = "INSERT INTO tmpPagosMes(idTrabajador,nombre,SS,IRPF,HT,HC,importe1,Importe2,"
    VariableCompartida = VariableCompartida & "NETO,preciohora1,Pagos,BRUTO,INGRESAR) VALUES ("
    'Son en realidad
    '  OFICIAL         TRABAJADA       NOMINA            BOLSA             IMPORTES
    '   D   H         D   HN  HC      D  H  H+      Antes   Despues    PAGOS  PLUS   ANTICPOS
    
    ' Dias Trabajados y duas nomina van en la tabla tmpHoras,, en campos Dias, HorasE
    
    For i = 1 To ListView1.ListItems.Count
        With ListView1.ListItems(i)
            SQL = .Text & ",'" & .SubItems(1) & "',"
            
            'OFICIALES
            SQL = SQL & Mid(.SubItems(2), 1, InStr(1, .SubItems(2), "/") - 1) & ",'"
            SQL = SQL & TransformaComasPuntos(Mid(.SubItems(2), InStr(1, .SubItems(2), "/") + 1)) & "',"
            
            
            'Horas Normales y compensables
            SQL = SQL & TransformaComasPuntos(.SubItems(4)) & "," & TransformaComasPuntos(.SubItems(5)) & ","
            
            'Horas Nomina y H+
            SQL = SQL & TransformaComasPuntos(.SubItems(9)) & "," & TransformaComasPuntos(.SubItems(10)) & ","
            
            'Bolsa antes y despues
            SQL = SQL & TransformaComasPuntos(.SubItems(11)) & "," & TransformaComasPuntos(.SubItems(12)) & ","
            
            'Importes: pagos, PULS y Anticpipos
            SQL = SQL & "0," & TransformaComasPuntos(CCur(.SubItems(13))) & ","
            SQL = SQL & TransformaComasPuntos(CCur(.SubItems(14))) & ")"
            conn.Execute VariableCompartida & SQL
            
            'Insertamos los dias en tmpHoras
            SQL = "INSERT INTO tmpHoras (trabajador,Dias,horasE) VALUES (" & .Text & ","
            SQL = SQL & .SubItems(3) & "," & .SubItems(8) & ")"
            conn.Execute SQL
        End With
    Next i
    
    If MiEmpresa.NominaAutomatica Then
        SQL = "Mes= """ & UCase(Combo1.List(Combo1.ListIndex)) & " " & Text1.Text & """|"
    Else
        SQL = "Mes= """ & UCase(Combo2.List(Combo2.ListIndex)) & " " & Text2.Text & """|"
    End If
    frmImprimir.opcion = 15
    frmImprimir.OtrosParametros = SQL
    frmImprimir.NumeroParametros = 1
    frmImprimir.Show vbModal
    
End Sub

Private Sub cmdModificaHoras_Click(Index As Integer)
    If Index = 1 Then
        PonerFrameHoras False
    Else
        Importe1 = CCur(txtHoras(3).Text)
        If Importe1 < 0 Then
            MsgBox "No puedes asignar mas horas de las que tiene.", vbExclamation
            Exit Sub
        End If
            
        Importe1 = CCur(txtHoras(2).Text)
        If Importe1 = 0 And txtDias.Text <> "0" Then
            MsgBox "No se pueden compensar dias en nómina sin compensar horas.", vbExclamation
            Exit Sub
        End If
            
        'Habria k volver a leer el dato
        Screen.MousePointer = vbHourglass
        ModificaDatosTrabajador
        Screen.MousePointer = vbDefault
        Me.Command3(1).SetFocus
    End If
End Sub



Private Sub PonerFrameHoras(Habilitado As Boolean)
    Me.FrameALZIRA.Visible = Habilitado
    FrameTipoAlzira.Enabled = Not Habilitado
    ListView1.Enabled = Not Habilitado

End Sub




Private Sub cmdQuitar_Click()
    'Eliminar los datos del trabjaodr
    On Error GoTo E1
            'Modificar datos trabajador
        If Me.ListView1.ListItems.Count = 0 Then Exit Sub
        If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
        
        SQL = "¿Desea eliminar de la nomina al trabajador: " & ListView1.SelectedItem.Text & " - " & ListView1.SelectedItem.SubItems(1) & "?"
        If MsgBox(SQL, vbQuestion + vbYesNo) = vbYes Then
            SQL = "DELETE FROM tmpDatosMes WHERE tmpDatosMes.trabajador =" & ListView1.SelectedItem.Text
            conn.Execute SQL
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
            
        End If


    Exit Sub
E1:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Combo2_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Command1_Click()
    ProcesoDeGeneracionNominas
End Sub


Private Sub ProcesoDeGeneracionNominas()
Dim B As Boolean
Dim Rs As ADODB.Recordset
Dim i As Integer

    If ListView1.ListItems.Count < 1 Then Exit Sub
    
    'Preguntamos si desea continuar
    SQL = "Seguro que desea generar las nóminas con estos valores?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
    
    
    'Vemos si ya se han generado las nominas del mes
    If MiEmpresa.NominaAutomatica Then
        i = DiasMes(Combo1.ListIndex + 1, CInt(Text1.Text))
        SQL = "#" & Text1.Text & "/" & Combo1.ListIndex + 1 & "/" & i & "#"
    Else
        i = DiasMes(Combo2.ListIndex + 1, CInt(Text2.Text))
        SQL = "#" & Text2.Text & "/" & Combo2.ListIndex + 1 & "/" & i & "#"
    End If
    SQL = "Select * from Nominas where Fecha = " & SQL
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = ""
    If Not Rs.EOF Then SQL = "SI"
    Rs.Close
    Set Rs = Nothing
    
    If SQL <> "" Then
        
        MsgBox "Ya se han generado las nominas de este mes.", vbExclamation
        
    End If
    
    
    'pondremos un transaccion
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    If MiEmpresa.NominaAutomatica Then
        B = GenerarNominas
    Else
        B = genNomALZIRA
    End If
    If B Then
        conn.CommitTrans
        MsgBox "Proceso finalizado", vbInformation
    Else
        conn.RollbackTrans
    End If
    Screen.MousePointer = vbDefault

End Sub



Private Sub Command2_Click(Index As Integer)
Dim i As Integer
Dim Rs As ADODB.Recordset
    Select Case Index
    Case 0
        'Guardar los datos
        SQL = "Desea guardar los cambios?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        ModificarRecuperar True
    Case 1
         'Recuperar datos
        SQL = "Desea recuperar los datos almacenados?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Screen.MousePointer = vbHourglass
        ModificarRecuperar False
        Screen.MousePointer = vbDefault
        
    Case 2
        'Imprimir
        cmdImprimir_Click
        
        
    Case 3
        'Modificar datos trabajador
        If Me.ListView1.ListItems.Count = 0 Then Exit Sub
        If Me.ListView1.SelectedItem Is Nothing Then Exit Sub
        
        If ListView1.SelectedItem.SubItems(10) <> "0.00" Then
            MsgBox "Quite primero los anticipos", vbExclamation
            Exit Sub
        End If
        
        frmCambiosDatosNomina.opcion = 0
        Load frmCambiosDatosNomina
        
        VariableCompartida = "" 'Si guarda o no guarda
        With ListView1.SelectedItem
            
            SQL = Combo1.ListIndex + 1 & "  / " & Combo1.Text
            frmCambiosDatosNomina.Caption = SQL
            frmCambiosDatosNomina.lblIdTra(0) = .Text             'Trabajador
            frmCambiosDatosNomina.lblTra(0) = " - " & .SubItems(1)              'Trabajador
            i = InStr(.SubItems(2), "/")
            'OFICIALES
            frmCambiosDatosNomina.txtDias(0).Text = Mid(.SubItems(2), 1, i - 1) '
            frmCambiosDatosNomina.txtHN(0).Text = Mid(.SubItems(2), i + 1)
            'TRABAJADAS
            frmCambiosDatosNomina.txtDias(1).Text = .SubItems(3)
            frmCambiosDatosNomina.txtHN(1).Text = .SubItems(4)
            frmCambiosDatosNomina.txtHC(1).Text = .SubItems(5)
            'Nomina
            frmCambiosDatosNomina.txtDias(2).Text = .SubItems(8)
            frmCambiosDatosNomina.txtHN(2).Text = .SubItems(4)
            frmCambiosDatosNomina.txtHC(2).Text = .SubItems(9)
            'Bolsa horas
            frmCambiosDatosNomina.txtBolsa(0).Text = .SubItems(11)
            frmCambiosDatosNomina.txtBolsa(1).Text = .SubItems(12)

            
            frmCambiosDatosNomina.Show vbModal
        End With
        If VariableCompartida <> "" Then
            'HA UPDATEADO LOS DATOS
            PonSQL ListView1.SelectedItem.Text
            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            If Not Rs.EOF Then
                
                PonLinea ListView1.SelectedItem, Rs
                ListView1.SelectedItem.SmallIcon = 4
            End If
            Rs.Close
        End If
    End Select
End Sub


Private Sub ModificarRecuperar(Guardar As Boolean)

    If Guardar Then
        'Borramos los datos de la 2
        SQL = "Delete from tmpDatosMes2"
        conn.Execute SQL
        
        'Insertamos tmp
        SQL = "INSERT INTo tmpDatosMes2 SELECT * from tmpDatosMES"
        conn.Execute SQL
        
        
        'UPDATEAMOS para guardar el año
        'Es decir, en la tabla tmpdatosmes2 habra en lugar del mes solo
        'habra yyyymm
        SQL = "UPDATE tmpDatosMEs2 SET mes=" & Text2.Text & Combo2.ListIndex + 1
        conn.Execute SQL
        
    Else
        'Borramos los datos de la 1
        SQL = "Delete from tmpDatosMes"
        conn.Execute SQL
        
        'truquito
        CadParam = "Error leyendo datos almacenados."
        SQL = DevuelveDesdeBD("mes", "tmpdatosmes2", "mes", "mes", "N")
        If SQL = "" Then
            MsgBox CadParam, vbExclamation
            Exit Sub
        End If
        
        Importe1 = Val(Mid(SQL, 1, 4))
        If Importe1 = 0 Then
            MsgBox CadParam, vbExclamation
            Exit Sub
        End If
        Text2.Text = Importe1
        
        
        Importe1 = Val(Mid(SQL, 5, 2))
        Importe1 = Importe1 - 1
        Combo2.ListIndex = CInt(Importe1)
        'UPDATEAMOS para dejar el mes solamente
        SQL = "UPDATE tmpDatosMEs2 SET mes=" & Importe1 + 1
        conn.Execute SQL
        
        
        'Insertamos tmp
        SQL = "INSERT INTo tmpDatosMes SELECT * from tmpDatosMES2"
        conn.Execute SQL
        
        
        
        'Volvemos a poner el año en el dato
        'UPDATEAMOS para guardar el año
        'Es decir, en la tabla tmpdatosmes2 habra en lugar del mes solo
        'habra yyyymm
        SQL = "UPDATE tmpDatosMEs2 SET mes=" & Text2.Text & Combo2.ListIndex + 1
        conn.Execute SQL
        
        'Cargamos datos
        CargaDatos
        
        
        
    End If
    
End Sub


Private Sub Command3_Click(Index As Integer)
Dim Indice As Integer
Dim OK As Boolean

    'Siguiente ITEM
    
    
    'Se trata de buscar el sigueinte ITEM
    OK = False
    Do
        If Index = 0 Then
            Indice = NOD.Index - 1
        Else
            Indice = NOD.Index + 1
        End If
        
        If Indice > ListView1.ListItems.Count Then
            SQL = "Es el ultimo trabajador. ¿Desea continuar por el principio?"
            If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Indice = 1
            
        Else
            If Indice < 1 Then
                SQL = "Es el primer trabajador. ¿Desea ir al final?"
                If MsgBox(SQL, vbQuestion + vbYesNo) = vbNo Then Exit Sub
                Indice = ListView1.ListItems.Count
            End If
        End If
        
        Set NOD = ListView1.ListItems(Indice)
        If NOD.SmallIcon <> 3 Then OK = True
     
    Loop Until OK
        
        
        Screen.MousePointer = vbHourglass
        Set ListView1.SelectedItem = NOD
        NOD.EnsureVisible
        'Ponemos los valores del nodo
        ModificarDesdeNodo

        'Ponemos el foco en horas
        Screen.MousePointer = vbDefault
        txtHoras(2).SetFocus
        
    
   
End Sub

Private Sub Form_Activate()
    If primeravez Then
        primeravez = False
        If opcion = 1 Then
            Combo2.SetFocus
        Else
            Combo1.SetFocus
        End If
    End If
End Sub

Private Sub Form_Load()
    primeravez = True
    FrameALZIRA.Visible = False
    Frame1.Visible = False
    FrameTipoAlzira.Visible = opcion = 1
    Me.FrameMes.Visible = opcion = 0
    If opcion = 0 Then
        CargaCombo Me.Combo1, Text1
        Command1.Enabled = vUsu.Nivel < 2 'Administrador
    Else
        CargaCombo Me.Combo2, Text2
        cmdGeneraAlz.Enabled = vUsu.Nivel < 2
    End If
    
    ListView1.SmallIcons = Me.ImageList1

    SQL = DevuelveDesdeBD("HorasJornada", "empresas", "idempresa", 1, "N")
    If SQL <> "" Then
        HorasxDia2 = CCur(SQL)
    Else
        HorasxDia2 = 0
    End If
    


    CargaComboSecciones cboSeccion, True
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub CargaCombo(ByRef C As ComboBox, ByRef T As TextBox)
Dim i As Integer
Dim F As Date

    For i = 1 To 12
        C.AddItem Format(CDate("01/" & i & "/2000"), "mmmm")
    Next i
    F = DateAdd("m", -1, Now)
    C.ListIndex = Month(F) - 1
    T.Text = Year(F)
End Sub


Private Sub CargaColumnas()
Dim Anch As Single
Dim clmX As ColumnHeader

'ListView1.ColumnHeaders.Clear
'Anch = ListView1.Width - 360
'Anch = Anch / 16
'
'
''Datos Trbajador
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Cod"
'clmX.Width = Anch
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Nombre"
'clmX.Width = Anch * 5
'
'
'
''OFICIALES
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Dias"
'clmX.Width = 510
'
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Horas"
'clmX.Width = Anch
'
''TRABAJADOS
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Dias"
'clmX.Width = 510
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Norm."
'clmX.Width = Anch
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Comp"
'clmX.Width = Anch
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "PLUS"
'clmX.Width = Anch
'
'
''Saldo
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Dias"
'clmX.Width = 510
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Horas"
'clmX.Width = Anch
'
'
''Bolsa
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Ant."
'clmX.Width = Anch
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Post"
'clmX.Width = Anch
'
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Anticipos"
'clmX.Width = Anch
'


ListView1.ColumnHeaders(14).Width = ListView1.Width - 320 - ListView1.ColumnHeaders(14).Left

For Each clmX In ListView1.ColumnHeaders
    
    If clmX.Index > 3 Then clmX.Alignment = lvwColumnRight
Next

'Las lineas
With ListView1


    Line1.X2 = .ColumnHeaders(3).Left - 30 + 160

    Label3.Left = .ColumnHeaders(3).Left + 160
    Line2.X1 = .ColumnHeaders(3).Left + 30 + 160
    Line2.X2 = .ColumnHeaders(4).Left - 30 + 160
    
    Label4.Left = .ColumnHeaders(4).Left + 160
    Line3.X1 = .ColumnHeaders(4).Left + 30 + 160
    Line3.X2 = .ColumnHeaders(7).Left - 30 + 160
    
    Label5.Left = .ColumnHeaders(7).Left + 160
    Line4.X1 = .ColumnHeaders(7).Left + 30 + 160
    Line4.X2 = .ColumnHeaders(9).Left - 30 + 160
    
    Label6.Left = .ColumnHeaders(9).Left + 160
    Line5.X1 = .ColumnHeaders(9).Left + 30 + 160
    Line5.X2 = .ColumnHeaders(12).Left - 30 + 160
    
    Label7.Left = .ColumnHeaders(12).Left + 160
    Line6.X1 = .ColumnHeaders(12).Left + 30 + 160
    Line6.X2 = .ColumnHeaders(14).Left - 30 + 160
    
    'Pequeño reajuste k borda las lineas
    .ColumnHeaders(3).Width = 1000
    .ColumnHeaders(14).Width = 1300
    'La ultima columna a 0
    .ColumnHeaders(15).Width = 0
End With
    


End Sub

Private Sub PonLinea(ByRef i As ListItem, ByRef Rs As ADODB.Recordset)
'Si tiene dias pendientes
Dim J As Integer
        
        
        If MiEmpresa.NominaAutomatica Then
            'Normal. Pica y cata
            If Rs!DiasTrabajados = 0 Then
                If Rs!mesdias = 0 Then
                    'ESTA DE BAJA
                    J = 3
                Else
                    J = 10
                End If
            Else
                If Rs!ControlNomina = 1 Then
                    'Normal
                    J = 1
                Else
                    If Rs!ControlNomina = 3 Then
                        'Jorandas semanas
                        J = 8
                    Else
                        'Tipo de liquidaciones
                        J = 6
                    End If
                End If
                If Rs!saldodias <> 0 Then J = J + 1
            End If
        
        Else
            'Como alzira
            If Rs!DiasTrabajados = 0 Then
                J = 3
            Else
                If Rs!bolsaperiodo <> 0 Then
                    J = 4
                Else
                    If Rs!saldodias > 0 Then
                        'Ya ha compensado
                        J = 2
                    Else
                        J = 1
                    End If
                End If
            End If
        End If
        i.SmallIcon = J
        i.Text = Rs!Trabajador
        i.SubItems(1) = Rs!nomtrabajador
        i.ToolTipText = Rs!nomtrabajador
        
        'Horas oficiles
        i.SubItems(2) = Rs!mesdias & "/" & Format(Rs!meshoras, "0.00")
        
        'Trabajados
        i.SubItems(3) = Rs!DiasTrabajados
        i.SubItems(4) = Format(Rs!horasn, "0.00")
        i.SubItems(5) = Format(Rs!horasc, "0.00")
        
        
        'Saldo
        i.SubItems(6) = Rs!saldodias
        i.SubItems(7) = Format(Rs!saldoh, "0.00")
        
        'Compensadas en NOMINA
        i.SubItems(8) = Rs!diasperiodo
        i.SubItems(9) = Format(Rs!extras, "0.00")
        
        
        
        '
        'Bolsa
        i.SubItems(11) = Rs!bolsaantes
        i.SubItems(12) = Format(Rs!bolsadespues, "0.00")
        
        Importe1 = Rs!Anticipos + Rs!plus
        i.SubItems(13) = Format(Importe1, "0.00")

        'PLUS
        If opcion = 0 Then
            i.SubItems(10) = "0.00"
            i.SubItems(14) = Format(DBLet(Rs!plus, "N"), "0.00")
        Else
            'ALZIRA. Aqui van las Horas extras en bruto
            i.SubItems(10) = DBLet(Rs!bolsaperiodo, "N")
            i.SubItems(14) = Format(DBLet(Rs!brutoplus, "N"), "0.00")
        End If
        'El tag
        i.Tag = Rs!ControlNomina
End Sub


Private Sub CargaDatos()
Dim i As ListItem
Dim Rs As ADODB.Recordset


    Set Rs = New ADODB.Recordset
    ListView1.ListItems.Clear
    PonSQL ""
    SQL = SQL & " order by "

    NParam = 1
    If MiEmpresa.NominaAutomatica Then
        If Option1(0).Value Then NParam = 0
    Else
        If Option2(0).Value Then NParam = 0
    End If
    
    
        If NParam = 0 Then
            SQL = SQL & "id"
        Else
            SQL = SQL & "nom"
        End If
        NParam = 0
    
    SQL = SQL & "Trabajador"
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Set i = ListView1.ListItems.Add
        PonLinea i, Rs
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
End Sub




Private Sub Form_Resize()
Dim H As Single
Dim W As Single

    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 7000 Then
        W = 7000
        Me.Width = W
    Else
        W = Me.Width
    End If
    If Me.Height < 3900 Then
        H = 3900
        Me.Height = H
    Else
        H = Me.Height
    End If
    Me.ListView1.Width = W - ListView1.Left - 210
    Me.ListView1.Height = H - ListView1.Top - 500
    CargaColumnas
End Sub

'Private Sub ListView1_Click()
'Dim i
'    SQL = ""
'    For i = 1 To ListView1.ColumnHeaders.Count
'        SQL = SQL & ListView1.ColumnHeaders(i).Text & ": " & ListView1.ColumnHeaders(i).Width & vbCrLf
'    Next i
'    MsgBox SQL
'End Sub


Private Sub PonerBaja(Baja As Boolean)
Dim Importe As Currency
Dim Imp1 As Currency
Dim Rs As ADODB.Recordset




    If Baja Then
    
    
        'Cuando ponemos la baja calculamos si tiene horas en bolsa despues.
        'las tranformamos en euros de mas en anticpos
        Importe = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(12))
        
        If Importe > 0 Then
            SQL = "SELECT Categorias.Importe1, Categorias.Importe2, Trabajadores.IdTrabajador,PorcSS,PorcIRPF"
            SQL = SQL & " FROM Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria"
            SQL = SQL & " WHERE Trabajadores.IdTrabajador=" & ListView1.SelectedItem.Text

            Set Rs = New ADODB.Recordset
            Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Rs.EOF Then
                MsgBox "Error leyendo datos trabajador", vbExclamation
            Else
                Importe = Importe * Rs.Fields(0) 'importe2
                Imp1 = (Importe * Rs!porcSS) + (Importe * Rs!porcirpf)
                Imp1 = Imp1 / 100
                Importe = Importe + Imp1
                Importe = Round(Importe, 2)
                'PLUS
                ListView1.SelectedItem.SubItems(14) = Format(Importe, FormatoImporte)
               
               
                'Importe origninal
                Imp1 = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(13))
                Importe = Importe + Imp1
                ListView1.SelectedItem.SubItems(12) = "0.00" 'Le quitamos la bolsa
                ListView1.SelectedItem.SubItems(13) = Format(Importe, FormatoImporte)

            End If
            Rs.Close
        End If
    Else
        'Reestablecemos los valores de tmpDatosmes
        PonSQL ListView1.SelectedItem.Text
        
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Rs.EOF Then
            MsgBox "Error leyendo datos tmpDatosMES del trabajador : " & ListView1.SelectedItem.Text, vbExclamation
        Else
            PonLinea ListView1.SelectedItem, Rs
        End If
        Rs.Close
    End If
    Set Rs = Nothing
End Sub


Private Sub PonSQL(Id As String)
    SQL = "Select tmpDatosMes.*,nomtrabajador,controlnomina from tmpDatosMes,Trabajadores"
    SQL = SQL & " WHERE tmpDatosMes.trabajador = Trabajadores.idTrabajador "
    If Id <> "" Then SQL = SQL & " AND tmpDatosMes.trabajador =" & Id
End Sub




Private Function GenerarNominas() As Boolean
Dim i As Integer
Dim Cad As String
Dim Importe As Currency

On Error GoTo EGenerarNominas
    GenerarNominas = False
    
    SQL = "INSERT INTO Nominas (Fecha,IdTrabajador,Dias,HN,HC,Plus,Anticipos,BolsaAntes,BolsaDespues,HP,DiasTra) VALUES (#"
    i = DiasMes(Combo1.ListIndex + 1, CInt(Text1.Text))
    SQL = SQL & Text1.Text & "/" & Combo1.ListIndex + 1 & "/" & i & "#,"
    
    'Primero generamos la tabla de  nominas con los importes marcados aqui
    For i = 1 To ListView1.ListItems.Count
        Cad = ListView1.ListItems(i).Text & "," & ListView1.ListItems(i).SubItems(8) & ","
        'Hnormales y Hcompensadas
        Cad = Cad & TransformaComasPuntos(ListView1.ListItems(i).SubItems(4)) & "," & TransformaComasPuntos(ListView1.ListItems(i).SubItems(9)) & ","
        
        'PLUS Y anticipos
        Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(14)))) & ","
        Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(13)))) & ","
        
        'Bolsa Antes. Para el informe conjunto a la nomina
        Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(11)))) & ","
        Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(12)))) & ","
        
        'Horas PLUS productividad
        Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(10)))) & ","
        'Dias trabajado
        Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(3))))
        
        
        Cad = Cad & ")"
        Cad = SQL & Cad
        conn.Execute Cad
        
        
        
        
        'Pondremos la bolsa de horas Y, hay bajas,
        'entonces actualizaremos la baja de cada trabajador
        'al ultimo dia trabajado
        Cad = "UPDATE Trabajadores SET Bolsahoras = " & TransformaComasPuntos(ListView1.ListItems(i).SubItems(12))
        Cad = Cad & " WHERE idTrabajador = " & ListView1.ListItems(i).Text
        conn.Execute Cad
        
        
        
'        'Si se da de baja le pongo fecha de baja
'        If ListView1.ListItems(I).SmallIcon = 3 Then
'            'SE DA DE BAJA
'            Cad = DevuelveDesdeBD("fecbaja", "trabajadores", "IdTrabajador", ListView1.ListItems(I).Text, "N")
'            If Cad = "" Then
'                'NO TIENE FECHA BAJA
'                Cad = DiasMes(Combo1.ListIndex + 1, Int(Text1.Text)) & "/" & CStr(Combo1.ListIndex + 1) & "/" & Text1.Text
'                Cad = Format(Cad, FormatoFecha)
'                Cad = "UPDATE Trabajadores SET fecbaja = #" & Cad & "#"
'                Cad = Cad & " WHERE idTrabajador = " & ListView1.ListItems(I).Text
'                Conn.Execute Cad
'            End If
'        End If
        
        
    Next i
    

    
    
    
    
    GenerarNominas = True
    Exit Function
EGenerarNominas:
    MuestraError Err.Number
End Function





Private Function PuedeCompensarDias() As Integer
Dim i As Integer

    PuedeCompensarDias = 0
    SQL = DevuelveDesdeBD("idHorario", "Trabajadores", "idTrabajador", ListView1.SelectedItem.Text, "N")
    i = Val(SQL)
    
    'En la tabla tmpHorasMesHorario, al cargar los datos
    'se han cargado las horas oficiales
    SQL = DevuelveDesdeBD("Dias", "tmpHorasMesHorario", "idHorario", CStr(i), "N")
    If SQL <> "" Then
        i = Val(SQL)
        i = i - Val(ListView1.SelectedItem.SubItems(8))
        If i > 0 Then PuedeCompensarDias = i
    End If
    
    
    
End Function


Private Sub CompensarDias(Dias As Integer)
Dim i As Integer
Dim Lab As Integer
Dim H As Currency
Dim H1 As Currency
Dim D1 As Integer
Dim Rs As ADODB.Recordset


    SQL = DevuelveDesdeBD("idHorario", "Trabajadores", "idTrabajador", ListView1.SelectedItem.Text, "N")
    i = Val(SQL)

    Lab = DiasLaborablesSemana(i)
    If Lab < 1 Then Exit Sub

    If Dias < Lab Then
        'Nos salimos pq no tengo bastantes dias para compensar un semana
        Exit Sub
    End If

    

    'QUiero saber las horas a la semana k puedo compensar
    Set Rs = New ADODB.Recordset
    SQL = "Select * from Horarios Where idHorario =" & i
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rs.EOF Then
    
        H = CCur(ListView1.SelectedItem.SubItems(11)) 'Las horas k le van a quedar en bolsa
        'Ya tengo el horario y los dias a compensar
        
        'A compensar
        D1 = 0 'Dias
        H1 = 0 'Horas
        
        'Por lo tanto veo cuantas semanas mas voy a compensar
        Do
            i = Dias \ Lab
            If i > 0 Then
                'Una semana seguro k puedo compensar. Vamos palla
                If H >= Rs!TotalHoras Then   'Horas semana
                    D1 = D1 + Lab
                    H1 = H1 + Rs!TotalHoras
                    H = H - Rs!TotalHoras
                End If
                Dias = Dias - Lab
            End If
        Loop Until i = 0
    End If
    Rs.Close
    
    
    'Si a compensado lo reflejo en la listview
    If D1 > 0 Then
        'Dias nomina
        
        
        'Horas para la nomina
        SQL = "Select Importe1,importe2,porcSS,porcIRPF from Categorias,Trabajadores WHERE Trabajadores.IdCategoria = Categorias.IdCategoria"
        SQL = SQL & " AND Trabajadores.idTrabajador =" & ListView1.SelectedItem.Text
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            'Ponemos ya las nuevas horas en horas normales
            H = CCur(ListView1.SelectedItem.SubItems(4)) + H1
            ListView1.SelectedItem.SubItems(4) = Format(H, FormatoImporte)
            'Bolsa
            H = CCur(ListView1.SelectedItem.SubItems(11)) - H1
            ListView1.SelectedItem.SubItems(11) = Format(H, FormatoImporte)
            
            'Precio bruto
            H = H1 * Rs!Importe1
            
            'Precio neto
            H1 = ((H * Rs!porcSS) + (H * Rs!porcirpf)) / 100
            
            H = Round(H - H1, 2)
            'Anticipos
            H1 = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(12))
            H = H + H1
            ListView1.SelectedItem.SubItems(12) = H
            
            'Dias nomina
            i = Val(ListView1.SelectedItem.SubItems(8)) + D1
            ListView1.SelectedItem.SubItems(8) = i
        End If
    End If
    Set Rs = Nothing
End Sub



Private Sub Form_Unload(Cancel As Integer)
    If opcion = 1 Then
        'ALZIRA
        If MsgBox("Seguro que desea salir de la edición de Nominas?", vbQuestion + vbYesNo) = vbNo Then Cancel = 1
    End If
End Sub



Private Sub ListView1_DblClick()
Dim vH As CHorarios
Dim F As Date
Dim F2 As Date
Dim MEDIOS As String

    With ListView1.SelectedItem
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        SQL = DevuelveDesdeBD("idHorario", "Trabajadores", "idTrabajador", .Text, "N")
        If SQL = "" Then
            MsgBox "Error leyendo datos trabajador", vbExclamation
        Else
                
            
            Set vH = New CHorarios
            vH.IdHorario = Val(SQL)
            SQL = ""
            If opcion = 0 Then
                F = CDate("01/" & Combo1.ListIndex + 1 & "/" & Text1.Text)
            Else
                F = CDate("01/" & Combo2.ListIndex + 1 & "/" & Text2.Text)
            End If
            F2 = F
            F = DateAdd("m", 1, F)
            F = DateAdd("d", -1, F)
            SQL = ""
            MEDIOS = vH.LeerMediosDias(vH.IdHorario, F2, F)
            
            SQL = vH.LeerDiasFestivos(vH.IdHorario, F2, F)
            frmVerDiasMesTrabajador3.DiasEnNomina = .SubItems(8)
            frmVerDiasMesTrabajador3.TodoElMEs = 0
            frmVerDiasMesTrabajador3.JornadasSemanales = (.Tag = 3)
            frmVerDiasMesTrabajador3.MediosDias = MEDIOS
            frmVerDiasMesTrabajador3.FESTIVOS = SQL
            frmVerDiasMesTrabajador3.Trabajador = .SubItems(1) & "|" & .Text & "|"
            frmVerDiasMesTrabajador3.FechaIni = F2
            frmVerDiasMesTrabajador3.HorasMinimoDia = HorasxDia2
            frmVerDiasMesTrabajador3.Show vbModal
            Set vH = Nothing
        End If
    End With
End Sub




Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Not MiEmpresa.NominaAutomatica Then
            PopupMenu mnPopup
        End If
    End If
End Sub

Private Sub ModificarDesdeNodo()
Dim RT As ADODB.Recordset
Dim Imp2 As Currency

   
    If NOD Is Nothing Then Exit Sub
    

    'Pongo los datos del text
    Label8.Caption = NOD.SubItems(1)
    Label8.Tag = NOD.Text
    
      
    SQL = "SELECT Trabajadores.IdTrabajador, Trabajadores.bolsaBRUTO, Trabajadores.bolsaNETO, Trabajadores.bolsahoras,"
    SQL = SQL & " tmpDatosMEs.HorasC, tmpDatosMEs.brutoperiodo, tmpDatosMEs.netoperiodo, tmpDatosMEs.DiasPeriodo,tmpDatosMEs.DiasTrabajados,mesdias,"
    SQL = SQL & " tmpDatosMEs.bolsaperiodo, Trabajadores.PorcSS, Trabajadores.porcIRPF,tmpDatosMEs.SaldoDias, "
    SQL = SQL & " tmpDatosMEs.bolsaDespues, tmpDatosMEs.netodespues, tmpDatosMEs.brutodespues, categorias.importe2"
    SQL = SQL & " FROM tmpDatosMEs INNER JOIN (Categorias INNER JOIN Trabajadores ON "
    SQL = SQL & " Categorias.IdCategoria = Trabajadores.idCategoria) ON tmpDatosMEs.Trabajador"
    SQL = SQL & " = Trabajadores.IdTrabajador"
    SQL = SQL & " WHERE idtrabajador =" & NOD.Text
    Set RT = New ADODB.Recordset
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RT.EOF Then
        MsgBox "Error leyendo trabajadores / tmpDatosMEs" & vbCrLf & Err.Description, vbExclamation
        RT.Close
    Else
        'Ponemos datos
        '---------------------------------------------
        'ANTERIOR
        Me.txtHoras(0).Text = DBLet(RT!bolsahoras, "N")
        Me.txtBruto(0).Text = DBLet(RT!bolsabruto, "N")
        Me.txtNeto(0).Text = DBLet(RT!bolsaneto, "N")
    
    
        'PERIODO
        '-----------------------------------
        Me.txtHoras(1).Text = RT!horasc
        
        'En el tag estan las sumas de antrior + periodo
        txtHoras(1).Tag = DBLet(RT!bolsahoras, "N") + RT!horasc
        
        
        Importe1 = RT!horasc * DBLet(RT!Importe2, "N")
        Importe1 = Round(Importe1, 2)
        Me.txtBruto(1).Text = Importe1
        Me.txtBruto(1).Tag = Importe1 + DBLet(RT!bolsabruto, "N")
        
        'neto
        Imp2 = DBLet(RT!porcSS, "N") + DBLet(RT!porcirpf)
        Imp2 = Imp2 / 100
        
        'En el TAG guardo el imp2 para futuros calculos
        'Importe/hora
        txtHoras(0).Tag = DBLet(RT!Importe2, "N")
        txtBruto(0).Tag = Imp2
        
        
        Imp2 = Imp2 * Importe1
        Importe1 = Round(Importe1 - Imp2, 2)
        Me.txtNeto(1).Text = Round(Importe1, 2)
        Me.txtNeto(1).Tag = Importe1 + DBLet(RT!bolsaneto, "N")
        
        

        
        
        'Actual
        Me.txtHoras(2).Text = RT!bolsaperiodo
        Me.txtBruto(2).Text = RT!brutoperiodo
        Me.txtNeto(2).Text = RT!netoperiodo
        
        'Posterior
        Me.txtHoras(3).Text = RT!bolsadespues
        Me.txtBruto(3).Text = RT!brutodespues
        Me.txtNeto(3).Text = RT!netodespues
        
        'Dias comensados
        If RT!saldodias > 0 Then
            Me.txtDias.Tag = RT!saldodias
            Label9(7).Caption = "DIAS (max. " & RT!saldodias & ")"
            Me.txtDias.Text = RT!diasperiodo - RT!DiasTrabajados
            Me.txtDias.Enabled = True
        Else
            Label9(7).Caption = "DIAS"
            Me.txtDias.Enabled = False
            Me.txtDias.Tag = 0
            Me.txtDias.Text = 0
        End If
        RT.Close
        PonerFrameHoras True
    End If
    Set RT = Nothing

End Sub


Private Sub mnModificaHoras_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    If ListView1.SelectedItem.SmallIcon = 3 Then
        MsgBox "Trabajador de baja", vbExclamation
        Exit Sub
    End If
    Set NOD = ListView1.SelectedItem
    ModificarDesdeNodo
    
End Sub

Private Sub mnVerDatos_Click()
    ListView1_DblClick
End Sub

Private Sub Text1_GotFocus()
    With Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text1_LostFocus()
    Text1.Text = Trim(Text1.Text)
    If Text1.Text <> "" Then
        If Not IsNumeric(Text1.Text) Then
            MsgBox "Año debe ser numérico. (" & Text1.Text & ")", vbExclamation
            Text1.Text = ""
        End If
    End If
    If Text1.Text = "" Then Text1.Text = Year(Now)
End Sub


Private Sub InsertarJornadasSemanales(Fin As Date, FFin As Date)
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim Anticipos As Currency
Dim H As Currency
Dim DiferenciaBolsa As Currency


    'A partir de la diferencia de bolsa de horas sabremos
    'cuantas a comepensado de el mes anterior
    


    'Buscamos en JornadasSemanles
    '                       1                                               2
    SQL = "SELECT Sum(JornadasSemanales.DiasOfi) AS SumaDeDiasOfi, Sum(JornadasSemanales.HorasOfi) AS SumaDeHorasOfi,"
    '                       3                                   4
    SQL = SQL & " Sum(JornadasSemanales.HN) AS SumaDeHN, Sum(JornadasSemanales.Dias) AS SumaDeDias, JornadasSemanales.idTrabajador"
    SQL = SQL & " , Categorias.Importe1, Categorias.Importe2,Trabajadores.porcirpf,Trabajadores.porcSS,Sum(JornadasSemanales.HE) AS SumaExtras"
    SQL = SQL & " FROM Categorias INNER JOIN (JornadasSemanales INNER JOIN Trabajadores ON JornadasSemanales.idTrabajador ="
    SQL = SQL & "Trabajadores.IdTrabajador) ON Categorias.IdCategoria = Trabajadores.idCategoria"
    SQL = SQL & " Where (((JornadasSemanales.Fecha) >= #" & Format(Fin, FormatoFecha)
    SQL = SQL & "# And (JornadasSemanales.Fecha) <= #" & Format(FFin, FormatoFecha)
    SQL = SQL & " #))"
    SQL = SQL & " GROUP BY JornadasSemanales.idTrabajador, Categorias.Importe1, Categorias.Importe2,Trabajadores.porcirpf,Trabajadores.porcSS;"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    'Cadena insercion
    SQL = "INSERT INTO tmpDatosMes(Mes, Trabajador, MesHoras, MesDias,bolsaAntes, bolsaDespues, DiasTrabajados, HorasN, HorasC , "
    SQL = SQL & "Anticipos,HorasPeriodo,DiasPeriodo,HorasPLUS,Extras) VALUES (" & Combo1.ListIndex + 1 & ","
    
    'Trabajador, MesHoras, MesDias, DiasTrabajados, HorasN, HorasT,bolsaAntes, bolsaDespues
    'para cada semana
    While Not Rs.EOF
        'Dias y horas oficiles
        Cad = Rs!idTrabajador & ","
        Cad = Cad & TransformaComasPuntos(CStr(Rs.Fields(1))) & "," & Rs.Fields(0) & ","
    
    
        'La bolsa es la bolsa anterior al primera semana y la bolsa posterior es la posterior a la ultima semana
        Cad = Cad & BolsaHorasTipo4(Fin, FFin, Rs!idTrabajador, DiferenciaBolsa)
    
        If DiferenciaBolsa < 0 Then DiferenciaBolsa = 0
    
        'Trabajados
        Cad = Cad & "," & Rs.Fields(3) & ","
        Anticipos = Rs!SumaDeHN - DiferenciaBolsa
        
        'Cad = Cad & TransformaComasPuntos(CStr(RS.Fields(2))) & ","
        Cad = Cad & TransformaComasPuntos(CStr(Anticipos)) & ","
        
        'Compensables
        Cad = Cad & TransformaComasPuntos(CStr(DiferenciaBolsa)) & ","
        
        
        'Anticipos
        Anticipos = (Rs!SumaDeHN * Rs!Importe1) 'Como se pagan igual las N k las C entonces pongo la suma
        
        'Quitamos IRPF y SS
        H = ((Anticipos * Rs!porcirpf) + (Anticipos * Rs!porcSS)) / 100

        Anticipos = Anticipos - H
        Anticipos = Round((Anticipos), 2)
        Cad = Cad & TransformaComasPuntos(CStr(Anticipos)) & ","
        
        
        'Horas periodo, dias periodo
        Cad = Cad & TransformaComasPuntos(CStr(Rs.Fields(2))) & ","
        Cad = Cad & Rs.Fields(3) & ","
        
        'Horas plus
        Cad = Cad & TransformaComasPuntos(CStr(DBLet(Rs.Fields(9), "N")))
        
        
        'Lo k hayamos compensado
        Cad = Cad & "," & TransformaComasPuntos(CStr(DiferenciaBolsa))
        
        
        Cad = SQL & Cad & ")"
        InsertarEnTemporalTrabajador Cad, Rs!idTrabajador
        
        
    
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing

End Sub


Private Sub InsertarEnTemporalTrabajador(ByRef dSQL As String, idTrabajador As Long)
    On Error Resume Next
    conn.Execute dSQL
    If Err.Number <> 0 Then
        dSQL = "Error insertando el trabajador : " & idTrabajador & " . Entrada duplicada"
        MsgBox dSQL, vbExclamation
    End If
End Sub


Private Function BolsaHorasTipo4(F1 As Date, F2 As Date, T As Long, Diferencia As Currency) As String
Dim vCad As String
Dim RT As ADODB.Recordset
Dim H As Currency
Dim h2 As Currency
    
    BolsaHorasTipo4 = "0,0"
    Diferencia = 0
    vCad = "SELECT Bolsaantes,bolsadespues"
    vCad = vCad & " From JornadasSemanales"
    vCad = vCad & " Where JornadasSemanales.Fecha >= #" & Format(F1, FormatoFecha)
    vCad = vCad & "# And JornadasSemanales.Fecha <= #" & Format(F2, FormatoFecha)
    vCad = vCad & " # AND idTrabajador = " & T
    vCad = vCad & " ORDER BY fecha"
    
    
    
    Set RT = New ADODB.Recordset
    RT.Open vCad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        RT.MoveFirst
        'Primer registro bolsa antes
        H = 0
        If Not IsNull(RT!bolsaantes) Then H = RT!bolsaantes
        BolsaHorasTipo4 = TransformaComasPuntos(CStr(H))
        h2 = H
        
        'Ultimo registro bolsa despues
        RT.MoveLast
        H = 0
        If Not IsNull(RT!bolsadespues) Then H = RT!bolsadespues
        BolsaHorasTipo4 = BolsaHorasTipo4 & "," & TransformaComasPuntos(CStr(H))
        Diferencia = h2 - H
       
    End If
    
    RT.Close
    Set RT = Nothing
End Function


Private Sub CalculaEntreFechasALZIRA(FI As Date, FF As Date)
Dim Rs As Recordset
Dim Horas As Currency
Dim Dias As Integer
Dim Seccion As Integer

    
    conn.Execute "DELETE FROM tmpHorasMesHorario"

    'Para comprobar si estando de baja han trabajado
    'En tmpPresencia voy a guardar
    conn.Execute "DELETE FROM tmpCombinada"

    Set Rs = New ADODB.Recordset
    Rs.Open "horarios", conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    
    Label1.Caption = "Obtener horarios"
    Label1.Refresh
    
    While Not Rs.EOF
        Horas = CalculaHorasHorarioALZ(Rs.Fields(0), Dias, FI, FF)
        If Horas > 0 Then
            'Insertamos en tmp HORAS
            conn.Execute "INSERT INTO tmpHorasMesHorario(idHorario,Horas,Dias) VALUES (" & Rs.Fields(0) & "," & TransformaComasPuntos(CStr(Horas)) & "," & Dias & ")"
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    Seccion = Me.cboSeccion.ItemData(cboSeccion.ListIndex)
    
    Label1.Caption = "Horas trabajadas"
    Label1.Refresh
    CalculaHorasTrabajadas FI, FF, 0, Seccion
    Me.Refresh
    

    
    Label1.Caption = "Datos periodo"
    Label1.Refresh
    CalculaDatosMes FI, FF, 0, Seccion
    
    Me.Refresh
    
    Label1.Caption = "Combina datos"
    Label1.Refresh
    CombinaDatos FI, FF
    
        
    'Ajustamos anticpos
    Label1.Caption = "Obtener anticipos"
    Label1.Refresh
    ObtenerAnticposPagadosPorPrograma FI, FF
    
    
    Label1.Caption = "Calcula diferencias dias/horas"
    Label1.Refresh
    CalculaDiferenciasDiasHoras

    Label1.Caption = "Pre-ajuste BOLSA"
    Label1.Refresh
    ValoresBolsaDespues
    
    'Ajustamos los que no hayan trabakado nada
    AjustaDatosBajaMesEntero
    


    


    'NUEVO 20 OCTUBRE
    'Metemos los datos de los k trabajan en jornads semanales de OCHO horas
    'Control Nomina=4
    Label1.Caption = "Joranadas semanales"
    Label1.Refresh
    InsertarJornadasSemanales FI, FF





    'Ajustamos los de jornadas semanales

    Label1.Caption = "Carga datos"
    Label1.Refresh
    CargaDatos



    'Ahora vamos a comprobar si alguno de los k ha estado de baja
    'En este periodo a trabajado
    If ListView1.ListItems.Count > 0 Then
        Label1.Caption = "Comprobar bajas con dias Tra."
        Label1.Refresh
        Rs.Open "Select * from tmpcombinada", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            If HaTrabajadoConBaja(Rs) Then
                Dias = 0
                Do
                    Dias = Dias + 1
                    If Dias <= ListView1.ListItems.Count Then
                        If ListView1.ListItems(Dias).Text = Rs!idTrabajador Then
                            'Pongo el icono distinto
                            ListView1.ListItems(Dias).SmallIcon = 5
                            'Salgo
                            Dias = 32000
                        End If
                    End If
                Loop Until Dias > ListView1.ListItems.Count
            End If   'De ha trabajado estando de baja
            'Siguiente caso
        Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    Set Rs = Nothing
End Sub

Private Sub CalcularSumasHorasNETOBRUTO()
Dim Im2 As Currency


    'Horas
    '------------------------------------
    Im2 = ImporteText(txtHoras(0))
    Importe1 = Im2
    Im2 = ImporteText(txtHoras(1))
    Importe1 = Importe1 + Im2
    txtHoras(1).Tag = Importe1
    'menos las utilizadas
    Im2 = ImporteText(txtHoras(2))
    Importe1 = Importe1 - Im2
    txtHoras(3).Text = Importe1
    
    
    'BRUTO
    '------------------------------------
    Im2 = ImporteText(txtBruto(0))
    Importe1 = Im2
    Im2 = ImporteText(txtBruto(1))
    Importe1 = Importe1 + Im2
    
    'lo guardamos en el tag
    txtBruto(1).Tag = Importe1
    
    'menos las utilizadas
    Im2 = ImporteText(txtBruto(2))
    Importe1 = Importe1 - Im2
    txtBruto(3).Text = Importe1
    
    
    'NETO
    '------------------------------------
    Im2 = ImporteText(txtNeto(0))
    Importe1 = Im2
    Im2 = ImporteText(txtNeto(1))
    Importe1 = Importe1 + Im2
    'lo guardamos en el tag
    txtNeto(1).Tag = Importe1
    
    
    'menos las utilizadas
    Im2 = ImporteText(txtNeto(2))
    Importe1 = Importe1 - Im2
    txtNeto(3).Text = Importe1
    
        
    
End Sub


Private Function ImporteText(ByRef T As TextBox) As Currency
    ImporteText = 0
    If T.Text <> "" Then
        ImporteText = CCur(TransformaPuntosComas(T.Text))
    End If
End Function



Private Sub Text2_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub



Private Sub txtDias_GotFocus()
    txtDias.SelStart = 0
    txtDias.SelLength = Len(txtDias.Text)
End Sub

Private Sub Txtdias_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Txtdias_LostFocus()
    If txtDias.Text <> "" Then
        If Not IsNumeric(txtDias.Text) Then
            MsgBox "Campo debe ser numérico", vbExclamation
            CompensaDias
        Else
            txtDias.Text = Val(txtDias.Text)
            If Val(txtDias.Text) > txtDias.Tag Then
                MsgBox "Máximo dias a comepensar: " & txtDias.Tag, vbExclamation
                CompensaDias
            End If
        End If
    End If
End Sub

Private Sub txtHoras_GotFocus(Index As Integer)
    If Index = 2 Then
        txtHoras(2).SelStart = 0
        txtHoras(2).SelLength = Len(txtHoras(2).Text)
    End If
End Sub

Private Sub txtHoras_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtHoras_LostFocus(Index As Integer)
Dim Imp2 As Currency

    If Index = 2 Then
        txtHoras(2).Text = Trim(txtHoras(2).Text)
        If txtHoras(2).Text = "" Then
            txtHoras(2).Text = 0
        Else
            If Not IsNumeric(txtHoras(2).Text) Then
                MsgBox "Campo no numerico o con formato incorrecto: " & txtHoras(2).Text, vbExclamation
                txtHoras(2).Text = "0"
            End If
        End If
        txtHoras(2).Tag = TransformaPuntosComas(txtHoras(2).Text)
        If txtHoras(2).Text = "0" Then
            txtBruto(2).Text = "0"
            txtNeto(2).Text = "0"
        Else
            
            'Calculo los importes de bruto
            Imp2 = txtBruto(1).Tag / txtHoras(1).Tag
            Importe1 = Round(CCur(txtHoras(2).Tag) * Imp2, 2)
            
            txtBruto(2).Text = Importe1
                
            'NETO
            Imp2 = txtNeto(1).Tag / txtHoras(1).Tag
            Importe1 = Round(CCur(txtHoras(2).Tag) * Imp2, 2)
            
            txtNeto(2).Text = Importe1
           
        End If
            
        'Y ponemos los dias si hay k poner
        CompensaDias
        
        'Racalculamos
        CalcularSumasHorasNETOBRUTO
    End If
End Sub


Private Sub CompensaDias()
    If txtDias.Tag > 0 Then
        If HorasxDia2 > 0 Then
            Importe1 = txtHoras(2).Tag \ HorasxDia2
            If Importe1 > Val(txtDias.Tag) Then Importe1 = Val(txtDias.Tag)
            txtDias.Text = Importe1
        Else
            txtDias.Text = 0
        End If
    Else
        txtDias.Text = 0
    End If
End Sub

'Le hemos
Private Sub ModificaDatosTrabajador()
Dim RT As ADODB.Recordset

    
    SQL = "UPDATE tmpDatosMES SET bolsaperiodo=" & TransformaComasPuntos(txtHoras(2).Text)
    SQL = SQL & " , brutoperiodo =" & TransformaComasPuntos(txtBruto(2).Text)
    SQL = SQL & " , netoperiodo =" & TransformaComasPuntos(txtNeto(2).Text)
    SQL = SQL & " , brutodespues =" & TransformaComasPuntos(txtBruto(3).Text)
    SQL = SQL & " , netodespues =" & TransformaComasPuntos(txtNeto(3).Text)
    SQL = SQL & " , bolsadespues =" & TransformaComasPuntos(txtHoras(3).Text)
    'EXTRAS
    
    
    'Dias trabajados
    If HorasxDia2 > 0 Then
        If txtDias.Text <> "" Then
            Importe1 = CCur(txtDias.Text)
            SQL = SQL & " , diasperiodo = diastrabajados + " & Importe1
        End If
    End If
    
    
    'El importe del PLUS
    SQL = SQL & " , plus =" & TransformaComasPuntos(txtNeto(2).Text)
    SQL = SQL & " , brutoplus =" & TransformaComasPuntos(txtBruto(2).Text)
    
    SQL = SQL & " WHERE Trabajador =" & NOD.Text
    conn.Execute SQL
    espera 0.5
    
    SQL = ""
    PonSQL NOD.Text
    Set RT = New ADODB.Recordset
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RT.EOF Then
        MsgBox "Error grave", vbCritical
    Else
        PonLinea NOD, RT
    End If
    RT.Close
    Set RT = Nothing
End Sub

Private Function genNomALZIRA() As Boolean
Dim i As Integer
Dim Cad As String
Dim Importe As Currency
Dim Rs As ADODB.Recordset

On Error GoTo EGenerarNominas
    genNomALZIRA = False

    SQL = "INSERT INTO Nominas (Fecha,IdTrabajador,Dias,HN,HC,Plus,HP,Anticipos,BolsaDespues,BolsaAntes"
    SQL = SQL & ",brutodespues,netodespues,importedelbote,brutoantes,netoantes) VALUES (#"
    i = DiasMes(Combo2.ListIndex + 1, CInt(Text2.Text))
    SQL = SQL & Text2.Text & "/" & Combo2.ListIndex + 1 & "/" & i & "#,"
    
    'Primero generamos la tabla de  nominas con los importes marcados aqui
    Cad = "SELECT tmpDatosMEs.*, Trabajadores.bolsaBRUTO, Trabajadores.bolsaNETO"
    Cad = Cad & " FROM tmpDatosMEs INNER JOIN Trabajadores ON tmpDatosMEs.Trabajador = Trabajadores.IdTrabajador"
    Cad = Cad & " Where Mes = " & Combo2.ListIndex + 1
    Cad = Cad & " ORDER BY tmpDatosMes.Trabajador"
    Set Rs = New ADODB.Recordset
    Rs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not Rs.EOF
        'IdTrabajador,Dias,HN,
        Cad = Rs!Trabajador & "," & Rs!diasperiodo & "," & TransformaComasPuntos(Rs!horasn)
        
        'HC,Plus,HP,Anticipos
        Importe1 = Rs!Anticipos + Rs!plus
        'Cad = Cad & "," & TransformaComasPuntos(RS!bolsaperiodo) & ",0,0," & TransformaComasPuntos(CStr(Importe1)) & ","
        Cad = Cad & ",0," & TransformaComasPuntos(DBLet(Rs!brutoplus, "N"))
        Cad = Cad & "," & TransformaComasPuntos(CStr(Rs!bolsaperiodo))
        Cad = Cad & "," & TransformaComasPuntos(CStr(Importe1)) & ","
        
        'Bolsadespues
        Cad = Cad & TransformaComasPuntos(Rs!bolsadespues) & "," & TransformaComasPuntos(Rs!bolsaantes)
        
        'brut despues, neto despues
        Cad = Cad & "," & TransformaComasPuntos(Rs!brutodespues) & "," & TransformaComasPuntos(Rs!netodespues)
        
        'importedelbote: cuanto de los anticipos ha sido cojido del bote
        Cad = Cad & "," & TransformaComasPuntos(Rs!plus)
        
        'brutoantes,netoantes
        Cad = Cad & "," & TransformaComasPuntos(DBLet(Rs!bolsabruto, "N")) & "," & TransformaComasPuntos(DBLet(Rs!bolsaneto, "N"))
        
        Cad = Cad & ")"
        Cad = SQL & Cad
        conn.Execute Cad
        
        
        
        
        'Pondremos la bolsa de horas Y, hay bajas,
        'entonces actualizaremos la baja de cada trabajador
        'al ultimo dia trabajado
        Cad = "UPDATE Trabajadores SET Bolsahoras = " & TransformaComasPuntos(Rs!bolsadespues)
        Cad = Cad & ",Bolsaneto = " & TransformaComasPuntos(Rs!netodespues)
        Cad = Cad & ",BolsaBruto = " & TransformaComasPuntos(Rs!brutodespues)
        Cad = Cad & " WHERE idTrabajador = " & Rs!Trabajador
        conn.Execute Cad
        
        
        

        
        'Sig
        Rs.MoveNext
    Wend
    

    
    Rs.Close
    
    
    
    
    genNomALZIRA = True
    Exit Function
EGenerarNominas:
    MuestraError Err.Number, Err.Description
    Set Rs = Nothing
End Function

