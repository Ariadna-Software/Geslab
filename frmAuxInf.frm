VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAuxInf 
   Caption         =   "Formulario de informes"
   ClientHeight    =   4875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9435
   Icon            =   "frmAuxInf.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4875
   ScaleWidth      =   9435
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox Check1 
      Caption         =   "PREVISUALIZAR"
      Height          =   315
      Left            =   4140
      TabIndex        =   7
      Top             =   4260
      Width           =   1575
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2655
      Left            =   180
      TabIndex        =   5
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   4683
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "FECHAS"
      Height          =   1215
      Left            =   180
      TabIndex        =   0
      Top             =   3360
      Width           =   3795
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   2100
         TabIndex        =   3
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   0
         Left            =   300
         TabIndex        =   1
         Top             =   660
         Width           =   1335
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   2400
         Picture         =   "frmAuxInf.frx":030A
         Top             =   420
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   780
         Picture         =   "frmAuxInf.frx":040C
         Top             =   420
         Width           =   240
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fin"
         Height          =   195
         Index           =   1
         Left            =   2100
         TabIndex        =   4
         Top             =   420
         Width           =   210
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Inicio"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   420
         Width           =   375
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   6
      Top             =   180
      Width           =   1080
   End
End
Attribute VB_Name = "frmAuxInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
CargaListview
End Sub




Private Sub CargaListview()
On Error GoTo ECargaListview




Exit Sub
ECargaListview:
    MuestraError Err.Number, "Cargando lista informes"
End Sub



