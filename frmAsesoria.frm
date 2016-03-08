VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmAsesoria 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asesoria"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Framecrear 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      Begin VB.CommandButton Command1 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   5
         Top             =   2340
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2220
         TabIndex        =   4
         Top             =   2340
         Width           =   975
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Traspasar festivos trabajados"
         Height          =   255
         Left            =   600
         TabIndex        =   3
         Top             =   1800
         Width           =   2475
      End
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   2280
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1140
         Width           =   855
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   1140
         Width           =   1455
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Crear fichero"
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
         Height          =   435
         Left            =   180
         TabIndex        =   13
         Top             =   180
         Width           =   4275
      End
      Begin VB.Label Label1 
         Caption         =   "Periodo"
         Height          =   195
         Left            =   600
         TabIndex        =   6
         Top             =   840
         Width           =   1035
      End
   End
   Begin VB.Frame FrameImportar 
      Height          =   3015
      Left            =   60
      TabIndex        =   7
      Top             =   0
      Width           =   4575
      Begin MSComDlg.CommonDialog cd1 
         Left            =   660
         Top             =   2460
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   120
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1200
         Width           =   4275
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   3
         Left            =   3240
         TabIndex        =   9
         Top             =   2340
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   2
         Left            =   2100
         TabIndex        =   8
         Top             =   2340
         Width           =   975
      End
      Begin VB.Image ImgFich 
         Height          =   240
         Left            =   840
         Picture         =   "frmAsesoria.frx":0000
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Fichero"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Importar"
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
         Height          =   435
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   4275
      End
   End
End
Attribute VB_Name = "frmAsesoria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public OPCION As Byte
    '0.- Generar fichero
    '1.- Importar fichero aseoria

Private Sub CargaCombo()
Dim i As Integer

    For i = 1 To 12
        Combo1.AddItem Format(CDate("01/" & i & "/2000"), "mmmm")
    Next i
    Combo1.ListIndex = Month(DateAdd("m", -1, Now)) - 1
    Text1.Text = Year(DateAdd("m", -1, Now))
End Sub

Private Sub Command1_Click(Index As Integer)
    If Index = 1 Or Index = 3 Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    FrameImportar.Visible = False
    Framecrear.Visible = (OPCION = 0)
    FrameImportar.Visible = Not Framecrear.Visible
    If OPCION = 0 Then
        CargaCombo
    Else
        Text2.Text = ""
    End If
End Sub

Private Sub ImgFich_Click()
    On Error GoTo EF
        cd1.CancelError = True
        cd1.DialogTitle = "Fichero asesoria ..."
        cd1.ShowOpen
        Text2.Text = cd1.FileName
    Exit Sub
    
EF:
    Err.Clear
End Sub
