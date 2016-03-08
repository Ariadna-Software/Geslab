VERSION 5.00
Begin VB.Form frmVarios2 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6750
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameErrorGrabando 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cerrar"
         Height          =   375
         Index           =   0
         Left            =   5280
         TabIndex        =   2
         Top             =   5040
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   4695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Text            =   "frmVarios2.frx":0000
         Top             =   240
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmVarios2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal1.Icon
    Caption = "Listado"
    FrameErrorGrabando.Visible = True
    Me.cmdCancel(0).Cancel = True
End Sub
