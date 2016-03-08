VERSION 5.00
Begin VB.Form frmPrenominaIndivi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Horas trabajador"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   6045
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   9
      Left            =   3000
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2880
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   8
      Left            =   1920
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   7
      Left            =   3000
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1995
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   6
      Left            =   1920
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   1995
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   5
      Left            =   5160
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   4080
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   3000
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1320
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "Dias compensados"
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
      Left            =   120
      TabIndex        =   19
      Top             =   3000
      Width           =   1695
   End
   Begin VB.Label Label9 
      Caption         =   "Dias compensados"
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
      Left            =   120
      TabIndex        =   18
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label Label8 
      Caption         =   "SALDO"
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
      Left            =   120
      TabIndex        =   15
      Top             =   2160
      Width           =   1230
   End
   Begin VB.Line Line1 
      X1              =   480
      X2              =   6000
      Y1              =   1800
      Y2              =   1800
   End
   Begin VB.Label Label7 
      Caption         =   "H.E."
      Height          =   255
      Left            =   5400
      TabIndex        =   13
      Top             =   1080
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "H.N."
      Height          =   255
      Left            =   4320
      TabIndex        =   12
      Top             =   1080
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "Horas"
      Height          =   195
      Left            =   3000
      TabIndex        =   11
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label4 
      Caption         =   "Dias"
      Height          =   195
      Left            =   1920
      TabIndex        =   10
      Top             =   600
      Width           =   315
   End
   Begin VB.Label Label3 
      Caption         =   "TRABAJADOS"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1230
   End
   Begin VB.Label Label2 
      Caption         =   "OFICIAL"
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
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   720
   End
   Begin VB.Label Label1 
      Caption         =   "DATOS TRABAJADOR"
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
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "frmPrenominaIndivi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

End Sub
