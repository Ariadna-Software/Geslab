VERSION 5.00
Begin VB.Form frmVarios 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6090
   Icon            =   "frmVarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrTrabajadores 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   5895
      Begin VB.Frame Frame2 
         Height          =   1455
         Left            =   1800
         TabIndex        =   9
         Top             =   1680
         Width           =   3615
         Begin VB.OptionButton optTrab3 
            Caption         =   "Todos los registros seleccionados"
            Height          =   255
            Index           =   1
            Left            =   360
            TabIndex        =   11
            Top             =   840
            Width           =   2895
         End
         Begin VB.OptionButton optTrab3 
            Caption         =   "Registro actual"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   3015
         End
      End
      Begin VB.CommandButton cmdListaTrabajadores 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   3480
         TabIndex        =   6
         Top             =   3360
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   4680
         TabIndex        =   5
         Top             =   3360
         Width           =   975
      End
      Begin VB.Frame Frame1 
         Height          =   735
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   3615
         Begin VB.OptionButton optTraba2 
            Caption         =   "Código"
            Height          =   255
            Index           =   1
            Left            =   2520
            TabIndex        =   8
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton optTraba2 
            Caption         =   "Tarjeta"
            Height          =   255
            Index           =   0
            Left            =   1200
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.Label Label2 
            Caption         =   "Ordenar:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.OptionButton optTraba1 
         Caption         =   "Tarjetas"
         Height          =   255
         Index           =   1
         Left            =   360
         TabIndex        =   2
         Top             =   1920
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton optTraba1 
         Caption         =   "Normal"
         Height          =   255
         Index           =   0
         Left            =   360
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Listado trabajadores"
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
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Byte
        
Private Sub cmdCancelar_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdListaTrabajadores_Click()
    If Me.optTraba1(0).Value Then
        
        VariableCompartida = "1|" & Abs(optTraba2(0).Value) & "|"
    Else
        'Tarjetas
         VariableCompartida = "2|" & Abs(optTrab3(0).Value) & "|"
    End If
    Unload Me
End Sub

        '   1.- Listado de trabajadores
        
        
        
Private Sub Form_Load()
    
    
    Select Case Opcion
    Case 1
        Me.Caption = "Imprimir trabajadores"
    
    End Select
End Sub
