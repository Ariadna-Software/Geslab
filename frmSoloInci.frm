VERSION 5.00
Begin VB.Form frmSoloInci 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introducción incidencias autmáticas"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmSoloInci.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   4
      Left            =   1920
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1920
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2520
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   315
      Index           =   1
      Left            =   3840
      TabIndex        =   4
      Top             =   3060
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   315
      Index           =   0
      Left            =   2460
      TabIndex        =   3
      Top             =   3060
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2220
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   1260
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1260
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   0
      Left            =   1500
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Horas decimal"
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
      Index           =   3
      Left            =   180
      TabIndex        =   11
      Top             =   2580
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Horas sexagesimal"
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
      Left            =   180
      TabIndex        =   10
      Top             =   1920
      Width           =   1755
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1140
      Picture         =   "frmSoloInci.frx":030A
      Top             =   1320
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00004080&
      Caption         =   "Nuevo marcaje"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   240
      TabIndex        =   9
      Top             =   120
      Width           =   4755
   End
   Begin VB.Label Label1 
      Caption         =   "Incidencia"
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
      Index           =   1
      Left            =   180
      TabIndex        =   7
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Empleado"
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
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   660
      Width           =   915
   End
End
Attribute VB_Name = "frmSoloInci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Public Event Seleccionar(vInci As Integer, vhoras As Single)

Public Inci As Integer
Public CadInci As String
Public Nombre As String
Public Horas As Single

'Private PrimeraVez As Boolean
Private SQL As String

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
    If DatosOk Then
        RaiseEvent Seleccionar(CInt(Text1(1).Text), Text1(3).Text)
        Else
        Exit Sub
    End If
End If
Unload Me
' MiDoble1, MiDoble2 son Doubles.
'MiDoble1 = 75.3421115: MiDoble2 = 75.3421555
'MiSimple1 = CSng(MiDoble1)   ' MiSimple1 contiene 75.34211.
'MiSimple2 = CSng(MiDoble2)   ' MiSimple2 contiene 75.34216.
End Sub

Private Sub Form_Activate()
    'If PrimeraVez Then
    '    PrimeraVez = False
    'End If
End Sub

Private Sub Form_Load()
    Text1(0).Text = Nombre
    Text1(1).Text = Inci
    Text1(2).Text = CadInci
    Text1(3).Text = Horas
    'En horas tenemos las horas en decimal
    Text1(4).Text = DevuelveHora(CSng(Horas))
    If Inci = 0 Then
        Label2.Caption = "Nueva incidencia"
    Else
        Label2.Caption = "Modificar"
    End If
    Text1(1).Enabled = (Inci = 0)
    'El sql sirve para cuando introduzcamos desde teclado la incidencia
    SQL = "SELECT NomInci FROM Incidencias WHERE IdInci="
    'PrimeraVez = True
End Sub


Private Function DatosOk() As Boolean
DatosOk = False

'Ahora la incidencia
If Text1(1).Text = "" Then
    MsgBox "Seleccione una incidencia.", vbExclamation
    Exit Function
End If

If Not IsNumeric(Text1(1).Text) Then
    MsgBox "Número de incidencia incorrecto.", vbExclamation
    Exit Function
End If

If CInt(Text1(1).Text) < 0 Then
    MsgBox "Número de incidencia incorrecta.", vbExclamation
    Exit Function
End If


'Numero de horas
If Text1(3).Text = "" Then
    Text1(3).Tag = 0
    Else
    If Not IsNumeric(Text1(3).Text) Then
        MsgBox "Número de horas incorrecto.", vbExclamation
        Exit Function
        Else
            Text1(3).Tag = Text1(3).Text
    End If
End If
DatosOk = True
End Function

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
Text1(1).Text = vCodigo
Text1(2).Text = vCadena
End Sub

Private Sub Image1_Click()
    Set frmB = New frmBusca
    frmB.Tabla = "Incidencias"
    frmB.CampoBusqueda = "NomInci"
    frmB.CampoCodigo = "IdInci"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "INCIDENCIAS"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub


Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim I As Integer
Dim RS As ADODB.Recordset

Select Case Index
Case 1
    If Text1(1).Text = "" Then Exit Sub
    If Not IsNumeric(Text1(1).Text) Then
        MsgBox "La incidencia tiene que ser un número."
        Text1(1).Text = -1
        Text1(2).Text = ""
        Text1(1).SetFocus
        Exit Sub
    End If
    'Incidencia
    Set RS = New ADODB.Recordset
    RS.Open SQL & Text1(1).Text, Conn, , , adCmdText
    If RS.EOF Then
        Text1(1).Text = -1
        Text1(2).Text = ""
        Text1(1).SetFocus
        Else
            Text1(2).Text = RS.Fields(0)
    End If
    RS.Close
    Set RS = Nothing
Case 3
    I = InStr(1, Text1(3).Text, ".")
    If I > 0 Then
        Text1(3).Text = Mid(Text1(3).Text, 1, I - 1) & "," & Mid(Text1(3).Text, I + 1)
    End If
    If IsNumeric(Text1(3).Text) Then _
        Text1(4).Text = DevuelveHora(CSng(Text1(3).Text))
    
Case 4
    Do
        I = InStr(1, Text1(4).Text, ".")
        If I > 0 Then
            Text1(4).Text = Mid(Text1(4).Text, 1, I - 1) & ":" & Mid(Text1(4).Text, I + 1)
        End If
    Loop Until I = 0
    
    If IsDate(Text1(4).Text) Then
        Text1(4).Text = Format(Text1(4).Text, "hh:mm:ss")
        Text1(3).Text = DevuelveValorHora(CDate(Text1(4).Text))
    End If
End Select
End Sub
