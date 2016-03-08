VERSION 5.00
Begin VB.Form frmSoloHora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introducción marcajes"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmSoloHora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   5235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   315
      Index           =   1
      Left            =   3720
      TabIndex        =   6
      Top             =   2280
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   315
      Index           =   0
      Left            =   2400
      TabIndex        =   5
      Top             =   2280
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   2220
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1680
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1680
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   660
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   840
      Width           =   1155
   End
   Begin VB.Label Label3 
      Caption         =   "Puedes  poner los dos puntos de las horas con el punto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   8
      Top             =   1140
      Width           =   4095
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1140
      Picture         =   "frmSoloHora.frx":030A
      Top             =   1740
      Width           =   240
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
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
      Height          =   495
      Left            =   240
      TabIndex        =   7
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
      TabIndex        =   2
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Hora"
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
      TabIndex        =   0
      Top             =   900
      Width           =   735
   End
End
Attribute VB_Name = "frmSoloHora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Public Event Seleccionar(vHora As Date, vInci As Integer)

Public Hora As String
Public Inci As Integer
Public CadInci As String

Private PrimeraVez As Boolean


Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
    If DatosOk Then
        RaiseEvent Seleccionar(CDate(Text1(0).Text), CInt(Text1(1).Text))
        Else
        Exit Sub
    End If
End If
Unload Me
End Sub

Private Sub Form_Activate()
If PrimeraVez Then
    PrimeraVez = False
    Me.Top = 1000
End If
End Sub

Private Sub Form_Load()
Text1(0).Text = Hora
Text1(1).Text = Inci
Text1(2).Text = CadInci
If Hora = "" Then
    Label2.Caption = "Nuevo marcaje"
    Else
    Label2.Caption = "Modificar"
End If

PrimeraVez = True
End Sub


Private Function DatosOk() As Boolean
DatosOk = False
If Text1(0).Text = "" Then
    MsgBox "Escriba una fecha", vbExclamation
    Exit Function
End If

If Not IsDate(Text1(0).Text) Then
    MsgBox "No es una fecha válida", vbExclamation
    Exit Function
End If
'Compruebo que en la cadena hay dos puntos
If InStr(1, Text1(0).Text, ":") = 0 Then
    MsgBox "No es un hora válida", vbExclamation
    Exit Function
End If

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
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub


Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim i As Integer
Dim C As String

Select Case Index
Case 0
    Do
        i = InStr(1, Text1(0).Text, ".")
        If i > 0 Then
            C = Mid(Text1(0).Text, i + 1)
            If Len(C) = 1 Then
                If Val(C) > 5 Then
                    C = "0" & C
                Else
                    C = C & "0"
                End If
            End If
            Text1(0).Text = Mid(Text1(0).Text, 1, i - 1) & ":" & C
        End If
    Loop While i <> 0
    
    If Text1(0).Text <> "" Then
        If Not IsDate(Text1(0).Text) Then
            MsgBox "Error en el campo hora: " & Text1(0).Text, vbExclamation
            Text1(0).Text = ""
            Text1(0).SetFocus
    
        End If
    End If
Case 1
    If Text1(1).Text = "" Then Exit Sub
    If Not IsNumeric(Text1(1).Text) Then
        MsgBox "La incidencia tiene que ser un número.", vbExclamation
        Text1(1).Text = -1
        Text1(2).Text = ""
        Text1(1).SetFocus
        Exit Sub
    End If
    'Incidencia
    C = DevuelveDesdeBD("nominci", "incidencias", "idinci", Text1(1).Text, "N")
    
    If C = "" Then
        
        Text1(2).Text = "NO EXISTE :" & Text1(1).Text
        Text1(1).Text = 0
        Text1(1).SetFocus
        Else
            Text1(2).Text = C
    End If
    
End Select

End Sub

Private Sub Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'ENTER
        KeyAscii = 0
        SendKeys "{tab}"
    ElseIf KeyAscii = 27 Then 'ESC
         Unload Me
    End If
End Sub
