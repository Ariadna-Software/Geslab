VERSION 5.00
Begin VB.Form frmRectificar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rectificar marcajes"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5235
   Icon            =   "frmRectificar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   3660
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1140
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1140
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Cancelar"
      Height          =   315
      Index           =   1
      Left            =   3660
      TabIndex        =   4
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   315
      Index           =   0
      Left            =   2340
      TabIndex        =   3
      Top             =   1800
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "horas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   8
      Top             =   180
      Width           =   4395
   End
   Begin VB.Label Label1 
      Caption         =   "Hora modificada"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   840
      Width           =   1395
   End
   Begin VB.Label Label1 
      Caption         =   "Hora fin"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   6
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Hora inicio"
      Height          =   255
      Index           =   0
      Left            =   180
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
End
Attribute VB_Name = "frmRectificar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Nuevo As Boolean
Public IdSeccion As Integer
Public Id As Long

Private Sub Command1_Click(Index As Integer)
Dim i As Integer
Dim Rs As ADODB.Recordset
Dim Cad As String

On Error GoTo Final
If Index = 1 Then GoTo Final


For i = 0 To 2
    If Trim(Text1(i).Text) = "" Then
        MsgBox Label1(i).Caption & " NO puede estar vacia.", vbExclamation
        Exit Sub
    End If
Next i
'Comprobamos que son fechas
For i = 0 To 2
    If Not IsDate(Text1(Index).Text) Then
        MsgBox Label1(i).Caption & " NO es una fecha correcta", vbExclamation
    End If
Next i

i = (CDate(Text1(1).Text) >= CDate(Text1(0).Text))
If i = 0 Then
    MsgBox "Hora inicio <= Hora fin ", vbExclamation
    Exit Sub
End If
'Llegados aqui modificamos o insertamos
Set Rs = New ADODB.Recordset
Rs.CursorType = adOpenKeyset
Rs.LockType = adLockOptimistic
If Nuevo Then

    'Calculamos el ultimo
    i = 1
    Cad = "Select max(id) from ModificarFichajes"
    Rs.Open Cad, Conn, , , adCmdText
    If Not Rs.EOF Then
       i = DBLet(Rs.Fields(0), "N") + 1
    End If
    Rs.Close
    Cad = "Select * from ModificarFichajes"
    Rs.Open Cad, Conn, , , adCmdText
    Rs.AddNew
    Rs!Id = i
    Rs!IdSeccion = IdSeccion
    Rs!inicio = Text1(0).Text
    Rs!Fin = Text1(1).Text
    Rs!modificada = Text1(2).Text
    Rs.Update
    
    
    'ELSE es modificar
    Else
        Cad = "Select * from ModificarFichajes"
        Cad = Cad & " WHERE id=" & Id
        Rs.Open Cad, Conn, , , adCmdText
        Rs!inicio = Text1(0).Text
        Rs!Fin = Text1(1).Text
        Rs!modificada = Text1(2).Text
        Rs.Update
 End If
Rs.Close
Set Rs = Nothing


Final:
    If Err.Number <> 0 Then MuestraError Err.Number
    Unload Me
End Sub

Private Sub Form_Activate()
 Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
If Nuevo Then
    Label2.Caption = "Nueva rectificación"
    Else
    Label2.Caption = "Modificar rectificación"
End If

End Sub



Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim i As Integer
Do
    i = InStr(1, Text1(Index).Text, ".")
    If i > 0 Then
        Text1(Index).Text = Mid(Text1(Index).Text, 1, i - 1) & ":" & Mid(Text1(Index).Text, i + 1)
    End If
Loop Until i = 0

'Si es hora correcta la formateamos, sino lo ponemos en blanco
If IsDate(Text1(Index).Text) Then
    Text1(Index).Text = Format(Text1(Index).Text, "hh:mm")
    Else
        Text1(Index).Text = ""
End If
End Sub
