VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmEliminar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Eliminar datos"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4620
   Icon            =   "frmEliminar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4620
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar Pb1 
      Height          =   375
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   1000
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3420
      TabIndex        =   0
      Top             =   2760
      Width           =   975
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Index           =   0
      Left            =   720
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2160
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2340
      TabIndex        =   4
      Top             =   2760
      Width           =   975
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   1620
      Picture         =   "frmEliminar.frx":0442
      Top             =   1860
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   3300
      Picture         =   "frmEliminar.frx":0544
      Top             =   1860
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   $"frmEliminar.frx":0646
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   240
      TabIndex        =   3
      Top             =   180
      Width           =   4215
   End
   Begin VB.Label Label3 
      Caption         =   "Fecha inicio"
      Height          =   255
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label4 
      Caption         =   "Fecha fin"
      Height          =   255
      Left            =   2400
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmEliminar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Dim Donde As String
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Tamanyo As Long
Dim Contador As Long

Private Sub Command1_Click()
Dim Cad As String
    'Primero comprobamos k ha puesto als dos fechas
    Text5(0).Text = Trim(Text5(0).Text)
    Text5(1).Text = Trim(Text5(1).Text)
    If Text5(0).Text = "" Or Text5(1).Text = "" Then
        MsgBox "Escriba las fechas del intervalo de supresion de datos.", vbExclamation
        Exit Sub
    End If
    
    
    
    '----------------------------------------
    'Pregunta del millon
    'David
    Cad = "Escriba la clave de operaciones para el borrado."
    Cad = InputBox(Cad)
    If Cad <> "" Then
        If UCase(Cad) = "ARIADNA" Then
            'Hacemos borrado
            HacerBorrado
        Else
            MsgBox "Lo sentimos mucho, pero no es correcta la clave", vbExclamation
        End If
    End If
        
            
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Pb1.Visible = False
    Text5(1).Text = Format(Now, "dd/mm/yyyy")
    Text5(0).Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text5(Val(Text5(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()

End Sub

Private Sub Image2_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Text5(Index).Text <> "" Then
        If IsDate(Text5(Index).Text) Then frmF.Fecha = CDate(Text5(Index).Text)
    End If
    Text5(0).Tag = Index
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub Text5_GotFocus(Index As Integer)
    GotFocus Text5(Index)
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub


Private Sub Text5_LostFocus(Index As Integer)
    With Text5(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not EsFechaOK(Text5(Index)) Then
                MsgBox "No es una fecha correcta: " & .Text, vbExclamation
                .Text = ""
                PonFoco Text5(Index)
            End If
        End If
    End With
    
End Sub

Private Sub Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub PonFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub GotFocus(ByRef T As TextBox)
    With T
        T.SelStart = 0
        T.SelLength = Len(T.Text)
    End With
End Sub




Private Sub HacerBorrado()
    Screen.MousePointer = vbHourglass
    Command1.Visible = False
    Pb1.Visible = True
    Conn.BeginTrans
    If BorrarTodosLosDatos Then
        Conn.CommitTrans
        Me.Command2.Caption = "Salir"
        MsgBox "La eliminacion de los datos generados ha sido correcta.", vbInformation
    Else
        Conn.RollbackTrans
    End If
    Pb1.Visible = False
    Screen.MousePointer = vbDefault
End Sub


Private Function BorrarTodosLosDatos() As Boolean
    
    On Error GoTo EBorrarTodosLosDatos
    BorrarTodosLosDatos = False
    
    'Eliminamos los marcajes
    Donde = "Eliminar marcajes"
    EliminarMarcajes
    
    
    Pb1.Value = 0
    Me.Refresh
    
    'Eliminaremos las tareas
    Donde = "Eliminar tareas realizadas"
    EliminarTareasRealizadas
    
    BorrarTodosLosDatos = True
    Exit Function
EBorrarTodosLosDatos:
    MuestraError Err.Number, Donde & vbCrLf & Err.Description
    
End Function


Private Function EliminarMarcajes()
Dim Marcaje As CMarcajes
Dim Fin As Boolean
Dim OK As Boolean



    'Ahora borramos los ticajes k kedaran en la introduccion
    SQL = "DELETE from EntradaFichajes where "
    SQL = SQL & " Fecha >= #" & Format(Text5(0).Text, "yyyy/mm/dd")
    SQL = SQL & "# AND Fecha <= #" & Format(Text5(1).Text, "yyyy/mm/dd") & "#"
    Conn.Execute SQL


    SQL = "from Marcajes where "
    SQL = SQL & " Fecha >= #" & Format(Text5(0).Text, "yyyy/mm/dd")
    SQL = SQL & "# AND Fecha <= #" & Format(Text5(1).Text, "yyyy/mm/dd") & "#"
    
    Set RS = New ADODB.Recordset
    RS.Open "Select count(*) " & SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Tamanyo = 0
    If Not RS.EOF Then Tamanyo = DBLet(RS.Fields(0), "N")
    RS.Close
    
    If Tamanyo = 0 Then Exit Function
    Contador = 0
    Fin = False
    OK = True
    Set Marcaje = New CMarcajes
    
    RS.Open "Select Entrada " & SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not Fin
        Contador = Contador + 1
        PonerBarra CInt((Contador / Tamanyo) * 1000)
    
        If Marcaje.Leer(RS.Fields(0)) = 1 Then
            OK = False
            Fin = True
        Else
            If Marcaje.Eliminar = 1 Then
                OK = False
                Fin = True
            Else
                'Comprobamos si es EOF
                RS.MoveNext
                Fin = RS.EOF
            End If
        End If
    Wend
    RS.Close
    Set RS = Nothing
    Set Marcaje = Nothing
    
    
    

End Function


Private Sub EliminarTareasRealizadas()
    
    Set RS = New ADODB.Recordset
    SQL = "from TareasRealizadas where "
    SQL = SQL & " Fecha >= #" & Format(Text5(0).Text, "yyyy/mm/dd")
    SQL = SQL & "# AND Fecha <= #" & Format(Text5(1).Text, "yyyy/mm/dd") & "#"


    RS.Open "Select count(*) " & SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    Tamanyo = 0
    If Not RS.EOF Then Tamanyo = DBLet(RS.Fields(0), "N")
    RS.Close
    
    If Tamanyo = 0 Then Exit Sub
    Contador = 0
    
    RS.Open "Select * " & SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not RS.EOF
        Contador = Contador + 1
        PonerBarra CInt((Contador / Tamanyo) * 1000)
        RS.Delete
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub

Private Sub PonerBarra(L As Integer)
    On Error Resume Next
    Pb1.Value = L
    If Err.Number <> 0 Then Err.Clear
End Sub
