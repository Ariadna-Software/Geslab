VERSION 5.00
Begin VB.Form frmDiasFest 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Dias festivos año"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmDiasFest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4800
      TabIndex        =   4
      Top             =   1020
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1500
      MaxLength       =   50
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1140
      Width           =   2715
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1500
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   195
      Left            =   1500
      TabIndex        =   7
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Seleccione una fecha y ponga una descripción para el dia festivo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   6
      Top             =   180
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   960
      Picture         =   "frmDiasFest.frx":030A
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Top             =   1140
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmDiasFest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Anyo As Integer
Public Fecha As Date
Public Descripcion As String
Public IdFestivo As Long
Public IdHor As Integer 'para saber si ya existe

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private Sub Command1_Click()
Dim Rs As ADODB.Recordset
Dim Sql As String
Dim Clave As Long

On Error GoTo ErrorCommand1

If Not DatosOk Then
    MsgBox "La fecha introducida NO es correcta o no se corresponde con el año " & Anyo, vbExclamation
    Exit Sub
End If
'Nuevo o modificar
Set Rs = New ADODB.Recordset
Rs.CursorType = adOpenKeyset
Rs.LockType = adLockOptimistic
If IdFestivo = 0 Then
    Sql = "SELECT max(id) FROM Festivos"
    Rs.Open Sql, Conn, , , adCmdText
    Clave = DBLet(Rs.Fields(0), "N") + 1
    Rs.Close
    Sql = "Select * from Festivos where "
    Sql = Sql & " Fecha=#" & Format(Text1(0).Text, "yyyy/mm/dd") & "#"
    Sql = Sql & " AND IdHorario=" & IdHor
    Rs.Open Sql, Conn, , , adCmdText
    If Rs.EOF Then
        'No existe ningun festivo en esa fecha y lo damos de alta
        Rs.AddNew
        Rs!Id = Clave
        Rs!Fecha = Text1(0).Text
        Rs!IdHorario = IdHor
        Rs!Descripcion = Text1(1).Text
        Rs!Anyo = Year(Text1(0).Text)
        Rs.Update
        Rs.Close
        Else
            MsgBox "Ya existe un dia festivo para " & vbCrLf & _
                "Fecha: " & Text1(0).Text & _
                "Descricpcion: " & Rs.Fields(4), vbExclamation
            Rs.Close
            Screen.MousePointer = vbDefault
            Exit Sub
    End If
    Else
        '----------------------------------
        'MODIFICAR
        '----------------------------------
        Sql = "Select * from Festivos where "
        Sql = Sql & " Id=" & IdFestivo
        Rs.Open Sql, Conn, , , adCmdText
        If Not Rs.EOF Then
            'No existe ningun festivo en esa fecha y lo damos de alta
            Rs!Descripcion = Text1(1).Text
            Rs.Update
            Else
                MsgBox "No existe el dia festivo: " & Text1(0).Text, vbCritical
        End If
        Rs.Close
End If
Set Rs = Nothing
Unload Me
Exit Sub
ErrorCommand1:
    MsgBox "Error: " & Err.Number & vbCrLf & "Descripción: " & _
        Err.Description, vbExclamation
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Caption = "Dias festivos para el año: " & Anyo
If IdFestivo > 0 Then
    'Modificar
    Command1.Caption = "Modificar"
    Label3.Caption = "Modificar"
    Label3.ForeColor = vbRed
    Text1(0).Text = Fecha
    Text1(1).Text = Descripcion
    Else
        'Nuevo
        Command1.Caption = "Insertar"
        Label3.Caption = "Insertar"
        Label3.ForeColor = vbBlue
        Text1(0).Text = ""
        Text1(1).Text = ""
End If
'Si es modificar NO se puede cambiar la fecha
Text1(0).Enabled = (IdFestivo = 0)
Image1.Enabled = (IdFestivo = 0)
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(0).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
Set frmC = New frmCal
frmC.Fecha = Day(Now) & "/" & Month(Now) & "/" & Anyo
frmC.Show vbModal
Set frmC = Nothing
End Sub


Private Function DatosOk() As Boolean
DatosOk = False
If Text1(0).Text = "" Then Exit Function
If Not IsDate(Text1(0).Text) Then Exit Function
If Year(Text1(0).Text) <> Anyo Then Exit Function
DatosOk = True
End Function
