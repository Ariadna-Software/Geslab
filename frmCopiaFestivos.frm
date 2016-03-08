VERSION 5.00
Begin VB.Form frmCopiaFestivos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "COPIA DE DIAS FESTIVOS"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmCopiaFestivos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3540
      TabIndex        =   3
      Top             =   3120
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Iniciar copia"
      Height          =   315
      Left            =   2040
      TabIndex        =   2
      Top             =   3120
      Width           =   1275
   End
   Begin VB.Frame Frame2 
      Caption         =   "H. DESTINO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1275
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   4575
      Begin VB.ComboBox Combo4 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmCopiaFestivos.frx":030A
         Left            =   3240
         List            =   "frmCopiaFestivos.frx":0338
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmCopiaFestivos.frx":0390
         Left            =   180
         List            =   "frmCopiaFestivos.frx":03BE
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   720
         Width           =   2835
      End
      Begin VB.Label Label1 
         Caption         =   "Año"
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   9
         Top             =   480
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Horario"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   420
         Width           =   1035
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "H. ORIGEN"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.ComboBox Combo2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmCopiaFestivos.frx":0416
         Left            =   3240
         List            =   "frmCopiaFestivos.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   660
         Width           =   1275
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   2835
      End
      Begin VB.Label Label1 
         Caption         =   "Año"
         Height          =   195
         Index           =   3
         Left            =   3240
         TabIndex        =   11
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "Horario"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   420
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmCopiaFestivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rs As ADODB.Recordset
Dim i As Long
Dim aux


Private Sub Command1_Click()
Dim RD As ADODB.Recordset
Dim vF As String

On Error GoTo ErrorCopiando

If Not DatosOk Then
    MsgBox "Seleccione datos de los cuadros combinados.", vbExclamation
    Exit Sub
End If

Set Rs = New ADODB.Recordset
i = 0
Rs.Open "Select max(id) from Festivos", Conn, , , adCmdText
If Not Rs.EOF Then
    i = DBLet(Rs.Fields(0), "N")
End If
Rs.Close
i = i + 1

Set RD = New ADODB.Recordset
    RD.CursorType = adOpenKeyset
    RD.LockType = adLockOptimistic
aux = "Select * From Festivos "
aux = aux & " WHERE IdHorario=" & Combo1.ItemData(Combo1.ListIndex)
aux = aux & " and anyo= " & Combo2.List(Combo2.ListIndex)
Rs.Open aux, Conn, , , adCmdText
While Not Rs.EOF
    'Creamos la fecha
    vF = "#" & Combo4.List(Combo4.ListIndex) & "/"
    vF = vF & Month(Rs!Fecha) & "/" & Day(Rs!Fecha) & "#"
    
    
    'Comprobamos si existe el campo
    aux = "Select * From Festivos "
    aux = aux & " WHERE IdHorario=" & Combo3.ItemData(Combo3.ListIndex)
    aux = aux & " and anyo= " & Combo4.List(Combo4.ListIndex)
    aux = aux & " and Fecha=" & vF
    
    RD.Open aux, Conn, , , adCmdText
    If RD.EOF Then
        'NUEVO Insertamos
        RD.AddNew
        RD!Id = i
        RD!IdHorario = Combo3.ItemData(Combo3.ListIndex)

        RD!Fecha = CDate(Mid(vF, 2, Len(vF) - 2))
        RD!Anyo = Combo4.List(Combo4.ListIndex)
        RD!Descripcion = Rs!Descripcion
        RD.Update
        i = i + 1
        'ELSE
        Else
            RD!Descripcion = Rs!Descripcion
            RD.Update
    End If
    RD.Close
    Rs.MoveNext
Wend

Rs.Close
Set Rs = Nothing
Set RD = Nothing

MsgBox "Copia finalizada con éxito.", vbInformation
Unload Me
Exit Sub
ErrorCopiando:
    MuestraError Err.Number
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
CargaCombos
End Sub



Private Sub CargaCombos()

Combo1.Clear
Combo2.Clear
Combo3.Clear
Combo4.Clear
On Error GoTo ErrorCarga

For i = 2000 To 2050
    Combo2.AddItem i
    Combo4.AddItem i
Next i
Combo2.ListIndex = Year(Now) - 2000
Combo4.ListIndex = Year(Now) - 2000

Set Rs = New ADODB.Recordset
aux = "Select idHorario,NomHorario from Horarios"
Rs.Open aux, Conn, , , adCmdText
i = 0
While Not Rs.EOF
    Combo1.AddItem Rs.Fields(1)
    Combo1.ItemData(i) = Rs.Fields(0)
    'Para el destino
    Combo3.AddItem Rs.Fields(1)
    Combo3.ItemData(i) = Rs.Fields(0)
    'i
    i = i + 1
    Rs.MoveNext
Wend
Combo1.ListIndex = 0
Combo3.ListIndex = 0
Rs.Close
Set Rs = Nothing
Exit Sub
ErrorCarga:
    MuestraError Err.Number
    Combo1.Clear
    Combo2.Clear
    Combo3.Clear
    Combo4.Clear
End Sub


Private Function DatosOk() As Boolean
DatosOk = False
If Combo1.ListIndex < 0 Then Exit Function
If Combo2.ListIndex < 0 Then Exit Function
If Combo3.ListIndex < 0 Then Exit Function
If Combo4.ListIndex < 0 Then Exit Function
DatosOk = True
End Function
