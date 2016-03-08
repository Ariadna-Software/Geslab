VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBusca 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   -510
   ClientTop       =   -750
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   ScaleHeight     =   6270
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3540
      TabIndex        =   3
      Top             =   5760
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Regresar"
      Default         =   -1  'True
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   5760
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   1020
      Width           =   4515
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4035
      Left            =   180
      TabIndex        =   1
      Top             =   1620
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   7117
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   435
      Left            =   780
      TabIndex        =   5
      Top             =   180
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Escriba tres carácteres (por lo menos)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   180
      TabIndex        =   4
      Top             =   720
      Width           =   2235
   End
End
Attribute VB_Name = "frmBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Tabla As String
Public CampoBusqueda As String
Public CampoCodigo As String
Public TipoDatos As Byte
Public Titulo As String
Public MostrarDeSalida As Boolean
'Evento publico
Public Event Seleccion(vCodigo As Long, vCadena As String)


Private RS As ADODB.Recordset
Dim Cad As String
Dim CadB As String
Dim CadB2 As String


Private Sub Command1_Click()
    On Error Resume Next
    'Si hay seleccionado algun item
    If Not ListView1.SelectedItem Is Nothing Then _
        RaiseEvent Seleccion(CLng(ListView1.SelectedItem.Tag), ListView1.SelectedItem.Text)
    If Err.Number = 0 Then
        Unload Me
        Else
            MsgBox "ERROR: " & Err.Description
    End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()

Label2.Caption = Titulo
Set RS = New ADODB.Recordset
Cad = "Select distinct " & CampoBusqueda & " , " & _
    CampoCodigo & " from " & Tabla
'Para la cadena de busqueda
Select Case TipoDatos
Case 1, 2
    CadB = CampoBusqueda & " = "
    CadB2 = ""
Case 3
    CadB = CampoBusqueda & " like '%"
    CadB2 = "%'"
End Select

'Solo para las incidencias
If Titulo = "INCIDENCIAS" Then Cad = Cad & " ORDER By nominci DESC"
ListView1.ColumnHeaders(1).Width = ListView1.Width - 120
'Si no queremos que aprezca nada, hasta que escriba tres letras
'hay que comentar este trozo
If MostrarDeSalida Then
    RS.Open Cad, Conn, , , adCmdText
    If Not RS.EOF Then
        Cargalistview
        Else
            RS.Close
    End If
End If
End Sub



Private Sub Cargalistview()
Dim itmX As ListItem

ListView1.ListItems.Clear
'iniciamos la carga
RS.MoveFirst
While Not RS.EOF
    Set itmX = ListView1.ListItems.Add(, , RS.Fields(0))
    itmX.Tag = RS.Fields(1)
    RS.MoveNext
Wend
'cada vez cierra el recodset
RS.Close
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set RS = Nothing
End Sub


Private Sub ListView1_DblClick()
If Not ListView1.SelectedItem Is Nothing Then Command1_Click
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If Len(Text1.Text) > 2 Then
    Cad = "Select " & CampoBusqueda & " , " & _
        CampoCodigo & " from " & Tabla & " WHERE "
    Cad = Cad & CadB & Text1.Text & Chr(KeyAscii) & CadB2
    RS.Open Cad, Conn, , , adCmdText
    If Not RS.EOF Then
        Cargalistview
        Else
        RS.Close
    End If
End If
End Sub

