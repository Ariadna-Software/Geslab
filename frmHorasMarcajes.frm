VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHorasMarcajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Marcajes máquina"
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6240
   Icon            =   "frmHorasMarcajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Volver"
      Height          =   375
      Index           =   1
      Left            =   5100
      TabIndex        =   0
      Top             =   5760
      Width           =   1035
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5640
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorasMarcajes.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Aceptar"
      Height          =   375
      Index           =   0
      Left            =   5100
      TabIndex        =   15
      Top             =   5760
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Eliminar"
      Height          =   495
      Index           =   2
      Left            =   5100
      Picture         =   "frmHorasMarcajes.frx":0624
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4140
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Modificar"
      Height          =   495
      Index           =   1
      Left            =   5100
      Picture         =   "frmHorasMarcajes.frx":0726
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3540
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Nuevo"
      Height          =   495
      Index           =   0
      Left            =   5100
      Picture         =   "frmHorasMarcajes.frx":0828
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   2940
      Width           =   975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3255
      Left            =   180
      TabIndex        =   8
      Top             =   2880
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5741
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Incidencia"
         Object.Width           =   7250
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "idInci"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1875
      Left            =   120
      TabIndex        =   2
      Top             =   540
      Width           =   5955
      Begin VB.TextBox txtHorario 
         Height          =   285
         Index           =   4
         Left            =   4500
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox txtHorario 
         Height          =   285
         Index           =   3
         Left            =   1380
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   1380
         Width           =   1095
      End
      Begin VB.TextBox txtHorario 
         Height          =   285
         Index           =   2
         Left            =   4500
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtHorario 
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtHorario 
         Height          =   285
         Index           =   0
         Left            =   2220
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   420
         Width           =   2775
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Salida segunda"
         Height          =   195
         Index           =   4
         Left            =   3180
         TabIndex        =   13
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Entrada segunda"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Salida primera"
         Height          =   195
         Index           =   2
         Left            =   3180
         TabIndex        =   11
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Entrada primera"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label Label2 
         Caption         =   "Horario del empleado"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   9
         Top             =   480
         Width           =   1695
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Marcajes"
      Height          =   195
      Left            =   240
      TabIndex        =   14
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   5475
   End
End
Attribute VB_Name = "frmHorasMarcajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmH As frmSoloHora
Attribute frmH.VB_VarHelpID = -1
Public Event HayModificacion(SiNo As Boolean, vOpcion As Byte)
Public OPCION As Byte
    '0.- Horas en entradamarcajes
    '1.- En HSITORICO

Public Nombre As String
Public vM As CMarcajes
Public VH As CHorarios


'Para saber a que item nos referimos
Private Secuencia As Long
Private SeHaModificado As Boolean


Private Sub Command1_Click(Index As Integer)
Dim RT As ADODB.Recordset
Dim cad As String
Dim RC As Byte

Select Case Index
Case 0
    'Nuevo
    Secuencia = -1
    Set frmH = New frmSoloHora
    frmH.Hora = ""
    frmH.Inci = 0
    frmH.CadInci = ""
    frmH.Show vbModal
    Set frmH = Nothing
Case 1
    'Modificar
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    Secuencia = ListView1.SelectedItem.Tag
    Set frmH = New frmSoloHora
    frmH.Hora = ListView1.SelectedItem
    frmH.Inci = ListView1.SelectedItem.SubItems(2)
    frmH.CadInci = ListView1.SelectedItem.SubItems(1)
    frmH.Show vbModal
    Set frmH = Nothing
Case 2
    'Eliminar
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    
    If OPCION = 0 Then
        cad = "Seguro que desea eliminar el ticaje efectuado " & vbCrLf
        cad = cad & " a las " & ListView1.SelectedItem.Text & vbCrLf
        If ListView1.SelectedItem.SubItems(2) <> 0 Then _
            cad = cad & "y con la incidencia : " & ListView1.SelectedItem.SubItems(1)
        RC = MsgBox(cad, vbQuestion + vbYesNo)
        If RC = vbYes Then
            'Eliminamos
            Set RT = New ADODB.Recordset
            cad = "Delete * from EntradaMarcajes where secuencia=" & ListView1.SelectedItem.Tag
            RT.Open cad, Conn, , , adCmdText
            Set RT = Nothing
            SeHaModificado = True
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
         End If
         
    Else
        'Elminamos del HCO
        cad = "Seguro que desea eliminar el ticaje efectuado " & vbCrLf
        cad = cad & " a las " & ListView1.SelectedItem.Text & vbCrLf
        If ListView1.SelectedItem.SubItems(2) <> 0 Then _
            cad = cad & "y con la incidencia : " & ListView1.SelectedItem.SubItems(1)
        RC = MsgBox(cad, vbQuestion + vbYesNo)
        If RC = vbYes Then
            'Eliminamos
            Set RT = New ADODB.Recordset
            cad = "Delete * from EntradaFichajes where secuencia=" & ListView1.SelectedItem.Tag
            RT.Open cad, Conn, , , adCmdText
            Set RT = Nothing
            SeHaModificado = True
            ListView1.ListItems.Remove ListView1.SelectedItem.Index
        End If
    End If
End Select
End Sub

Private Sub Command2_Click(Index As Integer)
If SeHaModificado Then
    RaiseEvent HayModificacion(True, OPCION)
End If
Unload Me
End Sub

Private Sub Form_Load()
Label1.Caption = Nombre
If OPCION = 0 Then
    Caption = "Modificar marcajes"
    Label1.ForeColor = vbBlack
Else
    Caption = "Modificar ENTRADAS fichajes"
    Label1.ForeColor = vbRed
End If

'Ponemos el horario oficial
SeHaModificado = False
txtHorario(0).Text = VH.NomHorario
txtHorario(1).Text = VH.HoraE1
txtHorario(2).Text = VH.HoraS1
txtHorario(3).Text = VH.HoraE2
txtHorario(4).Text = VH.HoraS2
Cargalistview
End Sub



Private Sub Cargalistview()
Dim rs As ADODB.Recordset
Dim SQL As String
Dim itm As ListItem

ListView1.ListItems.Clear
Set rs = New ADODB.Recordset
'Sql = "Select * from  ControlFichajes WHERE fecha=#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"


If OPCION = 0 Then
    SQL = "SELECT EntradaMarcajes.Hora, Incidencias.NomInci, EntradaMarcajes.idInci,Secuencia"
    SQL = SQL & " FROM EntradaMarcajes ,Incidencias WHERE EntradaMarcajes.idInci = Incidencias.IdInci"
    SQL = SQL & " AND idMarcaje=" & vM.Entrada
    SQL = SQL & " Order by Hora"
Else
    SQL = "    SELECT EntradaFichajes.Hora, Incidencias.NomInci, EntradaFichajes.idInci,Secuencia"
    SQL = SQL & " FROM EntradaFichajes INNER JOIN Incidencias ON EntradaFichajes.idInci = Incidencias.IdInci"
    SQL = SQL & " WHERE idTrabajador=" & vM.idTrabajador
    SQL = SQL & " AND Fecha = #" & Format(vM.Fecha, "yyyy/mm/dd") & "#"
    SQL = SQL & " Order by Hora"
End If



rs.Open SQL, Conn, , , adCmdText
While Not rs.EOF
    Set itm = ListView1.ListItems.Add(, , rs!Hora)
    If rs!IdInci = 0 Then
        itm.SubItems(1) = ""
        Else
        itm.SubItems(1) = rs!NomInci
    End If
    itm.SubItems(2) = rs!IdInci
    itm.Tag = rs!Secuencia
    itm.SmallIcon = 1
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
End Sub

Private Sub frmH_Seleccionar(vHora As Date, vInci As Integer)
Dim rs As ADODB.Recordset
Dim Valor As Long
Dim Tabla As String

On Error GoTo ErrFRMSelecionar
Set rs = New ADODB.Recordset
rs.CursorType = adOpenKeyset
rs.LockType = adLockOptimistic
If OPCION = 0 Then
    Tabla = "EntradaMarcajes"
Else
    Tabla = "EntradaFichajes"  'Los fichajes
End If

If Secuencia < 0 Then
    'Nuevo marcaje
    Valor = 1

      
    rs.Open "Select max(Secuencia) FROM " & Tabla, Conn, , , adCmdText
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then Valor = rs.Fields(0) + 1
    End If
    rs.Close
    
    'Añadimos
    rs.Open "Select * from " & Tabla, Conn, adOpenKeyset, , adCmdText
    rs.AddNew
        rs!Secuencia = Valor
        rs!idTrabajador = vM.idTrabajador
        If OPCION = 0 Then rs!idMarcaje = vM.Entrada
        rs!Fecha = vM.Fecha
        rs!Hora = vHora
        rs!horareal = vHora
        rs!IdInci = vInci
    rs.Update
    SeHaModificado = True
    Else
    'MODIFICAR
    rs.Open "Select * from " & Tabla & " where secuencia=" & Secuencia, Conn, , , adCmdText
    If rs.EOF Then
        MsgBox "Se ha producido un error leyendo el marcaje."
        Else
            rs!Hora = vHora
            rs!IdInci = vInci
            If OPCION = 1 Then
                'Modificamos tb el real
                rs!horareal = vHora
            End If
            rs.Update
            SeHaModificado = True
    End If
End If
rs.Close
espera 0.25
Set rs = Nothing
Cargalistview
Exit Sub
ErrFRMSelecionar:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

Private Sub ListView1_DblClick()
    Command1_Click (1)
End Sub
