VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDatosKimaldi 
   Caption         =   "Datos reloj"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9750
   Icon            =   "frmDatosKimaldi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   9750
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1200
      Top             =   6600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosKimaldi.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDatosKimaldi.frx":08A4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5895
      Left            =   0
      TabIndex        =   9
      Top             =   960
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10398
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "NODO"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "HORA"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Marcaje"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Nombre"
         Object.Width           =   61736
      EndProperty
   End
   Begin VB.Frame Frame2 
      Height          =   915
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   9735
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmDatosKimaldi.frx":6096
         Left            =   6000
         List            =   "frmDatosKimaldi.frx":60A3
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   480
         Width           =   975
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         Height          =   375
         Left            =   8640
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox txtTra 
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   480
         Width           =   915
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Text            =   "Text2"
         Top             =   480
         Width           =   3435
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Actualizar"
         Height          =   375
         Index           =   0
         Left            =   7200
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "TERMINAL"
         Height          =   195
         Index           =   1
         Left            =   6000
         TabIndex        =   10
         Top             =   240
         Width           =   930
      End
      Begin VB.Image ImgTrab 
         Height          =   240
         Left            =   2280
         Picture         =   "frmDatosKimaldi.frx":60B0
         Top             =   240
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   8
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   450
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmDatosKimaldi.frx":61B2
         Top             =   240
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmDatosKimaldi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Dim SQL As String
Dim Rs As ADODB.Recordset




Private Sub cmdBuscar_Click()
Dim i As Integer
Dim D As Integer
Dim J As Integer

    If ListView1.ListItems.Count = 0 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then
        i = 1
    Else
        i = ListView1.SelectedItem.Index + 1
    End If
    
    Do
        For J = i To ListView1.ListItems.Count
            If ListView1.ListItems(J).Bold Then
                'Es el selecionado
                Set ListView1.SelectedItem = ListView1.ListItems(J)
                ListView1.ListItems(J).EnsureVisible
                i = 0
                Exit For
            End If
        Next J
        If J > ListView1.ListItems.Count And i <> 0 Then
            If i <> 1 Then
                i = 1
                Beep
            Else
                i = 0
            End If
            
        End If
    Loop Until i = 0
End Sub



Private Sub Combo1_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Command2_Click(Index As Integer)
    CargaGrid
End Sub

Private Sub Form_Load()
    Text1(0).Text = Format(Now, "dd/mm/yyyy")
    Me.txtTra.Text = ""
    Me.Text2.Text = ""
    Combo1.ListIndex = 0
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    VariableCompartida = vCodigo
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(0).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Me.Text1(0).Text <> "" Then frmC.Fecha = CDate(Text1(0).Text)
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub ImgTrab_Click()
    VariableCompartida = ""
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "NomTrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPLEADOS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
    If VariableCompartida <> "" Then
        txtTra.Text = VariableCompartida
        txtTra_LostFocus
    End If
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
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then Exit Sub
    If Not EsFechaOK(Text1(Index)) Then
        MsgBox "No es una fecha correcta: " & Text1(Index).Text, vbExclamation
        Text1(Index).Text = ""
        Focus Text1(Index)
    End If
End Sub

Private Sub Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub Focus(ByRef Ob As Object)
    On Error Resume Next
    Ob.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub
Private Sub CargaGrid()
Dim IT As ListItem
Dim C As Long
Dim Tareas As ADODB.Recordset
    ListView1.ListItems.Clear
    
    
    
    Set Tareas = New ADODB.Recordset
    Tareas.Open "Select * from Tareas", conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    
        SQL = "SELECT MarcajesKimaldi.Nodo, MarcajesKimaldi.Fecha, MarcajesKimaldi.Hora, MarcajesKimaldi.Marcaje, MarcajesKimaldi.TipoMens, Trabajadores.NomTrabajador,Trabajadores.IdTrabajador"
        SQL = SQL & " FROM MarcajesKimaldi LEFT JOIN Trabajadores ON MarcajesKimaldi.Marcaje = Trabajadores.NumTarjeta"
        SQL = SQL & " Where 1=1 "
        '
        If Me.Text1(0).Text <> "" Then SQL = SQL & " AND MarcajesKimaldi.Fecha = #" & Format(Me.Text1(0).Text, FormatoFecha) & "#"
    If Combo1.ListIndex > 0 Then SQL = SQL & " AND NODO = " & Combo1.ItemData(Combo1.ListIndex)
    
    SQL = SQL & " ORDER BY Fecha, Hora;"
    
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    C = 1
    While Not Rs.EOF
        Set IT = ListView1.ListItems.Add(, "C" & C)
        IT.Text = "   " & Rs!Nodo
        IT.SubItems(1) = Format(Rs!Hora, "hh:mm:ss")
        IT.SubItems(2) = Rs!Marcaje
        If Not IsNull(Rs!nomtrabajador) Then
            IT.SubItems(3) = Rs!nomtrabajador
            'icono
            IT.SmallIcon = 1
            
            
            If Rs!idTrabajador = txtTra.Text Then
                IT.Bold = True
                IT.ForeColor = vbRed
                IT.ListSubItems(1).Bold = True
                IT.ListSubItems(1).ForeColor = vbRed
                IT.ListSubItems(2).Bold = True
                IT.ListSubItems(2).ForeColor = vbRed
                IT.ListSubItems(3).Bold = True
                IT.ListSubItems(3).ForeColor = vbRed
                
            End If
        Else
            IT.SubItems(3) = ""
            
            Tareas.Find "Tarjeta ='" & Rs!Marcaje & "'", , adSearchForward, 1
            If Not Tareas.EOF Then IT.SubItems(3) = "   ** " & Tareas!Descripcion
            
            'icono
            IT.SmallIcon = 2
        End If
        
        C = C + 1
        Rs.MoveNext
    Wend
    Rs.Close
    
End Sub

Private Sub txtTra_GotFocus()
    With txtTra
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub



Private Sub txtTra_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtTra_LostFocus()
Dim Cad As String
    If txtTra.Text <> "" Then
        If Not IsNumeric(txtTra.Text) Then
            Cad = ""
        Else
            Cad = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idTrabajador", txtTra.Text, "N")
        End If
        If Cad = "" Then
            MsgBox "Codigo incorrecto: " & txtTra.Text, vbExclamation
            txtTra.Text = ""
        Else
            Me.cmdBuscar.Visible = True
        End If
    Else
        Cad = ""
    End If
    Text2.Text = Cad
End Sub



