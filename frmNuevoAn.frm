VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmNuevoAn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anticpos"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5640
   Icon            =   "frmNuevoAn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Caption         =   "PAGADO"
      Height          =   315
      Left            =   1560
      TabIndex        =   4
      Tag             =   "Pagado|T|N|||Pagos|Pagado|||"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   1500
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Tag             =   "Tipo pago|N|N|||Pagos|tipo||S|"
      Top             =   480
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   4440
      TabIndex        =   12
      Top             =   3840
      Width           =   1035
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   3240
      TabIndex        =   11
      Top             =   3840
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3600
      Top             =   0
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Left            =   1500
      TabIndex        =   10
      Text            =   "Text2"
      Top             =   1260
      Width           =   4035
   End
   Begin VB.TextBox Text1 
      Height          =   885
      Index           =   3
      Left            =   120
      MaxLength       =   255
      MultiLine       =   -1  'True
      TabIndex        =   5
      Tag             =   "Observaciones|T|S|||Pagos|Observaciones|||"
      Text            =   "frmNuevoAn.frx":030A
      Top             =   2760
      Width           =   5355
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   120
      TabIndex        =   3
      Tag             =   "Importe|N|N|||Pagos|importe|||"
      Text            =   "Text1"
      Top             =   2040
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Tag             =   "Trabajador|N|N|||Pagos|Trabajador||S|"
      Text            =   "Text1"
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Fecha|F|N|||Pagos|fecha|dd/mm/yyyy|S|"
      Text            =   "Text1"
      Top             =   480
      Width           =   1215
   End
   Begin VB.Image ImageF 
      Height          =   240
      Left            =   600
      Picture         =   "frmNuevoAn.frx":0310
      Top             =   120
      Width           =   240
   End
   Begin VB.Image ImageT 
      Height          =   240
      Left            =   960
      Picture         =   "frmNuevoAn.frx":0412
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   255
      Index           =   4
      Left            =   1500
      TabIndex        =   13
      Top             =   180
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2460
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Importe €"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   8
      Top             =   1740
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   180
      Width           =   555
   End
End
Attribute VB_Name = "frmNuevoAn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public registro As String
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Private Cad As String


Private Sub Check1_KeyPress(KeyAscii As Integer)
     Keypress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
    If registro = "" Then
        If DatosOk Then
            
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                VariableCompartida = "OK"
                Unload Me
            End If
        End If
    Else
        'modificar
        If DatosOk Then
            If ModificaDesdeFormulario(Me) Then
                VariableCompartida = "OK"
                Unload Me
            End If
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    'Carga combo
    CargaCombo
    If registro = "" Then
        Limpiar Me
        Text1(0).Text = Format(Now, "dd/mm/yyyy")
        Text1(0).Enabled = True
        Check1.Value = 1
    Else
        Text1(0).Enabled = False
        
        'Ponemos los campos
        PonerCampos
        
        Combo1.Enabled = False
    End If
    Me.ImageF.Visible = Combo1.Enabled
    Me.ImageT.Visible = Combo1.Enabled
    Text1(1).Enabled = Text1(0).Enabled
End Sub

Private Sub PonerCampos()
    
    
    Cad = "Select * from Pagos Where trabajador="
    Cad = Cad & RecuperaValor(registro, 1)
    Cad = Cad & " AND Fecha = #" & Format(RecuperaValor(registro, 2), FormatoFecha)
    Cad = Cad & "# AND Tipo =" & RecuperaValor(registro, 3)
    adodc1.ConnectionString = conn
    adodc1.RecordSource = Cad
    adodc1.Refresh
    If adodc1.Recordset.EOF Then
        MsgBox "Error obteniendo datos desde BD.", vbExclamation
        Limpiar Me
        cmdAceptar.Enabled = False
    Else
        PonerCamposForma Me, adodc1
        'Ponemos el nombre del empleaqdo
        Text1_LostFocus 1
        Text1(2).Text = TransformaComasPuntos(Text1(2).Text)
    End If
End Sub

Private Function DatosOk() As Boolean
Dim B As Boolean
Dim Impo As Currency

    'Si es insertar
    If registro = "" Then
        If Text1(0).Text <> "" Then

        End If
    End If

    'Por si hay mas de dos decimales
    Impo = ImporteFormateadoAmoneda(Text1(2).Text)
    Impo = Round(Impo, 2)
    Text1(2).Text = Format(Impo, FormatoImporte)

    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    
    
    
    'Si esta embargado no dejamos pasar
    Cad = DevuelveDesdeBD("embargado", "trabajadores", "idtrabajador", Text1(1).Text)
    If Cad = "1" Then
        MsgBox "El trabajador esta en situacion de embargo. No puede generar pagos.", vbExclamation
    Else
        DatosOk = True
    End If
End Function

Private Sub PonFoco(ByRef Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    Text1(1).Text = vCodigo
    Text2.Text = vCadena
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(0).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub ImageF_Click()
        Set frmC = New frmCal
        frmC.Fecha = Now
        If Text1(0).Text <> "" Then
            If IsDate(Text1(0).Text) Then frmC.Fecha = CDate(Text1(0).Text)
        End If
        frmC.Show vbModal
        Set frmC = Nothing
End Sub

Private Sub ImageT_Click()

    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "NomTrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPLEADOS"
    frmB.MostrarDeSalida = True
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

    With Text1(Index)
        .Text = Trim(.Text)
        If .Text = "" Then
            If Index = 1 Then Text2.Text = ""
            Exit Sub
        End If
        
        Select Case Index
        Case 0
            If Not EsFechaOK(Text1(Index)) Then
                MsgBox "Error en la fecha: " & .Text, vbExclamation
                PonFoco Text1(Index)
            End If
        
        Case 1
            If Not IsNumeric(.Text) Then
                Cad = ""
            Else
                Cad = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idTrabajador", .Text, "N")
            End If
            Text2.Text = Cad
            If Cad = "" Then
                MsgBox "Codigo empleado incorrecto: " & .Text, vbExclamation
                .Text = ""
                PonFoco Text1(Index)
            End If
            
        Case 2
            'importe
            Cad = TextBoxAImporte(Text1(2))
            If Cad <> "" Then
                MsgBox Cad, vbExclamation
                Text1(2).Text = ""
                PonFoco Text1(2)
            End If
        End Select
    End With
End Sub


Private Sub CargaCombo()
Dim RT As ADODB.Recordset
    On Error GoTo EC
    Combo1.Clear
    Cad = "Select * from TipoPago"
    Set RT = New ADODB.Recordset
    RT.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RT.EOF
        Combo1.AddItem RT.Fields(1)
        Combo1.ItemData(Combo1.NewIndex) = RT.Fields(0)
        RT.MoveNext
    Wend
    RT.Close
    Set RT = Nothing
    Exit Sub
EC:
    MuestraError Err.Number, "Carga combo"
End Sub
