VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBajas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Bajas"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   Icon            =   "frmBajas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdT 
      Caption         =   "+"
      Height          =   290
      Left            =   4560
      TabIndex        =   16
      Top             =   5640
      Width           =   195
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Bajas actuales"
      Height          =   255
      Left            =   3240
      TabIndex        =   15
      Top             =   6120
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdCta 
      Caption         =   "+"
      Height          =   290
      Left            =   840
      TabIndex        =   14
      Top             =   5640
      Width           =   195
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   5
      Left            =   6300
      TabIndex        =   5
      Tag             =   "Fecha Alta|F|S|||bajas|fechaalta|dd/mm/yyyy||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   4
      Left            =   4800
      TabIndex        =   4
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   3
      Left            =   3840
      TabIndex        =   3
      Tag             =   "Cod. trabajador|N|N|||bajas|idtrab||S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   290
      Index           =   2
      Left            =   2400
      TabIndex        =   2
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7260
      TabIndex        =   6
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   8460
      TabIndex        =   7
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   1
      Left            =   900
      MaxLength       =   30
      TabIndex        =   1
      Tag             =   "idbaja|N|N|||bajas|idtipobaja|||"
      Text            =   "Dato2"
      Top             =   5640
      Width           =   1395
   End
   Begin VB.TextBox txtAux 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   290
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Tag             =   "Fecha baja|F|N|||bajas|fechabaja|dd/mm/yyyy|S|"
      Text            =   "Dat"
      Top             =   5640
      Width           =   800
   End
   Begin VB.CommandButton cmdRegresar 
      Caption         =   "&Regresar"
      Height          =   375
      Left            =   8460
      TabIndex        =   10
      Top             =   6000
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   120
      TabIndex        =   8
      Top             =   5895
      Width           =   2865
      Begin VB.Label lblIndicador 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   240
         Width           =   2550
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   9540
      _ExtentX        =   16828
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Poner MARCAJE como baja"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   5160
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmBajas.frx":000C
      Height          =   5325
      Left            =   60
      TabIndex        =   13
      Top             =   540
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   9393
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         AllowFocus      =   0   'False
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin MSAdodcLib.Adodc adodc1 
      Height          =   375
      Left            =   5970
      Top             =   0
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
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
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
         Shortcut        =   ^N
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnBarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
         Shortcut        =   ^S
      End
   End
End
Attribute VB_Name = "frmBajas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Tag: Nombre concepto|T|N|||sconam|nomconam|||
Private WithEvents frmC As frmTiposDiario
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
Public DatosADevolverBusqueda As String    'Tendra el n� de text que quiere que devuelva, empipados
Public Event DatoSeleccionado(CadenaSeleccion As String)
Private CadenaConsulta As String
Dim CadAncho As Boolean  'Para cuando llamemos al al form de lineas
Dim Modo As Byte
Dim JJ As Integer
Dim SQL As String

'----------------------------------------------
'----------------------------------------------
'   Deshabilitamos todos los botones menos
'   el de salir
'   Ademas mostramos aceptar y cancelar
'   Modo 0->  Normal
'   Modo 1 -> Lineas  INSERTAR
'   Modo 2 -> Lineas MODIFICAR
'   Modo 3 -> Lineas BUSCAR
'----------------------------------------------
'----------------------------------------------

Private Sub PonerModo(vModo)
Dim B As Boolean
Modo = vModo

B = (Modo = 0)


cmdCta.Visible = Modo > 0
cmdT.Visible = ((Modo = 1) Or Modo = 3)
For JJ = 0 To 5
    txtAux(JJ).Visible = Not B
Next JJ
mnOpciones.Enabled = B
Toolbar1.Buttons(1).Enabled = B
Toolbar1.Buttons(2).Enabled = B

cmdAceptar.Visible = Not B
cmdCancelar.Visible = Not B
'DataGrid1.Enabled = b

'Si es regresar
If DatosADevolverBusqueda <> "" Then
    cmdRegresar.Visible = B
End If
'Si estamo mod or insert
If Modo = 2 Then
   txtAux(0).BackColor = &H80000018
   txtAux(3).BackColor = &H80000018
   Else
    txtAux(0).BackColor = &H80000005
    txtAux(3).BackColor = &H80000005
End If
txtAux(0).Enabled = (Modo <> 2)
txtAux(3).Enabled = (Modo <> 2)

End Sub

Private Sub BotonAnyadir()
    Dim anc As Single
    
    'Obtenemos la siguiente numero de factura
    lblIndicador.Caption = "INSERTANDO"
    'Situamos el grid al final
    DataGrid1.AllowAddNew = True
    If Not adodc1.Recordset.EOF Then
        DataGrid1.HoldFields
        adodc1.Recordset.MoveLast
    End If
    
    
   
    If DataGrid1.Row < 0 Then
        anc = 755
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row + 1) + 545
    End If
    For JJ = 0 To 5
        txtAux(JJ).Text = ""
    Next JJ
    LLamaLineas anc, 0
    
    
    'Ponemos el foco
    txtAux(0).SetFocus
    
'    If FormularioHijoModificado Then
'        CargaGrid
'        BotonAnyadir
'        Else
'            'cmdCancelar.SetFocus
'            If Not Adodc1.Recordset.EOF Then _
'                Adodc1.Recordset.MoveFirst
'    End If
End Sub



Private Sub BotonVerTodos()
    CargaGrid ""
End Sub

Private Sub BotonBuscar()
    CargaGrid "idtrab = -1"
    'Buscar
    For JJ = 0 To txtAux.Count - 1
        txtAux(JJ).Text = ""
    Next JJ
    LLamaLineas DataGrid1.Top + 206, 2
    txtAux(0).SetFocus
End Sub

Private Sub BotonModificar()
    '---------
    'MODIFICAR
    '----------
    Dim Cad As String
    Dim anc As Single
    Dim i As Integer
    If adodc1.Recordset.EOF Then Exit Sub
    'If Adodc1.Recordset.RecordCount < 1 Then Exit Sub


    Screen.MousePointer = vbHourglass
    Me.lblIndicador.Caption = "MODIFICAR"
    
    If DataGrid1.Bookmark < DataGrid1.FirstRow Or DataGrid1.Bookmark > (DataGrid1.FirstRow + DataGrid1.VisibleRows - 1) Then
        i = DataGrid1.Bookmark - DataGrid1.FirstRow
        DataGrid1.Scroll 0, i
        DataGrid1.Refresh
    End If
    
    If DataGrid1.Row < 0 Then
        anc = 320
        Else
        anc = DataGrid1.RowTop(DataGrid1.Row) + 545
    End If

    'Llamamos al form
    For JJ = 0 To 5
        txtAux(JJ).Text = DataGrid1.Columns(JJ).Text
    Next JJ
    LLamaLineas anc, 1
   
   'Como es modificar
   PonerFoco txtAux(1)
    Screen.MousePointer = vbDefault
End Sub

Private Sub LLamaLineas(alto As Single, xModo As Byte)
DeseleccionaGrid
PonerModo xModo + 1
cmdCta.Top = alto
cmdT.Top = alto
'Fijamos el ancho
For JJ = 0 To 5
    txtAux(JJ).Top = alto
Next JJ
End Sub




Private Sub BotonEliminar()
Dim SQL As String
    On Error GoTo Error2
    'Ciertas comprobaciones
    If adodc1.Recordset.EOF Then Exit Sub
 
    
    '### a mano
    SQL = "Seguro que desea eliminar la baja :"
    SQL = SQL & vbCrLf & "Trabajador: " & adodc1.Recordset.Fields(4)
    SQL = SQL & vbCrLf & "Codigo: " & adodc1.Recordset.Fields(3)
    SQL = SQL & vbCrLf & "Fecha baja: " & adodc1.Recordset.Fields(0)
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
        'Hay que eliminar
        SQL = "Delete from bajas where idtrab = " & adodc1.Recordset.Fields(3)
        SQL = SQL & " AND FechaBaja = #" & Format(adodc1.Recordset.Fields(0), FormatoFecha) & "#"
        Conn.Execute SQL
        espera 0.5
        CargaGrid ""
        adodc1.Recordset.Cancel
    End If
    Exit Sub
Error2:
        Screen.MousePointer = vbDefault
        MuestraError Err.Number, "Eliminando registro" & vbCrLf & Err.Description
End Sub





Private Sub cmdAceptar_Click()
Dim i As Integer
Dim CadB As String
Select Case Modo
    Case 1
    If DatosOk Then
            '-----------------------------------------
            'Hacemos insertar
            If InsertarDesdeForm(Me) Then
                AjustarDiasBajas CLng(txtAux(3).Text), CDate(txtAux(1).Text)
                'MsgBox "Registro insertado.", vbInformation
                CargaGrid
                BotonAnyadir
            End If
        End If
    Case 2
            'Modificar
            If DatosOk Then
                '-----------------------------------------
                'Hacemos insertar
                If ModificaDesdeFormulario(Me) Then
                   
                    i = adodc1.Recordset.AbsolutePosition
                    PonerModo 0
                    CargaGrid
                    adodc1.Recordset.Move i - 1
                    lblIndicador.Caption = ""
                End If
            End If
    Case 3
        'HacerBusqueda
        CadB = ObtenerBusqueda(Me)
        If CadB <> "" Then
            PonerModo 0
            CargaGrid CadB
        End If
    End Select


End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 1
    DataGrid1.AllowAddNew = False
    'CargaGrid
    If Not adodc1.Recordset.EOF Then adodc1.Recordset.MoveFirst
    
Case 3
    CargaGrid
End Select
PonerModo 0
lblIndicador.Caption = ""
DataGrid1.SetFocus
End Sub

Private Sub cmdCta_Click()
    Set frmC = New frmTiposDiario
    frmC.DatosADevolverBusqueda = "0|1"
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub cmdRegresar_Click()
Dim Cad As String

If adodc1.Recordset.EOF Then
    MsgBox "Ning�n registro a devolver.", vbExclamation
    Exit Sub
End If

Cad = adodc1.Recordset.Fields(0) & "|"
Cad = Cad & adodc1.Recordset.Fields(1) & "|"
RaiseEvent DatoSeleccionado(Cad)
Unload Me
End Sub






Private Sub cmdT_Click()
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

Private Sub DataGrid1_DblClick()
If cmdRegresar.Visible Then cmdRegresar_Click
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub


Private Sub Form_Load()

          ' ICONITOS DE LA BARRA
    With Me.Toolbar1
        .ImageList = frmPpal1.imgListComun
        .Buttons(1).Image = 1
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        .Buttons(10).Image = 22
        .Buttons(11).Image = 16
        .Buttons(12).Image = 15
        .Buttons(14).Image = 6
        .Buttons(15).Image = 7
        .Buttons(16).Image = 8
        .Buttons(17).Image = 9
    End With

    '## A mano
    'Vemos como esta guardado el valor del check
    chkVistaPrevia.Value = CheckValueLeer(Name)
   
    cmdRegresar.Visible = (DatosADevolverBusqueda <> "")
    
    DespalzamientoVisible False
    PonerModo 0
    CadAncho = False

    'Cadena consulta
    CargaGrid
    lblIndicador.Caption = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    CheckValueGuardar Me.Name, Me.chkVistaPrevia.Value
End Sub


Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    txtAux(3).Text = vCodigo
    txtAux(4).Text = vCadena
End Sub

Private Sub frmC_DatoSeleccionado(CadenaSeleccion As String)
    txtAux(1).Text = RecuperaValor(CadenaSeleccion, 1)
    txtAux(2).Text = RecuperaValor(CadenaSeleccion, 2)
End Sub

Private Sub mnBuscar_Click()
    BotonBuscar
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnModificar_Click()
    BotonModificar
End Sub

Private Sub mnNuevo_Click()
    BotonAnyadir
End Sub

Private Sub mnSalir_Click()
    Screen.MousePointer = vbHourglass
    Unload Me
End Sub

Private Sub mnVerTodos_Click()
    BotonVerTodos
End Sub




Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index > 5 And Button.Index < 9 Then
        If vUsu.Nivel > 1 Then
            MsgBox "No tiene autorizacion para realizar cambios", vbExclamation
            Exit Sub
        End If
    End If
    Select Case Button.Index
    Case 1
            BotonBuscar
    Case 2
            BotonVerTodos
    Case 6
            BotonAnyadir
    Case 7
            BotonModificar
    Case 8
            BotonEliminar
    Case 10
            If Me.adodc1.Recordset.EOF Then Exit Sub
            AjustarDiasBajas adodc1.Recordset!idTrabajador, adodc1.Recordset!fechabaja
    Case 12
            Unload Me
    Case Else
    
    End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
    Dim i
    For i = 14 To 17
        Toolbar1.Buttons(i).Visible = bol
    Next i
End Sub

Private Sub CargaGrid(Optional vSQL As String)
    Dim J As Integer
    Dim TotalAncho As Integer
    Dim i As Integer
    
    adodc1.ConnectionString = Conn
    If Check1.Value = 1 Then
        SQL = " fechaalta is null"
    Else
        SQL = ""
    End If
    If vSQL <> "" Then
        If SQL <> "" Then SQL = SQL & " AND "
    End If
    vSQL = SQL & vSQL
    
    PonerSQL
    If vSQL <> "" Then SQL = SQL & " WHERE " & vSQL
    
    SQL = SQL & " ORDER BY fechabaja"
    adodc1.RecordSource = SQL
    adodc1.CursorType = adOpenDynamic
    adodc1.LockType = adLockOptimistic
    adodc1.Refresh
    
    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 290
    
    ' Fechabaja, idbaja, descbaja, IdTrabajador, NomTrabajador"
    'Cuenta contable
    i = 0
        DataGrid1.Columns(i).Caption = "Fec. Baja"
        DataGrid1.Columns(i).Width = 1000
        DataGrid1.Columns(i).NumberFormat = "dd/mm/yyyy"
    
    'Descripcion NOMMACTA
    i = 1
        DataGrid1.Columns(i).Caption = "Tipo"
        DataGrid1.Columns(i).Width = 500
        TotalAncho = TotalAncho + DataGrid1.Columns(i).Width
    
    'Entidad
    i = 2
        DataGrid1.Columns(i).Caption = "Descr."
        DataGrid1.Columns(i).Width = 2000
        
        
    i = 3
        DataGrid1.Columns(i).Caption = "Trab."
        DataGrid1.Columns(i).Width = 800
        
        
    'Codigo control
    i = 4
        DataGrid1.Columns(i).Caption = "Nombre"
        DataGrid1.Columns(i).Width = 3200
    'Cueta bancaria
    i = 5
        DataGrid1.Columns(i).Caption = "Alta"
        DataGrid1.Columns(i).Width = 1000
        DataGrid1.Columns(i).NumberFormat = "dd/mm/yyyy"
    
        
    For i = 0 To 3
        DataGrid1.Columns(i).AllowSizing = False
    Next i
        
        'Fiajamos el cadancho
    If Not CadAncho Then
        'La primera vez fijamos el ancho y alto de  los txtaux
        txtAux(0).Left = DataGrid1.Left + 340
        txtAux(0).Width = DataGrid1.Columns(0).Width - 60
        
        
        For JJ = 1 To 5
            txtAux(JJ).Width = DataGrid1.Columns(JJ).Width - 60
            txtAux(JJ).Left = txtAux(JJ - 1).Left + txtAux(JJ - 1).Width + 60
        Next JJ
        txtAux(5).Left = txtAux(5).Left + 15
        CadAncho = True
        
        'El botoncito para la cuenta
        cmdCta.Left = txtAux(2).Left - 180
        cmdT.Left = txtAux(4).Left - 180
    End If
'    'Habilitamos modificar y eliminar
'    If vUsu.Nivel < 2 Then
'        Toolbar1.Buttons(7).Enabled = Not adodc1.Recordset.EOF
'        Toolbar1.Buttons(8).Enabled = Not adodc1.Recordset.EOF
'    End If
End Sub

Private Sub txtaux_GotFocus(Index As Integer)
With txtAux(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtaux_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtAux_LostFocus(Index As Integer)
Dim RC As String
    txtAux(Index).Text = Trim(txtAux(Index).Text)
    If txtAux(Index).Text = "" Then
        'Si es tipobaja o empleado hay que poner a "" los correspondientes
        If Index = 1 Then txtAux(2).Text = ""
        If Index = 3 Then txtAux(4).Text = ""
        Exit Sub
    End If
    If Modo = 3 Then Exit Sub 'Busquedas
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 0, 5
            If Not EsFechaOK(txtAux(Index)) Then
                MsgBox "Fecha incorrecta: " & txtAux(Index).Text, vbExclamation
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
    Case 1, 3
            If Not IsNumeric(txtAux(Index).Text) Then
                MsgBox "Campo numerico incorrecto: " & txtAux(Index).Text, vbExclamation
                txtAux(Index).Text = ""
                PonerFoco txtAux(Index)
            End If
            
            If Index = 1 Then
                RC = DevuelveDesdeBD("descbaja", "tipobaja", "idbaja", txtAux(Index).Text)
            Else
                RC = DevuelveDesdeBD("nomtrabajador", "Trabajadores", "idTrabajador", txtAux(Index).Text)
            End If
            If RC = "" Then
                MsgBox "Codigo incorrecto.", vbExclamation
                txtAux(Index).Text = ""
            End If
            txtAux(Index + 1).Text = RC
            If RC = "" Then PonerFoco txtAux(Index)
    End Select
    Screen.MousePointer = vbDefault
End Sub


Private Function DatosOk() As Boolean
Dim B As Boolean
Dim RS As ADODB.Recordset

    DatosOk = False
    B = CompForm(Me)
    If Not B Then Exit Function
    
    'Datos bien. Ahora comprobaremos que si es insertar el trabajador no tiene ninguna
    ' inicdencia abierta
    If Modo = 1 Then
        SQL = "Select * from Bajas where idTrab =" & txtAux(3).Text
        SQL = SQL & " AND fechabaja >=#" & Format(txtAux(0).Text) & "#"
        SQL = SQL & " AND fechaalta is null"
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then B = False
        End If
        RS.Close
        Set RS = Nothing
    Else
        'Modificamos. Si ha puesto fecha alta, comprobaremos k la fecha de alta no es menor k la de baja
        If txtAux(5).Text <> "" Then
            If CDate(txtAux(0).Text) > CDate(txtAux(5).Text) Then
                MsgBox "Fecha alta es menor que la fecha de baja", vbExclamation
                B = False
            End If
        End If
        
    End If
    DatosOk = B
End Function

Private Sub DeseleccionaGrid()
    On Error GoTo EDeseleccionaGrid
    While DataGrid1.SelBookmarks.Count > 0
        DataGrid1.SelBookmarks.Remove 0
    Wend
    Exit Sub
EDeseleccionaGrid:
        Err.Clear
End Sub


Private Sub PonerSQL()
    SQL = "SELECT Bajas.Fechabaja, tipobaja.idbaja, tipobaja.descbaja, Trabajadores.IdTrabajador, Trabajadores.NomTrabajador,Fechaalta"
    SQL = SQL & " FROM Trabajadores INNER JOIN (tipobaja INNER JOIN Bajas ON tipobaja.idbaja = Bajas.idTipobaja) ON Trabajadores.IdTrabajador = Bajas.idTrab"
End Sub





Private Sub PonerFoco(Obj As Object)
    On Error Resume Next
    Obj.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Sub AjustarDiasBajas(idTRa As Long, Fec As Date)
Dim vM As CMarcajes
    
    Set vM = New CMarcajes
    If vM.Leer2(idTRa, Fec) = 0 Then
        'Ha trabajado ese dia. Entonces lo pongo en HT
        If vM.PonerDiaBaja Then vM.Modificar
        
    Else
        MsgBox "El trabajador no tiene MARCAJES procesados ese dia", vbExclamation
    End If
    Set vM = Nothing
End Sub
