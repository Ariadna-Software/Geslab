VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "msadodc.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmColAnticipos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado anticpos"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   9600
   Icon            =   "frmColAnticipos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox Combo2 
      Height          =   315
      ItemData        =   "frmColAnticipos.frx":000C
      Left            =   7920
      List            =   "frmColAnticipos.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   6960
      Width           =   1335
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmColAnticipos.frx":0028
      Left            =   6960
      List            =   "frmColAnticipos.frx":0035
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   6960
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Text            =   "Text2"
      Top             =   6960
      Width           =   3435
   End
   Begin VB.TextBox txtTra 
      Height          =   285
      Left            =   2400
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   6960
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1200
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   6960
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   60
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   6960
      Width           =   1035
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8460
      TabIndex        =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9600
      _ExtentX        =   16933
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Modificar Lineas"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Deshacer pago"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Anterior"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
      Begin VB.CheckBox chkVistaPrevia 
         Caption         =   "Vista previa"
         Height          =   195
         Left            =   4560
         TabIndex        =   2
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "frmColAnticipos.frx":0044
      Height          =   6045
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   10663
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
   Begin VB.Label Label3 
      Caption         =   "Tipo"
      Height          =   195
      Left            =   7920
      TabIndex        =   14
      Top             =   6720
      Width           =   795
   End
   Begin VB.Label Label2 
      Caption         =   "Pagado"
      Height          =   195
      Left            =   6960
      TabIndex        =   12
      Top             =   6720
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   195
      Index           =   2
      Left            =   2400
      TabIndex        =   10
      Top             =   6720
      Width           =   840
   End
   Begin VB.Image ImgTrab 
      Height          =   240
      Left            =   3240
      Picture         =   "frmColAnticipos.frx":0059
      Top             =   6720
      Width           =   240
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   1
      Left            =   1740
      Picture         =   "frmColAnticipos.frx":015B
      Top             =   6720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   1
      Left            =   1200
      TabIndex        =   7
      Top             =   6720
      Width           =   420
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   0
      Left            =   720
      Picture         =   "frmColAnticipos.frx":025D
      Top             =   6720
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   255
      Index           =   0
      Left            =   60
      TabIndex        =   4
      Top             =   6720
      Width           =   555
   End
End
Attribute VB_Name = "frmColAnticipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1




Private Sub DataGrid1_DblClick()
    Modificar
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        CargaGrid
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
     PrimeraVez = True
     CargaCombo
     With Me.Toolbar1
        .ImageList = frmPpal1.imgListComun
        '.Buttons(1).Image = 1
        .Buttons(1).Visible = False
        .Buttons(2).Image = 2
        .Buttons(6).Image = 3
        .Buttons(7).Image = 4
        .Buttons(8).Image = 5
        '.Buttons(10).Image = 10
        .Buttons(11).Image = 16
        
        .Buttons(14).Image = 14
        .Buttons(16).Image = 15
        
        'Desplazamiento NO visible
        .Buttons(14).Visible = True
        
        'Eliminar pago
        .Buttons(15).Visible = False
        'Salir
        .Buttons(16).Visible = True
        
        '
        .Buttons(17).Visible = False
        
    End With
    Combo1.ListIndex = 0
    txtTra.Text = ""
    Text2.Text = ""
    Text1(0).Text = ""
    Text1_LostFocus 0
    Text1(1).Text = ""
    Text1_LostFocus 1
End Sub


Private Sub CargaGrid()
Dim I As Integer


    adodc1.ConnectionString = Conn
    adodc1.RecordSource = DevuelveSQL
    adodc1.Refresh

    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    I = 0
    DataGrid1.Columns(I).Width = 1000
    DataGrid1.Columns(I).Caption = "Fecha"
    DataGrid1.Columns(I).NumberFormat = "dd/mm/yyyy"
    
    I = 1
    DataGrid1.Columns(I).Width = 600
    DataGrid1.Columns(I).Caption = "Cod"
    
    I = 2
    DataGrid1.Columns(I).Width = 3000
    DataGrid1.Columns(I).Caption = "Nombre"
    
    I = 3
    DataGrid1.Columns(I).Width = 1000
    DataGrid1.Columns(I).Caption = "Importe"
    DataGrid1.Columns(I).NumberFormat = FormatoImporte
    DataGrid1.Columns(I).Alignment = dbgRight
    
    I = 4
    DataGrid1.Columns(I).Width = 1200
    DataGrid1.Columns(I).Caption = "Tipo"
    
    I = 5
    DataGrid1.Columns(I).Width = 1600
    DataGrid1.Columns(I).Caption = "Obsr."
    
    I = 6
    DataGrid1.Columns(I).Width = 400
    DataGrid1.Columns(I).Caption = "P"
    
    For I = 0 To 6
        DataGrid1.Columns(I).AllowSizing = False
    Next I
    
   DataGrid1.Columns(7).Visible = False
End Sub

Private Function DevuelveSQL() As String

    DevuelveSQL = "SELECT Pagos.Fecha, Pagos.Trabajador, Trabajadores.NomTrabajador, "
    DevuelveSQL = DevuelveSQL & "Pagos.Importe, TipoPago.Descripcion, Pagos.Observaciones"
    DevuelveSQL = DevuelveSQL & ",IIf([Pagado],""Si"","""") AS P, Pagos.Tipo"
    DevuelveSQL = DevuelveSQL & " FROM (Pagos INNER JOIN TipoPago ON Pagos.Tipo = TipoPago.idTipopago) INNER JOIN "
    DevuelveSQL = DevuelveSQL & "Trabajadores ON Pagos.Trabajador = Trabajadores.IdTrabajador"
    DevuelveSQL = DevuelveSQL & " WHERE Fecha >=#" & Format(Text1(0).Text, "yyyy/mm/dd")
    DevuelveSQL = DevuelveSQL & "# AND Fecha <=#" & Format(Text1(1).Text, "yyyy/mm/dd") & "#"
    'Pagado
    If Combo1.ListIndex > 0 Then
        DevuelveSQL = DevuelveSQL & " AND ("
        If Combo1.ListIndex = 2 Then DevuelveSQL = DevuelveSQL & " NOT "
        DevuelveSQL = DevuelveSQL & " Pagado )"
    End If
    If Combo2.ListIndex > 0 Then
        DevuelveSQL = DevuelveSQL & " AND Tipo ="
        DevuelveSQL = DevuelveSQL & Combo2.ItemData(Combo2.ListIndex)
    End If
    If txtTra.Text <> "" Then DevuelveSQL = DevuelveSQL & " AND Pagos.Trabajador=" & txtTra.Text
    DevuelveSQL = DevuelveSQL & " ORDER BY Fecha,Pagos.Trabajador"
End Function


Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    VariableCompartida = vCodigo
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Text1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub ImgFech_Click(Index As Integer)
        Set frmC = New frmCal
        Text1(0).Tag = Index
        frmC.Fecha = CDate(Text1(Index).Text)
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
        HacerToolBar 2
    End If
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_LostFocus(Index As Integer)

    With Text1(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not EsFechaOK(Text1(Index)) Then
                .Text = ""
            End If
        End If
        
        If .Text = "" Then
            If Index = 0 Then
                .Text = "01/" & Format(DateAdd("m", -1, Now), "mm/yyyy")
            Else
                .Text = Format(Now, "dd/mm/yyyy")
            End If
        End If
     End With
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index > 5 And Button.Index < 9 Then
        If vUsu.Nivel > 1 Then 'solo admon
            MsgBox "No tiene autorizacion para realizar cambios", vbExclamation
            Exit Sub
        End If
    End If
    HacerToolBar Button.Index
End Sub

Private Sub HacerToolBar(Indice As Integer)
    Select Case Indice
    Case 2
        Screen.MousePointer = vbHourglass
        CargaGrid
        Screen.MousePointer = vbDefault
    
    Case 6
        VariableCompartida = ""
        frmNuevoAn.registro = ""
        frmNuevoAn.Show vbModal
        If VariableCompartida <> "" Then
            Me.Refresh
            Screen.MousePointer = vbHourglass
            CargaGrid
            Screen.MousePointer = vbDefault
        End If
    Case 7
        Modificar
    Case 8
        Eliminar
        
    Case 14
        'Deshacer pago
        frmPagosBanco2.Opcion = 1
        frmPagosBanco2.Show vbModal
        HacerToolBar 2
    Case 16
        Unload Me
    End Select
End Sub



Private Sub txtTra_GotFocus()
    With txtTra
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
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
        End If
    Else
        Cad = ""
    End If
    Text2.Text = Cad
End Sub

Private Sub Modificar()


    If adodc1.Recordset.EOF Then Exit Sub

        VariableCompartida = adodc1.Recordset!Trabajador & "|" & adodc1.Recordset!Fecha & "|" & adodc1.Recordset!Tipo & "|"
        frmNuevoAn.registro = VariableCompartida
        VariableCompartida = ""
        frmNuevoAn.Show vbModal
        If VariableCompartida <> "" Then
            Me.Refresh
            Screen.MousePointer = vbHourglass
            CargaGrid
            Screen.MousePointer = vbDefault
        End If



End Sub


Private Sub Eliminar()

On Error GoTo EEliminar
    If adodc1.Recordset.EOF Then Exit Sub
    
    VariableCompartida = "¿Desea eliminar la entrada: " & vbCrLf
    VariableCompartida = VariableCompartida & "Fecha: " & adodc1.Recordset!Fecha & vbCrLf
    VariableCompartida = VariableCompartida & "Trabajador: " & adodc1.Recordset!nomtrabajador & vbCrLf
    VariableCompartida = VariableCompartida & "Importe: " & adodc1.Recordset!Importe & "  ?"
    If MsgBox(VariableCompartida, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
        VariableCompartida = "DELETE FROM Pagos WHERE Fecha =#" & Format(adodc1.Recordset!Fecha, FormatoFecha) & "#"
        VariableCompartida = VariableCompartida & " AND Trabajador= " & adodc1.Recordset!Trabajador
        VariableCompartida = VariableCompartida & " AND Tipo =  " & adodc1.Recordset!Tipo
        Screen.MousePointer = vbHourglass
        Conn.Execute VariableCompartida
        espera 0.5
        CargaGrid
        Screen.MousePointer = vbDefault
    End If

    Exit Sub
EEliminar:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub CargaCombo()
Dim RS As ADODB.Recordset

    Set RS = New ADODB.Recordset
    RS.Open "Select * from TipoPago", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Combo2.Clear
    Combo2.AddItem "*"
    Combo2.ItemData(Combo2.NewIndex) = 0
    While Not RS.EOF
        Combo2.AddItem RS.Fields(1)
        Combo2.ItemData(Combo2.NewIndex) = RS.Fields(0)
        RS.MoveNext
    Wend
    Combo2.ListIndex = 0
    RS.Close
    Set RS = Nothing
End Sub
