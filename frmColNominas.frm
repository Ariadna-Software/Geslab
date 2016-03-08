VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColNominas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado Nominas"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   10815
   Icon            =   "frmColNominas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   10815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            Object.Tag             =   "2"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
      Bindings        =   "frmColNominas.frx":000C
      Height          =   6045
      Left            =   60
      TabIndex        =   3
      Top             =   540
      Width           =   10590
      _ExtentX        =   18680
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
      Picture         =   "frmColNominas.frx":0021
      Top             =   6720
      Width           =   240
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   1
      Left            =   1740
      Picture         =   "frmColNominas.frx":0123
      Top             =   6660
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
      Picture         =   "frmColNominas.frx":0225
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
Attribute VB_Name = "frmColNominas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim primeravez As Boolean
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmBu As frmBuscaGrid
Attribute frmBu.VB_VarHelpID = -1




Private Sub DataGrid1_DblClick()
    Modificar
End Sub

Private Sub Form_Activate()
    If primeravez Then
        primeravez = False
        CargaGrid
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
     primeravez = True
     
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
        .Buttons(12).Image = 15
        'Desplazamiento NO visible
        .Buttons(14).Visible = False
        .Buttons(15).Visible = False
        .Buttons(16).Visible = False
        .Buttons(17).Visible = False
    End With
    
    txtTra.Text = ""
    Text2.Text = ""
    Text1(0).Text = ""
    Text1_LostFocus 0
    Text1(1).Text = ""
    Text1_LostFocus 1
End Sub


Private Sub CargaGrid()
Dim i As Integer


    adodc1.ConnectionString = conn
    adodc1.RecordSource = DevuelveSQL
    adodc1.Refresh

    DataGrid1.AllowRowSizing = False
    DataGrid1.RowHeight = 320
    
    i = 0
    DataGrid1.Columns(i).Width = 600
    DataGrid1.Columns(i).Caption = "Cod"
    
    
    i = 1
    DataGrid1.Columns(i).Width = 3500
    DataGrid1.Columns(i).Caption = "Nombre"
    
    i = 2
    DataGrid1.Columns(i).Width = 1100
    DataGrid1.Columns(i).Caption = "Fecha"
    DataGrid1.Columns(i).NumberFormat = "dd/mm/yyyy"
    
    i = 3
    DataGrid1.Columns(i).Width = 600
    DataGrid1.Columns(i).Caption = "Dias"
    DataGrid1.Columns(i).Alignment = dbgRight
    
    i = 4
    DataGrid1.Columns(i).NumberFormat = FormatoImporte
    DataGrid1.Columns(i).Width = 800
    DataGrid1.Columns(i).Caption = "HN"
    DataGrid1.Columns(i).Alignment = dbgRight
    
    i = 5
    DataGrid1.Columns(i).Width = 800
    DataGrid1.Columns(i).Caption = "HC"
    DataGrid1.Columns(i).NumberFormat = FormatoImporte
    DataGrid1.Columns(i).Alignment = dbgRight
    
    i = 6
    DataGrid1.Columns(i).Width = 1300
    DataGrid1.Columns(i).Caption = "Anticipos"
    DataGrid1.Columns(i).NumberFormat = FormatoImporte
    DataGrid1.Columns(i).Alignment = dbgRight
    
    
    i = 7
    DataGrid1.Columns(i).Width = 1000
    DataGrid1.Columns(i).Caption = "Trabajados"
    DataGrid1.Columns(i).Alignment = dbgRight
    DataGrid1.Columns(i).Visible = MiEmpresa.QueEmpresa = 0
    
    
    For i = 0 To 7
        DataGrid1.Columns(i).AllowSizing = False
    Next i
    
   'DataGrid1.Columns(7).Visible = False
End Sub

Private Function DevuelveSQL() As String

    DevuelveSQL = "SELECT Nominas.idTrabajador, Trabajadores.NomTrabajador, Nominas.Fecha, Nominas.Dias, Nominas.HN, Nominas.HC, Nominas.Anticipos,"
    If MiEmpresa.QueEmpresa = 0 Then
        DevuelveSQL = DevuelveSQL & "DiasTra "
    Else
        DevuelveSQL = DevuelveSQL & "HN "
    End If
    DevuelveSQL = DevuelveSQL & " FROM Trabajadores INNER JOIN Nominas ON Trabajadores.IdTrabajador = Nominas.idTrabajador"
    DevuelveSQL = DevuelveSQL & " WHERE Fecha >=#" & Format(Text1(0).Text, "yyyy/mm/dd")
    DevuelveSQL = DevuelveSQL & "# AND Fecha <=#" & Format(Text1(1).Text, "yyyy/mm/dd") & "#"
   
'    'Pagado
'    If Combo1.ListIndex > 0 Then
'        DevuelveSQL = DevuelveSQL & " AND "
'        If Combo1.ListIndex = 2 Then DevuelveSQL = DevuelveSQL & " NOT "
'        DevuelveSQL = DevuelveSQL & " Pagado"
'    End If
'    If Combo2.ListIndex > 0 Then
'        DevuelveSQL = DevuelveSQL & " AND Tipo ="
'        DevuelveSQL = DevuelveSQL & Combo2.ItemData(Combo2.ListIndex)
'    End If
    If txtTra.Text <> "" Then DevuelveSQL = DevuelveSQL & " AND Nominas.idTrabajador=" & txtTra.Text
    
    DevuelveSQL = DevuelveSQL & " ORDER BY Nominas.idTrabajador"
End Function




Private Sub frmBu_Selecionado(CadenaDevuelta As String)
VariableCompartida = CadenaDevuelta
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
    Set frmBu = New frmBuscaGrid
    frmBu.vBusqueda = ""
    frmBu.vCampos = "Codigo|Trabajadores|idtrabajador|N|000|25·Nombre|Trabajadores|nomtrabajador|T||65·"
    frmBu.vDevuelve = "0|1|"
    frmBu.vselElem = 0
    frmBu.vSQL = ""
    frmBu.vTabla = "trabajadores"
    frmBu.vTitulo = "Trabajadores"
    frmBu.Show vbModal
    Set frmBu = Nothing

    If VariableCompartida <> "" Then
        txtTra.Text = RecuperaValor(VariableCompartida, 1)
        Text2.Text = RecuperaValor(VariableCompartida, 2)
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
            If Not IsDate(.Text) Then
                .Text = ""
            Else
                .Text = Format(.Text, "dd/mm/yyyy")
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
            MsgBox "No tiene autorización para realizar cambios", vbExclamation
            Exit Sub
        End If
    End If
    HacerToolBar Button.Index
End Sub

Private Sub HacerToolBar(Indice As Integer)
    Select Case Indice
    Case 2
        Screen.MousePointer = vbHourglass
        espera 1
        CargaGrid
        Screen.MousePointer = vbDefault
    
    Case 6
 
    Case 7
        Modificar
    Case 8
        Eliminar
    Case 11
         


        With frmImprimir
            .Opcion = 165
            .NumeroParametros = 2
            .OtrosParametros = "fechaini= """ & Text1(0).Text & """|" & "fechafin= """ & Text1(1).Text & """|"
            If txtTra.Text <> "" Then
                .NumeroParametros = 4
                .OtrosParametros = .OtrosParametros & "t1= " & txtTra.Text & "|" & "t2= " & txtTra.Text & "|"
            End If
            .Show vbModal
        End With
    Case 12
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
Dim L As Long

    If adodc1.Recordset.EOF Then Exit Sub

    If Not (MiEmpresa.QueEmpresa = 0) Then
        MsgBox "No disponible para no TCP3", vbExclamation
        Exit Sub
    End If

        frmCambiosDatosNomina.Opcion = 1
        Load frmCambiosDatosNomina
        
        VariableCompartida = "" 'Si guarda o no guarda
      
            
        With adodc1.Recordset
            frmCambiosDatosNomina.Caption = !Fecha
            frmCambiosDatosNomina.lblIdTra(1) = !idTrabajador
            frmCambiosDatosNomina.lblTra(1) = " - " & !nomtrabajador             'Trabajador
            
            'Dias
            frmCambiosDatosNomina.txtDias(3).Text = !Dias
            'TRABAJADAS
            frmCambiosDatosNomina.txtDias(4).Text = DBLet(!DiasTra, "N")
            'HN
            frmCambiosDatosNomina.txtHN(3).Text = !HN
            'HC
            frmCambiosDatosNomina.txtHN(4).Text = !HC
            'Anticipos
            frmCambiosDatosNomina.txtHN(5).Text = !Anticipos
            

            
            frmCambiosDatosNomina.Show vbModal
        End With
        If VariableCompartida <> "" Then
            Screen.MousePointer = vbHourglass
            L = adodc1.Recordset.AbsolutePosition
            espera 1
            CargaGrid
            If L > 1 Then adodc1.Recordset.Move L - 1, 1
            Screen.MousePointer = vbDefault
        End If
End Sub


Private Sub Eliminar()

On Error GoTo EEliminar
    If adodc1.Recordset.EOF Then Exit Sub
    
  

    Exit Sub
EEliminar:
    MuestraError Err.Number, Err.Description
End Sub


'Private Sub CargaCombo()
'Dim rs As ADODB.Recordset
'
'    Set rs = New ADODB.Recordset
'    rs.Open "Select * from TipoPago", Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Combo2.Clear
'    Combo2.AddItem "*"
'    Combo2.ItemData(Combo2.NewIndex) = 0
'    While Not rs.EOF
'        Combo2.AddItem rs.Fields(1)
'        Combo2.ItemData(Combo2.NewIndex) = rs.Fields(0)
'        rs.MoveNext
'    Wend
'    Combo2.ListIndex = 0
'    rs.Close
'    Set rs = Nothing
'End Sub
