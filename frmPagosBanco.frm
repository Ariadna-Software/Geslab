VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmPagosBanco2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pagos banco"
   ClientHeight    =   3675
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6300
   Icon            =   "frmPagosBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   6300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6120
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      ItemData        =   "frmPagosBanco.frx":030A
      Left            =   3600
      List            =   "frmPagosBanco.frx":0314
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Frame FrameDeshacer 
      Height          =   3615
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   1920
         TabIndex        =   14
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Todos los pagos del tipo que seleccione a la derecha y con fecha de pago  pasarán a estar pendientes de pago."
         Height          =   1215
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Deshacer pago."
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de pago"
         Height          =   255
         Index           =   4
         Left            =   240
         TabIndex        =   15
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Image ImgFech 
         Height          =   240
         Index           =   3
         Left            =   1440
         Picture         =   "frmPagosBanco.frx":032F
         Top             =   2280
         Width           =   240
      End
   End
   Begin VB.TextBox Text3 
      Height          =   1485
      Left            =   240
      MaxLength       =   34
      MultiLine       =   -1  'True
      TabIndex        =   18
      Top             =   1560
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   240
      MaxLength       =   34
      TabIndex        =   3
      Top             =   1680
      Width           =   2475
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Index           =   1
      Left            =   4980
      TabIndex        =   5
      Top             =   3120
      Width           =   1155
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Aceptar"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   4
      Top             =   3120
      Width           =   1155
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Fichero banco"
      Height          =   255
      Index           =   1
      Left            =   1500
      TabIndex        =   10
      Top             =   3240
      Width           =   1395
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2115
      Left            =   3600
      TabIndex        =   9
      Top             =   360
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   3731
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      _Version        =   393217
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
      NumItems        =   0
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   360
      Width           =   1095
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Listado"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   7
      Top             =   3240
      Value           =   -1  'True
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.ComboBox cboBanco 
      Height          =   315
      Left            =   240
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2520
      Width           =   3135
   End
   Begin VB.Label Label4 
      Caption         =   "Banco"
      Height          =   255
      Left            =   240
      TabIndex        =   23
      Top             =   2280
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Tipos de pago:"
      Height          =   255
      Index           =   6
      Left            =   3600
      TabIndex        =   20
      Top             =   120
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto INFORME"
      Height          =   255
      Index           =   5
      Left            =   240
      TabIndex        =   19
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Concepto BANCO"
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   1380
      Width           =   2175
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   2
      Left            =   1200
      Picture         =   "frmPagosBanco.frx":0431
      Top             =   720
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha orden"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   11
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "Fin"
      Height          =   255
      Index           =   1
      Left            =   1680
      TabIndex        =   8
      Top             =   120
      Width           =   555
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   1
      Left            =   2340
      Picture         =   "frmPagosBanco.frx":0533
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Inicio"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   6
      Top             =   120
      Width           =   555
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   0
      Left            =   900
      Picture         =   "frmPagosBanco.frx":0635
      Top             =   120
      Width           =   240
   End
End
Attribute VB_Name = "frmPagosBanco2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private Sub Command1_Click(Index As Integer)
Dim I As Integer


    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    
    If Opcion = 0 Then
    
        If Text1(0).Text = "" Or Text1(1).Text = "" Then
            MsgBox "Ponga las fechas", vbExclamation
            Exit Sub
        End If
    
        If Option1(1).Value Then
            If Text1(2).Text = "" Then
                MsgBox "Escriba la fecha del anticipo", vbExclamation
                Exit Sub
            End If
        End If
    Else
        If Text1(3).Text = "" Then
            MsgBox "Ponga las fecha", vbExclamation
            Exit Sub
        End If
    
    End If
    
    ListView1.Tag = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            ListView1.Tag = "D"
            Exit For
        End If
    Next I

    If ListView1.Tag = "" Then
        MsgBox "Seleccione algun modo de pago", vbExclamation
        Exit Sub
    End If
    
    
    If Opcion = 0 Then
            If Option1(1).Value Then
                If Combo1.ListIndex < 0 Then
                    MsgBox "Seleccione el tipo de transferencia", vbExclamation
                    Exit Sub
                End If
                    
                'BANCO
                ListView1.Tag = "Este proceso generara todos los pagos y grabara el fichero correspondiente"
                If cboBanco.ListIndex > 0 Then ListView1.Tag = ListView1.Tag & vbCrLf & vbCrLf & "Banco: " & cboBanco.List(cboBanco.ListIndex) & vbCrLf
                ListView1.Tag = ListView1.Tag & vbCrLf & "¿   Desea continuar  ?"
                If MsgBox(ListView1.Tag, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub
            End If
            
            'AHora generamos los valores
            Screen.MousePointer = vbHourglass
            'Generamos la tabla con los datos. tmpNorma34
            If GeneraTMPnorma34 Then
            
                'AHor los tenemos en la tabla tmpNorma34
            
                If Option1(1).Value Then
                    'BCANO
                    conn.BeginTrans
                    If GenerarDatosPagos Then
                       conn.CommitTrans
                       Command1(0).Enabled = False
                    Else
                        conn.RollbackTrans
                    End If
                Else
                    'LISTADITO
                    espera 1
                    ImprimeListadito
                End If
            End If
    Else
        
            Screen.MousePointer = vbHourglass
            DeshacerPago
        
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub ImprimeListadito()

     
    CadParam = "Campo1= 'Fecha orden'|"
    CadParam = CadParam & "Campo2= '" & Text1(2).Text & "'|"
    CadParam = CadParam & "Campo3= 'Concepto: '|"
    CadParam = CadParam & "Campo4= '" & Text3.Text & "'|"
    With frmImprimir
        .Opcion = 110
        .NumeroParametros = 4
        .OtrosParametros = CadParam
        .FormulaSeleccion = ""
        .Show vbModal
    End With

End Sub


Private Sub Form_Load()
    If Opcion = 0 Then
        Text1_LostFocus 0
        Text1_LostFocus 1
        Text1_LostFocus 2
        PonOpcionCombo True
    End If
    
    Combo1.Visible = Opcion = 0
    Me.FrameDeshacer.Visible = (Opcion = 1)
    CargaList
    CargaBanco
End Sub

Private Sub Form_Unload(Cancel As Integer)
    PonOpcionCombo False
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Text1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub ImgFech_Click(Index As Integer)
        Set frmC = New frmCal
        Text1(0).Tag = Index
        frmC.Fecha = Now
        If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
        frmC.Show vbModal
        Set frmC = Nothing
End Sub

Private Sub Option1_Click(Index As Integer)
    Label1(3).Visible = Option1(1).Value
    Label1(5).Visible = Not Label1(3).Visible
    Text2.Visible = Label1(3).Visible
    Text3.Visible = Not Text2.Visible
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
            If Not EsFechaOK(Text1(Index)) Then .Text = ""
            
        End If
        
        If .Text = "" Then
            If Index = 0 Then
                .Text = "01/" & Format(Now, "mm/yyyy")
            Else
                .Text = CStr(DiasMes(Month(Now), Year(Now))) & "/" & Format(Now, "mm/yyyy")
            End If
        End If
     End With
End Sub



Private Sub CargaList()
Dim RS As ADODB.Recordset
Dim itmX As ListItem

    Set RS = New ADODB.Recordset
    RS.Open "Select * from TipoPago order by idTipoPago", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    ListView1.ListItems.Clear
    While Not RS.EOF
        Set itmX = ListView1.ListItems.Add(, , RS.Fields(1))
        itmX.Tag = RS.Fields(0)
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    If ListView1.ListItems.Count > 1 Then ListView1.ListItems(2).Checked = True
    Command1(0).Enabled = (ListView1.ListItems.Count > 0)
End Sub


Private Function GenerarDatosPagos() As Boolean
Dim NIF As String
Dim Cta As String
Dim RT As ADODB.Recordset
Dim AUX As String
Dim Sufijo As String
    '
    'Hay que generar diskette del banco
    GenerarDatosPagos = False
    
    'Obtenemos datos de la empresa en lo referente a NIF, Cuenta bancaria
    ObtenerDatosEmpresa NIF, Cta, Sufijo
    If NIF = "" Then Exit Function
    
    
    
    
    If GeneraFicheroNorma34SEPA_(NIF, CDate(Text1(2).Text), Cta, CStr(Combo1.ItemData(Combo1.ListIndex)), Sufijo) Then
    
        'octubre 2011
        'Pediremos el destino. Si cancela, ya existe... no ejecutaremos el SQL
        cd1.InitDir = App.Path
        cd1.DialogTitle = "Seleccione destino fichero norma 34"
        cd1.CancelError = False
        cd1.ShowSave
        If cd1.FileTitle = "" Then Exit Function 'HA cancelado.
        Cta = "OK"
        
        If Dir(cd1.FileName, vbArchive) <> "" Then
            Cta = "El archivo " & cd1.FileName & " ya existe" & vbCrLf & vbCrLf & "¿Sobreescribir?"
            If MsgBox(Cta, vbQuestion + vbYesNo) = vbNo Then Cta = ""
        End If
        
        If Cta = "" Then Exit Function
        Cta = ""
        
        'Primero copiamos el fichero, y si va bien, actualizamos
        If CopiarFicheroNorma43_(cd1.FileName) Then

                'Actualizamos la BD poniendo pagado = true
                PonSQL Cta, False
                Set RT = New ADODB.Recordset
                RT.Open Cta, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                NIF = "UPDATE Pagos Set Pagado= TRUE  WHERE "
                While Not RT.EOF
                    Cta = NIF & "Trabajador =  " & RT!Trabajador
                    Cta = Cta & " AND Tipo = " & RT!Tipo
                    Cta = Cta & " AND Fecha = #" & Format(RT!Fecha, FormatoFecha) & "#"
                    conn.Execute Cta
            
            
                    RT.MoveNext
            
                Wend
                RT.Close
                Set RT = Nothing
                GenerarDatosPagos = True
        End If


        
        
        '
        
     End If
        
End Function


Private Function GeneraTMPnorma34() As Boolean
Dim vSQL As String
Dim I As Integer
Dim Cad As String
Dim AUX As String
Dim RS As ADODB.Recordset

    On Error GoTo EGeneraTMPnorma34
    GeneraTMPnorma34 = False

    conn.Execute "Delete from tmpNorma34"

    'Buscamos con los valores del form los pagos que faltan hacerse
    PonSQL vSQL, False
    
    Set RS = New ADODB.Recordset
    RS.Open vSQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    If RS.EOF Then
        'Ningun datos con esos valores
        MsgBox "Ningun registro con esos valores", vbExclamation
    Else
        vSQL = "INSERT INTO tmpNorma34 (CodSoc, Nombre, Banco1, Banco2, Banco3, Banco4, Domicilio, Codpos, Poblacion, Concepto, Importe,tipo) VALUES ("
    
        While Not RS.EOF
            Cad = RS!Trabajador & ",'"
            Cad = Cad & ASQL(RS!nomtrabajador) & "','"
            'Cerramos el rs
            If PonerDatosBancos(RS, Cad) Then
                Cad = Cad & ",'" & ASQL(DBLet(RS!domtrabajador))
                Cad = Cad & "','" & DBLet(RS!codpostrabajador)
                Cad = Cad & "','" & ASQL(DBLet(RS!pobtrabajador)) & "','" & Text2.Text & "',"
                Cad = Cad & TransformaComasPuntos(CStr(RS!Importe)) & ","
                Cad = Cad & RS!Tipo & ")"
                'Insertamos
                conn.Execute vSQL & Cad
                I = I + 1
            End If
            'Concepto
            RS.MoveNext
        
        Wend
    End If
    RS.Close
    Set RS = Nothing
    If I > 0 Then
        GeneraTMPnorma34 = True
    Else
        MsgBox "Ningun dato generado", vbExclamation
    End If
    Exit Function
EGeneraTMPnorma34:
    MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Function



Private Sub PonSQL(ByRef SQL As String, vDeshacerPago As Boolean)
Dim Cad As String
Dim I As Integer

    Cad = ""
    SQL = ""
    For I = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(I).Checked Then
            If Cad <> "" Then Cad = Cad & " OR "
            Cad = Cad & "tipo = " & ListView1.ListItems(I).Tag
            SQL = SQL & "1"
        End If
    Next I

    If Len(SQL) = ListView1.ListItems.Count Then
        'TODOS, luego cad no ponemos subconsultas
        Cad = ""
    Else
        Cad = "(" & Cad & ")"
    End If
            
    
    'Vamos a coger los pagos k faltan
    If vDeshacerPago Then
        'Es un update
        SQL = "UPDATE Pagos SET Pagado = FALSE "
    Else
        SQL = "SELECT Pagos.Trabajador, Pagos.Importe, Pagos.Pagado, Pagos.Fecha, "
        SQL = SQL & " Trabajadores.NomTrabajador, Trabajadores.DomTrabajador, Trabajadores.PobTrabajador"
        SQL = SQL & ", Trabajadores.CodPosTrabajador, Trabajadores.entidad, Trabajadores.oficina, Trabajadores.pagobancario,"
        SQL = SQL & " Trabajadores.controlcta, Trabajadores.cuenta,Pagos.tipo,Trabajadores.pagobancario"
        SQL = SQL & " FROM Trabajadores INNER JOIN Pagos ON Trabajadores.IdTrabajador = Pagos.Trabajador"
    End If
    
    SQL = SQL & " WHERE Pagos.Pagado= "
    
    If vDeshacerPago Then
        SQL = SQL & " TRUE"
    Else
        SQL = SQL & " False"
    End If
    'Las fechas
    If Not vDeshacerPago Then
        SQL = SQL & " AND Fecha >=#" & Format(Text1(0).Text, FormatoFecha) & "#"
        SQL = SQL & " AND Fecha <=#" & Format(Text1(1).Text, FormatoFecha) & "#"
    Else
        SQL = SQL & " AND Fecha = #" & Format(Text1(3).Text, FormatoFecha) & "#"
    End If
    'Los tipos
    If Cad <> "" Then SQL = SQL & " AND " & Cad
    
    
End Sub

Private Function PonerDatosBancos(ByRef RS1 As ADODB.Recordset, ByRef C As String) As Boolean
Dim OK As Boolean
    OK = False
        'Paga por banco o no
    If RS1!pagobancario Then
        If Not IsNull(RS1!Entidad) Then
            C = C & Format(RS1!Entidad, "0000") & "','"
            If Not IsNull(RS1!oficina) Then
                C = C & Format(RS1!oficina, "0000") & "','"
                If Not IsNull(RS1!controlcta) Then
                    C = C & Mid(DBLet(RS1!controlcta) & "  ", 1, 2) & "','"
                    If Not IsNull(RS1!cuenta) Then
                        C = C & Right("0000000000" & RS1!cuenta, 10) & "'"
                        OK = True
                    End If
                End If
            End If
        End If
    Else
        'NO .
        C = C & " ',' ','NO','BANCO'"
        If Option1(1).Value Then
            OK = False
        Else
            OK = True
        End If
    End If

    PonerDatosBancos = OK
'    If Not OK Then
'        MsgBox "Error en la cuenta bancaria para : " & RS1!nomtrabajador & vbCrLf & vbCrLf & C, vbExclamation
'        RS1.Close
'    End If
End Function





Private Sub ObtenerDatosEmpresa(ByRef vNIF As String, ByRef vCTA As String, ByRef Sufijo As String)
Dim RS As ADODB.Recordset

    
    vNIF = "Select * from empresas"
    Set RS = New ADODB.Recordset
    RS.Open vNIF, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    vNIF = ""
    If Not RS.EOF Then
        vNIF = DBLet(RS!CIF)
        
        'LA cuenta
        If vNIF <> "" Then
            
            If Me.cboBanco.ListIndex > 0 Then
                'Abro EL RS con bancos para el seleccionado en el combo
                RS.Close
                vCTA = "Select * from bancos where id = " & Me.cboBanco.ItemData(cboBanco.ListIndex)
                RS.Open vCTA, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            End If
            
            If IsNull(RS!Entidad) Or IsNull(RS!Sucursal) Or IsNull(RS!codcontrol) Or IsNull(RS!cuenta) Then
                vCTA = "Cuenta con datos vacios"
                vNIF = ""
            Else
                Sufijo = Right("000" & DBLet(RS!sufijoN34, "T"), 3)
                vCTA = RS!Entidad & "|" & RS!Sucursal & "|" & RS!codcontrol & "|" & RS!cuenta & "|" & Right("0000" & DBLet(RS!IBAN, "T"), 4) & "|"
                If Len(vCTA) <> 29 Then
                    vCTA = "Longitud de la cuenta bancaria incorrecta"
                    vNIF = ""
                End If
            End If
        Else
            vCTA = "NIF vacio"
        End If
    End If
    RS.Close
    
    If vNIF = "" Then MsgBox vCTA, vbExclamation
End Sub

Private Sub DeshacerPago()
Dim vSQL As String
    
    
    
    'Buscamos con los valores del form los pagos que ya estan hechos
    PonSQL vSQL, True

    conn.Execute vSQL
    espera 0.5
    MsgBox "Proceso finalizado", vbExclamation
End Sub


Private Sub PonOpcionCombo(Leer As Boolean)
'El combo tendra 2 opciones
    'transferencia : itemdata: 9
    ' nomina       :    ""     1


Dim NombrArchivo As String
    NombrArchivo = App.Path & "\Valnorma19.dat"
    If Leer Then
        If Dir(NombrArchivo) = "" Then
            Combo1.ListIndex = 0
        Else
            Combo1.ListIndex = 1
        End If
    Else
        If Combo1.ListIndex = 0 Then
            If Dir(NombrArchivo) <> "" Then Kill NombrArchivo
        Else
            If Dir(NombrArchivo) = "" Then
                Opcion = CByte(FreeFile)
                Open NombrArchivo For Output As CInt(Opcion)
                Close CInt(Opcion)
            End If
        End If
    End If
End Sub

Private Sub CargaBanco()
Dim RS As ADODB.Recordset
On Error GoTo ecboBanco
    
    cboBanco.Clear
    'Banco de parametros empresa
    cboBanco.AddItem "Habitual"
    
    Set RS = New ADODB.Recordset
    RS.Open "Select * from bancos order by id", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        cboBanco.AddItem RS!Observa & " (" & RS!Entidad & " / " & RS!Sucursal & ")"
        cboBanco.ItemData(cboBanco.NewIndex) = RS!Id
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
    cboBanco.ListIndex = 0
    
    
    
ecboBanco:
    Err.Clear
End Sub
