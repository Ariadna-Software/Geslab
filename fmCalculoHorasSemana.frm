VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCalculoHorasSemana 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HORAS SEMANA"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   Icon            =   "fmCalculoHorasSemana.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   11850
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   11775
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   6960
         TabIndex        =   17
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Código"
         Height          =   255
         Index           =   0
         Left            =   5760
         TabIndex        =   16
         Top             =   480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenHoras 
         Caption         =   "Calcular"
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cerrar semana"
         Height          =   375
         Left            =   8760
         TabIndex        =   2
         Top             =   360
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2760
         TabIndex        =   14
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Fin"
         Height          =   255
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   495
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   600
         Picture         =   "fmCalculoHorasSemana.frx":030A
         Top             =   360
         Width           =   240
      End
      Begin VB.Label Label8 
         Caption         =   "Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   495
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   4080
      TabIndex        =   9
      Top             =   4560
      Width           =   4095
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   315
         Left            =   300
         TabIndex        =   10
         Top             =   360
         Width           =   3615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   60
      Top             =   4800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasSemana.frx":040C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasSemana.frx":09A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasSemana.frx":0F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasSemana.frx":125A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasSemana.frx":17F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "fmCalculoHorasSemana.frx":1C46
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4215
      Left            =   60
      TabIndex        =   3
      Top             =   1320
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   7435
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   17
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "D / H"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "HN"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "HC"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "HE"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Hor"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "D"
         Object.Width           =   706
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "H"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "HE"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "Ant"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Post"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   14
         Text            =   "Anticipos"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(16) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   15
         Text            =   "Plus"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(17) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   16
         Text            =   "BORRAR"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Bolsa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   195
      Left            =   7380
      TabIndex        =   11
      Top             =   960
      Width           =   480
   End
   Begin VB.Line Line6 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   2
      X1              =   7260
      X2              =   8400
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00008080&
      BorderWidth     =   2
      X1              =   6000
      X2              =   7140
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00008000&
      BorderWidth     =   2
      X1              =   4740
      X2              =   5880
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00800000&
      BorderWidth     =   2
      X1              =   3420
      X2              =   4560
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   2100
      X2              =   3240
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   120
      X2              =   1860
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Semana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   195
      Left            =   6120
      TabIndex        =   8
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Saldo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   195
      Left            =   4980
      TabIndex        =   7
      Top             =   960
      Width           =   495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Trabajadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   3600
      TabIndex        =   6
      Top             =   960
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Oficial"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2220
      TabIndex        =   5
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Top             =   960
      Width           =   930
   End
End
Attribute VB_Name = "frmCalculoHorasSemana"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private SQL As String
Dim EmpresaHoraExtra As Boolean


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1




'Private Sub cmdBaja_Click()
'Dim DiasC As Integer
' 'Marcaremos esta casilla para decir k el trabajador ha sido dado de baja/alta si ya estaba
'    If ListView1.ListItems.Count = 0 Then Exit Sub
'    If ListView1.SelectedItem Is Nothing Then Exit Sub
'
'    'Esta de baja
'    SQL = "Va a cambiar la opcion BAJA del trabajador: " & ListView1.SelectedItem.SubItems(1)
'    SQL = SQL & vbCrLf & "     ¿Desea continuar ?"
'
'
'
'    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) = vbYes Then
'        'Si pone de baja pero puede compensar dias entonces preguntamos si compensadias o
'        ' lo manda todo a extras
'        If Not ListView1.SelectedItem.Bold Then
'            'Si tiene dias para compensar
'            DiasC = PuedeCompensarDias
'            If DiasC > 0 Then
'                'Hacemos la pregunta si desea
'                SQL = "El trabajador tiene la posibilidad de compensar dias." & vbCrLf
'                SQL = SQL & "  ¿Desea compensarle los  dias ?"
'                If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then CompensarDias (DiasC)
'            End If
'        End If
'
'
'        ListView1.SelectedItem.Bold = Not ListView1.SelectedItem.Bold
'
'        If ListView1.SelectedItem.Bold Then
'            ListView1.SelectedItem.SmallIcon = 3
'            PonerBaja True, 0
'        Else
'            PonerBaja False, 0
'        End If
'    End If
'End Sub

Private Sub cmdGenHoras_Click()
Dim D As Integer
Dim FI As Date
Dim FF As Date


    If Text1.Text = "" Then
        MsgBox "Fecha incorrecta.", vbExclamation
        Exit Sub
    End If
        
    If ListView1.ListItems.Count > 0 Then
        SQL = "Ya ha generado datos. ¿ Seguro que desea volverlos a generar ?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        ListView1.ListItems.Clear
    End If
        
    FI = CDate(Text1.Text)
    FF = CDate(Text2.Text)
        
        
    If ComprobarMarcajesCorrectos(FI, FF, True) = 0 Then
        SQL = "No existe marcajes entre las fechas."
        MsgBox SQL, vbExclamation
        Exit Sub
    End If
        
    If ComprobarMarcajesCorrectos(FI, FF, False) <> 0 Then
        SQL = "Existen marcajes incorrectos entre las fechas. ¿Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
          
    Label1.Caption = "Comienzo proceso"
    Frame1.Visible = True
    Me.Refresh
    
    Screen.MousePointer = vbHourglass
    
    ListView1.ListItems.Clear
    
    CalculaEntreFechas FI, FF
    Frame1.Visible = False
    Screen.MousePointer = vbDefault

End Sub


Private Sub CalculaEntreFechas(FI As Date, FF As Date)
Dim Rs As Recordset
Dim Horas As Currency
Dim Dias As Integer

    
    conn.Execute "DELETE FROM tmpHorasMesHorario"

    'Para comprobar si estando de baja han trabajado
    'En tmpPresencia voy a guardar
    conn.Execute "DELETE FROM tmpCombinada"

    Set Rs = New ADODB.Recordset
    Rs.Open "horarios", conn, adOpenForwardOnly, adLockPessimistic, adCmdTable
    
    Label1.Caption = "Obtener horarios"
    Label1.Refresh
    
    While Not Rs.EOF
        Horas = CalculaHorasHorario(Rs.Fields(0), Dias, FI, FF, False)
        If Horas > 0 Then
            'Insertamos en tmp HORAS
            conn.Execute "INSERT INTO tmpHorasMesHorario(idHorario,Horas,Dias) VALUES (" & Rs.Fields(0) & "," & TransformaComasPuntos(CStr(Horas)) & "," & Dias & ")"
        End If
        Rs.MoveNext
    Wend
    Rs.Close
    
    
    Label1.Caption = "Horas trabajadas"
    Label1.Refresh
    'Modificacion
    If Not EmpresaHoraExtra Then
        CalculaHorasTrabajadas FI, FF, 1, -1
    Else
        CalculaHorasTrabajadasConEXTRAS FI, FF, 1
    End If
    Me.Refresh
    
    Label1.Caption = "Datos periodo"
    Label1.Refresh
    CalculaDatosMes FI, FF, 1, -1
    
    Me.Refresh
    
    Label1.Caption = "Combina datos"
    Label1.Refresh
    CombinaDatos FI, FF
    
    'AHora realizamos los calculos de horas k kedan y demas
    Label1.Caption = "Datos a compensar"
    Label1.Refresh
    CalculoDatosACompensar
    
    Me.Refresh
    
    'Hacemos las comensaciones por horas
    Label1.Caption = "Compensaciones"
    Label1.Refresh
    HacerCompensaciones FI, FF, Label1
    
    
    
    'Hacemos compensacion semana
    'CATADAU no lo hace
    If Not EmpresaHoraExtra Then
        Label1.Caption = "Horas maximas / dias trabajdos"
        Label1.Refresh

        HacerCompensacionSememana FI, FF
    End If
    
    
    'No lleva bolsa horas, es para catadau
    If EmpresaHoraExtra Then PonHorasExtraDeBolsa
    
    
    
    'Ajustamos los que no hayan trabakado nada
    AjustaDatosBajaMesEntero
    

    Label1.Caption = "Carga datos"
    Label1.Refresh
    CargaDatos



    'Ahora vamos a comprobar si alguno de los k ha estado de baja
    'En este periodo a trabajado
    If ListView1.ListItems.Count > 0 Then
        Label1.Caption = "Comprobar bajas con dias Tra."
        Label1.Refresh
        Rs.Open "Select * from tmpcombinada", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        While Not Rs.EOF
            If HaTrabajadoConBaja(Rs) Then
                Dias = 0
                Do
                    Dias = Dias + 1
                    If Dias <= ListView1.ListItems.Count Then
                        If ListView1.ListItems(Dias).Text = Rs!idTrabajador Then
                            'Pongo el icono distinto
                            ListView1.ListItems(Dias).SmallIcon = 5
                            'Salgo
                            Dias = 32000
                        End If
                    End If
                Loop Until Dias > ListView1.ListItems.Count
            End If   'De ha trabajado estando de baja
            'Siguiente caso
        Rs.MoveNext
        Wend
        Rs.Close
    End If
    
    Set Rs = Nothing
End Sub






'Private Sub cmdHPlus_Click(Index As Integer)
'Dim Importe As Currency
'Dim Imp1 As Currency
'Dim RS As ADODB.Recordset
'
'    If ListView1.SelectedItem Is Nothing Then Exit Sub
'
'    If Index = 1 Then
'        SQL = "reestablecer horas plus"
'    Else
'        SQL = "añadir horas plus"
'    End If
'    If MsgBox("Desea continuar con la opción " & SQL & " ?", vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
'
'    If Index = 0 Then
'        'Si ya ha compensado le decimos k ya ha compensado
'        If ListView1.SelectedItem.SubItems(10) <> "0.00" Then
'            MsgBox "Ya ha compensando horas. Quite la comepnsacion primero", vbExclamation
'            Exit Sub
'        End If
'    Else
'        If ListView1.SelectedItem.SubItems(10) = "0.00" Then
'            MsgBox "Ya ha compensando horas. Quite la compensacion primero", vbExclamation
'            Exit Sub
'        End If
'    End If
'
'
'
'    If Index = 0 Then
'        'Cuando ponemos la baja calculamos si tiene horas en bolsa despues.
'        'las tranformamos en euros de mas en anticpos
'        Imp1 = -1
'        Importe = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(12))
'             Do
'                 SQL = "Introduzca las horas de PLUS para " & ListView1.SelectedItem.SubItems(1) & "." & vbCrLf & "Máximo:" & Format(Importe, "0.00")
'                 SQL = InputBox(SQL, "Horas +")
'                 If SQL <> "" Then
'                     If IsNumeric(SQL) Then
'                         SQL = TransformaPuntosComas(SQL)
'                         Imp1 = CCur(SQL)
'                         If Imp1 > 0 Then
'                            If Imp1 > Importe Then
'                                MsgBox "No puede poner mas horas de las que tiene", vbExclamation
'                                Imp1 = 0
'                            Else
'                                SQL = ""
'                            End If
'                        End If
'                     End If
'                 End If
'             Loop Until SQL = ""
'
'            If SQL = "" And Imp1 <= 0 Then Exit Sub
'
'
'      '  Importe = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(12))
'
'
'            SQL = "SELECT Categorias.Importe1, Categorias.Importe2, Trabajadores.IdTrabajador,PorcSS,PorcIRPF"
'            SQL = SQL & " FROM Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria"
'            SQL = SQL & " WHERE Trabajadores.IdTrabajador=" & ListView1.SelectedItem.Text
'
'            Set RS = New ADODB.Recordset
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            If RS.EOF Then
'                MsgBox "Error leyendo datos trabajador", vbExclamation
'            Else
'                'Le ponemos las horas de plus
'                ListView1.SelectedItem.SubItems(10) = Format(Imp1, FormatoImporte)
'                'En la bolsa le dejo las k tenia menos las k lleva al plus
'                Importe = Importe - Imp1
'                ListView1.SelectedItem.SubItems(12) = Format(Importe, FormatoImporte)
'
'                Importe = Imp1 * RS.Fields(0) 'importe2    horas * importe
'                Imp1 = (Importe * RS!porcSS) + (Importe * RS!porcirpf)
'                Imp1 = Imp1 / 100
'                Importe = Importe + Imp1
'                Importe = Round(Importe, 2)
'
'                'PLUS
'                ListView1.SelectedItem.SubItems(14) = Format(Importe, FormatoImporte)
'
'
'                'Importe origninal
'                Imp1 = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(13))
'                Importe = Importe + Imp1
'
'                'Ponemos las horas de plus
'                ListView1.SelectedItem.SubItems(13) = Format(Importe, FormatoImporte)
'
'                ListView1.SelectedItem.SmallIcon = 4 'Icono de h+
'
'            End If
'            RS.Close
'
'    Else
'
'        PonerBaja False, 0
'    End If
'End Sub

Private Sub Command1_Click()
Dim Rs As ADODB.Recordset
Dim i As Integer
Dim Cad As String


    If ListView1.ListItems.Count < 1 Then Exit Sub


    'Primera comprobacion. O es Lunes o es primero de mes
    If Weekday(CDate(Text1.Text), vbMonday) <> 1 Then
        If Day(CDate(Text1.Text)) <> 1 Then
            SQL = "El dia de inicio ni es Lunes ni es primero de mes."
            MsgBox SQL, vbExclamation
            Exit Sub
        End If
    End If

    'Preguntamos si desea continuar
    SQL = "Seguro que desea cerrar la semana del " & Text1.Text & " al " & Text2.Text & "?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Sub


    'Vemos si ya se han generado las nominas del mes para cada trabajador. Podria ser k no hubiera ninguno
    
    SQL = "#" & Format(Text2.Text, FormatoFecha) & "#"
    SQL = "Select * from JornadasSemanales where Fecha = " & SQL
    SQL = SQL & " AND idTrabajador = "
    Cad = ""
    Set Rs = New ADODB.Recordset
    
    'Recorremos
    For i = 1 To ListView1.ListItems.Count
        Rs.Open SQL & ListView1.ListItems(i).Text, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            Cad = Cad & ListView1.ListItems(i).Text & " - " & ListView1.ListItems(i).SubItems(1) & vbCrLf
            
            'Truco. Voy a utilizar el subitem (15) para marcarlos
            ListView1.ListItems(i).SubItems(15) = 1
        End If
        Rs.Close
    Next i
   
    If Cad <> "" Then
        Cad = "Los siguientes trabajadores ya han cerrado la semana." & vbCrLf & _
            "Si continua no se producirá ningun cambio sobre los mismos y serán borrados de esta lista" & vbCrLf & Cad
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
            
        'Llegado aqui hay k borrar
        For i = ListView1.ListItems.Count To 1 Step -1
            If ListView1.ListItems(i).SubItems(15) = 1 Then
                'Borrar
                ListView1.ListItems.Remove i
            End If
        Next i
        If ListView1.ListItems.Count < 1 Then
            MsgBox "Ningun dato para generar.", vbExclamation
            Exit Sub
        End If
    End If
    
    'pondremos un transaccion
    Screen.MousePointer = vbHourglass
    conn.BeginTrans
    If GenerarCierreSemana Then
        conn.CommitTrans
        MsgBox "Proceso finalizado", vbInformation
    Else
        conn.RollbackTrans
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

    
    Frame1.Visible = False

    
    OfertaFecha
    CalculaFinalSemana True
 
    Command1.Enabled = vUsu.Nivel < 2
    
    EmpresaHoraExtra = False
    SQL = DevuelveDesdeBD("EmpresaHoraExtra", "Empresas", "idEmpresa", 1, "N")
    If SQL <> "" Then
        If CBool(SQL) Then EmpresaHoraExtra = True
    End If
    
    
    ListView1.SmallIcons = Me.ImageList1

    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaColumnas()
'Dim Anch As Single
Dim clmX As ColumnHeader


'ListView1.ColumnHeaders.Clear
'Anch = ListView1.Width - 360
'Anch = Anch / 16
'
'
''Datos Trbajador
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Cod"
'clmX.Width = Anch
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Nombre"
'clmX.Width = Anch * 5
'
'
'
''OFICIALES
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Dias"
'clmX.Width = 510
'
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Horas"
'clmX.Width = Anch
'
''TRABAJADOS
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Dias"
'clmX.Width = 510
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Norm."
'clmX.Width = Anch
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Comp"
'clmX.Width = Anch
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "PLUS"
'clmX.Width = Anch
'
'
''Saldo
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Dias"
'clmX.Width = 510
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Horas"
'clmX.Width = Anch
'
'
''Bolsa
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Ant."
'clmX.Width = Anch
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Post"
'clmX.Width = Anch
'
'
'Set clmX = ListView1.ColumnHeaders.Add()
'clmX.Text = "Anticipos"
'clmX.Width = Anch
'


For Each clmX In ListView1.ColumnHeaders
    
    If clmX.Index > 3 Then clmX.Alignment = lvwColumnRight
Next

'Las lineas
With ListView1

    If EmpresaHoraExtra Then
        .ColumnHeaders(7).Width = 705
        .ColumnHeaders(12).Width = 705
    Else
        .ColumnHeaders(7).Width = 0
        .ColumnHeaders(12).Width = 0
    End If
    Line1.X2 = .ColumnHeaders(3).Left - 30 + 160

    Label3.Left = .ColumnHeaders(3).Left + 160
    Line2.X1 = .ColumnHeaders(3).Left + 30 + 160
    Line2.X2 = .ColumnHeaders(4).Left - 30 + 160
    
    Label4.Left = .ColumnHeaders(4).Left + 160
    Line3.X1 = .ColumnHeaders(4).Left + 30 + 160
    Line3.X2 = .ColumnHeaders(8).Left - 30 + 160
    
    Label5.Left = .ColumnHeaders(8).Left + 160
    Line4.X1 = .ColumnHeaders(8).Left + 30 + 160
    Line4.X2 = .ColumnHeaders(10).Left - 30 + 160
    
    Label6.Left = .ColumnHeaders(10).Left + 160
    Line5.X1 = .ColumnHeaders(10).Left + 30 + 160
    Line5.X2 = .ColumnHeaders(13).Left - 30 + 160
    
    Label7.Left = .ColumnHeaders(13).Left + 160
    Line6.X1 = .ColumnHeaders(13).Left + 30 + 160
    Line6.X2 = .ColumnHeaders(15).Left - 30 + 160
    
    'Pequeño reajuste k borda las lineas
    .ColumnHeaders(3).Width = 1000
    '.ColumnHeaders(14).Width = 1300
    .ColumnHeaders(15).Width = 0
    'La ultima columna a 0
    .ColumnHeaders(16).Width = 0
End With
    


End Sub

Private Sub PonLinea(ByRef i As ListItem, ByRef Rs As ADODB.Recordset)
Dim V As Boolean
'Si tiene dias pendientes

        
        If Rs!ControlNomina = 1 Then
          
            Else
            
        End If

        If Rs!saldodias <> 0 Then
            i.SmallIcon = 2
        Else
            If Rs!DiasTrabajados = 0 Then
                i.SmallIcon = 3
            Else
                If Rs!ControlNomina = 1 Then
                    i.SmallIcon = 1
                Else
                    i.SmallIcon = 6
                End If
            
                
            End If
        End If
        
        i.Text = Rs!Trabajador
        i.SubItems(1) = Rs!Nomtrabajador
        i.ToolTipText = Rs!Nomtrabajador
        
        'Horas oficiles
        i.SubItems(2) = Rs!mesdias & "/" & Format(Rs!meshoras, "0.00")
        
        'Trabajados
        i.SubItems(3) = Rs!DiasTrabajados
        i.SubItems(4) = Format(Rs!horasn, "0.00")
        i.SubItems(5) = Format(Rs!horasc, "0.00")
        i.SubItems(6) = Format(Rs!HorasE, "0.00")
        
        'Saldo
        i.SubItems(7) = Rs!saldodias
        i.SubItems(8) = Format(Rs!saldoh, "0.00")
        
        'ANTES
        'Compensadas en NOMINA
        'I.SubItems(8) = rs!diasperiodo
        'I.SubItems(9) = Format(rs!extras, "0.00")
        
        'AHORA
        
        i.SubItems(9) = Rs!diasperiodo  'Sigue igual
        
       
       
        If Rs!mesdias <> Rs!diasperiodo Then
            If EmpresaHoraExtra Then
                i.SubItems(10) = Format(Rs!extras + Rs!horasn, "0.00")
            Else
                i.SubItems(10) = Format(Rs!horasperiodo, "0.00")
            End If
        Else
            i.SubItems(10) = Format(Rs!horasperiodo, "0.00")
            
        End If
        
        
        'Luego lo utilizare para borrar los
        i.SubItems(11) = Format(Rs!Extrasperiodo, "0.00")
        
        '
        'Bolsa
        i.SubItems(12) = Rs!bolsaantes
        i.SubItems(13) = Format(Rs!bolsadespues, "0.00")
        i.SubItems(14) = Format(Rs!Anticipos, "0.00")

        'PLUS
        i.SubItems(15) = "0.00"
    
        i.SubItems(15) = "0"  'para borrar
        'El tag
        i.Tag = Rs!ControlNomina
End Sub


Private Sub CargaDatos()
Dim i As ListItem
Dim Rs As ADODB.Recordset

    Set Rs = New ADODB.Recordset
    ListView1.ListItems.Clear
    PonSQL ""
    SQL = SQL & " order by "
    If Option1(0).Value Then
        SQL = SQL & "id"
    Else
       SQL = SQL & "nom"
    End If
    SQL = SQL & "Trabajador"
    Rs.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not Rs.EOF
        Set i = ListView1.ListItems.Add

        PonLinea i, Rs
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
        
End Sub




Private Sub Form_Resize()
Dim H As Single
Dim W As Single

    If Me.WindowState = 1 Then Exit Sub
    If Me.Width < 7000 Then
        W = 7000
        Me.Width = W
    Else
        W = Me.Width
    End If
    If Me.Height < 3900 Then
        H = 3900
        Me.Height = H
    Else
        H = Me.Height
    End If
    Me.ListView1.Width = W - ListView1.Left - 210
    Me.ListView1.Height = H - ListView1.Top - 500
    CargaColumnas
End Sub

'Private Sub ListView1_Click()
'Dim i
'    SQL = ""
'    For i = 1 To ListView1.ColumnHeaders.Count
'        SQL = SQL & ListView1.ColumnHeaders(i).Text & ": " & ListView1.ColumnHeaders(i).Width & vbCrLf
'    Next i
'    MsgBox SQL
'End Sub


'Para ahorrar variables
Private Sub OfertaFecha()
Dim Rs As ADODB.Recordset


    
        'Poner fecha
        Text1.Text = ""
        SQL = "Select max(fecha) from JornadasSemanales"
        Set Rs = New ADODB.Recordset
        Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not Rs.EOF Then
            If Not IsNull(Rs.Fields(0)) Then
                Text1.Text = Format(DateAdd("d", 1, Rs.Fields(0)), "dd/mm/yyyy")
            End If
        End If
        Set Rs = Nothing
                
        If Text1.Text = "" Then CalculaFinalSemana False
    
End Sub


Private Sub PonSQL(Id As String)
    SQL = "Select tmpDatosMes.*,nomtrabajador,controlnomina from tmpDatosMes,Trabajadores"
    SQL = SQL & " WHERE tmpDatosMes.trabajador = Trabajadores.idTrabajador "
    If Id <> "" Then SQL = SQL & " AND tmpDatosMes.trabajador =" & Id
End Sub




Private Function GenerarCierreSemana() As Boolean
Dim i As Integer
Dim Cad As String
Dim Importe As Currency
Dim J As Integer
Dim Dias As Integer
Dim Horas As Currency

On Error GoTo EGenerarNominas
    GenerarCierreSemana = False

    SQL = "INSERT INTO JornadasSemanales(Fecha,IdTrabajador,DiasOfi,HorasOfi,Dias,HN,HC,BolsaAntes,BolsaDespues,HE) VALUES (#"
    SQL = SQL & Format(Text2.Text, FormatoFecha) & "#,"

    'Primero generamos la tabla de  nominas con los importes marcados aqui
    For i = 1 To ListView1.ListItems.Count
        Cad = ListView1.ListItems(i).Text & ","
        
        'Dias / Horas oficiales
        J = InStr(1, ListView1.ListItems(i).SubItems(2), "/")
        Dias = Val(Mid(ListView1.ListItems(i).SubItems(2), 1, J - 1))
        Cad = Cad & Dias & ","
        Horas = CCur(Mid(ListView1.ListItems(i).SubItems(2), J + 1))
        Cad = Cad & TransformaComasPuntos(CStr(Horas)) & ","
        
        'Dias trabajadps
        Cad = Cad & TransformaComasPuntos(ListView1.ListItems(i).SubItems(9)) & ","
        
        'Hnormales y Hcompensadas
        Cad = Cad & TransformaComasPuntos(ListView1.ListItems(i).SubItems(10)) & ",0,"

     
        'Bolsa Antes. Para el informe conjunto a la nomina
        Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(12)))) & ","
        Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(13)))) & ","
        
        If EmpresaHoraExtra Then
            Cad = Cad & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(11))))
        Else
            Cad = Cad & "0"
        End If
        
        Cad = Cad & ")"
        Cad = SQL & Cad
        
        J = 1
        If Dias = 0 And Horas = 0 Then
            'Las oficilaes son 0. Veremos las trabajadas
            If Val(TransformaComasPuntos(ListView1.ListItems(i).SubItems(8))) = 0 Then
                If CCur(ListView1.ListItems(i).SubItems(9)) = 0 Then J = 0
            End If
        End If
        If J = 1 Then
            conn.Execute Cad

          
            
            'Pondremos la bolsa de horas Y, hay bajas,
             'entonces actualizaremos la baja de cada trabajador
             'al ultimo dia trabajado
             Cad = "UPDATE Trabajadores SET Bolsahoras = " & TransformaComasPuntos(ListView1.ListItems(i).SubItems(13))
             Cad = Cad & " WHERE idTrabajador = " & ListView1.ListItems(i).Text
             conn.Execute Cad

        End If
'
'        'Si se da de baja le pongo fecha de baja
'        If ListView1.ListItems(I).SmallIcon = 3 Then
'            'SE DA DE BAJA
'            cad = DevuelveDesdeBD("fecbaja", "trabajadores", "IdTrabajador", ListView1.ListItems(I).Text, "N")
'            If cad = "" Then
'                'NO TIENE FECHA BAJA
'                cad = DiasMes(Combo1.ListIndex + 1, Int(Text1.Text)) & "/" & CStr(Combo1.ListIndex + 1) & "/" & Text1.Text
'                cad = Format(cad, FormatoFecha)
'                cad = "UPDATE Trabajadores SET fecbaja = #" & cad & "#"
'                cad = cad & " WHERE idTrabajador = " & ListView1.ListItems(I).Text
'                Conn.Execute cad
'            End If
'        End If
'
        
    Next i
    

    
    
    
    
    GenerarCierreSemana = True
    Exit Function
EGenerarNominas:
    MuestraError Err.Number, Err.Description
End Function





Private Function PuedeCompensarDias() As Integer
Dim i As Integer

    PuedeCompensarDias = 0
    SQL = DevuelveDesdeBD("idHorario", "Trabajadores", "idTrabajador", ListView1.SelectedItem.Text, "N")
    i = Val(SQL)
    
    'En la tabla tmpHorasMesHorario, al cargar los datos
    'se han cargado las horas oficiales
    SQL = DevuelveDesdeBD("Dias", "tmpHorasMesHorario", "idHorario", CStr(i), "N")
    If SQL <> "" Then
        i = Val(SQL)
        i = i - Val(ListView1.SelectedItem.SubItems(8))
        If i > 0 Then PuedeCompensarDias = i
    End If
    
    
    
End Function


'Private Sub CompensarDias(Dias As Integer)
'Dim I As Integer
'Dim Lab As Integer
'Dim H As Currency
'Dim H1 As Currency
'Dim D1 As Integer
'Dim RS As ADODB.Recordset
'
'
'    SQL = DevuelveDesdeBD("idHorario", "Trabajadores", "idTrabajador", ListView1.SelectedItem.Text, "N")
'    I = Val(SQL)
'
'    Lab = DiasLaborablesSemana(I)
'    If Lab < 1 Then Exit Sub
'
'    If Dias < Lab Then
'        'Nos salimos pq no tengo bastantes dias para compensar un semana
'        Exit Sub
'    End If
'
'
'
'    'QUiero saber las horas a la semana k puedo compensar
'    Set RS = New ADODB.Recordset
'    SQL = "Select * from Horarios Where idHorario =" & I
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'
'        H = CCur(ListView1.SelectedItem.SubItems(11)) 'Las horas k le van a quedar en bolsa
'        'Ya tengo el horario y los dias a compensar
'
'        'A compensar
'        D1 = 0 'Dias
'        H1 = 0 'Horas
'
'        'Por lo tanto veo cuantas semanas mas voy a compensar
'        Do
'            I = Dias \ Lab
'            If I > 0 Then
'                'Una semana seguro k puedo compensar. Vamos palla
'                If H >= RS!TotalHoras Then   'Horas semana
'                    D1 = D1 + Lab
'                    H1 = H1 + RS!TotalHoras
'                    H = H - RS!TotalHoras
'                End If
'                Dias = Dias - Lab
'            End If
'        Loop Until I = 0
'    End If
'    RS.Close
'
'
'    'Si a compensado lo reflejo en la listview
'    If D1 > 0 Then
'        'Dias nomina
'
'
'        'Horas para la nomina
'        SQL = "Select Importe1,importe2,porcSS,porcIRPF from Categorias,Trabajadores WHERE Trabajadores.IdCategoria = Categorias.IdCategoria"
'        SQL = SQL & " AND Trabajadores.idTrabajador =" & ListView1.SelectedItem.Text
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        If Not RS.EOF Then
'            'Ponemos ya las nuevas horas en horas normales
'            H = CCur(ListView1.SelectedItem.SubItems(4)) + H1
'            ListView1.SelectedItem.SubItems(4) = Format(H, FormatoImporte)
'            'Bolsa
'            H = CCur(ListView1.SelectedItem.SubItems(11)) - H1
'            ListView1.SelectedItem.SubItems(11) = Format(H, FormatoImporte)
'
'            'Precio bruto
'            H = H1 * RS!Importe1
'
'            'Precio neto
'            H1 = ((H * RS!porcSS) + (H * RS!porcirpf)) / 100
'
'            H = Round(H - H1, 2)
'            'Anticipos
'            H1 = ImporteFormateadoAmoneda(ListView1.SelectedItem.SubItems(12))
'            H = H + H1
'            ListView1.SelectedItem.SubItems(12) = H
'
'            'Dias nomina
'            I = Val(ListView1.SelectedItem.SubItems(8)) + D1
'            ListView1.SelectedItem.SubItems(8) = I
'        End If
'    End If
'    Set RS = Nothing
'End Sub



Private Sub frmC_Selec(vFecha As Date)
    Text1.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1.Text <> "" Then
        If IsDate(Text1.Text) Then frmC.Fecha = CDate(Text1.Text)
    End If
    frmC.Show vbModal
    Set frmC = Nothing
    CalculaFinalSemana True
End Sub

Private Sub ListView1_DblClick()
Dim vH As CHorarios
Dim F As Date
Dim MEDIOS As String
Dim i As Integer

    With ListView1.SelectedItem
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        SQL = DevuelveDesdeBD("idHorario", "Trabajadores", "idTrabajador", .Text, "N")
        If SQL = "" Then
            MsgBox "Error leyendo datos trabajador", vbExclamation
        Else
                
            
            Set vH = New CHorarios
            F = CDate(Text2.Text)
            MEDIOS = vH.LeerMediosDias(CInt(SQL), CDate(Text1.Text), F)
            
            SQL = vH.LeerDiasFestivos(CInt(SQL), CDate(Text1.Text), F)
            i = DateDiff("d", CDate(Text1.Text), F)
            
            frmVerDiasMesTrabajador3.DiasEnNomina = Val(.SubItems(12))
            frmVerDiasMesTrabajador3.TodoElMEs = i
            frmVerDiasMesTrabajador3.JornadasSemanales = False
            frmVerDiasMesTrabajador3.MediosDias = MEDIOS
            frmVerDiasMesTrabajador3.FESTIVOS = SQL
            frmVerDiasMesTrabajador3.Trabajador = .SubItems(1) & "|" & .Text & "|"
            frmVerDiasMesTrabajador3.FechaIni = Text1.Text
            frmVerDiasMesTrabajador3.Show vbModal
            
            Set vH = Nothing
        End If
    End With
End Sub

Private Sub Text1_GotFocus()
    With Text1
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text1_LostFocus()
    Text1.Text = Trim(Text1.Text)
    If Text1.Text <> "" Then
        If Not EsFechaOK(Text1) Then
            MsgBox "Fecha incorrecta. (" & Text1.Text & ")", vbExclamation
            Text1.SetFocus
        Else
            'Ponemos la siguiente fecha
            CalculaFinalSemana True
        End If
    End If
End Sub


Private Sub CalculaFinalSemana(Final As Boolean)
Dim F As Date
Dim F2 As Date
Dim i As Integer


If Final Then
    F = CDate(Text1.Text)
    i = Weekday(F, vbMonday)
    'Ya tengo k dia de la semana es
    i = 7 - i
    F2 = DateAdd("d", i, F)
    If Month(F2) <> Month(F) Then
        'Ha sumado y se pasa del mes
        i = DiasMes(Month(F), Year(F))
        F2 = CDate(i & "/" & Format(F, "mm/yyyy"))
    End If
    Text2.Text = Format(F2, "dd/mm/yyyy")
Else
    F = Now
    i = Weekday(F, vbMonday)
    If i <> 1 Then
        'NO ES LUNES
        If Day(F) < 6 Then
            F2 = CDate("01/" & Month(F) & "/" & Year(F))
        Else
            i = i - 1
            F2 = DateAdd("d", -i, F)
        End If
        F = F2
    End If
    Text1.Text = Format(F, "dd/mm/yyyy")
End If
    
End Sub



