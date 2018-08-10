VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmGeneraAnti 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anticipos"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11550
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   11550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneraAnti.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneraAnti.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmGeneraAnti.frx":0B34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame FrameMes 
      Height          =   855
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11475
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   6480
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   360
         Width           =   2415
      End
      Begin VB.CommandButton cmdListado 
         Caption         =   "Imprimir"
         Height          =   375
         Left            =   4560
         TabIndex        =   14
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "Recibos"
         Height          =   375
         Left            =   5520
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nombre"
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   11
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Codigo"
         Height          =   255
         Index           =   0
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generar"
         Height          =   375
         Left            =   10320
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   9000
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   420
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   0
         Left            =   240
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   420
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   1380
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   420
         Width           =   1095
      End
      Begin VB.CommandButton cmdGenHoras 
         Caption         =   "Calcular horas"
         Height          =   615
         Left            =   2520
         TabIndex        =   3
         Top             =   120
         Width           =   855
      End
      Begin VB.Image ImgFech 
         Height          =   240
         Index           =   2
         Left            =   9960
         Picture         =   "frmGeneraAnti.frx":320E
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Fec. anticipo"
         Height          =   195
         Index           =   2
         Left            =   9000
         TabIndex        =   9
         Top             =   180
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Desde"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   180
         Width           =   555
      End
      Begin VB.Image ImgFech 
         Height          =   240
         Index           =   0
         Left            =   900
         Picture         =   "frmGeneraAnti.frx":3310
         Top             =   180
         Width           =   240
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   1380
         TabIndex        =   6
         Top             =   180
         Width           =   420
      End
      Begin VB.Image ImgFech 
         Height          =   240
         Index           =   1
         Left            =   1860
         Picture         =   "frmGeneraAnti.frx":3412
         Top             =   180
         Width           =   240
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   6915
      Left            =   60
      TabIndex        =   8
      Top             =   900
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   12197
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   5821
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "HN"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Imp/h"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "HC"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "Imp/h"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "%SS"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "% IRPF"
         Object.Width           =   1191
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "Total"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Pagos"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "A ingresar"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   11
         Text            =   "T1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   12
         Text            =   "T2"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   13
         Text            =   "Bruto"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   6915
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   11475
      _ExtentX        =   20241
      _ExtentY        =   12197
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   14
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cod"
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Nombre"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "HN"
         Object.Width           =   1206
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Importe"
         Object.Width           =   1455
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Antig."
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   5
         Text            =   "SS"
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   6
         Text            =   "Ret."
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "TOTAL N"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "HC"
         Object.Width           =   1206
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   9
         Text            =   "Importe"
         Object.Width           =   1455
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   10
         Text            =   "Antig."
         Object.Width           =   1244
      EndProperty
      BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   11
         Text            =   "SS"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   12
         Text            =   "Ret"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   13
         Text            =   "TOTAL C"
         Object.Width           =   1587
      EndProperty
   End
End
Attribute VB_Name = "frmGeneraAnti"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Antiguedad As Byte
    '0 .-  Ninguno, ni horas N, ni horas C
    '1 .- Antiguedad sobre horas N
    '2 ,. Antiguedad sobre holras C
    '3 .- Antiguedad sobre las dos


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Dim RT As ADODB.Recordset
Dim SQLAUX As String
Dim PrimeraVez As Boolean

Private Sub cmdGenHoras_Click()
Dim Seccion As Integer

    If Text1(0).Text = "" Or Text1(1).Text = "" Then
        MsgBox "Escriba las fechas de inicio y fin", vbExclamation
        Exit Sub
    End If
    If ExistenIncorrectos Then Exit Sub
    Screen.MousePointer = vbHourglass
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
    
    Seccion = Me.Combo1.ItemData(Me.Combo1.ListIndex)
    
    CalculaHorasTrabajadas CDate(Text1(0).Text), CDate(Text1(1).Text), 3, Seccion
    
    'Si hay trabajadores de controlnomina=2  LIQUIDA COMEPNSABLES
    GenerarLiquidacionCompensables CDate(Text1(0).Text), CDate(Text1(1).Text)
    
    CargaDatos
    Screen.MousePointer = vbDefault
End Sub
Private Sub CargaDatos()
    If Antiguedad = 0 Then
        CargaDatosModo1
    Else
        CargaDatosAntiguedad
    End If
End Sub
Private Function ExistenIncorrectos() As Boolean
Dim FI As Date
Dim FF As Date

    ExistenIncorrectos = True
    If ComprobarMarcajesCorrectos(CDate(Text1(0).Text), CDate(Text1(1).Text), False) <> 0 Then
        SQLAUX = "Existen marcajes incorrectos entre las fechas. ¿Desea continuar?"
        If MsgBox(SQLAUX, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
    End If
    ExistenIncorrectos = False
End Function



Private Sub ImprimirNormal()
Dim i As Integer
Dim Importe As Currency
Dim Importe2 As Currency
Dim SS As Currency
Dim IRPF As Currency
Dim SQ As String
Dim inserttmpdatosmes As String

    If ListView1.ListItems.Count = 0 And ListView2.ListItems.Count = 0 Then Exit Sub
    
    
    Screen.MousePointer = vbHourglass
    
    SQLAUX = "DELETE FROM tmpPagosMes"
    conn.Execute SQLAUX
    
    
 
    If Antiguedad = 0 Then
        'modificado el 18 octubre 2004
        ' revisado el 29 de octubre
        'precio hora 1
        SQLAUX = "INSERT INTO tmpPagosMes(idTrabajador,nombre,HT,importe1,HC,Importe2,"
        SQLAUX = SQLAUX & "IRPF,SS,NETO,Pagos,INGRESAR,BRUTO,preciohora1) VALUES ("
    
    
        'SIN ANTIGUEDAD
            For i = 1 To ListView1.ListItems.Count
                With ListView1.ListItems(i)
                    SQ = .Text
                    SQ = SQ & ",""" & .SubItems(1) & ""","
                    'Horas T
                    Importe = CCur(.SubItems(11))
                    SQ = SQ & """" & .SubItems(2) & """," & TransformaComasPuntos(CStr(Importe)) & ","
                    
                    
                    'HORAS N
                    Importe = CCur(.SubItems(12))
                    SQ = SQ & """" & .SubItems(4) & """," & TransformaComasPuntos(CStr(Importe)) & ","
                    
                    
                    'IRPF y SS
                    SQ = SQ & "'" & .SubItems(7) & "','" & .SubItems(6) & "',"
                    
                    'Salario neot y pagos realizados
                    Importe = CCur(.SubItems(8))
                    SQ = SQ & TransformaComasPuntos(CStr(Importe)) & ","
                    
                    Importe = CCur(.SubItems(9))
                    SQ = SQ & TransformaComasPuntos(CStr(Importe)) & ","
                    
                    'Ingresar
                    Importe = CCur(.SubItems(10))
                    SQ = SQ & TransformaComasPuntos(CStr(Importe))
                    
                    'BRUTO
                    Importe = CCur(.SubItems(13))
                    SQ = SQ & "," & TransformaComasPuntos(CStr(Importe))
                    
                    
                    'Precio hora 1
                    Importe = CCur(.SubItems(3))
                    SQ = SQ & "," & TransformaComasPuntos(CStr(Importe))
                End With
                
                'Ejecutamos
                SQ = SQLAUX & SQ & ")"
                conn.Execute SQ
            Next i
    
    Else
        'CON ANTIGUEDAD
             'modificado el 18 octubre 2004
            'precio hora 1
            SQLAUX = "INSERT INTO tmpPagosMes(idTrabajador,nombre,HT,importe1,HC,Importe2,"
            SQLAUX = SQLAUX & "SS,IRPF,NETO,irpf2,ss2,Pagos,INGRESAR,BRUTO,importe3,PrecioHora1) VALUES ("
        
          For i = 1 To ListView2.ListItems.Count
                With ListView2.ListItems(i)
                    SQ = .Text
                    SQ = SQ & ",""" & .SubItems(1) & ""","
                    'Horas T
                    Importe = CCur(.SubItems(3))
                    Importe2 = CCur(.SubItems(2))
                    SQ = SQ & TransformaComasPuntos(CStr(Importe2)) & "," & TransformaComasPuntos(CStr(Importe)) & ","
                    Importe2 = Importe
                    
                    
                    'le sumo la antiguedad
                    Importe = CCur(.SubItems(4))
                    Importe2 = Importe2 + Importe
                    
                    
                    
                    'HORAS c
                    Importe = CCur(.SubItems(9))
                    SQ = SQ & "" & TransformaComasPuntos(.SubItems(8)) & "," & TransformaComasPuntos(CStr(Importe)) & ","
                    
                    
                    'esta no hace falta, tengo un importe separado
                    'Importe2 = Importe2 + Importe
                    
                    'Imprimimos el total
                    'SQ = SQ & TransformaComasPuntos(CStr(Importe2)) & ","
                    
                    'IRPF y SS
                    SS = CCur(.SubItems(5))
                    IRPF = CCur(.SubItems(6))
                    SQ = SQ & "'"
                    'SQ = SQ & TransformaComasPuntos(Format(SS, "0.00")) & ","
                    SQ = SQ & TransformaComasPuntos(.SubItems(5)) & "','"
                    'SQ = SQ & TransformaComasPuntos(Format(IRPF, "0.00")) & ","
                    SQ = SQ & TransformaComasPuntos(.SubItems(6)) & "',"
                    'Salario neto es el importe1
                    Importe2 = Importe2 - IRPF - SS
                    SQ = SQ & TransformaComasPuntos(CStr(Importe2)) & ","
                    
                    
                    
                    
                    
                    
                    'Las de las C
                    SS = CCur(.SubItems(11))
                    SQ = SQ & TransformaComasPuntos(Format(SS, "0.00")) & ","
                    IRPF = CCur(.SubItems(12))
                    SQ = SQ & TransformaComasPuntos(Format(IRPF, "0.00")) & ","
                    
                    
                    
                    'pagos realizados
                    SQ = SQ & "0,"
                    
                    'El a ingresar
                    Importe = Importe - 0 - IRPF - SS 'Pagos
                    SQ = SQ & TransformaComasPuntos(CStr(Importe)) & ","
                    'Bruto
                    SQ = SQ & TransformaComasPuntos(CStr(Importe2)) & ","
                
        
                    'Antiguedad
                    Importe = CCur(.SubItems(4))
                    SQ = SQ & TransformaComasPuntos(CStr(Importe)) & ","
                    
                    'Dias trabajados
                    SQ = SQ & .Tag
                    
                End With
                
                'Ejecutamos
                SQ = SQLAUX & SQ & ")"
                conn.Execute SQ
            Next i
    
    
    
    
    End If
    
    'Cadena texto
    SQ = "TEXTO= ""Desde " & Text1(0).Text & "    hasta " & Text1(1).Text
    SQ = SQ & "          ANTICIPO: " & Text1(2).Text & """|"
    If Antiguedad = 0 Then
        i = 1
    Else
        i = 7
    End If
    frmImprimir.Opcion = i
    frmImprimir.NumeroParametros = 1
    frmImprimir.OtrosParametros = SQ
    frmImprimir.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdImprimir_Click()
    If Antiguedad = 0 Then
        MsgBox "Opción no disponible"
    Else
        'Impresion nominas
        ImprimirAnticiposAlzira
    End If
End Sub

Private Sub cmdListado_Click()
    
        ImprimirNormal

End Sub

Private Sub Command1_Click()

    If Text1(2).Text = "" Then
        MsgBox "Ponga la fecha del anticipo", vbExclamation
        Exit Sub
    End If
    If Antiguedad = 0 Then
        If ListView1.ListItems.Count = 0 Then Exit Sub
    Else
        If ListView2.ListItems.Count = 0 Then Exit Sub
    End If
    
    
    
    SQLAUX = "Seguro que desea generar los anticipos con estos valores?"
    If MsgBox(SQLAUX, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    
    'Entre las fechas
    If YaExistenAnticiposFecha(CDate(Text1(2).Text)) Then
        SQLAUX = "Ya existen anticipos generados con esta fecha. " & vbCrLf
        MsgBox SQLAUX, vbQuestion
        Exit Sub
    End If
    
    
    
    
    conn.BeginTrans
    If CrearAnticipos Then
        'Imprimiremos el listado
        conn.CommitTrans
        MsgBox "Proceso finalizdo con exito", vbInformation
    Else
        conn.RollbackTrans
        MsgBox "Se han producido errores", vbExclamation
    End If
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        ListView1.Visible = (Antiguedad = 0)
        ListView2.Visible = Not ListView1.Visible
    End If
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    If Day(Now) < 16 Then
        Text1(0).Text = Format(CDate("01/" & Month(Now) & "/" & Year(Now)), "dd/mm/yyyy")
        SQLAUX = Format(CDate("15/" & Month(Now) & "/" & Year(Now)), "dd/mm/yyyy")
    Else
        Text1(0).Text = Format(CDate("16/" & Month(Now) & "/" & Year(Now)), "dd/mm/yyyy")
        SQLAUX = CStr(DiasMes(Month(Now), Year(Now))) & "/" & Month(Now) & "/" & Year(Now)
    End If
    Text1(1).Text = SQLAUX
    Text1(2).Text = SQLAUX
    SQLAUX = ""
    ListView2.ListItems.Clear
    ListView1.ListItems.Clear
    CargaComboSecciones Me.Combo1, True

End Sub



Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(frmC.Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub ImgFech_Click(Index As Integer)
    Set frmC = New frmCal
    frmC.Fecha = Now
    If Text1(Index).Text <> "" Then frmC.Fecha = CDate(Text1(Index).Text)
    frmC.Tag = Index
    frmC.Show vbModal
    Set frmC = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text <> "" Then
        If Not EsFechaOK(Text1(Index)) Then
            Text1(Index).SetFocus
        Else
            If Index = 1 Then Text1(2).Text = Text1(1).Text
        End If
    End If
    
End Sub


Private Sub CargaDatosModo1()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Importe As Currency
Dim Importe2 As Currency
Dim itmX As ListItem

    ListView1.ListItems.Clear
    ListView1.ColumnHeaders(12).Width = 0
    ListView1.ColumnHeaders(13).Width = 0
    ListView1.ColumnHeaders(14).Width = 0
    SQL = "SELECT Trabajadores.IdTrabajador, Trabajadores.NomTrabajador, "
    SQL = SQL & " tmpHoras.HorasT, Categorias.Importe1,[HorasT]*[Importe1] AS T1,"
    SQL = SQL & " tmpHoras.HorasC, Categorias.Importe2,[HorasC]*[Importe2] AS T2 "
    'SQL = SQL & " tmpHoras.HorasE, Categorias.Importe3,[HorasE]*[Importe2] AS T3 "
    SQL = SQL & " ,Trabajadores.PorcSS, Trabajadores.PorcIRPF, Trabajadores.ControlNomina,Trabajadores.Embargado  "
    SQL = SQL & " FROM (Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria) INNER JOIN tmpHoras ON Trabajadores.IdTrabajador = tmpHoras.trabajador"
    SQL = SQL & " ORDER BY "
    If Option1(0).Value Then
        SQL = SQL & "idTrabajador "
    Else
        SQL = SQL & "nomtrabajador"
    End If
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set itmX = ListView1.ListItems.Add(, , RS.Fields(0))
        itmX.SubItems(1) = RS!Nomtrabajador
        itmX.SubItems(2) = Format(RS!horast, FormatoImporte)
        Importe = Round(RS!Importe1 * RS!horast, 2)
        itmX.SubItems(3) = Format(RS!Importe1, FormatoImporte)

        'itmX.SubItems(3) = Format(Importe, FormatoImporte)
        
        itmX.SubItems(4) = Format(RS!horasc, FormatoImporte)
        
        
        If RS!ControlNomina = 2 Then
            Importe = 0

        Else
            Importe = RS!Importe2
        End If
        itmX.SubItems(5) = Format(Importe, FormatoImporte)

        itmX.SubItems(6) = Format(RS!porcSS, FormatoImporte)
        itmX.SubItems(7) = Format(RS!porcirpf, FormatoImporte)
        

        Importe = Round(RS!Importe1 * RS!horast, 2)
        itmX.SubItems(3) = Format(RS!Importe1, FormatoImporte)

        Importe = RS!T1
        itmX.SubItems(11) = Importe
        
        If RS!ControlNomina = 2 Then
            Importe2 = 0
        Else
            Importe2 = RS!T2
        End If
        itmX.SubItems(12) = Importe2
            
        Importe = Importe + Importe2
        Importe = Round(Importe, 2)
        'Bruto
        itmX.SubItems(13) = Importe
        
        'Iconito en funcion del tipo de control de nominas
        If RS!ControlNomina = 2 Then
            itmX.SmallIcon = 2
        Else
            itmX.SmallIcon = 1
        End If
        
        
        Importe2 = RS!porcSS + RS!porcirpf
        Importe2 = (Importe2 * Importe) / 100
        Importe2 = Round(Importe2, 2)
        Importe = Importe - Importe2  'BRUTO - IRPF - SS
        itmX.SubItems(8) = Format(Importe, FormatoImporte)   'TOTAL
        'Obtner pagos efectuados en el periodo
        Importe2 = ObtenerPagosPeriodo(RS.Fields(0))
        itmX.SubItems(9) = Format(Importe2, FormatoImporte)
        'TOTAL A INGRESAR
        Importe2 = Importe - Importe2
        
        
        If RS!embargado = 1 Then
            'Esta embargado, NO le pagamos nada de nada
            itmX.SubItems(10) = 0
            itmX.SmallIcon = 3
        Else
            itmX.SubItems(10) = Format(Importe2, FormatoImporte)
        End If
            
                    
            
            
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Function CrearAnticipos() As Boolean
Dim i As Integer
Dim Aux As String
Dim SQL As String
Dim PagosSeparados As Boolean
Dim Importe As Currency
Dim Aux2 As String
Dim Importe2 As Currency
    On Error GoTo ECrearAnticipos
    CrearAnticipos = False
    
    
    'Vemos si tiene puesto la marca de pagos separados en parametros
    SQL = DevuelveDesdeBD("abonosSeparados", "Empresas", "idEmpresa", MiEmpresa.IdEmpresa, "N")
    i = -1
    If SQL <> "" Then
        i = Val(SQL)
    End If
    
    If i < 0 Then
        'No esta definidao pagos separados
        Exit Function
    End If
    
    PagosSeparados = (i = 1)
    
    'Para ver la secuencia
    Set RT = New ADODB.Recordset
    
    'Para la insercion
    SQL = "INSERT INTO Pagos (Fecha,Observaciones,Pagado,Trabajador,Importe,Tipo) VALUES ("
    SQL = SQL & "#" & Format(Text1(2).Text, FormatoFecha)
    SQL = SQL & "#,'Anticipo: " & Format(Text1(0).Text, "dd/mm/yy") & " - " & Format(Text1(1).Text, "dd/mm/yy") & "',FALSE,"
    
    'Generaremos los anticipos
    If Antiguedad = 0 Then
        For i = 1 To ListView1.ListItems.Count
            Aux = SQL & ListView1.ListItems(i).Text & ","
            Aux = Aux & TransformaComasPuntos(CStr(ImporteFormateadoAmoneda(ListView1.ListItems(i).SubItems(10)))) & ",1)"  '1 de anticipo
            
            
            'Si es cero NO grabamos nada. Probablemente es que ha sido embargado
            If ListView1.ListItems(i).SubItems(10) = "0" Then
                'Embargado, casi seguro
                'no insertamos
            Else
                conn.Execute Aux
            End If
        Next i
    Else
        'Antiguedad
        'Si son abonos separados
        For i = 1 To ListView2.ListItems.Count
            'El trabajador es el mismo
            Aux = SQL & ListView2.ListItems(i).Text & ","
            Importe = ImporteFormateadoAmoneda(ListView2.ListItems(i).SubItems(7))
            Importe2 = ImporteFormateadoAmoneda(ListView2.ListItems(i).SubItems(13))
            
            If PagosSeparados Then
                'Primero pagamos el anticipo nomina
                If Importe > 0 Then
                    Aux2 = Aux & TransformaComasPuntos(CStr(Importe)) & ",1)"  '1 anticpo normal
                    conn.Execute Aux2
                End If
                'El anticpo de las horas extra
                If Importe2 > 0 Then
                    Aux2 = Aux & TransformaComasPuntos(CStr(Importe2)) & ",3)" '3 anticpo extras
                    conn.Execute Aux2
                End If
            
            Else
                'Pago unico
                Importe = Importe + Importe2
                If Importe > 0 Then
                    Aux2 = Aux & TransformaComasPuntos(CStr(Importe)) & ",1)" '1 anticpo normal
                    conn.Execute Aux2
                End If

        
            End If
        Next i
    End If
    
    CrearAnticipos = True
    Exit Function
ECrearAnticipos:

    MuestraError Err.Number, Err.Description & vbCrLf
    
End Function


'Private Function ObtenerSecuencia(Trab As Long) As Integer
'    RT.Open SQLAUX & Trab, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    ObtenerSecuencia = 0
'    If Not RT.EOF Then
'        If Not IsNull(RT.Fields(0)) Then
'            ObtenerSecuencia = RT.Fields(0)
'        End If
'    End If
'    RT.Close
'    ObtenerSecuencia = ObtenerSecuencia + 1
'End Function
'

Private Function YaExistenAnticiposFecha(F1 As Date) As Boolean
    SQLAUX = "Select * from pagos where fecha =#" & Format(F1, FormatoFecha) & "#"
    SQLAUX = SQLAUX & " AND (tipo = 1 or tipo =4)"  'ANTICIPO
    Set RT = New ADODB.Recordset
    RT.Open SQLAUX, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    YaExistenAnticiposFecha = False
    If Not RT.EOF Then YaExistenAnticiposFecha = True
    RT.Close
    Set RT = Nothing
End Function



Private Function ObtenerPagosPeriodo(Traba As Long) As Currency
Dim Impor As Currency
    SQLAUX = "Select importe from pagos where fecha >=#" & Format(Text1(0).Text, FormatoFecha) & "#"
    SQLAUX = SQLAUX & " AND  fecha <=#" & Format(Text1(1).Text, FormatoFecha) & "#"
    SQLAUX = SQLAUX & " AND tipo = 0"  'Pagos adelantados al trabajador
    SQLAUX = SQLAUX & " AND Trabajador = " & Traba
    Set RT = New ADODB.Recordset
    RT.Open SQLAUX, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Impor = 0
    If Not RT.EOF Then
        While Not RT.EOF
            Impor = Impor + RT.Fields(0)
            RT.MoveNext
        Wend
    End If
    RT.Close
    Set RT = Nothing
    ObtenerPagosPeriodo = Impor
End Function



Private Sub CargaDatosAntiguedad()
Dim SQL As String
Dim RS As ADODB.Recordset
Dim Importe As Currency
Dim Importe2 As Currency
Dim Importe3 As Currency
Dim itmX As ListItem



    'Si no tiene antiguedad en alguhna de sus modalidades oculto la columna
    If Antiguedad <> 3 Then
        If Antiguedad = 1 Then
            'OCoulto antiguedad en COMEPNSABLES
            ListView2.ColumnHeaders(11).Width = 0
        Else
            'raro raro raro
            ListView2.ColumnHeaders(5).Width = 0
        End If
    End If
'    ListView1.ListItems.Clear
'    ListView1.ColumnHeaders(12).Width = 0
'    ListView1.ColumnHeaders(13).Width = 0
'    ListView1.ColumnHeaders(14).Width = 0
    SQL = "SELECT Trabajadores.IdTrabajador, Trabajadores.NomTrabajador, "
    SQL = SQL & " tmpHoras.HorasT, Categorias.Importe1,[HorasT]*[Importe1] AS T1,"
    SQL = SQL & " tmpHoras.HorasC, Categorias.Importe2,[HorasC]*[Importe2] AS T2,Dias "
    'SQL = SQL & " tmpHoras.HorasE, Categorias.Importe3,[HorasE]*[Importe2] AS T3 "
    SQL = SQL & " ,Trabajadores.PorcSS, Trabajadores.PorcIRPF,Trabajadores.PorcAntiguedad,controlnomina"
    SQL = SQL & " FROM (Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria) INNER JOIN tmpHoras ON Trabajadores.IdTrabajador = tmpHoras.trabajador"
    SQL = SQL & " ORDER BY "
    If Option1(0).Value Then
        SQL = SQL & "idTrabajador "
    Else
        SQL = SQL & "nomtrabajador"
    End If
    Set RS = New ADODB.Recordset
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'List 2
        Set itmX = ListView2.ListItems.Add(, , RS.Fields(0))
        itmX.SubItems(1) = RS!Nomtrabajador
        itmX.SubItems(2) = Format(RS!horast, FormatoImporte)
        'Importe 1
        
        Importe = Round(RS!Importe1 * RS!horast, 2)
        itmX.SubItems(3) = Format(Importe, FormatoImporte)
        
        'Antiguedad
        If Antiguedad = 1 Or Antiguedad = 3 Then
            Importe2 = Round((Importe * RS!porcantiguedad) / 100, 2)
        Else
            Importe2 = 0
        End If
        itmX.SubItems(4) = Format(Importe2, FormatoImporte)
        Importe = Importe + Importe2
        
        'IRPF y RET sobre la misma BASE, importe2
        
        Importe2 = Round((Importe * RS!porcirpf) / 100, 2)
        itmX.SubItems(5) = Format(Importe2, FormatoImporte)
        
        Importe3 = Round((Importe * RS!porcSS) / 100, 2)
        itmX.SubItems(6) = Format(Importe3, FormatoImporte)
        
        Importe3 = Importe3 + Importe2
        Importe = Importe - Importe3
        itmX.SubItems(7) = Format(Importe, FormatoImporte)
        
        
        'HORAS COMPENSABLES
        itmX.SubItems(8) = Format(RS!horasc, FormatoImporte)
        If RS!ControlNomina = 2 Then
            Importe = 0
        Else
            Importe = RS!Importe2
        End If
        Importe = Round(Importe * RS!horasc, 2)
        itmX.SubItems(9) = Format(Importe, FormatoImporte)
        'Antiguedad
        If Antiguedad >= 2 Then
            Importe2 = Round((Importe * RS!porcantiguedad) / 100, 2)
        Else
            Importe2 = 0
        End If
        itmX.SubItems(10) = Format(Importe2, FormatoImporte)
        Importe = Importe + Importe2
        
        'IRPF y RET sobre la misma BASE, importe2
        
        Importe2 = Round((Importe * RS!porcirpf) / 100, 2)
        itmX.SubItems(11) = Format(Importe2, FormatoImporte)
        
        Importe3 = Round((Importe * RS!porcSS) / 100, 2)
        itmX.SubItems(12) = Format(Importe3, FormatoImporte)
        
        Importe3 = Importe3 + Importe2
        Importe = Importe - Importe3
        itmX.SubItems(13) = Format(Importe, FormatoImporte)
        
            
        'Ciconito en funcion del tipo de control de nominas
        If RS!ControlNomina = 2 Then
            itmX.SmallIcon = 2
        Else
            itmX.SmallIcon = 1
        End If
        itmX.Tag = DBLet(RS!Dias, "N")
            
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub





Private Sub ImprimirAnticiposAlzira()

    On Error GoTo EImprimirAnticiposAlzira
        
        
    SQLAUX = "EMPRESA= """ & MiEmpresa.NomEmpresa & """|"
    SQLAUX = SQLAUX & "Direccion= """ & MiEmpresa.DirEmpresa & "  -   " & MiEmpresa.CodPosEmpresa & " " & MiEmpresa.PobEmpresa & """|"
    SQLAUX = SQLAUX & "FFIN= """ & Text1(1).Text & """|"
    SQLAUX = SQLAUX & "FINI= """ & Text1(0).Text & """|"
    frmImprimir.Opcion = 2
    frmImprimir.NumeroParametros = 4
    frmImprimir.OtrosParametros = SQLAUX
    frmImprimir.Show vbModal
    Screen.MousePointer = vbDefault
    
    Exit Sub
EImprimirAnticiposAlzira:
    MuestraError Err.Number, Err.Description
End Sub



