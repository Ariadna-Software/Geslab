VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImprimir 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Impresión"
   ClientHeight    =   2925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6645
   Icon            =   "frmImprimir9.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
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
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Sin definir"
      Top             =   180
      Width           =   6315
   End
   Begin VB.Frame Frame2 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   6435
      Begin VB.CheckBox chkSoloImprimir 
         Caption         =   "Previsualizar"
         Height          =   255
         Left            =   4920
         TabIndex        =   4
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Default         =   -1  'True
      Height          =   375
      Left            =   3780
      TabIndex        =   2
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5220
      TabIndex        =   1
      Top             =   2340
      Width           =   1275
   End
   Begin VB.CommandButton cmdConfigImpre 
      Caption         =   "Sel. &impresora"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   2340
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1320
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   120
      Top             =   1200
      Width           =   615
   End
End
Attribute VB_Name = "frmImprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public opcion As Integer
        '1 .- Listado cobros pendientes por CLIENTE
        '2 .- Anticpos formato alzira, 1 pagina apaisada, detallado las fechas
        '3 .- Horas /precio fecha
        '4 .- Horas /precio Trabajador
        '5 .- Ticaje actual - Nombre
        '6 .-   "      "    - Codigo
        
        '  JORNADAS SEMANALES
        
        '8  .- Fecha cod
        '9  .- Fecha nom
        '10 .- Empleado cod
        '11 .- Empleado nom
        
        
        '12 .- VACIO
        
        '13 y 14 son igual k 3 y 4
        '13 .- Horas /precio fecha   ord por codigo
        '14 .- Horas /precio Trabajador  ord by cod
        
        '15 .- Impresion de la pantalla de Generar nominas
        
        
        
        
        'TENEMOS QUE SUBIR EL CRYSTAL DE 8 A 9
        'PARA ELLO MANDAREMOS TODOS LOS INFORMES AQUI
        'Y ANDE LUEGO PONDREMOS EL VISREPORT
        'PARA ESO VAMOS A EMPEZAR DESDE EL 100
        'LOS DATOS VENDRAN DESDE LOS FORMULARIOS POR ORDEN ALFABETICO
        
        '101    .-DIAS TRABAJADOS
        '102
        '103    .-NO TRAJADOS
        '104
        
        
        '105    .- Lisado trabajaodres
        '106    .-   por codigo
        
        
        '107    .- Informes horarios
        '108    .- EAN trabajadores
        '109    .-  "   tareas
        '110    .- Transferencias
        '111    .- Incedencias
        
        '114    .- LIstados de incidencias  generadas
        '           ""
        '116    .- Presencia ....
        '           ""
        
        '142    .- Nomina con bolsa de horas
        
        
        '       -----------------------------  13-diciembre -2005
        '143    .- TAREA ACTUAL PRODUCCION
        
        
        
        '145.- PRODUCCION
        '
        '   Reservo 15. Los siguientes en 160
        
        
        
        '160
                
        '165    - Listdo nominas
        
        '180    .- LIstados de incidencias  MANUALES
        '181           ""       x codigo

        
        
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |
                                   ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes


Private MostrarTree As Boolean
Private Nombre As String
Private MIPATH As String
Private Lanzado As Boolean
Private primeravez As Boolean



Private Sub cmdConfigImpre_Click()
Screen.MousePointer = vbHourglass
'Me.CommonDialog1.Flags = cdlPDPageNums
CommonDialog1.ShowPrinter
PonerNombreImpresora
Screen.MousePointer = vbDefault
End Sub


Private Sub cmdImprimir_Click()
    Imprime9
End Sub


'Private Function Imprime() As Boolean
'Dim I As Integer
'Dim Cad As String
'
'
'On Error GoTo EchkSoloImprimir
'    Screen.MousePointer = vbHourglass
'    CR1.ReportFileName = MIPATH & Nombre
'
'    For I = 1 To NumeroParametros
'        Cad = RecuperaValor(OtrosParametros, I)
'        CR1.Formulas(I) = Cad
'    Next I
'    CR1.WindowShowGroupTree = MostrarTree
'    CR1.WindowTitle = "Resumen horas incidencias "
'    If chkSoloImprimir.Value = 1 Then
'        CR1.Destination = crptToWindow
'        CR1.WindowState = crptMaximized
'    Else
'        CR1.Destination = crptToPrinter
'    End If
'    CR1.Action = 1
'    Screen.MousePointer = vbDefault
'    Exit Function
'EchkSoloImprimir:
'    MuestraError Err.Number, CR1.ReportFileName & vbCrLf & Err.Description
'End Function


Private Sub Imprime9()

    Screen.MousePointer = vbHourglass
    With frmVisReport
        .FormulaSeleccion = Me.FormulaSeleccion
        .SoloImprimir = (Me.chkSoloImprimir.Value = 0)
        .OtrosParametros = OtrosParametros
        .NumeroParametros = NumeroParametros
        .MostrarTree = MostrarTree
        .Informe = MIPATH & Nombre
        '.ExportarPDF = (chkEMAIL.Value = 1)
        .Show vbModal
    End With
    Unload Me
End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub



Private Sub Form_Activate()
If primeravez Then
    espera 0.1
    
    If SoloImprimir Then
        Imprime9
        Unload Me
    End If
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim Cad As String

primeravez = True
Lanzado = False
CargaICO
Cad = Dir(App.Path & "\impre.dat", vbArchive)


'ReestableceSoloImprimir = False
If Cad = "" Then
    chkSoloImprimir.Value = 0
    Else
    chkSoloImprimir.Value = 1
    'ReestableceSoloImprimir = True
End If
cmdImprimir.Enabled = True
If SoloImprimir Then
    chkSoloImprimir.Value = 0
    Me.Frame2.Enabled = False
    chkSoloImprimir.Visible = False
Else
    Frame2.Enabled = True
    chkSoloImprimir.Visible = True
End If
PonerNombreImpresora
MostrarTree = False



MIPATH = App.Path & "\Informes\"

Select Case opcion
Case 1
    Text1.Text = "Anticipos"
    Nombre = "anticipos.rpt"
    
Case 2
    Text1.Text = "Recibos"
    Nombre = "alzinomina.rpt"

Case 3
    Text1.Text = "Horas con importe Fecha"
    Nombre = "CosteFech.rpt"
    MostrarTree = True
Case 4
    Text1.Text = "Horas con importe Trabajador"
    Nombre = "CosteTra.rpt"
    MostrarTree = True

Case 5
    Text1.Text = "Ticaje actual - Nombre"
    Nombre = "TicajeActual.rpt"
    
Case 6
    Text1.Text = "Ticaje actual - Codigo"
    Nombre = "TicajeActualCod.rpt"
    
    
Case 7
    Text1.Text = "Anticipos con antiguedad"
    Nombre = "anticipos_sin.rpt"
    
    '  JORNADAS SEMANALES
        'JornadasEmpCod.rpt
        '8  .- Fecha cod
        '9  .- Fecha nom
        '10  .- Empleado cod
        '11 .- Empleado nom
        
Case 8 To 11
    Text1.Text = "Jornadas SEMANAS ("
    Nombre = "Jornadas"
    
    'Fecha
    If opcion < 10 Then
        Nombre = Nombre & "Fe"
        Text1.Text = Text1.Text & "Fecha/"
    Else
        MostrarTree = True
        Nombre = Nombre & "Emp"
        Text1.Text = Text1.Text & "Trabajador/"
    End If
        
        
    If opcion = 8 Or opcion = 10 Then
        Nombre = Nombre & "Cod"
        Text1.Text = Text1.Text & "Cod"
    Else
        Nombre = Nombre & "Nom"
        Text1.Text = Text1.Text & "Nombre"
    End If
    
    Nombre = Nombre & ".rpt"
    Text1.Text = Text1.Text & ")"
            
            
Case 13
    Text1.Text = "Horas con importe Fecha(Codigo)"
    Nombre = "CosteFech2.rpt"
    MostrarTree = True
Case 14
    Text1.Text = "Horas con importe Trabajador(Codigo)"
    Nombre = "CosteTra2.rpt"
    MostrarTree = True
            
Case 15
    Text1.Text = "Impresion generación nóminas"
    Nombre = "genNomina.rpt"
    
    
    
    
    
'---------------------------------------------------

Case 101
    Text1.Text = "Dias trabajados"
    Nombre = "Diahor.rpt"
    
Case 102
    Text1.Text = "Dias trabajados (Trabajador)"
    Nombre = "Diahoremp.rpt"
    MostrarTree = True
    
Case 103
    Text1.Text = "Dias NO trabajados (Trabajador)"
    Nombre = "NoTraEmp.rpt"
    MostrarTree = True
Case 104
    Text1.Text = "Dias NO trabajados"
    Nombre = "NoTraDia.rpt"
    
    
Case 105
    Text1.Text = "Listado trabajadores(Codigo)"
    Nombre = "List_tra.rpt"


Case 106
    Text1.Text = "Listado trabajadores(Tarjeta)"
    Nombre = "List_tra2.rpt"

Case 107
    Text1.Text = "Listado horarios"
    Nombre = "infHorario.rpt"

Case 108
    Text1.Text = "Tajetas trabajadores EAN"
    Nombre = "TarjetasT.rpt"

Case 109
    Text1.Text = "TAREAS"
    Nombre = "TarjetaTar.rpt"

Case 110
    Text1.Text = "Transferencia"
    Nombre = "tranferencias.rpt"

Case 111
    Text1.Text = "Listado incidencias"
    Nombre = "list_Inc.rpt"
    
    
Case 112
    Text1.Text = "Incidencia resumen"
    Nombre = "IncResT.rpt"
            
            
Case 113
    Text1.Text = "Incidencia resumen (cod)"
    Nombre = "IncResIn.rpt"
            
            
Case 114
    Text1.Text = "Incidencia generadas (Nombre)"
    Nombre = "incgenn.rpt"
    
            
            
Case 115
    Text1.Text = "Incidencia generadas (Código)"
    Nombre = "incgenc.rpt"
            
   
   
Case 116
    Text1.Text = "Presencia / fecha"
    Nombre = "pres_fe.rpt"
    
Case 117
    Text1.Text = "Presencia / nombre"
    Nombre = "pres_Nom.rpt"
    MostrarTree = True
    
Case 118
    Text1.Text = "Presencia / codigo"
    Nombre = "pres_cod.rpt"
    MostrarTree = True
    
Case 119
    Text1.Text = "Horas trabajadas"
    Nombre = "HTFecha.rpt"
    
    
Case 120
    Text1.Text = "Horas trabajadas (Empleado)"
    Nombre = "HTEmple.rpt"
    MostrarTree = True
    
    
Case 121
    Text1.Text = "Horas trabajadas (Fecha)"
    Nombre = "HorasFech.rpt"
    
Case 122
    Text1.Text = "Horas trabajadas (Trab.)"
    Nombre = "HorasTrab.rpt"
    MostrarTree = True
    
Case 123
    Text1.Text = "Horas trabajadas resumen(Fecha)"
    Nombre = "HorasFechRES.rpt"
    
Case 124
    Text1.Text = "Horas trabajadas resumen(Trab.)"
    Nombre = "HorasTrabRES.rpt"
    MostrarTree = True
Case 125
    Text1.Text = "Horas trabajadas resumen inc (Fecha)"
    Nombre = "HorasFech_c.rpt"
    
Case 126
    Text1.Text = "Horas trabajadas resumen inc(Trab.)"
    Nombre = "HorasTrab_c.rpt"
    MostrarTree = True

    
    
Case 127
    Text1.Text = "Presencia por nombre (2)"
    Nombre = "pres_nom2.rpt"
    MostrarTree = True
    
    
Case 128
    Text1.Text = "Combinado fecha/empleado(cod)"
    Nombre = "combifech.rpt"
Case 129
    Text1.Text = "Combinado empleado(cod)"
    Nombre = "combiemp.rpt"
    MostrarTree = True
    
Case 130
    Text1.Text = "Combinado fecha/empleado(nombre)"
    Nombre = "combifechn.rpt"
    
Case 131
    Text1.Text = "Combinado empleado(nombre)"
    Nombre = "combiempn.rpt"
    MostrarTree = True
    
    
    
Case 132
    Text1.Text = "Resumen nomina"
    Nombre = "nomin.rpt"
    



 
Case 133 To 136
        
        If opcion = 133 Then
                Nombre = "HOFFechaC"
        ElseIf opcion = 134 Then
                Nombre = "HOFFecha"
        ElseIf opcion = 135 Then
                Nombre = "HOFempCod"
        ElseIf opcion > 135 Then
                Nombre = "HOFempNom"
        End If
        Text1.Text = "Listado oficial"
        Nombre = Nombre & ".rpt"
        
Case 137
    Text1.Text = "Combinado HORAS"
    Nombre = "combinadohoras.rpt"
    MostrarTree = True
        
        
        
Case 138
    Text1.Text = "Resumen nomina A3"
    Nombre = "resunomina.rpt"
    
Case 139
    Text1.Text = "Resumen nomina A3"
    Nombre = "resunominaa4b.rpt"
    
    
    
'--------------------- 140 Y 141... Vacios
'Case 140


  '  Text1.Text = "Producción"
  '  Nombre = "ProdCatFec.rpt"
    
'Case 141
  '  Text1.Text = "Producción /Trabajadores"
  '  Nombre = "ProdCatFecS.rpt"
            
            

Case 142
    Text1.Text = "Resumen nomina/bolsa"
    Nombre = "infMesALZIRA.rpt"
    
            
         
Case 143
    Text1.Text = "Tarea actual desde produ."
    Nombre = "tareaactuprod.rpt"
    
    
Case 144

    Text1.Text = "Tarjeta trabajador"
    Nombre = "Tarjeta.rpt"
    
    
 Case 145
    Text1.Text = "Producción"
    Nombre = "ProdCatFec.rpt"
    
Case 146
    Text1.Text = "Producción /Trabajadores"
    Nombre = "ProdCatFecS.rpt"

 Case 147
    Text1.Text = "Producción"
    Nombre = "PCaTraTar.rpt"
    
Case 148
    Text1.Text = "Producción /Trabajadores"
    Nombre = "PCaTraTarS.rpt"


Case 165
    Text1.Text = "Mantenimiento nominas"
    Nombre = "rptNomina.rpt"

        '114 +60
Case 174
    Text1.Text = "Incidencia manual (Nombre)"
    Nombre = "incgenn.rpt"
    
            
        '115 + 70
Case 175
    Text1.Text = "Incidencia manual (Código)"
    Nombre = "incgenc.rpt"


Case 180
    Text1.Text = "Incidencia manuales (Nombre)"
    Nombre = "incmann.rpt"
    
            
            
Case 181
    Text1.Text = "Incidencia manuales (Código)"
    Nombre = "incmanc.rpt"

Case Else
    Text1.Text = "Opcion incorrecta"
    Me.cmdImprimir.Enabled = False
    
    
End Select



Screen.MousePointer = vbDefault
End Sub




'Private Function Imprime() As Boolean
'Dim I As Integer
'Dim Cad As String
'
'
'On Error GoTo EchkSoloImprimir
'    Screen.MousePointer = vbHourglass
'    CR1.ReportFileName = MIPATH & Nombre
'
'    For I = 1 To NumeroParametros
'        Cad = RecuperaValor(OtrosParametros, I)
'        CR1.Formulas(I) = Cad
'    Next I
'    CR1.WindowShowGroupTree = MostrarTree
'    CR1.WindowTitle = "Resumen horas incidencias "
'    If chkSoloImprimir.Value = 1 Then
'        CR1.Destination = crptToWindow
'        CR1.WindowState = crptMaximized
'    Else
'        CR1.Destination = crptToPrinter
'    End If
'    CR1.Action = 1
'    Screen.MousePointer = vbDefault
'    Exit Function
'EchkSoloImprimir:
'    MuestraError Err.Number, CR1.ReportFileName & vbCrLf & Err.Description
'End Function


Private Sub Form_Unload(Cancel As Integer)
    OperacionesArchivoDefecto
End Sub

Private Sub OperacionesArchivoDefecto()
Dim crear  As Boolean
On Error GoTo ErrOperacionesArchivoDefecto

crear = (Me.chkSoloImprimir.Value = 1)
'crear = crear And ReestableceSoloImprimir
If Not crear Then
    Kill App.Path & "\impre.dat"
    Else
        CrearArchivo
End If
ErrOperacionesArchivoDefecto:
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub CrearArchivo()
Dim i As Integer
    On Error GoTo ECrearArchivo
    i = FreeFile
    Open App.Path & "\impre.dat" For Output As #i
    Print #i, Now
    Close #i
    Exit Sub
ECrearArchivo:
    MuestraError Err.Number, "Guardando archivos valores por defecto"
End Sub

Private Sub Text1_DblClick()
Frame2.Tag = Val(Frame2.Tag) + 1
If Val(Frame2.Tag) > 2 Then
    Frame2.Enabled = True
    chkSoloImprimir.Visible = True
End If
End Sub

Private Sub PonerNombreImpresora()
On Error Resume Next
    Label1.Caption = Printer.DeviceName
    If Err.Number <> 0 Then
        Label1.Caption = "No hay impresora instalada"
        Err.Clear
    End If
End Sub

Private Sub CargaICO()
    On Error Resume Next
    Image1.Picture = LoadPicture(App.Path & "\iconos\printer.ico")
    If Err.Number <> 0 Then Err.Clear
End Sub


