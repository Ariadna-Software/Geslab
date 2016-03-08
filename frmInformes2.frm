VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmInformes2 
   Caption         =   "Informes"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8235
   Icon            =   "frmInformes2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8235
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameJornadas 
      Height          =   2295
      Left            =   120
      TabIndex        =   53
      Top             =   1800
      Width           =   7935
   End
   Begin VB.Frame FrameInformeAlzira 
      Height          =   4455
      Left            =   0
      TabIndex        =   49
      Top             =   600
      Width           =   8175
      Begin VB.Label Label6 
         Caption         =   "Generando impresión de anticipos para el periodo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   51
         Top             =   1440
         Width           =   7455
      End
      Begin VB.Label Label5 
         Caption         =   "Generando impresión de anticipos para el periodo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   50
         Top             =   600
         Width           =   7455
      End
   End
   Begin VB.Frame FrameResumNomina 
      BorderStyle     =   0  'None
      Height          =   3135
      Left            =   120
      TabIndex        =   42
      Top             =   840
      Visible         =   0   'False
      Width           =   8055
      Begin VB.CheckBox chkA3 
         Caption         =   "Impresión sobre A3"
         Height          =   255
         Left            =   1560
         TabIndex        =   48
         Top             =   840
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   4440
         TabIndex        =   46
         Text            =   "Text1"
         Top             =   240
         Width           =   615
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmInformes2.frx":030A
         Left            =   1560
         List            =   "frmInformes2.frx":0332
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   225
         Width           =   2175
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   0
         TabIndex        =   47
         Top             =   2760
         Width           =   7815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Año"
         Height          =   195
         Index           =   1
         Left            =   3960
         TabIndex        =   44
         Top             =   285
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "MES"
         Height          =   195
         Index           =   0
         Left            =   1080
         TabIndex        =   43
         Top             =   285
         Width           =   345
      End
   End
   Begin VB.CheckBox Check1 
      Caption         =   "RESUMEN"
      Height          =   375
      Left            =   240
      TabIndex        =   38
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Frame FramePresencia 
      Caption         =   "Ordenado por"
      Height          =   855
      Left            =   120
      TabIndex        =   17
      Top             =   4200
      Width           =   7755
      Begin VB.Frame Frame2 
         Height          =   615
         Left            =   4320
         TabIndex        =   39
         Top             =   120
         Width           =   3135
         Begin VB.OptionButton Option2 
            Caption         =   "Nombre"
            Height          =   255
            Index           =   1
            Left            =   1920
            TabIndex        =   41
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Codigo"
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   40
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optTrab 
         Caption         =   "Empleado"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   1395
      End
   End
   Begin VB.TextBox txtEmpleado 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   4920
      TabIndex        =   31
      Text            =   "Text2"
      Top             =   2400
      Width           =   3015
   End
   Begin VB.TextBox txtEmpleado 
      Height          =   285
      Index           =   2
      Left            =   4320
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   2400
      Width           =   555
   End
   Begin VB.TextBox txtEmpleado 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   900
      TabIndex        =   30
      Text            =   "Text2"
      Top             =   2400
      Width           =   2775
   End
   Begin VB.TextBox txtEmpleado 
      Height          =   285
      Index           =   0
      Left            =   240
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   2400
      Width           =   555
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   2
      Top             =   1200
      Width           =   1395
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   4920
      TabIndex        =   3
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox txtEmpresa 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   13020
      TabIndex        =   21
      Text            =   "Text2"
      Top             =   1320
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Index           =   2
      Left            =   12360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox txtEmpresa 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   9120
      TabIndex        =   20
      Text            =   "Text2"
      Top             =   1320
      Visible         =   0   'False
      Width           =   2715
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Index           =   0
      Left            =   8400
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   1320
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.CommandButton cmdInforme 
      Caption         =   "Informe"
      Height          =   375
      Left            =   4620
      TabIndex        =   14
      Top             =   5280
      Width           =   1455
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6360
      TabIndex        =   15
      Top             =   5280
      Width           =   1455
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   7200
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1155
      Left            =   240
      TabIndex        =   32
      Top             =   2880
      Width           =   7935
      Begin VB.TextBox txtIncidencia 
         Height          =   315
         Index           =   0
         Left            =   60
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox txtIncidencia 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   840
         TabIndex        =   34
         Text            =   "Text1"
         Top             =   660
         Width           =   2775
      End
      Begin VB.TextBox txtIncidencia 
         Height          =   315
         Index           =   2
         Left            =   4140
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   660
         Width           =   615
      End
      Begin VB.TextBox txtIncidencia 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   4860
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   660
         Width           =   2835
      End
      Begin VB.Label Label1 
         Caption         =   "Seccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Index           =   4
         Left            =   60
         TabIndex        =   37
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   9
         Left            =   60
         TabIndex        =   36
         Top             =   420
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   8
         Left            =   4140
         TabIndex        =   35
         Top             =   420
         Width           =   420
      End
      Begin VB.Image ImgIncidencia 
         Height          =   240
         Index           =   0
         Left            =   540
         Picture         =   "frmInformes2.frx":039B
         Top             =   390
         Width           =   240
      End
      Begin VB.Image ImgIncidencia 
         Height          =   240
         Index           =   1
         Left            =   4620
         Picture         =   "frmInformes2.frx":049D
         Top             =   390
         Width           =   240
      End
   End
   Begin VB.Frame FrameTrab 
      Caption         =   "Ordenado por"
      Height          =   855
      Left            =   120
      TabIndex        =   16
      Top             =   4200
      Width           =   7755
      Begin VB.OptionButton Option1 
         Caption         =   "HORAS"
         Height          =   255
         Index           =   3
         Left            =   6180
         TabIndex        =   11
         Top             =   420
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Fecha"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   420
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Empleado"
         Height          =   255
         Index           =   1
         Left            =   2220
         TabIndex        =   9
         Top             =   420
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cod. Empleado"
         Height          =   255
         Index           =   2
         Left            =   4200
         TabIndex        =   10
         Top             =   420
         Width           =   1635
      End
   End
   Begin VB.CheckBox chkSoloBolsaHoras 
      Caption         =   "Solo bolsa horas"
      Height          =   255
      Left            =   360
      TabIndex        =   52
      Top             =   3240
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.Image imgEmpleado 
      Height          =   240
      Index           =   1
      Left            =   4800
      Picture         =   "frmInformes2.frx":059F
      Top             =   2130
      Width           =   240
   End
   Begin VB.Image imgEmpleado 
      Height          =   240
      Index           =   0
      Left            =   780
      Picture         =   "frmInformes2.frx":06A1
      Top             =   2130
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Empleado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   29
      Top             =   1740
      Width           =   1695
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   28
      Top             =   2160
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   7
      Left            =   4320
      TabIndex        =   27
      Top             =   2160
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   4
      Left            =   300
      TabIndex        =   26
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   5
      Left            =   4200
      TabIndex        =   25
      Top             =   960
      Width           =   420
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmInformes2.frx":07A3
      Top             =   960
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   4680
      Picture         =   "frmInformes2.frx":08A5
      Top             =   900
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   24
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label lblEmpresa 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   1
      Left            =   12420
      TabIndex        =   23
      Top             =   1050
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.Label lblEmpresa 
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   8460
      TabIndex        =   22
      Top             =   1050
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Image ImageEmp 
      Height          =   240
      Index           =   1
      Left            =   12900
      Picture         =   "frmInformes2.frx":09A7
      Top             =   1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageEmp 
      Height          =   240
      Index           =   0
      Left            =   9000
      Picture         =   "frmInformes2.frx":0AA9
      Top             =   1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   5
      Left            =   8460
      TabIndex        =   19
      Top             =   660
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Label3"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   300
      TabIndex        =   18
      Top             =   60
      Width           =   7035
   End
End
Attribute VB_Name = "frmInformes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Opcion As Integer
 'Tenemos
 ' 1 .- Visualizacion de la entrada de fichajes
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1


Private NombreTabla As String ' en el informe C. report

'Para los labels que aparecerán en el informe
Dim NCampos As Byte 'De 1 a 3
Dim vLabel(5) As String
    'En el informe son 0 campo1, 1 campo 2  .....
Dim Indice As Integer
Dim vIndex As Integer

'Para saber si tiene horas festivas
Dim TieneFestivas As Boolean

Private Sub cmdInforme_Click()
Dim I As Integer
Dim Cad As String
Dim Ordenacion As Byte
Dim Formula
Dim Etiq As String

On Error GoTo EInf
'Cadena de etiquetitas a blancos
NCampos = 0
For I = 0 To 5
    vLabel(I) = ""
Next I
If Opcion = 1 And Me.Option1(3).Value Then
    RealizarHORAS  'Este es un informe paralelo
    Exit Sub
End If


Select Case Opcion
Case 1

    Formula = DevuelveCadenaSQL
    If Formula = "###" Then Exit Sub
    Screen.MousePointer = vbHourglass

    'PRESENCIA
    Ordenacion = DevuelveOrdenacion
    'CR1.Connect = Conn
    Select Case Ordenacion
    Case 0
        'Reales ordenados por nombre
        CR1.ReportFileName = App.Path & "\Informes\pres_fe.rpt"
        Cad = "(Fecha)"
    Case 1
        CR1.ReportFileName = App.Path & "\Informes\pres_Nom.rpt"
        Cad = "(Nombre)"
    Case 2
        CR1.ReportFileName = App.Path & "\Informes\pres_cod.rpt"
        Cad = "(Codigo trab.)"
    End Select
    
       
Case 2
    Screen.MousePointer = vbHourglass


    If TieneFestivas Then

            'Generaremos la tabla
            CargarTablaTemporal
            
            ' Informes horas trabajadas
            Ordenacion = DevuelveOrdenacion
            'CR1.Connect = Conn
            Select Case Ordenacion
            Case 0
                'Reales ordenados por nombre
                CR1.ReportFileName = App.Path & "\Informes\HTFecha.rpt"
                Cad = "(Fecha)"
            Case 1
                CR1.ReportFileName = App.Path & "\Informes\HTEmple.rpt"
                Cad = "(Nombre)"
            End Select
            

    Else
        'No tiene horasfestivas. Es como hasta ahora
        Formula = DevuelveCadenaSQL
        If Formula = "###" Then Exit Sub
        Screen.MousePointer = vbHourglass
    
        'PRESENCIA
        Ordenacion = DevuelveOrdenacion
        'CR1.Connect = Conn
        Select Case Ordenacion
        Case 0
            'Reales ordenados por nombre
            CR1.ReportFileName = App.Path & "\Informes\HorasFech"
            Cad = "(Fecha)"
        Case 1
            CR1.ReportFileName = App.Path & "\Informes\HorasTrab"
            Cad = "(Nombre)"
        End Select
        If Check1.Value Then
            CR1.ReportFileName = CR1.ReportFileName & "RES"
        Else
            'No es resumido. La ordenacion es importante
            If Option2(0).Value Then CR1.ReportFileName = CR1.ReportFileName & "_c"
        End If
        CR1.ReportFileName = CR1.ReportFileName & ".rpt"
        
  
  
  
  
    End If
    
    
Case 3
    'Listados combinados. Loa hacemos todo Aparte, y salimos
    HacerListadoCombinado
    Exit Sub
    
    
Case 4
    'Listados combinados. Loa hacemos todo Aparte, y salimos
    HacerListadoNominas
    Exit Sub
    
    
Case 5
    'lISTADO oficial
    'Es decir, entre los datos k nos piden iremos poniendo los
    'listados como máximo las horas de la jornada, si no las k ha trabajado
    HacerListadoOficial
    Exit Sub
    
    
Case 6, 7
    'Resumen para adjuntar en nominas. Origen: Picassent
    If Combo1.ListIndex < 0 Then
        MsgBox "Seleccione un mes", vbExclamation
        Exit Sub
    End If
    
    If Text1.Text = "" Then
        MsgBox "Seleccione un año", vbExclamation
        Exit Sub
    End If
    
    
    If Opcion = 6 Then
        HacerListadoResumenNomina
    Else
        HacerListadoHorasMEs_A3
    End If
    Label4.Caption = ""
    Exit Sub
    
    
Case 8
    'impresion datos liquidacion de alzira
    
    
    Exit Sub
    
    
Case 9
    'Immpresion con
    Screen.MousePointer = vbHourglass
    GenerarImpresionimportesCostesAlzira
    Exit Sub
Case 10
    If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
        MsgBox "Debe poner las dos fechas", vbExclamation
        Exit Sub
    End If
    ListadoHorasSemanales
    Exit Sub
Case Else
    Exit Sub  'Salimos por que ha habido un error en la seleccion
End Select
    
        'Debug.Print CR1.ReportFileName
        CR1.SelectionFormula = Formula
        Etiq = ""
        For I = 0 To 5
            If vLabel(I) <> "" Then
               Etiq = vLabel(I)
               vLabel(I) = "Campo" & I + 1 & "= """ & Etiq & """ "
            End If
        Next I
        For I = 0 To 5
            CR1.Formulas(I) = vLabel(I)
        Next I
        
        CR1.WindowTitle = "Resumen horas trabajadas " & Cad
        CR1.WindowState = crptMaximized
    
        CR1.Action = 1
        
EInf:
    If Err.Number <> 0 Then MuestraError Err.Number, "Mostrar informe"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

 NombreTabla = "ado"
 FrameResumNomina.Visible = False
 LimpiarCampos
 Label4.Caption = ""
 Check1.Visible = False
 FrameInformeAlzira.Visible = False
 Frame2.Visible = True
 Me.FrameJornadas.Visible = False
 Me.chkSoloBolsaHoras.Visible = False
 Select Case Opcion
    Case 1 'PRESENCIA
        Label3.Caption = "Presencia Real"
        FramePresencia.Visible = False
        FrameTrab.Visible = True
        Frame1.Visible = True

    Case 2  'RESUMEN
        Label3.Caption = "Informe resumen horas trabajadas"
        FramePresencia.Visible = True
        FrameTrab.Visible = False
        Frame1.Visible = True
        Check1.Visible = True
    Case 3
        'Informe combinados
        Label3.Caption = "Informe combinado"
        FramePresencia.Visible = True
        FrameTrab.Visible = False
        Frame1.Visible = True
        
    Case 4
        'Informe combinados
        Label3.Caption = "Informes resumidos HORAS"
        FramePresencia.Visible = False
        FrameTrab.Visible = False
        Frame1.Visible = False
        txtFecha(0).Text = "01/" & Month(Now) & "/" & Year(Now)
        txtFecha(1).Text = Format(Now, "dd/mm/yyyy")
        
    Case 5
        'oficial
        Label3.Caption = "Informe Horas"
        FramePresencia.Visible = True
        FrameTrab.Visible = False
        Frame1.Visible = True
        txtFecha(0).Text = "01/" & Month(Now) & "/" & Year(Now)
        txtFecha(1).Text = Format(Now, "dd/mm/yyyy")
        
        
    Case 6, 7
       'Informe combinados resumen nomina TIPO picassent
       'Informe combinados resumen nomina TIPO picassent
        FrameResumNomina.Visible = True
        If Opcion = 6 Then
            Label3.Caption = "Informes resumidos HORAS Nomi. "
        Else
            Label3.Caption = "LISTADO HORAS DIARIO "
        End If
        FramePresencia.Visible = False
        FrameTrab.Visible = False
        Frame1.Visible = False
        'If Month(Now) = 1 Then
        
        If Month(Now) = 1 Then
            Combo1.ListIndex = 11 'Diciembre año anterior
            Text1.Text = Year(Now) - 1
        Else
            Combo1.ListIndex = Month(Now) - 2
            Text1.Text = Year(Now)
        End If
        
        Me.chkA3.Visible = (Opcion = 7)
    
       
       
    Case 8
        'INFORME DE NOMINAS DE ALZIRA
        FrameInformeAlzira.Visible = True
        Label6.Caption = RecuperaValor(VariableCompartida, 1)
        
        
    Case 9
        'Informe COSTE TRABAJADOR pedido por ALZIRA
        Label3.Caption = "Informes Coste trabajador"
        FramePresencia.Visible = True
        Frame2.Visible = True
        FrameTrab.Visible = False
        Frame1.Visible = False
        Me.chkSoloBolsaHoras.Visible = True
        txtFecha(0).Text = "01/" & Month(Now) & "/" & Year(Now)
        Indice = DiasMes(Month(Now), Year(Now))
        txtFecha(1).Text = Format(Indice, "00") & "/" & Format(Month(Now), "00") & "/" & Year(Now)
        
        
    Case 10
        
        Label3.Caption = "Informes Jornadas semanales"
        FramePresencia.Visible = True
        Frame2.Visible = True
        FrameTrab.Visible = False
        Frame1.Visible = False
        Me.chkSoloBolsaHoras.Visible = True
        txtFecha(0).Text = "01/" & Month(Now) & "/" & Year(Now)
        Indice = DiasMes(Month(Now), Year(Now))
        txtFecha(1).Text = Format(Indice, "00") & "/" & Format(Month(Now), "00") & "/" & Year(Now)
        FrameJornadas.Visible = True
    Case Else
        
End Select
TieneFestivas = (Dir(App.Path & "\TFest.txt") <> "")

End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)

Select Case Indice
Case 0
    txtEmpresa(vIndex).Text = vCodigo
    txtEmpresa(vIndex + 1).Text = vCadena
Case 1
    txtEmpleado(vIndex).Text = vCodigo
    txtEmpleado(vIndex + 1).Text = vCadena
Case 2
    txtIncidencia(vIndex).Text = vCodigo
    txtIncidencia(vIndex + 1).Text = vCadena
End Select

End Sub

Private Sub frmF_Selec(vFecha As Date)
txtFecha(Indice).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
Indice = Index
Set frmF = New frmCal
frmF.Fecha = Now
frmF.Show vbModal
Set frmF = Nothing
End Sub

Private Function DevuelveOrdenacion() As Byte
Dim I As Integer
Select Case Opcion
Case 1
    'PRESENCIA
    For I = 0 To 2
        If Option1(I).Value Then
            DevuelveOrdenacion = I
            Exit Function
        End If
    Next I
Case 2
    'TRABAJADORES
    For I = 0 To 1
        If optTrab(I).Value Then
            DevuelveOrdenacion = I
            Exit Function
        End If
    Next I
End Select
DevuelveOrdenacion = 1
End Function



Private Function DevuelveCadenaSQL() As String
Dim I As Integer
Dim Cad As String
Dim Formula As String
Dim Nexo As String
Dim CADENA As String
Dim v1, v2
Dim F2 As String

DevuelveCadenaSQL = "###"
'LimpiarTags
Formula = ""
Nexo = ""

'---------------------------------------------------------------------------
'Empresa
v1 = 0
v2 = 999999999
I = 0
CADENA = ""
Cad = "Empresa desde "
txtEmpresa(I).Text = Trim(txtEmpresa(I).Text)
If txtEmpresa(I).Text <> "" Then
    If Not IsNumeric(txtEmpresa(I).Text) Then
        MsgBox Cad & " NO es un fecha correcta.", vbExclamation
        Exit Function
        Else
            v1 = CLng(txtEmpresa(I).Text)
            If v1 < 0 Then
                MsgBox Cad & " NO es correcta.", vbExclamation
                Exit Function
            End If
            Formula = Formula & Nexo & "{" & NombreTabla & ".IdEmpresa} >= " & v1
            Nexo = " AND "
            CADENA = CADENA & " desde " & Format(txtEmpresa(I).Text, "00000")
    End If
End If


I = 2
CADENA = ""
Cad = "Empresa hasta "
txtEmpresa(I).Text = Trim(txtEmpresa(I).Text)
If txtEmpresa(I).Text <> "" Then
    If Not IsNumeric(txtEmpresa(I).Text) Then
        MsgBox Cad & " NO es un número correcta.", vbExclamation
        Exit Function
        Else
            v2 = CLng(txtEmpresa(I).Text)
            If v2 < 0 Then
                MsgBox Cad & " NO es correcta.", vbExclamation
                Exit Function
            End If
            Formula = Formula & Nexo & "{" & NombreTabla & ".IdEmpresa} <= " & v2
            Nexo = " AND "
            CADENA = CADENA & " desde " & Format(txtEmpresa(I).Text, "00000")
    End If
End If

If v1 > v2 Then
    MsgBox "Empresa desde es mayor que empresa hasta. ", vbExclamation
    Exit Function
End If

'----------------------------------------------------------------------
'FECHA
v1 = "01/01/1900"
v2 = "31/12/2800"
'fecha desde
I = 0
CADENA = ""
Cad = "Fecha desde "
txtFecha(I).Text = Trim(txtFecha(I).Text)
If txtFecha(I).Text <> "" Then
    If Not IsDate(txtFecha(I).Text) Then
        MsgBox Cad & " NO es un fecha correcta.", vbExclamation
        Exit Function
        Else
            v1 = CDate(txtFecha(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".Fecha} >=#" & Format(txtFecha(I).Text, "yyyy/mm/dd") & "#"
            Nexo = " AND "
            CADENA = CADENA & " desde " & Format(txtFecha(I).Text, "dd/mm/yyyy")
    End If
End If
'fecha hasta
I = 1
Cad = "Fecha hasta "
txtFecha(I).Text = Trim(txtFecha(I).Text)
If txtFecha(I).Text <> "" Then
    If Not IsDate(txtFecha(I).Text) Then
        MsgBox Cad & " NO es un fecha correcta.", vbExclamation
        Exit Function
        Else
            v2 = CDate(txtFecha(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".Fecha} <=#" & Format(txtFecha(I).Text, "yyyy/mm/dd") & "#"
            Nexo = " AND "
            CADENA = CADENA & " hasta " & Format(txtFecha(I).Text, "dd/mm/yyyy")
    End If
End If
Dim aux
v1 = Format(v1, "yyyy/mm/dd")
v2 = Format(v2, "yyyy/mm/dd")
If v1 > v2 Then
    MsgBox "Fecha desde es mayor que fecha hasta. ", vbExclamation
    Exit Function
End If

'Para el encabezado del informe
If CADENA <> "" Then
    vLabel(NCampos) = "Fecha: "
    vLabel(NCampos + 1) = CADENA
    NCampos = NCampos + 2
    CADENA = ""
    
    
End If
'-------------------------------------------------------------------
'Empleado
I = 0
v1 = 0
v2 = 99999999
txtEmpleado(I).Text = Trim(txtEmpleado(I).Text)
Cad = "Empleado desde "
If txtEmpleado(I).Text <> "" Then
    If Not IsNumeric(txtEmpleado(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Function
        Else
            v1 = CLng(txtEmpleado(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".idTrabajador} >=" & txtEmpleado(I).Text
            Nexo = " AND "
            CADENA = CADENA & " desde " & Format(txtEmpleado(I).Text, "00000")
    End If
End If
'Empleado
I = 2
Cad = "Empleado hasta "
txtEmpleado(I).Text = Trim(txtEmpleado(I).Text)
If txtEmpleado(I).Text <> "" Then
    If Not IsNumeric(txtEmpleado(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Function
        Else
            v2 = CLng(txtEmpleado(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".idTrabajador} <=" & txtEmpleado(I).Text
            Nexo = " AND "
            CADENA = CADENA & " hasta " & Format(txtEmpleado(I).Text, "00000")
    End If
End If

If v1 > v2 Then
    MsgBox "Fecha desde es mayor que fecha hasta. ", vbExclamation
    Exit Function
End If

'Para el encabezado del informe
If CADENA <> "" Then
    vLabel(NCampos) = "Empleado: "
    vLabel(NCampos + 1) = CADENA
    NCampos = NCampos + 2
    CADENA = ""
End If



I = 0
v1 = 0
v2 = 99999999

    Select Case Opcion
    Case 1, 2
    
            'SECCION
            Cad = "Seccion desde "
            txtIncidencia(I).Text = Trim(txtIncidencia(I).Text)
            If txtIncidencia(I).Text <> "" Then
                If Not IsNumeric(txtIncidencia(I).Text) Then
                    MsgBox Cad & " NO es numérico.", vbExclamation
                    Exit Function
                    Else
                        v1 = CLng(txtIncidencia(I).Text)
                        Formula = Formula & Nexo & "{" & NombreTabla & ".idSeccion} >=" & txtIncidencia(I).Text
                        Nexo = " AND "
                        CADENA = CADENA & " desde " & Format(txtIncidencia(I).Text, "00000")
                End If
            End If
            'Incidencias
            I = 2
            Cad = "Incidencia hasta "
            If txtIncidencia(I).Text <> "" Then
                If Not IsNumeric(txtIncidencia(I).Text) Then
                    MsgBox Cad & " NO es numérico.", vbExclamation
                    Exit Function
                    Else
                        v2 = CLng(txtIncidencia(I).Text)
                        Formula = Formula & Nexo & "{" & NombreTabla & ".idSeccion} <=" & txtIncidencia(I).Text
                        Nexo = " AND "
                        CADENA = CADENA & " hasta " & Format(txtIncidencia(I).Text, "00000")
                End If
            End If
            
            If v1 > v2 Then
                MsgBox "Incidencia desde es mayor que Incidencia hasta. ", vbExclamation
                Exit Function
            End If

            If CADENA <> "" Then
                vLabel(NCampos) = "Sección: "
                vLabel(NCampos + 1) = CADENA
                NCampos = NCampos + 2
                CADENA = ""
            End If
    Case Else

        '-----------------------------------------------------------------------
        'INCIDENCIA
    
        Cad = "Incidencia desde "
        txtIncidencia(I).Text = Trim(txtIncidencia(I).Text)
        If txtIncidencia(I).Text <> "" Then
            If Not IsNumeric(txtIncidencia(I).Text) Then
                MsgBox Cad & " NO es numérico.", vbExclamation
                Exit Function
                Else
                    v1 = CLng(txtIncidencia(I).Text)
                    Formula = Formula & Nexo & "{" & NombreTabla & ".idInci} >=" & txtIncidencia(I).Text
                    Nexo = " AND "
                    CADENA = CADENA & " desde " & Format(txtIncidencia(I).Text, "00000")
            End If
        End If
        'Incidencias
        I = 2
        Cad = "Incidencia hasta "
        If txtIncidencia(I).Text <> "" Then
            If Not IsNumeric(txtIncidencia(I).Text) Then
                MsgBox Cad & " NO es numérico.", vbExclamation
                Exit Function
                Else
                    v2 = CLng(txtIncidencia(I).Text)
                    Formula = Formula & Nexo & "{" & NombreTabla & ".idInci} <=" & txtIncidencia(I).Text
                    Nexo = " AND "
                    CADENA = CADENA & " hasta " & Format(txtIncidencia(I).Text, "00000")
            End If
        End If
        
        If v1 > v2 Then
            MsgBox "Incidencia desde es mayor que Incidencia hasta. ", vbExclamation
            Exit Function
        End If
        
        
        'Para el encabezado del informe
        If CADENA <> "" Then
            vLabel(NCampos) = "Incidencias: "
            vLabel(NCampos + 1) = CADENA
            NCampos = NCampos + 2
            CADENA = ""
        End If
 End Select
    
    








'Devolvemos la cadena
'Ahora recorremos los textos para hallar la subconsulta
' y saber las etiquetas
 

DevuelveCadenaSQL = Formula
End Function


'Esta funcion modifica la tabla para mostrar el informe por lineas
Private Function DevuelveCadenaSQLTrab() As String
Dim RsBase As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim I As Integer
Dim CadenaSQL As String
Dim CADENA As String
Dim Cad As String
Dim Cad2 As String
Dim C As Long
Dim Fecha As Date
Dim Inc As Integer

DevuelveCadenaSQLTrab = "###"
ObtenCadenaSql CadenaSQL


'Devolvemos la cadena
'Ahora recorremos los textos para hallar la subconsulta
Cad = "SELECT Empresas.NomEmpresa, Trabajadores.NomTrabajador, Marcajes.Entrada, Marcajes.Fecha"
Cad = Cad & " ,Secciones.Nombre"
Cad = Cad & " FROM Empresas ,Trabajadores,Marcajes,Secciones "
Cad = Cad & " WHERE Empresas.IdEmpresa = Trabajadores.IdEmpresa "
Cad = Cad & " AND Trabajadores.IdTrabajador = Marcajes.idTrabajador"
Cad = Cad & " AND Trabajadores.Seccion = Secciones.IdSeccion"
'unimos la cadena sql
Cad = Cad & CadenaSQL


Set RsBase = New ADODB.Recordset
RsBase.Open Cad, Conn, , , adCmdText
If RsBase.EOF Then
    Set RsBase = Nothing
    Exit Function
End If
'Borramos los registros anteriores
Set RT = New ADODB.Recordset
RT.Open "Delete * from tmpPresencia", Conn, , , adCmdText
Set RT = Nothing
'Empezamos para insertar
Set RT = New ADODB.Recordset
RT.CursorType = adOpenKeyset
RT.LockType = adLockOptimistic
RT.Open "Select * from tmpPresencia", Conn, , , adCmdText

Set RS = New ADODB.Recordset
I = 1
While Not RsBase.EOF
    RT.AddNew
    Cad = "Select IdInci,Hora from EntradaMarcajes WHERE IdMarcaje=" & RsBase!Entrada
    Cad = Cad & " ORDER BY Hora"
    RS.Open Cad, Conn, , , adCmdText
    RT!Id = I
    RT!NomEmpresa = RsBase!NomEmpresa
    RT!nomtrabajador = RsBase!nomtrabajador
    RT!Fecha = RsBase!Fecha
    C = 1
    CADENA = ""
    While Not RS.EOF
        Fecha = RS!Hora
        Inc = RS!IdInci
        If Inc > 0 Then
            Cad2 = DevuelveTextoIncidencia(Inc)
            Else
                Cad2 = ""
        End If
        If Cad2 <> "" Then
            If CADENA = "" Then
                CADENA = ".- " & Cad2
                Else
                    CADENA = ".- " & "El marcaje tiene mas de una incidencia."
            End If
        End If
            
        Select Case C
        Case 1
            RT!H1 = Fecha
        Case 2
            RT!h2 = Fecha
        Case 3
            RT!H3 = Fecha
        Case 4
            RT!h4 = Fecha
        Case 5
            RT!h5 = Fecha
        Case 6
            RT!h6 = Fecha
        Case 7
            RT!h7 = Fecha
        Case 8
            RT!h8 = Fecha
        End Select
        RS.MoveNext
        C = C + 1
    Wend
    RT!Incidencias = CADENA
    RT!Seccion = RsBase!Nombre
    RT.Update
    RS.Close
    RsBase.MoveNext
    I = I + 1
Wend
RT.Close
RsBase.Close
Set RS = Nothing
Set RT = Nothing
Set RsBase = Nothing
DevuelveCadenaSQLTrab = "Todo bien"
Exit Function
ErrSQL:
    MsgBox "Error: " & Err.Description, vbExclamation
End Function



Private Sub LimpiarCampos()
Dim I As Integer
 
For I = 0 To 3
    txtEmpleado(I).Text = ""
Next I
For I = 0 To 3
    txtEmpresa(I).Text = ""
Next I
For I = 0 To 3
    txtIncidencia(I).Text = ""
Next I
For I = 0 To 1
    Me.txtFecha(I).Text = ""
Next I
End Sub


Private Sub ImageEmp_Click(Index As Integer)
    Indice = 0
    vIndex = (Index * 2)
    Set frmB = New frmBusca
    frmB.Tabla = "Empresas"
    frmB.CampoBusqueda = "NomEmpresa"
    frmB.CampoCodigo = "IdEmpresa"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPRESAS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub imgEmpleado_Click(Index As Integer)
    Indice = 1
    vIndex = (Index * 2)
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

Private Sub ImgIncidencia_Click(Index As Integer)

    'Ahora pasa a ser seccion
    Indice = 2
    vIndex = (Index * 2)
    Set frmB = New frmBusca
    frmB.Tabla = "Secciones"
    frmB.CampoBusqueda = "Nombre"
    frmB.CampoCodigo = "IdSeccion"
    frmB.TipoDatos = 3
    frmB.Titulo = "SECCIONES"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub Text1_LostFocus()
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then Exit Sub
    If Not IsNumeric(Text1.Text) Then
        MsgBox "Campo numérico", vbExclamation
        Text1.Text = ""
        Exit Sub
    End If
    
    If Val(Text1.Text) > 2100 Then
        MsgBox "Año incorrecto", vbExclamation
        Exit Sub
    End If
        
    
    
End Sub

Private Sub txtEmpleado_GotFocus(Index As Integer)
    With txtEmpleado(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtEmpleado_KeyPress(Index As Integer, KeyAscii As Integer)
 Keypress KeyAscii
End Sub

Private Sub txtEmpleado_LostFocus(Index As Integer)
Dim Cad As String
If Trim(txtEmpleado(Index).Text) = "" Then
    txtEmpleado(Index + 1).Text = ""
    Exit Sub
End If
   
If Not IsNumeric(txtEmpleado(Index).Text) Then
    txtEmpleado(Index).Text = "-1"
    txtEmpleado(Index + 1).Text = "Código de empleado erróneo."
    Else
        Cad = devuelveNombreTrabajador(CInt(txtEmpleado(Index).Text))
        If Cad = "" Then
            'txtEmpleado(Index).Text = "-1"
            'txtEmpleado(Index + 1).Text = "Código de empresa erróneo."
            txtEmpleado(Index + 1).Text = ""
            Else
                txtEmpleado(Index + 1).Text = Cad
        End If
End If
End Sub

Private Sub txtEmpresa_LostFocus(Index As Integer)
Dim Cad As String
If Trim(txtEmpresa(Index).Text) = "" Then
    txtEmpresa(Index + 1).Text = ""
    Exit Sub
End If
   
If Not IsNumeric(txtEmpresa(Index).Text) Then
    txtEmpresa(Index).Text = "-1"
    txtEmpresa(Index + 1).Text = "Código de empresa erróneo."
    Else
        Cad = DevuelveNombreEmpresa(CLng(txtEmpresa(Index).Text))
        If Cad = "" Then
            'txtEmpresa(Index).Text = "-1"
            'txtEmpresa(Index + 1).Text = "Código de empresa erróneo."
            txtEmpresa(Index + 1).Text = ""
            Else
                txtEmpresa(Index + 1).Text = Cad
        End If
End If
End Sub



Private Sub txtFecha_GotFocus(Index As Integer)
    With txtFecha(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
    
If txtFecha(Index).Text <> "" Then
    If EsFechaOK(txtFecha(Index)) Then
        
        Else
            txtFecha(Index).Text = ""
    End If
End If
End Sub


Private Sub Keypress(KeyAscii As Integer)
        If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If

End Sub

Private Sub txtIncidencia_KeyPress(Index As Integer, KeyAscii As Integer)
 Keypress KeyAscii
End Sub

Private Sub txtIncidencia_LostFocus(Index As Integer)
Dim Cad As String
If Trim(txtIncidencia(Index).Text) = "" Then
    txtIncidencia(Index + 1).Text = ""
    Exit Sub
End If
   
If Not IsNumeric(txtIncidencia(Index).Text) Then
    txtIncidencia(Index).Text = "-1"
    txtIncidencia(Index + 1).Text = "Código de sección erróneo."
    Else
        Cad = DevuelveNombreSeccion(CLng(txtIncidencia(Index).Text))
        If Cad = "" Then
            'txtIncidencia(Index).Text = "-1"
            'txtIncidencia(Index + 1).Text = "Código de sección erróneo."
            txtIncidencia(Index + 1).Text = ""
            Else
                txtIncidencia(Index + 1).Text = Cad
        End If
End If
End Sub



Private Sub RealizarHORAS()
Dim Cad As String
Dim Etiq As String
Dim I As Integer

Screen.MousePointer = vbHourglass
Cad = DevuelveCadenaSQLTrab
If Cad <> "###" Then
    'Mostrar el informe
    'CR1.Connect = Conn
    Etiq = ""
    For I = 0 To 5
        If vLabel(I) <> "" Then
           Etiq = vLabel(I)
           vLabel(I) = "Campo" & I + 1 & "= """ & Etiq & """ "
        End If
    Next I
    For I = 0 To 5
        CR1.Formulas(I) = vLabel(I)
    Next I
    CR1.SelectionFormula = ""
    CR1.ReportFileName = App.Path & "\Informes\pres_nom2.rpt"
    CR1.WindowTitle = "Resumen horas."
    CR1.WindowState = crptMaximized
    CR1.Action = 1
End If
Screen.MousePointer = vbDefault
End Sub


Private Sub CargarTablaTemporal()
Dim RO As ADODB.Recordset
Dim RD As ADODB.Recordset
Dim mSql As String
Dim v1 As Single
Dim v2 As Single
Dim v3 As Single
Dim AntFech As Date
Dim AntHora As Long
Dim mH As CHorarios
Dim mEm As CEmpresas
Dim C As Long
Dim aux As String

'Borramos la tabla temporal
Conn.Execute "Delete * from HorasTrabajadas"

'Obtenemos los subqueries
ObtenCadenaSql aux

mSql = "SELECT Secciones.Nombre, Empresas.idEmpresa, Marcajes.Fecha, "
mSql = mSql & " Trabajadores.NomTrabajador, Marcajes.HorasTrabajadas, "
mSql = mSql & " Marcajes.HorasIncid,Trabajadores.IdHorario,Marcajes.IncFinal"
mSql = mSql & " FROM Empresas,Secciones,Trabajadores,Marcajes "
mSql = mSql & " WHERE Empresas.IdEmpresa = Secciones.idEmpresa AND Secciones.IdSeccion = Trabajadores.Seccion AND "
mSql = mSql & " Empresas.IdEmpresa = Trabajadores.IdEmpresa AND "
mSql = mSql & " Trabajadores.IdTrabajador = Marcajes.idTrabajador "

'Condiciones de empresa y demas
If aux <> "" Then mSql = mSql & " " & aux

mSql = mSql & " ORDER BY Marcajes.Fecha,Trabajadores.IdHorario"

Set RO = New ADODB.Recordset
RO.Open mSql, Conn, , , adCmdText
If Not RO.EOF Then
    Set RD = New ADODB.Recordset
    Set mH = New CHorarios
    Set mEm = New CEmpresas
    RD.CursorType = adOpenKeyset
    RD.LockType = adLockOptimistic
    RD.Open "HorasTrabajadas", Conn, , , adCmdTable
    C = 1
    AntHora = -1
    AntFech = "01/01/1900"
    While Not RO.EOF
        With RD
            'Leemos los datos de la empresa
            If RO!IdEmpresa <> mEm.IdEmpresa Then mEm.Leer RO!IdEmpresa
            'Comprobamos si es dia festivo o no
            If AntFech <> RO!Fecha Or AntHora <> RO!IdHorario Then
                    If mH.Leer(RO!IdHorario, RO!Fecha) = 0 Then
                        AntFech = RO!Fecha
                        AntHora = RO!IdHorario
                    End If
            End If
            'Insertamos
            .AddNew
            !Id = C
            !empresa = mEm.NomEmpresa
            !Seccion = RO.Fields(0)
            !Fecha = RO.Fields(2)
            !Nombre = RO.Fields(3)
            If mH.EsDiaFestivo Then
                v3 = RO!HorasIncid
                v2 = 0
                v1 = 0
                Else
                    v3 = 0
                    If RO!IncFinal = mEm.IncHoraExtra Then
                        v2 = RO!HorasIncid
                        Else
                            v2 = 0
                    End If
                    v1 = RO!HorasTrabajadas - v2
            End If
            !horasn = v1
            !HorasE = v2
            !HorasF = v3
            .Update
            C = C + 1
        End With
        'Avanzamos
        RO.MoveNext
    Wend
    RD.Close
    Set RD = Nothing
    Set mH = Nothing
    Set mEm = Nothing
End If
RO.Close
Set RO = Nothing
End Sub




Private Sub ObtenCadenaSql(ByRef CadenaSQL As String)
Dim I As Integer
Dim C As Integer
Dim Cad As String
Dim CADENA As String




CadenaSQL = ""
'Limpiamos los tag y ahora
'fecha desde
I = 0
Cad = "Fecha desde "
If txtFecha(I).Text <> "" Then
    If Not IsDate(txtFecha(I).Text) Then
        MsgBox Cad & " NO es un fecha correcta.", vbExclamation
        Exit Sub
        Else
            CadenaSQL = CadenaSQL & " AND Fecha >=#" & Format(txtFecha(I).Text, "yyyy/mm/dd") & "#"
            CADENA = CADENA & " desde " & Format(txtFecha(I).Text, "dd/mm/yyyy")
    End If
End If
'fecha hasta
I = 1
Cad = "Fecha hasta "
If txtFecha(I).Text <> "" Then
    If Not IsDate(txtFecha(I).Text) Then
        MsgBox Cad & " NO es un fecha correcta.", vbExclamation
        Exit Sub
        Else
            CadenaSQL = CadenaSQL & " AND Fecha <=#" & Format(txtFecha(I).Text, "yyyy/mm/dd") & "#"
            CADENA = CADENA & " hasta " & Format(txtFecha(I).Text, "dd/mm/yyyy")
    End If
End If

'Para el encabezado del informe
If CADENA <> "" Then
    vLabel(NCampos) = "Fecha: "
    vLabel(NCampos + 1) = CADENA
    NCampos = NCampos + 2
    CADENA = ""
End If



'Empleado
I = 0
Cad = "Empleado desde "
If txtEmpleado(I).Text <> "" Then
    If Not IsNumeric(txtEmpleado(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Sub
        Else
            CadenaSQL = CadenaSQL & " AND Trabajadores.idTrabajador >=" & txtEmpleado(I).Text
            CADENA = CADENA & " desde " & Format(txtEmpleado(I).Text, "00000")
    End If
End If
'Empleado
I = 2
Cad = "Empleado hasta "
If txtEmpleado(I).Text <> "" Then
    If Not IsNumeric(txtEmpleado(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Sub
        Else
            CadenaSQL = CadenaSQL & " AND Trabajadores.idTrabajador <=" & txtEmpleado(I).Text
            CADENA = CADENA & " hasta " & Format(txtEmpleado(I).Text, "00000")
    End If
End If

'Para el encabezado del informe
If CADENA <> "" Then
    vLabel(NCampos) = "Empleado: "
    vLabel(NCampos + 1) = CADENA
    NCampos = NCampos + 2
    CADENA = ""
End If

'Las empresas
I = 0
Cad = "Empresa desde "
If txtEmpresa(I).Text <> "" Then
    If Not IsNumeric(txtEmpresa(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Sub
        Else
            CadenaSQL = CadenaSQL & " AND Empresas.idEmpresa >=" & txtEmpresa(I).Text
    End If
End If
I = 2
Cad = "Empresa hasta "
If txtEmpresa(I).Text <> "" Then
    If Not IsNumeric(txtEmpresa(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Sub
        Else
            CadenaSQL = CadenaSQL & " AND Empresas.IdEmpresa <=" & txtEmpresa(I).Text
    End If
End If

'Para el encabezado del informe
If CADENA <> "" Then
    vLabel(NCampos) = "Empresa: "
    vLabel(NCampos + 1) = CADENA
    NCampos = NCampos + 2
    CADENA = ""
End If



'La seccion

Cad = "Seccion desde "
If txtIncidencia(0).Text <> "" Then
    If Not IsNumeric(txtIncidencia(0).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Sub
        Else
            CadenaSQL = CadenaSQL & " AND Secciones.IdSeccion >=" & txtIncidencia(0).Text
            CADENA = CADENA & " Desde " & Format(txtIncidencia(0).Text, "00000")
    End If
End If

Cad = "Seccion hasta "
If txtIncidencia(2).Text <> "" Then
    If Not IsNumeric(txtIncidencia(2).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Sub
        Else
            CadenaSQL = CadenaSQL & " AND Secciones.IdSeccion <=" & txtIncidencia(2).Text
            CADENA = CADENA & " hasta " & Format(txtIncidencia(2).Text, "00000")
    End If
End If

'Para el encabezado del informe
If CADENA <> "" Then
    vLabel(NCampos) = "Seccion: "
    vLabel(NCampos + 1) = CADENA
    NCampos = NCampos + 2
    CADENA = ""
End If


End Sub




Private Sub HacerListadoCombinado()

    On Error GoTo EHacerListadoCombinado
    Screen.MousePointer = vbHourglass
    If CargaDatosCombinados Then
        espera 1
        CR1.SelectionFormula = ""
        If optTrab(0).Value Then
            vLabel(0) = "combifech"
        Else
            vLabel(0) = "combiemp"
        End If
        If Not Option2(0).Value Then
            'Por nombre
            vLabel(0) = vLabel(0) & "N"
        End If
        CR1.ReportFileName = App.Path & "\Informes\" & vLabel(0) & ".rpt"
        CR1.WindowTitle = "Informe combinado Presencia/Ticajes"
        CR1.WindowState = crptMaximized
        CR1.Action = 1
    End If
    Screen.MousePointer = vbDefault
    Exit Sub
EHacerListadoCombinado:
    MuestraError Err.Number, "Hacer Listado Combinado" & vbCrLf & Err.Description
End Sub


'estoy aqui, generando la tabla tmpcombinada por la cual
'listare el informe k ya he hecho. Para ello necesitare
'unos cuantos datos, y la insercion de horas, tal y como pone aqui bajo
'una vez acabado, los listare

'Esta funcion modifica la tabla para mostrar el informe por lineas
Private Function CargaDatosCombinados() As Boolean
Dim RsBase As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim C As Integer
Dim CadenaSQL As String
Dim Cad As String
Dim Fecha As Date
Dim NH As Currency

CargaDatosCombinados = False
ObtenCadenaSql CadenaSQL


'Devolvemos la cadena
'Ahora recorremos los textos para hallar la subconsulta
Cad = "SELECT Marcajes.entrada,Marcajes.idTrabajador,Marcajes.Fecha,Marcajes.HorasTrabajadas,Marcajes.HorasIncid,ExcesoDefecto"
Cad = Cad & " FROM Empresas ,Trabajadores,Marcajes,Incidencias,Secciones"
Cad = Cad & " WHERE Empresas.IdEmpresa = Trabajadores.IdEmpresa "
Cad = Cad & " AND Trabajadores.IdTrabajador = Marcajes.idTrabajador"
Cad = Cad & " AND Trabajadores.Seccion = Secciones.Idseccion"
Cad = Cad & " AND Incidencias.idInci = Marcajes.IncFinal"
Cad = Cad & " AND Marcajes.correcto = True"

'unimos la cadena sql
Cad = Cad & CadenaSQL


Set RsBase = New ADODB.Recordset
RsBase.Open Cad, Conn, , , adCmdText
If RsBase.EOF Then
    MsgBox "Ningun dato con esos valores", vbExclamation
    Set RsBase = Nothing
    Exit Function
End If
'Borramos los registros anteriores
Conn.Execute "Delete * from tmpCombinada"

'Empezamos para insertar
Set RT = New ADODB.Recordset
RT.CursorType = adOpenKeyset
RT.LockType = adLockOptimistic
RT.Open "Select * from tmpCombinada", Conn, , , adCmdText

Set RS = New ADODB.Recordset

While Not RsBase.EOF
    RT.AddNew
    Cad = "Select IdInci,Hora from EntradaMarcajes WHERE IdMarcaje=" & RsBase!Entrada
    Cad = Cad & " ORDER BY Hora"
    RS.Open Cad, Conn, , , adCmdText
    RT!idTrabajador = RsBase!idTrabajador
    RT!Fecha = RsBase!Fecha
    'Trbajadas
    NH = CCur(RsBase!HorasTrabajadas)
    RT!ht = NH
    'Las horas
    If RsBase!excesodefecto Then
        NH = CCur(RsBase!HorasIncid)
    Else
        NH = 0
    End If
    RT!HE = NH
    C = 1
    While Not RS.EOF
        Fecha = RS!Hora
        Select Case C
        Case 1
            RT!H1 = Fecha
        Case 2
            RT!h2 = Fecha
        Case 3
            RT!H3 = Fecha
        Case 4
            RT!h4 = Fecha
        Case 5
            RT!h5 = Fecha
        Case 6
            RT!h6 = Fecha
        Case 7
            RT!h7 = Fecha
        Case 8
            RT!h8 = Fecha
        End Select
        RS.MoveNext
        C = C + 1
    Wend
    RT.Update
    RS.Close
    RsBase.MoveNext
Wend
RT.Close
RsBase.Close
Set RS = Nothing
Set RT = Nothing
Set RsBase = Nothing
CargaDatosCombinados = True
Exit Function
ErrSQL:
    MsgBox "Error: " & Err.Description, vbExclamation
End Function


Private Sub HacerListadoNominas()
Dim Formula As String


On Error GoTo EHacerListadoNominas

        If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
            MsgBox "Escriba un intervalo de fechas", vbExclamation
            Exit Sub
        End If

        'Fechas
        Formula = "marcajes.fecha >= #" & Format(txtFecha(0).Text, "yyyy/mm/dd") & "#"
        Formula = Formula & " AND marcajes.fecha <= #" & Format(txtFecha(1).Text, "yyyy/mm/dd") & "#"
        
        If txtEmpleado(0).Text <> "" Then
            Formula = Formula & " AND idTrabajador >=" & txtEmpleado(0).Text
        End If
        If txtEmpleado(2).Text <> "" Then
            Formula = Formula & " AND idTrabajador <=" & txtEmpleado(2).Text
        End If
        Screen.MousePointer = vbHourglass
        Conn.Execute "Delete from tmpMarcajes "
        
        
        Formula = "Insert into tmpMarcajes SELECT * from Marcajes WHERE  " & Formula
        Conn.Execute Formula
        
        espera 1
        CR1.ReportFileName = App.Path & "\Informes\nomin.rpt"
        CR1.SelectionFormula = ""
        CR1.WindowTitle = "Resumen horas trabajadas "
        CR1.WindowState = crptMaximized
        
        CR1.Action = 1
            
EHacerListadoNominas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Screen.MousePointer = vbDefault
End Sub



Private Sub HacerListadoOficial()
Dim RS As ADODB.Recordset
Dim Formula As String
Dim SQL As String
Dim L As Long
Dim Horas As Currency
        Screen.MousePointer = vbHourglass
        
        On Error GoTo EHacerListadoNominas
        Conn.Execute "Delete from tmpMarcajes "
        
        If txtFecha(0).Text = "" Or txtFecha(1).Text = "" Then
            MsgBox "Escriba un intervalo de fechas", vbExclamation
            Exit Sub
        End If
        Formula = " WHERE Trabajadores.IdTrabajador = Marcajes.idTrabajador AND Marcajes.IncFinal = Incidencias.IdInci"

        'Fechas
        If txtFecha(0).Text <> "" Then Formula = Formula & " AND marcajes.fecha >= #" & Format(txtFecha(0).Text, "yyyy/mm/dd") & "#"
        If txtFecha(1).Text <> "" Then Formula = Formula & " AND marcajes.fecha <= #" & Format(txtFecha(1).Text, "yyyy/mm/dd") & "#"
        
        If txtEmpleado(0).Text <> "" Then
            Formula = Formula & " AND Trabajadores.idTrabajador >=" & txtEmpleado(0).Text
        End If
        If txtEmpleado(2).Text <> "" Then
            Formula = Formula & " AND Trabajadores.idTrabajador <=" & txtEmpleado(2).Text
        End If
        
       If Val(txtIncidencia(2).Text) > 0 Then
            Formula = Formula & " AND Trabajadores.seccion >=" & txtIncidencia(2).Text
        End If
       
        If Val(txtIncidencia(0).Text) > 0 Then
            Formula = Formula & " AND Trabajadores.seccion <=" & txtIncidencia(0).Text
        End If
       
       
        Formula = "Select Marcajes.*,Incidencias.ExcesoDefecto FROM Trabajadores ,Marcajes,Incidencias " & Formula
        Formula = Formula & " AND Marcajes.Correcto = True"
        Set RS = New ADODB.Recordset
        RS.Open Formula, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        L = 0
        Formula = "INSERT INTO tmpMarcajes(Entrada,IdTrabajador,Fecha,HorasTrabajadas) VALUES ("
        While Not RS.EOF
            L = L + 1
            SQL = L & "," & RS!idTrabajador & ",#" & Format(RS!Fecha, FormatoFecha) & "#,"
            
            If RS!IncFinal = 0 Then
                Horas = RS!HorasTrabajadas
            Else
                If RS!excesodefecto Then
                    Horas = RS!HorasTrabajadas - RS!HorasIncid
                Else
                    Horas = RS!HorasTrabajadas
                End If
            End If
            SQL = SQL & TransformaComasPuntos(CStr(Horas)) & ")"
            SQL = Formula & SQL
            Conn.Execute SQL
            'Sig
            RS.MoveNext
        Wend
        RS.Close
        Set RS = Nothing
        If L = 0 Then
            MsgBox "Ningun dato entre esos valores.", vbExclamation
            Exit Sub
        End If
        
        espera 1
        SQL = App.Path & "\Informes\"
        If optTrab(0).Value Then
            If Option2(0).Value Then
                Formula = "HOFFechaC"
            Else
                Formula = "HOFFecha"
            End If
        Else
            If Option2(0).Value Then
                Formula = "HOFempCod"
            Else
                Formula = "HOFempNom"
            End If
        End If
        SQL = SQL & Formula & ".rpt"
        CR1.ReportFileName = SQL
        CR1.SelectionFormula = ""
        CR1.WindowTitle = "HORAS trabajadas "
        CR1.WindowState = crptMaximized
        
        CR1.Action = 1
            
EHacerListadoNominas:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Screen.MousePointer = vbDefault
    

End Sub



Private Sub HacerListadoResumenNomina()
Dim SQL As String
Dim Insert As String
Dim RS As ADODB.Recordset
Dim I As Integer
Dim F1 As Date
Dim F2 As Date
Dim HP As Currency

    On Error GoTo EHacerListadoResumenNomina
    Screen.MousePointer = vbHourglass
    
    Conn.Execute "Delete * from tmpDatosMes"
    espera 0.2
        
    
    Set RS = New ADODB.Recordset
    F1 = CDate("01/" & Combo1.ListIndex + 1 & "/" & Text1.Text)
    I = DiasMes(CInt(Combo1.ListIndex + 1), CInt(Text1.Text))
    F2 = CDate(I & "/" & Combo1.ListIndex + 1 & "/" & Text1.Text)
    
    'Antes del 15 Nov. 2004.   Hemos añadido a la funcioin calcularhor... la opcion 3
    'Que calcula las horas para los tipos control nomina 1,2,3
    'CalculaHorasTrabajadas F1, F2, 0
    CalculaHorasTrabajadas F1, F2, 3
    espera 0.2
    SQL = "SELECT Nominas.*, tmpHoras.HorasT,tmpHoras.HorasC"
    SQL = SQL & " FROM Nominas LEFT JOIN tmpHoras ON Nominas.idTrabajador = tmpHoras.trabajador"
    'Inicio
    SQL = SQL & " WHERE Fecha>=#" & Format(F1, FormatoFecha)
    'Fecha fin
    SQL = SQL & "# AND Fecha <=#" & Format(F2, FormatoFecha) & "#"
    
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    Insert = "INSERT INTO tmpDatosMes(mes,DiasTrabajados ,Trabajador ,HORASN,BolsaAntes,Extras,HorasC,BolsaDespues) "
    Insert = Insert & " VALUES (" & Combo1.ListIndex + 1 & "," & Text1.Text & ","  'Para la fecha en el informe
    
    While Not RS.EOF
        I = I + 1
        
        SQL = RS!idTrabajador & ","

        HP = 0
        If Not IsNull(RS!horast) Then HP = HP + RS!horast
        If Not IsNull(RS!horasc) Then HP = HP + RS!horasc
        If HP = 0 Then HP = -1
        'Horas trabjadas
        SQL = SQL & TransformaComasPuntos(CStr(HP)) & ","
        
        
        'BolsaAntes
        SQL = SQL & TransformaComasPuntos(RS!bolsaantes) & ","
        
        'Horas PLUS
        SQL = SQL & TransformaComasPuntos(RS!HP) & ","
        
        'Horas compensadas en nomina
        SQL = SQL & TransformaComasPuntos(RS!HC) & ","
        
        'Bosa despues
        SQL = SQL & TransformaComasPuntos(RS!bolsadespues) & ")"
    
        
        'Insertamos
        SQL = Insert & SQL
         Conn.Execute SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    
'    'Borramos todas las entradas donde Horas trabajadas sean cero, o menor k cero
'    SQL = "DELETE FROM tmpDatosMes WHERE HORASN<=0"
'    Conn.Execute SQL
'
    
    If I > 0 Then
        'Mostramos el informe
        SQL = App.Path & "\Informes\CombinadoHoras.rpt"
        CR1.ReportFileName = SQL
        CR1.SelectionFormula = ""
        CR1.WindowTitle = "HORAS trabajadas "
        CR1.WindowState = crptMaximized
        CR1.Action = 1
    Else
        MsgBox "Ningun dato generado", vbExclamation
    End If
    
    
EHacerListadoResumenNomina:
    If Err.Number <> 0 Then MuestraError Err.Number, "HacerListadoResumenNomina"
    Set RS = Nothing
    Screen.MousePointer = vbDefault
End Sub



Private Sub HacerListadoHorasMEs_A3()

On Error GoTo EHacerListadoHorasMEs_A3

    Screen.MousePointer = vbHourglass
    If ObtenerHorasDiasyDemas Then
        'IMprimimos
        If chkA3.Value = 1 Then
            CR1.ReportFileName = App.Path & "\Informes\resunomina.rpt"
        Else
            'cr1.ReportFileName = App.Path & "\Informes\resunominaA4.rpt"
            CR1.ReportFileName = App.Path & "\Informes\resunominaA4B.rpt"
        End If
        CR1.SelectionFormula = ""
        CR1.Formulas(0) = "Fecha =""01/" & Combo1.ListIndex + 1 & "/" & Text1.Text & """"
        CR1.Formulas(1) = "Titulo = ""HORAS MES " & UCase(Combo1.List(Combo1.ListIndex)) & """"
        CR1.WindowTitle = "HORAS trabajadas mes"
        CR1.WindowState = crptMaximized
        CR1.Action = 1
    End If
    Screen.MousePointer = vbDefault


    Exit Sub
EHacerListadoHorasMEs_A3:
    MuestraError Err.Number, Err.Description
    Screen.MousePointer = vbDefault
End Sub


Private Function ObtenerHorasDiasyDemas() As Boolean
Dim RT As ADODB.Recordset
Dim RS As Recordset
Dim Horas As Currency
Dim HoraS2 As Currency
Dim Dias As Integer
Dim Cad As String
Dim FI As Date
Dim FF As Date
Dim F As Date
Dim FESTIVOS As String
Dim SABADOS As String
Dim Compuestos As String
Dim Fiesta As Boolean
Dim vH As CHorarios
Dim Insert As String
Dim VALUES As String
Dim I As Integer
Dim HN As Currency
Dim HC As Currency
Dim J As Integer

    Label4.Caption = "Preparar datos"
    Label4.Refresh
    ObtenerHorasDiasyDemas = False
    
    Cad = "DELETE FROM tmpInformehorasmes"
    Conn.Execute Cad
    
    'Cojemos todos los trabajadores del mes
    'Fecha incicio
    Cad = "01/" & Me.Combo1.ListIndex + 1 & "/" & Text1.Text
    FI = CDate(Cad)
    'Fin
    Dias = DiasMes((Combo1.ListIndex + 1), CInt(Text1.Text))
    Cad = Dias & "/" & Me.Combo1.ListIndex + 1 & "/" & Text1.Text
    FF = CDate(Cad)


    'Meto los sabados k hay entre las fechas
    'Para los k tienen el tipo de control nomina de jornadas semanales
    'Concatenare con festivos
    '---------------------------------------
'    F = FI
'    While F <= FF
'        'Si es sabados
'        If Weekday(F, vbMonday) = 6 Then SABADOS = SABADOS & Format(F, "dd/mm/yyyy") & "|"
'        F = DateAdd("d", 1, F)
'    Wend
'






    Cad = "Select trabajadores.idTrabajador,idhorario FROM Marcajes,trabajadores "
    Cad = Cad & " WHERE marcajes.idTrabajador = trabajadores.idtrabajador and "
    Cad = Cad & " Fecha >= #" & Format(FI, FormatoFecha) & "# AND Fecha <= #"
    Cad = Cad & Format(FF, FormatoFecha) & "# "
    Cad = Cad & " AND Trabajadores.idTrabajador <900"
    Cad = Cad & " GROUP BY trabajadores.idTrabajador,idhorario"
    Cad = Cad & " ORDER BY idHorario"
    Set RT = New ADODB.Recordset
    RT.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RT.EOF Then
        MsgBox "Ninguna dato para este mes", vbExclamation
        RT.Close
        Exit Function
    End If
    
    Set vH = New CHorarios
    vH.IdHorario = -1
    Set RS = New ADODB.Recordset
    While Not RT.EOF
        'PARA cada trabajador
        Cad = "SELECT Marcajes.*, Trabajadores.NomTrabajador, Trabajadores.IdHorario, Trabajadores.numDNI, Incidencias.ExcesoDefecto,Trabajadores.controlnomina"
        Cad = Cad & " FROM Trabajadores, Marcajes,Incidencias WHERE "
        Cad = Cad & " Trabajadores.IdTrabajador = Marcajes.idTrabajador  AND "
        Cad = Cad & "  Marcajes.IncFinal = Incidencias.IdInci"
        Cad = Cad & " AND Fecha >= #" & Format(FI, FormatoFecha) & "# AND Fecha <= #"
        Cad = Cad & Format(FF, FormatoFecha) & "# AND marcajes.idTrabajador = " & RT.Fields(0)
        Cad = Cad & " ORDER BY Fecha"

        
        RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        
        If RS!IdHorario <> vH.IdHorario Then
            'Ha cambiado
            FESTIVOS = vH.LeerDiasFestivos(RS!IdHorario, FI, FF)
        End If
        'Si el controlNOMINA es joranadas le sumamos
        Compuestos = FESTIVOS
        If RS!ControlNomina = 3 Then Compuestos = Compuestos & SABADOS
            
        'La cadena del INSERT
        'de insercion. La completamos
        Insert = "INSERT INTO  tmpInformehorasmes (idtrabajador,Nombre,Asesoria"
        
        VALUES = ") VALUES(" & RS!idTrabajador & ",'" & RS!nomtrabajador & "','" & DBLet(RS!numdni) & "'"
        
        
        Dias = 0
        HN = 0
        HC = 0
        J = 1
        Label4.Caption = RS!nomtrabajador
        Label4.Refresh
        While Not RS.EOF
            'Si es en un A4 voy a imprimir guioncitos de los dias k faltan
            If chkA3.Value = 0 Then
                If J < (Day(RS!Fecha) - 1) Then
                    For I = J To Day(RS!Fecha) - 1
                        'Normal
                        Insert = Insert & ",h" & I
                        VALUES = VALUES & ",'- -'"
                        'Compensable
                        Insert = Insert & ",c" & I
                        VALUES = VALUES & ",'- -'"
                    Next I
                End If
            End If
            J = Day(RS!Fecha) + 1
        
            'PARA CADA DIA
            I = InStr(1, Compuestos, Format(RS!Fecha, "dd/mm/yyyy") & "|")
            If I > 0 Then
                'Es festivo
                Fiesta = True
            Else
                Fiesta = False
                Dias = Dias + 1
            End If
            
            
            
            If RS!IncFinal = 0 Then
                Horas = RS!HorasTrabajadas
                HoraS2 = 0
            Else
                'Tiene incidencia
                If RS!excesodefecto Then
                
                    'HORAS EXTRA
                    Horas = RS!HorasTrabajadas - RS!HorasIncid
                    HoraS2 = RS!HorasIncid
                    
                Else
                    'Retraso
                    HoraS2 = 0
                    Horas = RS!HorasTrabajadas
                End If
            End If
            
            If Fiesta Then
                HoraS2 = HoraS2 + Horas
                Horas = 0
            End If
            
            'Horas----> noramles
            If Horas > 0 Then
                Insert = Insert & ",h" & Day(RS!Fecha)
                VALUES = VALUES & ",'" & TransformaComasPuntos(Format(Horas, "00.00'"))
            Else
                If Me.chkA3.Value = 0 Then
                    'Si es sobre a4 imprimo un guioncito
                    Insert = Insert & ",h" & Day(RS!Fecha)
                    VALUES = VALUES & ",'- -'"
                End If
            End If
            'Compensables
            If HoraS2 > 0 Then
                Insert = Insert & ",c" & Day(RS!Fecha)
                VALUES = VALUES & ",'" & TransformaComasPuntos(Format(HoraS2, "00.00'"))
            Else
                If Me.chkA3.Value = 0 Then
                    'Si es sobre a4 imprimo un guioncito
                    Insert = Insert & ",c" & Day(RS!Fecha)
                    VALUES = VALUES & ",'- -'"
                End If
            End If
            
            HN = HN + Horas
            HC = HC + HoraS2
            
            RS.MoveNext
            
            
            
        Wend
        
        
        ''Si es en un A4 voy a imprimir guioncitos de los dias k faltan
        If chkA3.Value = 0 Then
            Cad = Dias
            Dias = DiasMes((Combo1.ListIndex + 1), CInt(Text1.Text))
            If J < Dias Then
                For I = J To Dias
                    'Normal
                    Insert = Insert & ",h" & I
                    VALUES = VALUES & ",'- -'"
                    'Compensable
                    Insert = Insert & ",c" & I
                    VALUES = VALUES & ",'- -'"
                Next I
            End If
            Dias = Val(Cad)
        End If
        
        
        
        
        'Ahora ya tenemos las cadenas
        'de insercion. La completamos
        
        Insert = Insert & ",HT,HN,DT"
        'Valores
        VALUES = VALUES & "," & TransformaComasPuntos(CStr(HN)) & "," & TransformaComasPuntos(CStr(HC))
        VALUES = VALUES & "," & Dias & ")"
        Insert = Insert & VALUES
        Conn.Execute Insert
        
        RS.Close
        'Siuguiente trabajador
        RT.MoveNext
    Wend
    RT.Close
    Label4.Caption = "Actualizando tablas"
    Label4.Refresh
    espera 1.5
    ObtenerHorasDiasyDemas = True
End Function





Private Sub GenerarImpresionimportesCostesAlzira()
Dim F1 As Date
Dim F2 As Date
Dim SQL As String
Dim I As Long
Dim RS As Recordset
Dim Horas As Currency
Dim h2 As Currency

    On Error GoTo EgenerarImpresionimportesCostesAlzira
        
    If txtFecha(0).Text <> "" Then
        SQL = txtFecha(0).Text
    Else
        SQL = Format("01/01/2003", "dd/mm/yyyy")
    End If
    F1 = CDate(SQL)
    
    If txtFecha(1).Text <> "" Then
        SQL = txtFecha(1).Text
    Else
        SQL = Format(Now, "dd/mm/yyyy")
    End If
    F2 = CDate(SQL)
        
    If ComprobarMarcajesCorrectos(F1, F2, False) <> 0 Then
        SQL = "Existen marcajes incorrectos entre las fechas" & vbCrLf & "Desea continuar?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
        
    Conn.Execute "DELETE FROm tmpMarcajes"
    
    'Modificacion para ver los de la bolsa de horas unicamente
'    SQL = "Select *,excesodefecto from marcajes,incidencias where"
'    SQL = SQL & " marcajes.incfinal = incidencias.idinci"
'    SQL = SQL & " AND fecha >= #" & Format(F1, FormatoFecha)
'    SQL = SQL & "#  AND fecha <=#" & Format(F2, FormatoFecha) & "#"
    
    
    SQL = "Select marcajes.*,excesodefecto from marcajes,incidencias,trabajadores where"
    SQL = SQL & " marcajes.idTrabajador = trabajadores.idTrabajador"
    SQL = SQL & " AND marcajes.incfinal = incidencias.idinci"
    SQL = SQL & " AND fecha >= #" & Format(F1, FormatoFecha)
    SQL = SQL & "#  AND fecha <=#" & Format(F2, FormatoFecha) & "#"
    
    If Me.chkSoloBolsaHoras Then SQL = SQL & " AND ControlNomina = 1"
        
    
    
    'Trabajdo desde
    If Me.txtEmpleado(0).Text <> "" Then SQL = SQL & " AND idTrabajador >=" & txtEmpleado(0).Text
        
    'Trabajdo desde
    If Me.txtEmpleado(2).Text <> "" Then SQL = SQL & " AND idTrabajador <=" & txtEmpleado(2).Text
    
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = 0
    While Not RS.EOF
        SQL = "INSERT INTO tmpMarcajes(incFinal,entrada,idTrabajador,Fecha,HorasTrabajadas,HorasIncid) VALUES (0," & I & ","
        SQL = SQL & RS!idTrabajador & ",#"
        SQL = SQL & Format(RS!Fecha, FormatoFecha) & "#,"
        If RS!excesodefecto Then
            Horas = RS!HorasTrabajadas - RS!HorasIncid
            h2 = RS!HorasIncid
        Else
            Horas = RS!HorasTrabajadas
            h2 = 0
        End If
        SQL = SQL & TransformaComasPuntos(CStr(Horas)) & ","
        SQL = SQL & TransformaComasPuntos(CStr(h2)) & ")"
        Conn.Execute SQL
        'Sig
        RS.MoveNext
        I = I + 1
    Wend
    RS.Close
    
    If I = 0 Then
        MsgBox "no se ha generado ningun dato con esos valores", vbExclamation
        Exit Sub
    End If
    
    
    
    'Mostramos el informe
     'Ponemos cadena
    SQL = ""
    If Me.txtFecha(0).Text <> "" Then SQL = SQL & "   Desde : " & txtFecha(0).Text
    If Me.txtFecha(1).Text <> "" Then SQL = SQL & "   Hasta : " & txtFecha(1).Text
    If txtEmpleado(0).Text <> "" Then SQL = SQL & "   Desde : " & txtEmpleado(0).Text
    If txtEmpleado(3).Text <> "" Then SQL = SQL & "   Hasta : " & txtEmpleado(3).Text
    SQL = Trim(SQL)
    If SQL <> "" Then SQL = "Intervalo= """ & SQL & """|"
    Me.Tag = SQL
    
    
    'Obtenemos la SS de la empresa
    SQL = DevuelveDesdeBD("IRPF", "Empresas", "idEmpresa", "1", "N")
    If SQL = "" Then
        SQL = "0"
    Else
        SQL = TransformaComasPuntos(SQL)
    End If
    SQL = "SSEmpresa= " & SQL & "|"
    SQL = Me.Tag & SQL
    Me.Tag = ""
    
    If Me.optTrab(0).Value Then
        I = 3
    Else
        I = 4
    End If
    If Me.Option2(0).Value Then I = I + 10 'El informe 13 y 14
        
    frmImprimir.Opcion = I
    frmImprimir.OtrosParametros = SQL
    frmImprimir.NumeroParametros = 2
    frmImprimir.Show vbModal
    
    Exit Sub
EgenerarImpresionimportesCostesAlzira:
    MuestraError Err.Number, Err.Description
End Sub


Private Sub ListadoHorasSemanales()
Dim RS As ADODB.Recordset
        '8  .- Fecha cod
        '9  .- Fecha nom
        '10  .- Empleado cod
        '11 .- Empleado nom
    
    'Comprobamos si hay algun dato entre las fechas
    vLabel(0) = "SELECT * from JornadasSemanales WHERE "
    vLabel(0) = vLabel(0) & "Fecha >=#" & Format(txtFecha(0).Text, FormatoFecha) & "# AND Fecha <= #" & Format(txtFecha(1).Text, FormatoFecha) & "#"
    Indice = 1
    Set RS = New ADODB.Recordset
    RS.Open vLabel(0), Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then Indice = 0
    RS.Close
    Set RS = Nothing
    
    If Indice = 1 Then
        MsgBox "No hay datos para las fechas seleccionadas", vbExclamation
        Exit Sub
    End If
    
    If Me.optTrab(0).Value Then
        Indice = 8
    Else
        Indice = 10
    End If
    
    If Option2(1).Value Then Indice = Indice + 1
    
    
    frmImprimir.Opcion = Indice
    frmImprimir.OtrosParametros = "FINI= """ & txtFecha(0).Text & """|" & "FFIN= """ & txtFecha(1).Text & """|"
    frmImprimir.NumeroParametros = 2
    frmImprimir.Show vbModal
    

End Sub
