VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Programa para el control de presencia"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmMain2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":0764
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":0A7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":6D18
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":7172
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":748C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":8366
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":9240
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":F4DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":1135C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":16F7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":17C58
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain2.frx":19CDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Trabajadores"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Horarios"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Revisar correctos"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Revisar incorrectos"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Operaciones TCP3"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Procesar marcajes"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Traspaso ARIADNA"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Presencia"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Horas trabajadas"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Dias trabajados"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   13
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6060
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   2
      Top             =   7875
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   5292
            MinWidth        =   5292
            Text            =   "(c) Ariadna Software"
            TextSave        =   "(c) Ariadna Software"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10821
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "13:00"
         EndProperty
      EndProperty
   End
   Begin VB.Image Image33 
      Height          =   9000
      Left            =   0
      Picture         =   "frmMain2.frx":19FF4
      Top             =   600
      Width           =   12000
   End
   Begin VB.Image Image2 
      Height          =   570
      Left            =   300
      Top             =   60
      Width           =   1800
   End
   Begin VB.Label Label3 
      BackColor       =   &H00800000&
      Caption         =   "    AriPresencia: Gestión de presencia"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2655
   End
   Begin VB.Menu mnDatos 
      Caption         =   "&Datos básicos"
      Begin VB.Menu mnEmpresas 
         Caption         =   "&Empresas"
      End
      Begin VB.Menu mnsecciones 
         Caption         =   "&Secciones"
      End
      Begin VB.Menu mnTrabajadores 
         Caption         =   "&Trabajadores"
      End
      Begin VB.Menu mnTareas 
         Caption         =   "Tareas"
      End
      Begin VB.Menu mnCategorias 
         Caption         =   "&Categorias"
      End
      Begin VB.Menu mnHorarios 
         Caption         =   "&Horarios"
      End
      Begin VB.Menu mnIncidencias 
         Caption         =   "&Incidencias"
      End
      Begin VB.Menu mnbarr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnSelecImpresora 
         Caption         =   "Seleccionar impresora"
      End
      Begin VB.Menu mn_barra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnConfig 
         Caption         =   "Con&figuración"
      End
      Begin VB.Menu mnbarra14 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnOperaciones 
      Caption         =   "&Operaciones"
      Begin VB.Menu mnRevisar 
         Caption         =   "&Revisar incorrectos"
      End
      Begin VB.Menu mnEntrada 
         Caption         =   "Revisar &marcajes"
      End
      Begin VB.Menu mnPedirFecha 
         Caption         =   "Pedir fecha al revisar marcajes"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnbarra2_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnProcesar 
         Caption         =   "Procesar marcajes"
      End
      Begin VB.Menu mnbarra2_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnImportar 
         Caption         =   "Importar &fichero de datos"
      End
      Begin VB.Menu mnbarra2_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnTraspasar 
         Caption         =   "&Trapaso aplicaciones Ariadna"
      End
      Begin VB.Menu mnbarra2_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnOperacionesTCP3 
         Caption         =   "&Operaciones TCP-3"
      End
   End
   Begin VB.Menu mnGeneraInformes 
      Caption         =   "&Informes"
      Begin VB.Menu mnPresencia 
         Caption         =   "&Presencia"
      End
      Begin VB.Menu mnResumen 
         Caption         =   "&Resumen horas trabajadas"
      End
      Begin VB.Menu mnDiasTrabajados 
         Caption         =   "Dias trabajados"
      End
      Begin VB.Menu mnIncResumen 
         Caption         =   "&Incidencias RESUMEN"
      End
      Begin VB.Menu mnGeneradas 
         Caption         =   "Incidencias &Generadas"
      End
   End
   Begin VB.Menu mnAcerca 
      Caption         =   "Acerca de ..."
      Begin VB.Menu mnAcercaDef 
         Caption         =   "Control de Presencia"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Private FechaRevision As Date

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
'Text1.Text = "" 'Format(Now, "dd/mm/yyyy")
On Error Resume Next
    Image2.Picture = LoadPicture(App.Path & "\iconos\minilogo.bmp")
    If Err.Number > 0 Then
        MsgBox "No se ha encontrado el icono de ARIADNA", vbExclamation
    End If
    Me.Left = 9
    Me.Top = 0
    Me.Width = 12000
    Me.Height = 9000
    BaseForm = Label2.Height + 620
    PonerPedirFecha True
    mnOperacionesTCP3.Enabled = mConfig.TCP3
    Toolbar1.Buttons(7).Visible = mConfig.TCP3
    Toolbar1.Buttons(9).Visible = mConfig.Ariadna
    mnTraspasar.Enabled = mConfig.Ariadna
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    PonerPedirFecha False
    Conn.Close
    Set Conn = Nothing
End Sub

Private Sub frmF_Selec(vFecha As Date)
    'Text1.Text = Format(vFecha, "dd/mm/yyyy")
    FechaRevision = vFecha
End Sub


Private Sub Label1_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 4
        frmEmpresas.Show vbModal
    
    Case 6
        frmCategoria.Show vbModal
    Case 7
       
    Case 8
        frmIncidencias.Show vbModal
    Case 9
        frmRevision2.Todos = 2
        frmRevision2.vFecha = ""
        frmRevision2.Show vbModal
    Case 10
    '    If Text1.Text <> "" Then
    '        If Not IsDate(Text1.Text) Then
    '            MsgBox "La fecha seleccionada no es una fecha correcta.", vbExclamation
    '            Exit Sub
    '        End If
    '    End If
    '    frmRevision2.Todos = (Check1.Value + 1)
    '    frmRevision2.vFecha = Text1.Text
    '    frmRevision2.Show vbModal
    Case 11
        mnImportar_Click
    Case 12
    
    Case 13
    
    Case 14
        'configuracion
        Set frmConfig.vCon = mConfig
        frmConfig.Show vbModal
    Case 15
        'Aqui irá el logo de ARIADNA
        frmAbout.Show vbModal
    Case 16
        Unload Me
    Case 17
       
    Case 18
        
    Case 19
        frmSeccion.Show vbModal
    Case 20
        mnProcesar_Click
    End Select
    Screen.MousePointer = vbDefault
End Sub



Private Sub mnAcercaDef_Click()
Label1_Click 15
End Sub

Private Sub mnCategorias_Click()
Label1_Click 6
End Sub

Private Sub mnConfig_Click()
Label1_Click 14
End Sub


Private Sub mnDiasTrabajados_Click()
'Dias trabajados
frmDiasTrabajados.Show vbModal
End Sub

Private Sub mnEmpresas_Click()
Label1_Click 4
End Sub

Private Sub mnEntrada_Click()
Revisarmarcajes True
End Sub

Private Sub mnGeneradas_Click()
frmInfIncGen.Show vbModal
End Sub

Private Sub mnHorarios_Click()
frmHorario.Show vbModal
End Sub

Private Sub mnImportar_Click()
Screen.MousePointer = vbHourglass
frmTraspaso.Opcion = 0
frmTraspaso.Show vbModal
Screen.MousePointer = vbDefault
End Sub

Private Sub mnIncidencias_Click()
Label1_Click 8
End Sub

Private Sub mnIncResumen_Click()
 frmInfInc.Show vbModal
End Sub

Private Sub mnOperacionesTCP3_Click()
'Utilizaremos esta variable global para saber si hay que importar
'un nuevo ficehero de datos
MostrarErrores = False
frmTCP3.Show vbModal
If MostrarErrores Then
    'Hay que importar
    Screen.MousePointer = vbHourglass
    frmTraspaso.Opcion = 1  'PARA SABER QUE VENIMOS DESDE TCP3
    frmTraspaso.Show vbModal
    Screen.MousePointer = vbDefault
End If
End Sub

Private Sub mnPedirFecha_Click()
Me.mnPedirFecha.Checked = Not Me.mnPedirFecha.Checked
End Sub

Private Sub mnPresencia_Click()
    'Informes
    frmInformes.Opcion = 1
    frmInformes.Show vbModal
End Sub

Private Sub mnProcesar_Click()
MostrarErrores = False
frmProcMarcajes.Show vbModal
If MostrarErrores Then
    frmRevision2.vFecha = ""
    frmRevision2.Todos = 2  'Solo incorrectas
    frmRevision2.Show vbModal
End If
End Sub

Private Sub mnResumen_Click()
    'Informes
    frmInformes.Opcion = 2
    frmInformes.Show vbModal
End Sub

Private Sub mnRevisar_Click()
Revisarmarcajes False
End Sub

Private Sub mnSalir_Click()
Unload Me
End Sub

Private Sub mnsecciones_Click()
Label1_Click 19
End Sub

Private Sub mnSelecImpresora_Click()
On Error Resume Next
Screen.MousePointer = vbHourglass
cd1.DialogTitle = "SELECCIONA LA IMPRESORA"
cd1.ShowPrinter
Screen.MousePointer = vbDefault
End Sub

Private Sub mnTareas_Click()
frmTareas.Show
End Sub

Private Sub mnTrabajadores_Click()
Screen.MousePointer = vbHourglass
frmEmpleados.Show vbModal
End Sub

Private Sub mnTraspasar_Click()
frmUnix.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
    Case 1
        'QUITAR###
        frmKimaldi.Show
        Exit Sub
        mnTrabajadores_Click
    Case 2
        frmHorario.Show vbModal
    Case 4
        Revisarmarcajes True
    Case 5
        Revisarmarcajes False
    Case 7
        mnOperacionesTCP3_Click
    Case 8
        mnProcesar_Click
    Case 9
        mnTraspasar_Click
    Case 11
        mnPresencia_Click
    Case 12
        mnResumen_Click
    Case 13
        mnDiasTrabajados_Click
    Case 15
        Unload Me
    End Select
End Sub


Private Sub PonerPedirFecha(Leer As Boolean)
Dim Cad As String
Dim NF As Integer
On Error GoTo EPonerPedirFecha

Cad = App.Path & "\Pdirfech.dat"
If Leer Then
    If Dir(Cad, vbArchive) = "" Then
        Me.mnPedirFecha.Checked = False
    Else
        Me.mnPedirFecha.Checked = True
    End If
Else
    'Escribir
    If Me.mnPedirFecha.Checked Then
        If Dir(Cad, vbArchive) = "" Then
            'Si no existe el archivo lo creamos
            NF = FreeFile
            Open Cad For Output As #NF
            Print #NF, "Pedir la fecha"
            Close #NF
        End If
    Else
        'lo borramos
        If Dir(Cad, vbArchive) <> "" Then Kill Cad
    End If
End If
Exit Sub
EPonerPedirFecha:
    Err.Clear
End Sub


Private Sub Revisarmarcajes(Todos As Boolean)
    
    FechaRevision = "0:00:00"
    If Me.mnPedirFecha.Checked Then
        Set frmF = New frmCal
        frmF.Fecha = (Now - 1)
        frmF.Show vbModal
        Set frmF = Nothing
    End If
    If FechaRevision = "0:00:00" Then
        'Ha pulsado cancelar
        frmRevision2.vFecha = ""
        Else
        frmRevision2.vFecha = Format(FechaRevision, "dd/mm/yyyy")
    End If
    If Todos Then
        frmRevision2.Todos = 1
    Else
        frmRevision2.Todos = 2
    End If
    frmRevision2.Show vbModal
End Sub

