VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm frmPpal1 
   BackColor       =   &H00CBCFCD&
   Caption         =   "Control presencia y producción"
   ClientHeight    =   8100
   ClientLeft      =   60
   ClientTop       =   735
   ClientWidth     =   8880
   Icon            =   "frmPpal1.frx":0000
   LinkTopic       =   "MDIForm1"
   NegotiateToolbars=   0   'False
   Picture         =   "frmPpal1.frx":030A
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cd1 
      Left            =   5520
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   585
      Left            =   0
      TabIndex        =   0
      Top             =   7515
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1032
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3175
            MinWidth        =   3175
            Picture         =   "frmPpal1.frx":144A3
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7594
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1323
            MinWidth        =   1323
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   970
            MinWidth        =   970
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   873
            MinWidth        =   882
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "8:33"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
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
         NumListImages   =   17
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":17A65
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":1976F
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":19A89
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":1FD23
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":2017D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":20497
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":21371
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":2224B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":284E5
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":287FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":2E421
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":2F0FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":3117D
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":31497
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":317B1
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":37A4B
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":38925
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Trabajadores"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Horarios"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Revisar correctos"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Revisar incorrectos"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Procesar marcajes"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Traspaso ARIADNA"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Presencia"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Horas trabajadas"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Dias trabajados"
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Traer datos maquinas"
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgListComun 
      Left            =   2460
      Top             =   1380
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":397FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":3B509
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":417AF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":421C1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":479B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":4A165
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":4AA3F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":4B319
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":4BBF3
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":4C4CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":5262F
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":52A89
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":52B9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":52CAD
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":52DBF
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":530D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":58CFB
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":5E4ED
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":5EEFF
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":65199
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":6A98B
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPpal1.frx":7017D
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnDatos 
      Caption         =   "&Datos básicos"
      NegotiatePosition=   1  'Left
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
      Begin VB.Menu mnBancos 
         Caption         =   "Bancos"
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
      Begin VB.Menu mnEmpresas 
         Caption         =   "Datos &empresa"
      End
      Begin VB.Menu mnConfig 
         Caption         =   "Con&figuración"
      End
      Begin VB.Menu mnMantenUsuarios 
         Caption         =   "Mantenimiento de Usuarios"
      End
      Begin VB.Menu mnbarra14 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mnLaboral 
      Caption         =   "&Laboral"
      Begin VB.Menu mnNominas 
         Caption         =   "Nominas"
         Begin VB.Menu mnCalculoHoras 
            Caption         =   "Cálculo horas"
         End
         Begin VB.Menu mnCalculoHoras2 
            Caption         =   "Cálculo horas"
         End
         Begin VB.Menu mnColNominas 
            Caption         =   "Ver datos"
         End
         Begin VB.Menu mnbarraSabados 
            Caption         =   "-"
         End
         Begin VB.Menu mnSemanasEspeciales 
            Caption         =   "Datos semanas especiales"
            Begin VB.Menu mnListadoSemanasEspeciales 
               Caption         =   "Mantenimiento horas semanales"
            End
            Begin VB.Menu mnGenerarSemana 
               Caption         =   "Generar horas semanales"
            End
         End
         Begin VB.Menu mnbarra4 
            Caption         =   "-"
         End
         Begin VB.Menu mnFicheroAsesoria 
            Caption         =   "Generar fichero asesoria"
         End
         Begin VB.Menu mnImportarFichAsesoria 
            Caption         =   "Importar fichero asesoria"
         End
         Begin VB.Menu mnBajaTemporada 
            Caption         =   "Dar baja temporada"
         End
      End
      Begin VB.Menu mnAnticipos 
         Caption         =   "Anticipos"
         Begin VB.Menu mnListadoAnticipos 
            Caption         =   "Listado anticipos"
         End
         Begin VB.Menu mnGeneracionAnticpos 
            Caption         =   "Generacion desde horas"
         End
         Begin VB.Menu mnbarra10 
            Caption         =   "-"
         End
         Begin VB.Menu mnPagosBanco 
            Caption         =   "Generar pagos banco"
         End
      End
      Begin VB.Menu mnBajas2 
         Caption         =   "Bajas"
         Begin VB.Menu mnVisitaMedica 
            Caption         =   "Visita medica"
         End
         Begin VB.Menu mnBarra 
            Caption         =   "-"
         End
         Begin VB.Menu mnBajas 
            Caption         =   "Mantenimento bajas"
         End
         Begin VB.Menu mnBajaAlta 
            Caption         =   "Baja / Alta"
         End
      End
      Begin VB.Menu mnbarra2_8 
         Caption         =   "-"
      End
      Begin VB.Menu mnTiposContrato 
         Caption         =   "Tipos de contrato"
      End
      Begin VB.Menu mnTiposBaja 
         Caption         =   "Tipos de baja"
      End
   End
   Begin VB.Menu mnOperaciones 
      Caption         =   "&Operaciones"
      NegotiatePosition=   1  'Left
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
      Begin VB.Menu mnTicajeActual 
         Caption         =   "Consultar ticaje actual"
      End
      Begin VB.Menu mnbarra2_5 
         Caption         =   "-"
      End
      Begin VB.Menu mnCambioHorarioMenu 
         Caption         =   "Cambios horario"
         Begin VB.Menu mnCabioHorario 
            Caption         =   "Masivo"
         End
         Begin VB.Menu mnCambioHorarioAjuste 
            Caption         =   "Ajustes"
         End
      End
      Begin VB.Menu mnbarra2_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnProcesar 
         Caption         =   "Procesar marcajes"
      End
      Begin VB.Menu mnFichajemasivo 
         Caption         =   "Generacion de ticajes"
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
      Begin VB.Menu mnKreta 
         Caption         =   "Terminal &KRETA"
      End
      Begin VB.Menu mnXPass 
         Caption         =   "X-Pass"
      End
   End
   Begin VB.Menu mnProducción 
      Caption         =   "&Producción"
      Begin VB.Menu mnDatosProduccion 
         Caption         =   "Datos producción"
      End
      Begin VB.Menu mnTareaActual 
         Caption         =   "Tarea actual"
      End
      Begin VB.Menu mnVerTicajesTareas 
         Caption         =   "Ver ticajes/tareas"
      End
      Begin VB.Menu mnInsertarTicajeManual 
         Caption         =   "Insertar ticajes manual"
      End
      Begin VB.Menu mnTraerDatosProduccion 
         Caption         =   "Traer datos maquina"
      End
      Begin VB.Menu mnDatosMaquinaKimaldi 
         Caption         =   "Datos maquina"
      End
      Begin VB.Menu mnbarra51 
         Caption         =   "-"
      End
      Begin VB.Menu mnEliminarDatosKimaldi 
         Caption         =   "Eliminar datos para recalcular"
      End
   End
   Begin VB.Menu mnGeneraInformes 
      Caption         =   "&Informes"
      NegotiatePosition=   1  'Left
      Begin VB.Menu mnHorasOficial 
         Caption         =   "Horas ofi."
      End
      Begin VB.Menu mnbarra52 
         Caption         =   "-"
      End
      Begin VB.Menu mnPresencia 
         Caption         =   "&Presencia"
      End
      Begin VB.Menu mnResumen 
         Caption         =   "&Resumen horas trabajadas"
      End
      Begin VB.Menu mnListHorTrab 
         Caption         =   "Listado horas trabajadas"
         Begin VB.Menu mnCombinado 
            Caption         =   "Combinado"
         End
         Begin VB.Menu mnResumenMensual 
            Caption         =   "Resumen mensual"
         End
         Begin VB.Menu mnImportes 
            Caption         =   "Horas con importes"
         End
         Begin VB.Menu mnListadoHorasJornadas 
            Caption         =   "Horas Jornadas"
         End
      End
      Begin VB.Menu mnInformaesCominiados 
         Caption         =   "Combinados Nom."
         Begin VB.Menu mnResumenHorasNomin 
            Caption         =   "Horas totales"
         End
         Begin VB.Menu mnResumenCuartilla 
            Caption         =   "Resumen cuartilla"
         End
         Begin VB.Menu mnBarra20 
            Caption         =   "-"
         End
         Begin VB.Menu mnNominasBolsa 
            Caption         =   "Nominas/Bolsa"
         End
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
      Begin VB.Menu mnIncidenciaManual 
         Caption         =   "Incidencia &manual"
      End
      Begin VB.Menu mnBarraProd1 
         Caption         =   "-"
      End
      Begin VB.Menu mnInformesproduccion 
         Caption         =   "Produccion"
      End
      Begin VB.Menu mnbarra4_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnGenerar 
         Caption         =   "Generar codigo de barras"
         Begin VB.Menu mnEANTrabajadores 
            Caption         =   "Trabajadores"
         End
         Begin VB.Menu mnEANTareas 
            Caption         =   "Tareas"
         End
      End
   End
   Begin VB.Menu mnAcerca 
      Caption         =   "Acerca de ..."
      NegotiatePosition=   3  'Right
      Begin VB.Menu mnAcercaDef 
         Caption         =   "Control de Presencia y Gestión Laboral"
      End
   End
End
Attribute VB_Name = "frmPpal1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Private FechaRevision As Date
Private primeravez As Boolean


Dim CAPADO2 As Boolean

Private Sub frmF_Selec(vFecha As Date)
    'Text1.Text = Format(vFecha, "dd/mm/yyyy")
    FechaRevision = vFecha
End Sub


Private Sub Label1_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    Select Case Index
    Case 4
        frmEmpresas.Show 'vbModal
    
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
        If vUsu.Nivel < 2 Then
            Set frmConfig.vCon = mConfig
            frmConfig.Show vbModal
        End If
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



Private Sub MDIForm_Activate()
    Screen.MousePointer = vbDefault
    If primeravez Then
        primeravez = False
        If mConfig.TCP3_ Then
            If mConfig.ComprobarHoraReloj Then
                frmTCP3.Comprobar = True
                frmTCP3.Show vbModal
            End If
        End If
    End If
End Sub

Private Sub MDIForm_Load()
Dim B As Boolean

On Error Resume Next
     primeravez = True
   
    Me.Left = 9
    Me.Top = 0
    Me.Width = 12000
    Me.Height = 9000
    PonerPedirFecha True
    
    'Ponemos los dibujitos
    Toolbar1.Buttons(1).Image = 3   'Trabajadores
    Toolbar1.Buttons(2).Image = 2   'Horario
    Toolbar1.Buttons(4).Image = 7  'Revisar
    Toolbar1.Buttons(5).Image = 6  'revisar
    Toolbar1.Buttons(7).Image = 1  'TCP3
    Toolbar1.Buttons(8).Image = 9  'procesar
    Toolbar1.Buttons(9).Image = 11 'Traspasos
    Toolbar1.Buttons(11).Image = 4  'Presecia
    Toolbar1.Buttons(12).Image = 5  'Resumen
    Toolbar1.Buttons(13).Image = 8 'Dias trabajados
    Toolbar1.Buttons(15).Image = 14 'Mauqina
    Toolbar1.Buttons(17).Image = 13  'SAlir
    
    

    
    
    B = mConfig.TCP3_
    If B Then If MiEmpresa.QueEmpresa = 0 Then B = False
    mnOperacionesTCP3.Visible = B
    
    
    B = False
    If MiEmpresa.QueEmpresa = 0 Then B = True
    mnXPass.Visible = B

    Toolbar1.Buttons(7).Visible = mnXPass.Visible Or mnOperacionesTCP3.Visible
    If MiEmpresa.QueEmpresa = 0 Then
        Toolbar1.Buttons(7).ToolTipText = "Lectura relojes XPass"
    Else
        Toolbar1.Buttons(7).ToolTipText = "Operaciones TCP3"
    End If
    
    
    Toolbar1.Buttons(9).Visible = mConfig.Ariadna
    mnTraspasar.Enabled = mConfig.Ariadna
    
    'Si es reloj KIMALDI
    B = (vUsu.Nivel < 2) And mConfig.Kimaldi And MiEmpresa.QueEmpresa <> 1
    Me.Toolbar1.Buttons(15).Visible = B
    Me.mnProducción.Visible = B
    Me.mnEliminarDatosKimaldi.Enabled = B
    B = mConfig.Kimaldi And MiEmpresa.QueEmpresa <> 1
    Me.mnBarraProd1.Visible = B
    Me.mnInformesproduccion.Visible = B
    mnBarraProd1.Visible = B
    mnbarra4_3.Visible = B
    Me.mnGenerar.Visible = B
   
    
    
    'Si no lleva laboral. Antes JUNIO 2011
    'mnLaboral.Visible = mConfig.Laboral
    mnLaboral.Visible = MiEmpresa.LlevaLaboral
    
    
    'Cambiado 20 Octubre 2004
    'No dejamos visible importar ficherito
    '----
    ' mnbarra2_1.Visible = Not mConfig.Kimaldi
    
    B = mConfig.Kimaldi
    If B Then
        If MiEmpresa.QueEmpresa = 4 Then B = False 'para que sea IMPORTAR
    End If
    If Not B Then
        mnImportar.Caption = "Generar entradas presencia"
    Else
        mnImportar.Caption = "Importar &fichero de datos"
    End If
        
   
    Me.StatusBar1.Panels(2).Text = "  Usuario:    " & vUsu.Nombre
    
    'Para los l esten visibles aplicamos el nivel usuario
    'Nivel 0 y 1. ADministrador.
    B = vUsu.Nivel < 2
    Me.mnMantenUsuarios.Enabled = B
    Me.mnConfig.Enabled = B
    Me.mnPagosBanco.Enabled = B
    Me.mnTiposBaja.Enabled = B
    Me.mnTiposContrato = B
    Me.mnProcesar.Enabled = B
    Me.mnImportar.Enabled = B
    mnFicheroAsesoria.Enabled = B
    mnImportarFichAsesoria.Enabled = B
    mnColNominas.Enabled = B
    mnFichajemasivo.Enabled = B
    mnInsertarTicajeManual.Enabled = B
    mnGenerarSemana.Enabled = B
    mnTraerDatosProduccion.Enabled = B
    Toolbar1.Buttons(8).Enabled = B
    'NO PUEDE CONSULTA
    B = vUsu.Nivel > 2
    Toolbar1.Buttons(4).Enabled = Not B
    Toolbar1.Buttons(5).Enabled = Not B
    Me.mnRevisar.Enabled = Not B
    mnEntrada.Enabled = Not B
    
    
    'Si esta capado el programa
    CAPADO2 = False
    If CAPADO2 Then
        Me.mnPagosBanco.Visible = False
        mnFicheroAsesoria.Visible = False
        mnImportarFichAsesoria.Visible = False
        mnAnticipos.Visible = False
    End If
    
    
    
    'mnOperacionesTCP3.Visible = mConfig.TCP3
    'Me.Toolbar1.Buttons(7).Visible = mConfig.TCP3
    
    mnKreta.Visible = MiEmpresa.QueEmpresa <> 0 '<> --> Alzira y Belgida
    
    
    'Ponemos segun la opcion en empresa
    PonerOpcionRevisionNominas
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    PonerPedirFecha False
    conn.Close
    Set conn = Nothing
    End
End Sub

Private Sub mnAcercaDef_Click()
    Label1_Click 15
    'SubirCodigoTrabajadores
End Sub

Private Sub mnBajaAlta_Click()
    If CAPADO2 Then frmCambioEmpleado.Command1(0).Enabled = False
    frmCambioEmpleado.Show vbModal
End Sub

Private Sub mnBajas_Click()
    frmBajas.Show vbModal
End Sub

Private Sub mnBajaTemporada_Click()
    frmBajaTemporada.Show vbModal
End Sub

Private Sub mnBancos_Click()
    frmBancos2.Show vbModal
End Sub

Private Sub mnCabioHorario_Click()
    frmCambioHorario.Opcion = 0
    frmCambioHorario.Show vbModal
End Sub

Private Sub mnCalculoHoras_Click()
    fmCalculoHorasMes.Opcion = 0
    fmCalculoHorasMes.Show 'vbModal
End Sub

Private Sub mnCalculoHoras2_Click()
    fmCalculoHorasMes.Opcion = 1
    fmCalculoHorasMes.Show 'vbModal
End Sub

Private Sub mnCambioHorarioAjuste_Click()
    frmCambioHorario.Opcion = 6
    frmCambioHorario.Show vbModal
End Sub

Private Sub mnCategorias_Click()
    Label1_Click 6
End Sub

Private Sub mnColNominas_Click()
    frmColNominas.Show vbModal
End Sub

Private Sub mnCombinado_Click()
    'Informes
    frmInformes.Opcion = 3
    frmInformes.Show vbModal
End Sub

Private Sub mnConfig_Click()
    Label1_Click 14
End Sub

Private Sub mnDatosMaquinaKimaldi_Click()
    frmDatosKimaldi.Show
End Sub

Private Sub mnDatosProduccion_Click()
    frmProduccion.Show
End Sub

Private Sub mnDiasTrabajados_Click()
    'Dias trabajados
    frmDiasTrabajados.Show vbModal
End Sub

Private Sub mnEANTareas_Click()
    frmImpTarjetas.Opcion = 1
    frmImpTarjetas.Show vbModal
End Sub

Private Sub mnEANTrabajadores_Click()
    frmImpTarjetas.Opcion = 0
    frmImpTarjetas.Show vbModal
End Sub

Private Sub mnEliminarDatosKimaldi_Click()
    frmEliminar.Show vbModal
End Sub

Private Sub mnEmpresas_Click()

    Label1_Click 4
End Sub

Private Sub mnEntrada_Click()
    Revisarmarcajes True
End Sub

Private Sub mnFichajemasivo_Click()
    frmCambioHorario.Opcion = 1
    frmCambioHorario.Show vbModal
End Sub

Private Sub mnFicheroAsesoria_Click()
'    frmAsesoria.OPCION = 0
'    frmAsesoria.Show vbModal
    LanzaAsesoria "/N"
End Sub

Private Sub mnGeneracionAnticpos_Click()
Dim Cad As String
Dim RC As Byte

'Leemos los parametros
    Cad = DevuelveDesdeBD("AplicaAntiguedadHN", "empresas", "idempresa", "1", "N")
    If Cad = "1" Then
        RC = 1
    Else
        RC = 0
    End If
    Cad = DevuelveDesdeBD("AplicaAntiguedadHC", "empresas", "idempresa", "1", "N")
    If Cad = "1" Then
        If RC = 0 Then
            RC = 2 'Solo sobre Compensables  raro raro raro este caso
        Else
            RC = 3
        End If
    End If
    frmGeneraAnti.Antiguedad = RC
    frmGeneraAnti.Show vbModal
End Sub

Private Sub mnGeneradas_Click()
    frmInfIncGen2.Opcion = 0
    frmInfIncGen2.Show vbModal
End Sub

Private Sub mnGenerarSemana_Click()
frmCalculoHorasSemana.Show vbModal
End Sub

Private Sub mnHorarios_Click()
    frmHorario.Show
End Sub

Private Sub mnHorasOficial_Click()
    'Informes
    frmInformes.Opcion = 5
    frmInformes.Show vbModal
End Sub

Private Sub mnImportar_Click()
Dim B As Boolean
    Screen.MousePointer = vbHourglass
    B = mConfig.Kimaldi
    If B Then
        If MiEmpresa.QueEmpresa = 4 Then B = False
    End If
    'Antes febrero 2015
    'If Not B Then
    If B Then
        frmTraspaso.Opcion = 2
    Else
        frmTraspaso.Opcion = 0
    End If
    frmTraspaso.Show vbModal
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnImportarFichAsesoria_Click()
'    frmAsesoria.OPCION = 1
'    frmAsesoria.Show vbModal
    LanzaAsesoria "/I"
End Sub

Private Sub mnImportes_Click()
        'Informes
    frmInformes.Opcion = 9
    frmInformes.Show vbModal
End Sub

Private Sub mnIncidenciaManual_Click()
    frmInfIncGen2.Opcion = 1
    frmInfIncGen2.Show vbModal
End Sub

Private Sub mnIncidencias_Click()
    Label1_Click 8
End Sub

Private Sub mnIncResumen_Click()
     frmInfInc.Show vbModal
End Sub

Private Sub mnInformesproduccion_Click()
    frmInfProduccion.Show vbModal
End Sub

Private Sub mnInsertarTicajeManual_Click()
    frmTicajeManual.Show
End Sub

Private Sub mnKreta_Click()
    
    frmKreta.Show vbModal
End Sub

Private Sub mnListadoAnticipos_Click()
    Screen.MousePointer = vbHourglass
    frmColAnticipos.Show vbModal
End Sub

Private Sub mnListadoHorasJornadas_Click()
    frmInformes.Opcion = 10
    frmInformes.Show vbModal
End Sub

Private Sub mnListadoSemanasEspeciales_Click()
    frmColjorSemana.Show vbModal
End Sub

Private Sub mnMantenUsuarios_Click()
 
    frmMantenusu.Show vbModal
End Sub

Private Sub mnNominasBolsa_Click()
    'Informes
    frmInformes.Opcion = 11
    frmInformes.Show vbModal
End Sub

Private Sub mnOperacionesTCP3_Click()
    'Utilizaremos esta variable global para saber si hay que importar
    'un nuevo ficehero de datos
    MostrarErrores = False
    frmTCP3.Comprobar = False
    frmTCP3.Show vbModal
    If MostrarErrores Then
        'Hay que importar
        Screen.MousePointer = vbHourglass
        frmTraspaso.Opcion = 1  'PARA SABER QUE VENIMOS DESDE TCP3
        frmTraspaso.Show vbModal
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub mnPagosBanco_Click()
    frmPagosBanco2.Opcion = 0
    frmPagosBanco2.Show vbModal
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
frmProcMarcajes2.Opcion = 0
frmProcMarcajes2.Show vbModal
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

Private Sub mnResumenCuartilla_Click()
    'Informes
    frmInformes.Opcion = 4
    frmInformes.Show vbModal
End Sub

Private Sub mnResumenHorasNomin_Click()
    'Informes
    frmInformes.Opcion = 6
    frmInformes.Show vbModal
End Sub

Private Sub mnResumenMensual_Click()
    'Informes
    frmInformes.Opcion = 7
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

Private Sub mnTareaActual_Click()
    Screen.MousePointer = vbHourglass
    frmTareaActual.Opcion = 0
    frmTareaActual.Show
End Sub

Private Sub mnTareas_Click()
    frmTareas.Show
End Sub

Private Sub mnTicajeActual_Click()
    
    Screen.MousePointer = vbHourglass
    frmTareaActual.Opcion = 1
    frmTareaActual.Show
End Sub

Private Sub mnTiposBaja_Click()
    frmTiposDiario.Opcion = 0
    frmTiposDiario.Show vbModal
End Sub

Private Sub mnTiposContrato_Click()
    frmTiposDiario.Opcion = 1
    frmTiposDiario.Show vbModal
End Sub

Private Sub mnTrabajadores_Click()
    Screen.MousePointer = vbHourglass
    frmEmpleados.Show 'vbModal
End Sub

Private Sub mnTraerDatosProduccion_Click()
    Screen.MousePointer = vbDefault
    If MiEmpresa.QueEmpresa = 4 Then mnKreta_Click

    
    Exit Sub
    frmKimaldi.Show vbModal
End Sub

Private Sub mnTraspasar_Click()
    frmUnix.Show vbModal
End Sub

Private Sub mnVerTicajesTareas_Click()
    frmVerTicadasProdu.Show
End Sub

Private Sub mnVisitaMedica_Click()
    frmBajaMedico.Show vbModal
End Sub

Private Sub mnXPass_Click()
    Screen.MousePointer = vbHourglass
    frmXpass.Show vbModal
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Screen.MousePointer = vbHourglass
    Select Case Button.Index
    Case 1
    
        mnTrabajadores_Click
    Case 2
        frmHorario.Show
    Case 4
        Revisarmarcajes True
    Case 5
        Revisarmarcajes False
    Case 7
        If MiEmpresa.QueEmpresa = 0 Then
            mnXPass_Click
        Else
            mnOperacionesTCP3_Click
        End If
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
        'Traer datos reloj
        mnTraerDatosProduccion_Click
    Case 17
        Unload Me
    End Select
    Screen.MousePointer = vbDefault
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
        Screen.MousePointer = vbDefault
        frmF.Show vbModal
        Set frmF = Nothing
    End If
    Screen.MousePointer = vbHourglass
    If FechaRevision = "0:00:00" Then
        'Ha pulsado cancelar
        If Me.mnPedirFecha.Checked Then
            frmRevision2.vFecha = ""
            If Todos Then
                Set frmRevision2 = Nothing
                Exit Sub
            End If
            'frmRevision2.vFecha = Format(FechaRevision, "dd/mm/yyyy")
        Else
            frmRevision2.vFecha = ""
        End If
    Else
        frmRevision2.vFecha = FechaRevision
    End If
    If Todos Then
        frmRevision2.Todos = 1
    Else
        frmRevision2.Todos = 2
    End If
    frmRevision2.Show vbModal
End Sub




Private Sub LanzaAsesoria(Opcion As String)
    On Error GoTo ELanzaAsesoria
    
    If Dir(App.Path & "\Gestoria.exe") = "" Then
        MsgBox "No se encuentra el fichero de gestoria", vbExclamation
        Exit Sub
    End If
    
    Shell App.Path & "\Gestoria.exe " & Opcion, vbNormalFocus
        
    Exit Sub
ELanzaAsesoria:
    MuestraError Err.Number
End Sub





Private Sub PonerOpcionRevisionNominas()

'CadParam = DevuelveDesdeBD("NominaAutomatica", "Empresas", "idEmpresa", 1)
'NParam = Abs(CBool(CadParam))

    '0.-Nomina NO, NO, automatica, es lo de Alzira.Ver modulo
    '1.- Nomina automatica, Picassent
mnCalculoHoras.Visible = MiEmpresa.NominaAutomatica
mnCalculoHoras2.Visible = Not mnCalculoHoras.Visible

End Sub











'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'
'Private Sub SubirCodigoTrabajadores()
'Dim RTT As ADODB.Recordset
'Dim Cad As String
'Dim Errores As Integer
'Dim Bie As Integer
'
'    Set RTT = New ADODB.Recordset
'
'
'
'    Cad = "Select idTrabajador from Trabajadores WHERE idTrabajador >=67 and idTrabajador<200"
'    RTT.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    Errores = 0
'    While Not RTT.EOF
'        Conn.BeginTrans
'        If HacerCambioTrabajador(RTT.Fields(0)) Then
'            Conn.CommitTrans
'
'                'Borramos el trabajador
'            Cad = "DELETE FROM Trabajadores where idTrabajador =" & RTT.Fields(0)
'            Conn.Execute Cad
'
'
'            Bie = Bie + 1
'        Else
'            Conn.RollbackTrans
'            Errores = Errores + 1
'        End If
'       RTT.MoveNext
'    Wend
'
'
'
'
'    RTT.Close
'
'
'End Sub
'
'
'
'
'
'Private Function HacerCambioTrabajador(idTrabajador As Long) As Boolean
'Dim Cad As String
'Dim RS As Recordset
'Dim RT As Recordset
'Dim I As Integer
'
'
'
'
'
'
'
'    On Error GoTo EHacerCambioTrabajador
'    HacerCambioTrabajador = False
'    Set RS = New ADODB.Recordset
'    Cad = "Select * from Trabajadores where idTrabajador = " & idTrabajador
'    RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'    'El recordset de insertar
'    Set RT = New ADODB.Recordset
'    RT.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'    RT.AddNew
'    For I = 0 To RS.Fields.Count - 1
'        RT.Fields(I) = RS.Fields(I)
'    Next I
'
'
'
'    'Baja
'    RS!Numtarjeta = Null
'    RS.Update
'
'
'
'    'Ponemos los nuevos valores
'    RT!idTrabajador = idTrabajador + 300
'    RT.Update
'
'
'    RT.Close
'    RS.Close
'
'
'
'
'    'Actualizamos las entradafichajes
'    Cad = "UPDATE EntradaFichajes SET idTrabajador=" & idTrabajador + 300
'    Cad = Cad & " WHERE idTrabajador =" & idTrabajador
'    'Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
'    Conn.Execute Cad
'
'
'    'Actualizamos los marcajes
'    Cad = "UPDATE Marcajes SET idTrabajador=" & idTrabajador + 300
'    Cad = Cad & " WHERE idTrabajador =" & idTrabajador
'    'Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
'    Conn.Execute Cad
'
'
'
'    'Actualizamos Pagos
'    Cad = "UPDATE Pagos SET Trabajador=" & idTrabajador + 300
'    Cad = Cad & " WHERE Trabajador =" & idTrabajador
'    'Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
'    Conn.Execute Cad
'
'
'    'Nominas
'    Cad = "UPDATE Nominas SET idTrabajador=" & idTrabajador + 300
'    Cad = Cad & " WHERE idTrabajador =" & idTrabajador
'    'Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
'    Conn.Execute Cad
'
'
'    'Bajas
'    Cad = "UPDATE Bajas SET idTrab=" & idTrabajador + 300
'    Cad = Cad & " WHERE idTrab =" & idTrabajador
'    'Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
'    Conn.Execute Cad
'
'
'
'
'
'
'
'    HacerCambioTrabajador = True
'    Exit Function
'EHacerCambioTrabajador:
'    MsgBox Err.Description
'
'End Function
'
