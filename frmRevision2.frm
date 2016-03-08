VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRevision2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de revisión de marcajes"
   ClientHeight    =   8340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   Icon            =   "frmRevision2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8340
   ScaleWidth      =   11910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelaBusqueda 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   9240
      TabIndex        =   71
      Top             =   7920
      Width           =   1155
   End
   Begin VB.Frame Frame4 
      Height          =   4035
      Left            =   2280
      TabIndex        =   64
      Top             =   2400
      Width           =   6675
      Begin VB.Label Label19 
         Caption         =   "Label15"
         Height          =   255
         Left            =   4800
         TabIndex        =   70
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label18 
         Caption         =   "Label15"
         Height          =   255
         Left            =   1500
         TabIndex        =   69
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label17 
         Caption         =   "Incorrectos"
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
         Left            =   3660
         TabIndex        =   68
         Top             =   3240
         Width           =   975
      End
      Begin VB.Label Label16 
         Caption         =   "Correctos"
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
         Left            =   540
         TabIndex        =   67
         Top             =   3240
         Width           =   825
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "Label15"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1440
         TabIndex        =   66
         Top             =   1920
         Width           =   3690
      End
      Begin VB.Label Label14 
         Caption         =   "Revisión masiva incorrectos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   180
         TabIndex        =   65
         Top             =   1140
         Width           =   6375
      End
   End
   Begin VB.TextBox txtDec 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   1
      Left            =   8040
      TabIndex        =   6
      Text            =   "Text2"
      Top             =   1980
      Width           =   795
   End
   Begin VB.TextBox txtDec 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   0
      Left            =   8040
      TabIndex        =   5
      Text            =   "Text2"
      Top             =   1440
      Width           =   795
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Poner Ok"
      Height          =   315
      Left            =   4200
      TabIndex        =   56
      Top             =   1980
      Width           =   915
   End
   Begin VB.TextBox txtMarcaje 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   55
      Text            =   "Text1"
      Top             =   1200
      Width           =   855
   End
   Begin VB.Frame Frame3 
      Height          =   3795
      Left            =   120
      TabIndex        =   51
      Top             =   3840
      Visible         =   0   'False
      Width           =   11595
      Begin VB.CommandButton Command4 
         Caption         =   "Aceptar marcaje"
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
         Left            =   7380
         TabIndex        =   54
         Top             =   2160
         Width           =   3255
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Cancelar"
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
         Left            =   660
         TabIndex        =   53
         Top             =   2160
         Width           =   3255
      End
      Begin VB.Label Label8 
         Caption         =   $"frmRevision2.frx":030A
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   600
         TabIndex        =   52
         Top             =   720
         Width           =   10215
      End
   End
   Begin VB.CommandButton cmdRevisada 
      Caption         =   "&Revisada"
      Height          =   435
      Left            =   10140
      TabIndex        =   50
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      Height          =   435
      Left            =   9900
      TabIndex        =   47
      Top             =   840
      Width           =   1515
      Begin VB.CheckBox Check1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   48
         Top             =   60
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Correcto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   315
         Left            =   420
         TabIndex        =   49
         Top             =   60
         Width           =   1155
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "  &Mod. marcajes"
      Height          =   675
      Left            =   3480
      TabIndex        =   46
      Top             =   6960
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Index           =   2
      Left            =   11220
      Picture         =   "frmRevision2.frx":03AF
      Style           =   1  'Graphical
      TabIndex        =   45
      ToolTipText     =   "Borrar"
      Top             =   6060
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Index           =   1
      Left            =   11220
      Picture         =   "frmRevision2.frx":04B1
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Modificar"
      Top             =   5520
      Width           =   435
   End
   Begin VB.CommandButton Command1 
      Height          =   435
      Index           =   0
      Left            =   11220
      Picture         =   "frmRevision2.frx":05B3
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Añadir"
      Top             =   4980
      Width           =   435
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   9240
      TabIndex        =   9
      Top             =   7920
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   7860
      TabIndex        =   10
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7920
      Width           =   1155
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   555
      Left            =   0
      TabIndex        =   40
      Top             =   7785
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   979
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
            Text            =   "texto"
            TextSave        =   "texto"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   13758
            MinWidth        =   13758
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   420
      Top             =   4680
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
            Picture         =   "frmRevision2.frx":06B5
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":09CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":0CE9
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2415
      Left            =   180
      TabIndex        =   36
      Top             =   4260
      Width           =   3555
      _ExtentX        =   6271
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hora"
         Object.Width           =   2082
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Incidencia Manual"
         Object.Width           =   3528
      EndProperty
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   2400
      TabIndex        =   31
      Top             =   7320
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   5
      Left            =   300
      TabIndex        =   29
      Top             =   7320
      Width           =   615
   End
   Begin VB.TextBox txtMarcaje 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   4
      Left            =   6960
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1440
      Width           =   795
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1395
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   11595
      Begin VB.TextBox txtHorario 
         Height          =   315
         Index           =   6
         Left            =   9660
         TabIndex        =   35
         Text            =   "Text1"
         Top             =   270
         Width           =   435
      End
      Begin VB.TextBox txtHorario 
         Height          =   315
         Index           =   5
         Left            =   6780
         TabIndex        =   33
         Text            =   "Text1"
         Top             =   270
         Width           =   495
      End
      Begin VB.TextBox txtHorario 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   1
         Left            =   1320
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   780
         Width           =   1300
      End
      Begin VB.TextBox txtHorario 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   2
         Left            =   3900
         TabIndex        =   21
         Text            =   "Text1"
         Top             =   780
         Width           =   1300
      End
      Begin VB.TextBox txtHorario 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   3
         Left            =   6840
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   780
         Width           =   1300
      End
      Begin VB.TextBox txtHorario 
         Alignment       =   2  'Center
         Height          =   315
         Index           =   4
         Left            =   9600
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   780
         Width           =   1300
      End
      Begin VB.TextBox txtHorario 
         Height          =   315
         Index           =   0
         Left            =   1800
         TabIndex        =   18
         Text            =   "Text2"
         Top             =   270
         Width           =   3315
      End
      Begin VB.Image Image4 
         Height          =   255
         Left            =   1320
         Top             =   360
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre horario"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   38
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Nº de marcajes"
         Height          =   195
         Index           =   3
         Left            =   8520
         TabIndex        =   34
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "Horas jornada"
         Height          =   195
         Index           =   2
         Left            =   5700
         TabIndex        =   32
         Top             =   330
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Hora entrada 1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   7
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Hora salida 1"
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   8
         Left            =   2940
         TabIndex        =   25
         Top             =   840
         Width           =   930
      End
      Begin VB.Label Label1 
         Caption         =   "Hora entrada 2"
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   9
         Left            =   5700
         TabIndex        =   24
         Top             =   840
         Width           =   1065
      End
      Begin VB.Label Label1 
         Caption         =   "Hora salida 2"
         ForeColor       =   &H00004080&
         Height          =   195
         Index           =   10
         Left            =   8580
         TabIndex        =   23
         Top             =   840
         Width           =   930
      End
   End
   Begin VB.TextBox txtMarcaje 
      Alignment       =   1  'Right Justify
      Height          =   315
      Index           =   7
      Left            =   6960
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   1980
      Width           =   825
   End
   Begin VB.TextBox txtMarcaje 
      Enabled         =   0   'False
      Height          =   315
      Index           =   6
      Left            =   1140
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   1980
      Width           =   2985
   End
   Begin VB.TextBox txtMarcaje 
      Height          =   315
      Index           =   5
      Left            =   180
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1980
      Width           =   765
   End
   Begin VB.TextBox txtMarcaje 
      Height          =   315
      Index           =   1
      Left            =   7320
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   540
      Width           =   1155
   End
   Begin VB.TextBox txtMarcaje 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   1140
      TabIndex        =   11
      Text            =   "3"
      Top             =   1200
      Width           =   3015
   End
   Begin VB.TextBox txtMarcaje 
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   480
      Width           =   1275
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2415
      Left            =   6240
      TabIndex        =   37
      Top             =   4260
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Inidencia automática"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Horas (Sexage)"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "en decimal"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Incidencia"
         Object.Width           =   2
      EndProperty
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   510
      Left            =   1140
      Top             =   6120
      Visible         =   0   'False
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   900
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
      ConnectStringType=   3
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
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4200
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":1003
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":1115
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":1227
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":1339
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":144B
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":155D
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":1E37
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":2711
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision2.frx":2FEB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   39
      Top             =   0
      Width           =   11910
      _ExtentX        =   21008
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Todos"
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
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
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
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   1260
      TabIndex        =   57
      Top             =   7320
      Width           =   915
   End
   Begin MSComctlLib.ListView ListView3 
      Height          =   2415
      Left            =   3960
      TabIndex        =   62
      Top             =   4260
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hora"
         Object.Width           =   2258
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   353
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   5640
      Picture         =   "frmRevision2.frx":38C5
      Top             =   3960
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "TICAJES MAQUINA"
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
      Left            =   3960
      TabIndex        =   63
      Top             =   4020
      Width           =   1680
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Decimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   7020
      TabIndex        =   61
      Top             =   1140
      Width           =   690
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Sexagesimal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   210
      Left            =   7920
      TabIndex        =   60
      Top             =   1140
      Width           =   1110
   End
   Begin VB.Line Line1 
      X1              =   180
      X2              =   2280
      Y1              =   7020
      Y2              =   7020
   End
   Begin VB.Label Label10 
      Caption         =   "Sexagesimal"
      Height          =   195
      Left            =   1260
      TabIndex        =   59
      Top             =   7080
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "Decimal"
      Height          =   195
      Left            =   300
      TabIndex        =   58
      Top             =   7080
      Width           =   675
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   6900
      Picture         =   "frmRevision2.frx":42C7
      Top             =   600
      Width           =   240
   End
   Begin VB.Label Label6 
      Caption         =   "INCIDENCIAS GENERADAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6240
      TabIndex        =   42
      Top             =   4020
      Width           =   2955
   End
   Begin VB.Label Label2 
      Caption         =   "MARCAJES"
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
      Left            =   180
      TabIndex        =   41
      Top             =   4020
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Nº de marcajes"
      Height          =   195
      Index           =   1
      Left            =   2280
      TabIndex        =   30
      Top             =   7080
      Width           =   1155
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Horas Trabajadas"
      Height          =   195
      Left            =   5580
      TabIndex        =   28
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "Horas Trabajadas"
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
      Index           =   0
      Left            =   480
      TabIndex        =   27
      Top             =   6780
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   2
      Left            =   1980
      Picture         =   "frmRevision2.frx":43C9
      Top             =   1740
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   840
      Picture         =   "frmRevision2.frx":44CB
      Top             =   960
      Width           =   240
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Incidencia Resumen"
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
      Index           =   1
      Left            =   180
      TabIndex        =   16
      Top             =   1740
      Width           =   1740
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nº horas incidencia"
      Height          =   255
      Index           =   12
      Left            =   5520
      TabIndex        =   15
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Fecha"
      Height          =   195
      Index           =   2
      Left            =   5820
      TabIndex        =   13
      Top             =   600
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Empleado"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   12
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Nº Secuencia"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   540
      Width           =   1215
   End
   Begin VB.Menu mnOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnBuscar 
         Caption         =   "&Buscar"
      End
      Begin VB.Menu mnVerTodos 
         Caption         =   "&Ver todos"
      End
      Begin VB.Menu mnBarra1 
         Caption         =   "-"
      End
      Begin VB.Menu mnNuevo 
         Caption         =   "&Nuevo"
      End
      Begin VB.Menu mnModificar 
         Caption         =   "&Modificar"
      End
      Begin VB.Menu mnEliminar 
         Caption         =   "&Eliminar"
      End
      Begin VB.Menu mnbarra2 
         Caption         =   "-"
      End
      Begin VB.Menu mnSalir 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu mntipo 
      Caption         =   "&Marcajes"
      Begin VB.Menu mnCorrectas 
         Caption         =   "&Correctos"
      End
      Begin VB.Menu mnIncorrectas 
         Caption         =   "&Incorrectos"
      End
      Begin VB.Menu mnTodas 
         Caption         =   "&Todos"
      End
      Begin VB.Menu mnbarra21 
         Caption         =   "-"
      End
      Begin VB.Menu mnQuitarBUSQ 
         Caption         =   "Quitar búsqueda"
      End
   End
   Begin VB.Menu mnFecha 
      Caption         =   "&Fecha"
      Begin VB.Menu mnCualquiera 
         Caption         =   "Cualquiera"
      End
      Begin VB.Menu mnKfecha 
         Caption         =   "Fecha: "
      End
   End
   Begin VB.Menu mnOrdenar 
      Caption         =   "Ordenacion"
      Begin VB.Menu mnPorFecha 
         Caption         =   "Fecha"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnTrabajador 
         Caption         =   "Trabajador"
      End
   End
   Begin VB.Menu mnSeccion 
      Caption         =   "Secciones"
      Begin VB.Menu mnSeccionT 
         Caption         =   "Todas"
      End
      Begin VB.Menu mnSeccion1 
         Caption         =   "Seccion:"
      End
   End
   Begin VB.Menu mnredondeo 
      Caption         =   "Redondeo"
      Begin VB.Menu mnDecima 
         Caption         =   "Décima de hora"
      End
      Begin VB.Menu mncuartos 
         Caption         =   "1/4 de hora"
      End
      Begin VB.Menu mnMediaHora 
         Caption         =   "Medias horas"
      End
      Begin VB.Menu mnbarra4_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnSinRedondeo 
         Caption         =   "Sin redondear"
      End
   End
   Begin VB.Menu mnRevision 
      Caption         =   "Revision"
      Begin VB.Menu mnRevisarIncorrectos 
         Caption         =   "Revisar Incorrectos"
      End
   End
End
Attribute VB_Name = "frmRevision2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vFecha As String
Public Todos As Byte

'De momento private
Private idSeccion As Integer

Private WithEvents frmHoras As frmHorasMarcajes
Attribute frmHoras.VB_VarHelpID = -1
Private WithEvents frmI As frmSoloInci
Attribute frmI.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmc2 As frmCal
Attribute frmc2.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmS As frmBusca
Attribute frmS.VB_VarHelpID = -1
Private Modo As Integer
    'En este form se utilizara un modo mas, el 5
    'en el cual se mostrará el panel
    ' frame3

Dim vH As CHorarios
Dim vMar As CMarcajes
Dim SQL As String
Dim kCampo As Integer
Dim primeravez As String
Dim CadenaConsulta As String
Private kRegistro As Long
Private RegTotales As Long
Private IdI As Long
Private vIndice As Integer 'Indice para el formuario de seleccionar
Private SignoIncidencia As Single
Private CadenaPorTipo As String
Private Redondeo As Byte

Private Sub cmdAceptar_Click()
Dim Cad As String
Dim i As Integer

On Error GoTo ErrorAceptar:
 If Modo = 4 Then
        'Modificar
        ''Haremos las comprobaciones necesarias de los campos
        'Recordamos que el text(0) tiene el codigo y no lo puede cambiar
        If Not DatosOk(True) Then Exit Sub
 
        'Almacenamos para luego buscarlo
        Cad = vMar.Entrada
        'modificamos
        vMar.Fecha = txtMarcaje(1).Text
        vMar.idTrabajador = txtMarcaje(2).Text
        vMar.HorasIncid = CSng(txtMarcaje(7).Text)
        vMar.HorasTrabajadas = CSng(txtMarcaje(4).Text)
        vMar.IncFinal = CInt(txtMarcaje(5).Text)
        'vMar.Correcto = False
        If vMar.Modificar = 1 Then Exit Sub
        PonerModo 2
        'Hay que refresca el DAta1
        Adodc1.Refresh
        RegTotales = Adodc1.Recordset.RecordCount
        'Hay que volver a poner el registro donde toca
        Adodc1.Recordset.MoveFirst
        i = 1
        While i > 0
            If Adodc1.Recordset.Fields(0) = Cad Then
                i = 0
                Else
                    Adodc1.Recordset.MoveNext
                    If Adodc1.Recordset.EOF Then i = 0
            End If
        Wend
        If Adodc1.Recordset.EOF Then
            kRegistro = RegTotales
            Adodc1.Recordset.MoveLast
        End If
        '-----------------
        ' estamos añadiendo
        Else
        'If Not DatosOk Then Exit Sub
        'Si esta correcto modificamos los valores y seguimos
        If vMar.Modificar = 1 Then Exit Sub
        Adodc1.Refresh
        MsgBox "                Registro insertado.             ", vbInformation
        'Tendremos que mover el recordset hasta el
        i = 0
        Cad = ""
        Do
            If Not Adodc1.Recordset.EOF Then
                If Adodc1.Recordset.Fields(0) <> vMar.Entrada Then
                    Adodc1.Recordset.MoveNext
                    Else
                        i = 1
                End If
                Else
                    i = 1
                    Cad = "NO ESTA"
            End If
        Loop While i = 0
        If Cad <> "" Then
            'No esta segun el criterio de busqueda
            PonerModo 0
            LimpiarCampos
            Else
                PonerModo 2
                PonerTodosLosCampos
        End If
End If
ErrorAceptar:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelaBusqueda_Click()
    PonerModo 0
End Sub

Private Sub cmdCancelar_Click()
Select Case Modo
Case 3
    'Como estamos insertando
    LimpiarCampos
    PonerModo 0
Case 4
    PonerTodosLosCampos
    PonerModo 2
Case Else
    
End Select
End Sub

Private Sub cmdRevisada_Click()
Dim T1, T2
Dim Siguiente As Long


On Error GoTo ErrRevision
If Trim(Me.txtMarcaje(0).Text) = "" Then Exit Sub
Screen.MousePointer = vbHourglass
If DatosOk(True) Then
    'Para situar luego el recordset
    Siguiente = kRegistro
    If kRegistro >= RegTotales Then Siguiente = RegTotales - 1
    vMar.PonerCorrecta
    Check1.Value = 1
    Adodc1.Refresh
    '-----------------------
    espera 1
    '----------------------
    Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then
        kRegistro = 1
        RegTotales = Adodc1.Recordset.RecordCount
        'Me situo en siguiente
        While kRegistro <> Siguiente
            If Adodc1.Recordset.EOF Then
                Siguiente = kRegistro + 1
            Else
                Adodc1.Recordset.MoveNext
            End If
            kRegistro = kRegistro + 1
        Wend
        PonerTodosLosCampos
        Else
        LimpiarCampos
        MsgBox "Ningún registro por mostrar.", vbInformation
        'Unload Me
    End If
End If

ErrRevision:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MuestraError Err.Number
End Sub

Private Sub Command1_Click(Index As Integer)
Dim RT As ADODB.Recordset
Dim RC As Byte
Dim Cad As String

On Error GoTo ErrIncidencias
If Trim(Me.txtMarcaje(0).Text) = "" Then Exit Sub
If Modo = 4 Or Modo = 5 Then
    MsgBox "Esta modificando los datos de la cabecera.", vbExclamation
    Exit Sub
End If
IdI = 0
Select Case Index
Case 0
    'Nuevo
    IdI = -1
    Set frmI = New frmSoloInci
    frmI.CadInci = ""
    frmI.Inci = 0
    frmI.Horas = 0
    frmI.Nombre = txtMarcaje(3).Text
    frmI.Show vbModal
    Set frmI = Nothing
Case 1
    'Modificar
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    IdI = ListView2.SelectedItem.Tag
    Set frmI = New frmSoloInci
    frmI.CadInci = ListView2.SelectedItem.Text
    frmI.Inci = ListView2.SelectedItem.SubItems(3)
    frmI.Horas = ListView2.SelectedItem.SubItems(2)
    frmI.Nombre = txtMarcaje(3).Text
    frmI.Show vbModal
    Set frmI = Nothing
Case 2
    'Eliminar
    If ListView2.SelectedItem Is Nothing Then Exit Sub
    Cad = "Seguro que desea eliminar la incidencia: " & vbCrLf
    Cad = Cad & ListView2.SelectedItem.Text & vbCrLf
    RC = MsgBox(Cad, vbQuestion + vbYesNo)
    If RC = vbYes Then
        'Eliminamos
        Set RT = New ADODB.Recordset
        Cad = "DElete * from IncidenciasGeneradas Where id=" & ListView2.SelectedItem.Tag
        RT.Open Cad, conn, , , adCmdText
        Set RT = Nothing
        ListView2.ListItems.Remove ListView2.SelectedItem.Index
     End If
End Select
Exit Sub
ErrIncidencias:
    If Err.Number <> 0 Then MuestraError Err.Number
End Sub

Private Sub Command2_Click()
ListView1_DblClick
End Sub

Private Sub Command3_Click()
'Cancelar cuando se esta insertando
LimpiarCampos
PonerModo 2
If Not Adodc1.Recordset.EOF Then PonerTodosLosCampos
End Sub

Private Sub Command4_Click()
Dim H1 As Single
Dim h2 As Single
Dim Incid As Integer
Dim RT As ADODB.Recordset
Dim Cad As String


On Error GoTo ErrInsertando
'Comprobamos los datos
If txtMarcaje(2).Tag = "" Then
    MsgBox "Seleccione un empleado", vbExclamation
    Exit Sub
End If

If CInt(txtMarcaje(2).Tag) < 0 Then
    MsgBox "Seleccione un empleado", vbExclamation
    Exit Sub
End If

If Not IsDate(txtMarcaje(1).Text) Then
    MsgBox "Seleccione una fecha válida"
    Exit Sub
End If

'Por si acaso ha puesto horas trabajadas
If txtMarcaje(4).Text = "" Then
    H1 = 0
    Else
        If Not IsNumeric(txtMarcaje(4).Text) Then
            MsgBox "El numero de horas trabajadas tiene que ser un número.", vbExclamation
            Exit Sub
            Else
                H1 = CSng(txtMarcaje(4).Text)
        End If
End If

'Por si acaso ha puesto horas incidencias
If txtMarcaje(5).Text = "" Then
    Incid = 0
    Else
        If Not IsNumeric(txtMarcaje(5).Text) Then
            MsgBox "La incidencia tiene que ser numérica.", vbExclamation
            Exit Sub
            Else
                If CInt(txtMarcaje(5).Text) < 0 Then
                    MsgBox "La incidencia no es correcta.", vbExclamation
                    Exit Sub
                    Else
                        Incid = CInt(txtMarcaje(5).Text)
                End If
        End If
End If

If Incid > 0 Then
    If txtMarcaje(7).Text = "" Then
        h2 = 0
        Else
            If Not IsNumeric(txtMarcaje(7).Text) Then
                MsgBox "El numero de horas trabajadas tiene que ser un número.", vbExclamation
                Exit Sub
                Else
                    h2 = CSng(txtMarcaje(7).Text)
            End If
    End If
    Else
        h2 = 0
End If

If H1 + h2 > vH.TotalHoras Then
    MsgBox "La horas trabajadas y las de la incidencia exceden", vbExclamation
End If

'Comprobamos si esta de baja
If EsBajaTrabajo(CLng(txtMarcaje(2).Text)) Then
    MsgBox "El empleado " & txtMarcaje(3).Text & " ha causado baja en la empresa.", vbExclamation
    Exit Sub
End If

'Insertamos el marcaje
vMar.Fecha = txtMarcaje(1).Text
vMar.idTrabajador = txtMarcaje(2).Text
vMar.HorasIncid = h2
vMar.HorasTrabajadas = H1
vMar.IncFinal = Incid
vMar.Correcto = False
'Comprobamos, antes de insertar, que no existe un marcaje con esos valores
Set RT = New ADODB.Recordset
Cad = "Select * from Marcajes where IdTrabajador=" & vMar.idTrabajador
Cad = Cad & " AND Fecha=#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"
RT.Open Cad, conn, , , adCmdText
If Not RT.EOF Then
    MsgBox "Ya existe una entrada para ese empleado en esa fecha.", vbExclamation
    RT.Close
    Exit Sub
    Else
        'Todo ha ido bien
            RT.Close
            Set RT = Nothing
End If
vMar.Agregar
PonerModo 3   'insertando
'Ponemos los valores de las tikadas e incidencias
VerTikadas
VerIncidencias

Text1(5).Text = ""
Text1(6).Text = ""

'Como es nueva le mostraremos los horarios
ListView1_DblClick
Exit Sub
ErrInsertando:
    MsgBox "Error: " & Err.Number, vbExclamation
End Sub

Private Sub Command5_Click()
'Poner datos OK
Dim Siguiente As Long

If Adodc1.Recordset.EOF Then
    MsgBox "Ningún dato ha sido seleccionado"
    Exit Sub
End If
If vMar Is Nothing Then
    MsgBox "Ningún dato ha sido seleccionado"
    Exit Sub
End If
If vH Is Nothing Then
    MsgBox "Error al leer el horario ofcial del empleado."
    Exit Sub
End If
Screen.MousePointer = vbHourglass
vMar.HorasIncid = 0
vMar.HorasTrabajadas = vH.TotalHoras
vMar.IncFinal = 0
vMar.Correcto = True
vMar.Modificar

    'Para situar luego el recordset
    Siguiente = kRegistro
    If kRegistro = RegTotales Then Siguiente = RegTotales - 1
'refrescamos
'Dim T1, T2
    Adodc1.Refresh
'pequeña espera
    espera 1
    '----------------------
    Adodc1.Refresh
If Adodc1.Recordset.EOF Then
    LimpiarCampos
    PonerModo 0
    MsgBox "Ningun registro por mostrar."
    Else
        kRegistro = 1
        RegTotales = Adodc1.Recordset.RecordCount
        'Me situo en siguiente
        While kRegistro < Siguiente
            Adodc1.Recordset.MoveNext
            kRegistro = kRegistro + 1
        Wend
        StatusBar1.Panels(1).Text = ""
        PonerModo 2
        Desplazamiento (0)
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If primeravez Then
    primeravez = False
    
    Adodc1.ConnectionString = conn
    Adodc1.RecordSource = SQL
    Adodc1.Refresh
    
    If Not Adodc1.Recordset.EOF Then
        kRegistro = 1
        RegTotales = Adodc1.Recordset.RecordCount
        PonerModo 2
        PonerTodosLosCampos
        Else
        LimpiarCampos
        PonerModo 0
        MsgBox "Ningún registro por mostrar.", vbInformation
    End If
    Screen.MousePointer = vbDefault
End If
End Sub

Private Sub Form_Load()
Dim Nexo As String

Screen.MousePointer = vbHourglass
Frame4.Visible = False
Set vH = New CHorarios
Set vMar = New CMarcajes
'Image3.Visible = (Dir(App.Path & "\RevHco.dat", vbArchive) <> "")
kRegistro = 0
RegTotales = 0
primeravez = True
SignoIncidencia = 0
idSeccion = -1 'todas
PonerCadenaPorTipo
PonerOpcionFecha
PonerOpcionSeccion ""
mnQuitarBUSQ.Enabled = False
Command5.Enabled = (Dir(App.Path & "\HabiCmd5.cfg") <> "")
SQL = "SELECT Entrada FROM Marcajes "
If CadenaPorTipo <> "" Then
    SQL = SQL & " WHERE " & CadenaPorTipo
    Nexo = " AND "
    Else
        Nexo = "WHERE"
End If
If vFecha <> "" Then
    SQL = SQL & Nexo & " Fecha=#" & Format(vFecha, "yyyy/mm/dd") & "#"
End If

SQL = SQL & " ORDER BY Fecha,idtrabajador"

'Leemos el valor de redondeo
LeerRedondeos

'Seccion


'Situamos el form
Top = 0
Left = 0
Width = 12000
Height = 9000
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Set vH = Nothing
Set vMar = Nothing
'If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation
Err.Clear
GrabarRedondeos
End Sub


Private Sub PonerTodosLosCampos()
Dim Horario As Integer
'Leemos del adodc1
Screen.MousePointer = vbHourglass
If Adodc1.Recordset.EOF Then Exit Sub
If vMar.Leer(Adodc1.Recordset.Fields(0)) = 1 Then
    MsgBox "Error leyendo el marcaje Num: " & Adodc1.Recordset.Fields(0)
    LimpiarCampos
    GoTo ErrorPonerCampos
End If
cmdRevisada.Visible = Not vMar.Correcto
Check1.Value = Abs(vMar.Correcto)
txtMarcaje(0).Text = vMar.Entrada
txtMarcaje(1).Text = Format(vMar.Fecha, "dd/mm/yyyy")
txtMarcaje(2).Text = vMar.idTrabajador
txtMarcaje(3).Text = devuelveNombreTrabajador(vMar.idTrabajador, Horario)
txtMarcaje(4).Text = vMar.HorasTrabajadas
txtMarcaje(6).Text = DevuelveTextoIncidencia(vMar.IncFinal, SignoIncidencia)
txtMarcaje(5).Text = vMar.IncFinal
txtMarcaje(7).Text = vMar.HorasIncid

'Horario en seagesimal
txtDec(0).Text = Format(DevuelveHora(vMar.HorasTrabajadas), "hh:mm")
txtDec(1).Text = Format(DevuelveHora(vMar.HorasIncid), "hh:mm")
'Procesamos el horario oficial
PonerHorario Horario, vMar.Fecha
'Vamos las tikadas y demas
VerTikadas
'Ver Incidencias
VerIncidencias


'Finalmente ponemos el label identificando
StatusBar1.Panels(1).Text = kRegistro & " de " & RegTotales
ErrorPonerCampos:
Screen.MousePointer = vbDefault
End Sub


Private Sub PonerHorario(vHorario As Integer, vFecha)
Dim Dias As Byte
Dim Obtener As Boolean


Obtener = True
If vH.IdHorario = vHorario Then
     Dias = Weekday(vFecha, vbMonday)
    If Dias = vH.DiaSemana Then Obtener = False
End If

If Obtener Then
    If vH.Leer(vHorario, CDate(vFecha)) = 0 Then
        txtHorario(0).Text = vH.NomHorario
        txtHorario(1).Text = vH.HoraE1
        txtHorario(2).Text = vH.HoraS1
        txtHorario(3).Text = DBLet(vH.HoraE2)
        txtHorario(4).Text = DBLet(vH.HoraS2)
        txtHorario(5).Text = vH.TotalHoras
        txtHorario(6).Text = vH.NumTikadas
    End If
End If
End Sub



Private Sub VerTikadas()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim itm As ListItem
Dim Cont As Integer

ListView1.ListItems.Clear
ListView3.ListItems.Clear
Set Rs = New ADODB.Recordset

SQL = "SELECT EntradaMarcajes.Hora, Incidencias.NomInci, EntradaMarcajes.idInci,EntradaMarcajes.HoraReal"
SQL = SQL & " FROM EntradaMarcajes ,Incidencias WHERE EntradaMarcajes.idInci = Incidencias.IdInci"
SQL = SQL & " AND idMarcaje=" & vMar.Entrada
SQL = SQL & " Order by Hora"
Rs.Open SQL, conn, , , adCmdText
Cont = 0
While Not Rs.EOF
    Set itm = ListView1.ListItems.Add(, , Rs!Hora)
    If Rs!idInci = 0 Then
        itm.SubItems(1) = ""
        Else
        itm.SubItems(1) = Rs!NomInci
    End If
    itm.SmallIcon = 1
    
    
    
    
    'Insertamos el ticaje real
    Set itm = ListView3.ListItems.Add(, , Rs!HoraReal)
    itm.SmallIcon = 2
    
    Cont = Cont + 1
    Rs.MoveNext
Wend
Text1(6).Text = Cont
CalculaHoras
Rs.Close
Set Rs = Nothing
End Sub




Private Sub VerIncidencias()
Dim Rs As ADODB.Recordset
Dim SQL As String
Dim itm As ListItem


ListView2.ListItems.Clear
Set Rs = New ADODB.Recordset
SQL = "SELECT Incidencias.NomInci, Incidencias.IdInci,IncidenciasGeneradas.horas,IncidenciasGeneradas.Id"
SQL = SQL & " FROM IncidenciasGeneradas,Incidencias WHERE IncidenciasGeneradas.Incidencia = Incidencias.IdInci"
SQL = SQL & " AND EntradaMarcaje=" & vMar.Entrada
SQL = SQL & " ORDER BY Id"
Rs.Open SQL, conn, , , adCmdText
While Not Rs.EOF
    Set itm = ListView2.ListItems.Add(, , Rs!NomInci)
    itm.SubItems(1) = DevuelveHora(Rs!Horas)
    itm.SubItems(2) = Rs!Horas
    itm.SubItems(3) = Rs!idInci
    itm.SmallIcon = 3
    itm.Tag = Rs!Id
    Rs.MoveNext
Wend
Rs.Close
Set Rs = Nothing

End Sub






Private Sub PonerModo(Kmodo As Integer)
Dim i As Integer
Dim B As Boolean

Modo = Kmodo
B = Modo > 2 Or Modo = 1
Image1(0).Visible = (Modo = 3)
cmdAceptar.Visible = B
cmdCancelar.Visible = B
DespalzamientoVisible (Kmodo = 2)
textoMarcajes (B)
Image1(0).Visible = (Modo > 2)
Image1(2).Visible = (Modo > 2)
Image2.Visible = Modo > 2
Frame3.Visible = (Modo = 5)
Toolbar1.Buttons(6).Enabled = (Modo < 3)
Toolbar1.Buttons(7).Enabled = (Modo = 2)
Toolbar1.Buttons(8).Enabled = (Modo = 2)
Toolbar1.Buttons(1).Enabled = (Modo < 3)
Toolbar1.Buttons(2).Enabled = (Modo < 3)
cmdAceptar.Visible = (Modo = 3) Or (Modo = 4)
cmdCancelar.Visible = cmdAceptar.Visible
cmdRevisada.Visible = Modo = 2
Command5.Visible = Modo = 4
txtMarcaje(3).Enabled = Modo = 1
txtMarcaje(6).Enabled = Modo = 1
cmdCancelaBusqueda.Visible = Modo = 1

'Menus hbilitados o no
B = Modo = 2 Or Modo = 0
HabilitaMenus B

End Sub

Private Sub HabilitaMenus(Si As Boolean)
    
    Me.mnOpciones.Enabled = Si
    Me.mntipo.Enabled = Si
    Me.mnFecha.Enabled = Si
    Me.mnredondeo.Enabled = Si
    Me.mnRevision.Enabled = Si
End Sub


Private Sub textoMarcajes(Habilitado As Boolean)
Dim i
    For i = 0 To txtMarcaje.Count - 1
        txtMarcaje(i).Locked = Not Habilitado
    Next i
    txtDec(0).Locked = Not Habilitado
    txtDec(1).Locked = Not Habilitado
End Sub






Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
Dim Cad As String
    txtMarcaje(vIndice).Text = vCodigo
    txtMarcaje(vIndice + 1).Text = vCadena
    If vIndice = 2 Then
        txtMarcaje(vIndice).Tag = vCodigo
        'Es el empleado, luego intentamos poner los codigos
        ObtenerHorarios2
        'Else
        Else
            'Es la incidencia
            Cad = DevuelveTextoIncidencia(CInt((vCodigo)), SignoIncidencia)
    End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
    txtMarcaje(1).Text = Format(vFecha, "dd/mm/yyyy")
    ObtenerHorarios2
End Sub

Private Sub frmc2_Selec(vvfecha As Date)
    vFecha = vvfecha
    PonerOpcionFecha
End Sub

Private Sub frmHoras_HayModificacion(SiNo As Boolean, vOpcion As Byte)
Dim bol As Boolean
Dim TipoControl As Byte
Dim Fin As Boolean

Screen.MousePointer = vbHourglass
If SiNo Then
    'SI ha habido modificacion
    'Si es del HCO entonces con cargar el grid sobra
    If vOpcion = 1 Then
        'No se debe dar
    Else
        'Han modificado las horas, luego hay que repintar las horas y recalcular posibles incidencias etc
        'al igual que si todo esta correcto habra que refrescar el adodc1.recordset
        conn.BeginTrans
        TipoControl = DevuelveTipoControl(vMar.idTrabajador)
        Select Case TipoControl
            Case 2
                bol = ProcesarMarcajeRevision2(vMar)
            Case 3
                bol = ProcesarMarcajeRevision3
            Case Else
                bol = ProcesarMarcajeRevision1
        End Select
        
        If bol Then
            conn.CommitTrans
            espera 0.05
            'Si tiene produccion
                If mConfig.Kimaldi And MiEmpresa.QueEmpresa <> 1 Then
'                    ImpFechaIni = "#" & Format(vMar.Fecha, "yyyy/mm/dd") & "#"
'                    frmProcMarcajes2.ListaTrabajadores = vMar.idTrabajador & "|"
'                    frmProcMarcajes2.opcion = 1
'                    frmProcMarcajes2.Show vbModal
                End If
            
            If vMar.Correcto Then
                'Al cambiar las horas el marcaje a pasado a correcto
                'luego volvemos a refrescal el adodc1
                If Todos = 2 Then _
                    MsgBox "La entrada es ahora correcta. No se mostrará.", vbExclamation
            
                Adodc1.Refresh
                If Not Adodc1.Recordset.EOF Then
                    kRegistro = 1
                    RegTotales = Adodc1.Recordset.RecordCount
                    
                    'Movemos hasta situarnos en el siguiente al que estaba
                    Fin = False
                    While Not Fin
                        If Adodc1.Recordset.Fields(0) >= vMar.Entrada Then
                            Fin = True
                            Else
                                Adodc1.Recordset.MoveNext
                                kRegistro = kRegistro + 1
                                If Adodc1.Recordset.EOF Then
                                    Fin = True
                                    kRegistro = RegTotales
                                    Adodc1.Recordset.MoveLast
                                End If
                        End If
                    Wend
                    
                    
                    PonerTodosLosCampos
                    Else
                    LimpiarCampos
                    MsgBox "Ningún registro por mostrar.", vbInformation
                    'Unload Me
                End If
                'ELSE de marcaje correcto
                Else
                    'vemos las tikadas
                    VerTikadas
            End If
            'Volvemos a poner los datos
            If Modo = 2 Then PonerTodosLosCampos
        Else
                MsgBox "Se ha producido un error.", vbExclamation
                conn.RollbackTrans
        End If  'de correcto
    End If  'De opcion
End If
Screen.MousePointer = vbDefault
End Sub

Private Sub frmI_Seleccionar(vInci As Integer, vhoras As Single)
Dim RT As ADODB.Recordset
Dim Cad As String
Dim Valor As Long

Set RT = New ADODB.Recordset
RT.CursorType = adOpenKeyset
RT.LockType = adLockOptimistic
'Si es nueva insertamos
If IdI < 0 Then
    'Insert
    Cad = "Select * from IncidenciasGeneradas order By ID"
    RT.Open Cad, conn, , , adCmdText
    Valor = 1
    If Not RT.EOF Then
            RT.MoveLast
           If Not IsNull(RT!Id) Then Valor = RT!Id + 1
    End If
    Cad = "INSERT INTO IncidenciasGeneradas (id,EntradaMarcaje,Incidencia,Horas) VALUES "
    Cad = Cad & "(" & Valor & "," & vMar.Entrada & "," & vInci
    Cad = Cad & "," & TransformaComasPuntos(CStr(vhoras)) & ")"
    conn.Execute Cad
    Else
        'modificamos
        Cad = "Select * from IncidenciasGeneradas WHERE Id=" & IdI
        RT.Open Cad, conn, , , adCmdText
        If Not RT.EOF Then
            RT!Incidencia = vInci
            RT!Horas = vhoras
            RT.Update
        End If
End If
RT.Close
Set RT = Nothing
espera 0.5
VerIncidencias
Exit Sub
ErrSelec:
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

Private Sub frmS_Seleccion(vCodigo As Long, vCadena As String)
    SQL = vCodigo & "|" & vCadena & "|"
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmB = New frmBusca
    'Ponemos los valores para abrir
    If Index = 0 Then
        vIndice = 2
        frmB.Tabla = "Trabajadores"
        frmB.CampoBusqueda = "NomTrabajador"
        frmB.CampoCodigo = "IdTrabajador"
        frmB.TipoDatos = 3
        frmB.Titulo = "EMPLEADOS"
        Else
            vIndice = 5
            frmB.Tabla = "Incidencias"
            frmB.CampoBusqueda = "NomInci"
            frmB.CampoCodigo = "IdInci"
            frmB.TipoDatos = 3
            frmB.Titulo = "INCIDENCIAS"
    End If
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub Image2_Click()
Set frmC = New frmCal
frmC.Fecha = Now
frmC.Show vbModal
Set frmC = Nothing
End Sub

Private Sub Image3_Click()
   '''' LlamaHoras 1  'hco
End Sub

Private Sub ListView1_DblClick()
    LlamaHoras 0   'Marcajes
End Sub

Private Sub LlamaHoras(opcion As Byte)
If Trim(Me.txtMarcaje(0).Text) = "" Then Exit Sub
If Modo = 4 Or Modo = 5 Then
    MsgBox "Esta modificando los datos de la cabecera.", vbExclamation
    Exit Sub
End If
Set frmHoras = New frmHorasMarcajes
frmHoras.Nombre = txtMarcaje(3).Text
Set frmHoras.vH = vH
Set frmHoras.vM = vMar
frmHoras.opcion = opcion  'Marcajes
frmHoras.Show vbModal
Set frmHoras = Nothing
End Sub



Private Sub RefrescarDB()
Dim Nexo As String

SQL = "SELECT Entrada FROM Marcajes,Trabajadores WHERE Marcajes.idtrabajador=Trabajadores.idTrabajador "
Nexo = " AND "
If CadenaPorTipo <> "" Then SQL = SQL & Nexo & CadenaPorTipo
    

If vFecha <> "" Then SQL = SQL & Nexo & " Fecha=#" & Format(vFecha, "yyyy/mm/dd") & "#"

'Por seccion,
If idSeccion >= 0 Then SQL = SQL & Nexo & " Seccion = " & idSeccion


SQL = SQL & " ORDER BY "
If Me.mnPorFecha.Checked Then
    SQL = SQL & "Fecha,Marcajes.IdTrabajador"
Else
    SQL = SQL & "Marcajes.idTrabajador,Fecha"
End If
Adodc1.RecordSource = SQL
Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then
        PonerModo 2
        kRegistro = 1
        RegTotales = Adodc1.Recordset.RecordCount
        PonerTodosLosCampos
        Else
        LimpiarCampos
        MsgBox "Ningún registro por mostrar.", vbInformation
        'Unload Me
    End If
End Sub




Private Sub mnCorrectas_Click()
    Todos = 1
    PonerCadenaPorTipo
    RefrescarDB
End Sub

Private Sub mnCualquiera_Click()
    vFecha = ""
    PonerOpcionFecha
    RefrescarDB
End Sub

Private Sub mncuartos_Click()
    ClickRedondeos 2
End Sub

Private Sub mnDecima_Click()
    ClickRedondeos (1)
End Sub

Private Sub mnEliminar_Click()
    BotonEliminar
End Sub

Private Sub mnIncorrectas_Click()
    Todos = 2
    PonerCadenaPorTipo
    RefrescarDB
End Sub

Private Sub mnKfecha_Click()
Set frmc2 = New frmCal
If vFecha <> "" Then
    If IsDate(vFecha) Then
        frmc2.Fecha = CDate(vFecha)
    End If
    Else
        frmc2.Fecha = Now
End If
frmc2.Show vbModal
Set frmc2 = Nothing
PonerOpcionFecha
RefrescarDB
End Sub

Private Sub mnMediaHora_Click()
    ClickRedondeos 3
End Sub

Private Sub mnModificar_Click()
BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub Ordenacion(PorFecha As Boolean)
mnPorFecha.Checked = PorFecha
mnTrabajador.Checked = Not PorFecha
PonerCadenaPorTipo
RefrescarDB
End Sub

Private Sub mnPorFecha_Click()
Ordenacion True
End Sub

Private Sub mnQuitarBUSQ_Click()
RefrescarDB
End Sub

Private Sub mnRevisarIncorrectos_Click()
    SQL = "Se prodece a la revision masiva de marcajes" & vbCrLf & Space(25) & "¿Desea continuar?"
    If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    StatusBar1.Panels(1).Text = "Revisando incorrectos"
    Label15.Caption = "Revisión"
    Me.Refresh
    RevisionGeneralIncorectos
End Sub

Private Sub mnSalir_Click()
Unload Me
End Sub

Private Sub mnSeccion1_Click()
    Set frmS = New frmBusca
    SQL = ""
    'Ponemos los valores para abrir
    
    frmS.Tabla = "Secciones"
    frmS.CampoBusqueda = "Nombre"
    frmS.CampoCodigo = "Idseccion"
    frmS.TipoDatos = 3
    frmS.Titulo = "Secciones"
    frmS.MostrarDeSalida = True
    frmS.Show vbModal
    Set frmS = Nothing
    If SQL <> "" Then
        idSeccion = RecuperaValor(SQL, 1)
        SQL = RecuperaValor(SQL, 2)
        PonerOpcionSeccion SQL
        RefrescarDB
        SQL = ""
    End If
End Sub

Private Sub mnSeccionT_Click()
    idSeccion = -1
    PonerOpcionSeccion ""
    RefrescarDB
End Sub

Private Sub mnSinRedondeo_Click()
    ClickRedondeos 0
End Sub

Private Sub mnTodas_Click()
Todos = 0
PonerCadenaPorTipo
RefrescarDB
End Sub

Private Sub mnTrabajador_Click()
'    mnTrabajador.Checked = Not mnTrabajador.Checked
'    Me.mnPorFecha.Checked = Not mnTrabajador.Checked
    Ordenacion False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    BotonBuscar
Case 2
   ' BotonVerTodos
Case 6
    BotonAnyadir
Case 7
    BotonModificar
Case 8
    BotonEliminar
Case 14 To 17
    Desplazamiento (Button.Index - 14)
Case Else

End Select
End Sub


Private Sub Desplazamiento(Index As Integer)
If Adodc1.Recordset.EOF Then Exit Sub
Select Case Index
    Case 0
        Adodc1.Recordset.MoveFirst
        kRegistro = 1
    Case 1
        Adodc1.Recordset.MovePrevious
        kRegistro = kRegistro - 1
        If Adodc1.Recordset.BOF Then
            Adodc1.Recordset.MoveFirst
            kRegistro = 1
        End If
    Case 2
        Adodc1.Recordset.MoveNext
        kRegistro = kRegistro + 1
        If Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveLast
            kRegistro = RegTotales
        End If
    Case 3
        Adodc1.Recordset.MoveLast
        kRegistro = RegTotales
End Select
PonerTodosLosCampos
End Sub

Private Sub CalculaHoras()
Dim g As Integer
Dim i As Integer
Dim Horas As Single
Dim V As Single

g = ListView1.ListItems.Count
If g = 0 Then
    Text1(5).Text = 0
    Text1(0).Text = 0
    Exit Sub
End If
g = g \ 2
Horas = 0
For i = 1 To g
    V = DevuelveValorHora(CDate(ListView1.ListItems((i * 2))) - CDate(ListView1.ListItems((i * 2) - 1)))
    Horas = Horas + V
Next i
Text1(5).Text = Round(Horas, 2)
Text1(0).Text = DevuelveHora(Horas)
End Sub


Private Function ProcesarMarcajeRevision1() As Boolean
Dim Cad As String
Dim Rss As ADODB.Recordset
Dim vE As CEmpresas
Dim InciManual As Integer
Dim T1 As Single
Dim T2 As Single
Dim Valor As Long
Dim NumTikadas As Integer
Dim HoraH As Date
Dim Exceso As Date
Dim Retraso As Date
Dim i As Integer
Dim kIncidencia As Single
Dim V(3) As Single
Dim TieneIncidencia As Boolean
Dim TotalH As Currency
Dim N As Integer

On Error GoTo ErroProcesaMarcaje22
ProcesarMarcajeRevision1 = False
Set Rss = New ADODB.Recordset
Valor = -1
Cad = "select IdEmpresa from Trabajadores WHERE IdTrabajador=" & vMar.idTrabajador
Rss.Open Cad, conn, , , adCmdText
If Not Rss.EOF Then
    If Not IsNull(Rss.Fields(0)) Then Valor = Rss.Fields(0)
End If
Rss.Close

Set vE = New CEmpresas
If vE.Leer(Valor) = 1 Then
    MsgBox "Error leyendo la empresa para el trabajador." & vbCrLf & " Cod. Empresa: " & Valor, vbExclamation
    Set vE = Nothing
    Exit Function
End If

'Borramos todas las incidencias, si tiene, generadas autmoaticamente
Cad = "DELETE * FROM IncidenciasGeneradas WHERE EntradaMarcaje=" & vMar.Entrada
Cad = Cad & " AND Incidencia<>" & vE.IncMarcaje
Rss.Open Cad, conn, , , adCmdText
'No hace falta cerrarlo puesto que no devuelve ningun registro

Cad = "Select * from EntradaMarcajes WHERE IdMarcaje=" & vMar.Entrada
Cad = Cad & " ORDER BY Hora"
Rss.CursorType = adOpenStatic
Rss.Open Cad, conn, , , adCmdText

If Rss.EOF Then
    'Si no hay ninguna entrada
    Set Rss = Nothing
    Exit Function
End If

InciManual = 0
NumTikadas = Rss.RecordCount
If vH.EsDiaFestivo Then

   ' y todo pasa a ser horas extras
    If (NumTikadas Mod 2) > 0 Then
        'Numero de marcajes impares. No podemos calcular horas
        'trabajadas. Generamos error en marcaje
        vMar.IncFinal = vE.IncMarcaje
        vMar.HorasIncid = 0
        vMar.HorasTrabajadas = 0
        GeneraIncidencia vE.IncMarcaje, vMar.Entrada, 0
        Else
            N = NumTikadas \ 2
            TotalH = 0
            'NUMERO DE MARCAJES PAR
            Rss.MoveFirst
            For i = 1 To N
                T1 = DevuelveValorHora(Rss!Hora)
                Rss.MoveNext
                T2 = DevuelveValorHora(Rss!Hora)
                Rss.MoveNext
                TotalH = TotalH + (T2 - T1)
            Next i
            'Ahora ya sabemos las horas trabajadas
            TotalH = Round(TotalH, 2)
            
            '---------------------------------------------
            'Vemos si hay k kitar el almuerzo
            If QuitarAlmuerzo(vMar, vH) Then TotalH = TotalH - vH.DtoAlm
            

             vMar.HorasTrabajadas = TotalH
             vMar.HorasIncid = TotalH
             vMar.IncFinal = vE.IncHoraExtra
    End If  'de NUMTIKADAS es numero par





    Else
        '---------------
        'ELSE
        '--> NO es dia festivo
        If NumTikadas = vH.NumTikadas Then
            'Ha ticado las mismas veces que le correspondian
            'Comprobamos si ha habido algun retraso, o exceso
            Exceso = DevuelveHora(vE.MaxExceso)
            Retraso = DevuelveHora(vE.MaxRetraso)
            i = 0
            PrimerTicaje = Rss!Hora
            While Not Rss.EOF
                If Rss!idInci > 0 Then InciManual = Rss!idInci
                Select Case i
                Case 0
                    HoraH = vH.HoraE1
                Case 1
                    HoraH = vH.HoraS1
                Case 2
                    HoraH = vH.HoraE2
                Case 3
                    HoraH = vH.HoraS2
                End Select
                kIncidencia = EntraDentro(Rss!Hora, HoraH, Exceso, Retraso, (i Mod 2) = 0)
                V(i) = kIncidencia
                i = i + 1
                UltimoTicaje = Rss!Hora
                Rss.MoveNext
            Wend
            
            'Ahora ya tenmos si ha llegado tarde, ha salido antes etc, por lo tanto
            ' realizamos los calculos de las horas y genereamos, si cabe
            'las incidencias
            'En v() tenemos que si es 0 nada, pero si es menor tenemos la horas extras
            ' y si es mayor las horas de retraso
            'En t1 tendremos las horas en las incidencias
            T1 = 0
            TieneIncidencia = False
            For i = 0 To 3
                T1 = T1 + V(i)
                If V(i) > 0 Then
                  '  GeneraIncidencia vE.IncRetraso, vMar.Entrada, v(I)
                    TieneIncidencia = True
                    Else
                        If V(i) < 0 Then
                   '         GeneraIncidencia vE.IncHoraExceso, vMar.Entrada, Abs(v(I))
                            TieneIncidencia = True
                        End If
                End If
            Next i
            'Contabilizaremos los descuentos relativos al almuerzo y merienda
            'si tiene dto. Le sumaremos al valor obtenido en T1 el valor de los dtos
           'Comprobamos los dtos almuerzo merienda
            '----------------------------------------------
            'Modificacion septiempre 04. Comentamos. No se descuenta a nivel de horas
            T2 = vH.TotalHoras
            'T2 = vH.TotalHoras - vH.DtoMer
            'If QuitarAlmuerzo(vMar, vH) Then T2 = T2 - vH.DtoAlm
            
            If vH.DtoAlm > 0 Then
                If Not LeQuitamosElAmluerzo(Rss, vH) Then
                    If T1 < 0 Then
                        'HORAS EXTRA
                        T1 = T1 - vH.DtoAlm
                    Else
                        'Si no llega pues no hacemos nada
                        T1 = T1 - vH.DtoAlm
                    End If
                End If
            End If
            'Una vez asignadas calculamos las horas que le corresponden
            'Segun los horarios
            
            T2 = Round(T2 - T1, 2)
            vMar.HorasTrabajadas = T2
            
            
            
            
            
            
            'Asignaremos la incidencia
            'Si tiene manual se queda la manual, si no se queda, si tuviera, la automatica
            If InciManual > 0 Then
                vMar.IncFinal = InciManual
                vMar.HorasIncid = Round(vH.TotalHoras - vMar.HorasTrabajadas, 2)
                'Generamos la incidencia manual
                GeneraIncidencia InciManual, vMar.Entrada, vMar.HorasIncid
                Else
                    'Vemos si tiene automatica
                    If T1 = 0 Then
                        If TieneIncidencia Then
                            'La suma de horas da 0, pero tiene incidencias
                            vMar.IncFinal = vE.IncMarcaje
                            Else
                                vMar.IncFinal = 0
                        End If
                        Else
                            'Falta o sobran horas
                            If T1 > 0 Then
                                'Retraso
                                vMar.IncFinal = vE.IncRetraso
                                Else
                                    vMar.IncFinal = vE.IncHoraExtra
                            End If
                            vMar.HorasIncid = Round(Abs(T1), 2)
                    End If 't2=0
            End If
            
            
        '   El numero de tikadas no coincide
        Else
            While Not Rss.EOF
                 If Rss!idInci > 0 Then InciManual = Rss!idInci
                 Rss.MoveNext
            Wend
            If InciManual > 0 Then
                vMar.IncFinal = InciManual
               ' GeneraIncidencia InciManual, vMar.Entrada, 0
                Else
                    vMar.IncFinal = vE.IncMarcaje
               '     GeneraIncidencia vE.IncMarcaje, vMar.Entrada, 0
            End If
            
            'Si pares
            'Ahora pondremos las horas trabajadas por diferencias
            Rss.MoveFirst
            kIncidencia = 0
            If (Rss.RecordCount Mod 2) = 0 Then
                While Not Rss.EOF
                    'Son pares
                    T1 = DevuelveValorHora(Rss!Hora)
                    'Siguiente
                    Rss.MoveNext
                    T2 = DevuelveValorHora(Rss!Hora)
                    
                    T2 = T2 - T1
                    T2 = Round(T2, 2)
                    kIncidencia = kIncidencia + T2
                    'siguiente par
                    Rss.MoveNext
                Wend
                kIncidencia = Round(kIncidencia, 2)
            End If
            T1 = 0
            'Deberia haber trabajado
            If kIncidencia > 0 Then
                T2 = vH.TotalHoras - vH.DtoMer
                If QuitarAlmuerzo(vMar, vH) Then
                    T2 = T2 - vH.DtoAlm
                    T1 = vH.DtoAlm
                End If
                kIncidencia = kIncidencia - vH.DtoMer - T1
                T1 = kIncidencia - T2
                T1 = Abs(T1)
            End If
            vMar.HorasTrabajadas = kIncidencia
            vMar.HorasIncid = T1
        
        End If   'de incimanual>0
End If ''De es festivo

'Por ultimo marcamos o no el campo correcto
vMar.Correcto = vMar.IncFinal = 0

'Comprobamos si esta de baja
If EsBajaTrabajo(vMar.idTrabajador) Then
    vMar.Correcto = False
    If vMar.IncFinal <> vE.IncMarcaje Then vMar.IncFinal = vE.IncVacaciones
End If

'Grabamos el marcaje
vMar.Modificar
ProcesarMarcajeRevision1 = True
Set vE = Nothing
Exit Function
'Salimos
ErroProcesaMarcaje22:
MsgBox "Error: " & Err.Description, vbExclamation
ProcesarMarcajeRevision1 = False
If Not vE Is Nothing Then Set vE = Nothing
End Function



Private Sub txtDec_GotFocus(Index As Integer)
txtDec(Index).SelStart = 0
txtDec(Index).SelLength = Len(txtDec(Index).Text)
End Sub

Private Sub txtDec_LostFocus(Index As Integer)
Dim i As Integer
'Por si acaso escribe aqui las horas
'Primero comprobamos si ha introducido con puntros en lugar de con dos puntos
If Modo < 4 Then Exit Sub
Do
    i = InStr(1, txtDec(Index).Text, ".")
    If i > 0 Then _
        txtDec(Index).Text = Mid(txtDec(Index).Text, 1, i - 1) & ":" & Mid(txtDec(Index).Text, i + 1)
Loop Until i = 0

'Una vez quitados los puntos comprobamos que es un hora. Formateamos el valor y ponemos su correspondiente para
'la sexagesimal

If Not IsDate(txtDec(Index).Text) Then
    txtDec(Index).Text = ""
    Exit Sub
End If

txtDec(Index).Text = Format(txtDec(Index).Text, "h:mm")
If Index = 0 Then
    txtMarcaje(4).Text = DevuelveValorHora(CDate(txtDec(Index).Text))
    Else
        txtMarcaje(7).Text = DevuelveValorHora(CDate(txtDec(Index).Text))
End If
End Sub

Private Sub txtMarcaje_GotFocus(Index As Integer)
kCampo = Index
If Modo = 1 Then
    txtMarcaje(Index).BackColor = vbYellow
    Else
        txtMarcaje(Index).SelStart = 0
        txtMarcaje(Index).SelLength = Len(txtMarcaje(Index))
End If
End Sub

Private Sub txtMarcaje_KeyPress(Index As Integer, KeyAscii As Integer)
If Modo = 1 Then
    If KeyAscii = 13 Then
        'Ha pulsado enter, luego tenemos que hacer la busqueda
        txtMarcaje(Index).BackColor = vbWhite
        BotonBuscar
    End If
End If
End Sub

Private Sub txtMarcaje_LostFocus(Index As Integer)
Dim i As Integer
'Pierde el enfocque
txtMarcaje(Index).BackColor = vbWhite
Select Case Index
Case 1
    If IsDate(txtMarcaje(1).Text) Then
        txtMarcaje(1).Text = Format(txtMarcaje(1).Text, "dd/mm/yyyy")
        ObtenerHorarios2
        Else
            txtMarcaje(1).Text = ""
    End If
Case 2
    ObtenerTrabajador
    ObtenerHorarios2
Case 4, 7
    i = InStr(1, txtMarcaje(Index).Text, ".")
    If i > 0 Then
        If i = 1 Then
            txtMarcaje(Index).Text = "0," & Mid(txtMarcaje(Index).Text, 2)
            Else
                txtMarcaje(Index).Text = Mid(txtMarcaje(Index).Text, 1, i - 1) & "," & Mid(txtMarcaje(Index).Text, i + 1)
        End If
    End If
    If Index = 4 Then
        i = 0
        Else
            i = 1
    End If
    If IsNumeric(txtMarcaje(Index).Text) Then
        txtDec(i).Text = Format(DevuelveHora(CSng(txtMarcaje(Index).Text)), "hh:mm")
        Else
            txtDec(i).Text = ""
    End If
Case 5
    If txtMarcaje(5).Text = "" Then Exit Sub
    If Not IsNumeric(txtMarcaje(5).Text) Then
        txtMarcaje(5).Text = -1
        txtMarcaje(6).Text = " Incidencia erronea."
        Exit Sub
    End If
    txtMarcaje(6).Text = DevuelveTextoIncidencia(CInt((txtMarcaje(5).Text)), SignoIncidencia)
End Select
End Sub


Private Function DatosOk(MostrarMsgbox As Boolean) As Boolean
Dim H1 As Single
Dim h2 As Single
Dim P1 As Single
Dim P2 As Single
Dim TipoDeControl As Byte
Dim Cad As String

DatosOk = False
txtMarcaje(4).Text = Trim(txtMarcaje(4).Text)
If txtMarcaje(4).Text = "" Then
    If MostrarMsgbox Then _
    MsgBox "Escriba las horas trabajadas.", vbExclamation
    Exit Function
End If
If Not IsNumeric(txtMarcaje(4).Text) Then
    If MostrarMsgbox Then _
    MsgBox "Las horas trabajadas tiene que ser numéricas.", vbExclamation
    Exit Function
End If
H1 = CSng(txtMarcaje(4).Text)
'Vamos con las horas extras
txtMarcaje(7).Text = Trim(txtMarcaje(7).Text)
If txtMarcaje(7).Text = "" Then
    h2 = 0
    Else
        If Not IsNumeric(txtMarcaje(7).Text) Then
            MsgBox "Las horas trabajadas tiene que ser numéricas.", vbExclamation
            Exit Function
        End If
        h2 = CSng(txtMarcaje(7).Text)
End If

'Por si acaso no tenemos marcajes
If Text1(6).Text = "" Then Text1(6).Text = 0
If Not IsNumeric(CInt(Text1(6).Text)) Then Text1(6).Text = 0
If CInt(Text1(6).Text) = 0 Then
    If MostrarMsgbox Then _
    MsgBox "No tiene realizado ningún marcaje.", vbExclamation
    Exit Function
End If

'Llegados a este punto el control de horas se hace si el tipo de
'control seleccionado para el trabajador lo requiere
'  0,1,2    SI
'     3     NO
'LEEmoes el timpo de control
TipoDeControl = DevuelveTipoControl(CLng(txtMarcaje(2).Text))

'-----------------------------------------------------------
'Contro  del numero de tikajes en funcion de tipo de control
If TipoDeControl <= 1 Then

    If (CInt(Text1(6).Text) Mod 2) <> 0 Then
        If MostrarMsgbox Then _
        MsgBox "Error en el numero de tikajes.", vbExclamation
        Exit Function
    Else
        If CInt(txtHorario(6).Text) <> CInt(Text1(6).Text) Then
            Cad = "Error en el numero de tikajes. ¿Desea continuar igualmente?"
            If Not MostrarMsgbox Then
                Exit Function
            Else
                If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Function
                'Ponemos como horas trabajadas y incidencia las k corresponda
            End If
        End If
    End If
End If

'-----------------------------------------------
'Contro  de horas  en funcion de tipo de control
' De momento solo el tipo tres no controla esto

If TipoDeControl <> 3 Then
    h2 = Round(SignoIncidencia * h2, 2)
    If vH.EsDiaFestivo Then
        If Round(H1 + h2, 2) = 0 Then
        
        Else
            If MostrarMsgbox Then _
            MsgBox "Es dia festivo. Todas las horas trabajadas son extras.", vbExclamation
            Exit Function
        End If
    Else
        'If TipoDeControl = 2 Then
            P1 = vH.TotalHoras
            If QuitarAlmuerzo(vMar, vH) Then
                'Modificacion SEPT 04. Comentamos la linea de abajo
                'P1 = P1 - vH.DtoAlm
                
            End If
            If Round(H1 + h2, 2) <> P1 Then
                If MostrarMsgbox Then _
                MsgBox "Error en el computo de horas trabajadas.", vbExclamation
                Exit Function
            End If
        'End If
    End If
End If

'--------------------------------------------------------
'Vemos la horas redondeando pero segun el tipo de control
If TipoDeControl = 3 Then
    'Conversion sencilla
    'Horas trabajadas
    H1 = RealizaRedondeo(CSng(txtMarcaje(4)), Redondeo)
    txtMarcaje(4).Text = H1
    'Horas incidencias
    h2 = RealizaRedondeo(CSng(txtMarcaje(7)), Redondeo)
    H1 = Round(H1, 2)
    h2 = Round(h2, 2)
    vMar.HorasIncid = h2
    vMar.HorasTrabajadas = H1
    txtMarcaje(4).Text = H1
    txtMarcaje(7).Text = h2
    txtMarcaje(4).Refresh
    txtMarcaje(7).Refresh
    txtDec(0).Text = Format(DevuelveHora(H1), "hh:mm")
    txtDec(0).Refresh
    txtDec(1).Text = Format(DevuelveHora(h2), "hh:mm")
    txtDec(1).Refresh
    Else
        'Controla tb las horas totales
        '-----------------------------
        P1 = H1
        P2 = h2
        ConversionRedondeo P1, P2, True
        If P2 > 0 Then
            'Significa k tiene incidencia
        If P2 > 0 And Val(txtMarcaje(5).Text) = 0 Then
            'Tiene horas incidencia pero no tiene incidencia
            If MostrarMsgbox Then _
            MsgBox "Tiene horas de incidencia pero no tiene asignada incidencia.", vbExclamation
            Exit Function
        End If
        If P2 = 0 And Val(txtMarcaje(5).Text) > 0 Then
            'Tiene horas incidencia pero no tiene incidencia
            If MostrarMsgbox Then _
            MsgBox "No tiene horas de incidencia pero tiene asignada una incidencia.", vbExclamation
            Exit Function
        End If
         
    End If
        
    ConversionRedondeo H1, h2, False
End If

If h2 > 0 And Val(txtMarcaje(5).Text) = 0 Then
    'Tiene horas incidencia pero no tiene incidencia
    If MostrarMsgbox Then _
    MsgBox "Tiene horas de incidencia pero no tiene asignada incidencia.", vbExclamation
    Exit Function
End If
If h2 = 0 And Val(txtMarcaje(5).Text) > 0 Then
    'Tiene horas incidencia pero no tiene incidencia
    If MostrarMsgbox Then _
    MsgBox "No tiene horas de incidencia pero tiene asignada una incidencia.", vbExclamation
    Exit Function
End If


DatosOk = True
End Function



Private Sub BotonAnyadir()
LimpiarCampos
If Not (vMar Is Nothing) Then Set vMar = Nothing
Set vMar = New CMarcajes
txtMarcaje(0).Text = vMar.Siguiente
SignoIncidencia = 0
PonerModo 5
Me.StatusBar1.Panels(1).Text = "NUEVO MARCAJE"
End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim RT As ADODB.Recordset
If vMar Is Nothing Then Exit Sub
If txtMarcaje(2).Text = "" Then Exit Sub
If Not IsNumeric(txtMarcaje(2).Text) Then Exit Sub
If txtMarcaje(2).Text <> vMar.idTrabajador Then
    MsgBox "Seleccione un trabajador.", vbExclamation
    Exit Sub
End If
Cad = "¿Seguro que desea eliminar el marcaje nº: " & vMar.Entrada & vbCrLf
Cad = Cad & " para el trabajador " & txtMarcaje(3).Text
Cad = Cad & " con fecha : " & vMar.Fecha & " ?"

If MsgBox(Cad, vbQuestion + vbYesNo) = vbYes Then
    Set RT = New ADODB.Recordset
    conn.BeginTrans
    Cad = "Delete * from EntradaMarcajes WHERE idMarcaje=" & vMar.Entrada
    RT.Open Cad, conn, , , adCmdText
    'rt.Close
    Cad = "Delete * from IncidenciasGeneradas WHERE EntradaMarcaje=" & vMar.Entrada
    RT.Open Cad, conn, , , adCmdText
    'rt.Close
    If vMar.Eliminar = 0 Then
        conn.CommitTrans
        Else
            conn.RollbackTrans
    End If
    Set RT = Nothing
    'Refrescamos
    Adodc1.Refresh
    If Not Adodc1.Recordset.EOF Then
        kRegistro = 1
        RegTotales = Adodc1.Recordset.RecordCount
        PonerTodosLosCampos
        Else
        LimpiarCampos
        MsgBox "Ningún registro por mostrar.", vbInformation
        'Unload Me
    End If
End If
    
End Sub

Private Sub ObtenerTrabajador()
Dim RT As ADODB.Recordset


If txtMarcaje(2).Text = "" Then Exit Sub
If Not IsNumeric(txtMarcaje(2).Text) Then
    txtMarcaje(2).Text = -1
    txtMarcaje(3).Text = "Error en el empleado"
End If

Screen.MousePointer = vbHourglass
Set RT = New ADODB.Recordset
RT.Open "Select NomTrabajador,IdHorario from Trabajadores where IdTrabajador=" & txtMarcaje(2).Text, conn, , , adCmdText
If Not RT.EOF Then
    If Not IsNull(RT.Fields(0)) Then
        txtMarcaje(3).Text = RT.Fields(0)
        txtMarcaje(2).Tag = RT.Fields(1)
        Else
            txtMarcaje(2).Text = -1
            txtMarcaje(2).Tag = -1
            txtMarcaje(3).Text = "Error en el empleado:"
    End If
    Else 'EOF
        txtMarcaje(2).Text = -1
        txtMarcaje(2).Tag = -1
        txtMarcaje(3).Text = "Error en el empleado:"
End If
RT.Close
Set RT = Nothing
Screen.MousePointer = vbDefault
End Sub

Private Sub ObtenerHorarios2()
Dim Horario As Integer
If txtMarcaje(1).Text = "" Then Exit Sub
If txtMarcaje(2).Text = "" Then Exit Sub
If Not IsDate(txtMarcaje(1).Text) Then Exit Sub
If Not IsNumeric(txtMarcaje(2).Text) Then Exit Sub
If CLng(txtMarcaje(2).Text) < 0 Then Exit Sub
If Not IsNumeric(txtMarcaje(2).Tag) Then Exit Sub
If Not (vH Is Nothing) Then Set vH = Nothing
Set vH = New CHorarios

SQL = DevuelveDesdeBD("idHorario", "trabajadores", "idtrabajador", CStr(txtMarcaje(2).Tag))

If SQL = "" Then
    MsgBox "No tienen horario asignado.  DEBE ASIGNARSELO", vbExclamation
    SQL = "0"
End If
PonerHorario CInt(SQL), txtMarcaje(1).Text
End Sub

Private Sub LimpiarCampos()
Dim T As TextBox

For Each T In txtMarcaje
    T.Text = ""
Next

For Each T In txtHorario
    T.Text = ""
Next

For Each T In Text1
    T.Text = ""
Next
txtDec(0).Text = ""
txtDec(1).Text = ""
ListView1.ListItems.Clear
ListView2.ListItems.Clear
ListView3.ListItems.Clear
Check1.Value = 0
StatusBar1.Panels(1).Text = ""
End Sub

Private Sub PonerCadenaPorTipo()
mnCorrectas.Checked = (Todos = 1)
mnIncorrectas.Checked = (Todos = 2)
mnTodas.Checked = (Todos = 0)
Select Case Todos
Case 1
    'Solo los correctos
    CadenaPorTipo = " Correcto = True"
Case 2
    'Solo los incorrectos
    CadenaPorTipo = " Correcto = False"
Case Else
    'TOdos
    CadenaPorTipo = ""
End Select
End Sub


Private Sub PonerOpcionFecha()
mnCualquiera.Checked = (vFecha = "")
mnKfecha.Checked = (vFecha <> "")
If vFecha <> "" Then _
    mnKfecha.Caption = "Fecha: " & Format(vFecha, "dd/mm/yyyy")
End Sub


Private Sub PonerOpcionSeccion(Texto As String)
mnSeccionT.Checked = (idSeccion < 0)
mnSeccion1.Checked = (idSeccion > 0)
If idSeccion >= 0 Then Texto = idSeccion & " - " & Texto
mnSeccion1.Caption = "Seccion: " & Texto
End Sub



Private Sub BotonModificar()
'---------
'MODIFICAR
'----------
'Añadiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Modificar"
PonerModo 4
'Escondemos el navegador y ponemos insertando
'Como el campo 1 es clave primaria, NO se puede modificar
StatusBar1.Panels(1).Text = "Modificar"
'Ponemos el foco sobre el nombre
txtMarcaje(1).SetFocus
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
Dim i
For i = 14 To 17
    Toolbar1.Buttons(i).Visible = bol
Next i
End Sub

Private Sub BotonBuscar()
If Modo <> 1 Then
    LimpiarCampos
    PonerModo 1
    txtMarcaje(1).SetFocus
    Else
        HacerBusqueda
        If RegTotales = 0 Then
            txtMarcaje(kCampo).Text = ""
            txtMarcaje(kCampo).BackColor = vbYellow
            txtMarcaje(kCampo).SetFocus
            Else
                txtMarcaje(kCampo).BackColor = vbWhite
        End If
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim c1 As String   'el nombre del campo
Dim Tipo As Long
Dim aux1
Dim RT As ADODB.Recordset

On Error GoTo EHacerBusqueda

Select Case kCampo
Case 0
    'Secuencia
    If Not IsNumeric(txtMarcaje(0).Text) Then
        MsgBox "La secuencia debe de ser numérica.", vbExclamation
        Exit Sub
    End If
    CadB = "Select Entrada from Marcajes where secuencia=" & txtMarcaje(0).Text
Case 1
    If Not IsDate(txtMarcaje(1).Text) Then
        MsgBox "La fecha no es correcta.", vbExclamation
        Exit Sub
    End If
    CadB = "Select Entrada from Marcajes where fecha=#" & Format(txtMarcaje(1).Text, "yyyy/mm/dd") & "#"
Case 2
    If Not IsNumeric(txtMarcaje(2).Text) Then
        MsgBox "El código del trabajador debe de ser numérico.", vbExclamation
        Exit Sub
    End If
    CadB = "Select Entrada from Marcajes where IdTrabajador=" & txtMarcaje(2).Text
Case 3
    If txtMarcaje(3).Text = "" Then
        MsgBox "Escriba algun carácter del nombre de trabajador.", vbExclamation
        Exit Sub
    End If
    CadB = "Select Entrada from Marcajes,Trabajadores where "
    CadB = CadB & " Trabajadores.idTrabajador=marcajes.IdTrabajador AND "
    CadB = CadB & " Trabajadores.NomTrabajador like '*" & txtMarcaje(3).Text & "*'"
Case 4
    If IsNumeric(txtMarcaje(4).Text) Then
        txtMarcaje(4).Text = " = " & txtMarcaje(4).Text
    End If
    CadB = "Select Entrada from Marcajes where HorasTrabajadas " & txtMarcaje(4).Text
Case 5
    If Not IsNumeric(txtMarcaje(5).Text) Then
        MsgBox "El código de la incidencia debe de ser numérico.", vbExclamation
        Exit Sub
    End If
    CadB = "Select Entrada from Marcajes where IncFinal=" & txtMarcaje(5).Text
Case 6
    If txtMarcaje(6).Text = "" Then
        MsgBox "Escriba algun carácter de la incidencia.", vbExclamation
        Exit Sub
    End If
    CadB = "Select Entrada from Marcajes,Incidencias where "
    CadB = CadB & " marcajes.incFinal = Incidencias.IdInci AND "
    CadB = CadB & " Incidencias.NomInci like '*" & txtMarcaje(6).Text & "*'"
Case Else
    CadB = ""
End Select
If CadB = "" Then Exit Sub


If CadenaPorTipo <> "" Then
    CadB = CadB & " AND " & CadenaPorTipo
End If
If vFecha <> "" Then
    CadB = CadB & " AND " & " Fecha=#" & Format(vFecha, "yyyy/mm/dd") & "#"
End If
'ORDENACION
If Me.mnPorFecha.Checked Then
    CadB = CadB & " ORDER BY FECHA,idtrabajador"
Else
    CadB = CadB & " ORDER BY idTrabajador,fecha"
End If

Adodc1.RecordSource = CadB
Adodc1.Refresh
If Adodc1.Recordset.EOF Then
    MsgBox "No hay ningún registro con esos valores.", vbInformation
    Screen.MousePointer = vbDefault
    RegTotales = 0
    Exit Sub
    'StatusBar1.Panels(2).Text = ""
    'PonerModo 0
    Else
        DespalzamientoVisible True
        PonerModo 2
        'adodc1.Recordset.MoveLast
        Adodc1.Recordset.MoveFirst
        RegTotales = Adodc1.Recordset.RecordCount
        kRegistro = 1
        PonerTodosLosCampos
End If
Exit Sub
EHacerBusqueda:
    MuestraError Err.Number, "Hacer busqueda"
    Set RT = Nothing
    Limpiar Me
    Set Adodc1.Recordset = Nothing
    PonerModo 0
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass

Adodc1.RecordSource = CadenaConsulta
Adodc1.Refresh
If Adodc1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro con esos valores.", vbInformation
    Screen.MousePointer = vbDefault
    RegTotales = 0
    Exit Sub
    'StatusBar1.Panels(2).Text = ""
    'PonerModo 0
    Else
        DespalzamientoVisible True
        PonerModo 2
        'adodc1.Recordset.MoveLast
        Adodc1.Recordset.MoveFirst
        RegTotales = Adodc1.Recordset.RecordCount
        kRegistro = 1
        PonerTodosLosCampos
End If

Adodc1.ConnectionString = conn
Adodc1.RecordSource = CadenaConsulta
Adodc1.Refresh
RegTotales = Adodc1.Recordset.RecordCount
Screen.MousePointer = vbDefault
End Sub


Private Sub LeerRedondeos()
Dim NF As Integer
Dim Cad As String

On Error GoTo FinLeerRedondeos:
Redondeo = 0
NF = FreeFile
Cad = Dir(App.Path & "\rdndeo.val")
If Cad <> "" Then
    'Si que existe el fichero de redondeos
    Open App.Path & "\rdndeo.val" For Input As #NF
    Line Input #NF, Cad
    If Cad <> "" Then
        If IsNumeric(Cad) Then
            Redondeo = CByte(Val(Cad))
        End If
    End If
    Close #NF
End If

FinLeerRedondeos:
    ClickRedondeos Redondeo
End Sub

Private Sub GrabarRedondeos()
Dim NF As Integer

On Error GoTo GrabarRedondeos
    NF = FreeFile
    'Si que existe el fichero de redondeos
    Open App.Path & "\rdndeo.val" For Output As #NF
    Print #NF, Redondeo
    Close #NF
GrabarRedondeos:
End Sub

Private Sub ClickRedondeos(Indice As Byte)
Redondeo = Indice
mnSinRedondeo.Checked = (Redondeo = 0)
mnDecima.Checked = (Redondeo = 1)
mncuartos.Checked = (Redondeo = 2)
mnMediaHora.Checked = (Redondeo = 3)
End Sub



Private Sub ConversionRedondeo(ByRef T1 As Single, ByRef T2 As Single, Precalculo As Boolean)
Dim T3 As Single
Dim Entera As Single
Dim resto As Single
Dim Divisor As Integer '
Dim margen As Single
Dim cociente As Integer
Dim V As Single


'Si no hay que redondear
If Redondeo = 0 Then Exit Sub

'Seguimos
Select Case Redondeo
Case 2
    Divisor = 25
    margen = 18
Case 3
    Divisor = 50
    margen = 38
Case Else  'Por si acso los recogemos en ELSE que es decima de punto
    Divisor = 10
    margen = 3
End Select
T3 = T1 + T2
'Cambiamos el valor de t1
Entera = Int(T1)
resto = Round((T1 - Entera) * 100, 0)


V = resto Mod Divisor
cociente = resto \ Divisor
'No se redondea nunca hacia arriba, luego la instrucciones van comentadas
If V >= margen Then
        cociente = cociente + 1
End If
V = cociente * Divisor  'Resto redondeado
V = V / 100
T1 = Entera + V
T2 = Round(T3 - T1, 2)
T2 = Abs(T2)

If Precalculo Then Exit Sub

'ASignamos a Marcajes
vMar.HorasIncid = T2
vMar.HorasTrabajadas = T1
txtMarcaje(4).Text = T1
txtMarcaje(4).Refresh
txtMarcaje(7).Text = T2
txtMarcaje(7).Refresh
''Horario en seagesimal
txtDec(0).Text = Format(DevuelveHora(vMar.HorasTrabajadas), "hh:mm")
txtDec(0).Refresh
txtDec(1).Text = Format(DevuelveHora(vMar.HorasIncid), "hh:mm")
txtDec(1).Refresh
End Sub





'------------------------------------------------------------------
'Cuando solo vemos si el num de tikadas es par y calculamos las horas 'trabajadas
Private Function ProcesarMarcajeRevision2(ByRef vMar As CMarcajes) As Boolean
Dim Rss As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim NumTikadas As Integer
Dim T1 As Single
Dim T2 As Single
Dim i As Long
Dim Cad As String
Dim N As Integer
Dim TotalH As Single
Dim HoE As Single
Dim vE As CEmpresas
Dim InciMan As Integer
'Ahora ya tenemos las horas tikadas reflejadas
'Comprobamos las horas en funcion de los horarios
'  y calculamos las horas comprobadas

On Error GoTo ErrorProcesaMarcaje_Tipo2_v
ProcesarMarcajeRevision2 = False
Set Rss = New ADODB.Recordset


Set Rss = New ADODB.Recordset
i = -1
Cad = "select IdEmpresa from Trabajadores WHERE IdTrabajador=" & vMar.idTrabajador
Rss.Open Cad, conn, , , adCmdText
If Not Rss.EOF Then
    If Not IsNull(Rss.Fields(0)) Then i = Rss.Fields(0)
End If
Rss.Close

Set vE = New CEmpresas
If vE.Leer(i) = 1 Then
    MsgBox "Error leyendo la empresa para el trabajador." & vbCrLf & " Cod. Empresa: " & i, vbExclamation
    Set vE = Nothing
    Exit Function
End If

'Borramos todas las incidencias, si tiene, generadas autmoaticamente
Cad = "DELETE * FROM IncidenciasGeneradas WHERE EntradaMarcaje=" & vMar.Entrada
Cad = Cad & " AND Incidencia<>" & vE.IncMarcaje
Rss.Open Cad, conn, , , adCmdText
'No hace falta cerrarlo puesto que no devuelve ningun registro

'Seleccionamos todas las horas de este
Cad = "SELECT * FROM EntradaMarcajes WHERE IdMarcaje=" & vMar.Entrada
Cad = Cad & " ORDER BY Hora"
Rss.CursorType = adOpenStatic
Rss.Open Cad, conn, , , adCmdText

If Rss.EOF Then
    'Si no hay ninguna entrada
    Rss.Close
    GoTo ErrorProcesaMarcaje_Tipo2_v
End If


'Si el numero de tikadas es par entonces calculamos las horas
NumTikadas = Rss.RecordCount
If (NumTikadas Mod 2) > 0 Then
    'Numero de marcajes impares. No podemos calcular horas
    'trabajadas. Generamos error en marcaje
    vMar.IncFinal = vE.IncMarcaje
    vMar.HorasIncid = 0
    vMar.HorasTrabajadas = 0
    Else
        N = NumTikadas \ 2
        TotalH = 0
        'NUMERO DE MARCAJES PAR
        Rss.MoveFirst
        PrimerTicaje = Rss!Hora
        InciMan = 0
        For i = 1 To N
            T1 = DevuelveValorHora(Rss!Hora)
            If Rss!idInci > 0 Then InciMan = 1
            Rss.MoveNext
            UltimoTicaje = Rss!Hora
            If Rss!idInci > 0 Then InciMan = 1
            T2 = DevuelveValorHora(Rss!Hora)
            Rss.MoveNext
            TotalH = TotalH + (T2 - T1)
            
        Next i
        
        'Comprobamos los detos almuerzo merienda
        '******************************************************
        'Comprobamos si hay que quitar los minutos del almuerzo
        If vH.DtoAlm > 0 Then
            If PrimerTicaje < vH.HoraDtoAlm Then
                TotalH = TotalH - vH.DtoAlm
                If TotalH < 0 Then TotalH = 0
                
                'Modificacion SEPTIPEMPRE 2004. No le resto las horas al computo de horas
                'vH.TotalHoras = vH.TotalHoras - vH.DtoAlm
            End If
        End If
        '----------------------------------------------
            
        'Comprobamos si hay que quitar los minutos de la MER
        'Como esta ya en el ultimo
        If vH.DtoMer > 0 Then
            If UltimoTicaje > vH.HoraDtoMer Then
                TotalH = TotalH - vH.DtoMer
                If TotalH < 0 Then TotalH = 0
            End If
        End If
        '----------------------------------------------
        
        
        
        'Ahora ya sabemos las horas trabajadas
        TotalH = Round(TotalH, 2)
        HoE = EntraDentro2(TotalH, vH.TotalHoras, vE.MaxExceso, vE.MaxRetraso)
        If HoE = 0 Then
            vMar.HorasTrabajadas = vH.TotalHoras
            vMar.HorasIncid = 0
            vMar.IncFinal = 0
            vMar.Correcto = True
            Else
                vMar.Correcto = False
                If HoE < 0 Then
                    'Horas extras
                    vMar.HorasTrabajadas = vH.TotalHoras - HoE
                    vMar.HorasIncid = Abs(HoE)
                    vMar.IncFinal = vE.IncHoraExtra
                    Else
                        'retraso, no ha llegado al minimo exigible
                        vMar.HorasTrabajadas = vH.TotalHoras - HoE
                        vMar.HorasIncid = HoE
                        vMar.IncFinal = vE.IncRetraso
                End If
        End If
        
        
        
        
        'Si tiene incidencia manual la ponemos
        If InciMan > 0 Then
            Rss.MoveFirst
            
            While Not Rss.EOF
                i = 0
                NumTikadas = 0
                
                T1 = DevuelveValorHora(Rss!Hora)
                i = Rss!idInci
                Rss.MoveNext
                NumTikadas = Rss!idInci
                T2 = DevuelveValorHora(Rss!Hora)
                Rss.MoveNext
                
                
                If i <> 0 Or NumTikadas <> 0 Then
                    T2 = T2 - T1
                    If i = 0 Then i = NumTikadas
                    GeneraIncidencia CInt(i), vMar.Entrada, T2
                End If
            Wend
        End If
End If
'Grabamos el marcaje
vMar.Modificar

Set Rss = Nothing
Set RFin = Nothing
Set vE = Nothing
ProcesarMarcajeRevision2 = True
Exit Function
ErrorProcesaMarcaje_Tipo2_v:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
End Function


'------------------------------------------------------------------
'Cuando solo vemos si el num de tikadas es par y calculamos las horas 'trabajadas
Private Function ProcesarMarcajeRevision3() As Boolean
Dim Rss As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim NumTikadas As Integer
Dim T1 As Single
Dim T2 As Single
Dim i As Long
Dim Cad As String
Dim N As Integer
Dim TotalH As Single
Dim HoE As Single
Dim vE As CEmpresas

'Ahora ya tenemos las horas tikadas reflejadas
'Comprobamos las horas en funcion de los horarios
'  y calculamos las horas comprobadas

On Error GoTo ErrorProcesaMarcaje_Tipo2_3
ProcesarMarcajeRevision3 = False
Set Rss = New ADODB.Recordset


Set Rss = New ADODB.Recordset
i = -1
Cad = "select IdEmpresa from Trabajadores WHERE IdTrabajador=" & vMar.idTrabajador
Rss.Open Cad, conn, , , adCmdText
If Not Rss.EOF Then
    If Not IsNull(Rss.Fields(0)) Then i = Rss.Fields(0)
End If
Rss.Close

Set vE = New CEmpresas
If vE.Leer(i) = 1 Then
    MsgBox "Error leyendo la empresa para el trabajador." & vbCrLf & " Cod. Empresa: " & i, vbExclamation
    Set vE = Nothing
    Exit Function
End If

'Borramos todas las incidencias, si tiene, generadas autmoaticamente
Cad = "DELETE * FROM IncidenciasGeneradas WHERE EntradaMarcaje=" & vMar.Entrada
Cad = Cad & " AND Incidencia<>" & vE.IncMarcaje
Rss.Open Cad, conn, , , adCmdText
'No hace falta cerrarlo puesto que no devuelve ningun registro

'Seleccionamos todas las horas de este
Cad = "SELECT * FROM EntradaMarcajes WHERE IdMarcaje=" & vMar.Entrada
Cad = Cad & " ORDER BY Hora"

Rss.CursorType = adOpenStatic
Rss.Open Cad, conn, , , adCmdText

If Rss.EOF Then
    'Si no hay ninguna entrada
    Rss.Close
    GoTo ErrorProcesaMarcaje_Tipo2_3
End If


'Si el numero de tikadas es par entonces calculamos las horas
NumTikadas = Rss.RecordCount
If (NumTikadas Mod 2) > 0 Then
    'Numero de marcajes impares. No podemos calcular horas
    'trabajadas. Generamos error en marcaje
    vMar.IncFinal = vE.IncMarcaje
    vMar.HorasIncid = 0
    vMar.HorasTrabajadas = 0
    vMar.IncFinal = 0
    vMar.Correcto = False
    GeneraIncidencia vE.IncMarcaje, vMar.Entrada, 0
    Else
        N = NumTikadas \ 2
        TotalH = 0
        'NUMERO DE MARCAJES PAR
        Rss.MoveFirst
        'Para dto horas
        PrimerTicaje = Rss!Hora
        For i = 1 To N
            T1 = DevuelveValorHora(Rss!Hora)
            Rss.MoveNext
            T2 = DevuelveValorHora(Rss!Hora)
            UltimoTicaje = Rss!Hora
            Rss.MoveNext
            TotalH = TotalH + (T2 - T1)
        Next i
            
            
            '******************************************************
            'Comprobamos si hay que quitar los minutos del almuerzo
            If vH.DtoAlm > 0 Then
                If PrimerTicaje < vH.HoraDtoAlm Then
                    TotalH = TotalH - vH.DtoAlm
                    If TotalH < 0 Then TotalH = 0
                End If
            End If
            
            '----------------------------------------------
            'Comprobamos si hay que quitar los minutos de la MER
            'Como esta ya en el ultimo
            If vH.DtoMer > 0 Then
                If UltimoTicaje > vH.HoraDtoMer Then
                    TotalH = TotalH - vH.DtoMer
                    If TotalH < 0 Then TotalH = 0
                End If
            End If
            '----------------------------------------------
            '******************************************************
        
        
        
        
        'Ahora ya sabemos las horas trabajadas
        TotalH = Round(TotalH, 2)
        vMar.HorasTrabajadas = TotalH
        If vH.EsDiaFestivo Then
            vMar.HorasIncid = TotalH
            vMar.IncFinal = vE.IncHoraExtra
            Else
                vMar.IncFinal = 0
        End If
End If

'Grabamos el marcaje
vMar.Modificar

Set Rss = Nothing
Set RFin = Nothing
Set vE = Nothing
ProcesarMarcajeRevision3 = True
Exit Function
ErrorProcesaMarcaje_Tipo2_3:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
End Function



Private Function QuitarAlmuerzo(ByRef m As CMarcajes, H As CHorarios) As Boolean
Dim RT As ADODB.Recordset
Dim H1 As Date
Dim N As Integer
Dim i As Integer

On Error GoTo EQuitar
    QuitarAlmuerzo = False
    
    If H.DtoAlm <= 0 Then Exit Function
    Set RT = New ADODB.Recordset
    RT.Open "Select * from EntradaMarcajes WHERE idMarcaje=" & m.Entrada & " ORDER BY Hora", conn, adOpenKeyset, adLockPessimistic, adCmdText
    'Revisaremos por pares
    N = RT.RecordCount \ 2
    For i = 1 To N
        H1 = RT!Hora
        RT.MoveNext
        
        'Si
        If H1 <= H.HoraDtoAlm Then
            If RT!Hora > H.HoraDtoAlm Then QuitarAlmuerzo = True
        End If
        
        RT.MoveNext
    Next i
    RT.Close

EQuitar:
    If Err.Number <> 0 Then Err.Clear
    Set RT = Nothing
End Function



'Con esto lo que pretendo hacer es k no hace falta k vaya uno a uno revisando si no
'que si esta marcado incorrectos el progrma vaya 1 a uno revisnado todos los marcajes
'y si puede los pase a correctos

Private Sub RevisionGeneralIncorectos()
Dim bien As Integer
Dim mal As Integer

    If Adodc1.Recordset.EOF Then Exit Sub
    Screen.MousePointer = vbHourglass
    If Me.mnIncorrectas Then
    
        Adodc1.Recordset.MoveFirst
        Label15.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
        Label18.Caption = ""
        Label19.Caption = ""
        Frame4.Visible = True
        Me.Refresh
        Screen.MousePointer = vbHourglass
        bien = 0
        mal = 0
        While Not Adodc1.Recordset.EOF
            PonerTodosLosCampos
            StatusBar1.Panels(1).Text = "Revisando"
            Me.Refresh
            espera 0.2
            If RevisaUNO Then
                bien = bien + 1
            Else
                mal = mal + 1
            End If
            Adodc1.Recordset.MoveNext
            Label15.Caption = Adodc1.Recordset.AbsolutePosition & " de " & Adodc1.Recordset.RecordCount
            Label18.Caption = bien
            Label19.Caption = mal
            
            Me.Refresh
            Screen.MousePointer = vbHourglass
            
        Wend
        
        Frame4.Visible = False
        StatusBar1.Panels(1).Text = "Recalculando"
        Me.Refresh
        
        'Volvemos a cargar todo
        Screen.MousePointer = vbHourglass
        espera 2
        Adodc1.Refresh
        '-----------------------
        'pequeña espera
        espera 1
        '----------------------
        Adodc1.Refresh
        kRegistro = 1
        RegTotales = Adodc1.Recordset.RecordCount
        StatusBar1.Panels(1).Text = kRegistro & " de " & RegTotales
        If Not Adodc1.Recordset.EOF Then
            PonerTodosLosCampos
        Else
            LimpiarCampos
            MsgBox "Ningún registro por mostrar.", vbInformation
            'Unload Me
        End If
        Screen.MousePointer = vbDefault
    End If
    
End Sub



Private Function RevisaUNO() As Boolean

On Error Resume Next
    RevisaUNO = False
    If DatosOk(False) Then
        vMar.PonerCorrecta
        RevisaUNO = True
    End If
    If Err.Number <> 0 Then Err.Clear
End Function

