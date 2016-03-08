VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Programa para el control de presencia"
   ClientHeight    =   8310
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   11880
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   Begin MSComDlg.CommonDialog cd1 
      Left            =   6060
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   17
      Top             =   7320
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4020
      TabIndex        =   16
      Text            =   "fecha"
      Top             =   6900
      Width           =   1215
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   23
      Top             =   7875
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   16052
            MinWidth        =   16052
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "INS"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "NÚM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            TextSave        =   "MAYÚS"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Object.Width           =   1058
            MinWidth        =   1058
            TextSave        =   "18:43"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Procesar marcajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   20
      Left            =   1680
      MouseIcon       =   "frmMain.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   26
      Top             =   6300
      Width           =   2250
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   4
      X1              =   1560
      X2              =   1320
      Y1              =   6420
      Y2              =   6420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sección"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   19
      Left            =   1680
      MouseIcon       =   "frmMain.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   25
      Top             =   2076
      Width           =   990
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   1500
      X2              =   1260
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   3
      X1              =   1320
      X2              =   1320
      Y1              =   5400
      Y2              =   7020
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Index           =   3
      X1              =   600
      X2              =   600
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Index           =   3
      X1              =   780
      X2              =   600
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   3
      X1              =   3540
      X2              =   3720
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   3
      X1              =   3720
      X2              =   3720
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   3
      X1              =   1560
      X2              =   1320
      Y1              =   7020
      Y2              =   7020
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   3
      X1              =   1560
      X2              =   1320
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   2
      X1              =   7500
      X2              =   7500
      Y1              =   5400
      Y2              =   7080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Index           =   2
      X1              =   6780
      X2              =   6780
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Index           =   2
      X1              =   6960
      X2              =   6780
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   2
      X1              =   8520
      X2              =   8700
      Y1              =   5100
      Y2              =   5100
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   2
      X1              =   8700
      X2              =   8700
      Y1              =   5100
      Y2              =   5460
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   2
      X1              =   7740
      X2              =   7500
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   2
      X1              =   7800
      X2              =   7500
      Y1              =   6420
      Y2              =   6420
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   2
      X1              =   7740
      X2              =   7500
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   7500
      X2              =   7500
      Y1              =   1260
      Y2              =   3660
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Index           =   1
      X1              =   6780
      X2              =   6780
      Y1              =   1020
      Y2              =   1380
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Index           =   1
      X1              =   6960
      X2              =   6780
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   9360
      X2              =   9540
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   9540
      X2              =   9540
      Y1              =   1080
      Y2              =   1440
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   7740
      X2              =   7500
      Y1              =   3660
      Y2              =   3660
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   7740
      X2              =   7500
      Y1              =   2940
      Y2              =   2940
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   7800
      X2              =   7500
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   1
      X1              =   7740
      X2              =   7500
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1500
      X2              =   1260
      Y1              =   1620
      Y2              =   1620
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1500
      X2              =   1260
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1500
      X2              =   1260
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1500
      X2              =   1260
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1500
      X2              =   1260
      Y1              =   3960
      Y2              =   3960
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   3780
      X2              =   3780
      Y1              =   1020
      Y2              =   1380
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   3600
      X2              =   3780
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Index           =   0
      X1              =   720
      X2              =   540
      Y1              =   1020
      Y2              =   1020
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00004080&
      BorderWidth     =   2
      Index           =   0
      X1              =   540
      X2              =   540
      Y1              =   1020
      Y2              =   1380
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      Index           =   0
      X1              =   1260
      X2              =   1260
      Y1              =   1260
      Y2              =   4560
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Incorrectos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2040
      TabIndex        =   24
      Top             =   7260
      Width           =   1215
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
      TabIndex        =   22
      Top             =   0
      Width           =   9495
   End
   Begin VB.Label Label2 
      BackColor       =   &H00800000&
      Height          =   645
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incidencias generadas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   18
      Left            =   7860
      MouseIcon       =   "frmMain.frx":091E
      MousePointer    =   99  'Custom
      TabIndex        =   20
      Top             =   3480
      Width           =   2955
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incidencias   RESUMEN"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   17
      Left            =   7860
      MouseIcon       =   "frmMain.frx":0C28
      MousePointer    =   99  'Custom
      TabIndex        =   19
      Top             =   2820
      Width           =   3000
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   5280
      Picture         =   "frmMain.frx":0F32
      Top             =   6960
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuración"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   14
      Left            =   7800
      MouseIcon       =   "frmMain.frx":1034
      MousePointer    =   99  'Custom
      TabIndex        =   18
      Top             =   5640
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Operaciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Index           =   1
      Left            =   900
      MouseIcon       =   "frmMain.frx":133E
      TabIndex        =   2
      Top             =   4860
      Width           =   2490
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Otros"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Index           =   3
      Left            =   7080
      MouseIcon       =   "frmMain.frx":1648
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   16
      Left            =   7860
      MouseIcon       =   "frmMain.frx":1952
      MousePointer    =   99  'Custom
      TabIndex        =   15
      Top             =   6960
      Width           =   570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Acerca de ..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   15
      Left            =   7860
      MouseIcon       =   "frmMain.frx":1C5C
      MousePointer    =   99  'Custom
      TabIndex        =   14
      Top             =   6300
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Presencia"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   13
      Left            =   7860
      MouseIcon       =   "frmMain.frx":1F66
      MousePointer    =   99  'Custom
      TabIndex        =   13
      Top             =   1500
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Resumen horas trabajadas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   12
      Left            =   7860
      MouseIcon       =   "frmMain.frx":2270
      MousePointer    =   99  'Custom
      TabIndex        =   12
      Top             =   2160
      Width           =   3570
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Importar fichero de datos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   11
      Left            =   1680
      MouseIcon       =   "frmMain.frx":257A
      MousePointer    =   99  'Custom
      TabIndex        =   11
      Top             =   5640
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revisar marcajes"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   10
      Left            =   1680
      MouseIcon       =   "frmMain.frx":2884
      MousePointer    =   99  'Custom
      TabIndex        =   10
      Top             =   6900
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Revisar incorrectos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Index           =   9
      Left            =   9240
      MouseIcon       =   "frmMain.frx":2B8E
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   7500
      Visible         =   0   'False
      Width           =   2340
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Incidencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   8
      Left            =   1680
      MouseIcon       =   "frmMain.frx":2E98
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   4380
      Width           =   1470
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horarios"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   7
      Left            =   1680
      MouseIcon       =   "frmMain.frx":31A2
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3804
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Categorías"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   6
      Left            =   1680
      MouseIcon       =   "frmMain.frx":34AC
      MousePointer    =   99  'Custom
      TabIndex        =   6
      Top             =   3228
      Width           =   1425
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajadores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   5
      Left            =   1680
      MouseIcon       =   "frmMain.frx":37B6
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   2652
      Width           =   1800
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Empresas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Index           =   4
      Left            =   1680
      MouseIcon       =   "frmMain.frx":3AC0
      MousePointer    =   99  'Custom
      TabIndex        =   0
      Top             =   1500
      Width           =   1290
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Informes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Index           =   2
      Left            =   7080
      MouseIcon       =   "frmMain.frx":3DCA
      TabIndex        =   3
      Top             =   840
      Width           =   1845
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Datos básicos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Index           =   0
      Left            =   780
      MouseIcon       =   "frmMain.frx":40D4
      TabIndex        =   1
      Top             =   780
      Width           =   2760
   End
   Begin VB.Image Image33 
      Height          =   8940
      Left            =   0
      Picture         =   "frmMain.frx":43DE
      Top             =   -660
      Width           =   11850
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



Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Text1.Text = "" 'Format(Now, "dd/mm/yyyy")
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
mnOperacionesTCP3.Enabled = mConfig.TCP3
mnTraspasar.Enabled = mConfig.Ariadna
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Conn.Close
Set Conn = Nothing
End Sub

Private Sub frmF_Selec(vFecha As Date)
Text1.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
Set frmF = New frmCal
frmF.Fecha = (Now - 1)
frmF.Show vbModal
Set frmF = Nothing
End Sub



Private Sub Label1_Click(Index As Integer)
Screen.MousePointer = vbHourglass
Select Case Index
Case 4
    frmEmpresas.Show vbModal
Case 5
    frmEmpleados.Show vbModal
Case 6
    frmCategoria.Show vbModal
Case 7
    frmHorario.Show vbModal
Case 8
    frmIncidencias.Show vbModal
Case 9
    frmRevision2.Todos = 2
    frmRevision2.vFecha = ""
    frmRevision2.Show vbModal
Case 10
    If Text1.Text <> "" Then
        If Not IsDate(Text1.Text) Then
            MsgBox "La fecha seleccionada no es una fecha correcta.", vbExclamation
            Exit Sub
        End If
    End If
    frmRevision2.Todos = (Check1.Value + 1)
    frmRevision2.vFecha = Text1.Text
    frmRevision2.Show vbModal
Case 11
    mnImportar_Click
Case 12
    'Informes
    frmInformes.Opcion = 2
    frmInformes.Show vbModal
Case 13
    'Informes
    frmInformes.Opcion = 1
    frmInformes.Show vbModal
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
    frmInfInc.Show vbModal
Case 18
    frmInfIncGen.Show vbModal
Case 19
    frmSeccion.Show vbModal
Case 20
    mnProcesar_Click
End Select
Screen.MousePointer = vbDefault
End Sub


Private Sub Label4_Click()
If Check1.Value = 0 Then
    Check1.Value = 1
    Else
        Check1.Value = 0
End If
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


Private Sub mnEmpresas_Click()
Label1_Click 4
End Sub

Private Sub mnEntrada_Click()
Check1.Tag = Check1.Value
Check1.Value = 0
Label1_Click 10
Check1.Value = Check1.Tag
End Sub

Private Sub mnGeneradas_Click()
Label1_Click 18
End Sub

Private Sub mnHorarios_Click()
Label1_Click 7
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
Label1_Click 17
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

Private Sub mnPresencia_Click()
Label1_Click 13
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
Label1_Click 12
End Sub

Private Sub mnRevisar_Click()
Label1_Click 9
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

Private Sub mnTrabajadores_Click()
Label1_Click 5
End Sub

Private Sub mnTraspasar_Click()
frmUnix.Show vbModal
End Sub
