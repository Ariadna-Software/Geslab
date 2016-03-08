VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmHorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento horarios"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   11370
   Icon            =   "frmHorario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6465
   ScaleWidth      =   11370
   Visible         =   0   'False
   Begin VB.CheckBox Check2 
      Caption         =   "Recupera horas sábado"
      Height          =   255
      Left            =   7320
      TabIndex        =   108
      Top             =   720
      Width           =   3615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4755
      Left            =   120
      TabIndex        =   57
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   8387
      _Version        =   393216
      Tab             =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Horario semanal"
      TabPicture(0)   =   "frmHorario.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Line1(15)"
      Tab(0).Control(1)=   "Line1(14)"
      Tab(0).Control(2)=   "Line1(13)"
      Tab(0).Control(3)=   "Line1(11)"
      Tab(0).Control(4)=   "Line1(10)"
      Tab(0).Control(5)=   "Line1(9)"
      Tab(0).Control(6)=   "Line1(7)"
      Tab(0).Control(7)=   "Line1(6)"
      Tab(0).Control(8)=   "Line1(5)"
      Tab(0).Control(9)=   "Line1(4)"
      Tab(0).Control(10)=   "Line1(3)"
      Tab(0).Control(11)=   "Line1(2)"
      Tab(0).Control(12)=   "Line1(1)"
      Tab(0).Control(13)=   "Line1(0)"
      Tab(0).Control(14)=   "Label4(4)"
      Tab(0).Control(15)=   "Label2(2)"
      Tab(0).Control(16)=   "Label4(3)"
      Tab(0).Control(17)=   "Label4(2)"
      Tab(0).Control(18)=   "Label4(1)"
      Tab(0).Control(19)=   "Label4(0)"
      Tab(0).Control(20)=   "Label2(1)"
      Tab(0).Control(21)=   "Label3(6)"
      Tab(0).Control(22)=   "Label3(5)"
      Tab(0).Control(23)=   "Label3(4)"
      Tab(0).Control(24)=   "Label3(3)"
      Tab(0).Control(25)=   "Label3(2)"
      Tab(0).Control(26)=   "Label3(1)"
      Tab(0).Control(27)=   "Label3(0)"
      Tab(0).Control(28)=   "Label2(0)"
      Tab(0).Control(29)=   "Label4(5)"
      Tab(0).Control(30)=   "Check1(6)"
      Tab(0).Control(31)=   "Text1(33)"
      Tab(0).Control(32)=   "Text1(34)"
      Tab(0).Control(33)=   "Text1(35)"
      Tab(0).Control(34)=   "Text1(36)"
      Tab(0).Control(35)=   "Text1(37)"
      Tab(0).Control(36)=   "Check1(5)"
      Tab(0).Control(37)=   "Text1(28)"
      Tab(0).Control(38)=   "Text1(29)"
      Tab(0).Control(39)=   "Text1(30)"
      Tab(0).Control(40)=   "Text1(31)"
      Tab(0).Control(41)=   "Text1(32)"
      Tab(0).Control(42)=   "Check1(4)"
      Tab(0).Control(43)=   "Text1(23)"
      Tab(0).Control(44)=   "Text1(24)"
      Tab(0).Control(45)=   "Text1(25)"
      Tab(0).Control(46)=   "Text1(26)"
      Tab(0).Control(47)=   "Text1(27)"
      Tab(0).Control(48)=   "Check1(3)"
      Tab(0).Control(49)=   "Text1(18)"
      Tab(0).Control(50)=   "Text1(19)"
      Tab(0).Control(51)=   "Text1(20)"
      Tab(0).Control(52)=   "Text1(21)"
      Tab(0).Control(53)=   "Text1(22)"
      Tab(0).Control(54)=   "Check1(2)"
      Tab(0).Control(55)=   "Text1(13)"
      Tab(0).Control(56)=   "Text1(14)"
      Tab(0).Control(57)=   "Text1(15)"
      Tab(0).Control(58)=   "Text1(16)"
      Tab(0).Control(59)=   "Text1(17)"
      Tab(0).Control(60)=   "Check1(1)"
      Tab(0).Control(61)=   "Text1(8)"
      Tab(0).Control(62)=   "Text1(9)"
      Tab(0).Control(63)=   "Text1(10)"
      Tab(0).Control(64)=   "Text1(11)"
      Tab(0).Control(65)=   "Text1(12)"
      Tab(0).Control(66)=   "Text1(7)"
      Tab(0).Control(67)=   "Text1(6)"
      Tab(0).Control(68)=   "Text1(5)"
      Tab(0).Control(69)=   "Text1(4)"
      Tab(0).Control(70)=   "Check1(0)"
      Tab(0).Control(71)=   "Text1(39)"
      Tab(0).Control(72)=   "Text1(40)"
      Tab(0).Control(73)=   "Text1(41)"
      Tab(0).Control(74)=   "Text1(42)"
      Tab(0).Control(75)=   "Text1(43)"
      Tab(0).Control(76)=   "Text1(44)"
      Tab(0).Control(77)=   "Text1(3)"
      Tab(0).Control(78)=   "Text1(38)"
      Tab(0).ControlCount=   79
      TabCaption(1)   =   "Dias festivos"
      TabPicture(1)   =   "frmHorario.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label6"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label7"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label8"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Combo1"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Command1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Command2"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Command3"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command4"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "ListView1"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).ControlCount=   9
      TabCaption(2)   =   "Paradas"
      TabPicture(2)   =   "frmHorario.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame1"
      Tab(2).ControlCount=   1
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   38
         Left            =   -72420
         TabIndex        =   5
         Tag             =   "0"
         Text            =   "T"
         Top             =   1320
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   -71700
         TabIndex        =   6
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   44
         Left            =   -72420
         TabIndex        =   47
         Tag             =   "0"
         Text            =   "3"
         Top             =   4200
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   43
         Left            =   -72420
         TabIndex        =   40
         Tag             =   "0"
         Text            =   "T"
         Top             =   3720
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   42
         Left            =   -72420
         TabIndex        =   34
         Tag             =   "0"
         Text            =   "T"
         Top             =   3240
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   41
         Left            =   -72420
         TabIndex        =   26
         Tag             =   "0"
         Text            =   "t"
         Top             =   2760
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   40
         Left            =   -72420
         TabIndex        =   19
         Tag             =   "0"
         Text            =   "T"
         Top             =   2280
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   39
         Left            =   -72420
         TabIndex        =   12
         Tag             =   "0"
         Text            =   "T"
         Top             =   1800
         Width           =   500
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   3195
         Left            =   2040
         TabIndex        =   106
         Top             =   1140
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   5636
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList2"
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Fecha"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripción"
            Object.Width           =   7056
         EndProperty
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Copiar horario"
         Height          =   855
         Left            =   300
         TabIndex        =   105
         Top             =   3480
         Width           =   1395
      End
      Begin VB.Frame Frame1 
         Height          =   4155
         Left            =   -74880
         TabIndex        =   84
         Top             =   420
         Width           =   10935
         Begin VB.TextBox txtAlm 
            Alignment       =   1  'Right Justify
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   85
            Text            =   "Text2"
            Top             =   960
            Width           =   975
         End
         Begin VB.TextBox txtAlm 
            Alignment       =   2  'Center
            Height          =   255
            Index           =   1
            Left            =   3300
            TabIndex        =   86
            Text            =   "Text2"
            Top             =   960
            Width           =   1155
         End
         Begin VB.TextBox txtAlm 
            Alignment       =   2  'Center
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   87
            Text            =   "Text2"
            Top             =   1680
            Width           =   1035
         End
         Begin VB.TextBox txtMer 
            Alignment       =   1  'Right Justify
            Height          =   255
            Index           =   0
            Left            =   2280
            TabIndex        =   88
            Text            =   "Text2"
            Top             =   2880
            Width           =   975
         End
         Begin VB.TextBox txtMer 
            Alignment       =   2  'Center
            Height          =   255
            Index           =   1
            Left            =   3300
            TabIndex        =   89
            Text            =   "Text2"
            Top             =   2880
            Width           =   1155
         End
         Begin VB.TextBox txtMer 
            Alignment       =   2  'Center
            Height          =   255
            Index           =   2
            Left            =   2280
            TabIndex        =   90
            Text            =   "Text2"
            Top             =   3600
            Width           =   975
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Almuerzo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Index           =   0
            Left            =   120
            TabIndex        =   104
            Top             =   240
            Width           =   1050
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Merienda"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   270
            Index           =   1
            Left            =   180
            TabIndex        =   103
            Top             =   2100
            Width           =   1050
         End
         Begin VB.Label Label10 
            Caption         =   "Descuento "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   900
            TabIndex        =   102
            Top             =   600
            Width           =   1155
         End
         Begin VB.Label Label10 
            Caption         =   "Hora "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   900
            TabIndex        =   101
            Top             =   1440
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Hora a partir de la cual NO se contabilizará el almuerzo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   4680
            TabIndex        =   100
            Top             =   1740
            Width           =   5715
         End
         Begin VB.Label Label11 
            Caption         =   "Descuento en minutos que se descontarán por el almuerzo. Cero(0) es no descuento almuerzo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   1
            Left            =   4680
            TabIndex        =   99
            Top             =   900
            Width           =   5835
         End
         Begin VB.Label Label11 
            Caption         =   "Decimal"
            Height          =   255
            Index           =   2
            Left            =   2460
            TabIndex        =   98
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Sexagesimal"
            Height          =   255
            Index           =   3
            Left            =   3360
            TabIndex        =   97
            Top             =   720
            Width           =   1035
         End
         Begin VB.Label Label10 
            Caption         =   "Descuento "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   900
            TabIndex        =   96
            Top             =   2520
            Width           =   1155
         End
         Begin VB.Label Label10 
            Caption         =   "Hora "
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   900
            TabIndex        =   95
            Top             =   3360
            Width           =   1155
         End
         Begin VB.Label Label11 
            Caption         =   "Hora a partir de la cual se contabilizará la merienda"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   4680
            TabIndex        =   94
            Top             =   3660
            Width           =   5715
         End
         Begin VB.Label Label11 
            Caption         =   "Descuento en minutos que se descontarán por la merienda. Cero(0) es no descuento almuerzo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   5
            Left            =   4680
            TabIndex        =   93
            Top             =   2820
            Width           =   5835
         End
         Begin VB.Label Label11 
            Caption         =   "Decimal"
            Height          =   255
            Index           =   6
            Left            =   2460
            TabIndex        =   92
            Top             =   2640
            Width           =   735
         End
         Begin VB.Label Label11 
            Caption         =   "Sexagesimal"
            Height          =   255
            Index           =   7
            Left            =   3360
            TabIndex        =   91
            Top             =   2640
            Width           =   1035
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00C00000&
            BorderWidth     =   2
            Index           =   0
            X1              =   1260
            X2              =   10620
            Y1              =   360
            Y2              =   360
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000080&
            BorderWidth     =   2
            Index           =   1
            X1              =   1320
            X2              =   10620
            Y1              =   2220
            Y2              =   2220
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Eliminar dia festivo"
         Height          =   435
         Left            =   8520
         TabIndex        =   82
         Top             =   3120
         Width           =   2415
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Modificar dia festivo"
         Height          =   435
         Left            =   8520
         TabIndex        =   81
         Top             =   2280
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Nuevo dia festivo"
         Height          =   435
         Left            =   8520
         TabIndex        =   79
         Top             =   1500
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmHorario.frx":035E
         Left            =   300
         List            =   "frmHorario.frx":038C
         Style           =   2  'Dropdown List
         TabIndex        =   77
         Top             =   1200
         Width           =   1515
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Festivo"
         Height          =   195
         Index           =   0
         Left            =   -73620
         TabIndex        =   3
         Top             =   1380
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   -70500
         TabIndex        =   7
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   5
         Left            =   -68460
         TabIndex        =   8
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   -67200
         TabIndex        =   9
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   1320
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   -64680
         TabIndex        =   10
         Tag             =   "0"
         Text            =   "T"
         Top             =   1320
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   12
         Left            =   -64680
         TabIndex        =   17
         Tag             =   "0"
         Text            =   "T"
         Top             =   1800
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   -67200
         TabIndex        =   16
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   -68460
         TabIndex        =   15
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   -70500
         TabIndex        =   14
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   1800
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   8
         Left            =   -71700
         TabIndex        =   13
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   1800
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Festivo"
         Height          =   195
         Index           =   1
         Left            =   -73620
         TabIndex        =   11
         Top             =   1860
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   -64680
         TabIndex        =   24
         Tag             =   "0"
         Text            =   "T"
         Top             =   2280
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   16
         Left            =   -67200
         TabIndex        =   23
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   15
         Left            =   -68460
         TabIndex        =   22
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -70500
         TabIndex        =   21
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -71700
         TabIndex        =   20
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   2280
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Festivo"
         Height          =   195
         Index           =   2
         Left            =   -73620
         TabIndex        =   18
         Top             =   2340
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   22
         Left            =   -64680
         TabIndex        =   32
         Tag             =   "0"
         Text            =   "t"
         Top             =   2760
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   21
         Left            =   -67200
         TabIndex        =   31
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   20
         Left            =   -68460
         TabIndex        =   30
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   19
         Left            =   -70500
         TabIndex        =   29
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   18
         Left            =   -71700
         TabIndex        =   28
         Tag             =   "0"
         Text            =   "18"
         Top             =   2760
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Festivo"
         Height          =   195
         Index           =   3
         Left            =   -73620
         TabIndex        =   25
         Top             =   2820
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   27
         Left            =   -64680
         TabIndex        =   38
         Tag             =   "0"
         Text            =   "T"
         Top             =   3240
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   -67200
         TabIndex        =   27
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   25
         Left            =   -68460
         TabIndex        =   37
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   24
         Left            =   -70500
         TabIndex        =   36
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   3240
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   23
         Left            =   -71700
         TabIndex        =   35
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   3240
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Festivo"
         Height          =   195
         Index           =   4
         Left            =   -73620
         TabIndex        =   33
         Top             =   3300
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   32
         Left            =   -64680
         TabIndex        =   45
         Tag             =   "0"
         Text            =   "T"
         Top             =   3720
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   31
         Left            =   -67200
         TabIndex        =   44
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   30
         Left            =   -68460
         TabIndex        =   43
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   -70500
         TabIndex        =   42
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   3720
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   -71700
         TabIndex        =   41
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   3720
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Festivo"
         Height          =   195
         Index           =   5
         Left            =   -73620
         TabIndex        =   39
         Top             =   3780
         Width           =   915
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   37
         Left            =   -64680
         TabIndex        =   51
         Tag             =   "0"
         Text            =   "3"
         Top             =   4200
         Width           =   500
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   36
         Left            =   -67200
         TabIndex        =   80
         Tag             =   "0"
         Text            =   "40"
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   35
         Left            =   -68460
         TabIndex        =   4
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   34
         Left            =   -70500
         TabIndex        =   49
         Tag             =   "0"
         Text            =   "34"
         Top             =   4200
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   33
         Left            =   -71700
         TabIndex        =   48
         Tag             =   "0"
         Text            =   "Text1"
         Top             =   4200
         Width           =   855
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Festivo"
         Height          =   195
         Index           =   6
         Left            =   -73620
         TabIndex        =   46
         Top             =   4260
         Width           =   915
      End
      Begin VB.Label Label4 
         Caption         =   "Dias/Nómina"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   5
         Left            =   -72600
         TabIndex        =   107
         Top             =   900
         Width           =   795
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   4800
         TabIndex        =   83
         Top             =   4320
         Width           =   3555
      End
      Begin VB.Label Label7 
         Caption         =   "años"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   6180
         TabIndex        =   78
         Top             =   660
         Width           =   1635
      End
      Begin VB.Label Label6 
         Caption         =   "Calendario de dias festivos para el año:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   495
         Left            =   240
         TabIndex        =   76
         Top             =   660
         Width           =   5775
      End
      Begin VB.Label Label2 
         Caption         =   "HORAS 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   0
         Left            =   -71040
         TabIndex        =   75
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label3 
         Caption         =   "Lunes"
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
         Index           =   0
         Left            =   -74580
         TabIndex        =   74
         Top             =   1380
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Martes"
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
         Index           =   1
         Left            =   -74580
         TabIndex        =   73
         Top             =   1860
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Miércoles"
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
         Index           =   2
         Left            =   -74580
         TabIndex        =   72
         Top             =   2340
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Jueves"
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
         Index           =   3
         Left            =   -74580
         TabIndex        =   71
         Top             =   2820
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Viernes"
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
         Index           =   4
         Left            =   -74580
         TabIndex        =   70
         Top             =   3300
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Sábado"
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
         Index           =   5
         Left            =   -74580
         TabIndex        =   69
         Top             =   3780
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Domingo"
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
         Index           =   6
         Left            =   -74580
         TabIndex        =   68
         Top             =   4260
         Width           =   855
      End
      Begin VB.Label Label2 
         Caption         =   "HORAS 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   1
         Left            =   -67860
         TabIndex        =   67
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   -71580
         TabIndex        =   66
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   -70380
         TabIndex        =   65
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Entrada"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   2
         Left            =   -68400
         TabIndex        =   64
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label4 
         Caption         =   "Salida"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   3
         Left            =   -67080
         TabIndex        =   63
         Top             =   900
         Width           =   675
      End
      Begin VB.Label Label2 
         Caption         =   "DIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   195
         Index           =   2
         Left            =   -74520
         TabIndex        =   62
         Top             =   600
         Width           =   1155
      End
      Begin VB.Label Label4 
         Caption         =   "Horas/Día"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   4
         Left            =   -64860
         TabIndex        =   61
         Top             =   900
         Width           =   795
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   -66120
         X2              =   -64800
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   -69600
         X2              =   -68600
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   -66120
         X2              =   -64800
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   -69600
         X2              =   -68600
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   -66120
         X2              =   -64800
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         Index           =   5
         X1              =   -69600
         X2              =   -68600
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Line Line1 
         Index           =   6
         X1              =   -66120
         X2              =   -64800
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         Index           =   7
         X1              =   -69600
         X2              =   -68600
         Y1              =   2880
         Y2              =   2880
      End
      Begin VB.Line Line1 
         Index           =   9
         X1              =   -69600
         X2              =   -68600
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         Index           =   10
         X1              =   -69600
         X2              =   -68600
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         Index           =   11
         X1              =   -69600
         X2              =   -68600
         Y1              =   4320
         Y2              =   4320
      End
      Begin VB.Line Line1 
         Index           =   13
         X1              =   -66120
         X2              =   -64800
         Y1              =   3360
         Y2              =   3360
      End
      Begin VB.Line Line1 
         Index           =   14
         X1              =   -66120
         X2              =   -64800
         Y1              =   3840
         Y2              =   3840
      End
      Begin VB.Line Line1 
         Index           =   15
         X1              =   -66120
         X2              =   -64800
         Y1              =   4320
         Y2              =   4320
      End
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   120
      TabIndex        =   53
      Top             =   5820
      Width           =   3435
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Label5"
         Height          =   195
         Left            =   660
         TabIndex        =   54
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   5400
      TabIndex        =   2
      Tag             =   "Direccion|T|S|||"
      Text            =   "0"
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   10020
      TabIndex        =   60
      Top             =   6060
      Width           =   1300
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8520
      TabIndex        =   59
      Top             =   6060
      Width           =   1300
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   120
      Top             =   1680
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
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
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   6960
      TabIndex        =   58
      Top             =   6060
      Width           =   1300
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1800
      MaxLength       =   50
      TabIndex        =   1
      Tag             =   "Nombre|T|N|||"
      Text            =   "Text1"
      Top             =   720
      Width           =   3195
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   180
      TabIndex        =   0
      Tag             =   "Código|N|S|||"
      Text            =   "Text1"
      Top             =   720
      Width           =   1155
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   56
      Top             =   0
      Width           =   11370
      _ExtentX        =   20055
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
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
            ImageIndex      =   6
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
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":03E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":04F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":0608
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":071A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":082C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":093E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":1218
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":1AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":23CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":2CA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHorario.frx":2DB8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Horas semanales"
      Height          =   195
      Index           =   2
      Left            =   5400
      TabIndex        =   55
      Top             =   480
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre "
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   52
      Top             =   480
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo "
      Height          =   195
      Index           =   0
      Left            =   180
      TabIndex        =   50
      Top             =   480
      Width           =   1155
   End
End
Attribute VB_Name = "frmHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
        'y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar
'
Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private NumRegistro As Long
Private kCampo As Integer
Private TotalReg As Long
Private PrimeraVez As Boolean


Private Sub Check1_Click(Index As Integer)
Dim V As Integer

If Check1(Index).Value = 1 Then
    Check1(Index).Tag = (Index * 5) + 3
    For V = CInt(Check1(Index).Tag) To CInt(Check1(Index).Tag) + 4
        Text1(V).Text = ""
    Next V
    'Recalculamos la horas
    V = (5 * Index) + 3
    If V > 0 Then CalculaHorasDia V, 0, False
End If
End Sub

Private Sub Check1_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub cmdAceptar_Click()
Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer

Screen.MousePointer = vbHourglass
On Error GoTo Error1

If Modo = 3 Then
    If DatosOk Then
       If InsertarRegistro = 0 Then
            MsgBox "Registro insertado.", vbInformation
            PonerModo 0
        End If
    End If
    Else
    If Modo = 4 Then
        'Modificar
        If DatosOk Then
            'Almacenamos para luego buscarlo
            Cad = Text1(0).Text
            'MODIFICAMOS
            If ModificarRegistro = 0 Then
                MsgBox "Registro modificado.", vbInformation
                PonerModo 0
            End If
            PonerModo 2
            'Hay que refresca el DAta1
            Data1.Refresh
            'Hay que volver a poner el registro donde toca
            Data1.Recordset.MoveFirst
            I = 1
            While I > 0
                If Data1.Recordset.Fields(0) = Cad Then
                    I = 0
                    Else
                        Data1.Recordset.MoveNext
                        If Data1.Recordset.EOF Then I = 0
                End If
            Wend
            If Data1.Recordset.EOF Then
                NumRegistro = TotalReg
                Data1.Recordset.MoveLast
            End If
            PonerCampos
            Label5.Caption = NumRegistro & " de " & TotalReg
        End If 'de datos ok
    End If  'modo=4
End If 'Modo=3
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
LimpiarCampos
PonerModo 0
End Sub

Private Sub BotonAnyadir()
Dim Cad As String
LimpiarCampos
'Añadiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Aceptar"
PonerModo 3
'Escondemos el navegador y ponemos insertando
DespalzamientoVisible False
Label5.Caption = "INSERTANDO"
SugerirCodigoSiguiente
Text1(0).SetFocus

End Sub

Private Sub BotonBuscar()
'Buscar
If Modo <> 1 Then
    LimpiarCampos
    Label5.Caption = "Búsqueda"
    PonerModo 1
    Text1(0).SetFocus
    Else
        HacerBusqueda
        If TotalReg = 0 Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            Text1(kCampo).SetFocus
        End If
End If
End Sub

Private Sub BotonVerTodos()
'Ver todos
LimpiarCampos
PonerModo 2
CadenaConsulta = "Select * from " & NombreTabla
PonerCadenaBusqueda
End Sub

Private Sub Desplazamiento(Index As Integer)
On Error GoTo ErrDesp
Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
        NumRegistro = 1
    Case 1
        Data1.Recordset.MovePrevious
        NumRegistro = NumRegistro - 1
        If Data1.Recordset.BOF Then
            Data1.Recordset.MoveFirst
            NumRegistro = 1
        End If
    Case 2
        Data1.Recordset.MoveNext
        NumRegistro = NumRegistro + 1
        If Data1.Recordset.EOF Then
            Data1.Recordset.MoveLast
            NumRegistro = TotalReg
        End If
    Case 3
        Data1.Recordset.MoveLast
        NumRegistro = TotalReg
End Select
PonerCampos
Exit Sub
ErrDesp:
    MuestraError Err.Number
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
Text1(0).Locked = True
DespalzamientoVisible False
Label5.Caption = "Modificar"
Text1(1).SetFocus
End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim I As Integer
Dim rDel As ADODB.Recordset

On Error GoTo Error2

'Ciertas comprobaciones
If Data1.Recordset.RecordCount = 0 Then Exit Sub
'Pregunta
Cad = "Seguro que desea eliminar de la BD el registro:"
Cad = Cad & vbCrLf & "Cod: " & Data1.Recordset.Fields(0)
Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
I = MsgBox(Cad, vbQuestion + vbYesNo)
If I = vbYes Then
    Screen.MousePointer = vbHourglass
    'Establecemos el punto de vuelta atras
    conn.BeginTrans
    Set rDel = New ADODB.Recordset
    rDel.Open "SELECT * from SubHorarios where IdHorario=" & Data1.Recordset.Fields(0), conn, adOpenDynamic, adLockOptimistic, adCmdText
    While Not rDel.EOF
        Cad = rDel.Fields(0)
        rDel.Delete
        rDel.Update
        rDel.Requery
        'rDel.MoveNext
    Wend
    rDel.Close
    '---------------------------------------------------
    'Borramos los dias festivos asociados a este horario
    Set rDel = New ADODB.Recordset
    rDel.Open "SELECT * from Festivos where IdHorario=" & Data1.Recordset.Fields(0), conn, adOpenDynamic, adLockOptimistic, adCmdText
    While Not rDel.EOF
        Cad = rDel.Fields(0)
        rDel.Delete
        rDel.Update
        rDel.Requery
        'rDel.MoveNext
    Wend
    rDel.Close
    
    
    
    'Liberamos
    Set rDel = Nothing
    'Hay que eliminar
    Data1.Recordset.Delete
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        PonerModo 0
        Else
            If NumRegistro = TotalReg Then
                    Data1.Recordset.MoveLast
                    NumRegistro = NumRegistro - 1
                    Else
                        For I = 1 To NumRegistro - 1
                            Data1.Recordset.MoveNext
                        Next I
            End If
            TotalReg = TotalReg - 1
            PonerCampos
    End If
    'Si llega hasta aqui es que todo ha ido bien
    conn.CommitTrans
End If
Screen.MousePointer = vbDefault
Exit Sub
Error2:
    If Err.Number <> 0 Then
        MsgBox Err.Number & " - " & Err.Description
        conn.RollbackTrans
        Data1.Refresh
        Data1.Recordset.MoveFirst
        PonerCampos
    End If
    Screen.MousePointer = vbDefault
End Sub


Private Sub Combo1_Click()
If PrimeraVez Then Exit Sub
On Error GoTo ErrorCombo
Label7.Caption = Combo1.List(Combo1.ListIndex)
If Text1(0).Text <> "" Then
    PonerFestivos
End If
ErrorCombo:
End Sub

Private Sub Command1_Click()
If Not (Modo = 2 Or Modo = 4) Then Exit Sub


'Añadir dia festivo
frmDiasFest.Anyo = Combo1.List(Combo1.ListIndex)
frmDiasFest.IdFestivo = 0 'Para que sepa que es nuevo
frmDiasFest.IdHor = Data1.Recordset.Fields(0)
frmDiasFest.Show vbModal
'Por si acaso refrescamos el datasource
Screen.MousePointer = vbHourglass
espera 0.1
PonerFestivos  'Probando a ver si lo hago dos veces y lo memoriza
Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
'MODIFCAR DIA FESTIVO
'Añadir dia festivo
On Error GoTo ErrAdd
If Not (Modo = 2 Or Modo = 4) Then Exit Sub


If ListView1.ListItems.Count = 0 Then Exit Sub

If Not (ListView1.SelectedItem Is Nothing) Then
        frmDiasFest.Anyo = Combo1.List(Combo1.ListIndex)
        frmDiasFest.Fecha = ListView1.SelectedItem.Text
        frmDiasFest.Descripcion = ListView1.SelectedItem.SubItems(1)
        frmDiasFest.IdFestivo = ListView1.SelectedItem.Tag
        frmDiasFest.IdHor = Data1.Recordset.Fields(0)
        frmDiasFest.Show vbModal
        'Por si acaso refrescamos el datasource
        Screen.MousePointer = vbHourglass
        espera 0.1
        PonerFestivos  'Probando a ver si lo hago dos veces y lo memoriza
        Screen.MousePointer = vbDefault
        Else
            MsgBox "No ha seleccionado ningún dia festivo.", vbCritical
End If


Exit Sub
ErrAdd:
    MuestraError Err.Number, "Add/Mod Festivo"
    Err.Clear
End Sub

Private Sub Command3_Click()
Dim RC As Byte
Dim Cad As String

If Not (Modo = 2 Or Modo = 4) Then Exit Sub


'ELIMINAR DIA FESTIVO
If ListView1.SelectedItem Is Nothing Then
    
    MsgBox "No ha seleccionado ningún dia festivo.", vbCritical
    Else
    Cad = "Seguro que desea eliminar el dia festivo: " & vbCrLf
    Cad = Cad & " Fecha: " & ListView1.SelectedItem.Text & vbCrLf
    Cad = Cad & "Desc: " & ListView1.SelectedItem.SubItems(1)
    RC = MsgBox(Cad, vbQuestion + vbYesNo, "ARIPRES")
    If RC = vbYes Then
        Cad = "Delete * from Festivos where Id=" & ListView1.SelectedItem.Tag
        conn.Execute Cad
        espera 0.1
        PonerFestivos  'Probando a ver si lo hago dos veces y lo memoriza
    End If

End If
End Sub

Private Sub Command4_Click()
    'Este boton sirve para mostrar el formulario
    'que nos permitira copiar los festivos
    ' o bien desde otro año y/o desde otro horario
    Screen.MousePointer = vbHourglass
    frmCopiaFestivos.Show vbModal
    'Por si acaso hemos modificado algo relativo a este horario
    'entonces refrescaremos
    PonerFestivos
End Sub

Private Sub Command7_Click()
    Unload Me
End Sub

Private Sub DataGrid1_DblClick()
    Command2_Click
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
PrimeraVez = True
Left = 300
Top = 200
LimpiarCampos
Data1.ConnectionString = conn
NombreTabla = "Horarios"
Ordenacion = " ORDER BY NomHorario"



    Combo1.Clear
    kCampo = Year(Now) - 4
    For TotalReg = kCampo To Year(Now) + 2
        Combo1.AddItem CStr(TotalReg)
    Next

'ASignamos un SQL al DATA1
Data1.RecordSource = "Select * from " & NombreTabla
Data1.Refresh
PonerModo 0



'Hoja 0
SSTab1.Tab = 0
PrimeraVez = False
End Sub



Private Sub LimpiarCampos()
Dim I
For I = 0 To Text1.Count - 1
    Text1(I).Text = ""
Next I
For I = 0 To Check1.Count - 1
    Check1(I).Value = 0
Next I
For I = 0 To 2
    txtAlm(I).Text = ""
    txtMer(I).Text = ""
Next I

'Combo1.ListIndex = Val(Abs(Year(Now) - 2000))

'El año actual es el combo.count -2. Como empieza en cero -3
Combo1.ListIndex = Combo1.ListCount - 3
Label7.Caption = Combo1.List(Combo1.ListIndex)
ListView1.ListItems.Clear
End Sub

Private Sub mnBuscar_Click()
BotonBuscar
End Sub

Private Sub mnEliminar_Click()
BotonEliminar
End Sub

Private Sub mnModificar_Click()
BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub



Private Sub Text1_GotFocus(Index As Integer)
    kCampo = Index
    If Index > 1 Then Exit Sub
    If Modo = 1 Then
        Text1(Index).BackColor = vbYellow
    End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Modo = 1 And Index < 2 Then
    If KeyAscii = 13 Then
        'Ha pulsado enter, luego tenemos que hacer la busqueda
        Text1(Index).BackColor = vbWhite
        BotonBuscar
    End If
Else
    Keypress KeyAscii
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Valor As Integer
Dim ind As Integer
Dim I As Integer
Dim SoloSemanales As Boolean

If Modo = 1 Then
    Text1(Index).BackColor = vbWhite
    Else
        If Modo > 2 Then   'Si insertamos o modificamos
            Select Case Index
            Case 3 To 7
                'Lunes
                Valor = 3
            Case 8 To 12
                'Martes
                Valor = 8
            Case 13 To 17
                'Miercoles
                Valor = 13
            Case 18 To 22
                'Jueves
                Valor = 18
            Case 23 To 27
                'Viernes
                Valor = 23
            Case 28 To 32
                'Sabado
                Valor = 28
            Case 33 To 37
                'Domingo
                Valor = 33
                
            Case 38 To 44
                'Si tiene valor
                Text1(Index).Text = Trim(Text1(Index).Text)
                If Not IsNumeric(Text1(Index).Text) Then
                    MsgBox "Campo numérico", vbExclamation
                    Text1(Index).Text = ""
                End If
                Exit Sub
            Case Else
                
                Valor = 0
            End Select
            
            'Cambiamos puntos por dos puntos
            If Index <> (Valor + 4) Then
                Do
                    I = InStr(1, Text1(Index).Text, ".")
                    If I > 0 Then _
                        Text1(Index).Text = Mid(Text1(Index).Text, 1, I - 1) & ":" & Mid(Text1(Index).Text, I + 1)
                Loop Until I = 0
            
                If IsDate(Text1(Index).Text) Then _
                    Text1(Index).Text = Format(Text1(Index).Text, "hh:mm")
            End If
            
            ind = (Index - 3) Mod 5
            If Index = (Valor + 4) Then
                If Not IsNumeric(Text1(Index).Text) Then
                    Text1(Index).Text = ""
                Else
                    Text1(Index).Text = TransformaPuntosComas(Text1(Index).Text)
                End If
            End If
            SoloSemanales = True
            If Valor > 0 Then
                If Text1(Valor + 4).Text = "" Then SoloSemanales = False
                CalculaHorasDia Valor, ind, SoloSemanales
            End If
             
        End If
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim c1 As String   'el nombre del campo
Dim Tipo As Long
Dim aux1

If Text1(kCampo).Text = "" Then Exit Sub
c1 = Data1.Recordset.Fields(kCampo).Name
c1 = " WHERE " & c1
Tipo = DevuelveTipo2(Data1.Recordset.Fields(kCampo).Type)
'Devolvera uno de los tipos
'   1.- Numeros
'   2.- Booleanos
'   3.- Cadenas
'   4.- Fecha
'   0.- Error leyendo los tipos de datos
' segun sea uno u otro haremos una comparacion
Select Case Tipo
Case 1
    CadB = c1 & " = " & Text1(kCampo)
Case 2
    'Vemos si la cadena tiene un Falso o False
    If InStr(1, UCase(Text1(kCampo).Text), "F") Then
        aux1 = "False"
        Else
        aux1 = "True"
    End If
    CadB = c1 & " = " & aux1
Case 3
    CadB = c1 & " like '*" & Trim(Text1(kCampo)) & "*'"
Case 4

Case 5

End Select

CadenaConsulta = "select * from " & NombreTabla & CadB & " " & Ordenacion
PonerCadenaBusqueda
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla" & NombreTabla, vbInformation
    TotalReg = 0
    Label5.Caption = ""
    PonerModo 0
    Screen.MousePointer = vbDefault
    Else
        DespalzamientoVisible True
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        TotalReg = Data1.Recordset.RecordCount
        NumRegistro = 1
        PonerCampos
End If

Data1.ConnectionString = conn
Data1.RecordSource = CadenaConsulta
Data1.Refresh
TotalReg = Data1.Recordset.RecordCount
Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim I As Integer
Dim J As Integer
Dim Cad As String
Dim Rss As ADODB.Recordset
Dim Valor As Integer

    'DATOS BÁSICOS DEL HORARIO
    Text1(0).Text = Data1.Recordset.Fields(0)
    Text1(1).Text = Data1.Recordset.Fields(1)
    Text1(2).Text = Data1.Recordset.Fields(2)
    Check2.Value = Abs(Data1.Recordset!RecuperaSabados)
    'Ponemos los datos de cada dia de la semana
    Cad = "Select * From SubHorarios Where IdHorario=" & Data1.Recordset.Fields(0)
    Cad = Cad & " ORDER BY DiaSemana"
    Set Rss = New ADODB.Recordset
    Rss.Open Cad, conn, , , adCmdText
    Valor = 0
    While Not Rss.EOF
        Valor = Rss.Fields!DiaSemana - 1
        I = (Valor * 5) + 3
        Text1(38 + Valor).Text = ""
        If Not Rss!Festivo Then
            Check1(Valor).Value = 0
            Text1(I).Text = Format(DBLet(Rss.Fields!HEntrada1), "hh:mm")
            Text1(I + 1).Text = Format(DBLet(Rss.Fields!HSalida1), "hh:mm")
            Text1(I + 2).Text = Format(DBLet(Rss.Fields!HEntrada2), "hh:mm")
            Text1(I + 3).Text = Format(DBLet(Rss.Fields!HSalida2), "hh:mm")
            Text1(I + 4).Text = DBLet(Rss.Fields!HorasDia)
        
            'Dias nomina
            If Not IsNull(Rss.Fields!DiaNomina) Then
                If Rss.Fields!DiaNomina <> 0 Then
                    If Rss.Fields!DiaNomina <> Int(Rss.Fields!DiaNomina) Then
                        Text1(38 + Valor).Text = Format(Rss.Fields!DiaNomina, "0.00")
                    Else
                        Text1(38 + Valor).Text = Int(Rss.Fields!DiaNomina)
                    End If
                End If
            End If
        
        Else
            Check1(Valor).Value = 1
            Text1(I).Text = ""
            Text1(I + 1).Text = ""
            Text1(I + 2).Text = ""
            Text1(I + 3).Text = ""
            Text1(I + 4).Text = ""
            

        End If
        Valor = Valor + 1
        Rss.MoveNext
    Wend
    Rss.Close
    Set Rss = Nothing
    'Ahora ponemos los dias festivos
    PonerFestivos
    'Ponemos los datos de almuerzo y merienda
    PonerDatosAlmuerzoMerienda
    Label5.Caption = NumRegistro & " de " & TotalReg
End Sub

Private Sub PonerModo(Kmodo As Integer)
Dim I As Integer
Dim B As Boolean

If Modo = 1 Then
    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
End If
Modo = Kmodo
DespalzamientoVisible (Kmodo = 2)
cmdAceptar.Visible = (Kmodo >= 3)
cmdCancelar.Visible = (Kmodo >= 3)
Toolbar1.Buttons(6).Enabled = (Kmodo < 3)
Toolbar1.Buttons(7).Enabled = (Kmodo = 2)
Toolbar1.Buttons(8).Enabled = (Kmodo = 2)
Toolbar1.Buttons(1).Enabled = (Kmodo < 3)
Toolbar1.Buttons(2).Enabled = (Kmodo < 3)
If Kmodo = 0 Then _
    Label5.Caption = ""
B = (Modo = 2) Or Modo = 0

For I = 0 To Text1.Count - 1
    Text1(I).Locked = B
Next I
Check2.Enabled = Not B
Frame1.Enabled = Not B
B = Modo > 2
For I = 0 To 6
    Check1(I).Enabled = B
Next I
Command4.Visible = (Kmodo = 4)
End Sub


Private Function DatosOk() As Boolean
Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim Valor As Integer


DatosOk = False
'Haremos las comprobaciones necesarias de los campos
'Al final todo esta correcto
If Text1(1).Text = "" Then
    MsgBox "El nombre del horario no puede estar en blanco.", vbExclamation
    Exit Function
End If

For I = 0 To 6
    If Check1(I).Value = 0 Then
        J = (I * 5) + 3
        'Entrada 1
        If Not FechaOk(Text1(J).Text) Then
            MsgBox "La hora de entrada 1 es incorrecta.", vbExclamation
            Exit Function
        End If
        'Salida 1
        If Not FechaOk(Text1(J + 1).Text) Then
            MsgBox "La hora de salida 1 es incorrecta.", vbExclamation
            Exit Function
        End If
        'Entrada 2
        If Not FechaOk(Text1(J + 2).Text) Then
            MsgBox "La hora de entrada 2 es incorrecta.", vbExclamation
            Exit Function
        End If
        'Salida 2
        If Not FechaOk(Text1(J + 3).Text) Then
            MsgBox "La hora de salida 2 es incorrecta.", vbExclamation
            Exit Function
        End If
    End If
    
Next I

'Comprobamos los datos de los dtos
'ALMUERZO
If txtAlm(0).Text = "" Then txtAlm(0).Text = 0
If txtAlm(0).Text = "0" Then
    txtAlm(1).Text = ""
    txtAlm(2).Text = ""
    
    Else
        'Comprobamos su valor
        If Not IsNumeric(txtAlm(0).Text) Then
            MsgBox "El descuento de empleados por el almuerzo debe de ser numérico.", vbExclamation
            Exit Function
        End If
        'Llegados a este punto tiene dto. Comprobaremos la hora
        If txtAlm(2).Text = "" Then
            MsgBox "Ponga una hora almuerzo.", vbExclamation
            Exit Function
        End If
        If Not IsDate(txtAlm(2).Text) Then
            MsgBox "Hora almuerzo incorrecta. Formato fecha incorrecto.", vbExclamation
            Exit Function
        End If
End If

'MERIENDA
If txtMer(0).Text = "" Then txtMer(0).Text = 0
If txtMer(0).Text = "0" Then
    txtMer(1).Text = ""
    txtMer(2).Text = ""
    
    Else
        'Comprobamos su valor
        If Not IsNumeric(txtMer(0).Text) Then
            MsgBox "El descuento de empleados por la merienda debe de ser numérico.", vbExclamation
            Exit Function
        End If
        'Llegados a este punto tiene dto. Comprobaremos la hora
        If txtMer(2).Text = "" Then
            MsgBox "Ponga una hora merienda.", vbExclamation
            Exit Function
        End If
        If Not IsDate(txtMer(2).Text) Then
            MsgBox "Hora merienda incorrecta. Formato fecha incorrecto.", vbExclamation
            Exit Function
        End If
End If


For I = 38 To 44
    Text1(I).Text = Trim(Text1(I).Text)
    J = I - 38
    Valor = Abs(Check1(J).Value)
    If Text1(I).Text <> "" Then
        'Tiene puesto datos enb eñl text
        If Valor = 0 Then
            If Not IsNumeric(Text1(I).Text) Then
                MsgBox "Dias / Nomina   debe ser numérico.", vbExclamation
                Exit Function
            End If
            If Val(Text1(I).Text) > 1 Then
                MsgBox "Valor maximo Dia/nomina es 1", vbExclamation
                Exit Function
            End If
    
        End If
    Else
        If Valor = 0 Then
            MsgBox "Campo Dia/nomina requerido", vbExclamation
            Exit Function
        End If
    End If
Next I
    


DatosOk = True
End Function

Private Function DevuelveFecha(Texto As String) As Variant
If Texto = "" Then
    DevuelveFecha = Null
    Else
        DevuelveFecha = Texto
End If
End Function

'Si tiene valor tiene que ser fecha
Private Function FechaOk(Texto As String) As Boolean
FechaOk = True
If Texto <> "" Then FechaOk = IsDate(Texto)
End Function


Private Sub SugerirCodigoSiguiente()
Dim Cad
Dim RS
'Sugeriremos el codigo siguiente.
'Obviamente depende en TOTAL medida de que tabla estemos trabajando
Cad = "Select Max(IdHorario) from Horarios"

Text1(0).Text = 1
Set RS = New ADODB.Recordset
RS.Open Cad, conn, , , adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then
        Text1(0).Text = RS.Fields(0) + 1
    End If
End If
RS.Close
End Sub



Private Function InsertarRegistro() As Byte
Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim k As Integer
Dim con As Integer
Dim Contador As Integer

On Error GoTo ErrInsertando
    conn.BeginTrans
    Set RS = New ADODB.Recordset
    'Insertamos primero en HORARIO
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open "Horarios", conn, , , adCmdTable
    RS.AddNew
    RS!IdHorario = Text1(0).Text
    RS!NomHorario = Text1(1).Text
    RS!TotalHoras = DevNumero(Text1(2).Text) 'CalculaHoras
    
    'Insertamos datos almuerzo
    RS!DtoAlm = DevNumero(txtAlm(0).Text)
    RS!HoraDtoAlm = DevuelveFecha(txtAlm(2).Text)
    'Merienda
    RS!DtoMer = DevNumero(txtMer(0).Text)
    RS!HoraDtoMer = DevuelveFecha(txtMer(2).Text)
    RS!RecuperaSabados = (Check2.Value = 1)
    
    '--------------------
    RS.Update
    RS.Close
    '--------------------------------------
    'Ahora insertamos los horarios por dias
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open "SubHorarios", conn, , , adCmdTable
    If Not RS.EOF Then
        RS.MoveLast
        Contador = RS!idsubhorario + 1
        Else
            Contador = 1
    End If
        
    For I = 0 To 6
        RS.AddNew
        '---------------------
        RS!idsubhorario = Contador
        RS!DiaSemana = I + 1
        RS!Festivo = Check1(I).Value = 1
        If Check1(I).Value = 0 Then
            J = (I * 5) + 3
            'Introducimos los subhorarios para cada dia
            RS.Fields!HEntrada1 = DevuelveFecha(Text1(J).Text)
            RS.Fields!HSalida1 = DevuelveFecha(Text1(J + 1).Text)
            RS.Fields!HEntrada2 = DevuelveFecha(Text1(J + 2).Text)
            RS.Fields!HSalida2 = DevuelveFecha(Text1(J + 3).Text)
            RS.Fields!HorasDia = DevNumero(Text1(J + 4).Text)
            con = 0
            For k = J To J + 3
                If Text1(k).Text <> "" Then con = con + 1
            Next k
            Else
                con = 0
        End If
        
        k = I + 38
        If Text1(k).Text = "" Then
            RS!DiaNomina = 0
        Else
            RS!DiaNomina = TransformaComasPuntos(Text1(k).Text)
        End If
            
        
        RS!IdHorario = Text1(0).Text
        RS!N_Tikadas = con
        Contador = Contador + 1
        RS.Update
    Next I
    conn.CommitTrans
    Data1.Refresh
    Set RS = Nothing
    InsertarRegistro = 0
    Exit Function
    
ErrInsertando:
    MsgBox "Error: " & Err.Number & " - " & Err.Description
    conn.RollbackTrans
    InsertarRegistro = 1
End Function

Private Function DevNumero(Texto As String) As Variant
If IsNumeric(Texto) Then
    DevNumero = CSng(Texto)
    Else
    DevNumero = 0
End If

End Function

Private Sub CalculaHorasDia(V As Integer, vInd As Integer, SemanalesSolo As Boolean)
Dim I
Dim k As Integer
Dim T1 As Single
Dim T2 As Single
Dim v1(3) As Single


If Not SemanalesSolo Then
    For k = 0 To 3
        If Text1(V + k).Text <> "" Then
            If IsDate(Text1(V + k).Text) Then
                  v1(k) = DevuelveValorHora(Text1(V + k).Text)
            Else
                v1(k) = 0
            End If
        Else
            v1(k) = 0
        End If
    Next k
    
    
    'Ya tenemos en cada tag
    If v1(0) <> 0 And v1(1) <> 0 Then
        T1 = v1(1) - v1(0)
        Else
            T1 = 0
    End If
    If v1(3) <> 0 And v1(2) <> 0 Then
        T2 = v1(3) - v1(2)
        Else
            T2 = 0
    End If
    'Las horas totales del dia son la suma de ambas
    Text1(V + 4).Text = T1 + T2

End If
    
'Recalcularemos las horas totales semanales

T2 = 0
For I = 0 To 6
    If Text1(7 + (5 * I)).Text <> "" Then
        T1 = CSng(Text1(7 + (5 * I)).Text)
        T2 = T2 + T1
    End If
Next I
If T2 > 0 Then
    Text1(2).Text = T2
    Else
    Text1(2).Text = ""
End If
End Sub


Private Function ModificarRegistro() As Byte
Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer
Dim J As Integer
Dim idHORA As Integer
Dim Contador As Integer
Dim con, k


On Error GoTo ErrInsertando
    idHORA = Data1.Recordset!IdHorario
    conn.BeginTrans
    Set RS = New ADODB.Recordset
    'Insertamos primero en HORARIO
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open "Select * from Horarios where IdHorario=" & idHORA, conn, , , adCmdText
    
    
    RS!NomHorario = Text1(1).Text
    RS!TotalHoras = DevNumero(Text1(2).Text) 'CalculaHoras
    
    'Mod datos almuerzo
    RS!DtoAlm = DevNumero(txtAlm(0).Text)
    RS!HoraDtoAlm = DevuelveFecha(txtAlm(2).Text)
    'Merienda
    RS!DtoMer = DevNumero(txtMer(0).Text)
    RS!HoraDtoMer = DevuelveFecha(txtMer(2).Text)
    RS!RecuperaSabados = (Check2.Value = 1)
    '--------------------
    RS.Update
    RS.Close
    '--------------------------------------
    'Ahora modificamos los horarios por dias
    'Borramos todos los horarios para ese dia
    RS.Open "Delete * from SubHorarios where idHorario=" & idHORA, conn, , , adCmdText
    'Insertamos Otra vez los registros
    'Ahora insertamos los horarios por dias
    Set RS = New ADODB.Recordset
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open "SubHorarios", conn, , , adCmdTable
    If Not RS.EOF Then
        RS.MoveLast
        Contador = RS!idsubhorario + 1
        Else
            Contador = 1
    End If
        
    For I = 0 To 6
        RS.AddNew
        '---------------------
        RS!idsubhorario = Contador
        RS!DiaSemana = I + 1
        RS!Festivo = Check1(I).Value = 1
        If Check1(I).Value = 0 Then
            J = (I * 5) + 3
            'Introducimos los subhorarios para cada dia
            RS.Fields!HEntrada1 = DevuelveFecha(Text1(J).Text)
            RS.Fields!HSalida1 = DevuelveFecha(Text1(J + 1).Text)
            RS.Fields!HEntrada2 = DevuelveFecha(Text1(J + 2).Text)
            RS.Fields!HSalida2 = DevuelveFecha(Text1(J + 3).Text)
            RS.Fields!HorasDia = DevNumero(Text1(J + 4).Text)
            con = 0
            For k = J To J + 3
                If Text1(k).Text <> "" Then con = con + 1
            Next k
            Else
                con = 0
        End If
        k = I + 38
        If Text1(k).Text = "" Then
            RS!DiaNomina = 0
        Else
            RS!DiaNomina = TransformaPuntosComas(Text1(k).Text)
        End If
        RS!IdHorario = Text1(0).Text
        RS!N_Tikadas = con
        Contador = Contador + 1
        RS.Update
    Next I
    
    
    conn.CommitTrans
    Data1.Refresh
    Set RS = Nothing
    ModificarRegistro = 0
    Exit Function
    
ErrInsertando:
    MsgBox "Error: " & Err.Number & " - " & Err.Description
    conn.RollbackTrans
    ModificarRegistro = 1
End Function


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index > 5 And Button.Index < 9 Then
        If vUsu.Nivel > 2 Then 'no puede consulta
            MsgBox "No tiene autorización para realizar cambios", vbExclamation
            Exit Sub
        End If
    End If
Select Case Button.Index
Case 1
    BotonBuscar
Case 2
    BotonVerTodos
Case 6
    BotonAnyadir
Case 7
    BotonModificar
Case 8
    BotonEliminar
Case 14 To 17
    Desplazamiento (Button.Index - 14)
Case 19
    'Imprimir el listado
    With frmImprimir
        .Opcion = 107
        .NumeroParametros = 0
        .OtrosParametros = ""
        .FormulaSeleccion = ""
        .Show vbModal
    End With
Case Else

End Select
End Sub

Private Sub DespalzamientoVisible(bol As Boolean)
Dim I
For I = 14 To 17
    Toolbar1.Buttons(I).Visible = bol
Next I
End Sub


Private Sub PonerFestivos()
Dim Cad As String
Dim mA As String
Dim iH As String
Dim RS As ADODB.Recordset
Dim itmX As ListItem

On Error GoTo EPonerFestivos

Label8.Caption = ""
If Data1.Recordset.EOF Then
    mA = "-1"
    iH = "-1"
    Else
    If Data1.Recordset.RecordCount > 0 Then
            mA = Combo1.List(Combo1.ListIndex)
            iH = Data1.Recordset.Fields(0)
            Else
            iH = "-1"
            mA = "-1"
    End If
End If
    
    ListView1.ListItems.Clear
    Cad = "Select Fecha,Descripcion,Id from Festivos where IdHorario=" & iH
    Cad = Cad & " AND Anyo =" & mA
    Cad = Cad & " ORDER BY Fecha"
    
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenKeyset, adLockOptimistic, adCmdText
    While Not RS.EOF
    'Fijamos el ancho de columna
        Set itmX = ListView1.ListItems.Add
        itmX.Text = Format(RS!Fecha, "dd/mm/yyyy")
        itmX.SubItems(1) = RS!Descripcion
        itmX.Tag = RS!Id
        itmX.SmallIcon = 11
        RS.MoveNext
    Wend
    If RS.RecordCount > 0 Then
        Label8.Caption = "Dias festivos: " & RS.RecordCount
        Else
        Label8.Caption = ""
    End If
    RS.Close
    Set RS = Nothing
    
Exit Sub
EPonerFestivos:
    MuestraError Err.Number, "Poner dias festivos"
    Set RS = Nothing
End Sub

Private Sub PonerDatosAlmuerzoMerienda()
Dim Cad
Dim I As Single

On Error GoTo EPonerDatosAlmuerzoMerienda
'Limpiamos
txtAlm(0).Text = "": txtAlm(1).Text = "": txtAlm(2).Text = ""
txtMer(0).Text = "": txtMer(1).Text = "": txtMer(2).Text = ""
'Ponemos valores si procede
If Not Data1.Recordset.EOF Then
    'ALMUERZO
    I = DBLet(Data1.Recordset!DtoAlm, "N")
    If I > 0 Then
        txtAlm(0).Text = I
        If I <> 0 Then
            txtAlm(1).Text = DevuelveHora(Data1.Recordset!DtoAlm)
            txtAlm(2).Text = Format(Data1.Recordset!HoraDtoAlm, "hh:mm")
        End If
    End If
    I = DBLet(Data1.Recordset!DtoMer, "N")
    If I > 0 Then
        txtMer(0).Text = I
        If I <> 0 Then
            txtMer(1).Text = DevuelveHora(Data1.Recordset!DtoMer)
            txtMer(2).Text = Format(Data1.Recordset!HoraDtoMer, "hh:mm")
        End If
    End If
End If
Exit Sub
EPonerDatosAlmuerzoMerienda:
    MuestraError Err.Number, "Poner datos almuerzo/merienda"
End Sub

Private Sub txtAlm_GotFocus(Index As Integer)
txtAlm(Index).SelStart = 0
txtAlm(Index).SelLength = Len(txtAlm(Index))
End Sub

Private Sub txtAlm_LostFocus(Index As Integer)
Dim I As Integer

If Index > 0 Then
    I = InStr(1, txtAlm(Index).Text, ".")
    If I > 0 Then txtAlm(Index).Text = Format(Mid(txtAlm(Index).Text, 1, I - 1) & ":" & Mid(txtAlm(Index).Text, I + 1), "hh:mm")
    Else
        I = InStr(1, txtAlm(0).Text, ".")
        If I > 0 Then txtAlm(0).Text = Mid(txtAlm(0).Text, 1, I - 1) & "," & Mid(txtAlm(Index).Text, I + 1)
End If

If Modo > 2 Then
    If Index = 0 Then
        If Not IsNumeric(txtAlm(0).Text) Then
            txtAlm(0).Text = ""
            txtAlm(1).Text = ""
            Else
                txtAlm(1).Text = Format(DevuelveHora(CSng(txtAlm(0).Text)), "hh:mm")
        End If

        'ELSE de index=0
        Else
            If Index = 1 Then
                If Not IsDate(txtAlm(1).Text) Then
                    txtAlm(0).Text = ""
                    txtAlm(1).Text = ""
                    Else
                        txtAlm(0).Text = DevuelveValorHora(CDate(txtAlm(1).Text))
                End If
            End If
        End If
End If
End Sub

Private Sub txtMer_GotFocus(Index As Integer)
txtMer(Index).SelStart = 0
txtMer(Index).SelLength = Len(txtMer(Index).Text)
End Sub

Private Sub txtMer_LostFocus(Index As Integer)
Dim I As Integer

If Index > 0 Then
    I = InStr(1, txtMer(Index).Text, ".")
    If I > 0 Then txtMer(Index).Text = Format(Mid(txtMer(Index).Text, 1, I - 1) & ":" & Mid(txtMer(Index).Text, I + 1), "hh:mm")
    Else
        I = InStr(1, txtMer(0).Text, ".")
        If I > 0 Then txtMer(0).Text = Mid(txtMer(0).Text, 1, I - 1) & "," & Mid(txtMer(Index).Text, I + 1)
End If

If Modo > 2 Then
    If Index = 0 Then
        If Not IsNumeric(txtMer(0).Text) Then
            txtMer(0).Text = ""
            txtMer(1).Text = ""
            Else
                txtMer(1).Text = DevuelveHora(CSng(txtMer(0).Text))
        End If

        'ELSE de index=0
        Else
            If Index = 1 Then
                If Not IsDate(txtMer(1).Text) Then
                    txtMer(0).Text = ""
                    txtMer(1).Text = ""
                    Else
                        txtMer(0).Text = DevuelveValorHora(CDate(txtMer(1).Text))
                End If
            End If
        End If
End If
End Sub
