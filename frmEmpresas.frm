VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmpresas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EMPRESAS"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   75
   ClientWidth     =   10830
   Icon            =   "frmEmpresas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   10830
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   20
      Left            =   120
      TabIndex        =   66
      Text            =   "Text1"
      Top             =   4200
      Visible         =   0   'False
      Width           =   615
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Datos básicos"
      TabPicture(0)   =   "frmEmpresas.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Check1(5)"
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(2)=   "Text1(0)"
      Tab(0).Control(3)=   "Text1(1)"
      Tab(0).Control(4)=   "Text1(2)"
      Tab(0).Control(5)=   "Text1(3)"
      Tab(0).Control(6)=   "Text1(4)"
      Tab(0).Control(7)=   "Text1(5)"
      Tab(0).Control(8)=   "Text1(6)"
      Tab(0).Control(9)=   "Text1(15)"
      Tab(0).Control(10)=   "Label1(0)"
      Tab(0).Control(11)=   "Label1(1)"
      Tab(0).Control(12)=   "Label1(2)"
      Tab(0).Control(13)=   "Label1(3)"
      Tab(0).Control(14)=   "Label1(4)"
      Tab(0).Control(15)=   "Label1(5)"
      Tab(0).Control(16)=   "Label1(6)"
      Tab(0).Control(17)=   "Label1(15)"
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "Ajustes"
      TabPicture(1)   =   "frmEmpresas.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Text1(25)"
      Tab(1).Control(1)=   "Text1(7)"
      Tab(1).Control(2)=   "Text1(8)"
      Tab(1).Control(3)=   "Text3(0)"
      Tab(1).Control(4)=   "Text3(1)"
      Tab(1).Control(5)=   "Text1(16)"
      Tab(1).Control(6)=   "Text1(17)"
      Tab(1).Control(7)=   "Text1(18)"
      Tab(1).Control(8)=   "Combo5"
      Tab(1).Control(9)=   "Text1(19)"
      Tab(1).Control(10)=   "Label1(21)"
      Tab(1).Control(11)=   "Label1(7)"
      Tab(1).Control(12)=   "Label1(8)"
      Tab(1).Control(13)=   "Label1(16)"
      Tab(1).Control(14)=   "Label1(17)"
      Tab(1).Control(15)=   "Label1(18)"
      Tab(1).Control(16)=   "Label1(19)"
      Tab(1).Control(17)=   "Label1(20)"
      Tab(1).ControlCount=   18
      TabCaption(2)   =   "Incidencias"
      TabPicture(2)   =   "frmEmpresas.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Text1(9)"
      Tab(2).Control(1)=   "Text1(10)"
      Tab(2).Control(2)=   "Text1(11)"
      Tab(2).Control(3)=   "Text1(12)"
      Tab(2).Control(4)=   "Text1(13)"
      Tab(2).Control(5)=   "Text1(14)"
      Tab(2).Control(6)=   "Text2(0)"
      Tab(2).Control(7)=   "Text2(1)"
      Tab(2).Control(8)=   "Text2(2)"
      Tab(2).Control(9)=   "Text2(3)"
      Tab(2).Control(10)=   "Text2(4)"
      Tab(2).Control(11)=   "Text2(5)"
      Tab(2).Control(12)=   "Label1(9)"
      Tab(2).Control(13)=   "Label1(10)"
      Tab(2).Control(14)=   "Label1(11)"
      Tab(2).Control(15)=   "Label1(12)"
      Tab(2).Control(16)=   "Label1(13)"
      Tab(2).Control(17)=   "Label1(14)"
      Tab(2).Control(18)=   "Image1(0)"
      Tab(2).Control(19)=   "Image1(1)"
      Tab(2).Control(20)=   "Image1(2)"
      Tab(2).Control(21)=   "Image1(3)"
      Tab(2).Control(22)=   "Image1(4)"
      Tab(2).Control(23)=   "Image1(5)"
      Tab(2).ControlCount=   24
      TabCaption(3)   =   "Parámetros"
      TabPicture(3)   =   "frmEmpresas.frx":035E
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Label3"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Label4"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Label1(22)"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Label1(23)"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Text1(26)"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).Control(5)=   "Text1(27)"
      Tab(3).Control(5).Enabled=   0   'False
      Tab(3).Control(6)=   "Text1(28)"
      Tab(3).Control(6).Enabled=   0   'False
      Tab(3).Control(7)=   "Check1(0)"
      Tab(3).Control(7).Enabled=   0   'False
      Tab(3).Control(8)=   "Check1(1)"
      Tab(3).Control(8).Enabled=   0   'False
      Tab(3).Control(9)=   "Check1(2)"
      Tab(3).Control(9).Enabled=   0   'False
      Tab(3).Control(10)=   "Text1(29)"
      Tab(3).Control(10).Enabled=   0   'False
      Tab(3).Control(11)=   "Check1(4)"
      Tab(3).Control(11).Enabled=   0   'False
      Tab(3).Control(12)=   "Check1(3)"
      Tab(3).Control(12).Enabled=   0   'False
      Tab(3).Control(13)=   "Text1(30)"
      Tab(3).Control(13).Enabled=   0   'False
      Tab(3).Control(14)=   "Text1(31)"
      Tab(3).Control(14).Enabled=   0   'False
      Tab(3).Control(15)=   "Text1(32)"
      Tab(3).Control(15).Enabled=   0   'False
      Tab(3).Control(16)=   "FrameXpass"
      Tab(3).Control(16).Enabled=   0   'False
      Tab(3).ControlCount=   17
      Begin VB.CheckBox Check1 
         Caption         =   "SEPA XML"
         Height          =   255
         Index           =   5
         Left            =   -68280
         TabIndex        =   93
         Top             =   2640
         Width           =   1335
      End
      Begin VB.Frame FrameXpass 
         Caption         =   "X-Pass"
         Height          =   615
         Left            =   240
         TabIndex        =   84
         Top             =   2160
         Width           =   9615
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   38
            Left            =   8640
            MaxLength       =   250
            TabIndex        =   91
            Tag             =   "Ult reg|T|S|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   795
         End
         Begin VB.TextBox Text1 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Index           =   37
            Left            =   6240
            MaxLength       =   250
            PasswordChar    =   "*"
            TabIndex        =   89
            Tag             =   "Contraseña  X-Pass|T|S|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   36
            Left            =   3600
            MaxLength       =   250
            TabIndex        =   87
            Tag             =   "Usuario X-Pass|T|S|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   35
            Left            =   840
            MaxLength       =   40
            TabIndex        =   85
            Tag             =   "Server X-Pass|T|S|||"
            Text            =   "Text1"
            Top             =   240
            Width           =   1755
         End
         Begin VB.Label Label1 
            Caption         =   "Ult. registro"
            Height          =   195
            Index           =   28
            Left            =   7680
            TabIndex        =   92
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Contraseña"
            Height          =   195
            Index           =   27
            Left            =   5280
            TabIndex        =   90
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Usuario"
            Height          =   195
            Index           =   26
            Left            =   2880
            TabIndex        =   88
            Top             =   240
            Width           =   1305
         End
         Begin VB.Label Label1 
            Caption         =   "Server"
            Height          =   195
            Index           =   25
            Left            =   240
            TabIndex        =   86
            Top             =   240
            Width           =   1305
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   32
         Left            =   1560
         MaxLength       =   250
         TabIndex        =   81
         Tag             =   "Costes|T|S|||"
         Text            =   "Text1"
         Top             =   2880
         Width           =   7875
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   31
         Left            =   1560
         MaxLength       =   250
         TabIndex        =   79
         Tag             =   "Huellas|T|S|||"
         Text            =   "Text1"
         Top             =   2400
         Width           =   7875
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   30
         Left            =   8880
         TabIndex        =   78
         Tag             =   "Horario Nocturno|N|N|0|1|"
         Text            =   "Text1"
         Top             =   600
         Width           =   555
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Horas compensables (sem 40): Extra"
         Height          =   255
         Index           =   3
         Left            =   3960
         TabIndex        =   76
         Top             =   1200
         Width           =   3015
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Calculo horas nom. automática"
         Height          =   255
         Index           =   4
         Left            =   3960
         TabIndex        =   75
         Top             =   1800
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   29
         Left            =   5280
         TabIndex        =   73
         Tag             =   "IRPF empresa|N|S|0|100|"
         Text            =   "Text1"
         Top             =   600
         Width           =   675
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Abonos separados en anticpos (HN/HC)"
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   72
         Top             =   1800
         Width           =   3255
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Aplica Antiguedad HC"
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   71
         Top             =   1200
         Width           =   2535
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Aplica Antiguedad HN"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   70
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   28
         Left            =   8760
         TabIndex        =   69
         Tag             =   "3|N|N|0||"
         Text            =   "Text1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   27
         Left            =   7920
         TabIndex        =   68
         Tag             =   "2|N|N|0||"
         Text            =   "Text1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   26
         Left            =   7320
         TabIndex        =   67
         Tag             =   "1|N|N|0||"
         Text            =   "Text1"
         Top             =   1560
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   25
         Left            =   -66120
         TabIndex        =   64
         Tag             =   "Anulacion|N|N|||"
         Text            =   "Text1"
         Top             =   1875
         Width           =   615
      End
      Begin VB.Frame Frame1 
         Caption         =   "Cuenta bancaria"
         Height          =   855
         Left            =   -74760
         TabIndex        =   57
         Top             =   2280
         Width           =   6015
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   34
            Left            =   4680
            MaxLength       =   3
            TabIndex        =   63
            Tag             =   "sufio|T|S|||"
            Text            =   "999"
            Top             =   360
            Width           =   555
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   33
            Left            =   120
            MaxLength       =   4
            TabIndex        =   58
            Tag             =   "iban|T|S|||"
            Text            =   "9999"
            Top             =   360
            Width           =   555
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   24
            Left            =   2760
            MaxLength       =   10
            TabIndex        =   62
            Tag             =   "Cuenta|T|S|||"
            Text            =   "9999999999"
            Top             =   360
            Width           =   1155
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   23
            Left            =   2280
            MaxLength       =   2
            TabIndex        =   61
            Tag             =   "CodControl|T|S|||"
            Text            =   "99"
            Top             =   360
            Width           =   315
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   22
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   60
            Tag             =   "Sucursal|T|S|||"
            Text            =   "9999"
            Top             =   360
            Width           =   555
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
            Height          =   285
            Index           =   21
            Left            =   840
            MaxLength       =   4
            TabIndex        =   59
            Tag             =   "Entidad|T|S|||"
            Text            =   "9999"
            Top             =   360
            Width           =   555
         End
         Begin VB.Label Label1 
            Caption         =   "Sufijo"
            Height          =   195
            Index           =   24
            Left            =   4080
            TabIndex        =   83
            Top             =   360
            Width           =   705
         End
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   9
         Left            =   -72540
         TabIndex        =   50
         Tag             =   "#|N|N|0||"
         Text            =   "Text1"
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   10
         Left            =   -67440
         TabIndex        =   49
         Tag             =   "#|N|N|0||"
         Text            =   "Text1"
         Top             =   1080
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   11
         Left            =   -72540
         TabIndex        =   48
         Tag             =   "#|N|N|0||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   12
         Left            =   -67440
         TabIndex        =   47
         Tag             =   "#|N|N|0||"
         Text            =   "Text1"
         Top             =   1560
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   13
         Left            =   -72540
         TabIndex        =   46
         Tag             =   "#|N|N|0||"
         Text            =   "Text1"
         Top             =   2040
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   14
         Left            =   -67440
         TabIndex        =   45
         Tag             =   "#|N|N|0||"
         Text            =   "Text1"
         Top             =   2040
         Width           =   555
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   -71880
         TabIndex        =   44
         Text            =   "Text2"
         Top             =   1080
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   -66840
         TabIndex        =   43
         Text            =   "Text2"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   2
         Left            =   -71880
         TabIndex        =   42
         Text            =   "Text2"
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   -66840
         TabIndex        =   41
         Text            =   "Text2"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   4
         Left            =   -71880
         TabIndex        =   40
         Text            =   "Text2"
         Top             =   2040
         Width           =   2055
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   285
         Index           =   5
         Left            =   -66840
         TabIndex        =   39
         Text            =   "Text2"
         Top             =   2040
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   7
         Left            =   -72120
         TabIndex        =   31
         Tag             =   "#|N|S|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   8
         Left            =   -67200
         TabIndex        =   30
         Tag             =   "#|N|S|||"
         Text            =   "Text1"
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   0
         Left            =   -71400
         TabIndex        =   29
         Text            =   "Text3"
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Height          =   285
         Index           =   1
         Left            =   -66480
         TabIndex        =   28
         Text            =   "Text3"
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   16
         Left            =   -72600
         TabIndex        =   27
         Tag             =   "#|N|S|||"
         Text            =   "Text1"
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   17
         Left            =   -72600
         TabIndex        =   26
         Tag             =   "#|N|N|||"
         Text            =   "Text1"
         Top             =   1875
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   18
         Left            =   -70320
         TabIndex        =   25
         Tag             =   "#|N|N|||"
         Text            =   "Text1"
         Top             =   1875
         Width           =   615
      End
      Begin VB.ComboBox Combo5 
         Height          =   315
         ItemData        =   "frmEmpresas.frx":037A
         Left            =   -72780
         List            =   "frmEmpresas.frx":038A
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   2640
         Width           =   2835
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   19
         Left            =   -68220
         TabIndex        =   23
         Tag             =   "#|N|S|||"
         Text            =   "Text1"
         Top             =   2640
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   0
         Left            =   -73860
         TabIndex        =   14
         Tag             =   "Código|N|S|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   1155
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   1
         Left            =   -71160
         MaxLength       =   50
         TabIndex        =   13
         Tag             =   "Nombre|T|N|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   2
         Left            =   -73860
         MaxLength       =   50
         TabIndex        =   12
         Tag             =   "Direccion|T|S|||"
         Text            =   "Text1"
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   3
         Left            =   -73860
         MaxLength       =   50
         TabIndex        =   11
         Tag             =   "Poblacion|T|S|||"
         Text            =   "Text1"
         Top             =   1680
         Width           =   3315
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   4
         Left            =   -69540
         MaxLength       =   50
         TabIndex        =   10
         Tag             =   "Provincia|T|N|||"
         Text            =   "Text1"
         Top             =   1680
         Width           =   2715
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   5
         Left            =   -65940
         TabIndex        =   9
         Tag             =   "Teléfono|T|S|||"
         Text            =   "Text1"
         Top             =   1680
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Index           =   6
         Left            =   -68280
         TabIndex        =   8
         Tag             =   "Código Postal|N|S|||"
         Text            =   "Text1"
         Top             =   1200
         Width           =   1035
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         Height          =   285
         Index           =   15
         Left            =   -66240
         MaxLength       =   15
         TabIndex        =   7
         Tag             =   "CIF|T|S|||"
         Text            =   "Text1"
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Path fich costes"
         Height          =   195
         Index           =   23
         Left            =   240
         TabIndex        =   82
         Top             =   2880
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Dir huellas"
         Height          =   195
         Index           =   22
         Left            =   240
         TabIndex        =   80
         Top             =   2400
         Width           =   1065
      End
      Begin VB.Label Label4 
         Caption         =   "Horario nocturno"
         Height          =   255
         Left            =   7440
         TabIndex        =   77
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "IRPF empresa"
         Height          =   255
         Left            =   3960
         TabIndex        =   74
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Anulacion marcajes repetidos(min)"
         Height          =   195
         Index           =   21
         Left            =   -68640
         TabIndex        =   65
         Top             =   1920
         Width           =   2400
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencia hora extra"
         Height          =   195
         Index           =   9
         Left            =   -74760
         TabIndex        =   56
         Top             =   1080
         Width           =   1485
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencia retraso"
         Height          =   195
         Index           =   10
         Left            =   -69720
         TabIndex        =   55
         Top             =   1140
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencias marcaje"
         Height          =   195
         Index           =   11
         Left            =   -74760
         TabIndex        =   54
         Top             =   1560
         Width           =   1410
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencia por BAJA"
         Height          =   195
         Index           =   12
         Left            =   -69720
         TabIndex        =   53
         Top             =   1620
         Width           =   1605
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencia tarjeta erronea"
         Height          =   195
         Index           =   13
         Left            =   -74760
         TabIndex        =   52
         Top             =   2040
         Width           =   1800
      End
      Begin VB.Label Label1 
         Caption         =   "Incidencia hora en exceso"
         Height          =   195
         Index           =   14
         Left            =   -69720
         TabIndex        =   51
         Top             =   2100
         Width           =   1920
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   -72960
         Picture         =   "frmEmpresas.frx":03DA
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   -67800
         Picture         =   "frmEmpresas.frx":04DC
         Top             =   1080
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   2
         Left            =   -72960
         Picture         =   "frmEmpresas.frx":05DE
         Top             =   1560
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   3
         Left            =   -67800
         Picture         =   "frmEmpresas.frx":06E0
         Top             =   1620
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   4
         Left            =   -72900
         Picture         =   "frmEmpresas.frx":07E2
         Top             =   2040
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   5
         Left            =   -67800
         Picture         =   "frmEmpresas.frx":08E4
         Top             =   2100
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Máximo Retraso (decimal/Sexage.)"
         Height          =   195
         Index           =   7
         Left            =   -74760
         TabIndex        =   38
         Top             =   720
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Maximo exceso  (decimal/Sexage.)"
         Height          =   195
         Index           =   8
         Left            =   -69840
         TabIndex        =   37
         Top             =   720
         Width           =   2475
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ajuste horas redondeo (min)"
         Height          =   195
         Index           =   16
         Left            =   -74760
         TabIndex        =   36
         Top             =   1200
         Width           =   1965
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ajuste entrada (min)"
         Height          =   195
         Index           =   17
         Left            =   -74760
         TabIndex        =   35
         Top             =   1920
         Width           =   1395
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Ajuste salida (min)"
         Height          =   195
         Index           =   18
         Left            =   -71640
         TabIndex        =   34
         Top             =   1920
         Width           =   1260
      End
      Begin VB.Label Label1 
         Caption         =   "Recuperación dias nómina"
         Height          =   315
         Index           =   19
         Left            =   -74820
         TabIndex        =   33
         Top             =   2640
         Width           =   2025
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Horas jornada"
         Height          =   195
         Index           =   20
         Left            =   -69360
         TabIndex        =   32
         Top             =   2640
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo "
         Height          =   195
         Index           =   0
         Left            =   -74760
         TabIndex        =   22
         Top             =   765
         Width           =   540
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre "
         Height          =   195
         Index           =   1
         Left            =   -71880
         TabIndex        =   21
         Top             =   765
         Width           =   600
      End
      Begin VB.Label Label1 
         Caption         =   "Dirección"
         Height          =   195
         Index           =   2
         Left            =   -74760
         TabIndex        =   20
         Top             =   1245
         Width           =   675
      End
      Begin VB.Label Label1 
         Caption         =   "Población"
         Height          =   195
         Index           =   3
         Left            =   -74760
         TabIndex        =   19
         Top             =   1680
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Teléfono"
         Height          =   195
         Index           =   4
         Left            =   -66660
         TabIndex        =   18
         Top             =   1680
         Width           =   630
      End
      Begin VB.Label Label1 
         Caption         =   "Provincia "
         Height          =   195
         Index           =   5
         Left            =   -70320
         TabIndex        =   17
         Top             =   1740
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "C.P."
         Height          =   195
         Index           =   6
         Left            =   -69240
         TabIndex        =   16
         Top             =   1245
         Width           =   300
      End
      Begin VB.Label Label1 
         Caption         =   "C.I.F."
         Height          =   195
         Index           =   15
         Left            =   -66960
         TabIndex        =   15
         Top             =   765
         Width           =   630
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5760
      Top             =   -180
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
            Picture         =   "frmEmpresas.frx":09E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpresas.frx":0AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpresas.frx":0C0A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpresas.frx":0D1C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpresas.frx":0E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpresas.frx":0F40
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpresas.frx":181A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpresas.frx":20F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpresas.frx":29CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame3 
      Height          =   615
      Left            =   840
      TabIndex        =   3
      Top             =   6240
      Width           =   3615
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3195
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8220
      TabIndex        =   2
      Top             =   3960
      Width           =   1155
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9480
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3960
      Width           =   1155
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   360
      Top             =   3240
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
      Left            =   6900
      TabIndex        =   1
      Top             =   3960
      Width           =   1215
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   10830
      _ExtentX        =   19103
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Ver Todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmEmpresas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
        'y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private HaPulsadoEnter As Boolean
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private NumRegistro As Long
Private kCampo As Integer
Private TotalReg As Long
Private vIndice As Integer



Private Sub cmdAceptar_Click()
Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer

Screen.MousePointer = vbHourglass
On Error GoTo Error1
If Modo = 3 Then
    If DatosOk Then
        
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenKeyset
        RS.LockType = adLockOptimistic
        RS.Open NombreTabla, conn, , , adCmdTable
        RS.AddNew
        'Ahora insertamos  '19 Campos
        For I = 0 To RS.Fields.Count - 1
           
            If Text1(I).Text <> "" Then RS.Fields(I) = Text1(I).Text
        Next I
        'El combo
        RS!RecuperacionDias = Combo5.ListIndex
        
                
        
        
        
        
        '--------------------
        RS.Update
        RS.Close
        Data1.Refresh
        'MsgBox "                Registro insertado.             ", vbInformation
        PonerModo 0
    End If
    Else
    If Modo = 4 Then
        'Modificar
        If DatosOk Then
            ''Haremos las comprobaciones necesarias de los campos
            
            'transformamos el campo de las horas exceso
            Text1(7).Text = TransformaPuntosComas(Text1(7).Text)
            'transformamos el campo de las horas defecto
            Text1(7).Text = TransformaPuntosComas(Text1(7).Text)
            'Recordamos que el text(0) tiene el codigo y no lo puede cambiar
            For I = 1 To 30    'DEl 31 en adelante lo pongo ahi bajo lo voy poniendo ahio abajo
                If Text1(I).Tag <> "" Then
                    If Not CmpCam(Text1(I).Tag, Text1(I).Text) Then _
                        GoTo Error1
                End If
            Next I
            
            
            
            
            'Ahora modificamos
            Cad = "Select * from " & NombreTabla
            Cad = Cad & " WHERE IdEmpresa=" & Data1.Recordset.Fields(0)
            Set RS = New ADODB.Recordset
            RS.CursorType = adOpenKeyset
            RS.LockType = adLockOptimistic
            RS.Open Cad, conn, , , adCmdText
            'Almacenamos para luego buscarlo
            Cad = RS!IdEmpresa
            'modificamos
            For I = 1 To 29   ' A partir del 29 lo pondremos abajo
                If Text1(I).Text <> "" Then
                    RS.Fields(I).Value = Text1(I).Text
                    Else
                    RS.Fields(I).Value = Null
                End If
            Next I
            
            'El combo
            RS!RecuperacionDias = Combo5.ListIndex
            
            
            'Valores nuevos
            RS!EmpresaHoraExtra = Check1(3).Value
            RS!NominaAutomatica = Check1(4).Value
            
            
            'El 30. Hasta el 29 lo lee de arriba
            RS!HorarioNocturno = Val(Text1(30).Text)
            If MiEmpresa.QueEmpresa = 2 Then
                If Text1(31).Text = "" Then
                    RS!DirHuellas = Null
                Else
                    RS!DirHuellas = Text1(31).Text
                End If
            End If
            
            
            If Text1(32).Text = "" Then
                RS!pathCostes = Null
            Else
                RS!pathCostes = Text1(32).Text
            End If
        
            RS!IBAN = Text1(33).Text
            RS!sufijoN34 = Text1(34).Text
            
            
            
            If MiEmpresa.QueEmpresa = 0 Then
                'Solo para XPASSS
                RS!xpassserver = Text1(35).Text
                RS!xpassuser = Text1(36).Text
                RS!xpasspwd = Text1(37).Text
                RS!xpassultID = Val(Text1(38).Text)
            End If
            
            
            RS!SepaXML = Check1(5).Value
            
            RS.Update
            RS.Close
            'MsgBox "El registro ha sido modificado", vbInformation
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
        End If 'Datos ok modificar
    End If
End If
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
If Modo = 3 Then
    'Como estamos insertando
    LimpiarCampos
    PonerModo 0
ElseIf Modo = 4 Then
    PonerCampos
    PonerModo 2
End If
End Sub

Private Sub BotonAnyadir()
LimpiarCampos
'Añadiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Aceptar"
PonerModo 3
'Sugerimos el código
SugerirCodigoSiguiente
'Escondemos el navegador y ponemos insertando
Label2.Caption = "Insertar"
Text1(7).Text = 0
Text1(8).Text = 0
Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
'Buscar
If Modo <> 1 Then
    LimpiarCampos
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
Label2.Caption = "Modificar"
'Ponemos el foco sobre el nombre
Text1(1).SetFocus
End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim I

'Ciertas comprobaciones
If Data1.Recordset.RecordCount = 0 Then Exit Sub
'Pregunta
Cad = "Seguro que desea eliminar de la BD el registro:"
Cad = Cad & vbCrLf & "Cod: " & Data1.Recordset.Fields(0)
Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
I = MsgBox(Cad, vbQuestion + vbYesNo)
If I = vbYes Then
    'Hay que eliminar
    On Error GoTo Error2
    Screen.MousePointer = vbHourglass
    Data1.Recordset.Delete
    'Esperamos un tiempo prudencial de 1 seg
    I = Timer
    Do
    Loop Until Timer - I > 1
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
End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Activate()
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer
Dim J As String

Screen.MousePointer = vbHourglass


NombreTabla = "Empresas"
Ordenacion = " ORDER BY NomEmpresa"
HaPulsadoEnter = False
'ASignamos un SQL al DATA1
Data1.ConnectionString = conn
Data1.RecordSource = "Select * from " & NombreTabla
Data1.Refresh

'ATencion, detalle
'Como cada text1(i) le corresponde un label1(i) desde i=0 hasta count-1
' y como ademas en los tag de los text1(i) tenemos las cadens para la comprobacion
' y estas contienen el nombre del campo, que a vez es el del label(i) correspondiente
' entonces lo que hago es poner en el primer campo del tag
' una almohadilla que ahora sustuire por su label correspondiente
For I = 0 To Text1.Count - 1
    J = Mid(Text1(I).Tag, 1, 1)
    If J = "#" Then _
        Text1(I).Tag = Label1(I).Caption & Mid(Text1(I).Tag, 2)
Next I


'Path huellas solo visible belgida
Label1(22).Visible = MiEmpresa.QueEmpresa = 2
Text1(31).Visible = MiEmpresa.QueEmpresa = 2

'Picassent 2015. Ya no funciona el TCP3
FrameXpass.Visible = MiEmpresa.QueEmpresa = 0

SSTab1.Tab = 0
PonerModo 2
BotonVerTodos
End Sub



Private Sub LimpiarCampos()
Dim I
For I = 0 To Text1.Count - 1
    Text1(I).Text = ""
Next I
For I = 0 To Text2.Count - 1
    Text2(I).Text = ""
Next I
Text3(0).Text = ""
Text3(1).Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Esto es una prueba
'frmMain.SetFocus
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
Text2(vIndice).Text = vCadena
Text1(vIndice + 9).Text = vCodigo
End Sub

Private Sub Image1_Click(Index As Integer)
vIndice = Index
    Set frmB = New frmBusca
    frmB.Tabla = "Incidencias"
    frmB.CampoBusqueda = "NomInci"
    frmB.CampoCodigo = "IdInci"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "INCIDENCIAS"
    frmB.Show vbModal
    Set frmB = Nothing
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
Screen.MousePointer = vbHourglass
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
kCampo = Index
If Modo = 1 Then
    Text1(Index).BackColor = vbYellow
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Modo = 1 Then
    If KeyAscii = 13 Then
        'Ha pulsado enter, luego tenemos que hacer la busqueda
        Text1(Index).BackColor = vbWhite
        BotonBuscar
    End If
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String
If Modo = 1 Then
    Text1(Index).BackColor = vbWhite
    ElseIf Modo > 2 Then
        If Index > 8 And Index < 15 Then
            If Not IsNumeric(Text1(Index).Text) Then
                Text1(Index).Text = ""
                Text2(Index - 9).Text = ""
                Else
                    Cad = DevuelveTextoIncidencia(CInt(Text1(Index).Text))
                    If Cad = "" Then
                        Text1(Index).Text = -1
                        Text2(Index - 9).Text = "Incidencia incorrecta"
                        Else
                            Text2(Index - 9).Text = Cad
                    End If
            End If
            'ELSE de index>8 y index<15
            Else
                'Horas en formato decimal
                If Index > 6 And Index < 9 Then
                    Text1(Index).Text = TransformaPuntosComas(Text1(Index).Text)
                    If Not IsNumeric(Text1(Index).Text) Then
                        Text1(Index).Text = ""
                        Text3(Index - 7).Text = ""
                        Else
                            Cad = Index
                            Text3(Val(Cad) - 7).Text = DevuelveHora(CSng(Text1(Index)))
                    End If
            End If
        End If
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

If Text1(kCampo).Text = "" Then Exit Sub

'------------------------------------------------
'Prueba de pascual jajajaja
Dim C1 As String   'el nombre del campo
Dim Tipo As Long
Dim aux1

C1 = Data1.Recordset.Fields(kCampo).Name
C1 = " WHERE " & C1
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
    CadB = C1 & " = " & Text1(kCampo)
Case 2
    'Vemos si la cadena tiene un Falso o False
    If InStr(1, UCase(Text1(kCampo).Text), "F") Then
        aux1 = "False"
        Else
        aux1 = "True"
    End If
    CadB = C1 & " = " & aux1
Case 3
    CadB = C1 & " like '*" & Trim(Text1(kCampo)) & "*'"
Case 4

Case 5

End Select
    
CadenaConsulta = "select * from " & NombreTabla & CadB & " " & Ordenacion
PonerCadenaBusqueda
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass
On Error GoTo Error4
Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.EOF Then
    MsgBox "No hay ningún registro en la tabla" & NombreTabla, vbInformation
    Screen.MousePointer = vbDefault
    TotalReg = 0
    Exit Sub
    Label2.Caption = ""
    'PonerModo 0
    Else
        'DespalzamientoVisible True
        PonerModo 2
        Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        TotalReg = Data1.Recordset.RecordCount
        NumRegistro = 1
        PonerCampos
End If

'Data1.ConnectionString = Conn
'Data1.RecordSource = CadenaConsulta
'Data1.Refresh
'TotalReg = Data1.Recordset.RecordCount
Error4:
Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim I As Integer
'    Text1(0).Text = Data1.Recordset.Fields(0)
'    Text1(1).Text = Data1.Recordset.Fields(1)
'    Text1(2).Text = DBLet(Data1.Recordset.Fields(2))
'    Text1(3).Text = DBLet(Data1.Recordset!poblacion)
'    Text1(4).Text = DBLet(Data1.Recordset!Provincia)
'For i = 0 To Data1.Recordset.Fields.Count - 1

'Si algun campo puede ser nulo pondremos dblets
'Hasta el 29
For I = 0 To 29
   Text1(I).Text = DBLet(Data1.Recordset.Fields(I))
   If I > 8 And I < 15 Then Text2(I - 9).Text = DevuelveTextoIncidencia(Data1.Recordset.Fields(I))
Next I


I = DBLet(Data1.Recordset!RecuperacionDias, "N")
Combo5.ListIndex = I

Label2.Caption = NumRegistro & " de " & TotalReg

'Ponemos los campos de las horas en sexagesimal
If Text1(7).Text <> "" Then _
    Text3(0).Text = DevuelveHora(CSng(Text1(7).Text))
If Text1(8).Text <> "" Then _
    Text3(1).Text = DevuelveHora(CSng(Text1(8).Text))
    
    
'Campos nuevos
For I = 0 To 2
    Check1(I).Value = Text1(26 + I).Text
Next I

    'Compensables de Sema40 son extra, para CATADAU
    Check1(3).Value = Abs(Data1.Recordset!EmpresaHoraExtra)
    Check1(4).Value = Abs(Data1.Recordset!NominaAutomatica)
    Check1(5).Value = Abs(DBLet(Data1.Recordset!SepaXML, "N"))
    
Text1(30).Text = DBLet(Data1.Recordset!HorarioNocturno, "N")


Text1(31).Text = ""
If MiEmpresa.QueEmpresa = 2 Then Text1(31).Text = DBLet(Data1.Recordset!DirHuellas, "T")
Text1(32).Text = DBLet(Data1.Recordset!pathCostes, "T")
Text1(33).Text = DBLet(Data1.Recordset!IBAN, "T")
Text1(34).Text = DBLet(Data1.Recordset!sufijoN34, "T")


'Julio 2015
' Picassen XPass (Suprema)
If MiEmpresa.QueEmpresa = 0 Then
    Text1(35).Text = DBLet(Data1.Recordset!xpassserver, "T")
    Text1(36).Text = DBLet(Data1.Recordset!xpassuser, "T")
    Text1(37).Text = DBLet(Data1.Recordset!xpasspwd, "T")
    Text1(38).Text = DBLet(Data1.Recordset!xpassultID, "N")
End If
End Sub



'AGRUPAR PARA QUE NO HAGA TANTAS COMPARACIONES
Private Sub PonerModo(Kmodo As Integer)
Dim I As Integer
Dim B As Boolean

If Modo = 1 Then
    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
End If
Modo = Kmodo
'DespalzamientoVisible (Kmodo = 2)
cmdAceptar.Visible = (Kmodo >= 3)
cmdCancelar.Visible = (Kmodo >= 3)
Toolbar1.Buttons(6).Enabled = (Kmodo < 3)
Toolbar1.Buttons(7).Enabled = (Kmodo = 2)
Toolbar1.Buttons(8).Enabled = (Kmodo = 2)
Toolbar1.Buttons(1).Enabled = (Kmodo < 3)
Toolbar1.Buttons(2).Enabled = (Kmodo < 3)
Label2.Visible = (Kmodo = 2)
B = (Modo = 2) Or Modo = 0
For I = 0 To Text1.Count - 1
    Text1(I).Locked = B
Next I
Text3(0).Locked = B
Text3(1).Locked = B
Combo5.Enabled = Not B
For I = 0 To 5
    Check1(I).Enabled = Not B
Next I


B = Modo > 2
For I = 0 To Image1.Count - 1
    Image1(I).Visible = B
Next I
End Sub


Private Function DatosOk() As Boolean
Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer

DatosOk = False
'Haremos las comprobaciones necesarias de los campos
'Cad = ComprobarCampos
'If Cad <> "" Then
'    MsgBox Cad, vbExclamation
'    Exit Function
'End If

'transformamos el campo de las horas exceso
Text1(7).Text = TransformaPuntosComas(Text1(7).Text)
Do
    I = InStr(1, Text1(7).Text, ".")
    If I > 0 Then
        Text1(7).Text = Mid(Text1(7).Text, 1, I - 1) & "," & Mid(Text1(7).Text, I + 1)
    End If
    Loop Until I = 0
'transformamos el campo de las horas defecto
Do
    I = InStr(1, Text1(8).Text, ".")
    If I > 0 Then
        Text1(8).Text = Mid(Text1(8).Text, 1, I - 1) & "," & Mid(Text1(8).Text, I + 1)
    End If
    Loop Until I = 0






For I = 0 To Text1.Count - 1
    If Text1(I).Tag <> "" Then
    
        If I = 30 Then
            'horario Nocturno. Si esta vacio pongo un CERo
            If Trim(Text1(I).Text) = "" Then Text1(I).Text = "0"
        End If
        
        If Not CmpCam(Text1(I).Tag, Text1(I).Text) Then
            Exit Function
        End If
    End If
Next I



If Combo5.ListIndex < 0 Then
    MsgBox "Seleccione la opcion de recuperacion dias en nomina", vbExclamation
    Exit Function
End If

'Si la recuperacion es por Horas/Jornada(tipo 2) el campo de Horas es requerido
If Combo5.ListIndex = 2 Then
    If Text1(19).Text = "" Then
        MsgBox "Ponga las horas de Jornada para la recuperacion de dias", vbExclamation
        Exit Function
    End If
End If


'Comprobamos los campos de cuenta bacncaria
For I = 21 To 24
    If Text1(I).Text <> "" Then
        If Not IsNumeric(Text1(I).Text) Then
            MsgBox "Cuenta bancaria debe ser numérica", vbExclamation
            Exit Function
        End If
        'Comprobamos la longitud
        If Len(Text1(I).Text) <> Text1(I).MaxLength Then
            MsgBox "Longitud cuenta bancaria( Entidad/Sucursal/CC/Cuenta) incorrecta", vbExclamation
            Exit Function
        End If
    End If
Next I
Text1(25).Text = Round(Val(Text1(25).Text), 0)


'Los chekc
'Campos nuevos
For I = 0 To 2
    Text1(26 + I).Text = Check1(I).Value
Next I


'IRPF
Cad = TransformaComasPuntos(Text1(29).Text)
If Val(Cad) > 100 Then
    MsgBox "Valor maximo IRPF es 100", vbExclamation
    Exit Function
End If
Text1(29).Text = TransformaPuntosComas(Text1(29).Text)


If MiEmpresa.QueEmpresa = 2 Then
    'BELGIDA
    If Text1(31).Text <> "" Then
        If Dir(Text1(31).Text, vbDirectory) = "" Then
            MsgBox "No existe carpeta", vbExclamation
            Exit Function
        End If
    End If
End If


    If MiEmpresa.QueEmpresa = 0 Then
        'XPASS
        Cad = ""
        vIndice = 0
        For I = 35 To 38
            If Trim(Text1(I).Text) = "" Then
                Cad = Cad & "  - " & RecuperaValor(Text1(I).Tag, 1) & vbCrLf
                If vIndice = 0 Then vIndice = I
            End If
        Next
        If Cad <> "" Then
            Cad = "Campos obligatorios: " & vbCrLf & vbCrLf & Cad
            MsgBox Cad, vbExclamation
            PonerFoco Text1(vIndice)
            vIndice = 0
            Exit Function
        End If
    End If
'Al final todo esta correcto
DatosOk = True
End Function


Private Sub SugerirCodigoSiguiente()
Dim Cad
Dim RS
'Sugeriremos el codigo siguiente.
'Obviamente depende en TOTAL medida de que tabla estemos trabajando
Cad = "Select Max(IdEmpresa) from " & NombreTabla

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





Private Sub Text3_GotFocus(Index As Integer)
Text3(Index).SelStart = 0
Text3(Index).SelLength = Len(Text3(Index).Text)
End Sub

Private Sub Text3_LostFocus(Index As Integer)
If Modo > 2 Then
    Text3(Index).Text = TransformaPuntosHoras(Text3(Index))
    If IsDate(Text3(Index).Text) Then
        Text1(Index + 7).Text = DevuelveValorHora(CDate(Text3(Index)))
        Text3(Index).Text = Format(Text3(Index).Text, "h:mm:ss")
        Else
            Text1(Index + 7).Text = ""
            Text3(Index).Text = ""
    End If
End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If vUsu.Nivel > 1 Then Exit Sub
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
Case Else

End Select
End Sub

'Private Sub DespalzamientoVisible(bol As Boolean)
'Dim i
'For i = 14 To 17
'    Toolbar1.Buttons(i).Visible = bol
'Next i
'End Sub


