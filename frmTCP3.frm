VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmTCP3 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Interacci�n con el TCP-3"
   ClientHeight    =   6465
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7455
   Icon            =   "frmTCP3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6465
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDavid 
      Caption         =   "ProcesarTmpMar"
      Height          =   375
      Index           =   1
      Left            =   1200
      TabIndex        =   43
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton cmdDavid 
      Caption         =   "ProcesarTmpMar2"
      Height          =   375
      Index           =   0
      Left            =   3240
      TabIndex        =   42
      Top             =   6000
      Width           =   2295
   End
   Begin VB.CommandButton Command7 
      Caption         =   "prueba para muchos ticajes"
      Height          =   315
      Left            =   240
      TabIndex        =   25
      Top             =   6000
      Visible         =   0   'False
      Width           =   975
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   180
      TabIndex        =   1
      Top             =   180
      Width           =   7155
      _ExtentX        =   12621
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Operaciones"
      TabPicture(0)   =   "frmTCP3.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Text1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Configuracion"
      TabPicture(1)   =   "frmTCP3.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblConfig(0)"
      Tab(1).Control(1)=   "lblConfig(1)"
      Tab(1).Control(2)=   "lblConfig(2)"
      Tab(1).Control(3)=   "lblConfig(3)"
      Tab(1).Control(4)=   "txtConfig(0)"
      Tab(1).Control(5)=   "txtConfig(1)"
      Tab(1).Control(6)=   "txtConfig(2)"
      Tab(1).Control(7)=   "Command5"
      Tab(1).Control(8)=   "txtConfig(3)"
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame6 
         Height          =   5175
         Left            =   240
         TabIndex        =   36
         Top             =   360
         Width           =   6735
         Begin VB.CommandButton cmdMasAcciones 
            Caption         =   "Acciones"
            Height          =   375
            Left            =   5040
            TabIndex        =   37
            Top             =   4560
            Width           =   1455
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "ERRORES"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   975
            Left            =   120
            TabIndex        =   41
            Top             =   3360
            Width           =   6495
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "ERRORES"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   27.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   1335
            Left            =   120
            TabIndex        =   40
            Top             =   1080
            Width           =   6495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ORDENADOR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   39
            Top             =   2640
            Width           =   1830
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "RELOJ"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   120
            TabIndex        =   38
            Top             =   360
            Width           =   900
         End
         Begin VB.Line Line1 
            X1              =   120
            X2              =   6480
            Y1              =   2520
            Y2              =   2520
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Comprobar hora"
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
         Height          =   1635
         Left            =   360
         TabIndex        =   34
         Top             =   3720
         Width           =   2355
         Begin VB.CommandButton cmdCompruebaFecha 
            Caption         =   "Comprobar HORA"
            Height          =   375
            Left            =   240
            TabIndex        =   35
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.TextBox txtConfig 
         Height          =   315
         Index           =   3
         Left            =   -71700
         TabIndex        =   26
         Text            =   "Text4"
         Top             =   2940
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Poner en hora el reloj"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   1635
         Left            =   360
         TabIndex        =   19
         Top             =   1920
         Width           =   2355
         Begin VB.TextBox Text4 
            Height          =   315
            Index           =   1
            Left            =   1020
            TabIndex        =   24
            Top             =   540
            Width           =   1095
         End
         Begin VB.TextBox Text4 
            Height          =   315
            Index           =   0
            Left            =   120
            TabIndex        =   21
            Top             =   540
            Width           =   615
         End
         Begin VB.CommandButton Command6 
            Caption         =   "Programar fecha/hora"
            Height          =   375
            Left            =   240
            TabIndex        =   20
            Top             =   1140
            Width           =   1815
         End
         Begin VB.Label Label2 
            Caption         =   "Fecha"
            Height          =   195
            Index           =   1
            Left            =   1020
            TabIndex        =   23
            Top             =   300
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Hora"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   22
            Top             =   300
            Width           =   435
         End
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Guadar configuraci�n"
         Height          =   375
         Left            =   -71640
         TabIndex        =   18
         Top             =   4320
         Width           =   1935
      End
      Begin VB.TextBox txtConfig 
         Height          =   315
         Index           =   2
         Left            =   -71700
         TabIndex        =   16
         Text            =   "Text4"
         Top             =   2340
         Width           =   1935
      End
      Begin VB.TextBox txtConfig 
         Height          =   315
         Index           =   1
         Left            =   -71700
         TabIndex        =   14
         Text            =   "Text4"
         Top             =   1740
         Width           =   1935
      End
      Begin VB.TextBox txtConfig 
         Height          =   315
         Index           =   0
         Left            =   -71700
         TabIndex        =   12
         Text            =   "Text4"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         Height          =   1575
         Left            =   2880
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3840
         Width           =   4035
      End
      Begin VB.Frame Frame3 
         Caption         =   "Incidencias Reloj"
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
         Height          =   1455
         Left            =   2820
         TabIndex        =   6
         Top             =   420
         Width           =   4095
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   285
            Index           =   1
            Left            =   480
            TabIndex        =   9
            Text            =   "Elija una incidencia"
            Top             =   360
            Width           =   2415
         End
         Begin VB.CommandButton Command3 
            Caption         =   "Programar Incid."
            Height          =   375
            Left            =   1080
            TabIndex        =   8
            Top             =   840
            Width           =   1815
         End
         Begin VB.TextBox Text3 
            Height          =   315
            Left            =   3060
            TabIndex        =   7
            Top             =   360
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   180
            Picture         =   "frmTCP3.frx":0342
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            Caption         =   "Tecla"
            Height          =   195
            Left            =   3060
            TabIndex        =   10
            Top             =   180
            Width           =   555
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Empleados"
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
         Height          =   1635
         Left            =   2820
         TabIndex        =   4
         Top             =   1920
         Width           =   4095
         Begin VB.TextBox Text5 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   1080
            TabIndex        =   28
            Text            =   "Text5"
            Top             =   480
            Width           =   915
         End
         Begin VB.TextBox Text5 
            Height          =   315
            Index           =   0
            Left            =   180
            TabIndex        =   29
            Text            =   "Text5"
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox Text2 
            Enabled         =   0   'False
            Height          =   345
            Index           =   0
            Left            =   180
            TabIndex        =   5
            Text            =   "Elija un empleado"
            Top             =   1140
            Width           =   3435
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Programar tarjeta"
            Height          =   375
            Left            =   2220
            TabIndex        =   30
            Top             =   420
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   33
            Top             =   900
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "TARJETA"
            Height          =   195
            Index           =   3
            Left            =   1140
            TabIndex        =   32
            Top             =   240
            Width           =   795
         End
         Begin VB.Label Label2 
            Caption         =   "Codigo"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   31
            Top             =   240
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   840
            Picture         =   "frmTCP3.frx":0444
            Top             =   900
            Width           =   240
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Marcajes reloj"
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
         Height          =   1455
         Left            =   360
         TabIndex        =   2
         Top             =   420
         Width           =   2355
         Begin VB.CommandButton Command2 
            Caption         =   "Traer"
            Height          =   375
            Left            =   240
            TabIndex        =   3
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Label lblConfig 
         AutoSize        =   -1  'True
         Caption         =   "Tiempo espera borrado"
         Height          =   195
         Index           =   3
         Left            =   -73800
         TabIndex        =   27
         Top             =   3000
         Width           =   1635
      End
      Begin VB.Label lblConfig 
         AutoSize        =   -1  'True
         Caption         =   "Baudios"
         Height          =   195
         Index           =   2
         Left            =   -73800
         TabIndex        =   17
         Top             =   2400
         Width           =   570
      End
      Begin VB.Label lblConfig 
         AutoSize        =   -1  'True
         Caption         =   "N� TCP-3"
         Height          =   195
         Index           =   1
         Left            =   -73800
         TabIndex        =   15
         Top             =   1800
         Width           =   675
      End
      Begin VB.Label lblConfig 
         AutoSize        =   -1  'True
         Caption         =   "Puerto COMM"
         Height          =   195
         Index           =   0
         Left            =   -73800
         TabIndex        =   13
         Top             =   1140
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5940
      TabIndex        =   0
      Top             =   6000
      Width           =   1275
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2040
      Top             =   4080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      NullDiscard     =   -1  'True
      RTSEnable       =   -1  'True
      BaudRate        =   19200
   End
End
Attribute VB_Name = "frmTCP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Comprobar As Byte
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Private vIndice As Integer 'Indice para el formuario de seleccionar
Dim T1, T2
Dim Buffer$
Private NF As Integer 'Fichero
Private NombreFichero As String
Private kTCP As String  'Por si tiene + de un terminal TCP

'Datos configuracion
Private PuertoComm As Byte
Private Baudios As Long
Private NumTCP3 As Byte
Private EsperaBorrado As Integer 'En segundos




Private Sub cmdCompruebaFecha_Click()
Dim Cad As String
Dim HayErrores As Boolean

   Screen.MousePointer = vbHourglass
    'Comprobaremos la fecha y hora del reloj
    'Si esta configurado y demas
    
   
    
    Frame6.Visible = True
    Label5.Caption = "  LEYENDO "
    'Fecha PC
    Cad = UCase(Format(Now, "ddd d, mmm  hh:mm"))
    Label6.Caption = Cad

    Me.Refresh
    'Llegados aqui es donde empezamos a transmitir con el reloj
    'Abrimos el puerto
    MSComm1.PortOpen = True
    Text1.Text = ""
    HayErrores = True
    
    'Solictamos comando
    Text1.Text = Text1.Text & "Solicitando programaci�n hora/fecha reloj al TCP-3" & vbCrLf
    Text1.Refresh
    PonerTexto kTCP
    Cad = Leer(2, "Cmd:")
    LimpiaBufferRecepcion
    'Si respuesta afirmativa
    If Cad = "" Then
        GoTo Salida3
    End If
        
    'Ponemos comando 5 Leer hora/fecha en TCP-3
    Cad = "5"
    PonerTexto Cad
    Cad = Leer(2, "Cmd OK")
    LimpiaBufferRecepcion
    'No hay datos correctos
    If Cad = "" Then
        GoTo Salida3
    End If
    

    
    HayErrores = False
Salida3:
   
    
    If HayErrores Then
    
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
         Label5.Caption = "  E R R O R E S "
        End If
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    
    
    'En cad tenemos lo k ha llegador
    If Not HayErrores Then
        Label5.Caption = "Ajustando valores dev."
        Label5.Refresh
        If Not PonerHoras(Cad) Then
            Label5.Caption = "Valores devueltos erroneos"
            Label5.Refresh
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdDavid_Click(Index As Integer)

    If Index = 0 Then
        ProcesarFichero2 True, True   'solo tmpmar2
    Else
        ProcesarFichero2 False, True   'todo
    End If
        MsgBox "Proceso finalizado", vbExclamation
End Sub

Private Sub cmdMasAcciones_Click()
    Text1.Text = ""
    Frame6.Visible = False
End Sub

'PROGRAMAR TARJETA
Private Sub Command1_Click()
Dim Cad As String
Dim HayErrores As Boolean

If Text2(0).Tag = "" Then
    MsgBox "Seleccione un empleado.", vbExclamation
    Exit Sub
End If
Screen.MousePointer = vbHourglass

'Abrimos el puerto
MSComm1.PortOpen = True
Text1.Text = ""
HayErrores = True

'Solictamos comando
Text1.Text = Text1.Text & "Solicitando programacion tarjeta al TCP-3"
Text1.Refresh
PonerTexto kTCP
Cad = Leer(2, "Cmd:")
LimpiaBufferRecepcion
'Si respuesta afirmativa
If Cad = "" Then
    GoTo salida
End If
    
'Ponemos comando 18 de PRogramar TARJETA
PonerTexto "18"
Cad = Leer(2, "Zona:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo salida
End If
'Ponemos zona
PonerTexto "1"
Cad = Leer(2, "Planta:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo salida
End If
'Ponemos PLANTA
PonerTexto "18"
Cad = Leer(2, "Acceso:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo salida
End If
'
Text1.Text = Text1.Text & vbCrLf & "Trabajador: " & Text2(vIndice).Text
Text1.Refresh
'
'Ponemos ACCESO
PonerTexto "255"
Cad = Leer(2, "Fichaje:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo salida
End If
'Ponemos NUM tarjeta  ----------------
'que esta en el tag
'If Len(Text2(0).Tag) > 16 Then
'    Cad = Mid(Text2(0).Tag, 1, 16)
'    Else
'    Cad = Text2(0).Tag
'End If
Cad = Text2(0).Tag
PonerTexto Cad
Cad = Leer(2, "Identf.:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo salida
End If
'Ponemos Nombre  ----------------
'que esta en el tag
If Len(Text2(0).Text) > 16 Then
    Cad = Mid(Text2(0).Text, 1, 16)
    Else
    Cad = Text2(0).Text
End If
Text1.Text = Text1.Text & vbCrLf & "Enviando Datos tarjeta ....."
Text1.Refresh
PonerTexto Cad
Cad = Leer(15, "Cmd OK")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo salida
End If
'Todo correcto
Text1.Text = vbCrLf & vbCrLf & "Traspaso realizado con �xito."
Text1.Refresh
HayErrores = False
salida:
    If HayErrores Then _
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Screen.MousePointer = vbDefault
End Sub


'IMPORTAR MARCAJES
Private Sub Command2_Click()
Dim Cad As String
Dim Correcto As Boolean

Screen.MousePointer = vbHourglass

'Si ya existe el fichero marcaje
Cad = Dir(NombreFichero)
If Cad <> "" Then
    Cad = "Ya existe el archivo: " & NombreFichero & vbCrLf
    Cad = Cad & "Proc�selo primero"
    MsgBox Cad, vbExclamation
    Screen.MousePointer = vbDefault
    Exit Sub
End If


'Como el TCP3 no envia el a�o, se lo pondremos nosotros
'en funcion de la fecha actual.
'Primero generamos el archivo Fichajes.txt y luego procesaremos
'Sus lineas. Es ahi donde Le a�adimos el a�o.
'Problema:
'   Si hacemos el traspaso a primeros de enero y hay algun ticaje
'   de diciembre estos tendr�n el a�o como el de enero.
'   Solucion dividir el FICHAJES.txt en dos.
'   La primera mitad contendra todos los ticajes de diciembre,
'   del a�o anterior, y la segunda los de enero hasta la fecha
'   actual.
'   para procesar los de diciembre cambiaremos la fecha de sistema
'   al a�o anterior consiguiendo de ese modo cambiar el a�o.
'   Una vez procesado la primera parte volveremos a la fecha actual
'   y procearemos la segunda.
If Month(Now) = 1 And Day(Now) < 4 Then
    Cad = "Puede que hayan marcajes del a�o anterior " & Year(Now) - 1 & "." & vbCrLf
    Cad = Cad & "Si hay alguno tendra que dividir el fichero fichajes.txt" & vbCrLf
    Cad = Cad & "Del siguiente modo." & vbCrLf
    Cad = Cad & "1.- Haga una copia del FICHAJES.TXT a FICHAJES.BAK" & vbCrLf
    Cad = Cad & "2.- Abra el fichero y seleccione las entradas de enero y borrelas. " & vbCrLf
    Cad = Cad & "3.- Cambie la fecha de su ordenador y ponga 31 de diciembre de " & Year(Now) - 1 & vbCrLf
    Cad = Cad & "4.- Procese el fichero" & vbCrLf
    Cad = Cad & "5.- Copie FICHAJES.BAK como Fichajes.txt" & vbCrLf
    Cad = Cad & "6.- Vuelva a poner la fecha actual" & vbCrLf
    Cad = Cad & "7.- Procese otra vez el fichero" & vbCrLf
    MsgBox Cad, vbExclamation, "PRECAUCION"
End If
'Eliminamos los temporales
Cad = Dir(App.Path & "\tmpMar.txt")
If Cad <> "" Then
    Kill App.Path & "\tmpMar.txt"
End If
Cad = Dir(App.Path & "\tmpMar2.txt")
If Cad <> "" Then
    Kill App.Path & "\tmpMar2.txt"
End If

'Abrimos el puerto
MSComm1.PortOpen = True
Text1.Text = ""

'Probablemente haya que bloquear la unidad durante este proceso


'Solictamos comando
Text1.Text = Text1.Text & "Solicitando marcajes al TCP-3" & vbCrLf
Text1.Refresh
PonerTexto kTCP
Cad = Leer(2, "Cmd:")
'Debug.Print cad
LimpiaBufferRecepcion

'Si respuesta afirmativa
If Cad = "" Then
    Text1.Text = Text1.Text & "Error solicitando datos"
    GoTo salida
End If
    
'Ponemos comando 3 de solicitud de transferencias
PonerTexto "3"
Cad = Leer(2, "Reg. Ini:")

LimpiaBufferRecepcion
'Si respuesta afirmativa
If Cad = "" Then
    GoTo salida
End If

'Registro inicial 0
PonerTexto "0"
Cad = Leer(8, "Reg. Fin:")
LimpiaBufferRecepcion

'Si respuesta afirmativa
If Cad = "" Then
    GoTo salida
End If

'Registro Final 0
Text1.Text = Text1.Text & vbCrLf & "Devolviendo registros..."
Text1.Refresh
PonerTexto "0"

'Ahora leeremos todos los datos que nos devuelve la maquina
'sobre los tikajes
Correcto = False
If AbrirFichero = 0 Then
    Correcto = LeerDatos2(5, True) 'Leemos los datos
    Correcto = Correcto And (CerrarFichero = 0)
End If

If Not Correcto Then GoTo salida
Text1.Text = Text1.Text & vbCrLf & "Marcajes recibidos. Procesando fichero temporal."
Text1.Refresh
'Ahora procesamos el fichero
Correcto = ProcesarFichero2(False, False)
If Not Correcto Then GoTo salida
    
If Correcto Then
    Text1.Text = Text1.Text & "    TODO CORRECTO  "
    Else
        Text1.Text = Text1.Text & "    Error procesando el fichero."
        Text1.Refresh
End If

If Not Correcto Then
    'Borramos marcajes.txt
    Text1.Text = Text1.Text & vbCrLf & "Error procesando el fichero temporal"
    Text1.Refresh
    Kill NombreFichero
    GoTo salida
End If


'BORRAR MARCAJES
'Todo ha ido bien, luego borramos los marcajes de la maquina
'Solictamos comando de borrar
Text1.Text = "Solicitando borrado marcajes."
Text1.Refresh
LimpiaBufferRecepcion
PonerTexto kTCP
Cad = Leer(6, "Cmd:")
'Si respuesta afirmativa
If Cad = "" Then
    Text1.Text = Text1.Text & vbCrLf & "Error conectando con el TCP-3(Cmd: 4). Tiempo agotado"
    Text1.Refresh
    Kill NombreFichero
    GoTo salida
End If
LimpiaBufferRecepcion

'Ponemos comando 3 de solicitud de transferencias
PonerTexto "4"
Cad = Leer(2, "Confirmar:")
LimpiaBufferRecepcion
'Si respuesta afirmativa
If Cad = "" Then
    Text1.Text = Text1.Text & vbCrLf & "Error conectando con el TCP-3(Confirmar: Si). Tiempo agotado"
    Text1.Refresh
    Kill NombreFichero
    GoTo salida
End If

PonerTexto "si"
'Aqui ponemos mas tiempo por que tiene que borrar registros: EsperaBorrado
Cad = Leer(EsperaBorrado, "Cmd OK")
'Debug.Print cad
LimpiaBufferRecepcion

'Si no ha podido borrar entonces borramos el marcaje
If Cad = "" Then
    Text1.Text = Text1.Text & vbCrLf & "Ha sido imposible eliminar los marcajes en el TCP-3"
    Text1.Refresh
    Kill NombreFichero
    GoTo salida
End If
Text1.Text = Text1.Text & vbCrLf & "Marcajes eliminados en el TCP-3"
Text1.Refresh
'LLegados a este punto cambiaremos el valor de la variable
'MostrarErrores para que el formulario ppal sepa que
'tiene datos para importar
    
MostrarErrores = True
'Todo correcto
Text1.Text = vbCrLf & vbCrLf & "Traspaso realizado con �xito."
Text1.Refresh

salida:
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Screen.MousePointer = vbDefault
    If Correcto Then Unload Me
End Sub


Private Function Leer(Segundos As Integer, CadenaEsperada As String) As String
Dim i As Integer
Dim Fin As Boolean
Dim T1, T2
Dim Buffer2$
Leer = ""
Fin = False
T1 = Timer
'Bucle
While Not Fin
    
    Buffer2$ = Buffer2$ & MSComm1.Input
    'Debug.Print "R: " & Buffer2$
    If InStr(1, Buffer2$, CadenaEsperada) Then
        Fin = True
        Leer = Buffer2$
    Else
        T2 = Timer - T1
        Fin = (T2 > Segundos)
    End If
Wend

If vUsu.Codigo = 0 Then
    If Leer = "" Then
        If Buffer2$ <> "" Then MsgBox Buffer2$, vbCritical
    End If
End If

End Function



Private Sub PonerTexto(T As String)
MSComm1.Output = T & Chr(13)
End Sub


Private Sub espera(tiempo As Single)
T1 = Timer
Do
    T2 = Timer - T1
Loop Until T2 > tiempo
End Sub



Private Sub EscribeTextoFichero(Texto As String)
'ProcesaTexto texto
Print #NF, Texto
End Sub

Private Function AbrirFichero() As Byte
On Error GoTo ErrorAbrir
'Abrimos el fichero para escritura
AbrirFichero = 1
NF = FreeFile
        Open App.Path & "\tmpMar.txt" For Output As #NF
AbrirFichero = 0
Exit Function
ErrorAbrir:
    MsgBox "Error abriendo fichero: " & vbCrLf & "N�mero: " & Err.Number & _
        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
End Function

Private Function CerrarFichero() As Byte
On Error GoTo ErrorCerrar
'Abrimos el fichero para escritura
CerrarFichero = 1
Close NF
CerrarFichero = 0
Exit Function
ErrorCerrar:
    MsgBox "Error cerrando fichero: " & vbCrLf & "N�mero: " & Err.Number & _
        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
End Function

Private Sub LimpiaBufferRecepcion()
Dim Cade
Cade = MSComm1.Input
espera 0.25
End Sub


'Private Function ProcesarFichero() As Boolean
'Dim Linea As String
'Dim i As Long
'Dim TotalReg As Long
'Dim Fich As Integer
'
'
'Text1.Text = ""
'On Error GoTo ErrProcFich
'Fich = FreeFile
'NF = FreeFile + 1
'Open App.Path & "\tmpMar.txt" For Input As #NF
'Open NombreFichero For Output As #Fich
'i = 0 'Tendremos el contador
'While Not EOF(NF)
'    Line Input #NF, Linea
'    'Ahora procesamos la linea
'    Linea = Trim(Linea)
'    If Len(Linea) = 1 Then Linea = ""
'    'Segun este en blanco ponga unas cosas etcc..
'    If Linea <> "" Then
'        If Mid(Linea, 1, 10) = "Total Reg." Then
'            TotalReg = CLng(Val(Mid(Linea, 11)))
'            Else
'                If Mid(Linea, 1, 4) <> "Reg." Then
'                    If Mid(Linea, 1, 3) <> "Cmd" Then
'                        Print #Fich, Linea
'                        Text1.Text = Text1.Text & Linea
'                        i = i + 1
'                    End If
'                End If
'        End If
'    End If
'Wend
'Close #NF
'Close #Fich
''Vemos si esta correcto o no
'ProcesarFichero = (i = TotalReg)
'Exit Function
'ErrProcFich:
'    ProcesarFichero = False
'    MsgBox "Error procesando fichero: " & vbCrLf & "N�mero: " & Err.Number & _
'        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
'End Function



'INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS
'INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS INCIDENCIAS
Private Sub Command3_Click()
Dim tecla As Integer
Dim HayErrores As Boolean
Dim Cad As String


If Text2(1).Tag = "" And Text2(0).Text = "Elija una incidencia" Then
    MsgBox "Seleccione una incidencia.", vbExclamation
    Exit Sub
End If

If Text3.Text = "" Then
    MsgBox "Seleccione una incidencia y una tecla para asociarla.", vbExclamation
    Exit Sub
End If

If Not IsNumeric(Text3.Text) Then
    MsgBox "El codigo de tecla de incidencia tiene que ser numerico.", vbExclamation
    Exit Sub
End If

tecla = CInt(Text3.Text)
If tecla < 1 Then
    MsgBox "El c�digo de tecla tiene que ser mayor que cero.", vbExclamation
    Exit Sub
End If
If tecla > 9 Then
    MsgBox "El c�digo de tecla tiene que ser menor que 10 ( Del 0 al 9).", vbExclamation
    Exit Sub
End If

Screen.MousePointer = vbHourglass
'Llegados aqui es donde empezamos a transmitir con el reloj
'Abrimos el puerto
MSComm1.PortOpen = True
Text1.Text = ""
HayErrores = True
'Solictamos comando
Text1.Text = Text1.Text & "Solicitando grabaci�n de mensajes al TCP-3" & vbCrLf
Text1.Refresh
PonerTexto kTCP
Cad = Leer(2, "Cmd:")
LimpiaBufferRecepcion
'Si respuesta afirmativa
If Cad = "" Then
    GoTo Salida2
End If
    
'Ponemos comando 13 grabar mensajes usuario
PonerTexto "13"
Cad = Leer(2, "Usuario")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida2
End If
'Ponemos la tecla asignada
PonerTexto CStr(tecla)
Cad = Leer(2, "Ind:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida2
End If
'Ponemos PLANTA
If Len(Text2(1).Text) < 16 Then
    Cad = Text2(1).Text & "                           "
    Else
        Cad = Mid(Text2(1).Text, 1, 16)
End If
Cad = Mid(Cad, 1, 16)
Text1.Text = Text1.Text & " Incidencia: " & tecla & vbCrLf
Text1.Text = Text1.Text & " Mensaje: " & Cad & vbCrLf
Text1.Refresh
PonerTexto Cad

Cad = Leer(6, "Cmd OK")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida2
End If
Text1.Text = Text1.Text & vbCrLf & "Grabaci�n correcta" & vbCrLf
Text1.Refresh

HayErrores = False
Salida2:
    If HayErrores Then _
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Screen.MousePointer = vbDefault



End Sub

Private Sub Command4_Click()
If MSComm1.PortOpen Then MSComm1.PortOpen = False
Unload Me
End Sub


'-------------------------------------------------
'-------------------------------------------------
'  Configuracion terminal
'  ----------------------
Private Sub Command5_Click()
Dim i As Integer

On Error GoTo ErrorGuardar
For i = 0 To 3
    If Not IsNumeric(txtConfig(i).Text) Then
        MsgBox "Los campos de configuracion tienen que ser num�ricos.", vbExclamation
        Exit Sub
    End If
Next i
    i = FreeFile
    Open App.Path & "\TCPConf.cfg" For Output As #i
    Print #i, Trim(txtConfig(0)) & "|" & Trim(txtConfig(1)) & "|" & Trim(txtConfig(2)) & "|" & Trim(txtConfig(3)) & "|"
    Close #i
ErrorGuardar:
    If Err.Number <> 0 Then
        MsgBox "Se ha producido un error." & vbCrLf & _
            " N�mero: " & Err.Number & vbCrLf & _
            " Descripci�n: " & Err.Description, vbExclamation
            
        Else
            'Ponemos los valores al COMM
            MSComm1.CommPort = txtConfig(0).Text
            MSComm1.Settings = Trim(txtConfig(2).Text) & ",N,8,1"
            kTCP = "tcp" & Trim(txtConfig(1).Text)
            EsperaBorrado = CInt(txtConfig(3))
    End If
End Sub

'-------------------------------------------------
'-------------------------------------------------
'  Poner en hora el reloj
'  ----------------------
' �  este es el caracter que recibe por la �
' puede, no esta comprobado, que de errores en un futuro
Private Sub Command6_Click()
Dim HayErrores As Boolean
Dim Cad As String


If Text4(0).Text = "" Then
    MsgBox "Ponga una hora para transmitir al TCP3", vbExclamation
    Exit Sub
End If

If Text4(1).Text = "" Then
    MsgBox "Ponga una fecha para transmitir al TCP3", vbExclamation
    Exit Sub
End If

If Not IsDate(Text4(0).Text) Then
    MsgBox "Ponga una hora correcta para transmitir al TCP3", vbExclamation
    Exit Sub
End If

If Not IsDate(Text4(1).Text) Then
    MsgBox "Ponga una fecha correcta para transmitir al TCP3", vbExclamation
    Exit Sub
End If


Screen.MousePointer = vbHourglass
'Llegados aqui es donde empezamos a transmitir con el reloj
'Abrimos el puerto
MSComm1.PortOpen = True
Text1.Text = ""
HayErrores = True

'Solictamos comando
Text1.Text = Text1.Text & "Solicitando programaci�n hora/fecha reloj al TCP-3" & vbCrLf
Text1.Refresh
PonerTexto kTCP
Cad = Leer(2, "Cmd:")
LimpiaBufferRecepcion
'Si respuesta afirmativa
If Cad = "" Then
    GoTo Salida3
End If
    
'Ponemos comando 2 programar hora/fecha en TCP-3
'El 1 tb programa pero el 2 en nuestro terminal es mas completo
PonerTexto "2"
Cad = Leer(2, "a�o:")  'realmente � es la �
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida3
End If

'Ponemos el a�o
Cad = Year(CDate(Text4(1).Text))
Text1.Text = Text1.Text & "A�o: " & Cad & vbCrLf
Text1.Refresh
PonerTexto Cad
Cad = Leer(2, "mes:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida3
End If

'Ponemos el mes
Cad = Month(CDate(Text4(1).Text))
Text1.Text = Text1.Text & "Mes: " & Cad & vbCrLf
Text1.Refresh
PonerTexto Cad
Cad = Leer(2, "dia:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida3
End If
'Ponemos el dia
Cad = Day(CDate(Text4(1).Text))
Text1.Text = Text1.Text & "Dia: " & Cad & vbCrLf
Text1.Refresh
PonerTexto Cad
Cad = Leer(2, "ds:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida3
End If
'Ponemos el dia semana
Cad = Weekday(CDate(Text4(1).Text), vbMonday)
Text1.Text = Text1.Text & "Dia seman: " & Cad & "  (1 - Lunes ...)" & vbCrLf
Text1.Refresh
PonerTexto Format(Cad, "00")
Cad = Leer(2, "hora:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida3
End If
'Ponemos la hora
Cad = Hour(CDate(Text4(0).Text))
Text1.Text = Text1.Text & "Hora: " & Cad & vbCrLf
Text1.Refresh
PonerTexto Cad
Cad = Leer(2, "min:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida3
End If
'Los minutos
Cad = Minute(CDate(Text4(0).Text))
Text1.Text = Text1.Text & "Hora: " & Cad & vbCrLf
Text1.Refresh
PonerTexto Cad
Cad = Leer(2, "seg:")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida3
End If
'Los segundos
Cad = 30
Text1.Text = Text1.Text & "Segundos: " & Cad & vbCrLf
Text1.Refresh
PonerTexto Cad
Cad = Leer(2, "Cmd OK")
LimpiaBufferRecepcion
'No hay datos correctos
If Cad = "" Then
    GoTo Salida3
End If

'llegados  a este punto todo ha ido bien
Text1.Text = Text1.Text & vbCrLf & vbCrLf
Text1.Text = Text1.Text & vbCrLf & "Comando completado con �xito."
Text1.Refresh

'
HayErrores = False
Salida3:
    If HayErrores Then _
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Screen.MousePointer = vbDefault







End Sub

Private Sub Command7_Click()
ProcesarFichero2 False, False
End Sub


Private Sub Form_Activate()
    If Me.Tag <> "" Then
        'La primera vez
        Me.Tag = ""
        If Comprobar Then cmdCompruebaFecha_Click
    End If
End Sub

Private Sub Form_Load()
    Me.Tag = "OK"
    'Si solo es comprobar la fecha solo le dejo salir
    Me.cmdMasAcciones.Visible = Not Comprobar
    Text2(0).Tag = ""
    Text2(1).Tag = ""
    Text5(0).Text = ""
    Text5(1).Text = ""
    Frame6.Visible = False
    NombreFichero = mConfig.DirMarcajes & "\" & mConfig.NomFich
    ObtenerConfiguracion
    MSComm1.CommPort = txtConfig(0).Text
    MSComm1.Settings = Trim(txtConfig(2).Text) & ",N,8,1"
    kTCP = "tcp" & Trim(txtConfig(1).Text)
    SSTab1.Tab = 0
    Text4(0).Text = Format(Now, "hh:mm")
    Text4(1).Text = Format(Now, "dd/mm/yyyy")
    
    cmdDavid(0).Visible = vUsu.Codigo = 0 'solo root
    cmdDavid(1).Visible = vUsu.Codigo = 0 'solo root
    Screen.MousePointer = vbDefault
End Sub




Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    Dim Cad As String
    If vIndice = 0 Then
        Cad = devuelveNombreNTarjeta(CInt(vCodigo))
        If Cad <> "" Then
            'Necesito saber tambien el n�mero de la tarjeta
            Text2(0).Tag = Cad
            Text2(0).Text = vCadena
            Text5(0).Text = vCodigo
            Text5(1).Text = Cad
        End If
        Else
            Text2(1).Tag = vCodigo
            Text3.Text = vCodigo
            Text2(1).Text = vCadena
    End If
End Sub

Private Sub Image1_Click(Index As Integer)
    Set frmB = New frmBusca
    'Ponemos los valores para abrir
    If Index = 0 Then
        vIndice = 0
        frmB.Tabla = "Trabajadores"
        frmB.CampoBusqueda = "NomTrabajador"
        frmB.CampoCodigo = "IdTrabajador"
        frmB.TipoDatos = 3
        frmB.Titulo = "EMPLEADOS"
        Else
            vIndice = 1
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


Private Sub ObtenerConfiguracion()
Dim Cad As String
Dim NF As Integer
Dim i As Integer
Dim ini As Integer
Dim V(3) As String
Dim C As String
Dim L As Long

On Error GoTo ErrorObtener
V(0) = "1"
V(1) = "1"
V(2) = "19200"
V(3) = "40"
Cad = Dir(App.Path & "\TCPConf.cfg")
If Cad <> "" Then
    NF = FreeFile
    Open App.Path & "\TCPConf.cfg" For Input As #NF
    Input #NF, Cad
    Close #NF
    'Leemos los valores
    ini = 1
    For NF = 0 To 3
        i = InStr(ini, Cad, "|")
        If i > 0 Then
            C = Mid(Cad, ini, i - ini)
            If IsNumeric(C) Then V(NF) = C
            ini = i + 1
        End If
    Next NF
End If
'Valores por defecto
'NO existe el fichero. Ponemos los valores por defecto
ErrorObtener:
    L = CLng(V(0))
    If L > 255 Then L = 1
    PuertoComm = CByte(L)
    '-------------------
    L = CLng(V(1))
    If L > 255 Then L = 1
    NumTCP3 = CByte(L)
    '---------------------
    Baudios = CLng(V(2))
    '---------------------
    EsperaBorrado = CInt(V(3))
    For i = 0 To 3
        txtConfig(i) = V(i)
    Next i
End Sub




'Segunda prueba para leer mucho datos desde el TCP 3
Private Function LeerDatos2(Segundos As Integer, Leyendo As Boolean) As Boolean
Dim Seg As Integer
Dim i As Integer
Dim Fin As Boolean
Dim tmp As String
Dim Esta As Boolean
Dim J As Integer
Dim T3, T4
    'De esta forma haremos:
    'iremos leyendo del buffer
    ' Si hay texto cad segundos/2 (aproximadamente)
    'lo guardamos y
    'restauramos tiempo
    'Si no hay tiempo salimos sin hacer nada
    Buffer$ = ""
    Fin = False
    T1 = Timer
    J = Segundos + 1
    'Bucle
    Text1.Text = Text1.Text & vbCrLf
    If Leyendo Then
        Text1.Text = Text1.Text & "Leyendo registros de marcajes. " & vbCrLf
        Else
        Text1.Text = Text1.Text & "Eliminando registros de marcajes. " & vbCrLf
    End If
    Text1.Refresh
    tmp = Text1.Text
    i = 0
    T3 = Timer
    While Not Fin
        
        
        Do
            Buffer$ = Buffer$ & MSComm1.Input
            T2 = Timer - T1
            If Buffer$ <> "" Then J = 2
        Loop Until T2 > Segundos Or T2 > J
        
        
        T4 = Timer - T3
        If Buffer$ <> "" Then
            Text1.Text = tmp & vbCrLf & "Bloque: " & Format(i, "0000") & "     " & "  Seg: " & Format(T4, "0.00")
            Text1.Refresh
            i = i + 1
            
            EscribeTextoFichero Buffer$
            If InStr(1, Buffer$, "Cmd OK") Then
                Fin = True
                Esta = True
                Else
                    Buffer$ = ""
                    J = Segundos + 1
                    T1 = Timer
            End If
        Else
            'i=0
            'fin tiempo
            Esta = False
            Fin = True
        End If
    Wend
    LeerDatos2 = Esta
End Function






Private Function ProcesarFichero2(ProcesarTmp2 As Boolean, DesdeBotonRoot As Boolean) As Boolean
Dim Linea As String
Dim Aux As String
Dim i As Long
Dim TotalReg As Long
Dim Fich As Integer
Dim Fin As Boolean
Dim J As Integer
Dim Cad As String
Dim Aux3 As String
Dim K3 As Integer


'Del fichero leido lo ponemos en lineas que contienen unicamente numero
' lo leido por el fichero
' Lineas del tipo linea i  : Reg. i
'                 linea i+1: 09001,10,12,10,53,0000,0000,02302
'
On Error GoTo ErrProcFich
Fich = FreeFile
NF = FreeFile + 1


If Not ProcesarTmp2 Then

        'GoTo Aqui
        Open App.Path & "\tmpMar.txt" For Input As #NF
        Open App.Path & "\tmpMar2.txt" For Output As #Fich
        
        If EOF(NF) Then
            MsgBox "El archivo esta vac�o.", vbExclamation
            Close #NF
            Close #Fich
            ProcesarFichero2 = False
            Exit Function
        End If
        
        Fin = False
        
            
        'Leemos la primera linea k despreciamos
        Line Input #NF, Linea
        'Segunda linea
        Line Input #NF, Linea
        'Por si acaso NO lleva los vblf
        If DesdeBotonRoot Then
            If Mid(Linea, Len(Linea), 1) <> vbLf Then Linea = Linea & vbLf
        End If
        
        
        'El resto del fichero
        While Not EOF(NF)
        
            Line Input #NF, Cad
            J = InStr(1, Cad, "Total Reg")
            If J > 0 Then
                'Mandamos a imprimir lo que esta antes de totalreg
                Aux = Linea & Mid(Cad, 1, J - 1)
                'Escribimos en el fichero Aux
                Print #Fich, Aux
                'Recalculamos linea
                Linea = Mid(Cad, J)
            Else
                    i = InStr(1, Cad, "Reg.")
                    If i > 0 Then
                    
                        If i = 1 Then
                            If Linea <> "" Then
                                Print #Fich, Linea
                                Linea = ""
                            End If
                        End If
                        'Para que acabe en vblf
                        If DesdeBotonRoot Then
                                If Mid(Cad, Len(Cad), 1) <> vbLf Then Cad = Cad & vbLf
                        End If
                        
                        'If Mid(Cad, 1, 8) = "Reg. 522" Then Stop

                        Aux = Linea & Mid(Cad, 1, i - 1)
                        'Escribimos en el fichero Aux
                        If Aux <> "" Then Print #Fich, Aux
                        'Recalculamos linea
                        Linea = Mid(Cad, i)
                        'Sino
                    Else
                        If DesdeBotonRoot Then
                            If Mid(Cad, Len(Cad), 1) <> vbLf Then Cad = Cad & vbLf
                        End If
                        Linea = Linea & Cad
                    End If
            End If
        Wend
        
        
        'Lo ultimo del texto lo imprimimos
        If Linea <> "" Then Print #Fich, Linea
        
        'Cerramos los ficheros
        Close #NF
        Close #Fich

End If



'Aqui:
If ProcesarTmp2 Then NombreFichero = App.Path & "\Fichajes2.txt"

'Ya hemos tenemos en tmpMarcajes el fichero PRE-procesado
'ahora lo terminamos de procesar
Fich = FreeFile
NF = FreeFile + 1
Open App.Path & "\tmpMar2.txt" For Input As #NF
Open NombreFichero For Output As #Fich

i = 0 'Tendremos el contador.
Fin = False
Linea = ""
While Not Fin
    Line Input #NF, Cad
    J = InStr(1, Cad, "Total Reg")
        If J > 0 Then
            'Es la ultima linea que nos intersa
            TotalReg = CLng(Val(Mid(Cad, 11)))
            Fin = True
            Linea = ""
            Else
                'If Mid(Linea, 1, 7) = "Reg. 556" Then Stop
                
                'Por si acaso van dos lineas en una
                K3 = InStr(9, Cad, "Reg.")
                If K3 > 0 Then
                    Aux3 = Mid(Cad, K3)
                    Cad = Mid(Cad, 1, K3 - 1)
                    
                    'Ahora hay que procear cada linea
                    Linea = ProcesarLineaMarcaje(Aux3)
                    Print #Fich, Linea
                    i = i + 1
                    
                    
                End If
                'Ahora hay que procear cada linea
                Linea = ProcesarLineaMarcaje(Cad)
                Print #Fich, Linea
                i = i + 1
        End If 'De j>0
    Fin = Fin Or EOF(NF)
Wend

'Cerramos los ficheros
Close #NF
Close #Fich

'Vemos si esta correcto o no
ProcesarFichero2 = (i = TotalReg)

If ProcesarTmp2 Then MsgBox "Fin proceso: " & NombreFichero

Exit Function
ErrProcFich:
    ProcesarFichero2 = False
    MsgBox "Error procesando fichero: " & vbCrLf & "N�mero: " & Err.Number & _
        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
End Function


Private Function ProcesarLineaMarcaje(Linea As String) As String
Dim Aux As String
Dim i As Integer

On Error GoTo errorProcesarLineaMarcaje

ProcesarLineaMarcaje = ""
'Separamos las dos partes de la linea
i = InStr(2, Linea, vbLf)


Aux = Mid(Linea, i + 1)
Aux = Mid(Aux, 1, Len(Aux) - 1)
ProcesarLineaMarcaje = Aux
Exit Function
errorProcesarLineaMarcaje:
    Aux = "Una linea del fichero ha llegado con error." & vbCrLf
    Aux = Aux & "--> " & Linea
    Aux = Aux & vbCrLf & vbCrLf & "Revise luego a mano el fichero FICHAJES.TXT"
    MsgBox Aux, vbCritical
    ProcesarLineaMarcaje = Linea
End Function


Private Sub Label1_DblClick()
    Text2(1).Enabled = Not Text2(1).Enabled
End Sub

Private Sub PonerEmpleadoVacio()
            Text5(0).Text = ""
            Text5(1).Text = ""
            Text2(0).Text = ""
            Text2(0).Tag = ""
End Sub
Private Sub PonerEmpleado(Cod As String, Campo As String)
Dim RT As ADODB.Recordset
Dim SQL As String
    
    SQL = "Select * from Trabajadores where "
    SQL = SQL & Campo & " = " & Cod
    Set RT = New ADODB.Recordset
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RT.EOF Then
        'ponerempleadovacio
        PonerEmpleadoVacio
    Else
        'Ponemos los datos del empleado
        If IsNull(RT!NumTarjeta) Then
            MsgBox "No tiene codigo tarjeta asociado", vbExclamation
            PonerEmpleadoVacio
        Else
            Text5(0).Text = RT!idTrabajador
            Text5(1).Text = RT!NumTarjeta
            Text2(0).Text = RT!nomtrabajador
            Text2(0).Tag = RT!NumTarjeta
        End If
    End If
    RT.Close
    Set RT = Nothing
End Sub



Private Sub Text5_GotFocus(Index As Integer)
    With Text5(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text5_LostFocus(Index As Integer)
    Text5(Index).Text = Trim(Text5(Index).Text)
    If Text5(Index).Text <> "" Then
        If Not IsNumeric(Text5(Index).Text) Then
            MsgBox "Codigo incorrecto: " & Text5(Index).Text, vbExclamation
            Text5(Index).Text = ""
        End If
    End If
    If Text5(Index).Text = "" Then
        PonerEmpleadoVacio
    Else
        If Index = 0 Then
            PonerEmpleado Text5(Index).Text, "idTrabajador"
        Else
            PonerEmpleado "'" & Text5(Index).Text & "'", "NumTarjeta"
        End If
    End If
End Sub


'Pone la cadena devuelta por el reloj  y la fecha /hora del PC
Private Function PonerHoras(CADENA As String) As Boolean
Dim i As Integer
Dim Fecha As Date

    PonerHoras = False
    Buffer$ = ""
    '5 Reg. Fin:
    'mes,dia,ds,hora,min,seg:08,31,5,21,38,21
    ' a partir de los dos puntos
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        
        'BIEN
        CADENA = Mid(CADENA, i + 1)
        'Le aa�dimos una coma al final para facilitar
        
        
    
            NF = InStr(1, CADENA, Chr(13))
            If NF > 0 Then CADENA = Mid(CADENA, 1, NF - 2)
            CADENA = CADENA & ","
        
        
        NF = 0
        Do
            i = InStr(1, CADENA, ",")
            If i > 0 Then
                NF = NF + 1
                Buffer$ = Mid(CADENA, i + 1)
                CADENA = Mid(CADENA, 1, i - 1) & "|" & Buffer$
            End If
        Loop Until i = 0
        
        Buffer$ = ""
        If NF = 6 Then

            'FECHA OK
            '----
            'Dia semana lo calcularemos de la primerasemana del
            'mes de noviembre de 2004 que empieza ekl 1 Lunes
            i = Val(RecuperaValor(CADENA, 3))
            If i > 7 Then Exit Function
            'I = I + 1
            Buffer$ = Buffer$ & Format(i & "/11/2004", "ddd")
            
            
            i = Val(RecuperaValor(CADENA, 2))
            If i = 0 Then Exit Function
            Buffer$ = Buffer$ & " " & i & ","
            
            i = Val(RecuperaValor(CADENA, 1))
            If i = 0 Then Exit Function
            Buffer$ = Buffer$ & "  " & Format("01/" & i & "/2004", "mmm")
            
            'Hora
            i = Val(RecuperaValor(CADENA, 4))
            Buffer$ = Buffer$ & "  " & i & ":"
            
            'Minutos
            i = Val(RecuperaValor(CADENA, 5))
            Buffer$ = Buffer$ & Format(i, "00")
            
            
            Label5.Caption = UCase(Buffer$)
            PonerHoras = True
        End If
        
    End If
    
    Buffer$ = ""
End Function
