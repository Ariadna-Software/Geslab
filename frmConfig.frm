VERSION 5.00
Begin VB.Form frmConfig 
   Caption         =   "Configuración del Control de presencia"
   ClientHeight    =   5760
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7050
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check6 
      Caption         =   "Comprobar Hora al iniciar Programa"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   5280
      Width           =   3435
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Gestión laboral"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   4080
      Width           =   1515
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   2100
      MaxLength       =   1
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3600
      Width           =   375
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Utiliza relojes Kimaldi"
      Height          =   195
      Left            =   3660
      TabIndex        =   14
      Top             =   4860
      Width           =   3435
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Horas especiales en Sábado"
      Height          =   195
      Left            =   3660
      TabIndex        =   13
      Top             =   4440
      Width           =   3435
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Exportacion Ariadna Sofware"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   4440
      Width           =   3435
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Reloj de control TCP-3 tipo Lipsoft  Elec."
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   4860
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   2280
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3060
      Width           =   6855
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1500
      Width           =   6855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5400
      TabIndex        =   3
      Top             =   5340
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   5340
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   735
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmConfig.frx":030A
      Top             =   420
      Width           =   6855
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Esta en parametros de la empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   1680
      TabIndex        =   18
      Top             =   4080
      Width           =   2490
   End
   Begin VB.Label Label1 
      Caption         =   "Digito tarjetas trabajadores"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   3660
      Width           =   2010
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del fichero"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   10
      Top             =   2040
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   2400
      Picture         =   "frmConfig.frx":0310
      Top             =   2820
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   1980
      Picture         =   "frmConfig.frx":0412
      Top             =   1260
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta de ficheros procesados"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   2820
      Width           =   2250
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta ficheros marcajes"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1260
      Width           =   1830
   End
   Begin VB.Label Label1 
      Caption         =   "DSN"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1470
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public vCon As CFGControl

Dim EsNuevo As Boolean

Private Sub Check1_Click()
If Check1.Value = 1 Then Check4.Value = 0
End Sub

Private Sub Check4_Click()
    If Check4.Value = 1 Then Check1.Value = 0
End Sub

Private Sub Command1_Click()

If DatosOk Then
    'Guardamos los datos
    vCon.Guardar
    'Vemos si hay que reiniciar la aplicacion
    If EsNuevo Then
        'Copiamos el archivo cambiando los valores
        MsgBox "Debe de reiniciar la aplicación."
        End
        Else
            'Si ha cambiado la cadena de conexion
            'hay que reiniciar
            If mConfig.BaseDatos <> vCon.BaseDatos Then
                'Guardamos los datos
                MsgBox "Debe de reiniciar la aplicación."
                End
            End If
            MsgBox "Seria conveniente reiniciar la aplicacion.", vbCritical
    End If
    Set mConfig = vCon
    'Asignamos el objeto
    Unload Me
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Dim I As Integer
If vCon Is Nothing Then
    'Es nuevo
    For I = 0 To Text1.Count - 1
        Text1(I).Text = ""
    Next I
    Set vCon = New CFGControl
    EsNuevo = True
    Else
        EsNuevo = False
        'Modificar
        Text1(0).Text = vCon.BaseDatos
        Text1(1).Text = vCon.DirMarcajes
        Text1(2).Text = vCon.DirProcesados
        Text1(3).Text = vCon.NomFich
        Text1(4).Text = vCon.DigitoTrabajadores
        Check1.Value = Abs(vCon.TCP3_)
        Check2.Value = Abs(vCon.Ariadna)
        Check3.Value = Abs(vCon.SabadosHorasFestivas)
        Check4.Value = Abs(vCon.Kimaldi)
        Check5.Value = Abs(vCon.Laboral)
        Check6.Value = Abs(vCon.ComprobarHoraReloj)
        
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set vCon = Nothing
End Sub

Private Sub Image1_Click(Index As Integer)
Dim Cad As String
Cad = GetFolder(Label1(Index + 1).Caption)
If Cad <> "" Then Text1(1 + Index).Text = Cad
End Sub


Private Function DatosOk() As Boolean
Dim I As Integer
Dim Cad As String


    DatosOk = False
    On Error GoTo SalirDatosOk
    
    For I = 0 To Text1.Count - 1
        Text1(I).Text = Trim(Text1(I).Text)
    Next I
'Comprobamos ciertos datos

For I = 0 To Text1.Count - 1
    If Text1(I).Text = "" Then
        MsgBox Label1(I).Caption & " NO puede estar vacia.", vbExclamation
        Exit Function
    End If
Next I

    'Este trozo comprobara si las carpetas existen o no
'    Cad = Dir(Text1(1).Text, vbDirectory)
'    If Cad = "" Then
'        MsgBox "La carpeta donde se colocarán los ficheros de marcajes NO esiste.", vbExclamation
'        Exit Function
'    End If
'
'    Cad = Dir(Text1(2).Text, vbDirectory)
'    If Cad = "" Then
'        MsgBox "La carpeta donde se colocarán los ficheros procesados NO esiste.", vbExclamation
'        Exit Function
'    End If
'vCon = Text1(0).Text
vCon.BaseDatos = Text1(0).Text
vCon.DirMarcajes = Text1(1).Text
vCon.DirProcesados = Text1(2).Text
vCon.NomFich = Text1(3).Text
vCon.TCP3_ = (Check1.Value = 1)
vCon.Ariadna = (Check2.Value = 1)
vCon.SabadosHorasFestivas = (Check3.Value = 1)
vCon.Kimaldi = (Check4.Value = 1)
vCon.DigitoTrabajadores = Text1(4).Text
vCon.Laboral = (Check5.Value = 1)

'Si NO esta TCP3 no comprobamos
If Not vCon.TCP3_ Then
    vCon.ComprobarHoraReloj = False
Else
    vCon.ComprobarHoraReloj = (Check6.Value = 1)
End If

DatosOk = True
SalirDatosOk:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
End Function



'Private Function RealizarCopia() As Byte
'Dim NF As Integer
'
'
'
'NF = FreeFile
'If Not EsNuevo Then
'    FileCopy App.Path & "\control.cfg", App.Path & "\control.tmp"
'End If
'
'Open App.Path & "\control.cfg" For Output As NF
''abriremos el config e iremos copiando ls lineas hasta que tengamos lo que queremos
'Print #NF, "# Parametrizaci¢n para la aplicacion"
'Print #NF,
'Print #NF, "BaseDatos = " & Text1(0).Text
'Print #NF, "DirMarcajes = " & Text1(1).Text
'Print #NF, "DirProcesados = " & Text1(2).Text
'Print #NF,
'Print #NF, "# Fin parametros de configuracion"
'Close #NF
'End Function
