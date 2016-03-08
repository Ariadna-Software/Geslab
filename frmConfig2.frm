VERSION 5.00
Begin VB.Form frmConfig2 
   Caption         =   "Configuración del Control de presencia"
   ClientHeight    =   2445
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7170
   Icon            =   "frmConfig2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   120
      TabIndex        =   12
      Top             =   1740
      Width           =   3855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Exportacion Ariadna Sofware"
      Height          =   195
      Left            =   180
      TabIndex        =   11
      Top             =   4560
      Width           =   3435
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Reloj de control TCP-3 tipo Lipsoft  Elec."
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   4140
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   3
      Left            =   120
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   2880
      Width           =   3615
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   120
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3660
      Width           =   3435
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   2100
      Width           =   3615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   5580
      TabIndex        =   3
      Top             =   1860
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1860
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   1215
      Index           =   0
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmConfig2.frx":030A
      Top             =   420
      Width           =   6855
   End
   Begin VB.Label Label1 
      Caption         =   "Nombre del fichero"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   2640
      Width           =   1410
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   2400
      Picture         =   "frmConfig2.frx":0310
      Top             =   3420
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   1980
      Picture         =   "frmConfig2.frx":0412
      Top             =   1860
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta de ficheros procesados"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   3420
      Width           =   2250
   End
   Begin VB.Label Label1 
      Caption         =   "Carpeta ficheros marcajes"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   1860
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
Attribute VB_Name = "frmConfig2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public vCon As CFGControl
Public mConfig As CFGControl


Dim EsNuevo As Boolean

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
Dim i As Integer
If vCon Is Nothing Then
    'Es nuevo
    For i = 0 To Text1.Count - 1
        Text1(i).Text = ""
    Next i
    Set vCon = New CFGControl
    EsNuevo = True
    Else
        EsNuevo = False
        'Modificar
        Text1(0).Text = vCon.BaseDatos
        Text1(1).Text = vCon.DirMarcajes
        Text1(2).Text = vCon.DirProcesados
        Text1(3).Text = vCon.NomFich
        Check1.Value = Abs(vCon.TCP3)
        Check2.Value = Abs(vCon.Ariadna)
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


Public Function GetFolder(Titulo As String) As String

End Function


Private Function DatosOk() As Boolean
Dim i As Integer
Dim Cad As String


    DatosOk = False
    On Error GoTo SalirDatosOk
    
    For i = 0 To Text1.Count - 1
        Text1(i).Text = Trim(Text1(i).Text)
    Next i
'Comprobamos ciertos datos

For i = 0 To 0
    If Text1(i).Text = "" Then
        MsgBox Label1(i).Caption & " NO puede estar vacia.", vbExclamation
        Exit Function
    End If
Next i

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
vCon.TCP3 = (Check1.Value = 1)
vCon.Ariadna = (Check2.Value = 1)
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
