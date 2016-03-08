VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmInfIncGen2 
   Caption         =   "Informes incidencias generadas"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6450
   Icon            =   "frmInfIncGen2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Index           =   0
      Left            =   8280
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmpresa 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   9180
      TabIndex        =   26
      Top             =   1320
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Index           =   2
      Left            =   8280
      TabIndex        =   1
      Top             =   1680
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtEmpresa 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   9180
      TabIndex        =   25
      Top             =   1680
      Visible         =   0   'False
      Width           =   3315
   End
   Begin Crystal.CrystalReport CR1 
      Left            =   480
      Top             =   4920
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
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   4860
      TabIndex        =   5
      Top             =   2460
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   2100
      TabIndex        =   4
      Top             =   2460
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   4860
      TabIndex        =   12
      Top             =   5100
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   2640
      TabIndex        =   17
      Top             =   1440
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   1800
      TabIndex        =   3
      Top             =   1440
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   2640
      TabIndex        =   16
      Top             =   1020
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1800
      TabIndex        =   2
      Top             =   1020
      Width           =   675
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordenación"
      Height          =   855
      Left            =   180
      TabIndex        =   14
      Top             =   4020
      Width           =   6075
      Begin VB.OptionButton Option4 
         Caption         =   "Código incidencia"
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   10
         Top             =   360
         Width           =   1635
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Nombre trabajador"
         Height          =   255
         Index           =   0
         Left            =   1020
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo"
      Height          =   795
      Left            =   180
      TabIndex        =   13
      Top             =   3120
      Width           =   6075
      Begin VB.OptionButton Option1 
         Caption         =   "Defecto"
         Height          =   255
         Index           =   2
         Left            =   4500
         TabIndex        =   8
         Top             =   360
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Exceso"
         Height          =   255
         Index           =   1
         Left            =   2460
         TabIndex        =   7
         Top             =   360
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   420
         TabIndex        =   6
         Top             =   360
         Value           =   -1  'True
         Width           =   1035
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Informe"
      Height          =   375
      Index           =   0
      Left            =   3300
      TabIndex        =   11
      Top             =   5100
      Width           =   1455
   End
   Begin VB.Image ImageEmp 
      Height          =   240
      Index           =   0
      Left            =   7920
      Picture         =   "frmInfIncGen2.frx":030A
      Top             =   1320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageEmp 
      Height          =   240
      Index           =   1
      Left            =   7920
      Picture         =   "frmInfIncGen2.frx":040C
      Top             =   1740
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7140
      TabIndex        =   28
      Top             =   1320
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   7140
      TabIndex        =   27
      Top             =   1740
      Visible         =   0   'False
      Width           =   855
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
      Index           =   1
      Left            =   180
      TabIndex        =   24
      Top             =   2040
      Width           =   1335
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
      Left            =   6720
      TabIndex        =   23
      Top             =   780
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Incidencia"
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
      Left            =   180
      TabIndex        =   22
      Top             =   600
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   4320
      Picture         =   "frmInfIncGen2.frx":050E
      Top             =   2520
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   1560
      Picture         =   "frmInfIncGen2.frx":0610
      Top             =   2520
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   21
      Top             =   2520
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   780
      TabIndex        =   20
      Top             =   2520
      Width           =   915
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   1440
      Picture         =   "frmInfIncGen2.frx":0712
      Top             =   1500
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   1440
      Picture         =   "frmInfIncGen2.frx":0814
      Top             =   1080
      Width           =   240
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   19
      Top             =   1500
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Desde"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   18
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Informes incidencias generadas"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Index           =   0
      Left            =   780
      TabIndex        =   15
      Top             =   180
      Width           =   4695
   End
End
Attribute VB_Name = "frmInfIncGen2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1

Private NombreTabla As String
Private vIndice As Integer  'Indice de text incidencia
Private vI As Integer  'Indice de text  hora
Private ImgEmpr As Boolean
Private vLabel(5) As String
Private NCampos As Integer

Private Sub Command1_Click(Index As Integer)
Dim Formula As String
Dim I As Integer
Dim Cad As String
Dim Cad2 As String
Dim Etiq As String

On Error GoTo ErrorCommand
Select Case Index
        Case 0


        'Devolvemos la cadena
        'Ahora recorremos los textos para hallar la subconsulta

    NCampos = 0
    For I = 0 To 5
        vLabel(I) = ""
    Next I

    Formula = DevuelveCadenaSQL

    If Formula = "###" Then Exit Sub
    Formula = Formula & Cad2
    Screen.MousePointer = vbHourglass
    CR1.Connect = Conn
    If Option4(0).Value Then
        'Ordenado por nombre
        CR1.ReportFileName = App.Path & "\Informes\incgenN.rpt"
        Cad = "(Nombre incidencia)"
        Else
            'Ordenado por codigo
            CR1.ReportFileName = App.Path & "\Informes\incgenc.rpt"
            Cad = "(Código incidencia)"
    End If
    CR1.SelectionFormula = Formula
    
    'Los campos
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
    
    CR1.WindowTitle = "Resumen horas incidencias generadas" & Cad
    CR1.WindowState = crptMaximized
    CR1.Action = 1
    Screen.MousePointer = vbDefault
Case 1
    Unload Me
End Select
Exit Sub
ErrorCommand:
    If Err.Number <> 0 Then MsgBox "Error: " & Err.Description, vbExclamation
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
NombreTabla = "ado"
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
If ImgEmpr Then
    txtEmpresa(vIndice).Text = vCodigo
    txtEmpresa(vIndice + 1).Text = vCadena
    If vIndice = 0 Then
            vIndice = 2
            Else
                vIndice = 3
        End If
    Else
        'Incidencia o fecha
        Text1(vIndice).Text = vCodigo
        Text1(vIndice + 1).Text = vCadena
        If vIndice = 0 Then
            vIndice = 2
            Else
                vIndice = 4
        End If
End If
End Sub

Private Sub frmC_Selec(vFecha As Date)
Text1(vI).Text = Format(vFecha, "dd/mm/yyyy")
'Para pasar el set focus a laguien
If vI = 4 Then
    vI = 5
    Else
        vI = 0
End If
End Sub

Private Sub Image1_Click(Index As Integer)
    ImgEmpr = False
    vIndice = Index * 2
    Set frmB = New frmBusca
    frmB.Tabla = "Incidencias"
    frmB.CampoBusqueda = "NomInci"
    frmB.CampoCodigo = "IdInci"
    frmB.TipoDatos = 3
    frmB.Titulo = "INCIDENCIAS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
    If vIndice > 0 Then Text1(vIndice).SetFocus
End Sub

Private Sub Image2_Click(Index As Integer)
vI = Index + 4
Set frmC = New frmCal
frmC.Fecha = Now
frmC.Show vbModal
Set frmC = Nothing
If vI > 0 Then Text1(vI).SetFocus
End Sub

Private Sub ImageEmp_Click(Index As Integer)
    ImgEmpr = True
    vIndice = Index * 2
    Set frmB = New frmBusca
    frmB.Tabla = "Empresas"
    frmB.CampoBusqueda = "NomEmpresa"
    frmB.CampoCodigo = "IdEmpresa"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPRESAS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
    Select Case vIndice
    Case 2
            txtEmpresa(2).SetFocus
    Case 3
            Text1(0).SetFocus
    End Select
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String
Select Case Index
Case 0
    'Incidencia desde
    If Text1(0).Text = "" Then
            Text1(0).Text = ""
            Text1(1).Text = ""
            Exit Sub
    End If
    If Not IsNumeric(Text1(0).Text) Then
        Text1(0).Text = -1
        Text1(1).Text = "Error en la incidencia"
        Else
            Cad = DevuelveTextoIncidencia(CInt(Text1(0).Text))
            If Cad = "" Then
                Text1(0).Text = -1
                Text1(1).Text = "Error en la incidencia"
                Else
                    Text1(1).Text = Cad
            End If
    End If
Case 2
    If Text1(2).Text = "" Then
        Text1(2).Text = ""
        Text1(3).Text = ""
        Exit Sub
    End If
    If Not IsNumeric(Text1(2).Text) Then
        Text1(2).Text = -1
        Text1(3).Text = "Error en la incidencia"
        Else
            Cad = DevuelveTextoIncidencia(CInt(Text1(2).Text))
            If Cad = "" Then
                Text1(2).Text = -1
                Text1(3).Text = "Error en la incidencia"
                Else
                    Text1(3).Text = Cad
            End If
    End If
Case 4, 5
    If Not EsFechaOK(Text1(Index)) Then Text1(Index).Text = ""
    
End Select
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
            txtEmpresa(Index).Text = "-1"
            txtEmpresa(Index + 1).Text = "Código de empresa erróneo."
            Else
                txtEmpresa(Index + 1).Text = Cad
        End If
End If
End Sub



Private Function DevuelveCadenaSQL() As String
Dim I As Integer
Dim Cad As String
Dim Formula As String
Dim Nexo As String
Dim CADENA As String
Dim v1, v2
Dim Cad2

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
I = 4
CADENA = ""
Cad = "Fecha desde "
Text1(I).Text = Trim(Text1(I).Text)
If Text1(I).Text <> "" Then
    If Not IsDate(Text1(I).Text) Then
        MsgBox Cad & " NO es un fecha correcta.", vbExclamation
        Exit Function
        Else
            v1 = CDate(Text1(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".Fecha} >=#" & Format(Text1(I).Text, "yyyy/mm/dd") & "#"
            Nexo = " AND "
            CADENA = CADENA & " desde " & Format(Text1(I).Text, "dd/mm/yyyy")
    End If
End If
'fecha hasta
I = 5
Cad = "Fecha hasta "
Text1(I).Text = Trim(Text1(I).Text)
If Text1(I).Text <> "" Then
    If Not IsDate(Text1(I).Text) Then
        MsgBox Cad & " NO es un fecha correcta.", vbExclamation
        Exit Function
        Else
            v2 = CDate(Text1(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".Fecha} <=#" & Format(Text1(I).Text, "yyyy/mm/dd") & "#"
            Nexo = " AND "
            CADENA = CADENA & " hasta " & Format(Text1(I).Text, "dd/mm/yyyy")
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

'-----------------------------------------------------------------------
'INCIDENCIA
I = 0
v1 = 0
v2 = 99999999
Cad = "Incidencia desde "
Text1(I).Text = Trim(Text1(I).Text)
If Text1(I).Text <> "" Then
    If Not IsNumeric(Text1(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Function
        Else
            v1 = CLng(Text1(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".idInci} >=" & Text1(I).Text
            Nexo = " AND "
            CADENA = CADENA & " desde " & Format(Text1(I).Text, "00000")
    End If
End If
'Incidencias
I = 2
Cad = "Incidencia hasta "
Text1(I).Text = Trim(Text1(I).Text)
If Text1(I).Text <> "" Then
    If Not IsNumeric(Text1(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Function
        Else
            v2 = CLng(Text1(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".idInci} <=" & Text1(I).Text
            Nexo = " AND "
            CADENA = CADENA & " hasta " & Format(Text1(I).Text, "00000")
    End If
End If

If v1 > v2 Then
    MsgBox "Incidencia desde es mayor que fecha hasta. ", vbExclamation
    Exit Function
End If

'Para el encabezado del informe
If CADENA <> "" Then
    vLabel(NCampos) = "Incidencias: "
    vLabel(NCampos + 1) = CADENA
    NCampos = NCampos + 2
    CADENA = ""
End If



If v1 > v2 Then
    MsgBox "Fecha desde es mayor que fecha hasta. ", vbExclamation
    Exit Function
End If


        'Frame de todas, execeso defecto
        
        If Option1(1).Value Then
            'Queremos ver las de execeso
            Cad2 = " {" & NombreTabla & ".ExcesoDefecto} = True"
            CADENA = "EXCESO"
            Else
                If Option1(2).Value Then
                    'LAS de defecto
                    Cad2 = " {" & NombreTabla & ".ExcesoDefecto} = False"
                    CADENA = "DEFECTO"
                End If
        End If
        'Para el encabezado del informe
        If CADENA <> "" Then
            vLabel(NCampos) = "Tipo Inci.: "
            vLabel(NCampos + 1) = CADENA
            NCampos = NCampos + 2
            CADENA = ""
            If Formula <> "" Then Cad2 = " AND " & Cad2
            Formula = Formula & Cad2
        End If
       'Vemos que si son todas o exceso o continuadas
       'La incidencia de codigo cero(sin ncidencias NO se muestra)
        Cad2 = "{" & NombreTabla & ".IdInci}<>0"
        If Formula <> "" Then Cad2 = " AND " & Cad2
        Formula = Formula & Cad2




DevuelveCadenaSQL = Formula
End Function
