VERSION 5.00
Begin VB.Form frmInfInc 
   Caption         =   "Informes ordenados por incidencias"
   ClientHeight    =   5640
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   Icon            =   "frmInfInc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEmpleado 
      Height          =   285
      Index           =   0
      Left            =   2280
      TabIndex        =   4
      Top             =   2160
      Width           =   795
   End
   Begin VB.TextBox txtEmpleado 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   3180
      TabIndex        =   32
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox txtEmpleado 
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   5
      Top             =   2520
      Width           =   795
   End
   Begin VB.TextBox txtEmpleado 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   3180
      TabIndex        =   31
      Top             =   2520
      Width           =   3255
   End
   Begin VB.TextBox txtEmpresa 
      Enabled         =   0   'False
      Height          =   285
      Index           =   3
      Left            =   9900
      TabIndex        =   25
      Top             =   1500
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Index           =   2
      Left            =   9000
      TabIndex        =   1
      Top             =   1500
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtEmpresa 
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   9900
      TabIndex        =   24
      Top             =   1140
      Visible         =   0   'False
      Width           =   3315
   End
   Begin VB.TextBox txtEmpresa 
      Height          =   285
      Index           =   0
      Left            =   9000
      TabIndex        =   0
      Top             =   1140
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   5
      Left            =   5340
      TabIndex        =   7
      Top             =   3360
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   4
      Left            =   2160
      TabIndex        =   6
      Top             =   3360
      Width           =   1035
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Default         =   -1  'True
      Height          =   375
      Index           =   1
      Left            =   5220
      TabIndex        =   14
      Top             =   5160
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   3180
      TabIndex        =   19
      Top             =   1380
      Width           =   3315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   2
      Left            =   2280
      TabIndex        =   3
      Top             =   1380
      Width           =   795
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3180
      TabIndex        =   18
      Top             =   960
      Width           =   3315
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   795
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ordenación"
      Height          =   855
      Left            =   3540
      TabIndex        =   16
      Top             =   3900
      Width           =   3015
      Begin VB.OptionButton Option4 
         Caption         =   "Incidencia"
         Height          =   255
         Index           =   1
         Left            =   1500
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Trabajador"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo incidencia"
      Height          =   855
      Left            =   180
      TabIndex        =   15
      Top             =   3900
      Width           =   3255
      Begin VB.OptionButton Option1 
         Caption         =   "Defecto"
         Height          =   255
         Index           =   2
         Left            =   2100
         TabIndex        =   10
         Top             =   360
         Width           =   1035
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Exceso"
         Height          =   255
         Index           =   1
         Left            =   1140
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Todas"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   8
         Top             =   360
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Informe"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   13
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Image imgEmpleado 
      Height          =   240
      Index           =   0
      Left            =   1860
      Picture         =   "frmInfInc.frx":030A
      Top             =   2220
      Width           =   240
   End
   Begin VB.Image imgEmpleado 
      Height          =   240
      Index           =   1
      Left            =   1860
      Picture         =   "frmInfInc.frx":040C
      Top             =   2580
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
      Index           =   2
      Left            =   1080
      TabIndex        =   34
      Top             =   2220
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
      Index           =   2
      Left            =   1080
      TabIndex        =   33
      Top             =   2580
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   30
      Top             =   3060
      Width           =   1335
   End
   Begin VB.Image ImageEmp 
      Height          =   240
      Index           =   1
      Left            =   8460
      Picture         =   "frmInfInc.frx":050E
      Top             =   1560
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image ImageEmp 
      Height          =   240
      Index           =   0
      Left            =   8460
      Picture         =   "frmInfInc.frx":0610
      Top             =   1140
      Visible         =   0   'False
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
      Index           =   1
      Left            =   7680
      TabIndex        =   27
      Top             =   1560
      Visible         =   0   'False
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
      Index           =   1
      Left            =   7680
      TabIndex        =   26
      Top             =   1140
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   5040
      Picture         =   "frmInfInc.frx":0712
      Top             =   3420
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   1860
      Picture         =   "frmInfInc.frx":0814
      Top             =   3420
      Width           =   240
   End
   Begin VB.Label Label5 
      Caption         =   "hasta"
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
      Left            =   4380
      TabIndex        =   23
      Top             =   3420
      Width           =   795
   End
   Begin VB.Label Label4 
      Caption         =   "desde"
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
      Left            =   1080
      TabIndex        =   22
      Top             =   3420
      Width           =   855
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   1
      Left            =   1800
      Picture         =   "frmInfInc.frx":0916
      Top             =   1440
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Index           =   0
      Left            =   1800
      Picture         =   "frmInfInc.frx":0A18
      Top             =   1020
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
      Left            =   1020
      TabIndex        =   21
      Top             =   1440
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
      Left            =   1020
      TabIndex        =   20
      Top             =   1020
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Informes incidencias"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   555
      Index           =   0
      Left            =   840
      TabIndex        =   17
      Top             =   60
      Width           =   4635
   End
   Begin VB.Label Label1 
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   5
      Left            =   7020
      TabIndex        =   28
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Incidencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   1
      Left            =   240
      TabIndex        =   29
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Empleado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   35
      Top             =   1920
      Width           =   1335
   End
End
Attribute VB_Name = "frmInfInc"
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

'Para los labels que aparecerán en el informe
Dim NCampos As Byte 'De 1 a 3
Dim vLabel(5) As String
    'En el informe son 0 campo1, 1 campo 2  .....
    
    
Dim CadParam As String
Dim NParam As Integer
Dim Opc As Integer

Private Sub Command1_Click(Index As Integer)
Dim I As Integer
Dim Cad2
Dim Formula
Dim Cad


On Error GoTo ErrorCommand
Select Case Index
Case 0
    Screen.MousePointer = vbHourglass
     For I = 0 To 5
        vLabel(I) = ""
    Next I
    NCampos = 0
    
    If Option4(0).Value Then
        Formula = DevuelveSQL
        Formula = ""
        Else
            Formula = DevuelveCadenaSQL
    End If
    If Formula = "###" Then
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    'Vemos que si son todas o exceso o continuadas
    'La incidencia de codigo cero(sin ncidencias NO se muestra)
    Cad = ""
    For I = 0 To 5
        If vLabel(I) <> "" Then
           Cad = vLabel(I)
           vLabel(I) = "Campo" & I + 1 & "= """ & Cad & """ "
        End If
    Next I
    
    NParam = 0
    CadParam = ""
    For I = 0 To 5
        If vLabel(I) <> "" Then
            NParam = NParam + 1
            CadParam = CadParam & vLabel(I) & "|"
        End If
    Next I
    
    
    
    
    If Option4(0).Value Then
        'Ordenado por nombre trabajadores
    '    CR1.ReportFileName = App.Path & "\Informes\IncResT.rpt"
    '    Cad = "(Nombre incidencia)"
        Opc = 112
        Else
            'Ordenado por codigo incidencia
    '        CR1.ReportFileName = App.Path & "\Informes\IncResIn.rpt"
    '        Cad = "(Código incidencia)"
            Opc = 113
    End If
    
    
    With frmImprimir
        .FormulaSeleccion = Formula
        .Opcion = Opc
        .NumeroParametros = NParam
        .OtrosParametros = CadParam
        .Show vbModal
    End With
    
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
If frmB.Tabla = "Trabajadores" Then
    txtEmpleado(vIndice).Text = vCodigo
    txtEmpleado(vIndice + 1).Text = vCadena
    If vIndice = 0 Then
        vIndice = 2
        Else
            vIndice = 0
    End If
    Else
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




Private Sub imgEmpleado_Click(Index As Integer)
    vIndice = Index
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "NomTrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPLEADOS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
    If vIndice = 0 Then
        Text1(4).SetFocus
        Else
        txtEmpleado(2).SetFocus
    End If
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



Private Sub txtEmpleado_LostFocus(Index As Integer)
Dim Cad As String
If Trim(txtEmpleado(Index).Text) = "" Then
    txtEmpleado(Index + 1).Text = ""
    Exit Sub
End If
   
If Not IsNumeric(txtEmpleado(Index).Text) Then
    txtEmpleado(Index).Text = "-1"
    txtEmpleado(Index + 1).Text = "Código de empleado erróneo."
    Else
        Cad = devuelveNombreTrabajador(CInt(txtEmpleado(Index).Text))
        If Cad = "" Then
            txtEmpleado(Index).Text = "-1"
            txtEmpleado(Index + 1).Text = "Código de empresa erróneo."
            Else
                txtEmpleado(Index + 1).Text = Cad
        End If
End If
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


Private Function DevuelveSQL() As String
Dim Formula As String
Dim I As Integer
Dim Cad As String
Dim Cad2 As String
Dim CADENA As String
Dim v1, v2
Dim SQL As String

    
        DevuelveSQL = "###"
        For I = 0 To Text1.Count - 1
            Text1(I).Text = Trim(Text1(I).Text)
            Text1(I).Tag = ""
        Next I
        txtEmpresa(0).Tag = ""
        txtEmpresa(1).Tag = ""
        txtEmpresa(2).Tag = ""
        txtEmpresa(3).Tag = ""
        txtEmpleado(0).Tag = ""
        txtEmpleado(1).Tag = ""
        txtEmpleado(2).Tag = ""
        txtEmpleado(3).Tag = ""
        '---------------------------------------------
        '--------------------------------------------
        'Empresa
        CADENA = ""
        v1 = 0
        v2 = 999999999
        I = 0
        Cad = "Empresa desde "
        txtEmpresa(I).Text = Trim(txtEmpresa(I).Text)
        If txtEmpresa(I).Text <> "" Then
            If Not IsNumeric(txtEmpresa(I).Text) Then
                MsgBox Cad & " NO es un código de empresa correcto.", vbExclamation
                Exit Function
                Else
                    v1 = CLng(txtEmpresa(I).Text)
                    If v1 < 0 Then
                        MsgBox Cad & " NO es correcta.", vbExclamation
                        Exit Function
                    End If
                    'Formula = Formula & Nexo & "{" & NomT & ".IdEmpresa} >= " & v1
                    txtEmpresa(0).Tag = "Empresas.IdEmpresa >= " & v1
                    CADENA = CADENA & " desde " & Format(txtEmpresa(I).Text, "00000")
            End If
        End If
        
        
        I = 2
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
                    'Formula = Formula & Nexo & Corchete & NomT & ".IdEmpresa} <= " & v2
                    txtEmpresa(1).Tag = "Empresas.IdEmpresa <= " & v2
                    CADENA = CADENA & " desde " & Format(txtEmpresa(I).Text, "00000")
            End If
        End If
        
        If v1 > v2 Then
            MsgBox "Empresa desde es mayor que empresa hasta. ", vbExclamation
            Exit Function
        End If
        
        CADENA = ""
        '-------------------------------------------------------------------
        'Empleado
        I = 0
        v1 = 0
        v2 = 99999999
        txtEmpleado(I).Text = Trim(txtEmpleado(I).Text)
        Cad = "Empleado desde "
        If txtEmpleado(I).Text <> "" Then
            If Not IsNumeric(txtEmpleado(I).Text) Then
                MsgBox Cad & " NO es numérico.", vbExclamation
                Exit Function
                Else
                    v1 = CLng(txtEmpleado(I).Text)
                    txtEmpleado(I).Tag = "Trabajadores.IdTrabajador >= " & v1
                    CADENA = CADENA & " desde " & Format(txtEmpleado(I).Text, "00000")
            End If
        End If
        'Empleado
        I = 2
        Cad = "Empleado hasta "
        txtEmpleado(I).Text = Trim(txtEmpleado(I).Text)
        If txtEmpleado(I).Text <> "" Then
            If Not IsNumeric(txtEmpleado(I).Text) Then
                MsgBox Cad & " NO es numérico.", vbExclamation
                Exit Function
                Else
                    v2 = CLng(txtEmpleado(I).Text)
                    txtEmpleado(I).Tag = "Trabajadores.IdTrabajador <= " & v2
                    CADENA = CADENA & " hasta " & Format(txtEmpleado(I).Text, "00000")
            End If
        End If
        
        If v1 > v2 Then
            MsgBox "Código empleado 'desde' es mayor que el código de empleado 'hasta'. ", vbExclamation
            Exit Function
        End If
        
        'Para el encabezado del informe
        If CADENA <> "" Then
            vLabel(NCampos) = "Empleado: "
            vLabel(NCampos + 1) = CADENA
            NCampos = NCampos + 2
            CADENA = ""
        End If

        
        
        '----------------------------------------------------------------------
        'Limpiamos los tag y ahora
        'incidencias
        CADENA = ""
        v1 = 0
        v2 = 9999999
        I = 0
        Cad = "Incidencia inicial "
        If Text1(I).Text <> "" Then
            If Not IsNumeric(Text1(I).Text) Then
                MsgBox Cad & " NO es un campo numérico.", vbExclamation
                Exit Function
                Else
                    If CInt(Text1(I).Text) < 0 Then
                        MsgBox "incidencia inicial incorrecta.", vbExclamation
                        Exit Function
                        Else
                            Text1(I).Tag = "IdInci >=" & Text1(I).Text
                            CADENA = CADENA & " desde " & Format(Text1(I).Text, "00000")
                    End If
            End If
        End If
        'Incidencia hasta
        I = 2
        Cad = "Incidencia final "
        If Text1(I).Text <> "" Then
            If Not IsNumeric(Text1(I).Text) Then
                MsgBox Cad & " NO es un campo numérico.", vbExclamation
                Exit Function
                Else
                    If CInt(Text1(I).Text) < 0 Then
                        MsgBox "incidencia final incorrecta.", vbExclamation
                        Exit Function
                        Else
                            Text1(I).Tag = "IdInci <=" & Text1(I).Text
                            CADENA = CADENA & " hasta " & Format(Text1(I).Text, "00000")
                    End If
            End If
        End If
        
        
        
        'Para el encabezado del informe
        If CADENA <> "" Then
            vLabel(NCampos) = "Incidnecia: "
            vLabel(NCampos + 1) = CADENA
            NCampos = NCampos + 2
            CADENA = ""
        End If
        '-----------------------------------------------------
        'Fecha incial
        v1 = "1/1/1900"
        v2 = "31/12/2200"
        I = 4
        Cad = "Fecha inicial "
        If Text1(I).Text <> "" Then
            If Not IsDate(Text1(I).Text) Then
                MsgBox Cad & " NO es una fecha válida.", vbExclamation
                Exit Function
                Else
                    Text1(I).Tag = "fecha >=#" & Format(Text1(I).Text, "yyyy/mm/dd") & "#"
                    CADENA = CADENA & " desde " & Format(Text1(I).Text, "dd/mm/yyyy")
            End If
        End If
        'Fecha incial
        I = 5
        Cad = "Fecha final "
        If Text1(I).Text <> "" Then
            If Not IsDate(Text1(I).Text) Then
                MsgBox Cad & " NO es una fecha válida.", vbExclamation
                Exit Function
                Else
                    Text1(I).Tag = "fecha <=#" & Format(Text1(I).Text, "yyyy/mm/dd") & "#"
                    CADENA = CADENA & " hasta " & Format(Text1(I).Text, "dd/mm/yyyy")
            End If
        End If
        
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


        'Devolvemos la cadena
        'Ahora recorremos los textos para hallar la subconsulta
        Formula = ""
        For I = 0 To txtEmpresa.Count - 1
            If txtEmpresa(I).Tag <> "" Then
                Formula = Formula & " AND " & txtEmpresa(I).Tag & " "
            End If
        Next I
        For I = 0 To txtEmpleado.Count - 1
            If txtEmpleado(I).Tag <> "" Then
                Formula = Formula & " AND " & txtEmpleado(I).Tag & " "
            End If
        Next I
        For I = 0 To Text1.Count - 1
            If Text1(I).Tag <> "" Then
                Formula = Formula & " AND " & Text1(I).Tag & " "
            End If
        Next I
        Cad2 = ""
        If Option1(1).Value Then
            'Queremos ver las de execeso
            Cad2 = "ExcesoDefecto = True"
            CADENA = "EXCESO"
            Else
                If Option1(2).Value Then
                    'LAS de defecto
                    Cad2 = "ExcesoDefecto = False"
                    CADENA = "DEFECTO"
                End If
        End If
        'Para el encabezado del informe
        If Cad2 <> "" Then
            If NCampos < 5 Then
                vLabel(NCampos) = "Tipo Inci.: "
                vLabel(NCampos + 1) = CADENA
                NCampos = NCampos + 2
                CADENA = ""
            End If
            Formula = Formula & " AND " & Cad2
        End If
        DevuelveSQL = "TODO_OK"

        Dim RS As ADODB.Recordset
        Dim RT As ADODB.Recordset
        Set RS = New ADODB.Recordset
        Set RT = New ADODB.Recordset
        
        'Borramos en la tabla temporal
        RS.Open "DELETE * FROM tmpConIncRes", Conn, , , adCmdText
        'Rs.Update
        'Genereamos los nuevos resultados
        SQL = " SELECT Marcajes.*, Empresas.IdEmpresa, Empresas.NomEmpresa, Incidencias.NomInci, Incidencias.ExcesoDefecto, Trabajadores.NomTrabajador,Secciones.Nombre"
        SQL = SQL & " From Empresas, Marcajes, Trabajadores, Incidencias, Secciones"
        SQL = SQL & " WHERE"
        SQL = SQL & " Marcajes.idTrabajador = Trabajadores.idTrabajador AND Empresas.IdEmpresa = Trabajadores.IdEmpresa AND Marcajes.IncFinal = Incidencias.IdInci "
        SQL = SQL & " AND Trabajadores.Seccion=Secciones.IdSeccion"
        SQL = SQL & " AND Marcajes.Correcto = True"
        'Evitamos mostrar la incidencia 0
        SQL = SQL & " AND IncFinal<>0"
        'Ahora si hay subconsulta se pone
        If Formula <> "" Then SQL = SQL & Formula
        'abrimos el recordset
        RS.Open SQL, Conn, , , adCmdText
        'Abrimos el recodset final
        RT.CursorType = adOpenKeyset
        RT.LockType = adLockOptimistic
        RT.Open "SELECT * FROM TmpConIncRes", Conn, , , adCmdText
        I = 1
        While Not RS.EOF
            'Insertamos
            RT.AddNew
            RT!Id = I
            RT!IdEmpresa = RS!IdEmpresa
            RT!NomEmpresa = RS!NomEmpresa
            RT!IdIncidencia = RS!IncFinal
            RT!NomIncidencia = RS!NomInci
            RT!idTrabajador = RS!idTrabajador
            RT!nomtrabajador = RS!nomtrabajador
            RT!Fecha = RS!Fecha
            RT!excesodefecto = RS!excesodefecto
            RT!Seccion = RS!Nombre
            If RS!excesodefecto Then
                RT!HE = RS!HorasIncid
                RT!HD = 0
                Else
                    RT!HD = RS!HorasIncid
                    RT!HE = 0
            End If
            'actualizamos
            RT.Update
            'movemos
            I = I + 1
            RS.MoveNext
        Wend
        'Cerramos los recordsets
        On Error Resume Next
        'Hay que hacer una espera de 1 segundo
        v1 = Timer
        Do
            v2 = Timer
            Loop Until v2 - v1 > 1
        
        RT.Update
        RS.Close
        RT.Close
        Set RS = Nothing
        Set RT = Nothing
End Function



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
            CADENA = CADENA & " hasta " & Format(txtEmpresa(I).Text, "00000")
    End If
End If

If v1 > v2 Then
    MsgBox "Empresa desde es mayor que empresa hasta. ", vbExclamation
    Exit Function
End If
CADENA = ""
'-------------------------------------------------------------------
'Empleado
I = 0
v1 = 0
v2 = 99999999
txtEmpleado(I).Text = Trim(txtEmpleado(I).Text)
Cad = "Empleado desde "
If txtEmpleado(I).Text <> "" Then
    If Not IsNumeric(txtEmpleado(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Function
        Else
            v1 = CLng(txtEmpleado(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".idTrabajador} >=" & txtEmpleado(I).Text
            Nexo = " AND "
            CADENA = CADENA & " desde " & Format(txtEmpleado(I).Text, "00000")
    End If
End If
'Empleado
I = 2
Cad = "Empleado hasta "
txtEmpleado(I).Text = Trim(txtEmpleado(I).Text)
If txtEmpleado(I).Text <> "" Then
    If Not IsNumeric(txtEmpleado(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Function
        Else
            v2 = CLng(txtEmpleado(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".idTrabajador} <=" & txtEmpleado(I).Text
            Nexo = " AND "
            CADENA = CADENA & " hasta " & Format(txtEmpleado(I).Text, "00000")
    End If
End If

If v1 > v2 Then
    MsgBox "Fecha desde es mayor que fecha hasta. ", vbExclamation
    Exit Function
End If

'Para el encabezado del informe
If CADENA <> "" Then
    vLabel(NCampos) = "Empleado: "
    vLabel(NCampos + 1) = CADENA
    NCampos = NCampos + 2
    CADENA = ""
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

