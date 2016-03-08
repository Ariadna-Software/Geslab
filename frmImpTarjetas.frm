VERSION 5.00
Begin VB.Form frmImpTarjetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión con codigo de barras"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8115
   Icon            =   "frmImpTarjetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   8115
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   6960
      TabIndex        =   17
      Top             =   2820
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5820
      TabIndex        =   16
      Top             =   2820
      Width           =   975
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   2535
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   7815
      Begin VB.TextBox txtTarea 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   5040
         TabIndex        =   22
         Top             =   1260
         Width           =   2955
      End
      Begin VB.TextBox txtTarea 
         Height          =   315
         Index           =   2
         Left            =   4380
         TabIndex        =   21
         Top             =   1260
         Width           =   615
      End
      Begin VB.TextBox txtTarea 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1140
         TabIndex        =   20
         Top             =   1260
         Width           =   2775
      End
      Begin VB.TextBox txtTarea 
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   19
         Top             =   1260
         Width           =   615
      End
      Begin VB.Image ImgTarea 
         Height          =   240
         Index           =   1
         Left            =   4920
         Picture         =   "frmImpTarjetas.frx":030A
         Top             =   990
         Width           =   240
      End
      Begin VB.Image ImgTarea 
         Height          =   240
         Index           =   0
         Left            =   900
         Picture         =   "frmImpTarjetas.frx":040C
         Top             =   990
         Width           =   240
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   4380
         TabIndex        =   25
         Top             =   1020
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   24
         Top             =   1020
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Tarea"
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
         Index           =   0
         Left            =   300
         TabIndex        =   23
         Top             =   600
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   2655
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   7995
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   1155
         Left            =   60
         TabIndex        =   7
         Top             =   1440
         Width           =   7935
         Begin VB.TextBox txtIncidencia 
            Height          =   315
            Index           =   0
            Left            =   60
            TabIndex        =   2
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtIncidencia 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   840
            TabIndex        =   9
            Top             =   660
            Width           =   2775
         End
         Begin VB.TextBox txtIncidencia 
            Height          =   315
            Index           =   2
            Left            =   4080
            TabIndex        =   3
            Top             =   660
            Width           =   615
         End
         Begin VB.TextBox txtIncidencia 
            BackColor       =   &H80000018&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   4740
            TabIndex        =   8
            Top             =   660
            Width           =   2955
         End
         Begin VB.Label Label1 
            Caption         =   "Seccion"
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
            Left            =   60
            TabIndex        =   12
            Top             =   0
            Width           =   1695
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Desde"
            Height          =   195
            Index           =   9
            Left            =   60
            TabIndex        =   11
            Top             =   420
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Hasta"
            Height          =   195
            Index           =   8
            Left            =   4080
            TabIndex        =   10
            Top             =   420
            Width           =   420
         End
         Begin VB.Image ImgIncidencia 
            Height          =   240
            Index           =   0
            Left            =   600
            Picture         =   "frmImpTarjetas.frx":050E
            Top             =   390
            Width           =   240
         End
         Begin VB.Image ImgIncidencia 
            Height          =   240
            Index           =   1
            Left            =   4620
            Picture         =   "frmImpTarjetas.frx":0610
            Top             =   390
            Width           =   240
         End
      End
      Begin VB.TextBox txtEmpleado 
         Height          =   285
         Index           =   0
         Left            =   60
         TabIndex        =   0
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtEmpleado 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   720
         TabIndex        =   6
         Top             =   840
         Width           =   2775
      End
      Begin VB.TextBox txtEmpleado 
         Height          =   285
         Index           =   2
         Left            =   4140
         TabIndex        =   1
         Top             =   840
         Width           =   555
      End
      Begin VB.TextBox txtEmpleado 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   285
         Index           =   3
         Left            =   4740
         TabIndex        =   5
         Top             =   840
         Width           =   3015
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   4140
         TabIndex        =   15
         Top             =   600
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   60
         TabIndex        =   14
         Top             =   600
         Width           =   465
      End
      Begin VB.Label Label1 
         Caption         =   "Empleado"
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
         Index           =   3
         Left            =   60
         TabIndex        =   13
         Top             =   180
         Width           =   1695
      End
      Begin VB.Image imgEmpleado 
         Height          =   240
         Index           =   0
         Left            =   600
         Picture         =   "frmImpTarjetas.frx":0712
         Top             =   570
         Width           =   240
      End
      Begin VB.Image imgEmpleado 
         Height          =   240
         Index           =   1
         Left            =   4620
         Picture         =   "frmImpTarjetas.frx":0814
         Top             =   570
         Width           =   240
      End
   End
End
Attribute VB_Name = "frmImpTarjetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
    '0.- Imprime tarjetas TRABAJADORES
    '1.-   "          "    TAREAS
    
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Dim Indice As Integer
Dim vIndex As Integer

Private Sub Command1_Click()
Dim SQL As String


On Error GoTo EImp

    If Opcion = 0 Then
        SQL = ValoresInforme
        If SQL = "###" Then Exit Sub
        nOpcion = 9
    Else
        SQL = ValoresInformeTarea
        If SQL = "###" Then Exit Sub
        nOpcion = 8
    End If
    
    With frmImprimir
        .Opcion = nOpcion + 100
        .NumeroParametros = 0
        .OtrosParametros = SQL
        .FormulaSeleccion = ""
        .Show vbModal
    End With
EImp:
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Frame2.Visible = Opcion = 0
    Me.Frame3.Visible = Opcion = 1
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    Select Case Indice
    Case 0
'        txtEmpresa(vIndex).Text = vCodigo
'        txtEmpresa(vIndex + 1).Text = vCadena
    Case 1
        txtEmpleado(vIndex).Text = vCodigo
        txtEmpleado(vIndex + 1).Text = vCadena
    Case 2
        txtIncidencia(vIndex).Text = vCodigo
        txtIncidencia(vIndex + 1).Text = vCadena
    Case 3
        txtTarea(vIndex).Text = vCodigo
        txtTarea(vIndex + 1).Text = vCadena
    End Select
End Sub

Private Sub imgEmpleado_Click(Index As Integer)
    Indice = 1
    vIndex = (Index * 2)
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "NomTrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPLEADOS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub ImgIncidencia_Click(Index As Integer)

    'Ahora pasa a ser seccion
    Indice = 2
    vIndex = (Index * 2)
    Set frmB = New frmBusca
    frmB.Tabla = "Secciones"
    frmB.CampoBusqueda = "Nombre"
    frmB.CampoCodigo = "IdSeccion"
    frmB.TipoDatos = 3
    frmB.Titulo = "SECCIONES"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
End Sub


Private Sub ImgTarea_Click(Index As Integer)
 'Ahora pasa a ser seccion
    Indice = 3
    vIndex = (Index * 2)
    Set frmB = New frmBusca
    frmB.Tabla = "Tareas"
    frmB.CampoBusqueda = "Descripcion"
    frmB.CampoCodigo = "IdTarea"
    frmB.TipoDatos = 3
    frmB.Titulo = "TAREAS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub txtEmpleado_GotFocus(Index As Integer)
    PonFoco txtEmpleado(Index)
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

Private Sub PonFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub txtIncidencia_GotFocus(Index As Integer)
    PonFoco Me.txtIncidencia(Index)
End Sub

Private Sub txtIncidencia_LostFocus(Index As Integer)
Dim Cad As String
If Trim(txtIncidencia(Index).Text) = "" Then
    txtIncidencia(Index + 1).Text = ""
    Exit Sub
End If
   
If Not IsNumeric(txtIncidencia(Index).Text) Then
    txtIncidencia(Index).Text = "-1"
    txtIncidencia(Index + 1).Text = "Código de sección erróneo."
    Else
        Cad = DevuelveNombreSeccion(CLng(txtIncidencia(Index).Text))
        If Cad = "" Then
            txtIncidencia(Index).Text = "-1"
            txtIncidencia(Index + 1).Text = "Código de sección erróneo."
            Else
                txtIncidencia(Index + 1).Text = Cad
        End If
End If
End Sub


Private Function ValoresInformeTarea() As String
Dim Cad As String

    ValoresInformeTarea = "###"
    txtTarea(0).Text = Trim(txtTarea(0).Text)
    txtTarea(2).Text = Trim(txtTarea(2).Text)
    
    If txtTarea(0).Text <> "" And txtTarea(2).Text <> "" Then
        If Val(txtTarea(0).Text) > Val((txtTarea(2).Text)) Then
            MsgBox "Tarea desde mayor que tarea hasta.", vbExclamation
            Exit Function
        End If
    End If
    Cad = ""
    If txtTarea(0).Text <> "" Then Cad = "{ado.idtarea} >= " & Me.txtTarea(0).Text
    If txtTarea(2).Text <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & "{ado.idtarea} <= " & Me.txtTarea(2).Text
    End If
    ValoresInformeTarea = Cad
End Function


Private Function ValoresInforme() As String
Dim I As Integer
Dim Cad As String
Dim Formula As String
Dim Nexo As String
Dim CADENA As String
Dim v1, v2
Dim NombreTabla As String



ValoresInforme = "###"
'LimpiarTags
Formula = ""
Nexo = ""
NombreTabla = "ado"
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
    MsgBox "Empleado desde es mayor que empleado hasta. ", vbExclamation
    Exit Function
End If




'Seccion
I = 0
v1 = 0
v2 = 99999999
txtIncidencia(I).Text = Trim(txtIncidencia(I).Text)
Cad = "Seccion desde "
If txtIncidencia(I).Text <> "" Then
    If Not IsNumeric(txtIncidencia(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Function
        Else
            v1 = CLng(txtIncidencia(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".seccion} >=" & txtIncidencia(I).Text
            Nexo = " AND "
            CADENA = CADENA & " desde " & Format(txtIncidencia(I).Text, "00000")
    End If
End If
'Empleado
I = 2
Cad = "Seccion hasta "
txtIncidencia(I).Text = Trim(txtIncidencia(I).Text)
If txtIncidencia(I).Text <> "" Then
    If Not IsNumeric(txtIncidencia(I).Text) Then
        MsgBox Cad & " NO es numérico.", vbExclamation
        Exit Function
        Else
            v2 = CLng(txtIncidencia(I).Text)
            Formula = Formula & Nexo & "{" & NombreTabla & ".seccion} <=" & txtIncidencia(I).Text
            Nexo = " AND "
            CADENA = CADENA & " hasta " & Format(txtIncidencia(I).Text, "00000")
    End If
End If

If v1 > v2 Then
    MsgBox "Seccion desde es mayor que seccion hasta. ", vbExclamation
    Exit Function
End If



ValoresInforme = Formula
End Function

Private Sub txtTarea_GotFocus(Index As Integer)
    PonFoco txtTarea(Index)
End Sub

Private Sub txtTarea_LostFocus(Index As Integer)
Dim Cad As String
    If Trim(txtTarea(Index).Text) = "" Then
        txtTarea(Index + 1).Text = ""
        Exit Sub
    End If
       
    If Not IsNumeric(txtTarea(Index).Text) Then
        txtTarea(Index).Text = "-1"
        txtTarea(Index + 1).Text = "Código de tarea erróneo."
        Else
            Cad = DevuelveDesdeBD("descripcion", "tareas", "idTarea", txtTarea(Index).Text, "N")
            If Cad = "" Then
                txtTarea(Index).Text = "-1"
                txtTarea(Index + 1).Text = "Código de sección erróneo."
                Else
                    txtTarea(Index + 1).Text = Cad
            End If
    End If
End Sub
