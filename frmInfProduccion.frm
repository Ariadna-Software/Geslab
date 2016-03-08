VERSION 5.00
Begin VB.Form frmInfProduccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes producción"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6480
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "Tarea"
      Height          =   315
      Index           =   1
      Left            =   2040
      TabIndex        =   24
      Top             =   4560
      Width           =   975
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Trabajador"
      Height          =   315
      Index           =   3
      Left            =   5160
      TabIndex        =   23
      Top             =   4560
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Trabajador/Tarea"
      Height          =   315
      Index           =   2
      Left            =   3120
      TabIndex        =   22
      Top             =   4560
      Width           =   1815
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Tarea / Trabajador"
      Height          =   315
      Index           =   0
      Left            =   120
      TabIndex        =   21
      Top             =   4560
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Aceptar"
      Height          =   435
      Left            =   3600
      TabIndex        =   6
      Top             =   5040
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   435
      Left            =   5040
      TabIndex        =   20
      Top             =   5040
      Width           =   1275
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4020
      Width           =   1095
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Index           =   1
      Left            =   4260
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   4020
      Width           =   1155
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Index           =   1
      Left            =   1920
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   2940
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3060
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   2940
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Index           =   0
      Left            =   1920
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2220
      Width           =   975
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   3060
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   2220
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   1
      Left            =   1860
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   1140
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   3000
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   1140
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Index           =   0
      Left            =   1860
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   420
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Index           =   0
      Left            =   3000
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   420
      Width           =   3135
   End
   Begin VB.Label Label5 
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
      Height          =   360
      Index           =   2
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   900
   End
   Begin VB.Label Label5 
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
      Height          =   360
      Index           =   1
      Left            =   120
      TabIndex        =   18
      Top             =   1980
      Width           =   825
   End
   Begin VB.Label Label5 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Trabajador"
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
      Height          =   360
      Index           =   0
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   1530
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicio"
      Height          =   195
      Left            =   1920
      TabIndex        =   16
      Top             =   3780
      Width           =   855
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicio"
      Height          =   195
      Left            =   4260
      TabIndex        =   15
      Top             =   3780
      Width           =   855
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   2820
      Picture         =   "frmInfProduccion.frx":0000
      Top             =   3720
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   5160
      Picture         =   "frmInfProduccion.frx":0102
      Top             =   3720
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   1
      Left            =   1920
      TabIndex        =   14
      Top             =   2700
      Width           =   420
   End
   Begin VB.Image ImageTarea 
      Height          =   240
      Index           =   1
      Left            =   2520
      Picture         =   "frmInfProduccion.frx":0204
      Top             =   2700
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   12
      Top             =   1980
      Width           =   465
   End
   Begin VB.Image ImageTarea 
      Height          =   240
      Index           =   0
      Left            =   2520
      Picture         =   "frmInfProduccion.frx":0306
      Top             =   1980
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Hasta"
      Height          =   195
      Index           =   1
      Left            =   1860
      TabIndex        =   10
      Top             =   900
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   1860
      TabIndex        =   8
      Top             =   180
      Width           =   465
   End
   Begin VB.Image ImageTrabajador 
      Height          =   240
      Index           =   0
      Left            =   2400
      Picture         =   "frmInfProduccion.frx":0408
      Top             =   180
      Width           =   240
   End
   Begin VB.Image ImageTrabajador 
      Height          =   240
      Index           =   1
      Left            =   2340
      Picture         =   "frmInfProduccion.frx":050A
      Top             =   900
      Width           =   240
   End
End
Attribute VB_Name = "frmInfProduccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
Dim Cad As String


Private Sub Keypress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Sub PonFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub




Private Sub GotFocus(ByRef T As TextBox)
    With T
        T.SelStart = 0
        T.SelLength = Len(T.Text)
    End With
End Sub


Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()
Dim FECHAS As String
Dim Texto As String
Dim TextoDef As String
Dim TextoDef2 As String

    On Error GoTo EMinforme
    
    
    
    TextoDef = ""
    TextoDef2 = ""
    Cad = ""
    '----------------------------------------------------
    'Desde hasta  fecha
    If Text5(0).Text <> "" And Text5(1).Text <> "" Then
        If CDate(Text5(0).Text) > CDate(Text5(1).Text) Then
            MsgBox "Fecha incio mayor fecha fin", vbExclamation
            Exit Sub
        End If
    End If
    
    FECHAS = ""
    Texto = ""
    If Text5(0).Text <> "" Then
        FECHAS = "{ado.fecha} >= #" & Format(Text5(0).Text, "yyyy/mm/dd") & "#"
        Texto = "Desde " & Text5(0).Text
    End If
    If Text5(1).Text <> "" Then
        If FECHAS <> "" Then FECHAS = FECHAS & " AND "
        FECHAS = FECHAS & "{ado.fecha} <= #" & Format(Text5(1).Text, "yyyy/mm/dd") & "#"
        If Texto = "" Then
            Texto = "H"
        Else
            Texto = Texto & "  h"
        End If
        Texto = Texto & "asta " & Text5(1).Text
    End If
       
    'Formula seleccion
    Cad = FECHAS
    If Texto <> "" Then
        Texto = "Fechas: " & Texto
        TextoDef = TextoDef & Texto
    End If
    
    
    
    '-----------------------------------------------------------------------
    'Desde hasta TAREA
    If Text3(0).Text <> "" And Text3(1).Text <> "" Then
        If Val(Text3(0).Text) > Val(Text3(1).Text) Then
            MsgBox "Tarea incio mayor tarea fin", vbExclamation
            Exit Sub
        End If
    End If
    
    FECHAS = ""
    Texto = ""
    If Text3(0).Text <> "" Then
        FECHAS = "{ado.Tarea} >= " & Text3(0).Text
        Texto = "Desde " & Text3(0).Text
    End If
    If Text3(1).Text <> "" Then
        If FECHAS <> "" Then FECHAS = FECHAS & " AND "
        FECHAS = FECHAS & "{ado.tarea} <= " & Text3(1).Text
        If Texto = "" Then
            Texto = "H"
        Else
            Texto = Texto & "  h"
        End If
        Texto = Texto & "asta " & Text3(1).Text
    End If
       
    'Formula seleccion
    If FECHAS <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & FECHAS
    End If
    If Texto <> "" Then
        If TextoDef <> "" Then TextoDef = TextoDef & "   -   "
        TextoDef = TextoDef & " Tarea " & Texto
    End If
    
    'LA TAREA SALIDA NO LA CUENTO
    'La tarea cuyo tipo es el 1 es la tarea SALIDA
    Texto = DevuelveDesdeBD("idtarea", "tareas", "tipo", "1", "N")
    Texto = " AND {ado.tarea} <> " & Texto
    Cad = Cad & Texto
    
    '-----------------------------------------------------------------------
    'Desde hasta TRABAJADOR
    If Text1(0).Text <> "" And Text1(1).Text <> "" Then
        If Val(Text1(0).Text) > Val(Text1(1).Text) Then
            MsgBox "Trabajador incio mayor trabajador fin", vbExclamation
            Exit Sub
        End If
    End If
    
    FECHAS = ""
    Texto = ""
    If Text1(0).Text <> "" Then
        FECHAS = "{ado.idTrabajador} >= " & Text1(0).Text
        Texto = "Desde " & Text1(0).Text & " - " & Text2(0).Text
    End If
    If Text1(1).Text <> "" Then
        If FECHAS <> "" Then FECHAS = FECHAS & " AND "
        FECHAS = FECHAS & "{ado.idtrabajador} <= " & Text1(1).Text
        If Texto = "" Then
            Texto = "H"
        Else
            Texto = Texto & "  h"
        End If
        Texto = Texto & "asta " & Text1(1).Text & " - " & Text2(1).Text
    End If
       
    'Formula seleccion
    If FECHAS <> "" Then
        If Cad <> "" Then Cad = Cad & " AND "
        Cad = Cad & FECHAS
    End If
    'Va en textodef2. Sera otra linea en el informe
    If Texto <> "" Then
        TextoDef2 = "Trabajador    " & Texto
    End If
    
    
    
    
    
    
    FECHAS = ""
'    If Option1(0).Value Then
'        FECHAS = App.Path & "\Informes\ProdCatFec"
'    Else
'        If Option1(1).Value Then
'            FECHAS = App.Path & "\Informes\ProdCatFec"
'        Else
'            FECHAS = App.Path & "\Informes\ProdCatFec"
'        End If
'    End If
'    If Check1.Value = 1 Then
'        'kiere mostrar los trabajdores
'        FECHAS = FECHAS & "S"
'    End If
'    FECHAS = FECHAS & ".rpt"
'    CR1.ReportFileName = FECHAS
    
    nOpcion = 0
    For NParam = 0 To Option1.Count - 1
        If Option1(NParam).Value Then nOpcion = NParam
    Next NParam
    
    
    
    'Captions
    FECHAS = "Fecha= """ & TextoDef & """|"
    FECHAS = FECHAS & "Seleccion= """ & TextoDef2 & """|"
    
    With frmImprimir
        .Opcion = nOpcion + 145
        .NumeroParametros = 2
        .FormulaSeleccion = Cad
        .OtrosParametros = FECHAS
        '.FormulaSeleccion = ""
        .Show vbModal
    End With

    Exit Sub
EMinforme:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub Form_Load()

    Text1(0).Text = ""
    Text2(0).Text = ""
    Text3(0).Text = ""
    Text4(0).Text = ""
    Text1(1).Text = ""
    Text2(1).Text = ""
    Text3(1).Text = ""
    Text4(1).Text = ""
    Text5(0).Text = Format(Now - 1, "dd/mm/yyyy")
    Text5(1).Text = Format(Now - 1, "dd/mm/yyyy")
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
Dim RC As Integer
    RC = Val(ImageTrabajador(1).Tag)
    If ImageTrabajador(0).Tag = 0 Then
        Text1(RC).Text = vCodigo
        Text2(RC).Text = vCadena
    Else
        Text3(RC).Text = vCodigo
        Text4(RC).Text = vCadena
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text5(Val(Text5(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Text5(Index).Text <> "" Then
        If IsDate(Text5(Index).Text) Then frmF.Fecha = CDate(Text5(Index).Text)
    End If
    Text5(0).Tag = Index
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub ImageTarea_Click(Index As Integer)
    ImageTrabajador(0).Tag = 1
    ImageTrabajador(1).Tag = Index
    Set frmB = New frmBusca
    frmB.Tabla = "Tareas"
    frmB.CampoBusqueda = "Descripcion"
    frmB.CampoCodigo = "IdTarea"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "TAREAS"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub ImageTrabajador_Click(Index As Integer)
    ImageTrabajador(0).Tag = 0
    ImageTrabajador(1).Tag = Index
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "Nomtrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "TRABAJADORES"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub



Private Sub Text1_GotFocus(Index As Integer)
    PonFoco Text1(Index)
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).Text = Trim(Text1(Index).Text)
    If Text1(Index).Text = "" Then
        Cad = ""
    Else
        If Not IsNumeric(Text1(Index).Text) Then
            MsgBox "Campo deberia ser numérico: " & Text1(Index).Text, vbExclamation
            Cad = ""
            Text1(Index).Text = ""
        Else
            Cad = DevuelveDesdeBD("nomTrabajador", "Trabajadores", "idTrabajador", Text1(Index).Text, "N")
            'Antes
            'If Cad = "" Then MsgBox "Ningún trabajador con ese código : " & Text1(Index).Text, vbExclamation
            If Cad = "" Then Cad = "INEXISTENTE"
        End If
        If Cad = "" Then PonFoco Text1(Index)
    End If
    Text2(Index).Text = Cad
End Sub



Private Sub Text3_GotFocus(Index As Integer)
    PonFoco Text3(Index)
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index).Text)
    If Text3(Index).Text = "" Then
        Cad = ""
    Else
        If Not IsNumeric(Text3(Index).Text) Then
            MsgBox "Campo deberia ser numérico: " & Text3(Index).Text, vbExclamation
            Cad = ""
            Text3(Index).Text = ""
        Else
            Cad = DevuelveDesdeBD("descripcion", "Tareas", "idTarea", Text3(Index).Text, "N")
            'If Cad = "" Then MsgBox "Ninguna tarea con ese código : " & Text3(Index).Text, vbExclamation
            If Cad = "" Then Cad = "TAREA INEXISTENTE"
        End If
        If Cad = "" Then PonFoco Text3(Index)
    End If
    Text4(Index).Text = Cad
End Sub

Private Sub Text5_GotFocus(Index As Integer)
    GotFocus Text5(Index)
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub


Private Sub Text5_LostFocus(Index As Integer)

    If Not EsFechaOK(Text5(Index)) Then Text5(Index).Text = ""

'    With Text5(Index)
'        .Text = Trim(.Text)
'        If .Text <> "" Then
'            If Not IsDate(.Text) Then
'                MsgBox "No es una fecha correcta: " & .Text, vbExclamation
'                .Text = ""
'                PonFoco Text5(Index)
'            Else
'                .Text = Format(.Text, "dd/mm/yyyy")
'            End If
'        End If
'    End With
'
End Sub
