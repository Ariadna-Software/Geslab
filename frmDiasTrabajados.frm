VERSION 5.00
Begin VB.Form frmDiasTrabajados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   5115
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIncidencia 
      Enabled         =   0   'False
      Height          =   315
      Index           =   3
      Left            =   2160
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   2730
      Width           =   2835
   End
   Begin VB.TextBox txtIncidencia 
      Height          =   315
      Index           =   2
      Left            =   1440
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   2730
      Width           =   615
   End
   Begin VB.TextBox txtIncidencia 
      Enabled         =   0   'False
      Height          =   315
      Index           =   1
      Left            =   2220
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   2070
      Width           =   2775
   End
   Begin VB.TextBox txtIncidencia 
      Height          =   315
      Index           =   0
      Left            =   1440
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   2100
      Width           =   615
   End
   Begin VB.CheckBox Check4 
      Caption         =   "No trabajados"
      Height          =   255
      Left            =   2880
      TabIndex        =   11
      Top             =   3540
      Width           =   1935
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Ordenado por trabajador"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   4020
      Value           =   1  'Checked
      Width           =   2415
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Previsualizar"
      Height          =   255
      Left            =   300
      TabIndex        =   9
      Top             =   4680
      Value           =   1  'Checked
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Solo CONTROL tipo 2"
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   3540
      Width           =   2295
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   4620
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   4620
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   0
      Left            =   1440
      TabIndex        =   0
      Top             =   1080
      Width           =   1155
   End
   Begin VB.Image ImgIncidencia 
      Height          =   240
      Index           =   1
      Left            =   1140
      Picture         =   "frmDiasTrabajados.frx":0000
      Top             =   2820
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   8
      Left            =   660
      TabIndex        =   19
      Top             =   2820
      Width           =   420
   End
   Begin VB.Image ImgIncidencia 
      Height          =   240
      Index           =   0
      Left            =   1140
      Picture         =   "frmDiasTrabajados.frx":0102
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   9
      Left            =   600
      TabIndex        =   16
      Top             =   2220
      Width           =   465
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
      Left            =   180
      TabIndex        =   13
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   615
      Left            =   2880
      TabIndex        =   12
      Top             =   3900
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "Dias trabajados"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   375
      Left            =   360
      TabIndex        =   5
      Top             =   120
      Width           =   3795
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
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   1
      Left            =   3480
      Picture         =   "frmDiasTrabajados.frx":0204
      Top             =   780
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Index           =   0
      Left            =   2040
      Picture         =   "frmDiasTrabajados.frx":0306
      Top             =   840
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Hasta"
      Height          =   195
      Index           =   5
      Left            =   3000
      TabIndex        =   3
      Top             =   840
      Width           =   420
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Desde"
      Height          =   195
      Index           =   4
      Left            =   1440
      TabIndex        =   2
      Top             =   840
      Width           =   465
   End
End
Attribute VB_Name = "frmDiasTrabajados"
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
Dim Trab As ADODB.Recordset


Private Sub cmdImprimir_Click()
On Error GoTo EImp

txtIncidencia(0).Text = Trim(txtIncidencia(0).Text)
txtIncidencia(2).Text = Trim(txtIncidencia(2).Text)
If Val(txtIncidencia(0).Text) > Val(txtIncidencia(0).Text) Then
    MsgBox "Seccion desde mayor seccion hasta.", vbExclamation
    Exit Sub
End If


Screen.MousePointer = vbHourglass
If Not generadias Then
    Screen.MousePointer = vbDefault
    Exit Sub
End If


If Check4.Value = 0 Then
    Label4.Visible = True
    If Not GeneratemporalDiasTr Then
        Screen.MousePointer = vbDefault
        Label4.Visible = False
        Exit Sub
    End If
    
    espera 1
    'K informe
    If Check3.Value = 0 Then
        nOpcion = 1
    Else
        nOpcion = 2
    End If
    
    
    'Seleccion de registro
    If Check1.Value = 1 Then
        Cad = "{ado.Control}= 2"
    Else
        Cad = "{ado.Control}>0"
    End If
    
    
    
   
Else

    Label4.Visible = True
    If Not Generatemporal Then
        Screen.MousePointer = vbDefault
        Label4.Visible = False
        Exit Sub
    End If


    'Informmes para los no trabajados
    If Check3.Value = 1 Then
        nOpcion = 3
    Else
        nOpcion = 4
    End If
    Cad = ""
End If

    CadParam = Cad
    'Fechas
    Cad = ""
    If txtFecha(0).Text <> "" Then Cad = Cad & "Desde " & txtFecha(0).Text
    If txtFecha(0).Text <> "" Then
        If Cad <> "" Then Cad = Cad & "   "
        Cad = Cad & "Hasta " & txtFecha(1).Text
    End If
    If Me.txtIncidencia(0).Text <> "" Then Cad = Cad & " Desde seccion " & Me.txtIncidencia(0).Text & " - " & txtIncidencia(1).Text
    If Me.txtIncidencia(2).Text <> "" Then Cad = Cad & " Hasta seccion " & Me.txtIncidencia(2).Text & " - " & txtIncidencia(3).Text

    
    
    
    Cad = "Desc= """ & Cad & """"


    With frmImprimir
        .FormulaSeleccion = CadParam
        .Opcion = 100 + nOpcion
        .NumeroParametros = 1
        .OtrosParametros = Cad
        .Show vbModal
    End With

EImp:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Generando informe" & vbCrLf & Err.Description
        
        Err.Clear
    End If
    Label4.Visible = False
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Cad = "01/" & Month(Now) & "/" & Year(Now)
txtFecha(0).Text = Format(Cad, "dd/mm/yyyy")
txtFecha(1).Text = Format(Now, "dd/mm/yyyy")
Me.txtIncidencia(0).Text = ""
Me.txtIncidencia(1).Text = ""
Me.txtIncidencia(2).Text = ""
Me.txtIncidencia(3).Text = ""
Screen.MousePointer = vbDefault
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    txtIncidencia(Val(Me.ImgIncidencia(0).Tag)).Text = vCodigo
    txtIncidencia(Val(Me.ImgIncidencia(0).Tag) + 1).Text = vCadena
End Sub

Private Sub frmF_Selec(vFecha As Date)
Image2(0).Tag = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image2_Click(Index As Integer)
    Image2(0).Tag = ""
    Set frmF = New frmCal
    frmF.Fecha = Now
    If txtFecha(Index).Text <> "" Then
        frmF.Fecha = CDate(txtFecha(Index).Text)
    End If
    frmF.Show vbModal
    Set frmF = Nothing
    If Image2(0).Tag <> "" Then txtFecha(Index).Text = Image2(0).Tag
End Sub

Private Sub ImgIncidencia_Click(Index As Integer)
    Me.ImgIncidencia(0).Tag = (Index * 2)
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

Private Sub txtFecha_GotFocus(Index As Integer)
With txtFecha(Index)
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtFecha_LostFocus(Index As Integer)
With txtFecha(Index)
    .Text = Trim(.Text)
    If .Text = "" Then Exit Sub
    If Not EsFechaOK(txtFecha(Index)) Then
        MsgBox "Fecha incorrecta: " & .Text, vbExclamation
        .Text = ""
        .SetFocus
    End If

End With
End Sub


Private Function generadias() As Boolean
Dim RT As ADODB.Recordset
Dim D As Date
Dim DF As Date
Dim Sab As Boolean
Dim Dom As Boolean
Dim i As Integer
Dim Desc As String

On Error GoTo Egeneradias
generadias = False
If txtFecha(0).Text <> "" And txtFecha(1).Text <> "" Then
    If CDate(txtFecha(0).Text) > CDate(txtFecha(1).Text) Then
        MsgBox "Fecha inicio mayor que fecha fin"
        Exit Function
    End If
End If
If txtFecha(0).Text = "" Then
    txtFecha(0).Tag = "01/01/2002"
Else
    txtFecha(0).Tag = txtFecha(0).Text
End If

If txtFecha(1).Text = "" Then
    txtFecha(1).Tag = Format(Now, "dd/mm/yyyy")
Else
    txtFecha(1).Tag = txtFecha(1).Text
End If

'Ahora generamos la tabla
If Not GeneraTabla Then Exit Function

'Ahora vamos a poner las opciones
Set RT = New ADODB.Recordset
Cad = "SELECT Horarios.IdHorario, Horarios.NomHorario"
Cad = Cad & " FROM Horarios INNER JOIN Trabajadores ON Horarios.IdHorario = Trabajadores.IdHorario"
Cad = Cad & " Where Trabajadores.Control "
If Me.Check1.Value = 1 Then
    Cad = Cad & " = 2"
Else
    Cad = Cad & " > 0"
End If
Cad = Cad & " GROUP BY Horarios.IdHorario, Horarios.NomHorario"

RT.Open Cad, Conn, adOpenStatic, adLockOptimistic, adCmdText

If RT.EOF Then
    MsgBox "Ningun horario seleccionable", vbCritical
    RT.Close
    Exit Function
End If

'Insertaremos desde´- hasta la fecha
While Not RT.EOF
    Sab = FINDEfestivo(True, RT!IdHorario)
    Dom = FINDEfestivo(False, RT!IdHorario)
    D = CDate(txtFecha(0).Tag)
    DF = CDate(txtFecha(1).Tag)
    'Insertamos los dias
    While D <= DF
        
        i = Weekday(D, vbMonday)
        Desc = ""
        Select Case i
        Case 6
            If Sab Then Desc = "Sabado"
        Case 7
            If Dom Then Desc = "Domingo"
        Case Else
            'Comprobam
            Desc = EsFestivo(D, RT.Fields(0))
        End Select
        Conn.Execute "Insert into tmpFechas VALUES ('" & Format(D, "yyyy/mm/dd") & "','" & RT.Fields(1) & "','" & Desc & "'," & RT.Fields(0) & ");"
        D = D + 1
    Wend
    'Siguiente
    RT.MoveNext
Wend
RT.Close
Set RT = Nothing
generadias = True
Exit Function
Egeneradias:
    MuestraError Err.Number, "Genera dias " & vbCrLf & Err.Description
End Function


Private Function GeneraTabla() As Boolean
Dim RT As Recordset
On Error Resume Next

Set RT = New ADODB.Recordset
GeneraTabla = True
RT.Open "Select * from tmpFechas", Conn, adOpenKeyset, adLockOptimistic, adCmdText
If Err.Number <> 0 Then
    Err.Clear
    Cad = "CREATE TABLE tmpFechas (fechas DATE,Horario  STRING,Descr STRING, idHor INTEGER)"
    Conn.Execute Cad
    If Err.Number <> 0 Then
        MuestraError Err.Number, "EGenerando tabla temporal" & vbCrLf & "Error critico." & Conn.Errors(0).Description
        GeneraTabla = False
        Exit Function
    End If
    Else
        'Si k existia
        RT.Close
        Conn.Execute "Delete * from tmpFechas"
End If

'Segunda tabla temporal
RT.Open "Select * from tmpNoTrabajo", Conn, adOpenKeyset, adLockOptimistic, adCmdText
If Err.Number <> 0 Then
    Err.Clear
    Cad = "CREATE TABLE tmpNoTrabajo (idTra INTEGER, idFech DATE)"
    Conn.Execute Cad
    If Err.Number <> 0 Then
        MuestraError Err.Number, "EGenerando tabla temporal" & vbCrLf & "Error critico." & Conn.Errors(0).Description
        GeneraTabla = False
        Exit Function
    End If
    Else
        'Si k existia
        RT.Close
        Conn.Execute "Delete * from tmpNoTrabajo"
End If

Set RT = Nothing
End Function




Private Function FINDEfestivo(Sabado As Boolean, Horario As Integer) As Boolean
Dim R As Recordset
Set R = New ADODB.Recordset
Cad = "SELECT SubHorarios.Festivo From SubHorarios"
Cad = Cad & " WHERE (((SubHorarios.IdHorario)=" & Horario
Cad = Cad & " ) AND ((SubHorarios.DiaSemana)="
If Sabado Then
    Cad = Cad & "6"
Else
    Cad = Cad & "7"
End If
Cad = Cad & "))"
R.Open Cad, Conn, adOpenForwardOnly, adCmdText
FINDEfestivo = False
If Not R.EOF Then
    FINDEfestivo = R.Fields(0)
End If
R.Close
Set R = Nothing
End Function

Private Function EsFestivo(Fecha, IdHorario As Integer) As String
Dim Cad As String
Dim RF As ADODB.Recordset

'Devuelve una cadena que dice que fiesta es( y se añade el nombre de horario.")
EsFestivo = ""
Set RF = New ADODB.Recordset
Cad = "Select Descripcion from Festivos where IdHorario=" & IdHorario
Cad = Cad & " and Fecha=#" & Format(Fecha, "yyyy/mm/dd") & "#"
RF.Open Cad, Conn, , , adCmdText
If Not RF.EOF Then
    EsFestivo = RF.Fields(0)
End If
RF.Close
Set RF = Nothing
End Function



Private Function Generatemporal() As Boolean
Dim RT As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim Nexo As String

Set RS = New ADODB.Recordset
Generatemporal = False
RS.Open "Select DISTINCT (Fechas) from tmpFechas where descr=""""", Conn, adOpenKeyset, adLockOptimistic, adCmdText
If RS.EOF Then
    RS.Close
    MsgBox "Ningún dato en el temporal de fechas", vbExclamation
    Exit Function
End If


Cad = "SELECT Marcajes.idTrabajador, Marcajes.Fecha FROM Marcajes WHERE "
Nexo = ""
If Me.txtFecha(0).Text <> "" Then
    Cad = Cad & " Fecha >=#" & Format(txtFecha(0).Text, "yyyy/mm/dd") & "#"
    Nexo = " AND "
End If
If Me.txtFecha(1).Text <> "" Then
    Cad = Cad & Nexo & " Fecha <=#" & Format(txtFecha(1).Text, "yyyy/mm/dd") & "#"
End If
Cad = Cad & " ORDER By idTrabajador"
Set Trab = New ADODB.Recordset
Trab.Open Cad, Conn, adOpenStatic, adLockOptimistic, adCmdText



Cad = "SELECT Trabajadores.IdTrabajador, Trabajadores.Control From Trabajadores"
Cad = Cad & " WHERE (((Trabajadores.Control)=2))"
If Me.txtIncidencia(0).Text <> "" Then Cad = Cad & " AND Trabajadores.Seccion >=" & Me.txtIncidencia(0).Text
If Me.txtIncidencia(2).Text <> "" Then Cad = Cad & " AND Trabajadores.Seccion <=" & Me.txtIncidencia(2).Text

Set RT = New ADODB.Recordset
RT.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
While Not RT.EOF
    RS.MoveFirst
    While Not RS.EOF
        Label4.Caption = "Trab: " & RT!idTrabajador & " Fecha: " & RS!FECHAS
        Label4.Refresh
        If Not EncuentraDia(RT!idTrabajador, RS!FECHAS) Then
            'Insertamos
            Conn.Execute "INSERT INTO tmpNoTrabajo VALUES (" & RT!idTrabajador & ",'" & RS!FECHAS & "')"
        End If
        RS.MoveNext
    Wend
    RT.MoveNext
Wend
Trab.Close
RT.Close
RS.Close
Generatemporal = True

End Function


Private Function GeneratemporalDiasTr() As Boolean
Dim RT As ADODB.Recordset
Dim aux As String
Dim RS As Recordset
    Set RS = New ADODB.Recordset
    GeneratemporalDiasTr = False
    RS.Open "Select DISTINCT (Fechas) from tmpFechas ", Conn, adOpenKeyset, adLockOptimistic, adCmdText
    If RS.EOF Then
        RS.Close
        MsgBox "Ningún dato en el temporal de fechas", vbExclamation
        Exit Function
    End If
    'Cad = "SELECT Marcajes.IdTrabajador From Marcajes"
    Cad = "SELECT Marcajes.idTrabajador FROM Trabajadores"
    Cad = Cad & " INNER JOIN Marcajes ON Trabajadores.IdTrabajador ="
    Cad = Cad & " Marcajes.idTrabajador"
    
    Cad = Cad & " WHERE Fecha=#"
    While Not RS.EOF
        aux = Cad & Format(RS!FECHAS, "yyyy/mm/dd") & "#"
        If Check1.Value = 1 Then aux = aux & " AND Trabajadores.Control = 2"
        If Me.txtIncidencia(0).Text <> "" Then aux = aux & " AND Trabajadores.Seccion >=" & Me.txtIncidencia(0).Text
        If Me.txtIncidencia(2).Text <> "" Then aux = aux & " AND Trabajadores.Seccion <=" & Me.txtIncidencia(2).Text
        aux = aux & " GROUP BY Marcajes.idTrabajador;"
        Label4.Caption = " Fecha: " & RS!FECHAS
        Label4.Refresh
        Set RT = New ADODB.Recordset
        RT.Open aux, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        While Not RT.EOF
            'Insertamos
            Conn.Execute "INSERT INTO tmpNoTrabajo VALUES (" & RT!idTrabajador & ",'" & RS!FECHAS & "')"
            RT.MoveNext
        Wend
        RT.Close
        RS.MoveNext
    Wend
    RS.Close
    GeneratemporalDiasTr = True
End Function

Private Function EncuentraDia(T As Integer, D As Date) As Boolean
EncuentraDia = False
If Trab.RecordCount < 1 Then Exit Function
Trab.MoveFirst
While Not Trab.EOF
    If Trab!idTrabajador = T Then
        If Trab!Fecha = D Then
            EncuentraDia = True
            Exit Function
        End If
    End If
    Trab.MoveNext
Wend
End Function

Private Sub txtIncidencia_LostFocus(Index As Integer)
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
