VERSION 5.00
Begin VB.Form frmCambioEmpleado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio Empleado"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6420
   Icon            =   "frmCambioEmpleado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   6420
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text7 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   240
      TabIndex        =   22
      Top             =   2400
      Width           =   4395
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Index           =   2
      Left            =   4800
      TabIndex        =   20
      Top             =   1680
      Width           =   1035
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   240
      TabIndex        =   4
      Top             =   4320
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Index           =   1
      Left            =   5040
      TabIndex        =   18
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar"
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   17
      Top             =   5040
      Width           =   1215
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1920
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Tag             =   "Categoria|N|N|||Trabajadores|idCategoria|||"
      Top             =   4320
      Width           =   3795
   End
   Begin VB.TextBox Text4 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Top             =   3480
      Width           =   3435
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   2
      Top             =   3480
      Width           =   915
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   4800
      TabIndex        =   1
      Top             =   2400
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   4860
      TabIndex        =   3
      Top             =   3480
      Width           =   1095
   End
   Begin VB.TextBox txtTra 
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   1680
      Width           =   915
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Top             =   1680
      Width           =   3435
   End
   Begin VB.Label Label1 
      Caption         =   "Categoria"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   23
      Top             =   2160
      Width           =   840
   End
   Begin VB.Label Label1 
      Caption         =   "ALTA"
      Height          =   255
      Index           =   5
      Left            =   4800
      TabIndex        =   21
      Top             =   1440
      Width           =   555
   End
   Begin VB.Label Label1 
      Caption         =   "Cod asesoria"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   19
      Top             =   4080
      Width           =   1080
   End
   Begin VB.Label Label3 
      Caption         =   "A efectos de nomina tiene dos códigos distintos."
      Height          =   255
      Index           =   1
      Left            =   240
      TabIndex        =   16
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label3 
      Caption         =   "Da de baja  a un trabajador con una determinada categoria y lo da de alta en otra."
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   15
      Top             =   240
      Width           =   6735
   End
   Begin VB.Label Label5 
      Caption         =   "Categoría"
      Height          =   195
      Left            =   1920
      TabIndex        =   14
      Top             =   4080
      Width           =   705
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00800000&
      BorderWidth     =   3
      Index           =   1
      X1              =   1200
      X2              =   7800
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Label Label2 
      Caption         =   "ALTA"
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
      Left            =   240
      TabIndex        =   13
      Top             =   2760
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Nuevo codigo"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   12
      Top             =   3240
      Width           =   2880
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000080&
      BorderWidth     =   3
      Index           =   0
      X1              =   1200
      X2              =   7800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "BAJA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Index           =   0
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Width           =   780
   End
   Begin VB.Label Label1 
      Caption         =   "Baja"
      Height          =   255
      Index           =   0
      Left            =   4800
      TabIndex        =   9
      Top             =   2160
      Width           =   555
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   0
      Left            =   5580
      Picture         =   "frmCambioEmpleado.frx":030A
      Top             =   2160
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Alta"
      Height          =   195
      Index           =   1
      Left            =   4860
      TabIndex        =   8
      Top             =   3180
      Width           =   420
   End
   Begin VB.Image ImgFech 
      Height          =   240
      Index           =   1
      Left            =   5400
      Picture         =   "frmCambioEmpleado.frx":040C
      Top             =   3120
      Width           =   240
   End
   Begin VB.Image ImgTrab 
      Height          =   240
      Left            =   1080
      Picture         =   "frmCambioEmpleado.frx":050E
      Top             =   1440
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1440
      Width           =   840
   End
End
Attribute VB_Name = "frmCambioEmpleado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer


Private Sub Command1_Click(Index As Integer)
    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
    
    
    If UCase(Me.Command1(0).Caption) = "REPETIR" Then
        Limpiar Me
        Me.Combo2.ListIndex = -1
        Me.Command1(0).Caption = "Generar"
        Exit Sub
    End If
    
    
    Cad = "Campos no pueden estar vacios"
    If txtTra.Text <> "" Then
        If Text1(0).Text <> "" And Text1(1).Text <> "" Then
            If Text3.Text <> "" And Text4.Text <> "" Then Cad = ""
        End If
    End If
    
    
    If Cad <> "" Then
        MsgBox Cad, vbExclamation
        Exit Sub
    End If

    If CDate(Text1(0).Text) > CDate(Text1(1).Text) Then
        MsgBox "Fecha de baja mayor que la fecha del alta ", vbExclamation
        Exit Sub
    End If

    If DateDiff("d", CDate(Text1(0).Text), CDate(Text1(1).Text)) > 1 Then
        '------------------------------------------------------------------
        Cad = "La diferencia entre la fecha de baja y la de alta es mayor que un dia."
        Cad = Cad & "  ¿Desea continuar ?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    If Combo2.ListIndex < 0 Then
        MsgBox "Seleccione la nueva categoría", vbExclamation
        Exit Sub
    End If

    Set RS = New ADODB.Recordset
    'Ahora comprobamos varias cosas
    'Primero. Antes de la fecha de baja
    'Antencion ¿VALE la fecha de baja?  ###QUITAR###
    Cad = "Select count(*) from entradafichajes where idTrabajador = " & txtTra.Text
    Cad = Cad & " AND Fecha <=#" & Format(Text1(0).Text, FormatoFecha) & "#"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = DBLet(RS.Fields(0), "N")
    RS.Close
    
    If I > 0 Then
        Cad = "Existen entradas pendientes de procesar antes de la fecha de baja para este trabajador."
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    'Ahora comprobamos que no existen marcajes para el trabajador con la fecha superior a fecha baja
    Cad = "Select count(*) from marcajes where idTrabajador = " & txtTra.Text
    Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    I = DBLet(RS.Fields(0), "N")
    RS.Close

    If I > 0 Then
        Cad = "Existe marcajes marcajes con fecha posterior a la fecha de baja. ¿Desea continuar?"
        If MsgBox(Cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
    End If
    
    Cad = "El proceso es irreversible. ¿Desea continuar?"
    If MsgBox(Cad, vbCritical + vbYesNoCancel) <> vbYes Then Exit Sub
    

    Screen.MousePointer = vbHourglass
    Conn.BeginTrans
    If GenerarNuevotrabajador Then
        Conn.CommitTrans
        MsgBox "Proceso finalizado", vbInformation
        Me.Command1(0).Caption = "Repetir"
    Else
        Conn.RollbackTrans
    End If
    Set RS = Nothing
    Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Load()
'Categorias
    Combo2.Clear
    Cad = "Select IdCategoria,NomCategoria from Categorias order by nomCategoria"
    Set RS = New ADODB.Recordset
    RS.Open Cad, Conn, , , adCmdText
    I = 0
    While Not RS.EOF
        Combo2.AddItem RS.Fields(1) '& " - " & rs.Fields(0)
        Combo2.ItemData(I) = RS.Fields(0)
        I = I + 1
        RS.MoveNext
    Wend
    RS.Close
End Sub
Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    VariableCompartida = vCodigo
End Sub

Private Sub frmC_Selec(vFecha As Date)
    Text1(CInt(Text1(0).Tag)).Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub ImgFech_Click(Index As Integer)
        Set frmC = New frmCal
        Text1(0).Tag = Index
        frmC.Fecha = CDate(Text1(Index).Text)
        frmC.Show vbModal
        Set frmC = Nothing
End Sub

Private Sub ImgTrab_Click()
    VariableCompartida = ""
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "NomTrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPLEADOS"
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
    If VariableCompartida <> "" Then
        txtTra.Text = VariableCompartida
        txtTra_LostFocus
    End If
        
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text1_LostFocus(Index As Integer)

    With Text1(Index)
        .Text = Trim(.Text)
        If .Text <> "" Then
            If Not EsFechaOK(Text1(Index)) Then
                .Text = ""
            End If
        End If
        
        If .Text = "" Then
            If Index = 0 Then
                .Text = "01/" & Format(DateAdd("m", -1, Now), "mm/yyyy")
            Else
                .Text = Format(Now, "dd/mm/yyyy")
            End If
        End If
        If Index = 0 Then
            If Text1(1).Text = "" Then Text1(1).Text = Format(DateAdd("d", 1, CDate(Text1(0).Text)), "dd/mm/yyyy")
        End If
     End With
End Sub




Private Sub Text3_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub Text3_LostFocus()
    If Text3.Text <> "" Then
        If Not IsNumeric(Text3.Text) Then
            MsgBox "Campo debe ser numérico", vbExclamation
            Text3.Text = ""
        Else
            Cad = DevuelveDesdeBD("nomtrabajador", "trabajadores", "idTrabajador", Text3.Text, "N")
            If Cad <> "" Then
                Cad = "El codigo: " & Text3.Text & " pertenece a " & Cad
                MsgBox Cad, vbExclamation
                Text3.Text = ""
            End If
        End If
    End If
End Sub



Private Sub Text5_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtTra_GotFocus()
    With txtTra
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub



Private Sub txtTra_KeyPress(KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtTra_LostFocus()
    Screen.MousePointer = vbHourglass
    I = 0
    If txtTra.Text <> "" Then
        If Not IsNumeric(txtTra.Text) Then
            
            MsgBox "Debe ser un número", vbExclamation
        Else
           If PonTrabajador Then I = 1
        End If
    End If
    If I = 0 Then
        Text1(2).Text = ""
        Text2.Text = ""
        Text7.Text = ""
    End If
    Text4.Text = Text2.Text
    Screen.MousePointer = vbDefault
End Sub

Private Sub Keypress(ByRef KeyAscii As Integer)
        If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub


Private Function GenerarNuevotrabajador() As Boolean
Dim RT As ADODB.Recordset

    On Error GoTo EGenerarNuevotrabajador
    GenerarNuevotrabajador = False
    'Pasos:
    '
    '   1.- Crear nuevo trabajador.
    '   2.- Dar de baja( fecha y tarjeta) al trabajador
    '   3.- En entradafichajes updatear al nuevo trabajador a partir de la fecha
    '   4.- En entradamarcajes updatear al nuevo trabajador a partir de la fecha
    
   
    'Crear nuevo trabajador
    Cad = "Select * from Trabajadores where idTrabajador = " & txtTra.Text
    RS.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    'El recordset de insertar
    Set RT = New ADODB.Recordset
    RT.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    RT.AddNew
    For I = 0 To RS.Fields.Count - 1
        RT.Fields(I) = RS.Fields(I)
    Next I

    
    
    'Baja
    RS!fecbaja = CDate(Text1(0).Text)
    RS!Numtarjeta = Null
    RS.Update
    
    
    
    'Ponemos los nuevos valores
    RT!idTrabajador = Text3.Text
    RT!fecbaja = Null
    RT!fecalta = CDate(Text1(1).Text)
    RT!idCategoria = Combo2.ItemData(Combo2.ListIndex)
    RT!idasesoria = Text5.Text
    RT!bolsahoras = 0
    RT.Update
    
    
    RT.Close
    RS.Close
    
    
    
    
    'Actualizamos las entradafichajes
    Cad = "UPDATE EntradaFichajes SET idTrabajador=" & Text3.Text
    Cad = Cad & " WHERE idTrabajador =" & txtTra.Text
    Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
    Conn.Execute Cad
    
    
    'Actualizamos las entradas marcajes
    Cad = "UPDATE EntradaMarcajes SET idTrabajador=" & Text3.Text
    Cad = Cad & " WHERE idTrabajador =" & txtTra.Text
    Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
    Conn.Execute Cad
    
    
    'Actualizamos los marcajes
    Cad = "UPDATE Marcajes SET idTrabajador=" & Text3.Text
    Cad = Cad & " WHERE idTrabajador =" & txtTra.Text
    Cad = Cad & " AND Fecha >#" & Format(Text1(0).Text, FormatoFecha) & "#"
    Conn.Execute Cad
    
    
    
    GenerarNuevotrabajador = True
    Exit Function
EGenerarNuevotrabajador:
    MuestraError Err.Number, Err.Description
End Function




Private Function PonTrabajador() As Boolean
    PonTrabajador = False
    Set RS = New ADODB.Recordset
    Cad = "SELECT Trabajadores.NomTrabajador, Trabajadores.FecAlta, Categorias.nomCategoria"
    Cad = Cad & " FROM Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria"
    Cad = Cad & " WHERE Trabajadores.IdTrabajador= " & txtTra.Text
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
       MsgBox "No exsite el trabajador: " & Me.txtTra.Text
    
    Else
        PonTrabajador = True
        Text1(2).Text = Format(RS!fecalta, "dd/mm/yyyy")
        Text2.Text = RS!nomtrabajador
        Text7.Text = RS!nomCategoria
    End If
End Function
