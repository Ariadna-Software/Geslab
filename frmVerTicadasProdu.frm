VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVerTicadasProdu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ver ticajes / tareas"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "frmVerTicadasProdu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4590
   ScaleWidth      =   7035
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   240
      TabIndex        =   13
      Text            =   "Text3"
      Top             =   420
      Width           =   615
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   120
      Top             =   4140
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerTicadasProdu.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVerTicadasProdu.frx":0624
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdFec 
      Caption         =   ">"
      Height          =   255
      Index           =   1
      Left            =   6540
      TabIndex        =   12
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdFec 
      Caption         =   "<"
      Height          =   255
      Index           =   0
      Left            =   6180
      TabIndex        =   11
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdTrab 
      Caption         =   ">"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   10
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton cmdTrab 
      Caption         =   "<"
      Height          =   255
      Index           =   0
      Left            =   1680
      TabIndex        =   9
      Top             =   120
      Width           =   315
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   5820
      TabIndex        =   4
      Top             =   4080
      Width           =   1095
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   180
      TabIndex        =   2
      Top             =   1080
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Hora"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   5280
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   420
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   315
      Left            =   900
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   420
      Width           =   4155
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   2955
      Left            =   2220
      TabIndex        =   3
      Top             =   1080
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tarea"
         Object.Width           =   4445
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Inicio"
         Object.Width           =   1482
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Horas"
         Object.Width           =   1640
      EndProperty
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   3120
      Picture         =   "frmVerTicadasProdu.frx":093E
      ToolTipText     =   "Cambiar tarea"
      Top             =   840
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   5880
      Picture         =   "frmVerTicadasProdu.frx":1340
      Top             =   120
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   1320
      Picture         =   "frmVerTicadasProdu.frx":1442
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Tareas"
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
      Height          =   195
      Left            =   2280
      TabIndex        =   8
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Ticajes maquina"
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
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5280
      TabIndex        =   6
      Top             =   150
      Width           =   570
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   150
      Width           =   975
   End
End
Attribute VB_Name = "frmVerTicadasProdu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1


Dim Cad As String
Dim RS As ADODB.Recordset
Dim Trab As Long
Dim Fec As Date

Dim itmX As ListItem

Private Sub cmdFec_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    If Text1.Text <> "" And Text2.Text <> "" Then OtraFecha cmdTrab(Index).Caption
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdTrab_Click(Index As Integer)
    Screen.MousePointer = vbHourglass
    If Text1.Text <> "" And Text2.Text <> "" Then OtroTrabajador cmdTrab(Index).Caption
    Screen.MousePointer = vbDefault
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub


Private Sub Command4_Click()
    OtraFecha "<"
End Sub

Private Sub Command5_Click()
    OtraFecha ">"
End Sub

Private Sub Form_Load()
    ListView1.SmallIcons = Me.ImageList1
    ListView2.SmallIcons = Me.ImageList1
    Trab = -1
    Fec = "0:00:00"
    Limpiar
    Ofertar
    Set RS = New ADODB.Recordset
End Sub

Private Function OtroTrabajador(Signo As String)
'    Cad = "SELECT MarcajesKimaldi.Fecha, MarcajesKimaldi.Hora, Trabajadores.IdTrabajador"
    Cad = "SELECT  Trabajadores.IdTrabajador,Trabajadores.nomtrabajador"
    Cad = Cad & " FROM Trabajadores INNER JOIN MarcajesKimaldi ON Trabajadores.NumTarjeta = MarcajesKimaldi.Marcaje"
    Cad = Cad & " WHERE Fecha = #" & Format(Fec, "yyyy/mm/dd") & "#"
    Cad = Cad & " AND idTrabajador" & Signo & Trab
    Cad = Cad & " ORDER BY Trabajadores.IdTrabajador "
    If Signo = ">" Then
        Cad = Cad & "ASC"
    Else
        Cad = Cad & "DESC"
    End If
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cad = ""
    If Not RS.EOF Then
        Text1.Text = RS.Fields(1) & " (" & RS.Fields(0) & ")"
        Text1.Tag = RS.Fields(0)
        Text3.Text = RS.Fields(0)
        Cad = RS.Fields(0)
    End If
    RS.Close

    If Cad <> "" Then
        'Trab = Val(Cad)
        CambioTrabajador
    End If
    
End Function



Private Function OtraFecha(Signo As String)
'    Cad = "SELECT MarcajesKimaldi.Fecha, MarcajesKimaldi.Hora, Trabajadores.IdTrabajador"
    Cad = "SELECT  MarcajesKimaldi.Fecha"
    Cad = Cad & " FROM Trabajadores INNER JOIN MarcajesKimaldi ON Trabajadores.NumTarjeta = MarcajesKimaldi.Marcaje"
    Cad = Cad & " WHERE Fecha " & Signo & " #" & Format(Fec, "yyyy/mm/dd") & "#"
    Cad = Cad & " AND Trabajadores.idTrabajador = " & Trab
    Cad = Cad & " ORDER BY MarcajesKimaldi.Fecha "
    If Signo = ">" Then
        Cad = Cad & "ASC"
    Else
        Cad = Cad & "DESC"
    End If

    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Cad = ""
    If Not RS.EOF Then
        Cad = RS.Fields(0)
    End If
    RS.Close

    If Cad <> "" Then
        
        Text2.Text = Format(Cad, "dd/mm/yyyy")
        'Fec = CDate(Text2.Text)
        CambioFecha
    End If
    
End Function



Private Sub Ofertar()
    'De momento no oferto na de na
End Sub



Private Sub Limpiar()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    ListView1.ListItems.Clear
    ListView2.ListItems.Clear
End Sub

Private Sub CargarDatos()
   On Error GoTo ECargarDatos
   Screen.MousePointer = vbHourglass
   

   ListView1.ListItems.Clear
   ListView2.ListItems.Clear
   CargaTicajes
   Me.Refresh
   CargaTareas
   
ECargarDatos:
    If Err.Number <> 0 Then _
        MuestraError Err.Number, Err.Description

    Screen.MousePointer = vbDefault
End Sub


Private Sub CargaTicajes()
Dim Tarj As String
    Cad = "Select numTarjeta from Trabajadores where idTrabajador =" & Trab
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Tarj = "2"
    If Not RS.EOF Then
        Tarj = DBLet(RS.Fields(0))
    End If
    RS.Close
    If Tarj = "" Then
        MsgBox "No encontrada tarjeta. Fallo grave.", vbCritical
        Exit Sub
    End If
    
    Cad = "Select hora from marcajeskimaldi where Marcaje = '" & Tarj & "'"
    Cad = Cad & " AND Fecha = #" & Format(Fec, "yyyy/mm/dd") & "# ORDER BY Hora"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    While Not RS.EOF
        Set itmX = ListView1.ListItems.Add(, , Format(RS.Fields(0), "hh:mm"))
        itmX.SmallIcon = 1
        'INSERTAMOS  items
        RS.MoveNext
    Wend
    RS.Close
    
End Sub



Private Sub CargaTareas()
    Cad = "SELECT Tareas.Descripcion,Tareas.idTarea , TareasRealizadas.HoraInicio, TareasRealizadas.HorasTrabajadas"
    Cad = Cad & " FROM TareasRealizadas INNER JOIN Tareas ON TareasRealizadas.Tarea = Tareas.idTarea"
    Cad = Cad & " WHERE (((TareasRealizadas.Trabajador)=" & Trab
    Cad = Cad & " ) AND ((Tareas.Tipo =0 )) AND " 'Quitamos salida
    Cad = Cad & "((TareasRealizadas.Fecha)=#" & Format(Fec, "yyyy/mm/dd")
    Cad = Cad & "#)) ORDER By HoraInicio;"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Set itmX = ListView2.ListItems.Add(, , RS.Fields(0))
        itmX.SubItems(1) = Format(RS.Fields(2), "hh:mm")
        itmX.Tag = "#" & Format(RS.Fields(2), "hh:mm:ss") & "#"
        itmX.SubItems(2) = Format(RS.Fields(3), "#,##0.00")
        itmX.SmallIcon = 2
        RS.MoveNext
    Wend
    RS.Close
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RS = Nothing
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    If Image1.Tag = 0 Then
        Text1.Text = vCadena & " (" & vCodigo & ")"
        Text1.Tag = vCodigo
        Text3.Text = vCodigo
        
    Else
        Cad = vCodigo & "|" & vCadena & "|"
    End If
End Sub

Private Sub frmF_Selec(vFecha As Date)
    Text2.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
    Image1.Tag = 0
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "Nomtrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "TRABAJADORES"
    frmB.Show vbModal
    Set frmB = Nothing
    Me.Refresh
    CambioTrabajador
End Sub

Private Sub Image2_Click()
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Text2.Text <> "" Then
        If IsDate(Text2.Text) Then frmF.Fecha = CDate(Text2.Text)
    End If
    frmF.Show vbModal
    Set frmF = Nothing
    CambioFecha
End Sub

Private Sub CambioTrabajador()
    If Text1.Tag <> "" Then
        If Val(Text1.Tag) <> Trab Then
            Trab = Text1.Tag
            'Si hay fecha puesta
            If Text2.Text <> "" Then CargarDatos
        End If
    End If
End Sub


Private Sub CambioFecha()
    If Text2.Text <> "" Then
        If CDate(Text2.Text) <> Fec Then
            Fec = Text2.Text
            CargarDatos
        End If
    End If
End Sub

Private Sub Image3_Click()

    If ListView2.ListItems.Count = 0 Then Exit Sub
    If ListView2.SelectedItem Is Nothing Then
        MsgBox "Seleccione la tarea a modificar", vbExclamation
        Exit Sub
    End If
    
     
    Cad = ""

    Image1.Tag = 1
    Set frmB = New frmBusca
    
    
    frmB.Tabla = "Tareas"
    frmB.CampoBusqueda = "Descripcion"
    frmB.CampoCodigo = "IdTarea"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "TAREAS"
    frmB.Show vbModal
    
    Set frmB = Nothing
    
    If Cad <> "" Then
        
        VariableCompartida = Text1.Text & " - " & Text2.Text & vbCrLf
        VariableCompartida = VariableCompartida & "TAREA: " & ListView2.SelectedItem.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf
        VariableCompartida = VariableCompartida & vbCrLf & "¿Desea cambiar la tarea por : " & RecuperaValor(Cad, 1) & " - " & RecuperaValor(Cad, 2) & "?"
        If MsgBox(VariableCompartida, vbQuestion + vbYesNo) = vbYes Then
            'UPDATEO
            
            VariableCompartida = "UPDATE TareasRealizadas SET tarea =" & RecuperaValor(Cad, 1)
            VariableCompartida = VariableCompartida & " WHERE fecha = #" & Format(Text2.Text, "yyyy/mm/dd")
            VariableCompartida = VariableCompartida & "# and Horainicio = " & ListView2.SelectedItem.Tag
            VariableCompartida = VariableCompartida & " AND trabajador =" & Text3.Text
            Conn.Execute VariableCompartida
            ListView2.ListItems.Clear
            CargaTareas
        End If
        
    End If
    
End Sub

Private Sub Text2_LostFocus()
    Text2.Text = Trim(Text2.Text)
    Fec = "01/01/1900"
    If Text2.Text <> "" Then
        If Not EsFechaOK(Text2) Then
            Text2.Text = ""
        Else
            Fec = CDate(Text2.Text)
        End If
    End If
    CargarDatos
End Sub

Private Sub Text3_LostFocus()

    If Text3.Text = "" Then
        Text1.Text = ""
        Text1.Tag = ""
        Trab = -1
        CargarDatos
        Exit Sub
    End If
    
    If Text3.Text = Text1.Tag Then Exit Sub
    Cad = DevuelveDesdeBD("nomtrabajador", "Trabajadores", "idtrabajador", Text3.Text, "N")
    If Cad = "" Then
        'NO EXISTE EL TRABAJADDOR
        Text3.Text = Text1.Tag
        
    Else
        'Si k existe
        Text1.Text = Cad
        Text1.Tag = Text3.Text
        
    End If
    CambioTrabajador
End Sub
