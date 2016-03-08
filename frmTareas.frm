VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTareas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tareas"
   ClientHeight    =   2025
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7260
   Icon            =   "frmTareas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2025
   ScaleWidth      =   7260
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   3
      Left            =   6540
      TabIndex        =   4
      Tag             =   "Tipo|N|N|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   435
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   2
      Left            =   4920
      TabIndex        =   3
      Tag             =   "EAN|T|N|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   1395
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   60
      TabIndex        =   10
      Top             =   1440
      Width           =   3135
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   60
         TabIndex        =   11
         Top             =   180
         Width           =   2955
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4710
      TabIndex        =   6
      Top             =   1560
      Width           =   1035
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1560
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   2400
      Top             =   300
      Visible         =   0   'False
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3420
      TabIndex        =   5
      Top             =   1560
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   1620
      TabIndex        =   2
      Tag             =   "Nombre|T|N|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Tag             =   "Código|N|S|||"
      Text            =   "Text1"
      Top             =   960
      Width           =   1155
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   20
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver todos"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Primero"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Anterior"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Siguiente"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Último"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Listado"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":1A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":22F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTareas.frx":2BCC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Tipo"
      Height          =   195
      Index           =   3
      Left            =   6540
      TabIndex        =   13
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "EAN"
      Height          =   195
      Index           =   2
      Left            =   4980
      TabIndex        =   12
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Descripción"
      Height          =   195
      Index           =   1
      Left            =   1620
      TabIndex        =   7
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   1155
   End
End
Attribute VB_Name = "frmTareas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'-----------------------------
'Se distinguen varios modos
'   0.-  Formulario limpio sin nungun campo rellenado
'   1.-  Preparando para hacer la busqueda
'   2.-  Ya tenemos registros y los vamos a recorrer
        'y podemos editarlos Edicion del campo
'   3.-  Insercion de nuevo registro
'   4.-  Modificar

Private Modo As Byte
Private CadenaConsulta As String
Private Ordenacion As String
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private kCampo As Integer





Private Sub cmdAceptar_Click()
Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer

Screen.MousePointer = vbHourglass
On Error GoTo Error1
If Modo = 3 Then
    If DatosOk Then
        
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenKeyset
        RS.LockType = adLockOptimistic
        RS.Open NombreTabla, Conn, , , adCmdTable
        RS.AddNew
        'Para luego
        For I = 0 To Text1.Count - 1
            RS.Fields(I) = Text1(I).Text
        Next I

        '--------------------
        RS.Update
        RS.Close
        Data1.Refresh
        'MsgBox "Registro insertado.", vbInformation
        PonerModo 0
        Label2.Caption = "Insertado"
    End If
    Else
    If Modo = 4 Then
        'Modificar
        ''Haremos las comprobaciones necesarias de los campos
        For I = 1 To Text1.Count - 1
            If Not CmpCam(Text1(I).Tag, Text1(I).Text) Then Exit Sub
        Next I
        'Ahora modificamos
        Cad = "Select * from " & NombreTabla
        Cad = Cad & " WHERE idtarea=" & Data1.Recordset.Fields(0)
        Set RS = New ADODB.Recordset
        RS.CursorType = adOpenKeyset
        RS.LockType = adLockOptimistic
        RS.Open Cad, Conn, , , adCmdText
        'Almacenamos para luego buscarlo
        Cad = RS!Idtarea
        'Modificamos
        For I = 1 To Text1.Count - 1
            RS.Fields(I) = Text1(I).Text
        Next I

        RS.Update
        RS.Close
        'MsgBox "El registro ha sido modificado", vbInformation
        PonerModo 2
        'Hay que refresca el DAta1
        Data1.Refresh
        'Hay que volver a poner el registro donde toca
        Data1.Recordset.MoveFirst
        I = 1
        While I > 0
            If Data1.Recordset.Fields(0) = Cad Then
                I = 0
                Else
                    Data1.Recordset.MoveNext
                    If Data1.Recordset.EOF Then I = 0
            End If
        Wend
        If Data1.Recordset.EOF Then
            Data1.Recordset.MoveFirst
        End If
        Label2.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
    End If
End If
Error1:
    Screen.MousePointer = vbDefault
    If Err.Number <> 0 Then MsgBox Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdCancelar_Click()
LimpiarCampos
PonerModo 0
End Sub

Private Sub BotonAnyadir()
LimpiarCampos
'Añadiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Aceptar"
PonerModo 3
'Escondemos el navegador y ponemos insertando
DespalzamientoVisible False
Label2.Caption = "INSERTANDO"
SugerirCodigoSiguiente

Text1(0).SetFocus
End Sub

Private Sub BotonBuscar()
'Buscar
If Modo <> 1 Then
    LimpiarCampos
    Label2.Caption = "Búsqueda"
    PonerModo 1
    Text1(0).SetFocus
    Else
        HacerBusqueda
        If Data1.Recordset.EOF Then
            Text1(kCampo).Text = ""
            Text1(kCampo).BackColor = vbYellow
            Text1(kCampo).SetFocus
        End If
End If
End Sub

Private Sub BotonVerTodos()
'Ver todos
LimpiarCampos
PonerModo 2
CadenaConsulta = "Select * from " & NombreTabla
PonerCadenaBusqueda
End Sub

Private Sub Desplazamiento(Index As Integer)

Select Case Index
    Case 0
        Data1.Recordset.MoveFirst
    Case 1
        Data1.Recordset.MovePrevious
        If Data1.Recordset.BOF Then Data1.Recordset.MoveFirst
    Case 2
        Data1.Recordset.MoveNext
        If Data1.Recordset.EOF Then Data1.Recordset.MoveLast
    Case 3
        Data1.Recordset.MoveLast
End Select
PonerCampos
End Sub

Private Sub BotonModificar()
'---------
'MODIFICAR
'----------
'Añadiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Modificar"
PonerModo 4
'Escondemos el navegador y ponemos insertando
'Como el campo 1 es clave primaria, NO se puede modificar
Text1(0).Locked = True
DespalzamientoVisible False
Label2.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
Dim Cad As String
Dim I As Long

'Ciertas comprobaciones
If Data1.Recordset.RecordCount = 0 Then Exit Sub
'Pregunta
Cad = "Seguro que desea eliminar de la BD el registro:"
Cad = Cad & vbCrLf & "Cod: " & Data1.Recordset.Fields(0)
Cad = Cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
I = MsgBox(Cad, vbQuestion + vbYesNo)
If I = vbYes Then
    'Hay que eliminar
    On Error GoTo Error2
    Screen.MousePointer = vbHourglass
    Data1.Recordset.Delete
    'Esperamos un tiempo prudencial de 1 seg
    I = CLng(Timer)
    Do
    Loop Until Timer - I > 1
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        PonerModo 0
        Else
            Data1.Recordset.MoveFirst
            PonerCampos
    End If
End If
Error2:
    Screen.MousePointer = vbDefault
    If Err.Number > 0 Then MsgBox Err.Number & " - " & Err.Description
End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
LimpiarCampos
NombreTabla = "Tareas"
Ordenacion = " ORDER BY idTarea"
'Situamos el form
Left = 0
Top = BaseForm
'ASignamos un SQL al DATA1
Data1.ConnectionString = Conn
Data1.RecordSource = "Select * from " & NombreTabla
Data1.Refresh
PonerModo 0
End Sub



Private Sub LimpiarCampos()
Dim I
For I = 0 To Text1.Count - 1
    Text1(I).Text = ""
Next I
End Sub


Private Sub mnBuscar_Click()
BotonBuscar
End Sub

Private Sub mnEliminar_Click()
BotonEliminar
End Sub

Private Sub mnModificar_Click()
BotonModificar
End Sub

Private Sub mnNuevo_Click()
BotonAnyadir
End Sub

Private Sub mnSalir_Click()
Screen.MousePointer = vbHourglass
Unload Me
End Sub

Private Sub mnVerTodos_Click()
BotonVerTodos
End Sub

Private Sub Text1_GotFocus(Index As Integer)
kCampo = Index
If Modo = 1 Then
    Text1(Index).BackColor = vbYellow
Else
    With Text1(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End If
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If Modo = 1 Then
    If KeyAscii = 13 Then
        'Ha pulsado enter, luego tenemos que hacer la busqueda
        Text1(Index).BackColor = vbWhite
        BotonBuscar
    End If
Else
    
End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
Dim Cad As String
If Modo = 1 Then
    Text1(Index).BackColor = vbWhite
    
Else
    If Modo > 2 Then
        Select Case Index
        Case 3
            'Cad = TextBoxAImporte(Text1(3))
            If Cad <> "" Then
                MsgBox "El campo debe ser numérico", vbExclamation
                Text1(3).Text = ""
            End If
        Case 0
            If Not IsNumeric(Text1(0).Text) Then
                MsgBox "Campo debe ser numerico", vbExclamation
                Text1(0).Text = ""
            End If
        End Select
    End If
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String
Dim c1 As String   'el nombre del campo
Dim Tipo As Long
Dim aux1

If Text1(kCampo).Text = "" Then Exit Sub
c1 = Data1.Recordset.Fields(kCampo).Name
c1 = " WHERE " & c1
Tipo = DevuelveTipo2(Data1.Recordset.Fields(kCampo).Type)
'Devolvera uno de los tipos
'   1.- Numeros
'   2.- Booleanos
'   3.- Cadenas
'   4.- Fecha
'   0.- Error leyendo los tipos de datos
' segun sea uno u otro haremos una comparacion
Select Case Tipo
Case 1
    CadB = c1 & " = " & Text1(kCampo)
Case 2
    'Vemos si la cadena tiene un Falso o False
    If InStr(1, UCase(Text1(kCampo).Text), "F") Then
        aux1 = "False"
        Else
        aux1 = "True"
    End If
    CadB = c1 & " = " & aux1
Case 3
    CadB = c1 & " like '*" & Trim(Text1(kCampo)) & "*'"
Case 4

Case 5

End Select

CadenaConsulta = "select * from " & NombreTabla & CadB & " " & Ordenacion
PonerCadenaBusqueda
End Sub


Private Sub PonerCadenaBusqueda()
Screen.MousePointer = vbHourglass

Data1.RecordSource = CadenaConsulta
Data1.Refresh
If Data1.Recordset.RecordCount <= 0 Then
    MsgBox "No hay ningún registro en la tabla" & NombreTabla, vbInformation
    Screen.MousePointer = vbDefault
    Exit Sub
    'StatusBar1.Panels(2).Text = ""
    'PonerModo 0
    Else
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        PonerCampos
End If

Data1.ConnectionString = Conn
Data1.RecordSource = CadenaConsulta
Data1.Refresh

Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim I As Integer
    If Data1.Recordset.EOF Then Exit Sub

    For I = 0 To Text1.Count - 1
        Text1(I).Text = Data1.Recordset.Fields(I)
    Next I
    'Los dos check
    Text1(3).Text = Data1.Recordset.Fields(3)
    Label2.Caption = Data1.Recordset.AbsolutePosition & " de " & Data1.Recordset.RecordCount
  
End Sub

Private Sub PonerModo(Kmodo As Integer)
Dim I As Integer
Dim b As Boolean

If Modo = 1 Then
    For I = 0 To Text1.Count - 1
        Text1(I).BackColor = vbWhite
    Next I
End If
Modo = Kmodo
DespalzamientoVisible (Kmodo = 2)
cmdAceptar.Visible = (Kmodo >= 3)
cmdCancelar.Visible = (Kmodo >= 3)
Toolbar1.Buttons(6).Enabled = (Kmodo < 3)
Toolbar1.Buttons(7).Enabled = (Kmodo = 2)
Toolbar1.Buttons(8).Enabled = (Kmodo = 2)
Toolbar1.Buttons(1).Enabled = (Kmodo < 3)
Toolbar1.Buttons(2).Enabled = (Kmodo < 3)
If Kmodo = 0 Then _
    Label2.Caption = ""
b = (Modo = 2) Or Modo = 0
For I = 0 To Text1.Count - 1
    Text1(I).Locked = b
Next I
End Sub


Private Function DatosOk() As Boolean
Dim RS As ADODB.Recordset
Dim Cad As String
Dim I As Integer

DatosOk = False
'Haremos las comprobaciones necesarias de los campos
'Cad = ComprobarCampos
'If Cad <> "" Then
'    MsgBox Cad, vbExclamation
'    Exit Function
'End If


For I = 0 To Text1.Count - 1
    If Not CmpCam(Text1(I).Tag, Text1(I).Text) Then Exit Function
Next I
'Llegados a este punto los datos son correctos en valores
'Ahora comprobaremos otras cosas
'Este apartado dependera del formulario y la tabla
Cad = "Select * from " & NombreTabla
Cad = Cad & " WHERE idtarea=" & Text1(0).Text

Set RS = New ADODB.Recordset
RS.Open Cad, Conn, , , adCmdText
If Not RS.EOF Then
    MsgBox "Ya existe un registro con ese código.", vbExclamation
    RS.Close
    Exit Function
End If
RS.Close
'Al final todo esta correcto
DatosOk = True
End Function


Private Sub SugerirCodigoSiguiente()
Dim Cad
Dim RS
'Sugeriremos el codigo siguiente.
'Obviamente depende en TOTAL medida de que tabla estemos trabajando
Cad = "Select Max(idTarea) from " & NombreTabla
Text1(0).Text = 1
Set RS = New ADODB.Recordset
RS.Open Cad, Conn, , , adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then
        Text1(0).Text = RS.Fields(0) + 1
    End If
End If
RS.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
    BotonBuscar
Case 2
    BotonVerTodos
Case 6
    BotonAnyadir
Case 7
    BotonModificar
Case 8
    BotonEliminar
Case 14 To 17
    Desplazamiento (Button.Index - 14)
Case 20
    'Listado en crystal report
    Screen.MousePointer = vbHourglass

    Screen.MousePointer = vbDefault

Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
Dim I
For I = 14 To 17
    Toolbar1.Buttons(I).Visible = bol
Next I
End Sub

