VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBancos2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bancos"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   Icon            =   "frmBancos2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   7
      Left            =   8400
      MaxLength       =   3
      TabIndex        =   7
      Tag             =   "Sufijo|T|S|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   4680
      MaxLength       =   4
      TabIndex        =   2
      Tag             =   "iban|T|S|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   675
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   7080
      MaxLength       =   10
      TabIndex        =   6
      Tag             =   "Cuenta banco|T|N|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   4
      Left            =   6720
      MaxLength       =   2
      TabIndex        =   5
      Tag             =   "DC|T|N|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   3
      Left            =   6120
      MaxLength       =   4
      TabIndex        =   4
      Tag             =   "Importe 2|T|N|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   555
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   5460
      MaxLength       =   4
      TabIndex        =   3
      Tag             =   "Entidad|T|N|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   555
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   14
      Top             =   1440
      Width           =   3375
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Label2"
         Height          =   255
         Left            =   420
         TabIndex        =   15
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6480
      TabIndex        =   9
      Top             =   1620
      Width           =   1035
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7680
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   1620
      Width           =   1035
   End
   Begin MSAdodcLib.Adodc Data1 
      Height          =   330
      Left            =   4740
      Top             =   120
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
      Left            =   5280
      TabIndex        =   8
      Top             =   1620
      Width           =   1035
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   1
      Left            =   900
      MaxLength       =   40
      TabIndex        =   1
      Tag             =   "Descripcion|T|N|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   3375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Tag             =   "Id|N|S|||"
      Text            =   "Text1"
      Top             =   840
      Width           =   675
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   17
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
            Object.ToolTipText     =   "�ltimo"
            ImageIndex      =   9
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
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":113E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":1A18
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmBancos2.frx":22F2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Sufijo"
      Height          =   195
      Index           =   7
      Left            =   8400
      TabIndex        =   21
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "IBAN"
      Height          =   195
      Index           =   6
      Left            =   4680
      TabIndex        =   20
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Cuenta"
      Height          =   195
      Index           =   5
      Left            =   7080
      TabIndex        =   19
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "D.C."
      Height          =   195
      Index           =   4
      Left            =   6720
      TabIndex        =   18
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Oficina"
      Height          =   195
      Index           =   3
      Left            =   6120
      TabIndex        =   17
      Top             =   600
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "Entidad"
      Height          =   195
      Index           =   2
      Left            =   5460
      TabIndex        =   16
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label1 
      Caption         =   "Descripcion banco"
      Height          =   195
      Index           =   1
      Left            =   900
      TabIndex        =   12
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Codigo "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   615
   End
End
Attribute VB_Name = "frmBancos2"
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
Private HaPulsadoEnter As Boolean
Private NombreTabla As String  'Nombre de la tabla o de la consulta
Private NumRegistro As Long
Private kCampo As Integer
Private TotalReg As Long


'CREATE TABLE Bancos(
'    Id INTEGER PRIMARY KEY,
'    Observa TEXT(40),
'    Entidad TEXT(4),
'    Sucursal TEXT(4),
'    CodControl TEXT(2),
'    Cuenta TEXT(10))

Private Sub cmdAceptar_Click()
Dim rs As ADODB.Recordset
Dim cad As String
Dim i As Integer

Screen.MousePointer = vbHourglass
On Error GoTo Error1
If Modo = 3 Then
    If DatosOk Then
        
        Set rs = New ADODB.Recordset
        rs.CursorType = adOpenKeyset
        rs.LockType = adLockOptimistic
        rs.Open NombreTabla, conn, , , adCmdTable
        rs.AddNew
        '--------------------
   
        rs.Fields(0) = Text1(0).Text
        For i = 0 To Text1.Count - 1
            rs.Fields(i) = Text1(i).Text
        Next i
        rs.Update
        rs.Close
        espera 0.2
        Data1.Refresh
        MsgBox "Registro insertado.", vbInformation
        PonerModo 0
        Label2.Caption = "Insertado"
    End If
    Else
    If Modo = 4 Then
        'Modificar
'        'Recordamos que el text(0) tiene el codigo y no lo puede cambiar
'        For I = 1 To Text1.Count - 1
'            If Not CmpCam(Text1(I).Tag, Text1(I).Text) Then _
'                GoTo Error1
'        Next I
        If DatosOk Then
            'Ahora modificamos
            cad = "Select * from " & NombreTabla
            cad = cad & " WHERE id=" & Data1.Recordset.Fields(0)
            Set rs = New ADODB.Recordset
            rs.CursorType = adOpenKeyset
            rs.LockType = adLockOptimistic
            rs.Open cad, conn, , , adCmdText
            'Almacenamos para luego buscarlo
            cad = rs!Id
            

            For i = 1 To Text1.Count - 1
                rs.Fields(i) = Text1(i).Text
            Next i
                
            rs.Update
            rs.Close
            'MsgBox "El registro ha sido modificado", vbInformation
            PonerModo 2
            'Hay que refresca el DAta1
            Data1.Refresh
            'Hay que volver a poner el registro donde toca
            Data1.Recordset.MoveFirst
            i = 1
            While i > 0
                If Data1.Recordset.Fields(0) = cad Then
                    i = 0
                    Else
                        Data1.Recordset.MoveNext
                        If Data1.Recordset.EOF Then i = 0
                End If
            Wend
            If Data1.Recordset.EOF Then
                NumRegistro = TotalReg
                Data1.Recordset.MoveLast
            End If
            Label2.Caption = NumRegistro & " de " & TotalReg
        End If
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
'A�adiremos el boton de aceptar y demas objetos para insertar
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
    Label2.Caption = "B�squeda"
    PonerModo 1
    Text1(0).SetFocus
    Else
        HacerBusqueda
        If TotalReg = 0 Then
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
        NumRegistro = 1
    Case 1
        Data1.Recordset.MovePrevious
        NumRegistro = NumRegistro - 1
        If Data1.Recordset.BOF Then
            Data1.Recordset.MoveFirst
            NumRegistro = 1
        End If
    Case 2
        Data1.Recordset.MoveNext
        NumRegistro = NumRegistro + 1
        If Data1.Recordset.EOF Then
            Data1.Recordset.MoveLast
            NumRegistro = TotalReg
        End If
    Case 3
        Data1.Recordset.MoveLast
        NumRegistro = TotalReg
End Select
PonerCampos
End Sub

Private Sub BotonModificar()
'---------
'MODIFICAR
'----------

If Data1.Recordset.EOF Then Exit Sub

'A�adiremos el boton de aceptar y demas objetos para insertar
cmdAceptar.Caption = "Modificar"
PonerModo 4
'Escondemos el navegador y ponemos insertando
'Como el campo 1 es clave primaria, NO se puede modificar
Text1(0).Locked = True
DespalzamientoVisible False
Label2.Caption = "Modificar"
End Sub

Private Sub BotonEliminar()
Dim cad As String
Dim i As Integer

'Ciertas comprobaciones
If Data1.Recordset.RecordCount = 0 Then Exit Sub
'Pregunta
cad = "Seguro que desea eliminar de la BD el registro:"
cad = cad & vbCrLf & "Cod: " & Data1.Recordset.Fields(0)
cad = cad & vbCrLf & "Nombre: " & Data1.Recordset.Fields(1)
i = MsgBox(cad, vbQuestion + vbYesNo)
If i = vbYes Then
    'Hay que eliminar
    On Error GoTo Error2
    Screen.MousePointer = vbHourglass
    Data1.Recordset.Delete
    Data1.Refresh
    If Data1.Recordset.EOF Then
        'Solo habia un registro
        LimpiarCampos
        PonerModo 0
        Else
            If NumRegistro = TotalReg Then
                    Data1.Recordset.MoveLast
                    NumRegistro = NumRegistro - 1
                    Else
                        For i = 1 To NumRegistro - 1
                            Data1.Recordset.MoveNext
                        Next i
            End If
            TotalReg = TotalReg - 1
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
NombreTabla = "Bancos"
Ordenacion = " ORDER BY id"
HaPulsadoEnter = False

'ASignamos un SQL al DATA1
Data1.ConnectionString = conn
Data1.RecordSource = "Select * from " & NombreTabla & " WHERE id=-1"
Data1.Refresh
PonerModo 0
End Sub



Private Sub LimpiarCampos()
Dim i
For i = 0 To Text1.Count - 1
    Text1(i).Text = ""
Next i
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
        Keypress KeyAscii
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Modo = 1 Then
    Text1(Index).BackColor = vbWhite
Else
    If Modo >= 3 Then
        If Index <> 1 And Index < 6 Then
            Text1(Index).Text = Trim(Text1(Index).Text)
            If Text1(Index).Text = "" Then Exit Sub
            If Not IsNumeric(Text1(Index).Text) Then
                MsgBox "Campo num�rico : " & Text1(Index).Text, vbExclamation
                Text1(Index).Text = ""
                PonleFoco Text1(Index)
            Else
                Select Case Index
                Case 4
                    Text1(Index).Text = Format(Text1(Index).Text, "00")
                Case 5
                    Text1(Index).Text = Format(Text1(Index).Text, "0000000000")
                Case Else
                    Text1(Index).Text = Format(Text1(Index).Text, "0000")
                End Select
            End If
        End If
    End If
End If
End Sub


Private Sub HacerBusqueda()
Dim CadB As String

If Text1(kCampo).Text = "" Then Exit Sub
'Prueba de pascual jajajaja
Dim c1 As String   'el nombre del campo
Dim Tipo As Long
Dim aux1

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
    MsgBox "No hay ning�n registro en la tabla " & NombreTabla, vbInformation
    Screen.MousePointer = vbDefault
    TotalReg = 0
    Exit Sub
    'StatusBar1.Panels(2).Text = ""
    'PonerModo 0
    Else
        DespalzamientoVisible True
        PonerModo 2
        'Data1.Recordset.MoveLast
        Data1.Recordset.MoveFirst
        TotalReg = Data1.Recordset.RecordCount
        NumRegistro = 1
        PonerCampos
End If

Data1.ConnectionString = conn
Data1.RecordSource = CadenaConsulta
Data1.Refresh
TotalReg = Data1.Recordset.RecordCount
Screen.MousePointer = vbDefault
End Sub


Private Sub PonerCampos()
Dim i As Integer
    For i = 0 To Text1.Count - 1
        Text1(i).Text = DBLet(Data1.Recordset(i), "N")
    Next i
'    For i = 0 To Text1.Count - 1
'        If Text1(i).Text <> "" And i <> 1 Then
'                Select Case i
'                Case 4
'                    Text1(i).Text = Format(Text1(i).Text, "00")
'                Case 5
'                    Text1(i).Text = Format(Text1(i).Text, "0000000000")
'                Case Else
'                    Text1(i).Text = Format(Text1(i).Text, "0000")
'                End Select
'
'        End If
'    Next i
    Label2.Caption = NumRegistro & " de " & TotalReg
End Sub

Private Sub PonerModo(Kmodo As Integer)
Dim i As Integer
Dim B As Boolean

If Modo = 1 Then
    For i = 0 To Text1.Count - 1
        Text1(i).BackColor = vbWhite
    Next i
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
B = (Modo = 2) Or Modo = 0
For i = 0 To Text1.Count - 1
    Text1(i).Locked = B
Next i

End Sub


Private Function DatosOk() As Boolean
Dim rs As ADODB.Recordset
Dim cad As String
Dim i As Integer

DatosOk = False
'Haremos las comprobaciones necesarias de los campos
'Cad = ComprobarCampos
'If Cad <> "" Then
'    MsgBox Cad, vbExclamation
'    Exit Function
'End If


For i = 0 To Text1.Count - 1
    Text1(i).Text = Trim(Text1(i).Text)
    If Not CmpCam(Text1(i).Tag, Text1(i).Text) Then Exit Function
Next i



'Veremos el formato de los numeros
For i = 2 To 3
    If Text1(i).Text <> "" Then
        If InStr(1, Text1(i).Text, ",") = 0 Then
            'No tiene comas. Cambiamos el punto por comas
            Text1(i).Text = TransformaPuntosComas(Text1(i).Text)
        End If
    End If
Next i

'Llegados a este punto los datos son correctos en valores
'Ahora comprobaremos otras cosas
'Este apartado dependera del formulario y la tabla
If Modo = 3 Then
    cad = "Select * from " & NombreTabla
    cad = cad & " WHERE id=" & Text1(0).Text
    
    Set rs = New ADODB.Recordset
    rs.Open cad, conn, , , adCmdText
    If Not rs.EOF Then
        MsgBox "Ya existe un registro con ese c�digo.", vbExclamation
        rs.Close
        Exit Function
    End If
    rs.Close
    Set rs = Nothing
End If
'Al final todo esta correcto
DatosOk = True
End Function


Private Sub SugerirCodigoSiguiente()
Dim cad
Dim rs
'Sugeriremos el codigo siguiente.
'Obviamente depende en TOTAL medida de que tabla estemos trabajando
cad = "Select Max(id) from " & NombreTabla
Text1(0).Text = 1
Set rs = New ADODB.Recordset
rs.Open cad, conn, , , adCmdText
If Not rs.EOF Then
    If Not IsNull(rs.Fields(0)) Then
        Text1(0).Text = rs.Fields(0) + 1
    End If
End If
rs.Close
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index < 9 And Button.Index > 5 Then
    If vUsu.Nivel > 2 Then  'Prohibimos consulta
        MsgBox "No tiene autorizacion para cambiar datos", vbExclamation
        Exit Sub
    End If
End If
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
Case Else

End Select
End Sub


Private Sub DespalzamientoVisible(bol As Boolean)
Dim i
For i = 14 To 17
    Toolbar1.Buttons(i).Visible = bol
Next i
End Sub


