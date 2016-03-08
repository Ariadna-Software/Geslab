VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMantenusu 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de usuarios"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmMantenusu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameUsuario 
      Height          =   4815
      Left            =   2640
      TabIndex        =   16
      Top             =   480
      Width           =   4815
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "frmMantenusu.frx":030A
         Left            =   120
         List            =   "frmMantenusu.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2040
         Width           =   2415
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   20
         Text            =   "Text1"
         Top             =   3240
         Width           =   1575
      End
      Begin VB.CommandButton cmdFrameUsu 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   3360
         TabIndex        =   22
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdFrameUsu 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   2040
         TabIndex        =   21
         Top             =   4200
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   2280
         PasswordChar    =   "*"
         TabIndex        =   19
         Text            =   "Text1"
         Top             =   2760
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   315
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Text            =   "Text1"
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "NUEVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   28
         Top             =   480
         Width           =   2415
      End
      Begin VB.Shape Shape1 
         Height          =   1455
         Left            =   120
         Top             =   2520
         Width           =   4455
      End
      Begin VB.Label Label4 
         Caption         =   "Confirma Password"
         Height          =   255
         Index           =   3
         Left            =   360
         TabIndex        =   27
         Top             =   3360
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Password"
         Height          =   255
         Index           =   2
         Left            =   360
         TabIndex        =   26
         Top             =   2760
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Nivel"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Nombre completo"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Login"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Salir"
      Height          =   375
      Left            =   8040
      TabIndex        =   0
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Frame FrameNormal 
      BorderStyle     =   0  'None
      Height          =   5895
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   9255
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   3135
         Left            =   3480
         TabIndex        =   29
         Top             =   2280
         Width           =   5895
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   1680
         Top             =   5520
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMantenusu.frx":033C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   1815
         Left            =   3480
         TabIndex        =   8
         Top             =   360
         Width           =   5655
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "frmMantenusu.frx":08D6
            Left            =   240
            List            =   "frmMantenusu.frx":08E6
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1320
            Width           =   2415
         End
         Begin VB.TextBox Text4 
            Height          =   315
            Left            =   120
            TabIndex        =   9
            Text            =   "Text1"
            Top             =   480
            Width           =   4335
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre completo"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Nivel"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   960
            Width           =   2295
         End
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   0
         Left            =   120
         Picture         =   "frmMantenusu.frx":0919
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Nuevo usuario"
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton cmdEmp 
         Height          =   375
         Index           =   0
         Left            =   3480
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Nueva bloqueo empresa"
         Top             =   5400
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   1
         Left            =   600
         Picture         =   "frmMantenusu.frx":0A1B
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Modificar usuario"
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton cmdUsu 
         Height          =   375
         Index           =   2
         Left            =   1080
         Picture         =   "frmMantenusu.frx":0B1D
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Eliminar usuario"
         Top             =   5400
         Width           =   375
      End
      Begin VB.CommandButton cmdEmp 
         Height          =   375
         Index           =   1
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Eliminar bloqueo empresa"
         Top             =   5400
         Visible         =   0   'False
         Width           =   375
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   2775
         Left            =   3480
         TabIndex        =   7
         Top             =   2520
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4895
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5821
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Resum."
            Object.Width           =   2293
         EndProperty
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   4935
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   8705
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Login"
            Object.Width           =   3528
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Usuarios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3480
         TabIndex        =   14
         Top             =   120
         Width           =   2895
      End
      Begin VB.Label Label1 
         Caption         =   "Empresas NO permitidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   13
         Top             =   2280
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmMantenusu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PrimeraVez As Boolean
Dim SQL As String
Dim I As Integer
Dim miRsAux As ADODB.Recordset



'Private Sub cmdEmp_Click(Index As Integer)
'Dim Cont As Integer
'
'    If ListView1.SelectedItem Is Nothing Then
'        MsgBox "Seleccione un usuario", vbExclamation
'        Exit Sub
'    End If
'
'    If Index = 0 Then
'
'
'        'nueva Empresa bloqueada para el usuario
'        CadenaDesdeOtroForm = ""
'        frmMensajes.Opcion = 4
'        frmMensajes.Show vbModal
'        If CadenaDesdeOtroForm <> "" Then
'            Cont = RecuperaValor(CadenaDesdeOtroForm, 1)
'            If Cont = 0 Then Exit Sub
'            For I = 1 To Cont
'                'No hacemos nada
'            Next I
'            For I = 0 To Cont - 1
'                SQL = RecuperaValor(CadenaDesdeOtroForm, I + Cont + 2)
'                InsertarEmpresa CInt(SQL)
'            Next I
'
'        Else
'            Exit Sub
'        End If
'
'    Else
'        If ListView2.SelectedItem Is Nothing Then Exit Sub
'        SQL = "Va a  desbloquear el acceso" & vbCrLf
'        SQL = SQL & vbCrLf & "a la empresa:   " & ListView2.SelectedItem.SubItems(1) & vbCrLf
'        SQL = SQL & "para el usuario:   " & ListView1.SelectedItem.SubItems(1) & vbCrLf & vbCrLf & "     ¿Desea continuar?"
'        If MsgBox(SQL, vbQuestion + vbYesNoCancel) = vbYes Then
'            SQL = "Delete FROM Usuarios.usuarioempresa WHERE codusu =" & ListView1.SelectedItem.Text
'            SQL = SQL & " AND codempre = " & ListView2.SelectedItem.Text
'            Conn.Execute SQL
'        Else
'            Exit Sub
'        End If
'    End If
'    'Llegados aqui recargamos los datos del usuario
'    Screen.MousePointer = vbHourglass
'    DatosUsusario
'    Screen.MousePointer = vbDefault
'End Sub
'
'
'Private Sub InsertarEmpresa(Empresa As Integer)
'    SQL = "INSERT INTO Usuarios.usuarioempresa(codusu,codempre) VALUES ("
'    SQL = SQL & ListView1.SelectedItem.Text & "," & Empresa & ")"
'    On Error Resume Next
'    Conn.Execute SQL
'    If Err.Number <> 0 Then
'        MuestraError Err.Number, Err.Description
'    Else
'
'    End If
'
'End Sub


Private Sub cmdFrameUsu_Click(Index As Integer)



    If Index = 0 Then
        For I = 0 To Text2.Count - 1
            Text2(I).Text = Trim(Text2(I).Text)
            If Text2(I).Text = "" Then
                MsgBox Label4(I).Caption & " requerido.", vbExclamation
                Exit Sub
            End If
        Next I
        
        If Combo2.ListIndex < 0 Then
            MsgBox "Seleccione un nivel de acceso", vbExclamation
            Exit Sub
        End If
    
        'Password
        If Text2(2).Text <> Text2(3).Text Then
            MsgBox "Password y confirmacion de password no coinciden", vbExclamation
            Exit Sub
        End If
        
        'Compruebo que el login es unico
        If UCase(Label6.Caption) = "NUEVO" Then
            Set miRsAux = New ADODB.Recordset
            SQL = "Select login from Usuarios where login='" & Text2(0).Text & "'"
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            SQL = ""
            If Not miRsAux.EOF Then SQL = "Ya existe en la tabla usuarios uno con el login: " & miRsAux.Fields(0)
            miRsAux.Close
            Set miRsAux = Nothing
            If SQL <> "" Then
                MsgBox SQL, vbExclamation
                Exit Sub
            End If
        End If
        InsertarModificar
        
        
    End If
    'Cargar usuarios
    If UCase(Label6.Caption) = "NUEVO" Then
        CargaUsuarios
    Else
        'Pero cargamos el tag como coresponde
        ListView1.SelectedItem.Tag = Combo2.ItemData(Combo2.ListIndex) & "|" & Text2(1).Text & "|"
    
        DatosUsusario
    End If
    'Para ambos casos
    Me.FrameUsuario.Visible = False
    Me.FrameNormal.Enabled = True
End Sub


Private Sub InsertarModificar()
On Error GoTo EInsertarModificar

    Set miRsAux = New ADODB.Recordset
    If UCase(Label6.Caption) = "NUEVO" Then
        
        'Nuevo
        SQL = "Select max(codusu) from Usuarios"
        miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        I = 0
        If Not miRsAux.EOF Then I = DBLet(miRsAux.Fields(0), "N")
        miRsAux.Close
        I = I + 1
        
        SQL = "INSERT INTO usuarios (codusu, nomusu,  nivelusu, login, passwordpropio) VALUES ("
        SQL = SQL & I
        SQL = SQL & ",'" & Text2(1).Text & "',"
        'Combo
        SQL = SQL & Combo2.ItemData(Combo2.ListIndex) & ",'"
        SQL = SQL & Text2(0).Text & "','"
        SQL = SQL & Text2(3).Text & "')"
        
    Else
        SQL = "UPDATE Usuarios Set nomusu='" & Text2(1).Text
        
        'Si el combo es administrador compruebo que no fuera en un principio SUPERUSUARIO
        If Combo2.ListIndex = 2 Then
            'Si el combo1 es 3 entonces es super
            If Combo1.ListIndex = 3 Then
                I = 0
            Else
                I = 1
            End If
        Else
            I = Combo2.ItemData(Combo2.ListIndex)
        End If
        SQL = SQL & "' , nivelusu =" & I
        'SQL = SQL & "  , login = '" & Text2(2).Text
        SQL = SQL & "  , passwordpropio = '" & Text2(3).Text
        SQL = SQL & "' WHERE codusu = " & ListView1.SelectedItem.Text
    End If
    Conn.Execute SQL

    Exit Sub
EInsertarModificar:
    MuestraError Err.Number, "EInsertarModificar"
End Sub



Private Sub cmdUsu_Click(Index As Integer)
    
    
    Select Case Index
    Case 0, 1
        If Index = 0 Then
            'Nuevo usuario
            Limpiar Me
            Label6.Caption = "NUEVO"
            I = 0 'Para el foco
        Else
            'Modificar
            If ListView1.SelectedItem Is Nothing Then Exit Sub
            Label6.Caption = "MODIFICAR"
            Set miRsAux = New ADODB.Recordset
            SQL = "Select * from usuarios where codusu = " & ListView1.SelectedItem.Text
            miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If miRsAux.EOF Then
                MsgBox "Error inesperado: Leer datos usuarios", vbExclamation
            Else
                Text2(0).Text = miRsAux!Login
                Text2(1).Text = miRsAux!nomusu
                Text2(2).Text = miRsAux!passwordpropio
                Text2(3).Text = miRsAux!passwordpropio
                I = miRsAux!nivelusu
                If I < 2 Then
                    Combo2.ListIndex = 2
                Else
                    If I = 2 Then
                        Combo2.ListIndex = 1
                    Else
                        Combo2.ListIndex = 0
                    End If
                End If
            End If
            I = 1 'Para el foco
        End If
        Text2(0).Enabled = (Index = 0)
        Me.FrameNormal.Enabled = False
        Me.FrameUsuario.Visible = True
        Text2(I).SetFocus
    Case 2
        If ListView1.SelectedItem Is Nothing Then Exit Sub
        If ListView1.SelectedItem.Text = CStr(vUsu.Codigo) Then
            MsgBox "Es el usuario actual.", vbExclamation
            Exit Sub
        End If
        SQL = "Seguro que desea eliminar al usuario " & Trim(ListView1.SelectedItem.SubItems(1)) & " ?"
        If MsgBox(SQL, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        Screen.MousePointer = vbHourglass
        If Eliminar Then CargaUsuarios
        Screen.MousePointer = vbDefault
    End Select

End Sub

Private Function Eliminar() As Boolean
    On Error Resume Next
    Eliminar = False
    SQL = "DELETE from USuarios where codusu =" & ListView1.SelectedItem.Text
    Conn.Execute SQL
    If Err.Number <> 0 Then
        MuestraError Err.Number, Err.Description
    Else
        Eliminar = True
    End If
    
End Function
Private Sub Combo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        Me.ListView1.SmallIcons = ImageList1
        Me.ListView2.SmallIcons = ImageList1
        CargaUsuarios
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Me.FrameUsuario.Visible = False
    Me.FrameNormal.Enabled = True
End Sub



Private Sub CargaUsuarios()
Dim itm As ListItem

    ListView1.ListItems.Clear
    Set miRsAux = New ADODB.Recordset
    '                               Aquellos usuarios k tengan nivel usu -1 NO son de conta
    '  QUitamos codusu=0 pq es el usuario ROOT
    SQL = "Select * from Usuarios where nivelusu >=0 and codusu > 0 order by codusu"
    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not miRsAux.EOF
        Set itm = ListView1.ListItems.Add
        itm.Text = miRsAux!CodUsu
        itm.SubItems(1) = miRsAux!Login
        itm.SmallIcon = 1
        'Nombre y nivel de usuario
        SQL = miRsAux!nivelusu & "|" & miRsAux!nomusu & "|"
        itm.Tag = SQL
        'Sig
        miRsAux.MoveNext
    Wend
    miRsAux.Close
    Set miRsAux = Nothing
    If ListView1.ListItems.Count > 0 Then
        Set ListView1.SelectedItem = ListView1.ListItems(1)
        DatosUsusario
    End If

End Sub



Private Sub DatosUsusario()
Dim itmX As ListItem
On Error GoTo EDatosUsu

    If ListView1.SelectedItem Is Nothing Then
        Text4.Text = ""
        Combo1.ListIndex = -1
        Exit Sub
    End If


    Text4.Text = RecuperaValor(ListView1.SelectedItem.Tag, 2)
    'NIVEL
    SQL = RecuperaValor(ListView1.SelectedItem.Tag, 1)
    '                           COMBO                      en Bd
    '                       0.- Consulta                     3
    '                       1.- Normal                       2
    '                       2.- Administrador                1
    '                       3.- SuperUsuario (root)          0
    If Not IsNumeric(SQL) Then SQL = 3
    Select Case Val(SQL)
    Case 2
        Combo1.ListIndex = 1
    Case 1
        Combo1.ListIndex = 2
    Case 0
        Combo1.ListIndex = 3
    Case Else
        Combo1.ListIndex = 0
    End Select
    Exit Sub
'    ListView2.ListItems.Clear
'    SQL = ListView2.Tag & ListView1.SelectedItem.Text
'    Set miRsAux = New ADODB.Recordset
'    miRsAux.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'    While Not miRsAux.EOF
'        Set ItmX = ListView2.ListItems.Add
'        ItmX.Text = miRsAux.Fields(0)
'        ItmX.SubItems(1) = miRsAux!nomempre
'        ItmX.SubItems(2) = miRsAux!nomresum
'        ItmX.SmallIcon = 20
'
'        miRsAux.MoveNext
'    Wend
'    miRsAux.Close
'    Set miRsAux = Nothing
    Exit Sub
EDatosUsu:
    MuestraError Err.Number, Err.Description
End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
    Screen.MousePointer = vbHourglass
    DatosUsusario
    Screen.MousePointer = vbDefault
End Sub



Private Sub Text2_GotFocus(Index As Integer)
    With Text2(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

