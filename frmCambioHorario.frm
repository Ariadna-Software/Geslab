VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCambioHorario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio de horario masivo"
   ClientHeight    =   6570
   ClientLeft      =   150
   ClientTop       =   810
   ClientWidth     =   9315
   Icon            =   "frmCambioHorario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   9315
   Begin VB.Frame FrameModificarTicaje 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   1560
      TabIndex        =   24
      Top             =   1800
      Width           =   6555
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   1
         Left            =   5160
         TabIndex        =   35
         Top             =   1750
         Width           =   975
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Index           =   0
         Left            =   3960
         TabIndex        =   34
         Top             =   1750
         Width           =   975
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   5
         Left            =   5040
         TabIndex        =   32
         Top             =   1250
         Width           =   1035
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   4
         Left            =   3000
         TabIndex        =   30
         Top             =   1250
         Width           =   1035
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   3
         Left            =   1200
         TabIndex        =   26
         Top             =   1250
         Width           =   1035
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   2
         Left            =   1200
         TabIndex        =   25
         Top             =   720
         Width           =   1035
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "HORA"
         Height          =   315
         Index           =   2
         Left            =   4200
         TabIndex        =   33
         Top             =   1250
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "FIN"
         Height          =   315
         Index           =   1
         Left            =   2400
         TabIndex        =   31
         Top             =   1250
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "FECHA"
         Height          =   315
         Left            =   360
         TabIndex        =   29
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Modificar un ticaje para los trabajadores seleccionados"
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
         Height          =   315
         Left            =   240
         TabIndex        =   28
         Top             =   240
         Width           =   5775
      End
      Begin VB.Shape Shape3 
         BorderWidth     =   3
         Height          =   2115
         Left            =   0
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "INICIO"
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   27
         Top             =   1250
         Width           =   615
      End
   End
   Begin VB.Frame FrameTicada 
      BorderStyle     =   0  'None
      Height          =   2355
      Left            =   1440
      TabIndex        =   16
      Top             =   1800
      Width           =   6555
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   1
         Left            =   3720
         TabIndex        =   18
         Top             =   900
         Width           =   1035
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   17
         Top             =   900
         Width           =   1035
      End
      Begin VB.CommandButton cmdHacerCambio 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   3
         Left            =   3720
         TabIndex        =   20
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CommandButton cmdHacerCambio 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   2
         Left            =   4980
         TabIndex        =   22
         Top             =   1620
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "HORA"
         Height          =   315
         Left            =   3000
         TabIndex        =   23
         Top             =   960
         Width           =   615
      End
      Begin VB.Shape Shape2 
         BorderWidth     =   3
         Height          =   2115
         Left            =   120
         Top             =   120
         Width           =   6375
      End
      Begin VB.Label Label7 
         Caption         =   "Generar nuevo ticaje para los trabajadores seleccionados"
         Height          =   315
         Left            =   600
         TabIndex        =   21
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "FECHA"
         Height          =   315
         Left            =   720
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.Frame FrameNuevoH 
      BorderStyle     =   0  'None
      Height          =   2175
      Left            =   1440
      TabIndex        =   7
      Top             =   1800
      Width           =   6555
      Begin VB.CommandButton cmdHacerCambio 
         Caption         =   "&Cancelar"
         Height          =   375
         Index           =   1
         Left            =   4980
         TabIndex        =   14
         Top             =   1620
         Width           =   1095
      End
      Begin VB.CommandButton cmdHacerCambio 
         Caption         =   "Aceptar"
         Height          =   375
         Index           =   0
         Left            =   3720
         TabIndex        =   13
         Top             =   1620
         Width           =   1095
      End
      Begin VB.TextBox Text2 
         BackColor       =   &H80000018&
         Enabled         =   0   'False
         Height          =   345
         Left            =   2460
         TabIndex        =   10
         Top             =   1080
         Width           =   3615
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1560
         TabIndex        =   9
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Horario"
         Height          =   315
         Left            =   360
         TabIndex        =   11
         Top             =   1140
         Width           =   1035
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   2220
         Picture         =   "frmCambioHorario.frx":030A
         Top             =   1140
         Width           =   240
      End
      Begin VB.Label Label4 
         Caption         =   "Nuevo HORARIO para los trabajadores seleccionados"
         Height          =   315
         Left            =   420
         TabIndex        =   8
         Top             =   420
         Width           =   5775
      End
      Begin VB.Shape Shape1 
         BorderWidth     =   3
         Height          =   2055
         Left            =   60
         Top             =   60
         Width           =   6435
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   1275
      Left            =   2880
      TabIndex        =   4
      Top             =   2400
      Visible         =   0   'False
      Width           =   3495
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Insertando datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   3195
      End
      Begin VB.Label Label2 
         Caption         =   "Insertando datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   5
         Top             =   180
         Width           =   3195
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5340
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCambioHorario.frx":040C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   5880
      Width           =   9135
      Begin VB.CommandButton cmdCambiar 
         Caption         =   "&Cambiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6840
         TabIndex        =   12
         Top             =   180
         Width           =   975
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   375
         Left            =   8040
         TabIndex        =   3
         Top             =   180
         Width           =   915
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   1
         Left            =   600
         Picture         =   "frmCambioHorario.frx":09A6
         Top             =   240
         Width           =   240
      End
      Begin VB.Image imgCheck 
         Height          =   240
         Index           =   0
         Left            =   240
         Picture         =   "frmCambioHorario.frx":0AF0
         Top             =   240
         Width           =   240
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   9657
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Horario"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Categoria"
         Object.Width           =   4410
      EndProperty
   End
   Begin MSComctlLib.ListView ListView2 
      Height          =   5475
      Left            =   120
      TabIndex        =   36
      Top             =   360
      Width           =   9120
      _ExtentX        =   16087
      _ExtentY        =   9657
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Horario"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Categoria"
         Object.Width           =   4410
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Insercion TICADA masiva"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   300
      Index           =   1
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   9135
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Trabajadores para el cambio de horario"
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
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9030
   End
   Begin VB.Menu mnAgregar 
      Caption         =   "Agregar"
      Begin VB.Menu mnHorario 
         Caption         =   "Horario"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnSeccion 
         Caption         =   "Seccion"
         Shortcut        =   ^S
      End
      Begin VB.Menu mn1Trabajador 
         Caption         =   "Trabajador (Nombre)"
         Shortcut        =   ^N
      End
      Begin VB.Menu mn1TrabajaCod 
         Caption         =   "Trabajador (Cod)"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnCategoria 
         Caption         =   "Categoria"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnCodigoTarjeta 
         Caption         =   "Por c�digo tarjeta"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnEliminar 
      Caption         =   "Elimininar"
      Begin VB.Menu mnEliminarUno 
         Caption         =   "Quitar uno"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnLimpiar 
         Caption         =   "Limpiar"
         Shortcut        =   ^L
      End
   End
End
Attribute VB_Name = "frmCambioHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Opcion As Byte
    '0.- Cambio horario masivo
    '1.- Insercion marcajes masivo
    '2.- Insercion de ticaje desde ticaje actual
    '3.- Modificacion de unos ticajes para unos trabajadores
    
    
    '6.- Modificacion de horario, peeero de salida ponemos los trabajadores
    '    que estan ya de altas, estan marcados con un check


Public Fecha As Date

Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmBu As frmBuscaGrid
Attribute frmBu.VB_VarHelpID = -1
Private vIndice As Byte
Private Devolucion As String


Dim RS As ADODB.Recordset
Dim Total As Long
Dim i As Long
Dim PrimeraVez As Boolean

Private Sub AccionesBuscaGrid(indice As Integer)
    Devolucion = ""
    Set frmBu = New frmBuscaGrid
    'de momento solo esta el CERO. No hago nada
    'Desc|Tabla|Tipo|Porcentaje
    frmBu.vBusqueda = ""
    frmBu.vCampos = "Codigo|Trabajadores|idtrabajador|N|000|25�Nombre|Trabajadores|nomtrabajador|T||65�"
    frmBu.vDevuelve = "0|"
    frmBu.vselElem = 0
    frmBu.vSQL = ""
    frmBu.vTabla = "trabajadores"
    frmBu.vTitulo = "Trabajadores"
    frmBu.Show vbModal
    Set frmBu = Nothing
    If Devolucion <> "" Then
        Devolucion = RecuperaValor(Devolucion, 1)
        InsertarTrabajador
    End If
    
    
End Sub


Private Sub Acciones(Index As Integer)

    Set frmB = New frmBusca
    'Ponemos los valores para abrir
    vIndice = Index
    Devolucion = ""
    Select Case Index
    Case 0
        vIndice = 0
        frmB.Tabla = "Trabajadores"
        frmB.CampoBusqueda = "NomTrabajador"
        frmB.CampoCodigo = "IdTrabajador"
        frmB.TipoDatos = 3
        frmB.Titulo = "EMPLEADOS"
    Case 1
        '## Es eliminar
        If ListView1.ListItems.Count > 0 Then HacerEliminar
        Exit Sub
    Case 2
  
        frmB.Tabla = "Secciones"
        frmB.CampoBusqueda = "Nombre"
        frmB.CampoCodigo = "IdSeccion"
        frmB.TipoDatos = 3
        frmB.Titulo = "SECCIONES"
    
    Case 3
        frmB.Tabla = "Horarios"
        frmB.CampoBusqueda = "NomHorario"
        frmB.CampoCodigo = "IdHorario"
        frmB.TipoDatos = 3
        frmB.Titulo = "HORARIOS"
    
    Case 4
        'Agregar seccion
        frmB.Tabla = "Categorias"
        frmB.CampoBusqueda = "Nomcategoria"
        frmB.CampoCodigo = "Idcategoria"
        frmB.TipoDatos = 3
        frmB.Titulo = "CATEGORIAS"
    
    
    Case 10
        Devolucion = "Seguro que desea limpiar la lista de trabajadores?"
        If MsgBox(Devolucion, vbQuestion + vbYesNoCancel) = vbYes Then
            Screen.MousePointer = vbHourglass
            conn.Execute "Delete from tmpCambioHor"
            ListView1.ListItems.Clear
            Screen.MousePointer = vbDefault
        End If
        Exit Sub
        
    Case 11
        'AGREGAMOS TRABAJADORES DESDE HASTA CODIGO TARJETA
        PedirDesdeHastaTarjeta
        Exit Sub
        
    End Select
    frmB.MostrarDeSalida = True
    frmB.Show vbModal
    Set frmB = Nothing
    Me.Refresh
    If Devolucion <> "" Then
        Screen.MousePointer = vbHourglass
        'Insertamos
        If Index = 0 Then
            'Un trabajador
            InsertarTrabajador
        Else
            Label2.Caption = "Insertar datos"
            Label3.Caption = ""
            Frame2.Visible = True
            Me.Refresh
            'Insertamos toda la seccion o grupo de trabajadores
            InsertarSeccionHorario (Index)
            'Desahacemos
            Frame2.Visible = False
            Me.Refresh
        End If
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdCambiar_Click()

    If Opcion = 6 Then
        Total = 0
        For i = 1 To ListView2.ListItems.Count
            If ListView2.ListItems(i).Checked Then
                Total = 1
                Exit For
            End If
        Next i
        
        If Total = 0 Then
            MsgBox "Marque algun trabajador para cambiarle el horario", vbExclamation
            Exit Sub
        End If
        
        Text1.Text = ""
        Text2.Text = ""
        FrameNuevoH.Visible = True
        Text1.SetFocus

        
    
    Else
        If ListView1.ListItems.Count > 0 Then
            MenuEnable False
            Select Case Opcion
            Case 0
                Text1.Text = ""
                Text2.Text = ""
                FrameNuevoH.Visible = True
                Text1.SetFocus
            Case 1, 2
                If Opcion <> 2 Then
                    Text3(0).Text = ""
                    i = 0
                Else
                    i = 1
                End If
                Text3(1).Text = ""
                FrameTicada.Visible = True
                Text3(i).SetFocus
            Case 3
                FrameModificarTicaje.Visible = True
            End Select
        Else
            MsgBox "Seleccione trabajadores.", vbExclamation
        End If
    End If
End Sub

Private Sub cmdHacerCambio_Click(Index As Integer)
    Devolucion = ""
    Select Case Index
    Case 1, 2
        'Na

    Case 0
        If Text1.Text = "" Then
            MsgBox "Seleccione el nuevo horario", vbExclamation
            Exit Sub
        End If
        Devolucion = Text1.Text
        MenuEnable True
        If Devolucion <> "" Then
            Screen.MousePointer = vbHourglass
            Label2.Caption = "Hacer cambio"
            Label3.Caption = ""
            Frame2.Visible = True
            Me.Refresh
            HacerCambio
            Frame2.Visible = False
            Me.Refresh
            Screen.MousePointer = vbDefault
            MsgBox "Cambios realizados.", vbExclamation
        End If
    Case 3
        'Nueva ticaje
        Devolucion = ""
        If Text3(0).Text <> "" And Text3(1).Text <> "" Then
            If Not IsDate(Text3(0).Text) Then
                Devolucion = "Fecha incorrecta"
            Else
                If Not IsDate(Text3(1).Text) Then
                    Devolucion = "Hora incorrecta"
                End If
            End If
        Else
            Devolucion = "Ponga la fecha y la hora"
        End If
            
        If Devolucion = "" Then
            Screen.MousePointer = vbHourglass
            Me.FrameTicada.Visible = False
            Label2.Caption = "Generar ticada"
            Label3.Caption = ""
            Frame2.Visible = True
            Me.Refresh
            HacerTicada
            Frame2.Visible = False
            Me.Refresh
            Screen.MousePointer = vbDefault
            MsgBox "Proceso TICADA finalizado.", vbExclamation
            If Opcion = 2 Then
                Unload Me
                Exit Sub
            End If
        Else
            MsgBox "Error: " & Devolucion, vbExclamation
            Exit Sub
        End If
    End Select
    Devolucion = ""
    FrameNuevoH.Visible = False
    FrameTicada.Visible = False
    MenuEnable True
End Sub

Private Sub cmdModificar_Click(Index As Integer)
Dim cad As String

    If Index = 1 Then
        MenuEnable True
        Me.FrameModificarTicaje.Visible = False
    Else
        For Total = 3 To 5
            Text3(Total).Text = Trim(Text3(Total).Text)
            If Text3(Total).Text = "" Then
                MsgBox "Todos los campos requieren valor", vbExclamation
                Exit Sub
            Else
                If Not IsDate(Text3(Total).Text) Then
                    MsgBox "hora incorrecta", vbExclamation
                    Exit Sub
                End If
            End If
        Next Total
        If CDate(Text3(3).Text) > CDate(Text3(4).Text) Then
            MsgBox "Hora desde mayor hora hasta", vbExclamation
            Exit Sub
        End If
        
        cad = "Va a modificar los fichajes de los trabajadores seleccionados " & vbCrLf
        cad = cad & " para el dia : " & Text3(2).Text & vbCrLf
        cad = cad & " Hora inicio: " & Text3(3).Text & vbCrLf
        cad = cad & " Hora fin: " & Text3(4).Text & vbCrLf & vbCrLf
        cad = cad & " Hora MODIFICADA: " & Text3(5).Text & vbCrLf
        cad = cad & vbCrLf & "�Desea continuar?"
        If MsgBox(cad, vbQuestion + vbYesNoCancel) <> vbYes Then Exit Sub
        'Ahora
        Screen.MousePointer = vbHourglass
            Me.FrameModificarTicaje.Visible = False
            Label2.Caption = "Modificar ticada"
            Label3.Caption = ""
            Frame2.Visible = True
            Me.Refresh
            espera 0.2
            Set RS = New ADODB.Recordset
            cad = "select * from tmpCambioHor"
            RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            cad = "UPDATE EntradaFichajes SET HoraReal = #" & Format(Text3(5), "hh:mm") & ":00#"
            'Modificacion del 10 / Enero /2005
            cad = cad & ", Hora = #" & Format(Text3(5), "hh:mm") & ":00#"
            
            cad = cad & " WHERE HoraReal >= #" & Format(Text3(3), "hh:mm") & ":00#"
            cad = cad & " AND HoraReal <= #" & Format(Text3(4), "hh:mm") & ":00#"
            cad = cad & " AND Fecha = #" & Format(Text3(2), FormatoFecha) & "#"
            cad = cad & " AND IdTrabajador  = "

            While Not RS.EOF
            
                Label3.Caption = RS!Trabajador
                Me.Refresh
                conn.Execute cad & RS!Trabajador
                RS.MoveNext
                
            Wend
            RS.Close
            Frame2.Visible = False
            Me.Refresh
            espera 0.1
            Unload Me
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()

    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 2 Or Opcion = 3 Then
            CargarDatos
            cmdCambiar_Click
            
        Else
          If Opcion = 6 Then
            MenuEnable False
            Frame1.Enabled = True
            CargarDatos
           End If

        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    PrimeraVez = True
    Label1(0).Visible = False
    Label1(1).Visible = False
    If Opcion < 2 Then Label1(Opcion).Visible = True
    Frame2.Visible = False
    FrameNuevoH.Visible = False
    FrameModificarTicaje.Visible = False
    Me.FrameTicada.Visible = False
    'Borramos todos los datos de la tabla temporal
    If Opcion < 2 Then
        conn.Execute "Delete from tmpCambioHor"
    Else
        Text3(0).Text = Format(Fecha, "dd/mm/yyyy")
        Text3(2).Text = Text3(0).Text
    End If
    Text3(0).Enabled = Opcion <> 2
    Text3(2).Enabled = Opcion <> 3
    If Opcion < 6 Then
        ListView1.ListItems.Clear
        'Enlazamos imagenes
        ListView1.SmallIcons = ImageList1
        
    Else
        InsertarTodosTrabajadores
        ListView2.ListItems.Clear
        'Enlazamos imagenes
        ListView2.SmallIcons = ImageList1
    End If
    ListView1.Visible = Opcion < 6
    ListView2.Visible = Opcion >= 6
    imgCheck(0).Visible = Opcion = 6
    imgCheck(1).Visible = Opcion = 6
    'Caption
    Select Case Opcion
    Case 0, 6
        Me.Tag = "Cambio"
        Devolucion = "Cambio horarios"
    Case Else
        Me.Tag = "Generar"
        Devolucion = "Generacion ticajes"
    End Select
    Caption = Devolucion
    cmdCambiar.Caption = Me.Tag
    cmdCambiar.Enabled = vUsu.Nivel < 2
    Me.Tag = ""
    Devolucion = ""
End Sub

'INSERTARA LA SECCION, EL HORARIO O LA CATEGORIA
Private Sub InsertarSeccionHorario(Seccion As Integer)
Dim cad As String

    cad = " from Trabajadores,Horarios,Categorias where "
    cad = cad & " Trabajadores.idHorario = horarios.idhorario AND "
    cad = cad & " Trabajadores.idCategoria = categorias.idcategoria AND "
    Select Case Seccion
    Case 2
        cad = cad & "seccion"
    Case 3
        cad = cad & "Trabajadores.idhorario"
    Case Else
        cad = cad & "Trabajadores.idCategoria"
    End Select
    cad = cad & " = " & Devolucion
    Set RS = New ADODB.Recordset
    
    'Contador
    RS.Open "Select count(*) " & cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Total = 0
    If Not RS.EOF Then
        Total = DBLet(RS.Fields(0))
    End If
    RS.Close
    If Total = 0 Then
        MsgBox "Ningun trabajador con esos valores.", vbExclamation
        Set RS = Nothing
        Exit Sub
    End If
    
    cad = "Select idTrabajador,nomtrabajador,nomhorario,nomcategoria " & cad
    cad = cad & " order by idTrabajador"
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    cad = ""
    i = 0
    While Not RS.EOF
        If Insertar1trabajador(RS.Fields(0)) Then _
            A�adeListview1 RS.Fields(0), RS.Fields(1), RS.Fields(2), RS.Fields(3)
        i = i + 1
        Label3.Caption = i & " de " & Total
        Label3.Refresh
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
End Sub


Private Sub InsertarTrabajador()
Dim cad As String

    cad = "Select idTrabajador,nomtrabajador,nomhorario,nomcategoria from Trabajadores,Horarios,Categorias where "
    cad = cad & " Trabajadores.idHorario = horarios.idhorario AND "
    cad = cad & " Trabajadores.idCategoria = categorias.idCategoria AND "
    cad = cad & " idTrabajador = " & Devolucion
    Set RS = New ADODB.Recordset
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Insertar1trabajador(RS.Fields(0)) Then _
            A�adeListview1 RS.Fields(0), RS.Fields(1), RS.Fields(2), RS.Fields(3)

        RS.MoveNext
    End If
    RS.Close
    Set RS = Nothing
End Sub



Private Function Insertar1trabajador(Id As Long) As Boolean
    On Error Resume Next
    Insertar1trabajador = False
    conn.Execute "INSERT INTO tmpCambioHor(Trabajador) VALUES (" & Id & ");"
    If Err.Number Then
        Err.Clear
    Else
        Insertar1trabajador = True
    End If
End Function




Private Sub HacerCambio()
    
    If Opcion = 6 Then
    
        For i = 1 To ListView2.ListItems.Count
            Label3.Caption = i & " de " & ListView2.ListItems.Count
            Label3.Refresh
            If ListView2.ListItems(i).Checked Then
                If UpdatearTrabajador(ListView2.ListItems(i).Tag) Then ListView2.ListItems(i).SubItems(1) = Text2.Text
            End If
        Next i
    Else
    
    
        For i = 1 To ListView1.ListItems.Count
              Label3.Caption = i & " de " & ListView1.ListItems.Count
              Label3.Refresh
              If UpdatearTrabajador(ListView1.ListItems(i).Tag) Then _
                  ListView1.ListItems(i).SubItems(1) = Text2.Text
        Next i
    End If
    
End Sub




Private Sub HacerEliminar()
'Dim Cad As String
Dim J As Integer

    Total = 0
    For i = 1 To ListView1.ListItems.Count
        If ListView1.ListItems(i).Selected Then Total = Total + 1
    Next i
    
    
    
    
    If Total = 0 Then
        MsgBox "Seleccione los trabajadores que desee eliminar de la lista", vbExclamation
    Else
        Label2.Caption = "Eliminar"
        Label3.Caption = ""
        Frame2.Visible = True
        Me.Refresh
        J = 0
        For i = ListView1.ListItems.Count To 1 Step -1
            If ListView1.ListItems(i).Selected Then
                J = J + 1
                Label3.Caption = J & " de " & Total
                Label3.Refresh
                If BorrarTrabajador(ListView1.ListItems(i).Tag) Then _
                    ListView1.ListItems.Remove ListView1.ListItems(i).Index
            End If
        Next i
    End If
    Frame2.Visible = False
End Sub


Private Function BorrarTrabajador(Id As String) As Boolean
    On Error Resume Next
    conn.Execute "Delete from tmpCambioHor where trabajador = " & Id & ";"
    If Err.Number <> 0 Then
        Err.Clear
        BorrarTrabajador = False
    Else
        BorrarTrabajador = True
    End If
End Function

Private Function UpdatearTrabajador(Id As String) As Boolean
    On Error Resume Next
    conn.Execute "Update Trabajadores set idHorario=" & Text1.Text & " where idtrabajador = " & Id & ";"
    If Err.Number <> 0 Then
        Err.Clear
        UpdatearTrabajador = False
    Else
        UpdatearTrabajador = True
    End If
End Function


Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    Devolucion = vCodigo
    If vIndice = 100 Then
        Text1.Text = vCodigo
        Text2.Text = vCadena
    End If
End Sub

Private Sub A�adeListview1(IdTra As Long, NomTra As String, nomHora As String, nomCat As String)
Dim itmX As ListItem
    If Opcion = 6 Then
        Set itmX = ListView2.ListItems.Add(, "C" & CStr(IdTra), NomTra)
    Else
        Set itmX = ListView1.ListItems.Add(, "C" & CStr(IdTra), NomTra)
    End If
    itmX.Tag = IdTra
    itmX.SubItems(1) = nomHora
    itmX.SubItems(2) = nomCat
    itmX.SmallIcon = 1
End Sub


Private Sub frmBu_Selecionado(CadenaDevuelta As String)
    Devolucion = CadenaDevuelta
End Sub

Private Sub Image1_Click()
        vIndice = 100
        Devolucion = ""
        Set frmB = New frmBusca
        frmB.Tabla = "Horarios"
        frmB.CampoBusqueda = "NomHorario"
        frmB.CampoCodigo = "IdHorario"
        frmB.TipoDatos = 3
        frmB.Titulo = "HORARIOS"
        frmB.MostrarDeSalida = True
        frmB.Show vbModal
        Set frmB = Nothing
End Sub




Private Sub imgCheck_Click(Index As Integer)
Dim B As Boolean
    B = Index = 0
    For i = 1 To ListView2.ListItems.Count
        ListView2.ListItems(i).Checked = B
    Next i
End Sub

Private Sub mn1TrabajaCod_Click()
    AccionesBuscaGrid 0
End Sub

Private Sub mn1Trabajador_Click()
    Acciones 0
End Sub

Private Sub mnCategoria_Click()
    Acciones 4
End Sub

Private Sub mnCodigoTarjeta_Click()
    Acciones 11
End Sub

Private Sub mnEliminarUno_Click()
    Acciones 1
End Sub

Private Sub mnHorario_Click()
    Acciones 3
End Sub

Private Sub mnLimpiar_Click()
    Acciones 10
End Sub

Private Sub mnSeccion_Click()
    Acciones 2
End Sub

Private Sub Text1_LostFocus()
Dim cH As CHorarios

    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then
        Text2.Text = ""
        Exit Sub
    End If
    
    
    If Not IsNumeric(Text1.Text) Then
        MsgBox "El campo debe ser num�rico", vbExclamation
        Text1.Text = ""
        Text2.Text = ""
        Exit Sub
    End If
    
    'Ponemos el horario
    Set cH = New CHorarios
    If cH.Leer(CInt(Text1.Text), Now) = 0 Then
        Text1.Text = cH.IdHorario
        Text2.Text = cH.NomHorario
    Else
        Text1.Text = ""
        Text2.Text = ""
    End If
    Set cH = Nothing
End Sub

Private Sub MenuEnable(B As Boolean)
    mnAgregar.Enabled = B
    mnEliminar.Enabled = B
    Frame1.Enabled = B
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Private Sub Text3_LostFocus(Index As Integer)
    Text3(Index).Text = Trim(Text3(Index).Text)
    If Text3(Index).Text = "" Then Exit Sub
    
    If Index = 0 Then
        'Fecha
        If Not EsFechaOK(Text3(Index)) Then
            MsgBox "Debe ser una fecha:" & Text3(Index).Text, vbExclamation
            Text3(Index).Text = ""
            Exit Sub
        End If
        
    Else
        Devolucion = TransformaPuntosHoras(Text3(Index).Text)
        If Not IsDate(Devolucion) Then
            MsgBox "Se esperaba una hora: " & Devolucion, vbExclamation
            Text3(Index).Text = ""
            Exit Sub
        End If
        Text3(Index).Text = Format(Text3(Index).Text, "hh:mm")
    End If
        
        
        
        
End Sub


'Para cada trabajador, veremos si ha trabajado ese dia, y
' en ENTRADATICAJES vere si hay
'            -SI  -> Genero una nueva
'            -NO  -> A�ado errores (NO esta calaro todavia)

Private Sub HacerTicada()
Dim JJ As Integer
    Set RS = New ADODB.Recordset
    
    'Veremos el contador
    Total = 0
    RS.Open "Select max(SECUENCIA) from entradafichajes", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        Total = DBLet(RS.Fields(0), "N")
    End If
    RS.Close
    
    Total = Total + 1   'la siguiente
    
    'Ponemos los tag
    Text3(0).Tag = "#" & Format(Text3(0).Text, "yyyy/mm/dd") & "#"
    Text3(1).Tag = "#" & Format(Text3(1).Text, "hh:mm:ss") & "#"
    'i ->errores
    i = 0
    For JJ = 1 To ListView1.ListItems.Count
          Label3.Caption = JJ & " de " & ListView1.ListItems.Count
          Label3.Refresh
          If TicadaTrabajador(ListView1.ListItems(JJ).Tag) = False Then i = i + 1

    Next JJ
    Set RS = Nothing
    Text3(0).Tag = ""
    Text3(1).Tag = ""
End Sub


Private Function TicadaTrabajador(Trab As String) As Boolean
Dim Tiene As Boolean
Dim cad As String

    On Error GoTo ETicadaTrabajador
    
    TicadaTrabajador = False
    'Para el trabajador veo si tiene ticadas
    cad = "Select secuencia from Entradafichajes where  idTrabajador=" & Trab & " AND fecha = " & Text3(0).Tag
    RS.Open cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If RS.EOF Then
        cad = ""
    End If
    RS.Close
   
    'Si k tiene ticajes
    'Insertamos el nuestro
    cad = "insert into entradafichajes (secuencia,idtrabajador,fecha,hora,idinci,horareal) VALUES ("
    cad = cad & Total & "," & Trab & "," & Text3(0).Tag & "," & Text3(1).Tag & ",0," & Text3(1).Tag & ")"
    conn.Execute cad
    Total = Total + 1
    TicadaTrabajador = True
    Exit Function
ETicadaTrabajador:
    MuestraError Err.Number, Err.Description
End Function


Private Function PedirDesdeHastaTarjeta()
Dim Fin As String
Dim Inicio As String
    
    Inicio = InputBox("Introduzca el codigo de tarjeta INICIAL")
    Fin = InputBox("Introduzca el codigo de tarjeta FINAL")
    Inicio = Trim(Inicio): Fin = Trim(Fin)
    If Inicio = "" And Fin = "" Then
        MsgBox "Debe poner el codigo incial  y/o  el codigo final", vbExclamation
        Exit Function
    End If
    If Inicio <> "" Then
        If Not IsNumeric(Inicio) Then
            MsgBox "Codigo de tarjeta inicial debe ser num�rico", vbExclamation
            Exit Function
        End If
    End If
    If Fin <> "" Then
        If Not IsNumeric(Fin) Then
            MsgBox "Codigo de tarjeta final debe ser num�rico", vbExclamation
            Exit Function
        End If
    End If
    
    
    'Llegados aqui insertaremos
    Devolucion = "Select idTrabajador,nomtrabajador,nomhorario,nomcategoria from Trabajadores,Horarios,Categorias where "
    Devolucion = Devolucion & " Trabajadores.idHorario = horarios.idhorario AND "
    Devolucion = Devolucion & " Trabajadores.idCategoria = categorias.idCategoria "
    If Inicio <> "" Then
        Devolucion = Devolucion & " AND Trabajadores.numTarjeta >= '" & Inicio & "'"
    End If
    If Inicio <> "" Then
        Devolucion = Devolucion & " AND Trabajadores.numTarjeta <= '" & Fin & "'"
    End If

    Set RS = New ADODB.Recordset
    RS.Open Devolucion, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        If Insertar1trabajador(RS.Fields(0)) Then _
            A�adeListview1 RS.Fields(0), RS.Fields(1), RS.Fields(2), RS.Fields(3)

        RS.MoveNext
    Wend
    RS.Close
    Devolucion = ""
    Set RS = Nothing

        
End Function





Private Function CargarDatos()

    
    
    'Llegados aqui insertaremos
    Devolucion = "Select tmpCambioHor.Trabajador,nomtrabajador,nomhorario,nomcategoria from tmpCambioHor,Trabajadores,Horarios,Categorias where "
    Devolucion = Devolucion & " Trabajadores.idHorario = horarios.idhorario AND "
    Devolucion = Devolucion & " Trabajadores.idCategoria = categorias.idCategoria AND "
    Devolucion = Devolucion & " tmpCambioHor.trabajador =  trabajadores.idtrabajador"
    
    Set RS = New ADODB.Recordset
    RS.Open Devolucion, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        
        A�adeListview1 RS.Fields(0), RS.Fields(1), RS.Fields(2), RS.Fields(3)

        RS.MoveNext
    Wend
    RS.Close
    Devolucion = ""
    Set RS = Nothing

        
End Function





Private Sub InsertarTodosTrabajadores()
    conn.Execute "Delete from tmpCambioHor"
    Devolucion = "INSERT INTO tmpCambioHor(trabajador) SELECT idtrabajador FROM trabajadores where  "
    Devolucion = Devolucion & " (Trabajadores.FecBaja Is Null Or fecbaja < #" & Format(Now, "yyyy/mm/dd")
    Devolucion = Devolucion & "#) AND Trabajadores.FecAlta<#" & Format(Now, "yyyy/mm/dd") & "#"
    conn.Execute Devolucion
    Devolucion = ""
End Sub
