VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUnix 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enviar Multibase"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6105
   Icon            =   "frmUnix.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Traspaso"
      TabPicture(0)   =   "frmUnix.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Image2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "ProgressBar1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command2(0)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtEmpresa(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtEmpresa(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Configuracion"
      TabPicture(1)   =   "frmUnix.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Command2(1)"
      Tab(1).Control(4)=   "txtConfig(0)"
      Tab(1).Control(5)=   "txtConfig(1)"
      Tab(1).Control(6)=   "Command3"
      Tab(1).Control(7)=   "Check1"
      Tab(1).Control(8)=   "Check2"
      Tab(1).ControlCount=   9
      Begin VB.CheckBox Check2 
         Caption         =   "Enviar campo Tipo Horas"
         Height          =   255
         Left            =   -71700
         TabIndex        =   20
         Top             =   1740
         Width           =   2175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Escribir horas en decimal"
         Height          =   255
         Left            =   -74580
         TabIndex        =   12
         Top             =   1740
         Width           =   2175
      End
      Begin VB.TextBox txtEmpresa 
         Enabled         =   0   'False
         Height          =   285
         Index           =   1
         Left            =   2160
         TabIndex        =   19
         Top             =   600
         Width           =   3255
      End
      Begin VB.TextBox txtEmpresa 
         Height          =   285
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   600
         Width           =   555
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Guardar config."
         Height          =   315
         Left            =   -73860
         TabIndex        =   13
         Top             =   2700
         Width           =   1395
      End
      Begin VB.TextBox txtConfig 
         Height          =   285
         Index           =   1
         Left            =   -73680
         TabIndex        =   11
         Text            =   "Text2"
         Top             =   1200
         Width           =   4035
      End
      Begin VB.TextBox txtConfig 
         Height          =   285
         Index           =   0
         Left            =   -73680
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   720
         Width           =   4035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   315
         Index           =   1
         Left            =   -71940
         TabIndex        =   15
         Top             =   2700
         Width           =   1275
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Salir"
         Height          =   315
         Index           =   0
         Left            =   4140
         TabIndex        =   4
         Top             =   2880
         Width           =   1275
      End
      Begin VB.Frame Frame1 
         Caption         =   "Fecha      "
         Height          =   915
         Left            =   420
         TabIndex        =   7
         Top             =   1200
         Width           =   5055
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   1
            Left            =   3420
            TabIndex        =   2
            Text            =   "Text1"
            Top             =   420
            Width           =   1215
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Index           =   0
            Left            =   1200
            TabIndex        =   1
            Text            =   "Text1"
            Top             =   420
            Width           =   1215
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   1
            Left            =   3120
            Picture         =   "frmUnix.frx":0342
            Top             =   420
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Index           =   0
            Left            =   900
            Picture         =   "frmUnix.frx":0444
            Top             =   420
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fin"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   9
            Top             =   480
            Width           =   210
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Inicio"
            Height          =   195
            Index           =   0
            Left            =   480
            TabIndex        =   8
            Top             =   480
            Width           =   375
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Iniciar"
         Height          =   315
         Left            =   2760
         TabIndex        =   3
         Top             =   2895
         Width           =   1275
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   255
         Left            =   420
         TabIndex        =   6
         Top             =   2460
         Width           =   5055
         _ExtentX        =   8916
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1140
         Picture         =   "frmUnix.frx":0546
         Top             =   660
         Width           =   240
      End
      Begin VB.Label Label5 
         Caption         =   "Empresa:"
         Height          =   255
         Left            =   420
         TabIndex        =   18
         Top             =   660
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Si modifica algún valor no olvide guardar la configuración."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   -74640
         TabIndex        =   17
         Top             =   2220
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre"
         Height          =   195
         Left            =   -74580
         TabIndex        =   16
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "PATH "
         Height          =   255
         Left            =   -74580
         TabIndex        =   14
         Top             =   780
         Width           =   675
      End
   End
End
Attribute VB_Name = "frmUnix"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Private Sub Command1_Click()
Dim RC As Byte
Dim Cad As String

'Inciar creación del fichero
'---------------------------------------------------------
'Comprobamos las fechas del traspaso
If Not DatosOk Then Exit Sub
'luego comprobamos que todos los marcajes estan correctos
RC = ComprobarMarcajesCorrectos(CDate(Text1(0).Text), CDate(Text1(1).Text), False)
If RC > 0 Then
    Cad = "En el periodo de fechas indicado( " & Text1(0).Text & " - " & Text1(1).Text & ")" & vbCrLf
    Cad = Cad & " existen marcajes incorrectos." & vbCrLf
    Cad = Cad & " ¿Desea continuar de todas formas?"
    RC = MsgBox(Cad, vbQuestion + vbYesNoCancel)
End If
If RC = vbNo Then Exit Sub
Screen.MousePointer = vbHourglass
'Generemos el archivo de texto y lo colocamos donde nos digan
RC = GeneraArchivo
If RC = 0 Then MsgBox "Fichero de traspaso realizado con exito", vbInformation
Screen.MousePointer = vbDefault
End Sub





Private Function DatosOk() As Boolean
DatosOk = False
'Incio
'La empresa
If txtEmpresa(0).Text = "" Then
    MsgBox "El codigo de empresa debe de ser numérico.", vbExclamation
    Exit Function
End If
If Not IsNumeric(txtEmpresa(0).Text) Then
    MsgBox "El codigo de empresa debe de ser numérico.", vbExclamation
    Exit Function
End If
If Trim(txtEmpresa(0).Text) = "-1" Then
    MsgBox "La empresa seleccionada no existe.", vbExclamation
    Exit Function
End If


If Text1(0).Text = "" Then
    MsgBox "Fecha inicio en blanco.", vbExclamation
    Exit Function
End If
If Not IsDate(Text1(0).Text) Then
    MsgBox "Fecha incio incorrecta", vbExclamation
    Exit Function
End If
'fin
If Text1(1).Text = "" Then
    MsgBox "Fecha fin en blanco.", vbExclamation
    Exit Function
End If
If Not IsDate(Text1(1).Text) Then
    MsgBox "Fecha fin incorrecta", vbExclamation
    Exit Function
End If
'Comprobamos que incio>fin
If CDate(Text1(0).Text) > CDate(Text1(1).Text) Then
    MsgBox "Fecha incio es mayor que la fecha final.", vbExclamation
    Exit Function
End If

'Los datos del archivo destino
txtConfig(0).Text = Trim(txtConfig(0).Text)
txtConfig(1).Text = Trim(txtConfig(1).Text)
If txtConfig(0).Text = "" Then
    MsgBox "La capeta donde se creará el archivo no puede estar vacia.", vbExclamation
    Exit Function
End If

If Dir(txtConfig(0).Text, vbDirectory) = "" Then
    MsgBox "La capeta donde se creará el archivo no existe.", vbExclamation
    Exit Function
End If


'El nombre del archivo
If txtConfig(1).Text = "" Then
    MsgBox "El nombre del archivo no puede estar vacio.", vbExclamation
    Exit Function
End If

'Datos bien
DatosOk = True
End Function

Private Sub Command2_Click(Index As Integer)
Unload Me
End Sub

Private Sub Command3_Click()
Dim Cad As String
Dim NF As Integer
On Error GoTo ErrorGuardar
Cad = Dir(txtConfig(0).Text, vbDirectory)
If Cad = "" Or Trim(txtConfig(1).Text) = "" Then
    MsgBox "El directorio " & txtConfig(0).Text & " no existe " & _
        " y/o el nombre de archivo no puede estar en blanco.", vbExclamation
    Else
        NF = FreeFile
        Open App.Path & "\TrasConf.cfg" For Output As #NF
        Cad = Trim(txtConfig(0).Text) & "|" 'path
        Cad = Cad & Trim(txtConfig(1).Text) & "|"  'Nombre
        Cad = Cad & Check1.Value & "|"  'Dec/sexagesimal
        Cad = Cad & Check2.Value & "|"  'Envio tipo de horas
                        'N: normales( si no son extra)  F: Festivas
        Print #NF, Cad
        Close #NF
        MsgBox "La configuración ha sido guardada correctamente.", vbInformation
End If
ErrorGuardar:
    If Err.Number <> 0 Then MuestraError Err.Number
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Cad As String

LeerConfiguracion
Text1(0).Tag = "ESTE"
Text1(0).Text = Format(Now - 30, "dd/mm/yyyy")
Text1(1).Text = Format(Now, "dd/mm/yyyy")

'Como opcion vamos a ver si podemos cargar la primera de las empresas
i = 1
Cad = ""
Do
    If i < 3 Then
        Cad = DevuelveNombreEmpresa(CLng(i))
        If Cad <> "" Then
            Me.txtEmpresa(0).Text = i
            txtEmpresa(1).Text = Cad
        End If
        i = i + 1
        Else
            Cad = "SAL"
    End If
Loop Until Cad <> ""
SSTab1.Tab = 0
Screen.MousePointer = vbDefault
End Sub


Private Sub LeerConfiguracion()
Dim Cad As String
Dim NF As Integer
'Dim ini As Integer
Dim i As Integer
On Error GoTo ErrorLeerConf
txtConfig(0).Text = App.Path
txtConfig(1).Text = "daripres.txt"
Check1.Value = 0
Check2.Value = 0
Cad = Dir(App.Path & "\TrasConf.cfg")
If Cad <> "" Then
    NF = FreeFile
    Open App.Path & "\TrasConf.cfg" For Input As #NF
    Input #NF, Cad
    Close #NF
    'Leemos los valores
    If Trim(Cad) <> "" Then
        For NF = 0 To 1
            i = InStr(1, Cad, "|")
            If i > 0 Then
                txtConfig(NF).Text = Mid(Cad, 1, i - 1)
                Cad = Mid(Cad, i + 1)
            End If
        Next NF
        'Para el check 1
        NF = InStr(1, Cad, "|")
        If NF > 0 Then
            Check1.Tag = Mid(Cad, 1, NF - 1)
            Cad = Mid(Cad, NF + 1)
        End If
        NF = Val(Check1.Tag)
        Check1.Tag = ""
        If NF = 1 Then Check1.Value = NF
        'Para el check 2
        NF = InStr(1, Cad, "|")
        If NF > 0 Then
            Check2.Tag = Mid(Cad, 1, NF - 1)
            Cad = Mid(Cad, NF + 1)
        End If
        NF = Val(Check2.Tag)
        Check2.Tag = ""
        If NF = 1 Then Check2.Value = NF
    End If
End If
Exit Sub
ErrorLeerConf:
    If Err.Number <> 0 Then MuestraError Err.Number
End Sub




Private Function GeneraArchivo() As Byte
Dim NF As Integer
Dim Rs As ADODB.Recordset
Dim Cad As String
Dim IncHoraExtra As Integer
Dim Horas As Date
Dim vE As CEmpresas
Dim AntFech As Date
Dim AntHorario As Integer
Dim LetraTipoHorario As String


    On Error GoTo ErrGenerarArchivo
    GeneraArchivo = 1
    
    IncHoraExtra = -1
    Set vE = New CEmpresas
    If vE.Leer(CLng(Val(txtEmpresa(0).Text))) = 1 Then
        MsgBox "Error leyendo los datos de la empresa nª: " & txtEmpresa(0).Text
        Else
            IncHoraExtra = vE.IncHoraExtra
    End If
    Set vE = Nothing
    If IncHoraExtra < 0 Then Exit Function
    
    Set Rs = New ADODB.Recordset
    Rs.CursorType = adOpenKeyset
    Rs.LockType = adLockOptimistic
    'SQL. Marcajes incorrectos entre las dos fechas
    Cad = "Select Trabajadores.idTrabajador, Marcajes.HorasTrabajadas,"
    Cad = Cad & " Marcajes.Fecha , Marcajes.HorasIncid , Marcajes.IncFinal, Trabajadores.IdHorario"
    Cad = Cad & " From Secciones, Trabajadores, Marcajes"
    Cad = Cad & " WHERE  Secciones.IdSeccion = Trabajadores.Seccion "
    Cad = Cad & " AND Trabajadores.idTrabajador = Marcajes.idTrabajador And Secciones.Nominas = True"
    Cad = Cad & " AND Trabajadores.IdEmpresa =" & txtEmpresa(0).Text
    Cad = Cad & " AND Fecha>=#" & Format(Text1(0).Text, "yyyy/mm/dd") & "#"
    Cad = Cad & " AND Fecha<=#" & Format(Text1(1).Text, "yyyy/mm/dd") & "#"
    Cad = Cad & " ORDER BY Trabajadores.IdHorario,Fecha"
    Rs.Open Cad, Conn, , , adCmdText
    If Rs.EOF Then
        Cad = "Ningún dato para traspasar entre esas fechas: " & vbCrLf & _
            "Incio: " & Format(Text1(0).Text, "dd/mm/yyyy") & vbCrLf & _
            "Fin: " & Format(Text1(1).Text, "dd/mm/yyyy") & vbCrLf
        Cad = Cad & "Para la empresa: " & vbCrLf & "          " & _
            "         " & txtEmpresa(0).Text & " - " & txtEmpresa(1).Text
        MsgBox Cad, vbExclamation
        Rs.Close
        Exit Function
        'ELSE
        Else
           
            'Valores por defecto para los horarios
            AntFech = "31/12/1900"
            AntHorario = -1
            
            If Rs.RecordCount > 32000 Then
                ProgressBar1.Max = 32000
                Else
                    ProgressBar1.Max = Rs.RecordCount
            End If
            NF = FreeFile
            Open txtConfig(0).Text & "\" & txtConfig(1).Text For Output As #NF
            While Not Rs.EOF
                'Por si acaso enviamos la letra de
                If Check2.Value Then
                    'Enviamos datos horario
                    If AntHorario <> Rs!IdHorario Then
                        LetraTipoHorario = DevuelveHorario(Rs!Fecha, Rs!IdHorario)
                        AntHorario = Rs!IdHorario
                        AntFech = Rs!Fecha
                        Else
                            If AntFech <> Rs!Fecha Then
                                LetraTipoHorario = DevuelveHorario(Rs!Fecha, Rs!IdHorario)
                                AntFech = Rs!Fecha
                            End If
                    End If
                End If
                Cad = "A|"
                Cad = Cad & Rs!idTrabajador & "|"
                Cad = Cad & Format(Rs!Fecha, "dd/mm/yyyy") & "|"
                '-----------------------------------------------
                'Esto es para no reflejar los retrasos como horas extras
                'Es decir, si ha llegado una hora tarde y tenia que trabajar 8
                'no ponemos :  Horas trabjad: 7  Horas extras -1
                '     si no que en horas extras ponemos directamente 0
                If Rs!IncFinal <> IncHoraExtra Then
                    Cad = Cad & PonHoras(Rs!HorasTrabajadas) & "|"
                    Cad = Cad & PonHoras(0) & "|"
                    Else
                        Cad = Cad & PonHoras(Rs!HorasTrabajadas - Rs!HorasIncid) & "|"
                        Cad = Cad & PonHoras(Rs!HorasIncid) & "|"
                End If
                '-------------------------------------------------
                Cad = Cad & "0|"
                'Si pide que pongamos que tipo de horas extras son entonces
                If Check2.Value Then
                    Cad = Cad & LetraTipoHorario & "|"
                End If
                Print #NF, Cad
                Rs.MoveNext
                If ProgressBar1.Max > ProgressBar1.Value Then _
                    ProgressBar1.Value = ProgressBar1.Value + 1
            Wend
            Close #NF
    End If
    Rs.Close
    Set Rs = Nothing
    GeneraArchivo = 0
    Exit Function
ErrGenerarArchivo:
    MsgBox "Error generando fichero traspaso. " & vbCrLf & "Número: " & Err.Number & _
        vbCrLf & "Descripcion: " & Err.Description, vbExclamation
End Function


Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
txtEmpresa(0).Text = vCodigo
txtEmpresa(1).Text = vCadena
End Sub

Private Sub frmC_Selec(vFecha As Date)
If Text1(0).Tag <> "" Then
    Text1(0).Text = Format(vFecha, "dd/mm/yyyy")
    Else
        Text1(1).Text = Format(vFecha, "dd/mm/yyyy")
End If
End Sub

Private Sub Image1_Click(Index As Integer)
Text1(1).Tag = Now
If Index = 0 Then
    Text1(0).Tag = "ESTE"
    If IsDate(Text1(0).Text) Then Text1(1).Tag = Text1(0).Text
    Else
        Text1(0).Tag = ""
        If IsDate(Text1(1).Text) Then Text1(1).Tag = Text1(1).Text
End If
Set frmC = New frmCal
frmC.Fecha = Text1(1).Tag
frmC.Show vbModal
Set frmC = Nothing
End Sub

Private Sub Image2_Click()
    Set frmB = New frmBusca
    frmB.Tabla = "Empresas"
    frmB.CampoBusqueda = "NomEmpresa"
    frmB.CampoCodigo = "IdEmpresa"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPRESAS"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub Text1_GotFocus(Index As Integer)
Text1(Index).SelStart = 0
Text1(Index).SelLength = Len(Text1(Index).Text)
End Sub

Private Sub Text1_LostFocus(Index As Integer)
If Trim(Text1(Index).Text) = "" Then Exit Sub
If IsDate(Text1(Index).Text) Then Text1(Index).Text = Format(Text1(Index).Text, "dd/mm/yyyy")
End Sub

Private Sub txtEmpresa_GotFocus(Index As Integer)
txtEmpresa(Index).SelStart = 0
txtEmpresa(Index).SelLength = Len(txtEmpresa(Index).Text)
End Sub

Private Sub txtEmpresa_LostFocus(Index As Integer)
If Index = 0 Then
    If txtEmpresa(0).Text = "" Then
        txtEmpresa(1).Text = ""
        Exit Sub
    End If
    If Not IsNumeric(txtEmpresa(0).Text) Then
        txtEmpresa(0).Text = "-1"
        txtEmpresa(1).Text = "Empresa incorrecta"
        Exit Sub
    End If
    txtEmpresa(1).Text = DevuelveNombreEmpresa(CLng(txtEmpresa(0).Text))
    If txtEmpresa(1).Text = "" Then
        txtEmpresa(0).Text = "-1"
        txtEmpresa(1).Text = "Empresa incorrecta"
    End If
End If
End Sub


Private Function PonHoras(Horas As Single) As String
Dim H As Date
If Check1.Value = 1 Then
    'DECIMAL    DECIMAL    DECIMAL    DECIMAL    DECIMAL
    PonHoras = Format(Horas, "0.00")
    Else
        H = DevuelveHora(Horas)
        PonHoras = Format(H, "hh:mm")
End If
End Function


Private Function DevuelveHorario(Fecha As Date, IdHor As Integer) As String
Dim VH As CHorarios

Set VH = New CHorarios
DevuelveHorario = "N" 'Por defecto ponemos NORMAL
If VH.Leer(IdHor, Fecha) = 0 Then
    If VH.EsDiaFestivo Then
        DevuelveHorario = "F"
        Else
            If mConfig.SabadosHorasFestivas Then
                If Weekday(Fecha) = 7 Then DevuelveHorario = "F"
            End If
    End If
End If
Set VH = Nothing
End Function
