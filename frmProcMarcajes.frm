VERSION 5.00
Object = "{FE0065C0-1B7B-11CF-9D53-00AA003C9CB6}#1.1#0"; "COMCT232.OCX"
Begin VB.Form frmProcMarcajes2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Procesar marcajes"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmProcMarcajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdComprobar 
      Caption         =   "Comprobar"
      Height          =   315
      Left            =   180
      TabIndex        =   11
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1155
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   5415
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   1
         Left            =   3840
         TabIndex        =   1
         Top             =   660
         Width           =   1395
      End
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   0
         Left            =   1200
         TabIndex        =   0
         Top             =   660
         Width           =   1395
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   1
         Left            =   3540
         Picture         =   "frmProcMarcajes.frx":030A
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmProcMarcajes.frx":040C
         Top             =   720
         Width           =   240
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Fin"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3180
         TabIndex        =   8
         Top             =   720
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Inicio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Fechas entre las cuales se procesarán los marcajes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   5175
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   3840
      TabIndex        =   3
      Top             =   2640
      Width           =   1515
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Iniciar"
      Height          =   315
      Left            =   2400
      TabIndex        =   2
      Top             =   2640
      Width           =   1275
   End
   Begin ComCtl2.Animation Animation1 
      Height          =   915
      Left            =   480
      TabIndex        =   4
      Top             =   2040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   1614
      _Version        =   327681
      BackColor       =   12632256
      FullWidth       =   305
      FullHeight      =   61
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   10
      Top             =   1620
      Width           =   5175
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   180
      TabIndex        =   9
      Top             =   1260
      Width           =   5175
   End
End
Attribute VB_Name = "frmProcMarcajes2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Opcion As Byte

    '0.- Cualquiera menos....
    '1.- Cuando añadimos un marcaje de presencia A MANO
    '   entonces, tenemos que volver a recalcular las tareas para ese
    '   currante. En ImpFechaIni tendremos la fecha del recalculo, ya en #fecha#
    
'----------  Para los recalculos. IMPORTANTE: lista ordenada por codigo
Public ListaTrabajadores As String


Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Public ContadorSecuencia As Long
'Public EsImportarFichero As Boolean
'Public EsProcesarMarcajes As Boolean


Dim PrimeraVez As Boolean
Dim Vector() As String

Private Sub cmdAceptar_Click()
'Dim Valor As Byte
Dim CADENA As String
Dim OK As Integer

    'Coprobamos las fechas
    If Not DatosOk Then Exit Sub
    Screen.MousePointer = vbHourglass
    cmdAceptar.Visible = False
    cmdSalir.Visible = False
    cmdComprobar.Visible = False
    Me.Refresh
    Me.Animation1.Open App.Path & "\ICONOS\FILEDELR.AVI"
    Me.Animation1.Play


    'Si trabajamos con el modulo KIMALDi entonces
    'Modificacion del 22 Octubre 04
    'Los datos de entradafichejes YA deben haber sido creados
    'Para ello, todo el modulo: GeneraMarcajesKimaldi
    'Sera copiado a procesar marcajes
    'If mConfig.Kimaldi Then GeneraMarcajesKimaldi
  
  
    'Ahora hago la rectificacion de marcajes
    conn.BeginTrans
    OK = RectificacionDeMarcajes
    If OK = 0 Then
        conn.CommitTrans
    Else
        conn.RollbackTrans
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
  
  
    'Generamos los marcajes
    MostrarErrores = GeneracionMarcajes
    
    'Para que no crezca el numero de secuencia en entrada fichero
    'renumeramos la secuencia
    ReOrdenarEntradaFichajes

    
    
    
    'Si es de produccion generamos las tareas
    If mConfig.Kimaldi And MiEmpresa.QueEmpresa <> 1 Then
        HazProduccion
        'Eliminamos las entradas k pudieran haber kedado sueltas
        EliminaEntradas
    End If
        
Fin:
    'Paramos el avi
    Me.Animation1.Stop
    Me.Animation1.Close
    'Restauramos lo del avi
    'cmdAceptar.Visible = True
    cmdSalir.Visible = True
    cmdComprobar.Visible = True
    Me.Refresh
    ComprobarMarcajesPendientes
    Label11.Caption = "Importación finalizada."
    Label11.Refresh
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdComprobar_Click()
Dim SQL As String
Dim RS As ADODB.Recordset

    Me.Tag = ""
    
    Set RS = New ADODB.Recordset
    
    'Comprobamos lo ultimo en marcajes
    SQL = "Select max(fecha) from marcajes"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Me.Tag = Format(RS.Fields(0), "dd/mm/yyyy")
    End If
    RS.Close
    
    If Me.Tag <> "" Then
        Me.Tag = "Marcajes.       Fecha: " & Me.Tag
    End If
    
    If mConfig.Kimaldi Then
        SQL = "Select max(fecha) from TareasRealizadas"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        SQL = ""
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then SQL = Format(RS.Fields(0), "dd/mm/yyyy")
        End If
        RS.Close
        If SQL <> "" Then
            SQL = "Tareas            Fecha: " & SQL
            If Me.Tag <> "" Then Me.Tag = Me.Tag & vbCrLf
            Me.Tag = Me.Tag & SQL
        End If
    End If
    
    
    If Me.Tag <> "" Then _
        MsgBox "Datos devueltos." & vbCrLf & vbCrLf & Me.Tag, vbInformation
    Set RS = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If Opcion = 1 Then
            If FechaHazProduccion = "" Then PrimeraVezDeFechaHazProduccion
            
            'Estamos recalculando las tareas para un trabajador en particular
            HazProduccion
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub


''''''''''
''''''''''Public Function ProcesaFichero() As Byte
'''''''''''---------------------
'''''''''''Valores que devuelve la function
''''''''''' 0.- Todo correcto
''''''''''' 1.- No existe el fichero o esta vacio
''''''''''' 2.- Algun fallo
''''''''''Dim Cad As String
''''''''''Dim NombreFichero As String
''''''''''Dim NF As Integer
''''''''''Dim Errores As Byte
''''''''''
''''''''''On Error GoTo ErrProcesaFichero
''''''''''
'''''''''''------------------------------------------------------------------
'''''''''''------------------------------------------------------------------
'''''''''''Para saber que tablas auxiliares utilizara y como se procesaran
'''''''''''las lineas
'''''''''''Estas lineas tendran que ser parametrizables puesto que son excluyentes dos a dos
''''''''''
''''''''''
''''''''''
'''''''''''En lineas generales TIPOALZICOOP = True significa k es un control de produccion
''''''''''TIPOALZICOOP = Not mConfig.TCP3
''''''''''NombreFichero = mConfig.DirMarcajes & "\" & mConfig.NomFich
'''''''''''TIPOALZICOOP = True
''''''''''
''''''''''
''''''''''
'''''''''''Obviamente a partir de la fecha obtendremos el nombre
'''''''''''del fichero. Sus datos seran almacenados en la
''''''''''' tabla: controlfichajes
''''''''''If Dir(NombreFichero) = "" Then
''''''''''    MsgBox "El fichero no esta presente " & _
''''''''''        " o ha sido eliminado." & vbCrLf & "Ruta: " & NombreFichero, vbCritical
''''''''''    ProcesaFichero = 1
''''''''''    Exit Function
''''''''''End If
''''''''''
'''''''''''Cerramos el fichero de errores
''''''''''FinErroresLinea
''''''''''
'''''''''''Pasamos el fichero a carpeta procesados
'''''''''''Indicativos
''''''''''Label11.Caption = "Moviendo fichero a procesados"
''''''''''Label11.Refresh
''''''''''Cad = Dir(mConfig.DirProcesados, vbDirectory)
''''''''''If Cad <> "" Then
''''''''''    'SI QUE EXISTE
''''''''''    Cad = "\PR" & Format(Now, "yymmdd") & ".tik"
''''''''''    Cad = mConfig.DirProcesados & Cad
''''''''''    Else
''''''''''        Cad = "No existe la carpeta procesados." & vbCrLf
''''''''''        Cad = Cad & "Se copiará sobre la misma carpeta de la aplicación."
''''''''''        MsgBox Cad, vbExclamation
''''''''''        Cad = App.Path & "\PR" & Format(Now, "yymmdd") & ".tik"
''''''''''End If
''''''''''FileCopy NombreFichero, Cad
''''''''''Kill NombreFichero
'''''''''''Indicativos
''''''''''Label11.Caption = "Fichero en procesados: " & Cad
''''''''''Label11.Refresh
''''''''''
''''''''''If TIPOALZICOOP Then
''''''''''    'No hace falta traspasar el temporal, puesto que ya los hemos traspasado
''''''''''    'antes
''''''''''    'ya estan los datos donde queremos
''''''''''    ProcesaFichero = 0
''''''''''    Exit Function
''''''''''End If
''''''''''
''''''''''
''''''''''' quitamos los marcajes de temporal
''''''''''Errores = TraspasaTemporal
'''''''''''mensaje de todo correcto
''''''''''If Errores = 0 Then
''''''''''    ProcesaFichero = 0
''''''''''    'MsgBox "El proceso importacion ha finalizado correctamente.", vbInformation
''''''''''    Else
''''''''''        If Errores < 125 Then
''''''''''            'MsgBox "Se han producido " & Errores & " error(es).", vbExclamation
''''''''''            ProcesaFichero = 2
''''''''''            Else
''''''''''                If Errores = 125 Then
''''''''''                    'MsgBox "Se ha producido un número elevado de errores( + de 125).", vbCritical
''''''''''                    ProcesaFichero = 2
''''''''''                    Else
''''''''''                        If Errores = 127 Then
''''''''''                            'MsgBox "Ningún dato a traspasar. ", vbCritical
''''''''''                            ProcesaFichero = 1
''''''''''                        End If
''''''''''                End If
''''''''''        End If
''''''''''End If
''''''''''Exit Function
''''''''''ErrProcesaFichero:
''''''''''    Cad = "Se ha producido un error mientras se procesaba el fichero de marcajes." & vbCrLf
''''''''''    Cad = Cad & " RUTA: " & NombreFichero & vbCrLf
''''''''''    Cad = Cad & " ERROR: " & vbCrLf
''''''''''    Cad = Cad & "      Número: " & Err.Number & vbCrLf
''''''''''    Cad = Cad & "      Descripción: " & Err.Description
''''''''''    MsgBox Cad, vbCritical
''''''''''    ProcesaFichero = 1
''''''''''End Function



Private Sub PrimeraVezDeFechaHazProduccion()
    On Error Resume Next
        
        ImpFechaFin = "select * into tmpTareasRealizadas2 from tmpTareasRealizadas where trabajador=-1"
        conn.Execute ImpFechaFin
        If Err.Number <> 0 Then
            'Error
        End If
End Sub


'Segun sea un control normal, o un control tipo ALZICOOP
Private Sub ObtenerPrimeraClave()
Dim RS As ADODB.Recordset
Dim Cad As String

Set RS = New ADODB.Recordset
ContadorSecuencia = 1
If TIPOALZICOOP Then
        Cad = "SELECT MAX(secuencia) FROM TipoAlzicoop"
        RS.Open Cad, conn, , , adCmdText
        If Not RS.EOF Then
                If Not IsNull(RS.Fields(0)) Then
                ContadorSecuencia = RS.Fields(0) + 1
            End If
        End If
    'ELSE
    Else
        Cad = "SELECT MAX(secuencia) FROM TemporalFichajes"
        RS.Open Cad, conn, , , adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then
                ContadorSecuencia = RS.Fields(0) + 1
            End If
        End If
End If
RS.Close
End Sub




'Los errores que devuelve son
' 0.- Ningun error
' 1..124 Nº de errores que se han producido
' 125 .- Mas errors que 124
' 126 .- Error grabando datos
' 127 .- Ningun marcaje a trasapasar
Private Function TraspasaTemporal() As Byte
Dim RIni As ADODB.Recordset
Dim RFin As ADODB.Recordset
Dim RError As ADODB.Recordset
Dim CadInci As String
Dim Codigo As Long
Dim ContError As Long
Dim ContFich As Long
Dim CuantosErrores As Integer
Dim Totales As Long
Dim KReg As Long

Set RIni = New ADODB.Recordset
RIni.CursorType = adOpenKeyset
RIni.Open "TemporalFichajes", conn, , , adCmdTable
If RIni.EOF Then
    TraspasaTemporal = 127
    RIni.Close
    Set RIni = Nothing
    Exit Function
End If
Totales = RIni.RecordCount
'Indicativos
Label11.Caption = "Pasando temporal a tabla control de fichajes "
Label11.Refresh
'marcamos por si error
On Error GoTo ErrTraspaso
conn.BeginTrans
Set RFin = New ADODB.Recordset
Set RError = New ADODB.Recordset
RFin.CursorType = adOpenKeyset
RFin.LockType = adLockOptimistic
RError.CursorType = adOpenKeyset
RError.LockType = adLockOptimistic
'Abrimos los regsitros destino
' el correcto y el erroneo
RError.Open "ErrorTarjetas", conn, , , adCmdTable
If RError.EOF Then
    ContError = 1
    Else
        RError.MoveLast
        ContError = RError!Secuencia + 1
End If
RFin.Open "EntradaFichajes", conn, , , adCmdTable
If RFin.EOF Then
    ContFich = 1
    Else
        RFin.MoveLast
        ContFich = RFin!Secuencia + 1
End If
CuantosErrores = 0
KReg = 0
While Not RIni.EOF
    KReg = KReg + 1
    If RIni!idInci = 0 Then
        CadInci = "OK"
        Else
        CadInci = DevuelveTextoIncidencia(RIni!idInci)
    End If
    Codigo = DevuelveCodigo(RIni!Numtarjeta)
    If CadInci = "" Or Codigo < 0 Then
        CuantosErrores = CuantosErrores + 1
        'Ha habido un error
        RError.AddNew
        RError!Secuencia = ContError
        RError!Fecha = RIni!Fecha
        RError!Hora = RIni!Hora
        RError!idInci = RIni!idInci
        RError!Numtarjeta = RIni!Numtarjeta
        If CadInci = "" Then RError!Error = "La incidencia no es correcta"
        If Codigo < 0 Then RError!Error = RError!Error & "  ---  El codigo de tarjeta no corresponde con ningun trabajador"
        RError.Update
        ContError = ContError + 1
        'Si es correcto
        Else
            RFin.AddNew
            RFin!Secuencia = ContFich
            RFin!Fecha = RIni!Fecha
            RFin!Hora = RIni!Hora
            RFin!idInci = RIni!idInci
            RFin!idTrabajador = Codigo
            ContFich = ContFich + 1
            RFin.Update
        End If
    RIni.MoveNext
    'Indicativos
    Label11.Caption = "Registro temporal:  " & KReg & " de " & Totales
    Label11.Refresh
Wend
RIni.Close
RFin.Close
RError.Close
Set RFin = Nothing
Set RError = Nothing
'--------------------
'Borramos el temporal
RIni.CursorType = 0
RIni.Open "Delete * from TemporalFichajes", conn, , , adCmdText
Set RIni = Nothing
conn.CommitTrans
'Regresamos con un valor
If CuantosErrores > 124 Then
    TraspasaTemporal = 125
    Else
        TraspasaTemporal = CByte(CuantosErrores)
End If
Exit Function
ErrTraspaso:
    conn.RollbackTrans
    MsgBox "Se ha producido un error traspasando." & Err.Description, vbExclamation
    TraspasaTemporal = 126
End Function


Private Function GeneracionMarcajes() As Boolean
Dim B As Boolean
Dim Aux As Boolean
Dim Cad As String
Dim RsFechas As ADODB.Recordset



GeneracionMarcajes = False

Set RsFechas = New ADODB.Recordset

Cad = "SELECT DISTINCT EntradaFichajes.IdTrabajador"
Cad = Cad & " FROM EntradaFichajes "
Cad = Cad & " WHERE "
Cad = Cad & " EntradaFichajes.Fecha >=" & ImpFechaIni
Cad = Cad & " AND EntradaFichajes.Fecha <=" & ImpFechaFin

RsFechas.Open Cad, conn, , , adCmdText
If RsFechas.EOF Then
    Cad = vbCrLf
    If txtFecha(0).Text <> "" Then Cad = " desde: " & txtFecha(0) & vbCrLf
    If txtFecha(1).Text <> "" Then Cad = Cad & " hasta: " & txtFecha(1) & vbCrLf
    MsgBox "No hay ningún dato a importar " & Cad, vbExclamation
    'Rs.Close
    Exit Function
End If



    

'Ahora procesaremos los marcajes
'----------------------------------
B = True
While Not RsFechas.EOF
    

    'procesamos
    Aux = GeneraEntradasMarcajes(RsFechas.Fields(0))
    B = B And Aux
    RsFechas.MoveNext
Wend
GeneracionMarcajes = B
End Function



Private Function GeneraEntradasMarcajes(vCod As Long) As Boolean
Dim RsCodigos As ADODB.Recordset
Dim Rss As ADODB.Recordset
Dim Reg As ADODB.Recordset
Dim vM As CMarcajes
Dim Cad As String
Dim cHorario As Long
Dim vEmpresa As Long
Dim TipoControl As Byte
Dim HayMarcajes As Long
Dim RC As Byte


GeneraEntradasMarcajes = False
Set vH = New CHorarios
Set vE = New CEmpresas


'AHora leemos los codigos de la empresa y del horario para ese trabajador
Set Rss = New ADODB.Recordset
Rss.Open "Select IdHorario,IdEmpresa from Trabajadores where IdTrabajador=" & vCod, conn, , , adCmdText
If Not Rss.EOF Then
    vEmpresa = DBLet(Rss!IdEmpresa, "N")
    cHorario = DBLet(Rss!IdHorario, "N")
    Else
        'Error leyendo los datos de la empresa y el horario del empleado
        'salimos
        Exit Function
End If
Rss.Close
Set Rss = Nothing


'Ahora vemos el tipo de control que hacemos sobre el trabajador
'Si es total, parcial, solo marcajes ...
Set Rss = New ADODB.Recordset
Rss.Open "Select Control from Trabajadores where IdTrabajador=" & vCod, conn, , , adCmdText
If Not Rss.EOF Then
    TipoControl = DBLet(Rss.Fields(0))
End If
Rss.Close
Set Rss = Nothing


'Ahora leemos los datos de la empresa
If vE.Leer(vEmpresa) = 1 Then
    'Error leyendo los datos
    Exit Function
    Set vE = Nothing
End If


Cad = "SELECT DISTINCT EntradaFichajes.Fecha"
Cad = Cad & " FROM EntradaFichajes "
Cad = Cad & " WHERE EntradaFichajes.IdTrabajador=" & vCod
Cad = Cad & " AND EntradaFichajes.Fecha >=" & ImpFechaIni
Cad = Cad & " AND EntradaFichajes.Fecha <=" & ImpFechaFin
Cad = Cad & " ORDER BY Fecha"

Set RsCodigos = New ADODB.Recordset
RsCodigos.Open Cad, conn, , , adCmdText
If RsCodigos.EOF Then
    MsgBox "Ninguna entrada para esta fecha.", vbExclamation
    'Rs.Close
    Exit Function
End If

Label10.Caption = "Generando marcajes: "
Label10.Refresh
While Not RsCodigos.EOF
    '-----------------------------------------------------
    'Comprobamos si existen ya marcajes para esos valores
    HayMarcajes = YaExistenMarcajes(CInt(vCod), RsCodigos.Fields(0))
    RC = vbYes
    If HayMarcajes > 0 Then
        Cad = "Ya existen marcajes para el trabajador cod: " & vCod & "   y fecha: " & RsCodigos.Fields(0) & vbCrLf
        Cad = Cad & "  ¿ Quiere eliminar el antiguo marcaje. ?" & vbCrLf
        Cad = Cad & "   .- Si --> Eliminamos los antiguos" & vbCrLf
        Cad = Cad & "   .- No --> Dejamos de procesar estos datos" & vbCrLf
        RC = MsgBox(Cad, vbQuestion + vbYesNo)
        If RC = vbYes Then
            Set vM = New CMarcajes
            If vM.Leer(HayMarcajes) = 0 Then vM.Eliminar
            Set vM = Nothing
        End If
    End If
    If RC = vbYes Then
        'Horario para ese dia
        If vH.Leer(CInt(cHorario), RsCodigos.Fields(0)) = 0 Then
            Set vM = New CMarcajes
            'ASignamos al objeto Entrada Marcaje tanto el trabajador
            'como la fecha en la que estamos
            vM.Siguiente
            vM.Fecha = RsCodigos!Fecha
            vM.idTrabajador = vCod
            If vM.Agregar = 1 Then
                'Si no se puede insertar
                Else

                    'Pa saber por donde vamos
                    Label11.Caption = "Trabajador " & vCod & " - Fecha: " & vM.Fecha
                    Label11.Refresh
                    
                    If MiEmpresa.QueEmpresa = 4 Then
                        Cad = "UPDATE EntradaFichajes set idinci=0  WHERE IdTrabajador=" & vM.idTrabajador
                        Cad = Cad & " AND Fecha=#" & Format(vM.Fecha, "yyyy/mm/dd") & "#"
                        Cad = Cad & " AND idinci=2"
                        conn.Execute Cad
                    
                        espera 0.25
                    End If
                    'procesando el marcaje, aunque depende del
                    'tipo de control que se le hace
                    Select Case TipoControl
                    Case 0, 1  'El tipo 0 y uno lo encuadramos en el mismo tipo 0
                        ProcesarMarcaje_Tipo1 vM
                    Case 2
                        ProcesarMarcaje_Tipo2 vM
                    Case 3
                        ProcesarMarcaje_Tipo3 vM
                    End Select
            End If
            Set vM = Nothing
        End If  'de leer horario
    End If ' de rc=vbyes
    'Ultima sentencia: avanzar registro
    RsCodigos.MoveNext
Wend
Set vH = Nothing
Set vE = Nothing
GeneraEntradasMarcajes = True
End Function




Private Sub Form_Load()
PrimeraVez = True
Label10.Caption = ""
Label11.Caption = ""
If Opcion = 0 Then SugerirFechasImportacion

Frame1.Visible = Opcion = 0
Me.cmdAceptar.Visible = Opcion = 0
Me.cmdComprobar.Visible = Opcion = 0
Me.cmdSalir.Visible = Opcion = 0
End Sub

Private Sub frmC_Selec(vFecha As Date)
'Segun si este "" o con texto el tag del txtFecha(0) sera el incio el el fin
If txtFecha(0).Tag = "" Then
    'Es fin
    txtFecha(1).Text = Format(vFecha, "dd/mm/yyyy")
    Else
        txtFecha(0).Text = Format(vFecha, "dd/mm/yyyy")
End If
End Sub

Private Sub Image1_Click(Index As Integer)
If Index = 0 Then
    txtFecha(0).Tag = "ESTE"
    Else
        txtFecha(0).Tag = ""
End If
Set frmC = New frmCal
frmC.Fecha = Now
frmC.Show vbModal
Set frmC = Nothing
End Sub









Private Sub txtFecha_GotFocus(Index As Integer)
txtFecha(Index).SelStart = 0
txtFecha(Index).SelLength = Len(txtFecha(Index).Text)
End Sub



Private Sub txtFecha_LostFocus(Index As Integer)
If txtFecha(Index) = "" Then Exit Sub
If Not EsFechaOK(txtFecha(Index)) Then
    MsgBox "Fecha incorrecta: " & txtFecha(Index).Text, vbExclamation
    txtFecha(Index).Text = ""
End If

End Sub


Private Function DatosOk() As Boolean
Dim RS As ADODB.Recordset
Dim Cad As String

DatosOk = False
If txtFecha(0) = "" Or txtFecha(1) = "" Then
    MsgBox "Ponga las fechas del intervalo", vbExclamation
    Exit Function
End If
If txtFecha(0) = "" Then
    ImpFechaIni = "1900/01/01"
    Else
        If Not IsDate(txtFecha(0).Text) Then
            MsgBox "Fecha inicio incorrecta.", vbExclamation
            Exit Function
        End If
        ImpFechaIni = Format(txtFecha(0).Text, "yyyy/mm/dd")
End If
'Fecha fin
If txtFecha(1) = "" Then
    ImpFechaFin = Year(Now) + 10 & "/01/01"
    Else
        If Not IsDate(txtFecha(1).Text) Then
            MsgBox "Fecha fin incorrecta.", vbExclamation
            Exit Function
        End If
        ImpFechaFin = Format(txtFecha(1).Text, "yyyy/mm/dd")
End If
'Comprobamos fecha inicio<fin
If CDate(ImpFechaIni) > CDate(ImpFechaFin) Then
    MsgBox "Fecha de incio es mayor que la fecha final.", vbExclamation
    Exit Function
End If
ImpFechaFin = "#" & ImpFechaFin & "#"
ImpFechaIni = "#" & ImpFechaIni & "#"

Set RS = New ADODB.Recordset
Cad = "Select count(*) from marcajes where fecha >=" & ImpFechaIni & " AND Fecha <=" & ImpFechaFin
RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
Cad = ""
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then
        If RS.Fields(0) <> 0 Then Cad = "TIENE"
    End If
End If
RS.Close
Set RS = Nothing
If Cad <> "" Then
    Cad = "Ya hay datos en presencia para las fechas a procesar."
    Cad = Cad & vbCrLf & " ¿Desea continuar?"
    If MsgBox(Cad, vbQuestion + vbYesNoCancel + vbDefaultButton2) <> vbYes Then Exit Function
End If

DatosOk = True
End Function



Private Sub ComprobarMarcajesPendientes()
Dim RS As ADODB.Recordset
Dim i As Integer
Dim Cad As String

Screen.MousePointer = vbDefault
'Auqi diremos si quedan marcajes pendientes
'pendientes de procesar
Cad = "Select DISTINCT(fecha) from EntradaFichajes"
Cad = Cad & " WHERE Fecha<= " & ImpFechaIni
Cad = Cad & " OR Fecha>=" & ImpFechaFin

Set RS = New ADODB.Recordset
RS.Open Cad, conn, , , adCmdText
If RS.EOF Then GoTo salida
i = 0
Cad = ""
While Not RS.EOF
    Cad = Cad & Format(RS!Fecha, "dd/mm/yyyy") & "   "
    i = i + 1
    If i = 5 Then
        i = 0
        Cad = Cad & vbCrLf
    End If
    RS.MoveNext
Wend
RS.Close

Cad = "Estan pendientes de procesar las siguientes fechas: " & vbCrLf & vbCrLf & Cad
MsgBox Cad, vbExclamation, "ARIPRES"

'------------------------------------
'Por si caso quedan fechas entre ellas
Cad = "Select DISTINCT(idTrabajador) from EntradaFichajes"
Cad = Cad & " WHERE Fecha>= " & ImpFechaIni
Cad = Cad & " AND Fecha<=" & ImpFechaFin
RS.Open Cad, conn, , , adCmdText
If RS.EOF Then GoTo salida
i = 0
Cad = ""
While Not RS.EOF
    Cad = Cad & RS!idTrabajador & "   "
    i = i + 1
    If i = 5 Then
        i = 0
        Cad = Cad & vbCrLf
    End If
    RS.MoveNext
Wend
RS.Close


'Para poner el puntero y cerrar rs
salida:
    Set RS = Nothing
    Screen.MousePointer = vbDefault
End Sub


Private Function ObtnerNumSecuenciaEntradaMarcajes() As Long
Dim RS As ADODB.Recordset

ObtnerNumSecuenciaEntradaMarcajes = 1
Set RS = New ADODB.Recordset
RS.Open "Select MAX(Secuencia) from EntradaFichajes", conn, , , adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then _
        ObtnerNumSecuenciaEntradaMarcajes = RS.Fields(0) + 1
End If
RS.Close
Set RS = Nothing
End Function



Private Sub ReOrdenarEntradaFichajes()
Dim RS As ADODB.Recordset
Dim Minimo As Long
 
Minimo = 1
Set RS = New ADODB.Recordset
RS.Open "Select MIN(Secuencia) from EntradaFichajes", conn, , , adCmdText
If Not RS.EOF Then
    If Not IsNull(RS.Fields(0)) Then _
        Minimo = RS.Fields(0)
End If
RS.Close

'Si el minimo no es ..... salimos
If Minimo > 20000 Then
    Label10.Caption = "Reordenando tabla entrada de fichajes."
    Label10.Refresh
    Minimo = 1
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    RS.Open "SELECT * FROM EntradaFichajes ORDER BY Secuencia", conn, , , adCmdText
    While Not RS.EOF
        RS!Secuencia = Minimo
        RS.Update
        Label11.Caption = " Valor : " & Minimo
        Label11.Refresh
        Minimo = Minimo + 1
        RS.MoveNext
        
    Wend
    RS.Close
End If
Label10.Caption = ""
Label10.Refresh
'Salimos
Set RS = Nothing
End Sub








''-------------------------------------------------------------------------
'' Coje de la tabla de MarcajesKimaldi y para cada trabajador, y fecha, genera
'' las entradas en la tabla entradamarcajes para luego procesarlos
''
'' Todos los registros de entradafichajes los generaremos a partir de la tabla de kimaldi
''
''
'Private Sub GeneraMarcajesKimaldi()
'Dim Rs As ADODB.Recordset
'Dim RT As ADODB.Recordset
'Dim SQL As String
'Dim INSE As String
'Dim con As Long
'Dim Trab As Long
'Dim FechaANT As Date
'Dim Insertar As Boolean
'Dim Hora As Date
'Dim CodTarea As String
'Dim EsperoSalida As Boolean
'
'    'Los pasamos a tmpMarcajesKimaldi
'    Set Rs = New ADODB.Recordset
'    Set RT = New ADODB.Recordset
'    Label11.Caption = "Creando  tabla intermedia"
'    Label11.Refresh
'    SQL = "Delete * from tmpMarcajesKimaldi"
'    Conn.Execute SQL
'    Label11.Caption = "Pasando a tabla intermedia"
'    Label11.Refresh
'    SQL = "Insert into tmpMarcajesKimaldi Select * from MarcajesKimaldi"
'    SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
'    Conn.Execute SQL
'
'
'    'Marcar SALIDA MASIVA
'    'Si tiene salida masiva meteremos la salida masiva, generando un ticaje
'    'para cada trabjador vinculado a esa marca
'    ' Recorro los ticajes viendo la tarea k es
'    ' Cuando encuentro "SALIDA" , y mientras encuentre trabajadores, inserto
'    ' en entradamarcajes
'    SQL = "Select Tarjeta from Tareas where Tipo=1"   'Salida masiva
'    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    CodTarea = ""
'    If Not Rs.EOF Then
'        If Not IsNull(Rs.Fields(0)) Then CodTarea = Rs.Fields(0)
'    End If
'    Rs.Close
'    If CodTarea <> "" Then
'        'OK, hay una tarea k es ticada masiva de salida
'        'Entre las fechas solicitadas. Buscaremos la tarea
'        ' y los marcajes k siguen son salidas, y los modificaremos poniendo una S
'        SQL = "Select * from tmpMarcajesKimaldi"
'        SQL = SQL & "  WHERE (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
'        SQL = SQL & " ORDER BY nodo,Fecha, Hora"
'        RT.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'        Trab = -1  'Sera el nodo
'        Insertar = False
'        While Not RT.EOF
'            If Trab <> RT!nodo Then
'                Trab = RT!nodo
'                Insertar = False
'            End If
'            'Si no es insertar
'            If Insertar Then
'                'Si el ticaje empieza por codigo trabajador
'                If Mid(RT!Marcaje, 1, 1) = mConfig.DigitoTrabajadores Then
'                        SQL = "UPDATE tmpMarcajesKimaldi SET TipoMens ='S' "
'                        SQL = SQL & " WHERE Nodo =" & RT!nodo
'                        SQL = SQL & " AND Fecha  = #" & Format(RT!Fecha, "yyyy/mm/dd") & "#"
'                        SQL = SQL & " AND Hora = #" & Format(RT!Hora, "hh:mm:ss") & "#"
'                        SQL = SQL & " AND Marcaje='" & RT!Marcaje & "'"
'                        Conn.Execute SQL
'                Else
'                    Insertar = False
'                End If
'            End If
'            If Not Insertar Then
'                'Estoy buscando las tarea de salida masiva
'                If RT!Marcaje = CodTarea Then
'                    'A partir de aqui son ticajes masivos de salida
'                    Insertar = True
'                End If
'            End If
'            RT.MoveNext
'        Wend
'        RT.Close
'    End If
'
'
'    'Eliminando datos de tareas
'    SQL = "Select * from tmpMarcajesKimaldi"
'    Rs.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
'    con = 0
'    While Not Rs.EOF
'        con = con + 1
'        Label11.Caption = "Registro: " & con
'        Label11.Refresh
'        If Mid(Rs!Marcaje, 1, 1) <> mConfig.DigitoTrabajadores Then Rs.Delete
'        Rs.MoveNext
'    Wend
'    Rs.Close
'
'    If con = 0 Then Exit Sub
'
'    'Obtenemos el max de secuencia para seguir insertando
'    con = 0
'    SQL = "Select max(secuencia) from Entradafichajes"
'
'    Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not Rs.EOF Then con = DBLet(Rs.Fields(0), "N")
'    Rs.Close
'    con = con + 1
'
'
'    'Ya tenemos la secuencia, ahora para cada fecha, y trbajador, introducimos en secuencia
'    INSE = "INSERT INTO entradafichajes (Secuencia,idTrabajador,Fecha,Hora,idInci,HoraReal) VALUES ("
'
'    'Cojemos cada uno de los trabajadores
'    SQL = "Select Marcaje from tmpMarcajesKimaldi  "
'    SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
'    SQL = SQL & " GROUP BY marcaje"    'El coodin es %
'    SQL = SQL & " HAVING (((tmpMarcajesKimaldi.Marcaje) Like '" & mConfig.DigitoTrabajadores & "%'));"
'    RT.Open SQL, Conn, adOpenDynamic, adLockOptimistic, adCmdText
'
'    While Not RT.EOF
'
'            'Trabajador
'            Trab = DevuelveTrabajador(RT!Marcaje, Rs)
'            If Trab > 0 Then
'                'Fecha anterior
'                FechaANT = CDate("01/01/1900")
'
'                SQL = "Select * from tmpMarcajesKimaldi where Marcaje = '" & RT!Marcaje & "'"
'                SQL = SQL & " AND (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
'                SQL = SQL & " ORDER BY Fecha, Hora"
'                Rs.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                While Not Rs.EOF
'                    If FechaANT <> Rs!Fecha Then
'                        'Insertamos un marcaje de entrada
'                        Insertar = True
'                        FechaANT = Rs!Fecha
'                        EsperoSalida = (DBLet(Rs!tipomens) <> "S")
'                    Else
'                        If EsperoSalida Then
'                            If DBLet(Rs!tipomens) = "S" Then
'                                Insertar = True
'                                EsperoSalida = False
'                            End If
'                        Else
'                            'NO espero la salida
'                            If DBLet(Rs!tipomens) <> "S" Then
'                                Insertar = True
'                                EsperoSalida = True
'                            Else
'                                'Si la salida difiere en mas de 15 minutos entonces la tico aunque no
'                                'genere incorrecto
'                                If DateDiff("n", Hora, Rs!Hora) > 15 Then Insertar = True
'                            End If
'                        End If
'                    End If
'                    Hora = Rs!Hora
'
'                    If Insertar Then
'                        '(Secuencia,idTrabajador,Fecha,Hora,idInci) VALUES ("
'                        SQL = con & "," & Trab
'                        SQL = SQL & ",#" & Format(Rs!Fecha, "yyyy/mm/dd") & "#"
'                        SQL = SQL & ",#" & Format(Rs!Hora, "hh:mm:ss") & "#"
'                        SQL = SQL & ",0,#" & Format(Rs!Hora, "hh:mm:ss") & "#"
'                        SQL = INSE & SQL & ")"
'                        Conn.Execute SQL
'                        con = con + 1
'                    End If
'
'                    Insertar = False
'                    'Siguiente
'                    Rs.MoveNext
'                Wend
'                Rs.Close
'
'            Else
'                'El numero de tarjeta no pertenece a ninugun trabajador
'                'Esto no deberia de pasar porque cuando traemos los marcajes
'                'ya compruebo si es de untrabajador o de una tarea
'
'            End If
'
'        'Siguiente
'        RT.MoveNext
'    Wend
'    RT.Close
'
'
'
'    '---------------------------------------
'    'Rectificamos si tiene rectificar
'    Dim RTRa As ADODB.Recordset
'    Dim Aux As String
'    Dim cad As String
'    Dim h2 As String    'Hora fin
'    Dim H3 As String    'Hora modificada
'    Dim H1 As String
'
'    cad = "SELECT DISTINCT Trabajadores.Seccion"
'    cad = cad & " FROM EntradaFichajes ,Trabajadores WHERE "
'    cad = cad & " EntradaFichajes.idTrabajador = Trabajadores.idTrabajador"
'
'    Set RT = New ADODB.Recordset
'    Set Rs = New ADODB.Recordset
'    Set RTRa = New ADODB.Recordset
'
'    'Ponemos el cursortype para los Rs, para que no de fallo
'    Rs.CursorType = adOpenKeyset
'    Rs.LockType = adLockOptimistic
'
'    RT.Open cad, Conn, , , adCmdText
'    While Not RT.EOF
'        Aux = RT.Fields(0)
'        'Para cada seccion vemos si tiene Rs
'        cad = "SELECT * FROM ModificarFichajes "
'        cad = cad & " WHERE IdSeccion= " & Aux
'        Rs.Open cad, Conn, , , adCmdText
'        While Not Rs.EOF
'            Label11.Caption = "Rectifcando marcajes para seccion: " & Aux & " /" & Rs.Fields(0)
'            Label11.Refresh
'            H1 = "#" & Format(Rs.Fields(2), "hh:mm") & "#"
'            h2 = "#" & Format(Rs.Fields(3), "hh:mm") & "#"
'            H3 = "#" & Format(Rs.Fields(4), "hh:mm") & "#"
'            'Creamos la consulta de acutalizacion
'            'Para cada Rs modificamos la tabla
'            cad = "SELECT DISTINCT(EntradaFichajes.idTrabajador)" ', Trabajadores.Seccion"
'            cad = cad & " FROM EntradaFichajes INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador"
'            cad = cad & " WHERE (((Trabajadores.Seccion)=" & Aux & "));"
'            'Abrimos
'            RTRa.Open cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'            cad = "UPDATE EntradaFichajes"
'            cad = cad & " SET Hora = " & H3
'            cad = cad & " WHERE Hora >= " & H1
'            cad = cad & " AND Hora <= " & h2
'            cad = cad & " AND idTrabajador ="
'            While Not RTRa.EOF
'                'Ejecutamos el SQL
'                Conn.Execute cad & RTRa.Fields(0)
'                RTRa.MoveNext
'            Wend
'            RTRa.Close
'            Rs.MoveNext
'        Wend
'        'Cerramos el recordset de modificar marcjaes
'        Rs.Close
'        'Movemos al siguiente
'        RT.MoveNext
'    Wend
'    'Cerramos los recordset
'    RT.Close
'
'
'
'
'
'    'Vamos a eliminar las entradas k transcurran menos de X minutos
'    'SQL = "SELECT EntradaFichajes.Fecha, EntradaFichajes.Hora, EntradaFichajes.idTrabajador"
'    SQL = "SELECT * From EntradaFichajes"
'    SQL = SQL & " ORDER BY EntradaFichajes.Fecha, EntradaFichajes.Hora, EntradaFichajes.idTrabajador;"
'    FechaANT = "01/01/1900"
'    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Insertar = False
'    Label10.Caption = "Busqueda duplicados"
'    Me.Refresh
'    While Not RT.EOF
'        Label11.Caption = RT!Fecha & " - " & RT!Hora
'        Label11.Refresh
'        If FechaANT = RT!Fecha Then
'            If DateDiff("n", Hora, RT!Hora) <= 3 Then
'                If Trab = RT!idTrabajador Then Insertar = True
'            End If
'        Else
'            FechaANT = RT!Fecha
'        End If
'        Hora = RT!Hora
'        Trab = RT!idTrabajador
'        If Insertar Then
'            'Realmente es ELIMINAR
'            RT.Delete
'            Insertar = False
'        End If
'
'        'Siguiente
'        RT.MoveNext
'    Wend
'    RT.Close
'
'
'
'
'
'
'
'
'
'
'
'
'
'    Set Rs = Nothing
'    Set RT = Nothing
'
'End Sub


'Private Function DevuelveTrabajador(ByRef Texto, ByRef R As ADODB.Recordset) As Long
'Dim SQL As String
'    DevuelveTrabajador = -1
'    SQL = "Select idTrabajador from Trabajadores where numtarjeta = '" & Texto & "';"
'    R.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not R.EOF Then
'        If Not IsNull(R.Fields(0)) Then DevuelveTrabajador = R.Fields(0)
'    End If
'    R.Close
'End Function



'Buscaremos cual es la primerafecha en entradafichajes. Esa tablka es la tabla donde kedan datos por procesar.
'Luego en fin un dia menso a la actual
Private Sub SugerirFechasImportacion()
Dim RS As ADODB.Recordset
Dim Fec1 As Date


    Set RS = New ADODB.Recordset
    Fec1 = Now
    If mConfig.Kimaldi Then
            'Comprobamos lo ultimo en marcajes
            RS.Open "Select max(fecha) from marcajes", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If Not RS.EOF Then
                If Not IsNull(RS.Fields(0)) Then
                    Fec1 = RS.Fields(0)
                    Fec1 = Fec1 + 1
                End If
            End If
    Else
        RS.Open "Select min(fecha) from entradafichajes", conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then Fec1 = RS.Fields(0)
        End If
    End If
    RS.Close
    Set RS = Nothing
    txtFecha(0).Text = Format(Fec1, "dd/mm/yyyy")
    If DateDiff("d", Now, Fec1) <= 0 Then
        Fec1 = DateAdd("d", -1, Now)
        txtFecha(1).Text = Format(Fec1, "dd/mm/yyyy")
    Else
        txtFecha(1).Text = ""
    End If
End Sub



Private Sub InsertaEnTemporalTareaParaRevisiondesdeTicajes(Fecha As Date, Trabajador As Integer, Hora As Date, Tarea As String)
Dim SQL As String
    SQL = "INSERT into tmpTareasRealizadas (Fecha,Hora,  Trabajador,Tarea) VALUES ("
    SQL = SQL & "#" & Format(Fecha, "yyyy/mm/dd") & "#"
    SQL = SQL & ",#" & Format(Hora, "hh:mm:ss") & "#,"
    SQL = SQL & Trabajador & ","
    SQL = SQL & Tarea & ")"
    conn.Execute SQL
End Sub


'--------------------------------------------------------------------------
'Produccion.
' Insertaremos en la tabla tareas realizadas una entrada por cada trabajador, en k fecha y hora, la tarea correspondiente
'
'
Private Sub HazProduccion()
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim SQL As String
Dim AntTarea As Long
Dim Trabajador As Long
Dim Procesar As Boolean
Dim salida As Boolean
Dim Insertar As Boolean
Dim H1 As Date
Dim h2 As Date
Dim Fecha As Date
Dim DEc As Currency
Dim mH As CHorarios
Dim TotalHoras As Currency
Dim BucleRevision As Integer
Dim Revisiones As Integer

    Set RS = New ADODB.Recordset
    Set RT = New ADODB.Recordset
    
    Label10.Caption = "Produccion"
    Label11.Caption = "Preparando datos"
    Me.Refresh
    
    'Eliminamos datos temporales
    conn.Execute "delete from tmpTareasRealizadas"
    
    
    'Deberiamos comprobar k no existen datos en las fechas señaladas
    If Opcion = 0 Then
        'Cuando la opcion es la noraml comprobamos
        SQL = "Select * from TareasRealizadas "
        SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then
                MsgBox "Ya existen datos procesados para esas fechas", vbExclamation
                RS.Close
                Exit Sub
            End If
        End If
        RS.Close
    End If
    
    

    'Recalulo de tareas para los trabajadores
    'AHora las borramos
    If Opcion = 1 Then
        Label10.Caption = "Eliminando anteriores:"
        Label11.Caption = ""
        Me.Refresh
        'AQUI###
        'Si la opcion es regenerar las tareas para un trabajador, entonces
        Vector = Split(ListaTrabajadores, "|")
        
        
        
        For BucleRevision = 0 To UBound(Vector) - 1
            Label11.Caption = "Cod: " & Vector(BucleRevision)
            Label11.Refresh
            
'            Procesar = True  'El primero siempre lo insertare
'            SQL = "Select * from  TareasRealizadas HoraInicio"
'            SQL = SQL & " where Fecha = " & ImpFechaIni & " and trabajador = " & Vector(BucleRevision)
'            SQL = SQL & " ORDER by Horainicio"
'            RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'            While Not RS.EOF
'
'                If Procesar Then
'                    'Es el primer registro
'                    InsertaEnTemporalTareaParaRevisiondesdeTicajes RS!Fecha, RS!Trabajador, RS!horainicio, RS!Tarea
'                    Fecha = RS!Fecha
'                    Procesar = False
'                    H1 = RS!horafin
'
'                Else
'
'                    'A la hora que acaba la tarea anterior, es la que empieza la siguiente, luego
'                    InsertaEnTemporalTareaParaRevisiondesdeTicajes RS!Fecha, RS!Trabajador, RS!horainicio, RS!Tarea
'                    H1 = RS!horafin
'
'                End If
'
'                RS.MoveNext
'            Wend
'            RS.Close
'            'Metemos la salida de la ultima tarea
'            InsertaEnTemporalTareaParaRevisiondesdeTicajes Fecha, CInt(Vector(BucleRevision)), H1, "7072"
            
            
            SQL = "DELETE from TareasRealizadas "
            SQL = SQL & " where Fecha = " & ImpFechaIni & " and trabajador = " & Vector(BucleRevision)
            conn.Execute SQL
        Next BucleRevision
        ImpFechaFin = ImpFechaIni
    End If
    
    
    'Obtenemos la anterior ultima tarea k estaban realizando
    AntTarea = -1
    SQL = "Select Tarea from TareasRealizadas order by Fecha,Horafin"
    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        RS.MoveLast 'Vemos el ultimo registro
        AntTarea = DBLet(RS!Tarea, "N")
    End If
    RS.Close
    
    
    
    
    'Recorremos la tabla Kimaldi entre las fechas seleccionadas
    ' y para cada registro de trabajador le insertamos su tarea correspondiente
    Label10.Caption = "Preparando datos tareas realizadas"
    Label10.Refresh


    BucleRevision = 1 'Por defecto. la opcion 0(revision tooodo el dia ya lo tiene
    If Opcion = 1 Then
        If FechaHazProduccion <> "" Then
            'Si la fecha de ahora, no es la misma k la k habia
            If ImpFechaIni = FechaHazProduccion Then BucleRevision = 0
        End If
    End If
    
    
    If BucleRevision = 1 Then

        SQL = "Select * from MarcajesKimaldi "
        SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
        SQL = SQL & " ORDER BY Nodo,Fecha,Hora"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
        
        
    
        
        While Not RS.EOF
            Label11.Caption = RS!Fecha & "    " & RS!Hora
            Label11.Refresh
            Procesar = True
            salida = False
            If DBLet(RS!tipomens) <> "" Then
                If RS!tipomens <> "S" Then
                    Procesar = False
                Else
                    salida = True
                End If
            End If
            
            If Procesar Then
                Insertar = False
                If Not salida Then
                    'Veremos si es marcaje de trabajador o tarea
                    If Mid(RS!Marcaje, 1, 1) = mConfig.DigitoTrabajadores Then
                        'Trabajador
                        Insertar = CodigoCorrecto(True, RS!Marcaje, Trabajador)
                    Else
                        'Tarea
                        CodigoCorrecto False, RS!Marcaje, AntTarea
                        
                    End If
                Else
                    AntTarea = -1
                    Insertar = True
                    'Hay k ver k trabajador
                    CodigoCorrecto True, RS!Marcaje, Trabajador
                End If
                
                If Insertar Then
                    SQL = "INSERT into tmpTareasRealizadas (Fecha,Hora,  Trabajador,Tarea) VALUES ("
                    SQL = SQL & "#" & Format(RS!Fecha, "yyyy/mm/dd") & "#"
                    SQL = SQL & ",#" & Format(RS!Hora, "hh:mm:ss") & "#,"
                    SQL = SQL & Trabajador & ","
                    SQL = SQL & AntTarea & ")"
                    conn.Execute SQL
                End If
            End If
            
            
            
            'Siguiente
            RS.MoveNext
        Wend
        RS.Close
        
        
        
        'Si la opcion es 1, lo meto en la variable
        If Opcion = 1 Then
            SQL = "Delete from tmpTareasRealizadas2"
            conn.Execute SQL
            SQL = "INSERT INTO tmpTareasRealizadas2 SELECT * from tmpTareasRealizadas"
            conn.Execute SQL
            FechaHazProduccion = ImpFechaIni
        End If
    Else
        'Cuando vamos a revisar los marcajes antiguos
        SQL = "INSERT INTO tmpTareasRealizadas SELECT * from tmpTareasRealizadas2"
        conn.Execute SQL
    End If
    
    'Ahora, si la opcion es 1, eliminaremos las tareas que no tengan que ver con
    'los trabajadores implicados
    If Opcion = 1 Then
        Label10.Caption = "Eliminando datos innecesarios"
        Label11.Caption = ""
        Me.Refresh
        
        SQL = ""
        For Trabajador = 0 To UBound(Vector) - 1
            Label11.Caption = Vector(Trabajador)
            Label11.Refresh
            If SQL = "" Then
                SQL = Vector(Trabajador)
                ListaTrabajadores = "Delete from tmptareasrealizadas where trabajador <" & SQL
            Else
                ListaTrabajadores = "Delete from tmptareasrealizadas where trabajador >" & SQL
                SQL = Vector(Trabajador)
                ListaTrabajadores = ListaTrabajadores & " and trabajador <" & SQL
            End If
            conn.Execute ListaTrabajadores
        Next Trabajador
        'El resto por arriba
        ListaTrabajadores = "Delete from tmptareasrealizadas where trabajador >" & SQL
        conn.Execute ListaTrabajadores
        
        
        
        'AHora volvemos a meter. los ticajes de los trabajadores selecciondas
        For Trabajador = 0 To UBound(Vector) - 1
            Label11.Caption = "Comprobando nuevas entradas para: " & Vector(Trabajador)
            Label11.Refresh
            
            SQL = "Select * from tmptareasrealizadas where trabajador =" & Vector(Trabajador) & " ORDER BY hora"
            RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            'En RT pongo los marcjaes de presencia. Recorrere RT
            ' viendo si es entrada o salida.
            SQL = "Select * from entradamarcajes where idtrabajador =" & Vector(Trabajador) & " and fecha =" & ImpFechaIni & " ORDER BY horareal"
            RT.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            salida = False
            'Para el INSERT
            ListaTrabajadores = "INSERT INTO tmptareasrealizadas(Fecha,trabajador,hora,tarea) VALUES ("
            ListaTrabajadores = ListaTrabajadores & ImpFechaIni & "," & Vector(Trabajador) & ",#"
            While Not RT.EOF
                SQL = ""
                If RS.EOF Then
                    'INSERTAMOS SEGURO
                    SQL = ListaTrabajadores & Format(RT!HoraReal, "hh:mm:ss") & "#,"
                    
                        If salida Then
                            SQL = SQL & "7072)"
                         Else
                            SQL = SQL & "9999)"
                         End If
                         conn.Execute SQL
                        salida = Not salida
                        RT.MoveNext
                    
                Else
                    AntTarea = Abs(DateDiff("n", RS!Hora, RT!HoraReal))
                    
                    If AntTarea > 10 Then
                        'INSERTAR ESTA TAREA
                        SQL = ListaTrabajadores & Format(RT!HoraReal, "hh:mm:ss") & "#,"
                        
                        
                    Else
                        RT.MoveNext
                    End If
                    
                    If SQL <> "" Then
                        'INSERTAMOS en tmptareasrealizadas
                         If salida Then
                            SQL = SQL & "7072)"
                         Else
                            SQL = SQL & "9999)"
                         End If
                         conn.Execute SQL
                         RT.MoveNext
                    Else
                        RS.MoveNext
                    End If
                    
                    salida = Not salida
                End If
            Wend
            RS.Close
            RT.Close
        Next Trabajador
        
        
        
    End If
        
        
    '--------------------------------------------------------------------
    'Ahora aplicaremos la correccion en los ticajes. Para ello iremos viendo
    RectificarMarcajesTareas
    
    
    '---------------------------------------------------------------------
    'Para cada tarea vemos si esta duplicada
    
    SQL = "SELECT tmpTareasRealizadas.Fecha, tmpTareasRealizadas.Hora, tmpTareasRealizadas.trabajador, tmpTareasRealizadas.Tarea"
    SQL = SQL & " From tmpTareasRealizadas"
    SQL = SQL & " ORDER BY tmpTareasRealizadas.Fecha, tmpTareasRealizadas.Hora, tmpTareasRealizadas.trabajador,"
    SQL = SQL & " tmpTareasRealizadas.Tarea DESC;"
    RT.Open SQL, conn, adOpenStatic, adLockPessimistic, adCmdText
    h2 = "01/01/1900"
    Insertar = False
    While Not RT.EOF
        Label11.Caption = RT!Fecha & " - " & RT!Hora
        Label11.Refresh
        If h2 = RT!Fecha Then
            If DateDiff("n", H1, RT!Hora) <= 3 Then
                If Trabajador = RT!Trabajador Then
                    If AntTarea = RT!Tarea Then
                        'SQL = "DELETE FROM tmpTareasRealizadas WHERE"
                        'SQL = SQL & " Fecha= #" & Format(RT!Fecha, "yyyy/mm/dd") & "#"
                        'SQL = SQL & " AND Hora = #" & Format(RT!Hora, "hh:mm:ss") & "#"
                        'SQL = SQL & " AND Trabajador = " & RT!Trabajador
                        'SQL = SQL & " AND Tarea = " & RT!Tarea
                        Insertar = True
                    End If
                End If
            End If
        Else
            h2 = RT!Fecha
        End If
        H1 = RT!Hora
        Trabajador = RT!Trabajador
        AntTarea = RT!Tarea
        If Insertar Then
            RT.Delete
            Insertar = False
        End If
        'Siguiente
        RT.MoveNext
    Wend
    RT.Close
    
    
    'Ahora ya tenemos para cada trabajador las tareas k tiene
    'Reacorremos la tabla tmpTareasRealizadas, ordenada por fecha
    'e insertamos los datos
    SQL = "Select Trabajador,Fecha from tmpTareasRealizadas "
    SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
    SQL = SQL & " GROUP BY Trabajador,Fecha ORDER By Trabajador,Fecha"
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
        'POnemos los labels k hgan falta para k los  muestre el form
        Label11.Caption = RT!Trabajador & " - " & RT.Fields(1)
        Label11.Refresh
        
        'ASignamos intervalos de horas, con las tareas
        SQL = "Select * from tmpTareasRealizadas where Trabajador=" & RT!Trabajador
        SQL = SQL & " AND Fecha =#" & Format(RT.Fields(1), "yyyy/mm/dd") & "# order by hora"
        RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If Not RS.EOF Then
            AntTarea = RS!Tarea
            H1 = RS!Hora
            RS.MoveNext
            
            'A partir de aqui vamos construyendo los siguientes
            While Not RS.EOF
                h2 = RS!Hora
                Insertar = False   'Nos indicara si hemos insertado el valor
                SQL = "INSERT into TareasRealizadas (Trabajador,Fecha,Horainicio,HoraFin,horasTrabajadas,Tarea,Horas1,Horas2,Horas3) VALUES ("
                SQL = SQL & RT.Fields(0)
                SQL = SQL & ",#" & Format(RT!Fecha, "yyyy/mm/dd") & "#"
                SQL = SQL & ",#" & Format(H1, "hh:mm:ss") & "#"
                SQL = SQL & ",#" & Format(h2, "hh:mm:ss") & "#,"
                'Horas trabajadas
                DEc = CCur(DevuelveValorHora(h2))
                DEc = DEc - CCur(DevuelveValorHora(H1))
                SQL = SQL & TransformaComasPuntos(CStr(DEc)) & ","
                
                'Tarea
                SQL = SQL & AntTarea & ","
                'Horas normasles HORAS1
                SQL = SQL & TransformaComasPuntos(CStr(DEc)) & ","
                'HORAS EXTRAS, por defecto 0
                SQL = SQL & "0,"
                'Horas tipo 2. De mometno no tratadas
                SQL = SQL & "0)"
                
                conn.Execute SQL
                
                'Para el siguiente intervalo de horas
                H1 = h2
                
                'Si la tarea es salir entonces, deberia ser la ultima
                If RS!Tarea = -1 Then
                    Insertar = True
                    'No cambio la tarea,Si ha marcado mal las horas seran incorrectas
                Else
                    AntTarea = RS!Tarea
                End If
                
                'Siguiente
                RS.MoveNext
            Wend
            If Insertar = False Then
                'Stop
            End If
        End If
        'Cerramos RS
        RS.Close
        
        'Siguiente Trabajador, dia
        RT.MoveNext
    Wend
    RT.Close
        
    '---------------------------------------------------------------------
    'QUitaremos todas aquellas tareas k de tiempo sean menores a 2 minutos
    Label11.Caption = "Eliminar tareas realizadas tiempo = 0"
    Me.Refresh
    SQL = "DELETE from TareasRealizadas "
    SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
    SQL = SQL & " AND HorasTrabajadas < 0.05"
    conn.Execute SQL
    
    
    
    
    'Dim BucleRevision As Integer
    'Dim Revisiones As Integer
    If Opcion = 0 Then
        Revisiones = 0
    Else
        Revisiones = UBound(Vector) - 1
    End If
    
    For BucleRevision = 0 To Revisiones
    
            'Ahora para cada trabajador veremos si tiene horas extras o no
            'Esto nos lo dira pq el horario es exaustivo.
            'Entonces yo generare las horas extras y demas dentro de cada ticajes
            SQL = "SELECT Trabajadores.idTrabajador, TareasRealizadas.HorasTrabajadas, TareasRealizadas.Fecha,"
            SQL = SQL & "TareasRealizadas.HoraInicio,TareasRealizadas.HoraFin, Trabajadores.idHorario,Trabajadores.Control,TareasRealizadas.Fecha"
            SQL = SQL & " FROM Trabajadores INNER JOIN TareasRealizadas ON Trabajadores.IdTrabajador = TareasRealizadas.Trabajador"
            SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
            SQL = SQL & " AND Control <=1"   'Tipo de control EXAUSTIVO
            If Opcion = 1 Then SQL = SQL & " AND Trabajadores.idTrabajador =" & Vector(BucleRevision)
            
            RT.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            
            'Para los horarios
            Set mH = New CHorarios
            mH.IdHorario = -1
            Fecha = "31/10/1900"
            While Not RT.EOF
                Label11.Caption = "Trab: " & RT!idTrabajador
                Label11.Refresh
                'Horario
                If mH.IdHorario = RT!IdHorario Then
                    If Fecha <> RT!Fecha Then
                        Procesar = True
                    Else
                        Procesar = False
                    End If
                Else
                    Procesar = True
                End If
                
                If Procesar Then
                    'HAY k leer el horario
                    If mH.Leer(RT!IdHorario, RT!Fecha) = 1 Then
                        MsgBox "Error leyendo horario: " & RT!IdHorario & " - " & RT!Fecha, vbExclamation
                        Exit Sub
                    End If
                End If
                H1 = RT!horainicio
                h2 = RT!horafin
                
                TotalHoras = HorasExtra1(H1, h2, mH)
                If TotalHoras > 0 Then
                    'Stop
                    'Updateamos para poner horas extra 1
                    DEc = RT!HorasTrabajadas
                    SQL = "UPDATE TareasRealizadas set Horas2 =" & TransformaComasPuntos(CStr(TotalHoras))
                    DEc = DEc - TotalHoras
                    SQL = SQL & " , Horas1 =" & TransformaComasPuntos(CStr(DEc))
                    SQL = SQL & " WHERE trabajador = " & RT!idTrabajador & " AND Fecha = #" & Format(RT!Fecha, "yyyy/mm/dd") & "#"
                    SQL = SQL & " AND HoraInicio = #" & Format(RT!horainicio, "hh:mm:ss") & "#"
                    conn.Execute SQL
                End If
            
                'Siguiente
                RT.MoveNext
            Wend
            RT.Close
            
            '-----------------------------------------------------
            'Vemos los descuento por almuerzo y merienda
            '
            'Descontaremos primero si tiene en horas extras, y despues,
            'lo k kede en horas
            'Haremos un selec para los horarios
            'Y para cada horario, si tiene dtos, los aplicampos
            
            
            SQL = "SELECT Horarios.IdHorario, TareasRealizadas.Fecha"
            SQL = SQL & " FROM Horarios INNER JOIN (TareasRealizadas INNER JOIN"
            SQL = SQL & " Trabajadores ON TareasRealizadas.Trabajador = Trabajadores.IdTrabajador)"
            SQL = SQL & " ON Horarios.IdHorario = Trabajadores.IdHorario"
            SQL = SQL & " GROUP BY Horarios.IdHorario, TareasRealizadas.Fecha, Horarios.DtoAlm, Horarios.DtoMer"
            SQL = SQL & " HAVING (((TareasRealizadas.Fecha)>=" & ImpFechaIni
            SQL = SQL & " And (TareasRealizadas.Fecha)<=" & ImpFechaFin
            SQL = SQL & " ) AND ((Horarios.DtoAlm)>0)) OR ((("
            SQL = SQL & " TareasRealizadas.Fecha)>=" & ImpFechaIni
            SQL = SQL & " And (TareasRealizadas.Fecha)<=" & ImpFechaFin
            SQL = SQL & " ) AND ((Horarios.DtoMer)>0));"
            
            Label10.Caption = "Dto en produccion de Almuerzo-merienda"
            Me.Refresh
            RT.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            Set mH = New CHorarios
            While Not RT.EOF
                If mH.Leer(RT.Fields(0), RT.Fields(1)) = 0 Then
                    SQL = "Select TareasRealizadas.Trabajador FROM TareasRealizadas"
                    
                    SQL = "SELECT TareasRealizadas.Trabajador, Trabajadores.IdHorario"
                    SQL = SQL & " FROM TareasRealizadas INNER JOIN Trabajadores ON TareasRealizadas.Trabajador = Trabajadores.IdTrabajador"
                    SQL = SQL & " WHERE (((TareasRealizadas.Fecha) = #" & Format(RT.Fields(1), "yyyy/mm/dd") & "#"
                    'Solo el trabajador
                    If Opcion = 1 Then SQL = SQL & " AND idTrabajador =" & Vector(BucleRevision)
                    SQL = SQL & ")) GROUP BY TareasRealizadas.Trabajador, Trabajadores.IdHorario"
                    SQL = SQL & " HAVING (((Trabajadores.IdHorario)=" & mH.IdHorario & "))"
                    
                    RS.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    
                    'Calculamos la hora FIN almuerzo y merienda
                    If mH.DtoAlm > 0 Then
                        DEc = mH.DtoAlm * 60
                        H1 = DateAdd("n", DEc, mH.HoraDtoAlm)  'Hora inicio almuerzo
                    Else
                        H1 = "0:00:00"
                    End If
                    If mH.DtoMer > 0 Then
                        DEc = mH.DtoMer * 60
                        h2 = DateAdd("n", DEc, mH.HoraDtoMer)  'Hora inicio almuerzo
                    Else
                        h2 = "0:00:00"
                    End If
                    While Not RS.EOF
                        Label11.Caption = "Trabajador: " & RS.Fields(0)
                        Label11.Refresh
                        'Quitamos
                        QuitarHorasAlmuerzo RS.Fields(0), RT!Fecha, H1, h2, mH
                            
                        'Siguiente
                        RS.MoveNext
                    Wend
                    RS.Close
                End If
                'Siguiente
                RT.MoveNext
            Wend
            RT.Close
            Set mH = Nothing
        
            'Ahora calcularemos importes en funcion de los trabajadores, y la categoria
            '
            Label10.Caption = "Calculando costes produccion"
            Label11.Caption = ""
            Me.Refresh
            SQL = "SELECT TareasRealizadas.Trabajador, TareasRealizadas.Horas1, TareasRealizadas.Horas2,TareasRealizadas.Fecha,"
            SQL = SQL & " TareasRealizadas.Horas3, Trabajadores.PorcAntiguedad, Trabajadores.PorcSS, Categorias.Importe1,"
            SQL = SQL & " Categorias.Importe2, Categorias.Importe3,TareasRealizadas.HoraInicio"
            SQL = SQL & " FROM TareasRealizadas INNER JOIN (Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria ="
            SQL = SQL & " Trabajadores.idCategoria) ON TareasRealizadas.Trabajador = Trabajadores.IdTrabajador"
            SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
            'Solo el trabajador
            If Opcion = 1 Then SQL = SQL & " AND idTrabajador =" & Vector(BucleRevision)
            SQL = SQL & " ORDER by Categorias.idCategoria,Trabajadores.idTrabajador"
            
            RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                    'Los labels
                    Label11.Caption = "Trabajador: " & RT!Trabajador
                    Label11.Refresh
            
                    SQL = "UPDATE TareasRealizadas set "
                    'Total importe
                    TotalHoras = 0
                    
                    '--------------
                    'Importe 1
                    DEc = (RT!HoraS1 * RT!Importe1)
                    DEc = DEc + ((DEc * RT!porcantiguedad) / 100)
                    DEc = Round(DEc, 2)
                    TotalHoras = TotalHoras + DEc
                    SQL = SQL & " Importe1 = " & TransformaComasPuntos(CStr(DEc))
            
                    '--------------
                    'Importe 2
                    DEc = (RT!HoraS2 * RT!Importe2)
                    DEc = DEc + ((DEc * RT!porcantiguedad) / 100)
                    DEc = Round(DEc, 2)
                    TotalHoras = TotalHoras + DEc
                    SQL = SQL & " ,Importe2 = " & TransformaComasPuntos(CStr(DEc))
                    
                    
        ''''            'Si lo habilitamos
        ''''            '--------------
        ''''            'Importe 3
        ''''            DEc = (RT!HoraS3 * RT!importe3)
        ''''            DEc = DEc + ((DEc * RT!porcantiguedad) / 100)
        ''''            DEc = Round(DEc, 2)
        ''''            TotalHoras = TotalHoras + DEc
        ''''            SQL = SQL & " ,Importe3 = " & TransformaComasPuntos(CStr(DEc))
        ''''
        
                '----------
                'Total
                'con Seguridad social
                DEc = Round(((TotalHoras * RT!porcSS) / 100), 2)
                TotalHoras = TotalHoras + DEc
                SQL = SQL & " ,Total = " & TransformaComasPuntos(CStr(TotalHoras))
                
                SQL = SQL & " WHERE trabajador = " & RT!Trabajador & " AND Fecha = #" & Format(RT!Fecha, "yyyy/mm/dd") & "#"
                SQL = SQL & " AND HoraInicio = #" & Format(RT!horainicio, "hh:mm:ss") & "#"
                
                
                'Ejecutar SQL
                conn.Execute SQL
                
                'Siguiente
                RT.MoveNext
            Wend
            RT.Close


    Next BucleRevision  'Si la opcion =0 entonces
                        ' lo hace una vez el bucle, pero para todos los trabajadors
                        ' SI NO, lo hace 1 vez por trabajador por cada trabajador

End Sub



Private Function CodigoCorrecto(Trabajador As Boolean, Marcaje As String, Valor As Long) As Boolean
Dim SQL As String
Dim RT As ADODB.Recordset

    Set RT = New ADODB.Recordset
    CodigoCorrecto = False
    If Trabajador Then
        SQL = "Select idTrabajador from Trabajadores where numtarjeta = '" & Marcaje & "';"
    Else
        SQL = "Select idTarea from Tareas where tarjeta = '" & Marcaje & "';"
    End If

        
    RT.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RT.EOF Then
        CodigoCorrecto = True
        Valor = RT.Fields(0)
    Else
        Valor = -1
    End If
    RT.Close
    Set RT = Nothing
End Function



Private Sub EliminaEntradas()
Dim SQL As String
    SQL = "Delete from EntradaFichajes"
    SQL = SQL & " where (Fecha >= " & ImpFechaIni & ") AND (Fecha <= " & ImpFechaFin & ")"
    On Error Resume Next
    conn.Execute SQL
    If Err.Number <> 0 Then MuestraError Err.Number, "Eliminar entradas generadas" & Err.Description
End Sub

'Antes
'La hora de almuerzo sera la hora fin almuerzo y le kitaremos la duracion
'Ahora
'La hora de almuerzo es la de inicio, en Horafin tengo el fin
Private Sub QuitarHorasAlmuerzo(Trab As Long, Fecha As Date, HoraFinAlmuerzo As Date, HoraFinMerienda As Date, ByRef vH As CHorarios)
Dim RS As ADODB.Recordset
Dim SQL As String
Dim Total As Currency
Dim salida As Boolean
Dim Intervalo As Currency
Dim HTemp As Date

    On Error GoTo EQuitarHorasAlmuerzo
    Set RS = New ADODB.Recordset
    SQL = "SELECT * FROM TareasRealizadas WHERE Trabajador = " & Trab
    SQL = SQL & " AND Fecha = #" & Format(Fecha, "yyyy/mm/dd") & "#"
    SQL = SQL & " ORDER BY HoraInicio"
    RS.Open SQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If RS.EOF Then
        RS.Close
        GoTo EQuitarHorasAlmuerzo
    End If
    'Comprobamos el almuerzo
    Total = 0
    If vH.DtoAlm > 0 Then
        salida = False
        'Hora inicio almuerzo
        Total = vH.DtoAlm
        Do
            If RS!horainicio <= vH.HoraDtoAlm Then
                If RS!horafin < vH.HoraDtoAlm Then
                    'No es para kitar toavia
                    RS.MoveNext
                Else
                    If RS!horafin >= HoraFinAlmuerzo Then
                        'Se quita totalmente
                        QuitarTiempoTarea Trab, RS!Tarea, Fecha, RS!horainicio, Total, RS!HorasTrabajadas, RS!HoraS1
                        'Movemos siguiente
                        RS.MoveNext
                        'Salimos
                        salida = True
                    Else
                        'Quitamos la parte
                        Intervalo = CCur(DevuelveValorHora(RS!horafin))
                        Intervalo = Intervalo - CCur(DevuelveValorHora(vH.HoraDtoAlm))
                        QuitarTiempoTarea Trab, RS!Tarea, Fecha, RS!horainicio, Intervalo, RS!HorasTrabajadas, RS!HoraS1
                        Total = Total - Intervalo
                        'A la sigueinte tarea le kitamos el resto
                        RS.MoveNext
                    End If
                End If
            Else
                'hora inicio mayor k hora salida almuerzo, luego ya no kitamos almuerzo
                ' si la hora inicio mayor k la fecha fin almuerzo
                If RS!horainicio > HoraFinAlmuerzo Then
                    'La tarea no nos sirve
                    salida = True
                Else

                    Intervalo = CCur(DevuelveValorHora(HoraFinAlmuerzo))
                    Intervalo = Intervalo - CCur(DevuelveValorHora(RS!horainicio))
                    QuitarTiempoTarea Trab, RS!Tarea, Fecha, RS!horainicio, Intervalo, RS!HorasTrabajadas, RS!HoraS1
                    
                    'Salimos
                    salida = True
                End If
            End If
            If RS.EOF Then salida = True
        Loop Until salida
    End If
            

    '--------------------------------------------------------------------
    '
    'Ahora la merienda
    '
    Total = 0
    If vH.DtoMer > 0 Then
        RS.MoveFirst
        salida = False
        'Hora inicio almuerzo
        Total = vH.DtoMer
        Do
            If RS!horainicio <= vH.HoraDtoMer Then
                If RS!horafin < vH.HoraDtoMer Then
                    'No es para kitar toavia
                    RS.MoveNext
                Else
                    If RS!horafin >= HoraFinMerienda Then
                        'Se quita totalmente
                        QuitarTiempoTarea Trab, RS!Tarea, Fecha, RS!horainicio, Total, RS!HorasTrabajadas, RS!HoraS1
                        'Movemos siguiente
                        RS.MoveNext
                        'Salimos
                        salida = True
                    Else
                        'Quitamos la parte
                        Intervalo = CCur(DevuelveValorHora(RS!horafin))
                        Intervalo = Intervalo - CCur(DevuelveValorHora(vH.HoraDtoMer))
                        QuitarTiempoTarea Trab, RS!Tarea, Fecha, RS!horainicio, Intervalo, RS!HorasTrabajadas, RS!HoraS1
                        Total = Total - Intervalo
                        'A la sigueinte tarea le kitamos el resto
                        RS.MoveNext
                    End If
                End If
            Else
                'hora inicio mayor k hora salida almuerzo, luego ya no kitamos almuerzo
                ' si la hora inicio mayor k la fecha fin almuerzo
                If RS!horainicio > HoraFinMerienda Then
                    'La tarea no nos sirve
                    salida = True
                Else

                    Intervalo = CCur(DevuelveValorHora(HoraFinMerienda))
                    Intervalo = Intervalo - CCur(DevuelveValorHora(RS!horainicio))
                    QuitarTiempoTarea Trab, RS!Tarea, Fecha, RS!horainicio, Intervalo, RS!HorasTrabajadas, RS!HoraS1
                    
                    'Salimos
                    salida = True
                End If
            End If
            If RS.EOF Then salida = True
        Loop Until salida
    End If
            
    
            
            
EQuitarHorasAlmuerzo:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Sub

'En Quitar tendremos el tiempo k hay k kitar.
'En horas las horsas que hay. Si la diferencia es menor o igual k cero
'no kitamos nada
Private Sub QuitarTiempoTarea(ByRef Trab As Long, ByRef Tarea As Integer, ByRef Fecha As Date, ByRef HIni As Date, Total_A_Quitar As Currency, TotalHoras As Currency, HoraS1 As Currency)
Dim SQL As String
Dim H As Currency
    
    If TotalHoras - Total_A_Quitar <= 0 Then
        'Eliminamos la tarea
        SQL = "DELETE FROM TareasRealizadas"
    Else
        SQL = "UPDATE TareasRealizadas SET Horas1 ="
        H = HoraS1 - Total_A_Quitar
        SQL = SQL & TransformaComasPuntos(CStr(H))
        SQL = SQL & ", HorasTrabajadas = "
        H = TotalHoras - Total_A_Quitar
        SQL = SQL & TransformaComasPuntos(CStr(H))
    End If
    SQL = SQL & " WHERE Trabajador =" & Trab
    SQL = SQL & " AND Fecha = #" & Format(Fecha, "yyyy/mm/dd")
    SQL = SQL & "# AND HoraInicio = #" & Format(HIni, "hh:mm:ss")
    SQL = SQL & "# AND Tarea = " & Tarea
    On Error Resume Next
    conn.Execute SQL
    If Err.Number <> 0 Then MuestraError Err.Number, "Quitar timepo tarea"
End Sub




Private Sub RectificarMarcajesTareas()
    'Rectificamos si tiene rectificar
    Dim RTRa As ADODB.Recordset
    Dim RS As ADODB.Recordset
    Dim RT As ADODB.Recordset
    Dim Aux As String
    Dim Cad As String
    Dim h2 As String    'Hora fin
    Dim H3 As String    'Hora modificada
    Dim H1 As String

    Cad = "SELECT DISTINCT Trabajadores.Seccion"
    Cad = Cad & " FROM tmpTareasRealizadas ,Trabajadores WHERE "
    Cad = Cad & " tmpTareasRealizadas.Trabajador = Trabajadores.idTrabajador"

    Set RT = New ADODB.Recordset
    Set RS = New ADODB.Recordset
    Set RTRa = New ADODB.Recordset
    
    'Ponemos el cursortype para los Rs, para que no de fallo
    RS.CursorType = adOpenKeyset
    RS.LockType = adLockOptimistic
    
    RT.Open Cad, conn, , , adCmdText
    While Not RT.EOF
        Aux = RT.Fields(0)
        'Para cada seccion vemos si tiene Rs
        Cad = "SELECT * FROM ModificarFichajes "
        Cad = Cad & " WHERE IdSeccion= " & Aux
        RS.Open Cad, conn, , , adCmdText
        While Not RS.EOF
            Label11.Caption = "Rectifcando marcajes TAREA para seccion: " & Aux & " /" & RS.Fields(0)
            Label11.Refresh
            H1 = "#" & Format(RS.Fields(2), "hh:mm") & "#"
            h2 = "#" & Format(RS.Fields(3), "hh:mm") & "#"
            H3 = "#" & Format(RS.Fields(4), "hh:mm") & "#"
            'Creamos la consulta de acutalizacion
            'Para cada Rs modificamos la tabla
            Cad = "SELECT DISTINCT(tmpTareasRealizadas.Trabajador)" ', Trabajadores.Seccion"
            Cad = Cad & " FROM tmpTareasRealizadas INNER JOIN Trabajadores ON tmpTareasRealizadas.Trabajador = Trabajadores.IdTrabajador"
            Cad = Cad & " WHERE (((Trabajadores.Seccion)=" & Aux & "));"
            'Abrimos
            RTRa.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            Cad = "UPDATE tmpTareasRealizadas"
            Cad = Cad & " SET Hora = " & H3
            Cad = Cad & " WHERE Hora >= " & H1
            Cad = Cad & " AND Hora <= " & h2
            Cad = Cad & " AND Trabajador ="
            While Not RTRa.EOF
                'Ejecutamos el SQL
                conn.Execute Cad & RTRa.Fields(0)
                RTRa.MoveNext
            Wend
            RTRa.Close
            RS.MoveNext
        Wend
        'Cerramos el recordset de modificar marcjaes
        RS.Close
        'Movemos al siguiente
        RT.MoveNext
    Wend
    'Cerramos los recordset
    RT.Close
    Set RT = Nothing
    Set RS = Nothing
    Set RTRa = Nothing
End Sub



'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
'
'       R E C T I F I C A C I O N      D E       M A R C A J E S
'
'-------------------------------------------------------------------------------
'-------------------------------------------------------------------------------
Private Function RectificacionDeMarcajes() As Integer
Dim RsSeccion As ADODB.Recordset
Dim Recortes As ADODB.Recordset
Dim vRs As ADODB.Recordset
Dim Cad As String
Dim Aux As String
Dim H1 As String    'Hora inicio
Dim h2 As String    'Hora fin
Dim H3 As String    'Hora modificada
Dim miHora As Date  'Una hora
Dim HoraAnt As Date
Dim Hora As Date
Dim i As Integer
Dim Redondeo As Integer
Dim Trabajador As Integer
Dim Fecha As Date
Dim AjusteE As Integer
Dim AjusteS As Integer
Dim Minutos As Integer


On Error GoTo ErrorRectificacionDeMarcajes
RectificacionDeMarcajes = 1

Cad = "SELECT DISTINCT Trabajadores.Seccion,  Secciones.RedondearCadaTicaje"
Cad = Cad & " FROM EntradaFichajes INNER JOIN (Secciones INNER JOIN Trabajadores ON Secciones.IdSeccion = Trabajadores.Seccion) ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador"


Set RsSeccion = New ADODB.Recordset
Set Recortes = New ADODB.Recordset
'Ponemos el cursortype para los recortes, para que no de fallo
Recortes.CursorType = adOpenKeyset
Recortes.LockType = adLockOptimistic

RsSeccion.Open Cad, conn, , , adCmdText
While Not RsSeccion.EOF
    Aux = RsSeccion.Fields(0)
    Label10.Caption = "Rectifcando marcajes para seccion: " & Aux
    Label10.Refresh
    
    'Para cada seccion vemos si tiene recortes
    If RsSeccion!RedondearCadaTicaje = 0 Then
        'Tenemos k recortar en funcion de lo k haya puewto
        'En ajuste manuales
        Cad = "SELECT * FROM ModificarFichajes "
        Cad = Cad & " WHERE IdSeccion= " & Aux
        Recortes.Open Cad, conn, , , adCmdText
        While Not Recortes.EOF
            
            H1 = "#" & Format(Recortes.Fields(2), "hh:mm") & "#"
            h2 = "#" & Format(Recortes.Fields(3), "hh:mm") & "#"
            'H3 = "#" & Format(Recortes.Fields(4), "hh:mm") & "#"
            H3 = Format(Recortes.Fields(4), "hh:mm")
            'Label
            Label11.Caption = H1 & " - " & h2 & "   --> " & H3
            Label11.Refresh
            'Creamos la consulta de acutalizacion
            'Para cada recortes modificamos la tabla

            
            'SELECT EntradaFichajes.idTrabajador, Trabajadores.IdTrabajador, Secciones.IdSeccion
            'FROM Secciones INNER JOIN (EntradaFichajes INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador) ON Secciones.IdSeccion = Trabajadores.Seccion
            'WHERE (((Secciones.IdSeccion)=1));
            Set vRs = New ADODB.Recordset
            Cad = "SELECT EntradaFichajes.*"
            Cad = Cad & " FROM Secciones INNER JOIN (EntradaFichajes INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador) ON Secciones.IdSeccion = Trabajadores.Seccion"
            Cad = Cad & " WHERE EntradaFichajes.Hora>=" & H1 & " AND EntradaFichajes.Hora<" & h2 & " AND Secciones.IdSeccion=" & RsSeccion!Seccion & " AND "
            Cad = Cad & " Fecha >=" & ImpFechaIni & " AND Fecha<=" & ImpFechaFin
            Set vRs = New ADODB.Recordset
            vRs.Open Cad, conn, adOpenKeyset, adLockPessimistic, adCmdText
            While Not vRs.EOF
                vRs!Hora = H3
                vRs.Update
                'Siguiente
                vRs.MoveNext
            Wend
            vRs.Close

            Recortes.MoveNext
        Wend
        'Cerramos el recordset de modificar marcjaes
        Recortes.Close
    
    Else
        'Podremos o redondear segun valor o ajustes de entrada salida
        If RsSeccion!RedondearCadaTicaje < 3 Then
            Label11.Caption = "Ajuste por fraccion"
            Label11.Refresh
            '------------------------
            'Ajustes por redondeo
            '-----------------------
          'Tenemos k redondear a cuartos, o a media hora en funcion del valor en datos de empresa
          'Entonces, a partir de las doce de la mañana vamos haciendo hasta las 11:30 de la noche
          If RsSeccion!RedondearCadaTicaje = 1 Then
              Aux = "15"
          Else
              Aux = "30"
          End If
          
          'Primero vemos los minutos para cortar los intervalos
          H1 = DevuelveDesdeBD("minutosredondeo", "Empresas", "idEmpresa", 1, "N")
          If H1 = "" Then H1 = 0
          
          If Val(H1) = 0 Then
              'MsgBox "No hay minutos redondeo.  No se ajustara ningun ticaje", vbExclamation
              Set RsSeccion = Nothing
              Exit Function
          End If
          Redondeo = Val(H1)
          
          'Hacemos el primer cambio k sera desde las 0:00 hasta los minutos redondeo son las 12
          H1 = "00:00"
          H3 = H1
          miHora = CDate(H1)
          miHora = DateAdd("n", Redondeo, miHora)
          h2 = miHora
              H1 = "#" & Format(H1, "hh:mm") & "#"
              h2 = "#" & Format(h2, "hh:mm") & "#"
              H3 = Format(H3, "hh:mm")
              
             
              
            Set vRs = New ADODB.Recordset
            Cad = "SELECT EntradaFichajes.*"
            Cad = Cad & " FROM Secciones INNER JOIN (EntradaFichajes INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador) ON Secciones.IdSeccion = Trabajadores.Seccion"
            Cad = Cad & " WHERE EntradaFichajes.Hora>=" & H1 & " AND EntradaFichajes.Hora<" & h2 & " AND Secciones.iDseccion=" & RsSeccion!Seccion & " AND "
            Cad = Cad & " Fecha >=" & ImpFechaIni & " AND Fecha<=" & ImpFechaFin
            Set vRs = New ADODB.Recordset
            
            
            vRs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not vRs.EOF
                vRs!Hora = H3
                vRs.Update
                'Siguiente
                vRs.MoveNext
            Wend
            vRs.Close
              
              
              
              
          
          'El bucle
              Hora = CDate("00:00")
              Hora = DateAdd("n", Val(Aux), Hora)
              HoraAnt = DateAdd("n", 1, miHora)
              While Hora <= CDate("23:45")
    
                      miHora = DateAdd("n", Redondeo, Hora)
                  
                      H1 = "#" & Format(HoraAnt, "hh:mm:ss") & "#"
                      h2 = "#" & Format(miHora, "hh:mm") & ":59#"
                      'H3 = "#" & Format(Hora, "hh:mm") & "#"
                      H3 = Format(Hora, "hh:mm") & ":00"
                      
                      Label11.Caption = H1 & " - " & h2 & "   --> " & H3
                      Label11.Refresh
                      Cad = "SELECT EntradaFichajes.*"
                      Cad = Cad & " FROM Secciones INNER JOIN (EntradaFichajes INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador) ON Secciones.IdSeccion = Trabajadores.Seccion"
                      Cad = Cad & " WHERE EntradaFichajes.Hora>=" & H1 & " AND EntradaFichajes.Hora<" & h2 & " AND Secciones.IdSeccion=" & RsSeccion!Seccion & " AND "
                      Cad = Cad & " Fecha >=" & ImpFechaIni & " AND Fecha<=" & ImpFechaFin
                      Set vRs = New ADODB.Recordset
                      vRs.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
                        While Not vRs.EOF
                            vRs!Hora = H3
                            vRs.Update
                            'Siguiente
                            vRs.MoveNext
                        Wend
                        vRs.Close
                          
                      'Subimos hora y hora post
                      Hora = DateAdd("n", Val(Aux), Hora)
                      HoraAnt = DateAdd("n", 1, miHora)
              Wend
          
          'Hacemos el ultimo, desde las 12-algo de la noche hasta las 23:59 son las 23:59
                      H1 = "#" & Format(HoraAnt, "hh:mm") & "#"
                      h2 = "#23:59#"
                      
                      Cad = "SELECT EntradaFichajes.*"
                      Cad = Cad & "  FROM Secciones INNER JOIN (EntradaFichajes INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador"
                      Cad = Cad & " = Trabajadores.IdTrabajador) ON Secciones.IdSeccion = Trabajadores.Seccion"
                      Cad = Cad & " WHERE EntradaFichajes.Hora>=" & H1 & " AND "
                      Cad = Cad & " Secciones.IdSeccion=" & RsSeccion!Seccion & " AND "
                      Cad = Cad & " Fecha >=" & ImpFechaIni & " AND Fecha<=" & ImpFechaFin
                      Set vRs = New ADODB.Recordset
                      vRs.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        While Not vRs.EOF
                            vRs!Hora = H3
                            vRs.Update
                            'Siguiente
                            vRs.MoveNext
                        Wend
                        vRs.Close
                      'Ejecutamos el SQL
                      conn.Execute Cad
                      
                      
                      
                      
                      
                      
                      
            Else
                         
                   '----------------------------
                   '  AJUSTES por entrada salida
                   '----------------------------
                      
                   'Cojeremos para cada trabajador, cada fecha e iremos viendo entrada salida
                   'Los marcajes, y por conteo iremos viendo
                   ' Entrada--> ajuste entrada.... salida---> ajuste salida
                
                         
                    If RsSeccion!RedondearCadaTicaje = 3 Then
                        Aux = "15"
                    Else
                        Aux = "30"   'Entradas salidas cada media hora
                    End If
                             
                   
                   
          
                
                    'Primero vemos los ajustes. Medias horas, cuartos
                    Cad = "AjusteSalida"
                    AjusteE = Val(DevuelveDesdeBD("AjusteEntrada", "Empresas", "IdEmpresa", "1", "N", Cad))
                    AjusteS = Val(Cad)
                    

                    Cad = "SELECT EntradaFichajes.*"
                    Cad = Cad & " FROM Secciones INNER JOIN (EntradaFichajes INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador) ON Secciones.IdSeccion = Trabajadores.Seccion"
                    Cad = Cad & " WHERE "
                    Cad = Cad & " Secciones.IdSeccion=" & RsSeccion!Seccion & " AND "
                    Cad = Cad & " Fecha >=" & ImpFechaIni & " AND Fecha<=" & ImpFechaFin


                    Cad = Cad & " ORDER By EntradaFichajes.idTrabajador,Fecha,Hora"
                    Trabajador = -1
                   
                    Recortes.Open Cad, conn, , , adCmdText
                    While Not Recortes.EOF
                        If Trabajador <> Recortes!idTrabajador Then
                            'label
                             Trabajador = Recortes!idTrabajador
                             Fecha = "01/01/1900"
                             Me.Refresh
                        End If
                                      
                                      
                        If Fecha <> Recortes!Fecha Then
                            'label
                            Label11.Caption = Recortes!idTrabajador & " - " & Recortes!Fecha
                            Label11.Refresh
                            i = 0
                            Fecha = Recortes!Fecha
                        End If
                                
                        
                        If (i Mod 2) = 0 Then
                            'Entrada
                            Hora = HoraRectificada(Recortes!Hora, AjusteE, CInt(Aux))
                        Else
                            'Salida
                            Hora = HoraRectificada(Recortes!Hora, AjusteS, CInt(Aux))
                        End If
                    
                        Recortes!Hora = Hora
                        Recortes.Update
                         'Siguiente
                         i = i + 1
                         Recortes.MoveNext
                     Wend
                     Recortes.Close
                      
            End If
    End If
    
    'Movemos al siguiente
    RsSeccion.MoveNext
Wend
'Cerramos los recordset
RsSeccion.Close
Set Recortes = Nothing
Set RsSeccion = Nothing
'Todo correcto
RectificacionDeMarcajes = 0
Exit Function
ErrorRectificacionDeMarcajes:
    MuestraError Err.Number
    RectificacionDeMarcajes = 1
End Function




'Private Function HoraEntrada(Hora As Date, Ajuste As Integer, FraccionHora As Integer) As Date
'Dim Nueva As Date
'Dim Minu As Integer
'Dim Salir As Boolean
'
'
'        HoraEntrada = Hora
'        Nueva = CDate(Hour(Hora) & ":00")
'        Salir = False
'        Do
'            Minu = DateDiff("n", Nueva, Hora)
'            If Minu > Ajuste Then
'                Nueva = DateAdd("n", FraccionHora, Nueva)
'            Else
'                Salir = True
'                'Llega antes k la entrada
'                HoraEntrada = Nueva
'            End If
'        Loop Until Salir
'
'End Function
'
'
'
'Private Function HoraSalida(Hora As Date, Ajuste As Integer, FraccionHora As Integer) As Date
'Dim Nueva As Date
'Dim Minu As Integer
'Dim Salir As Boolean
'
'
'        HoraSalida = Hora
'        Nueva = CDate(Hour(Hora) & ":00")
'        Salir = False
'        Do
'            Minu = DateDiff("n", Hora, Nueva)
'            If Minu > Ajuste Then
'                Nueva = DateAdd("n", FraccionHora, Nueva)
'            Else
'                Salir = True
'                If Minu >= 0 Then
'                    'Llega antes k la entrada
'                    HoraSalida = Nueva
'                Else
'                    HoraSalida = DateAdd("n", -1 * FraccionHora, Nueva)
'                End If
'            End If
'        Loop Until Salir
'
'End Function

'Es lo mismo para la entrada k para la salida.
'Lo k cambia es el ajuste
Private Function HoraRectificada(Hora As Date, Ajuste As Integer, FraccionHora As Integer) As Date
Dim Nueva As Date
Dim Minu As Integer
Dim Salir As Boolean


        HoraRectificada = Hora
        Nueva = CDate(Hour(Hora) & ":00")
        Salir = False
        Do
            Minu = DateDiff("n", Nueva, Hora)
            If Minu > Ajuste Then
                If DateDiff("n", Nueva, CDate("23:59")) <= 30 Then
                    HoraRectificada = CDate("23:59")
                    Exit Function
                End If
                Nueva = DateAdd("n", FraccionHora, Nueva)
            Else
                Salir = True
                HoraRectificada = Nueva
            End If
        Loop Until Salir
        
End Function

