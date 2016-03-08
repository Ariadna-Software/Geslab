Attribute VB_Name = "LibPresencia"
Option Explicit

'Public Const CadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\ControlPresencia\BDatos.mdb;Persist Security Info=False"
Public conn As Connection
Public mConfig As CFGControl
Public BaseForm As Integer 'Tendremos la base donde mostrar
                            'los formularios
                            
Public MostrarErrores As Boolean

'En estas variables tendremos las fechas de inicio
'y fin para la importacion del fichero
Public ImpFechaIni As String
Public ImpFechaFin As String

'Para los redondeos de la merienda y del almuerzo
'Son comunes para todos los procesamarcaje...
Public PrimerTicaje As Date
Public UltimoTicaje As Date


Public MiEmpresa As CEmpresas

Public VariableCompartida As String

Public FormatoFecha As String

Public FormatoImporte As String



Public myMonday As Integer  'PAra el primer dia del frmcal

Public vUsu As Usuario


'Estas variable :
'   -  Es para cuando estamos revisando los marcajes, si
'       tiene produccion, se generea unos dtos a partir de la tabla de
'       kimaldi. Ese proceso es lento, con lo cual, la primera vez
'       lo genero para una fecha, y luego los guardo en una temporal.
'       Cuando vuelvo a necesitar la produccion, si es para la misma fecha
'       NO hace falta generarlo, con un simple Select INTO desde esta
'       nueva temporal sobre la tmpTareasRea
Public FechaHazProduccion As String




'Para los informes
Public CadParam As String
Public NParam As Integer
Public nOpcion As Integer


Public Sub Main()
On Error Resume Next
Screen.MousePointer = vbHourglass

'Vemos si ya se esta ejecutando
If App.PrevInstance Then
    MsgBox "Ya se está ejecutando el programa ARIPRES (Tenga paciencia).", vbCritical
    Screen.MousePointer = vbDefault
    Exit Sub
End If


Set mConfig = New CFGControl
If mConfig.Leer = 1 Then
    frmConfig.Show vbModal
    End
End If
If mConfig.BaseDatos = "" Then
    MsgBox "Falta cadena de conexion del la base de datos", vbCritical
    End
End If



'Ahora abrimos por ODBC luego no quiero la cadena de conexion
'si no el nombre del DSN correspondiente y por lo tanto
'no abro con un connectionstring sino del siguiente modo
' conn.open  nombreDSN
' Hecahs unas pruebas y comprueba que con DSN
'no se puede utilizar Begintrans, committrans y rollback
Set conn = New ADODB.Connection
If Err.Number <> 0 Then
    MuestraError Err.Number, "Creando Objeto CONN ADODB.Connection"
    End
    Exit Sub
End If
conn.ConnectionString = mConfig.BaseDatos
conn.CursorLocation = adUseServer
conn.Open
If Err.Number <> 0 Then
    MuestraError Err.Number, Err.Description
    MsgBox "Error en la cadena de conexion" & vbCrLf & mConfig.BaseDatos, vbCritical
    End
End If



FormatoImporte = "#,###,##0.00"
FormatoFecha = "yyyy/mm/dd"
FechaHazProduccion = ""

Set MiEmpresa = New CEmpresas
MiEmpresa.Leer (1)   'La uno pq solo trabajo con una empresa



'-- PARA HUELLA
'If MiEmpresa.QueEmpresa = 1 Or MiEmpresa.QueEmpresa = 2 Or MiEmpresa.QueEmpresa = 3 Then
' Enero 2012
If MiEmpresa.QueEmpresa >= 1 And MiEmpresa.QueEmpresa <= 4 Then
    'Es BELGIDA. Trabaja con lectores BIometricos
    'Y no hace nada mas
    'Alzira TB entra aqui
    'Catadau. TB
    ' La bD esta en el ODBC driver de MDB y se llama accGestorHuella
    AbrirBaseDatos
End If



    'En PICASSENT
    If MiEmpresa.QueEmpresa = 0 Then
           ' mConfig.TCP3 = False 'YA no lleva TCP3  DERIAMOS COCAMBIAR LA VARIABLE para no utilizarla mas
            
        
    End If

'Fijamos el primer dia de la semana para el frmCal
FijarPrimerDiaSemana

'Veremos si esta registrado o no el programa
'Cargamos en memoria los dos formularios
'Veremos si esta registrado o no el programa
'Cargamos en memoria los dos formularios
Screen.MousePointer = vbHourglass
PonerValoresConstantes_BD

VariableCompartida = ""
frmIdentifica.Show vbModal
If VariableCompartida = "" Then
    'No se ha identificado
    conn.Close
    End
End If

Load frmPpal1
'Load frmLLave

'Vemos el registro
'If frmLLave.ActiveLock1.RegisteredUser Then
'    Unload frmLLave
'    Else
'        frmLLave.Show vbModal
'End If

 'frmMain.Show
frmPpal1.Show
End Sub


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosComas(CADENA As String) As String
Dim i As Integer
Do
    i = InStr(1, CADENA, ".")
    If i > 0 Then
        CADENA = Mid(CADENA, 1, i - 1) & "," & Mid(CADENA, i + 1)
    End If
    Loop Until i = 0
TransformaPuntosComas = CADENA
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaPuntosHoras(CADENA As String) As String
Dim i As Integer
Do
    i = InStr(1, CADENA, ".")
    If i > 0 Then
        CADENA = Mid(CADENA, 1, i - 1) & ":" & Mid(CADENA, i + 1)
    End If
    Loop Until i = 0
TransformaPuntosHoras = CADENA
End Function


'Cambia los puntos de los numeros decimales
'por comas
Public Function TransformaComasPuntos(CADENA As String) As String
    Dim i As Integer
    Do
        i = InStr(1, CADENA, ",")
        If i > 0 Then
            CADENA = Mid(CADENA, 1, i - 1) & "." & Mid(CADENA, i + 1)
        End If
        Loop Until i = 0
    TransformaComasPuntos = CADENA
End Function




Public Function DBLet(vData As Variant, Optional Tipo As String) As Variant
    If IsNull(vData) Then
        DBLet = ""
        If Tipo <> "" Then
            Select Case Tipo
                Case "T"
                    DBLet = ""
                Case "N"
                    DBLet = 0
                Case "F"
                    DBLet = "0:00:00"
                Case "D"
                    DBLet = 0
                Case "B"  'Boolean
                    DBLet = False
                Case Else
                    DBLet = ""
            End Select
        End If
    Else
        DBLet = vData
    End If
End Function



Public Function CmpCam(CADENA As String, Valor As Variant) As Boolean
    '------------------------------------------------------------------------
    'CmpCam: Comprueba la validez de un dato 'Valor' en función de una
    'cadena de comprobación 'Cadena' con los valores
    '   V1|V2|V3|V4|V5|
    '   V1: Nombre del campo, para mensajes.
    '   V2: Tipo de campo N=Numérico,T=Texto,F=Fecha(dd/mm/aaaa),H=Hora(hh:mm)
    '   V3: S=Se permiten nulos, N=No se permiten
    '   V4: El valor ha de ser mayor o igual que él, puede ir vacío
    '   V5: El valor a de ser menor o igual que él, puede ir vacío
    'Si V4 o V5 van vacíos no se hace comprobación de rangos
    '------------------------------------------------------------------------
    Dim V(5) As Variant
    Dim i As Integer
    Dim i2 As Integer
    Dim mCadena As String
    Dim Mc As String
    Dim Faltantes As Boolean
    Dim Mens As String
    
    
    '-- Limpiamos los campos
    For i = 1 To 5
        V(i) = "#"
    Next i
    '-- Cargamos los datos
    mCadena = ""
    Mc = ""
    i2 = 0
    For i = 1 To Len(CADENA)
        Mc = Mid(CADENA, i, 1)
        If Mc = "|" Then
            i2 = i2 + 1
            V(i2) = mCadena
            mCadena = ""
        Else
            mCadena = mCadena & Mc
        End If
    Next i
    '-- Comprobamos que no se han dejado ningún campo
    For i = 1 To 5
        If V(i) = "#" Then Faltantes = True
    Next i
    If Faltantes Then
        Mens = "Faltan parámetros en la etiqueta de comprobación"
        MsgBox Mens, vbInformation, "Comprobador de campos"
        Exit Function
    End If
    '-- Comenzamos las comprobaciones
    If V(1) = "" Then
        Mens = "El nombre de campo no puede estar vacío"
        MsgBox Mens, vbInformation, "Comprobador de campos"
        Exit Function
    End If
    
    
    'Comprobamos si permite nulos
    If V(3) = "N" Then
        '-- No se permiten nulos.
        If Valor = "" Then
            Mens = "El valor de " & V(1) & " no puede ser nulo."
            MsgBox Mens, vbInformation, "Comprobador de campos"
            Exit Function
        End If
    End If
    
    Select Case V(2)
        Case "N"
            If Valor <> "" Then
                If Not IsNumeric(Valor) Then
                    Mens = "El valor de " & V(1) & " debe ser numérico."
                    MsgBox Mens, vbInformation, "Comprobador de campos"
                    Exit Function
                End If
            End If
        Case "T"
        
        Case "F"
                If Valor <> "" Then
                    If Not IsDate(Valor) Then
                        Mens = "El valor de " & V(1) & " debe ser una fecha (dd/mm/aaaa)."
                        MsgBox Mens, vbInformation, "Comprobador de campos"
                        Exit Function
                    End If
                End If
        Case "H"
            If Not IsDate(Valor) Then
                Mens = "El valor de " & V(1) & " debe ser una hora (hh:mm)."
                MsgBox Mens, vbInformation, "Comprobador de campos"
                Exit Function
            End If
        Case Else
            Mens = "Tipo de campo desconocido"
            MsgBox Mens, vbInformation, "Comprobador de campos"
            Exit Function
    End Select
    
    'Si es numerico y NO permitido nulos enonces
    If V(2) = "N" And V(3) = "N" Then
        If V(4) <> "" Then
          If IsNumeric(V(4)) Then
              If Val(Valor) < Val(V(4)) Then
                  Mens = "El valor de " & V(1) & " debe ser mayor o igual que " & V(4)
                  MsgBox Mens, vbInformation, "Comprobador de campos"
                  Exit Function
              End If
          End If
        End If
    
       If V(5) <> "" Then
            If IsNumeric(V(5)) Then
              If Val(Valor) > Val(V(5)) Then
                  Mens = "El valor de " & V(1) & " debe ser mayor o igual que " & V(5)
                  MsgBox Mens, vbInformation, "Comprobador de campos"
                  Exit Function
              End If
            End If
        End If
    End If
    CmpCam = True
End Function





Public Function DevuelveNumeroTikadas(vIdTrabajador As Long) As Byte
Dim rs As ADODB.Recordset
Dim sql As String

    DevuelveNumeroTikadas = 0
    Set rs = New ADODB.Recordset
    sql = "SELECT Horarios.NumTikadas" & _
        " FROM Horarios,Trabajadores WHERE " & _
        " Horarios.IdHorario = Trabajadores.IdHorario " & _
        " AND Trabajadores.IdTrabajador=" & vIdTrabajador
    rs.Open sql, conn, , , adCmdTable
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then _
            DevuelveNumeroTikadas = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
End Function


Public Function DevuelveTextoIncidencia(vId As Integer, Optional ByRef vSigno As Single) As String
Dim rs As ADODB.Recordset
Dim sql As String

    DevuelveTextoIncidencia = ""
    vSigno = 0
    Set rs = New ADODB.Recordset
    sql = "SELECT * From Incidencias " & _
        " WHERE Incidencias.IdInci =" & vId
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveTextoIncidencia = rs.Fields(1)
            If rs!excesodefecto Then
                vSigno = -1
                Else
                vSigno = 1
            End If
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function DevuelveNombreEmpresa(vId As Long) As String
Dim rs As ADODB.Recordset
Dim sql As String

    DevuelveNombreEmpresa = ""
    Set rs = New ADODB.Recordset
    sql = "SELECT NomEmpresa From Empresas " & _
        " WHERE IdEmpresa =" & vId
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveNombreEmpresa = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
End Function


Public Function DevuelveNombreHorario(vId As Long) As String
Dim rs As ADODB.Recordset
Dim sql As String

    DevuelveNombreHorario = ""
    Set rs = New ADODB.Recordset
    sql = "SELECT Nomhorario From Horarios " & _
        " WHERE IdHorario =" & vId
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveNombreHorario = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
End Function


Public Function DevuelveCodigo(vNUmTar) As Long
Dim rs As ADODB.Recordset
Dim sql As String
    DevuelveCodigo = -1
    Set rs = New ADODB.Recordset
    sql = "SELECT idTrabajador From Trabajadores " & _
        " WHERE NumTarjeta ='" & vNUmTar & "'"
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveCodigo = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
End Function

'Public Sub DevuelveMargenesHorarios(ByRef HEx As Date, ByRef HDef As Date, vIdTrabajador As Long)
'Dim Rs As adodb.Recordset
'Dim Sql As String
'Dim v1 As Single, v2 As Single
'
'    HEx = "0:00:00"
'    HDef = "0:00:00"
'    Set Rs = New adodb.Recordset
'    Sql = "SELECT  Empresas.MaxRetraso, Empresas.MaxExceso" & _
'        " FROM Empresas ,Trabajadores WHERE Empresas.IdEmpresa = Trabajadores.IdEmpresa" & _
'        " AND Trabajadores.IdTrabajador=" & vIdTrabajador
'    Rs.Open Sql, Conn, , , adCmdText
'    If Not Rs.EOF Then
'        v1 = Rs.Fields(0)
'        v2 = Rs.Fields(1)
'        HEx = DevuelveHora(v1)
'        HDef = DevuelveHora(v2)
'    End If
'    Rs.Close
'    Set Rs = Nothing
'End Sub



Public Function DevuelveINC_MARCAJE(vTrabajador As Long) As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sql As String

    sql = "SELECT Empresas.IncMarcaje " & _
         "FROM Empresas ,Trabajadores WHERE Empresas.IdEmpresa = Trabajadores.IdEmpresa" & _
         " AND Trabajadores.IdTrabajador=" & vTrabajador
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveINC_MARCAJE = DBLet(rs.Fields(0), "N")
        Else
            DevuelveINC_MARCAJE = 0
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function EsBajaTrabajo(vTrabajador As Long) As Boolean
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sql As String

    sql = "SELECT Trabajadores.FecBaja" & _
         " FROM Trabajadores WHERE Trabajadores.IdTrabajador=" & vTrabajador
    rs.Open sql, conn, , , adCmdText
    If rs.EOF Then
            EsBajaTrabajo = True
        Else
            EsBajaTrabajo = Not IsNull(rs.Fields(0))
    End If
    rs.Close
    Set rs = Nothing
End Function


Public Function DevuelveINC_RETRASO(vTrabajador As Long) As Integer
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim sql As String

    sql = "SELECT Empresas.IncRetraso" & _
         " FROM Empresas ,Trabajadores WHERE Empresas.IdEmpresa = Trabajadores.IdEmpresa" & _
         " AND Trabajadores.IdTrabajador=" & vTrabajador
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveINC_RETRASO = DBLet(rs.Fields(0), "N")
        Else
            DevuelveINC_RETRASO = -1
    End If
    rs.Close
    Set rs = Nothing
End Function

'Public Function DevuelveINC_HORAEXTRA(vTrabajador As Long) As Integer
'Dim Rs As ADODB.Recordset
'Set Rs = New ADODB.Recordset
'Dim SQL As String
'
'    SQL = "SELECT Empresas.IncHoraExtra" & _
'         " FROM Empresas ,Trabajadores WHERE Empresas.IdEmpresa = Trabajadores.IdEmpresa" & _
'         " AND Trabajadores.IdTrabajador=" & vTrabajador
'    Rs.Open SQL, Conn, , , adCmdText
'    If Not Rs.EOF Then
'            DevuelveINC_HORAEXTRA = DBLet(Rs.Fields(0), "N")
'        Else
'            DevuelveINC_HORAEXTRA = -1
'    End If
'    Rs.Close
'    Set Rs = Nothing
'End Function



Public Function devuelveNombreTrabajador(vId As Long, Optional ByRef vHor As Integer) As String
Dim rs As ADODB.Recordset
Dim sql As String

    devuelveNombreTrabajador = ""
    Set rs = New ADODB.Recordset
    sql = "SELECT NomTrabajador,IdHorario From Trabajadores " & _
        " WHERE idTrabajador =" & vId
    rs.Open sql, conn, , , adCmdText
    vHor = 0
    If Not rs.EOF Then
            vHor = rs.Fields(1)
            devuelveNombreTrabajador = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
End Function


Public Function devuelveNombreNTarjeta(vId As Integer, Optional ByRef vHor As Integer) As String
Dim rs As ADODB.Recordset
Dim sql As String

    devuelveNombreNTarjeta = ""
    Set rs = New ADODB.Recordset
    sql = "SELECT NumTarjeta From Trabajadores " & _
        " WHERE idTrabajador =" & vId
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then
            devuelveNombreNTarjeta = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing
End Function

Public Function DevuelveNombreSeccion(vId As Long) As String
Dim rs As ADODB.Recordset
Dim sql As String

    DevuelveNombreSeccion = ""
    Set rs = New ADODB.Recordset
    sql = "SELECT Nombre FROM Secciones  " & _
        " WHERE idSeccion =" & vId
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then DevuelveNombreSeccion = rs.Fields(0)
    rs.Close
    Set rs = Nothing
End Function

Public Function DevuelveTipoControl(vId As Long) As Byte
Dim rs As ADODB.Recordset
Dim sql As String

    DevuelveTipoControl = 127
    Set rs = New ADODB.Recordset
    sql = "SELECT Control From Trabajadores " & _
        " WHERE idTrabajador =" & vId
    rs.Open sql, conn, , , adCmdText
    If Not rs.EOF Then
            DevuelveTipoControl = rs.Fields(0)
    End If
    rs.Close
    Set rs = Nothing


End Function







Public Function ComprobarMarcajesCorrectos(FI As Date, FF As Date, Correctos As Boolean) As Byte
Dim rs As ADODB.Recordset
Dim cad As String
Dim C As Long
    
    ComprobarMarcajesCorrectos = 127
    C = 0
    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    'SQL. Marcajes incorrectos entre las dos fechas
    cad = "Select count(Entrada) "
    cad = cad & " From Secciones, Trabajadores, Marcajes"
    cad = cad & " WHERE  Secciones.IdSeccion = Trabajadores.Seccion AND"
    cad = cad & " Trabajadores.idTrabajador = Marcajes.idTrabajador"
    'Esto era para llevar a euroagr
    'And Secciones.Nominas = True"
    cad = cad & " AND Trabajadores.IdEmpresa=" & MiEmpresa.IdEmpresa
    cad = cad & " AND Fecha>=#" & Format(FI, "yyyy/mm/dd") & "#"
    cad = cad & " AND Fecha<=#" & Format(FF, "yyyy/mm/dd") & "#"
    cad = cad & " AND Correcto="
    If Correctos Then
        cad = cad & " True"
    Else
        cad = cad & " False"
    End If
    rs.Open cad, conn, , , adCmdText
    If Not rs.EOF Then
        If Not IsNull(rs.Fields(0)) Then C = rs.Fields(0)
    End If
    rs.Close
    'Si c>0 entonces tiene marcajes incorrectos
    If C > 0 Then
        ComprobarMarcajesCorrectos = 1
        Else
            ComprobarMarcajesCorrectos = 0
    End If
End Function










'Public Function CopiaSeguridad() As Boolean
'Dim cad As String
'Dim i As Integer
'Dim j As Integer
'
'
'
'
'On Error GoTo ErrCopiaSeg
''Copiamos la BD
''primero tenemos que saber donde la vamos a copiar
''en la carpeta BACKUP dentro de app.path
'If Dir(App.Path & "\Backup", vbDirectory) = "" Then _
'    MkDir App.Path & "\Backup"
''Obtenemos el nombre de la BD  'Source=C:\ControlPresencia\BDatos.mdb;
'i = InStr(1, mConfig.BaseDatos, "Data Source=")
'If i = 0 Then
'    MsgBox "Error en la cadena de conexion.", vbExclamation
'    Exit Function
'End If
'j = InStr(i, mConfig.BaseDatos, ";")
'If j = 0 Then
'    MsgBox "Error en la cadena de conexion.", vbExclamation
'    Exit Function
'End If
'i = i + 12 'Por que hemos buscado Data source=
'cad = Mid(mConfig.BaseDatos, i, j - i)
''Ya tenemos el nombre de la BD
''Ahora la vemos si esta, por si acaso
'If Dir(cad) = "" Then
'    MsgBox "No se encuentra la BD de la aplicación.", vbExclamation
'    Exit Function
'End If
'
''Copiamos la BD
'Dim aux
'aux = App.Path & "\Backup\BD" & Format(Now, "yymmdd") & ".mdb"
'FileCopy cad, aux
'
'
'
'Exit Function
'ErrCopiaSeg:
'    MsgBox "Se ha producido un error." & vbCrLf & _
'        "Número: " & Err.Number & vbCrLf & _
'        "Descripción: " & Err.Description, vbExclamation
'End Function

Public Sub MuestraError(Numero As Long, Optional CADENA As String)
Dim cad As String
'Con este sub pretendemos unificar el msgbox para todos los errores
'que se produzcan
On Error Resume Next
cad = "Se ha producido un error: " & vbCrLf
If CADENA <> "" Then
    cad = cad & vbCrLf & CADENA & vbCrLf & vbCrLf
End If
cad = cad & "Número: " & Numero & vbCrLf & "Descripción: " & Error(Numero)
MsgBox cad, vbExclamation, "ARIPRES"

End Sub


Public Function espera(Segundos As Single)
Dim T1
T1 = Timer
Do
Loop Until Timer - T1 > Segundos
End Function





Public Function TextBoxAImporte(ByRef T As TextBox) As String
Dim Mon As Currency

    T = Trim(T.Text)
    If Not IsNumeric(T.Text) Then
        TextBoxAImporte = "No es campo numerico"
        Exit Function
    End If
    
    If InStr(1, T.Text, ",") Then
        'Ya esta formateado
        TextBoxAImporte = ""
        Exit Function
    End If
    
    'Llegados aqui solo hay un punto. Luego lo pasamos a moneda
    Mon = CCur(TransformaPuntosComas(T.Text))
    T.Text = Format(Mon, "##,###,##0.00")
        
    
End Function

'Se presupone k el texto esta formateado
Public Function ImporteFormateadoAmoneda(ByVal Texto As String) As Currency
Dim i As Integer

    ImporteFormateadoAmoneda = 0
    Do
        i = InStr(1, Texto, ".")
        If i > 0 Then Texto = Mid(Texto, 1, i - 1) & Mid(Texto, i + 1)
    Loop Until i = 0
    'Ahora solo queda con el punto
    ImporteFormateadoAmoneda = CCur(Texto)
    
End Function




Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
    Dim rs As Recordset
    Dim cad As String
    Dim AUX As String
    
    On Error GoTo EDevuelveDesdeBD
    DevuelveDesdeBD = ""
    cad = "Select " & kCampo
    If OtroCampo <> "" Then cad = cad & ", " & OtroCampo
    cad = cad & " FROM " & Ktabla
    cad = cad & " WHERE " & Kcodigo & " = "
    If Tipo = "" Then Tipo = "N"
    Select Case Tipo
    Case "N"
        'No hacemos nada
        cad = cad & ValorCodigo
    Case "T"
    
        cad = cad & "'" & ValorCodigo & "'"
        
    Case "F"
        cad = cad & "#" & ValorCodigo & "#"
    Case Else
        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
        Exit Function
    End Select
    
    
    
    'Creamos el sql
    Set rs = New ADODB.Recordset
    rs.Open cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    If Not rs.EOF Then
        DevuelveDesdeBD = DBLet(rs.Fields(0))
        If OtroCampo <> "" Then OtroCampo = DBLet(rs.Fields(1))
    End If
    rs.Close
    Set rs = Nothing
    Exit Function
EDevuelveDesdeBD:
        MuestraError Err.Number, "Devuelve DesdeBD." & Err.Description
End Function


Public Function HorasDiff(ByRef Hora1 As Date, ByRef Hora2 As Date) As Currency
Dim v1 As Currency
Dim v2 As Currency
    v1 = CCur(DevuelveValorHora(Hora1))
    v2 = CCur(DevuelveValorHora(Hora2))
    HorasDiff = v2 - v1
End Function

'Esto es para las tareas
Public Function HorasExtra1(ByVal H1 As Date, ByVal h2 As Date, ByRef vHor As CHorarios) As Currency
Dim SumaE As Currency
Dim DEc As Currency

    'Comprobamos si es dia festivo
    If vHor.EsDiaFestivo Then
        DEc = CCur(DevuelveValorHora(h2))
        DEc = DEc - CCur(DevuelveValorHora(H1))
        HorasExtra1 = DEc
        Exit Function
    End If
    
    SumaE = 0
    If H1 < vHor.HoraE1 Then
        If h2 < vHor.HoraE1 Then
            DEc = CCur(DevuelveValorHora(h2))
            DEc = DEc - CCur(DevuelveValorHora(H1))
            HorasExtra1 = DEc
            Exit Function
            
        Else
            DEc = CCur(DevuelveValorHora(vHor.HoraE1))
            DEc = DEc - CCur(DevuelveValorHora(H1))
            SumaE = SumaE + DEc
            H1 = vHor.HoraE1
        End If
    End If
        

    If H1 <= vHor.HoraS1 Then
        If h2 <= vHor.HoraS1 Then
            'No ha hecho mas horas extras(si es ke ha hecho
            HorasExtra1 = SumaE
            Exit Function
        
        Else
            H1 = vHor.HoraS1
        End If
    End If
    
    
    'No tiene horario por las tardes
    If vHor.HoraE2 = "0:00:00" Then
            DEc = CCur(DevuelveValorHora(h2))
            DEc = DEc - CCur(DevuelveValorHora(H1))
            HorasExtra1 = SumaE + DEc
            Exit Function
    End If
    
    
    If H1 < vHor.HoraE2 Then
        If h2 < vHor.HoraE2 Then
            DEc = CCur(DevuelveValorHora(h2))
            DEc = DEc - CCur(DevuelveValorHora(H1))
            HorasExtra1 = SumaE + DEc
            Exit Function
        Else
            DEc = CCur(DevuelveValorHora(vHor.HoraE2))
            DEc = DEc - CCur(DevuelveValorHora(H1))
            SumaE = SumaE + DEc
            H1 = vHor.HoraE2
        End If
    End If
    
    If H1 < vHor.HoraS2 Then
        If h2 <= vHor.HoraS2 Then
            HorasExtra1 = SumaE
            Exit Function
        Else
            H1 = vHor.HoraS2
        End If
    End If
    
    DEc = CCur(DevuelveValorHora(h2))
    DEc = DEc - CCur(DevuelveValorHora(H1))
    HorasExtra1 = SumaE + DEc
    
    
End Function



'Public Function DevuelveDesdeBD(kCampo As String, Ktabla As String, Kcodigo As String, ValorCodigo As String, Optional Tipo As String, Optional ByRef OtroCampo As String) As String
'    Dim Rs As Recordset
'    Dim Cad As String
'    Dim AUx As String
'
'    On Error GoTo EDevuelveDesdeBD
'    DevuelveDesdeBD = ""
'    Cad = "Select " & kCampo
'    If OtroCampo <> "" Then Cad = Cad & ", " & OtroCampo
'    Cad = Cad & " FROM " & Ktabla
'    Cad = Cad & " WHERE " & Kcodigo & " = "
'    If Tipo = "" Then Tipo = "N"
'    Select Case Tipo
'    Case "N"
'        'No hacemos nada
'        Cad = Cad & ValorCodigo
'    Case "T", "F"
'        Cad = Cad & "'" & ValorCodigo & "'"
'    Case Else
'        MsgBox "Tipo : " & Tipo & " no definido", vbExclamation
'        Exit Function
'    End Select
'
'
'
'    'Creamos el sql
'    Set Rs = New ADODB.Recordset
'    Rs.Open Cad, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'    If Not Rs.EOF Then
'        DevuelveDesdeBD = DBLet(Rs.Fields(0))
'        If OtroCampo <> "" Then OtroCampo = DBLet(Rs.Fields(1))
'    End If
'    Rs.Close
'    Set Rs = Nothing
'    Exit Function
'EDevuelveDesdeBD:
'        MuestraError Err.Number, "Devuelve DesdeBD.", Err.Description
'End Function


Public Function SeparaValorBusqueda(ByRef CADENA As String, ByRef Oper As String, ByRef Valor As String) As Boolean
Dim C As String
Dim i As Integer
    On Error GoTo ESeparaValorBusqueda
    SeparaValorBusqueda = False
    
    
    C = Mid(CADENA, 1, 1)
    Oper = ""
    i = 1
    If Operacion(C) Then
        Oper = C
        C = Mid(CADENA, 2, 1)
        i = 2
        If Operacion(C) Then
            Oper = Oper & C
            i = 3
        End If
    End If
    Valor = Mid(CADENA, i)
    
    
    SeparaValorBusqueda = True
    Exit Function
ESeparaValorBusqueda:
    MuestraError Err.Number, "Separa Valor Busqueda"
End Function



Private Function Operacion(cad As String) As Boolean
    Select Case cad
    Case ">", "<", "="
        Operacion = True
    Case Else
        Operacion = False
    End Select
End Function





'Para los nombre que pueden tener ' . Para las comillas habra que hacer dentro otro INSTR
Public Sub NombreSQL(ByRef CADENA As String)
Dim J As Integer
Dim i As Integer
Dim AUX As String
    J = 1
    Do
        i = InStr(J, CADENA, "'")
        If i > 0 Then
            AUX = Mid(CADENA, 1, i - 1) & "'"
            CADENA = AUX & Mid(CADENA, i)
            J = i + 2
        End If
    Loop Until i = 0
End Sub


Public Function ASQL(CADENA As String) As String
Dim J As Integer
Dim i As Integer
Dim AUX As String
    J = 1
    Do
        i = InStr(J, CADENA, "'")
        If i > 0 Then
            AUX = Mid(CADENA, 1, i - 1) & "'"
            CADENA = AUX & Mid(CADENA, i)
            J = i + 2
        End If
    Loop Until i = 0
    ASQL = CADENA
End Function



Public Function EsFechaOK(ByRef T As TextBox) As Boolean
Dim cad As String
    
    cad = T.Text
    If InStr(1, cad, "/") = 0 Then
        If Len(T.Text) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T.Text) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If
    
    If IsDate(cad) Then
        EsFechaOK = True
        T.Text = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOK = False
    End If
End Function



Public Function EsFechaOKString(ByRef T As String) As Boolean
Dim cad As String
    
    cad = T
    If InStr(1, cad, "/") = 0 Then
        If Len(T) = 8 Then
            cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        Else
            If Len(T) = 6 Then cad = Mid(cad, 1, 2) & "/" & Mid(cad, 3, 2) & "/" & Mid(cad, 5)
        End If
    End If
    If IsDate(cad) Then
        EsFechaOKString = True
        T = Format(cad, "dd/mm/yyyy")
    Else
        EsFechaOKString = False
    End If
End Function



Public Function DiasMes(Mes As Integer, Anyo As Integer) As Integer
    
    Select Case Mes
    Case 1, 3, 5, 7, 8, 10, 12
        DiasMes = 31
    Case 2
        DiasMes = 28
        If (Anyo Mod 4) = 0 Then DiasMes = 29
    Case Else
        DiasMes = 30
    End Select
End Function



Public Function DiasLaborablesSemana(Horario As Integer) As Integer
Dim sql As String
Dim rs As ADODB.Recordset
    sql = "SELECT Count(*) From SubHorarios Where SubHorarios.IdHorario = " & Horario & " And SubHorarios.Festivo = False"
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If rs.EOF Then
        DiasLaborablesSemana = -1
    Else
        DiasLaborablesSemana = DBLet(rs.Fields(0), "N")
    End If
    rs.Close
    Set rs = Nothing
End Function



'-----------------------------------------------------------
'Devolverá entre un intervalo de fechas , y para un horario
'los dias k son festivos
'  -> Iran empipados los dias con formato normal de fecha dd/mm/yyyy|dd/mm/yyyy|
Public Function DevuelveDiasFestivos(Horario As Integer, Fini As Date, FFin As Date) As String
Dim cad As String
Dim rs As ADODB.Recordset

        
        cad = "Select * from Festivos WHERE idHorario = " & "#"
        
    
End Function

Public Sub Keypress(ByRef KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{tab}"
    End If
End Sub

Public Sub PonerFoco(ByRef T As TextBox)
    T.SelStart = 0
    T.SelLength = Len(T.Text)
End Sub

Public Sub PonleFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub FijarPrimerDiaSemana()
Dim C As String
    On Error GoTo EFijarPrimerDiaSemana
    myMonday = vbMonday
    C = Dir(App.Path & "\*.dia", vbArchive)
    If C <> "" Then
        C = Mid(C, 1, (InStr(1, C, ".") - 1))
        myMonday = CInt(C)
    End If
    Exit Sub
EFijarPrimerDiaSemana:
    Err.Clear
End Sub




'------------------------------------------------------------------
'------------------------------------------------------------------
'
'   Despues de leer fichajes/tareas,
'
'   Llegan tareas y salidas. Tengo que revisar el dia y veo la primera entrada y hasta que encuentre salida las borro
Public Sub ProcesaEntradaFichajesCatadau(ByRef lb As Label)
Dim sql As String
Dim rs As ADODB.Recordset
Dim Col As Collection
Dim J As Integer
Dim F1 As Date
Dim H1 As Date
Dim YaHaEntrado As Boolean

    On Error GoTo eProcesaEntradaFichajesCatadau

    lb.Caption = "leyendo trabajadores"
    lb.Refresh
    sql = "Select idTrabajador from entradafichajes "   'where fecha>=#" & Format(PrimerDia, "yyyy/mm/dd") & "#"
    sql = sql & " GROUP BY idTrabajador ORDER BY idTrabajador"
    Set Col = New Collection
    Set rs = New ADODB.Recordset
    rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not rs.EOF
        sql = rs!idTrabajador
        Col.Add sql
        
        rs.MoveNext
    Wend
    rs.Close
    
    
    For J = 1 To Col.Count
        lb.Caption = "Proceso: " & J & " / " & Col.Count
        lb.Refresh
        sql = "Select * from entradafichajes "  'where fecha>=#" & Format(PrimerDia, "yyyy/mm/dd") & "#"
        sql = sql & " WHERE idtrabajador = " & Col.Item(J)
        sql = sql & " ORDER BY idTrabajador,fecha,hora"

        rs.Open sql, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        sql = "" 'Pondre las entradas que vamos a eliminar
        If Not rs.EOF Then
            YaHaEntrado = True
            If rs!idInci = 2 Then
                'MAL, primer fichaje una salida, de mometno lo permito
                YaHaEntrado = False
            End If
            F1 = rs!Fecha
            H1 = rs!HoraReal
            
           ' Debug.Print RS!idTrabajador
            
            'Muevo al siguiente
            rs.MoveNext
            
            While Not rs.EOF
                If rs!Fecha <> F1 Then
                    F1 = rs!Fecha
                    If rs!idInci = 2 Then
                        'Mal, perimero una salida, pero contionuo
                        YaHaEntrado = False
                    Else
                        YaHaEntrado = True
                    End If
                Else
                    If YaHaEntrado Then
                        'Esoy esperando una salida
                        If rs!idInci <> 2 Then
                            'Estoy esperando una salida
                            sql = sql & ", " & rs!Secuencia
                        Else
                            'ok
                            YaHaEntrado = False  'ya salio
                        End If
                       
                    Else
                        'SALIO, estoy esperando una entrada
                        If rs!idInci <> 2 Then
                            'ok
                       
                            YaHaEntrado = True  'OK ahora otra entrada
                        Else
                            'Esperaba una entrad. Esta la borro
                            sql = sql & ", " & rs!Secuencia
                        End If
                    End If
                End If
                rs.MoveNext
            Wend
        End If
        rs.Close
        
        
        
        If sql <> "" Then
            'Hay que borrar unas entradas
            sql = Mid(sql, 2)
            sql = "DELETE FROM entradafichajes where idtrabajador = " & Col.Item(J) & " AND secuencia IN (" & sql & ")"
            conn.Execute sql
            
            'Las incidencias 2 de dias anteriores DEBERIA quitarlas
                        
        End If
    Next J




    Exit Sub
eProcesaEntradaFichajesCatadau:
    MuestraError Err.Number, "Procesando tarea/fichajes"
    Set rs = Nothing
End Sub
