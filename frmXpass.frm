VERSION 5.00
Begin VB.Form frmXpass 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lectura fichajes X-Pass"
   ClientHeight    =   2790
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5865
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLeer 
      Caption         =   "Leer"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   2160
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   4560
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   2565
      Left            =   0
      Picture         =   "frmXpass.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   1290
   End
   Begin VB.Label lblInd 
      Caption         =   "Label1"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   4095
   End
End
Attribute VB_Name = "frmXpass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Cn As ADODB.Connection

Dim Cad As String
Dim RS As ADODB.Recordset


Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdLeer_Click()
Dim AUX As String
Dim Col As Collection
Dim J As Integer
Dim Incremento As Long
Dim Ultimo_nEventLogIdn As Long
Dim LaFecha As Date
Dim Trabajadores As ADODB.Recordset
Dim Tarje As String


    On Error GoTo ecmdLeer


    Screen.MousePointer = vbHourglass

    lblInd.Caption = "Preparando datos"
    lblInd.Refresh
    
    
    Cad = "DElete from temporalfichajes"
    conn.Execute Cad
    
    
    Cad = "select nEventLogIdn,nUserID,FROM_UNIXTIME(nDateTime) lafecha from tb_event_log where "
    'Cad = Cad & " nisuseta=1"
    Cad = Cad & " neventidn = 47"   'RAFA Septiembre 2015
    
    Cad = Cad & " AND nEventLogIdn > " & cmdLeer.Tag
    'Cad = Cad & " AND nEventLogIdn < " & cmdLeer.Tag + 2350
    Cad = Cad & " ORDER BY nEventLogIdn "
    Set RS = New ADODB.Recordset
    RS.Open Cad, Cn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    AUX = ""
    While Not RS.EOF
        lblInd.Caption = "Reg: " & RS!nEventLogIdn
        lblInd.Refresh
         
         
        LaFecha = RS!LaFecha
        Incremento = DevuelveIncrementoUTC(LaFecha)
        LaFecha = DateAdd("h", -Incremento, LaFecha)
         
        AUX = " (" & RS!nEventLogIdn & "," & RS!nUserID & ",#" & Format(LaFecha, FormatoFecha) & "#,#"
        AUX = AUX & Format(LaFecha, "hh:mm:ss") & "#,0)"
        AUX = "INSERT INTO temporalfichajes  VALUES " & AUX & ";"
        conn.Execute AUX
        
        
        Ultimo_nEventLogIdn = RS!nEventLogIdn   'Me guardo el ultimo eventlog
        
        RS.MoveNext
    Wend
    RS.Close
    
    If AUX = "" Then
        MsgBox "Ningun dato pendiende de traspasar", vbExclamation
        GoTo ecmdLeer
    End If
    
    Set Trabajadores = New ADODB.Recordset
'    Cad = "Select * from     "
'
'
'    Cad = "Select fecha from temporalfichajes group by fecha order by fecha"
'    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    Set Col = New Collection
'    While Not RS.EOF
'        Col.Add CStr(RS!Fecha)
'        RS.MoveNext
'    Wend
'    RS.Close
'
'
'    For J = 1 To Col.Count
'        lblInd.Caption = "Ajuste hora " & Col.Item(J)
'        lblInd.Refresh
'        Incremento = DevuelveIncrementoUTC(CDate(Col.Item(J)))
'
'        Cad = "UPDATE temporalfichajes set hora = dateadd(""h"",-" & Incremento & ",hora)"
'        Cad = Cad & " WHERE fecha = #" & Format(Col.Item(J), FormatoFecha) & "#"
'
'        conn.Execute Cad
'
'    Next J
'
    
    
    
    'Ponemos UN negativo para luego hacer la transcripicon tarjeta-trabajador
    
    Cad = "UPDATE temporalfichajes set numtarjeta= ""T"" & numtarjeta "
    conn.Execute Cad
    
    'OK. Ya tenemos leidos los fichajes desde los Xpass, ahora comprobamos datos
    'Miraremos que todos los trabajadores existen
    lblInd.Caption = "Comprobar trabajadores"
    lblInd.Refresh
    Set Col = Nothing
    Set Col = New Collection
    Cad = "Select numtarjeta from temporalfichajes group by numtarjeta order by numtarjeta"
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        Col.Add CStr(RS!numtarjeta)
        RS.MoveNext
    Wend
    RS.Close
        
    AUX = "" 'Llevare registros errorneos
    For J = 1 To Col.Count
        Tarje = Mid(Col.Item(J), 2)
        Tarje = Format(Tarje, "00000")
        
        lblInd.Caption = "Trabajador tarjeta: " & Tarje
        lblInd.Refresh
        
        
        
        
        Cad = "Select idtrabajador from trabajadores where numtarjeta = """ & Tarje & """"
        
        
        
        RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RS.EOF Then
            AUX = AUX & Col.Item(J) & vbCrLf
            
        Else
            Cad = "UPDATE temporalfichajes set numtarjeta=" & RS!idTrabajador & " where numtarjeta = """ & Col.Item(J) & """"
            conn.Execute Cad
        End If
        RS.Close
        
        
    Next J
    
    If AUX <> "" Then
        AUX = "Identificacion trabajadores incorrecta" & vbCrLf & String(40, "=") & vbCrLf & AUX
        Err.Raise 513, "Codigos incorrectos", AUX
        
    End If
    
    
    
    
    
    
    'OK. fichajes desde los Xpass y comprobados ahora hay que meterlos en entradafichajees
    'Metemos todos los fichajes en entradaficahjes
    Cad = "Select max(secuencia) FROM entradafichajes"
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = "0"
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then Cad = RS.Fields(0)
    End If
    RS.Close
    Incremento = Val(Cad)
    
    
    
    
    conn.BeginTrans
    If InsertarEnEntradaFichajes(Incremento) Then
        Cad = "UPDATE empresas set xpassultID = " & Ultimo_nEventLogIdn
        conn.Execute Cad
    
        conn.CommitTrans
        
        Unload Me
        
    Else
        conn.RollbackTrans
    End If
    
    
    
    
ecmdLeer:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
    Set Trabajadores = Nothing
    lblInd.Caption = ""
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    If Cn Is Nothing Then
        Set Cn = New ADODB.Connection
        AbrirConexion
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal1.Icon
    
End Sub


Private Sub AbrirConexion()

    On Error GoTo eAbrirConexion
    lblInd.Caption = "Abriendo conexion BD X-Pass"
    lblInd.Refresh
    cmdLeer.Enabled = False
    cmdLeer.Tag = 0   'Ultima entrada leida
    Cad = "Select xpassserver,xpassuser,xpasspwd,xpassultID from empresas"
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not RS.EOF Then
        If Not IsNull(RS!xpassserver) Then
            Cad = "DRIVER={MySQL ODBC 3.51 Driver};DESC=;DATABASE=biostar;SERVER=" & RS!xpassserver
            Cad = Cad & ";UID=" & DBLet(RS!xpassuser, "T")
            Cad = Cad & ";PWD=" & DBLet(RS!xpasspwd, "T")
            Cad = Cad & ";PORT=3306;OPTION=3;STMT=;"
            cmdLeer.Tag = DBLet(RS!xpassultID, "N")
            
            Set Cn = New ADODB.Connection
            Cn.ConnectionString = Cad
            Cn.Open
            lblInd.Caption = ""
            lblInd.Refresh
            If vUsu.Codigo < 3 Then cmdLeer.Enabled = True
        Else
            lblInd.Caption = "Parametro servidor X-Pass"
            lblInd.Refresh
            
        End If
    Else
        lblInd.Caption = "Error leendo parametros empresa"
        lblInd.Refresh
    End If

    
    
    
eAbrirConexion:
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description
    Set RS = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set RS = New ADODB.Recordset
    Set Cn = Nothing
End Sub

Private Function DevuelveIncrementoUTC(Fecha As Date) As Integer
Dim diaCambio As Integer
Dim Mes As Integer

    Mes = Month(Fecha)

    If Mes < 3 Or Mes > 10 Then
        DevuelveIncrementoUTC = 1
    Else
        If Mes = 3 Or Mes = 10 Then
            diaCambio = DiaCambioHora(Mes, Year(Fecha))
            If Mes = 3 Then
                If Day(Fecha) < diaCambio Then
                    DevuelveIncrementoUTC = 1
                Else
                    DevuelveIncrementoUTC = 2
                End If
            Else
                'OCTBURE
                If Day(Fecha) > diaCambio Then
                    DevuelveIncrementoUTC = 1
                Else
                    DevuelveIncrementoUTC = 2
                End If
            End If
        Else
            DevuelveIncrementoUTC = 2
        End If
    End If
        
    
End Function


Private Function DiaCambioHora(Mes As Integer, Anyo As Integer) As Integer
Dim F As Date
Dim DiaSem
Dim Dia As Integer

    DiaCambioHora = 0
    Dia = 31
    Do
        F = CDate(Dia & "/" & Format(Mes, "00") & "/" & Anyo)
        DiaSem = Weekday(F, vbMonday)
        If DiaSem = 7 Then
            DiaCambioHora = Day(F)
            Dia = 1
        
        End If
        Dia = Dia - 1
    Loop Until Dia <= 0
        
    
    If DiaCambioHora = 0 Then
        MsgBox "Error calculando dia cambio horario verano-invierno", vbExclamation
        End
    End If

    
End Function




'*********************************************
Private Function InsertarEnEntradaFichajes(Secuen As Long) As Boolean
Dim PrimeraInsercio As Long
Dim Minutos As Long
Dim Fecha As Date
Dim Hora As Date
Dim Repeticion As Integer


    On Error GoTo eInsertarEnEntradaFichajes
    
    InsertarEnEntradaFichajes = False
    PrimeraInsercio = -1
    Cad = "Select * from temporalfichajes  order by fecha,hora"
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        lblInd.Caption = "Registro: " & Secuen
        lblInd.Refresh
    
        Secuen = Secuen + 1
        If PrimeraInsercio < 0 Then PrimeraInsercio = Secuen
        
        Cad = "INSERT INTO entradafichajes(secuencia,idtrabajador,fecha,hora,idinci,horareal) VALUES ("
        Cad = Cad & Secuen & "," & RS!numtarjeta & ",#" & Format(RS!Fecha, FormatoFecha) & "#,#"
        Cad = Cad & Format(RS!Hora, "hh:mm:ss") & "#,0,#" & Format(RS!Hora, "hh:mm:ss") & "#)"
    
        conn.Execute Cad
        
        RS.MoveNext
    Wend
    RS.Close
    
    'Repetidos
    lblInd.Caption = "Eliminando marcajes repeticion "
    lblInd.Caption = ""
    Me.Refresh
    espera 0.5
    
    Cad = DevuelveDesdeBD("repeticion", "Empresas", "idEmpresa", "1", "N")
    Repeticion = Val(Cad)
    If Repeticion > 0 Then
        'Obtenemos la fecha mas baja
       
        Cad = "Select min(fecha) from EntradaFichajes WHERE Secuencia >= " & PrimeraInsercio
        RS.Open Cad, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
        
        Cad = "#1900/01/01#"
        If Not RS.EOF Then
            If Not IsNull(RS.Fields(0)) Then Cad = "#" & Format(RS.Fields(0), FormatoFecha) & "#"
        End If
        RS.Close
        Cad = " Fecha >= " & Cad
    
        
        'Ya tenemos a partir de k fecha, y con k cadencia vamos a eliminar repetidos
        Cad = "Select * from Entradafichajes WHERE " & Cad
        Cad = Cad & " ORDER BY idTrabajador,Fecha,Hora"
        RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        Secuen = 0 'Tendremos el codigo del trabajador
        Cad = "DELETE from EntradaFichajes WHERE Secuencia = "
        While Not RS.EOF
            If RS!idTrabajador <> Secuen Then
                lblInd.Caption = "Trabajador: " & RS!idTrabajador
                lblInd.Refresh
                'Nuevo trabajador
                Secuen = RS!idTrabajador
                Fecha = RS!Fecha
                Hora = RS!Hora
            Else
                'Es el mismo trabajador.
                'Veamos la fecha
                If RS!Fecha <> Fecha Then
                    Fecha = RS!Fecha
                    Hora = RS!Hora
                Else
                    'MISMO TRABAJADOR , MISMA FECHA
                    Minutos = DateDiff("n", Hora, RS!Hora)
                    If Minutos > Repeticion Then
                        'Las horas se diferencian. NO elimino
                        Hora = RS!Hora
                    Else
                        'SI elimino
                        conn.Execute Cad & RS!Secuencia
                    End If
                End If
            End If
            'Siguiente
            RS.MoveNext
        Wend
        RS.Close
    
    End If  'Eliminacion marcajes repetidos

    
    
    'Conmprobacion de bajas
    lblInd.Caption = "Comprobar bajas"
    lblInd.Refresh
    Cad = "SELECT Trabajadores.NomTrabajador, Bajas.idTrab"
    Cad = Cad & " FROM (EntradaFichajes INNER JOIN Bajas ON EntradaFichajes.idTrabajador = Bajas.idTrab) INNER JOIN Trabajadores ON EntradaFichajes.idTrabajador = Trabajadores.IdTrabajador"
    Cad = Cad & " WHERE (((Bajas.FechaAlta) Is Null) AND ((EntradaFichajes.Secuencia)>= " & PrimeraInsercio
    Cad = Cad & ")) group by  Trabajadores.NomTrabajador, Bajas.idTrab"
    Set RS = New ADODB.Recordset
    RS.Open Cad, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    While Not RS.EOF
        Cad = Cad & vbCrLf & "    - " & RS!nomtrabajador & " (" & RS!idTrab & ")"
        RS.MoveNext
    Wend
    RS.Close
    If Cad <> "" Then
        Cad = "Hay trabajadores que estan de baja y han fichado. " & vbCrLf & vbCrLf & Cad
        MsgBox Cad, vbExclamation
    End If

    
    
    
    
    
    
    InsertarEnEntradaFichajes = True
    Exit Function
eInsertarEnEntradaFichajes:
    MuestraError Err.Number, Err.Description & vbCrLf & Cad
End Function
