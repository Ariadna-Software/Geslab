VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHorarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'Variables locales que contienen valores de propiedad
Private mvarIdHorario As Integer 'copia local
Private mvarNomHorario As String 'copia local
Private mvarDiaSemana As Byte 'copia local
Private mvarHoraE1 As Date 'copia local
Private mvarHoraS1 As Date 'copia local
Private mvarHoraE2 As Date 'copia locdal
Private mvarHoraS2 As Date 'copia local
Private mvarNumTikadas As Byte 'copia local
Private mvarTotalHoras As Single 'copia local
'Private mvarTotalHoras As byte 'copia local
Private mvarEsDiaFestivo As Boolean 'copia local
Private mvarDtoAlm As Single 'copia local
Private mvarHoraDtoAlm As Date 'copia local
Private mvarDtoMer As Single 'copia local
Private mvarHoraDtoMer As Date 'copia local
Private mvarDiaNomina As Currency 'copia local



Private SQL As String  'Para la sentencia
Private Mens As String    'Para los mensajes
Private RS As Recordset
Private RC As Byte



'------------------------------------------------
'Propiedades del modulo CHorarios
'------------------------------------------------


Public Property Let IdHorario(ByVal vData As Integer)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.IdHorario = 5
     mvarIdHorario = vData
End Property

Public Property Get IdHorario() As Integer
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.IdHorario
     IdHorario = mvarIdHorario
End Property


Public Property Let NomHorario(ByVal vData As String)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NomHorario = 5
     mvarNomHorario = vData
End Property

Public Property Get NomHorario() As String
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NomHorario
     NomHorario = mvarNomHorario
End Property


Public Property Let DiaSemana(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DiaSemana = 5
     mvarDiaSemana = vData
End Property

Public Property Get DiaSemana() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DiaSemana
     DiaSemana = mvarDiaSemana
End Property


Public Property Let HoraE1(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraE1 = 5
     mvarHoraE1 = vData
End Property

Public Property Get HoraE1() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraE1
     HoraE1 = mvarHoraE1
End Property


Public Property Let HoraS1(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraS1 = 5
     mvarHoraS1 = vData
End Property

Public Property Get HoraS1() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraS1
     HoraS1 = mvarHoraS1
End Property


Public Property Let HoraE2(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraE2 = 5
     mvarHoraE2 = vData
End Property

Public Property Get HoraE2() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraE2
     HoraE2 = mvarHoraE2
End Property


Public Property Let HoraS2(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraS2 = 5
     mvarHoraS2 = vData
End Property

Public Property Get HoraS2() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraS2
     HoraS2 = mvarHoraS2
End Property


Public Property Let NumTikadas(ByVal vData As Byte)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NumTikadas = 5
     mvarNumTikadas = vData
End Property

Public Property Get NumTikadas() As Byte
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NumTikadas
     NumTikadas = mvarNumTikadas
End Property

Public Property Let TotalHoras(ByVal vData As Single)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NumTikadas = 5
     mvarTotalHoras = vData
End Property

Public Property Get TotalHoras() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NumTikadas
     TotalHoras = mvarTotalHoras
End Property



Public Property Let EsDiaFestivo(ByVal vData As Boolean)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.NumTikadas = 5
     mvarEsDiaFestivo = vData
End Property

Public Property Get EsDiaFestivo() As Boolean
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.NumTikadas
     EsDiaFestivo = mvarEsDiaFestivo
End Property



Public Property Let DtoAlm(ByVal vData As Single)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DtoAlm = 5
     mvarDtoAlm = vData
End Property

Public Property Get DtoAlm() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DtoAlm
     DtoAlm = mvarDtoAlm
End Property


Public Property Let HoraDtoAlm(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraDtoAlm = 5
     mvarHoraDtoAlm = vData
End Property

Public Property Get HoraDtoAlm() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraDtoAlm
     HoraDtoAlm = mvarHoraDtoAlm
End Property


Public Property Let DtoMer(ByVal vData As Single)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.DtoMer = 5
     mvarDtoMer = vData
End Property

Public Property Get DtoMer() As Single
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.DtoMer
     DtoMer = mvarDtoMer
End Property


Public Property Let HoraDtoMer(ByVal vData As Date)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraDtoMer = 5
     mvarHoraDtoMer = vData
End Property

Public Property Get HoraDtoMer() As Date
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraDtoMer
     HoraDtoMer = mvarHoraDtoMer
End Property



Public Property Let DiaNomina(ByVal vData As Currency)
'se usa cuando se asigna un valor a la propiedad, en el
'lado izquierdo de la asignacion
'Syntax: X.HoraDtoMer = 5
     mvarDiaNomina = vData
End Property

Public Property Get DiaNomina() As Currency
'se usa cuando se asigna un valor a la propiedad, en el
'lado derecho de la asignacion
'Syntax: Debug.Print X.HoraDtoMer
     DiaNomina = mvarDiaNomina
End Property





'------------------------------------------------
'FUNCIONES del modulo CHorarios
'------------------------------------------------


'****   LEER  ***
Public Function Leer(vIdIdHorario As Integer, vFecha) As Byte
Dim RT As ADODB.Recordset
Dim DiaSem As Byte
Dim HayHorarios As Boolean
Dim EsFinD As Boolean 'Festivo de Fin de semana

    On Error GoTo Err3
    SQL = "SELECT * FROM Horarios"
    SQL = SQL & " WHERE "
    SQL = SQL & " idHorario = " & vIdIdHorario
    RS.Open SQL, Conn, , , adCmdText

    If RS.EOF Then
        Leer = 1
        'ELSE
        Else
            'Variables de HORARIOS
            mvarIdHorario = RS!IdHorario
            mvarNomHorario = RS!NomHorario
            mvarDtoAlm = DBLet(RS!DtoAlm, "N")
            mvarHoraDtoAlm = DBLet(RS!HoraDtoAlm, "F")
            mvarDtoMer = DBLet(RS!DtoMer, "N")
            mvarHoraDtoMer = DBLet(RS!HoraDtoMer, "F")
            
            
            DiaSem = Weekday(vFecha, vbMonday)
            mvarDiaSemana = DiaSem
            SQL = "select * from Subhorarios where idHorario=" & vIdIdHorario
            SQL = SQL & " AND DiaSemana=" & DiaSem
            Set RT = New ADODB.Recordset
            RT.Open SQL, Conn, , , adCmdText
            If RT.EOF Then
                HayHorarios = False
                Else
                    HayHorarios = True
                    mvarHoraE1 = DBLet(RT!HEntrada1, "F")
                    mvarHoraS1 = DBLet(RT!HSalida1, "F")
                    mvarHoraE2 = DBLet(RT!HEntrada2, "F")
                    mvarHoraS2 = DBLet(RT!HSalida2, "F")
                    mvarNumTikadas = RT!N_Tikadas
                    mvarTotalHoras = RT!HorasDia
                    If Not IsNull(RT!DiaNomina) Then
                        mvarDiaNomina = RT!DiaNomina
                    Else
                        mvarDiaNomina = 0
                    End If
                    EsFinD = DBLet(RT!Festivo, "B")
            End If
            RT.Close
            Set RT = Nothing
            'Comprobamos si es un dia festivo
            SQL = EsFestivo(vFecha)
            If SQL <> "" Then
                If Not HayHorarios Then
                    mvarHoraE1 = "0:00:00"
                    mvarHoraS1 = "0:00:00"
                    mvarHoraE2 = "0:00:00"
                    mvarHoraS2 = "0:00:00"
                End If
                mvarEsDiaFestivo = True
                mvarNumTikadas = 0
                mvarTotalHoras = 0
                mvarNomHorario = mvarNomHorario & " (" & SQL & ")"
                Leer = 0
                Else
                    'No es festivo de calendario
                    'Comprobaremos que no es festivo de
    
                    If HayHorarios Then
                        
                        If EsFinD Then
                            mvarEsDiaFestivo = True
                            mvarNomHorario = mvarNomHorario & " (Festivo)"
                        Else
                             mvarEsDiaFestivo = False
                        End If
                        Leer = 0
                        Else
                            Leer = 1
                    End If
            End If 'DE FESTIVO
    End If
    RS.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
    Mens = Mens & "Descripci�n: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    Leer = 1
End Function

Private Sub Class_Initialize()
Set RS = New ADODB.Recordset
End Sub

Private Sub Class_Terminate()
Set RS = Nothing
End Sub



Private Function EsFestivo(Fecha) As String
Dim Cad As String
Dim RF As ADODB.Recordset

'Devuelve una cadena que dice que fiesta es( y se a�ade el nombre de horario.")
EsFestivo = ""
Set RF = New ADODB.Recordset
Cad = "Select Descripcion from Festivos where IdHorario=" & mvarIdHorario
Cad = Cad & " and Fecha=#" & Format(Fecha, "yyyy/mm/dd") & "#"
RF.Open Cad, Conn, , , adCmdText
If Not RF.EOF Then
    EsFestivo = RF.Fields(0)
End If
RF.Close
Set RF = Nothing
End Function



'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
Public Function LeerDiasFestivos(vIdIdHorario As Integer, Fini As Date, FFin As Date) As String
Dim RT As ADODB.Recordset
Dim HayHorarios As Boolean
Dim EsFinD As Boolean 'Festivo de Fin de semana
Dim Fe As Date
Dim D As Integer
Dim FESTIVOS As String
Dim RecuperaEnSabado As Boolean

    On Error GoTo Err3
    SQL = "SELECT * FROM Horarios"
    SQL = SQL & " WHERE "
    SQL = SQL & " idHorario = " & vIdIdHorario
    RS.Open SQL, Conn, , , adCmdText
    LeerDiasFestivos = ""
    FESTIVOS = ""
    If Not RS.EOF Then
        
            'Variables de HORARIOS
            mvarIdHorario = RS!IdHorario
            mvarNomHorario = RS!NomHorario
            
            'Vemos si los sabados son recuperables
            RecuperaEnSabado = RecuperaDiasEnSabado(RS)
            
            Set RT = New ADODB.Recordset
                        
            SQL = "Select * from Festivos WHERE idHorario = " & mvarIdHorario
            SQL = SQL & " AND Fecha >= #" & Format(Fini, FormatoFecha) & "#"
            SQL = SQL & " AND Fecha <= #" & Format(FFin, FormatoFecha) & "#"
            SQL = SQL & " ORDER By FECHA"
            RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            While Not RT.EOF
                FESTIVOS = FESTIVOS & Format(RT!Fecha, "dd/mm/yyyy") & "|"
                RT.MoveNext
            Wend
            RT.Close
            
            SQL = "select * from Subhorarios where idHorario=" & vIdIdHorario
            RT.Open SQL, Conn, , , adCmdText
            If RT.EOF Then
                
                Else
                    'En el SQL Meteremos los dias de las semana k son festivos
                    'Es decir , empimaremos los dias k sean festivos
                    ' Doming|Sabado|  por ejemplo, pero en numeros 5|6|
                    SQL = ""
                    While Not RT.EOF
                        
                        If RT!Festivo Then
                           
                            If RecuperaEnSabado Then
                                If RT!DiaSemana <> 6 Then
                                    SQL = SQL & RT!DiaSemana & "|"
                                End If
                            Else
                                SQL = SQL & RT!DiaSemana & "|"
                            End If
                        End If
                        RT.MoveNext
                    Wend
                    
                    
                    Fe = Fini
                    While Fe <= FFin
                        D = Weekday(Fe, vbMonday)
                        Mens = D & "|"
                        If InStr(1, SQL, Mens) > 0 Then
                                FESTIVOS = FESTIVOS & Format(Fe, "dd/mm/yyyy") & "|"
                        End If
                            
                        'Dia mas
                        Fe = DateAdd("d", 1, Fe)
                    Wend
            End If
            RT.Close
            Set RT = Nothing
            'Comprobamos si es un dia festivo
            LeerDiasFestivos = FESTIVOS
    
    End If
    RS.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
    Mens = Mens & "Descripci�n: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    LeerDiasFestivos = ""
End Function





'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
'----------------------------------------------------------------
Public Function LeerMediosDias(vIdIdHorario As Integer, Fini As Date, FFin As Date) As String
Dim RT As ADODB.Recordset
Dim HayHorarios As Boolean
Dim EsFinD As Boolean 'Festivo de Fin de semana
Dim Fe As Date
Dim D As Integer
Dim MediosDias As String

    On Error GoTo Err3
    SQL = "SELECT * FROM Horarios"
    SQL = SQL & " WHERE "
    SQL = SQL & " idHorario = " & vIdIdHorario
    RS.Open SQL, Conn, , , adCmdText
    LeerMediosDias = ""
    MediosDias = ""
    If Not RS.EOF Then
        
            'Variables de HORARIOS
            mvarIdHorario = RS!IdHorario
            mvarNomHorario = RS!NomHorario
            Set RT = New ADODB.Recordset
            
            SQL = "select * from Subhorarios where idHorario=" & vIdIdHorario
            RT.Open SQL, Conn, , , adCmdText
            If RT.EOF Then
                
                Else
                    'En el SQL Meteremos los dias de las semana k s trabajan 0.5 horas
                    ' Miercoles|Sabado|  por ejemplo, pero en numeros 3|6|
                    SQL = ""
                    While Not RT.EOF
                        If Not RT!Festivo Then
                            If RT!DiaNomina < 1 Then SQL = SQL & RT!DiaSemana & "|"
                        End If
                        RT.MoveNext
                    Wend
                    
                    
                    Fe = Fini
                    While Fe <= FFin
                        D = Weekday(Fe, vbMonday)
                        Mens = D & "|"
                        If InStr(1, SQL, Mens) > 0 Then
                                MediosDias = MediosDias & Format(Fe, "dd/mm/yyyy") & "|"
                        End If
                            
                        'Dia mas
                        Fe = DateAdd("d", 1, Fe)
                    Wend
            End If
            RT.Close
            Set RT = Nothing
            'Comprobamos si es un dia festivo
            LeerMediosDias = MediosDias
    
    End If
    RS.Close
    Exit Function

Err3:
    Mens = "Se ha producido un error." & vbCrLf
    Mens = Mens & "N�mero: " & Err.Number & vbCrLf
    Mens = Mens & "Descripci�n: " & Err.Description
    MsgBox Mens, vbExclamation
    RS.Close
    LeerMediosDias = ""
End Function



Private Function RecuperaDiasEnSabado(ByRef R As Recordset) As Boolean
    On Error Resume Next
    RecuperaDiasEnSabado = False
    RecuperaDiasEnSabado = RS!RecuperaSabados
    If Err.Number <> 0 Then Err.Clear
End Function
