Attribute VB_Name = "modHorasGesLab"
 Option Explicit

Private Const HoraIntermediaMiercolesSabado = "14:00:00"


'Para las PUTAS compensaciones de los miercoles / sabado
Private SemanaMesPrimera As Integer
Private SemanaMesUltima As Integer


Public Function CalculaHorasHorario(IdHor As Integer, ByRef Dias As Integer, Fini As Date, FFin As Date, CalculoDeBaja As Boolean) As Currency
Dim Sum As Currency
Dim vH As CHorarios
Dim F As Date
Dim D As Currency
Dim Semana As Integer
Dim UltimoMiercolesTrabajado As Integer
    Set vH = New CHorarios
    
    CalculaHorasHorario = -1
    Sum = 0
    D = 0
    F = Fini
    
    'Para cada dia del mes
    
    Do
        If vH.Leer(IdHor, F) = 1 Then Exit Function
        
        If Not vH.EsDiaFestivo Then
            If vH.DiaNomina = 0.5 Then
                If CalculoDeBaja Then
                    'normal  no hacemos la cosa rara de los miercoles sabado
                    D = D + vH.DiaNomina
                    
                Else
                    'medio dia
                    Semana = Format(F, "ww")
                    If Weekday(F) = 4 Then
                        UltimoMiercolesTrabajado = Semana
                        D = D + 1
                    Else
                        If UltimoMiercolesTrabajado <> Semana Then D = D + 1
                    End If
            
                End If
            Else
                D = D + vH.DiaNomina
            End If
            Sum = Sum + vH.TotalHoras

        End If
        F = DateAdd("d", 1, F)
    Loop Until F > FFin
    'Redondeamos siempre hacia arriba
    Dias = Int(D)
    If D > Int(D) Then
        'Tiene fraccion de dia
        Dias = Int(D) + 1
    End If
    CalculaHorasHorario = Sum
End Function


'Calculo de horas. Simplemente es dias * 8
Public Function CalculaHorasHorarioALZ(IdHor As Integer, ByRef Dias As Integer, Fini As Date, FFin As Date) As Currency
Dim vH As CHorarios
Dim F As Date
Dim D As Currency

    Set vH = New CHorarios
    
    CalculaHorasHorarioALZ = -1
'    Sum = 0
    D = 0
    F = Fini
    'Para cada dia del mes
    Do
        If vH.Leer(IdHor, F) = 1 Then Exit Function
        
        If Not vH.EsDiaFestivo Then
            

            D = D + vH.DiaNomina

        End If
        F = DateAdd("d", 1, F)
    Loop Until F > FFin
    'Redondeamos siempre hacia arriba
    Dias = Int(D)
    If D > Int(D) Then
        'Tiene fraccion de dia
        Dias = Int(D) + 1
    End If
    CalculaHorasHorarioALZ = Dias * 8
End Function





'Calcula las horas trabajadas para los trabajadores k tiene la marca puesta
Public Sub CalculaHorasTrabajadas(Fini As Date, FFin As Date, ControlNomina As Byte)
Dim FAux As Date
Dim FAux2 As Date
Dim RS As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim Dias As Currency
Dim Trabajador As Long
Dim Aux As String
Dim SQL As String
Dim vH As CHorarios
Dim FESTIVOS As String
Dim MEDIODIA As String
Dim strControlNomina As String
'----------------------------------------------
'FALTA### parametrizar esto
Dim UltimoMiercolesTrabajado As Integer
Dim Semana As Integer
    

    'Modificacion 25 Noviembre.  Vamos a cambiar algunas cosas
    '
    

    
    'IMPORTANTE
    'Ahora hay un control nomina mas, k es el 2
    'El tipo de control 2: Tiene un suledo fijo al mes
    'Pero en anticpos solo anticipa hNormales
    'luego el calculo de horas es el mismo que el 1
    ' por lo tanto donde ponia
        'SQL = SQL & " AND Trabajadores.ControlNomina = 1"
    ' pondra ahora
        'SQL = SQL & " AND Trabajadores.ControlNomina > 0"


    'Otro MAS. El tipo 3
    '   40 Horas semanales. 5 dias semana
    '
    'Con lo cual si en
    ' controlnomina
        ' 1.-   NORMAL ControlNomina >0 and ControlNomina <3
        ' 2.- Solo para el tipo  3
    Select Case ControlNomina
    Case 0
        strControlNomina = " AND Trabajadores.ControlNomina >0  AND Trabajadores.ControlNomina <=2 "
    Case 1
        strControlNomina = " AND Trabajadores.ControlNomina = 3"
    Case 2
        strControlNomina = " AND (Trabajadores.ControlNomina =1  OR Trabajadores.ControlNomina =3) "
    Case 3
        'Sera para el listado que se entraga a los trabbajdores en PICASSENT
        ' Es para los tipos 1,2,3
        strControlNomina = " AND Trabajadores.ControlNomina >0"
    Case Else
        strControlNomina = ""
    End Select
    
    
    
    Conn.Execute "Delete from tmpHoras"
    
    'Calculamos las horas para el mes
    'Primero las normales con un simple insert into
    SQL = "INSERT INTO tmpHoras(trabajador,HorasT) "
    SQL = SQL & "SELECT Marcajes.idTrabajador, Sum(Marcajes.HorasTrabajadas) AS SumaDeHorasTrabajadas"
    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
    SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
    
    SQL = SQL & strControlNomina
    SQL = SQL & " GROUP BY Marcajes.idTrabajador;"
    Conn.Execute SQL
    
    
    
    '----HORAS COMPENSAR
    'Las horas para la bolsa de trabajo
    
    SQL = "SELECT Marcajes.idTrabajador,Sum(Marcajes.Horasincid) AS SumaDeHoras"
    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
    SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
    SQL = SQL & strControlNomina
    
    'Como las horas extra se consideran a compensar
    SQL = SQL & " And IncFinal =" & MiEmpresa.IncHoraExtra
    SQL = SQL & " GROUP BY Marcajes.idTrabajador;"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF
        'If RS!idTrabajador = 63 Then Stop
        SQL = "UPDATE tmpHoras Set HorasC = " & TransformaComasPuntos(RS!sumadehoras)
        '        SQL = SQL & " ,HorasT = HorasT - " & TransformaComasPuntos(RS!sumadehoras)
        SQL = SQL & " WHERE Trabajador = " & RS!idTrabajador
        Conn.Execute SQL
        
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    'Updatemos con los dias trabajados.
    '
    'Acciones:
    '       -En una variable cargaremos los dias festivos de
    '       -En Otra Cargaremos los medios dias.
    '       -Para cada dia trabajado, para cada trabajador, veremos
    '       - Si los dias trabajados es un festivo o unidad fraccionarai
    
    SQL = "SELECT idHorario"
    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
    SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
    SQL = SQL & strControlNomina
    SQL = SQL & " GROUP BY Trabajadores.idHorario;"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    
    
    While Not RS.EOF
        Set vH = New CHorarios
        If vH.Leer(RS!IdHorario, Now) = 0 Then
            FESTIVOS = vH.LeerDiasFestivos(vH.IdHorario, Fini, FFin)
            MEDIODIA = vH.LeerMediosDias(vH.IdHorario, Fini, FFin)
            
    
            'AHora para cada trabajador k haya trabajado entre las fechas le sumare dias trabajados
            '.... o no(ej festivo)
            SQL = "SELECT Marcajes.*"
            SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
            SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
            SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
            SQL = SQL & strControlNomina
            SQL = SQL & " And Trabajadores.IdHorario = " & RS!IdHorario
            SQL = SQL & " ORDER BY Marcajes.idTrabajador, Fecha"
            Set RS2 = New ADODB.Recordset
            RS2.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
            
            If Not RS2.EOF Then
                Trabajador = -1
                Do
                   If Trabajador <> RS2!idTrabajador Then
                         
                        
                         
                        If Trabajador > 0 Then
                            
                            SQL = "UPDATE tmpHoras Set Dias = "
                            If Dias > Int(Dias) Then
                                Dias = Int(Dias) + 1
                            Else
                                Dias = Int(Dias)
                            End If
                            SQL = SQL & Int(Dias)
                            SQL = SQL & " WHERE Trabajador = " & Trabajador
                            Conn.Execute SQL
                        End If
                   
                        Trabajador = RS2!idTrabajador
                        Dias = 0
                        UltimoMiercolesTrabajado = 0
                    End If
    
                    'Si el dia esta en FESTIVOS no lo sumo
                    Aux = Format(RS2!Fecha, "dd/mm/yyyy") & "|"
    
                    'NO esta en festivos
                    If InStr(1, FESTIVOS, Aux) = 0 Then
                        'Si es medio dia sumo medio
                        If InStr(1, MEDIODIA, Aux) > 0 Then
                            Semana = Format(RS2!Fecha, "ww")
                            If Weekday(RS2!Fecha) = 4 Then
                                Dias = Dias + 1
                                UltimoMiercolesTrabajado = Semana
                            Else
                                If UltimoMiercolesTrabajado <> Semana Then Dias = Dias + 1

                            End If
                        Else
                            Dias = Dias + 1
                        End If
                    End If
    
                    'Sig
                    RS2.MoveNext
                Loop Until RS2.EOF
                
                'Ahora faltara por hacer el ultimo trabajador
                SQL = "UPDATE tmpHoras Set Dias = "
                If Dias > Int(Dias) Then
                    Dias = Int(Dias) + 1
                Else
                    Dias = Int(Dias)
                End If
                SQL = SQL & Int(Dias)
                SQL = SQL & " WHERE Trabajador = " & Trabajador
                Conn.Execute SQL
            End If
        End If
        RS.MoveNext 'Siguiente horario
    Wend
        
        
        
        
        
        
        
        
        
        
        
    'Por si acaso algun trabajador tiene numeros negativos
    SQL = "UPDATE tmpHoras Set Dias = 0"
    SQL = SQL & " WHERE Dias < 0 "
    Conn.Execute SQL
    
    
    Set RS = Nothing
End Sub




Public Sub CalculaDatosMes(Fini As Date, FFin As Date, ControlNomina As Byte)
Dim FAux As Date
Dim FAux2 As Date
Dim RS As ADODB.Recordset
Dim Horas As Currency
Dim D As Integer
Dim Aux As String
Dim SQL As String
Dim strControlNomina As String
Dim D22 As Integer
Dim h22 As Currency
Dim IDT As Integer
Dim IDH As Integer
Dim vM As CMarcajes
Dim HorasBaja As Currency

    'IMPORTANTE
    'Ahora hay un control nomina mas, k es el 2
    'El tipo de control 2: Tiene un suledo fijo al mes
    'Pero en anticpos solo anticipa hNormales
    'luego el calculo de horas es el mismo que el 1
    ' por lo tanto donde ponia
        'SQL = SQL & " AND Trabajadores.ControlNomina = 1"
    ' pondra ahora
        'SQL = SQL & " AND Trabajadores.ControlNomina > 0"


    'Otro MAS. El tipo 3
    '   40 Horas semanales. 5 dias semana
    '
    'Con lo cual si en
    ' controlnomina
        ' 1.-   NORMAL ControlNomina >0 and ControlNomina <3
        ' 2.- Solo para el tipo  3
    If ControlNomina = 0 Then
        strControlNomina = " AND Trabajadores.ControlNomina >0  AND Trabajadores.ControlNomina <3"
    Else
        strControlNomina = " AND Trabajadores.ControlNomina = 3"
    End If



    
    '-------   Datos teroicos del mes
    Conn.Execute "Delete from tmpDatosMes"
    
    'Creamos todos los trabajadores con las horas y dias k
    'Deberian haber trabajado en el mes completo( y no esten de baja)
    SQL = "INSERT INTO tmpDatosMes(Mes,Trabajador,MesHoras,MesDias)"
    SQL = SQL & " SELECT " & Month(Fini) & ", Trabajadores.IdTrabajador, tmpHorasMesHorario.Horas, tmpHorasMesHorario.Dias"   ', Trabajadores.FecBaja"
    SQL = SQL & " FROM Trabajadores INNER JOIN tmpHorasMesHorario ON Trabajadores.IdHorario = tmpHorasMesHorario.idHorario"
    SQL = SQL & " WHERE (Trabajadores.FecBaja Is Null) "
    SQL = SQL & strControlNomina
    SQL = SQL & " AND (Trabajadores.FecAlta < #" & Format(Fini, FormatoFecha) & "#)"
    Conn.Execute SQL


    Set RS = New ADODB.Recordset
    
    'Ahora vemo los k entraron a trabajar este periodo.
    '¡Descontaremos de las horas laborables los dias k no han trabajado
    'o dicho de otra forma. Le contamos solo las horas k debia haber trabajado en fechas de alta
    SQL = "Select idTrabajador,idHorario,fecalta,fecbaja from Trabajadores WHERE"
    SQL = SQL & " fecalta >=#" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " and fecalta <=#" & Format(FFin, FormatoFecha) & "#"
    SQL = SQL & strControlNomina
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "INSERT INTO tmpDatosMes(Mes,Trabajador,MesHoras,MesDias) VALUES (" & Month(Fini) & ","
    While Not RS.EOF
    
        
    
        If IsNull(RS!fecbaja) Then
            FAux = FFin
        Else
            FAux = RS!fecbaja
            If FAux > FFin Then FAux = FFin
        End If
        Horas = CalculaHorasHorario(RS!IdHorario, D, RS!fecalta, FAux, False)
        
        
        Aux = RS.Fields!idTrabajador & "," & TransformaComasPuntos(CStr(Horas)) & "," & D & ")"
        Conn.Execute SQL & Aux
        RS.MoveNext
    Wend
    RS.Close
    
    
    'AHora vemos los k se han dado de baja en este periodo
    SQL = "Select idTrabajador,idHorario,fecalta,fecbaja from Trabajadores WHERE"
    SQL = SQL & " fecalta <#" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " AND fecbaja >=#" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " AND fecbaja <=#" & Format(FFin, FormatoFecha) & "#"
    SQL = SQL & strControlNomina
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "INSERT INTO tmpDatosMes(Mes,Trabajador,MesHoras,MesDias) VALUES (" & Month(Fini) & ","
    While Not RS.EOF
        Horas = CalculaHorasHorario(RS!IdHorario, D, Fini, RS!fecbaja, True)
        Aux = RS.Fields!idTrabajador & "," & TransformaComasPuntos(CStr(Horas)) & "," & D & ")"
        Conn.Execute SQL & Aux
        RS.MoveNext
    Wend
    RS.Close
    
 
    'Aquellos que entran de baja enfermedad durante este mes
    'Cambio 3 Diciembre
        '-----------------
        'Calcularemos los dias que tenia que haber trabajado,
        'no los que le faltban para completar el mes y leugo restar
        
        
    SQL = "Select bajas.*,trabajadores.idHorario from bajas,trabajadores where idtrab=idTrabajador"
    SQL = SQL & strControlNomina
    SQL = SQL & " AND fechabaja >=#" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " AND fechabaja <=#" & Format(FFin, FormatoFecha) & "#"
    SQL = SQL & " ORDER BY idtrabajador"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Aux = ""
    D = 0
    While Not RS.EOF
        If D <> RS!idTrab Then
            D = RS!idTrab
            Aux = Aux & D & "|"
        End If
        
        RS.MoveNext
    Wend
    RS.Close
    
    'Ya tengo los trabajadores. Ahora ire uno a uno por si han tenido mas dias de baja y eso
    While Aux <> ""
        
        D = InStr(1, Aux, "|")
        If D = 0 Then
            Aux = ""
        Else
            SQL = "Select bajas.*,trabajadores.idHorario from bajas,trabajadores where idtrab=idTrabajador"
            SQL = SQL & strControlNomina
            SQL = SQL & " AND fechabaja >=#" & Format(Fini, FormatoFecha) & "#"
            SQL = SQL & " AND fechabaja <=#" & Format(FFin, FormatoFecha) & "# AND idtrab = "
            SQL = SQL & Mid(Aux, 1, D - 1)
            SQL = SQL & " ORDER BY fechabaja"
            Aux = Mid(Aux, D + 1)
            RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            FAux = Fini
            
            D22 = 0
            h22 = 0
            HorasBaja = 0
            While Not RS.EOF
                IDH = RS!IdHorario
                IDT = RS!idTrab
                'Tramo anterior a la baja
                If RS!Fechabaja > FAux Then
                    FAux2 = DateAdd("d", -1, RS!Fechabaja)
                    Horas = CalculaHorasHorario(RS!IdHorario, D, FAux, FAux2, True)
                    h22 = h22 + Horas
                    D22 = D22 + D
                End If
                
                
                
                
                FAux2 = FFin
                If Not IsNull(RS!fechaalta) Then
                    If RS!fechaalta < FFin Then FAux2 = RS!fechaalta
                End If
                
                
                'Vemos si trabajao el dia de la baja
                Set vM = New CMarcajes
                                                            'Trabajo el dia de la baja
                If vM.Leer2(CLng(IDT), RS!Fechabaja) = 0 Then HorasBaja = HorasBaja + vM.HorasIncid
                Set vM = Nothing
                    
                
                'Insertamos en temporal de bajas para comprobar luego quien ha estado
                SQL = "INSERT INTO tmpCombinada(idTrabajador,Fecha,H1) VALUES (" & IDT
                SQL = SQL & ",#" & Format(RS!Fechabaja, FormatoFecha) & "#,#" & Format(FAux2, FormatoFecha) & "#)"
                Conn.Execute SQL
                        
                
                
                RS.MoveNext
            Wend
            RS.Close
            
            If FAux2 < FFin Then
                'Significa que aun trabaja algo a final del mes
                FAux = DateAdd("d", 1, FAux2)
                FAux2 = FFin
                Horas = CalculaHorasHorario(IDH, D, FAux, FAux2, True)
                h22 = h22 + Horas
                D22 = D22 + D
            End If

            SQL = "UPDATE tmpDatosMes SET meshoras= " & TransformaComasPuntos(CStr(h22))
            SQL = SQL & " , mesdias = " & D22
            'If horas de baja >0 siginifica que trabajo. Luego tiene que tener +hc y -hn
            If HorasBaja > 0 Then
                SQL = SQL & " , HorasN = horasN - " & TransformaComasPuntos(CStr(HorasBaja))
                SQL = SQL & " , HorasC = horasC - " & TransformaComasPuntos(CStr(HorasBaja))
            End If
            SQL = SQL & " WHERE mes= " & Month(Fini) & " AND Trabajador =" & IDT
            Conn.Execute SQL
        
        End If
    Wend
    
    'LA FECHA DE ALTA NOOOOOOOO se trabaja
    'Cmprobar el proceimiento de bajao
    'Aquellos que entraron de baja en dias anteriores al mes
    'Y se dieron de alta en el mes de calculo
    SQL = "Select bajas.*,trabajadores.idHorario,fechaalta as altaTrabajador from bajas,trabajadores where idtrab=idTrabajador"
    SQL = SQL & strControlNomina
    SQL = SQL & " AND fechabaja <#" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " AND fechaalta >=#" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " AND fechaalta <=#" & Format(FFin, FormatoFecha) & "#"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
    Aux = ""
    D = 0
    While Not RS.EOF
        If D <> RS!idTrab Then
            D = RS!idTrab
            Aux = Aux & D & "|"
        End If
        
        RS.MoveNext
    Wend
    RS.Close

    
    'Ya tengo los trabajadores. Ahora ire uno a uno por si han tenido mas dias de baja y eso
    While Aux <> ""
        
        D = InStr(1, Aux, "|")
        If D = 0 Then
            Aux = ""
        Else
    
            'Empieza a trabajar este mes despues de una baja. No hacemos nada
            SQL = "Select bajas.*,trabajadores.idHorario,FecAlta as altaTrabajador from bajas,trabajadores where idtrab=idTrabajador"
            SQL = SQL & strControlNomina
            SQL = SQL & " AND fechabaja <#" & Format(Fini, FormatoFecha) & "#"
            SQL = SQL & " AND fechaalta >=#" & Format(Fini, FormatoFecha) & "#"
            SQL = SQL & " AND fechaalta <=#" & Format(FFin, FormatoFecha) & "#"
            SQL = SQL & " AND idtrab = " & Mid(Aux, 1, D - 1)
            RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
            IDH = RS!IdHorario
            IDT = RS!idTrab
            
            Aux = Mid(Aux, D + 1)
            'Si se ha dado de
            FAux = Fini
            If RS!altaTrabajador > FAux Then FAux = RS!altaTrabajador
            FAux2 = RS!fechaalta
            
            RS.Close
            

            
            Horas = CalculaHorasHorario(IDH, D, FAux, FAux2, True)
        

            SQL = "INSERT INTO tmpCombinada(idTrabajador,Fecha,H1) VALUES (" & IDT
            SQL = SQL & ",#" & Format(FAux, FormatoFecha) & "#,#" & Format(FAux2, FormatoFecha) & "#)"
            Ejecuta SQL
            

            SQL = "UPDATE tmpDatosMes SET meshoras= meshoras - " & TransformaComasPuntos(CStr(Horas))
            SQL = SQL & " , mesdias =mesdias - " & D
            SQL = SQL & " WHERE mes= " & Month(Fini) & " AND Trabajador =" & IDT
            Conn.Execute SQL
        End If
      
            
    Wend
    
    
    
    
    
    'A titulo informtivo pondremos aquellos trabajadores
    'que estan de baja todavia. Es decir la fecha de alta es menor
    'ALTA temporada<inicio
    'baja temporada o null o >ffin
    'En bajas esta con la fecha de alta a null y fecha baja < finicio
    If ControlNomina = 0 Then
    
        'PARA QUE APAREZCAN LAS BAJAS EN EL MOMENTO
    
'        SQL = "SELECT Bajas.idTrab"
'        SQL = SQL & " FROM Trabajadores INNER JOIN Bajas ON Trabajadores.IdTrabajador = Bajas.idTrab"
'        SQL = SQL & " WHERE Bajas.FechaAlta Is Null AND Trabajadores.FecAlta<#" & Format(Fini, FormatoFecha) & "# AND"
'        SQL = SQL & " (Trabajadores.FecBaja Is Null  OR Trabajadores.Fecbaja>#" & Format(FFin, FormatoFecha) & "#) AND"
'        SQL = SQL & " (Bajas.Fechabaja Is Null  OR Bajas.Fechabaja<#" & Format(Fini, FormatoFecha) & "#)"
'
'        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'        SQL = "INSERT INTO tmpDatosMes(Mes,Trabajador,MesHoras,MesDias) VALUES (" & Month(Fini) & ","
'        While Not RS.EOF
'
'            aux = RS.Fields!idTrab & ",0,0)"
'            'Conn.Execute SQL & Aux
'            RS.MoveNext
'        Wend
'        RS.Close
    End If
    
    Set RS = Nothing
End Sub

Private Sub Ejecuta(SQL As String)
    On Error Resume Next
    Conn.Execute SQL
    If Err.Number <> 0 Then MuestraError Err.Number, Err.Description & vbCrLf & "El proceso continua"
End Sub

'Cojeremos y uniremos en tmpDatosMes todos los datos relativos a los trabajadores , para
'el periodo procesado anteriormente
Public Sub CombinaDatos(Fini As Date, FFin As Date)
Dim RS As ADODB.Recordset
Dim RT As ADODB.Recordset
Dim i As Integer
Dim Tot As Currency
'Dim Importe As Currency
Dim Aux As String
Dim SQL As String
Dim RS2 As ADODB.Recordset

    Set RS = New ADODB.Recordset
    SQL = "Select Trabajador,MEsDias from tmpDatosMes "
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQL = "SELECT tmpHoras.trabajador, tmpHoras.HorasT, tmpHoras.HorasC, tmpHoras.HorasE, tmpHoras.Dias, Trabajadores.bolsahoras"
    SQL = SQL & " FROM Trabajadores INNER JOIN tmpHoras ON Trabajadores.IdTrabajador = tmpHoras.trabajador"
    SQL = SQL & " WHERE tmpHoras.trabajador = "

    Set RT = New ADODB.Recordset
    While Not RS.EOF
        RT.Open SQL & RS.Fields(0), Conn, adOpenForwardOnly, adLockPessimistic, adCmdText

        If Not RT.EOF Then
            Aux = "UPDATE tmpDatosMes Set HorasN=" & TransformaComasPuntos(CStr(RT!horast))
            Aux = Aux & " ,HorasC=" & TransformaComasPuntos(CStr(RT!horasc))
            Aux = Aux & " ,HorasE=" & TransformaComasPuntos(CStr(RT!HorasE))
            Tot = RT!horasc + RT!horast
            Aux = Aux & " ,HorasT=" & TransformaComasPuntos(CStr(Tot))
            i = RS!mesdias - RT!Dias
            If i < 0 Then
                i = RS!mesdias
            Else
                i = RT!Dias
            End If
            Aux = Aux & " ,DiasTrabajados=" & i
            Aux = Aux & " ,BolsaAntes =" & TransformaComasPuntos(CStr(DBLet(RT!bolsahoras, "N")))
            'ANTES Tot = ObtenerAnticipos(FIni, FFin, Rs.Fields(0))
            Aux = Aux & " ,Anticipos = " & "0"    ' & TransformaComasPuntos(CStr(Tot))
            Aux = Aux & " WHERE trabajador = " & RS.Fields(0)
            Conn.Execute Aux
        Else
            'MIRARE SI TIENE BOLSA DE HORAS. Con lo cual puede que no haya trabajado NINGUN dia
            'pero si tenia bolsa le seguiremos generando dias
            RT.Close
            Aux = "SELECT Trabajadores.bolsahoras FROM Trabajadores WHERE idtrabajador = " & RS.Fields(0)
            RT.Open Aux, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            If DBLet(RT!bolsahoras, "N") > 0 Then
                Aux = "UPDATE tmpDatosMes SET "
                Aux = Aux & " BolsaAntes =" & TransformaComasPuntos(CStr(DBLet(RT!bolsahoras, "N")))
                Aux = Aux & " WHERE trabajador = " & RS.Fields(0)
                Conn.Execute Aux
            End If
        End If
        RT.Close
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    '----------------------------------------------------------------------------
    'Como la fecha de alta es la fecha de antiguedad como tal, no la fecha en la que empieza
    'a trabajar.
    'Para todos aquellos trabajadores k no han trabajado ningun dia ,y no estan de baja
    'Los elimino de la entrada de datos
    ' Lo quito pq  no debe borrar las entradas
    ' Falta revisar este trozo para meses posteriores
    '
'    SQL = "SELECT tmpDatosMEs.DiasTrabajados, Trabajadores.IdTrabajador, Trabajadores.FecAlta,"
'    SQL = SQL & " Trabajadores.FecBaja FROM tmpDatosMEs INNER JOIN Trabajadores ON tmpDatosMEs.Trabajador"
'    SQL = SQL & " = Trabajadores.IdTrabajador WHERE (((tmpDatosMEs.DiasTrabajados)=0) AND"
'    SQL = SQL & " ((Trabajadores.FecAlta)<#" & Format(Fini, FormatoFecha)
'    SQL = SQL & "#) AND ((Trabajadores.FecBaja) Is Null)) OR "
'    SQL = SQL & " (((Trabajadores.FecBaja)>#" & Format(Fini, FormatoFecha)
'    SQL = SQL & "#));"
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'    If Not RS.EOF Then
'        Set RS2 = New ADODB.Recordset
'        'Dejamos preparado el SQL
'        SQL = "SELECT Bajas.idTrab, Bajas.FechaAlta, Bajas.Fechabaja From Bajas"
'        SQL = SQL & " WHERE (((Bajas.FechaAlta)<#" & Format(Fini, FormatoFecha)
'        SQL = SQL & "#) AND ((Bajas.Fechabaja) Is Not Null))"
'        SQL = SQL & " AND idTrab = "
'
'        While Not RS.EOF
'            'Veo si es k esta de baja
'            RS2.Open SQL & RS.Fields(1), Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'            If RS2.EOF Then
'               'NO esta de baja
'                'Booramos la entrada                                        OJO RS no RS2
'                'Conn.Execute "DELETE FROM tmpDatosMes WHERE Trabajador = " & RS.Fields(1)
'            End If
'            RS2.Close
'            RS.MoveNext
'        Wend
'        Set RS2 = Nothing
'    End If
'    RS.Close
    
    
    Set RS = Nothing
End Sub




'Total horas y total dias
Public Sub CalculoDatosACompensar()
Dim RS As ADODB.Recordset
Dim i As Integer
Dim SQL As String

    SQL = "Select * from tmpDatosMes"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    While Not RS.EOF
        RS!saldoh = RS!horast - RS!meshoras
        RS!saldodias = RS!mesdias - RS!DiasTrabajados
        RS.Update
        
        'sgi
        RS.MoveNext
    Wend
    RS.Close
    Set RS = Nothing
        
End Sub


Public Sub HacerCompensaciones(FInicio As Date, FFin As Date, lbl As Label)
Dim HCompMes As Currency
Dim HPaBolsa As Currency
Dim DiasOF As Integer
Dim HorasOf As Currency
Dim H As Currency
Dim SQL As String
Dim ModoCompensacion_2 As String
Dim HorasJornadaRecuperacion As Currency
Dim Horario As Integer
Dim vH As CHorarios
Dim FESTIVOS As String
Dim MEDIODIA As String
Dim DiasReajusteXSTrabajados As Integer
    Dim RS As ADODB.Recordset


    'Vemos cual es el modo de compensacion
    '   0 .- NO compensa
    '   1 .- A partir de los dias trabajados del trabajador
    '         vemos cuantos dias le puedo compensar
    '   2 .- X horas hacen una jornada laboral a compensar
    '   3 .- Picassen cotubre 2008.
    '           -Compensaran por semana /dia con cuidado a los miercoles sabados
    '           -si trabaja una hora un dia, el resto de horas NO las tiene que compensar para la nomina
    SQL = "HorasJornada"
    ModoCompensacion_2 = DevuelveDesdeBD("RecuperacionDias", "Empresas", "idEmpresa", "1", "N", SQL)
    

    If ModoCompensacion_2 = "" Then
        ModoCompensacion_2 = "0"
        HorasJornadaRecuperacion = 0
    Else
        HorasJornadaRecuperacion = CCur(SQL)
    End If
    


    'De momento NO lo necesito
    If ModoCompensacion_2 = "3" Then
    
        'Fijo cual es ñla primera semana del mes, y la utima
        SemanaMesPrimera = Format(FInicio, "ww")
        SemanaMesUltima = Format(FFin, "ww")
    
        'Ajustes ponemos HN las que tiene menos las que sean extra
        lbl.Caption = "Ajuste horas normales"
        AjustarHorasNormales
        
        'Utlizaremos una tabla mas para guardar lios dias que XyS no deberean ser contabilizados como tal en nomina
        Conn.Execute "DELETE FROM tmpNOTrabajo"
        
        'VEmos el miercoles
        RecalculoHorasMiercolesSabados FInicio, FFin, lbl, True
        'Sabado
        RecalculoHorasMiercolesSabados FInicio, FFin, lbl, False

        'Vemos cuantos miercoles /sabado han trabajado pero no deben entrar en nomina
        lbl.Caption = "Procesar datos"
        lbl.Refresh
    End If

  
    SQL = "Select tmpDatosMes.*,idHorario,FecAlta,FecBaja,controlnomina from tmpDatosMes,Trabajadores"
    SQL = SQL & " WHERE tmpDatosMes.trabajador = Trabajadores.idTrabajador"
    SQL = SQL & " ORDER BY idHorario"
    Horario = -1
    FESTIVOS = ""
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    While Not RS.EOF
    
        If ModoCompensacion_2 = "1" Or ModoCompensacion_2 = "3" Then
            If Horario <> RS!IdHorario Then
                Set vH = Nothing
                Set vH = New CHorarios
                Horario = RS!IdHorario
                FESTIVOS = vH.LeerDiasFestivos(Horario, FInicio, FFin)
                MEDIODIA = vH.LeerMediosDias(Horario, FInicio, FFin)
            End If
        End If
            
        
        DiasReajusteXSTrabajados = 0
        'Dias trabajados y horas oficiales
        'If RS!saldodias > 0 Then
        If True Then
            'Me debe dias trabajados
            'Tengo k ver si en las horas que tiene tiene suficiente
            'Para esos dias trabajados. Si no no le compenso los dias
     
        
            
            'Primero. Veo si tiene bastantes horas para compensar los dias
            H = RS!horasc + RS!bolsaantes
            If (RS!horasn + H) >= RS!meshoras Then
            
                
            
                If ModoCompensacion_2 <> "3" Then
                        'Tiene bastantes horas para compensar el mes entero
                        DiasOF = RS!mesdias
                        HorasOf = RS!meshoras
                        
                        HCompMes = RS!meshoras - RS!horasn
                        If RS!horasc >= HCompMes Then
                            'Las coje todas de las compensadas de este mes
                            H = RS!horasc - HCompMes
                            HPaBolsa = RS!bolsaantes + H
                    
                        Else
                            H = HCompMes - RS!horasc
                            'Necesito h horas de la bolsa
                            HPaBolsa = RS!bolsaantes - H
                        End If
                Else
                    'Picassent 2008
                    'Stop
                    
                    HCompMes = H
                    DiasOF = CompensacionesDiaTrabajadoYSemana(RS!saldodias, RS, FESTIVOS, FInicio, FFin, vH, HorasJornadaRecuperacion, HCompMes, DiasReajusteXSTrabajados)
                    HPaBolsa = H - HCompMes
                    HorasOf = RS!horasn
                    DiasOF = RS!DiasTrabajados + DiasOF
                    
                End If
                
            Else
                
                'Si no tengo bastante le dejo la bolsa como esta
                'Y le pongo los dias k ha hecho, sin modificar
                
                If H = 0 Then   'Si NO tiene nada a compensar
                    DiasOF = RS!DiasTrabajados
                    HorasOf = RS!horast
                    HCompMes = 0   'Este mes no le quedara nada para compensar
                    HPaBolsa = 0
                Else
                
               
                   HPaBolsa = 0
                   'En funcion del tipo de compensacion
                   Select Case Val(ModoCompensacion_2)
                   Case 1, 2
                                     
                        HorasOf = RS!horasn + H
                    
                        'Vemos esas h -horas cuantos dias me puede compensar
                        If ModoCompensacion_2 = 2 Then
                            DiasOF = CuantosDiasCompensas(RS!saldodias, H, HorasJornadaRecuperacion)
                  
                        Else
                            'EN HorasJornadaRecuperacion: tengo el minimo de horas para que le compensen a una persona un dia sin llegar a las 8 horas
                            DiasOF = CompensacionesDiaTrabajado(RS!saldodias, H, RS, FESTIVOS, FInicio, FFin, vH, HorasJornadaRecuperacion)
                           
                        End If
                  
                        DiasOF = RS!DiasTrabajados + DiasOF
                        HCompMes = H
                   Case 3
                        'Nueva forma de compensar en PICASSENT. Oct 2008
                        
                        'EN HorasJornadaRecuperacion: tengo el minimo de horas para que le compensen a una persona un dia sin llegar a las 8 horas
                        HCompMes = H
                        DiasOF = CompensacionesDiaTrabajadoYSemana(RS!saldodias, RS, FESTIVOS, FInicio, FFin, vH, HorasJornadaRecuperacion, HCompMes, DiasReajusteXSTrabajados)
                        DiasOF = RS!DiasTrabajados + DiasOF
                        
                        'Aui esta la difencia
                        'Las horas que no se utilizan, no se compensan
                        HPaBolsa = H - HCompMes
                        HorasOf = RS!horast - RS!horasc
                        'Ahora
                        HorasOf = RS!horasn
                        
                   Case Else
                        'NOOOOOO compensamos nada
                        DiasOF = RS!DiasTrabajados
                        HPaBolsa = 0
                        HorasOf = RS!horast
                        HCompMes = 0
                   End Select
               
                End If
                
            End If
        Else
        
'            'Todo normal. A nivel de dias trabajados
'            DiasOF = RS!mesdias
'
'            'Ahora no van todas las horas, sololas compensadas
'            'HorasOf = RS!MesHoras
'            HorasOf = RS!horasn - RS!horasc
'            HPaBolsa = RS!bolsaantes + RS!horasc
'            HCompMes = 0
          
                        HCompMes = H
                        DiasOF = CompensacionesDiaTrabajadoYSemana(RS!saldodias, RS, FESTIVOS, FInicio, FFin, vH, HorasJornadaRecuperacion, HCompMes, DiasReajusteXSTrabajados)
                        DiasOF = RS!DiasTrabajados + DiasOF
                        
                        'Aui esta la difencia
                        'Las horas que no se utilizan, no se compensan
                        HPaBolsa = H - HCompMes
                        HorasOf = RS!horast - RS!horasc
          
          
          
          
            
        End If   'Diastrabajados no es igual k los k debia hber trabajado


        'Updateamos con los valores calculados
        SQL = "UPDATE tmpDatosMes SET"
        If RS!ControlNomina = 2 Then
            'No puede tener horas en bolsa
            HPaBolsa = 0
        End If
        
        SQL = SQL & "  BolsaDespues =" & TransformaComasPuntos(CStr(HPaBolsa))
        SQL = SQL & ", HorasPeriodo = " & TransformaComasPuntos(CStr(HorasOf))
        SQL = SQL & ", DiasPeriodo  = " & TransformaComasPuntos(CStr(DiasOF))
        If DiasReajusteXSTrabajados > 0 Then
            SQL = SQL & ", DiasTrabajados  = DiasTrabajados - " & DiasReajusteXSTrabajados
     
        End If
        'Para PICASSENT, machaco los datos
        SQL = SQL & ", HorasN = " & TransformaComasPuntos(CStr(HorasOf))
        
        'Las horas extras
        If HCompMes < 0 Then HCompMes = 0
        SQL = SQL & ", Extras = " & TransformaComasPuntos(CStr(HCompMes))
        'Trabajador
        SQL = SQL & " WHERE Trabajador = " & RS!Trabajador
        Conn.Execute SQL
        'sgi
        RS.MoveNext
    Wend
    RS.Close
    espera 0.5
    
    
    'AHora obtenemos los anticpos en NOMINA
    '-----------------------------------------
    SQL = "SELECT Trabajador,tmpDatosMEs.HorasN, tmpDatosMEs.extras, Categorias.Importe1, Categorias.Importe2, Trabajadores.PorcSS, Trabajadores.PorcIRPF"
    SQL = SQL & " FROM tmpDatosMEs INNER JOIN (Categorias INNER JOIN Trabajadores ON Categorias.IdCategoria = Trabajadores.idCategoria) ON tmpDatosMEs.Trabajador = Trabajadores.IdTrabajador"
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RS.EOF

        HorasOf = (RS!horasn * RS!Importe1) + (RS!extras * RS!Importe2)
        'Quitamos IRPF y SS
        H = (HorasOf * RS!porcirpf) + (HorasOf * RS!porcSS)
        H = Round((H / 100), 2)
        HorasOf = HorasOf - H
        SQL = "UPDATE tmpDatosMes SET"
        SQL = SQL & " Anticipos = " & TransformaComasPuntos(CStr(HorasOf))
        'Trabajador
        SQL = SQL & " WHERE Trabajador = " & RS!Trabajador
    
        Conn.Execute SQL
    
        'Sig
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    
    
    
    Set RS = Nothing
    
    

End Sub


Private Function MiercolesSabadoNoCuentaTrabajado(T As String, F As Date) As Boolean
Dim RT As ADODB.Recordset
Dim C As String

    C = "Select * from tmpNoTrabajo where idtra=" & T & " AND idFech=#" & Format(F, "yyyy/mm/dd") & "#"
    Set RT = New ADODB.Recordset
    RT.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    MiercolesSabadoNoCuentaTrabajado = RT.EOF
    RT.Close
    Set RT = Nothing
        
End Function








Private Function CuantosDiasCompensas(Dias As Integer, HorasCompensar As Currency, HorasJornadaCompensable As Currency) As Integer
Dim i As Integer
    'Compensamos dias a partir de HorasJornada  horas trabajadas
    i = CInt(HorasCompensar / HorasJornadaCompensable)
    If i > Dias Then i = Dias
    CuantosDiasCompensas = i
End Function



Private Function CompensacionesDiaTrabajado(Dias As Integer, HorasCompensar As Currency, ByRef Rec As Recordset, ByRef FEST As String, ByVal FI As Date, ByVal FF As Date, ByRef vHO As CHorarios, HorasMinimoDia As Currency) As Integer
Dim RF As ADODB.Recordset
Dim Cad As String
Dim Fin  As Boolean
Dim Horas As Currency
Dim Sig As Boolean
Dim DiaC As Currency
Dim FechaReferencia As Date

On Error GoTo ECompensacionesDiaTrabajado

    CompensacionesDiaTrabajado = 0
    'Si fecha alta > fecha inicio mes enonces finicio mes=fecha alta
    If Rec!fecalta > FI Then FI = Rec!fecalta

    'Si fecha baja < fecha baja mes entonces finicio mes=fecha alta
    If Not IsNull(Rec!fecbaja) Then
        If Rec!fecbaja < FF Then FF = Rec!fecbaja
    End If

    Cad = "Select distinct(fecha) from marcajes"
    Cad = Cad & " WHERE Fecha >=#" & Format(FI, FormatoFecha) & "#"
    Cad = Cad & " AND Fecha <=#" & Format(FF, FormatoFecha) & "#"
    Cad = Cad & " AND idTrabajador = " & Rec!Trabajador

    Set RF = New ADODB.Recordset
    RF.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Horas = 0
    DiaC = 0
    Fin = False
    If Not RF.EOF Then
        '
        FechaReferencia = RF!Fecha
        While Not Fin
            
            If FI > FF Then
                Fin = True
                Sig = False 'Para k no mueva el recordset
            Else
                If FI = FechaReferencia Then
                    FI = DateAdd("d", 1, FI)
                    Sig = True
                Else
                    If InStr(1, FEST, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                        'Es un dia festivo
                        FI = DateAdd("d", 1, FI)
                        Sig = False
                    Else
                        'Es un dia k no ha  trabajado. Vemos cuantas horas son
                        vHO.Leer vHO.IdHorario, FI
                        'Ya tenog las horas k debia haber trabajado
                        If Horas + vHO.TotalHoras <= HorasCompensar Then
                            'Le puedo compensar este dia
                            DiaC = DiaC + vHO.DiaNomina
                            Horas = Horas + vHO.TotalHoras
                            
                        Else
                            'Este dia no se lo puedo compensar
                            'No hago nada
                        End If
                                     
                        FI = DateAdd("d", 1, FI)
                        Sig = False
                                     
                        'Por si acaso ya ha compensado todos los dias
                        If DiaC >= Dias Then
                            If DiaC > Dias Then DiaC = Dias
                            Fin = True
                        End If
                        
                        'Si no le kedan horas para compensar tampoco seguimos
                        If HorasCompensar - Horas < 3 Then Fin = True
                    End If
                End If
            
        
            End If
            If Sig Then
                If RF.EOF Then
                    'Deberiamos salir
                    'Stop
                Else
                    RF.MoveNext
                    'ANTES
                    'If RF.EOF Then Fin = True
                    If Not RF.EOF Then FechaReferencia = RF!Fecha
                End If
            End If
        Wend
    Else
        'NO HA TRABAJADO, pero tiene Horas de otros meses
        DiaC = HorasCompensar \ 8    'Cuantos dias de 8 horas le entran
        Horas = HorasCompensar - (DiaC * 8) 'Horas sobrantes
        If Horas >= HorasMinimoDia Then DiaC = DiaC + 1   'Veo si el resto me comepnsa un dia o no
        If DiaC >= Dias Then DiaC = Dias                  'NO puede compensar mas dias de los que pueden ir en nomina

    End If
    RF.Close
    Set RF = Nothing
    If DiaC > Int(DiaC) Then
        DiaC = Int(DiaC) + 1
        If DiaC > Dias Then DiaC = Dias
    End If
        
    CompensacionesDiaTrabajado = DiaC
        
    Exit Function
ECompensacionesDiaTrabajado:
    MuestraError Err.Number, "CompensacionesDiaTrabajado"

End Function









Public Sub AjustaDatosBajaMesEntero()
Dim SQL As String
    SQL = "UPDATE tmpDatosMes SET "
    SQL = SQL & " MesHoras=0, Mesdias = 0, SaldoH=0, SaldoDias=0,HorasPeriodo =0, BolsaDespues=0, DiasPeriodo=0"
    SQL = SQL & " WHERE (((tmpDatosMEs.DiasTrabajados)=0) AND ((tmpDatosMEs.HorasN)=0) AND ((tmpDatosMEs.HorasC)=0) AND ((tmpDatosMEs.bolsaAntes)=0)) ;"
    Conn.Execute SQL
End Sub




'Un trabajador, entre unas fechas, si ha trabajado
Public Function HaTrabajadoConBaja(ByRef R As ADODB.Recordset) As Boolean
Dim Rec As ADODB.Recordset
Dim SQL As String

    HaTrabajadoConBaja = False
    SQL = "Select * from Marcajes WHERE"
    SQL = SQL & " idTrabajador =" & R!idTrabajador
    SQL = SQL & " AND fecha >=#" & Format(R!Fecha, FormatoFecha) & "#"
    'Ambos inclusive de baja
    SQL = SQL & " AND fecha <=#" & Format(R!H1, FormatoFecha) & "#"
    Set Rec = New ADODB.Recordset
    Rec.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If Not Rec.EOF Then HaTrabajadoConBaja = True
    Rec.Close
    Set Rec = Nothing
        
End Function




'Calcula las horas trabajadas para los trabajadores k tiene la marca puesta
Public Sub CalculaHorasTrabajadasConEXTRAS(Fini As Date, FFin As Date, ControlNomina As Byte)
Dim FAux As Date
Dim FAux2 As Date
Dim RS As ADODB.Recordset
Dim RS2 As ADODB.Recordset
Dim Dias As Currency
Dim Trab As Long
Dim Aux As String
Dim SQL As String
Dim vH As CHorarios
Dim FESTIVOS As String
Dim MEDIODIA As String
Dim strControlNomina As String
Dim Horario As Integer
Dim HC As Currency
Dim HE As Currency

    'IMPORTANTE
    'Ahora hay un control nomina mas, k es el 2
    'El tipo de control 2: Tiene un suledo fijo al mes
    'Pero en anticpos solo anticipa hNormales
    'luego el calculo de horas es el mismo que el 1
    ' por lo tanto donde ponia
        'SQL = SQL & " AND Trabajadores.ControlNomina = 1"
    ' pondra ahora
        'SQL = SQL & " AND Trabajadores.ControlNomina > 0"


    'Otro MAS. El tipo 3
    '   40 Horas semanales. 5 dias semana
    '
    'Con lo cual si en
    ' controlnomina
        ' 1.-   NORMAL ControlNomina >0 and ControlNomina <3
        ' 2.- Solo para el tipo  3
    Select Case ControlNomina
    Case 0
        strControlNomina = " AND Trabajadores.ControlNomina >0  AND Trabajadores.ControlNomina <2 "
    Case 1
        strControlNomina = " AND Trabajadores.ControlNomina = 3"
    Case 2
        strControlNomina = " AND (Trabajadores.ControlNomina =1  OR Trabajadores.ControlNomina =3) "
    Case 3
        'Sera para el listado que se entraga a los trabbajdores en PICASSENT
        ' Es para los tipos 1,2,3
        strControlNomina = " AND Trabajadores.ControlNomina >0"
    Case Else
        strControlNomina = ""
    End Select
    


    Conn.Execute "Delete from tmpHoras"
    
    'Calculamos las horas para el mes
    'Primero las normales con un simple insert into
    SQL = "INSERT INTO tmpHoras(trabajador,HorasT) "
    SQL = SQL & "SELECT Marcajes.idTrabajador, Sum(Marcajes.HorasTrabajadas) AS SumaDeHorasTrabajadas"
    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
    SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
    
    SQL = SQL & strControlNomina
    SQL = SQL & " GROUP BY Marcajes.idTrabajador;"
    Conn.Execute SQL
    
    
    
    '----HORAS COMPENSAR
    'Las horas para la bolsa de trabajor
    SQL = "SELECT Marcajes.idTrabajador,Marcajes.Horasincid,Fecha,Trabajadores.idHorario,IncFinal"
    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
    SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
    SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
    SQL = SQL & strControlNomina
    

    SQL = SQL & " ORDER BY idHorario,Marcajes.idTrabajador,Fecha"
    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Horario = -1
    Trab = -1
    Set vH = New CHorarios
    While Not RS.EOF
        If RS!IdHorario <> Horario Then
            
            If vH.Leer(RS!IdHorario, Now) = 0 Then
                FESTIVOS = vH.LeerDiasFestivos(vH.IdHorario, Fini, FFin)
                MEDIODIA = vH.LeerMediosDias(vH.IdHorario, Fini, FFin)
            Else
                MsgBox "Error leyendo datos del horario:" & RS!IdHorario & ". El programa finalizara", vbExclamation
                Exit Sub
            End If
            Horario = RS!IdHorario
        End If
        
        If Trab <> RS!idTrabajador Then
            If Trab <> -1 Then
                If Dias > Int(Dias) Then
                    Dias = Int(Dias) + 1
                Else
                    Dias = Int(Dias)
                End If
                UpdateaHoras HC, HE, Dias, Trab
            End If
        
        
            HE = 0
            HC = 0
            Trab = RS!idTrabajador
            Dias = 0
                    
        End If
        
        'Si el dia esta en FESTIVOS no lo sumo
        Aux = Format(RS!Fecha, "dd/mm/yyyy") & "|"

        
        If InStr(1, FESTIVOS, Aux) = 0 Then
            'Si es medio dia sumo medio
            'NO esta en festivos  'NO esta en festivos   'NO esta en festivos  'NO esta en festivos
            If RS!IncFinal = MiEmpresa.IncHoraExtra Then
                HC = HC + RS!HorasIncid
            End If
                
            If InStr(1, MEDIODIA, Aux) > 0 Then
                Dias = Dias + 0.5
            Else
                Dias = Dias + 1
            End If
        Else
            'FIESTA
            HE = HE + RS!HorasIncid
            

        End If

        
        RS.MoveNext
    Wend
    RS.Close
    
    
    
    'Updatemaos el ultimo
    If Trab > 0 Then
        If Dias > Int(Dias) Then
            Dias = Int(Dias) + 1
        Else
            Dias = Int(Dias)
        End If
        UpdateaHoras HC, HE, Dias, Trab
    End If
    
'    'Updatemos con los dias trabajados.
'    '
'    'Acciones:
'    '       -En una variable cargaremos los dias festivos de
'    '       -En Otra Cargaremos los medios dias.
'    '       -Para cada dia trabajado, para cada trabajador, veremos
'    '       - Si los dias trabajados es un festivo o unidad fraccionarai
'
'    SQL = "SELECT idHorario"
'    SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
'    SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
'    SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
'    SQL = SQL & strControlNomina
'    SQL = SQL & " GROUP BY Trabajadores.idHorario;"
'    RS.Open SQL, Conn, adOpenForwardOnly, adLockOptimistic, adCmdText
'
'
'    While Not RS.EOF
'        Set vH = New CHorarios
'        If vH.Leer(RS!IdHorario, Now) = 0 Then
'            FESTIVOS = vH.LeerDiasFestivos(vH.IdHorario, Fini, FFin)
'            MEDIODIA = vH.LeerMediosDias(vH.IdHorario, Fini, FFin)
'
'
'            'AHora para cada trabajador k haya trabajado entre las fechas le sumare dias trabajados
'            '.... o no(ej festivo)
'            SQL = "SELECT Marcajes.*"
'            SQL = SQL & " FROM Trabajadores INNER JOIN Marcajes ON Trabajadores.IdTrabajador = Marcajes.idTrabajador"
'            SQL = SQL & " Where Marcajes.Fecha >= #" & Format(Fini, FormatoFecha) & "#"
'            SQL = SQL & " and Marcajes.Fecha <= #" & Format(FFin, FormatoFecha) & "#"
'            SQL = SQL & strControlNomina
'            SQL = SQL & " And Trabajadores.IdHorario = " & RS!IdHorario
'            SQL = SQL & " ORDER BY Marcajes.idTrabajador, Fecha"
'            Set RS2 = New ADODB.Recordset
'            RS2.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'
'            If Not RS2.EOF Then
'                Trabajador = -1
'                Do
'                   If Trabajador <> RS2!idTrabajador Then
'
'                        If Trabajador > 0 Then
'
'                            SQL = "UPDATE tmpHoras Set Dias = "
'                            If Dias > Int(Dias) Then
'                                Dias = Int(Dias) + 1
'                            Else
'                                Dias = Int(Dias)
'                            End If
'                            SQL = SQL & Int(Dias)
'                            SQL = SQL & " WHERE Trabajador = " & Trabajador
'                            Conn.Execute SQL
'                        End If
'
'                        Trabajador = RS2!idTrabajador
'                        Dias = 0
'
'                    End If
'
'
'                    'Sig
'                    RS2.MoveNext
'                Loop Until RS2.EOF
'
'                'Ahora faltara por hacer el ultimo trabajador
'                SQL = "UPDATE tmpHoras Set Dias = "
'                If Dias > Int(Dias) Then
'                    Dias = Int(Dias) + 1
'                Else
'                    Dias = Int(Dias)
'                End If
'                SQL = SQL & Int(Dias)
'                SQL = SQL & " WHERE Trabajador = " & Trabajador
'                Conn.Execute SQL
'            End If
'        End If
'        RS.MoveNext 'Siguiente horario
'    Wend
'
        
    'Por si acaso algun trabajador tiene numeros negativos
    SQL = "UPDATE tmpHoras Set Dias = 0"
    SQL = SQL & " WHERE Dias < 0 "
    Conn.Execute SQL
    
    Set RS = Nothing
End Sub

Private Sub UpdateaHoras(ByRef vHC As Currency, ByRef vHE As Currency, ByRef vDias As Currency, ByRef T As Long)
Dim SQL As String

        SQL = "UPDATE tmpHoras Set HorasC = " & TransformaComasPuntos(CStr(vHC))
        SQL = SQL & " ,HorasE =  " & TransformaComasPuntos(CStr(vHE))
        SQL = SQL & " ,Dias =  " & vDias
        '
        vHC = vHC + vHE
        SQL = SQL & " ,HorasT = HorasT - " & TransformaComasPuntos(CStr(vHC))
        SQL = SQL & " WHERE Trabajador = " & T
        Conn.Execute SQL
End Sub



Public Sub PonHorasExtraDeBolsa()
        Conn.Execute "UPDATE tmpDatosMEs set ExtrasPeriodo = HorasE + Bolsadespues"
        espera 0.2
        Conn.Execute "UPDATE tmpDatosMEs set Bolsadespues=0"
End Sub



'Este sub hay k mejorarlo , de moento esta asi pq es para uno solo
'Esta puesto asi para PICASSENT
Public Sub GenerarLiquidacionCompensables(FI As Date, FF As Date)
Dim SQLAUX As String
Dim RT As ADODB.Recordset
Dim H As Currency
Dim Def As Currency
Dim D As Integer


    


    SQLAUX = "SELECT Trabajadores.IdTrabajador, tmpHoras.HorasT, tmpHoras.HorasC, Trabajadores.ControlNomina, Trabajadores.IdHorario"
    SQLAUX = SQLAUX & " FROM tmpHoras INNER JOIN Trabajadores ON tmpHoras.trabajador = Trabajadores.IdTrabajador"
    SQLAUX = SQLAUX & " WHERE (((Trabajadores.ControlNomina)=2));"
    
    Set RT = New ADODB.Recordset

    RT.Open SQLAUX, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
      
    While Not RT.EOF
            H = CalculaHorasHorario(RT!IdHorario, D, FI, FF, False)
            'Horas k faltan para completar las horas oficiales
            H = H - RT!horast
            If H > 0 Then
                If H <= RT!horasc Then
                    Def = H
                Else
                    Def = RT!horasc
                End If
            Else
                Def = 0

            End If

            SQLAUX = "UPDATE tmpHoras Set Horast=horast + " & TransformaComasPuntos(CStr(Def))
            SQLAUX = SQLAUX & " , Horasc=horasc - " & TransformaComasPuntos(CStr(Def))
            SQLAUX = SQLAUX & " WHERE Trabajador = " & RT!idTrabajador

            'MOVEMOS AL SIGUIENTE
            RT.MoveNext
            If Def > 0 Then Conn.Execute SQLAUX


    Wend

    RT.Close


    SQLAUX = "SELECT Trabajadores.IdTrabajador, tmpHoras.HorasT, tmpHoras.HorasC, Trabajadores.ControlNomina, Trabajadores.IdHorario"
    SQLAUX = SQLAUX & " FROM tmpHoras INNER JOIN Trabajadores ON tmpHoras.trabajador = Trabajadores.IdTrabajador"
    SQLAUX = SQLAUX & " WHERE (((Trabajadores.ControlNomina)=1));"
    
    RT.Open SQLAUX, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
            
            
            'PICASSENT
            H = RT!horasc
            If H > 0 Then
                Def = H
                
            Else
                Def = 0
                
            End If
            
            SQLAUX = "UPDATE tmpHoras Set Horast=horast - " & TransformaComasPuntos(CStr(Def))
''            SQLAUX = SQLAUX & " , Horasc=horasc - " & TransformaComasPuntos(CStr(Def))
            SQLAUX = SQLAUX & " WHERE Trabajador = " & RT!idTrabajador
            
            'MOVEMOS AL SIGUIENTE
            RT.MoveNext
            If Def > 0 Then Conn.Execute SQLAUX
            
            
    Wend
    RT.Close
    
    
    SQLAUX = "SELECT Trabajadores.IdTrabajador, tmpHoras.HorasT, tmpHoras.HorasC, Trabajadores.ControlNomina, Trabajadores.IdHorario"
    SQLAUX = SQLAUX & " FROM tmpHoras INNER JOIN Trabajadores ON tmpHoras.trabajador = Trabajadores.IdTrabajador"
    SQLAUX = SQLAUX & " WHERE horast<0;"
        
    RT.Open SQLAUX, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    SQLAUX = ""
    Def = 0
    While Not RT.EOF
        SQLAUX = SQLAUX & Format(RT!idTrabajador, "00000") & "    "
        Def = Def + 1
        If Def > 5 Then
            SQLAUX = SQLAUX & vbCrLf
            Def = 0
        End If

        RT.MoveNext
    Wend
    RT.Close
    If SQLAUX <> "" Then
        MsgBox "Hay trabajadores con horas negativas. Consulte soporte tecnico", vbExclamation
    End If
    
    
    
    
    
    'Para picassent. Control nomina =1

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    Set RT = Nothing
    
End Sub



'------------------------------------------------
'
' El objetivo final de los trabajadores semana
' es k trabajan 5 dias a la semans 8 horas
' Por lo tanto, no va a poder trabajar a la semana mas horas de
' las k las oficiales. Por lo tanto , con este sub
' Revisamos k las horas trabajadas no son mas de:
' dias * 8
' Si fuera mayor le sumariamos la diferencia a la bolsa k tuviera
' y en horas como mucho pondriamos
Public Sub HacerCompensacionSememana(FI As Date, FF As Date)
Dim SQLAUX As String
Dim RT As ADODB.Recordset

Dim H As Currency
Dim Def As Currency
Dim D As Integer


    SQLAUX = "SELECT  *"
    SQLAUX = SQLAUX & " , [diasperiodo]*8 AS Expr1, [HorasPeriodo]-[expr1] AS Diferencia"
    SQLAUX = SQLAUX & " FROM tmpDatosMes"
    D = Month(FI)
    SQLAUX = SQLAUX & " WHERE mes = " & D

    
    Set RT = New ADODB.Recordset

    RT.Open SQLAUX, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    While Not RT.EOF
            If RT!Diferencia > 0 Then
                Def = Abs(RT!Diferencia)
                SQLAUX = "UPDATE tmpdatosmes SET HorasPeriodo = " & TransformaComasPuntos(CStr((RT!expr1)))
                Def = Def + RT!bolsadespues
                SQLAUX = SQLAUX & " , BolsaDespues = " & TransformaComasPuntos(CStr(Def))
                SQLAUX = SQLAUX & " WHERE Mes = " & D & " AND Trabajador = " & RT!Trabajador
            Else
                SQLAUX = ""
            End If
            RT.MoveNext
            If SQLAUX <> "" Then Conn.Execute SQLAUX
    Wend
    RT.Close
    Set RT = Nothing
    
End Sub




'Obtener anticpos pagados
'Pondremos los tipos
'               0, Pagos
'               1.- Anticpos

Public Sub ObtenerAnticposPagadosPorPrograma(FI As Date, FF As Date)
Dim SQLAUX As String
Dim RT As ADODB.Recordset

    SQLAUX = "Select sum(importe) as impor,trabajador from pagos where tipo <2 "
    SQLAUX = SQLAUX & " AND fecha>=#" & Format(FI, FormatoFecha)
    SQLAUX = SQLAUX & "# AND fecha<=#" & Format(FF, FormatoFecha) & "#"
    SQLAUX = SQLAUX & " GROUP BY trabajador"
    Set RT = New ADODB.Recordset
    RT.Open SQLAUX, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RT.EOF
        SQLAUX = "UPDATE tmpDatosMEs SET Anticipos=" & TransformaComasPuntos(RT!Impor)
        SQLAUX = SQLAUX & " WHERE trabajador = " & RT!Trabajador
        Conn.Execute SQLAUX
        RT.MoveNext
    Wend
    RT.Close
    Set RT = Nothing
End Sub


Public Sub CalculaDiferenciasDiasHoras()
Dim SQLAUX As String


    SQLAUX = "UPDATE tmpDatosMes SET SaldoH = Meshoras - HorasN, "
    SQLAUX = SQLAUX & " SaldoDias= MesDias - DiasTrabajados"
    SQLAUX = SQLAUX & " ,DiasPeriodo = DiasTrabajados"
    Conn.Execute SQLAUX
End Sub




'--------------------------------
' EN tipo Alz la bolsa de horas
' pas directamente las HORASC as bolsa de horas
Public Sub ValoresBolsaDespues()
Dim SQLAUX As String

Dim RT As ADODB.Recordset
Dim I1 As Currency
Dim i2 As Currency
Dim Bruto As Currency
    SQLAUX = "SELECT Trabajadores.bolsahoras, Trabajadores.bolsaBRUTO, Trabajadores.IdTrabajador"
    SQLAUX = SQLAUX & " ,Trabajadores.bolsaNETO, tmpDatosMEs.HorasC, Categorias.Importe2"
    SQLAUX = SQLAUX & " ,Trabajadores.porcss,Trabajadores.porcIRPF"
    SQLAUX = SQLAUX & " FROM tmpDatosMEs INNER JOIN (Categorias INNER JOIN Trabajadores ON"
    SQLAUX = SQLAUX & " Categorias.IdCategoria = Trabajadores.idCategoria) ON tmpDatosMEs.Trabajador = Trabajadores.IdTrabajador;"

    Set RT = New ADODB.Recordset
    RT.Open SQLAUX, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    While Not RT.EOF
        
        'Bolsa despues
        i2 = DBLet(RT!bolsahoras, "N") + RT!horasc
        'SQL
        SQLAUX = "UPDATE tmpDatosMes SET bolsadespues = " & TransformaComasPuntos(CStr(i2))
        
        'La bolsa k importe supone bruto
        I1 = RT!horasc * RT!Importe2
        I1 = Round(I1, 2)
        Bruto = I1
        i2 = DBLet(RT!bolsabruto, "N")
        i2 = i2 + I1
        SQLAUX = SQLAUX & ",brutodespues = " & TransformaComasPuntos(CStr(i2))
        
        'El neto
        i2 = DBLet(RT!porcSS, "N") + DBLet(RT!porcirpf, "N")
        i2 = i2 / 100
        i2 = Round(Bruto * i2, 2)
        i2 = Bruto - i2
        
        i2 = i2 + DBLet(RT!bolsaneto, "N")
        SQLAUX = SQLAUX & ",netodespues = " & TransformaComasPuntos(CStr(i2))
    
    
        'idTrabajador
        SQLAUX = SQLAUX & " WHERE Trabajador = " & RT!idTrabajador
        
        RT.MoveNext
        'Ejecutamos
        Conn.Execute SQLAUX
    Wend
    RT.Close
    Set RT = Nothing
End Sub



'-------------------------------------------------------------------------------------------------------------------
'-------------------------------------------------------------------------------------------------------------------
'
'   Compensacion PICASSENT a partir de octubre 2008
'
'   -La compensación de horas extras se realizar siempre por días que hayan faltado al trabajo y
'    NUNCA para completar horas de los días que no hayan realizado las 8 horas.
'
'    -El minimo de horas necesario para compensar un día de trabajo es de 8 horas.
'
'   - Si un dia, las horas minimo no llegan al minimo por dia NO entra en nomina

'las horas a compensar se pasan por referencia
Private Function CompensacionesDiaTrabajadoYSemana(Dias As Integer, ByRef Rec As Recordset, ByRef FEST As String, ByVal FI As Date, ByVal FF As Date, ByRef vHO As CHorarios, HorasMinimoDia2 As Currency, ByRef HorasCompensarQueUtiliza As Currency, ByRef DiasQueReajusteXSTrabajados As Integer) As Integer
Dim RF As ADODB.Recordset
Dim Cad As String
Dim Fin  As Boolean
Dim Horas As Currency
Dim Sig As Boolean
Dim DiaC As Currency
Dim FechaReferencia As Date
Dim DiaSem As Byte
Dim Semana As Integer
Dim Js As Integer
Dim TrabajaMier As Boolean
Dim TrabajaSab As Boolean
Dim CompensaMitad As Boolean
Dim vFechaINicio As Date
Dim vfechaFin As Date
Dim HorasCompensar2 As Currency
Dim Rhoras As ADODB.Recordset
Dim XS_NoCuentaTrabajados As String
Dim XS_NoCtanYPuedoCompensarlos As String
Dim HorasTrabajadas As Currency
Dim BAJAS As String
Dim SemanaNormal As Boolean   'Si tiene la semana miercoles y sabado. Para la primera del mes
On Error GoTo ECompensacionesDiaTrabajado

    CompensacionesDiaTrabajadoYSemana = 0
    'Si fecha alta > fecha inicio mes enonces finicio mes=fecha alta
    If Rec!fecalta > FI Then FI = Rec!fecalta
    vFechaINicio = FI
    
    'Si fecha baja < fecha baja mes entonces finicio mes=fecha alta
    If Not IsNull(Rec!fecbaja) Then
        If Rec!fecbaja < FF Then FF = Rec!fecbaja
    End If
    vfechaFin = FF
    
    DiasQueReajusteXSTrabajados = 0
    HorasCompensar2 = HorasCompensarQueUtiliza
    XS_NoCuentaTrabajados = ""

    If Rec!Trabajador = 10 Or Rec!Trabajador = 10 Then Stop
    Set RF = New ADODB.Recordset
    
    
    'Las bajas
    
    Cad = "Select * from tmpcombinada where idtrabajador=" & Rec!Trabajador
    RF.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    BAJAS = ""
    If Not RF.EOF Then
        
        BAJAS = "|"
        While Not RF.EOF
            FechaReferencia = RF!Fecha
            Do
                BAJAS = BAJAS & Format(FechaReferencia, "dd/mm/yyyy") & "|"
                FechaReferencia = DateAdd("d", 1, FechaReferencia)
            Loop Until FechaReferencia > RF!H1
            RF.MoveNext
        Wend
    End If
    RF.Close


    Cad = "Select fecha,horastrabajadas from marcajes"
    Cad = Cad & " WHERE Fecha >=#" & Format(FI, FormatoFecha) & "#"
    Cad = Cad & " AND Fecha <=#" & Format(FF, FormatoFecha) & "#"
    Cad = Cad & " AND idTrabajador = " & Rec!Trabajador
    Cad = Cad & " ORDER BY fecha"
    RF.Open Cad, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    
    
    
    Horas = 0
    DiaC = 0
    Fin = False
    Semana = 0
    Set Rhoras = New ADODB.Recordset
    HorasTrabajadas = 0
    
    If Not RF.EOF Then
        'PRIMERA PASADA--------------------------------------------------------------------------
        
        
        FechaReferencia = RF!Fecha
        If HorasCompensar2 - Horas < HorasMinimoDia2 Then Fin = True
        While Not Fin
            
            DiaSem = Weekday(FechaReferencia, vbMonday)
            If DiaSem = 3 Or DiaSem = 6 Then
                Js = CInt(Format(FechaReferencia, "ww"))
            End If
            
            If Js <> Semana Then
                'HA cambiado de semana
                If Semana > 0 Then
                               
                    SemanaNormal = True
                    'Tenemos que ver si la primera semana tiene solo sabado
                    If Semana = SemanaMesPrimera Then
                         DiaSem = Weekday(vFechaINicio, vbMonday)
                         If DiaSem > 3 Then
                            'SOLO TIENE SABADO
                            SemanaNormal = False
                             If Not TrabajaSab Then Stop
                            
                            
                        End If
                    End If
                
                
                    'Vemos si compensa alguno de los dias
                    'Ha trabajado UNO de los dos dias solo
                    If SemanaNormal Then
                        If Not TrabajaSab And Not TrabajaMier Then
                            DiaC = DiaC + 1
                            Horas = Horas + 8
                            If DiaC > Dias Then DiaC = Dias 'NO puede compensar mas dias de los que pueden ir en nomina
                            If HorasCompensar2 - Horas < 8 Then Fin = True
                        Else
                            'Si entre los dos dias no suma 3.5 tb le compensaremos
                            
                            If HorasTrabajadas < HorasMinimoDia2 Then   'HorasMinimoDia2=3.5
                                MsgBox "CompensacionesDiaTrabajadoYSemana: " & Rec!Trabajador & " "
                                
                                'Le debemos compensar las horas hasta
                                HorasTrabajadas = HorasMinimoDia2 - HorasTrabajadas
                                Horas = Horas + HorasTrabajadas
                            End If
                        End If
                    End If
                Else
                    'Primera semana que trabaja. Luego tendremos que comprobar si
                    
                    
                    
                    
                End If
                TrabajaSab = False
                TrabajaMier = False
                HorasTrabajadas = 0
                Semana = Js
            End If
            If FI > FF Then
                Fin = True
                Sig = False 'Para k no mueva el recordset
            Else
                If FI = FechaReferencia Then
                    
                    If InStr(1, BAJAS, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                        'Estaba de baja
                        DiaSem = 0
                    Else
                        'Ha trabajado este dia. Compruebo si es X o S, para ver si compenso
                        DiaSem = Weekday(FechaReferencia, vbMonday)
                    End If
                    If DiaSem = 3 Or DiaSem = 6 Then
                    
       
                    
                        'Comprobaremos con el sQL si le cuento como trabajado o no
                        If DiaSem = 3 Then
                            Cad = "Select max(hora) from entradamarcajes where fecha = #"
                            Cad = Cad & Format(FI, "yyyy/mm/dd") & "# AND idTrabajador = " & Rec!Trabajador
                        Else
                            Cad = "Select min(hora) from entradamarcajes where fecha = #"
                            Cad = Cad & Format(FI, "yyyy/mm/dd") & "# AND idTrabajador = " & Rec!Trabajador
                        End If
                        Rhoras.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                        If Rhoras.EOF Then
                            MsgBox "Error grave: " & Cad, vbExclamation
                            End
                        End If
                        
                        
                        'Si que es miercoles o sabado
                        If DiaSem = 3 Then
                            If Rhoras.Fields(0) > CDate(HoraIntermediaMiercolesSabado) Then
                                TrabajaMier = True
                                HorasTrabajadas = HorasTrabajadas + RF!HorasTrabajadas
                            Else
                                XS_NoCuentaTrabajados = XS_NoCuentaTrabajados & FI & "|"
                            End If
                            
                            
                        Else
                            If Rhoras.Fields(0) < CDate(HoraIntermediaMiercolesSabado) Then
                                TrabajaSab = True
                                HorasTrabajadas = HorasTrabajadas + RF!HorasTrabajadas
                            Else
                                XS_NoCuentaTrabajados = XS_NoCuentaTrabajados & FI & "|"
                            End If
                        End If
                        Rhoras.Close
                    End If
                    FI = DateAdd("d", 1, FI)
                    Sig = True
                Else
                    If InStr(1, FEST, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                        'Es un dia festivo
                        FI = DateAdd("d", 1, FI)
                        Sig = False
                    Else
                        FI = DateAdd("d", 1, FI)
                        Sig = False
                     
                    End If
                End If
            
        
            End If
            If Sig Then
                If RF.EOF Then
                    'Deberiamos salir
                    MsgBox "Rs.eof: " & RF.Source, vbExclamation
                    
                Else
                    RF.MoveNext
                    'ANTES
                    'If RF.EOF Then Fin = True
                    If Not RF.EOF Then FechaReferencia = RF!Fecha
                End If
            End If
        Wend
    
    
        'Para la ultima semana y para la preimera
        'Veremos si se ha trabajado la ultima(primera) semana , y cuanto
        Js = CInt(Format(FF, "ww"))
        If Js <> Semana Then
                'HA cambiado de semana
                'Si estaba de baja vemos cuando acabo la baja
                If BAJAS <> "" Then
                    Cad = Right(BAJAS, 11)
                    Cad = Left(Cad, 10)
                    FI = CDate(Cad)
                    If FI >= FF Then Semana = 0 'Ha estado de baja hasta el final de mes
                    
                End If
                
                If Semana > 0 Then
                
                    'Ahora veremos si tenia que haber trabajado miercoles y/o sabado
                    DiaSem = Weekday(FF, vbMonday)
                    
                    If DiaSem > 2 Then   'El ultimo dia es miercoles o mas
                    
                        If DiaSem >= 6 Then
                            
                            'Vemos si compensa alguno de los dias
                            'Ha trabajado UNO de los dos dias solo
                            If Not TrabajaSab And Not TrabajaMier Then
                                DiaC = DiaC + 1
                                Horas = Horas + 8
                                If DiaC > Dias Then DiaC = Dias 'NO puede compensar mas dias de los que pueden ir en nomina
                                
                            Else
                                'Si entre los dos dias no suma 3.5 tb le compensaremos
                            
                                If HorasTrabajadas < HorasMinimoDia2 Then   'HorasMinimoDia2=3.5
                                    
                                    'Le debemos compensar las horas hasta
                                    HorasTrabajadas = HorasMinimoDia2 - HorasTrabajadas
                                    Horas = Horas + HorasTrabajadas
                                End If
                            End If
                            
                            
                        Else
                            'Enero 2009
                            'Esta ultima semana no tiene Sabado. Luego COMPENSAMOS  3.5 Horas del miercoles
                            DiaC = DiaC + 1
                            If DiaC > Dias Then DiaC = Dias 'NO puede compensar mas dias de los que pueden ir en nomina
                            Horas = Horas + 3.5
                            
                        
                            
                        End If
                    End If  'De diasem>2
                End If
        End If
    
    
    
    
    
    
        '--------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------
        '--------------------------------------------------------------------------------
        If XS_NoCuentaTrabajados <> "" Then
            'Significa que algunos de los dias que ha trabajado NO cuenta en nomina
    
           
            While XS_NoCuentaTrabajados <> ""
                DiaSem = InStr(1, XS_NoCuentaTrabajados, "|")
                FI = CDate(Mid(XS_NoCuentaTrabajados, 1, DiaSem - 1))
                XS_NoCuentaTrabajados = Mid(XS_NoCuentaTrabajados, DiaSem + 1)
                Fin = False
                'Veo si es miercoles o sabado
                DiaSem = Weekday(FI, vbMonday)
                If DiaSem = 3 Then
                    'miercoles compruebo si el sabado se ha trabajado
                    FechaReferencia = DateAdd("d", 3, FI)
                    If InStr(1, XS_NoCuentaTrabajados, FechaReferencia) Then
                        'si que trabjo el sabado siguiente, pero tampoco lo contamos en nomina
                        'luego descontamos un dia en nomina
                        Fin = True
                        'lo quito de procesar
                        DiaSem = InStr(1, XS_NoCuentaTrabajados, "|")
                        XS_NoCuentaTrabajados = Mid(XS_NoCuentaTrabajados, DiaSem + 1)
                    Else
                        'vemaos que lo trabajo aunque fuera normal
                        RF.MoveFirst
                        Sig = False  'Para saber si lo descontamos
                        While Not RF.EOF
                            If RF!Fecha < FechaReferencia Then
                                'no hacemos nada
                            Else
                                If RF!Fecha > FechaReferencia Then
                                    RF.MoveLast
                                Else
                                    'FEcha igual. HA trabajado el sabado en modo normal,
                                    'con lo cual NO descontare el dai en trabajados
                                    Sig = True
                                End If
                            End If
                            RF.MoveNext
                        Wend
                        If Not Sig Then Fin = True
                        
                   End If
                Else
                    '-------
                    'Sabado
                    FechaReferencia = DateAdd("d", -3, FI)
                    RF.MoveFirst
                    Sig = False
                    While Not RF.EOF
                        If RF!Fecha < FechaReferencia Then
                            'no hacemos nada
                        Else
                            If RF!Fecha > FechaReferencia Then
                                RF.MoveLast
                            Else
                                'FEcha igual. HA trabajado el sabado en modo normal,
                                'con lo cual NO descontare el dai en trabajados
                                Sig = True
                            End If
                        End If
                        RF.MoveNext
                    Wend
                    If Not Sig Then Fin = True
                End If
                'SI FIN siginifaca que tengo que descontar un dia enNOMINA
                If Fin Then
                    DiasQueReajusteXSTrabajados = DiasQueReajusteXSTrabajados + 1
                    XS_NoCtanYPuedoCompensarlos = XS_NoCtanYPuedoCompensarlos & Format(FI, "dd/mm/yyyy") & "|"
                End If
            Wend
            

        End If  ' De XS_NoCuentaTrabajados
    
        
    
    
    
    
    
    
    
    
    
    
    
        'SEGUNDA PASADA--------------------------------------------------------------------------
        'SEGUNDA PASADA--------------------------------------------------------------------------
        'SEGUNDA PASADA--------------------------------------------------------------------------
        'SEGUNDA PASADA--------------------------------------------------------------------------
        'busco compensar dias de 8 en 8 horas
        RF.MoveFirst
        FI = vFechaINicio
        FF = vfechaFin
        Fin = False
        

        'Si no tiene OCHO horas no puedo compensarle NI UN SOLO DIA mas
        If HorasCompensar2 - Horas < 8 Then Fin = True
        FechaReferencia = RF!Fecha
        While Not Fin
            
            If FI > FF Then
                Fin = True
                Sig = False 'Para k no mueva el recordset
            Else
                If FI = FechaReferencia Then
                    FI = DateAdd("d", 1, FI)
                    Sig = True
                    If RF!Fecha = FechaReferencia Then
                        If InStr(1, BAJAS, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                            'Ha trabajaod estando de baja
                            'El primer dia puede darse el caso
                            'Stop
                            'Debug.Print FI & " " & Rec!Trabajador
                        Else
                            
                            DiaSem = Weekday(FechaReferencia, vbMonday)
                            If DiaSem <> 3 And DiaSem <> 6 Then
                                If RF!HorasTrabajadas < HorasMinimoDia2 Then
                                    'Auqi compensamos
                                    'Le debemos compensar las horas hasta
                                    HorasTrabajadas = HorasMinimoDia2 - RF!HorasTrabajadas
                                    Horas = Horas + HorasTrabajadas
                                    'Si no le kedan horas para compensar tampoco seguimos
                                    If HorasCompensar2 - Horas < HorasMinimoDia2 Then Fin = True
                                End If
                            Else
                                'Podria ser que habienddo trabajado NO cuente. Ej. Trabaja miercoles de 9 a 2
                                If InStr(1, XS_NoCtanYPuedoCompensarlos, Format(FechaReferencia, "dd/mm/yyyy") & "|") > 0 Then
                                    vHO.Leer vHO.IdHorario, FechaReferencia
                                    'Ya tenog las horas k debia haber trabajado
                                    If Horas + vHO.TotalHoras <= HorasCompensar2 Then
                                    
                                        'Si el dia es de miercoles o sabado SI que quito las horas
                                        'Le puedo compensar este dia
                                        DiaC = DiaC + 1
                                        If DiaC > Dias Then DiaC = Dias 'NO puede compensar mas dias de los que pueden ir en nomina
                                        Horas = Horas + vHO.TotalHoras
                                     End If
                                    
                                    
                                End If
                            End If
                        End If
                    End If
                Else
                    If InStr(1, FEST, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                        'Es un dia festivo
                        FI = DateAdd("d", 1, FI)
                        Sig = False
                    Else
                        'Es de bajas
                        If InStr(1, BAJAS, Format(FI, "dd/mm/yyyy") & "|") > 0 Then
                            FI = DateAdd("d", 1, FI)
                            Sig = False
                        Else
                            DiaSem = Weekday(FI, vbMonday)
                            If DiaSem = 3 Or DiaSem = 6 Then
                                'YA LO HEMOS PROCESADO
                                'Pero podria ser que tiene a compensar. ha trabajado por la mañana cuando debia trabajar por la tarde
                                FI = DateAdd("d", 1, FI)
                                Sig = False
                            Else
                                    'Es un dia k no ha  trabajado. Vemos cuantas horas son
                                    vHO.Leer vHO.IdHorario, FI
                                    'Ya tenog las horas k debia haber trabajado
                                    If Horas + vHO.TotalHoras <= HorasCompensar2 Then
                                    
                                        'Si el dia es de miercoles o sabado SI que quito las horas
                                        'Le puedo compensar este dia
                                        DiaC = DiaC + vHO.DiaNomina
                                        
                                        Horas = Horas + vHO.TotalHoras
                                        
                                    Else
                                        'Este dia no se lo puedo compensar
                                        'No hago nada
                                    End If
                                                 
                                    FI = DateAdd("d", 1, FI)
                                    Sig = False
                                                 
                                    'Por si acaso ya ha compensado todos los dias
                                    If DiaC >= Dias Then
                                        If DiaC > Dias Then DiaC = Dias
                                        'Fin = True
                                        'No pongo el FIN, pq puede compensarles horas todavia
                                        
                                    End If
                                    
                                    'Si no le kedan horas para compensar tampoco seguimos
                                    If HorasCompensar2 - Horas < 8 Then Fin = True
                                    
                            End If 'de diasem
                        End If 'de bajas
                    End If
                End If
            
        
            End If
            If Sig Then
                If RF.EOF Then
                    'Deberiamos salir
                    'Stop
                Else
                    RF.MoveNext
                    'ANTES
                    'If RF.EOF Then Fin = True
                    If Not RF.EOF Then FechaReferencia = RF!Fecha
                End If
            End If
        Wend
        
        
        
        
        
        
    Else   'rt.eof
        'NO HA TRABAJADO, pero tiene Horas de otros meses
        DiaC = HorasCompensar2 \ 8    'Cuantos dias de 8 horas le entran
        Horas = HorasCompensar2 - (DiaC * 8) 'Horas sobrantes
        If Horas >= HorasMinimoDia2 Then DiaC = DiaC + 1   'Veo si el resto me comepnsa un dia o no
        If DiaC >= Dias Then DiaC = Dias                  'NO puede compensar mas dias de los que pueden ir en nomina

    End If
    RF.Close
    Set RF = Nothing
    Set Rhoras = Nothing
    If DiaC > Int(DiaC) Then
        DiaC = Int(DiaC) + 1
        If DiaC > Dias Then DiaC = Dias 'NO puede compensar mas dias de los que pueden ir en nomina
    End If
        
    HorasCompensarQueUtiliza = Horas
    CompensacionesDiaTrabajadoYSemana = DiaC
        
    Exit Function
ECompensacionesDiaTrabajado:
    MuestraError Err.Number, "CompensacionesDiaTrabajado"

End Function






Private Sub RecalculoHorasMiercolesSabados(F1 As Date, F2 As Date, vLbl As Label, Miercoles As Boolean)
Dim Cad As String
Dim RF As ADODB.Recordset
Dim Trab As Long
Dim HT As Currency
Dim Horas As Currency

    vLbl.Caption = "Recaluclo horas miercoles"
    vLbl.Refresh
    If Miercoles Then
        Trab = 4
    Else
        Trab = 7
    End If
    Cad = "SELECT EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha, Weekday([Fecha]) AS Expr1"
    Cad = Cad & " From EntradaMarcajes"
    Cad = Cad & " Where EntradaMarcajes.Fecha >= #" & Format(F1, "yyyy/mm/dd") & "# And"
    Cad = Cad & " EntradaMarcajes.Fecha <= #" & Format(F2, "yyyy/mm/dd") & "# And "
    Cad = Cad & " Weekday([Fecha]) = " & Trab & " And Hora "
    If Miercoles Then
        Cad = Cad & " <"
    Else
        Cad = Cad & " >"
    End If
    Cad = Cad & " #14:00:00# group by  EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha,  Weekday([Fecha])"
    Cad = Cad & " ORDER BY EntradaMarcajes.idTrabajador, EntradaMarcajes.Fecha"
    Set RF = New ADODB.Recordset
    RF.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Trab = -1
    While Not RF.EOF
    
        'If RF!idTrabajador = 73 Then Stop
    
        If Trab <> RF!idTrabajador Then
            vLbl.Caption = "R.H. (" & Val(RF!expr1) - 1 & ")  trab:" & RF!idTrabajador
            vLbl.Refresh
            'Nuevo trabajador
            If Trab > 0 Then
                'Updateamos los nuevos valores en tmphoras
                If HT > 0 Then UpdateaNuevosValoresMiercolesSabado Trab, HT, Month(RF!Fecha)
            End If
            'Reseteamos variables
            Trab = RF!idTrabajador
            HT = 0
            
        End If
        
        Horas = NuevoCalculoHorasDiaXS(Trab, RF!Fecha, Miercoles)
        HT = HT + Horas
        RF.MoveNext
        
    Wend
    RF.Close
    Set RF = Nothing
    'EL ultimo
    If Trab > 0 And HT > 0 Then UpdateaNuevosValoresMiercolesSabado Trab, HT, Month(F1)
End Sub
'falta pasar variable fecha
Private Sub UpdateaNuevosValoresMiercolesSabado(IdTra As Long, Hor As Currency, KMes As Integer)
Dim Cad As String


    Cad = "UPDATE tmpDatosMes SET horasN= horasN - " & TransformaComasPuntos(CStr(Hor))
    Cad = Cad & " , horasc=horasc +  " & TransformaComasPuntos(CStr(Hor))
    Cad = Cad & " WHERE mes= " & KMes & " AND Trabajador =" & IdTra
    Conn.Execute Cad
End Sub

Public Function NuevoCalculoHorasDiaXS(Trabajador As Long, Fecha As Date, DeMiercoles As Boolean) As Currency
Dim RH As ADODB.Recordset
Dim C As String
Dim T1 As Currency
Dim T2 As Currency
Dim E As Boolean
Dim Seguir As Boolean
Dim NuevaHora As Currency

    NuevoCalculoHorasDiaXS = 0
    Set RH = New ADODB.Recordset
    
    C = "Select * from entradamarcajes where idtrabajador=" & Trabajador
    C = C & " AND fecha = #" & Format(Fecha, "yyyy/mm/dd") & "# "
    
    'C = C & " AND fecha = #" & Format(Fecha, "yyyy/mm/dd") & "# AND hora "
    'If DeMiercoles Then
    '    C = C & "<"
    'Else
    '    C = C & ">"
    'End If
    'C = C & "#14:00:00# ORDER BY hora"
    C = C & " ORDER BY hora"

    RH.Open C, Conn, adOpenKeyset, adLockPessimistic, adCmdText
    
    Seguir = True
    'Si tiene algun --> Miercoles y es por la tarde ---->NORMAL, no hago nada
    '                     sabado   y es por la mañana ---> "
    If Not RH.EOF Then
        If DeMiercoles Then
       
            If RH!Hora > CDate(HoraIntermediaMiercolesSabado) Then
                Seguir = False
            Else
                'RH.MoveFirst
            End If
        Else
            RH.MoveLast
           'sabado
            If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                Seguir = False
            Else
                RH.MoveFirst
            End If
            
        End If
    End If
    
    
    If Not Seguir Then
        RH.Close
        Exit Function
    End If
    E = True
    NuevaHora = 0
    While Not RH.EOF

        If E Then
            If DeMiercoles Then
                If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                    T1 = CCur(DevuelveValorHora(RH!Hora))
                Else
                    'Ya es cuando le toca
                    RH.MoveLast
                End If
                
             Else
                If RH!Hora >= CDate(HoraIntermediaMiercolesSabado) Then
                    T1 = CCur(DevuelveValorHora(RH!Hora))
                Else
                    'Ya es cuando le toca
                  
                End If
             End If
                
        Else
            'Si tiene valor t1 calculamos dif
            If T1 > 0 Then
                T2 = CCur(DevuelveValorHora(RH!Hora))
                T1 = T2 - T1
            
                NuevaHora = NuevaHora + T1
                T1 = 0

            End If
        End If
        E = Not E
        RH.MoveNext
    Wend
        
        
    'veremos si este dia realmente es como si no lo trabajadra
    'ya que si tenia que haber venido por la mañana, pero solo viene por
    'la tarde a efectos de nomina no lo cuento
    
    If Trabajador < 700 Then
        If DeMiercoles Then
            RH.MoveLast
            'nos vamos al ultimo. Si la hora es mayor que las 2 NO loañado a la lista
            If RH!Hora < CDate(HoraIntermediaMiercolesSabado) Then
                'Este dia no lo contare para la nomina
                C = "INSERT INTO tmpNoTrabajo (idtra,idfech) VALUES (" & Trabajador & ",#" & Format(Fecha, "yyyy/mm/dd") & "#)"
                Conn.Execute C
            End If
        Else
            RH.MoveFirst
            If RH!Hora > CDate(HoraIntermediaMiercolesSabado) Then
                
                C = "INSERT INTO tmpNoTrabajo (idtra,idfech) VALUES (" & Trabajador & ",#" & Format(Fecha, "yyyy/mm/dd") & "#)"
                Conn.Execute C
            End If
         End If
    End If
    RH.Close    'para que no coja los 700 y 900
    If NuevaHora > 0 And Trabajador < 700 Then
        C = "Select * from marcajes where idtrabajador=" & Trabajador
        C = C & " AND fecha = #" & Format(Fecha, "yyyy/mm/dd") & "# "
        RH.Open C, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        If RH.EOF Then
            MsgBox "Mal: " & C
        Else
            
            
            
            
            'Si deberia tenre mas horas comepnsables
            
            If RH!IncFinal = MiEmpresa.IncRetraso Then
                'No hacemos nada, ya que el valor calculado sera el bueno
             
            Else
                If NuevaHora > RH!HorasIncid Then
                    T1 = NuevaHora - RH!HorasIncid
                    'Tenemos lo que incrementa y decrementa en comepnsables y en normales
                    'nunca puede aumentar mas que lo que ha trabajador
                    T2 = RH!HorasTrabajadas - RH!HorasIncid
                    If T1 > T2 Then T1 = T2
                    NuevaHora = T1
                Else
                   'No hacemos nada
                   NuevaHora = 0
                End If
                
            End If
            RH.Close
            NuevoCalculoHorasDiaXS = NuevaHora
        End If
    End If
End Function

Private Sub AjustarHorasNormales()
Dim RT As ADODB.Recordset
    
    Conn.Execute "UPDATE tmpdatosmes set horasn=horasn-horasc"
    espera 0.5
    Set RT = New ADODB.Recordset
    RT.Open "Select * from tmpdatosmes where horasn<0 ", Conn, adOpenForwardOnly, adLockOptimistic
    If Not RT.EOF Then
        MsgBox "Avise soporte tecnico." & vbCrLf & "HorasN<0", vbExclamation
    End If
    RT.Close
End Sub

'----------------------------
'dentro del nuevo proc
'                        If TrabajaSab Then
'                            'COMPENSA miercoles.
'                            'Habra que comprobar que el miercoles de esa semana
'                            'Esta en el periodo de calculo
'                            '
'                            CompensaMitad = True
'                            If CInt(Format(vFechaINicio, "ww")) = Js Then
'                                'Primera semana del calculo.
'                                'Si el dia es mayor que miercoles NO tiene miercoels a comepnsar
'                                DiaSem = Weekday(vFechaINicio, vbMonday)
'                                If DiaSem > 3 Then CompensaMitad = False                             'NO TIENE miercoles
'                            End If
'                            If CompensaMitad Then Horas = Horas + 3.5
'
'                        End If
'
'                        If TrabajaMier Then
'                            'COMPENSA SABADO.
'                            'Habra que comprobar que el sabado de esa semana
'                            'Esta en el periodo de calculo
'                            '
'                            CompensaMitad = True
'                            If CInt(Format(vfechaFin, "ww")) = Js Then
'                                'Primera semana del calculo.
'                                'Si el dia es mayor que miercoles NO tiene miercoels a comepnsar
'                                DiaSem = Weekday(vfechaFin, vbMonday)
'                                If DiaSem < 6 Then CompensaMitad = False                 'NO TIENE sabado
'                            End If
'                            If CompensaMitad Then Horas = Horas + 4.5
'
'                        End If
'                        If HorasCompensar2 - Horas < 4.5 Then Fin = True
'                    End If
'                End If

