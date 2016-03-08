Attribute VB_Name = "ModFunciones"
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
'   En este modulo estan las funciones que recorren el form
'   usando el each for
'   Estas son
'
'   CamposSiguiente -> Nos devuelve el el text siguiente en
'           el orden del tabindex
'
'   CompForm -> Compara los valores con su tag
'
'   InsertarDesdeForm - > Crea el sql de insert e inserta
'
'   Limpiar -> Pone a "" todos los objetos text de un form
'
'   ObtenerBusqueda -> A partir de los text crea el sql a
'       partir del WHERE ( sin el).
'
'   ModifcarDesdeFormulario -> Opcion modificar. Genera el SQL
'
'   PonerDatosForma -> Pone los datos del RECORDSET en el form
'////////////////////////////////////////////////////////////
'////////////////////////////////////////////////////////////
Option Explicit

Public Const ValorNulo = "Null"

Public Function CompForm(ByRef Formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Carga As Boolean
    Dim Correcto As Boolean
       
    CompForm = False
    Set mTag = New CTag
    For Each Control In Formulario.Controls
        'TEXT BOX
        If TypeOf Control Is TextBox Then
            Carga = mTag.Cargar(Control)
            If Carga = True Then
                Correcto = mTag.Comprobar(Control)
                If Not Correcto Then Exit Function
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        'COMBOBOX
        ElseIf TypeOf Control Is ComboBox Then
            'Comprueba que los campos estan bien puestos
            If Control.Tag <> "" Then
                Carga = mTag.Cargar(Control)
                If Carga = False Then
                    MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                    Exit Function
                    
                Else
                    If mTag.Vacio = "N" And Control.ListIndex < 0 Then
                            MsgBox "Seleccione una dato para: " & mTag.Nombre, vbExclamation
                            Exit Function
                    End If
                End If
            End If
        End If
    Next Control
    CompForm = True
End Function


Public Sub Limpiar(ByRef Formulario As Form)
    Dim Control As Object
    
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Control.Text = ""
        End If
    Next Control
End Sub


Public Function CampoSiguiente(ByRef Formulario As Form, Valor As Integer) As Control
Dim Fin As Boolean
Dim Control As Object

On Error GoTo ECampoSiguiente

    'Debug.Print "Llamada:  " & Valor
    'Vemos cual es el siguiente
    Do
        Valor = Valor + 1
        For Each Control In Formulario.Controls
            'Debug.Print "-> " & Control.Name & " - " & Control.TabIndex
            'Si es texto monta esta parte de sql
            If Control.TabIndex = Valor Then
                    Set CampoSiguiente = Control
                    Fin = True
                    Exit For
            End If
        Next Control
        If Not Fin Then
            Valor = -1
        End If
    Loop Until Fin
    Exit Function
ECampoSiguiente:
    Set CampoSiguiente = Nothing
    Err.Clear
End Function




Private Function ValorParaSQL(Valor, ByRef vTag As CTag) As String
Dim Dev As String
Dim D As Single
Dim i As Integer

Dim V
    Dev = ""
    If Valor <> "" Then
        Select Case vTag.TipoDato
        Case "N"
            V = Valor
            If InStr(1, Valor, ",") Then
                V = CSng(Valor)
                Valor = V
            End If
            Dev = TransformaComasPuntos(CStr(Valor))
            
        Case "F"
            'Dev = "'" & Format(Valor, FormatoFecha) & "'"      ' EN MYSQL
            Dev = "#" & Format(Valor, FormatoFecha) & "#"       ' EN ACCESS
        Case "T"
            Dev = CStr(Valor)
            NombreSQL Dev
            Dev = "'" & Dev & "'"
        Case Else
            Dev = "'" & Valor & "'"
        End Select
        
    Else
        'Si se permiten nulos, la "" ponemos un NULL
        If vTag.Vacio = "S" Then Dev = ValorNulo
    End If
    ValorParaSQL = Dev
End Function

Public Function InsertarDesdeForm(ByRef Formulario As Form) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Izda As String
    Dim Der As String
    Dim Cad As String
    
    On Error GoTo EInsertarF
    'Exit Function
    Set mTag = New CTag
    InsertarDesdeForm = False
    Der = ""
    Izda = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        If Izda <> "" Then Izda = Izda & ","
                        'Access
                        'Izda = Izda & "[" & mTag.Columna & "]"
                        Izda = Izda & "" & mTag.columna & ""
                    
                        'Parte VALUES
                        Cad = ValorParaSQL(Control.Text, mTag)
                        If Der <> "" Then Der = Der & ","
                        Der = Der & Cad
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Izda <> "" Then Izda = Izda & ","
                'Access
                'Izda = Izda & "[" & mTag.Columna & "]"
                Izda = Izda & "" & mTag.columna & ""
                If Control.Value = 1 Then
                    Cad = "1"
                    Else
                    Cad = "0"
                End If
                If Der <> "" Then Der = Der & ","
                If mTag.TipoDato = "N" Then Cad = Abs(CBool(Cad))
                Der = Der & Cad
            End If
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Izda <> "" Then Izda = Izda & ","
                    'Izda = Izda & "[" & mTag.Columna & "]"
                    Izda = Izda & "" & mTag.columna & ""
                    If Control.ListIndex = -1 Then
                        Cad = ValorNulo
                        Else
                        Cad = Control.ItemData(Control.ListIndex)
                    End If
                    If Der <> "" Then Der = Der & ","
                    Der = Der & Cad
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo
    'INSERT INTO Empleados (Nombre,Apellido, Cargo) VALUES ('Carlos', 'Sesma', 'Prácticas');
    
    Cad = "INSERT INTO " & mTag.Tabla
    Cad = Cad & " (" & Izda & ") VALUES (" & Der & ");"
    
   
    conn.Execute Cad, , adCmdText
    espera 0.5
    
    InsertarDesdeForm = True
Exit Function
EInsertarF:
    MuestraError Err.Number, "Inserta. " & conn.Errors(0).Description
End Function





Public Function PonerCamposForma(ByRef Formulario As Form, ByRef vData As Adodc) As Boolean
    Dim Control As Object
    Dim mTag As CTag
    Dim Cad As String
    Dim Valor As Variant
    Dim Campo As String  'Campo en la base de datos
    Dim i As Integer


    On Error GoTo EPonerCamposForma:

    Set mTag = New CTag
    PonerCamposForma = False

    For Each Control In Formulario.Controls
        'TEXTO
        'Debug.Print Control.Tag
        If TypeOf Control Is TextBox Then
            'Comprobamos que tenga tag
            mTag.Cargar Control
            If Control.Tag <> "" Then
                If mTag.Cargado Then
                    'Columna en la BD
                    If mTag.columna <> "" Then
                        Campo = mTag.columna
                        If mTag.Vacio = "S" Then
                            Valor = DBLet(vData.Recordset.Fields(Campo))
                        Else
                            Valor = vData.Recordset.Fields(Campo)
                        End If
                        If mTag.Formato <> "" And CStr(Valor) <> "" Then
                            If mTag.TipoDato = "N" Then
                                'Es numerico, entonces formatearemos y sustituiremos
                                ' La coma por el punto
                                Cad = Format(Valor, mTag.Formato)
                                'Antiguo
                                'Control.Text = TransformaComasPuntos(cad)
                                'nuevo
                                Control.Text = Cad
                            Else
                                Control.Text = Format(Valor, mTag.Formato)
                            End If
                        Else
                            Control.Text = Valor
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Columna en la BD
                    Campo = mTag.columna
                    If IsNull(vData.Recordset.Fields(Campo)) Then
                        Valor = False
                    Else
                        Valor = vData.Recordset.Fields(Campo)
                    End If
                    Else
                        Valor = False
                End If
                Control.Value = Abs(Valor)
            End If
            
         'COMBOBOX
         ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    Campo = mTag.columna
                    If Not IsNull(vData.Recordset.Fields(Campo)) Then
                        Valor = vData.Recordset.Fields(Campo)
                        i = 0
                        For i = 0 To Control.ListCount - 1
                            If Control.ItemData(i) = Val(Valor) Then
                                Control.ListIndex = i
                                Exit For
                            End If
                        Next i
                    Else
                        i = 32000
                    End If
                    If i >= Control.ListCount Then Control.ListIndex = -1
                End If 'de cargado
            End If 'de <>""
        End If
    Next Control
    
    'Veremos que tal
    PonerCamposForma = True
Exit Function
EPonerCamposForma:
    MuestraError Err.Number, "Poner campos formulario. "
End Function

Public Function ObtenerMaximoMinimo(ByRef vSQL As String) As String
Dim Rs As Recordset
ObtenerMaximoMinimo = ""
Set Rs = New ADODB.Recordset
Rs.Open vSQL, conn, adOpenForwardOnly, adLockOptimistic, adCmdText
If Not Rs.EOF Then
    If Not IsNull(Rs.EOF) Then
        ObtenerMaximoMinimo = CStr(Rs.Fields(0))
    End If
End If
Rs.Close
Set Rs = Nothing
End Function


Public Function ObtenerBusqueda(ByRef Formulario As Form) As String
    Dim Control As Object
    Dim Carga As Boolean
    Dim mTag As CTag
    Dim Aux As String
    Dim Cad As String
    Dim SQL As String
    Dim Tabla As String
    Dim RC As Byte

    On Error GoTo EObtenerBusqueda

    'Exit Function
    Set mTag = New CTag
    ObtenerBusqueda = ""
    SQL = ""

    'Recorremos los text en busca de ">>" o "<<"
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            Aux = Trim(Control.Text)
            If Aux = ">>" Or Aux = "<<" Then
                Carga = mTag.Cargar(Control)
                If Carga Then
                    If Aux = ">>" Then
                        Cad = " MAX(" & mTag.columna & ")"
                    Else
                        Cad = " MIN(" & mTag.columna & ")"
                    End If
                    SQL = "Select " & Cad & " from " & mTag.Tabla
                    SQL = ObtenerMaximoMinimo(SQL)
                    Select Case mTag.TipoDato
                    Case "N"
                        SQL = mTag.Tabla & "." & mTag.columna & " = " & TransformaComasPuntos(SQL)
                    Case "F"
                        SQL = mTag.Tabla & "." & mTag.columna & " = '" & Format(SQL, "yyyy-mm-dd") & "'"
                    Case Else
                        SQL = mTag.Tabla & "." & mTag.columna & " = '" & SQL & "'"
                    End Select
                    SQL = "(" & SQL & ")"
                End If
            End If
        End If
    Next

    

    'Recorremos los textbox
    For Each Control In Formulario.Controls
        If TypeOf Control Is TextBox Then
            'Cargamos el tag
            Carga = mTag.Cargar(Control)
            If Carga Then
                Aux = Trim(Control.Text)
                If Aux <> "" Then
                    If mTag.Tabla <> "" Then
                        Tabla = mTag.Tabla & "."
                        Else
                        Tabla = ""
                    End If
                    RC = SeparaCampoBusqueda(mTag.TipoDato, Tabla & mTag.columna, Aux, Cad)
                    If RC = 0 Then
                        If SQL <> "" Then SQL = SQL & " AND "
                        SQL = SQL & "(" & Cad & ")"
                    End If
                End If
            Else
                MsgBox "Carga de tag erronea en el control " & Control.Text & " -> " & Control.Tag
                Exit Function
            End If
        
        
        
        'COMBO BOX
        ElseIf TypeOf Control Is ComboBox Then
            mTag.Cargar Control
            If mTag.Cargado Then
                If Control.ListIndex > -1 Then
                    Cad = Control.ItemData(Control.ListIndex)
                    Cad = mTag.Tabla & "." & mTag.columna & " = " & Cad
                    If SQL <> "" Then SQL = SQL & " AND "
                    SQL = SQL & "(" & Cad & ")"
                End If
            End If
        
        End If
    Next Control
    ObtenerBusqueda = SQL
Exit Function
EObtenerBusqueda:
    ObtenerBusqueda = ""
    MuestraError Err.Number, "Obtener búsqueda. "
End Function




Public Function ModificaDesdeFormulario(ByRef Formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    If mTag.columna <> "" Then
                        'Sea para el where o para el update esto lo necesito
                        Aux = ValorParaSQL(Control.Text, mTag)
                        'Si es campo clave NO se puede modificar y se utiliza como busqueda
                        'dentro del WHERE
                        If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                             If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                             cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                             
                        Else
                            If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                            cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                        End If
                    End If
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                
                
                
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then

                
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                    
                    
                    If mTag.EsClave Then
                            'Lo pondremos para el WHERE
                        If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                        cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                    End If
                    
                End If
            End If
        End If
    Next Control
    'Construimos el SQL
    'Ejemplo:
    'Update Pedidos
    'SET ImportePedido = ImportePedido * 1.1,
    'Cargo = Cargo * 1.03
    'WHERE PaísDestinatario = 'México';
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText

    espera 0.5




ModificaDesdeFormulario = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function

Public Function ParaGrid(ByRef Control As Control, AnchoPorcentaje As Integer, Optional Desc As String) As String
Dim mTag As CTag
Dim Cad As String

'Montamos al final: "Cod Diag.|idDiag|N|10·"

ParaGrid = ""
Cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Desc <> "" Then
                Cad = Desc
            Else
                Cad = mTag.Nombre
            End If
            Cad = Cad & "|"
            Cad = Cad & mTag.columna & "|"
            Cad = Cad & mTag.TipoDato & "|"
            Cad = Cad & AnchoPorcentaje & "·"
            
                
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            
        ElseIf TypeOf Control Is ComboBox Then
        
        
        End If 'De los elseif
    End If
Set mTag = Nothing
ParaGrid = Cad
End If



End Function

'////////////////////////////////////////////////////
' Monta a partir de una cadena devuelta por el formulario
'de busqueda el sql para situar despues el datasource
Public Function ValorDevueltoFormGrid(ByRef Control As Control, ByRef CadenaDevuelta As String, Orden As Integer) As String
Dim mTag As CTag
Dim Cad As String
Dim Aux As String
'Montamos al final: " columnatabla = valordevuelto "

ValorDevueltoFormGrid = ""
Cad = ""
Set mTag = New CTag
mTag.Cargar Control
If mTag.Cargado Then
    If Control.Tag <> "" Then
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            Aux = RecuperaValor(CadenaDevuelta, Orden)
            If Aux <> "" Then Cad = mTag.columna & " = " & ValorParaSQL(Aux, mTag)
                
            
            
                
        'CheckBOX
       ' ElseIf TypeOf Control Is CheckBox Then
       '
       ' ElseIf TypeOf Control Is ComboBox Then
       '
       '
        End If 'De los elseif
    End If
End If
Set mTag = Nothing
ValorDevueltoFormGrid = Cad
End Function


Public Sub FormateaCampo(vTex As TextBox)
    Dim mTag As CTag
    Dim Cad As String
    On Error GoTo EFormateaCampo
    Set mTag = New CTag
    mTag.Cargar vTex
    If mTag.Cargado Then
        If vTex.Text <> "" Then
            If mTag.Formato <> "" Then
                Cad = TransformaPuntosComas(vTex.Text)
                Cad = Format(Cad, mTag.Formato)
                vTex.Text = Cad
            End If
        End If
    End If
EFormateaCampo:
    If Err.Number <> 0 Then Err.Clear
    Set mTag = Nothing
End Sub


'recupera valor desde una cadena con pipes(acabada en pipes)
'Para ello le decimos el orden  y ya ta
Public Function RecuperaValor(ByRef CADENA As String, Orden As Integer) As String
Dim i As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String

i = 0
Cont = 1
Cad = ""
Do
    J = i + 1
    i = InStr(J, CADENA, "|")
    If i > 0 Then
        If Cont = Orden Then
            Cad = Mid(CADENA, J, i - J)
            i = Len(CADENA) 'Para salir del bucle
            Else
                Cont = Cont + 1
        End If
    End If
Loop Until i = 0
RecuperaValor = Cad
End Function




'-----------------------------------------------------------------------
'Deshabilitar ciertas opciones del menu
'EN funcion del nivel de usuario
'Esto es a nivel general, cuando el Toolba es el mismo



'Este modifica las claves prinipales y todo
'la sentenca del WHERE cod=1 and .. viene en claves
Public Function ModificaDesdeFormularioClaves(ByRef Formulario As Form, Claves As String) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim cadUPDATE As String
Dim i As Integer

On Error GoTo EModificaDesdeFormulario
    ModificaDesdeFormularioClaves = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        'CheckBOX
        ElseIf TypeOf Control Is CheckBox Then
            'Partimos de la base que un booleano no es nunca clave primaria
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If Control.Value = 1 Then
                    Aux = "TRUE"
                    Else
                    Aux = "FALSE"
                End If
                If mTag.TipoDato = "N" Then Aux = Abs(CBool(Aux))
                If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                'Esta es para access
                'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
            End If
            
        ElseIf TypeOf Control Is ComboBox Then
            If Control.Tag <> "" Then
                mTag.Cargar Control
                If mTag.Cargado Then
                    If Control.ListIndex = -1 Then
                        Aux = ValorNulo
                        Else
                        Aux = Control.ItemData(Control.ListIndex)
                    End If
                    If cadUPDATE <> "" Then cadUPDATE = cadUPDATE & " , "
                    'cadUPDATE = cadUPDATE & "[" & mTag.Columna & "] = " & aux
                    cadUPDATE = cadUPDATE & "" & mTag.columna & " = " & Aux
                End If
            End If
        End If
    Next Control
    cadWHERE = Claves
    'Construimos el SQL
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        Exit Function
    End If
    Aux = "UPDATE " & mTag.Tabla
    Aux = Aux & " SET " & cadUPDATE & " WHERE " & cadWHERE
    conn.Execute Aux, , adCmdText






ModificaDesdeFormularioClaves = True
Exit Function
EModificaDesdeFormulario:
    MuestraError Err.Number, "Modificar. " & Err.Description
End Function







Public Function BLOQUEADesdeFormulario(ByRef Formulario As Form) As Boolean
Dim Control As Object
Dim mTag As CTag
Dim Aux As String
Dim cadWHERE As String
Dim AntiguoCursor As Byte

On Error GoTo EBLOQUEADesdeFormulario
    BLOQUEADesdeFormulario = False
    Set mTag = New CTag
    Aux = ""
    cadWHERE = ""
    AntiguoCursor = Screen.MousePointer
    Screen.MousePointer = vbHourglass
    For Each Control In Formulario.Controls
        'Si es texto monta esta parte de sql
        If TypeOf Control Is TextBox Then
            If Control.Tag <> "" Then
              
                mTag.Cargar Control
                If mTag.Cargado Then
                    'Sea para el where o para el update esto lo necesito
                    Aux = ValorParaSQL(Control.Text, mTag)
                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
                    'dentro del WHERE
                    If mTag.EsClave Then
                        'Lo pondremos para el WHERE
                         If cadWHERE <> "" Then cadWHERE = cadWHERE & " AND "
                         cadWHERE = cadWHERE & "(" & mTag.columna & " = " & Aux & ")"
                    End If
                End If
            End If
        End If
    Next Control
    
    If cadWHERE = "" Then
        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
        
    Else
        Aux = "select * FROM " & mTag.Tabla
        Aux = Aux & " WHERE " & cadWHERE & " FOR UPDATE"
        
        'Intenteamos bloquear
        'PreparaBloquear
        conn.Execute Aux, , adCmdText
        BLOQUEADesdeFormulario = True
    End If
EBLOQUEADesdeFormulario:
    If Err.Number <> 0 Then
        MuestraError Err.Number, "Bloqueo tabla"
        'TerminaBloquear
    End If
    Screen.MousePointer = AntiguoCursor
End Function




'Public Function BloqueaRegistroForm(ByRef Formulario As Form) As Boolean
'Dim Control As Object
'Dim mTag As CTag
'Dim Aux As String
'Dim AuxDef As String
'Dim AntiguoCursor As Byte
'
'On Error GoTo EBLOQ
'    BloqueaRegistroForm = False
'    Set mTag = New CTag
'    Aux = ""
'    AuxDef = ""
'    AntiguoCursor = Screen.MousePointer
'    Screen.MousePointer = vbHourglass
'    For Each Control In Formulario.Controls
'        'Si es texto monta esta parte de sql
'        If TypeOf Control Is TextBox Then
'            If Control.Tag <> "" Then
'
'                mTag.Cargar Control
'                If mTag.Cargado Then
'                    'Si es campo clave NO se puede modificar y se utiliza como busqueda
'                    'dentro del WHERE
'                    If mTag.EsClave Then
'                        Aux = ValorParaSQL(Control.Text, mTag)
'                        AuxDef = AuxDef & Aux & "|"
'                    End If
'                End If
'            End If
'        End If
'    Next Control
'
'    If AuxDef = "" Then
'        MsgBox "No se ha definido ninguna clave principal.", vbExclamation
'
'    Else
'        Aux = "Insert into zBloqueos(codusu,tabla,clave) VALUES(" & vUsu.Codigo & ",'" & mTag.Tabla
'        Aux = Aux & "',""" & AuxDef & """)"
'        Conn.Execute Aux
'        BloqueaRegistroForm = True
'    End If
'EBLOQ:
'    If Err.Number <> 0 Then
'        Aux = ""
'        If Conn.Errors.Count > 0 Then
'            If Conn.Errors(0).NativeError = 1062 Then
'                '¡Ya existe el registro, luego esta bloqueada
'                Aux = "BLOQUEO"
'            End If
'        End If
'        If Aux = "" Then
'            MuestraError Err.Number, "Bloqueo tabla"
'        Else
'            MsgBox "Registro bloqueado por otro usuario", vbExclamation
'        End If
'    End If
'    Screen.MousePointer = AntiguoCursor
'End Function
'
'
'Public Function DesBloqueaRegistroForm(ByRef TextBoxConTag As TextBox) As Boolean
'Dim mTag As CTag
'Dim SQL As String
'
''Solo me interesa la tabla
'On Error Resume Next
'    Set mTag = New CTag
'    mTag.Cargar TextBoxConTag
'    If mTag.Cargado Then
'        SQL = "DELETE from zBloqueos where codusu=" & vUsu.Codigo & " and tabla='" & mTag.Tabla & "'"
'        Conn.Execute SQL
'        If Err.Number <> 0 Then
'            Err.Clear
'        End If
'    End If
'    Set mTag = Nothing
'End Function
'
'
'
'
'











Public Function SeparaCampoBusqueda(Tipo As String, Campo As String, CADENA As String, ByRef DevSQL As String) As Byte
Dim Cad As String
Dim Aux As String
Dim cH As String
Dim Fin As Boolean
Dim i, J As String

On Error GoTo ErrSepara
SeparaCampoBusqueda = 1
DevSQL = ""
Cad = ""
Select Case Tipo
Case "N"
    '----------------  NUMERICO  ---------------------
    i = CararacteresCorrectos(CADENA, "N")
    If i > 0 Then Exit Function  'Ha habido un error y salimos
    'Comprobamos si hay intervalo ':'
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        'Intervalo numerico
        Cad = Mid(CADENA, 1, i - 1)
        Aux = Mid(CADENA, i + 1)
        If Not IsNumeric(Cad) Or Not IsNumeric(Aux) Then Exit Function  'No son numeros
        'Intervalo correcto
        'Construimos la cadena
        DevSQL = Campo & " >= " & Cad & " AND " & Campo & " <= " & Aux
        '----
        'ELSE
        Else
            'Prueba
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                DevSQL = "1=1"
             Else
                    Fin = False
                    i = 1
                    Cad = ""
                    Aux = "NO ES NUMERO"
                    While Not Fin
                        cH = Mid(CADENA, i, 1)
                        If cH = ">" Or cH = "<" Or cH = "=" Then
                            Cad = Cad & cH
                            Else
                                Aux = Mid(CADENA, i)
                                Fin = True
                        End If
                        i = i + 1
                        If i > Len(CADENA) Then Fin = True
                    Wend
                    'En aux debemos tener el numero
                    If Not IsNumeric(Aux) Then Exit Function
                    'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                    If Cad = "" Then Cad = " = "
                    DevSQL = Campo & " " & Cad & " " & Aux
            End If
        End If
Case "F"
     '---------------- FECHAS ------------------
    i = CararacteresCorrectos(CADENA, "F")
    If i = 1 Then Exit Function
    'Comprobamos si hay intervalo ':'
    i = InStr(1, CADENA, ":")
    If i > 0 Then
        'Intervalo de fechas
        Cad = Mid(CADENA, 1, i - 1)
        Aux = Mid(CADENA, i + 1)
        If Not EsFechaOKString(Cad) Or Not EsFechaOKString(Aux) Then Exit Function  'Fechas incorrectas
        'Intervalo correcto
        'Construimos la cadena
        Cad = Format(Cad, FormatoFecha)
        Aux = Format(Aux, FormatoFecha)
        'En my sql es la ' no el #
        'DevSQL = Campo & " >=#" & Cad & "# AND " & Campo & " <= #" & AUX & "#"
        DevSQL = Campo & " >='" & Cad & "' AND " & Campo & " <= '" & Aux & "'"
        '----
        'ELSE
        Else
            'Comprobamos que no es el mayor
            If CADENA = ">>" Or CADENA = "<<" Then
                  DevSQL = "1=1"
            Else
                Fin = False
                i = 1
                Cad = ""
                Aux = "NO ES FECHA"
                While Not Fin
                    cH = Mid(CADENA, i, 1)
                    If cH = ">" Or cH = "<" Or cH = "=" Then
                        Cad = Cad & cH
                        Else
                            Aux = Mid(CADENA, i)
                            Fin = True
                    End If
                    i = i + 1
                    If i > Len(CADENA) Then Fin = True
                Wend
                'En aux debemos tener el numero
                If Not EsFechaOKString(Aux) Then Exit Function
                'Si que es numero. Entonces, si Cad="" entronces le ponemos =
                Aux = "'" & Format(Aux, FormatoFecha) & "'"
                If Cad = "" Then Cad = " = "
                DevSQL = Campo & " " & Cad & " " & Aux
            End If
        End If
    
    
    
    
Case "T"
    '---------------- TEXTO ------------------
    i = CararacteresCorrectos(CADENA, "T")
    If i = 1 Then Exit Function
    
    'Comprobamos que no es el mayor
     If CADENA = ">>" Or CADENA = "<<" Then
        DevSQL = "1=1"
        Exit Function
    End If
    'Cambiamos el * por % puesto que en ADO es el caraacter para like
    i = 1
    Aux = CADENA
    While i <> 0
        i = InStr(1, Aux, "*")
        If i > 0 Then Aux = Mid(Aux, 1, i - 1) & "%" & Mid(Aux, i + 1)
    Wend
    'Cambiamos el ? por la _ pue es su omonimo
    i = 1
    While i <> 0
        i = InStr(1, Aux, "?")
        If i > 0 Then Aux = Mid(Aux, 1, i - 1) & "_" & Mid(Aux, i + 1)
    Wend
    Cad = Mid(CADENA, 1, 2)
    If Cad = "<>" Then
        Aux = Mid(CADENA, 3)
        DevSQL = Campo & " LIKE '!" & Aux & "'"
        Else
        DevSQL = Campo & " LIKE '" & Aux & "'"
    End If
    


    
Case "B"
    'Como vienen de check box o del option box
    'los escribimos nosotros luego siempre sera correcta la
    'sintaxis
    'Los booleanos. Valores buenos son
    'Verdadero , Falso, True, False, = , <>
    'Igual o distinto
    i = InStr(1, CADENA, "<>")
    If i = 0 Then
        'IGUAL A valor
        Cad = " = "
        Else
            'Distinto a valor
        Cad = " <> "
    End If
    'Verdadero o falso
    i = InStr(1, CADENA, "V")
    If i > 0 Then
            Aux = "True"
            Else
            Aux = "False"
    End If
    'Ponemos la cadena
    DevSQL = Campo & " " & Cad & " " & Aux
    
Case Else
    'No hacemos nada
        Exit Function
End Select
SeparaCampoBusqueda = 0
ErrSepara:
    If Err.Number <> 0 Then MuestraError Err.Number
End Function


Private Function CararacteresCorrectos(vCad As String, Tipo As String) As Byte
Dim i As Integer
Dim cH As String
Dim Error As Boolean

CararacteresCorrectos = 1
Error = False
Select Case Tipo
Case "N"
    'Numero. Aceptamos numeros, >,< = :
    For i = 1 To Len(vCad)
        cH = Mid(vCad, i, 1)
        Select Case cH
            Case "0" To "9"
            Case "<", ">", ":", "=", ".", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next i
Case "T"
    'Texto aceptamos numeros, letras y el interrogante y el asterisco
    For i = 1 To Len(vCad)
        cH = Mid(vCad, i, 1)
        Select Case cH
            Case "a" To "z"
            Case "A" To "Z"
            Case "0" To "9"
            Case "*", "%", "?", "_", "\", ":", " " ' estos son para un caracter sol no esta demostrado , "%", "&"
            'Esta es opcional
            Case "<", ">"
            Case "Ñ", "ñ"
            Case Else
                Error = True
                Exit For
        End Select
    Next i
Case "F"
    'Numeros , "/" ,":"
    For i = 1 To Len(vCad)
        cH = Mid(vCad, i, 1)
        Select Case cH
            Case "0" To "9"
            Case "<", ">", ":", "/", "="
            Case Else
                Error = True
                Exit For
        End Select
    Next i
Case "B"
    'Numeros , "/" ,":"
    For i = 1 To Len(vCad)
        cH = Mid(vCad, i, 1)
        Select Case cH
            Case "0" To "9"
            Case "<", ">", ":", "/", "=", " "
            Case Else
                Error = True
                Exit For
        End Select
    Next i
End Select
'Si no ha habido error cambiamos el retorno
If Not Error Then CararacteresCorrectos = 0
End Function







Public Sub CargaComboSecciones(ByRef CBO As ComboBox, AñadirTodas As Boolean)
Dim SQL As String
Dim Rs As ADODB.Recordset

    CBO.Clear
    SQL = "select IdSeccion,nombre from secciones order by NOMBRE"
    Set Rs = New ADODB.Recordset
    Rs.Open SQL, conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    If AñadirTodas Then
        CBO.AddItem "Todas las secciones"
        CBO.ItemData(CBO.NewIndex) = -1
    End If
    
    While Not Rs.EOF
        CBO.AddItem Rs!Nombre & " (" & Rs!idSeccion & ")"
        CBO.ItemData(CBO.NewIndex) = Rs!idSeccion
        Rs.MoveNext
    Wend
    Rs.Close
    Set Rs = Nothing
    If AñadirTodas Then CBO.ListIndex = 0
    
End Sub

