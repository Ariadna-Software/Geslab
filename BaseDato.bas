Attribute VB_Name = "BaseDato"
Option Explicit

Public dbBinary As Long
Public dbBoolean As Long
Public dbByte As Long
Public dbChar As Long
Public dbCurrency As Long
Public dbDate As Long
Public dbDecimal As Long
Public dbDouble As Long
Public dbFloat As Long
Public dbInteger As Long
Public dbLong As Long
Public dbMemo As Long
Public dbNumeric As Long
Public dbSingle As Long
Public dbText As Long
Public dbTime As Long



Public Sub PonerValoresConstantes_BD()

'Estos valores son los de DAO
dbBinary = 9
dbBoolean = 1
dbByte = 2
'dbchar= 18
dbChar = 202
dbCurrency = 5
dbDate = 8
dbDecimal = 20
dbDouble = 7
dbFloat = 21
dbInteger = 3
dbLong = 4
dbMemo = 12
dbNumeric = 19
dbSingle = 6
dbText = 10
dbTime = 22
End Sub



'Para los tipos de datos de los campos de las BD
Public Function DevuelveTipo(Valor As Long) As String

DevuelveTipo = "NO"
If Valor = dbBinary Then
        'Binary
         DevuelveTipo = "Byte"
ElseIf Valor = dbBoolean Then
        '  Boolean
         DevuelveTipo = "Boolean"
ElseIf Valor = dbByte Then
        'Byte
         DevuelveTipo = "Byte"
ElseIf Valor = dbChar Then
        'Char
         DevuelveTipo = "String"
ElseIf Valor = dbCurrency Then
        'Currency
         DevuelveTipo = "Currency"
ElseIf Valor = dbDate Then
        'Date / Time
         DevuelveTipo = "Date"
ElseIf Valor = dbDecimal Then
        'Decimal
         DevuelveTipo = "Decimal"
ElseIf Valor = dbDouble Then
        '  Double
         DevuelveTipo = "Double"
ElseIf Valor = dbFloat Then
        'Float
         DevuelveTipo = "Float"
ElseIf Valor = dbInteger Then
        'Integer
         DevuelveTipo = "Integer"
ElseIf Valor = dbLong Then
        'Long
         DevuelveTipo = "Long"
ElseIf Valor = dbMemo Then
        'Memo
         DevuelveTipo = "String"
ElseIf Valor = dbNumeric Then
        'Numeric
         DevuelveTipo = "NO"
ElseIf Valor = dbSingle Then
        'Single
         DevuelveTipo = "Single"
ElseIf Valor = dbText Then
        'Text
         DevuelveTipo = "String"
'Case dbTime
'        'Time
'         DevuelveTipo = "NO"
'Case dbTimeStamp
'        'TimeStamp
'         DevuelveTipo = "NO"
'Case dbVarBinary
'        'VarBinary
'         DevuelveTipo = "NO"
End If
End Function

Public Function DevuelveTipo2(Valor As Long) As Byte


'   1.- Numeros
'   2.- Booleanos
'   3.- Cadenas
'   4.- Fecha
'   0.- Error leyendo los tipos de datos
DevuelveTipo2 = 0
If Valor = dbBinary Then
        'Binary
         DevuelveTipo2 = 1
ElseIf Valor = dbBoolean Then
        '  Boolean
         DevuelveTipo2 = 2
ElseIf Valor = dbByte Then
'        'Byte
         DevuelveTipo2 = 1
ElseIf Valor = dbChar Then
'        'Char
        DevuelveTipo2 = 3
ElseIf Valor = dbCurrency Then
'        'Currency
        DevuelveTipo2 = 1
ElseIf Valor = dbDate Then
'        'Date / Time
        DevuelveTipo2 = 4
ElseIf Valor = dbDecimal Then
        DevuelveTipo2 = 1
ElseIf Valor = dbDouble Then
'        '  Double
        DevuelveTipo2 = 1
ElseIf Valor = dbFloat Then
'        'Float
        DevuelveTipo2 = 1
ElseIf Valor = dbInteger Then
'        'Integer
        DevuelveTipo2 = 1
ElseIf Valor = dbLong Then
'        'Long
        DevuelveTipo2 = 1
ElseIf Valor = dbMemo Then
'        'Memo
        DevuelveTipo2 = 3
ElseIf Valor = dbNumeric Then
'        'Numeric
        DevuelveTipo2 = 1
ElseIf Valor = dbSingle Then
'        'Single
        DevuelveTipo2 = 1
ElseIf Valor = dbText Then
'        'Text
        DevuelveTipo2 = 3
ElseIf Valor = dbTime Then
'        'Time
        DevuelveTipo2 = 4
End If
End Function
