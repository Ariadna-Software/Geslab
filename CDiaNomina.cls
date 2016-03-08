VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDiaNomina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Private mvarDiaSemana As Byte 'copia local
Private mvarFestivo As Boolean 'copia local
Private mvarHoraRealT As Currency 'copia local
Private mvarHoraRealC As Currency 'copia local
Public Property Let HoraRealC(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HoraRealC = 5
    mvarHoraRealC = vData
End Property


Public Property Get HoraRealC() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HoraRealC
    HoraRealC = mvarHoraRealC
End Property



Public Property Let HoraRealT(ByVal vData As Currency)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.HoraRealT = 5
    mvarHoraRealT = vData
End Property


Public Property Get HoraRealT() As Currency
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.HoraRealT
    HoraRealT = mvarHoraRealT
End Property



Public Property Let Festivo(ByVal vData As Boolean)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Festivo = 5
    mvarFestivo = vData
End Property


Public Property Get Festivo() As Boolean
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Festivo
    Festivo = mvarFestivo
End Property



Public Property Let DiaSemana(ByVal vData As Byte)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DiaSemana = 5
    mvarDiaSemana = vData
End Property


Public Property Get DiaSemana() As Byte
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DiaSemana
    DiaSemana = mvarDiaSemana
End Property



