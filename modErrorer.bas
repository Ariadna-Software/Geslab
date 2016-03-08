Attribute VB_Name = "modErrorer"
Option Explicit

Private Fich_ErrorLineas As Integer
Private B_FErroresLin As Boolean
Private Fecha_FicheErrorLinea As Date



'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'                 ERRORES PROCESANDO LAS LINEAS
'
'
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------
'-----------------------------------------------------------------------------


Public Sub InicializaErroresLinea(vF As Date)
B_FErroresLin = False
Fecha_FicheErrorLinea = vF
End Sub

Public Sub FinErroresLinea()
Dim cad As String

If B_FErroresLin Then
    Close #Fich_ErrorLineas
    'Optativo
    cad = Format(Fecha_FicheErrorLinea, "ddmmyyyy")
    cad = App.Path & "\ErrorLineas" & cad & ".log"
    MsgBox "Se han producido errores en la importacion de archivos." & vbCrLf & _
        "Vease el archivo " & cad & " para más información.", vbExclamation
End If
End Sub

Private Sub AbrirFicheroErrores()
Dim cad As String

On Error GoTo Errores1
cad = Format(Fecha_FicheErrorLinea, "ddmmyy")
cad = App.Path & "\ErrorLineas" & cad & ".log"
Fich_ErrorLineas = FreeFile
Open cad For Output As #Fich_ErrorLineas
B_FErroresLin = True
Exit Sub
Errores1:
    MsgBox "Error: " & vbCrLf & Err.Number & " - " & Err.Description
End Sub


Public Sub EscribeErrorLinea(Lin As String)
If B_FErroresLin Then
    Print #Fich_ErrorLineas, Lin
    Else
        AbrirFicheroErrores
        If B_FErroresLin Then Print #Fich_ErrorLineas, Lin
End If
End Sub
