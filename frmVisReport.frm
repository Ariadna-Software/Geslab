VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmVisReport 
   Caption         =   "Visor de informes"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5925
   WindowState     =   2  'Maximized
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   3855
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      lastProp        =   600
      _cx             =   10186
      _cy             =   6800
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
   End
End
Attribute VB_Name = "frmVisReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Informe As String

'estas varriables las trae del formulario de impresion
Public FormulaSeleccion As String
Public SoloImprimir As Boolean
Public OtrosParametros As String   ' El grupo acaba en |                            ' param1=valor1|param2=valor2|
Public NumeroParametros As Integer   'Cuantos parametros hay.  EMPRESA(EMP) no es parametro. Es fijo en todos los informes
Public MostrarTree As Boolean

Public ExportarPDF As Boolean


Dim mapp As CRAXDRT.Application
Dim mrpt As CRAXDRT.Report
Dim Argumentos() As String
Dim PrimeraVez As Boolean



Private Sub Form_Activate()
    If PrimeraVez Then
        PrimeraVez = False
        If SoloImprimir Or Me.ExportarPDF Then
            Screen.MousePointer = vbHourglass
            Unload Me
        End If
    End If
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()

On Error GoTo Err_Carga
    
    Dim i As Integer
    Screen.MousePointer = vbHourglass
    Set mapp = CreateObject("CrystalRuntime.Application")
    'Informe = "C:\Programas\Conta\Contabilidad\InformesD\sumas12.rpt"
    Set mrpt = mapp.OpenReport(Informe)

    For i = 1 To mrpt.Database.Tables.Count
       'mrpt.Database.Tables(I).SetLogOnInfo "vUsuarios", "Usuarios", vConfig.User, vConfig.Password
       mrpt.Database.Tables(i).SetLogOnInfo "aripres"
    Next i

    'LA variable otros parametros debe empezar con la barra |
    If OtrosParametros <> "" Then
        If Mid(OtrosParametros, 1, 1) <> "|" Then OtrosParametros = "|" & OtrosParametros
    End If

    PrimeraVez = True
    CargaArgumentos
    CRViewer1.EnableGroupTree = MostrarTree
    CRViewer1.DisplayGroupTree = MostrarTree
    
    
    If FormulaSeleccion <> "" Then
        mrpt.RecordSelectionFormula = FormulaSeleccion
    End If
    'Si es a mail
    If Me.ExportarPDF Then
       ' Exportar
        Exit Sub
    End If
    
    
    'lOS MARGENES
    PonerMargen
    
    CRViewer1.ReportSource = mrpt
    If SoloImprimir Then
        mrpt.PrintOut False
    Else
        CRViewer1.ViewReport
    End If
    Exit Sub
Err_Carga:
    MsgBox "Error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Informe, vbCritical
    Set mapp = Nothing
    Set mrpt = Nothing
End Sub

Private Sub Form_Resize()
    CRViewer1.Top = 0
    CRViewer1.Left = 0
    CRViewer1.Height = ScaleHeight
    CRViewer1.Width = ScaleWidth
End Sub

Private Sub CargaArgumentos()
Dim Parametro As String
Dim i As Integer
    'El primer parametro es el nombre de la empresa para todas las empresas
    ' Por lo tanto concaatenaremos con otros parametros
    ' Y sumaremos uno
    'Luego iremos recogiendo para cada formula su valor y viendo si esta en
    ' La cadena de parametros
    'Si esta asignaremos su valor
    
    'OtrosParametros = "|Emp= """ & vEmpresa.nomempre & """|" & OtrosParametros
    'NumeroParametros = NumeroParametros + 1
    
    
    
    For i = 1 To mrpt.FormulaFields.Count
        Parametro = mrpt.FormulaFields(i).Name
        Parametro = Mid(Parametro, 3)  'Quitamos el {@
        Parametro = Mid(Parametro, 1, Len(Parametro) - 1) ' el } del final
        'Debug.Print Parametro
        If DevuelveValor(Parametro) Then mrpt.FormulaFields(i).Text = Parametro
        'Debug.Print " -- " & Parametro
    Next i
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mrpt = Nothing
    Set mapp = Nothing
End Sub


Private Function DevuelveValor(ByRef Valor As String) As Boolean
Dim i As Integer
Dim J As Integer
    Valor = "|" & Valor & "="
    DevuelveValor = False
    i = InStr(1, OtrosParametros, Valor, vbTextCompare)
    If i > 0 Then
        i = i + Len(Valor) + 1
        J = InStr(i, OtrosParametros, "|")
        If J > 0 Then
            Valor = Mid(OtrosParametros, i, J - i)
            If Valor = "" Then
                Valor = " "
            Else
                CompruebaComillas Valor
            End If
            DevuelveValor = True
        End If
    End If
End Function


Private Sub CompruebaComillas(ByRef Valor1 As String)
Dim Aux As String
Dim J As Integer
Dim i As Integer

    If Mid(Valor1, 1, 1) = Chr(34) Then
        'Tiene comillas. Con lo cual tengo k poner las dobles
        Aux = Mid(Valor1, 2, Len(Valor1) - 2)
        i = -1
        Do
            J = i + 2
            i = InStr(J, Aux, """")
            If i > 0 Then
              Aux = Mid(Aux, 1, i - 1) & """" & Mid(Aux, i)
            End If
        Loop Until i = 0
        Aux = """" & Aux & """"
        Valor1 = Aux
    End If
End Sub

'Private Sub Exportar()
'    mrpt.ExportOptions.DiskFileName = App.Path & "\docum.pdf"
'    mrpt.ExportOptions.DestinationType = crEDTDiskFile
'    mrpt.ExportOptions.PDFExportAllPages = True
'    mrpt.ExportOptions.FormatType = crEFTPortableDocFormat
'    mrpt.Export False
'    'Si ha generado bien entonces
'    CadenaDesdeOtroForm = "OK"
'End Sub

Private Sub PonerMargen()
Dim Cad As String
Dim i As Integer
    On Error GoTo EPon
    Cad = Dir(App.Path & "\*.mrg")
    If Cad <> "" Then
        i = InStr(1, Cad, ".")
        If i > 0 Then
            Cad = Mid(Cad, 1, i - 1)
            If IsNumeric(Cad) Then
                If Val(Cad) > 4000 Then Cad = "4000"
                If Val(Cad) > 0 Then
                    mrpt.BottomMargin = mrpt.BottomMargin + Val(Cad)
                End If
            End If
        End If
    End If
    
    Exit Sub
EPon:
    Err.Clear
End Sub
