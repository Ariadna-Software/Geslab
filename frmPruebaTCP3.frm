VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form frmPruebaTCP3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Comprobar TCP"
   ClientHeight    =   6900
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8625
   Icon            =   "frmPruebaTCP3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6900
   ScaleWidth      =   8625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSecuencia 
      Height          =   375
      Left            =   7800
      Picture         =   "frmPruebaTCP3.frx":1082
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox Text4 
      Height          =   1575
      Left            =   5280
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "frmPruebaTCP3.frx":1A84
      Top             =   600
      Width           =   3015
   End
   Begin VB.CommandButton cmdEnviar 
      Height          =   375
      Left            =   3600
      Picture         =   "frmPruebaTCP3.frx":1A94
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1800
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   2415
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2880
      TabIndex        =   4
      Text            =   "4"
      Top             =   1800
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   4335
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2400
      Width           =   8055
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   7320
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdReloj 
      Caption         =   "Prueba reloj"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label3 
      Caption         =   "Secuencia (comando |espera|)"
      Height          =   255
      Left            =   5280
      TabIndex        =   11
      Top             =   240
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   5040
      X2              =   5040
      Y1              =   240
      Y2              =   2160
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4920
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label2 
      Caption         =   "Texto a enviar"
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "Espera"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   1560
      Width           =   615
   End
   Begin VB.Label Label6 
      Height          =   255
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Label Label5 
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Width           =   4095
   End
End
Attribute VB_Name = "frmPruebaTCP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit




Private vIndice As Integer 'Indice para el formuario de seleccionar
Dim T1, T2
Dim Buffer$
Private NF As Integer 'Fichero
Private NombreFichero As String
Private kTCP As String  'Por si tiene + de un terminal TCP

'Datos configuracion
Private PuertoComm As Byte
Private Baudios As Long
Private NumTCP3 As Byte
Private EsperaBorrado As Integer 'En segundos


Private Sub cmdEnviar_Click()
Dim Cad As String
    If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
    
    
    Text1.Text = Text1.Text & Text3.Text & vbCrLf
    Text1.Refresh
    PonerTexto Text3.Text
    
    Cad = Leer(5, "espero:")
    Text3.SetFocus
End Sub

Private Sub CmdReloj_Click()
Dim Cad As String
Dim HayErrores As Boolean

   Screen.MousePointer = vbHourglass
    'Comprobaremos la fecha y hora del reloj
    'Si esta configurado y demas
    If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
   
    
    Text1.Text = ""
    Label5.Caption = "  LEYENDO "
    'Fecha PC
    Cad = UCase(Format(Now, "ddd d, mmm  hh:mm"))
    Label6.Caption = Cad

    Me.Refresh
    'Llegados aqui es donde empezamos a transmitir con el reloj
    'Abrimos el puerto
    
    Text1.Text = ""
    HayErrores = True
    
    'Solictamos comando
    Text1.Text = Text1.Text & "Solicitando programación hora/fecha reloj al TCP-3" & vbCrLf
    Text1.Refresh
    PonerTexto kTCP
    
    Cad = Leer(5, "Cmd:")
    LimpiaBufferRecepcion
    'Si respuesta afirmativa
    If Cad = "" Then
        GoTo Salida3
    End If
        
    'Ponemos comando 5 Leer hora/fecha en TCP-3
    Cad = "5"
    PonerTexto Cad
    Cad = Leer(2, "Cmd OK")
    LimpiaBufferRecepcion
    'No hay datos correctos
    If Cad = "" Then
        GoTo Salida3
    End If
    

    
    HayErrores = False
Salida3:
   
    
    If HayErrores Then
    
        Text1.Text = Text1.Text & vbCrLf & vbCrLf & _
            "Se han producido errores."
         Label5.Caption = "  E R R O R E S "
        End If
    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    
    
    'En cad tenemos lo k ha llegador
    If Not HayErrores Then
        Label5.Caption = "Ajustando valores dev."
        Label5.Refresh
        If Not PonerHoras(Cad) Then
            Label5.Caption = "Valores devueltos erroneos"
            Label5.Refresh
        End If
    End If
    
    Screen.MousePointer = vbDefault
End Sub


'Segunda prueba para leer mucho datos desde el TCP 3
Private Function LeerDatos2(Segundos As Integer, Leyendo As Boolean) As Boolean
Dim Seg As Integer
Dim I As Integer
Dim Fin As Boolean
Dim tmp As String
Dim Esta As Boolean
Dim J As Integer
Dim T3, T4
'De esta forma haremos:
'iremos leyendo del buffer
' Si hay texto cad segundos/2 (aproximadamente)
'lo guardamos y
'restauramos tiempo
'Si no hay tiempo salimos sin hacer nada
Buffer$ = ""
Fin = False
T1 = Timer
J = Segundos + 1
'Bucle
Text1.Text = Text1.Text & vbCrLf
If Leyendo Then
    Text1.Text = Text1.Text & "Leyendo registros de marcajes. " & vbCrLf
    Else
    Text1.Text = Text1.Text & "Eliminando registros de marcajes. " & vbCrLf
End If
Text1.Refresh
tmp = Text1.Text
I = 0
T3 = Timer
While Not Fin
    
    
    Do
        Buffer$ = Buffer$ & MSComm1.Input
        T2 = Timer - T1
        If Buffer$ <> "" Then J = 2
    Loop Until T2 > Segundos Or T2 > J
    T4 = Timer - T3
    If Buffer$ <> "" Then
        Text1.Text = tmp & vbCrLf & "Bloque: " & Format(I, "0000") & "     " & "  Seg: " & Format(T4, "0.00")
        Text1.Refresh
        I = I + 1
        EscribeTextoFichero Buffer$
        If InStr(1, Buffer$, "Cmd OK") Then
            Fin = True
            Esta = True
            Else
                Buffer$ = ""
                J = Segundos + 1
                T1 = Timer
        End If
        Else
            'i=0
            'fin tiempo
            Esta = False
            Fin = True
    End If
Wend
LeerDatos2 = Esta
End Function

Private Sub LimpiaBufferRecepcion()
Dim Cade
Cade = MSComm1.Input
If Cade <> "" Then Cade = String(40, "*") & vbCrLf & " L I M P I A R       B U F F E R" & vbCrLf & Cade & vbCrLf & String(40, "*")
Text1.Text = Text1.Text & Cade & vbCrLf
espera 0.25
End Sub

Private Sub PonerTexto(T As String)
'MSComm1.Output = T & vbCrLf
MSComm1.Output = T & Chr(13)

End Sub

Private Sub espera(tiempo As Single)
T1 = Timer
Do
    T2 = Timer - T1
Loop Until T2 > tiempo
End Sub



'Pone la cadena devuelta por el reloj  y la fecha /hora del PC
Private Function PonerHoras(CADENA As String) As Boolean
Dim I As Integer
Dim Fecha As Date

    PonerHoras = False
    Buffer$ = ""
    '5 Reg. Fin:
    'mes,dia,ds,hora,min,seg:08,31,5,21,38,21
    ' a partir de los dos puntos
    I = InStr(1, CADENA, ":")
    If I > 0 Then
        
        'BIEN
        CADENA = Mid(CADENA, I + 1)
        'Le aañdimos una coma al final para facilitar
        
        
    
            NF = InStr(1, CADENA, Chr(13))
            If NF > 0 Then CADENA = Mid(CADENA, 1, NF - 2)
            CADENA = CADENA & ","
        
        
        NF = 0
        Do
            I = InStr(1, CADENA, ",")
            If I > 0 Then
                NF = NF + 1
                Buffer$ = Mid(CADENA, I + 1)
                CADENA = Mid(CADENA, 1, I - 1) & "|" & Buffer$
            End If
        Loop Until I = 0
        
        Buffer$ = ""
        If NF = 6 Then

            'FECHA OK
            '----
            'Dia semana lo calcularemos de la primerasemana del
            'mes de noviembre de 2004 que empieza ekl 1 Lunes
            I = Val(RecuperaValor(CADENA, 3))
            If I > 7 Then Exit Function
            'I = I + 1
            Buffer$ = Buffer$ & Format(I & "/11/2004", "ddd")
            
            
            I = Val(RecuperaValor(CADENA, 2))
            If I = 0 Then Exit Function
            Buffer$ = Buffer$ & " " & I & ","
            
            I = Val(RecuperaValor(CADENA, 1))
            If I = 0 Then Exit Function
            Buffer$ = Buffer$ & "  " & Format("01/" & I & "/2004", "mmm")
            
            'Hora
            I = Val(RecuperaValor(CADENA, 4))
            Buffer$ = Buffer$ & "  " & I & ":"
            
            'Minutos
            I = Val(RecuperaValor(CADENA, 5))
            Buffer$ = Buffer$ & Format(I, "00")
            
            
            Label5.Caption = UCase(Buffer$)
            PonerHoras = True
        End If
        
    End If
    
    Buffer$ = ""
End Function



Private Function Leer(Segundos As Integer, CadenaEsperada As String) As String
Dim I As Integer
Dim Fin As Boolean
Dim T1, T2
Dim Buffer2$
Dim C As String

Leer = ""
Fin = False
T1 = Timer
I = 1
'Bucle
While Not Fin
    C = MSComm1.Input
    Buffer2$ = Buffer2$ & C
    Text1.Text = Text1.Text & C
    If InStr(1, Buffer2$, CadenaEsperada) Then
        Fin = True
        Leer = Buffer2$
        Else
            T2 = Timer - T1
            If T2 > I Then
                Text1.Text = Text1.Text & Format(Now, "hh:mm:ss") & vbCrLf
                Text1.Refresh
                I = Round(T2, 0) + 1
            End If
            Fin = (T2 > Segundos)
    End If
Wend
Text1.SelStart = Len(Text1.Text)
End Function



Private Sub EscribeTextoFichero(Texto As String)
'ProcesaTexto texto
Print #NF, Texto
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdSecuencia_Click()
Dim C1 As String
Dim C2 As String
Dim I As Integer
Dim T As String


   If Text4.Text = "" Then Exit Sub
    T = Text4.Text
   Screen.MousePointer = vbHourglass
   
    'Comprobaremos la fecha y hora del reloj
    'Si esta configurado y demas
    If Not MSComm1.PortOpen Then MSComm1.PortOpen = True
   
        
    Text1.Text = ""
 
    
    
    'Solictamos comando
    Do
        If Text4.Text = "" Then
            'naaaa
        Else
            I = InStr(1, Text4.Text, vbCrLf)
            If I > 0 Then
                C1 = Mid(Text4.Text, 1, I - 1)
                Text4.Text = Mid(Text4.Text, I + 2)
                
            Else
                C1 = Text4.Text
                Text4.Text = ""
            End If
            C2 = RecuperaValor(C1, 2)
            C1 = RecuperaValor(C1, 1)
            
            If C1 = "" Or C2 = "" Then
                'Mal
                
            Else
                I = CInt(Val(C2)) 'espera
                If I = 0 Then I = 1
                Text1.Text = Text1.Text & "-->" & C1 & vbCrLf
                PonerTexto kTCP
                C1 = Leer(I, "OK")
                LimpiaBufferRecepcion
            End If
        End If
    Loop Until Text4.Text = ""
    
    
    Text1.Text = Text1.Text & vbCrLf & finalizado
    Text4.Text = T
    
    
    

    

Salida3:
   
    

    'Cerramos el puerto
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    MSComm1.CommPort = "1"
    MSComm1.Settings = "19200,N,8,1"
    kTCP = "tcp1"
    Me.Caption = Me.Caption & "  " & App.Major & "." & App.Minor & "." & App.Revision
End Sub


Private Function RecuperaValor(ByRef CADENA As String, Orden As Integer) As String
Dim I As Integer
Dim J As Integer
Dim Cont As Integer
Dim Cad As String

I = 0
Cont = 1
Cad = ""
Do
    J = I + 1
    I = InStr(J, CADENA, "|")
    If I > 0 Then
        If Cont = Orden Then
            Cad = Mid(CADENA, J, I - J)
            I = Len(CADENA) 'Para salir del bucle
            Else
                Cont = Cont + 1
        End If
    End If
Loop Until I = 0
RecuperaValor = Cad
End Function



Private Sub Text3_GotFocus()
    Text3.SelStart = 0
    Text3.SelLength = Len(Text3.Text)
        
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmdEnviar.SetFocus
        
End Sub
