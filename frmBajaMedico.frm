VERSION 5.00
Begin VB.Form frmBajaMedico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form2"
   ClientHeight    =   3525
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6780
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6780
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameBaja 
      Height          =   3495
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6615
      Begin VB.TextBox txtFecha 
         Height          =   315
         Index           =   0
         Left            =   1560
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtIncid 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.TextBox txtDesIncid 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1560
         Width           =   3615
      End
      Begin VB.TextBox txtDesTra 
         BackColor       =   &H80000018&
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox txtTra 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   0
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   0
         Left            =   1560
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtHora 
         Height          =   285
         Index           =   1
         Left            =   3960
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   2280
         Width           =   975
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancelar"
         Height          =   375
         Index           =   0
         Left            =   5160
         TabIndex        =   6
         Top             =   2880
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   3840
         TabIndex        =   5
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha "
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   1020
         Width           =   495
      End
      Begin VB.Image Image2 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmBajaMedico.frx":0000
         Top             =   960
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Incidencia"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Trabajador"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   765
      End
      Begin VB.Image ImgTrab 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmBajaMedico.frx":0102
         Top             =   360
         Width           =   240
      End
      Begin VB.Image ImgInc 
         Height          =   240
         Index           =   0
         Left            =   1080
         Picture         =   "frmBajaMedico.frx":0204
         Top             =   1560
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hora INCIO"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   11
         Top             =   2280
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hora FIN"
         Height          =   195
         Index           =   4
         Left            =   3120
         TabIndex        =   10
         Top             =   2280
         Width           =   645
      End
   End
End
Attribute VB_Name = "frmBajaMedico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmC As frmCal
Attribute frmC.VB_VarHelpID = -1
Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1

Public opcion As Byte
    '0.- Nueva baja
    '1.- Listado bajas



Dim Cad As String
Dim primeravez As Boolean

Private Sub cmdAceptar_Click()
Dim L As Long

    On Error GoTo EA
    If txtTra(0).Text = "" Or txtIncid(0).Text = "" Or txtFecha(0).Text = "" Or txtHora(0).Text = "" Or txtHora(1).Text = "" Then
        MsgBox "Todos los campos obligatorios", vbExclamation
        Exit Sub
    End If
    
    
    If CDate(txtHora(0).Text) > CDate(txtHora(1).Text) Then
        MsgBox "Hora fina mayor hora inicio", vbExclamation
        Exit Sub
    End If
    
    'Creamos dos inserciones en entradafichajes
    
    Cad = txtTra(0).Text & " AND fecha = #" & Format(txtFecha(0).Text, FormatoFecha) & "#"
    Cad = Cad & " AND idinci = " & txtIncid(0).Text
    Cad = DevuelveDesdeBD("hora", "entradafichajes", "idtrabajador", Cad)
    If Cad <> "" Then
        'Ya exsite un marcaje de esa incidencia
        MsgBox "Ya tiene una incidencia: " & Me.txtDesIncid(0).Text, vbExclamation
        Exit Sub
    End If
    
    Cad = "Select max(secuencia) from entradafichajes"
    L = ObtenerMaximoMinimo(Cad) + 1
    
    Cad = "INSERT INTO entradafichajes(secuencia,idtrabajador,fecha,hora,idinci,horareal) VALUES ("
    Cad = Cad & L & "," & txtTra(0).Text & ",#" & Format(txtFecha(0).Text, FormatoFecha) & "#,#" & txtHora(0).Text & "#,"
    Cad = Cad & txtIncid(0).Text & ",#" & txtHora(0).Text & "#)"
    Conn.Execute Cad
    
    L = L + 1
    Cad = "INSERT INTO entradafichajes(secuencia,idtrabajador,fecha,hora,idinci,horareal) VALUES ("
    Cad = Cad & L & "," & txtTra(0).Text & ",#" & Format(txtFecha(0).Text, FormatoFecha) & "#,#" & txtHora(1).Text & "#,"
    Cad = Cad & "0,#" & txtHora(1).Text & "#)"
    Conn.Execute Cad
    VariableCompartida = "0K"
    Unload Me
    Exit Sub
EA:
    MuestraError Err.Number, Cad
End Sub


Private Sub cmdCancel_Click(Index As Integer)
    Unload Me
End Sub

Private Sub Form_Activate()
    If primeravez Then
        primeravez = False
        If txtFecha(0).Text <> "" Then
            PonleFoco txtHora(0)
        Else
            PonleFoco Me.txtTra(0)
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Icon = frmPpal1.Icon
    primeravez = True
    Limpiar Me
    Caption = "Visita medica"
    Me.FrameBaja.Visible = opcion = 0
    'Pongo la baja medica a mano. Igual deberiamos parametrizarlo
    If opcion = 0 Then
        txtIncid(0).Text = MiEmpresa.IncVisitaMedica
    End If
    cmdCancel(opcion).Cancel = True
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
    Cad = vCodigo & "|" & vCadena & "|"
End Sub

Private Sub Image2_Click(Index As Integer)
    Cad = ""
    Set frmC = New frmCal
    frmC.Fecha = Now
    If txtFecha(Index).Text <> "" Then frmC.Fecha = CDate(txtFecha(Index))
    frmC.Show vbModal
    Set frmC = Nothing
    If Cad <> "" Then
        txtFecha(Index).Text = Cad
    End If
End Sub

Private Sub ImgInc_Click(Index As Integer)
    Cad = ""
    Set frmB = New frmBusca
    frmB.Tabla = "Incidencias"
    frmB.CampoBusqueda = "NomInci"
    frmB.CampoCodigo = "IdInci"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "INCIDENCIAS"
    frmB.Show vbModal
    Set frmB = Nothing
    If Cad <> "" Then
        Me.txtIncid(Index).Text = RecuperaValor(Cad, 1)
        Me.txtDesIncid(Index).Text = RecuperaValor(Cad, 2)
    End If
End Sub

Private Sub ImgTrab_Click(Index As Integer)
    Cad = ""
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "NomTrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.TipoDatos = 3
    frmB.Titulo = "EMPLEADOS"
    frmB.Show vbModal
    Set frmB = Nothing
    If Cad <> "" Then
        Me.txtTra(Index).Text = RecuperaValor(Cad, 1)
        Me.txtDesTra(Index).Text = RecuperaValor(Cad, 2)
    End If
End Sub



Private Sub txtFecha_GotFocus(Index As Integer)
    PonerFoco txtFecha(Index)
End Sub

Private Sub txtFecha_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub


Private Sub txtFecha_LostFocus(Index As Integer)
    If Not EsFechaOK(txtFecha(Index)) Then txtFecha(Index).Text = ""
End Sub

Private Sub txtHora_GotFocus(Index As Integer)
    PonerFoco txtHora(Index)
End Sub

Private Sub txtHora_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtHora_LostFocus(Index As Integer)
Dim i As Integer

    txtHora(Index).Text = Trim(txtHora(Index).Text)
    If txtHora(Index).Text = "" Then Exit Sub

    Do
        i = InStr(1, txtHora(Index).Text, ".")
        If i > 0 Then
            Cad = Mid(txtHora(Index).Text, i + 1)
            If Len(Cad) = 1 Then
                If Val(Cad) > 5 Then
                    Cad = "0" & Cad
                Else
                    Cad = Cad & "0"
                End If
            End If
            txtHora(Index).Text = Mid(txtHora(Index).Text, 1, i - 1) & ":" & Cad
        End If
    Loop While i <> 0
    
    If txtHora(0).Text <> "" Then
        If Not IsDate(txtHora(Index).Text) Then
            MsgBox "Error en el campo hora: " & txtHora(Index).Text, vbExclamation
            txtHora(Index).Text = ""
            PonleFoco txtHora(Index)
    
        End If
    End If

End Sub

Private Sub txtIncid_GotFocus(Index As Integer)
    PonerFoco txtIncid(Index)
End Sub

Private Sub txtIncid_LostFocus(Index As Integer)
  If Not IsNumeric(txtIncid(Index).Text) Then
        txtIncid(Index).Text = -1
        txtDesIncid(Index).Text = "Error en la incidencia"
        Else
            Cad = DevuelveTextoIncidencia(CInt(txtIncid(Index).Text))
            If Cad = "" Then
                txtIncid(Index).Text = -1
                txtDesIncid(Index).Text = "Error en la incidencia"
                Else
                    txtDesIncid(Index).Text = Cad
            End If
    End If
End Sub

Private Sub txtIncid_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub


Private Sub txtTra_GotFocus(Index As Integer)
    PonerFoco txtTra(Index)
End Sub

Private Sub txtTra_KeyPress(Index As Integer, KeyAscii As Integer)
    Keypress KeyAscii
End Sub

Private Sub txtTra_LostFocus(Index As Integer)
  If Not IsNumeric(txtTra(Index).Text) Then
        txtTra(Index).Text = ""
        txtDesTra(Index).Text = "Campo numerico"
        Else
            Cad = devuelveNombreTrabajador(CInt(txtTra(Index).Text))
            If Cad = "" Then
                txtTra(Index).Text = -1
                txtDesTra(Index).Text = "Error en el trabajador"
                Else
                    txtDesTra(Index).Text = Cad
            End If
    End If
End Sub

