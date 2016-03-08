VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmTicajeManual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Introducción manual de tareas"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmTicajeManual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4665
   ScaleWidth      =   6000
   Begin VB.CommandButton Command3 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   4920
      TabIndex        =   11
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&GENERAR"
      Height          =   375
      Left            =   1260
      TabIndex        =   10
      Top             =   4200
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   180
      TabIndex        =   9
      Top             =   4200
      Width           =   975
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   5040
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   420
      Width           =   675
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   3900
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   420
      Width           =   975
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000018&
      Enabled         =   0   'False
      Height          =   315
      Left            =   900
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   420
      Width           =   2835
   End
   Begin VB.TextBox Text1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   420
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   420
      Top             =   3360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTicajeManual.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2955
      Left            =   120
      TabIndex        =   7
      Top             =   1140
      Width           =   5760
      _ExtentX        =   10160
      _ExtentY        =   5212
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Tarjeta"
         Object.Width           =   6174
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   1320
      Picture         =   "frmTicajeManual.frx":08A4
      Top             =   900
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   1020
      Picture         =   "frmTicajeManual.frx":09A6
      Top             =   900
      Width           =   240
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   4440
      Picture         =   "frmTicajeManual.frx":0AA8
      Top             =   180
      Width           =   240
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   600
      Picture         =   "frmTicajeManual.frx":0BAA
      Top             =   180
      Width           =   240
   End
   Begin VB.Label Label4 
      Caption         =   "Trabajador"
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   900
      Width           =   825
   End
   Begin VB.Label Label3 
      Caption         =   "Hora"
      Height          =   195
      Left            =   5040
      TabIndex        =   6
      Top             =   180
      Width           =   345
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha"
      Height          =   195
      Left            =   3900
      TabIndex        =   5
      Top             =   180
      Width           =   450
   End
   Begin VB.Label Label1 
      Caption         =   "Tarea "
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   180
      Width           =   465
   End
End
Attribute VB_Name = "frmTicajeManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1
Private WithEvents frmF As frmCal
Attribute frmF.VB_VarHelpID = -1

Dim Cad As String


Private Sub Command1_Click()
    Cad = "¿Seguro que desea borrar los trabajadores"
    If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then ListView1.ListItems.Clear
End Sub

Private Sub Command2_Click()
Dim RS As ADODB.Recordset

    'Comprobaciones
    If Text1.Text = "" Then
        MsgBox "Escriba la tarea.", vbExclamation
        Exit Sub
    End If
    If Text3.Text = "" Then
        MsgBox "Escriba la fecha.", vbExclamation
        Exit Sub
    End If
    
    If Text4.Text = "" Then
        MsgBox "Escriba la hora.", vbExclamation
        Exit Sub
    End If
    
    If ListView1.ListItems.Count <= 0 Then
        MsgBox "Añada algun trabajador", vbExclamation
        Exit Sub
    End If
    
    
    Cad = DevuelveDesdeBD("Tarjeta", "Tareas", "idTarea", Text1.Text, "N")
    Text1.Tag = Cad
    If Cad = "" Then
        MsgBox "Codigo tarjeta para la tarea no encontrado", vbExclamation
        Exit Sub
    End If
    
    'Comprobamos k no esta cerrado
    Set RS = New ADODB.Recordset
    Cad = "Select count(*) from TareasRealizadas WHERE fecha = #" & Format(Text3.Text, "yyyy/mm/dd") & "#"
    RS.Open Cad, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    Cad = ""
    If Not RS.EOF Then
        If Not IsNull(RS.Fields(0)) Then
            If RS.Fields(0) > 0 Then
                Cad = CStr(RS.Fields(0))
            End If
        End If
    End If
    RS.Close
    
    If Cad <> "" Then
        MsgBox "El dia ya ha sido procesado", vbExclamation
    Else
        Screen.MousePointer = vbHourglass
        'O todos o nadie
        Conn.BeginTrans
        If GenerarLaTicada Then
            Conn.CommitTrans
            MsgBox "El ticaje se ha generado con éxito", vbExclamation
        Else
            Conn.RollbackTrans
        End If
        Screen.MousePointer = vbDefault
    End If
    
    Set RS = Nothing
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text4.Text = ""
    ListView1.ListItems.Clear
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
Dim itmX As ListItem
    
On Error GoTo EIntroduce
    If Image1.Tag = 1 Then
        Text1.Text = vCodigo
        Text2.Text = vCadena
    Else
        'Añadimos trabajador
        Cad = DevuelveDesdeBD("numTarjeta", "Trabajadores", "idTrabajador", CStr(vCodigo), "N")
        If Cad <> "" Then
            Set itmX = ListView1.ListItems.Add(, "C" & vCodigo, vCadena)
            itmX.SubItems(1) = Cad
            itmX.SmallIcon = 1
            itmX.Tag = vCodigo
        Else
            MsgBox "Error leyendo tarjeta", vbExclamation
        End If
    End If
    Exit Sub
EIntroduce:
    MuestraError Err.Number, vCadena & vbCrLf & Err.Description
End Sub



Private Sub frmF_Selec(vFecha As Date)
    Text3.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
    Image1.Tag = 1
    Set frmB = New frmBusca
    frmB.Tabla = "Tareas"
    frmB.CampoBusqueda = "Descripcion"
    frmB.CampoCodigo = "IdTarea"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "TAREAS"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub Image2_Click()
    Set frmF = New frmCal
    frmF.Fecha = Now
    If Text3.Text <> "" Then
        If IsDate(Text3.Text) Then frmF.Fecha = CDate(Text3.Text)
    End If
    frmF.Show vbModal
    Set frmF = Nothing
End Sub

Private Sub Image3_Click()
    Image1.Tag = 0
    Set frmB = New frmBusca
    frmB.Tabla = "Trabajadores"
    frmB.CampoBusqueda = "Nomtrabajador"
    frmB.CampoCodigo = "IdTrabajador"
    frmB.MostrarDeSalida = True
    frmB.TipoDatos = 3
    frmB.Titulo = "TRABAJADORES"
    frmB.Show vbModal
    Set frmB = Nothing
End Sub

Private Sub Image4_Click()
    If ListView1.SelectedItem Is Nothing Then
        MsgBox "Seleccione un trabajador a eliminar de la lista", vbExclamation
        Exit Sub
    End If
    
    'Eliminar de la lista
    Cad = ListView1.SelectedItem.Text & " (" & ListView1.SelectedItem.SubItems(1) & ")  ?"
    Cad = "Seguro que desea quitar de la lista al trabajdor: " & vbCrLf & Cad
    If MsgBox(Cad, vbQuestion + vbYesNoCancel) = vbYes Then
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
    End If
End Sub

Private Sub PonFoco(ByRef T As TextBox)
    On Error Resume Next
    T.SetFocus
    If Err.Number <> 0 Then Err.Clear
End Sub

Private Sub Text1_LostFocus()
    Text1.Text = Trim(Text1.Text)
    If Text1.Text = "" Then
        Text2.Text = ""
        Exit Sub
    End If
    
    If Not IsNumeric(Text1.Text) Then
        MsgBox "Campo debe ser numérico", vbExclamation
        Cad = ""
    Else
        Cad = DevuelveDesdeBD("descripcion", "tareas", "idTarea", Text1.Text, "N")
        If Cad <> "" Then
            Text2.Text = Cad
        Else
            MsgBox "No exise la tarea codigo: " & Text1.Text, vbExclamation
        End If
    End If
    If Cad = "" Then
        Text1.Text = ""
        Text2.Text = ""
        PonFoco Text1
    End If
End Sub

Private Sub Text3_LostFocus()
    Text3.Text = Trim(Text3.Text)
    If Text3.Text = "" Then Exit Sub
    If Not IsDate(Text3.Text) Then
        MsgBox "No es una fecha válida: " & Text3.Text, vbExclamation
        Text3.Text = ""
        PonFoco Text3
    Else
        Text3.Text = Format(Text3.Text, "dd/mm/yyyy")
    End If
End Sub



Private Sub Text4_LostFocus()
    Text4.Text = Trim(Text4.Text)
    If Text4.Text = "" Then Exit Sub
    Text4.Text = TransformaPuntosHoras(Text4.Text)
    If Not IsDate(Text4.Text) Then
        MsgBox "No es una hora válida: " & Text4.Text, vbExclamation
        Text4.Text = ""
        PonFoco Text4
    Else
        Text4.Text = Format(Text4.Text, "hh:mm")
    End If
End Sub


Private Function GenerarLaTicada() As Boolean
Dim Fec As Date
Dim HInsertar As Date
Dim RS As ADODB.Recordset
Dim H1 As Date
Dim Fin As Boolean
Dim Insertar As Boolean
Dim I As Integer

    On Error GoTo EGenerarLaTicada
    GenerarLaTicada = False
    Set RS = New ADODB.Recordset
    
  
    
    
    'Buscamos el hueco
    Cad = "SELECT * FROM MarcajesKimaldi where fecha = #" & Format(Text3.Text, "yyyy/mm/dd") & "#"
    Fec = CDate(Text4.Text & ":00")
    Cad = Cad & " AND Hora >= #" & Format(Fec, "hh:mm") & ":00#"
    H1 = DateAdd("n", 5, Fec)
    Cad = Cad & " AND Hora <= #" & Format(H1, "hh:mm:ss") & "#"
    Cad = Cad & " ORDER BY Hora"
    RS.Open Cad, Conn, adOpenKeyset, adLockOptimistic, adCmdText
    Insertar = False
    'Hora a insertar
    HInsertar = CDate(Text4.Text & ":00")
    
    If RS.EOF Then
        Insertar = True
    Else
            
        H1 = HInsertar
        Fin = False
        Do While Not Fin
            If Mid(RS!Marcaje, 1, 1) <> mConfig.DigitoTrabajadores Then
                If DateDiff("s", H1, RS!Hora) > 2 Then
                    'Significa k esto es una tarea y el marcaje a insertar esta
                    'Tres segundos pod delante
                    Insertar = True
                    Fin = True
                Else
                    H1 = RS!Hora
                End If
            Else
                H1 = RS!Hora
            End If
            RS.MoveNext
            If RS.EOF Then Fin = True
        Loop
    
    End If
    
    
    'Si hay k insertar generaremos un marcaje
    If Insertar Then
        Fec = CDate(Text3.Text)
        'La tarea
        HInsertar = DateAdd("s", 1, H1)
        Cad = "INSERT INTO MarcajesKimaldi(Nodo,Fecha,Hora,TipoMens,Marcaje) VALUES "
        Cad = Cad & "(999,#" & Format(Fec, "yyyy/mm/dd") & "#,#"
        Cad = Cad & Format(HInsertar, "hh:mm:ss") & "#,'','" & Text1.Tag & "')"
        Conn.Execute Cad
        'Los trabajadores
        HInsertar = DateAdd("s", 2, H1)
        Cad = "INSERT INTO MarcajesKimaldi(Nodo,Fecha,Hora,TipoMens,Marcaje) VALUES "
        Cad = Cad & "(999,#" & Format(Fec, "yyyy/mm/dd") & "#,#" & Format(HInsertar, "hh:mm:ss") & "#,'','"
        For I = 1 To ListView1.ListItems.Count
             Conn.Execute Cad & ListView1.ListItems(I).SubItems(1) & "')"
        Next I
    Else
        MsgBox "La aplicación no puede encotrar hueco para esa hora e inserar la nueva tarea", vbExclamation
        Exit Function
    End If
    GenerarLaTicada = True
    Exit Function
EGenerarLaTicada:
    MuestraError Err.Number, "Ticadas"
End Function



