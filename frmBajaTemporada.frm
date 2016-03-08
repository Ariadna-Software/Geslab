VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmBajaTemporada 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Baja temporada"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6825
   Icon            =   "frmBajaTemporada.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDelete 
      Height          =   375
      Left            =   660
      Picture         =   "frmBajaTemporada.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eliminar"
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton cmdAdd 
      Height          =   375
      Left            =   180
      Picture         =   "frmBajaTemporada.frx":040C
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Añadir"
      Top             =   3120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Baja"
      Height          =   375
      Left            =   3900
      TabIndex        =   3
      Top             =   3120
      Width           =   1275
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   3120
      Width           =   1275
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   4471
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   10583
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Trabajadores"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmBajaTemporada"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents frmB As frmBusca
Attribute frmB.VB_VarHelpID = -1



Private Sub cmdAdd_Click()
        Set frmB = New frmBusca
        frmB.Tabla = "Trabajadores"
        frmB.CampoBusqueda = "NomTrabajador"
        frmB.CampoCodigo = "IdTrabajador"
        frmB.TipoDatos = 3
        frmB.Titulo = "EMPLEADOS"
        frmB.Show vbModal
        Set frmB = Nothing
End Sub

Private Sub cmdDelete_Click()
Dim SQ As String
    If ListView1.ListItems.Count < 1 Then Exit Sub
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    SQ = "Va a eliminar de la lista al trabjador : " & ListView1.SelectedItem.Text & vbCrLf
    SQ = SQ & Space(30) & "¿Desea continuar?"
    If MsgBox(SQ, vbQuestion + vbYesNoCancel) = vbYes Then
        ListView1.ListItems.Remove ListView1.SelectedItem.Index
        Me.Refresh
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub frmB_Seleccion(vCodigo As Long, vCadena As String)
Dim It As ListItem
    On Error GoTo EV
        
    Set It = ListView1.ListItems.Add(, CStr("C" & vCodigo))
    It.Text = vCadena
    It.Tag = vCodigo
    
    Exit Sub
EV:
    MuestraError Err.Number, "Insertar trabajador" & vbCrLf & Err.Description
End Sub
