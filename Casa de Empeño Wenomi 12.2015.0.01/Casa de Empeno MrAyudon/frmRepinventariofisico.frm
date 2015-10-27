VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRepinventariofisico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inventario Físico"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepinventariofisico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1380
   ScaleWidth      =   3510
   Begin VB.OptionButton opTipos 
      Appearance      =   0  'Flat
      Caption         =   "Tipos de prenda"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1470
   End
   Begin VB.ComboBox cmbTipo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmRepinventariofisico.frx":000C
      Left            =   240
      List            =   "frmRepinventariofisico.frx":0022
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   600
      Width           =   1935
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   2280
      TabIndex        =   2
      Top             =   840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "&Salir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmRepinventariofisico.frx":003E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   2280
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Mostrar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmRepinventariofisico.frx":00CF
   End
End
Attribute VB_Name = "frmRepinventariofisico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fl() As cFlatControl

Private Sub cmbTipo_GotFocus()
cmbTipo.BackColor = &HC0FFFF
End Sub

Private Sub cmbTipo_LostFocus()
cmbTipo.BackColor = vbWhite
End Sub

Private Sub cmdAceptar_Click()
Screen.MousePointer = vbHourglass
frmInventariofisico.MuestraArticulos IIf(cmbTipo.ItemData(cmbTipo.ListIndex) = 0, -1, cmbTipo.ItemData(cmbTipo.ListIndex))
Screen.MousePointer = default
Unload Me
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass
Poner_Flat fl, Me.Controls, Me
CentrarForm Me, frmMDI
Cargar_Combos "Descripcion", "Tipo", cmbTipo
cmbTipo.ListIndex = 0
Screen.MousePointer = vbDefault
End Sub

'cargamos el tipo de prenda
Private Sub Cargar_Combos(Campo As String, Tabla As String, Combo As ComboBox)
On Error GoTo error

Dim rcTipos As New ADODB.Recordset

rcTipos.Open "SELECT * FROM " & Tabla, dbDatos, adOpenDynamic, adLockOptimistic

Combo.Clear
Combo.AddItem "(TODOS)"
With rcTipos
   While Not .EOF
      Combo.AddItem .Fields(Campo)
      Combo.ItemData(Combo.NewIndex) = !ID
      .MoveNext
   Wend
End With

rcTipos.Close
error:
Maneja_Error Err

Set rcTipos = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
Quitar_Flat fl()
End Sub
