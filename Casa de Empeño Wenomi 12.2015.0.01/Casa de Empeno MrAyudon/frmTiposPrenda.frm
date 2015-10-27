VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmTiposPrenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Prenda"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4125
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTiposPrenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbTipoPrenda 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   600
      Width           =   2655
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2880
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
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmTiposPrenda.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Aceptar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmTiposPrenda.frx":055E
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo prenda:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1530
   End
End
Attribute VB_Name = "frmTiposPrenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim TipoPrenda As Long

Private Sub cmbTipoPrenda_GotFocus()
    Cambiar_Color True, cmbTipoPrenda
End Sub

Private Sub cmbTipoPrenda_LostFocus()
    Cambiar_Color False, cmbTipoPrenda
End Sub

Private Sub cmdAceptar_Click()
    
    If cmbTipoPrenda.ListIndex = 0 Then
        
        'Si selecciona TODOS
        TipoPrenda = 0
    
    ElseIf cmbTipoPrenda.ListIndex = (cmbTipoPrenda.ListCount - 1) Then
        
        'Si selecciona AUTOS
        TipoPrenda = -1
    Else
        
        'Si selecciona cualquier otro
        TipoPrenda = cmbTipoPrenda.ItemData(cmbTipoPrenda.ListIndex)
    End If
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    TipoPrenda = -2
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Poner_Flat Fl, Me.Controls, Me
    cmbTipoPrenda.AddItem "(TODOS)"
    Cargar_Combos "Descripcion", "tipo", cmbTipoPrenda, , , False
    cmbTipoPrenda.AddItem "AUTOS"
    cmbTipoPrenda.ListIndex = 0
    CentrarForm Me, frmMDI
End Sub

Public Function Mostrar() As Long
    TipoPrenda = 0
    Me.Show vbModal
    Mostrar = TipoPrenda
End Function

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub
