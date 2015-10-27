VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmAgregaMedio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Medio"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgregaMedio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4230
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDescripcion 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      MaxLength       =   50
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   735
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmAgregaMedio.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   735
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Aceptar"
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
      TextColor       =   4210752
      Object.ToolTipText     =   ""
      Picture         =   "frmAgregaMedio.frx":055E
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Medio:"
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
      TabIndex        =   3
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmAgregaMedio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim IDMedio As Integer

Private Sub cmdAceptar_Click()
    
    If Trim(txtDescripcion.text) <> "" Then

        dbDatos.Execute "INSERT INTO medios (Descripcion) VALUES ('" & _
                        Trim(txtDescripcion.text) & "')"
        
        IDMedio = SacaValor("medios", "MAX(ID)")
        Unload Me
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Screen.MousePointer = vbHourglass
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Texto txtDescripcion
    Cambiar_Color True, txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescripcion_LostFocus()
    Cambiar_Color False, txtDescripcion
End Sub

Public Function Mostrar() As Integer
    IDMedio = 0
    Me.Show vbModal
    Mostrar = IDMedio
End Function
