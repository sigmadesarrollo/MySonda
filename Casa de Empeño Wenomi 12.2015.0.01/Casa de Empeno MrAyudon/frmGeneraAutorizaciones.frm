VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmGeneraAutorizaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generador de autorizaciones"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGeneraAutorizaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1335
   ScaleWidth      =   4215
   Begin VB.TextBox txtCodigo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   2790
   End
   Begin VB.TextBox txtClaveGenerada 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1005
      Width           =   2790
   End
   Begin DevPowerFlatBttn.FlatBttn cmdGenerar 
      Default         =   -1  'True
      Height          =   375
      Left            =   3060
      TabIndex        =   2
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Generar"
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
      Object.ToolTipText     =   ""
      Picture         =   "frmGeneraAutorizaciones.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3060
      TabIndex        =   5
      Top             =   915
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
      Picture         =   "frmGeneraAutorizaciones.frx":055E
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Clave por Autorizar:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   270
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Clave Autorización"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   270
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   2265
   End
End
Attribute VB_Name = "frmGeneraAutorizaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdGenerar_Click()
    GeneraCodigos
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtClaveGenerada_GotFocus()
    Seleccionar_Texto txtClaveGenerada
    Cambiar_Color True, txtClaveGenerada
End Sub

Private Sub txtClaveGenerada_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtClaveGenerada_LostFocus()
    Cambiar_Color False, txtClaveGenerada
End Sub

Sub GeneraCodigos()
Dim strParte1 As String, strParte2 As String
Dim strConvertido As String, i As Integer

    strParte1 = Mid(Trim(txtCodigo.text), 1, 8)
    strParte2 = Mid(Trim(txtCodigo.text), 9, Len(Trim(txtCodigo.text)))
    
    strConvertido = ""
    For i = 1 To 8
    
        Select Case Mid(strParte1, i, 1)
        Case "A"
        
            strConvertido = strConvertido & "1"
        
        Case "B"
            
            strConvertido = strConvertido & "2"
        
        Case "C"
            
            strConvertido = strConvertido & "3"
        
        Case "D"
            
            strConvertido = strConvertido & "4"
        
        Case "E"
            
            strConvertido = strConvertido & "5"
        
        Case "F"
            
            strConvertido = strConvertido & "6"
        
        Case Else
        
            strConvertido = strConvertido & Mid(strParte1, i, 1)
        End Select
    Next i
    
    strParte1 = Oct(strConvertido)
    strParte2 = Hex(strParte2)
    
    txtClaveGenerada.text = strParte1 & strParte2
End Sub

Private Sub txtCodigo_GotFocus()
    Seleccionar_Texto txtCodigo
    Cambiar_Color True, txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCodigo_LostFocus()
    Cambiar_Color False, txtCodigo
End Sub
