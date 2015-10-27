VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmAutorizacionPrestamo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Autorización Préstamo"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAutorizacionPrestamo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   4380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtClaveAutorizada 
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
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1005
      Width           =   3015
   End
   Begin VB.TextBox txtClaveAutorizar 
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
      Height          =   255
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   420
      Width           =   3015
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   285
      Width           =   1005
      _ExtentX        =   1773
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
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmAutorizacionPrestamo.frx":000C
      PictureDisabled =   "frmAutorizacionPrestamo.frx":0376
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   5
      Top             =   840
      Width           =   1005
      _ExtentX        =   1773
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
      Picture         =   "frmAutorizacionPrestamo.frx":04D0
   End
   Begin VB.Label Label2 
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
   Begin VB.Label Label1 
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
      TabIndex        =   2
      Top             =   120
      Width           =   2460
   End
End
Attribute VB_Name = "frmAutorizacionPrestamo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim Bandera As Boolean, Autorizacion As Long

Public Function GeneraCodigo(strCodigo As String, ByRef CodigoValido As Boolean, IDAutorizacion As Long)
Dim strParte1 As String, strParte2 As String

    strParte1 = Mid(Trim(strCodigo), 1, 8)
    strParte2 = Mid(Trim(strCodigo), 9, 16)
    
    txtClaveAutorizar = Hex(strParte1) & Oct(strParte2)
    Bandera = False
    Autorizacion = 0
    Me.Show vbModal
    CodigoValido = Bandera
    IDAutorizacion = Autorizacion
End Function

Private Sub cmdAceptar_Click()
Dim strParte1 As String, strParte2 As String
Dim strConvertido As String, i As Integer

On Error GoTo error

    strParte1 = Mid(Trim(txtClaveAutorizar.text), 1, 8)
    strParte2 = Mid(Trim(txtClaveAutorizar.text), 9, Len(Trim(txtClaveAutorizar.text)))
    
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
    
    If Trim(txtClaveAutorizada.text) = strParte1 & strParte2 Then
        
        Bandera = True
        dbDatos.Execute "INSERT INTO autorizaciones (Fecha,IDUsuario,IDSucursal,Codigo,Opcion) VALUES ('" & _
                        Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & strParte1 & strParte2 & "',1)"
        
        Autorizacion = SacaValor("autorizaciones", "MAX(ID)")
        Unload Me
    Else
        
        Bandera = False
        MsgBox "Clave de autorización inválida, favor de verificarla !!", vbInformation, "Autorización Préstamo"
    End If

error:
    Maneja_Error Err
End Sub

Private Sub cmdSalir_Click()
    Bandera = False
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtClaveAutorizada_GotFocus()
    Seleccionar_Texto txtClaveAutorizada
    Cambiar_Color True, txtClaveAutorizada
End Sub

Private Sub txtClaveAutorizada_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtClaveAutorizada_LostFocus()
    Cambiar_Color False, txtClaveAutorizada
End Sub

Private Sub txtClaveAutorizar_GotFocus()
    Seleccionar_Texto txtClaveAutorizar
    Cambiar_Color True, txtClaveAutorizar
End Sub

Private Sub txtClaveAutorizar_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtClaveAutorizar_LostFocus()
    Cambiar_Color False, txtClaveAutorizar
End Sub
