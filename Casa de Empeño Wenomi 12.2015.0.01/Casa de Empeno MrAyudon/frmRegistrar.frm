VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRegistrar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activación MySonda"
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4755
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRegistrar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4755
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCodigoSeguridad 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1920
      TabIndex        =   0
      Top             =   1560
      Width           =   2775
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   1920
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "  &Aceptar"
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
      Picture         =   "frmRegistrar.frx":000C
      PictureDisabled =   "frmRegistrar.frx":0376
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Top             =   1920
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
      Picture         =   "frmRegistrar.frx":04D0
   End
   Begin VB.Label lblDiasUso 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   3120
      TabIndex        =   9
      Top             =   840
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dias uso:"
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   840
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Clave de activación:"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Código del software:"
      Height          =   195
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Versión:"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   585
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmRegistrar.frx":0A22
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      Caption         =   "ACTIVAR MYSONDA"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   240
      Width           =   3450
   End
End
Attribute VB_Name = "frmRegistrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As New cFlatControl
Public Salir As Integer

Private Sub cmdAceptar_Click()
    
    If Trim(txtCodigo.text) = "" Then
        
        MsgBox "Introduzca el código de activación del softwate !!", vbCritical, "Activación MySonda"
        txtCodigo.SetFocus
    Else
        
        frmMDI.ActiveLock2.LiberationKey = Trim(txtCodigo.text)
        If frmMDI.ActiveLock2.RegisteredUser Then
            
            MsgBox "Clave de activación correcta !!" & Chr(13) & "Su Software ha sido registrado con éxito !!", vbInformation, "Activación"
            End
        Else
            
            MsgBox "Clave de activación incorrecta !!", vbCritical, "Activación MySonda"
        End If
    
    End If
End Sub

Private Sub cmdSalir_Click()
    If Salir <> 0 Then
        
        Unload Me
    Else
    
        Quitar_Flat Fl
        End
    End If
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    Poner_Flat Fl, Me.Controls, Me
    lblDiasUso.Caption = frmMDI.ActiveLock2.UsedDays
    txtCodigoSeguridad.text = frmMDI.ActiveLock2.SoftwareCode
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtCodigo_GotFocus()
    Seleccionar_Texto txtCodigo
    Cambiar_Color True, txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCodigo_LostFocus()
    Cambiar_Color False, txtCodigo
End Sub

'Ponemos los controles en modo flat
Public Sub Poner_Flat(ByRef Fl() As cFlatControl, Controles As Object, forma As Form)
Dim Contador As Integer
Dim Control As Object
   
   For Each Control In Controles
      
      If TypeOf Control Is TextBox Or TypeOf Control Is MaskEdBox Then
         
         ReDim Preserve Fl(0 To Contador)
         Set Fl(Contador) = New cFlatControl
         Fl(Contador).hWndAttach Control.hWnd, forma.hWnd, False
         Contador = Contador + 1
      
      ElseIf TypeOf Control Is ComboBox Then
         
         ReDim Preserve Fl(0 To Contador)
         Set Fl(Contador) = New cFlatControl
         Fl(Contador).hWndAttach Control.hWnd, forma.hWnd, True
         Contador = Contador + 1
      
      End If
   
   Next
End Sub

Public Sub Quitar_Flat(ByRef Fl() As cFlatControl)
Dim i As Integer
   
    'Descargamos de memoria el flat
    For i = LBound(Fl) To UBound(Fl)
      
      Set Fl(i) = Nothing
    
    Next i
    
End Sub
