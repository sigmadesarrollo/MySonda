VERSION 5.00
Begin VB.Form frmCambio 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cambio"
   ClientHeight    =   3465
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambio.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCambio 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   855
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "0.00"
      Top             =   2160
      Width           =   3975
   End
   Begin VB.TextBox txtEfectivo 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   120
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   600
      Width           =   3975
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   33.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   975
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "0.00"
      Top             =   3840
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "CAMBIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1620
      TabIndex        =   4
      Top             =   1680
      Width           =   1440
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "EFECTIVO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   375
      Left            =   1410
      TabIndex        =   2
      Top             =   120
      Width           =   1740
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IMPORTE"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1470
      TabIndex        =   0
      Top             =   3360
      Visible         =   0   'False
      Width           =   1620
   End
End
Attribute VB_Name = "frmCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then Unload Me
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

Public Sub Mostrar(crImporte As Double)
    txtImporte.text = Format(crImporte, FMoneda)
    Me.Show vbModal
End Sub

Private Sub txtCambio_GotFocus()
    Seleccionar_Texto txtCambio
    Cambiar_Color True, txtCambio
End Sub

Private Sub txtCambio_LostFocus()
    Cambiar_Color False, txtCambio
End Sub

Private Sub txtEfectivo_Change()
    CalculaCambio
End Sub

Private Sub txtEfectivo_GotFocus()
    Seleccionar_Texto txtEfectivo
    Cambiar_Color True, txtEfectivo
End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtEfectivo_LostFocus()
    Cambiar_Color False, txtEfectivo
End Sub

Private Sub txtImporte_GotFocus()
    Seleccionar_Texto txtImporte
    Cambiar_Color True, txtImporte
End Sub

Private Sub txtImporte_LostFocus()
    Cambiar_Color False, txtImporte
End Sub

Sub CalculaCambio()
Dim crImporte As Double, crEfectivo As Double
    
    If Val(txtImporte.text) > 0 Or Trim(txtImporte.text) <> "" Then
        
        crImporte = CDbl(txtImporte.text)
    Else
        
        crImporte = 0
    End If
    
    If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
        
        crEfectivo = CDbl(txtEfectivo.text)
    Else
        
        crEfectivo = 0
    End If
    
    
    If crEfectivo > 0 Then
        
        txtCambio.text = Format(crEfectivo - crImporte, FMoneda)
    Else
        
        txtCambio.text = Format(0, FMoneda)
    End If
    
End Sub
