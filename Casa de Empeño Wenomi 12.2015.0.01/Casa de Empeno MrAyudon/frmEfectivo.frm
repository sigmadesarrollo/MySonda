VERSION 5.00
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmEfectivo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   2355
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4830
   Icon            =   "frmEfectivo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   2310
      Left            =   4800
      Top             =   0
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   4075
      Orientation     =   0
      ShadowColor     =   8421631
      LigthColor      =   8421631
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   30
      Index           =   0
      Left            =   -120
      Top             =   2280
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   53
      ShadowColor     =   8421631
      LigthColor      =   8421631
      LineWidth       =   3
   End
   Begin VB.TextBox txtEfectivo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   35.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   75
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   960
      Width           =   4680
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   2310
      Left            =   0
      Top             =   0
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   4075
      Orientation     =   0
      ShadowColor     =   8421631
      LigthColor      =   8421631
      LineWidth       =   3
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   30
      Index           =   1
      Left            =   0
      Top             =   0
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   53
      ShadowColor     =   8421631
      LigthColor      =   8421631
      LineWidth       =   10
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "EFECTIVO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   27
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   660
      Left            =   780
      TabIndex        =   1
      Top             =   225
      Width           =   3285
   End
End
Attribute VB_Name = "frmEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim crImporte As Double, crEfectivo As Double, Pestana As Integer

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = vbKeyEscape Then crEfectivo = -1: Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtEfectivo_Change()
'''''Dim Pos As Integer
    
'''''    txtEfectivo.text = Format(txtEfectivo.text, FMoneda)
'''''    Pos = InStr(1, Trim(txtEfectivo.text), ".")
'''''    txtEfectivo.SelStart = IIf(Pos > 0, Pos - 1, 0)
End Sub

Private Sub txtEfectivo_GotFocus()
    Seleccionar_Texto txtEfectivo
    Cambiar_Color True, txtEfectivo
End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
    
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtEfectivo.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    If KeyAscii = vbKeyReturn Then
        
        crEfectivo = 0
        If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
            
            crEfectivo = txtEfectivo.text
        End If
                
        If crEfectivo < crImporte Then
            
            MsgBox "Monto insuficiente, favor de revisar el efectivo !!", vbInformation, IIf(Pestana = 1, "Desempeño", "Refrendo")
        End If
        
        Unload Me
    End If
    
End Sub

Private Sub txtEfectivo_LostFocus()
    Cambiar_Color False, txtEfectivo
End Sub

Public Function RegresaCambio(crImp As Double, Pesta As Integer) As Double
    crImporte = crImp
    Pestana = Pesta
    Me.Show vbModal
    RegresaCambio = crEfectivo
End Function

Sub Inicializar()
    Poner_Flat Fl, Me.Controls, Me
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
