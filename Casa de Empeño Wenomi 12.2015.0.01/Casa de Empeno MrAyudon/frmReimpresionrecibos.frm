VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmReimpresionrecibos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-impresión recibos"
   ClientHeight    =   1080
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReimpresionrecibos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1080
   ScaleWidth      =   2925
   Begin VB.TextBox txtFolio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   1455
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   2
      Top             =   585
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
      Picture         =   "frmReimpresionrecibos.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   150
      Width           =   1095
      _ExtentX        =   1931
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
      Object.ToolTipText     =   ""
      Picture         =   "frmReimpresionrecibos.frx":055E
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Folio:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   810
   End
End
Attribute VB_Name = "frmReimpresionrecibos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim Folio As Long, strTabla As String, strCampo As String, strCondicion As String, EjecutaConsulta As Boolean

Private Sub cmdAceptar_Click()
Dim rcReimpresion As New ADODB.Recordset
    
    If Trim(txtFolio.text) <> "" Then
        
        Folio = 0
        rcReimpresion.Open "SELECT " & strCampo & " FROM " & strTabla & strCondicion & Val(txtFolio.text), dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcReimpresion.BOF And Not rcReimpresion.EOF Then
            
            Folio = rcReimpresion.Fields(0)

        End If
        rcReimpresion.Close
        Set rcReimpresion = Nothing
    End If
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl()
End Sub

Private Sub txtFolio_GotFocus()
    Seleccionar_Texto txtFolio
    Cambiar_Color True, txtFolio
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
End Sub

Private Sub txtFolio_LostFocus()
    Cambiar_Color False, txtFolio
End Sub

Public Function ReImprimir(Tabla As String, Campo As String, Condicion As String) As Long
    Folio = -1
    strTabla = Tabla
    strCampo = Campo
    strCondicion = Condicion
    Me.Show vbModal
    ReImprimir = Folio
End Function
