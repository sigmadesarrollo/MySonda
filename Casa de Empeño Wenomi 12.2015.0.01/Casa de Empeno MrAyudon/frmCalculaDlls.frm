VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCalculaDlls 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calcula Dlls"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCalculaDlls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2985
   ScaleWidth      =   5685
   Begin VB.TextBox txtResultado 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   39.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   1440
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   2520
      TabIndex        =   2
      Top             =   120
      Width           =   3150
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   630
         Left            =   2040
         TabIndex        =   7
         Top             =   120
         Width           =   1065
         Begin VB.OptionButton optCompra 
            Appearance      =   0  'Flat
            Caption         =   "Compra"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   75
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton optVenta 
            Appearance      =   0  'Flat
            Caption         =   "Venta"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   345
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbDivisas 
         Height          =   315
         Left            =   135
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   720
         Width           =   1965
      End
      Begin VB.OptionButton optDivisa 
         Appearance      =   0  'Flat
         Caption         =   "Divisa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   450
         Width           =   855
      End
      Begin VB.OptionButton optMoneda 
         Appearance      =   0  'Flat
         Caption         =   "Moneda Nacional"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   180
         Value           =   -1  'True
         Width           =   1620
      End
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   24.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   720
      Left            =   120
      TabIndex        =   0
      Text            =   "0.00"
      Top             =   480
      Width           =   2295
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   4515
      TabIndex        =   6
      Top             =   2520
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
      Picture         =   "frmCalculaDlls.frx":000C
   End
   Begin VB.Label lblCambio 
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1185
   End
End
Attribute VB_Name = "frmCalculaDlls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Sub Inicializar()
    Cargar_Combos "Descripcion", "monedas", cmbDivisas, , "Descripcion", , "Clave"
    cmbDivisas.ListIndex = 0
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub cmbDivisas_Click()
    Calcula
End Sub

Private Sub cmbDivisas_GotFocus()
    Cambiar_Color True, cmbDivisas
End Sub

Private Sub cmbDivisas_LostFocus()
    Cambiar_Color False, cmbDivisas
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub optCompra_Click()
    Calcula
End Sub

Private Sub optDivisa_Click()
    Calcula
End Sub

Private Sub optMoneda_Click()
    Calcula
End Sub

Private Sub optVenta_Click()
    Calcula
End Sub

Private Sub txtCantidad_Change()
    Calcula
End Sub

Private Sub txtCantidad_GotFocus()
    Seleccionar_Texto txtCantidad
    Cambiar_Color True, txtCantidad
End Sub

Private Sub txtCantidad_LostFocus()
    Cambiar_Color False, txtCantidad
End Sub

Sub Calcula()
Dim Cantidad As Double, TipoCambio As Double, crCambio As Double, crTotal As Double, strTipoCambio As String
    
    lblCambio.Visible = False: Cantidad = 0: crCambio = 0: TipoCambio = 0: strTipoCambio = ""
    If Val(txtCantidad.text) > 0 Or Trim(txtCantidad.text) <> "" Then
        
        Cantidad = txtCantidad.text
    End If
    
    strTipoCambio = SacaValor("cotizaciones", IIf(optCompra.Value, "Compra", "Venta"), " WHERE ID=" & Val(SacaValor("cotizaciones", "MAX(ID)", " WHERE IDMoneda=" & cmbDivisas.ItemData(cmbDivisas.ListIndex))))
    If Trim(strTipoCambio) <> "" Then
        
        TipoCambio = CDbl(strTipoCambio)
    End If
    
    If optMoneda.Value And TipoCambio > 0 Then
        
        crTotal = Int(Cantidad / TipoCambio)
        crCambio = Cantidad - (crTotal * TipoCambio)
    Else
        
        crTotal = Cantidad * TipoCambio
    End If
    
    txtResultado.text = Format(crTotal, IIf(optDivisa.Value, FMoneda, ""))
    
    If optMoneda.Value And crCambio > 0 And optCompra.Value = False Then
        lblCambio.Visible = True
        lblCambio.Caption = "Cambio: " & Format(crCambio, FMoneda)
    End If
End Sub

Private Sub txtResultado_GotFocus()
    Seleccionar_Texto txtResultado
    Cambiar_Color True, txtResultado
End Sub

Private Sub txtResultado_LostFocus()
    Cambiar_Color False, txtResultado
End Sub
