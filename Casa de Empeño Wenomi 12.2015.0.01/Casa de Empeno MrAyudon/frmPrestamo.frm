VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmPrestamos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prestamos"
   ClientHeight    =   1575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrestamo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1575
   ScaleWidth      =   8775
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4680
      MaxLength       =   7
      TabIndex        =   3
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtConcepto 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   3975
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7560
      TabIndex        =   10
      Top             =   1080
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
      Picture         =   "frmPrestamo.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   6360
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AutoSize        =   0   'False
      Caption         =   "&Aceptar"
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
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Left            =   8400
      TabIndex        =   8
      Top             =   600
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      AlignCaption    =   4
      AlignPicture    =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmPrestamo.frx":009D
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosClave 
      Height          =   285
      Left            =   4080
      TabIndex        =   11
      Top             =   480
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      AutoSize        =   0   'False
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
   End
   Begin VB.Label lblFolio 
      AutoSize        =   -1  'True
      Caption         =   "<Folio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   7200
      TabIndex        =   5
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Folio:"
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
      Left            =   6360
      TabIndex        =   4
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
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
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1050
   End
End
Attribute VB_Name = "frmPrestamos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 26/04/2002
' Modulo frmPrestamo - frmPrestamo.frm
' Ultima Modificacion - 26/04/2002
'
'////////////////////////////////////////////////////////////////


Option Explicit

Dim Fl() As cFlatControl


Private Sub cmdMosClave_Click()
'frmMostrarUsuarios.Ver Me, txtConcepto
End Sub

Private Sub cmdMosFecha_Click()
   txtFecha.Text = frmCalendario.Fecha(txtFecha.Text)
End Sub

Private Sub txtFecha_GotFocus()
   Seleccionar_Texto txtFecha
   Cambiar_Color True, txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtFecha_LostFocus()
   Cambiar_Color False, txtFecha
End Sub

Private Sub cmdAceptar_Click()
  If Validar Then Grabar_Datos
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  inicializar
End Sub

Private Sub inicializar()
  Screen.MousePointer = vbHourglass
  CentrarForm Me, frmMDI
  txtFecha.Text = Format(Date, "DD/MM/YY")
  lblFolio.Caption = Regresa_Movimiento(False, "FolioPrestamos")
  Poner_Flat Fl, Me.Controls, Me
  Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Quitar_Flat Fl
End Sub

Private Sub txtConcepto_GotFocus()
  Seleccionar_Texto txtConcepto
  Cambiar_Color True, txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
  Pasar_Foco KeyAscii
End Sub

Private Sub txtConcepto_LostFocus()
  Cambiar_Color False, txtConcepto
End Sub

Private Sub txtImporte_GotFocus()
  Seleccionar_Texto txtImporte
  Cambiar_Color True, txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
  KeyAscii = Solo_Numeros(KeyAscii, 1)
  Pasar_Foco KeyAscii
End Sub

Private Sub txtImporte_LostFocus()
  Cambiar_Color False, txtImporte
End Sub

'Grabamos los datos
Private Sub Grabar_Datos()
  Dim Movimiento As Long
  Dim Folio As Long
  Dim Importe As Currency
  
  
  Movimiento = Regresa_Movimiento(False)
  Folio = Regresa_Movimiento(False, "FolioPrestamos")
  Regresa_Movimiento True
  Regresa_Movimiento True, "FolioPrestamos"
  Importe = Format(CCur(Val(txtImporte.Text)), "###########0.00")
  
  dbDatos.Execute "INSERT INTO Prestamos (Folio,Fecha,Concepto,Importe,IDUsuario) VALUES " & _
                  "(" & Folio & ",#" & Format(txtFecha.Text, "MM/DD/YY") & "#,'" & txtConcepto.Text & "'," & Importe & "," & Val(txtConcepto.Tag) & ")"
                  
  dbDatos.Execute "UPDATE Usuarios SET Prestamo=Prestamo+" & CCur(txtImporte.Text) & " WHERE ID=" & Val(txtConcepto.Tag) & ""
                    
  'Grabamos el cargo
  dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#," & Movimiento & "," & Folio & ",'PR50','110150'," & Importe & "," & TIPO_ABONO & ",0,'" & txtConcepto.Text & "','" & Nombre_PC & "')"
                  
  dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#," & Movimiento & "," & Folio & ",'PR50','199450'," & Importe & "," & TIPO_ABONO & ",0,'" & txtConcepto.Text & "','" & Nombre_PC & "')"
                  
                  
  'Grabamos el abono
  dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC) VALUES " & _
                  "(#" & Format(Date, "MM/DD/YY") & "#," & Movimiento & "," & Folio & ",'PR01','151301'," & Importe & "," & TIPO_CARGO & ",0,'" & txtConcepto.Text & "','" & Nombre_PC & "')"
  
  
  
  limpiar
  lblFolio.Caption = Regresa_Movimiento(False, "FolioPrestamos")
End Sub

'Validamos que esten correctos los datos
Private Function Validar() As Boolean
  Validar = True
  
  If Trim(txtConcepto.Text) = "" Then
    MsgBox "Imposible grabar la venta, Datos incompletos", vbOKOnly + vbCritical
    txtConcepto.SetFocus
    Validar = False
    Exit Function
  End If
  
  If Trim(txtImporte.Text) = "" Then
    MsgBox "Imposible grabar la venta, Datos incompletos", vbOKOnly + vbCritical
    txtImporte.SetFocus
    Validar = False
    Exit Function
  End If
  
  If Not IsDate(txtFecha.Text) Then
   MsgBox "Imposible de grabar el deposito, Favor de poner una fecha valida", vbOKOnly + vbCritical
   Validar = False
   txtFecha.SetFocus
  End If
    
End Function

'Limpiamos los campos
Private Sub limpiar()
  txtConcepto.Text = ""
  txtImporte.Text = ""
  txtFecha.Text = Format(Date, "DD/MM/YY")
  lblFolio.Caption = ""
End Sub


