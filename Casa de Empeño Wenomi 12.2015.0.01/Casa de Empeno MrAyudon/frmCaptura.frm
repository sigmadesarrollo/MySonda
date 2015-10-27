VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.10#0"; "FlatBtn2.ocx"
Begin VB.Form frmCaptura 
   Caption         =   "Captura de Datos"
   ClientHeight    =   4365
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5985
   Icon            =   "frmCaptura.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4365
   ScaleWidth      =   5985
   Begin VB.TextBox txtSucursal 
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
      Left            =   1440
      TabIndex        =   8
      Top             =   240
      Width           =   4335
   End
   Begin VB.TextBox txtHaber 
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
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   12
      Top             =   2520
      Width           =   2055
   End
   Begin VB.TextBox txtDebe 
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
      Left            =   3000
      MaxLength       =   9
      TabIndex        =   11
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   3
      Top             =   1320
      Width           =   5655
      Begin VB.TextBox txtSaldo 
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
         Left            =   2880
         MaxLength       =   9
         TabIndex        =   10
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Anterior:"
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
         Left            =   600
         TabIndex        =   9
         Top             =   240
         Width           =   1860
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Debe:"
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
         Left            =   1770
         TabIndex        =   7
         Top             =   720
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Haber:"
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
         Left            =   1650
         TabIndex        =   6
         Top             =   1200
         Width           =   825
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   1770
         TabIndex        =   5
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "<Total>"
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
         Left            =   2730
         TabIndex        =   4
         Top             =   1680
         Width           =   2175
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4560
      TabIndex        =   13
      Top             =   3720
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
      Picture         =   "frmCaptura.frx":030A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   3720
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
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal:"
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
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1110
   End
   Begin VB.Label Label3 
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
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   795
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "<Fecha>"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   960
   End
End
Attribute VB_Name = "frmCaptura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.I. Jorge Gabriel Colio Ramos
' Mazatlan, Sin. 15/08/2002
' Modulo frmImportarInformacion - frmImportarInformacion.frm
' Ultima Modificacion - 16/08/2002
'
'////////////////////////////////////////////////////////////////



Option Explicit
Dim fl() As cFlatControl

Private Sub cmdAceptar_Click()
   If txtSucursal = "" Or txtSaldo = "" Or txtDebe = "" Or txtHaber = "" Then
      MsgBox ("Favor de rellenar las cajas de texto requeridas")
      If txtSucursal = "" Then
         txtSucursal.SetFocus
      ElseIf txtSaldo = "" Then
         txtSaldo.SetFocus
      ElseIf txtDebe = "" Then
         txtDebe.SetFocus
      ElseIf txtHaber = "" Then
         txtHaber.SetFocus
      End If
   Else
      If MsgBox("Estan correctos los Datos que va a ingresar", vbOKCancel) = vbOK Then
         Grabar
         Limpiar
      Else
         txtSaldo.SetFocus
      End If
   End If
End Sub

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   Inicializa
End Sub

Private Sub Inicializa()
   Screen.MousePointer = vbHourglass
   Me.Height = 4770
   Me.Width = 6105
   CentrarForm Me, frmMDI
   lblFecha.Caption = Format(Date, "DD/MMMM/YYYY")
   Poner_Flat fl, Me.Controls, Me
   frmCaptura.Caption = "Captura Datos de Surcusales " + Regresa_Valor("MONTEPIO", "Sucursal", "")
   Screen.MousePointer = vbDefault
End Sub

Private Sub Grabar()
   dbReportes.Execute "INSERT INTO CorteCajaVentanilla (Saldo,Debe,Haber) VALUES " & _
   "('" & txtSaldo.Text & "','" & txtDebe.Text & "','" & txtHaber.Text & "' )"
End Sub

Private Sub txtDebe_Change()
   Total
End Sub

Private Sub txtDebe_GotFocus()
   Seleccionar_Texto txtDebe
   Cambiar_Color True, txtDebe
End Sub

Private Sub txtDebe_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtDebe_LostFocus()
   Cambiar_Color False, txtDebe
End Sub

Private Sub txtHaber_Change()
   Total
End Sub

Private Sub txtHaber_GotFocus()
   Seleccionar_Texto txtHaber
   Cambiar_Color True, txtHaber
End Sub

Private Sub txtHaber_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtHaber_LostFocus()
   Cambiar_Color False, txtHaber
End Sub

Private Sub txtSaldo_Change()
   Total
End Sub

Private Sub txtSaldo_GotFocus()
   Seleccionar_Texto txtSaldo
   Cambiar_Color True, txtSaldo
End Sub

Private Sub txtSaldo_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtSaldo_LostFocus()
   Cambiar_Color False, txtSaldo
End Sub

Private Sub txtSucursal_GotFocus()
   Seleccionar_Texto txtSucursal
   Cambiar_Color True, txtSucursal
End Sub

Private Sub txtSucursal_KeyPress(KeyAscii As Integer)
  Pasar_Foco KeyAscii
End Sub

Private Sub txtSucursal_LostFocus()
   Cambiar_Color False, txtSucursal
End Sub

Private Sub Total()
lblTotal = (Val(txtSaldo.Text) + Val(txtDebe.Text)) - Val(txtHaber.Text)
End Sub

Private Sub Limpiar()
   txtSucursal = ""
   txtSaldo = ""
   txtDebe = ""
   txtHaber = ""
End Sub
