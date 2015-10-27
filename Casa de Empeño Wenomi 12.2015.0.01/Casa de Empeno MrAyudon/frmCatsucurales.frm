VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCatsucursales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de sucursales"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatsucurales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   6870
   Begin VB.Frame Frame3 
      Caption         =   "Registro Público del Contrato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   32
      Top             =   5760
      Width           =   6615
      Begin VB.TextBox txtNumeroRegistro 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2280
         TabIndex        =   13
         Top             =   840
         Width           =   3855
      End
      Begin VB.TextBox txtContratoAdhesion 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   2280
         TabIndex        =   12
         Top             =   480
         Width           =   3855
      End
      Begin MSMask.MaskEdBox txtFechaRegistro 
         Height          =   240
         Left            =   2280
         TabIndex        =   36
         Top             =   1200
         Width           =   1620
         _ExtentX        =   2858
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   10
         Format          =   "dd-mmmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de Registro:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Width           =   1560
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Número de Registro:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   34
         Top             =   840
         Width           =   1725
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contrato de Adhesión:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   1875
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Dudas, Aclaraciones y Reclamaciones"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   28
      Top             =   4080
      Width           =   6615
      Begin VB.TextBox txtCorreoAclaraciones 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         TabIndex        =   11
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtTelefonoAclaraciones 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtDomicilioAclaraciones 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         TabIndex        =   9
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correo Electrónico:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   31
         Top             =   1200
         Width           =   1590
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   30
         Top             =   840
         Width           =   780
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Domicilio:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   29
         Top             =   480
         Width           =   810
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Sucursal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtNomcomercial 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   2
         Top             =   1200
         Width           =   4215
      End
      Begin VB.TextBox txtRazonsocial 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   1
         Top             =   840
         Width           =   4215
      End
      Begin VB.TextBox txtDireccion 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   4
         Top             =   1920
         Width           =   4215
      End
      Begin VB.TextBox txtCiudad 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   60
         TabIndex        =   5
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox txtEstado 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txtTelefono 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   25
         TabIndex        =   7
         Top             =   3000
         Width           =   1215
      End
      Begin VB.TextBox txtRfc 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtCp 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   5
         TabIndex        =   8
         Top             =   3360
         Width           =   1215
      End
      Begin VB.TextBox txtClave 
         BorderStyle     =   0  'None
         Height          =   225
         Left            =   1920
         MaxLength       =   100
         TabIndex        =   0
         Top             =   480
         Width           =   735
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente2 
         Height          =   255
         Left            =   6135
         TabIndex        =   18
         Top             =   840
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         AlignCaption    =   4
         AutoSize        =   0   'False
         Caption         =   ". . ."
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Comercial:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   27
         Top             =   1200
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Razón Social:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   840
         Width           =   1110
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dirección:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   25
         Top             =   1920
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ciudad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Estado:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   23
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Teléfono:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   22
         Top             =   3000
         Width           =   780
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "RFC:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   21
         Top             =   1560
         Width           =   360
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "C.P."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   20
         Top             =   3360
         Width           =   300
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Clave:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   510
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   4470
      TabIndex        =   15
      Top             =   7455
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Limpiar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   255
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmCatsucurales.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5550
      TabIndex        =   17
      Top             =   7455
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
      Picture         =   "frmCatsucurales.frx":0110
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3285
      TabIndex        =   14
      Top             =   7455
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
      TextColor       =   4210752
      Object.ToolTipText     =   ""
      Picture         =   "frmCatsucurales.frx":0662
   End
End
Attribute VB_Name = "frmCatsucursales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()

    Dim Sql As String

    If Valida Then
    
        If txtRazonsocial.Tag <> "" Then
        
            If MsgBox("Desea guardar los cambios ??", vbQuestion + vbYesNo + vbDefaultButton2, "Catálogo de sucursales") = vbYes Then
            
                Sql = "UPDATE Sucursales SET Clave=" & Val(txtClave.text) & ", " & "Razonsocial='" & Trim(txtRazonsocial.text) & "', " & "Nombrecomercial='" & Trim(txtNomcomercial.text) & "', " & "Direccion='" & Trim(txtDireccion.text) & "', " & "Rfc='" & Trim(txtRFC.text) & "', " & "Ciudad='" & Trim(txtCiudad.text) & "', " & "Estado='" & Trim(txtEstado.text) & "', " & "Telefono='" & Trim(txtTelefono.text) & "', " & "Cp=" & Val(txtCP.text) & ", " & _
                    "DomicilioAclaraciones='" & Trim(txtDomicilioAclaraciones.text) & "', " & "TelefonoAclaraciones='" & Trim(txtTelefonoAclaraciones.text) & "', " & "CorreoAclaraciones='" & Trim(txtCorreoAclaraciones.text) & "', " & _
                    "ContratoRegistrado='" & Trim(txtNumeroRegistro.text) & "', " & "FechaContratoRegistrado='" & Format(txtFechaRegistro.text, "YYYY-MM-DD") & "',CodProfeco='" & Trim(txtContratoAdhesion.text) & "'" & _
                    " WHERE ID=" & Val(txtRazonsocial.Tag) & ""
                    
                dbDatos.Execute Sql
            End If
        End If

        Limpiar
        txtRazonsocial.SetFocus
    End If
    
End Sub

Private Sub cmdLimpiar_Click()
    Limpiar
    txtClave.SetFocus
End Sub

Private Sub cmdMosCliente2_Click()
    frmMostrarSucursales.Ver Me, txtRazonsocial, True, True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
    Inicializar
    Screen.MousePointer = vbDefault
    
End Sub

Sub Limpiar()
    txtClave.text = ""
    txtNomcomercial.text = ""
    txtRazonsocial.text = ""
    txtRazonsocial.Tag = ""
    txtDireccion.text = ""
    txtCiudad.text = ""
    txtEstado.text = ""
    txtTelefono.text = ""
    txtCP.text = ""
    txtRFC.text = ""
    txtDomicilioAclaraciones.text = ""
    txtTelefonoAclaraciones.text = ""
    txtCorreoAclaraciones.text = ""
    txtContratoAdhesion.text = ""
    txtNumeroRegistro.text = ""
    txtFechaRegistro.text = Date
End Sub

Private Sub txtCiudad_GotFocus()
    Seleccionar_Texto txtCiudad
    Cambiar_Color True, txtCiudad
End Sub

Private Sub txtCiudad_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCiudad_LostFocus()
    Cambiar_Color False, txtCiudad
End Sub

Private Sub txtClave_GotFocus()
    Seleccionar_Texto txtClave
    Cambiar_Color True, txtClave
End Sub

Private Sub txtClave_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtClave_LostFocus()
    Cambiar_Color False, txtClave
End Sub

Private Sub txtCP_GotFocus()
    Seleccionar_Texto txtCP
    Cambiar_Color True, txtCP
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCP_LostFocus()
    Cambiar_Color False, txtCP
End Sub

Private Sub txtDireccion_GotFocus()
    Seleccionar_Texto txtDireccion
    Cambiar_Color True, txtDireccion
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDireccion_LostFocus()
    Cambiar_Color False, txtDireccion
End Sub

Private Sub txtEstado_GotFocus()
    Seleccionar_Texto txtEstado
    Cambiar_Color True, txtEstado
End Sub

Private Sub txtEstado_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEstado_LostFocus()
    Cambiar_Color False, txtEstado
End Sub

Private Sub txtNomcomercial_GotFocus()
    Seleccionar_Texto txtNomcomercial
    Cambiar_Color True, txtNomcomercial
End Sub

Private Sub txtNomcomercial_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNomcomercial_LostFocus()
    Cambiar_Color False, txtNomcomercial
End Sub

Private Sub txtRazonsocial_GotFocus()
    Seleccionar_Texto txtRazonsocial
    Cambiar_Color True, txtRazonsocial
End Sub

Private Sub txtRazonsocial_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRazonsocial_LostFocus()
    Cambiar_Color False, txtRazonsocial
End Sub

Private Sub txtRfc_GotFocus()
    Seleccionar_Texto txtRFC
    Cambiar_Color True, txtRFC
End Sub

Private Sub txtRfc_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRfc_LostFocus()
    Cambiar_Color False, txtRFC
End Sub

Private Sub txtTelefono_GotFocus()
    Seleccionar_Texto txtTelefono
    Cambiar_Color True, txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTelefono_LostFocus()
    Cambiar_Color False, txtTelefono
End Sub

Private Sub txtDomicilioAclaraciones_GotFocus()
    Seleccionar_Texto txtDomicilioAclaraciones
    Cambiar_Color True, txtDomicilioAclaraciones
End Sub

Private Sub txtDomicilioAclaraciones_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDomicilioAclaraciones_LostFocus()
    Cambiar_Color False, txtDomicilioAclaraciones
End Sub

Private Sub txtTelefonoAclaraciones_GotFocus()
    Seleccionar_Texto txtTelefonoAclaraciones
    Cambiar_Color True, txtTelefonoAclaraciones
End Sub

Private Sub txtTelefonoAclaraciones_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTelefonoAclaraciones_LostFocus()
    Cambiar_Color False, txtTelefonoAclaraciones
End Sub

Private Sub txtCorreoAclaraciones_GotFocus()
    Seleccionar_Texto txtCorreoAclaraciones
    Cambiar_Color True, txtCorreoAclaraciones
End Sub

Private Sub txtCorreoAclaraciones_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCorreoAclaraciones_LostFocus()
    Cambiar_Color False, txtCorreoAclaraciones
End Sub

Private Sub txtContratoAdhesion_GotFocus()
    Seleccionar_Texto txtContratoAdhesion
    Cambiar_Color True, txtContratoAdhesion
End Sub

Private Sub txtContratoAdhesion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtContratoAdhesion_LostFocus()
    Cambiar_Color False, txtContratoAdhesion
End Sub

Private Sub txtNumeroRegistro_GotFocus()
    Seleccionar_Texto txtNumeroRegistro
    Cambiar_Color True, txtNumeroRegistro
End Sub

Private Sub txtNumeroRegistro_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumeroRegistro_LostFocus()
Cambiar_Color False, txtNumeroRegistro
End Sub

Private Sub txtFechaRegistro_GotFocus()
    Seleccionar_Texto txtFechaRegistro
    Cambiar_Color True, txtFechaRegistro
End Sub

Private Sub txtFechaRegistro_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaRegistro_LostFocus()
    Cambiar_Color False, txtFechaRegistro
End Sub


Public Function BuscarSucursal(IDSucursal As Long)

On Error GoTo Error

    Set rcConsulta = dbDatos.Execute("select * from sucursales where Id=" & IDSucursal & "")

    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        With rcConsulta
            txtRazonsocial.Tag = !ID
            If Not IsNull(!Clave) Then txtClave.text = !Clave Else txtClave.text = ""
            If Not IsNull(!RazonSocial) Then txtRazonsocial.text = !RazonSocial Else txtRazonsocial.text = ""
            If Not IsNull(!NombreComercial) Then txtNomcomercial.text = !NombreComercial Else txtNomcomercial.text = ""
            If Not IsNull(!RFC) Then txtRFC.text = !RFC Else txtRFC.text = ""
            If Not IsNull(!Direccion) Then txtDireccion.text = !Direccion Else txtDireccion.text = ""
            If Not IsNull(!Ciudad) Then txtCiudad.text = !Ciudad Else txtCiudad.text = ""
            If Not IsNull(!Estado) Then txtEstado.text = !Estado Else txtEstado.text = ""
            If Not IsNull(!Telefono) Then txtTelefono.text = !Telefono Else txtTelefono.text = ""
            If Not IsNull(!CP) Then txtCP.text = !CP Else txtCP.text = ""
            If Not IsNull(!DomicilioAclaraciones) Then txtDomicilioAclaraciones.text = !DomicilioAclaraciones Else txtDomicilioAclaraciones.text = ""
            If Not IsNull(!TelefonoAclaraciones) Then txtTelefonoAclaraciones.text = !TelefonoAclaraciones Else txtTelefonoAclaraciones.text = ""
            If Not IsNull(!CorreoAclaraciones) Then txtCorreoAclaraciones.text = !CorreoAclaraciones Else txtCorreoAclaraciones.text = ""
            If Not IsNull(!ContratoRegistrado) Then txtNumeroRegistro.text = !ContratoRegistrado Else txtNumeroRegistro.text = ""
            If Not IsNull(!FechaContratoRegistrado) Then txtFechaRegistro.text = !FechaContratoRegistrado Else txtFechaRegistro.text = Date
            If Not IsNull(!CodProfeco) Then txtContratoAdhesion.text = !CodProfeco Else txtContratoAdhesion.text = ""
        End With
    
    End If

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function

Sub Inicializar()

    Dim rc As ADODB.Recordset

On Error GoTo Error

    Set rc = dbDatos.Execute("SELECT * FROM sucursales WHERE activa = 1")
    
    If Not rc.BOF And Not rc.EOF Then
    
        With rc
            txtRazonsocial.Tag = rc!ID
            If Not IsNull(!Clave) Then txtClave.text = !Clave Else txtClave.text = ""
            If Not IsNull(!RazonSocial) Then txtRazonsocial.text = !RazonSocial Else txtRazonsocial.text = ""
            If Not IsNull(!NombreComercial) Then txtNomcomercial.text = !NombreComercial Else txtNomcomercial.text = ""
            If Not IsNull(!RFC) Then txtRFC.text = !RFC Else txtRFC.text = ""
            If Not IsNull(!Direccion) Then txtDireccion.text = !Direccion Else txtDireccion.text = ""
            If Not IsNull(!Ciudad) Then txtCiudad.text = !Ciudad Else txtCiudad.text = ""
            If Not IsNull(!Estado) Then txtEstado.text = !Estado Else txtEstado.text = ""
            If Not IsNull(!Telefono) Then txtTelefono.text = !Telefono Else txtTelefono.text = ""
            If Not IsNull(!CP) Then txtCP.text = !CP Else txtCP.text = ""
            
            If Not IsNull(!DomicilioAclaraciones) Then txtDomicilioAclaraciones.text = !DomicilioAclaraciones Else txtDomicilioAclaraciones.text = ""
            If Not IsNull(!TelefonoAclaraciones) Then txtTelefonoAclaraciones.text = !TelefonoAclaraciones Else txtTelefonoAclaraciones.text = ""
            If Not IsNull(!CorreoAclaraciones) Then txtCorreoAclaraciones.text = !CorreoAclaraciones Else txtCorreoAclaraciones.text = ""
            
            If Not IsNull(!CodProfeco) Then txtContratoAdhesion.text = !CodProfeco Else txtContratoAdhesion.text = ""
            If Not IsNull(!ContratoRegistrado) Then txtNumeroRegistro.text = !ContratoRegistrado Else txtNumeroRegistro.text = ""
            If Not IsNull(!FechaContratoRegistrado) Then txtFechaRegistro.text = !FechaContratoRegistrado 'Else txtFechaRegistro.text = ""
        End With
        
    End If

Error:
    Maneja_Error Err
    Set rc = Nothing
End Sub

Public Function Valida() As Boolean
    
    Valida = True

    If txtClave.text = "" Then
        MsgBox "Introduzca la clave !!", vbInformation, "Catálogo de sucursales"
        Valida = False
        txtClave.SetFocus
        Exit Function
    End If
    
    If txtRazonsocial.text = "" Then
        MsgBox "Introduzca la razón social !!", vbInformation, "Catálogo de sucursales"
        Valida = False
        txtRazonsocial.SetFocus
        Exit Function
    End If
    
    If txtNomcomercial.text = "" Then
        MsgBox "Introduzca el nombre comercial !!", vbInformation, "Catálogo de sucursales"
        Valida = False
        txtNomcomercial.SetFocus
        Exit Function
    End If
    
    If txtRFC.text = "" Then
        MsgBox "Introduzca el Rfc !!", vbInformation, "Catálogo de sucursales"
        Valida = False
        txtRFC.SetFocus
        Exit Function
    End If
    
    If txtCiudad.text = "" Then
        MsgBox "Introduzca la ciudad !!", vbInformation, "Catálogo de sucursales"
        Valida = False
        txtCiudad.SetFocus
        Exit Function
    End If
    
    If txtEstado.text = "" Then
        MsgBox "Introduzca el estado !!", vbInformation, "Catálogo de sucursales"
        Valida = False
        txtEstado.SetFocus
        Exit Function
    End If
    
    If txtCP.text = "" Then
        MsgBox "Introduzca el Cp !!", vbInformation, "Catálogo de sucursales"
        Valida = False
        txtCP.SetFocus
        Exit Function
    End If
End Function
