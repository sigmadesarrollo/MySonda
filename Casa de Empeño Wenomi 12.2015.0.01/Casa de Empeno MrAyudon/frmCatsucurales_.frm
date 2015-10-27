VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatsucursales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de sucursales"
   ClientHeight    =   4530
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
   Icon            =   "frmCatsucurales.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4530
   ScaleWidth      =   5685
   Begin VB.TextBox txtEmail 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2280
      MaxLength       =   80
      TabIndex        =   9
      Top             =   3480
      Width           =   3255
   End
   Begin VB.TextBox txtClave 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   120
      MaxLength       =   100
      TabIndex        =   0
      Top             =   480
      Width           =   735
   End
   Begin VB.TextBox txtCp 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4440
      MaxLength       =   5
      TabIndex        =   8
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox txtRfc 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   120
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1680
      Width           =   2055
   End
   Begin VB.TextBox txtTelefono 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   25
      TabIndex        =   7
      Top             =   3480
      Width           =   2055
   End
   Begin VB.TextBox txtEstado 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   6
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtCiudad 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   120
      MaxLength       =   60
      TabIndex        =   5
      Top             =   2880
      Width           =   2055
   End
   Begin VB.TextBox txtDireccion 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   120
      MaxLength       =   100
      TabIndex        =   4
      Top             =   2280
      Width           =   5415
   End
   Begin VB.TextBox txtRazonsocial 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   960
      MaxLength       =   100
      TabIndex        =   1
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox txtNomcomercial 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   120
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1080
      Width           =   5415
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosCliente2 
      Height          =   255
      Left            =   5175
      TabIndex        =   18
      Top             =   465
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
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   3150
      TabIndex        =   19
      Top             =   3975
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
      Left            =   4230
      TabIndex        =   21
      Top             =   3975
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
      Left            =   1965
      TabIndex        =   22
      Top             =   3975
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
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Correo Electrónico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2295
      TabIndex        =   23
      Top             =   3240
      Width           =   3240
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Clave"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
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
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   4440
      TabIndex        =   17
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "RFC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   135
      TabIndex        =   16
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   135
      TabIndex        =   15
      Top             =   3240
      Width           =   2070
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   2295
      TabIndex        =   14
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Ciudad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   135
      TabIndex        =   13
      Top             =   2640
      Width           =   2055
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   2040
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Razón Social"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   975
      TabIndex        =   11
      Top             =   240
      Width           =   4200
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00404040&
      Caption         =   "Nombre Comercial"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   840
      Width           =   5415
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
    If txtRazonSocial.Tag <> "" Then
        If MsgBox("Desea guardar los cambios ??", vbQuestion + vbYesNo + vbDefaultButton2, "Catálogo de sucursales") = vbYes Then
            Sql = "update Sucursales set Clave=" & Val(txtClave.text) & ",Razonsocial='" & Trim(txtRazonSocial.text) & "',Nombrecomercial='" & Trim(txtNomcomercial.text) & "',Direccion='" & Trim(txtDireccion.text) & "',Rfc='" & Trim(txtRfc.text) & "',Ciudad='" & Trim(txtCiudad.text) & "',Estado='" & Trim(txtEstado.text) & "',Telefono='" & Trim(txtTelefono.text) & "',Cp=" & Val(txtCP.text) & ",Email='" & Trim(txtEmail.text) & "' where ID=" & Val(txtRazonSocial.Tag) & ""
        Else
            GoTo 125
        End If
    Else
        Sql = "insert into Sucursales (Clave,Razonsocial,NombreComercial,Direccion,Rfc,Ciudad,Estado,Telefono,CP,Email)values(" & Val(txtClave.text) & ",'" & Trim(txtRazonSocial.text) & "','" & Trim(txtNomcomercial.text) & "','" & Trim(txtDireccion.text) & "','" & Trim(txtRfc.text) & "','" & Trim(txtCiudad.text) & "','" & Trim(txtEstado.text) & "','" & Trim(txtTelefono.text) & "','" & Trim(txtCP.text) & "','" & Trim(txtEmail.text) & "')"
    End If
    
    dbDatos.Execute Sql
125:
    Limpiar
    txtRazonSocial.SetFocus
End If
End Sub

Private Sub cmdLimpiar_Click()
Limpiar
txtClave.SetFocus
End Sub

Private Sub cmdMosCliente2_Click()
frmMostrarSucursales.Ver Me, txtRazonSocial, True, True
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
txtRazonSocial.text = ""
txtRazonSocial.Tag = ""
txtDireccion.text = ""
txtCiudad.text = ""
txtEstado.text = ""
txtTelefono.text = ""
txtCP.text = ""
txtRfc.text = ""
txtEmail.text = ""
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
Seleccionar_Texto txtRazonSocial
Cambiar_Color True, txtRazonSocial
End Sub

Private Sub txtRazonsocial_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtRazonsocial_LostFocus()
Cambiar_Color False, txtRazonSocial
End Sub

Private Sub txtRfc_GotFocus()
Seleccionar_Texto txtRfc
Cambiar_Color True, txtRfc
End Sub

Private Sub txtRfc_KeyPress(KeyAscii As Integer)
KeyAscii = Mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtRfc_LostFocus()
Cambiar_Color False, txtRfc
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

Public Function BuscarSucursal(IDSucursal As Long)
Dim rcConsulta As New ADODB.Recordset
On Error GoTo Error

Set rcConsulta = dbDatos.Execute("select * from sucursales where Id=" & IDSucursal & "")
If Not rcConsulta.BOF And Not rcConsulta.EOF Then
    With rcConsulta
        txtRazonSocial.Tag = !ID
        If Not IsNull(!RazonSocial) Then txtRazonSocial.text = !RazonSocial Else txtRazonSocial.text = ""
        If Not IsNull(!NombreComercial) Then txtNomcomercial.text = !NombreComercial Else txtNomcomercial.text = ""
        If Not IsNull(!RFC) Then txtRfc.text = !RFC Else txtRfc.text = ""
        If Not IsNull(!Direccion) Then txtDireccion.text = !Direccion Else txtDireccion.text = ""
        If Not IsNull(!Ciudad) Then txtCiudad.text = !Ciudad Else txtCiudad.text = ""
        If Not IsNull(!Estado) Then txtEstado.text = !Estado Else txtEstado.text = ""
        If Not IsNull(!Telefono) Then txtTelefono.text = !Telefono Else txtTelefono.text = ""
        If Not IsNull(!CP) Then txtCP.text = !CP Else txtCP.text = ""
        If Not IsNull(!Clave) Then txtClave.text = !Clave Else txtClave.text = ""
        If Not IsNull(!Email) Then txtEmail.text = !Email Else txtEmail.text = ""
    End With
End If

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function

Sub Inicializar()
Dim rc As ADODB.Recordset

On Error GoTo Error

Set rc = dbDatos.Execute("select * from sucursales where Activa=1")
If Not rc.BOF And Not rc.EOF Then
    With rc
        txtRazonSocial.Tag = rc!ID
        If Not IsNull(!RazonSocial) Then txtRazonSocial.text = !RazonSocial Else txtRazonSocial.text = ""
        If Not IsNull(!NombreComercial) Then txtNomcomercial.text = !NombreComercial Else txtNomcomercial.text = ""
        If Not IsNull(!RFC) Then txtRfc.text = !RFC Else txtRfc.text = ""
        If Not IsNull(!Direccion) Then txtDireccion.text = !Direccion Else txtDireccion.text = ""
        If Not IsNull(!Ciudad) Then txtCiudad.text = !Ciudad Else txtCiudad.text = ""
        If Not IsNull(!Estado) Then txtEstado.text = !Estado Else txtEstado.text = ""
        If Not IsNull(!Telefono) Then txtTelefono.text = !Telefono Else txtTelefono.text = ""
        If Not IsNull(!CP) Then txtCP.text = !CP Else txtCP.text = ""
        If Not IsNull(!Clave) Then txtClave.text = !Clave Else txtClave.text = ""
        If Not IsNull(!Email) Then txtEmail.text = !Email Else txtEmail.text = ""
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

If txtRazonSocial.text = "" Then
    MsgBox "Introduzca la razón social !!", vbInformation, "Catálogo de sucursales"
    Valida = False
    txtRazonSocial.SetFocus
    Exit Function
End If

If txtNomcomercial.text = "" Then
    MsgBox "Introduzca el nombre comercial !!", vbInformation, "Catálogo de sucursales"
    Valida = False
    txtNomcomercial.SetFocus
    Exit Function
End If

If txtRfc.text = "" Then
    MsgBox "Introduzca el Rfc !!", vbInformation, "Catálogo de sucursales"
    Valida = False
    txtRfc.SetFocus
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

'MLD-MODIF
Private Sub txtEmail_GotFocus()
    Seleccionar_Texto txtEmail
    Cambiar_Color True, txtEmail
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    KeyAscii = minusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEmail_LostFocus()
    Cambiar_Color False, txtEmail
End Sub
