VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de clientes"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2970
   ScaleWidth      =   9510
   Begin VB.TextBox txtEdad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4725
      MaxLength       =   30
      TabIndex        =   26
      Top             =   1800
      Width           =   765
   End
   Begin VB.TextBox txtFecNacimiento 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   2220
      MaxLength       =   30
      TabIndex        =   25
      Top             =   1800
      Width           =   1470
   End
   Begin VB.TextBox txtMensaje 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1065
      MaxLength       =   120
      TabIndex        =   9
      Top             =   2160
      Width           =   8325
   End
   Begin VB.ComboBox cmbSexo 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmClientes.frx":000C
      Left            =   6465
      List            =   "frmClientes.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   1755
      Width           =   2925
   End
   Begin VB.TextBox txtApellidos 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5760
      MaxLength       =   60
      TabIndex        =   1
      Top             =   360
      Width           =   3630
   End
   Begin VB.TextBox txtNombre 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1080
      MaxLength       =   20
      TabIndex        =   0
      Top             =   360
      Width           =   3105
   End
   Begin VB.TextBox txtDireccion 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1140
      MaxLength       =   30
      TabIndex        =   2
      Top             =   735
      Width           =   8250
   End
   Begin VB.TextBox txtMunicipio 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   5220
      MaxLength       =   30
      TabIndex        =   4
      Top             =   1080
      Width           =   2760
   End
   Begin VB.TextBox txtColonia 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   555
      MaxLength       =   30
      TabIndex        =   3
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtEstado 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   915
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1440
      Width           =   2250
   End
   Begin VB.TextBox txtTelefono 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   4215
      MaxLength       =   20
      TabIndex        =   7
      Top             =   1440
      Width           =   1650
   End
   Begin VB.TextBox txtIdentificacion 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   7470
      MaxLength       =   30
      TabIndex        =   8
      Top             =   1440
      Width           =   1920
   End
   Begin VB.TextBox txtCP 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   8475
      MaxLength       =   5
      TabIndex        =   5
      Top             =   1080
      Width           =   915
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   7155
      TabIndex        =   11
      Top             =   2520
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
      Picture         =   "frmClientes.frx":002F
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosCliente2 
      Height          =   240
      Left            =   4230
      TabIndex        =   21
      Top             =   375
      Width           =   390
      _ExtentX        =   688
      _ExtentY        =   423
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
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   3720
      TabIndex        =   27
      Top             =   1755
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
      Picture         =   "frmClientes.frx":0133
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8250
      TabIndex        =   29
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
      Picture         =   "frmClientes.frx":0248
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   6000
      TabIndex        =   30
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Aceptar"
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
      Picture         =   "frmClientes.frx":079A
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Edad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4125
      TabIndex        =   28
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label114 
      AutoSize        =   -1  'True
      Caption         =   "Mensaje:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   24
      Top             =   2160
      Width           =   885
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de nacimiento:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   23
      Top             =   1800
      Width           =   2040
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Sexo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5880
      TabIndex        =   22
      Top             =   1800
      Width           =   540
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4710
      TabIndex        =   20
      Top             =   360
      Width           =   945
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   19
      Top             =   360
      Width           =   810
   End
   Begin VB.Label Label64 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   18
      Top             =   720
      Width           =   960
   End
   Begin VB.Label Label65 
      AutoSize        =   -1  'True
      Caption         =   "Municipio:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4200
      TabIndex        =   17
      Top             =   1080
      Width           =   960
   End
   Begin VB.Label Label66 
      AutoSize        =   -1  'True
      Caption         =   "Col:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   360
   End
   Begin VB.Label Label71 
      AutoSize        =   -1  'True
      Caption         =   "Estado:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label72 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3270
      TabIndex        =   14
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label Label73 
      AutoSize        =   -1  'True
      Caption         =   "Identificación:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   6000
      TabIndex        =   13
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label Label91 
      AutoSize        =   -1  'True
      Caption         =   "Cp:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   8115
      TabIndex        =   12
      Top             =   1080
      Width           =   315
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbSexo_GotFocus()
    Cambiar_Color True, cmbSexo
End Sub

Private Sub cmbSexo_LostFocus()
    Cambiar_Color False, cmbSexo
End Sub

Private Sub cmdAceptar_Click()
Dim FechaNac As String, Sexo As Integer

    If Valida Then
        
        If Trim(txtFecNacimiento.text) = "" Then
            
            FechaNac = "Null"
        Else
            
            FechaNac = "'" & Format(txtFecNacimiento.text, "YYYY/MM/DD") & "'"
        End If
            
        If cmbSexo.ListIndex = -1 Then
            
            Sexo = 0
        Else
            
            Sexo = cmbSexo.ItemData(cmbSexo.ListIndex)
        End If
            
        If txtNombre.Tag = "" Then
            
            dbDatos.Execute "INSERT INTO clientes (Iniciales,Nombre,Apellido,Direccion,Colonia,Municipio,Estado,Tel,Identificacion,CP,FecNac,Sexo,Notas,FecRegistro) VALUES ('" & _
                            Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "','" & Trim(txtNombre.text) & "','" & Trim(txtApellidos.text) & "','" & Trim(txtDireccion.text) & "','" & Trim(txtColonia.text) & "','" & Trim(txtMunicipio.text) & "','" & Trim(txtEstado.text) & "','" & Trim(txtTelefono.text) & "','" & Trim(txtIdentificacion.text) & "','" & Trim(txtCP.text) & "'," & FechaNac & "," & Sexo & ",'" & Trim(txtMensaje.text) & "','" & Format(Date, "YYYY/MM/DD") & "')"
        
        Else
            
            If MsgBox("Desea guardar los cambios realizados ??", vbQuestion + vbYesNo + vbDefaultButton1, "Clientes") = vbYes Then
                    
                dbDatos.Execute "UPDATE clientes SET Iniciales='" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "',Nombre='" & Trim(txtNombre.text) & "',Apellido='" & Trim(txtApellidos.text) & "',Direccion='" & txtDireccion.text & "',Colonia='" & txtColonia.text & "',Municipio='" & txtMunicipio.text & "'," & _
                                "Estado='" & txtEstado.text & "',Tel='" & txtTelefono.text & "',Identificacion='" & txtIdentificacion.text & "',CP='" & txtCP.text & "',FecNac=" & FechaNac & ",Sexo=" & Sexo & ",Notas='" & Trim(txtMensaje.text) & "' WHERE ID = " & Val(txtNombre.Tag)
            End If
            
        End If
        
        Limpiar
        txtNombre.SetFocus
    
    End If

End Sub

Private Sub cmdLimpiar_Click()
    Limpiar
    txtNombre.SetFocus
End Sub

Private Sub cmdMosCliente2_Click()
    frmMostrarCliente.ver Me, txtNombre, True, 0
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
Dim Fecha As String
    
    Fecha = IIf(txtFecNacimiento.text = "", Date, txtFecNacimiento.text)
    Fecha = frmCalendario.Fecha(Trim(Fecha))
    
    If Fecha <> "" Then
        
        txtFecNacimiento.text = Format(Fecha, "DD/MMM/YYYY")
        txtEdad.text = Calcula_Edad(CDate(Fecha))
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtApellidos_GotFocus()
    Seleccionar_Texto txtApellidos
    Cambiar_Color True, txtApellidos
End Sub

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtApellidos_LostFocus()
    Cambiar_Color False, txtApellidos
End Sub

Private Sub txtColonia_GotFocus()
    Seleccionar_Texto txtColonia
    Cambiar_Color True, txtColonia
End Sub

Private Sub txtColonia_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtColonia_LostFocus()
    Cambiar_Color False, txtColonia
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

Private Sub txtIdentificacion_GotFocus()
    Seleccionar_Texto txtIdentificacion
    Cambiar_Color True, txtIdentificacion
End Sub

Private Sub txtIdentificacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtIdentificacion_LostFocus()
    Cambiar_Color False, txtIdentificacion
End Sub

Private Sub txtMensaje_GotFocus()
    Seleccionar_Texto txtMensaje
    Cambiar_Color True, txtMensaje
End Sub

Private Sub txtMensaje_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMensaje_LostFocus()
    Cambiar_Color False, txtMensaje
End Sub

Private Sub txtMunicipio_GotFocus()
    Seleccionar_Texto txtMunicipio
    Cambiar_Color True, txtMunicipio
End Sub

Private Sub txtMunicipio_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMunicipio_LostFocus()
    Cambiar_Color False, txtMunicipio
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Texto txtNombre
    Cambiar_Color True, txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNombre_LostFocus()
    Cambiar_Color False, txtNombre
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

Sub Limpiar()
    txtNombre.text = ""
    txtNombre.Tag = ""
    txtApellidos.text = ""
    txtDireccion.text = ""
    txtColonia.text = ""
    txtMunicipio.text = ""
    txtCP.text = ""
    txtEstado.text = ""
    txtTelefono.text = ""
    txtIdentificacion.text = ""
    txtMensaje.text = ""
    txtFecNacimiento.text = ""
    txtEdad.text = ""
    cmbSexo.ListIndex = -1
End Sub

Function Valida() As Boolean
    
    Valida = True
    
    'si no tiene nombre
    If Trim(txtNombre.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Valida = False
        txtNombre.SetFocus
        Exit Function
    End If
      
    'si no tiene apellido
    If Trim(txtApellidos.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Valida = False
        txtApellidos.SetFocus
        Exit Function
    End If
    
    'si no tiene direccion
    If Trim(txtDireccion.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Valida = False
        txtDireccion.SetFocus
        Exit Function
    End If
    
    'si no tiene estado
    If Trim(txtEstado.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Valida = False
        txtEstado.SetFocus
        Exit Function
    End If
    
    'si no tiene colonia
    If Trim(txtColonia.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Valida = False
        txtColonia.SetFocus
        Exit Function
    End If
    
    'si no tiene municipio
    If Trim(txtMunicipio.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Valida = False
        txtMunicipio.SetFocus
        Exit Function
    End If
    
    'si no tiene cp
    If Trim(txtCP.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Valida = False
        txtCP.SetFocus
        Exit Function
    End If
    
    'si no identificacion
    If Trim(txtIdentificacion.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Valida = False
        txtIdentificacion.SetFocus
        Exit Function
    End If

'''''    'Fecha de nacimiento
'''''    If Trim(txtDia.text) = "" Then
'''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'''''        Valida = False
'''''        txtDia.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If Trim(txtMes.text) = "" Then
'''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'''''        Valida = False
'''''        txtMes.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If Trim(txtYear.text) = "" Then
'''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'''''        Valida = False
'''''        txtYear.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    'si no tiene sexo
'''''    If cmbSexo.ListIndex = -1 Then
'''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'''''        Valida = False
'''''        cmbSexo.SetFocus
'''''        Exit Function
'''''    End If

End Function

Public Sub Buscar_Cliente(ID As Long)
Dim rcClientes As New ADODB.Recordset
   
On Error GoTo error

    rcClientes.Open "SELECT * FROM clientes WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
    With rcClientes
        txtNombre.text = !Nombre
        txtApellidos.text = !apellido
        txtDireccion.text = !Direccion
        txtColonia.text = IIf(IsNull(!Colonia), "", !Colonia)
        txtMunicipio.text = IIf(IsNull(!Municipio), "", !Municipio)
        txtEstado.text = IIf(IsNull(!Estado), "", !Estado)
        txtTelefono.text = IIf(IsNull(!Tel), "", !Tel)
        txtIdentificacion.text = IIf(IsNull(!identificacion), "", !identificacion)
        txtCP.text = IIf(IsNull(!CP), "", !CP)
        txtMensaje.text = IIf(IsNull(!notas), "", !notas)
        txtFecNacimiento.text = IIf(IsNull(!FecNac), "", Format(!FecNac, "DD/MMM/YYYY"))
        If IsNull(!FecNac) Then
            txtEdad.text = ""
        Else
            txtEdad.text = Calcula_Edad(!FecNac)
        End If
        cmbSexo.ListIndex = ComboInformacion(cmbSexo, IIf(IsNull(!Sexo), -1, !Sexo))
        txtNombre.Tag = ID
    End With
    rcClientes.Close
    Set rcClientes = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcClientes = Nothing
End Sub

Private Sub txtFecNacimiento_GotFocus()
    Seleccionar_Texto txtFecNacimiento
    Cambiar_Color True, txtFecNacimiento
End Sub

Private Sub txtFecNacimiento_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtFecNacimiento_LostFocus()
    Cambiar_Color False, txtFecNacimiento
End Sub
