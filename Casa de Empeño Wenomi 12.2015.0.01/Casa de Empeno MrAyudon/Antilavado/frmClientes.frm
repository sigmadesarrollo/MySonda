VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmClientes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de clientes"
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11040
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
   ScaleHeight     =   6810
   ScaleWidth      =   11040
   Begin VB.ComboBox txtEstado 
      Height          =   315
      ItemData        =   "frmClientes.frx":000C
      Left            =   7515
      List            =   "frmClientes.frx":000E
      TabIndex        =   29
      Text            =   "txtEstado"
      Top             =   5280
      Width           =   2250
   End
   Begin VB.ComboBox CmbTipoIdentificacion 
      Height          =   315
      ItemData        =   "frmClientes.frx":0010
      Left            =   240
      List            =   "frmClientes.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3450
      Width           =   4335
   End
   Begin VB.TextBox txtNumIdentificacion 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   8850
      MaxLength       =   30
      TabIndex        =   20
      Top             =   3480
      Width           =   1950
   End
   Begin VB.TextBox txtOtraIdentificacion 
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
      Left            =   4680
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   3480
      Width           =   3975
   End
   Begin VB.TextBox txtRFC 
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   240
      MaxLength       =   80
      TabIndex        =   21
      Top             =   4080
      Width           =   2070
   End
   Begin VB.ComboBox cmbFisicaMoral 
      Height          =   315
      ItemData        =   "frmClientes.frx":0014
      Left            =   1410
      List            =   "frmClientes.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2190
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   7560
      TabIndex        =   49
      Top             =   9960
      Visible         =   0   'False
      Width           =   3255
      Begin MSMask.MaskEdBox txtFechaExpiracion 
         Height          =   240
         Left            =   1560
         TabIndex        =   36
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   15
         Format          =   "dd-mmmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblFecExpiración 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fec. Expiración:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1470
      End
   End
   Begin VB.TextBox txtEmail 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   4800
      MaxLength       =   80
      TabIndex        =   23
      Top             =   4080
      Width           =   6045
   End
   Begin VB.TextBox txtNoExterior 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   9690
      MaxLength       =   30
      TabIndex        =   26
      Top             =   4680
      Width           =   1125
   End
   Begin VB.TextBox txtNoInterior 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   8370
      MaxLength       =   30
      TabIndex        =   25
      Top             =   4680
      Width           =   1125
   End
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
      Left            =   4845
      MaxLength       =   30
      TabIndex        =   44
      Top             =   11040
      Visible         =   0   'False
      Width           =   765
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
      Left            =   240
      MaxLength       =   120
      TabIndex        =   31
      Top             =   5880
      Width           =   10595
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
      Left            =   240
      MaxLength       =   30
      TabIndex        =   24
      Top             =   4680
      Width           =   7875
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
      Left            =   4620
      MaxLength       =   30
      TabIndex        =   28
      Top             =   5280
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
      Left            =   240
      MaxLength       =   30
      TabIndex        =   27
      Top             =   5280
      Width           =   4215
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
      Left            =   2415
      MaxLength       =   20
      TabIndex        =   22
      Top             =   4080
      Width           =   2250
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
      Left            =   9915
      MaxLength       =   5
      TabIndex        =   30
      Top             =   5280
      Width           =   915
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   6225
      TabIndex        =   34
      Top             =   6270
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
      Picture         =   "frmClientes.frx":0018
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   7320
      TabIndex        =   35
      Top             =   6270
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
      Picture         =   "frmClientes.frx":011C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
      Height          =   375
      Left            =   5040
      TabIndex        =   33
      Top             =   6270
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
      Picture         =   "frmClientes.frx":066E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdFoto 
      Height          =   375
      Left            =   3720
      TabIndex        =   32
      Top             =   6270
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Fotografia"
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
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmClientes.frx":0BC0
   End
   Begin VB.Frame FrameFisica 
      Caption         =   "DATOS DE LA PERSONA FISICA"
      Height          =   2535
      Left            =   120
      TabIndex        =   60
      Top             =   600
      Width           =   10815
      Begin VB.ComboBox CmbOcupaciones 
         Height          =   315
         ItemData        =   "frmClientes.frx":1095
         Left            =   2280
         List            =   "frmClientes.frx":1097
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2070
         Width           =   8415
      End
      Begin VB.TextBox txtCurp 
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   120
         MaxLength       =   80
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2070
         Width           =   2070
      End
      Begin VB.ComboBox CmbPaisNacionalidad 
         Height          =   315
         ItemData        =   "frmClientes.frx":1099
         Left            =   8250
         List            =   "frmClientes.frx":109B
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1170
         Width           =   2430
      End
      Begin VB.ComboBox CmbEstadoNac 
         Height          =   315
         ItemData        =   "frmClientes.frx":109D
         Left            =   6000
         List            =   "frmClientes.frx":109F
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1170
         Width           =   2190
      End
      Begin VB.ComboBox cmdPaisNacimiento 
         Height          =   315
         ItemData        =   "frmClientes.frx":10A1
         Left            =   3480
         List            =   "frmClientes.frx":10A3
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1170
         Width           =   2430
      End
      Begin VB.ComboBox CmbSexo 
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
         ItemData        =   "frmClientes.frx":10A5
         Left            =   120
         List            =   "frmClientes.frx":10AF
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1170
         Width           =   1485
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
         Left            =   120
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   3225
      End
      Begin VB.TextBox txtApellidoPaterno 
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
         Left            =   3480
         MaxLength       =   60
         TabIndex        =   2
         Top             =   480
         Width           =   3105
      End
      Begin VB.TextBox txtApellidoMaterno 
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
         Left            =   6720
         MaxLength       =   60
         TabIndex        =   3
         Top             =   480
         Width           =   3105
      End
      Begin MSMask.MaskEdBox txtFecNacimiento 
         Height          =   240
         Left            =   1710
         TabIndex        =   5
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
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente2 
         Height          =   240
         Left            =   10080
         TabIndex        =   72
         Top             =   360
         Width           =   510
         _ExtentX        =   900
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
      Begin VB.Image icoP_CURP 
         Height          =   240
         Left            =   120
         Picture         =   "frmClientes.frx":10C8
         Top             =   1800
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Label lblOcupación 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ocupación o Actividad del Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   2295
         TabIndex        =   71
         Top             =   1815
         Width           =   8370
      End
      Begin VB.Label lblCurp 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "CURP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   69
         ToolTipText     =   "Click Aquí para Generar CURP"
         Top             =   1815
         Width           =   2070
      End
      Begin VB.Label lblPaisNacionalidad 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pais Nacionalidad"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   8250
         TabIndex        =   68
         Top             =   960
         Width           =   2400
      End
      Begin VB.Label lblEstadoNac 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Estado Nacimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   1
         Left            =   6000
         TabIndex        =   67
         Top             =   960
         Width           =   2175
      End
      Begin VB.Label lblPaisNac 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Pais Nacimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   3480
         TabIndex        =   66
         Top             =   960
         Width           =   2385
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fec. Nacimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   23
         Left            =   1710
         TabIndex        =   65
         Top             =   960
         Width           =   1620
      End
      Begin VB.Label lblSexo 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Sexo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   64
         Top             =   960
         Width           =   1485
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Nombre(s)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   63
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label lblApellidoPaterno1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apellido Paterno:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   3480
         TabIndex        =   62
         Top             =   240
         Width           =   3105
      End
      Begin VB.Label lblApellidoMaterno 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Apellido Materno:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6720
         TabIndex        =   61
         Top             =   240
         Width           =   3105
      End
   End
   Begin VB.Frame FrameMoral 
      Caption         =   "DATOS DE LA PERSONA MORAL"
      Height          =   2535
      Left            =   120
      TabIndex        =   52
      Top             =   600
      Visible         =   0   'False
      Width           =   10815
      Begin VB.Frame Frame3 
         Caption         =   "DATOS DEL REPRESENTANTE LEGAL"
         Height          =   1575
         Left            =   120
         TabIndex        =   55
         Top             =   840
         Width           =   10575
         Begin VB.TextBox txtRLCurp 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2280
            MaxLength       =   80
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   1205
            Width           =   2070
         End
         Begin VB.TextBox txtRLRFC 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   120
            MaxLength       =   80
            TabIndex        =   16
            Top             =   1205
            Width           =   2070
         End
         Begin VB.TextBox txtRLApellidoMaterno 
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
            Left            =   6600
            MaxLength       =   60
            TabIndex        =   15
            Top             =   600
            Width           =   3105
         End
         Begin VB.TextBox txtRLApellidoPaterno 
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
            Left            =   3360
            MaxLength       =   60
            TabIndex        =   14
            Top             =   600
            Width           =   3105
         End
         Begin VB.TextBox txtRLNombre 
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
            Left            =   120
            MaxLength       =   20
            TabIndex        =   13
            Top             =   600
            Width           =   3105
         End
         Begin VB.Label lblRLCURP 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "CURP"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   2280
            TabIndex        =   73
            Top             =   960
            Width           =   2070
         End
         Begin VB.Label lblRFC 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "RFC"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   59
            Top             =   960
            Width           =   2070
         End
         Begin VB.Label lblApellidoMaterno1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Materno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   6600
            TabIndex        =   58
            Top             =   360
            Width           =   3105
         End
         Begin VB.Label lblApellidoPaterno1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Paterno"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   2
            Left            =   3360
            TabIndex        =   57
            Top             =   360
            Width           =   3105
         End
         Begin VB.Label lblNombre 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Nombre"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   56
            Top             =   360
            Width           =   3105
         End
      End
      Begin VB.TextBox txtRazonSocial 
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
         Left            =   120
         MaxLength       =   60
         TabIndex        =   11
         Top             =   480
         Width           =   8415
      End
      Begin MSMask.MaskEdBox txtAltaRazonSocial 
         Height          =   240
         Left            =   8760
         TabIndex        =   12
         Top             =   480
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   423
         _Version        =   393216
         BorderStyle     =   0
         MaxLength       =   10
         Format          =   "dd-mmmm-yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblApellidoPaterno1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Razon Social"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   8415
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00FFC0C0&
         Caption         =   "Fecha Alta"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   0
         Left            =   8760
         TabIndex        =   53
         Top             =   240
         Width           =   1860
      End
   End
   Begin VB.Image icoP_RFC 
      Height          =   240
      Left            =   240
      Picture         =   "frmClientes.frx":140A
      Top             =   3840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Label lblId 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Tipo Identificación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   22
      Left            =   240
      TabIndex        =   76
      Top             =   3240
      Width           =   4305
   End
   Begin VB.Label lblNoIdentificación 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "N° Identificación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   27
      Left            =   8850
      TabIndex        =   75
      Top             =   3240
      Width           =   1950
   End
   Begin VB.Label lblId 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   "Otra Identificación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   4680
      TabIndex        =   74
      Top             =   3240
      Width           =   3990
   End
   Begin VB.Label lblRFC 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "RFC"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   255
      TabIndex        =   70
      ToolTipText     =   "Click Aquí para Generar RFC"
      Top             =   3840
      Width           =   2025
   End
   Begin VB.Label Label 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Persona"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   1
      Left            =   240
      TabIndex        =   51
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label lblEmail 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Correo Electrónico"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   9
      Left            =   4800
      TabIndex        =   48
      Top             =   3855
      Width           =   6015
   End
   Begin VB.Label lblNoExt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "N° Exterior"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   1
      Left            =   9690
      TabIndex        =   47
      Top             =   4440
      Width           =   1125
   End
   Begin VB.Label lblNoInt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "N° Interior"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   8370
      TabIndex        =   46
      Top             =   4440
      Width           =   1125
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
      Left            =   4245
      TabIndex        =   45
      Top             =   11040
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Label Label114 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Mensaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   43
      Top             =   5640
      Width           =   10590
   End
   Begin VB.Label Label64 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Dirección"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   42
      Top             =   4440
      Width           =   7875
   End
   Begin VB.Label Label65 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Municipio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   4620
      TabIndex        =   41
      Top             =   5040
      Width           =   2760
   End
   Begin VB.Label Label66 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Colonia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   240
      TabIndex        =   40
      Top             =   5040
      Width           =   4215
   End
   Begin VB.Label Label71 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   7515
      TabIndex        =   39
      Top             =   5040
      Width           =   2250
   End
   Begin VB.Label Label72 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "Teléfono"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   2415
      TabIndex        =   38
      Top             =   3840
      Width           =   2250
   End
   Begin VB.Label Label91 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      Caption         =   "C.P."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   9915
      TabIndex        =   37
      Top             =   5040
      Width           =   915
   End
End
Attribute VB_Name = "frmClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private bolMostrar As Boolean
Private Cliente As clientes

Private Sub cmbFisicaMoral_Click()
    If cmbFisicaMoral.ListIndex >= 0 Then
        If cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex) = 1 Then
            FrameFisica.Visible = True
            FrameMoral.Visible = False
'            txtNombre.Enabled = True
'            txtApellidoPaterno.Enabled = True
'            txtApellidoMaterno.Enabled = True
'            txtRazonSocial.Enabled = False
'            txtAltaRazonSocial.Enabled = False
'            txtRLNombre.Enabled = False
'            txtRLApellidoPaterno.Enabled = False
'            txtRLApellidoMaterno.Enabled = False
'            txtRLRFC.Enabled = False
        Else
            FrameMoral.Visible = True
            FrameFisica.Visible = False
'            txtNombre.Enabled = False
'            txtApellidoPaterno.Enabled = False
'            txtApellidoMaterno.Enabled = False
'            txtRazonSocial.Enabled = True
'            txtAltaRazonSocial.Enabled = True
'            txtRLNombre.Enabled = True
'            txtRLApellidoPaterno.Enabled = True
'            txtRLApellidoMaterno.Enabled = True
'            txtRLRFC.Enabled = True
        End If
    End If
End Sub

Private Sub cmbSexo_GotFocus()
    cmbSexo.BackColor = &HC0FFFF
End Sub

Private Sub cmbSexo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbSexo_LostFocus()
    cmbSexo.BackColor = vbWhite
    If Cliente.ID = 0 Then
        If cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex) = 1 Then
            txtCurp.text = GenerarCURP(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, cmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"), CmbEstadoNac.text)
            txtRfc.text = GeneraRFC(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"))
        End If
    End If
End Sub

Private Sub cmdAgregar_Click()
    Dim FechaNac As String, Sexo As Integer

    If Valida Then
        With Cliente
            If cmbFisicaMoral.ListIndex < 0 Then
                .FisicaMoral = 0
            Else
                .FisicaMoral = IIf(cmbFisicaMoral.ListIndex < 0, 0, cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex))
            End If
            .Nombre = txtNombre.text
            .ApellidoPaterno = txtApellidoPaterno.text
            .ApellidoMaterno = txtApellidoMaterno.text
            .RazonSocial = txtRazonsocial.text
            .FechaAltaRazonSocial = IIf(.FisicaMoral = 2, txtAltaRazonSocial.text, "1900/01/01")
            .FechaNacimiento = IIf(.FisicaMoral = 2, "1900/01/01", txtFecNacimiento.text)
            .Direccion = txtDireccion.text
            .NoExterior = txtNoExterior.text
            .NoInterior = txtNoInterior.text
            .Colonia = txtColonia.text
            .Municipio = txtMunicipio.text
            .Estado = txtEstado.text
            .CodigoPostal = txtCP.text
            .Telefono = txtTelefono.text
            .Email = txtEmail.text
            '.FechaNacimiento = txtFecNacimiento.text
            .Mensaje = txtMensaje.text
            .Curp = txtCurp.text
            .RFC = txtRfc.text
            If cmbSexo.ListIndex < 0 Then
                .Sexo = 0
            Else
                .Sexo = IIf(cmbSexo.ListIndex < 0, 0, cmbSexo.ItemData(cmbSexo.ListIndex))
            End If
            If CmbOcupaciones.ListIndex < 0 Then
                .IDOcupacion = 0
            Else
                .IDOcupacion = CmbOcupaciones.ItemData(CmbOcupaciones.ListIndex)
            End If
            If CmbTipoIdentificacion.ListIndex < 0 Then
                .IDTipoIdentificacion = 0
            Else
                .IDTipoIdentificacion = CmbTipoIdentificacion.ItemData(CmbTipoIdentificacion.ListIndex)
            End If
            .NumeroIdentificacion = txtNumIdentificacion.text
'            .FechaExpiracion = txtFechaExpiracion.text
            If CmbPaisNacionalidad.ListIndex < 0 Then
                .IDPaisNacionalidad = 0
            Else
                .IDPaisNacionalidad = CmbPaisNacionalidad.ItemData(CmbPaisNacionalidad.ListIndex)
            End If
            If CmbEstadoNac.ListIndex < 0 Then
                .IDEstadoNacimiento = 0
            Else
                .IDEstadoNacimiento = CmbEstadoNac.ItemData(CmbEstadoNac.ListIndex)
            End If
            If cmdPaisNacimiento.ListIndex < 0 Then
                .IDPaisNacimiento = 0
            Else
                .IDPaisNacimiento = cmdPaisNacimiento.ItemData(cmdPaisNacimiento.ListIndex)
            End If
            .RL_Nombre = txtRLNombre.text
            .RL_ApellidoPaterno = txtRLApellidoPaterno.text
            .RL_ApellidoMaterno = txtRLApellidoMaterno.text
            .RL_RFC = txtRLRFC.text
            .RL_Curp = txtRLCurp.text
            .DesIdentificacionOtro = txtOtraIdentificacion.text
        End With
        If bolMostrar = False Then
            If Cliente.Grabar Then
                Limpiar
                txtNombre.SetFocus
            End If
        Else
            bolMostrar = False
            Unload Me
        End If
                
    End If
End Sub

Private Sub cmdFoto_Click()
Dim strCliente As String

    If txtNombre.Tag = "" And (txtNombre.text = "" Or txtApellidoPaterno.text = "" Or txtApellidoMaterno.text = "") Then
        
        MsgBox "Seleccione un cliente o introduzca un nombre !!", vbInformation, "Empeño"
        txtNombre.SetFocus
        Exit Sub
    Else
        
        strCliente = txtNombre.text & " " & txtApellidoPaterno.text & " " & txtApellidoMaterno.text
    End If

    'Mando a llamar el formulario
    'frmCapturaImagen.Ver CInt(strCliente), 1
    frmCapturaImagenBiometrico.Ver txtNombre.Tag, strCliente, 1
End Sub

Private Sub cmdLimpiar_Click()
    Limpiar
    txtNombre.SetFocus
End Sub

Private Sub cmdMosCliente2_Click()
    frmMostrarCliente.Ver Me, txtNombre, True, 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentrarForm Me, frmMDI
    If Not bolMostrar Then Set Cliente = New clientes
    Inicializar
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> vbFormCode Then
        Cancel = bolMostrar
    End If
End Sub

Private Sub lblCurp_Click(Index As Integer)
    txtCurp.text = GenerarCURP(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, cmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"), CmbEstadoNac.text)
    If txtCurp.text <> "" And (Trim(txtNombre.text) <> "" And Trim(txtApellidoPaterno.text) <> "" And Trim(txtApellidoMaterno.text) <> "" And Trim(cmbSexo.text) <> "" And (Trim(txtFecNacimiento.text) <> "__/__/____" Or Trim(txtFecNacimiento.text) <> "") And Trim(CmbEstadoNac.text) <> "") Then
        icoP_CURP.Visible = False
    End If
End Sub

Private Sub lblRFC_Click(Index As Integer)
    txtRfc.text = GeneraRFC(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"))
    If txtRfc.text <> "" And (Trim(txtNombre.text) <> "" And Trim(txtApellidoPaterno.text) <> "" And Trim(txtApellidoMaterno.text) <> "" And Trim(cmbSexo.text) <> "" And (Trim(txtFecNacimiento.text) <> "__/__/____" Or Trim(txtFecNacimiento.text) <> "")) Then
        icoP_RFC.Visible = False
    End If
End Sub

Private Sub txtApellidoMaterno_GotFocus()
    Seleccionar_Texto txtApellidoMaterno
    Cambiar_Color True, txtApellidoMaterno
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtApellidoMaterno_LostFocus()
    Cambiar_Color False, txtApellidoMaterno
    If Cliente.ID = 0 Then
        txtCurp.text = GenerarCURP(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, cmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"), CmbEstadoNac.text)
        
        'If cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex) = 1 Then
        '    txtRFC.text = GeneraRFC(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, CmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"))
        'End If
        
    End If
    If Trim(txtNombre.text) <> "" And Trim(txtApellidoPaterno.text) <> "" And Trim(txtApellidoMaterno.text) <> "" And Val(txtNombre.Tag) <= 0 Then Mostrar_Seleccionar_Cliente Trim(txtNombre.text), Trim(txtApellidoPaterno.text), Trim(txtApellidoMaterno.text)
End Sub

Private Sub txtApellidoPaterno_GotFocus()
    Seleccionar_Texto txtApellidoPaterno
    Cambiar_Color True, txtApellidoPaterno
End Sub

Private Sub txtApellidoPaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtApellidoPaterno_LostFocus()
    Cambiar_Color False, txtApellidoPaterno
    If Cliente.ID = 0 Then
        txtCurp.text = GenerarCURP(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, cmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"), CmbEstadoNac.text)
        
        'If cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex) = 1 Then
        '    txtRFC.text = GeneraRFC(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, CmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"))
        'End If
        
    End If
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

Private Sub txtNumIdentificacion_GotFocus()
    Seleccionar_Texto txtNumIdentificacion
    Cambiar_Color True, txtNumIdentificacion
End Sub

Private Sub txtNumIdentificacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumIdentificacion_LostFocus()
    Cambiar_Color False, txtNumIdentificacion
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
    If Cliente.ID = 0 Then
        txtCurp.text = GenerarCURP(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, cmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"), CmbEstadoNac.text)
        
        'If cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex) = 1 Then
        '    txtRFC.text = GeneraRFC(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, CmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"))
        'End If
        
    End If
End Sub

Private Sub txtRLCurp_GotFocus()
    Seleccionar_Texto txtRLCurp
    Cambiar_Color True, txtRLCurp
End Sub

Private Sub txtRLCurp_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRLCurp_LostFocus()
    Cambiar_Color False, txtRLCurp
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
    txtNombre.Tag = ""
    txtNombre.text = ""
    txtApellidoPaterno.text = ""
    txtApellidoMaterno.text = ""
    txtDireccion.text = ""
    txtNoExterior.text = ""
    txtNoInterior.text = ""
    txtColonia.text = ""
    txtMunicipio.text = ""
    txtEstado.ListIndex = -1 'txtEstado.text = ""
    txtCP.text = ""
    txtTelefono.text = ""
    txtEmail.text = ""
    txtFecNacimiento.Mask = ""
    txtFecNacimiento.Mask = "##/##/####"
    txtFecNacimiento.text = CDate("1900/01/01")
    txtEdad.text = ""
    txtMensaje.text = ""
    txtCurp.text = ""
    txtRfc.text = ""
    cmbSexo.ListIndex = -1
    CmbOcupaciones.ListIndex = -1
    CmbTipoIdentificacion.ListIndex = -1
    txtNumIdentificacion.text = ""
    txtFechaExpiracion.Mask = ""
    txtFechaExpiracion.Mask = "##/##/####"
    txtFechaExpiracion.text = CDate("1900/01/01")
    CmbPaisNacionalidad.ListIndex = -1
    CmbEstadoNac.ListIndex = -1
    cmdPaisNacimiento.ListIndex = -1
    txtRazonsocial.text = ""
    cmbFisicaMoral.ListIndex = -1
    txtAltaRazonSocial.Mask = ""
    txtAltaRazonSocial.Mask = "##/##/####"
    txtAltaRazonSocial.text = CDate("1900/01/01")
    txtRLNombre.text = ""
    txtRLApellidoPaterno.text = ""
    txtRLApellidoMaterno.text = ""
    txtRLRFC.text = ""
    txtRLCurp.text = ""
    txtOtraIdentificacion.text = ""
    
    PonerDefault
    
    Cliente.Limpiar
End Sub

Function Valida() As Boolean
    If Not bolMostrar Then
        
        If cmbFisicaMoral.ListIndex = -1 Then
            Valida = False
            cmbFisicaMoral.SetFocus
            Exit Function
        End If
        
        If cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex) = 1 Then
            'si no tiene nombre
            If Trim(txtNombre.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtNombre.SetFocus
                Exit Function
            End If
              
            'si no tiene apellido
            If Trim(txtApellidoPaterno.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtApellidoPaterno.SetFocus
                Exit Function
            End If
            
            'si no tiene apellido
            If Trim(txtApellidoMaterno.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtApellidoMaterno.SetFocus
                Exit Function
            End If
            
            If Trim(txtFecNacimiento.text) = "" Or Trim(txtFecNacimiento.text) = "__/__/____" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtFecNacimiento.SetFocus
                Exit Function
            End If
        
            'si no tiene sexo
            If cmbSexo.ListIndex = -1 Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                cmbSexo.SetFocus
                Exit Function
            End If
            
            If CmbOcupaciones.ListIndex = -1 Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                CmbOcupaciones.SetFocus
                Exit Function
            End If
        
            If CmbPaisNacionalidad.ListIndex = -1 Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                CmbPaisNacionalidad.SetFocus
                Exit Function
            End If
            
            If CmbEstadoNac.ListIndex = -1 Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                CmbEstadoNac.SetFocus
                Exit Function
            End If
            
            If cmdPaisNacimiento.ListIndex = -1 Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                cmdPaisNacimiento.SetFocus
                Exit Function
            End If
            
            'si no tiene apellido
            If Trim(txtCurp.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtCurp.SetFocus
                Exit Function
            End If
        Else
            'si no tiene razon social
            If Trim(txtRazonsocial.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtRazonsocial.SetFocus
                Exit Function
            End If
            
            'si no tiene nombre
            If Trim(txtRLNombre.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtRLNombre.SetFocus
                Exit Function
            End If
              
            'si no tiene apellido
            If Trim(txtRLApellidoPaterno.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtRLApellidoPaterno.SetFocus
                Exit Function
            End If
            
            'si no tiene apellido
            If Trim(txtRLApellidoMaterno.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtRLApellidoMaterno.SetFocus
                Exit Function
            End If
        
            'si no tiene apellido
            If Trim(txtRLRFC.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtRLRFC.SetFocus
                Exit Function
            End If
            
            'si no tiene apellido
            If Trim(txtRLCurp.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtRLCurp.SetFocus
                Exit Function
            End If
            
            If Trim(txtAltaRazonSocial.text) = "" Or Trim(txtAltaRazonSocial.text) = "__/__/____" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtAltaRazonSocial.SetFocus
                Exit Function
            End If
        End If
        
        If CmbTipoIdentificacion.ListIndex = -1 Then
            MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Valida = False
            CmbTipoIdentificacion.SetFocus
            Exit Function
        End If
        
        If Trim(txtNumIdentificacion.text) = "" Then
            MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Valida = False
            txtNumIdentificacion.SetFocus
            Exit Function
        End If
        
        If CmbTipoIdentificacion.ItemData(CmbTipoIdentificacion.ListIndex) >= 11 And CmbTipoIdentificacion.ItemData(CmbTipoIdentificacion.ListIndex) <= 13 Then
            If Trim(txtOtraIdentificacion.text) = "" Then
                MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
                Valida = False
                txtOtraIdentificacion.SetFocus
                Exit Function
            End If
        End If
        
        If Trim(txtRfc.text) = "" Then
            MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Valida = False
            txtRfc.SetFocus
            Exit Function
        End If
        
        'si no tiene direccion
        If Trim(txtDireccion.text) = "" Then
            MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Valida = False
            txtDireccion.SetFocus
            Exit Function
        End If
        
        'si no tiene NoExterior
        If Trim(txtNoExterior.text) = "" Then
            MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Valida = False
            txtNoExterior.SetFocus
            Exit Function
        End If
        
        'si no tiene colonia
        If Trim(txtColonia.text) = "" Then
            MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Valida = False
            txtColonia.SetFocus
            Exit Function
        End If
        
        'si no tiene estado
        If Trim(txtEstado.text) = "" Then
            MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Valida = False
            txtEstado.SetFocus
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
        
        If Trim(txtFechaExpiracion.text) = "" Or Trim(txtFechaExpiracion.text) = "__/__/____" Then
        '    txtFechaExpiracion.text = "1900-01-01"
        End If
        
    End If

    Valida = True
End Function

Private Sub Inicializar()
    cmbSexo.Clear
    cmbSexo.AddItem "MASCULINO"
    cmbSexo.ItemData(cmbSexo.NewIndex) = 1
    cmbSexo.AddItem "FEMENINO"
    cmbSexo.ItemData(cmbSexo.NewIndex) = 2
    cmbFisicaMoral.Clear
    cmbFisicaMoral.AddItem "FISICA"
    cmbFisicaMoral.ItemData(cmbFisicaMoral.NewIndex) = 1
    cmbFisicaMoral.AddItem "MORAL"
    cmbFisicaMoral.ItemData(cmbFisicaMoral.NewIndex) = 2
    Cargar_Combos "Descripcion", "mld_actividades_economicas", CmbOcupaciones, , "Descripcion"
    Cargar_Combos "Descripcion", "mld_paises", CmbPaisNacionalidad, , "Descripcion"
    Cargar_Combos "Descripcion", "estadospais", CmbEstadoNac, , "Descripcion"
    Cargar_Combos "Descripcion", "mld_paises", cmdPaisNacimiento, , "Descripcion"
    Cargar_Combos "Descripcion", "mld_tipo_identificaciones", CmbTipoIdentificacion, , "Descripcion"
    
    Cargar_Combos "Descripcion", "estadospais", txtEstado, , "Descripcion"
    
    txtAltaRazonSocial.Mask = ""
    txtAltaRazonSocial.Mask = "##/##/####"
    txtAltaRazonSocial.text = CDate("1900/01/01")
    
    PonerDefault
End Sub

Private Sub PonerDefault()
    cmdPaisNacimiento.ListIndex = ComboInformacion(cmdPaisNacimiento, Val(SacaValor("mld_paises", "Id", " WHERE RegDefault=1")))
    CmbPaisNacionalidad.ListIndex = ComboInformacion(CmbPaisNacionalidad, Val(SacaValor("mld_paises", "Id", " WHERE RegDefault=1")))
    cmbFisicaMoral.ListIndex = ComboInformacion(cmbFisicaMoral, 1)
    CmbTipoIdentificacion.ListIndex = ComboInformacion(CmbTipoIdentificacion, Val(SacaValor("mld_tipo_identificaciones", "Id", " WHERE RegDefault=1")))
End Sub

Private Function Iniciales() As String
    Dim Cadena As String, Nombre As String, Apellidos As String
   
    Nombre = Trim(txtNombre.text)
    Apellidos = Trim(txtApellidoPaterno.text)
    
    Cadena = Mid(Nombre, 1, 1)
    If InStr(1, Nombre, " ") <> 0 Then Cadena = Cadena & Mid(Nombre, InStr(1, Nombre, " ") + 1, 1)
    
    Cadena = Cadena & Mid(Apellidos, 1, 1)
    If InStr(1, Apellidos, " ") <> 0 Then Cadena = Cadena & Mid(Apellidos, InStr(1, Apellidos, " ") + 1, 1)
       
    Iniciales = Cadena
End Function

Public Sub Buscar(ID As Long)
    If Cliente.Buscar(ID) = True Then
        MostrarDatos Cliente
    End If
End Sub

Private Sub txtFecNacimiento_GotFocus()
    Seleccionar_Texto txtFecNacimiento
    Cambiar_Color True, txtFecNacimiento
End Sub

Private Sub txtFecNacimiento_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFecNacimiento_LostFocus()
    Cambiar_Color False, txtFecNacimiento
    If Not IsDate(txtFecNacimiento.text) And Trim(txtFecNacimiento.text) <> "__/__/____" Then
        
        MsgBox "Introduzca una fecha válida !!", vbInformation, "Empeño"
        txtFecNacimiento.SetFocus
    ElseIf IsDate(txtFecNacimiento.text) Then
        
        txtEdad.text = Calcula_Edad(CDate(txtFecNacimiento.text))
        If Cliente.ID = 0 Then
            If cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex) = 1 Then
                txtCurp.text = GenerarCURP(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, cmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"), CmbEstadoNac.text)
                txtRfc.text = GeneraRFC(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"))
            End If
        End If
    End If
End Sub

'-------------------------------------------------------------------------------------------------------------------------------
Private Sub txtNoInterior_GotFocus()
    Seleccionar_Texto txtNoInterior
    Cambiar_Color True, txtNoInterior
End Sub

Private Sub txtNoInterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNoInterior_LostFocus()
    Cambiar_Color False, txtNoInterior
End Sub

Private Sub txtNoExterior_GotFocus()
    Seleccionar_Texto txtNoExterior
    Cambiar_Color True, txtNoExterior
End Sub

Private Sub txtNoExterior_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNoExterior_LostFocus()
    Cambiar_Color False, txtNoExterior
End Sub

Private Sub txtEmail_GotFocus()
    Seleccionar_Texto txtEmail
    Cambiar_Color True, txtEmail
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
    'KeyAscii = mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEmail_LostFocus()
    Cambiar_Color False, txtEmail
End Sub

Private Sub txtCurp_GotFocus()
    Seleccionar_Texto txtCurp
    Cambiar_Color True, txtCurp
End Sub

Private Sub txtCurp_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCurp_LostFocus()
    Cambiar_Color False, txtCurp
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

'txtFechaExpiracion
Private Sub txtFechaExpiracion_GotFocus()
    Seleccionar_Texto txtFechaExpiracion
    Cambiar_Color True, txtFechaExpiracion
End Sub

Private Sub txtFechaExpiracion_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaExpiracion_LostFocus()
    Cambiar_Color False, txtFechaExpiracion
    If Not IsDate(txtFechaExpiracion.text) And Trim(txtFechaExpiracion.text) <> "__/__/____" Then
        
        MsgBox "Introduzca una fecha válida !!", vbInformation, "Empeño"
        txtFechaExpiracion.SetFocus
    End If
End Sub

Private Sub CmbEstadoNac_GotFocus()
    CmbEstadoNac.BackColor = &HC0FFFF
End Sub

Private Sub CmbEstadoNac_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub CmbEstadoNac_LostFocus()
    CmbEstadoNac.BackColor = vbWhite
    If Cliente.ID = 0 Then
        If cmbFisicaMoral.ItemData(cmbFisicaMoral.ListIndex) = 1 Then
            txtCurp.text = GenerarCURP(txtNombre.text, txtApellidoPaterno.text, txtApellidoMaterno.text, cmbSexo.text, Format(IIf(Trim(txtFecNacimiento.text) = "__/__/____" Or Trim(txtFecNacimiento.text) = "", "1900-01-01", txtFecNacimiento.text), "YYYY-MM-DD"), CmbEstadoNac.text)
        End If
    End If
End Sub

Private Sub CmbOcupaciones_GotFocus()
    CmbOcupaciones.BackColor = &HC0FFFF
End Sub

Private Sub CmbOcupaciones_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub CmbOcupaciones_LostFocus()
    CmbOcupaciones.BackColor = vbWhite
End Sub

Private Sub CmbTipoIdentificacion_GotFocus()
    CmbTipoIdentificacion.BackColor = &HC0FFFF
End Sub

Private Sub CmbTipoIdentificacion_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub CmbTipoIdentificacion_LostFocus()
    CmbTipoIdentificacion.BackColor = vbWhite
End Sub

Private Sub CmbPaisNacionalidad_GotFocus()
    CmbPaisNacionalidad.BackColor = &HC0FFFF
End Sub

Private Sub CmbPaisNacionalidad_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub CmbPaisNacionalidad_LostFocus()
    CmbPaisNacionalidad.BackColor = vbWhite
End Sub

Private Sub cmdPaisNacimiento_GotFocus()
    cmdPaisNacimiento.BackColor = &HC0FFFF
End Sub

Private Sub cmdPaisNacimiento_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmdPaisNacimiento_LostFocus()
    cmdPaisNacimiento.BackColor = vbWhite
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

Private Sub cmbFisicaMoral_GotFocus()
    cmbFisicaMoral.BackColor = &HC0FFFF
End Sub

Private Sub cmbFisicaMoral_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbFisicaMoral_LostFocus()
    cmbFisicaMoral.BackColor = vbWhite
End Sub

'txtAltaRazonSocial
Private Sub txtAltaRazonSocial_GotFocus()
    Seleccionar_Texto txtAltaRazonSocial
    Cambiar_Color True, txtAltaRazonSocial
End Sub

Private Sub txtAltaRazonSocial_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAltaRazonSocial_LostFocus()
    Cambiar_Color False, txtAltaRazonSocial
    If Not IsDate(txtAltaRazonSocial.text) And Trim(txtAltaRazonSocial.text) <> "__/__/____" Then
        
        MsgBox "Introduzca una fecha válida !!", vbInformation, "Empeño"
        txtAltaRazonSocial.SetFocus
    End If
End Sub

Private Sub txtRLNombre_GotFocus()
    Seleccionar_Texto txtRLNombre
    Cambiar_Color True, txtRLNombre
End Sub

Private Sub txtRLNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRLNombre_LostFocus()
    Cambiar_Color False, txtRLNombre
End Sub

Private Sub txtRLApellidoMaterno_GotFocus()
    Seleccionar_Texto txtRLApellidoMaterno
    Cambiar_Color True, txtRLApellidoMaterno
End Sub

Private Sub txtRLApellidoMaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRLApellidoMaterno_LostFocus()
    Cambiar_Color False, txtRLApellidoMaterno
End Sub

Private Sub txtRLApellidoPaterno_GotFocus()
    Seleccionar_Texto txtRLApellidoPaterno
    Cambiar_Color True, txtRLApellidoPaterno
End Sub

Private Sub txtRLApellidoPaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRLApellidoPaterno_LostFocus()
    Cambiar_Color False, txtRLApellidoPaterno
End Sub

Private Sub txtRLRfc_GotFocus()
    Seleccionar_Texto txtRLRFC
    Cambiar_Color True, txtRLRFC
End Sub

Private Sub txtRLRfc_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRLRfc_LostFocus()
    Cambiar_Color False, txtRLRFC
End Sub

Private Sub txtOtraIdentificacion_GotFocus()
    Seleccionar_Texto txtOtraIdentificacion
    Cambiar_Color True, txtOtraIdentificacion
End Sub

Private Sub txtOtraIdentificacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtOtraIdentificacion_LostFocus()
    Cambiar_Color False, txtOtraIdentificacion
End Sub

Public Sub Mostrar(ByRef DatosCliente As clientes)
    bolMostrar = True
    cmdAgregar.Caption = " Aceptar"
    cmdLimpiar.Visible = False
    cmdSalir.Visible = False
    cmdFoto.Visible = False
    cmdMosCliente2.Visible = False
    Inicializar
    Set Cliente = DatosCliente
    
    MostrarDatos DatosCliente
    
    If Cliente.ID = 0 Then PonerDefault
    
    Me.Show vbModal
End Sub

'MLD-MODIF.
Private Sub Mostrar_Seleccionar_Cliente(ByVal sNombre As String, ByVal sApellidoPaterno As String, ByVal sApellidoMaterno As String)
   Dim Seleccionar As New frmSeleccionarClientes
   If Val(SacaValor("Clientes", "COUNT(ID)", " WHERE Nombre LIKE '%" & sNombre & "%' AND Apellido LIKE '%" & Trim(sApellidoPaterno & " " & sApellidoMaterno) & "%'")) > 0 Then
      Seleccionar.Nombre = sNombre
      Seleccionar.Apellido = Trim(sApellidoPaterno & " " & sApellidoMaterno)
      Seleccionar.Show vbModal, frmMDI
      If Seleccionar.IDCliente <> 0 Then Buscar Seleccionar.IDCliente
      Unload Seleccionar
   End If
End Sub

Private Sub MostrarDatos(ByVal Datos As clientes)
    With Datos
        txtNombre.Tag = .ID
        cmbFisicaMoral.ListIndex = ComboInformacion(cmbFisicaMoral, .FisicaMoral)
        txtNombre.text = .Nombre
        txtApellidoPaterno.text = .ApellidoPaterno
        txtApellidoMaterno.text = .ApellidoMaterno
        txtRazonsocial.text = .RazonSocial
        txtAltaRazonSocial.text = .FechaAltaRazonSocial
        txtDireccion.text = .Direccion
        txtNoExterior.text = .NoExterior
        txtNoInterior.text = .NoInterior
        txtColonia.text = .Colonia
        txtMunicipio.text = .Municipio
        'txtEstado.text = .Estado
        If .Estado = "" Then txtEstado.ListIndex = -1 Else txtEstado.text = .Estado
        txtCP.text = .CodigoPostal
        txtTelefono.text = .Telefono
        txtEmail.text = .Email
        txtFecNacimiento.text = .FechaNacimiento
        txtEdad.text = Calcula_Edad(.FechaNacimiento)
        txtMensaje.text = .Mensaje
        txtCurp.text = .Curp
        txtRfc.text = .RFC
        cmbSexo.ListIndex = ComboInformacion(cmbSexo, .Sexo)
        CmbOcupaciones.ListIndex = ComboInformacion(CmbOcupaciones, .IDOcupacion)
        CmbTipoIdentificacion.ListIndex = ComboInformacion(CmbTipoIdentificacion, .IDTipoIdentificacion)
        txtNumIdentificacion.text = .NumeroIdentificacion
'        txtFechaExpiracion.text = .FechaExpiracion
        CmbPaisNacionalidad.ListIndex = ComboInformacion(CmbPaisNacionalidad, IIf(.IDPaisNacionalidad = -1, Val(SacaValor("mld_paises", "Id", " WHERE RegDefault=1")), .IDPaisNacionalidad))
        CmbEstadoNac.ListIndex = ComboInformacion(CmbEstadoNac, .IDEstadoNacimiento)
        cmdPaisNacimiento.ListIndex = ComboInformacion(cmdPaisNacimiento, IIf(.IDPaisNacimiento = -1, Val(SacaValor("mld_paises", "Id", " WHERE RegDefault=1")), .IDPaisNacimiento))
        txtRLNombre.text = .RL_Nombre
        txtRLApellidoPaterno.text = .RL_ApellidoPaterno
        txtRLApellidoMaterno.text = .RL_ApellidoMaterno
        txtRLRFC.text = .RL_RFC
        txtOtraIdentificacion.text = .DesIdentificacionOtro
        txtRLCurp.text = .RL_Curp
        
        If Datos.ID > 0 Then
            If Trim(.Curp) = "" Then icoP_CURP.Visible = True
            If Trim(.RFC) = "" Then icoP_RFC.Visible = True
        End If
        
    End With
End Sub
