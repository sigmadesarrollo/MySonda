VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmTarjetaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarjeta Beneficio"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTarjetaCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   6060
   Begin VB.TextBox txtVencimiento 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4200
      Locked          =   -1  'True
      MaxLength       =   50
      TabIndex        =   25
      Top             =   6600
      Width           =   1515
   End
   Begin VB.TextBox txtCedula 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   24
      Top             =   6600
      Width           =   1845
   End
   Begin VB.TextBox txtCartilla 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   50
      TabIndex        =   23
      Top             =   6600
      Width           =   2085
   End
   Begin VB.TextBox txtPasaporte 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   22
      Top             =   6120
      Width           =   1725
   End
   Begin VB.TextBox txtIfe 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   21
      Top             =   6120
      Width           =   1845
   End
   Begin VB.TextBox txtTelTrabajo 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   35
      TabIndex        =   20
      Top             =   6120
      Width           =   2085
   End
   Begin VB.TextBox txtTrabajo 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   4200
      MaxLength       =   50
      TabIndex        =   19
      Top             =   5640
      Width           =   1710
   End
   Begin VB.TextBox txtIngreso 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2280
      MaxLength       =   50
      TabIndex        =   18
      Top             =   5640
      Width           =   1845
   End
   Begin VB.TextBox txtAntiguedadTra 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   50
      TabIndex        =   17
      Top             =   5640
      Width           =   2085
   End
   Begin VB.TextBox txtPuesto 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3480
      MaxLength       =   80
      TabIndex        =   16
      Top             =   5160
      Width           =   2445
   End
   Begin VB.TextBox txtGiro 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   80
      TabIndex        =   15
      Top             =   5160
      Width           =   3165
   End
   Begin VB.TextBox txtDomicilioEmpresa 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   100
      TabIndex        =   14
      Top             =   4680
      Width           =   5805
   End
   Begin VB.TextBox txtEmpresa 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   100
      TabIndex        =   13
      Top             =   4200
      Width           =   5805
   End
   Begin VB.TextBox txtSituacion 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3240
      MaxLength       =   50
      TabIndex        =   12
      Top             =   3720
      Width           =   2670
   End
   Begin VB.TextBox txtDueño 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   80
      TabIndex        =   8
      Top             =   2760
      Width           =   5715
   End
   Begin VB.TextBox txtAntiguedad 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   50
      TabIndex        =   11
      Top             =   3720
      Width           =   2880
   End
   Begin VB.TextBox txtPersonas 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3240
      MaxLength       =   30
      TabIndex        =   10
      Top             =   3240
      Width           =   2655
   End
   Begin VB.TextBox txtRenta 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   50
      TabIndex        =   9
      Top             =   3240
      Width           =   2880
   End
   Begin VB.TextBox txtEmail 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   50
      TabIndex        =   5
      Top             =   2280
      Width           =   2280
   End
   Begin VB.ComboBox cmbEstadoCivil 
      Height          =   315
      ItemData        =   "frmTarjetaCliente.frx":000C
      Left            =   2460
      List            =   "frmTarjetaCliente.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2250
      Width           =   1410
   End
   Begin VB.TextBox txtTelefono 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3960
      MaxLength       =   35
      TabIndex        =   7
      Top             =   2280
      Width           =   1860
   End
   Begin VB.TextBox txtLugarNac 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   3000
      MaxLength       =   50
      TabIndex        =   4
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txtConyuge 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      MaxLength       =   50
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin VB.TextBox txtSucursal 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.TextBox txtCliente 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   80
      TabIndex        =   1
      Top             =   840
      Width           =   4335
   End
   Begin VB.TextBox txtNumTarjeta 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1440
      MaxLength       =   12
      TabIndex        =   0
      Top             =   270
      Width           =   1215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdGuardar 
      Height          =   375
      Left            =   2520
      TabIndex        =   53
      Top             =   6990
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Guardar"
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
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmTarjetaCliente.frx":0010
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4845
      TabIndex        =   54
      Top             =   6990
      Width           =   1035
      _ExtentX        =   1826
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
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmTarjetaCliente.frx":00A0
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   3705
      TabIndex        =   55
      Top             =   6990
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
      Picture         =   "frmTarjetaCliente.frx":0131
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosusuario 
      Height          =   225
      Left            =   4470
      TabIndex        =   56
      Top             =   825
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   397
      AlignCaption    =   4
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
   Begin DevPowerFlatBttn.FlatBttn cmdMuestraSucursal 
      Height          =   225
      Left            =   4485
      TabIndex        =   57
      Top             =   1305
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   397
      AlignCaption    =   4
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
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   5715
      TabIndex        =   59
      Top             =   6525
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
      Picture         =   "frmTarjetaCliente.frx":0235
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "Vencimiento:"
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
      Left            =   4200
      TabIndex        =   58
      Top             =   6360
      Width           =   1185
   End
   Begin VB.Label lblFecha 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4080
      TabIndex        =   52
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label26 
      AutoSize        =   -1  'True
      Caption         =   "Cédula:"
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
      Left            =   2280
      TabIndex        =   51
      Top             =   6360
      Width           =   675
   End
   Begin VB.Label Label25 
      AutoSize        =   -1  'True
      Caption         =   "Cartilla:"
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
      Left            =   120
      TabIndex        =   50
      Top             =   6360
      Width           =   690
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "Pasaporte:"
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
      Left            =   4200
      TabIndex        =   49
      Top             =   5880
      Width           =   990
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "IFE:"
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
      Left            =   2280
      TabIndex        =   48
      Top             =   5880
      Width           =   330
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "Tel. Trabajo Ant.:"
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
      Left            =   120
      TabIndex        =   47
      Top             =   5880
      Width           =   1590
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Trabajo anterior:"
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
      Left            =   4200
      TabIndex        =   46
      Top             =   5400
      Width           =   1545
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Ingreso:"
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
      Left            =   2280
      TabIndex        =   45
      Top             =   5400
      Width           =   765
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Antiguedad:"
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
      Left            =   120
      TabIndex        =   44
      Top             =   5400
      Width           =   1140
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Puesto:"
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
      Left            =   3480
      TabIndex        =   43
      Top             =   4920
      Width           =   705
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Giro Empresa:"
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
      Left            =   120
      TabIndex        =   42
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "Domicilio Empresa:"
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
      Left            =   120
      TabIndex        =   41
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "Empresa:"
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
      Left            =   120
      TabIndex        =   40
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "Situación Vivienda:"
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
      Left            =   3240
      TabIndex        =   39
      Top             =   3480
      Width           =   1740
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Dueño o Banco:"
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
      Left            =   120
      TabIndex        =   38
      Top             =   2520
      Width           =   1455
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "Antiguedad Dom.:"
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
      Left            =   120
      TabIndex        =   37
      Top             =   3480
      Width           =   1680
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Personas Dep.:"
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
      Left            =   3240
      TabIndex        =   36
      Top             =   3000
      Width           =   1365
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Renta Hip.:"
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
      Left            =   120
      TabIndex        =   35
      Top             =   3000
      Width           =   1035
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Telefono(s):"
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
      Left            =   3960
      TabIndex        =   34
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Estado Civil:"
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
      Left            =   2490
      TabIndex        =   33
      Top             =   2040
      Width           =   1110
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Email:"
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
      Left            =   120
      TabIndex        =   32
      Top             =   2040
      Width           =   525
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "L. Nacimiento:"
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
      Left            =   3000
      TabIndex        =   31
      Top             =   1560
      Width           =   1290
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Cónyuge:"
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
      Left            =   120
      TabIndex        =   30
      Top             =   1560
      Width           =   870
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal:"
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
      Left            =   120
      TabIndex        =   29
      Top             =   1080
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   3360
      TabIndex        =   28
      Top             =   240
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cliente:"
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
      Left            =   120
      TabIndex        =   27
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Num. Tarjeta:"
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
      Left            =   120
      TabIndex        =   26
      Top             =   240
      Width           =   1230
   End
End
Attribute VB_Name = "frmTarjetaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbEstadoCivil_GotFocus()
Cambiar_Color True, cmbEstadoCivil
End Sub

Private Sub cmbEstadoCivil_KeyPress(KeyAscii As Integer)
Pasar_Foco KeyAscii
End Sub

Private Sub cmbEstadoCivil_LostFocus()
Cambiar_Color False, cmbEstadoCivil
End Sub

Private Sub cmdGuardar_Click()
Dim sql As String, EstadoCivil As Integer
Dim rcConsulta As New ADODB.Recordset

On Error GoTo error

If ValidaDatos Then
    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Tarjeta Clientes") = vbNo Then GoTo error
    
    rcConsulta.Open "select NumTarjeta from tarjetas where NumTarjeta='" & Trim(txtNumTarjeta.Text) & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF And Val(txtNumTarjeta.Tag) = 0 Then
        MsgBox "El número de Tarjeta que desea grabar ya existe !!", vbCritical, "Tarjeta Clientes"
        txtNumTarjeta.SetFocus
        GoTo error
    End If
    rcConsulta.Close
    
    If cmbEstadoCivil.ListIndex > -1 Then
        EstadoCivil = cmbEstadoCivil.ItemData(cmbEstadoCivil.ListIndex)
    Else
        EstadoCivil = 0
    End If
    
    If Val(txtNumTarjeta.Tag) = 0 Then
        sql = "insert into tarjetas (Fecha,IDCliente,IDSucursal,Conyuge,LugarNacimiento,Email,EstadoCivil,Telefono,Dueno,Renta,Personas,Antiguedad,Situacion,Empresa,DomicilioEmpresa,Giro,Puesto,AntiguedadTrabajo,Ingreso,TrabajoAnterior,TelTrabajo,IFE,Pasaporte,Cartilla,Cedula,Vencimiento,NumTarjeta)values " _
            & "('" & Format(Date, "YYYY/MM/DD") & "'," & Val(txtCliente.Tag) & "," & Val(txtSucursal.Tag) & ",'" & Trim(txtConyuge.Text) & "','" & txtLugarNac.Text & "','" & txtEmail.Text & "'," & EstadoCivil & ",'" & txtTelefono.Text & "','" & txtDueño.Text & "','" & txtRenta.Text & "','" & txtPersonas.Text & "'," _
            & "'" & txtAntiguedad.Text & "','" & txtSituacion.Text & "','" & txtEmpresa.Text & "','" & txtDomicilioEmpresa.Text & "','" & txtGiro.Text & "','" & txtPuesto.Text & "','" & txtAntiguedadTra.Text & "','" & txtIngreso.Text & "','" & txtTrabajo.Text & "','" & txtTelTrabajo.Text & "','" & txtIfe.Text & "','" & txtPasaporte.Text & "','" & txtCartilla.Text & "','" & txtCedula.Text & "','" & Format(txtVencimiento.Text, "YYYY/MM/DD") & "','" & Trim(txtNumTarjeta.Text) & "')"
        
        dbDatos.Execute "update clientes set NumTarjeta='" & Trim(txtNumTarjeta.Text) & "' where ID=" & Val(txtCliente.Tag) & ""
    
    Else
        If MsgBox("Desea guardar los cambios realizados ??", vbQuestion + vbYesNo + vbDefaultButton1, "Tarjeta Clientes") = vbYes Then
                sql = "update tarjetas set IDSucursal=" & Val(txtSucursal.Tag) & ",Conyuge='" & Trim(txtConyuge.Text) & "',LugarNacimiento='" & txtLugarNac.Text & "',Email='" & Trim(txtEmail.Text) & "',EstadoCivil=" & EstadoCivil & ",Telefono='" & txtTelefono.Text & "',Dueno='" & Trim(txtDueño.Text) & "',Renta='" & Trim(txtRenta.Text) & "',Personas='" & Trim(txtPersonas.Text) & "',Antiguedad='" & Trim(txtAntiguedad.Text) & "',Situacion='" & Trim(txtSituacion.Text) & "',Empresa='" & Trim(txtEmpresa.Text) & "',DomicilioEmpresa='" & txtDomicilioEmpresa.Text & "',Giro='" & Trim(txtGiro.Text) & "',Puesto='" & txtPuesto.Text & "'," _
                & "AntiguedadTrabajo='" & Trim(txtAntiguedadTra.Text) & "',Ingreso='" & Trim(txtIngreso.Text) & "',TrabajoAnterior='" & Trim(txtTrabajo.Text) & "',TelTrabajo='" & Trim(txtTelTrabajo.Text) & "',IFE='" & Trim(txtIfe.Text) & "',Pasaporte='" & Trim(txtPasaporte.Text) & "',Cartilla='" & Trim(txtCartilla.Text) & "',Cedula='" & Trim(txtCedula.Text) & "',Vencimiento='" & Format(txtVencimiento.Text, "YYYY/MM/DD") & "' where ID=" & Val(txtNumTarjeta.Tag) & ""
            
        Else
            GoTo error
        End If
    End If
    
    dbDatos.Execute sql
    
    Limpiar
    txtNumTarjeta.SetFocus
End If

error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub cmdLimpiar_Click()
Limpiar
txtNumTarjeta.SetFocus
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
txtVencimiento = frmCalendario.Fecha(Trim(txtVencimiento.Text), 1)
End Sub

Private Sub cmdMosusuario_Click()
frmMostrarCliente.Ver Me, txtCliente, True, 0
End Sub

Private Sub cmdMuestraSucursal_Click()
frmMostrarSucursales.Ver Me, txtSucursal, True, True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
inicializar
End Sub

Sub inicializar()
Poner_Flat Fl, Me.Controls, Me
Cargar_Combos "Descripcion", "estadocivil", cmbEstadoCivil
lblFecha.Caption = Format(Date, "DD/MMM/YYYY")
CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
Quitar_Flat Fl
End Sub

Private Sub txtAntiguedad_GotFocus()
Seleccionar_Texto txtAntiguedad
Cambiar_Color True, txtAntiguedad
End Sub

Private Sub txtAntiguedad_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtAntiguedad_LostFocus()
Cambiar_Color False, txtAntiguedad
End Sub

Private Sub txtAntiguedadTra_GotFocus()
Seleccionar_Texto txtAntiguedadTra
Cambiar_Color True, txtAntiguedadTra
End Sub

Private Sub txtAntiguedadTra_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtAntiguedadTra_LostFocus()
Cambiar_Color False, txtAntiguedadTra
End Sub

Private Sub txtCartilla_GotFocus()
Seleccionar_Texto txtCartilla
Cambiar_Color True, txtCartilla
End Sub

Private Sub txtCartilla_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtCartilla_LostFocus()
Cambiar_Color False, txtCartilla
End Sub

Private Sub txtCedula_GotFocus()
Seleccionar_Texto txtCedula
Cambiar_Color True, txtCedula
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtCedula_LostFocus()
Cambiar_Color False, txtCedula
End Sub

Private Sub txtCliente_GotFocus()
Seleccionar_Texto txtCliente
Cambiar_Color True, txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtCliente_LostFocus()
Cambiar_Color False, txtCliente
End Sub

Private Sub txtConyuge_GotFocus()
Seleccionar_Texto txtConyuge
Cambiar_Color True, txtConyuge
End Sub

Private Sub txtConyuge_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtConyuge_LostFocus()
Cambiar_Color False, txtConyuge
End Sub

Private Sub txtDomicilioEmpresa_GotFocus()
Seleccionar_Texto txtDomicilioEmpresa
Cambiar_Color True, txtDomicilioEmpresa
End Sub

Private Sub txtDomicilioEmpresa_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtDomicilioEmpresa_LostFocus()
Cambiar_Color False, txtDomicilioEmpresa
End Sub

Private Sub txtDueño_GotFocus()
Seleccionar_Texto txtDueño
Cambiar_Color True, txtDueño
End Sub

Private Sub txtDueño_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtDueño_LostFocus()
Cambiar_Color False, txtDueño
End Sub

Private Sub txtEmail_GotFocus()
Seleccionar_Texto txtEmail
Cambiar_Color True, txtEmail
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtEmail_LostFocus()
Cambiar_Color False, txtEmail
End Sub

Private Sub txtEmpresa_GotFocus()
Seleccionar_Texto txtEmpresa
Cambiar_Color True, txtEmpresa
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtEmpresa_LostFocus()
Cambiar_Color False, txtEmpresa
End Sub

Private Sub txtGiro_GotFocus()
Seleccionar_Texto txtGiro
Cambiar_Color True, txtGiro
End Sub

Private Sub txtGiro_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtGiro_LostFocus()
Cambiar_Color False, txtGiro
End Sub

Private Sub txtIfe_GotFocus()
Seleccionar_Texto txtIfe
Cambiar_Color True, txtIfe
End Sub

Private Sub txtIfe_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtIfe_LostFocus()
Cambiar_Color False, txtIfe
End Sub

Private Sub txtIngreso_GotFocus()
Seleccionar_Texto txtIngreso
Cambiar_Color True, txtIngreso
End Sub

Private Sub txtIngreso_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtIngreso_LostFocus()
Cambiar_Color False, txtIngreso
End Sub

Private Sub txtLugarNac_GotFocus()
Seleccionar_Texto txtLugarNac
Cambiar_Color True, txtLugarNac
End Sub

Private Sub txtLugarNac_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtLugarNac_LostFocus()
Cambiar_Color False, txtLugarNac
End Sub

Private Sub txtNumTarjeta_GotFocus()
Seleccionar_Texto txtNumTarjeta
Cambiar_Color True, txtNumTarjeta
End Sub

Private Sub txtNumTarjeta_KeyPress(KeyAscii As Integer)
KeyAscii = Solo_Numeros(KeyAscii)
If KeyAscii = vbKeyReturn Then
    BuscaTarjeta Trim(txtNumTarjeta.Text)
End If
End Sub

Private Sub txtNumTarjeta_LostFocus()
Cambiar_Color False, txtNumTarjeta
End Sub

Private Sub txtPasaporte_GotFocus()
Seleccionar_Texto txtPasaporte
Cambiar_Color True, txtPasaporte
End Sub

Private Sub txtPasaporte_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtPasaporte_LostFocus()
Cambiar_Color False, txtPasaporte
End Sub

Private Sub txtPersonas_GotFocus()
Seleccionar_Texto txtPersonas
Cambiar_Color True, txtPersonas
End Sub

Private Sub txtPersonas_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtPersonas_LostFocus()
Cambiar_Color False, txtPersonas
End Sub

Private Sub txtPuesto_GotFocus()
Seleccionar_Texto txtPuesto
Cambiar_Color True, txtPuesto
End Sub

Private Sub txtPuesto_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtPuesto_LostFocus()
Cambiar_Color False, txtPuesto
End Sub

Private Sub txtRenta_GotFocus()
Seleccionar_Texto txtRenta
Cambiar_Color True, txtRenta
End Sub

Private Sub txtRenta_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtRenta_LostFocus()
Cambiar_Color False, txtRenta
End Sub

Private Sub txtSituacion_GotFocus()
Seleccionar_Texto txtSituacion
Cambiar_Color True, txtSituacion
End Sub

Private Sub txtSituacion_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtSituacion_LostFocus()
Cambiar_Color False, txtSituacion
End Sub

Private Sub txtSucursal_GotFocus()
Seleccionar_Texto txtSucursal
Cambiar_Color True, txtSucursal
End Sub

Private Sub txtSucursal_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtSucursal_LostFocus()
Cambiar_Color False, txtSucursal
End Sub

Private Sub txtTelefono_GotFocus()
Seleccionar_Texto txtTelefono
Cambiar_Color True, txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtTelefono_LostFocus()
Cambiar_Color False, txtTelefono
End Sub

Private Sub txtTelTrabajo_GotFocus()
Seleccionar_Texto txtTelTrabajo
Cambiar_Color True, txtTelTrabajo
End Sub

Private Sub txtTelTrabajo_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtTelTrabajo_LostFocus()
Cambiar_Color False, txtTelTrabajo
End Sub

Private Sub txtTrabajo_GotFocus()
Seleccionar_Texto txtTrabajo
Cambiar_Color True, txtTrabajo
End Sub

Private Sub txtTrabajo_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtTrabajo_LostFocus()
Cambiar_Color False, txtTrabajo
End Sub

Public Function Buscar(IDCliente As Long)
Dim rcCliente As New ADODB.Recordset

On Error GoTo error

rcCliente.Open "select ID,concat(Nombre,' ',Apellido) as Cliente from clientes where ID=" & IDCliente, dbDatos, adOpenForwardOnly, adLockOptimistic
If Not rcCliente.BOF And Not rcCliente.EOF Then
    txtCliente.Text = rcCliente!cliente
    txtCliente.Tag = rcCliente!ID
End If
rcCliente.Close

error:
    Maneja_Error Err
    Set rcCliente = Nothing
End Function

Function Limpiar()
Dim ctrl As Control

For Each ctrl In Controls
    On Error Resume Next
    If TypeOf ctrl Is TextBox Then ctrl.Text = ""
    If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
    ctrl.Tag = ""
Next
End Function

Public Function BuscarSucursal(IDSucursal As Long)
Dim rcSucursal As New ADODB.Recordset

On Error GoTo error

rcSucursal.Open "select ID,RazonSocial from sucursales where ID=" & IDSucursal, dbDatos, adOpenForwardOnly, adLockOptimistic
If Not rcSucursal.BOF And Not rcSucursal.EOF Then
    txtSucursal.Text = rcSucursal!RazonSocial
    txtSucursal.Tag = rcSucursal!ID
End If
rcSucursal.Close

error:
    Maneja_Error Err
    Set rcSucursal = Nothing
End Function

Function ValidaDatos() As Boolean
ValidaDatos = True

If txtNumTarjeta.Text = "" Then
    MsgBox "Introduzca el número de tarjeta !!", vbInformation, "Tarjeta Clientes"
    ValidaDatos = False
    txtNumTarjeta.SetFocus
    Exit Function
End If

If txtCliente.Text = "" Then
    MsgBox "Seleccione el cliente !!", vbInformation, "Tarjeta Clientes"
    ValidaDatos = False
    txtCliente.SetFocus
    Exit Function
End If

If txtVencimiento.Text = "" Then
    MsgBox "Seleccione el vencimiento de la tarjeta !!", vbInformation, "Tarjeta Clientes"
    ValidaDatos = False
    txtVencimiento.SetFocus
    Exit Function
End If

End Function

Private Sub txtVencimiento_GotFocus()
Seleccionar_Texto txtVencimiento
Cambiar_Color True, txtVencimiento
End Sub

Private Sub txtVencimiento_KeyPress(KeyAscii As Integer)
Pasar_Foco KeyAscii
End Sub

Private Sub txtVencimiento_LostFocus()
Cambiar_Color False, txtVencimiento
End Sub

Function BuscaTarjeta(NumTarjeta As String)
Dim rcConsulta As New ADODB.Recordset

On Error GoTo error

rcConsulta.Open "select tarjetas.*,concat(clientes.Nombre,' ',clientes.Apellido) as Cliente,sucursales.razonsocial as Sucursal from tarjetas Inner Join clientes on tarjetas.IDCliente=clientes.ID Left Join sucursales on tarjetas.IDSucursal=sucursales.ID where tarjetas.NumTarjeta='" & Trim(NumTarjeta) & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
    With rcConsulta
        If Not .BOF And Not .EOF Then
            txtCliente.Text = !cliente
            txtNumTarjeta.Tag = !ID
            txtSucursal.Text = IIf(IsNull(!Sucursal), "", !Sucursal)
            txtConyuge.Text = IIf(IsNull(!Conyuge), "", !Conyuge)
            txtLugarNac.Text = IIf(IsNull(!LugarNacimiento), "", !LugarNacimiento)
            txtEmail.Text = IIf(IsNull(!Email), "", !Email)
            cmbEstadoCivil.ListIndex = ComboInformacion(cmbEstadoCivil, !EstadoCivil)
            txtTelefono.Text = IIf(IsNull(!Telefono), "", !Telefono)
            txtDueño.Text = IIf(IsNull(!Dueno), "", !Dueno)
            txtRenta.Text = IIf(IsNull(!Renta), "", !Renta)
            txtPersonas.Text = IIf(IsNull(!Personas), "", !Personas)
            txtAntiguedad.Text = IIf(IsNull(!Antiguedad), "", !Antiguedad)
            txtSituacion.Text = IIf(IsNull(!Situacion), "", !Situacion)
            txtEmpresa.Text = IIf(IsNull(!Empresa), "", !Empresa)
            txtDomicilioEmpresa.Text = IIf(IsNull(!DomicilioEmpresa), "", !DomicilioEmpresa)
            txtGiro.Text = IIf(IsNull(!Giro), "", !Giro)
            txtPuesto.Text = IIf(IsNull(!Puesto), "", !Puesto)
            txtAntiguedadTra.Text = IIf(IsNull(!AntiguedadTrabajo), "", !AntiguedadTrabajo)
            txtIngreso.Text = IIf(IsNull(!Ingreso), "", !Ingreso)
            txtTrabajo.Text = IIf(IsNull(!TrabajoAnterior), "", !TrabajoAnterior)
            txtTelTrabajo.Text = IIf(IsNull(!TelTrabajo), "", !TelTrabajo)
            txtIfe.Text = IIf(IsNull(!Ife), "", !Ife)
            txtPasaporte.Text = IIf(IsNull(!Pasaporte), "", !Pasaporte)
            txtCartilla.Text = IIf(IsNull(!Cartilla), "", !Cartilla)
            txtCedula.Text = IIf(IsNull(!Cedula), "", !Cedula)
            txtVencimiento.Text = IIf(IsNull(!Vencimiento), "", Format(!Vencimiento, "DD/MMM/YYYY"))
        End If
    End With
rcConsulta.Close

error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function
