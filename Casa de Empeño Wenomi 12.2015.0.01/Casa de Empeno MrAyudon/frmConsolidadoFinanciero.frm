VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmConsolidadoFinanciero 
   Caption         =   "Consolidado Financiero"
   ClientHeight    =   6735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6735
   ScaleWidth      =   10740
   Begin VB.TextBox txtFechaIni 
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
      Left            =   1080
      TabIndex        =   54
      Top             =   240
      Width           =   1095
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9480
      TabIndex        =   17
      Top             =   6240
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
      Picture         =   "frmConsolidadoFinanciero.frx":0000
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   8160
      TabIndex        =   18
      Top             =   6240
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
      Index           =   0
      Left            =   2280
      TabIndex        =   55
      Top             =   240
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
      Picture         =   "frmConsolidadoFinanciero.frx":0091
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   2760
      TabIndex        =   56
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16777215
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmConsolidadoFinanciero.frx":01A6
   End
   Begin VB.Label Label8 
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
      TabIndex        =   57
      Top             =   240
      Width           =   795
   End
   Begin VB.Label lblTotales 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<PreTotal>"
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
      Index           =   6
      Left            =   9195
      TabIndex        =   53
      Top             =   4920
      Width           =   1245
   End
   Begin VB.Label lblPrestamos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Prestamos>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   6960
      TabIndex        =   52
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblPrestamos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Prestamos>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5160
      TabIndex        =   51
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblPrestamos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Prestamos>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   3360
      TabIndex        =   50
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label lblPrestamos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Prestamos>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   1680
      TabIndex        =   49
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Prestamos:"
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
      TabIndex        =   48
      Top             =   4920
      Width           =   1380
   End
   Begin VB.Label lblGranTotal 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<GranTota>"
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
      Left            =   8745
      TabIndex        =   47
      Top             =   5640
      Width           =   1725
   End
   Begin VB.Label lblTotales 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<FalTotal>"
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
      Index           =   5
      Left            =   9225
      TabIndex        =   46
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Label lblTotales 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<ApaTotal>"
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
      Index           =   4
      Left            =   9105
      TabIndex        =   45
      Top             =   3720
      Width           =   1320
   End
   Begin VB.Label lblTotales 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<JoyTotal>"
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
      Index           =   3
      Left            =   9165
      TabIndex        =   44
      Top             =   3120
      Width           =   1260
   End
   Begin VB.Label lblTotales 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<EmpTotal>"
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
      Index           =   2
      Left            =   9045
      TabIndex        =   43
      Top             =   2520
      Width           =   1380
   End
   Begin VB.Label lblTotales 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<BanTotal>"
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
      Index           =   1
      Left            =   9135
      TabIndex        =   42
      Top             =   1920
      Width           =   1290
   End
   Begin VB.Label lblTotales 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<BovTotal>"
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
      Index           =   0
      Left            =   8715
      TabIndex        =   41
      Top             =   1320
      Width           =   1665
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Totales"
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
      Left            =   8655
      TabIndex        =   40
      Top             =   840
      Width           =   1785
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "<Total>"
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
      Index           =   3
      Left            =   6960
      TabIndex        =   39
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "<Total>"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   38
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "<Total>"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   37
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label lblBoveda 
      Alignment       =   1  'Right Justify
      Caption         =   "<Boveda>"
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
      Index           =   3
      Left            =   6960
      TabIndex        =   36
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblBancos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Bancos>"
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
      Index           =   3
      Left            =   6960
      TabIndex        =   35
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblEmpeño 
      Alignment       =   1  'Right Justify
      Caption         =   "<Empeño>"
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
      Index           =   3
      Left            =   6960
      TabIndex        =   34
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblJoyeria 
      Alignment       =   1  'Right Justify
      Caption         =   "<Joyeria>"
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
      Index           =   3
      Left            =   6960
      TabIndex        =   33
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblApartados 
      Alignment       =   1  'Right Justify
      Caption         =   "<Apartados>"
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
      Index           =   3
      Left            =   6960
      TabIndex        =   32
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblFaltante 
      Alignment       =   1  'Right Justify
      Caption         =   "<Faltante>"
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
      Index           =   3
      Left            =   6960
      TabIndex        =   31
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblBoveda 
      Alignment       =   1  'Right Justify
      Caption         =   "<Boveda>"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   30
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblBancos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Bancos>"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   29
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblEmpeño 
      Alignment       =   1  'Right Justify
      Caption         =   "<Empeño>"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   28
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblJoyeria 
      Alignment       =   1  'Right Justify
      Caption         =   "<Joyeria>"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   27
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblApartados 
      Alignment       =   1  'Right Justify
      Caption         =   "<Apartados>"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   26
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblFaltante 
      Alignment       =   1  'Right Justify
      Caption         =   "<Faltante>"
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
      Index           =   2
      Left            =   5160
      TabIndex        =   25
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label lblBoveda 
      Alignment       =   1  'Right Justify
      Caption         =   "<Boveda>"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   24
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label lblBancos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Bancos>"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   23
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label lblEmpeño 
      Alignment       =   1  'Right Justify
      Caption         =   "<Empeño>"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   22
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label lblJoyeria 
      Alignment       =   1  'Right Justify
      Caption         =   "<Joyeria>"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   21
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label lblApartados 
      Alignment       =   1  'Right Justify
      Caption         =   "<Apartados>"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label lblFaltante 
      Alignment       =   1  'Right Justify
      Caption         =   "<Faltante>"
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
      Index           =   1
      Left            =   3360
      TabIndex        =   19
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label31 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Morelos"
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
      Left            =   7365
      TabIndex        =   16
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label30 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Bazareño"
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
      Left            =   5460
      TabIndex        =   15
      Top             =   840
      Width           =   1185
   End
   Begin VB.Label Label29 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "La Nacional"
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
      Left            =   3495
      TabIndex        =   14
      Top             =   840
      Width           =   1425
   End
   Begin VB.Label Label28 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Ocampo"
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
      Left            =   1680
      TabIndex        =   13
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label lblBoveda 
      Alignment       =   1  'Right Justify
      Caption         =   "<Boveda>"
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
      Index           =   0
      Left            =   1605
      TabIndex        =   12
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Boveda:"
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
      TabIndex        =   11
      Top             =   1320
      Width           =   990
   End
   Begin VB.Label lblBancos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Bancos>"
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
      Index           =   0
      Left            =   1605
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Bancos:"
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
      TabIndex        =   9
      Top             =   1920
      Width           =   945
   End
   Begin VB.Label lblEmpeño 
      Alignment       =   1  'Right Justify
      Caption         =   "<Empeño>"
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
      Index           =   0
      Left            =   1605
      TabIndex        =   8
      Top             =   2520
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Empeño:"
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
      TabIndex        =   7
      Top             =   2520
      Width           =   1065
   End
   Begin VB.Label lblJoyeria 
      Alignment       =   1  'Right Justify
      Caption         =   "<Joyeria>"
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
      Index           =   0
      Left            =   1605
      TabIndex        =   6
      Top             =   3120
      Width           =   1575
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Joyeria:"
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
      TabIndex        =   5
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label lblApartados 
      Alignment       =   1  'Right Justify
      Caption         =   "<Apartados>"
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
      Index           =   0
      Left            =   1605
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Apartados:"
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
      TabIndex        =   3
      Top             =   3720
      Width           =   1350
   End
   Begin VB.Label lblFaltante 
      Alignment       =   1  'Right Justify
      Caption         =   "<Faltante>"
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
      Index           =   0
      Left            =   1605
      TabIndex        =   2
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "Faltante:"
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
      Top             =   4320
      Width           =   1110
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      X1              =   1560
      X2              =   10560
      Y1              =   5400
      Y2              =   5400
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      X1              =   1560
      X2              =   10560
      Y1              =   5520
      Y2              =   5520
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "<Total>"
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
      Index           =   0
      Left            =   1560
      TabIndex        =   0
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "frmConsolidadoFinanciero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    Imprimir
End Sub

Private Sub cmdBuscar_Click()
For Each ctrl In Controls
    If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = "0.00"
Next
If txtFechaIni.Text = "" Then
    MsgBox "Favor de escoger una fecha para realizar la busqueda", vbCritical
Else
    Cargar_Montos txtFechaIni.Text
    Poner_Totales
End If
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
txtFechaIni.Text = frmCalendario.Fecha(txtFechaIni.Text)
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
inicializar
End Sub

Private Sub inicializar()
Dim ctrl As Control
Screen.MousePointer = vbHourglass
    Me.Height = 7245
    Me.Width = 10860
    CentrarForm frmConsolidadoFinanciero, frmMDI
    For Each ctrl In Controls
        If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = "0.00"
    Next
Screen.MousePointer = vbDefault
End Sub

Private Sub Cargar_Montos(Fecha As String)
Dim rcBD As New ADODB.Recordset

rcBD.Open "SELECT * FROM Financiero WHERE Fecha=#" & Format(Fecha, "MM/DD/YY") & "# ORDER BY Sucursal", dbDatos, adOpenDynamic, adLockOptimistic

With rcBD
    If rcBD.EOF = False And rcBD.BOF = False Then
        While Not .EOF
            lblBoveda(!Sucursal - 1).Caption = Format(rcBD!Boveda, "###,###,##0.00")
            lblBancos(!Sucursal - 1).Caption = Format(rcBD!Bancos, "###,###,#0.00")
            lblEmpeño(!Sucursal - 1).Caption = Format(rcBD!empeño, "###,###,##0.00")
            lblJoyeria(!Sucursal - 1).Caption = Format(rcBD!Joyeria, "###,###,##0.00")
            lblApartados(!Sucursal - 1).Caption = Format(rcBD!Apartados, "###,###,##0.00")
            lblFaltante(!Sucursal - 1).Caption = Format(rcBD!Faltante, "###,###,##0.00")
            lblPrestamos(!Sucursal - 1).Caption = Format(rcBD!Prestamos, "###,###,##0.00")
            .MoveNext
        Wend
    End If
    .Close
End With

End Sub

Private Sub Poner_Totales()
Dim Band As Integer

For Band = 0 To 3
    lblTotal(Band).Caption = Format(CCur(lblBoveda(Band).Caption) + CCur(lblBancos(Band).Caption) + CCur(lblEmpeño(Band).Caption) + CCur(lblJoyeria(Band).Caption) + CCur(lblApartados(Band).Caption) + CCur(lblFaltante(Band).Caption), "###,###,###,#0.00")
Next Band

lblTotales(0).Caption = Format(CCur(lblBoveda(0).Caption) + CCur(lblBoveda(1).Caption) + CCur(lblBoveda(2).Caption) + CCur(lblBoveda(3).Caption), "###,###,##0.00")
lblTotales(1).Caption = Format(CCur(lblBancos(0).Caption) + CCur(lblBancos(1).Caption) + CCur(lblBancos(2).Caption) + CCur(lblBancos(3).Caption), "###,###,##0.00")
lblTotales(2).Caption = Format(CCur(lblEmpeño(0).Caption) + CCur(lblEmpeño(1).Caption) + CCur(lblEmpeño(2).Caption) + CCur(lblEmpeño(3).Caption), "###,###,##0.00")
lblTotales(3).Caption = Format(CCur(lblJoyeria(0).Caption) + CCur(lblJoyeria(1).Caption) + CCur(lblJoyeria(2).Caption) + CCur(lblJoyeria(3).Caption), "###,###,##0.00")
lblTotales(4).Caption = Format(CCur(lblApartados(0).Caption) + CCur(lblApartados(1).Caption) + CCur(lblApartados(2).Caption) + CCur(lblApartados(3).Caption), "###,###,##0.00")
lblTotales(5).Caption = Format(CCur(lblFaltante(0).Caption) + CCur(lblFaltante(1).Caption) + CCur(lblFaltante(2).Caption) + CCur(lblFaltante(3).Caption), "###,###,##0.00")
lblTotales(6).Caption = Format(CCur(lblPrestamos(0).Caption) + CCur(lblPrestamos(1).Caption) + CCur(lblPrestamos(2).Caption) + CCur(lblPrestamos(3).Caption), "###,###,##0.00")

lblGranTotal.Caption = Format(CCur(lblTotales(0).Caption) + CCur(lblTotales(1).Caption) + CCur(lblTotales(2).Caption) + CCur(lblTotales(3).Caption) + CCur(lblTotales(4).Caption) + CCur(lblTotales(5).Caption) + CCur(lblTotales(6).Caption), "###,###,##0.00")

End Sub

Private Sub Imprimir()
  With frmMDI.Cr
    .DataFiles(0) = path & "\Base De Datos\Datos.mdb"
    .DiscardSavedData = True
    .WindowShowPrintSetupBtn = True
    .ReportFileName = path & "\Reportes\ConsolidadoFinanciero.rpt"
    .Formulas(0) = "Boveda=" & CCur(lblBoveda(0).Caption) & ""
    .Formulas(1) = "Bancos=" & CCur(lblBancos(0).Caption) & ""
    .Formulas(2) = "Empeño=" & CCur(lblEmpeño(0).Caption) & ""
    .Formulas(3) = "Joyeria=" & CCur(lblJoyeria(0).Caption) & ""
    .Formulas(4) = "Apartados=" & CCur(lblApartados(0).Caption) & ""
    .Formulas(5) = "Faltante=" & CCur(lblFaltante(0).Caption) & ""
    .Formulas(6) = "Total=" & CCur(lblTotal(0).Caption) & ""
    .Formulas(7) = "Boveda2=" & CCur(lblBoveda(1).Caption) & ""
    .Formulas(8) = "Bancos2=" & CCur(lblBancos(1).Caption) & ""
    .Formulas(9) = "Empeño2=" & CCur(lblEmpeño(1).Caption) & ""
    .Formulas(10) = "Joyeria2=" & CCur(lblJoyeria(1).Caption) & ""
    .Formulas(11) = "Apartados2=" & CCur(lblApartados(1).Caption) & ""
    .Formulas(12) = "Faltante2=" & CCur(lblFaltante(1).Caption) & ""
    .Formulas(13) = "Total2=" & CCur(lblTotal(1).Caption) & ""
    .Formulas(14) = "Boveda3=" & CCur(lblBoveda(2).Caption) & ""
    .Formulas(15) = "Bancos3=" & CCur(lblBancos(2).Caption) & ""
    .Formulas(16) = "Empeño3=" & CCur(lblEmpeño(2).Caption) & ""
    .Formulas(17) = "Joyeria3=" & CCur(lblJoyeria(2).Caption) & ""
    .Formulas(18) = "Apartados3=" & CCur(lblApartados(2).Caption) & ""
    .Formulas(19) = "Faltante3=" & CCur(lblFaltante(2).Caption) & ""
    .Formulas(20) = "Total3=" & CCur(lblTotal(2).Caption) & ""
    .Formulas(21) = "Boveda4=" & CCur(lblBoveda(3).Caption) & ""
    .Formulas(22) = "Bancos4=" & CCur(lblBancos(3).Caption) & ""
    .Formulas(23) = "Empeño4=" & CCur(lblEmpeño(3).Caption) & ""
    .Formulas(24) = "Joyeria4=" & CCur(lblJoyeria(3).Caption) & ""
    .Formulas(25) = "Apartados4=" & CCur(lblApartados(3).Caption) & ""
    .Formulas(26) = "Faltante4=" & CCur(lblFaltante(3).Caption) & ""
    .Formulas(27) = "Total4=" & CCur(lblTotal(3).Caption) & ""
    .Formulas(28) = "TotBoveda=" & CCur(lblTotales(0).Caption) & ""
    .Formulas(29) = "TotBancos=" & CCur(lblTotales(1).Caption) & ""
    .Formulas(30) = "TotEmpeño=" & CCur(lblTotales(2).Caption) & ""
    .Formulas(31) = "TotJoyeria=" & CCur(lblTotales(3).Caption) & ""
    .Formulas(32) = "TotApartados=" & CCur(lblTotales(4).Caption) & ""
    .Formulas(33) = "TotFaltante=" & CCur(lblTotales(5).Caption) & ""
    .Formulas(34) = "GranTotal=" & CCur(lblGranTotal.Caption) & ""
    .Formulas(35) = "Prestamos=" & CCur(lblPrestamos(0).Caption) & ""
    .Formulas(36) = "Prestamos2=" & CCur(lblPrestamos(1).Caption) & ""
    .Formulas(37) = "Prestamos3=" & CCur(lblPrestamos(2).Caption) & ""
    .Formulas(38) = "Prestamos4=" & CCur(lblPrestamos(3).Caption) & ""
    .Formulas(39) = "TotPrestamos=" & CCur(lblTotales(6).Caption) & ""
    .Destination = crptToWindow
    .Action = 1
  End With
End Sub
