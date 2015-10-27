VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmConfiguracionTasas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración Tasas"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14490
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfiguracionTasas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   14490
   Begin VB.ComboBox cmbPersona 
      Height          =   315
      ItemData        =   "frmConfiguracionTasas.frx":000C
      Left            =   10320
      List            =   "frmConfiguracionTasas.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   37
      Top             =   1150
      Width           =   1245
   End
   Begin VB.TextBox txtDiasGracias 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8760
      MaxLength       =   3
      TabIndex        =   13
      Top             =   1170
      Width           =   1305
   End
   Begin VB.TextBox txtDesempenoExt 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10320
      MaxLength       =   6
      TabIndex        =   11
      Top             =   1560
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txtInteresAnual 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   8760
      MaxLength       =   6
      TabIndex        =   5
      Top             =   465
      Width           =   1305
   End
   Begin VB.TextBox txtDiasMinimos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7080
      MaxLength       =   3
      TabIndex        =   12
      Top             =   1170
      Width           =   1305
   End
   Begin VB.OptionButton opElectronicos 
      Appearance      =   0  'Flat
      Caption         =   "Electrónicos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   12960
      TabIndex        =   32
      Top             =   120
      Width           =   1425
   End
   Begin VB.OptionButton opAutomovil 
      Appearance      =   0  'Flat
      Caption         =   "Automóvil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11880
      TabIndex        =   31
      Top             =   420
      Width           =   1320
   End
   Begin VB.OptionButton opMetales 
      Appearance      =   0  'Flat
      Caption         =   "Metales"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11880
      TabIndex        =   30
      Top             =   120
      Width           =   1035
   End
   Begin VB.TextBox txtAlmacenaje 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5490
      MaxLength       =   6
      TabIndex        =   10
      Top             =   1170
      Width           =   1305
   End
   Begin VB.TextBox txtAlmacenajeAnual 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   10320
      MaxLength       =   6
      TabIndex        =   6
      Top             =   480
      Width           =   1305
   End
   Begin VB.TextBox txtCat 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7155
      MaxLength       =   6
      TabIndex        =   4
      Top             =   480
      Width           =   1305
   End
   Begin VB.TextBox txtPrestamo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5490
      MaxLength       =   5
      TabIndex        =   3
      Top             =   480
      Width           =   1305
   End
   Begin Line3D.ucLine3D ucLine3D5 
      Height          =   30
      Index           =   0
      Left            =   150
      Top             =   375
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D4 
      Height          =   30
      Left            =   150
      Top             =   1500
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   1410
      Index           =   0
      Left            =   5295
      Top             =   120
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2487
      Orientation     =   0
      LineWidth       =   2
   End
   Begin VB.TextBox txtTasaPreferente 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   3795
      MaxLength       =   6
      TabIndex        =   9
      Top             =   1170
      Width           =   1305
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Left            =   120
      Top             =   120
      Width           =   11595
      _ExtentX        =   20452
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1395
      Index           =   0
      Left            =   135
      Top             =   105
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2461
      Orientation     =   0
      LineWidth       =   2
   End
   Begin VB.TextBox txtTasaPromocion 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   8
      Top             =   1170
      Width           =   1305
   End
   Begin vbAcceleratorGrid6.vbalGrid grdTasas 
      Height          =   4275
      Left            =   120
      TabIndex        =   19
      Top             =   1680
      Width           =   14295
      _ExtentX        =   25215
      _ExtentY        =   7541
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
   End
   Begin VB.TextBox txtTasaTipica 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   330
      MaxLength       =   6
      TabIndex        =   7
      Top             =   1170
      Width           =   1305
   End
   Begin VB.ComboBox cmbTipoInteres 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   195
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   420
      Width           =   1560
   End
   Begin VB.ComboBox cmbPlazo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3660
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   420
      Width           =   1560
   End
   Begin VB.ComboBox cmbTipoPeriodo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1905
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   1560
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   13065
      TabIndex        =   20
      Top             =   1155
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
      Picture         =   "frmConfiguracionTasas.frx":0010
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   11940
      TabIndex        =   14
      Top             =   720
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
      Picture         =   "frmConfiguracionTasas.frx":0562
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   11940
      TabIndex        =   21
      Top             =   1155
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Cancelar"
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
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmConfiguracionTasas.frx":0AB4
      PictureDisabled =   "frmConfiguracionTasas.frx":0D03
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1410
      Index           =   1
      Left            =   1830
      Top             =   105
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2487
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1425
      Index           =   2
      Left            =   3540
      Top             =   90
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2514
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D5 
      Height          =   30
      Index           =   1
      Left            =   135
      Top             =   810
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D5 
      Height          =   30
      Index           =   2
      Left            =   135
      Top             =   1095
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   375
      Left            =   13065
      TabIndex        =   25
      Top             =   720
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Eliminar"
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
      Picture         =   "frmConfiguracionTasas.frx":18D5
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   1410
      Index           =   1
      Left            =   6960
      Top             =   120
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2487
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   1410
      Index           =   2
      Left            =   8595
      Top             =   120
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2487
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   1410
      Index           =   3
      Left            =   10200
      Top             =   120
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2487
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   1410
      Index           =   4
      Left            =   11720
      Top             =   120
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2487
      Orientation     =   0
      LineWidth       =   2
   End
   Begin VB.Label lblDiasGracia 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dias Gracia"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8880
      TabIndex        =   36
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label lblDesExt 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10725
      TabIndex        =   35
      Top             =   840
      Width           =   405
   End
   Begin VB.Label lblIntAnual 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Int. Anual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   8880
      TabIndex        =   34
      Top             =   135
      Width           =   1035
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Dias Mínimos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7080
      TabIndex        =   33
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Almacenaje"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5580
      TabIndex        =   29
      Top             =   840
      Width           =   1125
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Alm. Anual"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   10440
      TabIndex        =   28
      Top             =   135
      Width           =   1035
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CAT"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   7620
      TabIndex        =   27
      Top             =   135
      Width           =   375
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "% Préstamo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   5542
      TabIndex        =   26
      Top             =   135
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tasa Preferente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3660
      TabIndex        =   24
      Top             =   840
      Width           =   1575
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tasa Promoción"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1935
      TabIndex        =   22
      Top             =   840
      Width           =   1515
   End
   Begin VB.Label Label57 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tasa Típica"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   450
      TabIndex        =   18
      Top             =   840
      Width           =   1050
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Interés"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   390
      TabIndex        =   17
      Top             =   135
      Width           =   1170
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   4185
      TabIndex        =   16
      Top             =   135
      Width           =   510
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2325
      TabIndex        =   15
      Top             =   135
      Width           =   735
   End
   Begin VB.Label Label2 
      BackColor       =   &H00404040&
      Height          =   1395
      Left            =   130
      TabIndex        =   23
      Top             =   120
      Width           =   11580
   End
End
Attribute VB_Name = "frmConfiguracionTasas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fl() As cFlatControl

Private Sub cmbPersona_GotFocus()
    Cambiar_Color True, cmbPersona
End Sub

Private Sub cmbPersona_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPersona_LostFocus()
    Cambiar_Color False, cmbPersona
End Sub

Private Sub cmbTipoPeriodo_GotFocus()
    Cambiar_Color True, cmbTipoPeriodo
End Sub

Private Sub cmbTipoPeriodo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoPeriodo_LostFocus()
    Cambiar_Color False, cmbTipoPeriodo
End Sub

Private Sub cmbPlazo_GotFocus()
    Cambiar_Color True, cmbPlazo
End Sub

Private Sub cmbPlazo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPlazo_LostFocus()
    Cambiar_Color False, cmbPlazo
End Sub

Private Sub cmbTipoInteres_GotFocus()
    Cambiar_Color True, cmbTipoInteres
End Sub

Private Sub cmbTipoInteres_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoInteres_LostFocus()
    Cambiar_Color False, cmbTipoInteres
End Sub

Private Sub cmdAceptar_Click()
    
    If DatosValidos Then
        
        If Val(txtTasaTipica.Tag) = 0 Then
            
            dbDatos.Execute "INSERT INTO configuraciontasas (IDTipoInteres,IDTipoPeriodo,IDPlazo,TasaTipica,TasaPromocion,TasaPreferencial,PorPrestamo,Cat,Almacenaje,DMinimos,IntAnual,AlmAnual,DesExt,DGracia,Persona) VALUES (" & _
                            cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & "," & cmbTipoPeriodo.ItemData(cmbTipoPeriodo.ListIndex) & "," & cmbPlazo.ItemData(cmbPlazo.ListIndex) & "," & CDbl(txtTasaTipica.text) & "," & CDbl(txtTasaPromocion.text) & "," & CDbl(txtTasaPreferente.text) & "," & CDbl(txtPrestamo.text) & "," & CDbl(txtCat.text) & "," & CDbl(txtAlmacenaje.text) & "," & CInt(txtDiasMinimos.text) & "," & CDbl(txtInteresAnual.text) & "," & CDbl(txtAlmacenajeAnual.text) & ",0," & CInt(txtDiasGracias.text) & "," & IIf(opAutomovil.Value = True, cmbPersona.ItemData(cmbPersona.ListIndex), 0) & ")"
        Else
            dbDatos.Execute "UPDATE configuraciontasas SET TasaTipica = " & CDbl(txtTasaTipica.text) & ",TasaPromocion = " & CDbl(txtTasaPromocion.text) & ",TasaPreferencial = " & CDbl(txtTasaPreferente.text) & ",PorPrestamo = " & CDbl(txtPrestamo.text) & ",Cat=" & CDbl(txtCat.text) & ",Almacenaje=" & CDbl(txtAlmacenaje.text) & ",DMinimos=" & CInt(txtDiasMinimos.text) & ",IntAnual=" & CDbl(txtInteresAnual.text) & ",AlmAnual=" & CDbl(txtAlmacenajeAnual.text) & ",DesExt=0,DGracia=" & CInt(txtDiasGracias.text) & ",Persona=" & IIf(opAutomovil.Value = True, cmbPersona.ItemData(cmbPersona.ListIndex), 0) & " WHERE ID=" & Val(txtTasaTipica.Tag)
        End If
        
        'Cargo los Datos en el Grid
        CargarDatos
        
        cmbTipoInteres.ListIndex = -1
        cmbTipoPeriodo.ListIndex = -1
        cmbPlazo.ListIndex = -1
        
        txtTasaTipica.text = ""
        txtTasaTipica.Tag = ""
        txtTasaPromocion.text = ""
        txtTasaPreferente.text = ""
        txtAlmacenaje.text = ""
        'txtDesempenoExt.text = ""
        txtPrestamo.text = ""
        txtDiasMinimos.text = ""
        txtCat.text = ""
        txtInteresAnual.text = ""
        txtAlmacenajeAnual.text = ""
        txtDiasGracias.text = ""
        
        cmbTipoInteres.SetFocus
    End If

End Sub

Private Sub cmdEliminar_Click()

    If grdTasas.Rows > 0 Then
        
        If grdTasas.SelectedRow > 0 Then
            
            If MsgBox("Desea eliminar la opción seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Configuración Tasas") = vbYes Then
                
                dbDatos.Execute "DELETE FROM configuraciontasas WHERE ID=" & Val(grdTasas.CellItemData(grdTasas.SelectedRow, 4))
                grdTasas.RemoveRow grdTasas.SelectedRow
                
            End If
            
            grdTasas.ClearSelection
            cmbTipoInteres.SetFocus
        End If
        
    End If

End Sub

Private Sub cmdCancelar_Click()
    
    cmbTipoInteres.ListIndex = -1
    cmbTipoPeriodo.ListIndex = -1
    cmbPlazo.ListIndex = -1
    txtTasaTipica.text = ""
    txtTasaTipica.Tag = ""
    txtTasaPromocion.text = ""
    txtTasaPreferente.text = ""
    txtAlmacenaje.text = ""
    'txtDesempenoExt.text = ""
    txtPrestamo.text = ""
    txtCat.text = ""
    txtInteresAnual.text = ""
    txtAlmacenajeAnual.text = ""
    txtDiasMinimos.text = ""
    txtDiasGracias.text = ""
    
    
    grdTasas.ClearSelection
    cmbTipoInteres.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()

    Screen.MousePointer = vbHourglass
            
    'Creo los encabezados
    CrearEncabezados
    
    'Cargo los Datos
    opMetales.Value = True
    
    'Periodos
    Cargar_Combos "Descripcion", "tipoperiodo", cmbTipoPeriodo, , "Ordenamiento"
    
    'Vencimientos
    Cargar_Combos "Descripcion", "plazos", cmbPlazo, , "Descripcion"
    
    'TipoPersona
    cmbPersona.AddItem "FISICA"
    cmbPersona.ItemData(0) = 1
    
    cmbPersona.AddItem "MORAL"
    cmbPersona.ItemData(1) = 2
    
    cmbPersona.ListIndex = 0
    
    Poner_Flat fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    
    Screen.MousePointer = vbDefault
    
End Sub

Sub CrearEncabezados()

    With grdTasas
        .AddColumn "C1", "Tipo Interes", ecgHdrTextALignLeft, , 79, , , , , , , CCLSortString
        .AddColumn "C2", "Periodo", ecgHdrTextALignLeft, , 69, , , , , , , CCLSortString
        .AddColumn "C3", "Plazo", ecgHdrTextALignRight, , 45, , , , , , , CCLSortNumeric
        .AddColumn "C4", "T. Típica", ecgHdrTextALignRight, , 65, , , , , "0.000", , CCLSortNumeric
        .AddColumn "C5", "T. Prom", ecgHdrTextALignRight, , 65, , , , , "0.000", , CCLSortNumeric
        .AddColumn "C6", "T. Pref", ecgHdrTextALignRight, , 65, , , , , "0.000", , CCLSortNumeric
        .AddColumn "C7", "Alm", ecgHdrTextALignRight, , 65, , , , , "0.000", , CCLSortNumeric
        .AddColumn "C8", "Des Ext", ecgHdrTextALignRight, , 78, False, , , , "0.000", , CCLSortNumeric
        .AddColumn "C9", "%Prés", ecgHdrTextALignRight, , 65, , , , , , , CCLSortNumeric
        .AddColumn "C10", "D.Mín", ecgHdrTextALignRight, , 65, , , , , , , CCLSortNumeric
        .AddColumn "C11", "CAT", ecgHdrTextALignRight, , 70, , , , , "0.000", , CCLSortNumeric
        .AddColumn "C12", "Int Anual", ecgHdrTextALignRight, , 70, , , , , "0.000", , CCLSortNumeric
        .AddColumn "C13", "Alm Anual", ecgHdrTextALignRight, , 70, , , , , "0.000", , CCLSortNumeric
        .AddColumn "C14", "D. Gracia", ecgHdrTextALignRight, , 65, , , , , , , CCLSortNumeric
        .AddColumn "C15", "Tipo", ecgHdrTextALignRight, , 65, , , , , , , CCLSortString
    End With
    
End Sub

Sub CargarDatos()
Dim rcDatos As New ADODB.Recordset
Dim sFntUnread As New StdFont
    
    sFntUnread.Name = "Arial"
    sFntUnread.Size = 8
    sFntUnread.Bold = False
        
    rcDatos.Open "SELECT ti.Descripcion AS TipoInteres,tp.Descripcion AS TipoPeriodo,p.Descripcion AS Plazo, " _
        & "ct.ID AS IDConfiguracionTasa,ct.TasaTipica,ct.TasaPromocion,ct.TasaPreferencial,ct.IDTipoInteres,ct.IDTipoPeriodo,ct.IDPlazo,ct.PorPrestamo,ct.Cat,ct.Almacenaje,ct.Seguro,ct.DMinimos,ct.IntAnual,ct.AlmAnual,ct.DesExt,ct.DGracia,ct.Persona " _
        & "FROM tipointeres ti INNER JOIN configuraciontasas ct ON ti.ID=ct.IDTipoInteres INNER JOIN tipoperiodo tp ON tp.ID=ct.IDTipoPeriodo " _
        & "INNER JOIN plazos p ON p.ID=ct.IDPlazo WHERE Serie=" & Regresa_Serie _
        & " ORDER BY ti.Ordenamiento,tp.Ordenamiento,p.ID", dbDatos, adOpenForwardOnly, adLockReadOnly

    With grdTasas
        
        .Redraw = False
        .Clear
        While Not rcDatos.EOF
            .AddRow
            .CellText(.Rows, 1) = rcDatos!TipoInteres
            .CellItemData(.Rows, 1) = rcDatos!IDTipoInteres
            .CellTextAlign(.Rows, 1) = DT_LEFT
            .CellFont(.Rows, 1) = sFntUnread
            
            .CellText(.Rows, 2) = rcDatos!TipoPeriodo
            .CellItemData(.Rows, 2) = rcDatos!IDTipoPeriodo
            .CellTextAlign(.Rows, 2) = DT_LEFT
            .CellFont(.Rows, 2) = sFntUnread
            
            .CellText(.Rows, 3) = rcDatos!plazo
            .CellItemData(.Rows, 3) = rcDatos!IDPlazo
            .CellTextAlign(.Rows, 3) = DT_RIGHT
            .CellFont(.Rows, 3) = sFntUnread
            
            .CellText(.Rows, 4) = rcDatos!TasaTipica
            .CellItemData(.Rows, 4) = rcDatos!IDConfiguracionTasa
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            .CellFont(.Rows, 4) = sFntUnread
            
            .CellText(.Rows, 5) = rcDatos!TasaPromocion
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            .CellFont(.Rows, 5) = sFntUnread
            
            .CellText(.Rows, 6) = rcDatos!TasaPreferencial
            .CellTextAlign(.Rows, 6) = DT_RIGHT
            .CellFont(.Rows, 6) = sFntUnread
            
            .CellText(.Rows, 7) = rcDatos!Almacenaje
            .CellTextAlign(.Rows, 7) = DT_RIGHT
            .CellFont(.Rows, 7) = sFntUnread
            
'            .CellText(.Rows, 8) = rcDatos!DesExt
'            .CellTextAlign(.Rows, 8) = DT_RIGHT
'            .CellFont(.Rows, 8) = sFntUnread

            .CellText(.Rows, 8) = 0
            .CellTextAlign(.Rows, 8) = DT_RIGHT
            .CellFont(.Rows, 8) = sFntUnread
            
            .CellText(.Rows, 9) = rcDatos!PorPrestamo
            .CellTextAlign(.Rows, 9) = DT_RIGHT
            .CellFont(.Rows, 9) = sFntUnread
            
            .CellText(.Rows, 10) = rcDatos!DMinimos
            .CellTextAlign(.Rows, 10) = DT_RIGHT
            .CellFont(.Rows, 10) = sFntUnread
            
            .CellText(.Rows, 11) = rcDatos!Cat
            .CellTextAlign(.Rows, 11) = DT_RIGHT
            .CellFont(.Rows, 11) = sFntUnread
            
            .CellText(.Rows, 12) = rcDatos!IntAnual
            .CellTextAlign(.Rows, 12) = DT_RIGHT
            .CellFont(.Rows, 12) = sFntUnread
            
            .CellText(.Rows, 13) = rcDatos!AlmAnual
            .CellTextAlign(.Rows, 13) = DT_RIGHT
            .CellFont(.Rows, 13) = sFntUnread
            
            .CellText(.Rows, 14) = rcDatos!DGracia
            .CellTextAlign(.Rows, 14) = DT_RIGHT
            .CellFont(.Rows, 14) = sFntUnread
            
            .CellText(.Rows, 15) = IIf(opAutomovil.Value = True, IIf(rcDatos!Persona = 1, "FISICA", "MORAL"), 0)
            .CellItemData(.Rows, 15) = rcDatos!Persona
            .CellTextAlign(.Rows, 15) = DT_RIGHT
            .CellFont(.Rows, 15) = sFntUnread
            
            Colorea grdTasas, grdTasas.Rows, IIf(grdTasas.Rows Mod 2 > 0, RGB(236, 252, 222), RGB(255, 255, 255))
            
        rcDatos.MoveNext
        Wend
        
        .Redraw = True
    End With
    rcDatos.Close
    Set rcDatos = Nothing
End Sub

Private Sub grdTasas_DblClick(ByVal lRow As Long, ByVal lCol As Long)

    If grdTasas.Rows > 0 Then
        
        If grdTasas.SelectedRow > 0 Then
                
                txtTasaTipica.text = grdTasas.CellText(grdTasas.SelectedRow, 4)
                txtTasaTipica.Tag = grdTasas.CellItemData(grdTasas.SelectedRow, 4)
                txtTasaPromocion.text = grdTasas.CellText(grdTasas.SelectedRow, 5)
                txtTasaPreferente.text = grdTasas.CellText(grdTasas.SelectedRow, 6)
                txtAlmacenaje.text = grdTasas.CellText(grdTasas.SelectedRow, 7)
                'txtDesempenoExt.text = grdTasas.CellText(grdTasas.SelectedRow, 8)
                
                txtPrestamo.text = grdTasas.CellText(grdTasas.SelectedRow, 9)
                txtDiasMinimos.text = grdTasas.CellText(grdTasas.SelectedRow, 10)
                
                txtCat.text = grdTasas.CellText(grdTasas.SelectedRow, 11)
                txtInteresAnual.text = grdTasas.CellText(grdTasas.SelectedRow, 12)
                txtAlmacenajeAnual.text = grdTasas.CellText(grdTasas.SelectedRow, 13)
                txtDiasGracias.text = grdTasas.CellText(grdTasas.SelectedRow, 14)
                
                cmbTipoInteres.ListIndex = ComboInformacion(cmbTipoInteres, grdTasas.CellItemData(grdTasas.SelectedRow, 1))
                cmbTipoPeriodo.ListIndex = ComboInformacion(cmbTipoPeriodo, grdTasas.CellItemData(grdTasas.SelectedRow, 2))
                cmbPlazo.ListIndex = ComboInformacion(cmbPlazo, grdTasas.CellItemData(grdTasas.SelectedRow, 3))
                
                If opAutomovil.Value = True Then
                    cmbPersona.ListIndex = grdTasas.CellItemData(grdTasas.SelectedRow, 15) - 1
                End If
                
            grdTasas.ClearSelection
            cmbTipoInteres.SetFocus
        End If
        
    End If

End Sub

Private Sub opAutomovil_Click()

    lblDesExt.Visible = True
    cmbPersona.Visible = True

    cmbPersona.ListIndex = 0

    Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Serie=" & Regresa_Serie, "Ordenamiento"
    CargarDatos
End Sub

Private Sub opElectronicos_Click()

    lblDesExt.Visible = False
    cmbPersona.Visible = False

    Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Serie=" & Regresa_Serie, "Ordenamiento"
    CargarDatos
End Sub

Private Sub opMetales_Click()

    lblDesExt.Visible = False
    cmbPersona.Visible = False

    Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Serie=" & Regresa_Serie, "Ordenamiento"
    CargarDatos
End Sub

Private Sub txtAlmacenaje_GotFocus()
    Seleccionar_Texto txtAlmacenaje
    Cambiar_Color True, txtAlmacenaje
End Sub

Private Sub txtAlmacenaje_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAlmacenaje_LostFocus()
    Cambiar_Color False, txtAlmacenaje
End Sub

Private Sub txtCat_GotFocus()
    Seleccionar_Texto txtCat
    Cambiar_Color True, txtCat
End Sub

Private Sub txtCat_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCat_LostFocus()
    Cambiar_Color False, txtCat
End Sub

Private Sub txtDesempenoExt_GotFocus()
    Seleccionar_Texto txtDesempenoExt
    Cambiar_Color True, txtDesempenoExt
End Sub

Private Sub txtDesempenoExt_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDesempenoExt_LostFocus()
    Cambiar_Color False, txtDesempenoExt
End Sub

Private Sub txtDiasGracias_GotFocus()
    Seleccionar_Texto txtDiasGracias
    Cambiar_Color True, txtDiasGracias
End Sub

Private Sub txtDiasGracias_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDiasGracias_LostFocus()
    Cambiar_Color False, txtDiasGracias
End Sub

Private Sub txtDiasMinimos_GotFocus()
    Seleccionar_Texto txtDiasMinimos
    Cambiar_Color True, txtDiasMinimos
End Sub

Private Sub txtDiasMinimos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDiasMinimos_LostFocus()
    Cambiar_Color False, txtDiasMinimos
End Sub

Private Sub txtInteresAnual_GotFocus()
    Seleccionar_Texto txtInteresAnual
    Cambiar_Color True, txtInteresAnual
End Sub

Private Sub txtInteresAnual_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtInteresAnual_LostFocus()
    Cambiar_Color False, txtInteresAnual
End Sub

Private Sub txtPrestamo_GotFocus()
    Seleccionar_Texto txtPrestamo
    Cambiar_Color True, txtPrestamo
End Sub

Private Sub txtPrestamo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamo_LostFocus()
    Cambiar_Color False, txtPrestamo
End Sub

Private Sub txtAlmacenajeAnual_GotFocus()
    Seleccionar_Texto txtAlmacenajeAnual
    Cambiar_Color True, txtAlmacenajeAnual
End Sub

Private Sub txtAlmacenajeAnual_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAlmacenajeAnual_LostFocus()
    Cambiar_Color False, txtAlmacenajeAnual
End Sub

Private Sub txtTasaPreferente_GotFocus()
    Seleccionar_Texto txtTasaPreferente
    Cambiar_Color True, txtTasaPreferente
End Sub

Private Sub txtTasaPreferente_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTasaPreferente_LostFocus()
    Cambiar_Color False, txtTasaPreferente
End Sub

Private Sub txtTasaPromocion_GotFocus()
    Seleccionar_Texto txtTasaPromocion
    Cambiar_Color True, txtTasaPromocion
End Sub

Private Sub txtTasaPromocion_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTasaPromocion_LostFocus()
    Cambiar_Color False, txtTasaPromocion
End Sub

Private Sub txtTasaTipica_GotFocus()
    Seleccionar_Texto txtTasaTipica
    Cambiar_Color True, txtTasaTipica
End Sub

Private Sub txtTasaTipica_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTasaTipica_LostFocus()
    Cambiar_Color False, txtTasaTipica
End Sub

Function Regresa_Serie() As Integer
    If opMetales.Value = True Then Regresa_Serie = SERIE_A
    If opAutomovil.Value = True Then Regresa_Serie = SERIE_B
    If opElectronicos.Value = True Then Regresa_Serie = SERIE_D
End Function


Function DatosValidos() As Boolean
Dim Tasa As Double, Almacenaje As Double, Seguro As Double

    DatosValidos = True
    
    If cmbTipoInteres.ListIndex = -1 Then
        DatosValidos = False
        cmbTipoInteres.SetFocus
        Exit Function
    End If
    
    If cmbTipoPeriodo.ListIndex = -1 Then
        DatosValidos = False
        cmbTipoPeriodo.SetFocus
        Exit Function
    End If
    
    If cmbPlazo.ListIndex = -1 Then
        DatosValidos = False
        cmbPlazo.SetFocus
        Exit Function
    End If
    
    If Trim(txtPrestamo.text) = "" Then
        DatosValidos = False
        txtPrestamo.SetFocus
        Exit Function
    End If
    
    If Trim(txtTasaTipica.text) = "" Then
        DatosValidos = False
        txtTasaTipica.SetFocus
        Exit Function
    End If
    
    If Trim(txtTasaPromocion.text) = "" Then
        DatosValidos = False
        txtTasaPromocion.SetFocus
        Exit Function
    End If
    
    If Trim(txtTasaPreferente.text) = "" Then
        DatosValidos = False
        txtTasaPreferente.SetFocus
        Exit Function
    End If
    
    If Trim(txtCat.text) = "" Then
        DatosValidos = False
        txtCat.SetFocus
        Exit Function
    End If
    
    If Trim(txtAlmacenaje.text) = "" Then
        DatosValidos = False
        txtAlmacenaje.SetFocus
        Exit Function
    End If
    
    If Trim(txtDiasMinimos.text) = "" Then
        DatosValidos = False
        txtDiasMinimos.SetFocus
        Exit Function
    End If
    
    If Trim(txtInteresAnual.text) = "" Then
        DatosValidos = False
        txtInteresAnual.SetFocus
        Exit Function
    End If
    
    If Trim(txtAlmacenajeAnual.text) = "" Then
        DatosValidos = False
        txtAlmacenajeAnual.SetFocus
        Exit Function
    End If
    
'    If Trim(txtDesempenoExt.text) = "" Then
'        DatosValidos = False
'        txtDesempenoExt.SetFocus
'        Exit Function
'    End If
    
    If Trim(txtDiasGracias.text) = "" Then
        DatosValidos = False
        txtDiasGracias.SetFocus
        Exit Function
    End If
    
End Function
