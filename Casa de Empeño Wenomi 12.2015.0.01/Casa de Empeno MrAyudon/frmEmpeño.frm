VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{1781610F-46E8-4DD3-922D-8DEF1A9DA567}#28.0#0"; "Credencial.ocx"
Begin VB.Form frmEmpeño 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Empeño "
   ClientHeight    =   10200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12825
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEmpeño.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10200
   ScaleWidth      =   12825
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1440
      Top             =   9600
   End
   Begin VB.TextBox txtNoTarjeta 
      BackColor       =   &H0000FFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   8280
      MaxLength       =   60
      TabIndex        =   213
      Top             =   240
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   2040
      Top             =   9600
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCotizar 
      Height          =   375
      Left            =   5250
      TabIndex        =   132
      Top             =   9600
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "    &Cotización"
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
      Picture         =   "frmEmpeño.frx":000C
      PictureDisabled =   "frmEmpeño.frx":020C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   11580
      TabIndex        =   149
      Top             =   9600
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
      Picture         =   "frmEmpeño.frx":0366
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   10320
      TabIndex        =   150
      Top             =   9600
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   8537065
      Object.ToolTipText     =   ""
      Picture         =   "frmEmpeño.frx":08B8
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   9090
      TabIndex        =   151
      Top             =   9600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Re-Imprimir"
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
      Picture         =   "frmEmpeño.frx":0E0A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdFoto 
      Height          =   375
      Left            =   6570
      TabIndex        =   152
      Top             =   9600
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
      Picture         =   "frmEmpeño.frx":135C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelacion 
      Height          =   375
      Left            =   7890
      TabIndex        =   153
      Top             =   9600
      Width           =   1155
      _ExtentX        =   2037
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
      Picture         =   "frmEmpeño.frx":1831
      PictureDisabled =   "frmEmpeño.frx":1A80
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   4080
      TabIndex        =   192
      Top             =   9600
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
      Picture         =   "frmEmpeño.frx":2652
   End
   Begin DevPowerFlatBttn.FlatBttn cmdPagosFijos 
      Height          =   375
      Left            =   2760
      TabIndex        =   193
      Top             =   9600
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Pagos Fijos"
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
      Picture         =   "frmEmpeño.frx":29D7
   End
   Begin vbalTabStrip6.TabControl TPestañas 
      Height          =   8655
      Left            =   120
      TabIndex        =   72
      Top             =   720
      Width           =   12720
      _ExtentX        =   22437
      _ExtentY        =   15266
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FlatSeparators  =   -1  'True
      FlatButtons     =   -1  'True
      CoolTabs        =   1
      Begin VB.Frame frmAutomoviles 
         Caption         =   "Autos"
         Height          =   8070
         Left            =   120
         TabIndex        =   94
         Top             =   480
         Visible         =   0   'False
         Width           =   12555
         Begin VB.CheckBox chkAutoencirculacion 
            Caption         =   "Auto en circulación"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8280
            TabIndex        =   282
            Top             =   3000
            Width           =   2655
         End
         Begin VB.TextBox txtCostoMensualSeguro 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
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
            Left            =   10440
            TabIndex        =   274
            Top             =   2590
            Width           =   1575
         End
         Begin Line3D.ucLine3D ucLine3D5 
            Height          =   30
            Left            =   8160
            Top             =   2520
            Width           =   4140
            _ExtentX        =   7303
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin VB.CommandButton cmdHistorial2 
            Caption         =   "Historial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5280
            TabIndex        =   239
            Top             =   435
            Width           =   990
         End
         Begin VB.CommandButton cmdAlerta2 
            Caption         =   "Selec. Alerta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6375
            TabIndex        =   238
            ToolTipText     =   "Selección de Alerta para Ley Anti-lavado de Dinero."
            Top             =   435
            Width           =   1230
         End
         Begin VB.CommandButton cmdEditarCliente2 
            Caption         =   "Editar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4215
            TabIndex        =   237
            Top             =   435
            Width           =   990
         End
         Begin VB.TextBox txtCotitularApellidoMaterno2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   3600
            MaxLength       =   60
            TabIndex        =   33
            Top             =   3150
            Width           =   3255
         End
         Begin VB.TextBox txtCotitularApellidoPaterno2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   135
            MaxLength       =   60
            TabIndex        =   32
            Top             =   3150
            Width           =   3255
         End
         Begin VB.TextBox txtResponsable2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   120
            MaxLength       =   30
            TabIndex        =   31
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox txtApellidoMaterno2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   3615
            MaxLength       =   60
            TabIndex        =   30
            Top             =   1065
            Width           =   3255
         End
         Begin VB.TextBox txtApellidoPaterno2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   150
            MaxLength       =   60
            TabIndex        =   29
            Top             =   1065
            Width           =   3255
         End
         Begin VB.TextBox txtNombre2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   135
            MaxLength       =   20
            TabIndex        =   28
            Top             =   465
            Width           =   3255
         End
         Begin VB.CommandButton cmdEditarCotitular2 
            Caption         =   "Editar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4080
            TabIndex        =   236
            Top             =   2475
            Width           =   990
         End
         Begin VB.ComboBox cmbPromocion2 
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
            ItemData        =   "frmEmpeño.frx":2BC7
            Left            =   10260
            List            =   "frmEmpeño.frx":2BDD
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   2175
            Width           =   1980
         End
         Begin VB.ComboBox cmbPlazos2 
            BackColor       =   &H0000FFFF&
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
            ItemData        =   "frmEmpeño.frx":2C25
            Left            =   11445
            List            =   "frmEmpeño.frx":2C27
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1800
            Width           =   870
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   390
            Index           =   1
            Left            =   11385
            Top             =   1785
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   688
            Orientation     =   0
            LineWidth       =   2
         End
         Begin VB.ComboBox cmbPeriodo2 
            BackColor       =   &H0000FFFF&
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
            ItemData        =   "frmEmpeño.frx":2C29
            Left            =   9930
            List            =   "frmEmpeño.frx":2C2B
            Style           =   2  'Dropdown List
            TabIndex        =   36
            Top             =   1800
            Width           =   1455
         End
         Begin VB.ComboBox cmbTipoInteres2 
            BackColor       =   &H0000FFFF&
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
            Left            =   8295
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1800
            Width           =   1560
         End
         Begin VB.TextBox txtPrestamo2 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """$""#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2058
               SubFormatType   =   2
            EndProperty
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   10485
            MaxLength       =   12
            TabIndex        =   34
            Top             =   1268
            Width           =   1605
         End
         Begin VB.TextBox txtNotas2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   4920
            MaxLength       =   250
            TabIndex        =   60
            Top             =   7650
            Width           =   4830
         End
         Begin VB.ComboBox cmbMedio2 
            Appearance      =   0  'Flat
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
            ItemData        =   "frmEmpeño.frx":2C2D
            Left            =   9840
            List            =   "frmEmpeño.frx":2C2F
            Style           =   2  'Dropdown List
            TabIndex        =   95
            Top             =   7575
            Width           =   2595
         End
         Begin VB.TextBox txtMensaje2 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   120
            MaxLength       =   150
            TabIndex        =   59
            Top             =   7650
            Width           =   4665
         End
         Begin Line3D.ucLine3D ucLine3D4 
            Height          =   135
            Left            =   0
            Top             =   3480
            Width           =   12300
            _ExtentX        =   21696
            _ExtentY        =   238
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMosCliente2 
            Height          =   225
            Left            =   3645
            TabIndex        =   111
            Top             =   480
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   397
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
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   22
            Left            =   8235
            Top             =   105
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   345
            Index           =   24
            Left            =   10200
            Top             =   105
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   609
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   26
            Left            =   8235
            Top             =   435
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   285
            Index           =   27
            Left            =   11520
            Top             =   720
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   503
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   285
            Index           =   28
            Left            =   10500
            Top             =   720
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   503
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   285
            Index           =   29
            Left            =   9180
            Top             =   720
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   503
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   30
            Left            =   8235
            Top             =   975
            Width           =   4110
            _ExtentX        =   7250
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   34
            Left            =   8235
            Top             =   720
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   330
            Index           =   35
            Left            =   10245
            Top             =   1245
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   582
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   38
            Left            =   8235
            Top             =   1530
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   40
            Left            =   8235
            Top             =   1230
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   2
            Left            =   9870
            Top             =   1785
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   2880
            Index           =   41
            Left            =   12360
            Top             =   120
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   5080
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   42
            Left            =   8160
            Top             =   2955
            Width           =   4260
            _ExtentX        =   7514
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   43
            Left            =   8250
            Top             =   1785
            Width           =   4095
            _ExtentX        =   7223
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D25 
            Height          =   2895
            Index           =   1
            Left            =   8160
            Top             =   120
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   5106
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   10
            Left            =   8235
            Top             =   2160
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   4
            Left            =   10200
            Top             =   2160
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin VB.Frame Frame1 
            Caption         =   "DATOS DEL AUTOMÓVIL"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3420
            Left            =   0
            TabIndex        =   96
            Top             =   3720
            Width           =   12375
            Begin VB.TextBox txtVIN 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   9720
               MaxLength       =   17
               TabIndex        =   46
               Top             =   495
               Width           =   2400
            End
            Begin VB.ComboBox cmbTipoBlindaje 
               Height          =   315
               Left            =   4020
               Style           =   2  'Dropdown List
               TabIndex        =   53
               Top             =   1650
               Width           =   2040
            End
            Begin VB.TextBox txtREPUVE 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   6240
               MaxLength       =   8
               TabIndex        =   50
               Top             =   1095
               Width           =   1680
            End
            Begin VB.TextBox txtModelo 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   2055
               MaxLength       =   50
               TabIndex        =   40
               Top             =   510
               Width           =   1800
            End
            Begin VB.TextBox txtNumMotor 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   2055
               MaxLength       =   17
               TabIndex        =   48
               Top             =   1095
               Width           =   1800
            End
            Begin VB.TextBox txtTipoPoliza 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   6240
               MaxLength       =   50
               TabIndex        =   57
               Top             =   2250
               Width           =   1680
            End
            Begin VB.TextBox txtPoliza 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   2055
               MaxLength       =   50
               TabIndex        =   55
               Top             =   2250
               Width           =   1800
            End
            Begin VB.TextBox txtAseguradora 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   135
               MaxLength       =   50
               TabIndex        =   54
               Top             =   2250
               Width           =   1800
            End
            Begin VB.TextBox txtGas 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   135
               MaxLength       =   50
               TabIndex        =   51
               Top             =   1695
               Width           =   1800
            End
            Begin VB.TextBox txtKms 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   8640
               MaxLength       =   20
               TabIndex        =   45
               Top             =   510
               Width           =   975
            End
            Begin VB.TextBox txtSerieChasis 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   4020
               MaxLength       =   17
               TabIndex        =   49
               Top             =   1095
               Width           =   2040
            End
            Begin VB.TextBox txtTarjeta 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   2055
               MaxLength       =   35
               TabIndex        =   52
               Top             =   1695
               Width           =   1800
            End
            Begin VB.TextBox txtAgencia 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   120
               MaxLength       =   50
               TabIndex        =   47
               Top             =   1110
               Width           =   1800
            End
            Begin VB.TextBox txtFactura 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   7335
               MaxLength       =   50
               TabIndex        =   44
               Top             =   510
               Width           =   1095
            End
            Begin VB.TextBox txtPlacas 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   6240
               MaxLength       =   50
               TabIndex        =   43
               Top             =   510
               Width           =   975
            End
            Begin VB.TextBox txtColor 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   5040
               MaxLength       =   50
               TabIndex        =   42
               Top             =   510
               Width           =   1095
            End
            Begin VB.TextBox txtAño 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   4020
               MaxLength       =   5
               TabIndex        =   41
               Top             =   510
               Width           =   855
            End
            Begin VB.TextBox txtMarca 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   135
               MaxLength       =   50
               TabIndex        =   39
               Top             =   510
               Width           =   1800
            End
            Begin VB.TextBox txtObservaciones2 
               BorderStyle     =   0  'None
               Height          =   360
               Left            =   135
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   58
               Top             =   2880
               Width           =   8085
            End
            Begin VB.Frame Frame2 
               Caption         =   "DOCUMENTOS ENTREGADOS"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2010
               Left            =   8295
               TabIndex        =   97
               Top             =   990
               Width           =   4005
               Begin VB.CheckBox chkImportacion 
                  Appearance      =   0  'Flat
                  Caption         =   "Importación"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   2535
                  TabIndex        =   110
                  Top             =   600
                  Width           =   1455
               End
               Begin VB.CheckBox chkCopyLicencia 
                  Appearance      =   0  'Flat
                  Caption         =   "Copia de licencia"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   135
                  TabIndex        =   103
                  Top             =   1515
                  Width           =   2175
               End
               Begin VB.CheckBox chkPoliza 
                  Appearance      =   0  'Flat
                  Caption         =   "Póliza de seguro"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   135
                  TabIndex        =   102
                  Top             =   1215
                  Width           =   2055
               End
               Begin VB.CheckBox chkTenencia 
                  Appearance      =   0  'Flat
                  Caption         =   "Tenencias (Últimas 5)"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   135
                  TabIndex        =   101
                  Top             =   915
                  Width           =   2235
               End
               Begin VB.CheckBox chkCopiIfe 
                  Appearance      =   0  'Flat
                  Caption         =   "Copia IFE"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   135
                  TabIndex        =   100
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.CheckBox chkTarjeta 
                  Appearance      =   0  'Flat
                  Caption         =   "Tarjeta de circulación"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   135
                  TabIndex        =   99
                  Top             =   300
                  Width           =   2130
               End
               Begin VB.CheckBox chkFactura 
                  Appearance      =   0  'Flat
                  Caption         =   "Factura"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   255
                  Left            =   2535
                  TabIndex        =   98
                  Top             =   300
                  Width           =   975
               End
            End
            Begin MSMask.MaskEdBox txtFechaVenciPoliza 
               Height          =   240
               Left            =   4020
               TabIndex        =   56
               Top             =   2250
               Width           =   2040
               _ExtentX        =   3598
               _ExtentY        =   423
               _Version        =   393216
               BorderStyle     =   0
               MaxLength       =   10
               Format          =   "dd-mmmm-yyyy"
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label19 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "VIN (N° Id. Vehículo)"
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
               Left            =   9720
               TabIndex        =   272
               Top             =   255
               Width           =   2370
            End
            Begin VB.Label Label16 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo Blindaje"
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
               Left            =   4020
               TabIndex        =   271
               Top             =   1455
               Width           =   1170
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "REPUVE"
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
               Left            =   6240
               TabIndex        =   270
               Top             =   855
               Width           =   1680
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Modelo"
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
               Left            =   2100
               TabIndex        =   269
               Top             =   255
               Width           =   1800
            End
            Begin VB.Label Label98 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Número de motor"
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
               Left            =   2100
               TabIndex        =   268
               Top             =   855
               Width           =   1800
            End
            Begin VB.Label Label58 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tipo poliza"
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
               Left            =   6240
               TabIndex        =   267
               Top             =   2010
               Width           =   1680
            End
            Begin VB.Label Label109 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha venc."
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
               Left            =   4020
               TabIndex        =   266
               Top             =   2010
               Width           =   2040
            End
            Begin VB.Label Label103 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Póliza"
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
               Left            =   2100
               TabIndex        =   265
               Top             =   2010
               Width           =   1800
            End
            Begin VB.Label Label102 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Aseguradora"
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
               Left            =   135
               TabIndex        =   264
               Top             =   2010
               Width           =   1800
            End
            Begin VB.Label Label101 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gas"
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
               Left            =   135
               TabIndex        =   263
               Top             =   1455
               Width           =   1800
            End
            Begin VB.Label Label100 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Kms."
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
               Left            =   8640
               TabIndex        =   262
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label99 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Serie chasis"
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
               Left            =   4020
               TabIndex        =   261
               Top             =   855
               Width           =   2040
            End
            Begin VB.Label Label94 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tarjeta de circu."
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
               Left            =   2100
               TabIndex        =   260
               Top             =   1455
               Width           =   1800
            End
            Begin VB.Label Label49 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Agencia"
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
               Left            =   135
               TabIndex        =   259
               Top             =   840
               Width           =   1800
            End
            Begin VB.Label Label48 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Factura"
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
               Left            =   7335
               TabIndex        =   258
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label47 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Placas"
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
               Left            =   6240
               TabIndex        =   257
               Top             =   240
               Width           =   975
            End
            Begin VB.Label Label46 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Color"
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
               Left            =   5040
               TabIndex        =   256
               Top             =   240
               Width           =   1095
            End
            Begin VB.Label Label45 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Año"
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
               Left            =   4020
               TabIndex        =   255
               Top             =   240
               Width           =   855
            End
            Begin VB.Label Label42 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Marca"
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
               Left            =   135
               TabIndex        =   254
               Top             =   240
               Width           =   1800
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Observaciones:"
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
               Left            =   150
               TabIndex        =   253
               Top             =   2640
               Width           =   1935
            End
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMosCotitular2 
            Height          =   225
            Left            =   3615
            TabIndex        =   240
            Top             =   2475
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   397
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
         Begin Line3D.ucLine3D ucLine3D2 
            Height          =   135
            Left            =   135
            Top             =   1995
            Width           =   7980
            _ExtentX        =   14076
            _ExtentY        =   238
         End
         Begin VB.Label Label89 
            AutoSize        =   -1  'True
            Caption         =   "Costo Poliza Seguro:"
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
            Index           =   100
            Left            =   8280
            TabIndex        =   275
            Top             =   2640
            Width           =   1995
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Direccion"
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
            Left            =   135
            TabIndex        =   252
            Top             =   1395
            Width           =   1155
         End
         Begin VB.Label lblRFC2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6015
            TabIndex        =   251
            Top             =   1755
            Width           =   2085
         End
         Begin VB.Label lblCiudad2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1335
            TabIndex        =   250
            Top             =   1755
            Width           =   3525
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Ciudad"
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
            Left            =   135
            TabIndex        =   249
            Top             =   1755
            Width           =   1155
         End
         Begin VB.Label Label14 
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
            Height          =   210
            Left            =   4935
            TabIndex        =   248
            Top             =   1755
            Width           =   1065
         End
         Begin VB.Label lblDireccion2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1335
            TabIndex        =   247
            Top             =   1395
            Width           =   6765
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Materno"
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
            Index           =   3
            Left            =   3600
            TabIndex        =   246
            Top             =   2925
            Width           =   3255
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Paterno"
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
            Index           =   2
            Left            =   135
            TabIndex        =   245
            Top             =   2925
            Width           =   3255
         End
         Begin VB.Label Label93 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cotitular"
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
            TabIndex        =   244
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Paterno"
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
            Left            =   150
            TabIndex        =   243
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label41 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Materno"
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
            Index           =   1
            Left            =   3615
            TabIndex        =   242
            Top             =   840
            Width           =   3255
         End
         Begin VB.Label Label40 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Nombre"
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
            Left            =   150
            TabIndex        =   241
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Promoción:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   6
            Left            =   8520
            TabIndex        =   202
            Top             =   2190
            Width           =   1485
         End
         Begin VB.Label lbllabeltipo 
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   3
            Left            =   8280
            TabIndex        =   201
            Top             =   2160
            Width           =   1920
         End
         Begin VB.Label lblVencimiento2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<Vencimiento>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   10770
            TabIndex        =   189
            Top             =   165
            Width           =   1800
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vence:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   2
            Left            =   10245
            TabIndex        =   188
            Top             =   135
            Width           =   780
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   4
            Left            =   10305
            TabIndex        =   187
            Top             =   1560
            Width           =   705
         End
         Begin VB.Label lblTasa2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Tasa>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   8370
            TabIndex        =   148
            Top             =   750
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Interes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   8400
            TabIndex        =   146
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%Tasa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   8400
            TabIndex        =   145
            Top             =   465
            Width           =   645
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   11640
            TabIndex        =   144
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label lblTotAvaluo2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Avaluo>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   8925
            TabIndex        =   143
            Top             =   1275
            Width           =   930
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Préstamo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   10560
            TabIndex        =   141
            Top             =   1005
            Width           =   1395
         End
         Begin VB.Label Label81 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Avalúo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   8280
            TabIndex        =   140
            Top             =   1005
            Width           =   1995
         End
         Begin VB.Label lblIva2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<IVA>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   11640
            TabIndex        =   138
            Top             =   750
            Width           =   570
         End
         Begin VB.Label Label84 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%Seguro"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   10575
            TabIndex        =   137
            Top             =   465
            Width           =   870
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%I.V.A."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   11565
            TabIndex        =   136
            Top             =   465
            Width           =   720
         End
         Begin VB.Label Label83 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%Almacenaje"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   9240
            TabIndex        =   135
            Top             =   465
            Width           =   1245
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
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
            Index           =   1
            Left            =   6180
            TabIndex        =   134
            Top             =   165
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Contrato:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   8280
            TabIndex        =   133
            Top             =   135
            Width           =   1050
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Notas impresas en contrato:"
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
            Left            =   4920
            TabIndex        =   115
            Top             =   7320
            Width           =   2595
         End
         Begin VB.Label lblFolio2 
            AutoSize        =   -1  'True
            Caption         =   "<Folio>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   9345
            TabIndex        =   109
            Top             =   135
            Width           =   915
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "<Fecha>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   240
            Index           =   4
            Left            =   6810
            TabIndex        =   108
            Top             =   165
            Visible         =   0   'False
            Width           =   885
         End
         Begin VB.Label Label114 
            AutoSize        =   -1  'True
            Caption         =   "Mensaje interno:"
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
            TabIndex        =   107
            Top             =   7320
            Width           =   1545
         End
         Begin VB.Label Label113 
            AutoSize        =   -1  'True
            Caption         =   "Como se enteró:"
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
            Left            =   9855
            TabIndex        =   106
            Top             =   7320
            Width           =   1515
         End
         Begin VB.Label lblAlmacenaje2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<Almacenaje>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   9255
            TabIndex        =   105
            Top             =   750
            Width           =   1215
         End
         Begin VB.Label lblSeguro2 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<Seguro>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   10575
            TabIndex        =   104
            Top             =   750
            Width           =   855
         End
         Begin VB.Label Label51 
            BackColor       =   &H00404040&
            Height          =   285
            Index           =   1
            Left            =   8250
            TabIndex        =   139
            Top             =   465
            Width           =   4095
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Height          =   240
            Index           =   1
            Left            =   8160
            TabIndex        =   142
            Top             =   1005
            Width           =   4095
         End
         Begin VB.Label lblCostoMensual 
            BackColor       =   &H00404040&
            Height          =   285
            Index           =   1
            Left            =   8280
            TabIndex        =   147
            Top             =   1530
            Width           =   4095
         End
      End
      Begin VB.Frame frmDesempeño 
         Caption         =   "DESEMPEÑO"
         Height          =   8055
         Left            =   120
         TabIndex        =   74
         Top             =   480
         Visible         =   0   'False
         Width           =   12555
         Begin VB.Frame Frame4 
            Caption         =   "GPS y Seguro Auto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   1815
            Left            =   5400
            TabIndex        =   283
            Top             =   3360
            Width           =   4455
            Begin VB.TextBox txtCargoGPSDes 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   2
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2760
               TabIndex        =   286
               Top             =   720
               Width           =   1335
            End
            Begin VB.TextBox txtCargoSeguroDes 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   2
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2760
               TabIndex        =   285
               Top             =   1200
               Width           =   1335
            End
            Begin VB.CheckBox chkCirculacionDes 
               Caption         =   "En circulación"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2040
               TabIndex        =   284
               Top             =   360
               Width           =   2055
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cargo Renta GPS:"
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
               Left            =   960
               TabIndex        =   288
               Top             =   840
               Width           =   1635
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cargo por Seguro de Auto:"
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
               TabIndex        =   287
               Top             =   1320
               Width           =   2505
            End
         End
         Begin VB.CheckBox chkAutomovil 
            Appearance      =   0  'Flat
            Caption         =   "Automóvil"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   2640
            TabIndex        =   91
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtFolioDesempeño 
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
            Left            =   1320
            TabIndex        =   75
            Top             =   240
            Width           =   1140
         End
         Begin vbAcceleratorGrid6.vbalGrid grdDesempeño 
            Height          =   2880
            Left            =   5400
            TabIndex        =   77
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   5080
            RowMode         =   -1  'True
            GridLines       =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BackColor       =   16777215
            GridLineColor   =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderFlat      =   -1  'True
            BorderStyle     =   2
            ScrollBarStyle  =   1
            Editable        =   -1  'True
            DisableIcons    =   -1  'True
            DefaultRowHeight=   25
            Begin VB.TextBox txtEdit2 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   840
               TabIndex        =   78
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin Credencial.usCredencial DatosCliente 
            Height          =   2560
            Index           =   0
            Left            =   45
            TabIndex        =   209
            Top             =   615
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   4524
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   6
            AlingHeader     =   0
            AlingBody       =   0
            BodyIndent      =   0
            HeaderIndent    =   0
            HeaderText      =   " Datos del cliente"
            HeaderBackColor =   16766131
            HeightHeader    =   22
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   25
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   16744576
            BackColor       =   16777215
         End
         Begin Credencial.usCredencial DatosContrato 
            Height          =   2010
            Index           =   0
            Left            =   45
            TabIndex        =   210
            Top             =   3195
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   3545
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   5
            AlingHeader     =   0
            AlingBody       =   0
            BodyIndent      =   0
            HeaderIndent    =   0
            HeaderText      =   " Datos del contrato"
            HeaderBackColor =   16766131
            HeightHeader    =   22
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   25
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   16744576
            BackColor       =   16777215
         End
         Begin Credencial.usCredencial DetallesContrato 
            Height          =   1830
            Index           =   0
            Left            =   45
            TabIndex        =   211
            Top             =   5220
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   3228
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   5
            AlingHeader     =   0
            AlingBody       =   0
            BodyIndent      =   0
            HeaderIndent    =   0
            HeaderText      =   " Detalle del contrato"
            HeaderBackColor =   16766131
            HeightHeader    =   22
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   25
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   16744576
            BackColor       =   16777215
         End
         Begin VB.Label labelContratoDesemp 
            Caption         =   "CONTRATO EN ALMONEDA!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   975
            Left            =   5520
            TabIndex        =   273
            Top             =   5520
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label Leyenda 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   570
            Left            =   12285
            TabIndex        =   80
            Top             =   5040
            Width           =   165
         End
         Begin VB.Label TotalDesempeño 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   39.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   12120
            TabIndex        =   79
            Top             =   5580
            Width           =   270
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Contrato:"
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
            TabIndex        =   76
            Top             =   240
            Width           =   1170
         End
      End
      Begin VB.Frame frmRefrendos 
         Caption         =   "REFRENDOS"
         Height          =   8055
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Visible         =   0   'False
         Width           =   12555
         Begin VB.Frame Frame3 
            Caption         =   "GPS y Seguro Auto"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000001&
            Height          =   1815
            Left            =   5400
            TabIndex        =   276
            Top             =   3240
            Width           =   4455
            Begin VB.CheckBox chkCirculacionRef 
               Caption         =   "En circulación"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   2040
               TabIndex        =   281
               Top             =   360
               Width           =   2055
            End
            Begin VB.TextBox txtCargoSeguro 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   2
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2760
               TabIndex        =   280
               Top             =   1200
               Width           =   1335
            End
            Begin VB.TextBox txtCargoGPS 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty DataFormat 
                  Type            =   1
                  Format          =   """$""#,##0.00"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   2058
                  SubFormatType   =   2
               EndProperty
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   2760
               TabIndex        =   279
               Top             =   720
               Width           =   1335
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cargo por Seguro de Auto:"
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
               TabIndex        =   278
               Top             =   1320
               Width           =   2505
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cargo Renta GPS:"
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
               Left            =   960
               TabIndex        =   277
               Top             =   840
               Width           =   1635
            End
         End
         Begin Credencial.usCredencial DatosCliente 
            Height          =   2560
            Index           =   1
            Left            =   45
            TabIndex        =   207
            Top             =   615
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   4524
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   6
            AlingHeader     =   0
            AlingBody       =   0
            BodyIndent      =   0
            HeaderIndent    =   0
            HeaderText      =   " Datos del cliente"
            HeaderBackColor =   16766131
            HeightHeader    =   22
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   25
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   16744576
            BackColor       =   16777215
         End
         Begin VB.CheckBox chkAutomovil 
            Appearance      =   0  'Flat
            Caption         =   "Automóvil"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   2640
            TabIndex        =   92
            Top             =   240
            Width           =   1575
         End
         Begin VB.TextBox txtFolioRefrendo 
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
            Left            =   1320
            TabIndex        =   71
            Top             =   240
            Width           =   1140
         End
         Begin vbAcceleratorGrid6.vbalGrid grdRefrendos 
            Height          =   2880
            Left            =   5400
            TabIndex        =   83
            Top             =   360
            Width           =   7215
            _ExtentX        =   12726
            _ExtentY        =   5080
            RowMode         =   -1  'True
            GridLines       =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BackColor       =   16777215
            GridLineColor   =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderFlat      =   -1  'True
            BorderStyle     =   2
            ScrollBarStyle  =   1
            Editable        =   -1  'True
            DisableIcons    =   -1  'True
            DefaultRowHeight=   25
            Begin VB.TextBox txtEdit 
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
               Left            =   840
               TabIndex        =   84
               Top             =   0
               Visible         =   0   'False
               Width           =   1455
            End
         End
         Begin Line3D.ucLine3D ucLine3D1 
            Height          =   2520
            Index           =   20
            Left            =   -120
            Top             =   3600
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   4445
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Credencial.usCredencial DatosContrato 
            Height          =   2250
            Index           =   1
            Left            =   45
            TabIndex        =   208
            Top             =   3195
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   3969
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   5
            AlingHeader     =   0
            AlingBody       =   0
            BodyIndent      =   0
            HeaderIndent    =   0
            HeaderText      =   " Datos del contrato"
            HeaderBackColor =   16766131
            HeightHeader    =   22
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   25
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   16744576
            BackColor       =   16777215
         End
         Begin Credencial.usCredencial DetallesContrato 
            Height          =   1830
            Index           =   1
            Left            =   45
            TabIndex        =   212
            Top             =   5460
            Width           =   5310
            _ExtentX        =   9366
            _ExtentY        =   3228
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Calibri"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   5
            AlingHeader     =   0
            AlingBody       =   0
            BodyIndent      =   0
            HeaderIndent    =   0
            HeaderText      =   " Detalle del contrato"
            HeaderBackColor =   16766131
            HeightHeader    =   22
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   25
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   16744576
            BackColor       =   16777215
         End
         Begin VB.Label labelContratoAlmoneda 
            Caption         =   "CONTRATO EN ALMONEDA!"
            BeginProperty Font 
               Name            =   "Calibri"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   855
            Left            =   5400
            TabIndex        =   206
            Top             =   5160
            Visible         =   0   'False
            Width           =   1905
         End
         Begin VB.Label NuevoFolio 
            AutoSize        =   -1  'True
            Caption         =   "<Nvo. Folio>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   1920
            TabIndex        =   85
            Top             =   6360
            Visible         =   0   'False
            Width           =   1590
         End
         Begin VB.Label TotalRefrendo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   39.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   975
            Left            =   12120
            TabIndex        =   82
            Top             =   5580
            Width           =   270
         End
         Begin VB.Label LeyendaRef 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   24
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   570
            Left            =   12285
            TabIndex        =   81
            Top             =   5040
            Width           =   165
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Contrato:"
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
            TabIndex        =   70
            Top             =   240
            Width           =   1170
         End
      End
      Begin VB.Frame frmEmpeño 
         Caption         =   "Empeno"
         Height          =   8055
         Left            =   120
         TabIndex        =   68
         Top             =   480
         Visible         =   0   'False
         Width           =   12555
         Begin VB.CommandButton cmdHistorial 
            Caption         =   "Historial"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   5280
            TabIndex        =   221
            Top             =   435
            Width           =   990
         End
         Begin VB.CommandButton cmdAlerta 
            Caption         =   "Selec. Alerta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6360
            TabIndex        =   220
            ToolTipText     =   "Selección de Alerta para Ley Anti-lavado de Dinero."
            Top             =   435
            Width           =   1230
         End
         Begin VB.TextBox txtCotitularApellidoMaterno 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   3615
            MaxLength       =   70
            TabIndex        =   5
            Top             =   3150
            Width           =   3255
         End
         Begin VB.TextBox txtCotitularApellidoPaterno 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   135
            MaxLength       =   60
            TabIndex        =   4
            Top             =   3150
            Width           =   3255
         End
         Begin VB.TextBox txtResponsable 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   135
            MaxLength       =   30
            TabIndex        =   3
            Top             =   2520
            Width           =   3255
         End
         Begin VB.TextBox txtApellidoPaterno 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   135
            MaxLength       =   60
            TabIndex        =   1
            Top             =   1065
            Width           =   3255
         End
         Begin VB.TextBox txtApellidoMaterno 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   3615
            MaxLength       =   70
            TabIndex        =   2
            Top             =   1065
            Width           =   3255
         End
         Begin VB.TextBox txtNombre 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   120
            MaxLength       =   20
            TabIndex        =   0
            Top             =   465
            Width           =   3255
         End
         Begin VB.TextBox txtBeneficiario 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   7080
            MaxLength       =   30
            TabIndex        =   6
            Top             =   3150
            Width           =   5295
         End
         Begin VB.CommandButton cmdEditarCotitular 
            Caption         =   "Editar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4200
            TabIndex        =   219
            Top             =   2445
            Width           =   990
         End
         Begin VB.CommandButton cmdEditarCliente 
            Caption         =   "Editar"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   4200
            TabIndex        =   218
            Top             =   435
            Width           =   990
         End
         Begin vbAcceleratorGrid6.vbalGrid grdEmpeños 
            Height          =   3315
            Left            =   4080
            TabIndex        =   63
            Top             =   4095
            Width           =   8430
            _ExtentX        =   14870
            _ExtentY        =   5847
            RowMode         =   -1  'True
            GridLines       =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            HeaderFlat      =   -1  'True
            BorderStyle     =   2
            ScrollBarStyle  =   2
            Editable        =   -1  'True
            DisableIcons    =   -1  'True
         End
         Begin VB.ComboBox cmbPromocion 
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
            ItemData        =   "frmEmpeño.frx":2C31
            Left            =   10260
            List            =   "frmEmpeño.frx":2C53
            Style           =   2  'Dropdown List
            TabIndex        =   199
            Top             =   2175
            Width           =   2085
         End
         Begin vbalTabStrip6.TabControl TPrendas 
            Height          =   3705
            Left            =   0
            TabIndex        =   159
            Top             =   3720
            Width           =   4050
            _ExtentX        =   7144
            _ExtentY        =   6535
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FlatSeparators  =   -1  'True
            FlatButtons     =   -1  'True
            CoolTabs        =   1
            Begin DevPowerFlatBttn.FlatBttn cmdBorrar 
               Height          =   375
               Left            =   3015
               TabIndex        =   196
               Top             =   3270
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   661
               AlignCaption    =   4
               AlignPicture    =   2
               AutoSize        =   0   'False
               Caption         =   "     &Eliminar"
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
               Picture         =   "frmEmpeño.frx":2CC3
               PictureDisabled =   "frmEmpeño.frx":3215
            End
            Begin DevPowerFlatBttn.FlatBttn cmdDiamante 
               Height          =   375
               Left            =   1005
               TabIndex        =   191
               Top             =   3270
               Width           =   1050
               _ExtentX        =   1852
               _ExtentY        =   661
               AlignCaption    =   4
               AlignPicture    =   2
               AutoSize        =   0   'False
               Caption         =   "     &Diamante"
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
               Picture         =   "frmEmpeño.frx":3DE7
               PictureDisabled =   "frmEmpeño.frx":400B
            End
            Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
               Height          =   375
               Left            =   75
               TabIndex        =   17
               Top             =   3270
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   661
               AlignCaption    =   3
               AlignPicture    =   2
               AutoSize        =   0   'False
               Caption         =   " &Limpiar"
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
               Picture         =   "frmEmpeño.frx":4165
            End
            Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
               Height          =   375
               Left            =   2070
               TabIndex        =   16
               Top             =   3270
               Width           =   930
               _ExtentX        =   1640
               _ExtentY        =   661
               AlignCaption    =   4
               AlignPicture    =   2
               AutoSize        =   0   'False
               Caption         =   "    &Agregar"
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
               Picture         =   "frmEmpeño.frx":4269
               PictureDisabled =   "frmEmpeño.frx":45D3
            End
            Begin VB.Frame frmMetales 
               Caption         =   "Metales"
               Height          =   2970
               Left            =   15
               TabIndex        =   160
               Top             =   315
               Width           =   3990
               Begin VB.TextBox txtPiedras 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   2895
                  MaxLength       =   3
                  TabIndex        =   163
                  Top             =   1335
                  Width           =   1020
               End
               Begin VB.ComboBox cmbEstado 
                  Height          =   315
                  ItemData        =   "frmEmpeño.frx":472D
                  Left            =   2865
                  List            =   "frmEmpeño.frx":472F
                  Style           =   2  'Dropdown List
                  TabIndex        =   11
                  Top             =   690
                  Width           =   1095
               End
               Begin VB.ComboBox cmbTipo 
                  Height          =   315
                  ItemData        =   "frmEmpeño.frx":4731
                  Left            =   1065
                  List            =   "frmEmpeño.frx":4733
                  Style           =   2  'Dropdown List
                  TabIndex        =   8
                  Top             =   60
                  Width           =   2130
               End
               Begin VB.TextBox txtPesoPiedra 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   1095
                  MaxLength       =   20
                  TabIndex        =   162
                  Top             =   1335
                  Width           =   1020
               End
               Begin VB.ComboBox cmbPrenda 
                  Height          =   315
                  ItemData        =   "frmEmpeño.frx":4735
                  Left            =   1065
                  List            =   "frmEmpeño.frx":4737
                  Style           =   2  'Dropdown List
                  TabIndex        =   9
                  Top             =   375
                  Width           =   2895
               End
               Begin VB.TextBox txtObservaciones 
                  BorderStyle     =   0  'None
                  Height          =   720
                  Left            =   60
                  MaxLength       =   150
                  MultiLine       =   -1  'True
                  TabIndex        =   15
                  Top             =   2175
                  Width           =   3870
               End
               Begin VB.TextBox txtPrestamoo 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Left            =   1095
                  MaxLength       =   20
                  TabIndex        =   14
                  Top             =   1635
                  Width           =   1020
               End
               Begin VB.TextBox txtCantidad 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   1095
                  MaxLength       =   3
                  TabIndex        =   12
                  Top             =   1035
                  Width           =   1020
               End
               Begin VB.ComboBox cmbKilates 
                  Height          =   315
                  ItemData        =   "frmEmpeño.frx":4739
                  Left            =   1065
                  List            =   "frmEmpeño.frx":473B
                  Style           =   2  'Dropdown List
                  TabIndex        =   10
                  Top             =   690
                  Width           =   1110
               End
               Begin VB.TextBox txtAvaluo 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   240
                  Left            =   2895
                  MaxLength       =   20
                  TabIndex        =   161
                  Top             =   1635
                  Width           =   1020
               End
               Begin VB.TextBox txtPeso 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   2895
                  MaxLength       =   20
                  TabIndex        =   13
                  Top             =   1035
                  Width           =   1020
               End
               Begin VB.Label lblPrestamoMaximo 
                  Caption         =   "0.00"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   1680
                  TabIndex        =   204
                  Top             =   1920
                  Width           =   1215
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Préstamo Máximo:"
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
                  Index           =   39
                  Left            =   30
                  TabIndex        =   203
                  Top             =   1920
                  Width           =   1590
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Peso Piedra:"
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
                  Index           =   35
                  Left            =   30
                  TabIndex        =   190
                  Top             =   1365
                  Width           =   1035
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Avalúo:"
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
                  Index           =   8
                  Left            =   2175
                  TabIndex        =   168
                  Top             =   1665
                  Width           =   630
               End
               Begin VB.Label lblAvaluoDiamante 
                  BackColor       =   &H80000013&
                  Caption         =   "0"
                  Height          =   255
                  Left            =   3540
                  TabIndex        =   177
                  Top             =   1470
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Label lblPrestamoDiamante 
                  BackColor       =   &H80000013&
                  Caption         =   "0"
                  Height          =   255
                  Left            =   3255
                  TabIndex        =   176
                  Top             =   1050
                  Visible         =   0   'False
                  Width           =   555
               End
               Begin VB.Label lblPuntos 
                  BackColor       =   &H80000013&
                  Caption         =   "0"
                  Height          =   255
                  Left            =   2820
                  TabIndex        =   175
                  Top             =   1815
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Label lblCantidadPiedras 
                  BackColor       =   &H80000013&
                  Caption         =   "0"
                  Height          =   255
                  Left            =   3510
                  TabIndex        =   174
                  Top             =   1815
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Label lblPiedra 
                  BackColor       =   &H80000013&
                  Height          =   255
                  Left            =   2070
                  TabIndex        =   173
                  Top             =   1815
                  Visible         =   0   'False
                  Width           =   600
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
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
                  Index           =   4
                  Left            =   2175
                  TabIndex        =   172
                  Top             =   750
                  Width           =   615
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo:"
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
                  Index           =   0
                  Left            =   30
                  TabIndex        =   171
                  Top             =   120
                  Width           =   405
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Piedras:"
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
                  Index           =   6
                  Left            =   2175
                  TabIndex        =   170
                  Top             =   1365
                  Width           =   675
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Préstamo:"
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
                  Index           =   7
                  Left            =   30
                  TabIndex        =   169
                  Top             =   1665
                  Width           =   870
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Kilataje:"
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
                  Index           =   2
                  Left            =   30
                  TabIndex        =   167
                  Top             =   750
                  Width           =   690
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Peso:"
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
                  Index           =   5
                  Left            =   2175
                  TabIndex        =   166
                  Top             =   1065
                  Width           =   450
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Cantidad:"
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
                  Index           =   3
                  Left            =   30
                  TabIndex        =   165
                  Top             =   1065
                  Width           =   795
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Prenda:"
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
                  Index           =   1
                  Left            =   30
                  TabIndex        =   164
                  Top             =   435
                  Width           =   645
               End
            End
            Begin VB.Frame frmElectronicos 
               Caption         =   "Electronicos"
               Height          =   2970
               Left            =   15
               TabIndex        =   178
               Top             =   315
               Width           =   3990
               Begin VB.TextBox txtMarcaElec 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   1095
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   20
                  Top             =   727
                  Width           =   1020
               End
               Begin VB.TextBox txtFamiliaElec 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   1095
                  Locked          =   -1  'True
                  MaxLength       =   80
                  TabIndex        =   19
                  Top             =   412
                  Width           =   2850
               End
               Begin DevPowerFlatBttn.FlatBttn cmdMostrarCatPrendas 
                  Height          =   270
                  Left            =   3195
                  TabIndex        =   197
                  Top             =   90
                  Width           =   315
                  _ExtentX        =   556
                  _ExtentY        =   476
                  AutoSize        =   0   'False
                  Caption         =   ". . ."
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Tahoma"
                     Size            =   9
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
               Begin VB.TextBox txtColorElec 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   2895
                  MaxLength       =   50
                  TabIndex        =   23
                  Top             =   1035
                  Width           =   1050
               End
               Begin VB.TextBox txtTamañoElec 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   1095
                  MaxLength       =   50
                  TabIndex        =   22
                  Top             =   1035
                  Width           =   1020
               End
               Begin VB.TextBox txtNumSerieElec 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   1095
                  MaxLength       =   80
                  TabIndex        =   24
                  Top             =   1335
                  Width           =   2850
               End
               Begin VB.TextBox txtModeloElec 
                  BorderStyle     =   0  'None
                  Height          =   240
                  Left            =   2895
                  Locked          =   -1  'True
                  MaxLength       =   50
                  TabIndex        =   21
                  Top             =   727
                  Width           =   1050
               End
               Begin VB.TextBox txtAvaluoElec 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   240
                  Left            =   2850
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   27
                  Top             =   1635
                  Width           =   1095
               End
               Begin VB.TextBox txtPrestamooElec 
                  Alignment       =   1  'Right Justify
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   240
                  Left            =   1080
                  MaxLength       =   20
                  TabIndex        =   25
                  Top             =   1635
                  Width           =   1020
               End
               Begin VB.TextBox txtObservacionesElec 
                  BorderStyle     =   0  'None
                  Height          =   720
                  Left            =   60
                  MaxLength       =   250
                  MultiLine       =   -1  'True
                  TabIndex        =   26
                  Top             =   2175
                  Width           =   3870
               End
               Begin VB.ComboBox cmbTipoElec 
                  Height          =   315
                  ItemData        =   "frmEmpeño.frx":473D
                  Left            =   1065
                  List            =   "frmEmpeño.frx":473F
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   60
                  Width           =   2130
               End
               Begin VB.Label lblPrestamoMaximo2 
                  Caption         =   "0.00"
                  BeginProperty Font 
                     Name            =   "Tahoma"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000C0&
                  Height          =   195
                  Left            =   1680
                  TabIndex        =   205
                  Top             =   1920
                  Width           =   1215
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Color:"
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
                  Index           =   38
                  Left            =   2205
                  TabIndex        =   195
                  Top             =   1065
                  Width           =   480
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Tamaño:"
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
                  Index           =   36
                  Left            =   30
                  TabIndex        =   194
                  Top             =   1065
                  Width           =   735
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "No. Serie:"
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
                  Index           =   34
                  Left            =   30
                  TabIndex        =   186
                  Top             =   1365
                  Width           =   780
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Modelo:"
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
                  Index           =   31
                  Left            =   2205
                  TabIndex        =   185
                  Top             =   750
                  Width           =   660
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Marca:"
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
                  Index           =   29
                  Left            =   30
                  TabIndex        =   184
                  Top             =   750
                  Width           =   570
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Familia:"
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
                  Index           =   37
                  Left            =   30
                  TabIndex        =   183
                  Top             =   435
                  Width           =   645
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Préstamo:"
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
                  Index           =   33
                  Left            =   30
                  TabIndex        =   182
                  Top             =   1665
                  Width           =   870
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Préstamo Máximo:"
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
                  Index           =   32
                  Left            =   30
                  TabIndex        =   181
                  Top             =   1920
                  Width           =   1590
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Tipo:"
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
                  Index           =   30
                  Left            =   30
                  TabIndex        =   180
                  Top             =   120
                  Width           =   405
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Avalúo:"
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
                  Index           =   28
                  Left            =   2175
                  TabIndex        =   179
                  Top             =   1665
                  Width           =   630
               End
            End
         End
         Begin VB.ComboBox cmbPeriodo 
            BackColor       =   &H0000FFFF&
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
            ItemData        =   "frmEmpeño.frx":4741
            Left            =   9930
            List            =   "frmEmpeño.frx":4743
            Style           =   2  'Dropdown List
            TabIndex        =   158
            Top             =   1830
            Width           =   1455
         End
         Begin VB.ComboBox cmbPlazos 
            BackColor       =   &H0000FFFF&
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
            ItemData        =   "frmEmpeño.frx":4745
            Left            =   11445
            List            =   "frmEmpeño.frx":4747
            Style           =   2  'Dropdown List
            TabIndex        =   155
            Top             =   1830
            Width           =   870
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   18
            Left            =   8235
            Top             =   105
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   345
            Index           =   16
            Left            =   8235
            Top             =   120
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   609
            Orientation     =   0
            LineWidth       =   2
         End
         Begin VB.TextBox txtNumBolsa 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   8220
            MaxLength       =   15
            TabIndex        =   7
            Top             =   3465
            Width           =   1410
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   285
            Index           =   7
            Left            =   11520
            Top             =   720
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   503
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   285
            Index           =   6
            Left            =   10500
            Top             =   720
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   503
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   270
            Index           =   5
            Left            =   9180
            Top             =   735
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   476
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   0
            Left            =   9870
            Top             =   1815
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D29 
            Height          =   30
            Index           =   0
            Left            =   8235
            Top             =   1800
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   21
            Left            =   8235
            Top             =   2490
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   300
            Index           =   12
            Left            =   10245
            Top             =   1245
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   529
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   9
            Left            =   8235
            Top             =   1230
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   2
            Left            =   8235
            Top             =   975
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   2280
            Index           =   3
            Left            =   12390
            Top             =   240
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   4022
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   4
            Left            =   8235
            Top             =   720
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   2115
            Index           =   1
            Left            =   8235
            Top             =   390
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   3731
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   345
            Index           =   14
            Left            =   10200
            Top             =   105
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   609
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   345
            Index           =   15
            Left            =   12390
            Top             =   105
            Width           =   75
            _ExtentX        =   132
            _ExtentY        =   609
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   17
            Left            =   8235
            Top             =   435
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin VB.ComboBox cmbTipoInteres 
            BackColor       =   &H0000FFFF&
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
            ItemData        =   "frmEmpeño.frx":4749
            Left            =   8295
            List            =   "frmEmpeño.frx":474B
            Style           =   2  'Dropdown List
            TabIndex        =   113
            Top             =   1830
            Width           =   1560
         End
         Begin VB.TextBox txtEdad 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   4740
            Locked          =   -1  'True
            MaxLength       =   30
            TabIndex        =   62
            Top             =   4050
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox txtMensaje 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   105
            MaxLength       =   150
            TabIndex        =   64
            Top             =   7695
            Width           =   4665
         End
         Begin VB.ComboBox cmbMedio 
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
            ItemData        =   "frmEmpeño.frx":474D
            Left            =   9870
            List            =   "frmEmpeño.frx":474F
            Style           =   2  'Dropdown List
            TabIndex        =   67
            Top             =   7635
            Width           =   2535
         End
         Begin VB.TextBox txtNotas 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   4905
            MaxLength       =   250
            TabIndex        =   66
            Top             =   7695
            Width           =   4830
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
            Height          =   225
            Left            =   3645
            TabIndex        =   61
            Top             =   480
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   397
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
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   390
            Index           =   44
            Left            =   11400
            Top             =   1815
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   688
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   0
            Left            =   8250
            Top             =   1515
            Width           =   4185
            _ExtentX        =   7382
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   8
            Left            =   8235
            Top             =   2160
            Width           =   4200
            _ExtentX        =   7408
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   3
            Left            =   10200
            Top             =   2160
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D3 
            Height          =   135
            Left            =   120
            Top             =   1995
            Width           =   7980
            _ExtentX        =   14076
            _ExtentY        =   238
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMosCotitular 
            Height          =   225
            Left            =   3600
            TabIndex        =   222
            Top             =   2475
            Width           =   390
            _ExtentX        =   688
            _ExtentY        =   397
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
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Nombre"
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
            Index           =   14
            Left            =   120
            TabIndex        =   225
            Top             =   240
            Width           =   3255
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Materno"
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
            Index           =   17
            Left            =   3615
            TabIndex        =   235
            Top             =   2910
            Width           =   3255
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Paterno"
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
            Index           =   16
            Left            =   135
            TabIndex        =   234
            Top             =   2910
            Width           =   3255
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Cotitular Nombre"
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
            Index           =   26
            Left            =   135
            TabIndex        =   233
            Top             =   2280
            Width           =   3255
         End
         Begin VB.Label lblRFC 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   6000
            TabIndex        =   232
            Top             =   1755
            Width           =   2085
         End
         Begin VB.Label lblCiudad 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1320
            TabIndex        =   231
            Top             =   1755
            Width           =   3525
         End
         Begin VB.Label Label65 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Ciudad"
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
            TabIndex        =   230
            Top             =   1755
            Width           =   1155
         End
         Begin VB.Label Label7 
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
            Height          =   210
            Left            =   4920
            TabIndex        =   229
            Top             =   1755
            Width           =   1065
         End
         Begin VB.Label lblDireccion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   1320
            TabIndex        =   228
            Top             =   1395
            Width           =   6765
         End
         Begin VB.Label Label10 
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
            Index           =   0
            Left            =   120
            TabIndex        =   227
            Top             =   1395
            Width           =   1155
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Materno"
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
            Index           =   13
            Left            =   3615
            TabIndex        =   226
            Top             =   825
            Width           =   3255
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Apellido Paterno"
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
            Index           =   15
            Left            =   135
            TabIndex        =   224
            Top             =   825
            Width           =   3255
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Beneficiario"
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
            Index           =   41
            Left            =   7080
            TabIndex        =   223
            Top             =   2910
            Width           =   5295
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Promoción:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   5
            Left            =   8430
            TabIndex        =   200
            Top             =   2220
            Width           =   1485
         End
         Begin VB.Label lblCostoMensual 
            BackColor       =   &H00404040&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   315
            Index           =   2
            Left            =   8265
            TabIndex        =   198
            Top             =   2175
            Width           =   1920
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Periodo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   10305
            TabIndex        =   157
            Top             =   1560
            Width           =   705
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Plazo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   11640
            TabIndex        =   156
            Top             =   1560
            Width           =   480
         End
         Begin VB.Label lblVencimiento 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<Vencimiento>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   10815
            TabIndex        =   154
            Top             =   165
            Width           =   1800
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Num. Bolsa:"
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
            Index           =   27
            Left            =   7080
            TabIndex        =   131
            Top             =   3435
            Width           =   1065
         End
         Begin VB.Label Label82 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Avalúo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   8280
            TabIndex        =   128
            Top             =   1005
            Width           =   1995
         End
         Begin VB.Label lblIva 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<IVA>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   11640
            TabIndex        =   126
            Top             =   750
            Width           =   570
         End
         Begin VB.Label Label83 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%Almacenaje"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   9240
            TabIndex        =   124
            Top             =   465
            Width           =   1245
         End
         Begin VB.Label lblAlmacenaje 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<Almacenaje>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   9255
            TabIndex        =   123
            Top             =   750
            Width           =   1215
         End
         Begin VB.Label Label39 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%I.V.A."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   11565
            TabIndex        =   122
            Top             =   465
            Width           =   720
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vence:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   270
            Index           =   0
            Left            =   10245
            TabIndex        =   121
            Top             =   135
            Width           =   780
         End
         Begin VB.Label lblContrato 
            AutoSize        =   -1  'True
            Caption         =   "<Folio>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   9345
            TabIndex        =   120
            Top             =   150
            Width           =   915
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Contrato:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   8280
            TabIndex        =   119
            Top             =   135
            Width           =   1050
         End
         Begin VB.Label Label27 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%Tasa"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   8400
            TabIndex        =   118
            Top             =   465
            Width           =   645
         End
         Begin VB.Label lblTasa 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Tasa>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   210
            Left            =   8370
            TabIndex        =   117
            Top             =   750
            Width           =   705
         End
         Begin VB.Label txtPrestamo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Prestamo>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   10770
            TabIndex        =   116
            Top             =   1275
            Width           =   1170
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tipo Interes"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   8505
            TabIndex        =   114
            Top             =   1560
            Width           =   1110
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Edad:"
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
            Index           =   24
            Left            =   4200
            TabIndex        =   112
            Top             =   4050
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label lblTotAvaluo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Avaluo>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   210
            Left            =   8925
            TabIndex        =   93
            Top             =   1275
            Width           =   930
         End
         Begin VB.Label lblSeguro 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "<Seguro>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   10575
            TabIndex        =   90
            Top             =   750
            Width           =   870
         End
         Begin VB.Label Label84 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "%Seguro"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   10575
            TabIndex        =   89
            Top             =   465
            Width           =   870
         End
         Begin VB.Label Label80 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Préstamo"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   0
            Left            =   10560
            TabIndex        =   88
            Top             =   1005
            Width           =   1395
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Como se enteró:"
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
            Index           =   12
            Left            =   9885
            TabIndex        =   87
            Top             =   7425
            Width           =   1515
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Mensaje interno:"
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
            Index           =   10
            Left            =   120
            TabIndex        =   86
            Top             =   7470
            Width           =   1545
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Notas impresas en contrato:"
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
            Index           =   11
            Left            =   4920
            TabIndex        =   65
            Top             =   7470
            Width           =   2595
         End
         Begin VB.Label lblFolio 
            AutoSize        =   -1  'True
            Caption         =   "<Folio>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   270
            Left            =   11280
            TabIndex        =   73
            Top             =   150
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label1 
            BackColor       =   &H00404040&
            Height          =   240
            Index           =   0
            Left            =   8250
            TabIndex        =   125
            Top             =   1005
            Width           =   4185
         End
         Begin VB.Label Label51 
            BackColor       =   &H00404040&
            Height          =   255
            Index           =   0
            Left            =   8250
            TabIndex        =   129
            Top             =   465
            Width           =   4185
         End
         Begin VB.Label lblCostoMensual 
            BackColor       =   &H00404040&
            Height          =   300
            Index           =   0
            Left            =   8250
            TabIndex        =   130
            Top             =   1530
            Width           =   4185
         End
      End
   End
   Begin VB.Label lblInfoSemaforo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<InfoSemaforo>"
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
      Left            =   2040
      TabIndex        =   217
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label lblPuntosAcumulados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   12000
      TabIndex        =   216
      Top             =   240
      Width           =   120
   End
   Begin VB.Label lblPuntosAcumulados1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos Acumulados:"
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
      Left            =   9960
      TabIndex        =   215
      Top             =   240
      Width           =   1890
   End
   Begin VB.Label lblNoTarjeta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Tarjeta:"
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
      Left            =   7080
      TabIndex        =   214
      Top             =   240
      Width           =   1065
   End
   Begin VB.Image ImgSemaforo 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   120
      Stretch         =   -1  'True
      ToolTipText     =   "prueba para el semaforo"
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label lblAutorizacion 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   270
      Left            =   120
      TabIndex        =   127
      Top             =   8160
      Visible         =   0   'False
      Width           =   60
   End
End
Attribute VB_Name = "frmEmpeño"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 03/04/02
' Modulo frmEmpeño - frmEmpeño.frm
' Ultima Modificacion - 19/08/02 - L.I. Jorge Gabriel Colio Ramos
' Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'
'////////////////////////////////////////////////////////////////.

Option Explicit

Dim Fl() As cFlatControl
Dim m_Peso As Double, Bandera As Boolean, BanElec As Boolean, pIDUsuarioAutoriza As Integer, pTipoAutorizacion As Integer, pAlmoneda As Single
'***Puntos***
Dim TarjetaPuntos As New ClienteFrecuente

'MLD-MODIF.
Dim ClienteEmp As clientes
Dim CotitularEmp As clientes
Dim Titular As Boolean
Dim vTipoAlerta As TipoAlerta

Public Property Let IDUsuarioAutoriza(Valor As Integer)
    pIDUsuarioAutoriza = Valor
End Property

Public Property Get IDUsuarioAutoriza() As Integer
    IDUsuarioAutoriza = pIDUsuarioAutoriza
End Property

Public Property Let TipoAutorizacion(Valor As Integer)
    pTipoAutorizacion = Valor
End Property

Public Property Get TipoAutorizacion() As Integer
    TipoAutorizacion = pTipoAutorizacion
End Property

Public Property Let Almoneda(Valor As Integer)
    pAlmoneda = Valor
End Property

Public Property Get Almoneda() As Integer
    Almoneda = pAlmoneda
End Property



Private Sub chkCirculacionDes_Click()
    If grdDesempeño.Rows > 0 Then
        Poner_Totales_Desempeño
    End If
End Sub

Private Sub chkCirculacionRef_Click()
If grdRefrendos.Rows > 0 Then
    Poner_Totales_Refrendo
End If
    
End Sub

Private Sub cmbEstado_Click()
    If cmbKilates.ListIndex > -1 And cmbEstado.ListIndex > -1 Then Calcular_Avaluo
End Sub

Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbMedio_Click()
    Dim strMedio As String, IDMedio As Integer

    If cmbMedio.text = "[AGREGAR]" Then
        IDMedio = 0
        strMedio = ""
        IDMedio = frmAgregaMedio.Mostrar()
        If IDMedio > 0 Then
            cmbMedio.Clear
            cmbMedio.AddItem "[AGREGAR]"
            Cargar_Combos "Descripcion", "medios", cmbMedio, , "Descripcion", False
            cmbMedio.ListIndex = ComboInformacion(cmbMedio, IDMedio)
        Else
            cmbMedio.ListIndex = -1
        End If
    End If
End Sub

Private Sub cmbMedio2_Click()
Dim strMedio As String, IDMedio As Integer

    If cmbMedio2.text = "[AGREGAR]" Then
        IDMedio = 0
        strMedio = ""
        IDMedio = frmAgregaMedio.Mostrar()
        If IDMedio > 0 Then
            
            cmbMedio2.Clear
            cmbMedio2.AddItem "[AGREGAR]"
            Cargar_Combos "Descripcion", "medios", cmbMedio2, , "Descripcion", False
            cmbMedio2.ListIndex = ComboInformacion(cmbMedio2, IDMedio)
        
        Else
            
            cmbMedio2.ListIndex = -1
        End If
    
    End If
End Sub

Private Sub cmbPeriodo2_Click()
    
    If cmbPeriodo2.ListIndex > -1 Then
        
        Cargar_Combos "DISTINCT plazos.Descripcion", "configuraciontasas INNER JOIN plazos ON plazos.ID=configuraciontasas.IDPlazo", cmbPlazos2, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres2.ItemData(cmbTipoInteres2.ListIndex) & " AND configuraciontasas.IDTipoPeriodo=" & cmbPeriodo2.ItemData(cmbPeriodo2.ListIndex), "plazos.Descripcion", , "plazos.ID"
    
    End If
    
    If cmbPlazos2.ListCount > 0 Then cmbPlazos2.ListIndex = 0
End Sub

Private Sub cmbPeriodo2_GotFocus()
    'Cambiar_Color True, cmbPeriodo2
End Sub

Private Sub cmbPeriodo2_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPeriodo2_LostFocus()
    'Cambiar_Color False, cmbPeriodo2
End Sub

Private Sub cmbPlazos2_Click()
Dim crPrestamo As Double
    
    If Val(txtPrestamo2.text) = 0 Or Trim(txtPrestamo2.text) = "" Then
        
        crPrestamo = 0
    Else
        
        crPrestamo = txtPrestamo2.text
    End If
        
    If Bandera = False Then
        
        SacaTasa crPrestamo, cmbTipoInteres2.ItemData(cmbTipoInteres2.ListIndex), cmbPeriodo2.ItemData(cmbPeriodo2.ListIndex), cmbPlazos2.ItemData(cmbPlazos2.ListIndex), IIf(Val(txtNombre2.Tag) = 0, False, True)
        Calcular_Avaluo
        Recalcula
    
    Else
        
        SacaTasa CDbl(txtPrestamo2.text), cmbTipoInteres2.ItemData(cmbTipoInteres2.ListIndex), cmbPeriodo2.ItemData(cmbPeriodo2.ListIndex), cmbPlazos2.ItemData(cmbPlazos2.ListIndex), IIf(Val(txtNombre2.Tag) = 0, False, True)
    End If

    Bandera = False
End Sub

Private Sub cmbPlazos2_GotFocus()
    'Cambiar_Color True, cmbPlazos2
End Sub

Private Sub cmbPlazos2_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPlazos2_LostFocus()
    'Cambiar_Color False, cmbPlazos2
End Sub

'''''Private Sub cmbPrendaElec_Click()
'''''Dim strPrenda As String, IDPrenda As Long
'''''
'''''    If cmbPrendaElec.text = "[AGREGAR]" Then
'''''        IDPrenda = 0
'''''        strPrenda = ""
'''''        IDPrenda = frmAgregaPrenda.Mostrar(cmbTipoElec.ItemData(cmbTipoElec.ListIndex))
'''''        If IDPrenda > 0 Then
'''''
'''''            cmbPrendaElec.Clear
'''''            cmbPrendaElec.AddItem "[AGREGAR]"
'''''            Cargar_Combos "Descripcion", "tipoprenda", cmbPrendaElec, " WHERE IDTipo=" & cmbTipoElec.ItemData(cmbTipoElec.ListIndex), "Descripcion", False
'''''            cmbPrendaElec.ListIndex = ComboInformacion(cmbPrendaElec, IDPrenda)
'''''
'''''        Else
'''''
'''''            cmbPrendaElec.ListIndex = -1
'''''        End If
'''''
'''''    End If
'''''End Sub

'''''Private Sub cmbPrendaElec_GotFocus()
'''''    Cambiar_Color True, cmbPrendaElec
'''''End Sub
'''''
'''''Private Sub cmbPrendaElec_KeyPress(KeyAscii As Integer)
'''''    Pasar_Foco KeyAscii
'''''End Sub
'''''
'''''Private Sub cmbPrendaElec_LostFocus()
'''''    Cambiar_Color False, cmbPrendaElec
'''''End Sub

Private Sub cmbKilates_Click()
    If cmbKilates.ListIndex > -1 And cmbEstado.ListIndex > -1 Then Calcular_Avaluo
End Sub

Private Sub cmbKilates_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPeriodo_Click()
    
    If cmbPeriodo.ListIndex > -1 Then
        
        Cargar_Combos "DISTINCT plazos.Descripcion", "configuraciontasas INNER JOIN plazos ON plazos.ID=configuraciontasas.IDPlazo", cmbPlazos, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND configuraciontasas.IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex), "plazos.Descripcion", , "plazos.ID"
    
    End If
    
    If cmbPlazos.ListCount > 0 Then cmbPlazos.ListIndex = 0
End Sub

Private Sub cmbPeriodo_GotFocus()
    'Cambiar_Color True, cmbPeriodo
End Sub

Private Sub cmbPeriodo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPeriodo_LostFocus()
    'Cambiar_Color False, cmbPeriodo
End Sub

Private Sub cmbPlazos_Click()

    If Bandera = False Then
        
        SacaTasa CCur(txtPrestamo.Caption), cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), IIf(Val(txtNombre.Tag) = 0, False, True)
        Calcular_Avaluo
        Recalcula
    
    Else
        
        SacaTasa CCur(txtPrestamo.Caption), cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), IIf(Val(txtNombre.Tag) = 0, False, True)
    End If

    Bandera = False
End Sub

Private Sub cmbPlazos_GotFocus()
    'Cambiar_Color True, cmbPlazos
End Sub

Private Sub cmbPlazos_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPlazos_LostFocus()
    'Cambiar_Color False, cmbPlazos
End Sub

Private Sub cmbPrenda_Click()
Dim IDPrenda As Integer

    If cmbPrenda.text = "[1. AGREGAR PRENDA]" Then
        
        IDPrenda = 0
        IDPrenda = frmAgregaPrendaOro.Mostrar()
        If IDPrenda > 0 Then
            
            cmbPrenda.Clear
            cmbPrenda.AddItem "[1. AGREGAR PRENDA]"
            Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Descripcion", False
            cmbPrenda.ListIndex = ComboInformacion(cmbPrenda, IDPrenda)
        
        Else
            
            cmbPrenda.ListIndex = -1
        End If
    
    End If
End Sub

Private Sub cmbPrenda_GotFocus()
    Seleccionar_Texto cmbPrenda
    Cambiar_Color True, cmbPrenda
End Sub

Private Sub cmbPrenda_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPrenda_LostFocus()
    Cambiar_Color False, cmbPrenda
End Sub

Private Sub cmbPromocion_GotFocus()
    Cambiar_Color True, cmbPromocion
End Sub

Private Sub cmbPromocion_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPromocion_LostFocus()
    Cambiar_Color False, cmbPromocion
End Sub

Private Sub cmbPromocion2_GotFocus()
    Cambiar_Color True, cmbPromocion2
End Sub

Private Sub cmbPromocion2_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPromocion2_LostFocus()
    Cambiar_Color False, cmbPromocion2
End Sub



'''''Private Sub cmbTipoElec_Click()
'''''
'''''    If cmbTipoElec.ListIndex > -1 Then
'''''        cmbPrendaElec.Clear
'''''        cmbPrendaElec.AddItem "[AGREGAR]"
'''''        Cargar_Combos "Descripcion", "tipoprenda", cmbPrendaElec, " WHERE IDTipo<>1", "Descripcion", False
'''''    End If
'''''
'''''End Sub

Private Sub cmbTipoElec_GotFocus()
    Cambiar_Color True, cmbTipoElec
End Sub

Private Sub cmbTipoElec_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoElec_LostFocus()
    Cambiar_Color False, cmbTipoElec
End Sub

Private Sub cmbTipoInteres_Click()

    If cmbTipoInteres.ListIndex > -1 Then
        
        cmbPeriodo.Clear
        cmbPlazos.Clear
        Cargar_Combos "DISTINCT tipoperiodo.Descripcion", "configuraciontasas INNER JOIN tipoperiodo ON tipoperiodo.ID=configuraciontasas.IDTipoPeriodo", cmbPeriodo, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), "tipoperiodo.Ordenamiento", , "tipoperiodo.ID"
    
    End If
    
    If cmbPeriodo.ListCount > 0 Then cmbPeriodo.ListIndex = 0
End Sub

Private Sub cmbTipoInteres_GotFocus()
    'Cambiar_Color True, cmbTipoInteres
End Sub

Private Sub cmbTipoInteres_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoInteres_LostFocus()
    'Cambiar_Color False, cmbTipoInteres
End Sub

Private Sub cmbTipoInteres2_Click()

    If cmbTipoInteres2.ListIndex > -1 Then
        
        cmbPeriodo2.Clear
        cmbPlazos2.Clear
        Cargar_Combos "DISTINCT tipoperiodo.Descripcion", "configuraciontasas INNER JOIN tipoperiodo ON tipoperiodo.ID=configuraciontasas.IDTipoPeriodo", cmbPeriodo2, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres2.ItemData(cmbTipoInteres2.ListIndex), "tipoperiodo.Ordenamiento", , "tipoperiodo.ID"
    
    End If
    
    If cmbPeriodo2.ListCount > 0 Then cmbPeriodo2.ListIndex = 0
End Sub

Private Sub cmbTipoInteres2_GotFocus()
    'Cambiar_Color True, cmbTipoInteres2
End Sub

Private Sub cmbTipoInteres2_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoInteres2_LostFocus()
    'Cambiar_Color False, cmbTipoInteres2
End Sub

Private Sub cmbTipo_Click()

    If cmbTipo.ListIndex > -1 Then
        
        cmbPrenda.AddItem "[1. AGREGAR PRENDA]"
        Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Descripcion", False
        Cargar_Combos "Descripcion", "kilatajes", cmbKilates, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Ordenamiento"
        Cargar_Combos "Estado", "estado", cmbEstado, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Ordenamiento"
    
    End If

End Sub

Private Sub cmdAgregar_Click()
Dim i As Integer, Estado As Integer, Kilates As Integer, PrecioVenta As Double, crPrestamo As Double
Dim IDTipo As Integer, IDTipoPrenda As Long, strPrenda As String, Piedras As Integer, PesoPiedras As Double

    If ValidaArticulos(TPrendas.SelectedTab) Then
    
        With grdEmpeños
            
            If Val(txtCantidad.Tag) > 0 And TPrendas.SelectedTab = 1 Then i = Val(txtCantidad.Tag): GoTo Edicion
            For i = 1 To .Rows

                If Val(.CellText(i, 2)) = 0 Then Exit For
            Next i
            
Edicion:
            If i = 10 Then MsgBox "No se pueden agregar más prendas a la boleta !!", vbInformation, "Empeños":  Exit Sub
            
            If TPrendas.SelectedTab = 1 Then
            
                IDTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
            Else
            
                IDTipo = cmbTipoElec.ItemData(cmbTipoElec.ListIndex)
            End If
            .CellText(i, 1) = IIf(TPrendas.SelectedTab = 1, cmbTipo.text, cmbTipoElec.text)
            .CellItemData(i, 1) = IDTipo
        
            .CellText(i, 2) = IIf(TPrendas.SelectedTab = 1, Val(txtCantidad), "1")
            .CellTextAlign(i, 2) = DT_RIGHT
            
            If TPrendas.SelectedTab = 1 Then
                
                IDTipoPrenda = cmbPrenda.ItemData(cmbPrenda.ListIndex)
                strPrenda = cmbPrenda.text & " " & lblPiedra.Caption
            Else
                
                IDTipoPrenda = Val(txtFamiliaElec.Tag)
                strPrenda = txtFamiliaElec.text
            End If
            
            .CellText(i, 3) = strPrenda
            .CellItemData(i, 3) = IDTipoPrenda
        
            .CellText(i, 4) = IIf(TPrendas.SelectedTab = 1, txtPeso.text, "")
            .CellTextAlign(i, 4) = DT_RIGHT
            
            If TPrendas.SelectedTab = 1 Then
                
                Kilates = RegresaKilates(cmbKilates.text, cmbTipo.text)
            Else
                
                Kilates = 0
            End If
            .CellText(i, 5) = IIf(TPrendas.SelectedTab = 1, cmbKilates.text, "")
            .CellTextAlign(i, 5) = DT_RIGHT
            .CellItemData(i, 5) = Kilates
        
            .CellText(i, 6) = IIf(TPrendas.SelectedTab = 1, txtAvaluo.text, txtAvaluoElec.text)
            .CellTextAlign(i, 6) = DT_RIGHT
        
            .CellText(i, 7) = IIf(TPrendas.SelectedTab = 1, txtPrestamoo.text, txtPrestamooElec.text)
            .CellTextAlign(i, 7) = DT_RIGHT
        
            If TPrendas.SelectedTab = 1 And cmbEstado.ListIndex >= 0 Then
                
                Estado = cmbEstado.ItemData(cmbEstado.ListIndex)
            Else
                
                Estado = 0
            End If

            .CellText(i, 9) = IIf(TPrendas.SelectedTab = 1, cmbEstado.text, "")
            .CellItemData(i, 9) = Estado
        
            .CellText(i, 10) = PrecioVenta
            .CellTextAlign(i, 10) = DT_RIGHT
        
            .CellText(i, 11) = IIf(TPrendas.SelectedTab = 1, txtObservaciones.text, txtObservacionesElec.text)
            
            If Val(txtPiedras.text) > 0 Or Trim(txtPiedras.text) <> "" Then
                
                Piedras = Val(txtPiedras.text)
            Else
                
                Piedras = 0
            End If
            
            If Val(txtPesoPiedra.text) > 0 Or Trim(txtPesoPiedra.text) <> "" Then
                
                PesoPiedras = CDbl(txtPesoPiedra.text)
            Else
                
                PesoPiedras = 0
            End If
            
            .CellText(i, 12) = Piedras
            .CellText(i, 13) = PesoPiedras
            
            .CellText(i, 14) = lblCantidadPiedras.Caption
            .CellText(i, 15) = lblPuntos.Caption
            .CellText(i, 16) = lblPrestamoDiamante.Caption
            
            .CellText(i, 17) = Trim(txtMarcaElec.text)
            .CellText(i, 18) = Trim(txtModeloElec.text)
            .CellText(i, 19) = Trim(txtNumSerieElec.text)
            .CellText(i, 20) = Trim(txtColorElec.text)
            .CellText(i, 21) = Trim(txtTamañoElec.text)
            
            Total_Avaluos
            .ClearSelection
            LimpiaArticulos
            If TPrendas.SelectedTab = 1 Then cmbTipo.ListIndex = 0 Else cmbTipoElec.ListIndex = 0
            
            'Tomo el Préstamo
            If Val(txtPrestamo.Caption) = 0 Or Trim(txtPrestamo.Caption) = "" Then
                
                crPrestamo = 0
            Else
            
                crPrestamo = txtPrestamo.Caption
            End If
            
            'Checo la Tasa
            MuestraTasa cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), crPrestamo, IIf(Val(txtNombre.Tag) = 0, False, True), lblTasa, False
            
            txtCantidad.text = "1"
            If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
        End With

    End If

End Sub

Private Sub cmdBorrar_Click()
Dim i As Integer, crPrestamo As Double

    If grdEmpeños.SelectedRow > 0 Then
        
        If Trim(grdEmpeños.CellText(grdEmpeños.SelectedRow, 1)) <> "" Then
            
            If MsgBox("Desea eliminar la prenda seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton1, "Empeños") = vbYes Then
                
                grdEmpeños.RemoveRow grdEmpeños.SelectedRow
                Total_Avaluos
                For i = 1 To 11 - grdEmpeños.Rows
                    grdEmpeños.AddRow
                Next i
                
                'Tomo el Préstamo
                If Val(txtPrestamo.Caption) = 0 Or Trim(txtPrestamo.Caption) = "" Then
                    
                    crPrestamo = txtPrestamo.Caption
                Else
                    
                    crPrestamo = 0
                End If
                
                'Checo la Tasa
                MuestraTasa cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), crPrestamo, IIf(Val(txtNombre.Tag) = 0, False, True), lblTasa, False
            
            End If
        
        End If
    
    End If
    
    grdEmpeños.ClearSelection
    If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    frmBusqueda.Show
    BringWindowToTop frmBusqueda.hWnd
End Sub

Private Sub cmdCotizar_Click()
Dim crPrestamo As Double, crAvaluo As Double
    
    If Val(txtPrestamo.Caption) > 0 Or Trim(txtPrestamo.Caption) <> "" Then
        
        crPrestamo = CDbl(txtPrestamo.Caption)
    Else
        
        crPrestamo = 0
    End If
    
    If Val(lblTotAvaluo.Caption) > 0 Or Trim(lblTotAvaluo.Caption) <> "" Then
        
        crAvaluo = CDbl(lblTotAvaluo.Caption)
    Else
        
        crAvaluo = 0
    End If

    
    frmCotizar.Cotizacion crPrestamo, crAvaluo, cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), IIf(Val(txtNombre.Tag) > 0, True, False)
    BringWindowToTop frmCotizar.hWnd
End Sub

Private Sub cmdDiamante_Click()

    If TPrendas.SelectedTab = 1 Then
        lblPiedra.Caption = ""
        lblPuntos.Caption = ""
        lblAvaluoDiamante.Caption = ""
        lblPrestamoDiamante.Caption = ""
        lblCantidadPiedras.Caption = ""
        frmDiamante.Mostrar Me
    End If
    
End Sub


Private Sub cmdFoto_Click()
Dim Cliente As String
Dim IDCliente As Long

    If TPestañas.SelectedTab = 1 Then
        
        If txtNombre.Tag = "" And (txtNombre.text = "" Or txtApellidoPaterno.text = "" Or txtApellidoMaterno.text = "") Then
            
            MsgBox "Seleccione un cliente o introduzca un nombre !!", vbInformation, "Empeño"
            txtNombre.SetFocus
            Exit Sub
        Else
            
            Cliente = Trim(txtNombre.text) & " " & Trim(txtApellidoPaterno.text) & " " & Trim(txtApellidoMaterno.text)
        End If

    ElseIf TPestañas.SelectedTab = 2 Then

        If txtNombre2.Tag = "" And (txtNombre2.text = "" Or txtApellidoPaterno2.text = "" Or txtApellidoMaterno2.text = "") Then
            
            MsgBox "Seleccione un cliente o introduzca un nombre !!", vbInformation, "Autos"
            txtNombre2.SetFocus
            Exit Sub
        Else
            
            Cliente = Trim(txtNombre2.text) & " " & Trim(txtApellidoPaterno2.text) & " " & Trim(txtApellidoMaterno2.text)
        End If

    ElseIf TPestañas.SelectedTab = 3 Then
        
        Cliente = DatosCliente(0).Tag
    Else
        
        Cliente = DatosCliente(1).Tag
    End If

    IDCliente = Val(SacaValor("clientes", "ID", " WHERE CONCAT(Nombre,' ',Apellido)='" & Cliente & "'"))
    
    'Mando a llamar el formulario
    'frmCapturaImagen.Ver Cliente, TPestañas.SelectedTab
    frmCapturaImagenBiometrico.Ver IDCliente, Cliente, TPestañas.SelectedTab
End Sub


Private Sub cmdImprimir_Click()
Dim ID As Long
   
    ID = frmReimpresion.Ver
End Sub

Private Sub cmdLimpiar_Click()
    LimpiaArticulos
    If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
End Sub

Private Sub cmdMostrarCatPrendas_Click()
    MostrarCatPrendas
End Sub

Public Sub MostrarCatPrendas(Optional IDPrendaNueva As Long)
Dim IDPrenda As Long
Dim rcPrenda As New ADODB.Recordset
    
    IDPrenda = 0
    If IDPrendaNueva > 0 Then
        
        IDPrenda = IDPrendaNueva
    
    ElseIf IDPrendaNueva = 0 Then
        
        IDPrenda = frmCatVarios.Mostrar(cmbTipoElec.ItemData(cmbTipoElec.ListIndex))
    End If
    
    If IDPrenda > 0 Then
        
        LimpiaArticulos
        With rcPrenda
            
            .Open "SELECT tipoprenda.Descripcion AS Desc_Familia,marcas.Descripcion AS Desc_Marca,prendaselec.ID AS IDPrenda,prendaselec.Modelo,prendaselec.Minimo,prendaselec.Maximo,prendaselec.IDTipo FROM prendaselec INNER JOIN tipoprenda ON prendaselec.IDFamilia=tipoprenda.ID INNER JOIN marcas ON prendaselec.IDMarca=marcas.ID WHERE prendaselec.ID=" & IDPrenda, dbDatos, adOpenForwardOnly, adLockOptimistic
            
                cmbTipoElec.ListIndex = ComboInformacion(cmbTipoElec, !IDTipo)
                txtFamiliaElec.text = !Desc_Familia
                txtFamiliaElec.Tag = !IDPrenda
                txtMarcaElec.text = !Desc_Marca
                txtModeloElec.text = !Modelo
                txtPrestamooElec.Tag = !Maximo / 100 * ImgSemaforo.Tag
                txtPrestamooElec.text = Format(!Minimo, FMoneda) / 100 * ImgSemaforo.Tag
                lblPrestamoMaximo2.Caption = Format(!Maximo, FMoneda) / 100 * ImgSemaforo.Tag
            .Close
            Set rcPrenda = Nothing
        
        End With
        
    End If
End Sub

Private Sub cmdPagosFijos_Click()
    frmPagosFijos.Show
    BringWindowToTop frmPagosFijos.hWnd
End Sub


Private Sub grdDesempeño_Click(ByVal lRow As Long, ByVal lCol As Long)
    
    If grdDesempeño.Rows > 0 And grdDesempeño.SelectedRow > 0 And (grdDesempeño.SelectedRow < grdDesempeño.Rows) Then
        
        MuestraDatosContrato grdDesempeño.CellItemData(grdDesempeño.SelectedRow, 4), (TPestañas.SelectedTab - 3)
    
    End If
    
End Sub

Private Sub grdDesempeño_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    
    If grdDesempeño.Rows > 0 And grdDesempeño.SelectedRow > 0 And KeyCode = vbKeyDelete Then
        
        If MsgBox("Desea quitar el contrato seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Desempeño") = vbYes Then
            
            grdDesempeño.RemoveRow grdDesempeño.SelectedRow
            If grdDesempeño.Rows > 1 Then
                
                Poner_Totales_Desempeño
                grdDesempeño.ClearSelection
            Else
                
                Limpiar "DESEMPEÑO"
                Limpiar_Leyendas
            End If
        
        End If
        
        txtFolioDesempeño.SetFocus
    End If
End Sub

Private Sub grdEmpeños_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    
    If grdEmpeños.SelectedRow > 0 Then
            
        If Val(grdEmpeños.CellText(grdEmpeños.SelectedRow, 2)) > 0 Then
            
            txtCantidad.text = grdEmpeños.CellText(grdEmpeños.SelectedRow, 2)
            txtCantidad.Tag = grdEmpeños.SelectedRow
            cmbPrenda.ListIndex = ComboInformacion(cmbPrenda, grdEmpeños.CellItemData(grdEmpeños.SelectedRow, 3))
            cmbKilates.ListIndex = ComboInformacion(cmbKilates, grdEmpeños.CellItemData(grdEmpeños.SelectedRow, 5))
            cmbEstado.ListIndex = ComboInformacion(cmbEstado, grdEmpeños.CellItemData(grdEmpeños.SelectedRow, 9))
            txtPesoPiedra.text = grdEmpeños.CellText(grdEmpeños.SelectedRow, 13)
            txtPeso.text = grdEmpeños.CellText(grdEmpeños.SelectedRow, 4)
            txtPrestamoo.text = grdEmpeños.CellText(grdEmpeños.SelectedRow, 7)
            txtAvaluo.text = grdEmpeños.CellText(grdEmpeños.SelectedRow, 6)
            txtObservaciones.text = grdEmpeños.CellText(grdEmpeños.SelectedRow, 11)
            txtPiedras.text = grdEmpeños.CellText(grdEmpeños.SelectedRow, 12)
        End If
        
        grdEmpeños.ClearSelection
    End If
    
End Sub

Private Sub grdRefrendos_Click(ByVal lRow As Long, ByVal lCol As Long)
    
    If grdRefrendos.Rows > 0 And grdRefrendos.SelectedRow > 0 And (grdRefrendos.SelectedRow < grdRefrendos.Rows) Then
        
        MuestraDatosContrato grdRefrendos.CellItemData(grdRefrendos.SelectedRow, 2), (TPestañas.SelectedTab - 3)
    End If
End Sub

Private Sub grdRefrendos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    
    If grdRefrendos.Rows > 0 And grdRefrendos.SelectedRow > 0 And KeyCode = vbKeyDelete Then
        
        If MsgBox("Desea quitar el contrato seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Refrendo") = vbYes Then
            
            grdRefrendos.RemoveRow grdRefrendos.SelectedRow
            If grdRefrendos.Rows > 1 Then
                
                Poner_Totales_Refrendo
                grdRefrendos.ClearSelection
            Else
                
                Limpiar "REFRENDOS"
                Limpiar_Leyendas
            End If
            
        End If
        
    txtFolioRefrendo.SetFocus
    End If
End Sub

Private Sub lblPrestamoDiamante_Change()
    Calcular_Avaluo
End Sub

Private Sub Timer1_Timer()
    If labelContratoAlmoneda.Visible Then
        
        labelContratoAlmoneda.Visible = False
    Else
        
        labelContratoAlmoneda.Visible = True
    End If
        
End Sub

Private Sub Timer2_Timer()
    labelContratoDesemp.Visible = Not labelContratoDesemp.Visible
End Sub


Private Sub tPrendas_TabClick(ByVal lTab As Long)
    
    Select Case lTab

        Case 1
            Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Serie=" & SERIE_A, "Ordenamiento"
            If cmbTipoInteres.ListCount > 0 Then cmbTipoInteres.ListIndex = 0 Else cmbTipoInteres.ListIndex = -1
            
            BanElec = False
            LimpiaArticulos
            txtCantidad.text = "1"
            frmMetales.Visible = True
            frmElectronicos.Visible = False
            grdEmpeños.Clear
            grdEmpeños.Rows = 11
            cmbTipo.ListIndex = ComboInformacion(cmbTipo, 1)
        Case 2
            Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Serie=" & SERIE_D, "Ordenamiento"
            If cmbTipoInteres.ListCount > 0 Then cmbTipoInteres.ListIndex = 0 Else cmbTipoInteres.ListIndex = -1
            
            BanElec = True
            LimpiaArticulos
            frmElectronicos.Visible = True
            frmMetales.Visible = False
            grdEmpeños.Clear
            grdEmpeños.Rows = 11
            cmbTipoElec.ListIndex = 0
    End Select
    
End Sub

Private Sub txtAvaluo_Change()
Dim Kilataje As Boolean, Peso As Boolean, IDTipo As Integer, Avaluo As Double

If Bandera = False Then

    If cmbTipo.ListIndex > -1 Then
        IDTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
    Else
        IDTipo = 0
    End If

    If Val(txtAvaluo.text) > 0 Or Trim(txtAvaluo.text) <> "" Then
        Avaluo = txtAvaluo.text
    Else
        Avaluo = 0
    End If

    Kilataje = Val(SacaValor("Tipo", "Kilataje", " where ID=" & IDTipo))
    Peso = Val(SacaValor("Tipo", "Peso", " where ID=" & IDTipo))

    If Kilataje And Peso Then
        
        txtAvaluo.Locked = True
    Else
    
        txtAvaluo.Locked = False
    End If
End If
End Sub

Private Sub txtAvaluo_GotFocus()
    Seleccionar_Texto txtAvaluo
    Cambiar_Color True, txtAvaluo
End Sub

Private Sub txtAvaluo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAvaluo_LostFocus()
    Cambiar_Color False, txtAvaluo
    txtAvaluo.text = Format(txtAvaluo.text, FMoneda)
End Sub

Private Sub txtAvaluoElec_GotFocus()
    Seleccionar_Texto txtAvaluoElec
    Cambiar_Color True, txtAvaluoElec
End Sub

Private Sub txtAvaluoElec_LostFocus()
    txtAvaluoElec.text = Format(txtAvaluoElec.text, FMoneda)
    Cambiar_Color False, txtAvaluoElec
End Sub

Private Sub txtBeneficiario_GotFocus()
    Seleccionar_Texto txtBeneficiario
    Cambiar_Color True, txtBeneficiario
End Sub

Private Sub txtBeneficiario_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBeneficiario_LostFocus()
    Cambiar_Color False, txtBeneficiario
End Sub


Private Sub txtCantidad_GotFocus()
    Seleccionar_Texto txtCantidad
    Cambiar_Color True, txtCantidad
End Sub

Private Sub txtCantidad_LostFocus()
    Cambiar_Color False, txtCantidad
End Sub



Private Sub txtColorElec_GotFocus()
    Seleccionar_Texto txtColorElec
    Cambiar_Color True, txtColorElec
End Sub

Private Sub txtColorElec_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtColorElec_LostFocus()
    Cambiar_Color False, txtColorElec
End Sub




Private Sub cmbEstado_DropDown()
    Cambiar_Color True, cmbEstado
End Sub

Private Sub cmbEstado_GotFocus()
    Cambiar_Color True, cmbEstado
End Sub

Private Sub cmbEstado_LostFocus()
    Cambiar_Color False, cmbEstado
    grdEmpeños.CancelEdit
End Sub

Private Sub cmbKilates_DropDown()
    Cambiar_Color True, cmbKilates
End Sub

Private Sub cmbKilates_GotFocus()
    Cambiar_Color True, cmbKilates
End Sub

Private Sub cmbKilates_LostFocus()
    Cambiar_Color False, cmbKilates
End Sub

Private Sub cmbMedio_GotFocus()
    Cambiar_Color True, cmbMedio
End Sub

Private Sub cmbMedio_LostFocus()
    Cambiar_Color False, cmbMedio
End Sub

Private Sub cmbTipo_DropDown()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmdAceptar_Click()
Dim Prestamo As Double, Autorizado As Boolean, IDAutorizacion As Long

    Select Case TPestañas.SelectedTab

        Case 1
    
            If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Empeño") = vbYes Then
                If Validar_Empeno = False Then Exit Sub
                
                'Tomo el monto del préstamo
                Prestamo = txtPrestamo.Caption
                     
                Autorizado = True
                VerificaImporte CDbl(Prestamo), Autorizado, IDAutorizacion
                
                If Autorizado Then
                    
                    'Grabo el Empeño
                    Grabar_Empeno Val(txtNombre.Tag), IDAutorizacion
                
                End If
                
            End If
  
        Case 2

            If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Autos") = vbYes Then
                
                If Validar_Empeno_Auto Then Grabar_Empeno_Autos Val(txtNombre2.Tag)
            
            End If
    
        Case 3

            If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Desempeño") = vbYes Then
                
                If Validar_Desempeño Then
                    
                    Grabar_Desempeno
                Else
                    
                    txtFolioDesempeño.SetFocus
                End If
            
            End If

        Case 4

            If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton2, "Refrendo") = vbYes Then
                If Validar_Refrendo Then
                    
                    Grabar_Refrendo
                Else
                    
                    txtFolioRefrendo.SetFocus
                End If
            End If

    End Select

End Sub

'Grabamos el Empeno
Private Sub Grabar_Empeno(ID As Long, IDAutorizacion As Long)
    
    Dim strSql As String, Contrato As Long, Folio As Long, Movimiento As Long, Prestamo As Double, Vencimiento As String, IDEmpeno As Long
    Dim Tasa As Double, Kilates As Integer, Estado As Integer, Indice As Integer, Almacenaje As Double, Seguro As Double, Comision As Double
    Dim Iva As Integer, IDCliente As Long, Dias As Integer, GTOOperacion As Double, VenAlmoneda As Integer, strMarca As String, strModelo As String, strNumSerie As String, strColor As String, strTamaño As String
    Dim strIniciales As String, Codigo As String, Peso As Double, CantidadPiedras As Integer, PesoPiedras As Double, CantidadDiamantes As Integer, PuntosDiamantes As Double, crPrestamoDiamantes As Double, Periodo As Integer, VenPeriodo As Integer, Promocion As Integer, Hora As String
    Dim IDTipo As Long, Serie As Integer, CAT As Double
    
    '***Puntos***
    Dim PuntosAcumulados As Double
    
    Dim IDCotitular As Long
    Dim vTipoGarantia As Integer
    
On Error GoTo Error
    
    Screen.MousePointer = vbHourglass
    
'    If ID = 0 Then
'       ID = Grabar_Cliente
'       IDCliente = ID
'    Else
'       IDCliente = Actualizar_Cliente(ID, Trim(txtDireccion.Tag), Trim(txtApellidos.Tag))
'    End If
    
    '----------------------------------------------------
    'Grabar Cliente
    '----------------------------------------------------
    ClienteEmp.Grabar
    ID = ClienteEmp.ID
    IDCliente = ID

    '----------------------------------------------------
    'Grabar Cotitular
    '----------------------------------------------------
    IDCotitular = 0
    If Trim(txtResponsable.text) <> "" And Trim(txtCotitularApellidoPaterno.text) <> "" And Trim(txtCotitularApellidoMaterno.text) <> "" Then
        If CotitularEmp.Valida = True Then
            CotitularEmp.Grabar
            IDCotitular = CotitularEmp.ID
        Else
            MsgBox "Datos incompletos del CoTitular.", vbCritical, "Empeño de Inmuebles"
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    '----------------------------------------------------
    
    
    
    
    
    'Actualizo el Numero de contratos del cliente
    dbDatos.Execute "UPDATE clientes SET Boletas=Boletas+1 WHERE ID=" & IDCliente
    
    'Saco el Numero de Contrato
    Contrato = Regresa_NumContrato(False, IIf(TPrendas.SelectedTab, SERIE_A, SERIE_D))
    Regresa_NumContrato True, IIf(TPrendas.SelectedTab, SERIE_A, SERIE_D)
    Serie = IIf(TPrendas.SelectedTab, SERIE_A, SERIE_D)
    
    'Folio
    Folio = Regresa_NumContrato(False, SERIE_C)
    Regresa_NumContrato True, SERIE_C
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
'''''    'Tomo el tipo de promocion
'''''    Select Case cmbPromocion.ListIndex
'''''    Case 0
'''''        Promocion = 0
'''''    Case 1
'''''        Promocion = 15
'''''    Case 2
'''''        Promocion = 30
'''''    Case 3
'''''        Promocion = 1
'''''    Case 4
'''''        Promocion = 2
'''''    Case 5
'''''        Promocion = 3
'''''    Case 6
'''''        Promocion = 4
'''''    Case 7
'''''        Promocion = 5
'''''    Case 8
'''''        Promocion = 20
'''''    Case 9
'''''        Promocion = 50
'''''    End Select
    Promocion = IIf(cmbPromocion.ListIndex >= 0, cmbPromocion.ItemData(cmbPromocion.ListIndex), 0)
    
    Select Case cmbPeriodo.text
    Case "MENSUAL"
        Periodo = 30
    Case "QUINCENAL"
        Periodo = 15
    Case "SEMANAL"
        Periodo = 7
    Case "DIARIA"
        Periodo = 1
    End Select
    VenPeriodo = Val(cmbPlazos.text)
    Vencimiento = lblVencimiento.Caption
    Tasa = CDbl(Mid(lblTasa.Caption, 1, Len(lblTasa.Caption) - 1))
    CAT = Val(lblTasa.Tag)
    Prestamo = CDbl(txtPrestamo.Caption)
    Almacenaje = CDbl(Mid(lblAlmacenaje.Caption, 1, Len(lblAlmacenaje.Caption) - 1))
    Seguro = CDbl(Mid(lblSeguro.Caption, 1, Len(lblSeguro.Caption) - 1))
    Iva = CDbl(Mid(lblIva.Caption, 1, Len(lblIva.Caption) - 1))
    GTOOperacion = Regresa_Valor_BD("Operacion")
    Comision = Regresa_Valor_BD("Comision")
    VenAlmoneda = Regresa_Valor_BD("VenAlmoneda")
    
    
    strIniciales = Iniciales(Trim(txtNombre.text), Trim(txtApellidoPaterno.text) & " " & Trim(txtApellidoMaterno.text))
        
    If cmbEstado.ListIndex = -1 Then
        Estado = 0
    Else
        Estado = cmbEstado.ItemData(cmbEstado.ListIndex)
    End If
   
    strSql = "INSERT INTO empeno (Fecha,Movimiento,NumContrato,Folio,Prestamo,Avaluo,Origen,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Responsable,Valuador,Notas,Tasa,Almacenaje,Seguro,Operacion,Comision,IVA,Periodo,Venperiodo,VenAlmoneda,Tipointeres,TipoTasa,IDSucursal,IDUsuario,IDAutorizacion,NumBolsa,Ubicacion,Caja,Cajon,Fila,IDUsuarioAutoriza,TipoAutoriza,Promocion,PrestamoInicial,Beneficiario,IDCotitular,Cat) VALUES " & _
           "('" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & Contrato & "," & Folio & "," & ConvMoneda(Prestamo) & "," & ConvMoneda(lblTotAvaluo.Caption) & "," & OD_EMPENO & ",'" & Format(Vencimiento, "YYYY/MM/DD") & "'," & Folio & "," & IIf(cmbTipoInteres.text = "FIJA", SERIE_C, Serie) & ",'" & NombrePc & "'," & IDCliente & "," _
           & "'" & Trim(txtResponsable.text & " " & txtCotitularApellidoPaterno.text & " " & txtCotitularApellidoMaterno.text) & "','" & frmMDI.Usuario & "','" & Trim(txtNotas.text) & "'," & ConvMoneda(Tasa) & "," & ConvMoneda(Almacenaje) & "," & ConvMoneda(Seguro) & "," & ConvMoneda(GTOOperacion) & "," & ConvMoneda(Comision) & "," & ConvMoneda(Iva) & _
           "," & Periodo & "," & VenPeriodo & "," & VenAlmoneda & ",'" & cmbTipoInteres.text & "','" & cmbPeriodo.text & "'," & frmMDI.IDSucursal & "," & frmMDI.IDUsuario & "," & IDAutorizacion & ",'" & Trim(txtNumBolsa.text) & "','','','',''," & Me.IDUsuarioAutoriza & "," & Me.TipoAutorizacion & "," & Promocion & "," & ConvMoneda(Prestamo) & ",'" & _
           Trim(txtBeneficiario.text) & "'," & IDCotitular & "," & CAT & ")"
   
    Err.Clear
    dbDatos.Execute strSql
    
    'Saco el ID del Empeño
    IDEmpeno = SacaValor("empeno", "MAX(ID)")
    
    'MLD-MODIF.
    GuardarDatosLavadoDinero IDEmpeno, "empeno", MLD_INSTRUMENTO_MONETARIO, MLD_PRESTAMO, 0, vTipoAlerta.ID, vTipoAlerta.Descripcion
    
    
    'Grabamos el detalle del empeño
    With grdEmpeños
    
        For Indice = 1 To .Rows
            
            If grdEmpeños.CellText(Indice, 1) <> "" Then
            
                Codigo = CreaCodigoBarras(Trim(Format(frmMDI.IDSucursal, "000")), ENTRADAEMPENO, Trim(Contrato), Indice)
                Kilates = RegresaKilates(IIf(grdEmpeños.CellText(Indice, 1) = "ORO", grdEmpeños.CellText(Indice, 5), ""), grdEmpeños.CellText(Indice, 1))
                Peso = IIf(Val(.CellText(Indice, 4)) = 0 Or Trim(.CellText(Indice, 4)) = "", 0, .CellText(Indice, 4))
                                
                CantidadPiedras = IIf(Val(.CellText(Indice, 12)) = 0 Or Trim(.CellText(Indice, 12)) = "", 0, .CellText(Indice, 12))
                PesoPiedras = IIf(Val(.CellText(Indice, 13)) = 0 Or Trim(.CellText(Indice, 13)) = "", 0, .CellText(Indice, 13))
                
                CantidadDiamantes = IIf(Val(.CellText(Indice, 14)) = 0 Or Trim(.CellText(Indice, 14)) = "", 0, .CellText(Indice, 14))
                PuntosDiamantes = IIf(Val(.CellText(Indice, 15)) = 0 Or Trim(.CellText(Indice, 15)) = "", 0, .CellText(Indice, 15))
                crPrestamoDiamantes = IIf(Val(.CellText(Indice, 16)) = 0 Or Trim(.CellText(Indice, 16)) = "", 0, .CellText(Indice, 16))
                
                strMarca = Trim(.CellText(Indice, 17))
                strModelo = Trim(.CellText(Indice, 18))
                strNumSerie = Trim(.CellText(Indice, 19))
                strColor = Trim(.CellText(Indice, 20))
                strTamaño = Trim(.CellText(Indice, 21))
                
                '---- MLD-MODIF. Sacar el valor de TipoGarantia segun el Tipo de Prenda ----
                vTipoGarantia = 0
                vTipoGarantia = Val(SacaValor("tipo", "IdTipoGarantia", " WHERE ID=" & Val(.CellItemData(Indice, 1))))
                '---------------------------------------------------------------------------
                
                dbDatos.Execute "INSERT INTO detallesempeno (IDEmpeno,Codigo,Tipo,Cantidad,Articulo,Peso,Kilates,Avaluo,Prestamo,Estado,Origen,Destino,TipoPrenda,Observaciones,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante,Marca,Modelo,Serie,Color,Tamano,IdTipoGarantia) VALUES (" & _
                                IDEmpeno & ",'" & Trim(Codigo) & "'," & .CellItemData(Indice, 1) & "," & Val(.CellText(Indice, 2)) & ",'" & Trim(UCase(.CellText(Indice, 3))) & "'," & ConvMoneda(Peso) & "," & _
                                Kilates & "," & ConvMoneda(.CellText(Indice, 6)) & "," & ConvMoneda(.CellText(Indice, 7)) & ",'" & Trim(UCase(.CellText(Indice, 9))) & "'," & ENTRADAEMPENO & ",0," & Val(.CellItemData(Indice, 3)) & ",'" & Trim(.CellText(Indice, 11)) & "'," & CantidadPiedras & "," & ConvMoneda(PesoPiedras) & "," & CantidadDiamantes & "," & ConvMoneda(PuntosDiamantes) & "," & ConvMoneda(crPrestamoDiamantes) & ",'" & strMarca & "','" & strModelo & "','" & strNumSerie & "','" & strColor & "','" & strTamaño & "'," & vTipoGarantia & ")"
            
                IDTipo = .CellItemData(Indice, 1)
            End If
            
        Next Indice
    
    End With
    
    'IDTipoPrenda Empeno
    dbDatos.Execute "UPDATE empeno SET IDTipoPrenda = " & IDTipo & " WHERE ID = " & IDEmpeno
    
    'Tomo la Hora
    Hora = Time
    
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','201701'," & ConvMoneda(Prestamo) & "," & TIPO_CARGO & "," & IIf(cmbTipoInteres.text = "FIJA", SERIE_C, SERIE_A) & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110150'," & ConvMoneda(Prestamo) & "," & TIPO_ABONO & "," & IIf(cmbTipoInteres.text = "FIJA", SERIE_C, SERIE_A) & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

'''    'Grabamos abono 199450
'''    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'''                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','199450'," & ConvMoneda(Prestamo) & "," & TIPO_ABONO & "," & IIf(cmbTipoInteres.text = "FIJA", SERIE_C, SERIE_A) & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Muestro el Contrato y el Folio
    lblFolio.Caption = Contrato
    lblContrato.Caption = Contrato
             
    '***Puntos***
    If TarjetaPuntos.CuentaFrecuente.Folio <> "" Then
    
        dbDatos.Execute "UPDATE empeno SET SaldoPuntosAnteriorEmp = " & TarjetaPuntos.CuentaFrecuente.Puntos & " WHERE ID = " & IDEmpeno
    
        PuntosAcumulados = TarjetaPuntos.Acumula_Puntos(TipoMovimiento.Empeno, frmMDI.IDUsuario, CCur(Prestamo), Contrato)
        MsgBox "Puntos Acumulados: " & PuntosAcumulados, vbInformation Or vbOKOnly
            
        dbDatos.Execute "UPDATE empeno SET PuntosAcumuladosEmp=" & PuntosAcumulados & ",SaldoPuntosActualEmp=" & TarjetaPuntos.CuentaFrecuente.Puntos & ",IDTarjetaEmp=" & TarjetaPuntos.CuentaFrecuente.IDCuenta & " WHERE ID = " & IDEmpeno
        
    End If
             
    Sleep 1000
    'Imprimir_Boleta_CR IDEmpeno
    'Imprimir_Boleta_Profeco IDEmpeno, , True
    'Imprimir_Boleta_CR_Caidas IDEmpeno, False, True, False
    Imprimir_Boleta_CR_Caidas IDEmpeno, False, True, False
        
    Limpiar "Empeno"
    txtPrestamo.Caption = ""
    grdEmpeños.CancelEdit
    grdEmpeños.ClearItems
    grdEmpeños.ClearSelection
    IDUsuarioAutoriza = 0
    TipoAutorizacion = 0
    
    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\APAGADOS.bmp")
    ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
    ImgSemaforo.ToolTipText = ""
    lblAlmacenaje.Caption = Format(Regresa_Valor_BD("Almacenaje"), "0.00") & "%"
    lblSeguro.Caption = Format(Regresa_Valor_BD("Seguro"), "0.00") & "%"
    lblIva.Caption = Format(Regresa_Valor_BD("IVA"), "0.00") & "%"
    txtPrestamo.Caption = "0.00"
    lblTotAvaluo.Caption = "0.00"
    lblAutorizacion.Caption = ""
    lblAutorizacion.Visible = False
    cmbPromocion.ListIndex = 0
    cmbTipoInteres.ListIndex = 0
    cmbTipoInteres_Click
    cmbTipo.ListIndex = 0
    txtNotas.text = Regresa_Valor_BD("Notas")
    Default 1
    txtNombre.SetFocus
    
    '***Puntos***
    Limpiar_Tarjeta
    
    'MLD-MODIF. ----------------------
    ClienteEmp.Limpiar
    CotitularEmp.Limpiar
    InicializarAlerta vTipoAlerta, MLD_PRESTAMO
    cmdAlerta.Enabled = False
    '---------------------------------
    
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error:
    Maneja_Error Err
    Resume
    Screen.MousePointer = vbDefault
End Sub

'Grabamos el desempeño
Private Sub Grabar_Desempeno()

    Dim Serie As Integer, Movimiento As Long, Pago As Double, Cont As Integer, Folio As Long
    Dim crEfectivo As Double, crImporteTotal As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double
    Dim ImportePerdida As Double, ImporteIvaPerdida As Double, FolioNota As Long, IDEmpeno As Long, IDEmpeño As Long, strIniciales As String, ContratoAlmoneda As Integer, crImporteAlmoneda As Double, Hora As String
    Dim rcArticulos As New ADODB.Recordset
   '**** CARGO GPS Y SEGURO AUTO
    Dim rcCargoGPS As Double, rcCargoSeguroAuto As Double, EnCirculacion As Integer, IvaCargoRenta As Double

    '***Puntos***
    Dim Puntos As Currency, IDCliente As Long, Contrato As Long, crPuntos As Currency, PuntosAcumulados As Double

On Error GoTo Error
    
    crImporteTotal = CDbl(TotalDesempeño.Caption)
    crEfectivo = frmEfectivo.RegresaCambio(crImporteTotal, 1)
    If crEfectivo < crImporteTotal Then Exit Sub
    CalculaCambio crEfectivo, crImporteTotal, 1
    
    For Cont = 1 To grdDesempeño.Rows - 1
        
        ImportePerdida = 0
        ImporteIvaPerdida = 0
        
        '***Puntos***
        Puntos = Val(grdDesempeño.CellText(Cont, 13))
        crPuntos = Val(grdDesempeño.CellText(Cont, 15))
        
        'Tomo el Importe de Boleta Perdida
        ImportePerdida = grdDesempeño.CellText(Cont, 10)
        If ImportePerdida > 0 Then
                                        
            ImporteIvaPerdida = ImportePerdida - (ImportePerdida / (1 + (Regresa_Valor_BD("IVA") / 100)))
            ImportePerdida = ImportePerdida - ImporteIvaPerdida
            
        End If
        
        'Tomo los Intereses
        ContratoAlmoneda = Val(grdDesempeño.CellText(Cont, 9))
        Serie = grdDesempeño.CellItemData(Cont, 1)
        crIntereses = CDbl(grdDesempeño.CellText(Cont, 5))
        crAlmacenaje = CDbl(grdDesempeño.CellText(Cont, 6))
        crSeguro = CDbl(grdDesempeño.CellText(Cont, 7))
        crMoratorios = CDbl(grdDesempeño.CellText(Cont, 11))
        crIva = CDbl(grdDesempeño.CellText(Cont, 8))
        Folio = CLng(grdDesempeño.CellText(Cont, 4))
        Pago = CDbl(grdDesempeño.CellText(Cont, 2))
        IvaCargoRenta = CDbl(grdDesempeño.CellText(Cont, 19))
        crImporteAlmoneda = 0
        
        EnCirculacion = 2
            If Frame4.Visible = True Then
                If chkCirculacionDes.Value = 1 Then
                    EnCirculacion = 1
                    If Not txtCargoSeguroDes.text = "" Then
                        rcCargoSeguroAuto = CDbl(txtCargoSeguroDes.text)
                    End If
                    If Not txtCargoGPSDes.text = "" Then
                           rcCargoGPS = CDbl(txtCargoGPSDes.text)
                    End If
                End If
            End If
        
        
        '***Puntos***
        IDCliente = CDbl(grdDesempeño.CellText(Cont, 17))
        
        'Folio Notas
        FolioNota = Regresa_Movimiento(False, "FolioNotas")
        Regresa_Movimiento True, "FolioNotas"
    
        'Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
        
        'Actualizo el registro
        dbDatos.Execute "UPDATE empeno SET PC='" & NombrePc & "',Pago=" & ConvMoneda(Pago) & ",Intereses=" & ConvMoneda(crIntereses) & ",Destino=" & D_DESEMPEÑO & ",FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',IDUsuarioMov=" & frmMDI.IDUsuario & ",Pagado=1,ImporteIva=" & ConvMoneda(crIva + ImporteIvaPerdida) & ",ImportePerdida=" & ConvMoneda(ImportePerdida) & ",ImporteAlmacenaje=" & ConvMoneda(crAlmacenaje) & ",ImporteSeguro=" & ConvMoneda(crSeguro) & ",ImporteMoratorios=" & ConvMoneda(crMoratorios) & ",FolioNota=" & FolioNota & ",Efectivo=" & ConvMoneda(crEfectivo) & ", ImporteSeguroAuto =" & rcCargoSeguroAuto & ", ImporteRentaGPS = " & rcCargoGPS & ", Circulando =" & EnCirculacion & ", ImporteIVAGPSSeguro = " & IvaCargoRenta & " WHERE ID=" & Val(grdDesempeño.CellItemData(Cont, 4))
  
        strIniciales = SacaValor("empeno LEFT JOIN clientes ON empeno.IDCliente=clientes.ID", "clientes.Iniciales", " WHERE empeno.ID=" & grdDesempeño.CellItemData(Cont, 4))
        IDEmpeño = grdDesempeño.CellItemData(Cont, 4)
        
        'Tomo la Hora
        Hora = Time
        
        'Cuenta de Empeños****
        'Cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(Pago) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
        'Abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','201750'," & ConvMoneda(Pago) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
        'Cuenta de intereses***
        'Cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crIntereses) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
  
        'Abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','520450'," & ConvMoneda(crIntereses) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
        'Cuenta de Almacenaje***
        If crAlmacenaje > 0 Then
    
            'Cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crAlmacenaje) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
            'Abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','670350'," & ConvMoneda(crAlmacenaje) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
   
        End If
        ' Cargo de GPS y Seguro Auto
        If EnCirculacion = 1 Then
                'Cuenta de Renta GPS
                If rcCargoGPS > 0 Then
                    'Se agrega el abono a Renta deGPS
                      'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño Renta GPS'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(rcCargoGPS) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño Renta GPS'," & Movimiento & "," & Folio & ",'" & strIniciales & "','818150'," & ConvMoneda(rcCargoGPS) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

                End If
                'Cuenta de Seguro Auto
                If rcCargoSeguroAuto > 0 Then
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño Renta GPS'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(rcCargoSeguroAuto) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño Renta GPS'," & Movimiento & "," & Folio & ",'" & strIniciales & "','828250'," & ConvMoneda(rcCargoSeguroAuto) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

                End If
                If IvaCargoRenta > 0 Then
                   crIva = crIva - IvaCargoRenta
                   'Cargo
                   dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño Renta GPS'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(IvaCargoRenta) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                   
                   'Abono
                   dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño Renta GPS'," & Movimiento & "," & Folio & ",'" & strIniciales & "','120150'," & ConvMoneda(IvaCargoRenta) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
     
                End If
        End If
          
        'Cuenta de Seguro***
        If crSeguro > 0 Then
      
            'Cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crSeguro) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
            'Abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','680350'," & ConvMoneda(crSeguro) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
        End If
        
        'Cuenta de Moratorios***
        If crMoratorios > 0 Then
      
            'Cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crMoratorios) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
            'Abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','690350'," & ConvMoneda(crMoratorios) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
        End If
    
        'Cuenta de Iva***
        If crIva > 0 Then
    
            'Cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crIva) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
            'Abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','120150'," & ConvMoneda(crIva) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
        End If
    
        If ImportePerdida > 0 Then
        
            'Grabamos el cargo de la boleta
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Boleta perdida'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(ImportePerdida) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
  
            'Grabamos el abono de la boleta
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Boleta perdida'," & Movimiento & "," & Folio & ",'" & strIniciales & "','530150'," & ConvMoneda(ImportePerdida) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
            If ImporteIvaPerdida > 0 Then
        
                'Grabamos el cargo del iva de la boleta
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                                Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Boleta perdida'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(ImporteIvaPerdida) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
                'Grabamos el abono del iva de la boleta
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                                Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Boleta perdida'," & Movimiento & "," & Folio & ",'" & strIniciales & "','120150'," & ConvMoneda(ImporteIvaPerdida) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
            End If
        
        End If
    
'''        'Grabamos el cargo a 199401
'''        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
'''                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','199401'," & ConvMoneda(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva + ImportePerdida + ImporteIvaPerdida + Pago) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        'Grabamos el cargo a 199401
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(Pago) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
            
        '***Puntos***
        If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(Val(IDCliente)) Then
            
            dbDatos.Execute "UPDATE empeno SET saldopuntosanterior = " & TarjetaPuntos.CuentaFrecuente.Puntos & " WHERE ID = " & IDEmpeño
            
            'PuntosAcumulados = TarjetaPuntos.Acumula_Puntos(Desempeno, frmMDI.IDUsuario, ConvMoneda(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva + ImportePerdida + ImporteIvaPerdida + Pago), Contrato)
            PuntosAcumulados = TarjetaPuntos.Acumula_Puntos(Desempeno, frmMDI.IDUsuario, Val(grdDesempeño.CellText(Cont, 3)), Contrato)
            
            MsgBox "Puntos acumulados por el Desempeño: " & PuntosAcumulados, vbOKOnly Or vbInformation
        
            dbDatos.Execute "UPDATE empeno SET puntosacumulados = " & PuntosAcumulados & " WHERE ID = " & IDEmpeño
            
            'descontamos los puntos
            If crPuntos > 0 Then
        
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Redencion Puntos Desempeño'," & Movimiento & "," & Folio & ",'" & _
                    strIniciales & "','905501'," & ConvMoneda(crPuntos) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
'''                'Grabamos el cargo a 199450
'''                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
'''                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Redencion Puntos Desempeño'," & Movimiento & "," & Folio & ",'" & _
'''                    strIniciales & "','199450'," & ConvMoneda(crPuntos) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                            
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Redencion Puntos Desempeño'," & Movimiento & "," & Folio & ",'" & _
                    strIniciales & "','110150'," & ConvMoneda(crPuntos) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                TarjetaPuntos.Redimir_Puntos Desempeno, CLng(Puntos), grdDesempeño.CellText(Cont, 14), frmMDI.IDUsuario, Contrato
                    
                dbDatos.Execute "UPDATE empeno SET descuentoxpuntos = " & ConvMoneda(crPuntos) & ",puntosusados = " & CLng(Puntos) & " WHERE ID = " & IDEmpeño
                
            End If
            
            dbDatos.Execute "UPDATE empeno SET SaldoPuntosActual = " & TarjetaPuntos.CuentaFrecuente.Puntos - Val(Puntos) & ",IDTarjeta = " & TarjetaPuntos.CuentaFrecuente.IDCuenta & " WHERE ID = " & IDEmpeño
            
        End If
            
'******************************************************************************************************************************************
        If ContratoAlmoneda = 1 Then
            rcArticulos.Open "SELECT d.ID AS IDPrendaInventario,d.Codigo,d.Tipo,d.Cantidad,d.Descripcion AS Articulo,d.Peso,d.Kilates,d.Precio AS Avaluo,d.Costo AS Prestamo,d.Estado,d.Serie,d.TipoPrenda,d.Observaciones,d.PesoPiedras,d.Puntos,d.CantidadPiedras,d.CantidadDiamantes,d.PrestamoDiamante,d.Marca,d.Modelo,d.Color,d.Tamano,t.IdTipoGarantia " & _
                             "FROM detallesentradainventario d LEFT JOIN tipo t ON d.Tipo = t.Id WHERE d.Cantidad>0 AND d.IDEmpeno=" & IDEmpeño, dbDatos, adOpenForwardOnly, adLockReadOnly
          
            With rcArticulos
                While Not .EOF
                    'Lo doy de baja del Inventario
                    dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=0,TipoSalida=" & SALIDAVENTAPIGNORANTE & " WHERE ID=" & rcArticulos!IDPrendaInventario
                    
                    'Tomo el Importe para darlo de baja del Inventario
                    crImporteAlmoneda = crImporteAlmoneda + Val(rcArticulos!Prestamo)
                   
                    .MoveNext
                Wend
                
            End With
            rcArticulos.Close
                
            'Grabamos el Abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Desempeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','620350'," & ConvMoneda(crImporteAlmoneda) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        End If

'******************************************************************************************************************************************
        If Serie = 2 Then
             'Imprimo el recibo
        Imprimir_Nota_Auto IDEmpeño, D_DESEMPEÑO, 0, frmMDI.IDUsuario

        Else
             'Imprimo el recibo
        Imprimir_Nota IDEmpeño, D_DESEMPEÑO, 0, frmMDI.IDUsuario

        End If
               
        If EnCirculacion = 1 Then
            'Imprimo Recibo de GPS y Seguro Auto
            Imprimir_Nota_GPS_Seguro IDEmpeño, rcCargoGPS, rcCargoSeguroAuto, IvaCargoRenta, D_DESEMPEÑO, 0, frmMDI.IDUsuario
        End If
    Next Cont

    MsgBox "Los contratos se desempeñaron correctamente !!", vbInformation, "Desempeño"
    
    Limpiar "DESEMPEÑO"
    grdDesempeño.Clear
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

'Grabamos los datos del refrendo
Private Sub Grabar_Refrendo()
    Dim crAdeudo As Double, strSql1 As String, strSql2 As String, strSql3 As String, Indice As Integer
    Dim Movimiento As Long, Serie As Integer, Vencimiento As String, Pago As Double, crPrestamoPrenda As Double
    
    Dim crPrestamo As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double
    Dim FolioAnterior As Long, FolioNuevo As Long, Abono As Double, dia As Integer, Dias As Long
    Dim crEfectivo As Double, crImporteTotal As Double, ImportePerdida As Double, ImporteIvaPerdida As Double, IDEmpeno As Long, FolioNota As Long, IDEmpenoAnterior As Long, ContratoAlmoneda As Integer, crImporteAlmoneda As Double, Hora As String
    Dim rcEmpeño As New ADODB.Recordset
    Dim rcArticulos As New ADODB.Recordset, Vencido As Integer
    Dim rcArticulosAuto As New ADODB.Recordset
    '**** CARGO GPS Y SEGURO AUTO
    Dim rcCargoGPS As Double, rcCargoSeguroAuto As Double, EnCirculacion As Integer, IvaCargoRenta As Double

    '***Puntos***
    Dim Puntos As Currency, crPuntos As Currency, Contrato As Long, PuntosAcumulados As Double
    
    Dim TipoPago As Integer, strSql4 As String
    Dim DigitosTarjeta As String

    
On Error GoTo Error
    
    'MLD-MODIF. ---------- Cuando es pago Solo EFECTIVO
    TipoPago = Val(SacaValor("mld_instr_monetarios", "Id", " WHERE RegDefault=1"))
    DigitosTarjeta = ""
    '---------------------

    'Checo si el contrato no se ha pasado a Almoneda y si tiene el permiso para refrendarlo
    If Almoneda = 1 Then
        
        frmPasswords.ConexSuc = 0
        frmPasswords.DescuentoVentas = 0
        frmPasswords.PrecioVitrina = 0
        frmPasswords.Ventas = 0
        frmPasswords.ModificaPrecio = 0
        frmPasswords.ModificaCorte = 0
        frmPasswords.InteresDesempeño = 0
        frmPasswords.InteresRefrendo = 0
        frmPasswords.HacerCorte = 0
        frmPasswords.RecalculoPrecios = 0
        frmPasswords.AutorizaPrestamo = 0
        frmPasswords.Cancel = 0
        frmPasswords.CancelaCierre = 0
        frmPasswords.Vencido = 1
        
        If frmPasswords.Password(CANCELACION, 1) = False Then Exit Sub
        
    End If
    
    crImporteTotal = CDbl(TotalRefrendo.Caption)
    crEfectivo = frmEfectivo.RegresaCambio(crImporteTotal, 2)
    If crEfectivo < crImporteTotal Then Exit Sub
    CalculaCambio crEfectivo, crImporteTotal, 2
    
    For Indice = 1 To grdRefrendos.Rows - 1
    
            Screen.MousePointer = vbHourglass
                                    
            ImportePerdida = 0
            ImporteIvaPerdida = 0
            Pago = 0
            
            '***Puntos***
            Puntos = Val(grdRefrendos.CellText(Indice, 13))
            crPuntos = Val(grdRefrendos.CellText(Indice, 15))
            
            Vencido = grdRefrendos.CellText(Indice, 18)
            
            'Tomo el Importe de Boleta Perdida
            ImportePerdida = grdRefrendos.CellText(Indice, 10)
            If ImportePerdida > 0 Then
                ImporteIvaPerdida = ImportePerdida - (ImportePerdida / (1 + (Regresa_Valor_BD("IVA") / 100)))
                ImportePerdida = ImportePerdida - ImporteIvaPerdida
            End If
            
            'Checo si tiene Abono
            If Val(grdRefrendos.CellText(Indice, 3)) > 0 Or Trim(grdRefrendos.CellText(Indice, 3)) <> "" Then
                
                Pago = CDbl(grdRefrendos.CellText(Indice, 3))
                    
                'Verifico el abono que sea mayor o igual al configurado en parámetros
                If Pago > 0 Then
                    
                    If Pago < Val(Regresa_Valor_BD("AbonoMinimo")) Then
                        
                        MsgBox "El importe del abono es menor al autorizado !!", vbInformation, "Refrendos"
                        Screen.MousePointer = vbDefault
                        Exit Sub
                    
                    End If
                    
                End If
            End If
            
            'Tomo los Valores
            ContratoAlmoneda = Val(grdRefrendos.CellText(Indice, 9))
            Serie = grdRefrendos.CellItemData(Indice, 3)
            crIntereses = CDbl(grdRefrendos.CellText(Indice, 5))
            crAlmacenaje = CDbl(grdRefrendos.CellText(Indice, 6))
            crSeguro = CDbl(grdRefrendos.CellText(Indice, 7))
            crMoratorios = CDbl(grdRefrendos.CellText(Indice, 11))
            crIva = CDbl(grdRefrendos.CellText(Indice, 8))
            crPrestamo = CDbl(grdRefrendos.CellText(Indice, 2))
            IvaCargoRenta = CDbl(grdRefrendos.CellText(Indice, 20))
            crAdeudo = crPrestamo - Pago
            crImporteAlmoneda = 0
            EnCirculacion = 2
            If Frame3.Visible = True Then
                If chkCirculacionRef.Value = 1 Then
                    EnCirculacion = 1
                    If Not txtCargoSeguro.text = "" Then
                        rcCargoSeguroAuto = CDbl(txtCargoSeguro.text)
                    End If
                    If Not txtCargoGPS.text = "" Then
                           rcCargoGPS = CDbl(txtCargoGPS.text)
                    End If
                End If
            End If
           
            
            'Leemos los datos del Empeno original
            rcEmpeño.Open "SELECT empeno.*,clientes.Iniciales FROM empeno LEFT JOIN clientes on empeno.IDCliente=clientes.ID WHERE empeno.ID=" & grdRefrendos.CellItemData(Indice, 2), dbDatos, adOpenForwardOnly, adLockOptimistic
            FolioAnterior = rcEmpeño!Folio
            IDEmpenoAnterior = rcEmpeño!ID
            
            '***Puntos***
            If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(rcEmpeño!IDCliente) Then
            
                dbDatos.Execute "UPDATE empeno SET saldopuntosanterior = " & TarjetaPuntos.CuentaFrecuente.Puntos & " WHERE ID = " & IDEmpenoAnterior
            
                PuntosAcumulados = TarjetaPuntos.Acumula_Puntos(IIf(Vencido = 1, RefrendoExt, Refrendo), frmMDI.IDUsuario, Val(grdRefrendos.CellText(Indice, 4)), Contrato)
                MsgBox "Puntos Acumulados por el Refrendo: " & PuntosAcumulados, vbOKOnly Or vbInformation
                
                dbDatos.Execute "UPDATE empeno SET puntosacumulados = " & PuntosAcumulados & " WHERE ID = " & IDEmpenoAnterior
                
                If Puntos > 0 Then
                    TarjetaPuntos.Redimir_Puntos Refrendo, CLng(Puntos), grdRefrendos.CellText(Indice, 14), frmMDI.IDUsuario, Contrato
                    
                    dbDatos.Execute "UPDATE empeno SET descuentoxpuntos = " & ConvMoneda(crPuntos) & " WHERE ID = " & IDEmpenoAnterior
                    dbDatos.Execute "UPDATE empeno SET puntosusados = " & CLng(Puntos) & " WHERE ID = " & IDEmpenoAnterior
                End If
                
                dbDatos.Execute "UPDATE empeno SET SaldoPuntosActual = " & TarjetaPuntos.CuentaFrecuente.Puntos - Val(Puntos) & ",IDTarjeta = " & TarjetaPuntos.CuentaFrecuente.IDCuenta & " WHERE ID = " & IDEmpenoAnterior
                
            End If
            
            'Tomo el Nuevo Folio
            FolioNuevo = Regresa_NumContrato(False, SERIE_C)
            Regresa_NumContrato True, SERIE_C
            
            'Saco el Movimiento
            Movimiento = Regresa_Movimiento(False)
            Regresa_Movimiento True
            
            'Folio Notas
            FolioNota = Regresa_Movimiento(False, "FolioNotas")
            Regresa_Movimiento True, "FolioNotas"
            
            'MLD-MODIF. Ponemos el refrendo
            dbDatos.Execute "UPDATE empeno SET PC='" & NombrePc & "',Pago=" & ConvMoneda(Pago) & ",Destino=" & IIf(ContratoAlmoneda = 1, D_VENTA, OD_REFRENDO) & ",FolioDestino=" & FolioNuevo & ",Pagado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',IDUsuarioMov=" & frmMDI.IDUsuario & ",Intereses=" & ConvMoneda(crIntereses) & ",Importeiva=" & ConvMoneda(crIva + ImporteIvaPerdida) & ",ImportePerdida=" & ConvMoneda(ImportePerdida) & ",ImporteAlmacenaje=" & ConvMoneda(crAlmacenaje) & ",ImporteSeguro=" & ConvMoneda(crSeguro) & ",ImporteMoratorios=" & ConvMoneda(crMoratorios) & ",FolioNota=" & FolioNota & ",Efectivo=" & ConvMoneda(crEfectivo) & ",IdInstrumentoMonetario=" & TipoPago & ",UltDigitosTarj='" & DigitosTarjeta & "', ImporteSeguroAuto =" & rcCargoSeguroAuto & ", ImporteRentaGPS = " & rcCargoGPS & ", Circulando =" & EnCirculacion & ", ImporteIVAGPSSeguro = " & IvaCargoRenta & " WHERE ID=" & grdRefrendos.CellItemData(Indice, 2)
            
            'Saco la nueva fecha de vencimiento
''''            If Day(Date) = Regresa_Ultimo_Dia_Mes(Date) Then
''''
''''                Dias = Val(rcEmpeño!Periodo) * Val(rcEmpeño!VenPeriodo)
''''                Vencimiento = IIf(rcEmpeño!Periodo = 30, Format(DateAdd("M", rcEmpeño!VenPeriodo, Date), "DD/MM/YYYY"), Format(DateAdd("D", Dias - 1, Date), "DD/MM/YYYY"))
''''            Else
''''
''''                Dias = Val(rcEmpeño!Periodo) * Val(rcEmpeño!VenPeriodo)
''''                Vencimiento = IIf(rcEmpeño!Periodo = 30, Format(DateAdd("M", rcEmpeño!VenPeriodo, Date), "DD/MM/YYYY"), Format(DateAdd("D", Dias - 1, Date), "DD/MM/YYYY"))
''''            End If
            
            If rcEmpeño!TipoTasa = "MENSUAL" Or rcEmpeño!TipoTasa = "Mensual" Then
               Vencimiento = Format(DateAdd("D", Val(rcEmpeño!VenPeriodo) * 30, Date), "DD/MM/YYYY")
            Else
               Vencimiento = Format(DateAdd("D", Val(rcEmpeño!VenPeriodo), Date), "DD/MM/YYYY")
            End If
            
           
            
            'MLD-MODIF. Grabo el nuevo Empeño
            If rcEmpeño!Serie = 2 Then
                 strSql1 = "INSERT INTO empeno (Fecha,Movimiento,NumContrato,Folio,Prestamo,Avaluo,Origen,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Responsable,Valuador,Notas,Tasa,Almacenaje,Seguro,Operacion,Comision,IVA,Periodo,Venperiodo,VenAlmoneda,Tipointeres,TipoTasa,IDSucursal,IDUsuario,IDAutorizacion,NumBolsa,Ubicacion,Caja,Cajon,Fila,IDUsuarioAutoriza,TipoAutoriza,PrestamoInicial,Beneficiario,IDEmpenoOrigen,IdCotitular,IdTipoOperacion,ClaveTipoOperacion,ValorSalarioMin,ValorUDI,IdInstrumentoMonetario,IdTipoMoneda,IdTipoAlerta,DescTipoAlerta,Promocion,Cat,ImporteSeguroAuto,ImporteRentaGPS,Circulando) VALUES "
            
            Else
                 strSql1 = "INSERT INTO empeno (Fecha,Movimiento,NumContrato,Folio,Prestamo,Avaluo,Origen,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Responsable,Valuador,Notas,Tasa,Almacenaje,Seguro,Operacion,Comision,IVA,Periodo,Venperiodo,VenAlmoneda,Tipointeres,TipoTasa,IDSucursal,IDUsuario,IDAutorizacion,NumBolsa,Ubicacion,Caja,Cajon,Fila,IDUsuarioAutoriza,TipoAutoriza,PrestamoInicial,Beneficiario,IDEmpenoOrigen,IdCotitular,IdTipoOperacion,ClaveTipoOperacion,ValorSalarioMin,ValorUDI,IdInstrumentoMonetario,IdTipoMoneda,IdTipoAlerta,DescTipoAlerta,Promocion,Cat) VALUES "
    
            End If
           
            strSql2 = "('" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & rcEmpeño!NumContrato & "," & FolioNuevo & "," & ConvMoneda(crAdeudo) & "," & ConvMoneda(rcEmpeño!Avaluo) & "," & OD_REFRENDO & ",'" & Format(Vencimiento, "YYYY/MM/DD") & "'," & rcEmpeño!Folio & "," & rcEmpeño!Serie & ",'" & NombrePc & "'," & rcEmpeño!IDCliente & ",'" & rcEmpeño!Responsable & "',"
            strSql3 = "'" & rcEmpeño!Valuador & "','" & rcEmpeño!Notas & "'," & ConvMoneda(rcEmpeño!Tasa) & "," & ConvMoneda(rcEmpeño!Almacenaje) & "," & ConvMoneda(rcEmpeño!Seguro) & "," & ConvMoneda(rcEmpeño!Operacion) & "," & ConvMoneda(rcEmpeño!Comision) & "," & Regresa_Valor_BD("Iva") & "," & rcEmpeño!Periodo & "," & rcEmpeño!VenPeriodo & "," & rcEmpeño!VenAlmoneda & ",'" & rcEmpeño!TipoInteres & "','" & rcEmpeño!TipoTasa & "'," & _
                      rcEmpeño!IDSucursal & "," & rcEmpeño!IDUsuario & "," & Val(lblAutorizacion.Tag) & ",'" & rcEmpeño!NumBolsa & "','" & rcEmpeño!ubicacion & "','" & rcEmpeño!caja & "','" & rcEmpeño!Cajon & "','" & rcEmpeño!Fila & "'," & rcEmpeño!IDUsuarioAutoriza & "," & rcEmpeño!TipoAutoriza & "," & ConvMoneda(rcEmpeño!PrestamoInicial) & ",'" & rcEmpeño!Beneficiario & "'," & rcEmpeño!ID & ""
            If rcEmpeño!Serie = 2 Then
                strSql4 = "," & rcEmpeño!IDCotitular & "," & rcEmpeño!IdTipoOperacion & "," & rcEmpeño!ClaveTipoOperacion & "," & rcEmpeño!ValorSalarioMin & "," & rcEmpeño!ValorUDI & "," & rcEmpeño!IdInstrumentoMonetario & "," & rcEmpeño!IdTipoMoneda & "," & rcEmpeño!IDTipoAlerta & ",'" & rcEmpeño!DescTipoAlerta & "'," & rcEmpeño!Promocion & "," & rcEmpeño!CAT & "," & rcCargoSeguroAuto & "," & rcCargoGPS & "," & EnCirculacion & ")"
            
            Else
              strSql4 = "," & rcEmpeño!IDCotitular & "," & rcEmpeño!IdTipoOperacion & "," & rcEmpeño!ClaveTipoOperacion & "," & rcEmpeño!ValorSalarioMin & "," & rcEmpeño!ValorUDI & "," & rcEmpeño!IdInstrumentoMonetario & "," & rcEmpeño!IdTipoMoneda & "," & rcEmpeño!IDTipoAlerta & ",'" & rcEmpeño!DescTipoAlerta & "'," & rcEmpeño!Promocion & "," & rcEmpeño!CAT & ")"
            
            End If
            
            
            
            dbDatos.Execute strSql1 & strSql2 & strSql3 & strSql4
            
            'Tomo el ID del Empeño
            IDEmpeno = SacaValor("empeno", "MAX(ID)")
            
            dbDatos.Execute "UPDATE empeno SET FolioNota='" & FolioNota & "', IDEmpenoDestino = " & IDEmpeno & " WHERE ID = " & rcEmpeño!ID
            
            'Tomo la Hora
            Hora = Time
            
            If Serie <> SERIE_B Then
                
                If ContratoAlmoneda = 0 Then
                    
                    rcArticulos.Open "SELECT * FROM detallesempeno WHERE IDEmpeno=" & rcEmpeño!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
                Else
                    'MLD-MODIF.
                    rcArticulos.Open "SELECT d.ID AS IDPrendaInventario,d.Codigo,d.Tipo,d.Cantidad,d.Descripcion AS Articulo,d.Peso,d.Kilates,d.Precio AS Avaluo,d.Costo AS Prestamo,d.Estado,d.Serie,d.TipoPrenda,d.Observaciones,d.PesoPiedras,d.Puntos,d.CantidadPiedras,d.CantidadDiamantes,d.PrestamoDiamante,d.Marca,d.Modelo,d.Color,d.Tamano,t.IdTipoGarantia " & _
                                     "FROM detallesentradainventario d LEFT JOIN tipo t ON d.Tipo = t.Id WHERE d.IDEmpeno=" & rcEmpeño!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
'                                      "FROM detallesentradainventario d LEFT JOIN tipo t ON d.Tipo = t.Id WHERE d.Cantidad>0 AND d.IDEmpeno=" & rcEmpeño!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
                End If
                
            Else
                If ContratoAlmoneda = 0 Then
                    
                    rcArticulos.Open "SELECT * FROM detallesempenoautos WHERE IDEmpeno=" & rcEmpeño!ID, dbDatos, adOpenForwardOnly, adLockReadOnly

                Else
                    'MLD-MODIF.
                    rcArticulosAuto.Open "SELECT d.ID AS IDPrendaInventario,d.Codigo,d.Tipo,d.Cantidad,d.Descripcion AS Articulo,d.Peso,d.Kilates,d.Precio AS Avaluo,d.Costo AS Prestamo,d.Estado,d.Serie,d.TipoPrenda,d.Observaciones,d.PesoPiedras,d.Puntos,d.CantidadPiedras,d.CantidadDiamantes,d.PrestamoDiamante,d.Marca,d.Modelo,d.Color,d.Tamano,t.IdTipoGarantia " & _
                                     "FROM detallesentradainventario d LEFT JOIN tipo t ON d.Tipo = t.Id WHERE d.IDEmpeno=" & rcEmpeño!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
'                                      "FROM detallesentradainventario d LEFT JOIN tipo t ON d.Tipo = t.Id WHERE d.Cantidad>0 AND d.IDEmpeno=" & rcEmpeño!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
                    rcArticulos.Open "SELECT * FROM detallesempenoautos WHERE IDEmpeno=" & rcEmpeño!ID, dbDatos, adOpenForwardOnly, adLockReadOnly

                End If
                            
            End If
          
            With rcArticulos
            
                While Not .EOF
                        
                    If Serie <> SERIE_B Then
                        
                        'Verifico si se abono al préstamo
                        crPrestamoPrenda = (!Prestamo * 100) / IIf(Pago > 0, rcEmpeño!Prestamo, crAdeudo)
                        crPrestamoPrenda = Redondeo(crAdeudo * (crPrestamoPrenda / 100))
                                                    
                        'MLD-MODIF.
                        dbDatos.Execute "INSERT INTO detallesempeno (IDEmpeno,Codigo,Tipo,Cantidad,Articulo,Peso,Kilates,Avaluo,Prestamo,Origen,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante,Observaciones,TipoPrenda,Estado,Marca,Modelo,Serie,Color,Tamano,IdTipoGarantia) VALUES (" & _
                                        IDEmpeno & ",'" & !Codigo & "'," & !Tipo & "," & IIf(!Cantidad = 0, 1, !Cantidad) & ",'" & !Articulo & "'," & ConvMoneda(!Peso) & "," & !Kilates & "," & ConvMoneda(!Avaluo) & "," & ConvMoneda(crPrestamoPrenda) & ",1," & _
                                        IIf(IsNull(!CantidadPiedras), 0, !CantidadPiedras) & "," & IIf(IsNull(!PesoPiedras), 0, !PesoPiedras) & "," & !CantidadDiamantes & "," & !Puntos & "," & !PrestamoDiamante & ",'" & IIf(IsNull(!Observaciones), "", !Observaciones) & "'," & !TipoPrenda & ",'" & !Estado & "','" & IIf(IsNull(!Marca), "", !Marca) & "','" & IIf(IsNull(!Modelo), "", !Modelo) & "','" & IIf(IsNull(!Serie), "", !Serie) & "','" & IIf(IsNull(!Color), "", !Color) & "','" & IIf(IsNull(!Tamano), "", !Tamano) & "'," & !IdTipoGarantia & ")"
                        
                        If ContratoAlmoneda = 1 Then
                            
                            'Lo doy de baja del Inventario
                            dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=0,TipoSalida=" & SALIDAREEMPENO & " WHERE ID=" & rcArticulos!IDPrendaInventario
                            
                            'Tomo el Importe para darlo de baja del Inventario
                            crImporteAlmoneda = crImporteAlmoneda + crPrestamoPrenda
                        End If
                        
                    Else
                    
                       'Verifico si se abono al préstamo
                       ' crPrestamoPrenda = (!Prestamo * 100) / IIf(Pago > 0, rcEmpeño!Prestamo, crAdeudo)
                        'crPrestamoPrenda = Redondeo(crAdeudo * (crPrestamoPrenda / 100))
                        
                        'MLD-MODIF.
                        dbDatos.Execute "INSERT INTO detallesempenoautos(IDEmpeno,marcaymodelo,año,color,placas,factura,agencia,numtarjetacircu,nummotor,seriechasis,kms,gas,aseguradora,poliza,fechavenci,tipo,factu,tarjetacircu,copiaife,tenencias,polizaseguro,copialicencia,importacion,IdTipoGarantia,IdTipoBlindajeAutos,Marca,Modelo,VIN,RePuVe)values" _
                                        & "(" & IDEmpeno & ",'" & !MarcayModelo & "'," & !Año & ",'" & !Color & "','" & !Placas & "','" & !Factura & "','" & !Agencia & "','" & !NumTarjetaCircu & "','" & !NumMotor & "','" & !SerieChasis & "','" & !Kms & "','" & !Gas & "','" & !Aseguradora & "','" & !Poliza & "'," & IIf(IsNull(!FechaVenci), "Null", "'" & Format(!FechaVenci, "YYYY/MM/DD") & "'") & ",'" & !Tipo & "'," & !Factu & "," & !TarjetaCircu & "," & !CopiaIfe & "," & !Tenencias & "," & !PolizaSeguro & "," & !CopiaLicencia & "," & !Importacion & "," & !IdTipoGarantia & "," & !IdTipoBlindajeAutos & ",'" & Trim(!Marca) & "','" & Trim(!Modelo) & "','" & Trim(!VIN) & "','" & Trim(!RePuVe) & "')"
                    
                       If ContratoAlmoneda = 1 Then
                            
                            'Lo doy de baja del Inventario
                            dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=0,TipoSalida=" & SALIDAREEMPENO & " WHERE ID=" & rcArticulosAuto!IDPrendaInventario
                            
                            'Tomo el Importe para darlo de baja del Inventario
                            'crImporteAlmoneda = crImporteAlmoneda + crPrestamoPrenda
                        End If
                    End If
                    
                .MoveNext
                Wend
                
            End With
            rcArticulos.Close
            
            'Grabamos si se hizo un abono
            If Pago > 0 Then
            
                'Grabamos el cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Abono Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(Pago) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Grabamos el abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Abono Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','201750'," & ConvMoneda(Pago) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            End If
            
            
            'Cuenta de Intereses
            'Cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(crIntereses) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            'Abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','520450'," & ConvMoneda(crIntereses) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                        
            
            'Cuenta de Almacenaje
            If crAlmacenaje > 0 Then
                
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(crAlmacenaje) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','670350'," & ConvMoneda(crAlmacenaje) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            End If
            
            
            
            'Cuenta de Seguro
            If crSeguro > 0 Then
                
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(crSeguro) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','680350'," & ConvMoneda(crSeguro) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            End If
            
            
            'Cuenta de Moratorios
            If crMoratorios > 0 Then
                
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(crMoratorios) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','690350'," & ConvMoneda(crMoratorios) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            End If
            
            If EnCirculacion = 1 Then
                'Cuenta de Renta GPS
                If rcCargoGPS > 0 Then
                    'Se agrega el abono a Renta deGPS
                      'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo Renta GPS'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(rcCargoGPS) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo Renta GPS'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','818150'," & ConvMoneda(rcCargoGPS) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

                End If
                'Cuenta de Seguro Auto
                If rcCargoSeguroAuto > 0 Then
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo Renta GPS'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(rcCargoSeguroAuto) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo Renta GPS'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','828250'," & ConvMoneda(rcCargoSeguroAuto) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

                End If
                If IvaCargoRenta > 0 Then
                   crIva = crIva - IvaCargoRenta
                   'Cargo
                   dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo Renta GPS'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(IvaCargoRenta) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                   
                   'Abono
                   dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo Renta GPS'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','120150'," & ConvMoneda(IvaCargoRenta) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
     
                End If
            End If
            
          
            
            'Cuenta de Iva
            If crIva > 0 Then
            
                'Cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(crIva) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                
                'Abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','120150'," & ConvMoneda(crIva) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            End If
            
         
            'Grabamos si pago por boleta perdida
            If ImportePerdida > 0 Then
                
                'Grabamos el cargo de la boleta perdida
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Boleta perdida'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(ImportePerdida) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Grabamos el abono de la boleta perdida
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Boleta perdida'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','530150'," & ConvMoneda(ImportePerdida) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                
                If ImporteIvaPerdida > 0 Then
                
                    'Grabamos el cargo del iva de la boleta perdida
                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Boleta perdida'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(ImporteIvaPerdida) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                
                    'Grabamos el abono del iva de la boleta perdida
                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Boleta perdida'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','120150'," & ConvMoneda(ImporteIvaPerdida) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                End If
                
            End If
            
            
'''            'Grabamos el cargo a 199401
'''            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
'''                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','199401'," & ConvMoneda(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva + ImportePerdida + ImporteIvaPerdida + Pago) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            'Grabamos el cargo a 110101
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','110101'," & ConvMoneda(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva + ImportePerdida + ImporteIvaPerdida + Pago) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                        
            'Ponemos la entrada del nuevo Folio
            'Grabamos el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioNuevo & ",'" & rcEmpeño!Iniciales & "','201701'," & ConvMoneda(crAdeudo) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            
            'Verifico si es un Reempeño
            If ContratoAlmoneda = 0 Then
                
                'Grabamos el abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','201750'," & ConvMoneda(crAdeudo) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            Else
                
                'Grabamos el cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & rcEmpeño!Iniciales & "','620350'," & ConvMoneda(crImporteAlmoneda) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            End If
                 
            '***Puntos***
            'descontamos los puntos
            If crPuntos > 0 Then
                            
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Redencion Puntos Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & _
                                rcEmpeño!Iniciales & "','905501'," & ConvMoneda(crPuntos) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
'''                'Grabamos el cargo a 199450
'''                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
'''                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Redencion Puntos Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & _
'''                                rcEmpeño!Iniciales & "','199450'," & ConvMoneda(crPuntos) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                
                'Grabamos el cargo a 199450
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Redencion Puntos Refrendo'," & Movimiento & "," & FolioAnterior & ",'" & _
                                rcEmpeño!Iniciales & "','110150'," & ConvMoneda(crPuntos) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
            End If
                 
''''            'Imprimo el nuevo contrato
''''            If Serie = SERIE_A Then
''''
                ''Imprimir_Boleta_CR IDEmpeno
''''            Else
''''
''''                Imprimir_Boleta_CR_Auto IDEmpeno
''''            End If
            
            rcEmpeño.Close
            rcEmpeño.Open "SELECT empeno.Prestamo,empeno.Avaluo,empeno.Fecha,empeno.TipoInteres,empeno.Serie FROM empeno WHERE ID=" & IDEmpeno, dbDatos, adOpenForwardOnly, adLockOptimistic
                                                
                'Opciones de Pago
                If Serie <> SERIE_B Then
                    OpcionesPago rcEmpeño!Prestamo, rcEmpeño!Avaluo, rcEmpeño!Fecha, IDEmpeno, rcEmpeño!TipoInteres, IIf(rcEmpeño!Serie = SERIE_B, True, False)

                Else
                    OpcionesPago rcEmpeño!Prestamo, rcEmpeño!Avaluo, rcEmpeño!Fecha, IDEmpeno, rcEmpeño!TipoInteres, IIf(rcEmpeño!Serie = SERIE_B, True, False)

                End If
                                
            rcEmpeño.Close
            If Serie = 2 Then
                  'Imprimo el Recibo
                Imprimir_Nota_Auto IDEmpenoAnterior, OD_REFRENDO, Pago, frmMDI.IDUsuario, DateAdd("D", Regresa_Valor_BD("DiasEnajenacion"), Vencimiento)
            
            Else
                  'Imprimo el Recibo
                Imprimir_Nota IDEmpenoAnterior, OD_REFRENDO, Pago, frmMDI.IDUsuario, DateAdd("D", Regresa_Valor_BD("DiasEnajenacion"), Vencimiento)
            
            End If
          
             If EnCirculacion = 1 Then
                'Imprimo Recibo de GPS y Seguro Auto
                Imprimir_Nota_GPS_Seguro IDEmpenoAnterior, rcCargoGPS, rcCargoSeguroAuto, IvaCargoRenta, OD_REFRENDO, Pago, frmMDI.IDUsuario, DateAdd("D", Regresa_Valor_BD("DiasEnajenacion") + 1, Vencimiento)
             End If
                
    Next Indice
        
    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\APAGADOS.bmp")
    ImgSemaforo.ToolTipText = ""
    Limpiar "REFRENDOS"
    grdRefrendos.Clear
    
Error:
    Maneja_Error Err
    Set rcEmpeño = Nothing
    Set rcArticulos = Nothing
    Screen.MousePointer = vbDefault
End Sub

'Validamos el refrendo
Private Function Validar_Refrendo() As Boolean
    
    Validar_Refrendo = True
  
    If grdRefrendos.Rows = 0 Then
        MsgBox "Introduzca los contratos que desea refrendar !!", vbInformation, "Refrendo"
        Validar_Refrendo = False
        Exit Function
    End If

End Function

'Validamos que esten los datos correctos en Empeno
Private Function Validar_Empeno() As Boolean
Dim i As Integer, x As Integer

    Validar_Empeno = True
      
    If Trim(txtNombre.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtNombre.SetFocus
        Exit Function
    End If
  
    'si no tiene apellido
    If Trim(txtApellidoPaterno.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtApellidoPaterno.SetFocus
        Exit Function
    End If
    
    'si no tiene apellido
    If Trim(txtApellidoMaterno.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtApellidoMaterno.SetFocus
        Exit Function
    End If
    
    If Not ClienteEmp.Valida Then
        MsgBox "Datos requeridos del Cliente incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        cmdEditarCliente_Click
        Exit Function
    End If
    
    If Trim(txtResponsable.text) <> "" Or Trim(txtCotitularApellidoPaterno.text) <> "" Or Trim(txtCotitularApellidoMaterno.text) <> "" Then
        If Not CotitularEmp.Valida Then
            MsgBox "Datos requeridos del Cotitular incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Validar_Empeno = False
            cmdEditarCotitular_Click
            Exit Function
        End If
    End If
        
'    'si no tiene nombre
'    If Trim(txtNombre.text) = "" Then
'        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'        Validar_Empeno = False
'        txtNombre.SetFocus
'        Exit Function
'    End If
'
'    'si no tiene apellido
'    If Trim(txtApellidos.text) = "" Then
'        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'        Validar_Empeno = False
'        txtApellidos.SetFocus
'        Exit Function
'    End If
'
'    '''''    'si no tiene direccion
'    '''''    If Trim(txtDireccion.Text) = "" Then
'    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'    '''''        Validar_Empeno = False
'    '''''        txtDireccion.SetFocus
'    '''''        Exit Function
'    '''''    End If
'    '''''
'    '''''    'si no tiene estado
'    '''''    If Trim(txtEstado.Text) = "" Then
'    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'    '''''        Validar_Empeno = False
'    '''''        txtEstado.SetFocus
'    '''''        Exit Function
'    '''''    End If
'    '''''
'    '''''    'si no tiene colonia
'    '''''    If Trim(txtColonia.Text) = "" Then
'    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'    '''''        Validar_Empeno = False
'    '''''        txtColonia.SetFocus
'    '''''        Exit Function
'    '''''    End If
'    '''''
'    '''''    'si no tiene municipio
'    '''''    If Trim(txtMunicipio.Text) = "" Then
'    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'    '''''        Validar_Empeno = False
'    '''''        txtMunicipio.SetFocus
'    '''''        Exit Function
'    '''''    End If
'    '''''
'    '''''    'si no tiene cp
'    '''''    If Trim(txtCp.Text) = "" Then
'    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'    '''''        Validar_Empeno = False
'    '''''        txtCp.SetFocus
'    '''''        Exit Function
'    '''''    End If
'
'    'Si no identificacion
'    If Trim(txtIdentificacion.text) = "" Then
'        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
'        Validar_Empeno = False
'        txtIdentificacion.SetFocus
'        Exit Function
'    End If
        
    If Trim(cmbTipoInteres.text) = "" Then
        MsgBox "Seleccione el tipo de tasa !!", vbCritical, "Empeño"
        Validar_Empeno = False
        cmbTipoInteres.SetFocus
        Exit Function
    End If
         
    'si el prestamo es 0 o menor a 0
    If Val(txtPrestamo.Caption) < 0 Then
        MsgBox "El Préstamo no puede ser igual o menor a 0, favor de llenar correctamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        'txtPrestamo.SetFocus
        Exit Function
    End If
        
'''    If Trim(txtNumBolsa.text) = "" And BanElec = False Then
'''        MsgBox "Introduzca el número de bolsa !!", vbCritical + vbOKOnly, "Empeño"
'''        Validar_Empeno = False
'''        txtNumBolsa.SetFocus
'''        Exit Function
'''    End If
    
    If cmbMedio.ListIndex = -1 Then
        MsgBox "Seleccione el medio por el cual se enteró !!", vbCritical + vbOKOnly, "Empeño"
        Validar_Empeno = False
        cmbMedio.SetFocus
        Exit Function
    End If
    
    'Checo si el grid de empeños tiene información
    x = 0

    For i = 1 To grdEmpeños.Rows

        If Val(grdEmpeños.CellText(i, 2)) = 0 Then x = x + 1
    Next i
    
    If x = grdEmpeños.Rows Then
        MsgBox "Favor de agregar las prendas al contrato !!", vbInformation, "Empeño"
        Validar_Empeno = False
    End If
    
End Function

Private Function Validar_Desempeño() As Boolean
    Validar_Desempeño = True
  
    If grdDesempeño.Rows = 0 Then
        
        MsgBox "Introduzca los contratos que desea desempeñar !!", vbInformation, "Desempeño"
        Validar_Desempeño = False
    End If
End Function

Private Sub cmdAceptar_GotFocus()
    cmdAceptar.BackColor = vb3DShadow
End Sub

Private Sub cmdAceptar_LostFocus()
    cmdAceptar.BackColor = vbButtonFace
End Sub

Private Sub cmdCancelacion_Click()
    frmPasswords.ConexSuc = 0
    frmPasswords.DescuentoVentas = 0
    frmPasswords.PrecioVitrina = 0
    frmPasswords.Ventas = 0
    frmPasswords.ModificaPrecio = 0
    frmPasswords.ModificaCorte = 0
    frmPasswords.InteresDesempeño = 0
    frmPasswords.InteresRefrendo = 0
    frmPasswords.HacerCorte = 0
    frmPasswords.RecalculoPrecios = 0
    frmPasswords.AutorizaPrestamo = 0
    frmPasswords.Vencido = 0
    frmPasswords.CancelaCierre = 0
    frmPasswords.Cancel = 1

    If frmPasswords.Password(CANCELACION, 1) Then
    
        frmCancelaciones.Show
        BringWindowToTop frmCancelaciones.hWnd
    End If

End Sub



Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdSalir_GotFocus()
    cmdSalir.BackColor = vb3DShadow
End Sub

Private Sub cmdSalir_LostFocus()
    cmdSalir.BackColor = vbButtonFace
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    
    Screen.MousePointer = vbHourglass
    
    '--------- MLD-MODIF. ---------------
    Set ClienteEmp = New clientes
    Set CotitularEmp = New clientes
    ClienteEmp.FechaExpiracion = "1900-01-01"
    ClienteEmp.FechaNacimiento = "1900-01-01"
    ClienteEmp.FechaAltaRazonSocial = "1900-01-01"
    CotitularEmp.FechaExpiracion = "1900-01-01"
    CotitularEmp.FechaNacimiento = "1900-01-01"
    CotitularEmp.FechaAltaRazonSocial = "1900-01-01"
    cmbTipoBlindaje.AddItem "NINGUNO": cmbTipoBlindaje.ItemData(cmbTipoBlindaje.NewIndex) = 0: cmbTipoBlindaje.ListIndex = 0
    Cargar_Combos "Descripcion", "mld_vehiculos_tipo_blindaje", cmbTipoBlindaje, , "Descripcion", False
    InicializarAlerta vTipoAlerta, MLD_PRESTAMO
    cmdAlerta.Enabled = False
    cmdAlerta2.Enabled = False
    '---------------------------------------
    
    Set TarjetaPuntos.CONEXION = dbDatos
    
    frmEmpeño.BorderStyle = 0
    frmRefrendos.BorderStyle = 0
    frmDesempeño.BorderStyle = 0
    frmAutomoviles.BorderStyle = 0
    frmMetales.BorderStyle = 0
    frmElectronicos.BorderStyle = 0
    
    Limpiar "Empeno"
    Limpiar "DESEMPEÑO"
    Limpiar "REFRENDOS"
    
    Crear_Pestañas
    Crear_Encabezados
    
    Bandera = True
    BanElec = False
    
    Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Serie=" & SERIE_A, "Ordenamiento"
    Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres2, " WHERE Serie=" & SERIE_B, "Ordenamiento"
    Cargar_Combos "Descripcion", "tipo", cmbTipo, " WHERE (Kilataje=1 OR Peso=1)", "Ordenamiento"
    Cargar_Combos "Descripcion", "tipo", cmbTipoElec, " WHERE (Kilataje=0 AND Peso=0)", "Ordenamiento"
    cmbPromocion.Clear
    cmbPromocion.AddItem "Sin Promoción"
    cmbPromocion.ItemData(cmbPromocion.NewIndex) = 0
    Cargar_Combos "Descripcion", "Promociones", cmbPromocion, " WHERE Activa=1", , False
    
    cmbMedio.AddItem "[AGREGAR]"
    cmbMedio2.AddItem "[AGREGAR]"
    Cargar_Combos "Descripcion", "medios", cmbMedio, , , False
    Cargar_Combos "Descripcion", "medios", cmbMedio2, , , False
    
    lblIva.Caption = Format(Regresa_Valor_BD("IVA"), "0.00") & "%"
    
    txtPrestamo.Caption = "0.00"
    lblTotAvaluo.Caption = "0.00"
    
    cmbTipoInteres.ListIndex = 0
    cmbTipo.ListIndex = ComboInformacion(cmbTipo, 1)
    cmbPromocion.ListIndex = 0
    
    txtNotas.text = Regresa_Valor_BD("Notas")
    
    Default 1
    
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
    
    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\APAGADOS.bmp")
    ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
    lblInfoSemaforo.Caption = ""
    Frame3.Visible = False
    Frame4.Visible = False
    Titular = True
End Sub

'Creamos las pestañas del tab
Private Sub Crear_Pestañas()

    With TPestañas
        .AddTab "Empeño", , , "K1"
        .AddTab "Autos", , , "K4"
        .AddTab "Desempeño", , , "K2"
        .AddTab "Refrendo", , , "K3"
    End With
    
    With TPrendas
        .AddTab "Metales", , , "K1"
        .AddTab "Electrónicos/Varios", , , "K4"
    End With
End Sub

'Creamos los encabezados del grid
Private Sub Crear_Encabezados()

    With grdDesempeño
        .AddColumn "K1", "Contrato", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortString
        .AddColumn "K2", "Préstamo", ecgHdrTextALignRight, , 165, , , , , FMoneda, , CCLSortString
        .AddColumn "K3", "Interés", ecgHdrTextALignRight, , 155, , , , , FMoneda, , CCLSortString
        .AddColumn "K4", " ", ecgHdrTextALignLeft, , 1, , , , , , , CCLSortString
      
        .AddColumn "K5", "Interés", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K6", "Almacenaje", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K7", "Seguro", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K8", "Iva", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K9", "Almoneda", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K10", "Importe Perdida", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K11", "Moratorios", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        
        .AddColumn "K12", "Control", ecgHdrTextALignRight, , 90, False, , , , , , CCLSortNumeric
        
        '***Puntos***
        .AddColumn "K13", "Puntos", ecgHdrTextALignRight, , 90, , , , , , , CCLSortNumeric
        .AddColumn "K14", "Total", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K15", "- Puntos", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K16", "Total a Pagar", ecgHdrTextALignRight, , 150, , , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "K17", "IDCliente", ecgHdrTextALignRight, , , False, , , , , , CCLSortNumeric
        .AddColumn "K18", "IDEmpeno", ecgHdrTextALignRight, , , False, , , , , , CCLSortNumeric
        .AddColumn "K19", "IvaCargo", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
    End With
    
    With grdRefrendos
        .AddColumn "K1", "Contrato", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K2", "Préstamo", ecgHdrTextALignRight, , 145, , , , , FMoneda, , CCLSortString
        .AddColumn "K3", "Abono", ecgHdrTextALignRight, , 113, , , , , FMoneda, , CCLSortString
        .AddColumn "K4", "Interés", ecgHdrTextALignRight, , 127, , , , , FMoneda, , CCLSortString
      
        .AddColumn "K5", "Interés", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K6", "Almacenaje", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K7", "Seguro", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K8", "Iva", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K9", "Almoneda", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K10", "Importe Perdida", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K11", "Moratorios", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        
        .AddColumn "K12", "Control", ecgHdrTextALignRight, , 90, False, , , , , , CCLSortNumeric
        
        '***Puntos***
        .AddColumn "K13", "Puntos", ecgHdrTextALignRight, , 90, , , , , , , CCLSortNumeric
        .AddColumn "K14", "Total", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K15", "- Puntos", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K16", "Total a Pagar", ecgHdrTextALignRight, , 150, , , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "K17", "IDCliente", ecgHdrTextALignRight, , , False, , , , , , CCLSortNumeric
        
        .AddColumn "K18", "Vencido", ecgHdrTextALignRight, , , False, , , , , , CCLSortNumeric
        .AddColumn "K19", "IDEmpeno", ecgHdrTextALignRight, , , False, , , , , , CCLSortNumeric
        .AddColumn "K20", "IvaCargo", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString

    End With
   
    With grdEmpeños
        .AddColumn "K1", "Tipo", ecgHdrTextALignLeft, , 90, False, , , , , , CCLSortString
        .AddColumn "K2", "Cant.", ecgHdrTextALignRight, , 39, , , , , , , CCLSortNumeric
        .AddColumn "K3", "Prenda", ecgHdrTextALignLeft, , 125, , , , , , , CCLSortNumeric
        .AddColumn "K4", "Peso", ecgHdrTextALignRight, , 48, , , , , , , CCLSortNumeric
        .AddColumn "K5", "Kílates", ecgHdrTextALignRight, , 58, , , , , , , CCLSortString
        .AddColumn "K6", "Avalúo", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Préstamo", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "Modelo", ecgHdrTextALignLeft, , 63, False, , , , , , CCLSortString
        .AddColumn "K9", "Hechura", ecgHdrTextALignLeft, , 63, False, , , , , , CCLSortString
        .AddColumn "K10", "Precio V.", ecgHdrTextALignRight, , 80, False, , , , , , CCLSortString
        .AddColumn "K11", "Observaciones", ecgHdrTextALignLeft, , 125, , , , , , , CCLSortNumeric
        
        .AddColumn "K12", "Cantidad Piedras", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortNumeric
        .AddColumn "K13", "Peso Piedras", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortNumeric
        
        .AddColumn "K14", "Cantidad Diamantes", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortNumeric
        .AddColumn "K15", "Puntos", ecgHdrTextALignLeft, , 115, False, , , , , , CCLSortNumeric
        .AddColumn "K16", "Prestamo Diamantes", ecgHdrTextALignLeft, , 115, False, , , , , , CCLSortNumeric
        
        .AddColumn "K17", "Marca", ecgHdrTextALignLeft, , 115, False, , , , , , CCLSortString
        .AddColumn "K18", "Modelo", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortString
        .AddColumn "K19", "NoSerie", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortString
        .AddColumn "K20", "Color", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortString
        .AddColumn "K21", "Tamaño", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortString
        .Rows = 11
    End With

End Sub

Private Sub Form_LostFocus()
    frmMDI.Com.PortOpen = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Quitar_Flat Fl
    '''''DetachMessage Me, Me.hwnd, WM_CONFIGURACION
    If frmMDI.Com.PortOpen Then
        
        frmMDI.Com.PortOpen = False
    End If

End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
    ISubclass_MsgResponse = emrPreprocess
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
'
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  Select Case iMsg
    Case WM_CONFIGURACION:
      'Cargar_Vencimiento
  End Select
End Function

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEdad_GotFocus()
    Cambiar_Color True, txtEdad
    Seleccionar_Texto txtEdad
End Sub

Private Sub txtEdad_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEdad_LostFocus()
    Cambiar_Color False, txtEdad
End Sub


Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtedit2_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtFamiliaElec_GotFocus()
    Seleccionar_Texto txtFamiliaElec
    Cambiar_Color True, txtFamiliaElec
End Sub

Private Sub txtFamiliaElec_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFamiliaElec_LostFocus()
    Cambiar_Color False, txtFamiliaElec
End Sub

Private Sub txtFechaVenciPoliza_GotFocus()
    Seleccionar_Texto txtFechaVenciPoliza
    Cambiar_Color True, txtFechaVenciPoliza
End Sub

Private Sub txtFechaVenciPoliza_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaVenciPoliza_LostFocus()
    Cambiar_Color False, txtFechaVenciPoliza
End Sub




Private Sub txtMarcaElec_GotFocus()
    Seleccionar_Texto txtMarcaElec
    Cambiar_Color True, txtMarcaElec
End Sub

Private Sub txtMarcaElec_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMarcaElec_LostFocus()
    Cambiar_Color False, txtMarcaElec
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

Private Sub txtModeloElec_GotFocus()
    Seleccionar_Texto txtModeloElec
    Cambiar_Color True, txtModeloElec
End Sub

Private Sub txtModeloElec_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtModeloElec_LostFocus()
    Cambiar_Color False, txtModeloElec
End Sub



'***Puntos***
Private Sub txtNoTarjeta_GotFocus()
    Titular = True
    Seleccionar_Texto txtNoTarjeta
    'Cambiar_Color True, txtNoTarjeta
End Sub
'***Puntos***
Private Sub txtNoTarjeta_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
   'Pasar_Foco KeyAscii
   If KeyAscii = vbKeyReturn Then
      If TarjetaPuntos.CuentaFrecuente.FindCuentaByFolio(txtNoTarjeta.text) Then
         lblPuntosAcumulados.Caption = TarjetaPuntos.CuentaFrecuente.Puntos
         'Buscar_Cliente TarjetaPuntos.CuentaFrecuente.IDCliente, True
         Buscar TarjetaPuntos.CuentaFrecuente.IDCliente, True
         
      Else
         lblPuntosAcumulados.Caption = "0"
         Seleccionar_Texto txtNoTarjeta
         MsgBox "No se encuentra la tarjeta de cliente frecuente", vbOKOnly Or vbInformation
      End If
   End If
End Sub
'***Puntos***
Private Sub txtNoTarjeta_LostFocus()
    'Cambiar_Color False, txtNoTarjeta
End Sub

Private Sub txtNotas_GotFocus()
    Cambiar_Color True, txtNotas
    Seleccionar_Texto txtNotas
End Sub

Private Sub txtNotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNotas_LostFocus()
    Cambiar_Color False, txtNotas
End Sub

Private Sub tPestañas_TabClick(ByVal lTab As Long)

    '***Puntos***
    lblNoTarjeta.Visible = False
    txtNoTarjeta.Visible = False
    lblPuntosAcumulados1.Visible = False
    lblPuntosAcumulados.Visible = False
    
    Timer1.Enabled = False
    Timer2.Enabled = False
    Select Case lTab

        Case 1
                
            '***Puntos***
            lblNoTarjeta.Visible = True
            txtNoTarjeta.Visible = True
            lblPuntosAcumulados1.Visible = True
            lblPuntosAcumulados.Visible = True
            
            Limpiar "Empeno"
            LimpiaArticulos
            grdEmpeños.ClearItems
            grdEmpeños.Rows = 11
            Default 1
            txtNotas.text = Regresa_Valor_BD("Notas")
            lblIva.Caption = Format(Regresa_Valor_BD("IVA"), "0.00") & "%"
            cmbPromocion.ListIndex = 0
            cmbTipoInteres.ListIndex = 0
            cmbTipoInteres_Click
            lblTotAvaluo.Caption = "0.00"
            frmRefrendos.Visible = False
            frmDesempeño.Visible = False
            frmAutomoviles.Visible = False
            frmEmpeño.Visible = True
            ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\APAGADOS.bmp")
            ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
            ImgSemaforo.ToolTipText = ""
            
        Case 2
                                    
            '***Puntos***
            lblNoTarjeta.Visible = True
            txtNoTarjeta.Visible = True
            lblPuntosAcumulados1.Visible = True
            lblPuntosAcumulados.Visible = True
                                    
            Limpiar "Autos"
            Limpiar "DOCUMENTOS ENTREGADOS"
            Limpiar "DATOS DEL AUTOMÓVIL"
            Default 2
            txtNotas2.text = Regresa_Valor_BD("Notas")
            lblIva2.Caption = Regresa_Valor_BD("IVA") & "%"
            lblFecha(4).Caption = Format(Date, "DD/MMM/YY")
            cmbTipoInteres2.ListIndex = 0
            cmbTipoInteres2_Click
            cmbPromocion2.ListIndex = 0
            lblTotAvaluo2.Caption = "0.00"
            frmEmpeño.Visible = False
            frmRefrendos.Visible = False
            frmDesempeño.Visible = False
            frmAutomoviles.Visible = True
            ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\APAGADOS.bmp")
            ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
            ImgSemaforo.ToolTipText = ""
            
        Case 3
            
            
            chkAutomovil(0).Value = 0
            Limpiar "DESEMPEÑO"
            Limpiar_Leyendas
            labelContratoDesemp.Visible = False
            NuevoFolio.Caption = ""
            frmEmpeño.Visible = False
            frmRefrendos.Visible = False
            frmAutomoviles.Visible = False
            frmDesempeño.Visible = True
            txtFolioDesempeño.SetFocus
            ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\APAGADOS.bmp")
            ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
            ImgSemaforo.ToolTipText = ""
            
        Case 4
                        
            chkAutomovil(1).Value = 0
            Limpiar "REFRENDOS"
            Limpiar_Leyendas
            labelContratoAlmoneda.Visible = False
            frmEmpeño.Visible = False
            frmDesempeño.Visible = False
            frmAutomoviles.Visible = False
            frmRefrendos.Visible = True
            txtFolioRefrendo.SetFocus
            ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\APAGADOS.bmp")
            ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
            ImgSemaforo.ToolTipText = ""
            
    End Select

    '***Puntos***
    Limpiar_Tarjeta
    
    'MLD-MODIF.--------------------------------
    InicializarAlerta vTipoAlerta, MLD_PRESTAMO
    cmdAlerta.Enabled = False
    cmdAlerta2.Enabled = False
    ClienteEmp.Limpiar
    CotitularEmp.Limpiar
    '------------------------------------------
    Titular = True
End Sub




Private Sub txtFolioDesempeño_Change()
    txtFolioDesempeño.Tag = ""
End Sub

Private Sub txtFolioRefrendo_Change()
    txtFolioRefrendo.Tag = ""
End Sub



Private Sub txtNumBolsa_GotFocus()
    Seleccionar_Texto txtNumBolsa
    Cambiar_Color True, txtNumBolsa
End Sub

Private Sub txtNumBolsa_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumBolsa_LostFocus()
    Cambiar_Color False, txtNumBolsa
End Sub

Private Sub txtNumSerieElec_GotFocus()
    Seleccionar_Texto txtNumSerieElec
    Cambiar_Color True, txtNumSerieElec
End Sub

Private Sub txtNumSerieElec_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumSerieElec_LostFocus()
    Cambiar_Color False, txtNumSerieElec
End Sub

Private Sub txtObservaciones_GotFocus()
    Seleccionar_Texto txtObservaciones
    Cambiar_Color True, txtObservaciones
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtObservaciones_LostFocus()
    Cambiar_Color False, txtObservaciones
End Sub

Private Sub txtObservaciones2_GotFocus()
    Seleccionar_Texto txtObservaciones2
    Cambiar_Color True, txtObservaciones2
End Sub

Private Sub txtObservaciones2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtObservaciones2_LostFocus()
    Cambiar_Color False, txtObservaciones2
End Sub

Private Sub txtObservacionesElec_GotFocus()
    Seleccionar_Texto txtObservacionesElec
    Cambiar_Color True, txtObservacionesElec
End Sub

Private Sub txtObservacionesElec_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        cmdAgregar.SetFocus
    End If
End Sub

Private Sub txtObservacionesElec_LostFocus()
    Cambiar_Color False, txtObservacionesElec
End Sub

Private Sub txtPeso_Change()
    Calcular_Avaluo
End Sub

Private Sub txtPeso_GotFocus()
    Seleccionar_Texto txtPeso
    Cambiar_Color True, txtPeso
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtPeso.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPeso_LostFocus()
    Cambiar_Color False, txtPeso
End Sub

Private Sub txtPesoPiedra_Change()
    Calcular_Avaluo
End Sub

Private Sub txtPesoPiedra_GotFocus()
    Seleccionar_Texto txtPesoPiedra
    Cambiar_Color True, txtPesoPiedra
End Sub

Private Sub txtPesoPiedra_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtPesoPiedra.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPesoPiedra_LostFocus()
    Cambiar_Color False, txtPesoPiedra
End Sub

Private Sub txtPiedras_GotFocus()
    Seleccionar_Texto txtPiedras
    Cambiar_Color True, txtPiedras
End Sub

Private Sub txtPiedras_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPiedras_LostFocus()
    Cambiar_Color False, txtPiedras
End Sub

Private Sub txtPrestamo2_Change()
Dim crPrestamo As Double
    
   If Val(txtPrestamo2.text) = 0 Or Trim(txtPrestamo2.text) = "" Then
        
        crPrestamo = 0
    Else
        
        crPrestamo = CDbl(txtPrestamo2.text)
    End If
    
    lblTotAvaluo2.Caption = Format(Calcula_Prestamo(crPrestamo, Regresa_Valor_BD("PrestamoAvaluoAutos"), False), FMoneda)
    SacaTasa crPrestamo, cmbTipoInteres2.ItemData(cmbTipoInteres2.ListIndex), cmbPeriodo2.ItemData(cmbPeriodo2.ListIndex), cmbPlazos2.ItemData(cmbPlazos2.ListIndex), IIf(Val(txtNombre2.Tag) = 0, False, True)
End Sub

Private Sub txtFolioDesempeño_GotFocus()

    Seleccionar_Texto txtFolioDesempeño
    Cambiar_Color True, txtFolioDesempeño
    
    If Leyenda.Tag = "1" Then
        
        Leyenda.Tag = ""
        Leyenda.Caption = ""
        TotalDesempeño.Caption = ""
        TotalDesempeño.ForeColor = &HFF0000
    End If

End Sub

Private Sub txtFolioDesempeño_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn And Trim(txtFolioDesempeño.text) <> "" Then
            
        If VerificaContratoDuplicado(txtFolioDesempeño.text, grdDesempeño, 1) = False Then
            
            Buscar_Empeno txtFolioDesempeño, 1
        Else
            
            txtFolioDesempeño.text = ""
        End If
        
    End If
End Sub

Private Sub txtFolioDesempeño_LostFocus()
    Cambiar_Color False, txtFolioDesempeño
End Sub

Private Sub txtFolioRefrendo_GotFocus()
    
    Seleccionar_Texto txtFolioRefrendo
    Cambiar_Color True, txtFolioRefrendo
    
    If LeyendaRef.Tag = "1" Then
        
        LeyendaRef.Tag = ""
        LeyendaRef.Caption = ""
        TotalRefrendo.Caption = ""
        TotalRefrendo.ForeColor = &HFF0000
        NuevoFolio.Caption = ""
    End If

End Sub

Private Sub txtFolioRefrendo_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn And Trim(txtFolioRefrendo.text) <> "" Then
        
        If VerificaContratoDuplicado(txtFolioRefrendo.text, grdRefrendos, 2) = False Then
            
            Buscar_Empeno txtFolioRefrendo, 2
        Else
            
            txtFolioRefrendo.text = ""
        End If
        
    End If
End Sub

Private Sub txtFolioRefrendo_LostFocus()
    Cambiar_Color False, txtFolioRefrendo
End Sub



Private Sub cmbTipo_GotFocus()
    Seleccionar_Texto cmbTipo
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
    grdEmpeños.CancelEdit
End Sub

Private Sub Limpiar(Contededor As String)
Dim ctrl As Control
  
    For Each ctrl In Controls
        
        On Error Resume Next

        If ctrl.Container.Caption = Contededor Then
            
            If TypeOf ctrl Is TextBox And ctrl.Name <> "NotaRef" And ctrl.Name <> "FechaCap" And ctrl.Name <> "VencimientoCap" Then ctrl.text = ""
            If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
            If TypeOf ctrl Is ComboBox And ctrl.Name <> "cmbTipoInteres" And ctrl.Name <> "cmbTipoInteres2" And ctrl.Name <> "cmbPeriodo" And ctrl.Name <> "cmbPeriodo2" And ctrl.Name <> "cmbPlazos" And ctrl.Name <> "cmbPlazos2" Then ctrl.ListIndex = -1
            If TypeOf ctrl Is MaskEdBox Then ctrl.Mask = "": ctrl.text = "": ctrl.Mask = "##/##/####"
            If TypeOf ctrl Is CheckBox Then ctrl.Value = 0
            If TypeOf ctrl Is usCredencial Then ctrl.Clear: ctrl.Tag = ""
            
            On Error Resume Next
            If ctrl.Name <> "LeyendaRef" And ctrl.Name <> "Leyenda" Then ctrl.Tag = ""
        
        End If

    Next

End Sub
'Buscamos el Empeno para ponerlo en el desempeño
Private Sub Buscar_Empeno(ByRef tNumContrato As TextBox, Indice As Long)

    Dim rcEmpeño As New ADODB.Recordset
    Dim rcTmp As New ADODB.Recordset
    
    Dim Band As Integer, Folio As Long, ImportePerdida As Double, ContratoAlmoneda As Integer, i As Integer, Serie As Integer, DiasProm  As Integer
    Dim crPrestamo As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double, crMinimo As Double, sqlPrendas As String, DiasMinimos As Integer, DiasTrans As Integer, DiasMes As Integer, FechaMes As Date, strDescripcion As String, crInteresDiario As Double, crTotalIntereses As Double, crInteresesPlazo As Double, crAlmacenajePlazo As Double, crSeguroPlazo As Double
    Dim Semaforo As String, Vencido As Integer
    Dim crCargoGPS As Double, crCargoSeguroAuto As Double, crIvaCargoGPSSeguro As Double
    Dim diasEnajenacion As Integer
    Dim FechaComercializacion As Date
    Dim diasTranscurridos As Integer

    Frame3.Visible = False
    Frame4.Visible = False
On Error GoTo Error

    Screen.MousePointer = vbHourglass
    
    If Trim(tNumContrato.text) = "" Then
        MsgBox "Introduzca el número de contrato a consultar !!", vbCritical, TPestañas.TabText(TPestañas.TabKey(TPestañas.SelectedTab))
        tNumContrato.SetFocus
        GoTo Error
    End If
    
    If chkAutomovil(Indice - 1).Value = 0 Then
    
        Serie = SERIE_A
    Else
        
        Serie = SERIE_B
        Frame3.Visible = True
        Frame4.Visible = True
    End If
    
    ImportePerdida = 0
    Folio = Val(tNumContrato)
    Limpiar IIf(Indice = 1, "DESEMPEÑO", "REFRENDOS")
        
    '***Puntos***
''''''    rcEmpeño.Open "SELECT e.*,c.ID AS IDCliente,CONCAT(c.nombre, ' ' ,c.apellido) AS Cliente,c.direccion,c.colonia,c.municipio,c.Estado,c.CP,c.Notas " & _
''''''                  "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.NumContrato=" & Folio & IIf(chkAutomovil(Indice - 1).Value = 1, " AND e.Serie=" & Serie, "") & " AND e.Cancelado=0 AND (e.Destino=0" & IIf(Indice = 1, ")", " OR e.Destino=" & D_ALMONEDA & ")"), dbDatos, adOpenForwardOnly, adLockOptimistic
  If Serie = 2 Then
       rcEmpeño.Open "SELECT e.*,c.ID AS IDCliente,CONCAT(c.nombre, ' ' ,c.apellido) AS Cliente,c.direccion,c.colonia,c.municipio,c.Estado,c.CP,c.Notas " & _
                  "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.NumContrato=" & Folio & " AND e.Serie=" & Serie & " AND e.Cancelado=0 AND (e.Destino=0 OR e.Destino=" & D_ALMONEDA & ")", dbDatos, adOpenForwardOnly, adLockOptimistic

  Else
       rcEmpeño.Open "SELECT e.*,c.ID AS IDCliente,CONCAT(c.nombre, ' ' ,c.apellido) AS Cliente,c.direccion,c.colonia,c.municipio,c.Estado,c.CP,c.Notas " & _
                  "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.NumContrato=" & Folio & IIf(chkAutomovil(Indice - 1).Value = 1, " AND e.Serie=" & Serie, "") & " AND e.Cancelado=0 AND (e.Destino=0 OR e.Destino=" & D_ALMONEDA & ")", dbDatos, adOpenForwardOnly, adLockOptimistic

  End If
     
    
    With rcEmpeño
        
        If .EOF Or .BOF Then
            
            MsgBox "No se encontró el contrato especificado !!", vbInformation, TPestañas.TabText(TPestañas.TabKey(TPestañas.SelectedTab))
            lblFolio.Tag = ""
                    
        Else
        
            If rcEmpeño!Perdida = 1 Then
                
                If MsgBox("Contrato marcado como perdido, desea cargar el importe de pérdida ??", vbQuestion + vbYesNo + vbDefaultButton2, TPestañas.TabText(TPestañas.TabKey(TPestañas.SelectedTab))) = vbYes Then
                    
                    ImportePerdida = Redondeo(Regresa_Valor_BD("ImportePerdida") * (1 + (Regresa_Valor_BD("Iva") / 100)))
                Else
                    
                    ImportePerdida = 0
                End If
            
            End If
                                                
            While Not .EOF
            
                If Band = 1 Then MsgBox "Contratos duplicados, verifique cual va a refrendar o desempeñar", vbCritical, TPestañas.TabText(TPestañas.TabKey(TPestañas.SelectedTab))
                
                If Indice = 1 Then txtFolioDesempeño.Tag = !ID Else txtFolioRefrendo.Tag = !ID
                
                tNumContrato.Tag = !ID
                Almoneda = !Almoneda
                
                Timer1.Enabled = False: labelContratoAlmoneda.Visible = False: labelContratoDesemp.Visible = False
                If !Destino = D_ALMONEDA Then
                    If Indice = 1 Then
                        labelContratoDesemp.Caption = "CONTRATO EN ALMONEDA!"
                        Timer2.Enabled = True
                    Else
                        labelContratoAlmoneda.Caption = "CONTRATO EN ALMONEDA!"
                        Timer1.Enabled = True
                    End If
                    Vencido = 1
                ElseIf Date > DateAdd("D", Val(Regresa_Valor_BD("DiasGracia")), !Vencimiento) Then
                    If Indice = 1 Then
                        labelContratoDesemp.Caption = "CONTRATO VENCIDO!"
                        Timer2.Enabled = True
                    Else
                        labelContratoAlmoneda.Caption = "CONTRATO VENCIDO!"
                        Timer1.Enabled = True
                    End If
                    Vencido = 1
                Else
                    Vencido = 0
                End If
                
                '***Puntos***
                'busco la tarjeta de cliente frecuente
                TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente !IDCliente
                '***Puntos***
                'si no fue buscado por la tarjeta de puntos
                If Indice = 2 Or Indice = 1 Then
                    If SacaValor("tarjetaspuntos", "count(id)", " where activa = 1") > 0 Then
                        If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(!IDCliente) = False Then
                           If MsgBox("El Cliente no cuenta con tarjeta de cliente frecuente" & vbCrLf & "Desea asignarle una tarjeta?", vbYesNoCancel Or vbQuestion) = vbYes Then
                              TarjetaPuntos.ShowAsignarTarjeta !IDCliente, frmMDI.IDUsuario
                              If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(!IDCliente) = False Then
                                  MsgBox "No se agregó la tarjeta al cliente", vbCritical, "Refrendo"
                              End If
                           End If
                        End If
                    End If
                End If
                
                'Se verifica eu sera prestamo de auto para calcular lo que se cobrara por GPS y Seguro de Auto
                  
                If !Serie = 2 Then
                If Indice = 1 Then
                     If !Circulando = 1 Then
                           chkCirculacionDes.Value = 1
                           'chkCirculacionRef.Enabled = False
                                                  
                        End If
                        
                        If !ImporteSeguroAuto > 0 Then
                            txtCargoSeguroDes.text = !ImporteSeguroAuto
                            crCargoSeguroAuto = !ImporteSeguroAuto
                        End If
                        txtCargoGPSDes.text = Regresa_Valor_BD("RentaGPS")
                        crCargoGPS = Regresa_Valor_BD("RentaGPS")
                Else
                     If !Circulando = 1 Then
                           chkCirculacionRef.Value = 1
                           'chkCirculacionRef.Enabled = False
                                                  
                        End If
                        
                        If !ImporteSeguroAuto > 0 Then
                            txtCargoSeguro.text = !ImporteSeguroAuto
                            crCargoSeguroAuto = !ImporteSeguroAuto
                        End If
                        txtCargoGPS.text = Regresa_Valor_BD("RentaGPS")
                        crCargoGPS = Regresa_Valor_BD("RentaGPS")
                End If
                       
                        
                End If
                
                
                ' Calcula los intereses
                                ' Calcula los intereses
                '//////// roger
                DiasMinimos = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres = ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo = tp.ID INNER JOIN plazos p ON ct.IDPlazo = p.ID", "DMinimos", " WHERE ti.Descripcion = '" & !TipoInteres & "' AND ti.Serie = " & !Serie & " AND tp.Descripcion='" & !TipoTasa & "' AND p.Descripcion=" & !VenPeriodo))
                '//////
                If !TipoInteres = "COMPLETO" Then
                    crIntereses = GeneraInteresesPeriodoCompleto(!ID, "Tasa")
                    crAlmacenaje = GeneraInteresesPeriodoCompleto(!ID, "Almacenaje")
                    crSeguro = GeneraInteresesPeriodoCompleto(!ID, "Seguro")
                Else
                    '/////////roger
                        crIntereses = GeneraIntereses(!ID, "Tasa", DiasMinimos)
                        crAlmacenaje = GeneraIntereses(!ID, "Almacenaje", DiasMinimos)
                        crSeguro = GeneraIntereses(!ID, "Seguro", DiasMinimos)
'                    //////////
                End If
                
                
'                If !TipoInteres = "COMPLETO" Then
'                    crIntereses = GeneraInteresesPeriodoCompleto(!ID, "Tasa")
'                    crAlmacenaje = GeneraInteresesPeriodoCompleto(!ID, "Almacenaje")
'                    crSeguro = GeneraInteresesPeriodoCompleto(!ID, "Seguro")
'                Else
'                        crIntereses = GeneraIntereses(!ID, "Tasa")
'                        crAlmacenaje = GeneraIntereses(!ID, "Almacenaje")
'                        crSeguro = GeneraIntereses(!ID, "Seguro")
'                End If

                crInteresesPlazo = Regresa_Intereses_Plazo(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoInteres)
                crAlmacenajePlazo = Regresa_Almacenaje_Plazo(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoTasa)
                crSeguroPlazo = Regresa_Seguro_Plazo(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoTasa)
                crTotalIntereses = (crInteresesPlazo + crAlmacenajePlazo + crSeguroPlazo)
                
                crMoratorios = 0
                If !Serie = 2 Then
                    If Date > DateAdd("D", Regresa_Valor_BD("DiasGraciaAuto"), !Vencimiento) Then crMoratorios = crTotalIntereses * (Regresa_Valor_BD("Operacion") / 100)
             
                Else
                    If Date > !Vencimiento Then crMoratorios = crTotalIntereses * (Regresa_Valor_BD("Operacion") / 100)
              
                End If
                
                crIva = Regresa_Iva(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crCargoGPS + crCargoSeguroAuto, !ID)
                crIvaCargoGPSSeguro = Redondeo(Regresa_Iva(crCargoGPS + crCargoSeguroAuto, !ID))
                crMinimo = Redondeo(Regresa_Valor_BD("PagoMinimo") * (1 + (Regresa_Valor_BD("IVA") / 100)))
                
                
                '***************************
                'Muestro los datos
                DatosCliente(Indice - 1).Tag = !Cliente
                DatosCliente(Indice - 1).Add "<bold> " & !Cliente & "</bold>"
                DatosCliente(Indice - 1).Add " " & !Direccion & " " & !Colonia & vbCrLf & _
                                             " " & !Municipio & ", " & !Estado & " C.P. " & !CP & vbCrLf & _
                                            IIf(IsNull(!Notas) Or Trim(!Notas) = "", "", " MENSAJE: " & !Notas & vbCrLf) & IIf(IsNull(!Responsable) Or Trim(!Responsable) = "", "", " COTITULAR: " & !Responsable & vbCrLf) & IIf(!NumBolsa <> "", " NUM. BOLSA: " & !NumBolsa & vbCrLf, "") & " USUARIO: " & SacaValor("usuarios", "Nombre", " WHERE ID=" & !IDUsuario)
                
                '***Puntos***
                DatosCliente(Indice - 1).Add "No. Tarjeta: " & TarjetaPuntos.CuentaFrecuente.Folio & vbCrLf & _
                                            "Puntos Acumulados: " & TarjetaPuntos.CuentaFrecuente.Puntos
                 'Saco dias enajenacio
                 diasEnajenacion = SacaValor("parametros", "diasEnajenacion")
                 FechaComercializacion = DateAdd("d", diasEnajenacion, !Vencimiento)
                 diasTranscurridos = DateDiff("d", !Fecha, Date)
                 diasTranscurridos = IIf(diasTranscurridos = 0, 1, diasTranscurridos)
                
                If !Serie = 2 Or !Periodo <> 1 Then
                    DatosContrato(Indice - 1).Add " FECHA EMPEÑO: " & Format(!Fecha, "DD/MMM/YYYY HH:MM:SS AM/PM") & vbCrLf & _
                                            " FECHA VENCIMIENTO: " & Format(!Vencimiento, "DD/MMM/YYYY") & vbCrLf & _
                                            " FECHA COMERCIALIZACION: " & Format(FechaComercializacion, "DD/MMM/YYYY") & vbCrLf & _
                                            " DIAS TRANSCURRIDOS: " & diasTranscurridos & vbCrLf & _
                                            " PLAZO: " & !TipoInteres & " " & !VenPeriodo & " " & IIf(!TipoTasa = "MENSUAL", "MESES", IIf(!TipoTasa = "QUINCENAL", "QUINCENAS", IIf(!TipoTasa = "SEMANAL", "SEMANAS", "DIAS"))) & vbCrLf & _
                                            " TASA: " & Format(((!Tasa + !Almacenaje + !Seguro) * (1 + (!Iva / 100))), "0.00") & "%" & vbCrLf & _
                                            IIf(!Promocion > 0, " CONTRATO PROMOCIÓN " & LeyendaPromocion(!Promocion), "")
                    
                Else
                    DatosContrato(Indice - 1).Add " FECHA EMPEÑO: " & Format(!Fecha, "DD/MMM/YYYY HH:MM:SS AM/PM") & vbCrLf & _
                                            " FECHA VENCIMIENTO: " & Format(!Vencimiento, "DD/MMM/YYYY") & vbCrLf & _
                                            " FECHA COMERCIALIZACION: " & Format(FechaComercializacion, "DD/MMM/YYYY") & vbCrLf & _
                                            " DIAS TRANSCURRIDOS: " & diasTranscurridos & vbCrLf & _
                                            " PLAZO: " & !TipoInteres & " " & !VenPeriodo & " " & IIf(!TipoTasa = "MENSUAL", "MESES", IIf(!TipoTasa = "QUINCENAL", "QUINCENAS", IIf(!TipoTasa = "SEMANAL", "SEMANAS", "DIAS"))) & vbCrLf & _
                                            " TASA: " & Format(((!Tasa + !Almacenaje + !Seguro) * (1 + (!Iva / 100))), "0.00") & "%" & vbCrLf & _
                                            " INTERES DIARIO: " & Format(((crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva) / diasTranscurridos), "0.00") & vbCrLf & _
                                            IIf(!Promocion > 0, " CONTRATO PROMOCIÓN " & LeyendaPromocion(!Promocion), "")
                
                End If
                
                                                           
                 'Semaforo
                Semaforo = Regresa_Semaforo(!IDCliente)
    
                If Semaforo = "Verde" Then
                    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\VERDE.bmp")
                ElseIf Semaforo = "Amarillo" Then
                    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\AMARILLO.bmp")
                Else
                    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\ROJO.bmp")
                End If
                                                           
                'Saco las Prendas
                If rcEmpeño!Destino = 0 Then
                    
                    If Not !Serie = 2 Then
                    sqlPrendas = "SELECT d.Cantidad,d.Articulo AS Descripcion,d.Peso,d.Prestamo,d.Marca,d.Modelo,kilatajes.Descripcion AS Kilataje " & _
                                "FROM detallesempeno d LEFT JOIN kilatajes ON d.Kilates=kilatajes.ID WHERE d.IDEmpeno=" & !ID
                    Else
                     sqlPrendas = "SELECT * FROM detallesempenoautos  WHERE IDEmpeno=" & !ID

                    End If
                    ContratoAlmoneda = 0
                Else
                    
                    sqlPrendas = "SELECT d.Cantidad,d.Descripcion,d.Peso,d.Costo AS Prestamo,d.Marca,d.Modelo,kilatajes.Descripcion AS Kilataje " & _
                                "FROM detallesentradainventario d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE  d.IDEmpeno=" & !ID
'                                                                "FROM detallesentradainventario d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE d.Cantidad>0 AND d.IDEmpeno=" & !ID
                    
                    ContratoAlmoneda = 1
                End If
                
                crPrestamo = 0
                rcTmp.Open sqlPrendas, dbDatos, adOpenForwardOnly, adLockReadOnly
                While Not rcTmp.EOF
                    If Not !Serie = 2 Then
                        For i = 1 To Len(rcTmp!Cantidad & " " & rcTmp!Descripcion & IIf(rcTmp!Peso > 0, " " & rcTmp!Peso & " Grms.", "") & IIf(IsNull(rcTmp!Kilataje) Or Trim(rcTmp!Kilataje) = "", "", " " & rcTmp!Kilataje) & IIf(IsNull(rcTmp!Marca) Or Trim(rcTmp!Marca) = "", "", " MARCA: " & rcTmp!Marca) & IIf(IsNull(rcTmp!Modelo) Or Trim(rcTmp!Modelo) = "", "", " MODELO: " & rcTmp!Modelo)) Step 50
                            strDescripcion = strDescripcion & IIf(Trim(strDescripcion) <> "", vbCrLf, "") & " " & Trim(Mid(rcTmp!Cantidad & " " & rcTmp!Descripcion & IIf(rcTmp!Peso > 0, " " & rcTmp!Peso & " Grms.", "") & IIf(IsNull(rcTmp!Kilataje) Or Trim(rcTmp!Kilataje) = "", "", " " & rcTmp!Kilataje) & IIf(IsNull(rcTmp!Marca) Or Trim(rcTmp!Marca) = "", "", " MARCA: " & rcTmp!Marca) & IIf(IsNull(rcTmp!Modelo) Or Trim(rcTmp!Modelo) = "", "", " MODELO: " & rcTmp!Modelo), i * 1, 50))
                        Next i
                    Else
                    If ContratoAlmoneda = 1 Then
                         For i = 1 To Len(rcTmp!Descripcion) Step 50
                            strDescripcion = strDescripcion & IIf(Trim(strDescripcion) <> "", vbCrLf, "") & " " & Trim(Mid(rcTmp!Descripcion, i * 1, 50))
                        Next i
                    Else
                         For i = 1 To Len(rcTmp!MarcayModelo & " " & rcTmp!Año & IIf(IsNull(rcTmp!Color) Or Trim(rcTmp!Color) = "", " ", rcTmp!Color) & IIf(IsNull(rcTmp!Placas) Or Trim(rcTmp!Placas) = "", "", " " & rcTmp!Placas) & IIf(IsNull(rcTmp!SerieChasis) Or Trim(rcTmp!SerieChasis) = "", "", " SERIE: " & rcTmp!SerieChasis) & IIf(IsNull(rcTmp!Kms) Or Trim(rcTmp!Kms) = "", "", " KMS: " & rcTmp!Kms)) Step 50
                            strDescripcion = strDescripcion & IIf(Trim(strDescripcion) <> "", vbCrLf, "") & " " & Trim(Mid(rcTmp!MarcayModelo & " " & rcTmp!Año & IIf(IsNull(rcTmp!Color) Or Trim(rcTmp!Color) = "", " ", rcTmp!Color) & IIf(IsNull(rcTmp!Placas) Or Trim(rcTmp!Placas) = "", "", " " & rcTmp!Placas) & IIf(IsNull(rcTmp!SerieChasis) Or Trim(rcTmp!SerieChasis) = "", "", " SERIE: " & rcTmp!SerieChasis) & IIf(IsNull(rcTmp!Kms) Or Trim(rcTmp!Kms) = "", "", " KMS: " & rcTmp!Kms), i * 1, 50))
                        Next i
                    End If
                        
                    End If
                rcTmp.MoveNext
                Wend
               
                
'                If rcTmp.EOF And rcTmp.BOF Then
'                If Indice = 1 Then
'                  MsgBox "La prenda no se puede Desempeñar", vbCritical
'                Limpiar "DESEMPEÑO"
'                Screen.MousePointer = vbDefault
'                Exit Sub
'                End If
'
'                If Indice = 2 Then
'                MsgBox "La prenda no se puede Refrendar", vbCritical
'                Limpiar "REFRENDOS"
'                Screen.MousePointer = vbDefault
'                Exit Sub
'                End If
'                End If
                 rcTmp.Close
                Set rcTmp = Nothing
                
                'Tomo el Préstamo
                crPrestamo = !Prestamo
                
                'Imprimo la descripción
                DetallesContrato(Indice - 1).Add " " & strDescripcion
              
'                If Serie = SERIE_A Then
'
'                    'Saco los Intereses de los dias que han pasado
'                    FechaMes = Format(!Fecha, "DD/MM/YYYY")
'                    FechaMes = DateAdd(IIf(!TipoTasa = "MENSUAL", "M", "D"), IIf(!TipoTasa = "MENSUAL", 1, IIf(!TipoTasa = "QUINCENAL", 15 * 2, 7 * 4)), FechaMes)
'                    DiasMes = DateDiff("D", !Fecha, FechaMes)
'                    DiasTrans = DateDiff("D", !Fecha, Date)
'                    If DiasTrans = 0 Then DiasTrans = 1
'                    DiasMinimos = 0
'
'                    crIntereses = Redondeo(((!Prestamo * ((!Tasa / 100) / !Periodo)) * DiasTrans) - ChecaPromocion(!ID, 0, "Tasa", IIf(DiasMinimos > DiasTrans, DiasMinimos, IIf(DiasTrans > DiasMes, DiasMes, DiasTrans)), True))
'                    crAlmacenaje = Redondeo(((!Prestamo * ((!Almacenaje / 100) / !Periodo)) * DiasTrans) - ChecaPromocion(!ID, 0, "Almacenaje", IIf(DiasMinimos > DiasTrans, DiasMinimos, IIf(DiasTrans > DiasMes, DiasMes, DiasTrans)), True))
'                    crSeguro = Redondeo(((!Prestamo * ((!Seguro / 100) / !Periodo)) * DiasTrans) - ChecaPromocion(!ID, 0, "Seguro", IIf(DiasMinimos > DiasTrans, DiasMinimos, IIf(DiasTrans > DiasMes, DiasMes, DiasTrans)), True))
'                Else
'
'                    crIntereses = Redondeo(Regresa_Intereses(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoTasa))
'                    crAlmacenaje = Redondeo(Regresa_Almacenaje(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoTasa))
'                    crSeguro = Redondeo(Regresa_Seguro(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoTasa))
'                    crInteresDiario = Redondeo((crIntereses + crAlmacenaje + crSeguro) / DateDiff("D", !Fecha, !Vencimiento))
'                End If
                
'                If !TipoInteres = "COMPLETO" Then
'                    crIntereses = GeneraInteresesPeriodoCompleto(!ID, "Tasa")
'                    crAlmacenaje = GeneraInteresesPeriodoCompleto(!ID, "Almacenaje")
'                    crSeguro = GeneraInteresesPeriodoCompleto(!ID, "Seguro")
'                Else
'                    crIntereses = GeneraIntereses(!ID, "Tasa")
'                    crAlmacenaje = GeneraIntereses(!ID, "Almacenaje")
'                    crSeguro = GeneraIntereses(!ID, "Seguro")
'                End If
'
'                crMoratorios = 0
'                If Date > !Vencimiento Then crMoratorios = Redondeo(!Prestamo * (Regresa_Valor_BD("Operacion") / 100))
'                crIva = Redondeo(Regresa_Iva(crIntereses + crAlmacenaje + crSeguro + crMoratorios, !ID))
'                crMinimo = Redondeo(Regresa_Valor_BD("PagoMinimo") * (1 + (Regresa_Valor_BD("IVA") / 100)))
'
                If Indice = 1 Then
                
                    'Cargamos los datos
                    grdDesempeño.Redraw = False
                    grdDesempeño.AddRow
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 1) = !NumContrato
                    grdDesempeño.CellItemData(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 1) = !Serie
                    grdDesempeño.CellTextAlign(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 1) = DT_CENTER Or DT_WORD_ELLIPSIS
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 2) = IIf(Serie = SERIE_B, !Prestamo, crPrestamo)
                    grdDesempeño.CellTextAlign(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 3) = IIf((crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva) < crMinimo, crMinimo + ImportePerdida, (crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva) + ImportePerdida)
                    grdDesempeño.CellItemData(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 3) = ImportePerdida
                    grdDesempeño.CellTextAlign(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 4) = !Folio
                    grdDesempeño.CellItemData(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 4) = tNumContrato.Tag
                                                          
                    'Checo si el pago mínimo es mayor que el importe de los intereses
                    If (crIntereses + crAlmacenaje + crSeguro + crIva) < crMinimo Then
                        crIntereses = Redondeo(crMinimo / (1 + (Val(Regresa_Valor_BD("IVA")) / 100)))
                        crIva = Redondeo(crMinimo - crIntereses)
                        crAlmacenaje = 0
                        crSeguro = 0
                    End If
                
                    'Intereses
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 5) = crIntereses
                    'Almacenaje
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 6) = crAlmacenaje
                    'Seguro
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 7) = crSeguro
                    'Iva
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 8) = crIva
                    'Si es contrato de Almoneda
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 9) = ContratoAlmoneda
                    'Importe por boleta perdida
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 10) = ImportePerdida
                    'Moratorios
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 11) = crMoratorios
                    
                    '***Puntos***
                    'IDCliente
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 17) = !IDCliente
                    'IDEmpeno
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 18) = !ID
                   'Iva Cargo Seguro y GPS
                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 19) = crIvaCargoGPSSeguro
                   
                    If grdDesempeño.Rows = 1 Then grdDesempeño.AddRow
                    grdDesempeño.Redraw = True
                    Poner_Totales_Desempeño
                    
                ElseIf Indice = 2 Then
                
                    grdRefrendos.Redraw = False
                    grdRefrendos.AddRow
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 1) = !NumContrato
                    grdRefrendos.CellItemData(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 1) = tNumContrato.Tag
                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 1) = DT_CENTER Or DT_WORD_ELLIPSIS
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 2) = IIf(Serie = SERIE_B, !Prestamo, crPrestamo)
                    grdRefrendos.CellItemData(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 2) = !ID
                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 3) = "0"
                    grdRefrendos.CellItemData(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 3) = !Serie
                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 3) = DT_RIGHT
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 4) = IIf((crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva) < crMinimo, crMinimo + ImportePerdida, (crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva) + ImportePerdida)
                    grdRefrendos.CellItemData(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 4) = ImportePerdida
                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 4) = DT_RIGHT Or DT_WORD_ELLIPSIS
                    
                    'Checo si el pago mínimo es mayor que el importe de los intereses
                    If (crIntereses + crAlmacenaje + crSeguro + crIva) < crMinimo Then
                        crIntereses = Redondeo(crMinimo / (1 + (Val(Regresa_Valor_BD("IVA")) / 100)))
                        crIva = Redondeo(crMinimo - crIntereses)
                        crAlmacenaje = 0
                        crSeguro = 0
                    End If
                    
                    'Intereses
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 5) = crIntereses
                    'Almacenaje
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 6) = crAlmacenaje
                    'Seguro
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 7) = crSeguro
                    'Iva
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 8) = crIva
                    'Si es contrato de Almoneda
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 9) = ContratoAlmoneda
                    'Importe Boleta Perdida
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 10) = ImportePerdida
                    'Moratorios
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 11) = crMoratorios
                    'Vencido
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 18) = Vencido
                    'IDEmpeno
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 19) = !ID
                   'Iva Cargo Seguro y GPS
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 20) = crIvaCargoGPSSeguro
                   
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 13) = 0
                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 13) = DT_RIGHT Or DT_WORD_ELLIPSIS
                    
                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 15) = 0
                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 15) = DT_RIGHT Or DT_WORD_ELLIPSIS
                    
                    If grdRefrendos.Rows = 1 Then grdRefrendos.AddRow
                    grdRefrendos.Redraw = True
                    
                    Poner_Totales_Refrendo
                End If
                
            .MoveNext
            Band = Band + 1
            Wend
        End If
    End With
    rcEmpeño.Close
    Set rcEmpeño = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcEmpeño = Nothing
    Set rcTmp = Nothing
    Screen.MousePointer = vbDefault
End Sub
''Buscamos el Empeno para ponerlo en el desempeño
'Private Sub Buscar_Empeno(ByRef tNumContrato As TextBox, Indice As Long)
'
'    Dim rcEmpeño As New ADODB.Recordset
'    Dim rcTmp As New ADODB.Recordset
'
'    Dim Band As Integer, Folio As Long, ImportePerdida As Double, ContratoAlmoneda As Integer, i As Integer, Serie As Integer, DiasProm  As Integer
'    Dim crPrestamo As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double, crMinimo As Double, sqlPrendas As String, DiasMinimos As Integer, DiasTrans As Integer, DiasMes As Integer, FechaMes As Date, strDescripcion As String, crInteresDiario As Double, crTotalIntereses As Double, crInteresesPlazo As Double, crAlmacenajePlazo As Double, crSeguroPlazo As Double
'    Dim Semaforo As String, Vencido As Integer
'    Dim crCargoGPS As Double, crCargoSeguroAuto As Double, crIvaCargoGPSSeguro As Double
'    Dim DiasEnajenacion As Integer
'    Dim FechaComercializacion As Date
'    Dim diasTranscurridos As Integer
'
'    Frame3.Visible = False
'    Frame4.Visible = False
'On Error GoTo error
'
'    Screen.MousePointer = vbHourglass
'
'    If Trim(tNumContrato.text) = "" Then
'        MsgBox "Introduzca el número de contrato a consultar !!", vbCritical, TPestañas.TabText(TPestañas.TabKey(TPestañas.SelectedTab))
'        tNumContrato.SetFocus
'        GoTo error
'    End If
'
'    If chkAutomovil(Indice - 1).Value = 0 Then
'
'        Serie = SERIE_A
'    Else
'
'        Serie = SERIE_B
'        Frame3.Visible = True
'        Frame4.Visible = True
'    End If
'
'    ImportePerdida = 0
'    Folio = Val(tNumContrato)
'    Limpiar IIf(Indice = 1, "DESEMPEÑO", "REFRENDOS")
'
'    '***Puntos***
'''''''    rcEmpeño.Open "SELECT e.*,c.ID AS IDCliente,CONCAT(c.nombre, ' ' ,c.apellido) AS Cliente,c.direccion,c.colonia,c.municipio,c.Estado,c.CP,c.Notas " & _
'''''''                  "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.NumContrato=" & Folio & IIf(chkAutomovil(Indice - 1).Value = 1, " AND e.Serie=" & Serie, "") & " AND e.Cancelado=0 AND (e.Destino=0" & IIf(Indice = 1, ")", " OR e.Destino=" & D_ALMONEDA & ")"), dbDatos, adOpenForwardOnly, adLockOptimistic
'  If Serie = 2 Then
'       rcEmpeño.Open "SELECT e.*,c.ID AS IDCliente,CONCAT(c.nombre, ' ' ,c.apellido) AS Cliente,c.direccion,c.colonia,c.municipio,c.Estado,c.CP,c.Notas " & _
'                  "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.NumContrato=" & Folio & " AND e.Serie=" & Serie & " AND e.Cancelado=0 AND (e.Destino=0 OR e.Destino=" & D_ALMONEDA & ")", dbDatos, adOpenForwardOnly, adLockOptimistic
'
'  Else
'       rcEmpeño.Open "SELECT e.*,c.ID AS IDCliente,CONCAT(c.nombre, ' ' ,c.apellido) AS Cliente,c.direccion,c.colonia,c.municipio,c.Estado,c.CP,c.Notas " & _
'                  "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.NumContrato=" & Folio & IIf(chkAutomovil(Indice - 1).Value = 1, " AND e.Serie=" & Serie, "") & " AND e.Cancelado=0 AND (e.Destino=0 OR e.Destino=" & D_ALMONEDA & ")", dbDatos, adOpenForwardOnly, adLockOptimistic
'
'  End If
'
'
'    With rcEmpeño
'
'        If .EOF Or .BOF Then
'
'            MsgBox "No se encontró el contrato especificado !!", vbInformation, TPestañas.TabText(TPestañas.TabKey(TPestañas.SelectedTab))
'            lblFolio.Tag = ""
'
'        Else
'
'            If rcEmpeño!Perdida = 1 Then
'
'                If MsgBox("Contrato marcado como perdido, desea cargar el importe de pérdida ??", vbQuestion + vbYesNo + vbDefaultButton2, TPestañas.TabText(TPestañas.TabKey(TPestañas.SelectedTab))) = vbYes Then
'
'                    ImportePerdida = Redondeo(Regresa_Valor_BD("ImportePerdida") * (1 + (Regresa_Valor_BD("Iva") / 100)))
'                Else
'
'                    ImportePerdida = 0
'                End If
'
'            End If
'
'            While Not .EOF
'
'                If Band = 1 Then MsgBox "Contratos duplicados, verifique cual va a refrendar o desempeñar", vbCritical, TPestañas.TabText(TPestañas.TabKey(TPestañas.SelectedTab))
'
'                If Indice = 1 Then txtFolioDesempeño.Tag = !ID Else txtFolioRefrendo.Tag = !ID
'
'                tNumContrato.Tag = !ID
'                Almoneda = !Almoneda
'
'                Timer1.Enabled = False: labelContratoAlmoneda.Visible = False: labelContratoDesemp.Visible = False
'                If !Destino = D_ALMONEDA Then
'                    If Indice = 1 Then
'                        labelContratoDesemp.Caption = "CONTRATO EN ALMONEDA!"
'                        Timer2.Enabled = True
'                    Else
'                        labelContratoAlmoneda.Caption = "CONTRATO EN ALMONEDA!"
'                        Timer1.Enabled = True
'                    End If
'                    Vencido = 1
'                ElseIf Date > DateAdd("D", Val(Regresa_Valor_BD("DiasGracia")), !Vencimiento) Then
'                    If Indice = 1 Then
'                        labelContratoDesemp.Caption = "CONTRATO VENCIDO!"
'                        Timer2.Enabled = True
'                    Else
'                        labelContratoAlmoneda.Caption = "CONTRATO VENCIDO!"
'                        Timer1.Enabled = True
'                    End If
'                    Vencido = 1
'                Else
'                    Vencido = 0
'                End If
'
'                '***Puntos***
'                'busco la tarjeta de cliente frecuente
'                TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente !IDCliente
'                '***Puntos***
'                'si no fue buscado por la tarjeta de puntos
'                If Indice = 2 Or Indice = 1 Then
'                    If SacaValor("tarjetaspuntos", "count(id)", " where activa = 1") > 0 Then
'                        If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(!IDCliente) = False Then
'                           If MsgBox("El Cliente no cuenta con tarjeta de cliente frecuente" & vbCrLf & "Desea asignarle una tarjeta?", vbYesNoCancel Or vbQuestion) = vbYes Then
'                              TarjetaPuntos.ShowAsignarTarjeta !IDCliente, frmMDI.IDUsuario
'                              If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(!IDCliente) = False Then
'                                  MsgBox "No se agregó la tarjeta al cliente", vbCritical, "Refrendo"
'                              End If
'                           End If
'                        End If
'                    End If
'                End If
'
'                'Se verifica eu sera prestamo de auto para calcular lo que se cobrara por GPS y Seguro de Auto
'
'                If !Serie = 2 Then
'                If Indice = 1 Then
'                     If !Circulando = 1 Then
'                           chkCirculacionDes.Value = 1
'                           'chkCirculacionRef.Enabled = False
'
'                        End If
'
'                        If !ImporteSeguroAuto > 0 Then
'                            txtCargoSeguroDes.text = !ImporteSeguroAuto
'                            crCargoSeguroAuto = !ImporteSeguroAuto
'                        End If
'                        txtCargoGPSDes.text = Regresa_Valor_BD("RentaGPS")
'                        crCargoGPS = Regresa_Valor_BD("RentaGPS")
'                Else
'                     If !Circulando = 1 Then
'                           chkCirculacionRef.Value = 1
'                           'chkCirculacionRef.Enabled = False
'
'                        End If
'
'                        If !ImporteSeguroAuto > 0 Then
'                            txtCargoSeguro.text = !ImporteSeguroAuto
'                            crCargoSeguroAuto = !ImporteSeguroAuto
'                        End If
'                        txtCargoGPS.text = Regresa_Valor_BD("RentaGPS")
'                        crCargoGPS = Regresa_Valor_BD("RentaGPS")
'                End If
'
'
'                End If
'
'
'                ' Calcula los intereses
'                '//////// roger
'                DiasMinimos = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres = ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo = tp.ID INNER JOIN plazos p ON ct.IDPlazo = p.ID", "DMinimos", " WHERE ti.Descripcion = '" & !TipoInteres & "' AND ti.Serie = " & !Serie & " AND tp.Descripcion='" & !TipoTasa & "' AND p.Descripcion=" & !VenPeriodo))
'                '//////
'                If !TipoInteres = "COMPLETO" Then
'                    crIntereses = GeneraInteresesPeriodoCompleto(!ID, "Tasa")
'                    crAlmacenaje = GeneraInteresesPeriodoCompleto(!ID, "Almacenaje")
'                    crSeguro = GeneraInteresesPeriodoCompleto(!ID, "Seguro")
'                Else
'                    '/////////roger
'                        crIntereses = GeneraIntereses(!ID, "Tasa", DiasMinimos)
'                        crAlmacenaje = GeneraIntereses(!ID, "Almacenaje", DiasMinimos)
'                        crSeguro = GeneraIntereses(!ID, "Seguro", DiasMinimos)
'                    '//////////
'                End If
'
'
''                If !TipoInteres = "COMPLETO" Then
''                    crIntereses = GeneraInteresesPeriodoCompleto(!ID, "Tasa")
''                    crAlmacenaje = GeneraInteresesPeriodoCompleto(!ID, "Almacenaje")
''                    crSeguro = GeneraInteresesPeriodoCompleto(!ID, "Seguro")
''                Else
''                        crIntereses = GeneraIntereses(!ID, "Tasa")
''                        crAlmacenaje = GeneraIntereses(!ID, "Almacenaje")
''                        crSeguro = GeneraIntereses(!ID, "Seguro")
''                End If
'
''                crInteresesPlazo = Regresa_Intereses_Plazo(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoInteres)
''                crAlmacenajePlazo = Regresa_Almacenaje_Plazo(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoInteres)
''                crSeguroPlazo = Regresa_Seguro_Plazo(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoInteres)
''                crTotalIntereses = (crInteresesPlazo + crAlmacenajePlazo + crSeguroPlazo)
'
'                crMoratorios = 0
'                If !Serie = 2 Then
'                    If Date > DateAdd("D", Regresa_Valor_BD("DiasGraciaAuto"), !Vencimiento) Then
'
'                    crMoratorios = Redondeo(!Prestamo * (Regresa_Valor_BD("Operacion") / 100))
'
'                Else
'                    If Date > !Vencimiento Then crMoratorios = Redondeo(!Prestamo * (Regresa_Valor_BD("Operacion") / 100))
'
'                End If
'
'                crIva = Redondeo(Regresa_Iva(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crCargoGPS + crCargoSeguroAuto, !ID))
'                crIvaCargoGPSSeguro = Redondeo(Regresa_Iva(crCargoGPS + crCargoSeguroAuto, !ID))
'                crMinimo = Redondeo(Regresa_Valor_BD("PagoMinimo") * (1 + (Regresa_Valor_BD("IVA") / 100)))
'
'
'                '***************************
'
'                'Muestro los datos
'                DatosCliente(Indice - 1).Tag = !Cliente
'                DatosCliente(Indice - 1).Add "<bold> " & !Cliente & "</bold>"
'                DatosCliente(Indice - 1).Add " " & !Direccion & " " & !Colonia & vbCrLf & _
'                                             " " & !Municipio & ", " & !Estado & " C.P. " & !CP & vbCrLf & _
'                                            IIf(IsNull(!Notas) Or Trim(!Notas) = "", "", " MENSAJE: " & !Notas & vbCrLf) & IIf(IsNull(!Responsable) Or Trim(!Responsable) = "", "", " COTITULAR: " & !Responsable & vbCrLf) & IIf(!NumBolsa <> "", " NUM. BOLSA: " & !NumBolsa & vbCrLf, "") & " USUARIO: " & SacaValor("usuarios", "Nombre", " WHERE ID=" & !IDUsuario)
'
'                '***Puntos***
'                DatosCliente(Indice - 1).Add "No. Tarjeta: " & TarjetaPuntos.CuentaFrecuente.Folio & vbCrLf & _
'                                            "Puntos Acumulados: " & TarjetaPuntos.CuentaFrecuente.Puntos
'
'
'
'                 'Saco dias enajenacio
'                 DiasEnajenacion = SacaValor("parametros", "diasEnajenacion")
'                 FechaComercializacion = DateAdd("d", DiasEnajenacion, !Vencimiento)
'                 diasTranscurridos = DateDiff("d", !Fecha, Date)
'                 diasTranscurridos = IIf(diasTranscurridos = 0, 1, diasTranscurridos)
'
'                If !Serie = 2 Or !Periodo <> 1 Then
'                    DatosContrato(Indice - 1).Add " FECHA EMPEÑO: " & Format(!Fecha, "DD/MMM/YYYY HH:MM:SS AM/PM") & vbCrLf & _
'                                            " FECHA VENCIMIENTO: " & Format(!Vencimiento, "DD/MMM/YYYY") & vbCrLf & _
'                                            " FECHA COMERCIALIZACION: " & Format(FechaComercializacion, "DD/MMM/YYYY") & vbCrLf & _
'                                            " DIAS TRANSCURRIDOS: " & diasTranscurridos & vbCrLf & _
'                                            " PLAZO: " & !TipoInteres & " " & !VenPeriodo & " " & IIf(!TipoTasa = "MENSUAL", "MESES", IIf(!TipoTasa = "QUINCENAL", "QUINCENAS", IIf(!TipoTasa = "SEMANAL", "SEMANAS", "DIAS"))) & vbCrLf & _
'                                            " TASA: " & Format(((!Tasa + !Almacenaje + !Seguro) * (1 + (!Iva / 100))), "0.00") & "%" & vbCrLf & _
'                                            IIf(!Promocion > 0, " CONTRATO PROMOCIÓN " & LeyendaPromocion(!Promocion), "")
'
'                Else
'                    DatosContrato(Indice - 1).Add " FECHA EMPEÑO: " & Format(!Fecha, "DD/MMM/YYYY HH:MM:SS AM/PM") & vbCrLf & _
'                                            " FECHA VENCIMIENTO: " & Format(!Vencimiento, "DD/MMM/YYYY") & vbCrLf & _
'                                            " FECHA COMERCIALIZACION: " & Format(FechaComercializacion, "DD/MMM/YYYY") & vbCrLf & _
'                                            " DIAS TRANSCURRIDOS: " & diasTranscurridos & vbCrLf & _
'                                            " PLAZO: " & !TipoInteres & " " & !VenPeriodo & " " & IIf(!TipoTasa = "MENSUAL", "MESES", IIf(!TipoTasa = "QUINCENAL", "QUINCENAS", IIf(!TipoTasa = "SEMANAL", "SEMANAS", "DIAS"))) & vbCrLf & _
'                                            " TASA: " & Format(((!Tasa + !Almacenaje + !Seguro) * (1 + (!Iva / 100))), "0.00") & "%" & vbCrLf & _
'                                            " INTERES DIARIO: " & Format(((crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva) / diasTranscurridos), "0.00") & vbCrLf & _
'                                            IIf(!Promocion > 0, " CONTRATO PROMOCIÓN " & LeyendaPromocion(!Promocion), "")
'
'                End If
'
'
'                 'Semaforo
'                 Semaforo = Regresa_Semaforo(!IDCliente)
'
'                If Semaforo = "Verde" Then
'                    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\VERDE.bmp")
'                ElseIf Semaforo = "Amarillo" Then
'                    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\AMARILLO.bmp")
'                Else
'                    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\ROJO.bmp")
'                End If
'
'                'Saco las Prendas
'                If rcEmpeño!Destino = 0 Then
'
'                    If Not !Serie = 2 Then
'                    sqlPrendas = "SELECT d.Cantidad,d.Articulo AS Descripcion,d.Peso,d.Prestamo,d.Marca,d.Modelo,kilatajes.Descripcion AS Kilataje " & _
'                                "FROM detallesempeno d LEFT JOIN kilatajes ON d.Kilates=kilatajes.ID WHERE d.IDEmpeno=" & !ID
'                    Else
'                     sqlPrendas = "SELECT * FROM detallesempenoautos  WHERE IDEmpeno=" & !ID
'
'                    End If
'                    ContratoAlmoneda = 0
'                Else
'
'                    sqlPrendas = "SELECT d.Cantidad,d.Descripcion,d.Peso,d.Costo AS Prestamo,d.Marca,d.Modelo,kilatajes.Descripcion AS Kilataje " & _
'                                "FROM detallesentradainventario d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE  d.IDEmpeno=" & !ID
''                                                                "FROM detallesentradainventario d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE d.Cantidad>0 AND d.IDEmpeno=" & !ID
'
'                    ContratoAlmoneda = 1
'                End If
'
'                crPrestamo = 0
'                rcTmp.Open sqlPrendas, dbDatos, adOpenForwardOnly, adLockReadOnly
'                While Not rcTmp.EOF
'                    If Not !Serie = 2 Then
'                        For i = 1 To Len(rcTmp!Cantidad & " " & rcTmp!Descripcion & IIf(rcTmp!Peso > 0, " " & rcTmp!Peso & " Grms.", "") & IIf(IsNull(rcTmp!Kilataje) Or Trim(rcTmp!Kilataje) = "", "", " " & rcTmp!Kilataje) & IIf(IsNull(rcTmp!Marca) Or Trim(rcTmp!Marca) = "", "", " MARCA: " & rcTmp!Marca) & IIf(IsNull(rcTmp!Modelo) Or Trim(rcTmp!Modelo) = "", "", " MODELO: " & rcTmp!Modelo)) Step 50
'                            strDescripcion = strDescripcion & IIf(Trim(strDescripcion) <> "", vbCrLf, "") & " " & Trim(Mid(rcTmp!Cantidad & " " & rcTmp!Descripcion & IIf(rcTmp!Peso > 0, " " & rcTmp!Peso & " Grms.", "") & IIf(IsNull(rcTmp!Kilataje) Or Trim(rcTmp!Kilataje) = "", "", " " & rcTmp!Kilataje) & IIf(IsNull(rcTmp!Marca) Or Trim(rcTmp!Marca) = "", "", " MARCA: " & rcTmp!Marca) & IIf(IsNull(rcTmp!Modelo) Or Trim(rcTmp!Modelo) = "", "", " MODELO: " & rcTmp!Modelo), i * 1, 50))
'                        Next i
'                    Else
'                    If ContratoAlmoneda = 1 Then
'                         For i = 1 To Len(rcTmp!Descripcion) Step 50
'                            strDescripcion = strDescripcion & IIf(Trim(strDescripcion) <> "", vbCrLf, "") & " " & Trim(Mid(rcTmp!Descripcion, i * 1, 50))
'                        Next i
'                    Else
'                         For i = 1 To Len(rcTmp!MarcayModelo & " " & rcTmp!Año & IIf(IsNull(rcTmp!Color) Or Trim(rcTmp!Color) = "", " ", rcTmp!Color) & IIf(IsNull(rcTmp!Placas) Or Trim(rcTmp!Placas) = "", "", " " & rcTmp!Placas) & IIf(IsNull(rcTmp!SerieChasis) Or Trim(rcTmp!SerieChasis) = "", "", " SERIE: " & rcTmp!SerieChasis) & IIf(IsNull(rcTmp!Kms) Or Trim(rcTmp!Kms) = "", "", " KMS: " & rcTmp!Kms)) Step 50
'                            strDescripcion = strDescripcion & IIf(Trim(strDescripcion) <> "", vbCrLf, "") & " " & Trim(Mid(rcTmp!MarcayModelo & " " & rcTmp!Año & IIf(IsNull(rcTmp!Color) Or Trim(rcTmp!Color) = "", " ", rcTmp!Color) & IIf(IsNull(rcTmp!Placas) Or Trim(rcTmp!Placas) = "", "", " " & rcTmp!Placas) & IIf(IsNull(rcTmp!SerieChasis) Or Trim(rcTmp!SerieChasis) = "", "", " SERIE: " & rcTmp!SerieChasis) & IIf(IsNull(rcTmp!Kms) Or Trim(rcTmp!Kms) = "", "", " KMS: " & rcTmp!Kms), i * 1, 50))
'                        Next i
'                    End If
'
'                    End If
'                rcTmp.MoveNext
'                Wend
'
'
''                If rcTmp.EOF And rcTmp.BOF Then
''                If Indice = 1 Then
''                  MsgBox "La prenda no se puede Desempeñar", vbCritical
''                Limpiar "DESEMPEÑO"
''                Screen.MousePointer = vbDefault
''                Exit Sub
''                End If
''
''                If Indice = 2 Then
''                MsgBox "La prenda no se puede Refrendar", vbCritical
''                Limpiar "REFRENDOS"
''                Screen.MousePointer = vbDefault
''                Exit Sub
''                End If
''                End If
'                 rcTmp.Close
'                Set rcTmp = Nothing
'
'                'Tomo el Préstamo
'                crPrestamo = !Prestamo
'
'                'Imprimo la descripción
'                DetallesContrato(Indice - 1).Add " " & strDescripcion
'
''                If Serie = SERIE_A Then
''
''                    'Saco los Intereses de los dias que han pasado
''                    FechaMes = Format(!Fecha, "DD/MM/YYYY")
''                    FechaMes = DateAdd(IIf(!TipoTasa = "MENSUAL", "M", "D"), IIf(!TipoTasa = "MENSUAL", 1, IIf(!TipoTasa = "QUINCENAL", 15 * 2, 7 * 4)), FechaMes)
''                    DiasMes = DateDiff("D", !Fecha, FechaMes)
''                    DiasTrans = DateDiff("D", !Fecha, Date)
''                    If DiasTrans = 0 Then DiasTrans = 1
''                    DiasMinimos = 0
''
''                    crIntereses = Redondeo(((!Prestamo * ((!Tasa / 100) / !Periodo)) * DiasTrans) - ChecaPromocion(!ID, 0, "Tasa", IIf(DiasMinimos > DiasTrans, DiasMinimos, IIf(DiasTrans > DiasMes, DiasMes, DiasTrans)), True))
''                    crAlmacenaje = Redondeo(((!Prestamo * ((!Almacenaje / 100) / !Periodo)) * DiasTrans) - ChecaPromocion(!ID, 0, "Almacenaje", IIf(DiasMinimos > DiasTrans, DiasMinimos, IIf(DiasTrans > DiasMes, DiasMes, DiasTrans)), True))
''                    crSeguro = Redondeo(((!Prestamo * ((!Seguro / 100) / !Periodo)) * DiasTrans) - ChecaPromocion(!ID, 0, "Seguro", IIf(DiasMinimos > DiasTrans, DiasMinimos, IIf(DiasTrans > DiasMes, DiasMes, DiasTrans)), True))
''                Else
''
''                    crIntereses = Redondeo(Regresa_Intereses(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoTasa))
''                    crAlmacenaje = Redondeo(Regresa_Almacenaje(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoTasa))
''                    crSeguro = Redondeo(Regresa_Seguro(!ID, !Fecha, !Prestamo, !Avaluo, !Folio, !Vencimiento, !TipoTasa))
''                    crInteresDiario = Redondeo((crIntereses + crAlmacenaje + crSeguro) / DateDiff("D", !Fecha, !Vencimiento))
''                End If
'
''                If !TipoInteres = "COMPLETO" Then
''                    crIntereses = GeneraInteresesPeriodoCompleto(!ID, "Tasa")
''                    crAlmacenaje = GeneraInteresesPeriodoCompleto(!ID, "Almacenaje")
''                    crSeguro = GeneraInteresesPeriodoCompleto(!ID, "Seguro")
''                Else
''                    crIntereses = GeneraIntereses(!ID, "Tasa")
''                    crAlmacenaje = GeneraIntereses(!ID, "Almacenaje")
''                    crSeguro = GeneraIntereses(!ID, "Seguro")
''                End If
''
''                crMoratorios = 0
''                If Date > !Vencimiento Then crMoratorios = Redondeo(!Prestamo * (Regresa_Valor_BD("Operacion") / 100))
''                crIva = Redondeo(Regresa_Iva(crIntereses + crAlmacenaje + crSeguro + crMoratorios, !ID))
''                crMinimo = Redondeo(Regresa_Valor_BD("PagoMinimo") * (1 + (Regresa_Valor_BD("IVA") / 100)))
''
'                If Indice = 1 Then
'
'                    'Cargamos los datos
'                    grdDesempeño.Redraw = False
'                    grdDesempeño.AddRow
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 1) = !NumContrato
'                    grdDesempeño.CellItemData(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 1) = !Serie
'                    grdDesempeño.CellTextAlign(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 1) = DT_CENTER Or DT_WORD_ELLIPSIS
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 2) = IIf(Serie = SERIE_B, !Prestamo, crPrestamo)
'                    grdDesempeño.CellTextAlign(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 3) = IIf((crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva) < crMinimo, crMinimo, (crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva)) + ImportePerdida
'                    grdDesempeño.CellItemData(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 3) = ImportePerdida
'                    grdDesempeño.CellTextAlign(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 4) = !Folio
'                    grdDesempeño.CellItemData(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 4) = tNumContrato.Tag
'
'                    'Checo si el pago mínimo es mayor que el importe de los intereses
'                    If (crIntereses + crAlmacenaje + crSeguro + crIva) < crMinimo Then
'                        crIntereses = Redondeo(crMinimo / (1 + (Val(Regresa_Valor_BD("IVA")) / 100)))
'                        crIva = Redondeo(crMinimo - crIntereses)
'                        crAlmacenaje = 0
'                        crSeguro = 0
'                    End If
'
'                    'Intereses
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 5) = crIntereses
'                    'Almacenaje
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 6) = crAlmacenaje
'                    'Seguro
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 7) = crSeguro
'                    'Iva
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 8) = crIva
'                    'Si es contrato de Almoneda
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 9) = ContratoAlmoneda
'                    'Importe por boleta perdida
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 10) = ImportePerdida
'                    'Moratorios
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 11) = crMoratorios
'
'                    '***Puntos***
'                    'IDCliente
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 17) = !IDCliente
'                    'IDEmpeno
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 18) = !ID
'                   'Iva Cargo Seguro y GPS
'                    grdDesempeño.CellText(IIf(grdDesempeño.Rows > 1, grdDesempeño.Rows - 1, grdDesempeño.Rows), 19) = crIvaCargoGPSSeguro
'
'                    If grdDesempeño.Rows = 1 Then grdDesempeño.AddRow
'                    grdDesempeño.Redraw = True
'                    Poner_Totales_Desempeño
'
'                ElseIf Indice = 2 Then
'
'                    grdRefrendos.Redraw = False
'                    grdRefrendos.AddRow
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 1) = !NumContrato
'                    grdRefrendos.CellItemData(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 1) = tNumContrato.Tag
'                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 1) = DT_CENTER Or DT_WORD_ELLIPSIS
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 2) = IIf(Serie = SERIE_B, !Prestamo, crPrestamo)
'                    grdRefrendos.CellItemData(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 2) = !ID
'                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 3) = "0"
'                    grdRefrendos.CellItemData(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 3) = !Serie
'                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 3) = DT_RIGHT
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 4) = IIf((crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva) < crMinimo, crMinimo, (crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva)) + ImportePerdida
'                    grdRefrendos.CellItemData(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 4) = ImportePerdida
'                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 4) = DT_RIGHT Or DT_WORD_ELLIPSIS
'
'                    'Checo si el pago mínimo es mayor que el importe de los intereses
'                    If (crIntereses + crAlmacenaje + crSeguro + crIva) < crMinimo Then
'                        crIntereses = Redondeo(crMinimo / (1 + (Val(Regresa_Valor_BD("IVA")) / 100)))
'                        crIva = Redondeo(crMinimo - crIntereses)
'                        crAlmacenaje = 0
'                        crSeguro = 0
'                    End If
'
'                    'Intereses
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 5) = crIntereses
'                    'Almacenaje
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 6) = crAlmacenaje
'                    'Seguro
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 7) = crSeguro
'                    'Iva
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 8) = crIva
'                    'Si es contrato de Almoneda
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 9) = ContratoAlmoneda
'                    'Importe Boleta Perdida
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 10) = ImportePerdida
'                    'Moratorios
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 11) = crMoratorios
'                    'Vencido
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 18) = Vencido
'                    'IDEmpeno
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 19) = !ID
'                   'Iva Cargo Seguro y GPS
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 20) = crIvaCargoGPSSeguro
'
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 13) = 0
'                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 13) = DT_RIGHT Or DT_WORD_ELLIPSIS
'
'                    grdRefrendos.CellText(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 15) = 0
'                    grdRefrendos.CellTextAlign(IIf(grdRefrendos.Rows > 1, grdRefrendos.Rows - 1, grdRefrendos.Rows), 15) = DT_RIGHT Or DT_WORD_ELLIPSIS
'
'                    If grdRefrendos.Rows = 1 Then grdRefrendos.AddRow
'                    grdRefrendos.Redraw = True
'
'                    Poner_Totales_Refrendo
'                End If
'
'            .MoveNext
'            Band = Band + 1
'            Wend
'        End If
'    End With
'    rcEmpeño.Close
'    Set rcEmpeño = Nothing
'    Screen.MousePointer = vbDefault
'    Exit Sub
'
'error:
'    Maneja_Error Err
'    Set rcEmpeño = Nothing
'    Set rcTmp = Nothing
'    Screen.MousePointer = vbDefault
'End Sub

Private Sub txtPrestamooElec_Change()
    Calcula_Avaluo_Elec
End Sub

Private Sub txtPrestamooElec_GotFocus()
    Seleccionar_Texto txtPrestamooElec
    Cambiar_Color True, txtPrestamooElec
End Sub

Private Sub txtPrestamooElec_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamooElec_LostFocus()
    txtPrestamooElec.text = Format(txtPrestamooElec.text, FMoneda)
    Cambiar_Color False, txtPrestamooElec
End Sub

Private Sub txtPrestamoo_Change()
Dim crPeso As Double, crPrecio As Double, crPrestamo As Double, Prestamo As Double, PorPrestamo As Double, PrestamoDiamante As Double, PesoPiedra As Double
Dim rcTmp As ADODB.Recordset

On Error GoTo Error
    
    If cmbTipo.ListIndex > -1 Then
        
        Set rcTmp = dbDatos.Execute("SELECT Kilataje,Peso FROM Tipo WHERE ID=" & cmbTipo.ItemData(cmbTipo.ListIndex))
        If Not rcTmp.BOF And Not rcTmp.EOF Then
        
            If rcTmp!Kilataje = 1 Or rcTmp!Peso = 1 Then
                
                If cmbKilates.ListIndex > -1 Then
                    
                    If Val(txtPeso.text) > 0 Or (Trim(txtPeso.text) <> "" And Trim(txtPeso.text) <> ".") Then
                        
                        crPeso = txtPeso.text
                    Else
                    
                        crPeso = 0
                    End If
                    
                    If Val(txtPesoPiedra.text) > 0 Or (Trim(txtPesoPiedra.text) <> "" And Trim(txtPesoPiedra.text) <> ".") Then
                
                        PesoPiedra = CDbl(txtPesoPiedra.text)
                    Else
                        
                        PesoPiedra = 0
                    End If
                    
                    If Val(lblPrestamoDiamante.Caption) > 0 Or (Trim(lblPrestamoDiamante.Caption) <> "" And Trim(lblPrestamoDiamante.Caption) <> ".") Then
                        
                        PrestamoDiamante = lblPrestamoDiamante.Caption
                    Else
                        
                        PrestamoDiamante = 0
                    End If
                    
                    If cmbTipo.ListIndex >= 0 And cmbKilates.ListIndex >= 0 And cmbEstado.ListIndex >= 0 Then
                        
                        Set rcTmp = dbDatos.Execute("SELECT Precio FROM PreciosKilataje WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND IDKilataje=" & cmbKilates.ItemData(cmbKilates.ListIndex) & " AND IDHechura=" & cmbEstado.ItemData(cmbEstado.ListIndex))
                        If Not rcTmp.BOF Then
                                crPrecio = rcTmp!Precio
                        Else
                                 MsgBox "Ingrese Precio de Kilataje !!", vbInformation, "Empeño"
                        End If
                    Else
                        
                        crPrecio = 0
                    End If
                                                    
                    PorPrestamo = SacaValor("configuraciontasas", "PorPrestamo", " WHERE IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex) & " AND IDPlazo=" & cmbPlazos.ItemData(cmbPlazos.ListIndex))
                    crPrestamo = Redondeo((crPeso - PesoPiedra) * crPrecio) * (PorPrestamo / 100)
                    PorPrestamo = Val(Regresa_Valor_BD("Negociacion")) / 100
                    
                    If Val(txtPrestamoo.text) > 0 Or Trim(txtPrestamoo.text) <> "" And crPeso > 0 Then
                        
                        Prestamo = txtPrestamoo.text
                        
                        If Prestamo > Redondeo((crPrestamo + (crPrestamo * PorPrestamo)) + PrestamoDiamante) / 100 * ImgSemaforo.Tag Then
                            MsgBox "El préstamo sobrepasa el margen de negociación !!", vbInformation, "Empeño"
                            Calcular_Avaluo
                        End If
                        
                    End If
                    
                Else
                    
                    txtPrestamoo.text = "0.00"
                End If
            
            Else
                
                If Val(txtPrestamoo.text) > 0 Then
                    
                    crPrestamo = txtPrestamoo.text
                Else
                    
                    crPrestamo = 0
                End If
                
                txtAvaluo.text = Format(Calcular_Avaluo(), FMoneda)
            End If
            
        End If
        
    End If
    Set rcTmp = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Private Sub txtPrestamoo_GotFocus()
    Seleccionar_Texto txtPrestamoo
    Cambiar_Color True, txtPrestamoo
End Sub

Private Sub txtPrestamoo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamoo_LostFocus()
    Cambiar_Color False, txtPrestamoo
    txtPrestamoo.text = Format(txtPrestamoo.text, FMoneda)
End Sub


Private Sub txtTamañoElec_GotFocus()
    Seleccionar_Texto txtTamañoElec
    Cambiar_Color True, txtTamañoElec
End Sub

Private Sub txtTamañoElec_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTamañoElec_LostFocus()
    Cambiar_Color False, txtTamañoElec
End Sub


Private Sub Poner_Totales_Desempeño()
    
    Dim Prestamo As Double, Interes As Double
    Dim Renglon As Integer, columna As Integer
    
    Dim crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crIvaCargo As Double
    Dim crMoratorios As Double, crCargoGPS As Double, crCargoSeguroAuto As Double, crIva As Double, crImportePerdida As Double
   
    
    '***Puntos***
    Dim Puntos As Long
    Dim Total As Currency
    Dim ImportePuntos As Currency
    Dim TotalPagar As Currency

    On Error GoTo Error

    'Hago la sumatoria de los totales (Cargos, Abonos, Saldos) desde el renglon 1 hasta el numero de renglones del GRID
    For Renglon = 1 To grdDesempeño.Rows - 1
    
    '* Se agrega para GPS
    
    'Se calcula de nuevo el interes
            crIntereses = CDbl(grdDesempeño.CellText(Renglon, 5))
            crAlmacenaje = CDbl(grdDesempeño.CellText(Renglon, 6))
            crSeguro = CDbl(grdDesempeño.CellText(Renglon, 7))
            crMoratorios = CDbl(grdDesempeño.CellText(Renglon, 11))
            crImportePerdida = CDbl(grdDesempeño.CellText(Renglon, 10))
         If Frame4.Visible = True Then
            If chkCirculacionDes.Value = 1 Then
            'si el carro esta en circulacion
            crCargoSeguroAuto = CDbl(IIf(txtCargoSeguroDes.text = "", 0, txtCargoSeguroDes.text))
            crCargoGPS = CDbl(IIf(txtCargoGPSDes.text = "", 0, txtCargoGPSDes.text))
            End If
         
         End If
        
        crIva = Regresa_Iva(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crCargoGPS + crCargoSeguroAuto, CLng(grdDesempeño.CellText(Renglon, 18)))
        crIvaCargo = Redondeo(Regresa_Iva(crCargoGPS + crCargoSeguroAuto, CLng(grdDesempeño.CellText(Renglon, 18))))
        
        grdDesempeño.CellText(Renglon, 8) = crIva
        grdDesempeño.CellText(Renglon, 19) = crIvaCargo
        grdDesempeño.CellText(Renglon, 3) = Redondeo(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva + crImportePerdida)
        
        '***********
    
    
    
        '***Puntos***
        grdDesempeño.CellDetails Renglon, 14, CDbl(grdDesempeño.CellText(Renglon, 2)) + CDbl(grdDesempeño.CellText(Renglon, 3)), DT_RIGHT Or DT_WORD_ELLIPSIS
        grdDesempeño.CellDetails Renglon, 16, CDbl(grdDesempeño.CellText(Renglon, 14)) - CDbl(grdDesempeño.CellText(Renglon, 15)), DT_RIGHT Or DT_WORD_ELLIPSIS
    
        Puntos = Puntos + CDbl(grdDesempeño.CellText(Renglon, 13))
        Total = Total + CDbl(grdDesempeño.CellText(Renglon, 14))
        ImportePuntos = ImportePuntos + CDbl(grdDesempeño.CellText(Renglon, 15))
        TotalPagar = Redondeo(TotalPagar + CDbl(grdDesempeño.CellText(Renglon, 16)))
        
        Prestamo = Prestamo + IIf(Val(grdDesempeño.CellText(Renglon, 2)) = 0, 0, grdDesempeño.CellText(Renglon, 2))
        Interes = Interes + IIf(Val(grdDesempeño.CellText(Renglon, 3)) = 0, 0, grdDesempeño.CellText(Renglon, 3))
    Next Renglon
                  
    If Frame4.Visible = True Then
        If chkCirculacionDes.Value = 1 Then
        
            TotalPagar = TotalPagar + CDbl(IIf(txtCargoSeguroDes.text = "", 0, txtCargoSeguroDes.text)) + CDbl(IIf(txtCargoGPSDes.text = "", 0, txtCargoGPSDes.text))
            
        End If
        
    End If
    'En la ultima linea del GRID cargo los totales (Cargos, Abonos, Saldos) y cambio el color de la linea
    grdDesempeño.CellText(grdDesempeño.Rows, 2) = Format(Prestamo, "Currency")
    grdDesempeño.CellTextAlign(grdDesempeño.Rows, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdDesempeño.CellText(grdDesempeño.Rows, 3) = Format(Interes, "Currency")
    grdDesempeño.CellTextAlign(grdDesempeño.Rows, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
        
    Leyenda.Caption = "TOTAL A PAGAR:"
    Leyenda.Tag = "0"
    'TotalDesempeño.Caption = Format(Redondeo(Prestamo + Interes), FMoneda)
    
    '***Puntos***
    TotalDesempeño.Caption = Format(TotalPagar, FMoneda)
    
    For Renglon = 1 To grdDesempeño.Rows
        
        For columna = 1 To grdDesempeño.Columns - 1
            grdDesempeño.CellBackColor(Renglon, columna) = &HFFFFFF
            grdDesempeño.CellForeColor(Renglon, columna) = &H0&
        Next columna
    
    Next Renglon

    For columna = 1 To grdDesempeño.Columns - 1
        
        grdDesempeño.CellBackColor(grdDesempeño.Rows, columna) = RGB(223, 208, 102)
        grdDesempeño.CellForeColor(grdDesempeño.Rows, columna) = &HFF0000
    
    Next columna
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub Poner_Totales_Refrendo()

    Dim Prestamo As Double, Interes As Double, Abono As Double
    Dim Renglon As Double, columna As Integer
     
    Dim crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crIvaCargo As Double
    Dim crMoratorios As Double, crCargoGPS As Double, crCargoSeguroAuto As Double, crIva As Double, crImportePerdida As Double
    
    '***Puntos***
    Dim Puntos As Long
    Dim Total As Currency
    Dim ImportePuntos As Currency
    Dim TotalPagar As Currency

On Error GoTo Error

    'Hago la sumatoria de los totales (Cargos, Abonos, Saldos) desde el renglon 1 hasta el numero de renglones del GRID
    For Renglon = 1 To grdRefrendos.Rows - 1
        
         'Se calcula de nuevo el interes
            crIntereses = CDbl(grdRefrendos.CellText(Renglon, 5))
            crAlmacenaje = CDbl(grdRefrendos.CellText(Renglon, 6))
            crSeguro = CDbl(grdRefrendos.CellText(Renglon, 7))
            crMoratorios = CDbl(grdRefrendos.CellText(Renglon, 11))
            crImportePerdida = CDbl(grdRefrendos.CellText(Renglon, 10))
         If Frame3.Visible = True Then
            If chkCirculacionRef.Value = 1 Then
            'si el carro esta en circulacion
            crCargoSeguroAuto = CDbl(IIf(txtCargoSeguro.text = "", 0, txtCargoSeguro.text))
            crCargoGPS = CDbl(IIf(txtCargoGPS.text = "", 0, txtCargoGPS.text))
            End If
         
         End If
        
        crIva = Regresa_Iva(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crCargoGPS + crCargoSeguroAuto, CLng(grdRefrendos.CellText(Renglon, 19)))
        crIvaCargo = Redondeo(Regresa_Iva(crCargoGPS + crCargoSeguroAuto, CLng(grdRefrendos.CellText(Renglon, 19))))
        
        grdRefrendos.CellText(Renglon, 8) = crIva
        grdRefrendos.CellText(Renglon, 20) = crIvaCargo
        grdRefrendos.CellText(Renglon, 4) = Redondeo(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva + crImportePerdida)
        
        '***Puntos***
        grdRefrendos.CellDetails Renglon, 14, CDbl(grdRefrendos.CellText(Renglon, 3)) + CDbl(grdRefrendos.CellText(Renglon, 4)), DT_RIGHT Or DT_WORD_ELLIPSIS
        grdRefrendos.CellDetails Renglon, 16, CDbl(grdRefrendos.CellText(Renglon, 14)) - CDbl(grdRefrendos.CellText(Renglon, 15)), DT_RIGHT Or DT_WORD_ELLIPSIS
         
        
        
         
        Puntos = Puntos + CDbl(grdRefrendos.CellText(Renglon, 12))
        Total = Total + CDbl(grdRefrendos.CellText(Renglon, 13))
        ImportePuntos = ImportePuntos + CDbl(grdRefrendos.CellText(Renglon, 15))
        TotalPagar = Redondeo(TotalPagar + CDbl(grdRefrendos.CellText(Renglon, 16)))
        
        Abono = Abono + IIf(Val(grdRefrendos.CellText(Renglon, 3)) = 0 Or Trim(grdRefrendos.CellText(Renglon, 3)) = "", 0, grdRefrendos.CellText(Renglon, 3))
        Prestamo = Prestamo + CDbl(grdRefrendos.CellText(Renglon, 2))
        Interes = Interes + CDbl(grdRefrendos.CellText(Renglon, 4))
    Next Renglon
    If Frame3.Visible = True Then
    If chkCirculacionRef.Value = 1 Then
    
        TotalPagar = TotalPagar + CDbl(IIf(txtCargoSeguro.text = "", 0, txtCargoSeguro.text)) + CDbl(IIf(txtCargoGPS.text = "", 0, txtCargoGPS.text))
        
    End If
        
    End If
         
    'En la ultima linea del GRID cargo los totales (Cargos, Abonos, Saldos) y cambio el color de la linea
    grdRefrendos.CellText(grdRefrendos.Rows, 2) = Prestamo
    grdRefrendos.CellTextAlign(grdRefrendos.Rows, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellText(grdRefrendos.Rows, 3) = Abono
    grdRefrendos.CellTextAlign(grdRefrendos.Rows, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellText(grdRefrendos.Rows, 4) = Interes
    grdRefrendos.CellTextAlign(grdRefrendos.Rows, 4) = DT_RIGHT Or DT_WORD_ELLIPSIS

    '***Puntos***
    grdRefrendos.CellDetails grdRefrendos.Rows, 13, Puntos, DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellDetails grdRefrendos.Rows, 14, Total, DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellDetails grdRefrendos.Rows, 15, ImportePuntos, DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellDetails grdRefrendos.Rows, 16, TotalPagar, DT_RIGHT Or DT_WORD_ELLIPSIS

    LeyendaRef.Caption = "TOTAL A PAGAR:"
    LeyendaRef.Tag = "0"
    'TotalRefrendo.Caption = Format(Redondeo(Interes + Abono), FMoneda)
    
    '***Puntos***
    TotalRefrendo.Caption = Format(TotalPagar, FMoneda)
    
    For columna = 1 To grdRefrendos.Columns
        grdRefrendos.CellBackColor(grdRefrendos.Rows - 1, columna) = RGB(255, 255, 255)
        grdRefrendos.CellForeColor(grdRefrendos.Rows - 1, columna) = RGB(0, 0, 0)
    Next columna
    
    For columna = 1 To grdRefrendos.Columns
        grdRefrendos.CellBackColor(grdRefrendos.Rows, columna) = RGB(223, 208, 102)
        grdRefrendos.CellForeColor(grdRefrendos.Rows, columna) = &HFF0000
    Next columna
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub grdRefrendos_CancelEdit()
    txtEdit.Visible = False
End Sub

Private Sub grdRefrendos_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, sText As String
   
    grdRefrendos.CancelEdit
       
    '***Puntos***
    If (lCol <> 3 And lCol <> 13) Or lRow = grdRefrendos.Rows Then
        grdRefrendos.CancelEdit
        Exit Sub
    End If
    
'''''    If lCol = 4 Then
'''''        frmPasswords.ConexSuc = 0
'''''        frmPasswords.DescuentoVentas = 0
'''''        frmPasswords.Cancel = 0
'''''        frmPasswords.Ventas = 0
'''''        frmPasswords.ModificaCorte = 0
'''''        frmPasswords.HacerCorte = 0
'''''        frmPasswords.InteresDesempeño = 0
'''''        frmPasswords.ModificaPrecio = 0
'''''        frmPasswords.RecalculoPrecios = 0
'''''        frmPasswords.InteresDesempeño = 0
'''''        frmPasswords.AutorizaPrestamo = 0
'''''        frmPasswords.InteresRefrendo = 1
'''''        If frmPasswords.Password(GERENTE, 1) Then GoTo 125 Else Exit Sub
'''''    End If
    
125:
    grdRefrendos.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight

    If Not IsMissing(grdRefrendos.CellText(lRow, lCol)) Then
        sText = grdRefrendos.CellFormattedText(lRow, lCol)
    Else
        sText = ""
    End If
   
    iKeyAscii = Solo_Numeros(iKeyAscii)
    If (iKeyAscii > 13) Then
        sText = Chr$(iKeyAscii) & sText
        txtEdit.text = sText
        txtEdit.SelStart = 1
        txtEdit.SelLength = Len(sText)
    Else
        txtEdit.text = sText
        txtEdit.SelStart = 0
        txtEdit.SelLength = Len(sText)
    End If
   
    Set txtEdit.Font = grdRefrendos.CellFont(lRow, lCol)
    If grdRefrendos.CellBackColor(lRow, lCol) = -1 Then
        txtEdit.BackColor = grdRefrendos.BackColor
    Else
        txtEdit.BackColor = grdRefrendos.CellBackColor(lRow, lCol)
    End If
    txtEdit.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
    txtEdit.Visible = True
    txtEdit.ZOrder
    txtEdit.SetFocus
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim crNuevoImporte As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crIva As Double, crMoratorios As Double, crImportePerdida As Double
    Dim crAbono As Double, crPrestamo As Double, crDiferencia As Double
    Dim TotalAPagar As Double

    'Condicion para grabar los abonos (tecla ENTER y FOCUS en la columna 3)
    If (KeyCode = vbKeyReturn) Then
    
        If grdRefrendos.SelectedCol = 3 Then
        
           If Val(txtEdit.text) > 0 Or Trim(txtEdit.text) <> "" Then
                crAbono = CDbl(txtEdit.text)
                crPrestamo = CDbl(grdRefrendos.CellText(grdRefrendos.SelectedRow, 2))
                
                If crAbono > crPrestamo Then
                    MsgBox "El importe del abono no puede ser mayor al préstamo !!", vbCritical, "Refrendo"
                    txtEdit.text = 0
                    crAbono = 0
                End If
                
                'Cambiamos los valores de la celda de acuerdo a los datos ingresados
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 3) = Format(crAbono, "###,###,###,###0.00")
                grdRefrendos.CellTextAlign(grdRefrendos.SelectedRow, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
                grdRefrendos.CancelEdit
                txtEdit.Visible = False
                Poner_Totales_Refrendo
                grdRefrendos.SetFocus
            Else
                txtEdit.text = 0
                txtEdit.Visible = False
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 3) = Format(txtEdit.text, "##,###0.00")
                grdRefrendos.CellTextAlign(grdRefrendos.SelectedRow, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
                Poner_Totales_Refrendo
                grdRefrendos.SetFocus
            End If
            
        ElseIf grdRefrendos.SelectedCol = 4 Then
            
            If Val(txtEdit.text) > 0 Then
                                
                crNuevoImporte = txtEdit.text
                
                crIntereses = grdRefrendos.CellText(grdRefrendos.SelectedRow, 5)
                crAlmacenaje = grdRefrendos.CellText(grdRefrendos.SelectedRow, 6)
                crSeguro = grdRefrendos.CellText(grdRefrendos.SelectedRow, 7)
                crImportePerdida = grdRefrendos.CellText(grdRefrendos.SelectedRow, 10)
                crMoratorios = grdRefrendos.CellText(grdRefrendos.SelectedRow, 11)
                crIva = grdRefrendos.CellText(grdRefrendos.SelectedRow, 8)
                
                crDiferencia = (crIntereses + crAlmacenaje + crSeguro + crImportePerdida + crMoratorios + crIva) - crNuevoImporte
                
                'Intereses
                If crDiferencia >= crIntereses Then
                    
                    crDiferencia = Redondeo(crDiferencia - crIntereses)
                    crIntereses = 0
                Else
                    
                    crIntereses = Redondeo(crIntereses - crDiferencia)
                    crDiferencia = 0
                End If
                
                'Almacenaje
                If crDiferencia >= crAlmacenaje Then
                    
                    crDiferencia = Redondeo(crDiferencia - crAlmacenaje)
                    crAlmacenaje = 0
                Else
                    
                    crAlmacenaje = Redondeo(crAlmacenaje - crDiferencia)
                    crDiferencia = 0
                End If
                
                'Seguro
                If crDiferencia >= crSeguro Then
                    
                    crDiferencia = Redondeo(crDiferencia - crSeguro)
                    crSeguro = 0
                Else
                    
                    crSeguro = Redondeo(crSeguro - crDiferencia)
                    crDiferencia = 0
                End If
                
                'Moratorios
                If crDiferencia >= crMoratorios Then
                    
                    crDiferencia = Redondeo(crDiferencia - crMoratorios)
                    crMoratorios = 0
                Else
                    
                    crMoratorios = Redondeo(crMoratorios - crDiferencia)
                    crDiferencia = 0
                End If
                
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 5) = crIntereses
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 6) = crAlmacenaje
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 7) = crSeguro
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 10) = crImportePerdida
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 11) = crMoratorios
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 8) = crIva
                
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 4) = Format(crNuevoImporte, FMoneda)
                grdRefrendos.CellTextAlign(grdRefrendos.SelectedRow, 4) = DT_RIGHT Or DT_WORD_ELLIPSIS
                            
                grdRefrendos.CancelEdit
                txtEdit.Visible = False
                Poner_Totales_Refrendo
                grdRefrendos.ClearSelection
                
            Else
                
                grdRefrendos.CancelEdit
                txtEdit.text = 0
                txtEdit.Visible = False
                grdRefrendos.CellText(grdRefrendos.SelectedRow, 4) = Format(0, FMoneda)
                grdRefrendos.CellTextAlign(grdRefrendos.SelectedRow, 4) = DT_RIGHT Or DT_WORD_ELLIPSIS
                Poner_Totales_Refrendo
                grdRefrendos.ClearSelection
                
            End If
            
        '***Puntos***
        ElseIf grdRefrendos.SelectedCol = 13 Then
        
            If Val(txtEdit.text) <= TarjetaPuntos.CuentaFrecuente.Puntos Then
               
               TotalAPagar = CDbl(grdRefrendos.CellText(grdRefrendos.SelectedRow, 4)) '14
                
                If TarjetaPuntos.GetImporte(Val(txtEdit.text)) < TotalAPagar Then
                    grdRefrendos.CellDetails grdRefrendos.SelectedRow, 13, txtEdit.text, DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdRefrendos.CellDetails grdRefrendos.SelectedRow, 15, TarjetaPuntos.GetImporte(Val(txtEdit.text)), DT_RIGHT Or DT_WORD_ELLIPSIS
                    
                Else
                    grdRefrendos.CellDetails grdRefrendos.SelectedRow, 13, "0", DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdRefrendos.CellDetails grdRefrendos.SelectedRow, 15, TarjetaPuntos.GetImporte(Val("0")), DT_RIGHT Or DT_WORD_ELLIPSIS
                
                    MsgBox "Los puntos a utilizar es mayor a los intereses", vbOKOnly Or vbCritical
                End If
               
               
               grdRefrendos.CancelEdit
               txtEdit.text = ""
               txtEdit.Visible = False
               Poner_Totales_Refrendo
            Else
               MsgBox "Los puntos a utilizar no pueden ser mayor al saldo de puntos", vbOKOnly Or vbCritical
               txtEdit.text = ""
               grdRefrendos.CancelEdit
               txtEdit.Visible = False
            End If
            
        End If
        
    ElseIf (KeyCode = vbKeyEscape) Then
    
            txtEdit.Visible = False
            grdRefrendos.CancelEdit
            Poner_Totales_Refrendo
            grdRefrendos.ClearSelection
    
    End If
    KeyCode = 0
End Sub

Private Sub Quitar_Renglon()

On Error GoTo Error

    grdRefrendos.RemoveRow (grdRefrendos.SelectedRow)
    If grdRefrendos.Rows = 1 Then
        
        grdRefrendos.Clear
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub Limpiar_Leyendas()
Dim ctrl As Control
    
    Leyenda.Caption = ""
    LeyendaRef.Caption = ""
    TotalDesempeño.Caption = ""
    TotalRefrendo.Caption = ""
    TotalDesempeño.ForeColor = &HFF0000
    TotalRefrendo.ForeColor = &HFF0000
            
    For Each ctrl In Controls
        
        If TypeOf ctrl Is vbalGrid Then ctrl.Clear
    Next ctrl

End Sub

Private Sub grddesempeño_CancelEdit()
    txtEdit2.Visible = False
End Sub

Private Sub grddesempeño_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)

Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, sText As String

   If (lCol <> 13) Or lRow = grdDesempeño.Rows Then txtEdit2.Visible = False: Exit Sub

   grdDesempeño.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight

   If Not IsMissing(grdDesempeño.CellText(lRow, lCol)) Then
      sText = grdDesempeño.CellFormattedText(lRow, lCol)
   Else
      sText = ""
   End If

   iKeyAscii = Solo_Numeros(iKeyAscii)
   If (iKeyAscii > 13) Then
      sText = Chr$(iKeyAscii) & sText
      txtEdit2.text = sText
      txtEdit2.SelStart = 1
      txtEdit2.SelLength = Len(sText)
   Else
      txtEdit2.text = sText
      txtEdit2.SelStart = 0
      txtEdit2.SelLength = Len(sText)
   End If
   Set txtEdit2.Font = grdDesempeño.CellFont(lRow, lCol)
   If grdDesempeño.CellBackColor(lRow, lCol) = -1 Then
      txtEdit2.BackColor = grdDesempeño.BackColor
   Else
      txtEdit2.BackColor = grdDesempeño.CellBackColor(lRow, lCol)
   End If
   txtEdit2.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
   txtEdit2.Visible = True
   txtEdit2.ZOrder
   txtEdit2.SetFocus

End Sub

Private Sub txtEdit2_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crImportePerdida As Double, crMoratorios As Double, crIva As Double, crNuevoImporte As Double, crDiferencia As Double
    Dim TotalAPagar As Double

    'Condicion para grabar los abonos (tecla ENTER y FOCUS en la columna 3)
    If (KeyCode = vbKeyReturn) Then
        
        If grdDesempeño.SelectedCol = 3 Then
                    
            If Val(txtEdit2.text) > 0 And Trim(txtEdit2.text) <> "" Then
                    
                crNuevoImporte = txtEdit2.text
                
                crIntereses = grdDesempeño.CellText(grdDesempeño.SelectedRow, 5)
                crAlmacenaje = grdDesempeño.CellText(grdDesempeño.SelectedRow, 6)
                crSeguro = grdDesempeño.CellText(grdDesempeño.SelectedRow, 7)
                crImportePerdida = grdDesempeño.CellText(grdDesempeño.SelectedRow, 10)
                crMoratorios = grdDesempeño.CellText(grdDesempeño.SelectedRow, 11)
                crIva = grdDesempeño.CellText(grdDesempeño.SelectedRow, 8)
                
                crDiferencia = (crIntereses + crAlmacenaje + crSeguro + crImportePerdida + crMoratorios + crIva) - crNuevoImporte
                
                'Intereses
                If crDiferencia >= crIntereses Then
                    
                    crDiferencia = Redondeo(crDiferencia - crIntereses)
                    crIntereses = 0
                Else
                    
                    crIntereses = Redondeo(crIntereses - crDiferencia)
                    crDiferencia = 0
                End If
                
                'Almacenaje
                If crDiferencia >= crAlmacenaje Then
                    
                    crDiferencia = Redondeo(crDiferencia - crAlmacenaje)
                    crAlmacenaje = 0
                Else
                    
                    crAlmacenaje = Redondeo(crAlmacenaje - crDiferencia)
                    crDiferencia = 0
                End If
                
                'Seguro
                If crDiferencia >= crSeguro Then
                    
                    crDiferencia = Redondeo(crDiferencia - crSeguro)
                    crSeguro = 0
                Else
                    
                    crSeguro = Redondeo(crSeguro - crDiferencia)
                    crDiferencia = 0
                End If
                
                'Moratorios
                If crDiferencia >= crMoratorios Then
                    
                    crDiferencia = Redondeo(crDiferencia - crMoratorios)
                    crMoratorios = 0
                Else
                    
                    crMoratorios = Redondeo(crMoratorios - crDiferencia)
                    crDiferencia = 0
                End If
                
                grdDesempeño.CellText(grdDesempeño.SelectedRow, 5) = crIntereses
                grdDesempeño.CellText(grdDesempeño.SelectedRow, 6) = crAlmacenaje
                grdDesempeño.CellText(grdDesempeño.SelectedRow, 7) = crSeguro
                grdDesempeño.CellText(grdDesempeño.SelectedRow, 10) = crImportePerdida
                grdDesempeño.CellText(grdDesempeño.SelectedRow, 11) = crMoratorios
                grdDesempeño.CellText(grdDesempeño.SelectedRow, 8) = crIva
                
                grdDesempeño.CellText(grdDesempeño.SelectedRow, 3) = Format(crNuevoImporte, FMoneda)
                grdDesempeño.CellTextAlign(grdDesempeño.SelectedRow, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
                            
                grdDesempeño.CancelEdit
                txtEdit2.Visible = False
                Poner_Totales_Desempeño
                grdDesempeño.ClearSelection
                
            Else
                
                grdDesempeño.CancelEdit
                txtEdit2.text = 0
                txtEdit2.Visible = False
                Poner_Totales_Refrendo
                grdDesempeño.ClearSelection
                
            End If
            
        '***Puntos***
        ElseIf grdDesempeño.SelectedCol = 13 Then
            
            If Val(txtEdit2.text) <= TarjetaPuntos.CuentaFrecuente.Puntos Then
            
                TotalAPagar = CDbl(grdDesempeño.CellText(grdDesempeño.SelectedRow, 4)) '14
                
                If TarjetaPuntos.GetImporte(Val(txtEdit2.text)) < TotalAPagar Then
                    grdDesempeño.CellDetails grdDesempeño.SelectedRow, 13, txtEdit2.text, DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdDesempeño.CellDetails grdDesempeño.SelectedRow, 15, TarjetaPuntos.GetImporte(Val(txtEdit2.text)), DT_RIGHT Or DT_WORD_ELLIPSIS
                Else
                    grdDesempeño.CellDetails grdDesempeño.SelectedRow, 13, "0", DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdDesempeño.CellDetails grdDesempeño.SelectedRow, 15, TarjetaPuntos.GetImporte(Val("0")), DT_RIGHT Or DT_WORD_ELLIPSIS
                    MsgBox "Los puntos a utilizar es mayor a los intereses", vbOKOnly Or vbCritical
                End If
               
                grdDesempeño.CancelEdit
                txtEdit2.text = ""
                txtEdit2.Visible = False
                Poner_Totales_Desempeño
            Else
               MsgBox "Los puntos a utilizar no pueden ser mayor al saldo de puntos", vbOKOnly Or vbCritical
               txtEdit2.text = ""
               grdDesempeño.CancelEdit
               txtEdit2.Visible = False
            End If
            
        End If
        
    ElseIf (KeyCode = vbKeyEscape) Then
    
        txtEdit2.Visible = False
        grdDesempeño.CancelEdit
        Poner_Totales_Desempeño
        grdDesempeño.SetFocus
        
    End If
    KeyCode = 0
End Sub

''Grabamos los datos del cliente
'Private Function Grabar_Cliente() As Long
'Dim Medio As Integer, FechaNac As String, Sexo As Integer
'
'On Error GoTo Error
'
'    If cmbMedio.ListIndex = -1 Then
'
'        Medio = 0
'    Else
'
'        Medio = cmbMedio.ItemData(cmbMedio.ListIndex)
'    End If
'
'    If Trim(txtFecNacimiento.text) = "" Or Trim(txtFecNacimiento.text) = "__/__/____" Then
'
'        FechaNac = "Null"
'    Else
'
'        FechaNac = "'" & Format(txtFecNacimiento.text, "YYYY/MM/DD") & "'"
'    End If
'
'    If cmbSexo.ListIndex = -1 Then
'
'        Sexo = 0
'    Else
'
'        Sexo = cmbSexo.ItemData(cmbSexo.ListIndex)
'    End If
'
'    dbDatos.Execute "INSERT INTO clientes (Nombre,Apellido,Iniciales,Direccion,Colonia,Municipio,Estado,Tel,Identificacion,IDMedio,Notas,CP,FecNac,Sexo,FecRegistro,NumeroIdentificacion,Celular,CorreoElectronico) VALUES ('" & _
'        Trim(txtNombre.text) & "','" & Trim(txtApellidos.text) & "','" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "','" & Trim(txtDireccion.text) & "','" & Trim(txtColonia.text) & "','" & Trim(txtMunicipio.text) & "','" & Trim(txtEstado.text) & "','" & Trim(txtTelefono.text) & "','" & Trim(txtIdentificacion.text) & "'," & Medio & ",'" & Trim(txtMensaje.text) & "','" & txtCP.text & "'," & FechaNac & "," & Sexo & ",'" & Format(Date, "YYYY/MM/DD") & "','" & Trim(txtIdentificacionNumero.text) & "','" & Trim(txtCelular.text) & "','" & Trim(txtCorreoElectronico.text) & "')"
'
'    Grabar_Cliente = SacaValor("clientes", "MAX(ID)")
'    Exit Function
'
'Error:
'    Maneja_Error Err
'End Function

''actualizamos los datos del cliente
'Private Function Actualizar_Cliente(ID As Long, strNombre As String, strApellidos As String) As Long
'Dim Medio As Integer, FechaNac As String, Sexo As Integer
'
'On Error GoTo Error
'
'    Actualizar_Cliente = ID
'
'    If (strNombre & " " & strApellidos) <> (Trim(txtNombre.text) & " " & Trim(txtApellidos.text)) Then
'
'        If MsgBox("La información del cliente " & strNombre & " " & strApellidos & Chr(13) & "será remplazada por la de " & Trim(txtNombre.text) & " " & Trim(txtApellidos.text) & " desea continuar ??" & Chr(13) & "Si selecciona Si se remplazará la información, en caso contrario se registrará como un nuevo cliente.", vbQuestion + vbYesNo + vbDefaultButton2, "Empeño") = vbNo Then
'
'            Actualizar_Cliente = Grabar_Cliente
'            Exit Function
'        End If
'
'    End If
'
'    If cmbMedio.ListIndex = -1 Then
'
'        Medio = 0
'    Else
'
'        Medio = cmbMedio.ItemData(cmbMedio.ListIndex)
'    End If
'
'    If Trim(txtFecNacimiento.text) = "" Or Trim(txtFecNacimiento.text) = "__/__/____" Then
'
'        FechaNac = "Null"
'    Else
'
'        FechaNac = "'" & Format(txtFecNacimiento.text, "YYYY/MM/DD") & "'"
'    End If
'
'    If cmbSexo.ListIndex = -1 Then
'
'        Sexo = 0
'    Else
'
'        Sexo = cmbSexo.ItemData(cmbSexo.ListIndex)
'    End If
'
'    dbDatos.Execute "UPDATE clientes SET nombre='" & Trim(txtNombre.text) & "',apellido='" & Trim(txtApellidos.text) & "',Iniciales='" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "',Direccion='" & Trim(txtDireccion.text) & "',Colonia='" & Trim(txtColonia.text) & "',Municipio='" & Trim(txtMunicipio.text) & "'," & _
'                    "Estado='" & Trim(txtEstado.text) & "',Tel='" & Trim(txtTelefono.text) & "',Identificacion='" & Trim(txtIdentificacion.text) & "',IDMedio=" & Medio & ",Notas='" & Trim(txtMensaje.text) & "',CP='" & txtCP.text & "',FecNac=" & FechaNac & ",Sexo=" & Sexo & ",NumeroIdentificacion='" & Trim(txtIdentificacionNumero.text) & "'" & _
'                    ",Celular='" & Trim(txtCelular.text) & "',CorreoElectronico='" & Trim(txtCorreoElectronico.text) & "' WHERE ID = " & ID
'    Exit Function
'
'Error:
'    Maneja_Error Err
'End Function

''Buscamos el id cliente
''***Puntos***
'Public Sub Buscar_Cliente(ID As Long, Optional Tarjeta As Boolean = False)
'
'    Dim rcClientes As New ADODB.Recordset
'    Dim crPrestamo As Double
'    Dim semaforo As String
'
'On Error GoTo Error
'
'    rcClientes.Open "SELECT * FROM clientes WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
'
'    Select Case TPestañas.SelectedTab
'
'    Case 1
'
'        With rcClientes
'            txtNombre.text = !Nombre
'            txtNombre.Tag = ID
'            txtApellidos.text = !Apellido
'            txtApellidos.Tag = !Apellido
'            txtDireccion.text = IIf(IsNull(!Direccion), "", !Direccion)
'            txtDireccion.Tag = !Nombre
'            txtColonia.text = IIf(IsNull(!Colonia), "", !Colonia)
'            txtMunicipio.text = IIf(IsNull(!Municipio), "", !Municipio)
'            txtEstado.text = IIf(IsNull(!Estado), "", !Estado)
'            txtTelefono.text = IIf(IsNull(!Tel), "", !Tel)
'            txtCelular.text = IIf(IsNull(!Celular), "", !Celular)
'            txtCorreoElectronico.text = IIf(IsNull(!Correoelectronico), "", !Correoelectronico)
'            txtIdentificacion.text = IIf(IsNull(!Identificacion), "", !Identificacion)
'            txtCP.text = IIf(IsNull(!CP), "", !CP)
'            txtMensaje.text = IIf(IsNull(!Notas), "", !Notas)
'            txtIdentificacionNumero.text = IIf(IsNull(!NumeroIdentificacion), "", !NumeroIdentificacion)
'            If Not IsNull(!FecNac) Then
'                txtFecNacimiento.text = !FecNac
'                txtEdad.text = Calcula_Edad(!FecNac)
'            Else
'                txtEdad.text = ""
'                txtFecNacimiento.Mask = ""
'                txtFecNacimiento.text = ""
'                txtFecNacimiento.Mask = "##/##/####"
'            End If
'            cmbSexo.ListIndex = ComboInformacion(cmbSexo, IIf(IsNull(!Sexo), -1, !Sexo))
'            cmbMedio.ListIndex = IIf(IsNull(!IDMedio) Or !IDMedio = 0, -1, ComboInformacion(cmbMedio, !IDMedio))
'            MuestraTasa cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), CCur(txtPrestamo.Caption), True, lblTasa, False
'        End With
'
'    Case 2
'
'        With rcClientes
'            txtNombre2.text = !Nombre
'            txtNombre2.Tag = ID
'            txtApellidos2.text = !Apellido
'            txtDireccion2.text = IIf(IsNull(!Direccion), "", !Direccion)
'            txtColonia2.text = IIf(IsNull(!Colonia), "", !Colonia)
'            txtMunicipio2.text = IIf(IsNull(!Municipio), "", !Municipio)
'            txtEstado2.text = IIf(IsNull(!Estado), "", !Estado)
'            txtTelefono2.text = IIf(IsNull(!Tel), "", !Tel)
'            txtCelular2.text = IIf(IsNull(!Celular), "", !Celular)
'            txtCorreoElectronico2.text = IIf(IsNull(!Correoelectronico), "", !Correoelectronico)
'            txtIdentificacion2.text = IIf(IsNull(!Identificacion), "", !Identificacion)
'            txtCp2.text = IIf(IsNull(!CP), "", !CP)
'            txtMensaje2.text = IIf(IsNull(!Notas), "", !Notas)
'            If Not IsNull(!FecNac) Then
'                txtFecNacimiento2.text = !FecNac
'                txtEdad3.text = Calcula_Edad(!FecNac)
'            Else
'                txtEdad3.text = ""
'                txtFecNacimiento2.Mask = ""
'                txtFecNacimiento2.text = ""
'                txtFecNacimiento2.Mask = "##/##/####"
'            End If
'
'            If Val(txtPrestamo2.text) > 0 Or Trim(txtPrestamo2.text) <> "" Then
'
'                crPrestamo = CDbl(txtPrestamo2.text)
'            Else
'
'                crPrestamo = 0
'            End If
'
'            cmbSexo2.ListIndex = ComboInformacion(cmbSexo2, IIf(IsNull(!Sexo), -1, !Sexo))
'            MuestraTasa cmbTipoInteres2.ItemData(cmbTipoInteres2.ListIndex), cmbPeriodo2.ItemData(cmbPeriodo2.ListIndex), cmbPlazos2.ItemData(cmbPlazos2.ListIndex), crPrestamo, True, lblTasa2, True
'        End With
'
'    End Select
'
'    semaforo = Regresa_Semaforo(ID)
'
'    If semaforo = "Verde" Then
'        ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\VERDE.bmp")
'        ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
'    ElseIf semaforo = "Amarillo" Then
'        ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\AMARILLO.bmp")
'        ImgSemaforo.Tag = SacaValor("parametros", "PrestamoAmarillo", "")
'    Else
'        ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\ROJO.bmp")
'        ImgSemaforo.Tag = SacaValor("parametros", "PrestamoRojo", "")
'    End If
'
'    '***Puntos***
'    'si no fue buscado por la tarjeta de puntos
'
'    If SacaValor("tarjetaspuntos", "count(id)", " where activa = 1") > 0 Then
'
'        If Not Tarjeta Then
'          If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(ID) Then
'             lblPuntosAcumulados.Caption = TarjetaPuntos.CuentaFrecuente.Puntos
'             txtNoTarjeta.text = TarjetaPuntos.CuentaFrecuente.Folio
'          Else
'             If MsgBox("El Cliente no cuenta con tarjeta de cliente frecuente" & vbCrLf & "Desea asignarle una tarjeta?", vbYesNoCancel Or vbQuestion) = vbYes Then
'                TarjetaPuntos.ShowAsignarTarjeta ID, frmMDI.IDUsuario
'
'                If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(ID) Then
'                    lblPuntosAcumulados.Caption = TarjetaPuntos.CuentaFrecuente.Puntos
'                    txtNoTarjeta.text = TarjetaPuntos.CuentaFrecuente.Folio
'                Else
'                    MsgBox "No se agregó la tarjeta al cliente", vbCritical, "Empeño"
'                End If
'
'             End If
'          End If
'        End If
'    End If
'
'    rcClientes.Close
'    Set rcClientes = Nothing
'    Exit Sub
'
'Error:
'    Maneja_Error Err
'    Set rcClientes = Nothing
'End Sub

'Calculamos el avaluo
Private Function Calcular_Avaluo() As Double
Dim crPrecio As Double, Peso As Double, PesoPiedra As Double, PrestamoDiamante As Double, PrestamoAvaluo As Double, AvaluoDiamante As Double, Margen As Double
Dim rcTmp As New ADODB.Recordset

On Error GoTo Error

    If cmbTipo.ListIndex >= 0 And cmbKilates.ListIndex >= 0 And cmbEstado.ListIndex >= 0 Then

        If Val(txtPeso.text) > 0 Or (Trim(txtPeso.text) <> "" And Trim(txtPeso.text) <> ".") Then
        
            Peso = CDbl(txtPeso.text)
        Else
        
            Peso = 0
        End If
    
        If Val(txtPesoPiedra.text) > 0 Or (Trim(txtPesoPiedra.text) <> "" And Trim(txtPesoPiedra.text) <> ".") Then
            
            PesoPiedra = CDbl(txtPesoPiedra.text)
        Else
            
            PesoPiedra = 0
        End If
    
        If Val(lblPrestamoDiamante.Caption) > 0 Or Trim(lblPrestamoDiamante.Caption) <> "" Then
            
            PrestamoDiamante = CDbl(lblPrestamoDiamante.Caption)
        Else
        
            PrestamoDiamante = 0
        End If
        
        If Val(lblAvaluoDiamante.Caption) > 0 Or Trim(lblAvaluoDiamante.Caption) <> "" Then
            
            AvaluoDiamante = CDbl(lblAvaluoDiamante.Caption)
        Else
            
            AvaluoDiamante = 0
        End If
        
        rcTmp.Open "SELECT Precio FROM PreciosKilataje WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND IDKilataje=" & RegresaKilates(cmbKilates.text, cmbTipo.text) & " AND IDHechura=" & cmbEstado.ItemData(cmbEstado.ListIndex), dbDatos, adOpenForwardOnly, adLockOptimistic
    
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Precio) Then
            
            crPrecio = rcTmp!Precio
        Else
            
            crPrecio = 0
        End If

        rcTmp.Close
    
        PrestamoAvaluo = SacaValor("configuraciontasas", "PorPrestamo", " WHERE IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex) & " AND IDPlazo=" & cmbPlazos.ItemData(cmbPlazos.ListIndex))
        Margen = Regresa_Valor_BD("Negociacion") / 100
        
        Calcular_Avaluo = (Peso - PesoPiedra) * crPrecio
        
        txtAvaluo.text = Format(Redondeo(Calcular_Avaluo + AvaluoDiamante), FMoneda)
        
        txtPrestamoo.text = Format(Calcula_Prestamo(CDbl(Calcular_Avaluo), PrestamoAvaluo) + PrestamoDiamante, FMoneda) / 100 * ImgSemaforo.Tag
        lblPrestamoMaximo.Caption = Format(Redondeo((Calcula_Prestamo(CDbl(Calcular_Avaluo), PrestamoAvaluo) * (1 + Margen)) + PrestamoDiamante), FMoneda) / 100 * ImgSemaforo.Tag


    End If

    Set rcTmp = Nothing
    Exit Function

Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

'Calculamos el total de los avaluos y prestamos
Private Sub Total_Avaluos()
Dim Indice As Integer, Peso As Double, crAvaluo As Double, crPrestamo As Double, crAvaluoTotal As Double, crPrestamoTotal As Double

    m_Peso = 0
   
    For Indice = 1 To grdEmpeños.Rows
        
        crAvaluo = IIf(Val(grdEmpeños.CellText(Indice, 6)) = 0 Or Trim(grdEmpeños.CellText(Indice, 6)) = "", 0, grdEmpeños.CellText(Indice, 6))
        crPrestamo = IIf(Val(grdEmpeños.CellText(Indice, 7)) = 0 Or Trim(grdEmpeños.CellText(Indice, 7)) = "", 0, grdEmpeños.CellText(Indice, 7))
        Peso = IIf(Val(grdEmpeños.CellText(Indice, 4)) = 0 Or Trim(grdEmpeños.CellText(Indice, 4)) = "", 0, grdEmpeños.CellText(Indice, 4))
        
        crAvaluoTotal = crAvaluoTotal + crAvaluo
        crPrestamoTotal = crPrestamoTotal + crPrestamo
        m_Peso = m_Peso + Peso
    Next Indice
   
    txtPrestamo.Caption = Format(crPrestamoTotal, FMoneda)
    lblTotAvaluo.Caption = Format(crAvaluoTotal, FMoneda)
End Sub












Private Sub txtKms_GotFocus()
    Seleccionar_Texto txtKms
    Cambiar_Color True, txtKms
End Sub

Private Sub txtKms_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtKms_LostFocus()
    Cambiar_Color False, txtKms
End Sub


Private Sub txtPrestamo2_GotFocus()
    Cambiar_Color True, txtPrestamo2
    Seleccionar_Texto txtPrestamo2
End Sub

Private Sub txtPrestamo2_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then cmdAceptar.SetFocus
End Sub

Private Sub txtPrestamo2_LostFocus()
    Cambiar_Color False, txtPrestamo2
    txtPrestamo2.text = Format(txtPrestamo2.text, FMoneda)
End Sub

Private Sub txtMarca_GotFocus()
    Seleccionar_Texto txtMarca
    Cambiar_Color True, txtMarca
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMarca_LostFocus()
    Cambiar_Color False, txtMarca
End Sub

Private Sub txtAño_GotFocus()
    Seleccionar_Texto txtAño
    Cambiar_Color True, txtAño
End Sub

Private Sub txtAño_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAño_LostFocus()
    Cambiar_Color False, txtAño
End Sub

Private Sub txtColor_GotFocus()
    Seleccionar_Texto txtColor
    Cambiar_Color True, txtColor
End Sub

Private Sub txtColor_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtColor_LostFocus()
    Cambiar_Color False, txtColor
End Sub

Private Sub txtPlacas_GotFocus()
    Seleccionar_Texto txtPlacas
    Cambiar_Color True, txtPlacas
End Sub

Private Sub txtPlacas_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPlacas_LostFocus()
    Cambiar_Color False, txtPlacas
End Sub

Private Sub txtFactura_GotFocus()
    Seleccionar_Texto txtFactura
    Cambiar_Color True, txtFactura
End Sub

Private Sub txtFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFactura_LostFocus()
    Cambiar_Color False, txtFactura
End Sub

Private Sub txtAgencia_GotFocus()
    Seleccionar_Texto txtAgencia
    Cambiar_Color True, txtAgencia
End Sub

Private Sub txtAgencia_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAgencia_LostFocus()
    Cambiar_Color False, txtAgencia
End Sub

Private Sub txtTarjeta_GotFocus()
    Seleccionar_Texto txtTarjeta
    Cambiar_Color True, txtTarjeta
End Sub

Private Sub txtTarjeta_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTarjeta_LostFocus()
    Cambiar_Color False, txtTarjeta
End Sub

Private Sub txtNummotor_GotFocus()
    Seleccionar_Texto txtNumMotor
    Cambiar_Color True, txtNumMotor
End Sub

Private Sub txtNummotor_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNummotor_LostFocus()
    Cambiar_Color False, txtNumMotor
End Sub

Private Sub txtSeriechasis_GotFocus()
    Seleccionar_Texto txtSerieChasis
    Cambiar_Color True, txtSerieChasis
End Sub

Private Sub txtSeriechasis_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtSeriechasis_LostFocus()
    Cambiar_Color False, txtSerieChasis
End Sub

Private Sub txtGas_GotFocus()
    Seleccionar_Texto txtGas
    Cambiar_Color True, txtGas
End Sub

Private Sub txtGas_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtGas_LostFocus()
    Cambiar_Color False, txtGas
End Sub

Private Sub txtAseguradora_GotFocus()
    Seleccionar_Texto txtAseguradora
    Cambiar_Color True, txtAseguradora
End Sub

Private Sub txtAseguradora_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAseguradora_LostFocus()
    Cambiar_Color False, txtAseguradora
End Sub

Private Sub txtPoliza_GotFocus()
    Seleccionar_Texto txtPoliza
    Cambiar_Color True, txtPoliza
End Sub

Private Sub txtPoliza_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPoliza_LostFocus()
    Cambiar_Color False, txtPoliza
End Sub

Private Sub txtTipoPoliza_GotFocus()
    Seleccionar_Texto txtTipoPoliza
    Cambiar_Color True, txtTipoPoliza
End Sub

Private Sub txtTipoPoliza_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTipoPoliza_LostFocus()
    Cambiar_Color False, txtTipoPoliza
End Sub

Private Sub txtMensaje2_GotFocus()
Seleccionar_Texto txtMensaje2
Cambiar_Color True, txtMensaje2
End Sub

Private Sub txtMensaje2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMensaje2_LostFocus()
    Cambiar_Color False, txtMensaje2
End Sub

Private Sub txtNotas2_GotFocus()
Seleccionar_Texto txtNotas2
    Cambiar_Color True, txtNotas2
End Sub

Private Sub txtNotas2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNotas2_LostFocus()
Cambiar_Color False, txtNotas2
End Sub

Private Sub cmbMedio2_GotFocus()
    Cambiar_Color True, cmbMedio2
End Sub

Private Sub cmbMedio2_LostFocus()
    Cambiar_Color False, cmbMedio2
End Sub

Private Sub Grabar_Empeno_Autos(ID As Long)

    Dim strSql As String, IDEmpeno As Long, Folio As Long, Movimiento As Long, Vencimiento As String, Tasa As Double, Almacenaje As Double, Seguro As Double, Iva As Double, IDCliente As Long, strIniciales As String, Prestamo As Double, Avaluo As Double
    Dim Dias As Long, GTOOperacion As Double, FecVenci As String, Comision As Double, Periodo As Integer, VenPeriodo As Integer, VenAlmoneda As Integer, NumContrato As Long, Promocion As Integer, CostoSeguro As Double, ImporteCostoRentaGPS As Double, EnCirculacion As Long
    '***Puntos***
    Dim PuntosAcumulados As Double, CAT As Double
    
    Dim IDCotitular As Long
    Dim vTipoGarantia As Integer

    
On Error GoTo Error
    
'    If ID = 0 Then
'        ID = Grabar_Cliente_Auto
'        IDCliente = ID
'    Else
'        IDCliente = ID
'        Actualizar_Cliente_Autos IDCliente
'    End If
    
    '----------------------------------------------------
    'MLD-MODIF - Grabar Cliente
    '----------------------------------------------------
    ClienteEmp.Grabar
    ID = ClienteEmp.ID
    IDCliente = ID
    '----------------------------------------------------
    'MLD-MODIF - Grabar Cotitular
    '----------------------------------------------------
    IDCotitular = 0
    If Trim(txtResponsable2.text) <> "" And Trim(txtCotitularApellidoPaterno2.text) <> "" And Trim(txtCotitularApellidoMaterno2.text) <> "" Then
        If CotitularEmp.Valida = True Then
            CotitularEmp.Grabar
            IDCotitular = CotitularEmp.ID
        Else
            MsgBox "Datos incompletos del CoTitular.", vbCritical, "Empeño de Inmuebles"
            Exit Sub
        End If
    End If
    '----------------------------------------------------
    
    'Actualizo el Numero de contratos del cliente
    dbDatos.Execute "UPDATE Clientes SET Boletas=Boletas+1 WHERE ID=" & IDCliente
    
    'Saco el Numero de contrato
    NumContrato = Regresa_NumContrato(False, SERIE_A)
    Regresa_NumContrato True, SERIE_A
    
    'Saco el Folio
    Folio = Regresa_NumContrato(False, SERIE_C)
    Regresa_NumContrato True, SERIE_C
    
    'Saco el Numero de Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Tomo el tipo de promocion
    Select Case cmbPromocion2.ListIndex
    Case 0
        Promocion = 0
    Case 1
        Promocion = 1
    Case 2
        Promocion = 2
    Case 3
        Promocion = 3
    Case 4
        Promocion = 4
    Case 5
        Promocion = 5
    End Select
    
    Select Case cmbPeriodo2.text
    Case "MENSUAL"
        Periodo = 30
    Case "QUINCENAL"
        Periodo = 15
    Case "SEMANAL"
        Periodo = 7
    Case "DIARIA"
        Periodo = 1
    End Select
    VenPeriodo = Val(cmbPlazos2.text)
    Vencimiento = lblVencimiento2.Caption
    Tasa = CDbl(Mid(lblTasa2.Caption, 1, Len(lblTasa2.Caption) - 1))
    Tasa = CDbl(Mid(lblTasa2.Caption, 1, Len(lblTasa2.Caption) - 1))
    CAT = Val(lblTasa2.Tag)
    Prestamo = CDbl(txtPrestamo2.text)
    Avaluo = CDbl(lblTotAvaluo2.Caption)
    Almacenaje = CDbl(Mid(lblAlmacenaje2.Caption, 1, Len(lblAlmacenaje2.Caption) - 1))
    Seguro = CDbl(Mid(lblSeguro2.Caption, 1, Len(lblSeguro2.Caption) - 1))
    Iva = CDbl(Mid(lblIva2.Caption, 1, Len(lblIva2.Caption) - 1))
    GTOOperacion = Regresa_Valor_BD("Operacion")
    Comision = Regresa_Valor_BD("Comision")
    VenAlmoneda = Regresa_Valor_BD("VenAlmoneda")
    ImporteCostoRentaGPS = Regresa_Valor_BD("RentaGPS")
    CostoSeguro = CDbl(txtCostoMensualSeguro.text)
    strIniciales = Iniciales(Trim(txtNombre2.text), Trim(txtApellidoPaterno2.text) & " " & Trim(txtApellidoMaterno2.text))
    If chkAutoencirculacion.Value = 1 Then
        EnCirculacion = 1
    Else
        EnCirculacion = 2
    End If
        
    strSql = "INSERT INTO empeno (Fecha,Movimiento,Numcontrato,Folio,Prestamo,Avaluo,Origen,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Responsable,Valuador,Notas,Tasa,Almacenaje,Seguro,Operacion,Comision,IVA,Periodo,VenPeriodo,VenAlmoneda,Tipointeres,TipoTasa,IDSucursal,IDUsuario,NumBolsa,Ubicacion,Caja,Cajon,Fila,PrestamoInicial,Promocion,IDCoTitular,Cat,ImporteSeguroAuto,ImporteRentaGPS,Circulando) VALUES " & _
            "('" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & NumContrato & "," & Folio & "," & Prestamo & "," & Avaluo & "," & OD_EMPENO & ",'" & Format(Vencimiento, "YYYY/MM/DD") & "'," & Folio & "," & SERIE_B & ",'" & NombrePc & _
            "'," & IDCliente & ",'" & Trim(txtResponsable2.text & " " & txtCotitularApellidoPaterno2.text & " " & txtCotitularApellidoMaterno2.text) & "','" & frmMDI.Usuario & "','" & Trim(txtNotas2.text) & "'," & Tasa & "," & Almacenaje & "," & Seguro & "," & GTOOperacion & "," & Comision & "," & Iva & "," & Periodo & "," & VenPeriodo & "," & VenAlmoneda & ",'" & cmbTipoInteres2.text & "','" & cmbPeriodo2.text & "'," & frmMDI.IDSucursal & "," & frmMDI.IDUsuario & ",0,'','','',''," & Prestamo & "," & Promocion & "," & IDCotitular & "," & CAT & "," & CostoSeguro & "," & ImporteCostoRentaGPS & "," & EnCirculacion & ")"
            
    Err.Clear
    dbDatos.Execute strSql
    
    'Saco el ID del Empeño
    IDEmpeno = SacaValor("empeno", "MAX(ID)")
        
    'MLD-MODIF.
    GuardarDatosLavadoDinero IDEmpeno, "empeno", MLD_INSTRUMENTO_MONETARIO, MLD_PRESTAMO, 0, vTipoAlerta.ID, vTipoAlerta.Descripcion
    
        
    If Trim(txtFechaVenciPoliza.text) = "__/__/____" Then
        
        FecVenci = "Null"
        
    ElseIf Not IsDate(txtFechaVenciPoliza.text) Then
        
        FecVenci = "Null"
    Else
    
        FecVenci = "'" & Format(CDate(txtFechaVenciPoliza.text), "YYYY/MM/DD") & "'"
    End If
    
    'MLD-MODIF.
    vTipoGarantia = 0
    If Val(SacaValor("tipo", "Id", " WHERE Descripcion LIKE '%AUTO%'")) = 0 Then
        vTipoGarantia = Val(SacaValor("mld_prestamos_tipo_garantias", "Id", " WHERE Descripcion LIKE '%Vehículo terrestre%'")) 'Vehículo terrestre
    Else
        vTipoGarantia = Val(SacaValor("tipo", "IdTipoGarantia", " WHERE Descripcion LIKE '%AUTO%'"))
    End If

    
    'Grabo el detalle del auto
    strSql = "INSERT INTO detallesempenoautos(IDEmpeno,MarcayModelo,Año,Color,Placas,Factura,Agencia,NumTarjetaCircu,NumMotor,SerieChasis,Kms,Gas,Aseguradora,Poliza,FechaVenci,Tipo,Factu,TarjetaCircu,CopiaIfe,Tenencias,PolizaSeguro,CopiaLicencia,Importacion,Observaciones,IdTipoGarantia,Marca,Modelo,VIN,RePuVe,IdTipoBlindajeAutos) VALUES (" & _
            IDEmpeno & ",'" & Trim(txtMarca.text) & "'," & Val(txtAño.text) & ",'" & Trim(txtColor.text) & "','" & Trim(txtPlacas.text) & "','" & Trim(txtFactura.text) & "','" & Trim(txtAgencia.text) & "','" & Trim(txtTarjeta.text) & "','" & Trim(txtNumMotor.text) & "','" & Trim(txtSerieChasis.text) & "','" & Trim(txtKms.text) & "','" & Trim(txtGas.text) & "','" & Trim(txtAseguradora.text) & "','" & Trim(txtPoliza.text) & "'," & FecVenci & ",'" & Trim(txtTipoPoliza.text) & "'," & chkFactura.Value & "," & chkTarjeta.Value & "," & chkCopiIfe.Value & "," & chkTenencia.Value & "," & chkPoliza.Value & "," & chkCopyLicencia.Value & "," & chkImportacion.Value & ",'" & Trim(txtObservaciones2.text) & "'," & vTipoGarantia & ",'" & Trim(txtMarca.text) & "','" & Trim(txtModelo.text) & "','" & Trim(txtVIN.text) & "','" & Trim(txtREPUVE.text) & "'," & Val(cmbTipoBlindaje.ItemData(cmbTipoBlindaje.ListIndex)) & ")"
    
    dbDatos.Execute strSql
  
    lblFolio2.Caption = NumContrato
    lblFolio2.Visible = True

    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','201701'," & Prestamo & "," & TIPO_CARGO & "," & SERIE_B & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110150'," & Prestamo & "," & TIPO_ABONO & "," & SERIE_B & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

'''    'Grabamos abono 199450
'''    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
'''                  "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','199450'," & Prestamo & "," & TIPO_ABONO & "," & SERIE_B & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
  
    '***Puntos***
    If TarjetaPuntos.CuentaFrecuente.Folio <> "" Then
        'MsgBox "Puntos Acumulados: " & TarjetaPuntos.Acumula_Puntos(EmpenoAutos, frmMDI.IDUsuario, CCur(Prestamo), NumContrato), vbInformation Or vbOKOnly
        
        dbDatos.Execute "UPDATE empeno SET SaldoPuntosAnteriorEmp = " & TarjetaPuntos.CuentaFrecuente.Puntos & " WHERE ID = " & IDEmpeno
    
        PuntosAcumulados = TarjetaPuntos.Acumula_Puntos(EmpenoAutos, frmMDI.IDUsuario, CCur(Prestamo), NumContrato)
        MsgBox "Puntos Acumulados: " & PuntosAcumulados, vbInformation Or vbOKOnly
            
        dbDatos.Execute "UPDATE empeno SET PuntosAcumuladosEmp=" & PuntosAcumulados & ",SaldoPuntosActualEmp=" & TarjetaPuntos.CuentaFrecuente.Puntos & ",IDTarjetaEmp=" & TarjetaPuntos.CuentaFrecuente.IDCuenta & " WHERE ID = " & IDEmpeno
        
    End If
    
    'Imprimo la Boleta
    'Imprimir_Boleta_CR_Auto IDEmpeno
    'Imprimir_Boleta_CR_Caidas_Autos IDEmpeno
    Imprimir_Boleta_CR_Caidas IDEmpeno, False, True, False
    'MLD-MODIF. ----------------------
    ClienteEmp.Limpiar
    CotitularEmp.Limpiar
    InicializarAlerta vTipoAlerta, MLD_PRESTAMO
    cmdAlerta2.Enabled = False
    '---------------------------------

    Limpiar "Autos"
    Limpiar "DATOS DEL AUTOMÓVIL"
    Limpiar "DOCUMENTOS ENTREGADOS"
    lblTotAvaluo2.Caption = "0.00"
    lblAlmacenaje2.Caption = Format(Regresa_Valor_BD("Almacenaje"), "0.00") & "%"
    lblSeguro2.Caption = Format(Regresa_Valor_BD("Seguro"), "0.00") & "%"
    lblIva2.Caption = Regresa_Valor_BD("IVA") & "%"
    Default 2
    ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\APAGADOS.bmp")
    ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
    ImgSemaforo.ToolTipText = ""
    cmbTipoInteres2.ListIndex = 0
    cmbTipoInteres2_Click
    cmbPromocion2.ListIndex = 0
    txtNotas2.text = Regresa_Valor_BD("Notas")
    
    '***Puntos***
    Limpiar_Tarjeta
    
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

''Grabamos los datos del cliente
'Private Function Grabar_Cliente_Auto() As Long
'Dim rc As New ADODB.Recordset
'Dim Medio As Integer, FechaNac As String, Sexo As Integer, strIniciales As String
'
'On Error GoTo Error
'
'    If cmbMedio2.ListIndex = -1 Then
'
'        Medio = 0
'    Else
'
'        Medio = cmbMedio2.ItemData(cmbMedio2.ListIndex)
'    End If
'
'    If Trim(txtFecNacimiento2.text) = "" Or Trim(txtFecNacimiento2.text) = "__/__/____" Then
'
'        FechaNac = "Null"
'    Else
'
'        FechaNac = "'" & Format(txtFecNacimiento2.text, "YYYY/MM/DD") & "'"
'    End If
'
'    If cmbSexo2.ListIndex = -1 Then
'
'        Sexo = 0
'    Else
'
'        Sexo = cmbSexo2.ItemData(cmbSexo2.ListIndex)
'    End If
'
'    'Saco las Iniciales
'    strIniciales = Iniciales(Trim(txtNombre2.text), Trim(txtApellidos2.text))
'
'    dbDatos.Execute "INSERT INTO Clientes (Nombre,Apellido,Iniciales,Direccion,Colonia,Municipio,Estado,Tel,Identificacion,IDMedio,Notas,CP,FecNac,Sexo) VALUES " & _
'                   "('" & Trim(txtNombre2.text) & "','" & Trim(txtApellidos2.text) & "','" & strIniciales & "','" & Trim(txtDireccion2.text) & "','" & Trim(txtColonia2.text) & "','" & Trim(txtMunicipio2.text) & "','" & _
'                   Trim(txtEstado2.text) & "','" & Trim(txtTelefono2.text) & "','" & Trim(txtIdentificacion2.text) & "'," & Medio & ",'" & Trim(txtMensaje2.text) & "','" & Trim(txtCp2.text) & "'," & FechaNac & "," & Sexo & ")"
'
'    rc.Open "SELECT MAX(ID) AS IDD FROM Clientes", dbDatos, adOpenForwardOnly, adLockOptimistic
'
'        Grabar_Cliente_Auto = rc!idd
'
'    rc.Close
'    Set rc = Nothing
'    Exit Function
'
'Error:
'    Maneja_Error Err
'    Set rc = Nothing
'End Function

Private Function Validar_Empeno_Auto() As Boolean
Dim Prestamo As Double, Avaluo As Double
  
    Validar_Empeno_Auto = True

    '-------------------------------------------------------------------------
    '----- MLD-MODIF.
    '-------------------------------------------------------------------------
    'si no tiene nombre
    If Trim(txtNombre2.text) = "" Then
        MsgBox "Introduzca el Nombre del Cliente !!", vbCritical, "Autos"
        Validar_Empeno_Auto = False
        txtNombre2.SetFocus
        Exit Function
    End If

    'si no tiene apellido
    If Trim(txtApellidoPaterno2.text) = "" Then
        MsgBox "Introduzca el Apellido Paterno del Cliente !!", vbCritical, "Autos"
        Validar_Empeno_Auto = False
        txtApellidoPaterno2.SetFocus
        Exit Function
    End If

    'si no tiene direccion
    If Trim(txtApellidoMaterno2.text) = "" Then
        MsgBox "Introduzca el Apellido Materno del Cliente !!", vbCritical, "Autos"
        Validar_Empeno_Auto = False
        txtApellidoMaterno2.SetFocus
        Exit Function
    End If

    If Not ClienteEmp.Valida Then
        MsgBox "Datos requeridos del Cliente incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno_Auto = False
        cmdEditarCliente2_Click
        Exit Function
    End If
    
    If Trim(txtResponsable.text) <> "" Or Trim(txtCotitularApellidoPaterno.text) <> "" Or Trim(txtCotitularApellidoMaterno.text) <> "" Then
        If Not CotitularEmp.Valida Then
            MsgBox "Datos requeridos del Cotitular incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
            Validar_Empeno_Auto = False
            cmdEditarCotitular2_Click
            Exit Function
        End If
    End If
    '-------------------------------------------------------------------------
    

'    'si no tiene nombre
'    If Trim(txtNombre2.text) = "" Then
'        MsgBox "Introduzca el Nombre del Cliente !!", vbCritical, "Autos"
'        Validar_Empeno_Auto = False
'        txtNombre2.SetFocus
'        Exit Function
'    End If
'
'    'si no tiene apellido
'    If Trim(txtApellidos2.text) = "" Then
'        MsgBox "Introduzca los Apellidos del Cliente !!", vbCritical, "Autos"
'        Validar_Empeno_Auto = False
'        txtApellidos2.SetFocus
'        Exit Function
'    End If
'
'    'si no tiene direccion
'    If Trim(txtDireccion2.text) = "" Then
'        MsgBox "Introduzca la Dirección del Cliente !!", vbCritical, "Autos"
'        Validar_Empeno_Auto = False
'        txtDireccion2.SetFocus
'        Exit Function
'    End If
'
'    'si no tiene colonia
'    If Trim(txtColonia2.text) = "" Then
'        MsgBox "Introduzca la Colonia del Cliente !!", vbCritical, "Autos"
'        Validar_Empeno_Auto = False
'        txtColonia2.SetFocus
'        Exit Function
'    End If
'
'    'si no tiene municipio
'    If Trim(txtMunicipio2.text) = "" Then
'        MsgBox "Introduzca el Municipio del Cliente !!", vbCritical, "Autos"
'        Validar_Empeno_Auto = False
'        txtMunicipio2.SetFocus
'        Exit Function
'    End If
'
'    'si no tiene cp
'    If Trim(txtCp2.text) = "" Then
'        MsgBox "Introduzca el CP del Cliente !!", vbCritical, "Autos"
'        Validar_Empeno_Auto = False
'        txtCp2.SetFocus
'        Exit Function
'    End If
'
'    'si no tiene estado
'    If Trim(txtEstado2.text) = "" Then
'        MsgBox "Introduzca el Estado del Cliente !!", vbCritical, "Autos"
'        Validar_Empeno_Auto = False
'        txtEstado2.SetFocus
'        Exit Function
'    End If
'
'    'si no identificacion
'    If Trim(txtIdentificacion2.text) = "" Then
'        MsgBox "Introduzca la Identificación del Cliente !!", vbCritical, "Autos"
'        Validar_Empeno_Auto = False
'        txtIdentificacion2.SetFocus
'        Exit Function
'    End If

    If txtMarca.text = "" Then
        MsgBox "Introduzca la Marca y el Modelo del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtMarca.SetFocus
        Exit Function
    End If
    
    'MLD-MODIF. --------------------------------------------------------------
    If txtModelo.text = "" Then
        MsgBox "Introduzca el Modelo del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtModelo.SetFocus
        Exit Function
    End If
    
    If txtVIN.text = "" Then
        MsgBox "Introduzca la Marca y el Modelo del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtVIN.SetFocus
        Exit Function
    End If
    
    If txtREPUVE.text = "" Then
        MsgBox "Introduzca el Registro Publico Vehicular (REPUVE) del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtREPUVE.SetFocus
        Exit Function
    End If
    
    If cmbTipoBlindaje.text = "" Then
        MsgBox "Seleccione el Tipo de Blindaje del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        cmbTipoBlindaje.SetFocus
        Exit Function
    End If
    '-------------------------------------------------------------------------
    
    If txtAño.text = "" Then
        MsgBox "Introduzca el Modelo del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtAño.SetFocus
        Exit Function
    End If

    If txtColor.text = "" Then
        MsgBox "Introduzca el Color del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtColor.SetFocus
        Exit Function
    End If

    If txtPlacas.text = "" Then
        MsgBox "Introduzca las Placas del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtPlacas.SetFocus
        Exit Function
    End If

    If txtFactura.text = "" Then
        MsgBox "Introduzca el Número de Factura del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtFactura.SetFocus
        Exit Function
    End If

    If txtAgencia.text = "" Then
        MsgBox "Introduzca la Agencia del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtAgencia.SetFocus
        Exit Function
    End If

    If txtTarjeta.text = "" Then
        MsgBox "Introduzca el Número de Tarjeta de Circulación del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtTarjeta.SetFocus
        Exit Function
    End If

    If txtNumMotor.text = "" Then
        MsgBox "Introduzca el Número de Motor del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtNumMotor.SetFocus
        Exit Function
    End If

    If txtSerieChasis.text = "" Then
        MsgBox "Introduzca la Serie Chasis del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtSerieChasis.SetFocus
        Exit Function
    End If

    If txtKms.text = "" Then
        MsgBox "Introduzca el Kilometraje del Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtKms.SetFocus
        Exit Function
    End If

    If txtGas.text = "" Then
        MsgBox "Introduzca si es de Gas el Automovil !!", vbInformation, "Autos"
        Validar_Empeno_Auto = False
        txtGas.SetFocus
        Exit Function
    End If
    If txtCostoMensualSeguro.text = "" Then
        txtCostoMensualSeguro.text = "0"
    End If
'''''    If txtAseguradora.Text = "" Then
'''''        MsgBox "Introduzca la Aseguradora del Automovil !!", vbInformation, "Autos"
'''''        Validar_Empeno_Auto = False
'''''        txtAseguradora.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If txtPoliza.Text = "" Then
'''''        MsgBox "Introduzca la Pólza del Automovil !!", vbInformation, "Autos"
'''''        Validar_Empeno_Auto = False
'''''        txtPoliza.SetFocus
'''''        Exit Function
'''''    End If

'''''    If Trim(txtFechaVenciPoliza.Text) = "__/__/____" Then
'''''        MsgBox "Introduzca la fecha de vencimiento de la póliza !!", vbInformation, "Autos"
'''''        Validar_Empeno_Auto = False
'''''        txtDia2.SetFocus
'''''        Exit Function
'''''    End If

'''''    If txtTipoPoliza.Text = "" Then
'''''        MsgBox "Introduzca el Tipo de Póliza del Automovil !!", vbInformation, "Autos"
'''''        Validar_Empeno_Auto = False
'''''        txtTipoPoliza.SetFocus
'''''        Exit Function
'''''    End If

    'Si no tiene medio
    If Trim(cmbMedio2.text) = "" Then
        MsgBox "Seleccione el medio por el cual se enteró !!", vbCritical, "Autos"
        Validar_Empeno_Auto = False
        cmbMedio2.SetFocus
        Exit Function
    End If

    'Si el Avaluo es 0 o menor
    'If txtAvaluoo.Text <> "" Then avaluo = txtAvaluoo.Text Else avaluo = 0
    'If avaluo <= 0 Then
    '  MsgBox "El Avalúo no puede ser menor o igual 0 !!", vbCritical, "Autos"
    '  Validar_Empeno_auto = False
    '  txtAvaluoo.SetFocus
    '  Exit Function
    'End If

    'si el prestamo es 0 o menor a 0
    If txtPrestamo2.text <> "" Then Prestamo = txtPrestamo2 Else Prestamo = 0

    If Prestamo <= 0 Then
        MsgBox "El Prestamo no puede ser menor o igual 0 !!", vbCritical, "Autos"
        Validar_Empeno_Auto = False
        txtPrestamo2.SetFocus
        Exit Function
    End If

End Function

'Private Sub Actualizar_Cliente_Autos(ID As Long)
'Dim Medio As Integer, strIniciales As String, FecNac As String, Sexo As Integer
'
'On Error GoTo Error
'
'    If cmbMedio2.ListIndex = -1 Then
'
'        Medio = 0
'    Else
'
'        Medio = cmbMedio2.ItemData(cmbMedio2.ListIndex)
'    End If
'
'    strIniciales = Iniciales(Trim(txtNombre2.text), Trim(txtApellidos2.text))
'
'    If Trim(txtFecNacimiento2.text) = "" Or Trim(txtFecNacimiento2.text) = "__/__/____" Then
'
'        FecNac = "Null"
'    Else
'
'        FecNac = "'" & Format(txtFecNacimiento2.text, "YYYY/MM/DD") & "'"
'    End If
'
'    If cmbSexo2.ListIndex = -1 Then
'
'        Sexo = 0
'    Else
'
'        Sexo = cmbSexo2.ItemData(cmbSexo2.ListIndex)
'    End If
'
'    dbDatos.Execute "UPDATE Clientes SET nombre='" & Trim(txtNombre2.text) & "',apellido='" & Trim(txtApellidos2.text) & "',Iniciales='" & strIniciales & "',Direccion='" & Trim(txtDireccion2.text) & "',Colonia='" & Trim(txtColonia2.text) & "',Municipio='" & Trim(txtMunicipio2.text) & "'," & _
'                    "Estado='" & Trim(txtEstado2.text) & "',Tel='" & Trim(txtTelefono2.text) & "',Identificacion='" & Trim(txtIdentificacion2.text) & "',IDMedio=" & Medio & ",Notas='" & Trim(txtMensaje2.text) & "',CP='" & Trim(txtCp2.text) & "',fecnac=" & FecNac & ",sexo=" & Sexo & _
'                    ",Celular='" & Trim(txtCelular2.text) & "',CorreoElectronico='" & Trim(txtCorreoElectronico2.text) & "' WHERE ID = " & ID
'    Exit Sub
'
'Error:
'    Maneja_Error Err
'End Sub

Public Sub Imprimir_Boleta_CR_Auto(ID As Long)
Dim i As Integer, crIntereses As Double, Contrato As String, CAT As Double
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error
                       
    'Leo los datos del empeño
    rcConsulta.Open "SELECT Prestamo,Avaluo,Folio,Fecha,NumContrato,Serie,TipoInteres,TipoTasa,Vencimiento,VenPeriodo FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockReadOnly
    
    'Opciones de Pago
    crIntereses = OpcionesPagoAutos(rcConsulta!Prestamo, rcConsulta!Avaluo, rcConsulta!Fecha, ID, rcConsulta!TipoTasa)
    
    'Saco el CAT
    CAT = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Cat", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie=" & IIf(rcConsulta!Serie = SERIE_A Or rcConsulta!Serie = SERIE_C, SERIE_A, rcConsulta!Serie) & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & rcConsulta!VenPeriodo))

    Contrato = rcConsulta!NumContrato
    For i = 1 To 6 - Len(Contrato)

        Contrato = "0" & Contrato
    Next i

    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\BoletaAutos.rpt"
        .SelectionFormula = "{empeno.ID}=" & ID
        .Formulas(0) = "CodigoBarras='*" & Contrato & "*'"
        .Formulas(1) = "GastosVenta=" & Regresa_Valor_BD("GtosVenta") & ""
        .Formulas(2) = "Cat=" & CAT
        .Formulas(3) = "ImporteRefrendo=" & crIntereses
        .Formulas(4) = "CantidadLetra='" & CantidadEnLetra(rcConsulta!Avaluo) & "'"
        .Formulas(5) = "FechaComercializacion='" & Format(DateAdd("D", Regresa_Valor_BD("DiasEnajenacion") + 1, rcConsulta!Vencimiento), "DD/MMM/YYYY") & "'"
        .Formulas(6) = "FechaFiniquito='" & Format(DateAdd("D", Regresa_Valor_BD("DiasGracia"), rcConsulta!Vencimiento), "DD/MMM/YYYY") & "'"
        .Formulas(7) = "RazonSocial='" & Sucursal.RazonSocial & "'"
        .Formulas(8) = "DireccionSuc='" & Sucursal.Direccion & " " & Sucursal.Ciudad & " " & Sucursal.Estado & "'"
        .Formulas(9) = "RfcSuc='" & Sucursal.RFC & "'"
        '*************
        .Formulas(10) = "CodProfeco='" & Regresa_Valor_BD("CodProfeco") & "'"
        .Formulas(11) = "Horario='" & Regresa_Valor_BD("HorarioSucursal") & "'"
        '*************
        
        .SubreportToChange = "OpcionPagos"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{opcionpagos.PC}='" & NombrePc & "'"
        .DiscardSavedData = True
        
        .SubreportToChange = "Articulos"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .DiscardSavedData = True
        
        .WindowTitle = "Contrato"
        .WindowState = crptMaximized
        .Action = 1
    End With
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Function Calcular_Avaluo_Auto(Empeno As Double, PrestamoAvaluo As Double) As Double
Dim crAvaluo As Double

On Error GoTo Error
       
    crAvaluo = (Empeno * (1 + PrestamoAvaluo))

    Calcular_Avaluo_Auto = crAvaluo
  
Error:
    Maneja_Error Err
   
End Function

Sub Imprimir_Nota_Auto(IDEmpeno As Long, Opcion As Integer, Optional Abono As Double, Optional IDUsuarioMov As Integer, Optional Comercializacion As Date)
Dim ImprDefault As Boolean

On Error GoTo Error
    
    If Opcion = OD_REFRENDO Then
    
    End If
    
    ImprDefault = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & IIf(Opcion = OD_REFRENDO, "\Reportes\NotaAuto.rpt", "\Reportes\NotaDesempeñoAuto.rpt")
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.ID}=" & IDEmpeno
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .Formulas(2) = "Opcion=" & Opcion & ""
        
        If Opcion = OD_REFRENDO Then
            .Formulas(3) = "Comercializacion='" & Format(Comercializacion, "DD-MMM-YYYY") & "'"
            .SubreportToChange = "OpcionesPagos"
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .SelectionFormula = "{opcionpagos.PC}='" & Nombre_Pc & "'"
            .DiscardSavedData = True
        End If
        
        .WindowState = crptMaximized
        .Destination = crptToPrinter
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                
        .WindowTitle = "Recibo"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub


Sub Imprimir_Nota(IDEmpeno As Long, Opcion As Integer, Optional Abono As Double, Optional IDUsuarioMov As Integer, Optional Comercializacion As Date)
Dim ImprDefault As Boolean

On Error GoTo Error
    
    If Opcion = OD_REFRENDO Then
    
    End If
    
    ImprDefault = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & IIf(Opcion = OD_REFRENDO, "\Reportes\Nota.rpt", "\Reportes\NotaDesempeño.rpt")
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.ID}=" & IDEmpeno
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .Formulas(2) = "Opcion=" & Opcion & ""
        
      If Opcion = OD_REFRENDO Then
            .Formulas(3) = "Comercializacion='" & Format(Comercializacion, "DD-MMM-YYYY") & "'"
            .SubreportToChange = "OpcionesPagos"
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .SelectionFormula = "{opcionpagos.PC}='" & Nombre_Pc & "'"
            .DiscardSavedData = True
        End If
        
        .WindowState = crptMaximized
        .Destination = crptToPrinter
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                
        .WindowTitle = "Recibo"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Sub Imprimir_Nota_GPS_Seguro(IDEmpeno As Long, CargoGPS As Double, CargoSeguro As Double, CargoIva As Double, Opcion As Integer, Optional Abono As Double, Optional IDUsuarioMov As Integer, Optional Comercializacion As Date)
Dim ImprDefault As Boolean

On Error GoTo Error
    
    
    
    ImprDefault = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\NotaGPS.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.ID}=" & IDEmpeno
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "IVA=" & CargoIva & ""
        '.Formulas(2) = "TOTAL=" & CDbl(CargoSeguro + CargoIva + CargoGPS) & ""
        
'        If Opcion = OD_REFRENDO Then
'            '.Formulas(3) = "Comercializacion='" & Format(Comercializacion, "DD-MMM-YYYY") & "'"
'            .SubreportToChange = "OpcionesPagos"
'            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'            .SelectionFormula = "{opcionpagos.PC}='" & Nombre_Pc & "'"
'            .DiscardSavedData = True
'        End If
        
        .WindowState = crptMaximized
        .Destination = crptToPrinter
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                
        .WindowTitle = "Recibo GPS y Seguro Auto"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Function Default(Opcion As Integer)
    
    If Opcion = 1 Then
        
        'txtMunicipio.text = ""
        'txtEstado.text = ""
        'txtCantidad.text = "1"
    Else
        
        'txtMunicipio2.text = ""
        'txtEstado2.text = ""
    End If
    
End Function

Function GeneraAutorizacion(TipoAutorizacion As Long) As Long
Dim Codigo As String, Consecutivo As String, i As Integer

On Error GoTo Error

    Consecutivo = Regresa_Movimiento(False, "FolioAutorizacion")
    Regresa_Movimiento True, "FolioAutorizacion"
    
    For i = 1 To 5 - Len(Consecutivo)
        Consecutivo = "0" & Consecutivo
    Next i
    
    Codigo = Format(Month(Date), "00") & Format(Day(Date), "00") & Format(frmMDI.IDSucursal, "000") & Consecutivo
    Codigo = CreaDigitoVerificador(Codigo)
    
    lblAutorizacion.Caption = "AUTORIZACIÓN: " & Codigo
    '''''lblAutorizacion.Visible = True
    
    dbDatos.Execute "INSERT INTO autorizaciones (Fecha,IDUsuario,Opcion,Status,IDSucursal,Codigo) VALUES ('" & Format(Date, "YYYY/MM/DD") & "'," & frmMDI.IDUsuario & "," & TipoAutorizacion & ",0," & frmMDI.IDSucursal & ",'" & Trim(Codigo) & "')"
    GeneraAutorizacion = SacaValor("autorizaciones", "MAX(ID)")

Error:
    Maneja_Error Err
End Function

Function VerificaImporte(crPrestamo As Double, ByRef Ban As Boolean, ByRef IDAutorizacion As Long)
Dim crLimite1 As Double, crLimite2 As Double
    
    Ban = False
    IDAutorizacion = 0
    crLimite1 = CDbl(Regresa_Valor_BD("Limite1"))
    crLimite2 = CDbl(Regresa_Valor_BD("Limite2"))
    IDUsuarioAutoriza = 0
    TipoAutorizacion = 0
    
    If crPrestamo < crLimite1 Then
        
        Ban = True
        
    ElseIf crPrestamo >= crLimite1 And crPrestamo < crLimite2 Then
        
        frmPasswords.ConexSuc = 0
        frmPasswords.PrecioVitrina = 0
        frmPasswords.Cancel = 0
        frmPasswords.Ventas = 0
        frmPasswords.ModificaCorte = 0
        frmPasswords.HacerCorte = 0
        frmPasswords.InteresDesempeño = 0
        frmPasswords.InteresRefrendo = 0
        frmPasswords.ModificaPrecio = 0
        frmPasswords.DescuentoVentas = 0
        frmPasswords.RecalculoPrecios = 0
        frmPasswords.Vencido = 0
        frmPasswords.CancelaCierre = 0
        frmPasswords.AutorizaPrestamo = 1
        
        If frmPasswords.Password(GERENTE, 1) Then
            
            Ban = True
            TipoAutorizacion = AUTORIZACIONLIMITE1
            
            dbDatos.Execute "INSERT INTO autorizaciones (Fecha,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            IDAutorizacion = SacaValor("autorizaciones", "MAX(ID)")
        End If
                
    ElseIf crPrestamo >= crLimite2 Then
        
        frmAutorizacionPrestamo.GeneraCodigo Regresa_Sucursal("Clave") & frmMDI.IDUsuario & Format(Date, "DDMMYY") & Format(Time, "HHMMSS"), Ban, IDAutorizacion
        TipoAutorizacion = AUTORIZACIONLIMITE2
        
    End If

End Function

Sub SacaTasa(crPrestamo As Double, TipoInteres As Integer, TipoPeriodo As Integer, TipoPlazo As Integer, ExisteCliente As Boolean)
Dim rcConsulta As New ADODB.Recordset
Dim Meses As Integer

On Error GoTo Error
    
    'Tasa
    rcConsulta.Open "SELECT ti.Descripcion AS TipoInteres,tp.Descripcion AS TipoPeriodo,tp.Periodo,p.Descripcion AS Vencimiento,ct.Almacenaje,ct.Seguro, ct.Cat " _
                    & "FROM configuraciontasas ct INNER JOIN plazos p ON ct.IDPlazo=p.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN tipointeres ti ON ct.IDTipoInteres=ti.ID WHERE " _
                    & "ct.IDTipoInteres=" & TipoInteres & " AND ct.IDTipoPeriodo=" & TipoPeriodo & " AND ct.IDPlazo=" & TipoPlazo, dbDatos, adOpenForwardOnly, adLockReadOnly

    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        Select Case rcConsulta!TipoPeriodo
        Case "MENSUAL"
            
            Meses = 1
        Case "QUINCENAL"
            
            Meses = 2
        Case "SEMANAL"
            
            Meses = 4
            
        End Select
                
        lblAlmacenaje.Caption = Format(rcConsulta!Almacenaje, "0.00") & "%"
        lblSeguro.Caption = Format(rcConsulta!Seguro, "0.00") & "%"
        
        lblAlmacenaje2.Caption = Format(rcConsulta!Almacenaje, "0.00") & "%"
        lblSeguro2.Caption = Format(rcConsulta!Seguro, "0.00") & "%"
        
        lblTasa.Tag = rcConsulta!CAT
        lblTasa2.Tag = rcConsulta!CAT
        
        'lblVencimiento.Caption = Format(IIf(cmbTipoInteres.text = "FIJA", DateAdd("M", 2, Date), IIf(rcConsulta!Periodo = 30, DateAdd("M", rcConsulta!Vencimiento, Date), DateAdd("D", rcConsulta!Periodo * rcConsulta!Vencimiento, Date))), "DD/MMM/YYYY")
        lblVencimiento.Caption = Format(IIf(cmbTipoInteres.text = "FIJA", DateAdd("D", (rcConsulta!Periodo * rcConsulta!Vencimiento) - 1, Date), IIf(rcConsulta!Periodo = 30, DateAdd("M", rcConsulta!Vencimiento, Date), DateAdd("D", rcConsulta!Periodo * rcConsulta!Vencimiento, Date))), "DD/MMM/YYYY")
        lblVencimiento2.Caption = Format(IIf(cmbTipoInteres2.text = "FIJA", DateAdd("D", rcConsulta!Periodo * rcConsulta!Vencimiento, Date), IIf(rcConsulta!Periodo = 30, DateAdd("M", rcConsulta!Vencimiento, Date), DateAdd("D", rcConsulta!Periodo * rcConsulta!Vencimiento, Date))), "DD/MMM/YYYY")
        
        'Muestro la Tasa
        MuestraTasa TipoInteres, TipoPeriodo, TipoPlazo, crPrestamo, ExisteCliente, IIf(TPestañas.SelectedTab = 1, lblTasa, lblTasa2), IIf(TPestañas.SelectedTab = 1, False, True)
    
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Function SacaTasaAutos()
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error

'Tasa
rcConsulta.Open "SELECT Tasa,Tarjeta,PrestamoAvaluo FROM tasas WHERE ID=" & cmbTipoInteres2.ItemData(cmbTipoInteres2.ListIndex), dbDatos, adOpenForwardOnly, adLockOptimistic
If Not rcConsulta.BOF And Not rcConsulta.EOF Then
    
    If Val(txtPrestamo2.text) > 0 Or Trim(txtPrestamo2.text) <> "" Then
        
        lblTotAvaluo2.Caption = Format(Calcular_Avaluo_Auto(CCur(txtPrestamo2.text), CDbl(SacaValor("tasas", "PrestamoAvaluo", " WHERE ID=" & cmbTipoInteres2.ItemData(cmbTipoInteres2.ListIndex))) / 100), "$###,###,###,###0.00")
    Else
        
        lblTotAvaluo2.Caption = Format(0, "$###,###,###,###0.00")
    End If
    
    lblTasa2.Caption = IIf(IsNull(rcConsulta!Tasa), 0, rcConsulta!Tasa) & "%"
End If
rcConsulta.Close

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function

Public Function ValidaArticulos(Pestaña As Integer) As Boolean
Dim rcTmp As ADODB.Recordset

On Error GoTo Error
    
    If Pestaña = 1 Then Set rcTmp = dbDatos.Execute("SELECT Kilataje,Peso FROM tipo WHERE ID=" & cmbTipo.ItemData(cmbTipo.ListIndex))
    
    ValidaArticulos = True

    If IIf(Pestaña = 1, cmbTipo.text, cmbTipoElec.text) = "" Then
        MsgBox "Seleccione el tipo !!", vbInformation, "Empeños"
        ValidaArticulos = False
        If Pestaña = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
        Exit Function
    End If
       
    If IIf(Pestaña = 1, cmbPrenda.text, txtFamiliaElec.text) = "" Then ''''' pestaña = 1 And cmbPrenda.ListIndex = -1 Then
        MsgBox "Seleccione la " & IIf(Pestaña = 1, "prenda", "familia") & " !!", vbInformation, "Empeños"
        ValidaArticulos = False
        If Pestaña = 1 Then cmbPrenda.SetFocus Else txtFamiliaElec.SetFocus
        Exit Function
    End If

    If txtCantidad.text = "" And Pestaña = 1 Then
        MsgBox "Introduzca la cantidad !!", vbInformation, "Empeños"
        ValidaArticulos = False
        txtCantidad.SetFocus
        Exit Function
    End If
    
    If Pestaña = 1 Then
        If cmbKilates.text = "" And rcTmp!Kilataje = 1 Then
            MsgBox "Seleccione el kilataje !!", vbInformation, "Empeños"
            ValidaArticulos = False
            cmbKilates.SetFocus
            Exit Function
        End If
        
        If cmbEstado.text = "" And rcTmp!Peso = 1 Then
            MsgBox "Seleccione la hechura !!", vbInformation, "Empeños"
            ValidaArticulos = False
            cmbEstado.SetFocus
            Exit Function
        End If
        
        If txtPeso.text = "" And rcTmp!Peso = 1 Then
            MsgBox "Introduzca el peso !!", vbInformation, "Empeños"
            ValidaArticulos = False
            txtPeso.SetFocus
            Exit Function
        End If
    End If
    
    If Pestaña = 2 And Trim(txtMarcaElec.text) = "" Then
        MsgBox "Introduzca la marca !!", vbInformation, "Empeños"
        ValidaArticulos = False
        txtMarcaElec.SetFocus
        Exit Function
    End If
    
'''''    If pestaña = 2 And Trim(txtModeloElec.text) = "" Then
'''''        MsgBox "Introduzca el modelo !!", vbInformation, "Empeños"
'''''        ValidaArticulos = False
'''''        txtModeloElec.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If pestaña = 2 And Trim(txtNumSerieElec.text) = "" Then
'''''        MsgBox "Introduzca el número de serie !!", vbInformation, "Empeños"
'''''        ValidaArticulos = False
'''''        txtNumSerieElec.SetFocus
'''''        Exit Function
'''''    End If
    
    If IIf(Pestaña = 1, txtPrestamoo.text, txtPrestamooElec) = "" Then
        MsgBox "Introduzca el préstamo !!", vbInformation, "Empeños"
        ValidaArticulos = False
        If Pestaña = 1 Then txtPrestamoo.SetFocus Else txtPrestamooElec.SetFocus
        Exit Function
    End If

Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

Sub LimpiaArticulos()
    lblPrestamoDiamante.Caption = "0"
    lblAvaluoDiamante.Caption = "0"
    lblPiedra.Caption = ""
    lblPuntos.Caption = "0"
    lblCantidadPiedras.Caption = "0"
    txtCantidad.text = ""
    txtCantidad.Tag = ""
    txtPesoPiedra.text = ""
    txtPesoPiedra.Tag = ""
    txtPeso.text = ""
    txtObservaciones.text = ""
    txtFamiliaElec.text = ""
    txtFamiliaElec.Tag = ""
    txtMarcaElec.text = ""
    txtModeloElec.text = ""
    txtNumSerieElec.text = ""
    txtTamañoElec.text = ""
    txtColorElec.text = ""
    txtPrestamooElec.text = ""
    txtPrestamooElec.Tag = ""
    txtAvaluoElec.text = ""
    txtObservacionesElec.text = ""
    txtPiedras.text = ""
    cmbPrenda.ListIndex = -1
    cmbTipo.ListIndex = 0
    cmbKilates.ListIndex = -1
    cmbEstado.ListIndex = -1
    cmbPrenda.ListIndex = -1
    cmbTipoElec.ListIndex = 0
    txtPrestamoo.text = ""
    txtAvaluo.text = ""
    lblPrestamoMaximo.Caption = "0.00"
    lblPrestamoMaximo2.Caption = "0.00"
End Sub

'Calculamos el prestamo
Private Function Calcula_Prestamo(Prestamo As Double, PorcentajePrestamo As Double, Optional Pestaña As Boolean = True) As Double

On Error GoTo Error:
    
    If Pestaña Then
        
        Calcula_Prestamo = Redondeo(Prestamo * (PorcentajePrestamo / 100))
    Else
    
        Calcula_Prestamo = Redondeo((Prestamo / PorcentajePrestamo) * 100)
    End If
    Exit Function
    
Error:
    Maneja_Error Err
   
End Function

Sub Recalcula()
Dim i As Integer, crPrestamoOriginal As Double, Avaluo As Double, Porcentaje As Double

    Porcentaje = SacaValor("configuraciontasas", "PorPrestamo", " WHERE IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex) & " AND IDPlazo=" & cmbPlazos.ItemData(cmbPlazos.ListIndex)) / 100

    With grdEmpeños

        For i = 1 To .Rows

            If Val(.CellText(i, 6)) > 0 And Trim(.CellText(i, 6)) <> "" And Val(.CellItemData(i, 1)) = 1 Then

'''''                crPrestamoOriginal = .CellText(i, 7)
                Avaluo = .CellText(i, 6)
                .CellText(i, 7) = Redondeo(Avaluo * Porcentaje)
                .CellTextAlign(i, 7) = DT_RIGHT
            End If

        Next i

        Total_Avaluos
    End With
End Sub

Sub MuestraDatosContrato(IDEmpeno As Long, Indice As Integer)
Dim rcAux As New ADODB.Recordset
Dim rcTmp As New ADODB.Recordset
Dim i As Integer, sqlPrendas As String, strDescripcion As String

On Error GoTo Error
            
    With rcAux
        
        '***Puntos***
        .Open "SELECT e.Fecha,e.Destino,e.Responsable,e.NumBolsa,e.Vencimiento,e.TipoInteres,e.TipoTasa,e.VenPeriodo,e.Tasa,e.Almacenaje,e.Seguro,e.IVA,e.IDUsuario,e.Promocion,c.ID AS IDCliente,concat(c.Nombre,' ',c.Apellido) AS Cliente,c.Direccion,c.Colonia,c.Municipio,c.Estado,c.Notas,c.CP FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.ID=" & IDEmpeno, dbDatos, adOpenForwardOnly, adLockOptimistic
            
            DatosCliente(Indice).Clear
            DatosContrato(Indice).Clear
            DetallesContrato(Indice).Clear
            
            '***Puntos***
            'buscamos la tarjeta del cliente
            TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente !IDCliente
            
            'Muestro los datos
            DatosCliente(Indice).Tag = !Cliente
            DatosCliente(Indice).Add "<bold> " & !Cliente & "</bold>"
            DatosCliente(Indice).Add " " & !Direccion & " " & !Colonia & vbCrLf & _
                                        " " & !Municipio & ", " & !Estado & " C.P. " & !CP & vbCrLf & _
                                        IIf(IsNull(!Notas) Or Trim(!Notas) = "", "", " MENSAJE: " & !Notas & vbCrLf) & IIf(IsNull(!Responsable) Or Trim(!Responsable) = "", "", " COTITULAR: " & !Responsable & vbCrLf) & IIf(!NumBolsa <> "", " NUM. BOLSA: " & !NumBolsa & vbCrLf, "") & " USUARIO: " & SacaValor("usuarios", "Nombre", " WHERE ID=" & !IDUsuario)
            
            '***Puntos***
            DatosCliente(Indice).Add "No. Tarjeta: " & TarjetaPuntos.CuentaFrecuente.Folio & vbCrLf & _
                                     "Puntos Acumulados: " & TarjetaPuntos.CuentaFrecuente.Puntos
            
            DatosContrato(Indice).Add " FECHA EMPEÑO: " & Format(!Fecha, "DD/MMM/YYYY HH:MM:SS AM/PM") & vbCrLf & _
                                        " FECHA VENCIMIENTO: " & Format(!Vencimiento, "DD/MMM/YYYY") & vbCrLf & _
                                        " PLAZO: " & !TipoInteres & " " & !VenPeriodo & " " & IIf(!TipoTasa = "MENSUAL", "MESES", IIf(!TipoTasa = "QUINCENAL", "QUINCENAS", IIf(!TipoTasa = "SEMANAL", "SEMANAS", "DIAS"))) & vbCrLf & _
                                        " TASA: " & Format(((!Tasa + !Almacenaje + !Seguro) * (1 + (!Iva / 100))), "0.00") & "%" & vbCrLf & _
                                        IIf(!Promocion > 0, " CONTRATO PROMOCIÓN " & LeyendaPromocion(!Promocion), "")
            
            Timer1.Enabled = False: labelContratoAlmoneda.Visible = False: labelContratoDesemp.Visible = False
            If !Destino = D_ALMONEDA Then
                If Indice = 1 Then
                    labelContratoDesemp.Caption = "CONTRATO EN ALMONEDA!"
                    Timer2.Enabled = True
                Else
                    labelContratoAlmoneda.Caption = "CONTRATO EN ALMONEDA!"
                    Timer1.Enabled = True
                End If
                
            ElseIf Date > DateAdd("D", Val(Regresa_Valor_BD("DiasGracia")), !Vencimiento) Then
                If Indice = 1 Then
                    labelContratoDesemp.Caption = "CONTRATO VENCIDO!"
                    Timer2.Enabled = True
                Else
                    labelContratoAlmoneda.Caption = "CONTRATO VENCIDO!"
                    Timer1.Enabled = True
                End If
            End If
            
            'Saco las Prendas
            If !Destino = 0 Then
                
                sqlPrendas = "SELECT d.Cantidad,d.Articulo AS Descripcion,d.Peso,d.Prestamo,d.Marca,d.Modelo,kilatajes.Descripcion AS Kilataje " & _
                            "FROM detallesempeno d LEFT JOIN kilatajes ON d.Kilates=kilatajes.ID WHERE d.IDEmpeno=" & IDEmpeno
                
            Else
                
                sqlPrendas = "SELECT d.Cantidad,d.Descripcion,d.Peso,d.Costo AS Prestamo,d.Marca,d.Modelo,kilatajes.Descripcion AS Kilataje " & _
                            "FROM detallesentradainventario d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE d.Cantidad>0 AND d.IDEmpeno=" & IDEmpeno
                
            End If
            
            rcTmp.Open sqlPrendas, dbDatos, adOpenForwardOnly, adLockReadOnly
            While Not rcTmp.EOF
                
                
                For i = 1 To Len(rcTmp!Cantidad & " " & rcTmp!Descripcion & IIf(rcTmp!Peso > 0, " " & rcTmp!Peso & " Grms.", "") & IIf(IsNull(rcTmp!Kilataje) Or Trim(rcTmp!Kilataje) = "", "", " " & rcTmp!Kilataje) & IIf(IsNull(rcTmp!Marca) Or Trim(rcTmp!Marca) = "", "", " MARCA: " & rcTmp!Marca) & IIf(IsNull(rcTmp!Modelo) Or Trim(rcTmp!Modelo) = "", "", " MODELO: " & rcTmp!Modelo)) Step 50

                    strDescripcion = strDescripcion & IIf(Trim(strDescripcion) <> "", vbCrLf, "") & Mid(rcTmp!Cantidad & " " & rcTmp!Descripcion & IIf(rcTmp!Peso > 0, " " & rcTmp!Peso & " Grms.", "") & IIf(IsNull(rcTmp!Kilataje) Or Trim(rcTmp!Kilataje) = "", "", " " & rcTmp!Kilataje) & IIf(IsNull(rcTmp!Marca) Or Trim(rcTmp!Marca) = "", "", " MARCA: " & rcTmp!Marca) & IIf(IsNull(rcTmp!Modelo) Or Trim(rcTmp!Modelo) = "", "", " MODELO: " & rcTmp!Modelo), i * 1, 50)
                Next i
                
            rcTmp.MoveNext
            Wend
            rcTmp.Close
            Set rcTmp = Nothing
                                        
            'Imprimo la descripción
            DetallesContrato(Indice).Add strDescripcion
                
        .Close
    End With
    Set rcAux = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcAux = Nothing
End Sub

Function MuestraTasa(TipoInteres As Integer, TipoPeriodo As Integer, TipoPlazo As Integer, crPrestamo As Double, ExisteCliente As Boolean, Etiqueta As Label, Autos As Boolean)
Dim rcTasas As New ADODB.Recordset
Dim TasaTipica As Double, TasaPromocion As Double, TasaPreferencial As Double, LimiteInferior As Double, LimiteSuperior As Double
    
    TasaTipica = 0
    TasaPromocion = 0
    TasaPreferencial = 0
    LimiteInferior = 0
    LimiteSuperior = 0
    
    LimiteInferior = Regresa_Valor_BD("LimiteInferior" & IIf(Autos, "Autos", ""))
    LimiteSuperior = Regresa_Valor_BD("LimiteSuperior" & IIf(Autos, "Autos", ""))
    
    rcTasas.Open "SELECT TasaTipica,TasaPromocion,TasaPreferencial FROM configuraciontasas WHERE IDTipoInteres=" & TipoInteres & " AND IDTipoPeriodo=" & TipoPeriodo & " AND IDPlazo=" & TipoPlazo, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not rcTasas.BOF And Not rcTasas.EOF Then
        
        TasaTipica = rcTasas!TasaTipica
        TasaPromocion = rcTasas!TasaPromocion
        TasaPreferencial = rcTasas!TasaPreferencial
    End If
    rcTasas.Close
    Set rcTasas = Nothing
    
    If ExisteCliente = False Then
        
        If crPrestamo >= LimiteSuperior Then
            
            MuestraTasa = TasaPreferencial
            
        ElseIf crPrestamo >= LimiteInferior Then
            
            MuestraTasa = TasaPromocion
        
        Else
            
            MuestraTasa = TasaTipica
        End If
    
    Else
        
        If crPrestamo >= LimiteInferior Then
            
            MuestraTasa = TasaPreferencial
        Else
            
            MuestraTasa = TasaTipica
        End If
        
    End If
    
    Etiqueta.Caption = Format(MuestraTasa, "0.00") & "%"
End Function

Sub Calcula_Avaluo_Elec()
Dim crPrestamo As Double, crMaximo As Double, PrestamoAvaluo As Double
    
    crPrestamo = 0
    crMaximo = 0
    If Val(txtPrestamooElec.text) > 0 Or Trim(txtPrestamooElec.text) <> "" Then
        
        'Tomo el prestamo
        crPrestamo = CDbl(txtPrestamooElec.text) '/ 100 * ImgSemaforo.Tag
        
        'Tomo el importe máximo
        If Val(txtPrestamooElec.Tag) > 0 Or Trim(txtPrestamooElec.Tag) <> "" Then
        
            crMaximo = CDbl(txtPrestamooElec.Tag) '/ 100 * ImgSemaforo.Tag
        End If
        
        If crPrestamo > crMaximo Then
            MsgBox "Ha sobrepasado el limite máximo permitido !!", vbInformation, "Empeño"
            crPrestamo = crMaximo
            txtPrestamooElec.text = crMaximo
        End If
        
        PrestamoAvaluo = Regresa_Valor_BD("PrestamoAvaluoElec") / 100
        
        txtAvaluoElec.text = Format(crPrestamo * (1 + PrestamoAvaluo), FMoneda)
    
    End If
        
End Sub

Function CalculaCambio(crEfectivo As Double, crImporte As Double, Pestana As Integer) As Boolean
Dim lblLeyenda As Label, lblOperacion As Label
    
    Set lblOperacion = IIf(Pestana = 1, TotalDesempeño, TotalRefrendo)
    Set lblLeyenda = IIf(Pestana = 1, Leyenda, LeyendaRef)
    
    lblOperacion.Caption = Format(crEfectivo - crImporte, FMoneda)
    lblOperacion.ForeColor = &HFF&
    lblLeyenda.Caption = "CAMBIO:"
    lblLeyenda.Tag = 1
    Abrir_Cajon
End Function

Function VerificaContratoDuplicado(NumContrato As Long, Grid As vbalGrid, Indice As Integer) As Boolean
Dim i As Integer, Bandera As Boolean
    
    VerificaContratoDuplicado = False
    For i = 1 To Grid.Rows - 1
        
        If NumContrato = Grid.CellText(i, 1) Then VerificaContratoDuplicado = True: Exit For
    
    Next i
    
    If VerificaContratoDuplicado Then MsgBox "Contrato duplicado, el contrato ya esta listo para el movimiento !!", vbInformation, IIf(Indice = 1, "Desempeño", "Refrendo")
End Function

'***Puntos***
Private Sub Limpiar_Tarjeta()
   txtNoTarjeta.text = ""
   lblPuntosAcumulados.Caption = "0"
   TarjetaPuntos.CuentaFrecuente.Clear
End Sub


'----------------------------------------------------------
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
    ClienteEmp.Nombre = Trim(txtNombre.text)
End Sub

Private Sub txtNombre2_GotFocus()
    Seleccionar_Texto txtNombre2
    Cambiar_Color True, txtNombre2
End Sub

Private Sub txtNombre2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNombre2_LostFocus()
    Cambiar_Color False, txtNombre2
    ClienteEmp.Nombre = Trim(txtNombre2.text)
End Sub
'---------------------------------------------------------
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
    ClienteEmp.ApellidoPaterno = Trim(txtApellidoPaterno.text)
End Sub

Private Sub txtApellidoPaterno2_GotFocus()
    Seleccionar_Texto txtApellidoPaterno2
    Cambiar_Color True, txtApellidoPaterno2
End Sub

Private Sub txtApellidoPaterno2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtApellidoPaterno2_LostFocus()
    Cambiar_Color False, txtApellidoPaterno2
    ClienteEmp.ApellidoPaterno = Trim(txtApellidoPaterno2.text)
End Sub
'---------------------------------------------------------
Private Sub txtApellidoMaterno_GotFocus()
   Seleccionar_Texto txtApellidoMaterno
   Cambiar_Color True, txtApellidoMaterno
End Sub

Private Sub txtApellidoMaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

'MLD-MODIF.
Private Sub txtApellidoMaterno_LostFocus()
    Cambiar_Color False, txtApellidoMaterno
    ClienteEmp.ApellidoMaterno = Trim(txtApellidoMaterno.text)
    Titular = True
    If Trim(txtNombre.text) <> "" And Trim(txtApellidoPaterno.text) <> "" And Trim(txtApellidoMaterno.text) <> "" And Val(txtNombre.Tag) <= 0 Then Mostrar_Seleccionar_Cliente Trim(txtNombre.text), Trim(txtApellidoPaterno.text), Trim(txtApellidoMaterno.text), Me
    If Val(txtNombre.Tag) = 0 Then
        ClienteEmp.Nombre = Trim(txtNombre.text)
        ClienteEmp.ApellidoPaterno = Trim(txtApellidoPaterno.text)
        ClienteEmp.ApellidoMaterno = Trim(txtApellidoMaterno.text)
        frmClientes.Mostrar ClienteEmp
        cmdAlerta.Enabled = True
        '--- Mostrar Datos Cliente ---
        lblDireccion.Caption = ClienteEmp.Direccion & IIf(ClienteEmp.NoExterior <> "", " #" & ClienteEmp.NoExterior, "") & IIf(ClienteEmp.NoInterior <> "", " INT." & ClienteEmp.NoInterior, "") & " COL." & ClienteEmp.Colonia & " C.P." & ClienteEmp.CodigoPostal
        lblCiudad.Caption = ClienteEmp.Municipio & ", " & ClienteEmp.Estado
        lblRFC.Caption = ClienteEmp.RFC
    End If
End Sub

Private Sub txtApellidoMaterno2_GotFocus()
   Seleccionar_Texto txtApellidoMaterno2
   Cambiar_Color True, txtApellidoMaterno2
End Sub

Private Sub txtApellidoMaterno2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtApellidoMaterno2_LostFocus()
    Cambiar_Color False, txtApellidoMaterno2
    ClienteEmp.ApellidoMaterno = Trim(txtApellidoMaterno2.text)
    Titular = True
    If Trim(txtNombre2.text) <> "" And Trim(txtApellidoPaterno2.text) <> "" And Trim(txtApellidoMaterno2.text) <> "" And Val(txtNombre2.Tag) <= 0 Then Mostrar_Seleccionar_Cliente Trim(txtNombre2.text), Trim(txtApellidoPaterno2.text), Trim(txtApellidoMaterno2.text), Me
    If Val(txtNombre2.Tag) = 0 Then
        ClienteEmp.Nombre = Trim(txtNombre2.text)
        ClienteEmp.ApellidoPaterno = Trim(txtApellidoPaterno2.text)
        ClienteEmp.ApellidoMaterno = Trim(txtApellidoMaterno2.text)
        frmClientes.Mostrar ClienteEmp
        cmdAlerta2.Enabled = True
        '--- Mostrar Datos Cliente ---
        lblDireccion2.Caption = ClienteEmp.Direccion & IIf(ClienteEmp.NoExterior <> "", " #" & ClienteEmp.NoExterior, "") & IIf(ClienteEmp.NoInterior <> "", " INT." & ClienteEmp.NoInterior, "") & " COL." & ClienteEmp.Colonia & " C.P." & ClienteEmp.CodigoPostal
        lblCiudad2.Caption = ClienteEmp.Municipio & ", " & ClienteEmp.Estado
        lblRFC2.Caption = ClienteEmp.RFC
    End If
End Sub
'-------------------------------------------------------------

Private Sub txtResponsable_GotFocus()
    Seleccionar_Texto txtResponsable
    Cambiar_Color True, txtResponsable
End Sub

Private Sub txtResponsable_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtResponsable_LostFocus()
    Cambiar_Color False, txtResponsable
    CotitularEmp.Nombre = Trim(txtResponsable.text)
End Sub

Private Sub txtResponsable2_GotFocus()
    Seleccionar_Texto txtResponsable2
    Cambiar_Color True, txtResponsable2
End Sub

Private Sub txtResponsable2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtResponsable2_LostFocus()
    Cambiar_Color False, txtResponsable2
End Sub

Private Sub txtCotitularApellidoPaterno_GotFocus()
   Seleccionar_Texto txtCotitularApellidoPaterno
   Cambiar_Color True, txtCotitularApellidoPaterno
End Sub

Private Sub txtCotitularApellidoPaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCotitularApellidoPaterno_LostFocus()
    Cambiar_Color False, txtCotitularApellidoPaterno
    CotitularEmp.ApellidoPaterno = Trim(txtCotitularApellidoPaterno.text)
End Sub

Private Sub txtCotitularApellidoPaterno2_GotFocus()
   Seleccionar_Texto txtCotitularApellidoPaterno2
   Cambiar_Color True, txtCotitularApellidoPaterno2
End Sub

Private Sub txtCotitularApellidoPaterno2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCotitularApellidoPaterno2_LostFocus()
    Cambiar_Color False, txtCotitularApellidoPaterno2
    CotitularEmp.ApellidoPaterno = Trim(txtCotitularApellidoPaterno2.text)
End Sub

Private Sub txtCotitularApellidoMaterno_GotFocus()
   Seleccionar_Texto txtCotitularApellidoMaterno
   Cambiar_Color True, txtCotitularApellidoMaterno
End Sub

Private Sub txtCotitularApellidoMaterno_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCotitularApellidoMaterno_LostFocus()
    Cambiar_Color False, txtCotitularApellidoMaterno
    CotitularEmp.ApellidoMaterno = Trim(txtCotitularApellidoMaterno.text)
    Titular = False
    If Trim(txtResponsable.text) <> "" And Trim(txtCotitularApellidoPaterno.text) <> "" And Trim(txtCotitularApellidoMaterno.text) <> "" And Val(txtResponsable.Tag) <= 0 Then Mostrar_Seleccionar_Cliente Trim(txtResponsable.text), Trim(txtCotitularApellidoPaterno.text), Trim(txtCotitularApellidoMaterno.text), Me
    If Val(txtResponsable.Tag) = 0 Then
        CotitularEmp.Nombre = Trim(txtResponsable.text)
        CotitularEmp.ApellidoPaterno = Trim(txtCotitularApellidoPaterno.text)
        CotitularEmp.ApellidoMaterno = Trim(txtCotitularApellidoMaterno.text)
        frmClientes.Mostrar CotitularEmp
    End If
End Sub

Private Sub txtCotitularApellidoMaterno2_GotFocus()
   Seleccionar_Texto txtCotitularApellidoMaterno2
   Cambiar_Color True, txtCotitularApellidoMaterno2
End Sub

Private Sub txtCotitularApellidoMaterno2_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCotitularApellidoMaterno2_LostFocus()
    Cambiar_Color False, txtCotitularApellidoMaterno2
    CotitularEmp.ApellidoMaterno = Trim(txtCotitularApellidoMaterno2.text)
    Titular = False
    If Trim(txtResponsable2.text) <> "" And Trim(txtCotitularApellidoPaterno2.text) <> "" And Trim(txtCotitularApellidoMaterno2.text) <> "" And Val(txtResponsable2.Tag) <= 0 Then Mostrar_Seleccionar_Cliente Trim(txtResponsable2.text), Trim(txtCotitularApellidoPaterno2.text), Trim(txtCotitularApellidoMaterno2.text), Me
    If Val(txtResponsable2.Tag) = 0 Then
        CotitularEmp.Nombre = Trim(txtResponsable2.text)
        CotitularEmp.ApellidoPaterno = Trim(txtCotitularApellidoPaterno2.text)
        CotitularEmp.ApellidoMaterno = Trim(txtCotitularApellidoMaterno2.text)
        frmClientes.Mostrar CotitularEmp
    End If
End Sub



'MLD-MODIF.
Private Sub cmdEditarCliente_Click()
    frmClientes.Mostrar ClienteEmp
    txtNombre.text = ClienteEmp.Nombre
    txtNombre.Tag = ClienteEmp.ID
    txtApellidoPaterno.text = ClienteEmp.ApellidoPaterno
    txtApellidoMaterno.text = ClienteEmp.ApellidoMaterno
    txtMensaje.text = ClienteEmp.Mensaje
    '--- Mostrar Datos Cliente ---
    If ClienteEmp.ID = 0 And (ClienteEmp.Nombre = "" And ClienteEmp.ApellidoPaterno = "" And ClienteEmp.ApellidoMaterno = "") Then
        lblDireccion.Caption = "": lblCiudad.Caption = "": lblRFC.Caption = ""
    Else
        lblDireccion.Caption = ClienteEmp.Direccion & IIf(ClienteEmp.NoExterior <> "", " #" & ClienteEmp.NoExterior, "") & IIf(ClienteEmp.NoInterior <> "", " INT." & ClienteEmp.NoInterior, "") & " COL." & ClienteEmp.Colonia & " C.P." & ClienteEmp.CodigoPostal
        lblCiudad.Caption = ClienteEmp.Municipio & ", " & ClienteEmp.Estado
        lblRFC.Caption = ClienteEmp.RFC
    End If
End Sub

'MLD-MODIF.
Private Sub cmdEditarCliente2_Click()
    frmClientes.Mostrar ClienteEmp
    txtNombre2.text = ClienteEmp.Nombre
    txtNombre2.Tag = ClienteEmp.ID
    txtApellidoPaterno2.text = ClienteEmp.ApellidoPaterno
    txtApellidoMaterno2.text = ClienteEmp.ApellidoMaterno
    txtMensaje2.text = ClienteEmp.Mensaje
    '--- Mostrar Datos Cliente ---
    If ClienteEmp.ID = 0 And (ClienteEmp.Nombre = "" And ClienteEmp.ApellidoPaterno = "" And ClienteEmp.ApellidoMaterno = "") Then
        lblDireccion2.Caption = "": lblCiudad2.Caption = "": lblRFC2.Caption = ""
    Else
        lblDireccion2.Caption = ClienteEmp.Direccion & IIf(ClienteEmp.NoExterior <> "", " #" & ClienteEmp.NoExterior, "") & IIf(ClienteEmp.NoInterior <> "", " INT." & ClienteEmp.NoInterior, "") & " COL." & ClienteEmp.Colonia & " C.P." & ClienteEmp.CodigoPostal
        lblCiudad2.Caption = ClienteEmp.Municipio & ", " & ClienteEmp.Estado
        lblRFC2.Caption = ClienteEmp.RFC
    End If
End Sub

'MLD-MODIF.
Private Sub cmdEditarCotitular_Click()
    frmClientes.Mostrar CotitularEmp
    txtResponsable.text = CotitularEmp.Nombre
    txtResponsable.Tag = CotitularEmp.ID
    txtCotitularApellidoPaterno.text = CotitularEmp.ApellidoPaterno
    txtCotitularApellidoMaterno.text = CotitularEmp.ApellidoMaterno
End Sub

'MLD-MODIF.
Private Sub cmdEditarCotitular2_Click()
    frmClientes.Mostrar CotitularEmp
    txtResponsable2.text = CotitularEmp.Nombre
    txtResponsable2.Tag = CotitularEmp.ID
    txtCotitularApellidoPaterno2.text = CotitularEmp.ApellidoPaterno
    txtCotitularApellidoMaterno2.text = CotitularEmp.ApellidoMaterno
End Sub

'---------------------------------------------------
Private Sub cmdMosCliente_Click()
    Titular = True
    '***Puntos***
    Limpiar_Tarjeta
    frmMostrarCliente.Ver Me, txtNombre, True, 0
End Sub

Private Sub cmdMosCliente2_Click()
    Titular = True
    frmMostrarCliente.Ver Me, txtNombre2, True, 0
End Sub


Private Sub cmdMosCotitular_Click()
    Titular = False
    frmMostrarCliente.Ver Me, txtResponsable, True, 0
End Sub

Private Sub cmdMosCotitular2_Click()
    Titular = False
    frmMostrarCliente.Ver Me, txtResponsable2, True, 0
End Sub
'---------------------------------------------------

'---------------------------------------------------
Private Sub cmdAlerta_Click()
    Dim pIdTipoAlerta As Integer, pDescAlerta As String
    
    pIdTipoAlerta = vTipoAlerta.ID: pDescAlerta = vTipoAlerta.Descripcion
    frmSelecAlertaLavado.Mostrar pIdTipoAlerta, pDescAlerta, MLD_PRESTAMO
    vTipoAlerta.ID = pIdTipoAlerta: vTipoAlerta.Descripcion = pDescAlerta
End Sub

Private Sub cmdAlerta2_Click()
    Dim pIdTipoAlerta As Integer, pDescAlerta As String
    
    pIdTipoAlerta = vTipoAlerta.ID: pDescAlerta = vTipoAlerta.Descripcion
    frmSelecAlertaLavado.Mostrar pIdTipoAlerta, pDescAlerta, MLD_PRESTAMO
    vTipoAlerta.ID = pIdTipoAlerta: vTipoAlerta.Descripcion = pDescAlerta
End Sub
'----------------------------------------------------

Private Sub cmdHistorial_Click()

    If Trim(txtNombre.text) = "" Then
        
        MsgBox "Seleccione el cliente del que desea consultar su historial !!", vbInformation, "Empeño"
        txtNombre.SetFocus
    Else
        
        frmHistorial.Ver Me, txtNombre, True, 1, CLng(txtNombre.Tag)
    End If

End Sub

Private Sub cmdHistorial2_Click()

    If Trim(txtNombre2.text) = "" Then
        
        MsgBox "Seleccione el cliente del que desea consultar su historial !!", vbInformation, "Empeño"
        txtNombre2.SetFocus
    Else
        
        frmHistorial.Ver Me, txtNombre2, True, 1, CLng(txtNombre2.Tag)
    End If

End Sub

'-----------------------------------------------------
'MLD-MODIF. - Buscamos el id cliente
Public Sub Buscar(ID As Long, Optional Tarjeta As Boolean = False)
Dim Semaforo As String
On Error GoTo Error


    If Titular Then
        ClienteEmp.Buscar ID
        
        If Not ClienteEmp.Valida Then
            frmClientes.Mostrar ClienteEmp
        End If
        Select Case TPestañas.SelectedTab
            Case 1
                txtNombre.text = ClienteEmp.Nombre
                txtNombre.Tag = ClienteEmp.ID
                txtApellidoPaterno.text = ClienteEmp.ApellidoPaterno
                txtApellidoMaterno.text = ClienteEmp.ApellidoMaterno
                txtMensaje.text = ClienteEmp.Mensaje
                cmdEditarCliente.Visible = True
                cmdAlerta.Enabled = True
                '--- Mostrar Datos Cliente ---
                lblDireccion.Caption = ClienteEmp.Direccion & IIf(ClienteEmp.NoExterior <> "", " #" & ClienteEmp.NoExterior, "") & IIf(ClienteEmp.NoInterior <> "", " INT." & ClienteEmp.NoInterior, "") & " COL." & ClienteEmp.Colonia & " C.P." & ClienteEmp.CodigoPostal
                lblCiudad.Caption = ClienteEmp.Municipio & ", " & ClienteEmp.Estado
                lblRFC.Caption = ClienteEmp.RFC
    
            Case 2
                txtNombre2.text = ClienteEmp.Nombre
                txtNombre2.Tag = ClienteEmp.ID
                txtApellidoPaterno2.text = ClienteEmp.ApellidoPaterno
                txtApellidoMaterno2.text = ClienteEmp.ApellidoMaterno
                txtMensaje2.text = ClienteEmp.Mensaje
                cmdEditarCliente2.Visible = True
                cmdAlerta2.Enabled = True
                '--- Mostrar Datos Cliente ---
                lblDireccion2.Caption = ClienteEmp.Direccion & IIf(ClienteEmp.NoExterior <> "", " #" & ClienteEmp.NoExterior, "") & IIf(ClienteEmp.NoInterior <> "", " INT." & ClienteEmp.NoInterior, "") & " COL." & ClienteEmp.Colonia & " C.P." & ClienteEmp.CodigoPostal
                lblCiudad2.Caption = ClienteEmp.Municipio & ", " & ClienteEmp.Estado
                lblRFC2.Caption = ClienteEmp.RFC
        End Select
        
        '**** TARJETA DE PUNTOS ****
        Semaforo = Regresa_Semaforo(ID)
    
        If Semaforo = "Verde" Then
            ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\VERDE.bmp")
            ImgSemaforo.Tag = SacaValor("parametros", "PrestamoVerde", "")
        ElseIf Semaforo = "Amarillo" Then
            ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\AMARILLO.bmp")
            ImgSemaforo.Tag = SacaValor("parametros", "PrestamoAmarillo", "")
        Else
            ImgSemaforo.Picture = LoadPicture(App.Path & "\Fotos\ROJO.bmp")
            ImgSemaforo.Tag = SacaValor("parametros", "PrestamoRojo", "")
        End If
        
        '***Puntos***
        'si no fue buscado por la tarjeta de puntos
        If SacaValor("tarjetaspuntos", "count(id)", " where activa = 1") > 0 Then
        
            If Not Tarjeta Then
              If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(ID) Then
                 lblPuntosAcumulados.Caption = TarjetaPuntos.CuentaFrecuente.Puntos
                 txtNoTarjeta.text = TarjetaPuntos.CuentaFrecuente.Folio
              Else
                 If MsgBox("El Cliente no cuenta con tarjeta de cliente frecuente" & vbCrLf & "Desea asignarle una tarjeta?", vbYesNoCancel Or vbQuestion) = vbYes Then
                    TarjetaPuntos.ShowAsignarTarjeta ID, frmMDI.IDUsuario
                    
                    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(ID) Then
                        lblPuntosAcumulados.Caption = TarjetaPuntos.CuentaFrecuente.Puntos
                        txtNoTarjeta.text = TarjetaPuntos.CuentaFrecuente.Folio
                    Else
                        MsgBox "No se agregó la tarjeta al cliente", vbCritical, "Empeño"
                    End If
                    
                 End If
              End If
            End If
        End If
        '***************************
        
    Else
        CotitularEmp.Buscar ID
        
        If Not CotitularEmp.Valida Then
            frmClientes.Mostrar CotitularEmp
        End If
        Select Case TPestañas.SelectedTab
            Case 1
                txtResponsable.text = CotitularEmp.Nombre
                txtResponsable.Tag = CotitularEmp.ID
                txtCotitularApellidoPaterno.text = CotitularEmp.ApellidoPaterno
                txtCotitularApellidoMaterno.text = CotitularEmp.ApellidoMaterno
                cmdEditarCotitular.Visible = True
            Case 2
                txtResponsable2.text = CotitularEmp.Nombre
                txtResponsable2.Tag = CotitularEmp.ID
                txtCotitularApellidoPaterno2.text = CotitularEmp.ApellidoPaterno
                txtCotitularApellidoMaterno2.text = CotitularEmp.ApellidoMaterno
                cmdEditarCotitular2.Visible = True
        End Select
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub


'MLD-MODIF.
Private Sub txtModelo_GotFocus()
    Seleccionar_Texto txtModelo
    Cambiar_Color True, txtModelo
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtModelo_LostFocus()
    Cambiar_Color False, txtModelo
End Sub

'MLD-MODIF.
Private Sub txtREPUVE_GotFocus()
    Seleccionar_Texto txtREPUVE
    Cambiar_Color True, txtREPUVE
End Sub

Private Sub txtREPUVE_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtREPUVE_LostFocus()
    Cambiar_Color False, txtREPUVE
End Sub

'MLD-MODIF.
Private Sub txtVIN_GotFocus()
    Seleccionar_Texto txtVIN
    Cambiar_Color True, txtVIN
End Sub

Private Sub txtVIN_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVIN_LostFocus()
    Cambiar_Color False, txtVIN
End Sub

'MLD-MODIF.
Private Sub cmbTipoBlindaje_GotFocus()
    Cambiar_Color True, cmbTipoBlindaje
End Sub

Private Sub cmbTipoBlindaje_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoBlindaje_LostFocus()
    Cambiar_Color False, cmbTipoBlindaje
End Sub

Private Sub txtCostoMensualSeguro_GotFocus()
    Cambiar_Color True, txtCostoMensualSeguro
    Seleccionar_Texto txtCostoMensualSeguro
End Sub

Private Sub txtCostoMensualSeguro_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then cmdAceptar.SetFocus
End Sub

Private Sub txtCostoMensualSeguro_LostFocus()
    Cambiar_Color False, txtCostoMensualSeguro
    txtCostoMensualSeguro = Format(txtCostoMensualSeguro.text, FMoneda)
End Sub


Private Sub txtCargoGPS_GotFocus()
    Cambiar_Color True, txtCargoGPS
    Seleccionar_Texto txtCargoGPS
End Sub

Private Sub txtCargoGPS_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then cmdAceptar.SetFocus
End Sub

Private Sub txtCargoGPS_LostFocus()
    Cambiar_Color False, txtCargoGPS
    txtCargoGPS = Format(txtCargoGPS.text, FMoneda)
    Poner_Totales_Refrendo
End Sub


Private Sub txtCargoSeguro_GotFocus()
    Cambiar_Color True, txtCargoSeguro
    Seleccionar_Texto txtCargoSeguro
    
End Sub

Private Sub txtCargoSeguro_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then cmdAceptar.SetFocus
End Sub

Private Sub txtCargoSeguro_LostFocus()
    Cambiar_Color False, txtCargoSeguro
    txtCargoSeguro = Format(txtCargoSeguro.text, FMoneda)
    Poner_Totales_Refrendo
End Sub

Private Sub txtCargoSeguroDes_GotFocus()
    Cambiar_Color True, txtCargoSeguroDes
    Seleccionar_Texto txtCargoSeguroDes
    
End Sub

Private Sub txtCargoSeguroDes_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then cmdAceptar.SetFocus
End Sub

Private Sub txtCargoSeguroDes_LostFocus()
    Cambiar_Color False, txtCargoSeguroDes
    txtCargoSeguroDes = Format(txtCargoSeguroDes.text, FMoneda)
    Poner_Totales_Desempeño
End Sub

Private Sub txtCargoGPSDes_GotFocus()
    Cambiar_Color True, txtCargoGPSDes
    Seleccionar_Texto txtCargoGPSDes
    
End Sub

Private Sub txtCargoGPSDes_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then cmdAceptar.SetFocus
End Sub

Private Sub txtCargoGPSDes_LostFocus()
    Cambiar_Color False, txtCargoGPSDes
    txtCargoGPSDes = Format(txtCargoGPSDes.text, FMoneda)
    Poner_Totales_Desempeño
End Sub
