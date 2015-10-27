VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{0BFA85A1-F9B8-11CF-8939-444553540000}#1.0#0"; "barcode.ocx"
Begin VB.Form frmCompras 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras/Dotaciones a Inventario"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14580
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   14580
   Begin VB.Frame Frame2 
      Height          =   2235
      Left            =   15
      TabIndex        =   26
      Top             =   120
      Width           =   14520
      Begin VB.TextBox txtApellidoMaterno 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3135
         MaxLength       =   70
         TabIndex        =   2
         Top             =   1080
         Width           =   2850
      End
      Begin VB.TextBox txtApellidoPaterno 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   135
         MaxLength       =   60
         TabIndex        =   1
         Top             =   1080
         Width           =   2850
      End
      Begin VB.TextBox txtNombre 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   120
         MaxLength       =   20
         TabIndex        =   0
         Top             =   480
         Width           =   2850
      End
      Begin VB.CommandButton cmdEditar 
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
         Left            =   3720
         TabIndex        =   75
         Top             =   360
         Width           =   990
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   8520
         TabIndex        =   59
         Top             =   840
         Width           =   4665
         Begin VB.ComboBox cmbMovimiento 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            ItemData        =   "frmCompras.frx":000C
            Left            =   315
            List            =   "frmCompras.frx":0016
            Style           =   2  'Dropdown List
            TabIndex        =   64
            Top             =   390
            Width           =   4050
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            Caption         =   "MOVIMIENTO"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   345
            Left            =   975
            TabIndex        =   61
            Top             =   0
            Width           =   2190
         End
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
         Height          =   225
         Left            =   3120
         TabIndex        =   27
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
      Begin VB.Label Label7 
         Alignment       =   2  'Center
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
         Left            =   6135
         TabIndex        =   84
         Top             =   840
         Width           =   2205
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
         Left            =   6135
         TabIndex        =   83
         Top             =   1080
         Width           =   2205
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
         Left            =   1815
         TabIndex        =   82
         Top             =   1800
         Width           =   4125
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFC0C0&
         Caption         =   "Ciudad"
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
         Left            =   135
         TabIndex        =   81
         Top             =   1800
         Width           =   1635
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
         Left            =   1815
         TabIndex        =   80
         Top             =   1440
         Width           =   6525
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
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
         Left            =   135
         TabIndex        =   79
         Top             =   1440
         Width           =   1635
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   135
         TabIndex        =   78
         Top             =   240
         Width           =   2850
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   15
         Left            =   135
         TabIndex        =   77
         Top             =   840
         Width           =   2850
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000000&
         Height          =   210
         Index           =   16
         Left            =   3135
         TabIndex        =   76
         Top             =   840
         Width           =   2850
      End
      Begin VB.Label lblFolio 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   345
         Left            =   10755
         TabIndex        =   66
         Top             =   375
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "FOLIO:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9585
         TabIndex        =   65
         Top             =   360
         Width           =   1110
      End
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCompras 
      Height          =   3990
      Left            =   4050
      TabIndex        =   25
      Top             =   2400
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   7038
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   -2147483626
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
      DisableIcons    =   -1  'True
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   13365
      TabIndex        =   28
      Top             =   6840
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
      Picture         =   "frmCompras.frx":0039
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   12225
      TabIndex        =   29
      Top             =   6840
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
      Picture         =   "frmCompras.frx":058B
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   10905
      TabIndex        =   30
      Top             =   6840
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
      Picture         =   "frmCompras.frx":0ADD
   End
   Begin vbalTabStrip6.TabControl TPrendas 
      Height          =   4350
      Left            =   0
      TabIndex        =   31
      Top             =   2400
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   7673
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
      Begin DevPowerFlatBttn.FlatBttn cmdDiamante 
         Height          =   375
         Left            =   1020
         TabIndex        =   70
         Top             =   3900
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
         Picture         =   "frmCompras.frx":102F
         PictureDisabled =   "frmCompras.frx":1253
      End
      Begin VB.Frame frmMetales 
         Caption         =   "Metales"
         Height          =   3570
         Left            =   45
         TabIndex        =   32
         Top             =   315
         Width           =   3990
         Begin VB.TextBox txtPesoPiedra 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1065
            MaxLength       =   20
            TabIndex        =   72
            Top             =   1335
            Width           =   1020
         End
         Begin VB.TextBox txtPiedras 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2895
            MaxLength       =   3
            TabIndex        =   71
            Top             =   1335
            Width           =   1020
         End
         Begin VB.TextBox txtPeso 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2895
            MaxLength       =   20
            TabIndex        =   8
            Top             =   1035
            Width           =   1020
         End
         Begin VB.TextBox txtPrecio 
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
            TabIndex        =   10
            Top             =   1635
            Width           =   1020
         End
         Begin VB.ComboBox cmbKilates 
            Height          =   315
            ItemData        =   "frmCompras.frx":13AD
            Left            =   1035
            List            =   "frmCompras.frx":13AF
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   690
            Width           =   1095
         End
         Begin VB.TextBox txtCantidad 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1065
            MaxLength       =   3
            TabIndex        =   7
            Top             =   1035
            Width           =   1020
         End
         Begin VB.TextBox txtCosto 
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
            Left            =   1065
            MaxLength       =   20
            TabIndex        =   9
            Top             =   1635
            Width           =   1020
         End
         Begin VB.TextBox txtDescripcion 
            BorderStyle     =   0  'None
            Height          =   1275
            Left            =   60
            MaxLength       =   150
            MultiLine       =   -1  'True
            TabIndex        =   11
            Top             =   2205
            Width           =   3870
         End
         Begin VB.ComboBox cmbPrenda 
            Height          =   315
            ItemData        =   "frmCompras.frx":13B1
            Left            =   1035
            List            =   "frmCompras.frx":13B3
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   375
            Width           =   2925
         End
         Begin VB.ComboBox cmbTipo 
            Height          =   315
            ItemData        =   "frmCompras.frx":13B5
            Left            =   1035
            List            =   "frmCompras.frx":13B7
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   60
            Width           =   1725
         End
         Begin VB.ComboBox cmbEstado 
            Height          =   315
            ItemData        =   "frmCompras.frx":13B9
            Left            =   2865
            List            =   "frmCompras.frx":13BB
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   690
            Width           =   1095
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
            TabIndex        =   74
            Top             =   1365
            Width           =   675
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
            TabIndex        =   73
            Top             =   1365
            Width           =   1035
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
            TabIndex        =   46
            Top             =   435
            Width           =   645
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
            TabIndex        =   45
            Top             =   1065
            Width           =   795
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
            TabIndex        =   44
            Top             =   1065
            Width           =   450
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
            TabIndex        =   43
            Top             =   750
            Width           =   690
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Costo:"
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
            Left            =   15
            TabIndex        =   42
            Top             =   1650
            Width           =   525
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Características:"
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
            Index           =   9
            Left            =   30
            TabIndex        =   41
            Top             =   1920
            Width           =   1320
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
            TabIndex        =   40
            Top             =   120
            Width           =   405
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
            TabIndex        =   39
            Top             =   750
            Width           =   615
         End
         Begin VB.Label lblPiedra 
            BackColor       =   &H80000013&
            Height          =   255
            Left            =   2070
            TabIndex        =   38
            Top             =   1815
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblCantidadPiedras 
            BackColor       =   &H80000013&
            Caption         =   "0"
            Height          =   255
            Left            =   3510
            TabIndex        =   37
            Top             =   1815
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblPuntos 
            BackColor       =   &H80000013&
            Caption         =   "0"
            Height          =   255
            Left            =   2820
            TabIndex        =   36
            Top             =   1815
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label lblPrestamoDiamante 
            BackColor       =   &H80000013&
            Caption         =   "0"
            Height          =   255
            Left            =   3255
            TabIndex        =   35
            Top             =   1050
            Visible         =   0   'False
            Width           =   555
         End
         Begin VB.Label lblAvaluoDiamante 
            BackColor       =   &H80000013&
            Caption         =   "0"
            Height          =   255
            Left            =   3540
            TabIndex        =   34
            Top             =   1470
            Visible         =   0   'False
            Width           =   600
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Precio:"
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
            TabIndex        =   33
            Top             =   1650
            Width           =   570
         End
      End
      Begin VB.Frame frmElectronicos 
         Caption         =   "Electronicos"
         Height          =   3570
         Left            =   45
         TabIndex        =   47
         Top             =   315
         Width           =   3990
         Begin VB.CheckBox chkEtiqueta 
            Appearance      =   0  'Flat
            Caption         =   "Etiqueta"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   50
            TabIndex        =   69
            Top             =   50
            Width           =   900
         End
         Begin VB.ComboBox cmbTipoElec 
            Height          =   315
            ItemData        =   "frmCompras.frx":13BD
            Left            =   990
            List            =   "frmCompras.frx":13BF
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   60
            Width           =   1770
         End
         Begin VB.TextBox txtDescripcionElec 
            BorderStyle     =   0  'None
            Height          =   1335
            Left            =   60
            MaxLength       =   250
            MultiLine       =   -1  'True
            TabIndex        =   24
            Top             =   2175
            Width           =   3870
         End
         Begin VB.TextBox txtCostoElec 
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
            Left            =   1015
            MaxLength       =   20
            TabIndex        =   22
            Top             =   1635
            Width           =   1020
         End
         Begin VB.TextBox txtPrecioElec 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Enabled         =   0   'False
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
            TabIndex        =   23
            Top             =   1635
            Width           =   1095
         End
         Begin VB.TextBox txtModeloElec 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2895
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   18
            Top             =   727
            Width           =   1050
         End
         Begin VB.TextBox txtNumSerieElec 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1015
            MaxLength       =   80
            TabIndex        =   21
            Top             =   1335
            Width           =   2925
         End
         Begin VB.TextBox txtTamañoElec 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1015
            MaxLength       =   50
            TabIndex        =   19
            Top             =   1035
            Width           =   1020
         End
         Begin VB.TextBox txtColorElec 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   2895
            MaxLength       =   50
            TabIndex        =   20
            Top             =   1035
            Width           =   1050
         End
         Begin VB.TextBox txtFamiliaElec 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1015
            Locked          =   -1  'True
            MaxLength       =   80
            TabIndex        =   16
            Top             =   412
            Width           =   2925
         End
         Begin VB.TextBox txtMarcaElec 
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   1015
            Locked          =   -1  'True
            MaxLength       =   50
            TabIndex        =   17
            Top             =   727
            Width           =   1020
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMostrarCatPrendas 
            Height          =   270
            Left            =   2760
            TabIndex        =   48
            Top             =   90
            Width           =   300
            _ExtentX        =   529
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
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Precio:"
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
            TabIndex        =   58
            Top             =   1665
            Width           =   570
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
            TabIndex        =   57
            Top             =   120
            Width           =   405
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Características:"
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
            TabIndex        =   56
            Top             =   1950
            Width           =   1320
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Costo:"
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
            TabIndex        =   55
            Top             =   1665
            Width           =   525
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
            TabIndex        =   54
            Top             =   435
            Width           =   645
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
            TabIndex        =   53
            Top             =   750
            Width           =   570
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
            TabIndex        =   52
            Top             =   750
            Width           =   660
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
            TabIndex        =   51
            Top             =   1365
            Width           =   780
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
            TabIndex        =   50
            Top             =   1065
            Width           =   735
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
            TabIndex        =   49
            Top             =   1065
            Width           =   480
         End
      End
      Begin VB.Frame Frame3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1860
         Left            =   960
         TabIndex        =   67
         Top             =   720
         Width           =   3405
         Begin BARCODELib.Barcode bcCodigo 
            Height          =   1410
            Left            =   75
            TabIndex        =   68
            Top             =   360
            Width           =   3300
            _Version        =   65536
            _ExtentX        =   5821
            _ExtentY        =   2487
            _StockProps     =   25
            Text            =   "12345678901212"
            TypeName        =   "EAN 13"
            Text            =   "12345678901212"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Borderwidth     =   0
            Borderheight    =   5
            NotchHeightInPercent=   15
         End
      End
      Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
         Height          =   375
         Left            =   2085
         TabIndex        =   12
         Top             =   3900
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
         Picture         =   "frmCompras.frx":13C1
         PictureDisabled =   "frmCompras.frx":172B
      End
      Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
         Height          =   375
         Left            =   75
         TabIndex        =   14
         Top             =   3900
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
         Picture         =   "frmCompras.frx":1885
      End
      Begin DevPowerFlatBttn.FlatBttn cmdBorrar 
         Height          =   375
         Left            =   3030
         TabIndex        =   13
         Top             =   3900
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
         Picture         =   "frmCompras.frx":1989
         PictureDisabled =   "frmCompras.frx":1EDB
      End
   End
   Begin VB.Label lblTotCosto 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   8580
      TabIndex        =   63
      Top             =   6450
      Width           =   420
   End
   Begin VB.Label lblTotPrecio 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
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
      Left            =   9765
      TabIndex        =   62
      Top             =   6450
      Width           =   420
   End
   Begin VB.Label lblTotalCosto 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00008080&
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
      Height          =   330
      Left            =   4050
      TabIndex        =   60
      Top             =   6405
      Width           =   10485
   End
End
Attribute VB_Name = "frmCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim BanElec As Boolean, pIDEmpeno As Long, pFolio As Long, Ban As Boolean
Dim ClienteCom As clientes

Public Property Let IDEmpeno(Valor As Integer)
    pIDEmpeno = Valor
End Property

Public Property Get IDEmpeno() As Integer
    IDEmpeno = pIDEmpeno
End Property

Public Property Let Folio(Valor As Long)
    pFolio = Valor
End Property

Public Property Get Folio() As Long
    Folio = pFolio
End Property

Private Sub cmbEstado_Click()
    Calcula_Costo
End Sub

Private Sub cmbEstado_GotFocus()
    Cambiar_Color True, cmbEstado
End Sub

Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbEstado_LostFocus()
    Cambiar_Color False, cmbEstado
End Sub

Private Sub cmbKilates_Click()
    Calcula_Costo
End Sub

Private Sub cmbKilates_GotFocus()
    Cambiar_Color True, cmbKilates
End Sub

Private Sub cmbKilates_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbKilates_LostFocus()
    Cambiar_Color False, cmbKilates
End Sub

Private Sub cmbMovimiento_GotFocus()
    Cambiar_Color True, cmbMovimiento
End Sub

Private Sub cmbMovimiento_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbMovimiento_LostFocus()
    Cambiar_Color False, cmbMovimiento
End Sub

Private Sub cmbPrenda_Click()
Dim strPrenda As String, IDPrenda As Integer

    If cmbPrenda.text = "[1. AGREGAR PRENDA]" Then
        IDPrenda = 0
        strPrenda = ""
        IDPrenda = frmAgregaPrendaOro.Mostrar()
        If IDPrenda > 0 Then
            
            cmbPrenda.Clear
            cmbPrenda.AddItem "[1. AGREGAR PRENDA]"
            Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Descripcion", False
            '''''Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda, , "Descripcion", False
            cmbPrenda.ListIndex = ComboInformacion(cmbPrenda, IDPrenda)
        
        Else
            
            cmbPrenda.ListIndex = -1
        End If
    
    End If
End Sub

Private Sub cmbPrenda_GotFocus()
    Cambiar_Color True, cmbPrenda
End Sub

Private Sub cmbPrenda_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPrenda_LostFocus()
    Cambiar_Color False, cmbPrenda
End Sub

Private Sub cmbTipo_Click()
        
    If cmbTipo.ListIndex > -1 Then
        
        cmbPrenda.AddItem "[1. AGREGAR PRENDA]"
        Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Descripcion", False
        Cargar_Combos "Descripcion", "kilatajes", cmbKilates, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Ordenamiento"
        Cargar_Combos "Estado", "estado", cmbEstado, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Ordenamiento"
    
    End If

End Sub

Private Sub cmbTipo_GotFocus()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
End Sub

Private Sub cmdAceptar_Click()

    If ValidaCompra Then
        
        If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Compras/Dotaciones a Inventario") = vbYes Then
            
            Screen.MousePointer = vbHourglass
            GrabarEntradas
            txtNombre.SetFocus
            Screen.MousePointer = vbDefault
        
        End If
        
    End If

End Sub

Private Sub cmdAgregar_Click()
'''''Dim i As Integer, Estado As Integer, Kilates As Integer
'''''
'''''    If ValidaArticulos(TPrendas.SelectedTab) Then
'''''
'''''        With grdCompras
'''''
'''''            For i = 1 To .Rows
'''''                If Val(.CellText(i, 2)) = 0 Then Exit For
'''''            Next i
'''''
'''''            If i = 12 Then MsgBox "Introduzca las piezas en otra compra !!", vbInformation, "Compras":  Exit Sub
'''''
'''''            .CellText(i, 1) = cmbTipo.text
'''''
'''''            .CellText(i, 2) = txtCantidad
'''''            .CellTextAlign(i, 2) = DT_RIGHT
'''''
'''''            .CellText(i, 3) = cmbPrenda.text
'''''            .CellItemData(i, 3) = cmbPrenda.ItemData(cmbPrenda.ListIndex)
'''''
'''''            'Kilates
'''''            Kilates = RegresaKilates(cmbKilates.text, cmbTipo.text)
'''''            .CellText(i, 4) = cmbKilates.text
'''''            .CellItemData(i, 4) = Kilates
'''''            .CellTextAlign(i, 4) = DT_CENTER
'''''
'''''            .CellText(i, 5) = txtPeso.text
'''''            .CellTextAlign(i, 5) = DT_RIGHT
'''''
'''''            .CellText(i, 6) = txtCosto.text
'''''            .CellTextAlign(i, 6) = DT_RIGHT
'''''
'''''            .CellText(i, 7) = txtPrecio.text
'''''            .CellTextAlign(i, 7) = DT_RIGHT
'''''
'''''            If cmbEstado.ListIndex >= 0 Then
'''''
'''''                Estado = cmbEstado.ItemData(cmbEstado.ListIndex)
'''''            Else
'''''
'''''                Estado = 0
'''''            End If
'''''            .CellText(i, 8) = cmbEstado.text
'''''            .CellItemData(i, 8) = Estado
'''''
'''''            .CellText(i, 9) = Trim(txtDescripcion.text)
'''''            .CellTextAlign(i, 9) = DT_RIGHT
'''''
'''''            Total_Costos
'''''            LimpiaArticulos
'''''            cmbTipo.SetFocus
'''''        End With
'''''    End If

Dim i As Integer, Estado As Integer, Kilates As Integer, PrecioVenta As Double, crPrestamo As Double
Dim IDTipo As Integer, IDTipoPrenda As Long, strPrenda As String, Piedras As Integer, PesoPiedras As Double, TipoMovimiento As Boolean
Dim strCodigoBarras As String

On Error GoTo Error

    If ValidaArticulos(TPrendas.SelectedTab) Then
    
        With grdCompras
            
            If Val(txtCantidad.Tag) > 0 And TPrendas.SelectedTab = 1 Then i = Val(txtCantidad.Tag): GoTo Edicion
            For i = 1 To .Rows

                If Val(.CellText(i, 2)) = 0 Then Exit For
            Next i
            
Edicion:
            If i = 41 Then MsgBox "No se pueden agregar más prendas !!", vbInformation, "Compras":  Exit Sub
            
            If Ban = False Then
                Ban = True
                Me.Folio = 0
                Me.IDEmpeno = 0
                CreaEncabezado
            End If
                                                                        
            'Saco el código de barras
            Select Case cmbMovimiento.ListIndex
            Case 0
                
                TipoMovimiento = True
            Case 1
                
                TipoMovimiento = False
            End Select
            strCodigoBarras = CreaCodigoBarras(Trim(Format(frmMDI.IDSucursal, "000")), IIf(TipoMovimiento, ENTRADACOMPRA, ENTRADADOTACION), Trim(Me.Folio), i)
            bcCodigo.text = ""
            bcCodigo.text = Mid(strCodigoBarras, 1, 12)
                
            If TPrendas.SelectedTab = 1 Then
            
                IDTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
            Else
            
                IDTipo = cmbTipoElec.ItemData(cmbTipoElec.ListIndex)
            End If
            .CellText(i, 1) = IIf(TPrendas.SelectedTab = 1, cmbTipo.text, cmbTipoElec.text)
            .CellItemData(i, 1) = IDTipo
        
            .CellText(i, 2) = IIf(TPrendas.SelectedTab = 1, Val(txtCantidad), "1")
            .CellTextAlign(i, 2) = DT_CENTER
            
            If TPrendas.SelectedTab = 1 Then
                
                IDTipoPrenda = cmbPrenda.ItemData(cmbPrenda.ListIndex)
                strPrenda = cmbPrenda.text & " " & lblPiedra.Caption
            Else
                
                IDTipoPrenda = Val(txtFamiliaElec.Tag)
                strPrenda = txtFamiliaElec.text
            End If
            
            .CellText(i, 3) = strPrenda
            .CellItemData(i, 3) = IDTipoPrenda
        
            .CellText(i, 4) = IIf(TPrendas.SelectedTab = 1, txtPeso.text, 0)
            .CellTextAlign(i, 4) = DT_RIGHT
            
            If TPrendas.SelectedTab = 1 Then
                
                Kilates = RegresaKilates(cmbKilates.text, cmbTipo.text)
            Else
                
                Kilates = 0
            End If
            .CellText(i, 5) = IIf(TPrendas.SelectedTab = 1, cmbKilates.text, "")
            .CellTextAlign(i, 5) = DT_RIGHT
            .CellItemData(i, 5) = Kilates
        
            .CellText(i, 6) = IIf(TPrendas.SelectedTab = 1, txtCosto.text, txtCostoElec.text)
            .CellTextAlign(i, 6) = DT_RIGHT
        
            .CellText(i, 7) = IIf(TPrendas.SelectedTab = 1, txtPrecio.text, txtPrecioElec.text)
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
        
            .CellText(i, 11) = IIf(TPrendas.SelectedTab = 1, txtDescripcion.text, txtDescripcionElec.text)
                                    
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
            
            'Mando a imprimir el código de barras
            If MsgBox("Desea imprimir la etiqueta ??", vbQuestion + vbYesNo + vbDefaultButton1, "Compras/Dotaciones a Inventario") = vbYes Then
                
                If TPrendas.SelectedTab = 1 Or (TPrendas.SelectedTab = 2 And chkEtiqueta.Value = vbChecked) Then
                    
                    .CellText(i, 22) = strCodigoBarras
                    Imprimir strCodigoBarras, .CellText(i, 4), CDbl(.CellText(i, 7)), Kilates, .CellText(i, 2), strPrenda
                
                End If
            
            End If
            
            Total_Costos
            .ClearSelection
            LimpiaArticulos
            If TPrendas.SelectedTab = 1 Then cmbTipo.ListIndex = 0 Else cmbTipoElec.ListIndex = 0
                                    
            txtCantidad.text = "1"
            If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
            
        End With

    End If
    Exit Sub
    
Error:
    Maneja_Error Err
    
End Sub

Private Sub cmdBorrar_Click()
Dim Res As Integer
Dim i As Integer

    If grdCompras.SelectedRow > 0 Then
            
        If Trim(grdCompras.CellText(grdCompras.SelectedRow, 1)) <> "" Then
            
            If MsgBox("Desea eliminar la prenda seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Compras") = vbNo Then
                
                grdCompras.ClearSelection
                If TPrendas.SelectedTab = 1 Then cmbTipo.ListIndex = 0 Else cmbTipoElec.ListIndex = 0
                Exit Sub
            End If
            
            grdCompras.RemoveRow grdCompras.SelectedRow
            Res = 12 - grdCompras.Rows
            For i = 1 To Res
                grdCompras.AddRow
            Next i
            grdCompras.ClearSelection
            Total_Costos
            If TPrendas.SelectedTab = 1 Then cmbTipo.ListIndex = 0 Else cmbTipoElec.ListIndex = 0
        
        End If
        
    End If
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

Private Sub cmdImprimir_Click()
Dim Folio As Long, TipoMovimiento As Integer

    Folio = frmReimpresionrecibos.ReImprimir("entradainventario", "Folio", " WHERE Folio=")
    If Folio > 0 Then
        
        TipoMovimiento = SacaValor("entradainventario", "TipoEntrada", " WHERE Folio=" & Folio)
                    
        If TipoMovimiento = ENTRADACOMPRA Then
            
            ImprimirTicket Folio
            
        ElseIf TipoMovimiento = ENTRADADOTACION Then
            
            ImprimirDotacion SacaValor("entradainventario", "ID", " WHERE Folio=" & Folio)
            
        End If
        
    Else
            
        MsgBox "No se encontró el folio especificado !!", vbInformation, "Compras/Dotaciones a Inventario"
    End If
    
End Sub


Private Sub cmdMostrarCatPrendas_Click()
Dim IDPrenda As Long
Dim rcPrenda As New ADODB.Recordset
        
    IDPrenda = frmCatVarios.Mostrar(cmbTipoElec.ItemData(cmbTipoElec.ListIndex))
    If IDPrenda > 0 Then
        
        LimpiaArticulos
        With rcPrenda
        
            .Open "SELECT prendaselec.IDTipo,tipoprenda.Descripcion AS Desc_Familia,marcas.Descripcion AS Desc_Marca,prendaselec.ID AS IDPrenda,prendaselec.Modelo,prendaselec.Minimo,prendaselec.Maximo FROM prendaselec INNER JOIN tipoprenda ON prendaselec.IDFamilia=tipoprenda.ID INNER JOIN marcas ON prendaselec.IDMarca=marcas.ID WHERE prendaselec.ID=" & IDPrenda, dbDatos, adOpenForwardOnly, adLockOptimistic
                        
            txtFamiliaElec.text = !Desc_Familia
            txtFamiliaElec.Tag = !IDPrenda
            txtMarcaElec.text = !Desc_Marca
            txtModeloElec.text = !Modelo
            txtCostoElec.Tag = !Maximo
            txtCostoElec.text = Format(!Minimo, FMoneda)
            
            .Close
            Set rcPrenda = Nothing
            
        End With
        
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    frmMetales.BorderStyle = 0
    frmElectronicos.BorderStyle = 0
    Crear_Pestañas
    Cargar_Combos "Descripcion", "tipo", cmbTipo, " WHERE ID=1", "Ordenamiento"
    Cargar_Combos "Descripcion", "tipo", cmbTipoElec, " WHERE ID<>1", "Ordenamiento"
    
    'MLD-MODIF.
    Set ClienteCom = New clientes
    ClienteCom.FechaExpiracion = "1900-01-01"
    ClienteCom.FechaNacimiento = "1900-01-01"
    ClienteCom.FechaAltaRazonSocial = "1900-01-01"
    
    Ban = False
    cmbTipo.ListIndex = 0
    Crear_Encabezados
    BanElec = False
    txtCantidad.text = "1"
    lblFolio.Caption = Regresa_Movimiento(False, "FolioInventario")
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

'Calculamos el prestamo
Private Function Calcula_Costo() As Double
Dim crCosto As Double, crPrecio As Double, Peso As Double, PesoPiedra As Double, PrestamoDiamante As Double, AvaluoDiamante As Double

On Error GoTo Error
   
    If cmbTipo.ListIndex >= 0 And cmbKilates.ListIndex >= 0 And cmbEstado.ListIndex >= 0 Then
        
        If Val(txtPeso.text) > 0 Or (Trim(txtPeso.text) <> "" And Trim(txtPeso.text) <> ".") Then
            
            Peso = txtPeso.text
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
        
        crCosto = Regresa_Valor_BD(cmbKilates.text)
        crPrecio = Regresa_Valor_BD("Venta" & cmbKilates.text)
        
        Calcula_Costo = (Peso - PesoPiedra) * crCosto
             
        txtCosto.text = Format(Redondeo(CCur(Calcula_Costo) + PrestamoDiamante), FMoneda)
        txtPrecio.text = Format(Redondeo((CCur(crPrecio * (Peso - PesoPiedra)) + AvaluoDiamante)), FMoneda)
    End If
    Exit Function
    
Error:
    Maneja_Error Err
End Function

Private Sub grdCompras_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Dim Res As Integer, i As Integer

    If grdCompras.SelectedRow > 0 And KeyCode = vbKeyDelete Then
            
        If Trim(grdCompras.CellText(grdCompras.SelectedRow, 1)) <> "" Then
            
            If MsgBox("Desea eliminar la prenda seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Compras") = vbNo Then
                
                grdCompras.ClearSelection
                If TPrendas.SelectedTab = 1 Then cmbTipo.ListIndex = 0 Else cmbTipoElec.ListIndex = 0
                Exit Sub
            End If
            
            grdCompras.RemoveRow grdCompras.SelectedRow
            Res = 40 - grdCompras.Rows
            For i = 1 To Res
                grdCompras.AddRow
            Next i
            grdCompras.ClearSelection
            Total_Costos
            If TPrendas.SelectedTab = 1 Then cmbTipo.ListIndex = 0 Else cmbTipoElec.ListIndex = 0
        
        End If
        
    End If
End Sub

Private Sub lblPrestamoDiamante_Change()
    Calcula_Costo
End Sub

Private Sub tPrendas_TabClick(ByVal lTab As Long)

    Select Case lTab

        Case 1
            
            BanElec = False
            LimpiaArticulos
            txtCantidad.text = "1"
            frmMetales.Visible = True
            frmElectronicos.Visible = False
            cmbTipo.ListIndex = 0
        Case 2
            
            BanElec = True
            LimpiaArticulos
            frmElectronicos.Visible = True
            frmMetales.Visible = False
            cmbTipoElec.ListIndex = 0
    End Select
End Sub


Private Sub txtCantidad_GotFocus()
    Seleccionar_Texto txtCantidad
    Cambiar_Color True, txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCantidad_LostFocus()
    Cambiar_Color False, txtCantidad
End Sub

Private Sub txtCosto_GotFocus()
    Seleccionar_Texto txtCosto
    Cambiar_Color True, txtCosto
End Sub

Private Sub txtCosto_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCosto_LostFocus()
    Cambiar_Color False, txtCosto
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Texto txtDescripcion
    Cambiar_Color True, txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescripcion_LostFocus()
    Cambiar_Color False, txtDescripcion
End Sub

Private Sub txtPeso_Change()
    Calcula_Costo
End Sub

Private Sub txtPeso_GotFocus()
    Seleccionar_Texto txtPeso
    Cambiar_Color True, txtPeso
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPeso_LostFocus()
    Cambiar_Color False, txtPeso
End Sub

Private Sub Crear_Encabezados()

    With grdCompras
'''''        .AddColumn "K1", "Tipo", ecgHdrTextALignRight, , 80, False, , , , , , CCLSortString
'''''        .AddColumn "K2", "Cantidad", ecgHdrTextALignRight, , 55, , , , , , , CCLSortString
'''''        .AddColumn "K3", "Descripción", ecgHdrTextALignLeft, , 290, , , , , , , CCLSortString
'''''        .AddColumn "K4", "Kilates", ecgHdrTextALignLeft, , 50, , , , , , , CCLSortNumeric
'''''        .AddColumn "K5", "Peso", ecgHdrTextALignRight, , 50, , , , , , , CCLSortNumeric
'''''        .AddColumn "K6", "Costo", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortString
'''''        .AddColumn "K7", "Precio", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
'''''        .AddColumn "K8", "Estado", ecgHdrTextALignRight, , 80, False, , , , , , CCLSortString
'''''        .AddColumn "K9", "Observaciones", ecgHdrTextALignRight, , 80, False, , , , , , CCLSortString
        
        .AddColumn "K1", "Tipo", ecgHdrTextALignLeft, , 90, False, , , , , , CCLSortString
        .AddColumn "K2", "Cant.", ecgHdrTextALignCentre, , 44, , , , , , , CCLSortNumeric
        .AddColumn "K3", "Descripción", ecgHdrTextALignLeft, , 162, , , , , , , CCLSortNumeric
        .AddColumn "K4", "Peso", ecgHdrTextALignRight, , 45, False, , , , , , CCLSortNumeric
        .AddColumn "K5", "Kilates", ecgHdrTextALignCentre, , 45, , , , , , , CCLSortString
        .AddColumn "K6", "Costo", ecgHdrTextALignRight, , 77, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Precio", ecgHdrTextALignRight, , 77, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "Modelo", ecgHdrTextALignLeft, , 62, False, , , , , , CCLSortString
        .AddColumn "K9", "Hechura", ecgHdrTextALignLeft, , 623, False, , , , , , CCLSortString
        .AddColumn "K10", "Precio V.", ecgHdrTextALignRight, , 77, False, , , , , , CCLSortString
        .AddColumn "K11", "Características", ecgHdrTextALignLeft, , 128, , , , , , , CCLSortNumeric
        
        .AddColumn "K12", "Cantidad Piedras", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortNumeric
        .AddColumn "K13", "Peso Piedras", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortNumeric
        
        .AddColumn "K14", "Cantidad Diamantes", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortNumeric
        .AddColumn "K15", "Puntos", ecgHdrTextALignLeft, , 115, False, , , , , , CCLSortNumeric
        .AddColumn "K16", "Prestamo Diamantes", ecgHdrTextALignLeft, , 115, False, , , , , , CCLSortNumeric
        
        .AddColumn "K17", "Marca", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
        .AddColumn "K18", "Modelo", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
        .AddColumn "K19", "NoSerie", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortString
        .AddColumn "K20", "Color", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortString
        .AddColumn "K21", "Tamaño", ecgHdrTextALignRight, , 50, False, , , , , , CCLSortString
        .AddColumn "K22", "Codigo", ecgHdrTextALignLeft, , 80, False, , , , , , CCLSortString
        
        .Rows = 40
    End With

End Sub

Public Function ValidaArticulos(Pestaña As Integer) As Boolean
'''''Dim rcTmp As ADODB.Recordset
'''''Dim Peso As Double, Costo As Double, Precio As Double
'''''
'''''    ValidaArticulos = True
'''''
'''''    If cmbTipo.ListIndex = -1 Then
'''''        MsgBox "Seleccione el tipo de prenda !!", vbInformation, "Compras"
'''''        ValidaArticulos = False
'''''        cmbTipo.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If cmbPrenda.ListIndex = -1 Then
'''''        MsgBox "Seleccione la prenda !!", vbInformation, "Compras"
'''''        ValidaArticulos = False
'''''        cmbPrenda.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If Trim(txtCantidad.text) = "" Or Val(txtCantidad.text) = 0 Then
'''''        MsgBox "Introduzca la cantidad !!", vbInformation, "Compras"
'''''        ValidaArticulos = False
'''''        txtCantidad.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    Set rcTmp = dbDatos.Execute("SELECT Kilataje,Peso FROM Tipo WHERE ID=" & cmbTipo.ItemData(cmbTipo.ListIndex))
'''''    If Trim(txtPeso.text) <> "" Or Val(txtPeso.text) > 0 Then
'''''
'''''        Peso = txtPeso.text
'''''    Else
'''''
'''''        Peso = 0
'''''    End If
'''''
'''''    If Peso = 0 And rcTmp!Peso = 1 Then
'''''        MsgBox "Introduzca el peso !!", vbInformation, "Compras"
'''''        ValidaArticulos = False
'''''        txtPeso.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If cmbKilates.ListIndex = -1 And rcTmp!Kilataje = 1 Then
'''''        MsgBox "Seleccione el kilataje !!", vbInformation, "Compras"
'''''        ValidaArticulos = False
'''''        cmbKilates.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If cmbEstado.ListIndex = -1 And rcTmp!Kilataje = 1 Then
'''''        MsgBox "Seleccione la hechura !!", vbInformation, "Compras"
'''''        ValidaArticulos = False
'''''        cmbEstado.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If (Trim(txtCosto.text) <> "" And Trim(txtCosto.text) <> ".") Or Val(txtCosto.text) > 0 Then
'''''
'''''        Costo = txtCosto.text
'''''    Else
'''''
'''''        Costo = 0
'''''    End If
'''''
'''''    If Costo = 0 Then
'''''        MsgBox "Introduzca el costo !!", vbInformation, "Compras"
'''''        ValidaArticulos = False
'''''        txtCosto.SetFocus
'''''        Exit Function
'''''    End If
'''''
'''''    If Trim(txtPrecio.text) <> "" Or Val(txtPrecio.text) > 0 Then
'''''        Precio = txtPrecio.text
'''''    Else
'''''        Precio = 0
'''''    End If
'''''
'''''    If Precio = 0 Then
'''''        MsgBox "Introduzca el precio !!", vbInformation, "Compras"
'''''        ValidaArticulos = False
'''''        txtPrecio.SetFocus
'''''        Exit Function
'''''    End If
'''''
''''''''''    If txtDescripcion.Text = "" Then
''''''''''        MsgBox "Introduzca las observaciones !!", vbInformation, "Compras"
''''''''''        ValidaArticulos = False
''''''''''        txtDescripcion.SetFocus
''''''''''        Exit Function
''''''''''    End If


Dim rcTmp As ADODB.Recordset

On Error GoTo Error
    
    If Pestaña = 1 Then Set rcTmp = dbDatos.Execute("SELECT Kilataje,Peso FROM tipo WHERE ID=" & cmbTipo.ItemData(cmbTipo.ListIndex))
    
    ValidaArticulos = True
    
    If cmbMovimiento.ListIndex = -1 Then
        MsgBox "Seleccione el movimiento !!", vbInformation, "Compras"
        ValidaArticulos = False
        If Pestaña = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
        Exit Function
    End If
    
    If IIf(Pestaña = 1, cmbTipo.text, cmbTipoElec.text) = "" Then
        MsgBox "Seleccione el tipo !!", vbInformation, "Compras"
        ValidaArticulos = False
        If Pestaña = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
        Exit Function
    End If
       
    If IIf(Pestaña = 1, cmbPrenda.text, txtFamiliaElec.text) = "" Then ''''' pestaña = 1 And cmbPrenda.ListIndex = -1 Then
        MsgBox "Seleccione la " & IIf(Pestaña = 1, "prenda", "familia") & " !!", vbInformation, "Compras"
        ValidaArticulos = False
        If Pestaña = 1 Then cmbPrenda.SetFocus Else txtFamiliaElec.SetFocus
        Exit Function
    End If

    If txtCantidad.text = "" And Pestaña = 1 Then
        MsgBox "Introduzca la cantidad !!", vbInformation, "Compras"
        ValidaArticulos = False
        txtCantidad.SetFocus
        Exit Function
    End If
    
    If Pestaña = 1 Then
        If cmbKilates.text = "" And rcTmp!Kilataje = 1 Then
            MsgBox "Seleccione el kilataje !!", vbInformation, "Compras"
            ValidaArticulos = False
            cmbKilates.SetFocus
            Exit Function
        End If
        
        If cmbEstado.text = "" And rcTmp!Peso = 1 Then
            MsgBox "Seleccione la hechura !!", vbInformation, "Compras"
            ValidaArticulos = False
            cmbEstado.SetFocus
            Exit Function
        End If
        
        If txtPeso.text = "" And rcTmp!Peso = 1 Then
            MsgBox "Introduzca el peso !!", vbInformation, "Compras"
            ValidaArticulos = False
            txtPeso.SetFocus
            Exit Function
        End If
    End If
    
    If Pestaña = 2 And Trim(txtMarcaElec.text) = "" Then
        MsgBox "Introduzca la marca !!", vbInformation, "Compras"
        ValidaArticulos = False
        txtMarcaElec.SetFocus
        Exit Function
    End If
    
    If IIf(Pestaña = 1, txtCosto.text, txtCostoElec) = "" Then
        MsgBox "Introduzca el préstamo !!", vbInformation, "Compras"
        ValidaArticulos = False
        If Pestaña = 1 Then txtCosto.SetFocus Else txtCostoElec.SetFocus
        Exit Function
    End If
    Set rcTmp = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

Private Sub txtPesoPiedra_Change()
    Calcula_Costo
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

Private Sub txtPrecio_GotFocus()
    Seleccionar_Texto txtPrecio
    Cambiar_Color True, txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrecio_LostFocus()
    txtPrecio.text = Format(txtPrecio.text, FMoneda)
    Cambiar_Color False, txtPrecio
End Sub

'Calculamos el total de los avaluos y prestamos
Private Sub Total_Costos()
Dim Indice As Integer, crTotalCosto As Double, crTotalPrecio As Double, crCosto As Double, crPrecio As Double, Cantidad As Integer

    For Indice = 1 To grdCompras.Rows
        crCosto = IIf(Val(grdCompras.CellText(Indice, 6)) > 0 Or Trim(grdCompras.CellText(Indice, 6)) <> "", CDbl(grdCompras.CellText(Indice, 6)), 0)
        crPrecio = IIf(Val(grdCompras.CellText(Indice, 7)) > 0 Or Trim(grdCompras.CellText(Indice, 7)) <> "", CDbl(grdCompras.CellText(Indice, 7)), 0)
        Cantidad = Val(grdCompras.CellText(Indice, 2))
        crTotalCosto = crTotalCosto + (crCosto * Cantidad)
        crTotalPrecio = crTotalPrecio + (crPrecio * Cantidad)
    Next Indice
    
    lblTotCosto.Caption = Format(Redondeo(CCur(crTotalCosto)), FMoneda)
    lblTotPrecio.Caption = Format(Redondeo(CCur(crTotalPrecio)), FMoneda)
End Sub

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
    txtDescripcion.text = ""
    txtFamiliaElec.text = ""
    txtFamiliaElec.Tag = ""
    txtMarcaElec.text = ""
    txtModeloElec.text = ""
    txtNumSerieElec.text = ""
    txtTamañoElec.text = ""
    txtColorElec.text = ""
    txtCostoElec.text = ""
    txtCostoElec.Tag = ""
    txtPrecioElec.text = ""
    txtDescripcionElec.text = ""
    txtPiedras.text = ""
    cmbPrenda.ListIndex = -1
    cmbTipo.ListIndex = 0
    cmbKilates.ListIndex = -1
    cmbEstado.ListIndex = -1
    cmbPrenda.ListIndex = -1
    cmbTipoElec.ListIndex = 0
    txtCosto.text = ""
    txtPrecio.text = ""
    chkEtiqueta.Value = Unchecked
    
'''''    cmbTipo.ListIndex = 0
'''''    cmbTipoElec.ListIndex = 0
'''''    txtFamiliaElec.text = ""
'''''    txtFamiliaElec.Tag = ""
'''''    txtMarcaElec.text = ""
'''''    txtModeloElec.text = ""
'''''    txtTamañoElec.text = ""
'''''    txtColorElec.text = ""
'''''    txtNumSerieElec.text = ""
'''''    txtCostoElec.text = ""
'''''    txtCostoElec.Tag = ""
'''''    txtPrecioElec.text = ""
'''''    cmbKilates.ListIndex = -1
'''''    cmbEstado.ListIndex = -1
'''''    cmbPrenda.ListIndex = -1
'''''    txtCantidad.text = ""
'''''    txtPeso.text = ""
'''''    txtPrecio.text = ""
'''''    txtCosto.text = ""
'''''    txtDescripcion.text = ""
'''''    txtDescripcionElec.text = ""
End Sub

'Grabamos todos los datos necesarios
Private Sub GrabarEntradas()
Dim Folio As Long, IDCliente As Long, IDEntrada As Long, IDCompra As Long, Iva As Double, crTotal As Double
Dim TipoMovimiento As Boolean

'    'Cliente
'    If Val(txtNombre.Tag) = 0 And Trim(txtNombre.text) <> "" Then
'
'        IDCliente = Grabar_Cliente()
'
'    ElseIf Val(txtNombre.Tag) > 0 Then
'
'        IDCliente = Val(txtNombre.Tag)
'        Actualizar_Cliente IDCliente
'    End If
    
    '--- MLD-MODIF.- Grabar el Cliente ---
    If Trim(txtNombre.text) <> "" And Trim(txtApellidoPaterno.text) <> "" And Trim(txtApellidoMaterno.text) <> "" Then
        If ClienteCom.Valida = True Then
            ClienteCom.Grabar
            IDCliente = ClienteCom.ID
        Else
            MsgBox "Datos incompletos del CoTitular.", vbCritical, Me.Caption
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    '-------------------------------------
    
    'Checo el tipo de movimiento
    If cmbMovimiento.ListIndex = 0 Then
        
        TipoMovimiento = True
    Else
        
        TipoMovimiento = False
    End If
    
    'Tomo el Total General
    crTotal = CDbl(lblTotCosto.Caption)
    
    'Tomo el Iva
    Iva = 0
    
    'Saco el Folio
    Folio = Me.Folio
    
    'Saco el ID de la Entrada
    IDEntrada = Me.IDEmpeno
    
    'Grabo la compra
    If TipoMovimiento Then
        
        dbDatos.Execute "INSERT INTO compras (Fecha,Folio,IDCliente,Total,Iva,IDUsuario,IDSucursal) VALUES ('" & _
                        Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & IDCliente & "," & ConvMoneda(crTotal) & "," & ConvMoneda(Iva) & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
        'Saco el ID de la Compra
        IDCompra = SacaValor("compras", "MAX(ID)")
    
    End If
            
    'Grabo el Detalle de la Compra
    GrabarInventario IDEntrada, IDCompra, Folio, crTotal, Iva, TipoMovimiento
    
    'Imprimo el comprobante de Compra
    If TipoMovimiento Then
    
        ImprimirTicket Folio
    Else
        
        ImprimirDotacion IDEntrada
    End If
    
    Ban = False
    Me.Folio = 0
    Me.IDEmpeno = 0
    LimpiarCliente
    LimpiaArticulos
    txtCantidad.text = "1"
    lblTotCosto.Caption = "0.00"
    lblTotPrecio.Caption = "0.00"
    lblFolio.Caption = Regresa_Movimiento(False, "FolioInventario")
    cmbMovimiento.ListIndex = -1
    grdCompras.Redraw = False
    grdCompras.Clear
    grdCompras.Rows = 40
    grdCompras.Redraw = True
End Sub

Private Sub GrabarInventario(IDEntrada As Long, IDCompra As Long, Folio As Long, crTotal As Double, Iva As Double, TipoMovimiento As Boolean)
Dim i As Integer, Peso As Double, Codigo As String, TipoPrenda As Integer, Movimiento As Long, Kilates As Integer, crIva As Double
Dim CantidadPiedras As Integer, PesoPiedras As Double, CantidadDiamantes As Integer, PuntosDiamantes As Double, crPrestamoDiamantes As Double
Dim strMarca As String, strModelo As String, strNumSerie As String, strColor As String, strTamaño As String

    With grdCompras
    
        For i = 1 To .Rows
            
            If Val(.CellText(i, 2)) > 0 Then
                
                Codigo = CreaCodigoBarras(Trim(Format(frmMDI.IDSucursal, "000")), IIf(TipoMovimiento, ENTRADACOMPRA, ENTRADADOTACION), Trim(Folio), i)
                Kilates = RegresaKilates(IIf(.CellText(i, 1) = "ORO", .CellText(i, 5), ""), .CellText(i, 1))
                
                Peso = IIf(Val(.CellText(i, 4)) = 0 Or Trim(.CellText(i, 4)) = "", 0, .CellText(i, 4))
                
                CantidadPiedras = IIf(Val(.CellText(i, 12)) = 0 Or Trim(.CellText(i, 12)) = "", 0, .CellText(i, 12))
                PesoPiedras = IIf(Val(.CellText(i, 13)) = 0 Or Trim(.CellText(i, 13)) = "", 0, .CellText(i, 13))
                
                CantidadDiamantes = IIf(Val(.CellText(i, 14)) = 0 Or Trim(.CellText(i, 14)) = "", 0, .CellText(i, 14))
                PuntosDiamantes = IIf(Val(.CellText(i, 15)) = 0 Or Trim(.CellText(i, 15)) = "", 0, .CellText(i, 15))
                crPrestamoDiamantes = IIf(Val(.CellText(i, 16)) = 0 Or Trim(.CellText(i, 16)) = "", 0, .CellText(i, 16))
                
                strMarca = Trim(.CellText(i, 17))
                strModelo = Trim(.CellText(i, 18))
                strNumSerie = Trim(.CellText(i, 19))
                strColor = Trim(.CellText(i, 20))
                strTamaño = Trim(.CellText(i, 21))
                                
                'DetallesEntradaInventario
                dbDatos.Execute "INSERT INTO detallesentradainventario (IDEntrada,Codigo,Tipo,Cantidad,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,SucursalOrigen,TipoEntrada,PrecioVitrina,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante) VALUES (" & _
                                IDEntrada & ",'" & Codigo & "'," & Val(.CellItemData(i, 1)) & "," & Val(.CellText(i, 2)) & ",'" & Trim(.CellText(i, 3)) & "'," & ConvMoneda(Peso) & "," & Kilates & "," & ConvMoneda(.CellText(i, 7)) & "," & ConvMoneda(.CellText(i, 6)) & ",'" & Trim(.CellText(i, 9)) & "','" & strMarca & "','" & strModelo & "','" & strNumSerie & "','" & strColor & "','" & strTamaño & "'," & Val(.CellItemData(i, 3)) & ",'" & Trim(.CellText(i, 11)) & "'," & frmMDI.IDSucursal & "," & IIf(TipoMovimiento, ENTRADACOMPRA, ENTRADADOTACION) & "," & ConvMoneda(.CellText(i, 7)) & "," & CantidadPiedras & "," & ConvMoneda(PesoPiedras) & "," & CantidadDiamantes & "," & ConvMoneda(PuntosDiamantes) & "," & ConvMoneda(crPrestamoDiamantes) & ")"
                
                If TipoMovimiento Then
                
                    'DetallesCompras
                    dbDatos.Execute "INSERT INTO detallescompras (IDCompra,Tipo,Codigo,Descripcion,Kilates,Estado,Cantidad,Peso,Costo,Precio,TipoPrenda,Observaciones) VALUES (" & _
                                    IDCompra & "," & Val(.CellItemData(i, 1)) & ",'" & Trim(Codigo) & "','" & Trim(.CellText(i, 3)) & "'," & Kilates & ",'" & Trim(.CellText(i, 9)) & "'," & Val(.CellText(i, 2)) & "," & ConvMoneda(Peso) & "," & ConvMoneda(.CellText(i, 6)) & "," & ConvMoneda(.CellText(i, 7)) & "," & Val(.CellItemData(i, 3)) & ",'" & Trim(.CellText(i, 11)) & "')"
                End If
                
            End If
        
        Next i
    
    End With
        
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True

    'Desgloso el Iva
    crIva = Redondeo(crTotal - (crTotal / (1 + (Iva / 100))))
    
    'Meto el total sin iva
    crTotal = Redondeo(crTotal - crIva)

    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'EN01','620301'," & ConvMoneda(crTotal) & "," & TIPO_CARGO & "," & IIf(TipoMovimiento, 0, 1) & ",'" & IIf(TipoMovimiento, "Compras", "Dotacion Inventario") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'EN50','" & IIf(TipoMovimiento, "110150", "200950") & "'," & ConvMoneda(crTotal) & "," & TIPO_ABONO & "," & IIf(TipoMovimiento, 0, 1) & ",'" & IIf(TipoMovimiento, "Compras", "Dotacion Inventario") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                        
'''    If TipoMovimiento Then
'''
'''        'Grabamos el abono
'''        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
'''                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'EN50','199450'," & ConvMoneda(crTotal + crIva) & "," & TIPO_ABONO & ",0,'Compras','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'''    End If
    
End Sub

Function ValidaCompra() As Boolean
Dim i As Integer, Prendas As Integer
    
    ValidaCompra = True
    
    If cmbMovimiento.ListIndex = -1 Then
        
        MsgBox "Seleccione el tipo de movimiento !!", vbInformation, "Compras/Dotaciones a Inventario"
        ValidaCompra = False
        cmbMovimiento.SetFocus
        Exit Function
    End If
    
    Prendas = grdCompras.Rows
    For i = 1 To grdCompras.Rows
        
        If Val(grdCompras.CellText(i, 2)) > 0 Then Prendas = Prendas - 1
    Next i
    
    If Prendas = grdCompras.Rows Then
        MsgBox "Favor de introducir las prendas !!", vbInformation, "Compras/Dotaciones a Inventario"
        ValidaCompra = False
    End If
    
End Function



Private Sub txtPrecioElec_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtPrecioElec.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    Pasar_Foco KeyAscii
End Sub



'Private Function Grabar_Cliente() As Long
'Dim rc As New ADODB.Recordset
'
'On Error GoTo Error
'
'    dbDatos.Execute "INSERT INTO clientes (Nombre,Apellido,Iniciales,Direccion,Colonia,Municipio,Estado,Tel,Identificacion,CP) VALUES ('" & _
'                    Trim(txtNombre.text) & "','" & Trim(txtApellidos.text) & "','" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "','" & Trim(txtDireccion.text) & "','" & Trim(txtColonia.text) & "','" & Trim(txtMunicipio.text) & "','" & Trim(txtEstado.text) & "','" & Trim(txtTelefono.text) & "','" & Trim(txtIdentificacion.text) & "','" & txtCP.text & "')"
'
'    rc.Open "SELECT MAX(ID) AS IDD FROM clientes", dbDatos, adOpenForwardOnly, adLockOptimistic
'
'        Grabar_Cliente = rc!idd
'
'    rc.Close
'    Set rc = Nothing
'    Exit Function
'
'Error:
'    Maneja_Error Err
'    Set rc = Nothing
'End Function
'
'Private Sub Actualizar_Cliente(ID As Long)
'
'On Error GoTo Error
'
'    dbDatos.Execute "UPDATE clientes SET nombre='" & Trim(txtNombre.text) & "',apellido='" & Trim(txtApellidos.text) & "',Iniciales='" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "',Direccion='" & Trim(txtDireccion.text) & "',Colonia='" & Trim(txtColonia.text) & "',Municipio='" & Trim(txtMunicipio.text) & "'," & _
'                    "Estado='" & Trim(txtEstado.text) & "',Tel='" & Trim(txtTelefono.text) & "',Identificacion='" & Trim(txtIdentificacion.text) & "',CP='" & txtCP.text & "' WHERE ID = " & ID
'    Exit Sub
'
'Error:
'    Maneja_Error Err
'End Sub

'Public Sub Buscar_Cliente(IDCliente As Long)
'Dim rcClientes As New ADODB.Recordset
'
'On Error GoTo Error
'
'    rcClientes.Open "SELECT * FROM clientes WHERE ID=" & IDCliente, dbDatos, adOpenForwardOnly, adLockOptimistic
'    With rcClientes
'
'        txtNombre.text = !Nombre
'        txtNombre.Tag = IDCliente
'        txtApellidos.text = !Apellido
'        txtDireccion.text = IIf(IsNull(!Direccion), "", !Direccion)
'        txtColonia.text = IIf(IsNull(!Colonia), "", !Colonia)
'        txtMunicipio.text = IIf(IsNull(!Municipio), "", !Municipio)
'        txtEstado.text = IIf(IsNull(!Estado), "", !Estado)
'        txtTelefono.text = IIf(IsNull(!Tel), "", !Tel)
'        txtIdentificacion.text = IIf(IsNull(!Identificacion), "", !Identificacion)
'        txtCP.text = IIf(IsNull(!CP), "", !CP)
'
'    End With
'    rcClientes.Close
'    Set rcClientes = Nothing
'    Exit Sub
'
'Error:
'    Maneja_Error Err
'    Set rcClientes = Nothing
'End Sub

Sub LimpiarCliente()
    ClienteCom.Limpiar
    txtNombre.text = ""
    txtNombre.Tag = ""
    txtApellidoPaterno.text = ""
    txtApellidoMaterno.text = ""
    lblDireccion.Caption = ""
    lblCiudad.Caption = ""
    lblRFC.Caption = ""

'    txtBuscar.text = ""
'    txtNombre.text = ""
'    txtNombre.Tag = ""
'    txtApellidos.text = ""
'    txtDireccion.text = ""
'    txtColonia.text = ""
'    txtMunicipio.text = ""
'    txtEstado.text = ""
'    txtTelefono.text = ""
'    txtIdentificacion.text = ""
'    txtCP.text = ""
End Sub

Sub ImprimirTicket(Folio As Long)
Dim ImprDefault As Boolean

    'Checo si hay impresora predeterminada
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
        
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{Compras.Folio}=" & Folio & ""
        .ReportFileName = Path & "\Reportes\NotaCompra.rpt"
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .Destination = crptToWindow
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
            
        .WindowState = crptMaximized
        .WindowTitle = "Recibo"
        .Action = 1
    End With

End Sub

'Creamos las pestañas del tab
Private Sub Crear_Pestañas()

    With TPrendas
        .AddTab "Metales", , , "K1"
        .AddTab "Electrónicos/Varios", , , "K4"
    End With
    
End Sub

Private Sub cmbTipoElec_GotFocus()
    Cambiar_Color True, cmbTipoElec
End Sub

Private Sub cmbTipoElec_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoElec_LostFocus()
    Cambiar_Color False, cmbTipoElec
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

Private Sub txtCostoElec_Change()
    Calcula_Avaluo_Elec
End Sub

Private Sub txtCostoElec_GotFocus()
    Seleccionar_Texto txtCostoElec
    Cambiar_Color True, txtCostoElec
End Sub

Private Sub txtCostoElec_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtCostoElec.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCostoElec_LostFocus()
    txtCostoElec.text = Format(txtCostoElec.text, FMoneda)
    Cambiar_Color False, txtCostoElec
End Sub

Private Sub txtPrecioElec_GotFocus()
    Seleccionar_Texto txtPrecioElec
    Cambiar_Color True, txtPrecioElec
End Sub

Private Sub txtPrecioElec_LostFocus()
    txtPrecioElec.text = Format(txtPrecioElec.text, FMoneda)
    Cambiar_Color False, txtPrecioElec
End Sub

Sub Calcula_Avaluo_Elec()
Dim crPrestamo As Double, crMaximo As Double, PrestamoAvaluo As Double
    
    crPrestamo = 0
    crMaximo = 0
    If Val(txtCostoElec.text) > 0 Or Trim(txtCostoElec.text) <> "" Then
        
        'Tomo el prestamo
        crPrestamo = CDbl(txtCostoElec.text)
        
        'Tomo el importe máximo
        If Val(txtCostoElec.Tag) > 0 Or Trim(txtCostoElec.Tag) <> "" Then
        
            crMaximo = CDbl(txtCostoElec.Tag)
        End If
        
        If crPrestamo > crMaximo Then
            MsgBox "Ha sobrepasado el limite máximo permitido !!", vbInformation, "Empeño"
            crPrestamo = crMaximo
            txtCostoElec.text = crMaximo
        End If
        
        PrestamoAvaluo = Regresa_Valor_BD("PrestamoAvaluoElec") / 100
        
        txtPrecioElec.text = Format(crPrestamo * (1 + PrestamoAvaluo), FMoneda)
    
    End If
        
End Sub

Private Sub txtDescripcionElec_GotFocus()
    Seleccionar_Texto txtDescripcionElec
    Cambiar_Color True, txtDescripcionElec
End Sub

Private Sub txtDescripcionElec_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        cmdAgregar.SetFocus
    End If
End Sub

Private Sub txtDescripcionElec_LostFocus()
    Cambiar_Color False, txtDescripcionElec
End Sub

Function ImprimirDotacion(IDEntrada As Long)

    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepDotacion.rpt"
        .SelectionFormula = "{entradainventario.ID}=" & IDEntrada
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado=''"
        .WindowTitle = "Reporte Dotación a Inventario"
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With

End Function

Function Imprimir(Codigo As String, Peso As Double, Precio As Double, Kilates As Integer, Cantidad As Integer, strPrenda As String)
Dim Impresora As Printer
Dim i As Integer

    DoEvents
    Sleep 500
    bcCodigo.text = Left(Codigo, 12)
    
    Set Impresora = Printer
    With Impresora

        For i = 1 To Cantidad
        
            bcCodigo.text = Left(Codigo, 12)
            .ScaleMode = vbMillimeters
            .Font = "Arial"
            .FontSize = 6.5
    
            'Imprimo el peso
            .CurrentX = Regresa_Valor("ETIQUETAS", "PesoX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "PesoY", 0)
            Impresora.Print Format(Peso, "##,###0.00") & " Grs."
    
            'Imprimo el Kilataje
            .CurrentX = Regresa_Valor("ETIQUETAS", "KilatesX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "KilatesY", 0)
            Impresora.Print SacaKilates(Kilates)
            
            'Imprimo la prenda
            .CurrentX = Regresa_Valor("ETIQUETAS", "PrendaX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "PrendaY", 0)
            Impresora.Print strPrenda
                
            'Imprimo el precio
            .CurrentX = Regresa_Valor("ETIQUETAS", "PrecioX", 0)
            .CurrentY = Regresa_Valor("ETIQUETAS", "PrecioY", 0)
            Impresora.Print Format(Precio, FMoneda)
        
            Sleep 500
            
            'Imprimo el Código de Barras
            .PaintPicture bcCodigo.Picture, Regresa_Valor("ETIQUETAS", "CodigoX", 0), Regresa_Valor("ETIQUETAS", "CodigoY", 0), Regresa_Valor("ETIQUETAS", "Anchocodigo", 0), Regresa_Valor("ETIQUETAS", "Altocodigo", 0)
        
        .EndDoc
        Next i

    End With

End Function

Function CreaEncabezado()
Dim Folio As Long, TipoMovimiento As Boolean
    
    Select Case cmbMovimiento.ListIndex
    Case 0
        
        TipoMovimiento = True
    Case 1
        
        TipoMovimiento = False
    End Select
    
    'Saco el Folio
    Folio = Regresa_Movimiento(False, "FolioInventario")
    Regresa_Movimiento True, "FolioInventario"
    
    'Grabo en Entrada Inventario
    dbDatos.Execute "INSERT INTO entradainventario (Fecha,Folio,TipoEntrada,IDUsuario,IDSucursal) VALUES ('" & _
                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & IIf(TipoMovimiento, ENTRADACOMPRA, ENTRADADOTACION) & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    lblFolio.Caption = Folio
    
    'Guardo el Folio
    Me.Folio = Folio
    
    'Tomo el Valor Máximo
    Me.IDEmpeno = SacaValor("entradainventario", "MAX(ID)")
    
End Function


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
    ClienteCom.Nombre = txtNombre.text
End Sub

'----------------------------------------------------------------

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
    ClienteCom.ApellidoPaterno = txtApellidoPaterno.text
End Sub


'-----------------------------------------------------------------
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
    ClienteCom.ApellidoMaterno = txtApellidoMaterno.text
    If Trim(txtNombre.text) <> "" And Trim(txtApellidoPaterno.text) <> "" And Trim(txtApellidoMaterno.text) <> "" And Val(txtNombre.Tag) <= 0 Then Mostrar_Seleccionar_Cliente Trim(txtNombre.text), Trim(txtApellidoPaterno.text), Trim(txtApellidoMaterno.text), Me
    If Val(txtNombre.Tag) = 0 Then
        ClienteCom.Nombre = Trim(txtNombre.text)
        ClienteCom.ApellidoPaterno = Trim(txtApellidoPaterno.text)
        ClienteCom.ApellidoMaterno = Trim(txtApellidoMaterno.text)
        frmClientes.Mostrar ClienteCom
        '--- Mostrar Datos Cliente ---
        If ClienteCom.ID = 0 Then
            lblDireccion.Caption = ""
            lblCiudad.Caption = ""
            lblRFC.Caption = ""
        Else
            lblDireccion.Caption = ClienteCom.Direccion & IIf(ClienteCom.NoExterior <> "", " #" & ClienteCom.NoExterior, "") & IIf(ClienteCom.NoInterior <> "", " INT." & ClienteCom.NoInterior, "") & " C.P." & ClienteCom.CodigoPostal
            lblCiudad.Caption = ClienteCom.Municipio & ", " & ClienteCom.Estado
            lblRFC.Caption = ClienteCom.RFC
        End If
    End If
End Sub


Private Sub cmdMosCliente_Click()
    frmMostrarCliente.Ver Me, txtNombre, True
End Sub


Private Sub cmdEditar_Click()
    frmClientes.Mostrar ClienteCom
    txtNombre.text = ClienteCom.Nombre
    txtNombre.Tag = ClienteCom.ID
    txtApellidoPaterno.text = ClienteCom.ApellidoPaterno
    txtApellidoMaterno.text = ClienteCom.ApellidoMaterno
    '--- Mostrar Datos Cliente ---
    If ClienteCom.ID = 0 Then
        lblDireccion.Caption = ""
        lblCiudad.Caption = ""
        lblRFC.Caption = ""
    Else
        lblDireccion.Caption = ClienteCom.Direccion & IIf(ClienteCom.NoExterior <> "", " #" & ClienteCom.NoExterior, "") & IIf(ClienteCom.NoInterior <> "", " INT." & ClienteCom.NoInterior, "") & " COL." & ClienteCom.Colonia & " C.P." & ClienteCom.CodigoPostal
        lblCiudad.Caption = ClienteCom.Municipio & ", " & ClienteCom.Estado
        lblRFC.Caption = ClienteCom.RFC
    End If
End Sub

'-------------------------------------------------------------------
'MLD-MODIF.- Buscamos el id cliente
Public Sub Buscar(ID As Long)
On Error GoTo Error

    ClienteCom.Buscar ID
    
    If Not ClienteCom.Valida Then
        frmClientes.Mostrar ClienteCom
    End If
    txtNombre.text = ClienteCom.Nombre
    txtNombre.Tag = ClienteCom.ID
    txtApellidoPaterno.text = ClienteCom.ApellidoPaterno
    txtApellidoMaterno.text = ClienteCom.ApellidoMaterno
    cmdEditar.Visible = True
    
    '--- Mostrar Datos Cliente ---
    lblDireccion.Caption = ClienteCom.Direccion & IIf(ClienteCom.NoExterior <> "", " #" & ClienteCom.NoExterior, "") & IIf(ClienteCom.NoInterior <> "", " INT." & ClienteCom.NoInterior, "") & " COL." & ClienteCom.Colonia & " C.P." & ClienteCom.CodigoPostal
    lblCiudad.Caption = ClienteCom.Municipio & ", " & ClienteCom.Estado
    lblRFC.Caption = ClienteCom.RFC
    
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub


