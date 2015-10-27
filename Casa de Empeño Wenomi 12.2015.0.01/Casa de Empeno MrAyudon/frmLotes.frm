VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmLotes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Deslotificación"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12765
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7065
   ScaleWidth      =   12765
   Begin VB.Frame frmSepararlote 
      Height          =   6465
      Left            =   30
      TabIndex        =   13
      Top             =   60
      Width           =   12690
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   30
         Left            =   30
         Top             =   2400
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   53
      End
      Begin VB.TextBox txtNumContrato 
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
         Height          =   210
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1215
      End
      Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
         Height          =   3705
         Left            =   4050
         TabIndex        =   32
         Top             =   2415
         Width           =   8595
         _ExtentX        =   15161
         _ExtentY        =   6535
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
         Begin VB.TextBox txtEmpeños 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   285
            Left            =   1800
            MaxLength       =   90
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   1215
         End
      End
      Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
         Height          =   375
         Left            =   2400
         TabIndex        =   12
         Top             =   285
         Width           =   1035
         _ExtentX        =   1826
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
         Picture         =   "frmLotes.frx":000C
      End
      Begin vbalTabStrip6.TabControl TPrendas 
         Height          =   3975
         Left            =   0
         TabIndex        =   47
         Top             =   2445
         Width           =   4035
         _ExtentX        =   7117
         _ExtentY        =   7011
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
         Begin VB.Frame frmMetales 
            Caption         =   "Metales"
            Height          =   3180
            Left            =   15
            TabIndex        =   51
            Top             =   315
            Width           =   3990
            Begin VB.TextBox txtPeso 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   2895
               MaxLength       =   20
               TabIndex        =   6
               Top             =   1035
               Width           =   1020
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
               TabIndex        =   52
               Top             =   1635
               Width           =   1020
            End
            Begin VB.ComboBox cmbKilates 
               Height          =   315
               ItemData        =   "frmLotes.frx":0391
               Left            =   1065
               List            =   "frmLotes.frx":0393
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   690
               Width           =   1110
            End
            Begin VB.TextBox txtCantidad 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1095
               MaxLength       =   3
               TabIndex        =   5
               Top             =   1035
               Width           =   1020
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
               TabIndex        =   9
               Top             =   1635
               Width           =   1020
            End
            Begin VB.TextBox txtObservaciones 
               BorderStyle     =   0  'None
               Height          =   960
               Left            =   60
               MaxLength       =   150
               MultiLine       =   -1  'True
               TabIndex        =   10
               Top             =   2175
               Width           =   3870
            End
            Begin VB.ComboBox cmbPrenda 
               Height          =   315
               ItemData        =   "frmLotes.frx":0395
               Left            =   1065
               List            =   "frmLotes.frx":0397
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   375
               Width           =   2895
            End
            Begin VB.TextBox txtPesoPiedra 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1095
               MaxLength       =   20
               TabIndex        =   7
               Top             =   1335
               Width           =   1020
            End
            Begin VB.ComboBox cmbTipo 
               Height          =   315
               ItemData        =   "frmLotes.frx":0399
               Left            =   1065
               List            =   "frmLotes.frx":039B
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   60
               Width           =   2130
            End
            Begin VB.ComboBox cmbEstado 
               Height          =   315
               ItemData        =   "frmLotes.frx":039D
               Left            =   2865
               List            =   "frmLotes.frx":039F
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   690
               Width           =   1095
            End
            Begin VB.TextBox txtPiedras 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   2895
               MaxLength       =   3
               TabIndex        =   8
               Top             =   1335
               Width           =   1020
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
               Top             =   750
               Width           =   690
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
               TabIndex        =   64
               Top             =   1665
               Width           =   870
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
               TabIndex        =   63
               Top             =   1950
               Width           =   1320
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
               TabIndex        =   62
               Top             =   1365
               Width           =   675
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
               TabIndex        =   61
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
               TabIndex        =   60
               Top             =   750
               Width           =   615
            End
            Begin VB.Label lblPiedra 
               BackColor       =   &H80000013&
               Height          =   255
               Left            =   2070
               TabIndex        =   59
               Top             =   1815
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblCantidadPiedras 
               BackColor       =   &H80000013&
               Caption         =   "0"
               Height          =   255
               Left            =   3510
               TabIndex        =   58
               Top             =   1815
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblPuntos 
               BackColor       =   &H80000013&
               Caption         =   "0"
               Height          =   255
               Left            =   2820
               TabIndex        =   57
               Top             =   1815
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblPrestamoDiamante 
               BackColor       =   &H80000013&
               Caption         =   "0"
               Height          =   255
               Left            =   3255
               TabIndex        =   56
               Top             =   1050
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.Label lblAvaluoDiamante 
               BackColor       =   &H80000013&
               Caption         =   "0"
               Height          =   255
               Left            =   3540
               TabIndex        =   55
               Top             =   1470
               Visible         =   0   'False
               Width           =   600
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
               TabIndex        =   54
               Top             =   1665
               Width           =   630
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
               TabIndex        =   53
               Top             =   1365
               Width           =   1035
            End
         End
         Begin VB.Frame frmElectronicos 
            Caption         =   "Electronicos"
            Height          =   3180
            Left            =   15
            TabIndex        =   69
            Top             =   315
            Width           =   3990
            Begin VB.ComboBox cmbTipoElec 
               Height          =   315
               ItemData        =   "frmLotes.frx":03A1
               Left            =   1065
               List            =   "frmLotes.frx":03A3
               Style           =   2  'Dropdown List
               TabIndex        =   80
               Top             =   60
               Width           =   2130
            End
            Begin VB.TextBox txtObservacionesElec 
               BorderStyle     =   0  'None
               Height          =   960
               Left            =   60
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   79
               Top             =   2175
               Width           =   3870
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
               Left            =   1095
               MaxLength       =   20
               TabIndex        =   78
               Top             =   1635
               Width           =   1020
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
               TabIndex        =   77
               Top             =   1635
               Width           =   1095
            End
            Begin VB.TextBox txtModeloElec 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   2895
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   76
               Top             =   727
               Width           =   1050
            End
            Begin VB.TextBox txtNumSerieElec 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1095
               MaxLength       =   80
               TabIndex        =   75
               Top             =   1335
               Width           =   2850
            End
            Begin VB.TextBox txtTamañoElec 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1095
               MaxLength       =   50
               TabIndex        =   74
               Top             =   1035
               Width           =   1020
            End
            Begin VB.TextBox txtColorElec 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   2895
               MaxLength       =   50
               TabIndex        =   73
               Top             =   1035
               Width           =   1050
            End
            Begin VB.TextBox txtFamiliaElec 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1095
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   71
               Top             =   412
               Width           =   2850
            End
            Begin VB.TextBox txtMarcaElec 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1095
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   70
               Top             =   727
               Width           =   1020
            End
            Begin DevPowerFlatBttn.FlatBttn cmdMostrarCatPrendas 
               Height          =   270
               Left            =   3195
               TabIndex        =   72
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
               TabIndex        =   90
               Top             =   1665
               Width           =   630
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
               TabIndex        =   89
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
               TabIndex        =   88
               Top             =   1950
               Width           =   1320
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
               TabIndex        =   87
               Top             =   1665
               Width           =   870
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
               TabIndex        =   86
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
               TabIndex        =   85
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
               TabIndex        =   84
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
               TabIndex        =   83
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
               TabIndex        =   82
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
               TabIndex        =   81
               Top             =   1065
               Width           =   480
            End
         End
         Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
            Height          =   375
            Left            =   2085
            TabIndex        =   11
            Top             =   3540
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
            Picture         =   "frmLotes.frx":03A5
            PictureDisabled =   "frmLotes.frx":070F
         End
         Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
            Height          =   375
            Left            =   90
            TabIndex        =   50
            Top             =   3540
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
            Picture         =   "frmLotes.frx":0869
         End
         Begin DevPowerFlatBttn.FlatBttn cmdDiamante 
            Height          =   375
            Left            =   1020
            TabIndex        =   49
            Top             =   3540
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
            Picture         =   "frmLotes.frx":096D
            PictureDisabled =   "frmLotes.frx":0B91
         End
         Begin DevPowerFlatBttn.FlatBttn cmdBorrar 
            Height          =   375
            Left            =   3030
            TabIndex        =   48
            Top             =   3540
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
            Picture         =   "frmLotes.frx":0CEB
            PictureDisabled =   "frmLotes.frx":123D
         End
      End
      Begin VB.Label lblNumContrato 
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   93
         Top             =   360
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label lblTasa 
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
         Height          =   210
         Left            =   870
         TabIndex        =   46
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tasa:"
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
         TabIndex        =   45
         Top             =   2160
         Visible         =   0   'False
         Width           =   465
      End
      Begin VB.Label lblPeso 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "0.000"
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
         Left            =   6840
         TabIndex        =   44
         Top             =   6150
         Width           =   540
      End
      Begin VB.Label lblAvaluo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "0.00"
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
         Left            =   8880
         TabIndex        =   43
         Top             =   6150
         Width           =   420
      End
      Begin VB.Label lblPrestamo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H80000018&
         Caption         =   "0.00"
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
         Left            =   10080
         TabIndex        =   42
         Top             =   6150
         Width           =   420
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Peso Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   10635
         TabIndex        =   41
         Top             =   1695
         Width           =   1500
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Avalúo Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   10335
         TabIndex        =   40
         Top             =   255
         Width           =   1800
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Préstamo Total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   345
         Left            =   9960
         TabIndex        =   39
         Top             =   975
         Width           =   2175
      End
      Begin VB.Label lblTotAvaluo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   12045
         TabIndex        =   38
         Top             =   615
         Width           =   90
      End
      Begin VB.Label lblTotPrestamo 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   12045
         TabIndex        =   37
         Top             =   1335
         Width           =   90
      End
      Begin VB.Label lblTotPeso 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   345
         Left            =   12045
         TabIndex        =   36
         Top             =   2055
         Width           =   90
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   4050
         TabIndex        =   35
         Top             =   6120
         Width           =   8595
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato:"
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
         TabIndex        =   34
         Top             =   360
         Width           =   900
      End
      Begin VB.Label lblIdentificacion 
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
         Height          =   210
         Left            =   6240
         TabIndex        =   31
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Label lblTelefono 
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
         Height          =   210
         Left            =   3315
         TabIndex        =   30
         Top             =   1800
         Width           =   1380
      End
      Begin VB.Label lblEstado 
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
         Height          =   210
         Left            =   870
         TabIndex        =   29
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lblCp 
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
         Height          =   210
         Left            =   7260
         TabIndex        =   28
         Top             =   1440
         Width           =   1395
      End
      Begin VB.Label lblMunicipio 
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
         Height          =   210
         Left            =   4200
         TabIndex        =   27
         Top             =   1440
         Width           =   2595
      End
      Begin VB.Label lblColonia 
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
         Height          =   210
         Left            =   600
         TabIndex        =   26
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lblDireccion 
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
         Height          =   210
         Left            =   1080
         TabIndex        =   25
         Top             =   1080
         Width           =   7575
      End
      Begin VB.Label lblApellido 
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
         Height          =   210
         Left            =   5400
         TabIndex        =   24
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblNombre 
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
         Height          =   210
         Left            =   960
         TabIndex        =   23
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos:"
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
         Left            =   4440
         TabIndex        =   22
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
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
         TabIndex        =   21
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label64 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
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
         TabIndex        =   20
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label Label65 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
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
         Left            =   3210
         TabIndex        =   19
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label Label66 
         AutoSize        =   -1  'True
         Caption         =   "Col:"
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
         TabIndex        =   18
         Top             =   1440
         Width           =   345
      End
      Begin VB.Label Label71 
         AutoSize        =   -1  'True
         Caption         =   "Estado:"
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
         TabIndex        =   17
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label Label72 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
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
         Left            =   2400
         TabIndex        =   16
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "Identificación:"
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
         Left            =   4860
         TabIndex        =   15
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "Cp:"
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
         Left            =   6870
         TabIndex        =   14
         Top             =   1440
         Width           =   300
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   11490
      TabIndex        =   91
      Top             =   6600
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
      Picture         =   "frmLotes.frx":1E0F
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   10320
      TabIndex        =   92
      Top             =   6600
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
      Picture         =   "frmLotes.frx":2361
   End
End
Attribute VB_Name = "frmLotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim BanElec As Boolean, Bandera As Boolean

Private Sub cmdAceptar_Click()
Dim Indice As Integer, TotPeso As Double, TotPrestamo As Double, Prestamo As Double, PesoUnitario As Double
Dim Codigo As String, Peso As Double, CantidadPiedras As Integer, PesoPiedras As Double, CantidadDiamantes As Integer, PuntosDiamantes As Double, crPrestamoDiamantes As Double, Kilates As Integer, strMarca As String, strModelo As String, strNumSerie As String, strColor As String, strTamaño As String
    
    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Deslotificación") = vbYes Then
        
        If txtNumContrato.Tag <> "" Then

            With grdArticulos
                    
                TotPeso = lblTotPeso.Caption
                TotPrestamo = lblTotPrestamo.Caption
                Peso = lblPeso.Caption
                Prestamo = lblPrestamo.Caption
            
'''''                If Peso <> TotPeso Then
'''''                    MsgBox "El peso no coincide con el del empeño !!", vbCritical, "Deslotificación"
'''''                    Exit Sub
'''''                End If
            
                If Prestamo <> TotPrestamo Then
                    MsgBox "El importe del préstamo no coincide con el del empeño !!", vbCritical, "Deslotificación"
                    Exit Sub
                End If
            
                dbDatos.Execute "DELETE FROM detallesempeno WHERE IDEmpeno=" & Val(txtNumContrato.Tag)

                For Indice = 1 To .Rows

                    If grdArticulos.CellText(Indice, 1) <> "" Then
            
                        Codigo = CreaCodigoBarras(Trim(Format(frmMDI.IDSucursal, "000")), ENTRADAEMPENO, Trim(lblNumContrato), Indice)
                        Kilates = RegresaKilates(IIf(grdArticulos.CellText(Indice, 1) = "ORO", grdArticulos.CellText(Indice, 5), ""), grdArticulos.CellText(Indice, 1))
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
                        
                        dbDatos.Execute "INSERT INTO detallesempeno (IDEmpeno,Codigo,Tipo,Cantidad,Articulo,Peso,Kilates,Avaluo,Prestamo,Estado,Origen,Destino,TipoPrenda,Observaciones,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante,Marca,Modelo,Serie,Color,Tamano) VALUES (" & _
                                        Val(txtNumContrato.Tag) & ",'" & Trim(Codigo) & "'," & .CellItemData(Indice, 1) & "," & Val(.CellText(Indice, 2)) & ",'" & Trim(UCase(.CellText(Indice, 3))) & "'," & Peso & "," & _
                                        Kilates & "," & CDbl(.CellText(Indice, 6)) & "," & CDbl(.CellText(Indice, 7)) & ",'" & Trim(UCase(.CellText(Indice, 9))) & "'," & ENTRADAEMPENO & ",0," & Val(.CellItemData(Indice, 3)) & ",'" & Trim(.CellText(Indice, 11)) & "'," & CantidadPiedras & "," & PesoPiedras & "," & CantidadDiamantes & "," & PuntosDiamantes & "," & crPrestamoDiamantes & ",'" & strMarca & "','" & strModelo & "','" & strNumSerie & "','" & strColor & "','" & strTamaño & "')"
                    End If

                Next Indice
            
                Limpiar frmSepararlote
                grdArticulos.Clear
                lblPrestamo.Caption = "0.00"
                lblAvaluo.Caption = "0.00"
                lblPeso.Caption = "0.00"
                txtNumContrato.SetFocus
            End With

        End If
    End If

End Sub

Private Sub cmdBuscar_Click()
    txtNumcontrato_KeyPress vbKeyReturn
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

Private Sub cmdLimpiar_Click()
    LimpiaArticulos
    If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
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
    Cargar_Combos "Descripcion", "tipo", cmbTipo, " WHERE (Kilataje=1 OR Peso=1)", "Ordenamiento"
    Cargar_Combos "Descripcion", "tipo", cmbTipoElec, " WHERE (Kilataje=0 OR Peso=0)", "Ordenamiento"
    Crear_Pestañas
    Crear_Encabezados
    lblPrestamo.Caption = "0.00"
    lblAvaluo.Caption = "0.00"
    lblPeso.Caption = "0.00"
    txtCantidad.text = "1"
    cmbTipo.ListIndex = 0
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Crear_Encabezados()

    With grdArticulos
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
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

'''''Private Sub grdArticulos_KeyUp(KeyCode As Integer, Shift As Integer)
'''''Dim Res As Integer, i As Integer
'''''
'''''    If grdArticulos.SelectedRow > 0 Then
'''''        If grdArticulos.CellText(grdArticulos.SelectedRow, 1) <> "" Then
'''''
'''''            If MsgBox("Desea eliminar la prenda seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Deslotificación") = vbNo Then
'''''                grdArticulos.ClearSelection
'''''                If TPrendas.SelectedTab = 1 Then cmbTipo.ListIndex = 0 Else cmbTipoElec.ListIndex = 0
'''''                Exit Sub
'''''            End If
'''''
'''''            grdArticulos.RemoveRow grdArticulos.SelectedRow
''''''''''            Res = 11 - grdArticulos.Rows
''''''''''
''''''''''            For i = 1 To Res
''''''''''                grdArticulos.AddRow
''''''''''            Next i
'''''
''''''''''            grdArticulos.ClearSelection
'''''            Total_Avaluos
'''''            If TPrendas.SelectedTab = 1 Then cmbTipo.ListIndex = 0 Else cmbTipoElec.ListIndex = 0
'''''
'''''        Else
'''''
'''''            grdArticulos.ClearSelection
'''''        End If
'''''
'''''    End If
'''''
'''''End Sub

Private Sub lblPrestamoDiamante_Change()
    Calcular_Avaluo
End Sub

Private Sub tPrendas_TabClick(ByVal lTab As Long)

    Select Case lTab

        Case 1
            
            BanElec = False
            LimpiaArticulos
            txtCantidad.text = "1"
            frmMetales.Visible = True
            frmElectronicos.Visible = False
            grdArticulos.Clear
            cmbTipo.ListIndex = ComboInformacion(cmbTipo, 1)
            cmbTipo.ListIndex = 0
        Case 2
            
            BanElec = True
            LimpiaArticulos
            frmElectronicos.Visible = True
            frmMetales.Visible = False
            grdArticulos.Clear
            cmbTipoElec.ListIndex = 0
    End Select
    
    Total_Avaluos
End Sub

Private Sub txtEmpeños_GotFocus()
    Cambiar_Color True, txtEmpeños
End Sub

Private Sub txtEmpeños_LostFocus()
    Cambiar_Color False, txtEmpeños
End Sub

'Calculamos el prestamo
Private Function Calcula_Prestamo(Prestamo As Double, PorcentajePrestamo As Double) As Double

On Error GoTo error:
    
    Calcula_Prestamo = Redondeo(Prestamo * (PorcentajePrestamo / 100))
    
error:
    Maneja_Error Err
   
End Function

'Calculamos el avaluo
Private Function Calcular_Avaluo() As Double
Dim crPrecio As Double, Peso As Double, PesoPiedra As Double, PrestamoDiamante As Double, PrestamoAvaluo As Double, AvaluoDiamante As Double
Dim rcTmp As New ADODB.Recordset

On Error GoTo error

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
    
        PrestamoAvaluo = Regresa_Valor_BD("PrestamoAvaluo")
    
        Calcular_Avaluo = (Peso - PesoPiedra) * crPrecio
        
        txtAvaluo.text = Format(Redondeo(Calcular_Avaluo + AvaluoDiamante), FMoneda)
        txtPrestamoo.text = Format(Calcula_Prestamo(CDbl(Calcular_Avaluo), PrestamoAvaluo) + PrestamoDiamante, FMoneda)
    End If

Set rcTmp = Nothing
Exit Function

error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

Private Sub txtNumcontrato_GotFocus()
    Seleccionar_Texto txtNumContrato
    Cambiar_Color True, txtNumContrato
End Sub

Private Sub txtNumcontrato_KeyPress(KeyAscii As Integer)
Dim rcConsulta As New ADODB.Recordset
Dim NumContrato As Long

On Error GoTo error

    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn And Trim(txtNumContrato.text) <> "" Then
            
            NumContrato = txtNumContrato.text
            Limpiar frmSepararlote
            txtNumContrato.text = NumContrato
            tPrendas_TabClick 1
            
            rcConsulta.Open "SELECT empeno.ID,clientes.Nombre,clientes.Apellido,clientes.Direccion,clientes.Colonia,clientes.Municipio,clientes.CP,clientes.Estado,clientes.Tel,clientes.Identificacion,empeno.Avaluo,empeno.Prestamo,empeno.TipoInteres,empeno.Serie " _
                            & "FROM empeno LEFT JOIN clientes ON empeno.IDCliente=clientes.ID WHERE empeno.NumContrato=" & Val(txtNumContrato.text) & " AND (empeno.Serie=" & SERIE_A & " OR empeno.Serie=" & SERIE_C & ") AND empeno.Pagado=0 AND empeno.Cancelado=0 AND empeno.Destino=0", dbDatos, adOpenForwardOnly, adLockOptimistic
            If Not rcConsulta.BOF And Not rcConsulta.EOF Then

                With rcConsulta
                    txtNumContrato.Tag = !ID
                    lblNumContrato.Caption = Val(txtNumContrato.text)
                    lblNombre.Caption = !Nombre
                    lblApellido.Caption = !apellido
                    lblDireccion.Caption = !Direccion
                    lblColonia.Caption = !Colonia
                    lblMunicipio.Caption = !Municipio
                    lblCP.Caption = !CP
                    lblEstado.Caption = !Estado
                    lblTelefono.Caption = !Tel
                    lblIdentificacion.Caption = !identificacion
                    lblTasa.Tag = !Serie
                    lblTotAvaluo.Caption = Format(!Avaluo, FMoneda)
                    lblTotPrestamo.Caption = Format(!Prestamo, FMoneda)
                    lblTotPeso.Caption = Format(SacaValor("detallesempeno", "SUM(Peso)", " WHERE IDEmpeno=" & !ID), "0.00")
                End With

            Else
                MsgBox "No se encontró el contrato especificado !!", vbCritical, "Deslotificación"
                txtNumContrato.SetFocus
            End If
            rcConsulta.Close
            Set rcConsulta = Nothing
            Exit Sub
    End If
    
error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub txtNumcontrato_LostFocus()
    Cambiar_Color False, txtNumContrato
End Sub

Private Sub Limpiar(Contededor As String)
    Dim ctrl As Control

    For Each ctrl In Controls
        On Error Resume Next

        If ctrl.Container.Caption = Contededor Then
            If TypeOf ctrl Is TextBox Then ctrl.text = "": ctrl.Tag = ""
            If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = "": ctrl.Tag = ""
            If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
            On Error Resume Next
            ctrl.Tag = ""
        End If

    Next

End Sub

Function SumaPeso() As Double
    Dim i As Integer, Peso As Double

    For i = 1 To grdArticulos.Rows
        Peso = Peso + IIf(grdArticulos.CellText(i, 4) = "", 0, grdArticulos.CellText(i, 4))
    Next i

    SumaPeso = Peso
End Function

Function SumaPrestamo() As Double
    Dim i As Integer, Prestamo As Double

    For i = 1 To grdArticulos.Rows
        Prestamo = Prestamo + IIf(grdArticulos.CellText(i, 7) = "", 0, grdArticulos.CellText(i, 7))
    Next i

    SumaPrestamo = Prestamo
End Function

Function RellenaCombo(ID As Long, Tabla As String, Campo As String, Combo As ComboBox, IDCriterio As String, Condicion As String)
    Dim rcTmp As ADODB.Recordset

    On Error GoTo error
    
    Combo.Clear
    Set rcTmp = dbDatos.Execute("SELECT Distinct PreciosKilatajes." & Campo & " as ID," & Tabla & ".Descripcion" & " FROM " & Tabla & " INNER JOIN PreciosKilatajes ON " & Tabla & ".ID = PreciosKilatajes." & Campo & " where PreciosKilatajes." & IDCriterio & "=" & ID & Condicion)
    
    If Not rcTmp.BOF And Not rcTmp.EOF Then
        rcTmp.MoveFirst
        While Not rcTmp.EOF
            Combo.AddItem rcTmp!Descripcion
            Combo.ItemData(Combo.NewIndex) = rcTmp!ID
            rcTmp.MoveNext
        Wend
    End If
    
error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

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

Private Sub Total_Avaluos()
Dim Indice As Integer, crAvaluo As Double, crPrestamo As Double, m_Peso As Double
    
    crPrestamo = 0
    crAvaluo = 0
    m_Peso = 0
   
    For Indice = 1 To grdArticulos.Rows
        crAvaluo = crAvaluo + IIf(Val(grdArticulos.CellText(Indice, 6)) > 0, grdArticulos.CellText(Indice, 6), 0)
        crPrestamo = crPrestamo + IIf(Val(grdArticulos.CellText(Indice, 7)) > 0, grdArticulos.CellText(Indice, 7), 0)
        m_Peso = m_Peso + IIf(Val(grdArticulos.CellText(Indice, 4)) > 0, grdArticulos.CellText(Indice, 4), 0)
    Next Indice
   
    lblPrestamo.Caption = Format(crPrestamo, FMoneda)
    lblAvaluo.Caption = Format(crAvaluo, FMoneda)
    lblPeso.Caption = Format(m_Peso, "0.000")
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

Private Sub txtAvaluoElec_GotFocus()
    Seleccionar_Texto txtAvaluoElec
    Cambiar_Color True, txtAvaluoElec
End Sub

Private Sub txtAvaluoElec_LostFocus()
    txtAvaluoElec.text = Format(txtAvaluoElec.text, FMoneda)
    Cambiar_Color False, txtAvaluoElec
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

Private Sub cmdMostrarCatPrendas_Click()
Dim IDPrenda As Long
Dim rcPrenda As New ADODB.Recordset
        
    IDPrenda = frmCatVarios.Mostrar(cmbTipoElec.ItemData(cmbTipoElec.ListIndex))
    If IDPrenda > 0 Then
        
        LimpiaArticulos
        With rcPrenda
            
            .Open "SELECT tipoprenda.Descripcion AS Desc_Familia,marcas.Descripcion AS Desc_Marca,prendaselec.ID AS IDPrenda,prendaselec.Modelo,prendaselec.Minimo,prendaselec.Maximo FROM prendaselec INNER JOIN tipoprenda ON prendaselec.IDFamilia=tipoprenda.ID INNER JOIN marcas ON prendaselec.IDMarca=marcas.ID WHERE prendaselec.ID=" & IDPrenda, dbDatos, adOpenForwardOnly, adLockOptimistic
            
            txtFamiliaElec.text = !Desc_Familia
            txtFamiliaElec.Tag = !IDPrenda
            txtMarcaElec.text = !Desc_Marca
            txtModeloElec.text = !Modelo
            txtPrestamooElec.Tag = !Maximo
            txtPrestamooElec.text = Format(!Minimo, FMoneda)
            
            .Close
            Set rcPrenda = Nothing
        
        End With
        
    End If

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
    Seleccionar_Texto cmbTipo
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
    grdArticulos.CancelEdit
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
    Seleccionar_Texto cmbPrenda
    Cambiar_Color True, cmbPrenda
End Sub

Private Sub cmbPrenda_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPrenda_LostFocus()
    Cambiar_Color False, cmbPrenda
End Sub

Private Sub cmbKilates_Click()
    If cmbKilates.ListIndex > -1 And cmbEstado.ListIndex > -1 Then Calcular_Avaluo
End Sub

Private Sub cmbKilates_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbKilates_GotFocus()
    Cambiar_Color True, cmbKilates
End Sub

Private Sub cmbKilates_LostFocus()
    Cambiar_Color False, cmbKilates
End Sub

Private Sub cmbEstado_Click()
    If cmbKilates.ListIndex > -1 And cmbEstado.ListIndex > -1 Then Calcular_Avaluo
End Sub

Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbEstado_GotFocus()
    Cambiar_Color True, cmbEstado
End Sub

Private Sub cmbEstado_LostFocus()
    Cambiar_Color False, cmbEstado
    grdArticulos.CancelEdit
End Sub

Private Sub txtCantidad_GotFocus()
    Seleccionar_Texto txtCantidad
    Cambiar_Color True, txtCantidad
End Sub

Private Sub txtCantidad_LostFocus()
    Cambiar_Color False, txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
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

Private Sub txtPrestamoo_Change()
Dim crPeso As Double, crPrecio As Double, crPrestamo As Double, Prestamo As Double, PorPrestamo As Double, PrestamoDiamante As Double, PesoPiedra As Double
Dim rcTmp As ADODB.Recordset

On Error GoTo error

    If cmbTipo.ListIndex > -1 And cmbPrenda.ListIndex > -1 Then
        
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
                        crPrecio = rcTmp!Precio
                    Else
                        
                        crPrecio = 0
                    End If
                                                    
                    PorPrestamo = Val(Regresa_Valor_BD("PrestamoAvaluo"))
                    crPrestamo = Redondeo((crPeso - PesoPiedra) * crPrecio) * (PorPrestamo / 100)
                    PorPrestamo = Val(Regresa_Valor_BD("Negociacion")) / 100
                    
                    If Val(txtPrestamoo.text) > 0 Or Trim(txtPrestamoo.text) <> "" And crPeso > 0 Then
                        
                        Prestamo = txtPrestamoo.text
                        If Prestamo > ((crPrestamo + (crPrestamo * PorPrestamo)) + PrestamoDiamante) Then
                            
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
    
error:
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

Private Sub cmdAgregar_Click()
Dim i As Integer, Estado As Integer, Kilates As Integer, PrecioVenta As Double, crPrestamo As Double
Dim IDTipo As Integer, IDTipoPrenda As Long, strPrenda As String, Piedras As Integer, PesoPiedras As Double

    If ValidaArticulos(TPrendas.SelectedTab) Then
    
        With grdArticulos
            
            If Val(txtCantidad.Tag) > 0 And TPrendas.SelectedTab = 1 Then i = Val(txtCantidad.Tag): GoTo Edicion
            
            .AddRow
            i = .Rows
Edicion:
                
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
                                    
            txtCantidad.text = "1"
            If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
        End With

    End If

End Sub

Private Sub cmdBorrar_Click()
Dim i As Integer, crPrestamo As Double

    If grdArticulos.SelectedRow > 0 Then
        
        If Trim(grdArticulos.CellText(grdArticulos.SelectedRow, 1)) <> "" Then
            
            If MsgBox("Desea eliminar la prenda seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton1, "Deslotificación") = vbYes Then
                
                grdArticulos.RemoveRow grdArticulos.SelectedRow
                Total_Avaluos
            End If
        
        End If
    
    End If
    
    grdArticulos.ClearSelection
    If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
End Sub

Sub Calcula_Avaluo_Elec()
Dim crPrestamo As Double, crMaximo As Double, PrestamoAvaluo As Double
    
    crPrestamo = 0
    crMaximo = 0
    If Val(txtPrestamooElec.text) > 0 Or Trim(txtPrestamooElec.text) <> "" Then
        
        'Tomo el prestamo
        crPrestamo = CDbl(txtPrestamooElec.text)
        
        'Tomo el importe máximo
        If Val(txtPrestamooElec.Tag) > 0 Or Trim(txtPrestamooElec.Tag) <> "" Then
        
            crMaximo = CDbl(txtPrestamooElec.Tag)
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

Public Function ValidaArticulos(Pestaña As Integer) As Boolean
Dim rcTmp As ADODB.Recordset

On Error GoTo error
    
    '''''If Pestaña = 1 Then Set rcTmp = dbDatos.Execute("SELECT Kilataje,Peso FROM tipo WHERE ID=" & cmbTipo.ItemData(cmbTipo.ListIndex))
    
    ValidaArticulos = True

    If IIf(Pestaña = 1, cmbTipo.text, cmbTipoElec.text) = "" Then
        MsgBox "Seleccione el tipo !!", vbInformation, "Deslotificación"
        ValidaArticulos = False
        If Pestaña = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
        Exit Function
    End If
       
    If IIf(Pestaña = 1, cmbPrenda.text, txtFamiliaElec.text) = "" Then ''''' pestaña = 1 And cmbPrenda.ListIndex = -1 Then
        MsgBox "Seleccione la " & IIf(Pestaña = 1, "prenda", "familia") & " !!", vbInformation, "Deslotificación"
        ValidaArticulos = False
        If Pestaña = 1 Then cmbPrenda.SetFocus Else txtFamiliaElec.SetFocus
        Exit Function
    End If

    If txtCantidad.text = "" And Pestaña = 1 Then
        MsgBox "Introduzca la cantidad !!", vbInformation, "Deslotificación"
        ValidaArticulos = False
        txtCantidad.SetFocus
        Exit Function
    End If
    
    If Pestaña = 1 Then
        
        Set rcTmp = dbDatos.Execute("SELECT Kilataje,Peso FROM tipo WHERE ID=" & cmbTipo.ItemData(cmbTipo.ListIndex))
        
        If cmbKilates.text = "" And rcTmp!Kilataje = 1 Then
            MsgBox "Seleccione el kilataje !!", vbInformation, "Deslotificación"
            ValidaArticulos = False
            cmbKilates.SetFocus
            Exit Function
        End If
        
        If cmbEstado.text = "" And rcTmp!Peso = 1 Then
            MsgBox "Seleccione la hechura !!", vbInformation, "Deslotificación"
            ValidaArticulos = False
            cmbEstado.SetFocus
            Exit Function
        End If
        
        If txtPeso.text = "" And rcTmp!Peso = 1 Then
            MsgBox "Introduzca el peso !!", vbInformation, "Deslotificación"
            ValidaArticulos = False
            txtPeso.SetFocus
            Exit Function
        End If
        
        Set rcTmp = Nothing
    End If
    
    If Pestaña = 2 And Trim(txtMarcaElec.text) = "" Then
        MsgBox "Introduzca la marca !!", vbInformation, "Deslotificación"
        ValidaArticulos = False
        txtMarcaElec.SetFocus
        Exit Function
    End If
    
    If IIf(Pestaña = 1, txtPrestamoo.text, txtPrestamooElec) = "" Then
        MsgBox "Introduzca el préstamo !!", vbInformation, "Deslotificación"
        ValidaArticulos = False
        If Pestaña = 1 Then txtPrestamoo.SetFocus Else txtPrestamooElec.SetFocus
        Exit Function
    End If
    
    Exit Function
    
error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function
