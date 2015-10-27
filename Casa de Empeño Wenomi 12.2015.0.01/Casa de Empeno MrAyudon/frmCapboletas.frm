VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCapboletas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura de contratos"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12810
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCapboletas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7305
   ScaleWidth      =   12810
   Begin VB.Frame frmEmpeño 
      Caption         =   "Captura de contratos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6750
      Left            =   60
      TabIndex        =   29
      Top             =   30
      Width           =   12750
      Begin VB.TextBox txtVencimiento 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   9555
         Locked          =   -1  'True
         TabIndex        =   117
         Top             =   1995
         Width           =   1350
      End
      Begin VB.TextBox txtNumContrato 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   10485
         TabIndex        =   115
         Top             =   2355
         Width           =   1350
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   330
         Left            =   10020
         Top             =   1620
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   582
         Orientation     =   0
         LineWidth       =   2
      End
      Begin VB.TextBox txtTasa 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   9120
         TabIndex        =   112
         Top             =   1680
         Width           =   855
      End
      Begin VB.TextBox txtFecha 
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
         Height          =   300
         Left            =   885
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   0
         Top             =   90
         Width           =   1380
      End
      Begin vbalTabStrip6.TabControl TPrendasPagos 
         Height          =   4110
         Left            =   4200
         TabIndex        =   105
         Top             =   2595
         Width           =   8505
         _ExtentX        =   15002
         _ExtentY        =   7250
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Frame frmPagos 
            Caption         =   "PAGOS"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   3675
            Left            =   30
            TabIndex        =   107
            Top             =   375
            Visible         =   0   'False
            Width           =   8430
            Begin vbAcceleratorGrid6.vbalGrid grdPagos 
               Height          =   3630
               Left            =   1320
               TabIndex        =   108
               Top             =   120
               Width           =   7035
               _ExtentX        =   12409
               _ExtentY        =   6403
               RowMode         =   -1  'True
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
               ScrollBarStyle  =   1
               DisableIcons    =   -1  'True
            End
            Begin DevPowerFlatBttn.FlatBttn cmdGenerar 
               Height          =   375
               Left            =   120
               TabIndex        =   109
               Top             =   0
               Width           =   1065
               _ExtentX        =   1879
               _ExtentY        =   661
               AlignCaption    =   4
               AlignPicture    =   2
               AutoSize        =   0   'False
               Caption         =   "    &Generar"
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
               Picture         =   "frmCapboletas.frx":000C
               PictureDisabled =   "frmCapboletas.frx":020C
            End
            Begin vbalIml6.vbalImageList lstIcons 
               Left            =   480
               Top             =   960
               _ExtentX        =   953
               _ExtentY        =   953
               Size            =   2296
               Images          =   "frmCapboletas.frx":0366
               Version         =   131072
               KeyCount        =   2
               Keys            =   "ÿ"
            End
         End
         Begin vbAcceleratorGrid6.vbalGrid grdEmpeños 
            Height          =   3675
            Left            =   30
            TabIndex        =   106
            Top             =   375
            Width           =   8430
            _ExtentX        =   14870
            _ExtentY        =   6482
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
      End
      Begin VB.TextBox txtNombre 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         MaxLength       =   20
         TabIndex        =   1
         Top             =   480
         Width           =   2970
      End
      Begin VB.TextBox txtMunicipio 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3735
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtColonia 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   405
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtEstado 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   735
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1545
         Width           =   1935
      End
      Begin VB.TextBox txtTelefono 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3735
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1545
         Width           =   1335
      End
      Begin VB.TextBox txtIdentificacion 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6540
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1545
         Width           =   1650
      End
      Begin VB.TextBox txtMensaje 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   105
         MaxLength       =   150
         TabIndex        =   75
         Top             =   6488
         Width           =   3870
      End
      Begin VB.TextBox txtResponsable 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   915
         MaxLength       =   30
         TabIndex        =   12
         Top             =   2265
         Width           =   4140
      End
      Begin VB.TextBox txtCP 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6615
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
      Begin VB.ComboBox cmbSexo 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmCapboletas.frx":0C7E
         Left            =   5880
         List            =   "frmCapboletas.frx":0C88
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1845
         Width           =   2325
      End
      Begin VB.TextBox txtEdad 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4380
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   74
         Top             =   1890
         Width           =   795
      End
      Begin VB.ComboBox cmbTipoInteres 
         Height          =   330
         ItemData        =   "frmCapboletas.frx":0CA1
         Left            =   8295
         List            =   "frmCapboletas.frx":0CA3
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   1290
         Width           =   1560
      End
      Begin VB.TextBox txtDireccion 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   975
         MaxLength       =   70
         TabIndex        =   3
         Top             =   840
         Width           =   7215
      End
      Begin VB.TextBox txtApellidos 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5175
         MaxLength       =   60
         TabIndex        =   2
         Top             =   480
         Width           =   3015
      End
      Begin VB.TextBox txtNumBolsa 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   6300
         MaxLength       =   15
         TabIndex        =   13
         Top             =   2265
         Width           =   1890
      End
      Begin VB.ComboBox cmbPlazos 
         Height          =   330
         ItemData        =   "frmCapboletas.frx":0CA5
         Left            =   11445
         List            =   "frmCapboletas.frx":0CA7
         Style           =   2  'Dropdown List
         TabIndex        =   72
         Top             =   1290
         Width           =   870
      End
      Begin VB.ComboBox cmbPeriodo 
         Height          =   330
         ItemData        =   "frmCapboletas.frx":0CA9
         Left            =   9930
         List            =   "frmCapboletas.frx":0CAB
         Style           =   2  'Dropdown List
         TabIndex        =   71
         Top             =   1290
         Width           =   1455
      End
      Begin vbalTabStrip6.TabControl TPrendas 
         Height          =   3705
         Left            =   0
         TabIndex        =   30
         Top             =   2520
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
         Begin VB.Frame frmMetales 
            Caption         =   "Metales"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2940
            Left            =   15
            TabIndex        =   34
            Top             =   315
            Width           =   3990
            Begin VB.TextBox txtPeso 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2895
               MaxLength       =   20
               TabIndex        =   19
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
               Left            =   2880
               MaxLength       =   20
               TabIndex        =   23
               Top             =   1635
               Width           =   1020
            End
            Begin VB.ComboBox cmbKilates 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmCapboletas.frx":0CAD
               Left            =   1065
               List            =   "frmCapboletas.frx":0CAF
               Style           =   2  'Dropdown List
               TabIndex        =   16
               Top             =   690
               Width           =   1110
            End
            Begin VB.TextBox txtCantidad 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1095
               MaxLength       =   3
               TabIndex        =   18
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
               TabIndex        =   22
               Top             =   1635
               Width           =   1020
            End
            Begin VB.TextBox txtObservaciones 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   60
               MaxLength       =   150
               MultiLine       =   -1  'True
               TabIndex        =   24
               Top             =   2175
               Width           =   3870
            End
            Begin VB.ComboBox cmbPrenda 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmCapboletas.frx":0CB1
               Left            =   1065
               List            =   "frmCapboletas.frx":0CC7
               Style           =   2  'Dropdown List
               TabIndex        =   15
               Top             =   375
               Width           =   2895
            End
            Begin VB.TextBox txtPesoPiedra 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1095
               MaxLength       =   20
               TabIndex        =   20
               Top             =   1335
               Width           =   1020
            End
            Begin VB.ComboBox cmbTipo 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmCapboletas.frx":0CE3
               Left            =   1065
               List            =   "frmCapboletas.frx":0CF9
               Style           =   2  'Dropdown List
               TabIndex        =   14
               Top             =   60
               Width           =   2130
            End
            Begin VB.ComboBox cmbEstado 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmCapboletas.frx":0D15
               Left            =   2865
               List            =   "frmCapboletas.frx":0D17
               Style           =   2  'Dropdown List
               TabIndex        =   17
               Top             =   690
               Width           =   1095
            End
            Begin VB.TextBox txtPiedras 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2895
               MaxLength       =   3
               TabIndex        =   21
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
               TabIndex        =   50
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
               TabIndex        =   49
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
               TabIndex        =   48
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
               TabIndex        =   47
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
               TabIndex        =   46
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
               TabIndex        =   45
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
               TabIndex        =   44
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
               TabIndex        =   43
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
               TabIndex        =   42
               Top             =   750
               Width           =   615
            End
            Begin VB.Label lblPiedra 
               BackColor       =   &H80000013&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2070
               TabIndex        =   41
               Top             =   1815
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblCantidadPiedras 
               BackColor       =   &H80000013&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3510
               TabIndex        =   40
               Top             =   1815
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblPuntos 
               BackColor       =   &H80000013&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   2820
               TabIndex        =   39
               Top             =   1815
               Visible         =   0   'False
               Width           =   600
            End
            Begin VB.Label lblPrestamoDiamante 
               BackColor       =   &H80000013&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3255
               TabIndex        =   38
               Top             =   1050
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.Label lblAvaluoDiamante 
               BackColor       =   &H80000013&
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   3540
               TabIndex        =   37
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
               TabIndex        =   36
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
               TabIndex        =   35
               Top             =   1365
               Width           =   1035
            End
         End
         Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
            Height          =   375
            Left            =   2070
            TabIndex        =   25
            Top             =   3285
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
            Picture         =   "frmCapboletas.frx":0D19
            PictureDisabled =   "frmCapboletas.frx":1083
         End
         Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
            Height          =   375
            Left            =   75
            TabIndex        =   33
            Top             =   3285
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
            Picture         =   "frmCapboletas.frx":11DD
         End
         Begin DevPowerFlatBttn.FlatBttn cmdDiamante 
            Height          =   375
            Left            =   1005
            TabIndex        =   32
            Top             =   3285
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
            Picture         =   "frmCapboletas.frx":12E1
            PictureDisabled =   "frmCapboletas.frx":1505
         End
         Begin DevPowerFlatBttn.FlatBttn cmdBorrar 
            Height          =   375
            Left            =   3015
            TabIndex        =   31
            Top             =   3285
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
            Picture         =   "frmCapboletas.frx":165F
            PictureDisabled =   "frmCapboletas.frx":1BB1
         End
         Begin VB.Frame frmElectronicos 
            Caption         =   "Electronicos"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2970
            Left            =   15
            TabIndex        =   51
            Top             =   315
            Width           =   3990
            Begin VB.ComboBox cmbTipoElec 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               ItemData        =   "frmCapboletas.frx":2783
               Left            =   1065
               List            =   "frmCapboletas.frx":2785
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   60
               Width           =   2130
            End
            Begin VB.TextBox txtObservacionesElec 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   720
               Left            =   60
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   60
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
               TabIndex        =   59
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
               TabIndex        =   58
               Top             =   1635
               Width           =   1095
            End
            Begin VB.TextBox txtModeloElec 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2895
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   57
               Top             =   727
               Width           =   1050
            End
            Begin VB.TextBox txtNumSerieElec 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1095
               MaxLength       =   80
               TabIndex        =   56
               Top             =   1335
               Width           =   2850
            End
            Begin VB.TextBox txtTamañoElec 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1095
               MaxLength       =   50
               TabIndex        =   55
               Top             =   1035
               Width           =   1020
            End
            Begin VB.TextBox txtColorElec 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   2895
               MaxLength       =   50
               TabIndex        =   54
               Top             =   1035
               Width           =   1050
            End
            Begin VB.TextBox txtFamiliaElec 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1095
               Locked          =   -1  'True
               MaxLength       =   80
               TabIndex        =   27
               Top             =   412
               Width           =   2850
            End
            Begin VB.TextBox txtMarcaElec 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1095
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   52
               Top             =   727
               Width           =   1020
            End
            Begin DevPowerFlatBttn.FlatBttn cmdMostrarCatPrendas 
               Height          =   270
               Left            =   3195
               TabIndex        =   53
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
               TabIndex        =   70
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
               TabIndex        =   69
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
               TabIndex        =   68
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
               TabIndex        =   67
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
               TabIndex        =   66
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
               TabIndex        =   65
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
               TabIndex        =   64
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
               TabIndex        =   63
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
               TabIndex        =   62
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
               TabIndex        =   61
               Top             =   1065
               Width           =   480
            End
         End
      End
      Begin Line3D.ucLine3D ucLine3D30 
         Height          =   360
         Index           =   0
         Left            =   9870
         Top             =   1275
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
         Top             =   1260
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D10 
         Height          =   300
         Index           =   12
         Left            =   10245
         Top             =   705
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
         Top             =   690
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D10 
         Height          =   30
         Index           =   2
         Left            =   8235
         Top             =   435
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D10 
         Height          =   1200
         Index           =   3
         Left            =   12300
         Top             =   420
         Width           =   75
         _ExtentX        =   132
         _ExtentY        =   2117
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D10 
         Height          =   1485
         Index           =   1
         Left            =   8235
         Top             =   450
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2619
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D10 
         Height          =   375
         Index           =   44
         Left            =   11400
         Top             =   1275
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   661
         Orientation     =   0
         LineWidth       =   2
      End
      Begin MSMask.MaskEdBox txtFecNacimiento 
         Height          =   240
         Left            =   1995
         TabIndex        =   10
         Top             =   1875
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
      Begin Line3D.ucLine3D ucLine3D10 
         Height          =   30
         Index           =   8
         Left            =   8235
         Top             =   1620
         Width           =   4125
         _ExtentX        =   7276
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
         Height          =   225
         Left            =   3825
         TabIndex        =   102
         Top             =   465
         Width           =   300
         _ExtentX        =   529
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
      Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
         Height          =   300
         Index           =   1
         Left            =   2295
         TabIndex        =   111
         Top             =   90
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
         Picture         =   "frmCapboletas.frx":2787
      End
      Begin Line3D.ucLine3D ucLine3D10 
         Height          =   30
         Index           =   0
         Left            =   8235
         Top             =   1920
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
         Height          =   300
         Index           =   0
         Left            =   10920
         TabIndex        =   118
         Top             =   2010
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
         Picture         =   "frmCapboletas.frx":289C
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         Index           =   11
         Left            =   8280
         TabIndex        =   116
         Top             =   2040
         Width           =   1185
      End
      Begin VB.Label lblNumContrato 
         AutoSize        =   -1  'True
         Caption         =   "Num. Contrato:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   8265
         TabIndex        =   114
         Top             =   2295
         Width           =   2190
      End
      Begin VB.Label Label28 
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
         Index           =   12
         Left            =   0
         TabIndex        =   110
         Top             =   105
         Width           =   795
      End
      Begin VB.Label Label28 
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
         Index           =   15
         Left            =   4230
         TabIndex        =   99
         Top             =   450
         Width           =   885
      End
      Begin VB.Label Label28 
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
         Index           =   14
         Left            =   15
         TabIndex        =   98
         Top             =   450
         Width           =   765
      End
      Begin VB.Label Label28 
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
         Index           =   16
         Left            =   15
         TabIndex        =   97
         Top             =   810
         Width           =   885
      End
      Begin VB.Label Label28 
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
         Index           =   18
         Left            =   2775
         TabIndex        =   96
         Top             =   1170
         Width           =   930
      End
      Begin VB.Label Label28 
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
         Index           =   17
         Left            =   15
         TabIndex        =   95
         Top             =   1170
         Width           =   345
      End
      Begin VB.Label Label70 
         AutoSize        =   -1  'True
         Caption         =   "Buscar:"
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
         TabIndex        =   94
         Top             =   2280
         Width           =   660
      End
      Begin VB.Label Label28 
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
         Index           =   20
         Left            =   15
         TabIndex        =   93
         Top             =   1530
         Width           =   690
      End
      Begin VB.Label Label28 
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
         Index           =   21
         Left            =   2775
         TabIndex        =   92
         Top             =   1530
         Width           =   855
      End
      Begin VB.Label Label28 
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
         Index           =   22
         Left            =   5175
         TabIndex        =   91
         Top             =   1530
         Width           =   1305
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
         TabIndex        =   90
         Top             =   6270
         Width           =   1545
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
         TabIndex        =   89
         Top             =   465
         Width           =   1395
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
         TabIndex        =   88
         Top             =   735
         Width           =   930
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Cotitular:"
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
         Left            =   15
         TabIndex        =   87
         Top             =   2235
         Width           =   870
      End
      Begin VB.Label Label28 
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
         Index           =   19
         Left            =   6255
         TabIndex        =   86
         Top             =   1170
         Width           =   300
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de nacimiento:"
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
         Index           =   23
         Left            =   15
         TabIndex        =   85
         Top             =   1890
         Width           =   1920
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Sexo:"
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
         Index           =   25
         Left            =   5325
         TabIndex        =   84
         Top             =   1890
         Width           =   510
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
         Left            =   3840
         TabIndex        =   83
         Top             =   1890
         Width           =   510
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
         Left            =   8505
         TabIndex        =   82
         Top             =   1020
         Width           =   1110
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
         TabIndex        =   81
         Top             =   735
         Width           =   1170
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
         Left            =   8325
         TabIndex        =   80
         Top             =   1680
         Width           =   645
      End
      Begin VB.Label Label87 
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
         TabIndex        =   79
         Top             =   465
         Width           =   1995
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
         Left            =   5160
         TabIndex        =   78
         Top             =   2235
         Width           =   1065
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
         TabIndex        =   77
         Top             =   1020
         Width           =   480
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
         TabIndex        =   76
         Top             =   1020
         Width           =   705
      End
      Begin VB.Label Label1 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   8250
         TabIndex        =   100
         Top             =   465
         Width           =   4095
      End
      Begin VB.Label Label55 
         BackColor       =   &H00404040&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   8250
         TabIndex        =   101
         Top             =   990
         Width           =   4095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Height          =   285
         Left            =   8280
         TabIndex        =   113
         Top             =   1650
         Width           =   795
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   10200
      TabIndex        =   28
      Top             =   6840
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Cancelar"
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
      Picture         =   "frmCapboletas.frx":29B1
      PictureDisabled =   "frmCapboletas.frx":2A27
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   11400
      TabIndex        =   103
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
      Picture         =   "frmCapboletas.frx":35F9
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   8970
      TabIndex        =   104
      Top             =   6840
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
      Object.ToolTipText     =   ""
      Picture         =   "frmCapboletas.frx":3B4B
   End
End
Attribute VB_Name = "frmCapboletas"
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
Dim m_Peso As Double, Bandera As Boolean, BanElec As Boolean, pIDUsuarioAutoriza As Integer, pTipoAutorizacion As Integer

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

Private Sub cmbEstado_Click()
    If cmbKilates.ListIndex > -1 And cmbEstado.ListIndex > -1 Then Calcular_Avaluo
End Sub

Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

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
    Cambiar_Color True, cmbPeriodo
End Sub

Private Sub cmbPeriodo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPeriodo_LostFocus()
    Cambiar_Color False, cmbPeriodo
End Sub

Private Sub cmbPlazos_Click()

    If Bandera = False Then
        
        SacaTasa CCur(txtPrestamo.Caption), cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), IIf(Val(txtNombre.Tag) = 0, False, True)
        Calcular_Avaluo
    
    Else
        
        SacaTasa CCur(txtPrestamo.Caption), cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), cmbPeriodo.ItemData(cmbPeriodo.ListIndex), cmbPlazos.ItemData(cmbPlazos.ListIndex), IIf(Val(txtNombre.Tag) = 0, False, True)
    End If

    Bandera = False
End Sub

Private Sub cmbPlazos_GotFocus()
    Cambiar_Color True, cmbPlazos
End Sub

Private Sub cmbPlazos_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPlazos_LostFocus()
    Cambiar_Color False, cmbPlazos
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

Private Sub cmbSexo_GotFocus()
    Cambiar_Color True, cmbSexo
End Sub

Private Sub cmbSexo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbSexo_LostFocus()
    Cambiar_Color False, cmbSexo
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

Private Sub cmbTipoInteres_Click()

    If cmbTipoInteres.ListIndex > -1 Then
        
        cmbPeriodo.Clear
        cmbPlazos.Clear
        Cargar_Combos "DISTINCT tipoperiodo.Descripcion", "configuraciontasas INNER JOIN tipoperiodo ON tipoperiodo.ID=configuraciontasas.IDTipoPeriodo", cmbPeriodo, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), "tipoperiodo.Ordenamiento", , "tipoperiodo.ID"
    
    End If
    
    If cmbPeriodo.ListCount > 0 Then cmbPeriodo.ListIndex = 0
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

Private Sub cmbTipo_Click()

    If cmbTipo.ListIndex > -1 Then
        
        Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Descripcion"
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
            '''''MuestraTasa crPrestamo, IIf(Val(txtNombre.Tag) = 0, False, True), lblTasa
            
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
                '''''MuestraTasa crPrestamo, IIf(Val(txtNombre.Tag) = 0, False, True), lblTasa
            
            End If
        
        End If
    
    End If
    
    grdEmpeños.ClearSelection
    If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
End Sub

Private Sub cmdDiamante_Click()

    If TPrendas.SelectedTab = 1 Then
        lblPiedra.Caption = ""
        lblPuntos.Caption = ""
        lblAvaluoDiamante.Caption = ""
        lblPrestamoDiamante.Caption = ""
        lblCantidadPiedras.Caption = ""
        frmDiamante.Show 1
    End If
End Sub

Private Sub cmdGenerar_Click()
Dim crPrestamo As Double, Tasa As Double, Almacenaje As Double, Seguro As Double, Iva As Double, plazo As Integer, Periodo As Integer, Meses As Integer, Fecha As String
    
    Fecha = ""
    crPrestamo = 0: Tasa = 0: Almacenaje = 0: Seguro = 0: Iva = 0
    
    If Val(txtTasa.text) > 0 Or Trim(txtTasa.text) <> "" Then
        
        Tasa = CDbl(txtTasa.text) / 100
    End If
    
    If Trim(txtFecha.text) <> "" Then
        
        Fecha = txtFecha.text
    End If
    
    If Val(txtPrestamo.Caption) > 0 Or Trim(txtPrestamo.Caption) <> "" Then
        
        crPrestamo = CDbl(txtPrestamo.Caption)
    End If
    
    Select Case cmbPeriodo.text
    Case "MENSUAL"
        
        Periodo = 30
        Meses = 1
    Case "QUINCENAL"
        
        Periodo = 15
        Meses = 2
    Case "SEMANAL"
        
        Periodo = 7
        Meses = 4
    End Select
    
    Almacenaje = 0 'CDbl(Mid(lblAlmacenaje.Caption, 1, Len(lblAlmacenaje.Caption) - 1)) / 100
    Seguro = 0 'CDbl(Mid(lblSeguro.Caption, 1, Len(lblSeguro.Caption) - 1)) / 100
    Iva = Regresa_Valor_BD("IVA") / 100 'CDbl(Mid(lblIva.Caption, 1, Len(lblIva.Caption) - 1)) / 100
    
    If crPrestamo > 0 And Fecha <> "" Then
        
        TasaFija crPrestamo, Tasa * (1 + Iva), Almacenaje * (1 + Iva), Seguro * (1 + Iva), Val(cmbPlazos.text) * Meses, Periodo, CDate(Fecha)
    End If
    
End Sub

Private Sub cmdLimpiar_Click()
    LimpiaArticulos
    If TPrendas.SelectedTab = 1 Then cmbTipo.SetFocus Else cmbTipoElec.SetFocus
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
    
    If Index = 1 Then
        
        txtFecha.text = frmCalendario.Fecha(txtFecha.text)
    Else
        
        txtVencimiento.text = frmCalendario.Fecha(txtVencimiento.text)
    End If

End Sub

Private Sub cmdMostrarCatPrendas_Click()
Dim IDPrenda As Long
Dim rcPrenda As New ADODB.Recordset
        
    IDPrenda = frmCatVarios.Mostrar(cmbTipoElec.ItemData(cmbTipoElec.ListIndex))
    If IDPrenda > 0 Then
        
        LimpiaArticulos
        rcPrenda.Open "SELECT tipoprenda.Descripcion AS Desc_Familia,marcas.Descripcion AS Desc_Marca,prendaselec.ID AS IDPrenda,prendaselec.Modelo,prendaselec.Minimo,prendaselec.Maximo FROM prendaselec INNER JOIN tipoprenda ON prendaselec.IDFamilia=tipoprenda.ID INNER JOIN marcas ON prendaselec.IDMarca=marcas.ID WHERE prendaselec.ID=" & IDPrenda, dbDatos, adOpenForwardOnly, adLockOptimistic
            
            txtFamiliaElec.text = rcPrenda!Desc_Familia
            txtFamiliaElec.Tag = rcPrenda!IDPrenda
            txtMarcaElec.text = rcPrenda!Desc_Marca
            txtModeloElec.text = rcPrenda!Modelo
            txtPrestamooElec.Tag = rcPrenda!Maximo
            txtPrestamooElec.text = Format(rcPrenda!Minimo, FMoneda)
            
        rcPrenda.Close
        Set rcPrenda = Nothing
    
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

Private Sub grdPagos_Click(ByVal lRow As Long, ByVal lCol As Long)

    If lCol = 1 And lRow > 0 Then
        
        If grdPagos.SelectedRow > 0 Then
            
            If grdPagos.CellItemData(lRow, 2) = 0 Then
                
                If grdPagos.CellIcon(lRow, lCol) = lstIcons.ItemIndex(1) Then
                    
                    grdPagos.CellIcon(lRow, lCol) = lstIcons.ItemIndex(2)
                Else
                
                    grdPagos.CellIcon(lRow, lCol) = lstIcons.ItemIndex(1)
                End If
            
            End If
            
        End If
        
    End If
    
End Sub

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
            grdEmpeños.Clear
            grdEmpeños.Rows = 11
            cmbTipo.ListIndex = 0
        Case 2
            
            BanElec = True
            LimpiaArticulos
            frmElectronicos.Visible = True
            frmMetales.Visible = False
            grdEmpeños.Clear
            grdEmpeños.Rows = 11
            cmbTipoElec.ListIndex = 0
    End Select
    
End Sub

Private Sub TPrendasPagos_BeforeClick(ByVal lTab As Long, bCancel As Boolean)
    Select Case lTab

        Case 1
            frmPagos.Visible = True
            
        Case 2
            
            frmPagos.Visible = False
    End Select
End Sub

Private Sub txtAvaluo_Change()
'''''Dim Kilataje As Boolean, Peso As Boolean, IDTipo As Integer, Avaluo As Double
'''''
'''''If Bandera = False Then
'''''
'''''    If cmbTipo.ListIndex > -1 Then
'''''        IDTipo = cmbTipo.ItemData(cmbTipo.ListIndex)
'''''    Else
'''''        IDTipo = 0
'''''    End If
'''''
'''''    If Val(txtAvaluo.text) > 0 Or Trim(txtAvaluo.text) <> "" Then
'''''        Avaluo = txtAvaluo.text
'''''    Else
'''''        Avaluo = 0
'''''    End If
'''''
'''''    Kilataje = Val(SacaValor("Tipo", "Kilataje", " where ID=" & IDTipo))
'''''    Peso = Val(SacaValor("Tipo", "Peso", " where ID=" & IDTipo))
'''''
'''''    If Kilataje And Peso Then
'''''
'''''        txtAvaluo.Locked = True
'''''    Else
'''''
'''''        txtAvaluo.Locked = False
'''''    End If
'''''End If
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
    txtAvaluo.text = Format(txtAvaluo.text, "###,###,###,###0.00")
End Sub

Private Sub txtAvaluoElec_GotFocus()
    Seleccionar_Texto txtAvaluoElec
    Cambiar_Color True, txtAvaluoElec
End Sub

Private Sub txtAvaluoElec_LostFocus()
    txtAvaluoElec.text = Format(txtAvaluoElec.text, "###,###,###,###0.00")
    Cambiar_Color False, txtAvaluoElec
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

Private Sub txtCP_GotFocus()
    Seleccionar_Texto txtCP
    Cambiar_Color True, txtCP
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = Solo_Numeros(KeyAscii)
End Sub

Private Sub txtCP_LostFocus()
    Cambiar_Color False, txtCP
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

Private Sub cmbTipo_DropDown()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmdAceptar_Click()
    
    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Empeño") = vbYes Then
        
        If Validar_Empeno = False Then Exit Sub
        
        'Grabo el Empeño
        Grabar_Empeno Val(txtNombre.Tag)
        
    End If
  
End Sub

'Grabamos el Empeno
Private Sub Grabar_Empeno(ID As Long)
Dim strSql As String, Contrato As Long, Folio As Long, Movimiento As Long, Prestamo As Double, crPrestamoVigente As Double, Vencimiento As String, IDEmpeno As Long
Dim Tasa As Double, Kilates As Integer, Estado As Integer, Indice As Integer, Almacenaje As Double, Seguro As Double, Comision As Double, Fecha As String
Dim Iva As Integer, IDCliente As Long, Dias As Integer, GTOOperacion As Double, VenAlmoneda As Integer, strMarca As String, strModelo As String, strNumSerie As String, strColor As String, strTamaño As String
Dim strIniciales As String, Codigo As String, crPrestamoPrenda As Double, Peso As Double, CantidadPiedras As Integer, PesoPiedras As Double, CantidadDiamantes As Integer, PuntosDiamantes As Double, crPrestamoDiamantes As Double, Periodo As Integer, VenPeriodo As Integer, Promocion As Integer

On Error GoTo error
    
    Screen.MousePointer = vbHourglass
    
    If ID = 0 Then
       ID = Grabar_Cliente
       IDCliente = ID
    Else
       IDCliente = ID
       Actualizar_Cliente IDCliente
    End If
    
    'Actualizo el Numero de contratos del cliente
    dbDatos.Execute "UPDATE clientes SET Boletas=Boletas+1 WHERE ID=" & IDCliente
    
    'Saco el Numero de Contrato
    Contrato = txtNumContrato.text
    
    'Folio
    Folio = Contrato
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    Select Case cmbPeriodo
    Case "MENSUAL"
    
        Periodo = 30
    Case "QUINCENAL"
        
        Periodo = 15
    Case "SEMANAL"
        
        Periodo = 7
    Case "DIARIA"
        
        Periodo = 1
    End Select
    
    Fecha = CDate(txtFecha.text)
    crPrestamoVigente = PrestamoVigente
    VenPeriodo = Val(cmbPlazos.text)
    Tasa = CDbl(txtTasa.text) 'CDbl(Mid(lblTasa.Caption, 1, Len(lblTasa.Caption) - 1))
    Prestamo = CDbl(txtPrestamo.Caption)
    Almacenaje = 0 'CDbl(Mid(lblAlmacenaje.Caption, 1, Len(lblAlmacenaje.Caption) - 1)) ' Regresa_Valor_BD("Almacenaje")
    Seguro = 0 'CDbl(Mid(lblSeguro.Caption, 1, Len(lblSeguro.Caption) - 1)) 'Regresa_Valor_BD("Seguro")
    Iva = Regresa_Valor_BD("IVA") 'CDbl(Mid(lblIva.Caption, 1, Len(lblIva.Caption) - 1)) 'Regresa_Valor_BD("IVA")
    GTOOperacion = 0 'Regresa_Valor_BD("Operacion")
    Comision = Regresa_Valor_BD("Comision")
    VenAlmoneda = Regresa_Valor_BD("VenAlmoneda")
    strIniciales = Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text))
        
    If cmbEstado.ListIndex = -1 Then
        
        Estado = 0
    Else
        
        Estado = cmbEstado.ItemData(cmbEstado.ListIndex)
    End If
    
    'Grabo en la tabla de Empeños
    strSql = "INSERT INTO empeno (Fecha,Movimiento,Numcontrato,Folio,Prestamo,Avaluo,Origen,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Responsable,Valuador,Notas,Tasa,Almacenaje,Seguro,Operacion,Comision,IVA,Periodo,Venperiodo,VenAlmoneda,Tipointeres,TipoTasa,IDSucursal,IDUsuario,IDAutorizacion,NumBolsa,Ubicacion,Caja,Cajon,Fila,IDUsuarioAutoriza,TipoAutoriza,Promocion,Captura,PrestamoInicial) VALUES " & _
           "('" & Format(Fecha, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & Contrato & "," & Folio & "," & Prestamo & "," & Val(lblTotAvaluo.Tag) & "," & OD_EMPENO & ",'" & Format(txtVencimiento.text, "YYYY/MM/DD") & "'," & Folio & "," & IIf(cmbTipoInteres.text = "FIJA", SERIE_C, SERIE_A) & ",'" & NombrePc & "'," & IDCliente & "," _
           & "'" & Trim(txtResponsable.text) & "','" & frmMDI.Usuario & "',''," & Tasa & "," & Almacenaje & "," & Seguro & "," & GTOOperacion & "," & Comision & "," & Iva & "," & Periodo & "," & VenPeriodo & "," & VenAlmoneda & ",'" & cmbTipoInteres.text & "','" & cmbPeriodo.text & "'," & frmMDI.IDSucursal & "," & frmMDI.IDUsuario & ",0,'" & Trim(txtNumBolsa.text) & "','','','',''," & Me.IDUsuarioAutoriza & "," & Me.TipoAutorizacion & "," & Promocion & ",1," & Prestamo & ")"
   
    Err.Clear
    dbDatos.Execute strSql
    
    'Saco el ID del Empeño
    IDEmpeno = SacaValor("empeno", "MAX(ID)")
    
    'Grabo los pagos
    GrabaPagos IDEmpeno
    
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
                
                'Recalculo el préstamo de las prendas
                crPrestamoPrenda = CDbl(.CellText(Indice, 7))
'''''                crPrestamoPrenda = Redondeo(crPrestamoVigente * (crPrestamoPrenda / 100))

                strMarca = Trim(.CellText(Indice, 17))
                strModelo = Trim(.CellText(Indice, 18))
                strNumSerie = Trim(.CellText(Indice, 19))
                strColor = Trim(.CellText(Indice, 20))
                strTamaño = Trim(.CellText(Indice, 21))
                
                dbDatos.Execute "INSERT INTO detallesempeno (IDEmpeno,Codigo,Tipo,Cantidad,Articulo,Peso,Kilates,Avaluo,Prestamo,Estado,Origen,Destino,TipoPrenda,Observaciones,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante,Marca,Modelo,Serie,Color,Tamano) VALUES (" & _
                                IDEmpeno & ",'" & Trim(Codigo) & "'," & .CellItemData(Indice, 1) & "," & Val(.CellText(Indice, 2)) & ",'" & Trim(UCase(.CellText(Indice, 3))) & "'," & Peso & "," & _
                                Kilates & "," & CDbl(.CellText(Indice, 6)) & "," & crPrestamoPrenda & ",'" & Trim(UCase(.CellText(Indice, 9))) & "'," & ENTRADAEMPENO & ",0," & Val(.CellItemData(Indice, 3)) & ",'" & Trim(.CellText(Indice, 11)) & "'," & CantidadPiedras & "," & PesoPiedras & "," & CantidadDiamantes & "," & PuntosDiamantes & "," & crPrestamoDiamantes & ",'" & strMarca & "','" & strModelo & "','" & strNumSerie & "','" & strColor & "','" & strTamaño & "')"
            End If
            
        Next Indice
    
    End With
    
'''''    'Grabamos el cargo
'''''    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'''''                    & "('" & Format(Fecha, "YYYY/MM/DD") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','201701'," & crPrestamoVigente & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'''''
'''''    'Grabamos el abono
'''''    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'''''                    & "('" & Format(Fecha, "YYYY/MM/DD") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110150'," & crPrestamoVigente & "," & TIPO_ABONO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'''''
'''''    'Grabamos abono 199450
'''''    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'''''                    & "('" & Format(Fecha, "YYYY/MM/DD") & "','Empeño'," & Movimiento & "," & Folio & ",'" & strIniciales & "','199450'," & crPrestamoVigente & "," & TIPO_ABONO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

    Limpiar "Captura de contratos"
    txtPrestamo.Caption = ""
    grdEmpeños.CancelEdit
    grdEmpeños.ClearItems
    grdEmpeños.ClearSelection
    grdPagos.Clear
    grdPagos.ClearSelection
    IDUsuarioAutoriza = 0
    TipoAutorizacion = 0
    txtPrestamo.Caption = "0.00"
    lblTotAvaluo.Caption = "0.00"
    cmbTipoInteres.ListIndex = 0
    cmbTipoInteres_Click
    cmbTipo.ListIndex = 0
    Default 1
    txtNombre.SetFocus
    
    
error:
    Maneja_Error Err
    Screen.MousePointer = vbDefault
End Sub

'Validamos que esten los datos correctos en Empeno
Private Function Validar_Empeno() As Boolean
Dim i As Integer, x As Integer

    Validar_Empeno = True
    
    'Checo el folio
    If Trim(txtNumContrato.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtNumContrato.SetFocus
        Exit Function
    End If
    
    'Checo la fecha
    If Trim(txtFecha.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtFecha.SetFocus
        Exit Function
    End If
    
     'Checo la fecha de Vencimiento
    If Trim(txtVencimiento.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtVencimiento.SetFocus
        Exit Function
    End If

    'si no tiene nombre
    If Trim(txtNombre.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtNombre.SetFocus
        Exit Function
    End If
  
    'si no tiene apellido
    If Trim(txtApellidos.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtApellidos.SetFocus
        Exit Function
    End If
    
    '''''    'si no tiene direccion
    '''''    If Trim(txtDireccion.Text) = "" Then
    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
    '''''        Validar_Empeno = False
    '''''        txtDireccion.SetFocus
    '''''        Exit Function
    '''''    End If
    '''''
    '''''    'si no tiene estado
    '''''    If Trim(txtEstado.Text) = "" Then
    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
    '''''        Validar_Empeno = False
    '''''        txtEstado.SetFocus
    '''''        Exit Function
    '''''    End If
    '''''
    '''''    'si no tiene colonia
    '''''    If Trim(txtColonia.Text) = "" Then
    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
    '''''        Validar_Empeno = False
    '''''        txtColonia.SetFocus
    '''''        Exit Function
    '''''    End If
    '''''
    '''''    'si no tiene municipio
    '''''    If Trim(txtMunicipio.Text) = "" Then
    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
    '''''        Validar_Empeno = False
    '''''        txtMunicipio.SetFocus
    '''''        Exit Function
    '''''    End If
    '''''
    '''''    'si no tiene cp
    '''''    If Trim(txtCp.Text) = "" Then
    '''''        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
    '''''        Validar_Empeno = False
    '''''        txtCp.SetFocus
    '''''        Exit Function
    '''''    End If
    
    'Si no identificacion
    If Trim(txtIdentificacion.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtIdentificacion.SetFocus
        Exit Function
    End If
        
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
        
'''''    If Trim(txtNumBolsa.text) = "" And BanElec = False Then
'''''        MsgBox "Introduzca el número de bolsa !!", vbCritical + vbOKOnly, "Empeño"
'''''        Validar_Empeno = False
'''''        txtNumBolsa.SetFocus
'''''        Exit Function
'''''    End If
    
    If Trim(txtTasa.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Empeno = False
        txtTasa.SetFocus
        Exit Function
    End If
    
'''''    'Checo si el grid de empeños tiene información
'''''    x = 0
'''''    For i = 1 To grdEmpeños.Rows
'''''
'''''        If Val(grdEmpeños.CellText(i, 2)) = 0 Then x = x + 1
'''''    Next i
'''''
'''''    If x = grdEmpeños.Rows Then
'''''        MsgBox "Favor de agregar las prendas al contrato !!", vbInformation, "Empeño"
'''''        Validar_Empeno = False
'''''        Exit Function
'''''    End If
'''''
'''''    If grdPagos.Rows = 0 Then
'''''        MsgBox "Favor de generar los pagos del contrato !!", vbInformation, "Empeño"
'''''        Validar_Empeno = False
'''''        Exit Function
'''''    End If
    
End Function

Private Sub cmdAceptar_GotFocus()
    cmdAceptar.BackColor = vb3DShadow
End Sub

Private Sub cmdAceptar_LostFocus()
    cmdAceptar.BackColor = vbButtonFace
End Sub

Private Sub cmdMosCliente_Click()
    frmMostrarCliente.Ver Me, txtNombre, True, 0
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
    frmEmpeño.BorderStyle = 0
    frmMetales.BorderStyle = 0
    frmPagos.BorderStyle = 0
    frmElectronicos.BorderStyle = 0
    Limpiar "Empeno"
    Crear_Pestañas
    Crear_Encabezados
    Bandera = True
    BanElec = False
    Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Descripcion='TRADICIONAL' AND Serie=" & SERIE_A, "Ordenamiento"
    Cargar_Combos "Descripcion", "tipo", cmbTipo, " WHERE ID=1", "Ordenamiento"
    Cargar_Combos "Descripcion", "tipo", cmbTipoElec, " WHERE ID<>1", "Ordenamiento"
    txtPrestamo.Caption = "0.00"
    lblTotAvaluo.Caption = "0.00"
    cmbTipoInteres.ListIndex = 0
    cmbTipo.ListIndex = 0
    Default 1
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

'Creamos las pestañas del tab
Private Sub Crear_Pestañas()
    
    With TPrendas
        .AddTab "Metales", , , "K1"
        .AddTab "Electrónicos/Varios", , , "K4"
    End With
    
    With TPrendasPagos
        .AddTab "Prendas"
'''''        .AddTab "Pagos Fijos"
    End With
    
End Sub

'Creamos los encabezados del grid
Private Sub Crear_Encabezados()
   
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
    
    With grdPagos
        .ImageList = lstIcons
        .AddColumn "K1", "Vencimiento", ecgHdrTextALignRight, , 99, , , , , "DD/MMM/YYYY", , CCLSortDate
        .AddColumn "K2", "Interes", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K3", "Almacenaje", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K4", "Seguro", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K5", "Pago Fijo", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Saldo", ecgHdrTextALignRight, , 80, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Amortizacion", ecgHdrTextALignRight, , 80, False, , , , FMoneda, , CCLSortNumeric
    End With
End Sub

Private Sub Form_LostFocus()
    frmMDI.Com.PortOpen = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Quitar_Flat Fl
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

Private Sub txtFecha_GotFocus()
    Seleccionar_Texto txtFecha
    Cambiar_Color True, txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFecha_LostFocus()
    Cambiar_Color False, txtFecha
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
    End If
End Sub

'''''Private Sub txtNumcontrato_GotFocus()
'''''    Seleccionar_Texto txtNumContrato
'''''    Cambiar_Color True, txtNumContrato
'''''End Sub
'''''
'''''Private Sub txtNumcontrato_KeyPress(KeyAscii As Integer)
'''''    KeyAscii = Solo_Numeros(KeyAscii)
'''''    Pasar_Foco KeyAscii
'''''End Sub
'''''
'''''Private Sub txtNumcontrato_LostFocus()
'''''    Cambiar_Color False, txtNumContrato
'''''End Sub

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
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
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

Private Sub txtNumcontrato_GotFocus()
    Seleccionar_Texto txtNumContrato
    Cambiar_Color True, txtNumContrato
End Sub

Private Sub txtNumcontrato_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumcontrato_LostFocus()
    Cambiar_Color False, txtNumContrato
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
    KeyAscii = Solo_Numeros(KeyAscii, 1)
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
    KeyAscii = Solo_Numeros(KeyAscii, 1)
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
            On Error Resume Next
            ctrl.Tag = ""
        End If

    Next

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

Private Sub txtPrestamoo_Change()
'''''Dim crPeso As Double, crPrecio As Double, crPrestamo As Double, Prestamo As Double, PorPrestamo As Double, PrestamoDiamante As Double, PesoPiedra As Double
'''''Dim rcTmp As ADODB.Recordset
'''''
'''''On Error GoTo error
'''''
'''''    If cmbTipo.ListIndex > -1 And cmbPrenda.ListIndex > -1 Then
'''''
'''''        Set rcTmp = dbDatos.Execute("SELECT Kilataje,Peso FROM Tipo WHERE ID=" & cmbTipo.ItemData(cmbTipo.ListIndex))
'''''        If Not rcTmp.BOF And Not rcTmp.EOF Then
'''''
'''''            If rcTmp!Kilataje = 1 Or rcTmp!Peso = 1 Then
'''''
'''''                If cmbKilates.ListIndex > -1 Then
'''''
'''''                    If Val(txtPeso.text) > 0 Or (Trim(txtPeso.text) <> "" And Trim(txtPeso.text) <> ".") Then
'''''
'''''                        crPeso = txtPeso.text
'''''                    Else
'''''
'''''                        crPeso = 0
'''''                    End If
'''''
'''''                    If Val(txtPesoPiedra.text) > 0 Or (Trim(txtPesoPiedra.text) <> "" And Trim(txtPesoPiedra.text) <> ".") Then
'''''
'''''                        PesoPiedra = CDbl(txtPesoPiedra.text)
'''''                    Else
'''''
'''''                        PesoPiedra = 0
'''''                    End If
'''''
'''''                    If Val(lblPrestamoDiamante.Caption) > 0 Or (Trim(lblPrestamoDiamante.Caption) <> "" And Trim(lblPrestamoDiamante.Caption) <> ".") Then
'''''
'''''                        PrestamoDiamante = lblPrestamoDiamante.Caption
'''''                    Else
'''''
'''''                        PrestamoDiamante = 0
'''''                    End If
'''''
'''''                    If cmbTipo.ListIndex >= 0 And cmbKilates.ListIndex >= 0 And cmbEstado.ListIndex >= 0 Then
'''''
'''''                        Set rcTmp = dbDatos.Execute("SELECT Precio FROM PreciosKilataje WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND IDKilataje=" & cmbKilates.ItemData(cmbKilates.ListIndex) & " AND IDHechura=" & cmbEstado.ItemData(cmbEstado.ListIndex))
'''''                        crPrecio = rcTmp!Precio
'''''                    Else
'''''
'''''                        crPrecio = 0
'''''                    End If
'''''
'''''                    PorPrestamo = Val(Regresa_Valor_BD("PrestamoAvaluo"))
'''''                    crPrestamo = Redondeo((crPeso - PesoPiedra) * crPrecio) * (PorPrestamo / 100)
'''''                    PorPrestamo = Val(Regresa_Valor_BD("Negociacion")) / 100
'''''
'''''                    If Val(txtPrestamoo.text) > 0 Or Trim(txtPrestamoo.text) <> "" And crPeso > 0 Then
'''''
'''''                        Prestamo = txtPrestamoo.text
'''''                        If Prestamo > ((crPrestamo + (crPrestamo * PorPrestamo)) + PrestamoDiamante) Then
'''''
'''''                            MsgBox "El préstamo sobrepasa el margen de negociación !!", vbInformation, "Empeño"
'''''                            Calcular_Avaluo
'''''                        End If
'''''
'''''                    End If
'''''
'''''                Else
'''''
'''''                    txtPrestamoo.text = "0.00"
'''''                End If
'''''
'''''            Else
'''''
'''''                If Val(txtPrestamoo.text) > 0 Then
'''''
'''''                    crPrestamo = txtPrestamoo.text
'''''                Else
'''''
'''''                    crPrestamo = 0
'''''                End If
'''''
'''''                txtAvaluo.text = Format(Calcular_Avaluo(), FMoneda)
'''''            End If
'''''
'''''        End If
'''''
'''''    End If
'''''    Set rcTmp = Nothing
'''''    Exit Sub
'''''
'''''error:
'''''    Maneja_Error Err
'''''    Set rcTmp = Nothing
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
    txtPrestamoo.text = Format(txtPrestamoo.text, "###,###,###,###0.00")
End Sub

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

Private Sub txtTasa_GotFocus()
    Seleccionar_Texto txtTasa
    Cambiar_Color True, txtTasa
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTasa_LostFocus()
    Cambiar_Color False, txtTasa
End Sub

Private Sub txtTelefono_GotFocus()
    Seleccionar_Texto txtTelefono
    Cambiar_Color True, txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTelefono_LostFocus()
    Cambiar_Color False, txtTelefono
End Sub

'Grabamos los datos del cliente
Private Function Grabar_Cliente() As Long
Dim rc As New ADODB.Recordset
Dim Medio As Integer, FechaNac As String, Sexo As Integer
   
On Error GoTo error
           
    Medio = 0
    
    If Trim(txtFecNacimiento.text) = "" Or Trim(txtFecNacimiento.text) = "__/__/____" Then
        
        FechaNac = "Null"
    Else
        
        FechaNac = "'" & Format(txtFecNacimiento.text, "YYYY/MM/DD") & "'"
    End If
   
    If cmbSexo.ListIndex = -1 Then
        
        Sexo = 0
    Else
        
        Sexo = cmbSexo.ItemData(cmbSexo.ListIndex)
    End If
   
    dbDatos.Execute "INSERT INTO clientes (Nombre,Apellido,Iniciales,Direccion,Colonia,Municipio,Estado,Tel,Identificacion,IDMedio,Notas,CP,FecNac,Sexo) VALUES ('" & _
                    Trim(txtNombre.text) & "','" & Trim(txtApellidos.text) & "','" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "','" & Trim(txtDireccion.text) & "','" & Trim(txtColonia.text) & "','" & Trim(txtMunicipio.text) & "','" & Trim(txtEstado.text) & "','" & Trim(txtTelefono.text) & "','" & Trim(txtIdentificacion.text) & "'," & Medio & ",'" & Trim(txtMensaje.text) & "','" & txtCP.text & "'," & FechaNac & "," & Sexo & ")"
      
    rc.Open "SELECT MAX(ID) AS IDD FROM clientes", dbDatos, adOpenForwardOnly, adLockOptimistic
   
        Grabar_Cliente = rc!idd
   
    rc.Close
    Set rc = Nothing
    Exit Function
    
error:
    Maneja_Error Err
    Set rc = Nothing
End Function

'actualizamos los datos del cliente
Private Sub Actualizar_Cliente(ID As Long)
Dim Medio As Integer, FechaNac As String, Sexo As Integer
   
On Error GoTo error

    Medio = 0
        
    If Trim(txtFecNacimiento.text) = "" Or Trim(txtFecNacimiento.text) = "__/__/____" Then
        
        FechaNac = "Null"
    Else
        
        FechaNac = "'" & Format(txtFecNacimiento.text, "YYYY/MM/DD") & "'"
    End If
   
    If cmbSexo.ListIndex = -1 Then
        
        Sexo = 0
    Else
        
        Sexo = cmbSexo.ItemData(cmbSexo.ListIndex)
    End If
   
    dbDatos.Execute "UPDATE clientes SET nombre='" & Trim(txtNombre.text) & "',apellido='" & Trim(txtApellidos.text) & "',Iniciales='" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "',Direccion='" & Trim(txtDireccion.text) & "',Colonia='" & Trim(txtColonia.text) & "',Municipio='" & Trim(txtMunicipio.text) & "'," & _
                    "Estado='" & Trim(txtEstado.text) & "',Tel='" & Trim(txtTelefono.text) & "',Identificacion='" & Trim(txtIdentificacion.text) & "',IDMedio=" & Medio & ",Notas='" & Trim(txtMensaje.text) & "',CP='" & txtCP.text & "',FecNac=" & FechaNac & ",Sexo=" & Sexo & " WHERE ID = " & ID
    Exit Sub

error:
    Maneja_Error Err
End Sub

'Buscamos el ID del cliente
Public Sub Buscar_Cliente(ID As Long)
Dim rcClientes As New ADODB.Recordset
   
On Error GoTo error

    rcClientes.Open "SELECT * FROM clientes WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
   
        With rcClientes
            txtNombre.text = !Nombre
            txtNombre.Tag = ID
            txtApellidos.text = !apellido
            txtDireccion.text = IIf(IsNull(!Direccion), "", !Direccion)
            txtColonia.text = IIf(IsNull(!Colonia), "", !Colonia)
            txtMunicipio.text = IIf(IsNull(!Municipio), "", !Municipio)
            txtEstado.text = IIf(IsNull(!Estado), "", !Estado)
            txtTelefono.text = IIf(IsNull(!Tel), "", !Tel)
            txtIdentificacion.text = IIf(IsNull(!identificacion), "", !identificacion)
            txtCP.text = IIf(IsNull(!CP), "", !CP)
            txtMensaje.text = IIf(IsNull(!notas), "", !notas)
            If Not IsNull(!FecNac) Then
                txtFecNacimiento.text = !FecNac
                txtEdad.text = Calcula_Edad(!FecNac)
            Else
                txtEdad.text = ""
                txtFecNacimiento.Mask = ""
                txtFecNacimiento.text = ""
                txtFecNacimiento.Mask = "##/##/####"
            End If
            cmbSexo.ListIndex = ComboInformacion(cmbSexo, IIf(IsNull(!Sexo), -1, !Sexo))
            '''''MuestraTasa CCur(txtPrestamo.Caption), True, lblTasa
        End With
        
    rcClientes.Close
    Set rcClientes = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcClientes = Nothing
End Sub

'Calculamos el avaluo
Private Function Calcular_Avaluo() As Double
Dim crPrecio As Double, Peso As Double, PesoPiedra As Double, PrestamoDiamante As Double, PrestamoAvaluo As Double, AvaluoDiamante As Double
Dim rcTmp As New ADODB.Recordset

On Error GoTo error

'''''    If cmbTipo.ListIndex >= 0 And cmbKilates.ListIndex >= 0 And cmbEstado.ListIndex >= 0 Then
'''''
'''''        If Val(txtPeso.text) > 0 Or (Trim(txtPeso.text) <> "" And Trim(txtPeso.text) <> ".") Then
'''''
'''''            Peso = CDbl(txtPeso.text)
'''''        Else
'''''
'''''            Peso = 0
'''''        End If
'''''
'''''        If Val(txtPesoPiedra.text) > 0 Or (Trim(txtPesoPiedra.text) <> "" And Trim(txtPesoPiedra.text) <> ".") Then
'''''
'''''            PesoPiedra = CDbl(txtPesoPiedra.text)
'''''        Else
'''''
'''''            PesoPiedra = 0
'''''        End If
'''''
'''''        If Val(lblPrestamoDiamante.Caption) > 0 Or Trim(lblPrestamoDiamante.Caption) <> "" Then
'''''
'''''            PrestamoDiamante = CDbl(lblPrestamoDiamante.Caption)
'''''        Else
'''''
'''''            PrestamoDiamante = 0
'''''        End If
'''''
'''''        If Val(lblAvaluoDiamante.Caption) > 0 Or Trim(lblAvaluoDiamante.Caption) <> "" Then
'''''
'''''            AvaluoDiamante = CDbl(lblAvaluoDiamante.Caption)
'''''        Else
'''''
'''''            AvaluoDiamante = 0
'''''        End If
'''''
'''''        rcTmp.Open "SELECT Precio FROM PreciosKilataje WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " AND IDKilataje=" & RegresaKilates(cmbKilates.text, cmbTipo.text) & " AND IDHechura=" & cmbEstado.ItemData(cmbEstado.ListIndex), dbDatos, adOpenForwardOnly, adLockOptimistic
'''''
'''''        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Precio) Then
'''''
'''''            crPrecio = rcTmp!Precio
'''''        Else
'''''
'''''            crPrecio = 0
'''''        End If
'''''
'''''        rcTmp.Close
'''''
'''''        PrestamoAvaluo = Regresa_Valor_BD("PrestamoAvaluo")
'''''
'''''        Calcular_Avaluo = (Peso - PesoPiedra) * crPrecio
'''''
'''''        txtAvaluo.text = Format(Redondeo(Calcular_Avaluo + AvaluoDiamante), FMoneda)
'''''        txtPrestamoo.text = Format(Calcula_Prestamo(CDbl(Calcular_Avaluo), PrestamoAvaluo) + PrestamoDiamante, FMoneda)
'''''    End If

    Set rcTmp = Nothing
    Exit Function

error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

'Calculamos el total de los avaluos y prestamos
Private Sub Total_Avaluos()
Dim Indice As Integer, Peso As Double, crAvaluo As Double, crPrestamo As Double, crAvaluoTotal As Double, crPrestamoTotal As Double, Cantidad As Integer

    m_Peso = 0
   
    For Indice = 1 To grdEmpeños.Rows
        
        Cantidad = Val(grdEmpeños.CellText(Indice, 2))
        crAvaluo = IIf(Val(grdEmpeños.CellText(Indice, 6)) = 0 Or Trim(grdEmpeños.CellText(Indice, 6)) = "", 0, grdEmpeños.CellText(Indice, 6))
        crPrestamo = IIf(Val(grdEmpeños.CellText(Indice, 7)) = 0 Or Trim(grdEmpeños.CellText(Indice, 7)) = "", 0, grdEmpeños.CellText(Indice, 7))
        Peso = IIf(Val(grdEmpeños.CellText(Indice, 4)) = 0 Or Trim(grdEmpeños.CellText(Indice, 4)) = "", 0, grdEmpeños.CellText(Indice, 4))
        
        crAvaluoTotal = crAvaluoTotal + (Cantidad * crAvaluo)
        crPrestamoTotal = crPrestamoTotal + (Cantidad * crPrestamo)
        m_Peso = m_Peso + (Peso * Cantidad)
    
    Next Indice
   
    txtPrestamo.Caption = Format(Redondeo(CCur(crPrestamoTotal)), FMoneda)
    lblTotAvaluo.Caption = Format(Redondeo(CCur(crAvaluoTotal)), FMoneda)
    lblTotAvaluo.Tag = crAvaluoTotal
End Sub

Function Calcular_Avaluo_Auto(Empeno As Double, PrestamoAvaluo As Double) As Double
Dim crAvaluo As Double

On Error GoTo error
       
    crAvaluo = (Empeno * (1 + PrestamoAvaluo))

    Calcular_Avaluo_Auto = crAvaluo
  
error:
    Maneja_Error Err
   
End Function

Sub Imprimir_Nota(IDEmpeno As Long, Opcion As Integer)

On Error GoTo error
        
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\Nota.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.ID}=" & IDEmpeno & ""
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "Usuario='" & Trim(UCase(frmMDI.Usuario)) & "'"
        .Formulas(2) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .Formulas(3) = "Opcion=" & Opcion & ""
        
        .SubreportToChange = "NuevosPagos"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .Formulas(0) = "Enajenacion=" & Regresa_Valor_BD("DiasEnajenacion") & ""
        .DiscardSavedData = True
        .WindowState = crptMaximized
        
        .Destination = crptToWindow
        .WindowTitle = "Recibo"
        .Action = 1
    End With
    Exit Sub
    
error:
    Maneja_Error Err
End Sub

Function Default(Opcion As Integer)
    
    If Opcion = 1 Then
        
        txtMunicipio.text = ""
        txtEstado.text = ""
        txtCantidad.text = "1"
    End If
    
End Function

Function VerificaImporte(crPrestamo As Double) As Boolean
Dim crLimite1 As Double, crLimite2 As Double
    
    crLimite1 = CDbl(Regresa_Valor_BD("Limite1"))
    crLimite2 = CDbl(Regresa_Valor_BD("Limite2"))
    VerificaImporte = False
    IDUsuarioAutoriza = 0
    TipoAutorizacion = 0
    
    If crPrestamo < crLimite1 Then
        
        VerificaImporte = True
        
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
        frmPasswords.CancelaCierre = 0
        frmPasswords.AutorizaPrestamo = 1
        
        If frmPasswords.Password(GERENTE, 1) Then
            
            VerificaImporte = True
            TipoAutorizacion = AUTORIZACIONLIMITE1
        End If
                
    ElseIf crPrestamo >= crLimite2 Then
        
        VerificaImporte = True
        TipoAutorizacion = AUTORIZACIONLIMITE2
    End If

End Function

Sub SacaTasa(crPrestamo As Double, TipoInteres As Integer, TipoPeriodo As Integer, TipoPlazo As Integer, ExisteCliente As Boolean)
Dim rcConsulta As New ADODB.Recordset
Dim Meses As Integer

On Error GoTo error
    
    'Tasa
    rcConsulta.Open "SELECT tipointeres.Descripcion AS TipoInteres,tipoperiodo.Descripcion AS TipoPeriodo,tipoperiodo.Periodo,plazos.Descripcion AS Vencimiento " _
                    & "FROM configuraciontasas INNER JOIN plazos ON configuraciontasas.IDPlazo=plazos.ID INNER JOIN tipoperiodo ON configuraciontasas.IDTipoPeriodo=tipoperiodo.ID INNER JOIN tipointeres ON configuraciontasas.IDTipoInteres=tipointeres.ID WHERE " _
                    & "configuraciontasas.IDTipoInteres=" & TipoInteres & " AND configuraciontasas.IDTipoPeriodo=" & TipoPeriodo & " AND configuraciontasas.IDPlazo=" & TipoPlazo, dbDatos, adOpenForwardOnly, adLockReadOnly

    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        Select Case rcConsulta!TipoPeriodo
        Case "MENSUAL"
            
            Meses = 1
        Case "QUINCENAL"
            
            Meses = 2
        Case "SEMANAL"
            
            Meses = 4
            
        End Select
                
        '''''lblAlmacenaje.Caption = Format((Regresa_Valor_BD("Almacenaje") / 30) * rcConsulta!Periodo, "0.00") & "%"
        '''''lblSeguro.Caption = Format((Regresa_Valor_BD("Seguro") / 30) * rcConsulta!Periodo, "0.00") & "%"
        '''''MuestraTasa crPrestamo, ExisteCliente, lblTasa
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub

error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Public Function ValidaArticulos(Pestaña As Integer) As Boolean
Dim rcTmp As ADODB.Recordset

On Error GoTo error
    
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

error:
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
    '''''chkDiamante.Value = 0
    cmbPrenda.ListIndex = -1
    cmbTipo.ListIndex = 0
    cmbKilates.ListIndex = -1
    cmbEstado.ListIndex = -1
    cmbPrenda.ListIndex = -1
    cmbTipoElec.ListIndex = 0
    txtPrestamoo.text = ""
    txtAvaluo.text = ""
End Sub

'Calculamos el prestamo
Private Function Calcula_Prestamo(Prestamo As Double, PorcentajePrestamo As Double, Optional Pestaña As Boolean = True) As Double

On Error GoTo error:
    
    If Pestaña Then
        
        Calcula_Prestamo = Redondeo(Prestamo * (PorcentajePrestamo / 100))
    Else
    
        Calcula_Prestamo = Redondeo((Prestamo / PorcentajePrestamo) * 100)
    End If
    Exit Function
    
error:
    Maneja_Error Err
   
End Function

Sub Recalcula()
Dim i As Integer, Avaluo As Double, Porcentaje As Double

    Porcentaje = Regresa_Valor_BD("PrestamoAvaluo") / 100

    With grdEmpeños

        For i = 1 To .Rows

            If Val(.CellText(i, 6)) > 0 And Trim(.CellText(i, 6)) <> "" And Val(.CellItemData(i, 1)) = 1 Then
                
                Avaluo = .CellText(i, 6)
                .CellText(i, 7) = Redondeo(Avaluo * Porcentaje)
                .CellTextAlign(i, 7) = DT_RIGHT
            End If

        Next i
    
        Total_Avaluos
    End With

End Sub

Function MuestraTasa(crPrestamo As Double, ExisteCliente As Boolean, Etiqueta As Label)
Dim rcTasas As New ADODB.Recordset
Dim TasaTipica As Double, TasaPromocion As Double, TasaPreferencial As Double, LimiteInferior As Double, LimiteSuperior As Double
    
    TasaTipica = 0
    TasaPromocion = 0
    TasaPreferencial = 0
    LimiteInferior = 0
    LimiteSuperior = 0
    
    LimiteInferior = Regresa_Valor_BD("LimiteInferior")
    LimiteSuperior = Regresa_Valor_BD("LimiteSuperior")
    
    rcTasas.Open "SELECT TasaTipica,TasaPromocion,TasaPreferencial FROM configuraciontasas WHERE IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex) & " AND IDPlazo=" & cmbPlazos.ItemData(cmbPlazos.ListIndex), dbDatos, adOpenForwardOnly, adLockOptimistic
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
            
            MuestraTasa = TasaPromocion
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

Sub TasaFija(crPrestamo As Double, Tasa As Double, Almacenaje As Double, Seguro As Double, plazo As Integer, Periodo As Integer, Fecha As Date)
Dim SaldoInsoluto As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double
Dim Vencimiento As Date, i As Integer, crSaldo As Double, crImporteTotal As Double, crPagoFijo As Double, crAmortizacion As Double, strIntervalo As String
    
    crPrestamo = crPrestamo
    SaldoInsoluto = crPrestamo
    crImporteTotal = Redondeo(Pmt((Tasa + Almacenaje + Seguro), plazo, -crPrestamo, 0, 0), 1) * plazo
    crPagoFijo = Redondeo(Pmt((Tasa + Almacenaje + Seguro), plazo, -crPrestamo, 0, 0), 2)
    crSaldo = crImporteTotal
    strIntervalo = "D"
    Vencimiento = DateAdd("D", IIf(strIntervalo = "D", -1, 0), Fecha)
    
    If Periodo = 30 Then
        Periodo = 1
        strIntervalo = "M"
    End If
    
    With grdPagos
    
        .Redraw = False
        .Clear
        
        For i = 1 To plazo
            
            Vencimiento = DateAdd(strIntervalo, Periodo, Vencimiento)
            crImporteTotal = crImporteTotal
            SaldoInsoluto = IIf(i = 1, crPrestamo, SaldoInsoluto - crAmortizacion)
            
            crIntereses = Redondeo(SaldoInsoluto * Tasa)
            crAlmacenaje = Redondeo(SaldoInsoluto * Almacenaje)
            crSeguro = Redondeo(SaldoInsoluto * Seguro)
            
            crAmortizacion = crPagoFijo - (crIntereses + crAlmacenaje + crSeguro)
            crSaldo = Redondeo(crSaldo - (crIntereses + crAlmacenaje + crSeguro + crAmortizacion))
                                                  
            .AddRow
            .CellIcon(.Rows, 1) = lstIcons.ItemIndex(1)
            .CellText(.Rows, 1) = Vencimiento
            .CellTextAlign(.Rows, 1) = DT_CENTER
            .CellText(.Rows, 2) = crIntereses
            .CellTextAlign(.Rows, 2) = DT_RIGHT
            .CellText(.Rows, 3) = crAlmacenaje
            .CellTextAlign(.Rows, 3) = DT_RIGHT
            .CellText(.Rows, 4) = crSeguro
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            .CellText(.Rows, 5) = crPagoFijo
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            .CellText(.Rows, 6) = crAmortizacion
            .CellTextAlign(.Rows, 6) = DT_RIGHT
            .CellText(.Rows, 7) = crSaldo
            .CellTextAlign(.Rows, 7) = DT_RIGHT
        Next i
        
        'Sombreo el Grid
        SombreaGrid grdPagos, 226, 220, 197, 238, 234, 221
                    
        grdPagos.Redraw = True
    
    End With
    
End Sub

Function PrestamoVigente() As Double
Dim crPrestamo As Double, i As Integer
    
    With grdPagos
                    
        For i = 1 To .Rows
            
            If .CellIcon(i, 1) = 0 Then
                
                crPrestamo = crPrestamo + CDbl(.CellText(i, 6))
            End If
            
        Next i
    
    End With
    
    PrestamoVigente = crPrestamo
End Function

Function RegresaVencimiento() As Date
Dim i As Integer

    With grdPagos
                    
        For i = 1 To .Rows
            
            If .CellIcon(i, 1) = 0 Then
                
                Exit For
            End If
            
        Next i
        
        RegresaVencimiento = DateAdd("M", 2, CDate(.CellText(IIf(i = 1, 1, (i - 1)), 1)))
    End With
        
End Function

Sub GrabaPagos(IDEmpeno As Long)
Dim i As Integer
Dim FechaMovimiento As String

    With grdPagos
                    
        For i = 1 To .Rows
            
            If .CellIcon(i, 1) > 0 Then
                
                FechaMovimiento = "'" & Format(Date, "YYYY/MM/DD") & "'"
            Else
                
                FechaMovimiento = "NULL"
            End If
            
            dbDatos.Execute "INSERT INTO pagosfijos (IDEmpeno,NumPago,Vencimiento,Pago,Interes,Almacenaje,Seguro,Amortizacion,Saldo,Pagado,FechaMovimiento) VALUES (" & _
                            IDEmpeno & "," & i & ",'" & Format(CDate(.CellText(i, 1)), "YYYY/MM/DD") & "'," & CDbl(.CellText(i, 5)) & "," & CDbl(.CellText(i, 2)) & "," & CDbl(.CellText(i, 3)) & "," & CDbl(.CellText(i, 4)) & "," & CDbl(.CellText(i, 6)) & "," & CDbl(.CellText(i, 7)) & "," & IIf(.CellIcon(i, 1) > 0, 1, 0) & "," & FechaMovimiento & ")"
        
        Next i
    
    End With
    
End Sub

Private Sub txtVencimiento_GotFocus()
    Seleccionar_Texto txtVencimiento
    Cambiar_Color True, txtVencimiento
End Sub

Private Sub txtVencimiento_LostFocus()
    Cambiar_Color False, txtVencimiento
End Sub
