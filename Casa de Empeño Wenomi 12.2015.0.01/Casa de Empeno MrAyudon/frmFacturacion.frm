VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmFacturacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12150
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFacturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   12150
   Begin VB.Frame frmEmpeño 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7395
      Left            =   45
      TabIndex        =   28
      Top             =   0
      Width           =   12015
      Begin VB.TextBox txtFolioVtaMay 
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
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   55
         Top             =   600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtFolioBillete 
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
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   54
         Top             =   600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.TextBox txtFolioVenta 
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
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   3
         Top             =   600
         Visible         =   0   'False
         Width           =   1140
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
         Left            =   4920
         MaxLength       =   8
         TabIndex        =   2
         Top             =   600
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.ComboBox cmbTipoFactura 
         Height          =   315
         ItemData        =   "frmFacturacion.frx":000C
         Left            =   120
         List            =   "frmFacturacion.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   600
         Width           =   3255
      End
      Begin VB.Frame FrameFacturacion 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   120
         TabIndex        =   37
         Top             =   840
         Width           =   11775
         Begin MSComCtl2.DTPicker dtpFecha 
            Height          =   495
            Left            =   8040
            TabIndex        =   57
            Top             =   2160
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   873
            _Version        =   393216
            Format          =   178716673
            CurrentDate     =   41891
         End
         Begin VB.TextBox txtMunicipioFac 
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
            Left            =   4860
            MaxLength       =   30
            TabIndex        =   16
            Top             =   1680
            Width           =   2625
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
            Height          =   210
            Left            =   6360
            MaxLength       =   120
            TabIndex        =   7
            Top             =   240
            Width           =   3345
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
            Height          =   210
            Left            =   1560
            MaxLength       =   120
            TabIndex        =   6
            Top             =   240
            Width           =   3345
         End
         Begin VB.TextBox txtRazonSocialFac 
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
            Left            =   1560
            MaxLength       =   120
            TabIndex        =   8
            Top             =   600
            Width           =   6465
         End
         Begin VB.TextBox txtRFCFac 
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
            Left            =   9000
            MaxLength       =   20
            TabIndex        =   9
            Top             =   600
            Width           =   2625
         End
         Begin VB.TextBox txtCalleFac 
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
            Left            =   900
            MaxLength       =   120
            TabIndex        =   10
            Top             =   975
            Width           =   7095
         End
         Begin VB.TextBox txtNumExtFac 
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
            Left            =   9000
            MaxLength       =   5
            TabIndex        =   11
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox txtNumIntFac 
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
            Left            =   10800
            MaxLength       =   5
            TabIndex        =   12
            Top             =   960
            Width           =   795
         End
         Begin VB.TextBox txtColoniaFac 
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
            Left            =   915
            MaxLength       =   80
            TabIndex        =   13
            Top             =   1320
            Width           =   7095
         End
         Begin VB.TextBox txtCiudadFac 
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
            Left            =   900
            MaxLength       =   30
            TabIndex        =   15
            Top             =   1680
            Width           =   2625
         End
         Begin VB.TextBox txtEstadoFac 
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
            Left            =   900
            MaxLength       =   30
            TabIndex        =   17
            Top             =   2040
            Width           =   2625
         End
         Begin VB.TextBox txtPaisFac 
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
            Left            =   4860
            MaxLength       =   30
            TabIndex        =   18
            Top             =   2040
            Width           =   2625
         End
         Begin VB.TextBox txtCPFac 
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
            Left            =   9000
            MaxLength       =   5
            TabIndex        =   14
            Top             =   1320
            Width           =   795
         End
         Begin VB.TextBox txtEmailFac 
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
            Left            =   2115
            MaxLength       =   60
            TabIndex        =   19
            Top             =   2400
            Width           =   5415
         End
         Begin VB.Label Label17 
            Caption         =   "FECHA A FACTURAR"
            Height          =   255
            Left            =   8160
            TabIndex        =   58
            Top             =   1800
            Width           =   1815
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   3840
            TabIndex        =   51
            Top             =   1680
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Apellidos(s):"
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
            Left            =   5040
            TabIndex        =   50
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nombre(s):"
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
            TabIndex        =   49
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Razón Social:"
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
            TabIndex        =   48
            Top             =   600
            Width           =   810
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "RFC:"
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
            Left            =   8280
            TabIndex        =   47
            Top             =   600
            Width           =   420
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Calle:"
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
            TabIndex        =   46
            Top             =   960
            Width           =   525
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Ext:"
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
            Left            =   8280
            TabIndex        =   45
            Top             =   960
            Width           =   660
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "N° Int:"
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
            Left            =   10080
            TabIndex        =   44
            Top             =   960
            Width           =   645
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Colonia:"
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
            TabIndex        =   43
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ciudad:"
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
            TabIndex        =   42
            Top             =   1680
            Width           =   720
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            TabIndex        =   41
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Pais:"
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
            Left            =   4320
            TabIndex        =   40
            Top             =   2040
            Width           =   465
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            Left            =   8280
            TabIndex        =   39
            Top             =   1320
            Width           =   315
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Correo Electrónico:"
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
            TabIndex        =   38
            Top             =   2400
            Width           =   1860
         End
      End
      Begin VB.TextBox txtBuscar 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   4
         Top             =   600
         Width           =   3705
      End
      Begin VB.TextBox txtFolio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   10275
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1560
      End
      Begin VB.TextBox lblTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   6945
         Width           =   1725
      End
      Begin VB.TextBox txtSubtotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   6315
         Width           =   1725
      End
      Begin VB.TextBox txtIva 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   6630
         Width           =   1725
      End
      Begin VB.TextBox txtNotas 
         BorderStyle     =   0  'None
         Height          =   555
         Left            =   4155
         MaxLength       =   250
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   6375
         Width           =   4710
      End
      Begin vbAcceleratorGrid6.vbalGrid grdFactura 
         Height          =   2505
         Left            =   120
         TabIndex        =   20
         Top             =   3720
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   4419
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
         ScrollBarStyle  =   1
         Editable        =   -1  'True
         DisableIcons    =   -1  'True
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
         Height          =   225
         Left            =   8040
         TabIndex        =   5
         Top             =   600
         Width           =   330
         _ExtentX        =   582
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
      Begin VB.Label lblTipo2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
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
         TabIndex        =   56
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Serie:"
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
         Left            =   9600
         TabIndex        =   53
         Top             =   600
         Width           =   750
      End
      Begin VB.Label lblSerieFacturacion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "X"
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
         Left            =   10440
         TabIndex        =   52
         Top             =   600
         Width           =   375
      End
      Begin VB.Label lblTipo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CLIENTE:"
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
         TabIndex        =   36
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folio:"
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
         Left            =   9600
         TabIndex        =   35
         Top             =   240
         Width           =   630
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   180
         TabIndex        =   34
         Top             =   240
         Width           =   750
      End
      Begin VB.Label lblFecha 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1035
         TabIndex        =   33
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal:"
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
         Left            =   9000
         TabIndex        =   32
         Top             =   6315
         Width           =   1020
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "IVA:"
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
         Left            =   9540
         TabIndex        =   31
         Top             =   6630
         Width           =   480
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   9375
         TabIndex        =   30
         Top             =   6945
         Width           =   645
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         Caption         =   "Notas:"
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
         Left            =   3525
         TabIndex        =   29
         Top             =   6360
         Width           =   585
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10950
      TabIndex        =   26
      Top             =   7470
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
      Picture         =   "frmFacturacion.frx":0010
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   9840
      TabIndex        =   25
      Top             =   7440
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
      Picture         =   "frmFacturacion.frx":0562
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   8415
      TabIndex        =   27
      Top             =   7470
      Visible         =   0   'False
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
      Picture         =   "frmFacturacion.frx":0AB4
   End
End
Attribute VB_Name = "frmFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim FechaIni As String
Dim FechaFin As String

Private Sub cmbTipoFactura_Click()
    Select Case cmbTipoFactura.ListIndex
         
        'Factura DEL DIA
        Case 0
            lblTipo.Caption = "CLIENTE:": Limpiar
            txtBuscar.Visible = True: cmdMosCliente.Visible = True
            txtFolioRefrendo.Visible = False
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = True: txtApellidos.Enabled = True
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
            BuscarFacturaDia CDate(dtpFecha.Value)
            
        'REFRENDO
        Case 1
            lblTipo.Caption = "FOLIO RECIBO:": Limpiar
            txtBuscar.Visible = False: cmdMosCliente.Visible = False
            txtFolioRefrendo.Visible = True
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = False: txtApellidos.Enabled = False
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
            
        'DESEMPEÑO
        Case 2
            lblTipo.Caption = "FOLIO RECIBO:": Limpiar
            txtBuscar.Visible = False: cmdMosCliente.Visible = False
            txtFolioRefrendo.Visible = True
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = False: txtApellidos.Enabled = False
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
            
        'VENTA MOSTRADOR
        Case 3
            lblTipo.Caption = "FOLIO VENTA:": Limpiar
            txtBuscar.Visible = False: cmdMosCliente.Visible = False
            txtFolioRefrendo.Visible = False
            txtFolioVenta.Visible = True
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = False: txtApellidos.Enabled = False
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
            
        'VENTA BILLETE O VENTA A CLIENTE
        Case 4
            lblTipo.Caption = "CONTRATO:": Limpiar
            txtBuscar.Visible = False: cmdMosCliente.Visible = False
            txtFolioRefrendo.Visible = False
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = True
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = False: txtApellidos.Enabled = False
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
        
        'VENTA MAYORISTA
        Case 5
            lblTipo.Caption = "FOLIO VENTA:": Limpiar
            txtBuscar.Visible = False: cmdMosCliente.Visible = False
            txtFolioRefrendo.Visible = False
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = True: lblTipo2.Caption = ""
            txtNombre.Enabled = True: txtApellidos.Enabled = True
            'txtBuscar.Left = 5355: cmdMosCliente.Left = 9105
         
        'REFRENDOS DIA
        Case 6
            lblTipo.Caption = "CLIENTE:": Limpiar
            txtBuscar.Visible = True: cmdMosCliente.Visible = True
            txtFolioRefrendo.Visible = False
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = True: txtApellidos.Enabled = True
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
            Buscar_Refrendos_Desempenos_Dia Date, 2, "REFRENDOS" 'Date
        
        'DESEMPEÑOS DIA
        Case 7
            lblTipo.Caption = "CLIENTE:": Limpiar
            txtBuscar.Visible = True: cmdMosCliente.Visible = True
            txtFolioRefrendo.Visible = False
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = True: txtApellidos.Enabled = True
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
            Buscar_Refrendos_Desempenos_Dia Date, 3, "DESEMPEÑOS" 'Date
            
        'VENTAS DIA
        Case 8
            lblTipo.Caption = "CLIENTE:": Limpiar
            txtBuscar.Visible = True: cmdMosCliente.Visible = True
            txtFolioRefrendo.Visible = False
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = True: txtApellidos.Enabled = True
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
            Buscar_Ventas_Dia Date
            
            Case 9
            lblTipo.Caption = "CLIENTE:": Limpiar
            txtBuscar.Visible = True: cmdMosCliente.Visible = True
            txtFolioRefrendo.Visible = False
            txtFolioVenta.Visible = False
            txtFolioBillete.Visible = False
            txtFolioVtaMay.Visible = False: lblTipo2.Caption = ""
            txtNombre.Enabled = True: txtApellidos.Enabled = True
            'txtBuscar.Left = 2955: cmdMosCliente.Left = 6705
            
            frmRangoFechas.Caption = "Reporte de contratos vencidos"
            frmRangoFechas.Fechas FechaIni, FechaFin
            If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
            
            BuscarFacturaPlazo FechaIni, FechaFin
            
    End Select
End Sub

Private Sub cmdAceptar_Click()

    Dim IDCliente As Integer, Subtotal As Double, Iva As Double, Total As Double
    Dim i As Integer, IDFactura As Integer, Cantidad As Integer, Importe As Double, PrecioUnit As Double
    Dim rcConsulta As New ADODB.Recordset
    Dim FolioDoc As Long, FolioComp As Long
    Dim sNombreArchivo As String

On Error GoTo Error

    If MsgBox("Estan correctos los datos ??", vbInformation + vbYesNo + vbDefaultButton1, "Facturación") = vbYes Then
        
        If Valida Then
        
            rcConsulta.Open "SELECT ID FROM facturas WHERE Folio='" & Trim(txtFolio.text) & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            
            If Not rcConsulta.BOF And Not rcConsulta.EOF Then
                MsgBox "El folio de la factura que desea registrar ya existe!!", vbCritical, "Facturación"
            Else
                
                Select Case cmbTipoFactura.ListIndex
                    Case 0: FolioDoc = 0: FolioComp = 0
                    Case 1: FolioDoc = txtFolioRefrendo.Tag: FolioComp = txtFolioRefrendo.text
                    Case 2: FolioDoc = txtFolioRefrendo.Tag: FolioComp = txtFolioRefrendo.text
                    Case 3: FolioDoc = txtFolioVenta.Tag: FolioComp = txtFolioVenta.text
                    Case 4: FolioDoc = txtFolioBillete.Tag: FolioComp = txtFolioBillete.text
                    Case 5: FolioDoc = txtFolioVtaMay.Tag: FolioComp = txtFolioVtaMay.text
                End Select
            
                If txtNombre.Tag = "" Then
                    IDCliente = Grabar_Cliente
                    txtNombre.Tag = IDCliente
                Else
                    Actualizar_Cliente (txtNombre.Tag)
                    IDCliente = txtNombre.Tag
                End If
                
                If Val(txtSubTotal.text) = 0 Or Trim(txtSubTotal.text) = "" Then
                    Subtotal = 0
                Else
                    Subtotal = txtSubTotal.text
                End If
            
                If Val(txtIva.text) = 0 Or Trim(txtIva.text) = "" Then
                    Iva = 0
                Else
                    Iva = txtIva.text
                End If
            
                Total = CDbl(lblTotal.text)
                
                'Grabo los datos en la tabla de facturas
                dbDatos.Execute "INSERT INTO facturas (Folio,Fecha,Cliente,Subtotal,Iva,Total,Notas,IDUsuario,IDSucursal,TipoFactura,IDDocOrigen) " & _
                                "VALUES ('" & Trim(txtFolio.text) & "','" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & IDCliente & "," & Subtotal & "," & Iva & "," & Total & ",'" & Trim(txtNotas.text) & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & cmbTipoFactura.ListIndex & "," & FolioDoc & ")"
                
                'Incrementar el Folio de control de las facturas
                Regresa_Movimiento True, "FolioFacturas"
                
                'Tomo el ID de la factura
                IDFactura = SacaValor("facturas", "MAX(ID)")
                
                For i = 1 To grdFactura.Rows
                    If Trim(grdFactura.CellText(i, 1)) <> "" Then
                        Cantidad = grdFactura.CellText(i, 1)
                        PrecioUnit = CDbl(grdFactura.CellText(i, 5))
                        Importe = CDbl(grdFactura.CellText(i, 6))
                        dbDatos.Execute "INSERT INTO detallefactura (IDFactura,Cantidad,Unidad,Codigo,Concepto,PrecioUnit,Importe) " & _
                                        "VALUES (" & IDFactura & "," & Cantidad & ",'" & grdFactura.CellText(i, 2) & "','" & grdFactura.CellText(i, 3) & "','" & grdFactura.CellText(i, 4) & "'," & PrecioUnit & "," & Importe & ")"
                                        
                    End If
                Next i
                
                '*** Marcar los registros de refrendos para factura del dia ***
                MarcarDocumentosFacturados FolioComp, cmbTipoFactura.ListIndex
                
                '*** Genera el Archivo de Texto de la factura ***
               ' sNombreArchivo = GeneraTxtFactura(txtFolio.text, CDate(lblFecha.Caption))
                CrarCfdiRV
                'MsgBox "La Factura N° " & Format(txtFolio.text, "000000") & " se Generó correctamente." & vbCrLf & _
                '       "Nombre de archivo: " & sNombreArchivo
                
                'Imprimir CLng(txtFolio.text)
                'grdFactura.CancelEdit
                Limpiar
            
            End If
            rcConsulta.Close
            Set rcConsulta = Nothing
        End If
    
    End If
        
    'If sNombreArchivo <> "" Then Crear_CFDi sNombreArchivo

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing

End Sub



Private Sub cmdImprimir_Click()
    CrearArchivoFactura "FAC_" & Format(txtFolio.text, "000000") & "_" & Trim(txtRFCFac.text) & "_" & CStr(Format(Day(Now), "00")) & CStr(Format(Month(Now), "00")) & CStr(Format(Year(Now), "00")) & ".txt"
End Sub

Private Sub cmdMosCliente_Click()
    frmMostrarCliente.Ver Me, txtBuscar, True, 0
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    frmEmpeño.BorderStyle = 0
    lblFecha.Caption = Format(Date, "dd/mmm/yyyy")
    'lblSerieFacturacion.Caption = MisParametros.SerieFacturas
    lblSerieFacturacion.Caption = Regresa_Valor_BD("SerieFacturas")
    CentrarForm Me, frmMDI
    Encabezado
    dtpFecha.Value = Date
'    Poner_Flat fl, Me.Controls, Me
    
    cmbTipoFactura.AddItem "FACTURA DEL DIA"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 1
    cmbTipoFactura.AddItem "REFRENDO"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 2
    cmbTipoFactura.AddItem "DESEMPEÑO"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 3
    cmbTipoFactura.AddItem "VENTA MOSTRADOR"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 4
    cmbTipoFactura.AddItem "VENTA CLIENTE"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 5
    cmbTipoFactura.AddItem "VENTA MAYORISTA"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 6
    cmbTipoFactura.AddItem "REFRENDOS DEL DIA"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 7
    cmbTipoFactura.AddItem "DESEMPEÑOS DEL DIA"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 8
    cmbTipoFactura.AddItem "VENTAS DEL DIA"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 8
    cmbTipoFactura.AddItem "FACTURA POR RANGO DE FECHAS"
    cmbTipoFactura.ItemData(cmbTipoFactura.NewIndex) = 9
    
    
    'cmbTipoFactura.ListIndex = 0
    lblTipo.Caption = "CLIENTE:"
    lblTipo2.Caption = ""
            
    'Mostrar el Folio Consecutivo de las facturas
    txtFolio.text = Regresa_Movimiento(False, "FolioFacturas")
            
End Sub

'MOD. 26-FEB-2014
Public Sub Buscar(ID As Long) 'Buscar_Cliente
Dim rcClientes As New ADODB.Recordset

On Error GoTo Error
   
    rcClientes.Open "SELECT * FROM clientes WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
   
    With rcClientes
        txtNombre.text = !Nombre
        txtNombre.Tag = ID
        txtApellidos.text = !Apellido
        
        txtRazonSocialFac.text = IIf(IsNull(!RazonSocial_Fac), "", !RazonSocial_Fac) '!RazonSocial_Fac
        txtRFCFac.text = IIf(IsNull(!RFC_Fac), "", !RFC_Fac) '!RFC_Fac
        txtCalleFac.text = IIf(IsNull(!Calle_Fac), "", !Calle_Fac) '!Calle_Fac
        txtNumExtFac.text = IIf(IsNull(!NumExt_Fac), "", !NumExt_Fac) '!NumExt_Fac
        txtNumIntFac.text = IIf(IsNull(!NumInt_Fac), "", !NumInt_Fac) '!NumInt_Fac
        txtColoniaFac.text = IIf(IsNull(!Colonia_Fac), "", !Colonia_Fac) '!Colonia_Fac
        txtCPFac.text = IIf(IsNull(!CP_Fac), "", !CP_Fac) '!CP_Fac
        txtCiudadFac.text = IIf(IsNull(!Ciudad_Fac), "", !Ciudad_Fac) '!Ciudad_Fac
        txtMunicipioFac.text = IIf(IsNull(!Municipio_Fac), "", !Municipio_Fac) '!Municipio_Fac
        txtEstadoFac.text = IIf(IsNull(!Estado_Fac), "", !Estado_Fac) '!Estado_Fac
        txtPaisFac.text = IIf(IsNull(!Pais_Fac), "", !Pais_Fac) '!Pais_Fac
        txtEmailFac.text = IIf(IsNull(!Email_Fac), "", !Email_Fac) '!Email_Fac
        
    End With

    rcClientes.Close
    Set rcClientes = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcClientes = Nothing
End Sub

Sub Encabezado()

    With grdFactura
        .AddColumn "K1", "Cantidad", ecgHdrTextALignRight, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Unidad", ecgHdrTextALignCentre, , 55, , , , , , , CCLSortString
        .AddColumn "K3", "Codigo", ecgHdrTextALignCentre, , 90, , , , , , , CCLSortString
        .AddColumn "K4", "Concepto", ecgHdrTextALignLeft, , 380, , , , , , , CCLSortString
        .AddColumn "K5", "P. Unitario", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Importe", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "K7", "IDDocOrigen", ecgHdrTextALignRight, , 85, False, , , , , , CCLSortNumeric
        
        .Gridlines = True
        '.Rows = 12
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    'Quitar_Flat fl
End Sub

Private Sub lblTotal_GotFocus()
    Seleccionar_Texto lblTotal
    Cambiar_Color True, lblTotal
End Sub

Private Sub lblTotal_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub lblTotal_LostFocus()
    Cambiar_Color False, lblTotal
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

Private Sub txtBuscar_GotFocus()
    Seleccionar_Texto txtBuscar
    Cambiar_Color True, txtBuscar
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBuscar_LostFocus()
    Cambiar_Color False, txtBuscar
End Sub

'Private Sub txtCantidad_GotFocus()
'    Cambiar_Color True, txtCantidad
'End Sub
'
'Private Sub txtCantidad_LostFocus()
'    Cambiar_Color False, txtCantidad
'End Sub

'Private Sub txtColonia_GotFocus()
'    Seleccionar_Texto txtColonia
'    Cambiar_Color True, txtColonia
'End Sub
'
'Private Sub txtColonia_KeyPress(KeyAscii As Integer)
'    KeyAscii = Mayusculas(KeyAscii)
'    Pasar_Foco KeyAscii
'End Sub
'
'Private Sub txtColonia_LostFocus()
'    Cambiar_Color False, txtColonia
'End Sub

'Private Sub txtCP_GotFocus()
'    Seleccionar_Texto txtCp
'    Cambiar_Color True, txtCp
'End Sub
'
'Private Sub txtCP_KeyPress(KeyAscii As Integer)
'    KeyAscii = Solo_Numeros(KeyAscii)
'    Pasar_Foco KeyAscii
'End Sub
'
'Private Sub txtCP_LostFocus()
'    Cambiar_Color False, txtCp
'End Sub

'Private Sub txtDireccion_GotFocus()
'    Seleccionar_Texto txtDireccion
'    Cambiar_Color True, txtDireccion
'End Sub
'
'Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
'    KeyAscii = Mayusculas(KeyAscii)
'    Pasar_Foco KeyAscii
'End Sub
'
'Private Sub txtDireccion_LostFocus()
'    Cambiar_Color False, txtDireccion
'End Sub

'Private Sub txtEstado_GotFocus()
'    Seleccionar_Texto txtEstado
'    Cambiar_Color True, txtEstado
'End Sub
'
'Private Sub txtEstado_KeyPress(KeyAscii As Integer)
'    KeyAscii = Mayusculas(KeyAscii)
'    Pasar_Foco KeyAscii
'End Sub
'
'Private Sub txtEstado_LostFocus()
'    Cambiar_Color False, txtEstado
'End Sub

Private Sub txtFolio_GotFocus()
    Seleccionar_Texto txtFolio
    Cambiar_Color True, txtFolio
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFolio_LostFocus()
    Cambiar_Color False, txtFolio
End Sub

'Private Sub txtImporte_GotFocus()
'    Cambiar_Color True, txtImporte
'End Sub

Private Sub txtFolioRefrendo_Change()
    txtFolioRefrendo.Tag = ""
End Sub

Private Sub txtFolioRefrendo_GotFocus()
    Seleccionar_Texto txtFolioRefrendo
    Cambiar_Color True, txtFolioRefrendo
End Sub

Private Sub txtFolioRefrendo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        
        If Trim(txtFolioRefrendo.text) <> "" Then
             BuscarContrato txtFolioRefrendo
        Else
            MsgBox "Introduzca el número de Contrato que desea Facturar !!", vbCritical, "Facturación"
        End If
        
    End If
    
End Sub

Private Sub txtFolioRefrendo_LostFocus()
    Cambiar_Color False, txtFolioRefrendo
End Sub

Private Sub txtFolioVenta_Change()
    txtFolioVenta.Tag = ""
End Sub

Private Sub txtFolioVenta_GotFocus()
    Seleccionar_Texto txtFolioVenta
    Cambiar_Color True, txtFolioVenta
End Sub

Private Sub txtFolioVenta_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtFolioVenta.text) <> "" Then
            BuscarVentaCte txtFolioVenta
        Else
            MsgBox "Introduzca el número de Folio Venta que desea Facturar !!", vbCritical, "Facturación"
        End If
    End If
End Sub

Private Sub txtFolioVenta_LostFocus()
    Cambiar_Color False, txtFolioVenta
End Sub

Private Sub txtIva_Change()
Dim Importe As Double, Iva As Double, Ivaa As String, Importee As String
    
    txtSubTotal.text = Format(txtSubTotal.text, FMoneda)
    txtIva.text = Format(txtIva.text, FMoneda)
    
    Importee = Trim(txtSubTotal.text)

    If Trim(Importee) = "." Then Importee = "0" & "."
    Importe = IIf(Trim(Importee) = "", 0, Importee)

    Ivaa = Trim(txtIva.text)

    If Ivaa = "." Then Ivaa = "0" & "."
    Iva = IIf(Trim(Ivaa) <> "", Ivaa, 0)

    lblTotal.text = Format(Importe + Iva, FMoneda)
End Sub

Private Sub txtIva_GotFocus()
    Seleccionar_Texto txtIva
    Cambiar_Color True, txtIva
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtIva_LostFocus()
    Cambiar_Color False, txtIva
End Sub

'Private Sub txtMunicipio_GotFocus()
'    Seleccionar_Texto txtMunicipio
'    Cambiar_Color True, txtMunicipio
'End Sub

'Private Sub txtMunicipio_KeyPress(KeyAscii As Integer)
'    KeyAscii = Mayusculas(KeyAscii)
'    Pasar_Foco KeyAscii
'End Sub
'
'Private Sub txtMunicipio_LostFocus()
'    Cambiar_Color False, txtMunicipio
'End Sub

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

Private Sub txtNotas_GotFocus()
    Seleccionar_Texto txtNotas
    Cambiar_Color True, txtNotas
End Sub

Private Sub txtNotas_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNotas_LostFocus()
    Cambiar_Color False, txtNotas
End Sub

'Private Sub txtRfc_GotFocus()
'    Seleccionar_Texto txtRfc
'    Cambiar_Color True, txtRfc
'End Sub
'
'Private Sub txtRfc_KeyPress(KeyAscii As Integer)
'    KeyAscii = Mayusculas(KeyAscii)
'    Pasar_Foco KeyAscii
'End Sub
'
'Private Sub txtRfc_LostFocus()
'    Cambiar_Color False, txtRfc
'End Sub

Private Sub txtSubtotal_Change()
Dim Importe As Double, Iva As Double, Ivaa As String
Dim Importee As String
    
    txtSubTotal.text = Format(txtSubTotal.text, FMoneda)
    txtIva.text = Format(txtIva.text, FMoneda)
    
    Importee = Trim(txtSubTotal.text)

    If Trim(Importee) = "." Then Importee = "0" & "."
    Importe = IIf(Trim(Importee) <> "", Importee, 0)

    Ivaa = Trim(txtIva.text)

    If Trim(Ivaa) = "." Then Ivaa = "0" & "."
    Iva = IIf(Trim(Ivaa) = "", 0, Ivaa)

    lblTotal.text = Format(Importe + Iva, FMoneda)
End Sub

Private Sub txtSubtotal_GotFocus()
    Seleccionar_Texto txtSubTotal
    Cambiar_Color True, txtSubTotal
End Sub

Private Sub txtSubtotal_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtSubtotal_LostFocus()
    Cambiar_Color False, txtSubTotal
End Sub

'Private Sub txtTelefono_GotFocus()
'    Seleccionar_Texto txtTelefono
'    Cambiar_Color True, txtTelefono
'End Sub
'
'Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
'    KeyAscii = Mayusculas(KeyAscii)
'    Pasar_Foco KeyAscii
'End Sub
'
'Private Sub txtTelefono_LostFocus()
'    Cambiar_Color False, txtTelefono
'End Sub

Sub Totales()
Dim i As Integer, Subtotal As Double, Iva As Double

    For i = 1 To grdFactura.Rows

        If Trim(grdFactura.CellText(i, 4)) <> "" Then Subtotal = Subtotal + CDbl(grdFactura.CellText(i, 4))
    Next i

    'Iva = MisParametros.Iva / 100
    Iva = Val(Regresa_Valor_BD("Iva")) / 100
    txtSubTotal.text = Format(Subtotal, FMoneda)
    txtIva.text = Format(Subtotal * Iva, FMoneda)
End Sub

Private Function Grabar_Cliente() As Long

On Error GoTo Error

    'dbDatos.Execute "INSERT INTO clientes " & _
    '"(Nombre,Apellido,Iniciales,Direccion,Colonia,Municipio,Estado,Tel,CP,rfc) VALUES ('" & _
    '                Trim(txtNombre.text) & "','" & Trim(txtApellidos.text) & "','" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "','" & Trim(txtDireccion.text) & "','" & Trim(txtColonia.text) & "','" & Trim(txtMunicipio.text) & "','" & Trim(txtEstado.text) & "','" & Trim(txtTelefono.text) & "','" & Trim(txtCp.text) & "','" & Trim(txtRfc.text) & "')"
    
    dbDatos.Execute "INSERT INTO clientes " & _
            "(Iniciales,Nombre,Apellido,Direccion,Colonia,Municipio,Estado,CP,rfc," & _
            "RazonSocial_Fac,RFC_Fac,Calle_Fac,NumExt_Fac,NumInt_Fac,Colonia_Fac,CP_Fac,Ciudad_Fac,Estado_Fac,Pais_Fac,Email_Fac,Municipio_Fac)" & _
            " VALUES " & _
            "('" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & " ','" & Trim(txtNombre.text) & "','" & Trim(txtApellidos.text) & "','" & Trim(txtCalleFac.text) & "','" & Trim(txtColoniaFac.text) & "','" & Trim(txtMunicipioFac.text) & "','" & Trim(txtEstadoFac.text) & "','" & Trim(txtCPFac.text) & "','" & Trim(txtRFCFac.text) & "'," & _
            "'" & Trim(txtRazonSocialFac.text) & "','" & Trim(txtRFCFac.text) & "','" & Trim(txtCalleFac.text) & "','" & Trim(txtNumExtFac.text) & "','" & Trim(txtNumIntFac.text) & "','" & Trim(txtColoniaFac.text) & "','" & Trim(txtCPFac.text) & "','" & Trim(txtCiudadFac.text) & "','" & Trim(txtEstadoFac.text) & "','" & Trim(txtPaisFac.text) & "','" & Trim(txtEmailFac.text) & "','" & Trim(txtMunicipioFac.text) & "')"
        
    Grabar_Cliente = SacaValor("clientes", "MAX(ID)")
    Exit Function
    
Error:
    Maneja_Error Err
End Function

Private Sub Actualizar_Cliente(ID As Long)

On Error GoTo Error

    'dbDatos.Execute "UPDATE clientes SET Iniciales='" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "',Direccion='" & Trim(txtDireccion.text) & "',Colonia='" & Trim(txtColonia.text) & "',Municipio='" & Trim(txtMunicipio.text) & "'," & "Estado='" & Trim(txtEstado.text) & "',Tel='" & Trim(txtTelefono.text) & "',CP='" & Trim(txtCp.text) & "',Rfc='" & Trim(txtRfc.text) & "' WHERE ID = " & ID
    
    dbDatos.Execute "UPDATE clientes SET Iniciales='" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "',Nombre='" & Trim(txtNombre.text) & "',Apellido='" & Trim(txtApellidos.text) & "',Direccion='" & txtCalleFac.text & "',Colonia='" & txtColoniaFac.text & "',Municipio='" & txtMunicipioFac.text & "'," & _
                    "Estado='" & txtEstadoFac.text & "',CP='" & txtCPFac.text & "'," & _
                    "RazonSocial_Fac='" & Trim(txtRazonSocialFac.text) & "', RFC_Fac='" & Trim(txtRFCFac.text) & "', Calle_Fac='" & Trim(txtCalleFac.text) & "', NumExt_Fac='" & Trim(txtNumExtFac.text) & "', NumInt_Fac='" & Trim(txtNumIntFac.text) & "', Colonia_Fac='" & Trim(txtColoniaFac.text) & "', CP_Fac='" & Trim(txtCPFac.text) & "', Ciudad_Fac='" & Trim(txtCiudadFac.text) & "', Estado_Fac='" & Trim(txtEstadoFac.text) & "', Pais_Fac='" & Trim(txtPaisFac.text) & "', Email_Fac='" & Trim(txtEmailFac.text) & "', Municipio_Fac='" & Trim(txtMunicipioFac.text) & "' " & _
                    " WHERE ID = " & ID
    
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Sub Limpiar()

    txtFolioRefrendo.text = "": txtFolioVenta.text = ""
    txtFolioRefrendo.Tag = "": txtFolioVenta.Tag = ""
    
    txtFolio.text = ""
    txtFolioBillete = "": txtFolioBillete.Tag = ""
    
    txtFolioVtaMay.text = "": txtFolioVtaMay.Tag = ""
    
    'Mostrar el Folio Consecutivo de las facturas
    txtFolio.text = Regresa_Movimiento(False, "FolioFacturas")
    
    txtNombre.Tag = ""
    txtNombre.text = ""
    txtApellidos.text = ""
    
    txtRazonSocialFac.text = ""
    txtRFCFac.text = ""
    txtCalleFac.text = ""
    txtNumExtFac.text = ""
    txtNumIntFac.text = ""
    txtCPFac.text = ""
    txtColoniaFac.text = ""
    txtCiudadFac.text = ""
    txtMunicipioFac.text = ""
    txtEstadoFac.text = ""
    txtPaisFac.text = ""
    txtEmailFac.text = ""
    
    txtSubTotal.text = ""
    txtIva.text = ""
    txtNotas.text = ""
    txtBuscar.text = ""
    
    grdFactura.Clear True
    Encabezado
End Sub

Function Valida() As Boolean
    
    Valida = True

    If txtFolio.text = "" Then
        MsgBox "Introduzca el folio de la factura !!", vbInformation, "Facturación": Valida = False: txtFolio.SetFocus
        Exit Function
    End If

    If txtNombre.text = "" Then
        MsgBox "Introduzca el nombre del cliente !!", vbInformation, "Facturación": Valida = False: txtNombre.SetFocus
        Exit Function
    End If
    
    If txtApellidos.text = "" Then
        MsgBox "Introduzca el Apellido del cliente !!", vbInformation, "Facturación": Valida = False: txtApellidos.SetFocus
        Exit Function
    End If
    
    If txtRazonSocialFac.text = "" Then
        MsgBox "Introduzca la Razón Social del cliente !!", vbInformation, "Facturación": Valida = False: txtRazonSocialFac.SetFocus
        Exit Function
    End If

    If txtRFCFac.text = "" Then
        MsgBox "Introduzca el Rfc !!", vbInformation, "Facturación": Valida = False: txtRFCFac.SetFocus
        Exit Function
    End If

    If txtCalleFac.text = "" Then
        MsgBox "Introduzca la Calle de Domicilio Fiscal del cliente !!", vbInformation, "Facturación": Valida = False: txtCalleFac.SetFocus
        Exit Function
    End If
    
    If txtNumExtFac.text = "" Then
        MsgBox "Introduzca N° Exterior de Domicilio Fiscal del cliente !!", vbInformation, "Facturación": Valida = False: txtNumExtFac.SetFocus
        Exit Function
    End If
    
    If txtColoniaFac.text = "" Then
        MsgBox "Introduzca la Colonia de Domicilio Fiscal del cliente !!", vbInformation, "Facturación": Valida = False: txtColoniaFac.SetFocus
        Exit Function
    End If
    
    If txtCPFac.text = "" Then
        MsgBox "Introduzca la Código Postal de Domicilio Fiscal del cliente !!", vbInformation, "Facturación": Valida = False: txtCPFac.SetFocus
        Exit Function
    End If
    
    If txtCiudadFac.text = "" Then
        MsgBox "Introduzca la Ciudad de Domicilio Fiscal del cliente!!", vbInformation, "Facturación": Valida = False: txtCiudadFac.SetFocus
        Exit Function
    End If

    If txtMunicipioFac.text = "" Then
        MsgBox "Introduzca Municipio de Domicilio Fiscal del cliente !!", vbInformation, "Facturación": Valida = False: txtMunicipioFac.SetFocus
        Exit Function
    End If

    If txtEstadoFac.text = "" Then
        MsgBox "Introduzca Estado de Domicilio Fiscal del cliente!!", vbInformation, "Facturación": Valida = False: txtEstadoFac.SetFocus
        Exit Function
    End If

    If txtPaisFac.text = "" Then
        MsgBox "Introduzca País de Domicilio Fiscal del cliente!!", vbInformation, "Facturación": Valida = False: txtPaisFac.SetFocus
        Exit Function
    End If

    If ValidaGrid = False Then Valida = False
    
    If ValidarEmisor(frmMDI.IDSucursal) = False Then Valida = False
    
End Function

Function ValidaGrid() As Boolean
    Dim i As Integer

    ValidaGrid = True

    With grdFactura
        If .Rows = 0 Then
            MsgBox "Introduzca el Detalle de la factura !!", vbInformation, "Facturación"
            ValidaGrid = False
        End If
    End With

End Function



'07-NOV-2011 ::: Buscar el Refrendo a Facturar
Private Sub BuscarContrato(ByVal Folio As Long)
    Dim RsEmp As New ADODB.Recordset
    Dim Sql As String
    Dim ImpIntereses As Double, ImpAlmacenaje As Double, ImpSeguro As Double, ImpMoratorios As Double, ImpPerdida As Double, ImpOtros, ImpIVA As Double
    Dim sTitulo As String
    Dim Subtotal As Double
    Dim Serie As Integer, IDEmpeno As Long
    
    Subtotal = 0
    
    If cmbTipoFactura.ListIndex = 1 Then
        sTitulo = "REFRENDO"
        Sql = "SELECT ID,Folio,Facturado,IDSucursal,IDCliente,Intereses as ImpIntereses, ImporteAlmacenaje as ImpAlmacenaje, ImporteSeguro as ImpSeguro, ImporteMoratorios as ImpMoratorios, ImportePerdida as ImpPerdida, ImporteOtros as ImpOtros, ImporteIVA as ImpIVA FROM empeno " & _
              "WHERE FolioNota = " & Folio & " AND " & _
              "IDSucursal =" & frmMDI.IDSucursal & " AND cancelado=0 " 'AND facturado=0 " 'AND (DESTINO=2) "
          
    ElseIf cmbTipoFactura.ListIndex = 2 Then
        sTitulo = "DESEMPEÑO"
        Sql = "SELECT ID,Folio,Facturado,IDSucursal,IDCliente,Intereses as ImpIntereses, ImporteAlmacenaje as ImpAlmacenaje, ImporteSeguro as ImpSeguro, ImporteMoratorios as ImpMoratorios, ImportePerdida as ImpPerdida, ImporteOtros as ImpOtros, ImporteIVA as ImpIVA FROM empeno " & _
              "WHERE FolioNota = " & Folio & " AND " & _
              "IDSucursal =" & frmMDI.IDSucursal & " AND cancelado=0 " 'AND facturado=0 " 'AND (Destino=3) "
                
    End If
          
    RsEmp.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    ImpIntereses = 0: ImpAlmacenaje = 0: ImpSeguro = 0: ImpMoratorios = 0: ImpPerdida = 0: ImpOtros = 0
    IDEmpeno = 0: Subtotal = 0
    If RsEmp.EOF Then
    
        MsgBox "No se encontró el contrato especificado !!", vbInformation, "Facturación"
        txtFolioRefrendo.Tag = ""
        Limpiar
        
    Else
        
        If RsEmp!Facturado = 1 Then
            
            If MsgBox("La Nota de " & sTitulo & " ya ha sido facturada. ¿Desea Volver a Facturar?", vbYesNo + vbExclamation, "Facturación") = vbNo Then
                txtFolioRefrendo.Tag = ""
                Limpiar
                If RsEmp.State = 1 Then
                    RsEmp.Close
                    Set RsEmp = Nothing
                End If
                Exit Sub
            
            End If
            
        ElseIf RsEmp!Facturado = 2 Then
        
            MsgBox "La Nota de " & sTitulo & " ya se facturó en Factura del Día. No se puede refacturar!!", vbInformation, "Facturación"
            txtFolioRefrendo.Tag = ""
            Limpiar
            If RsEmp.State = 1 Then
                RsEmp.Close
                Set RsEmp = Nothing
            End If
            Exit Sub
                
        End If
        
        Buscar RsEmp!IDCliente
        
        txtFolioRefrendo.Tag = RsEmp!Folio
        ImpIntereses = RsEmp!ImpIntereses
        ImpAlmacenaje = RsEmp!ImpAlmacenaje
        ImpSeguro = RsEmp!ImpSeguro
        ImpMoratorios = RsEmp!ImpMoratorios
        ImpIVA = RsEmp!ImpIVA
        ImpPerdida = RsEmp!ImpPerdida
        ImpOtros = RsEmp!ImpOtros
        IDEmpeno = RsEmp!Folio
        'Do While Not RsEmp.EOF
        'Loop
    End If
    RsEmp.Close
    Set RsEmp = Nothing
    
    With grdFactura
        .Clear
        If ImpIntereses <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "INT": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "INTERESES " & sTitulo & " NOTA " & Folio
            .CellText(.Rows, 5) = Format(ImpIntereses, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpIntereses, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 7) = IDEmpeno
        End If
        If ImpAlmacenaje <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "ALM": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "ALMACENAJE " & sTitulo & " NOTA " & Folio
            .CellText(.Rows, 5) = Format(ImpAlmacenaje, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpAlmacenaje, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 7) = IDEmpeno
        End If
        If ImpSeguro <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "SEG": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "SEGURO " & sTitulo & " NOTA " & Folio
            .CellText(.Rows, 5) = Format(ImpSeguro, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpSeguro, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 7) = IDEmpeno
        End If
        If ImpMoratorios <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "MOR": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "MORATORIOS / RECARGOS " & sTitulo & " NOTA " & Folio
            .CellText(.Rows, 5) = Format(ImpMoratorios, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpMoratorios, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 7) = IDEmpeno
        End If
    End With
    
    Subtotal = ImpIntereses + ImpAlmacenaje + ImpSeguro + ImpMoratorios
    
    txtSubTotal.text = Subtotal
''''txtIva.text = SubTotal * (Regresa_Valor_BD("IvaVentas") / 100)
    txtIva.text = ImpIVA
    
End Sub

'07-NOV-2011 ::: Buscar la Venta a Cliente a Facturar
Private Sub BuscarVentaCte(ByVal Folio As Long)
    
    '*** OBTENER LA VENTA MOSTRADOR (DESGLOCE POR ARTICULO) ***
    Dim RsVta As New ADODB.Recordset
    Dim Sql As String
    Dim Subtotal As Double
    Dim ImporteIva As Double
    Dim IvaFacturacionVentas As Double
    
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.ImporteIva, v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE v.Folio=" & Folio & " AND " & "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND v.TipoVenta=" & VENTAMOSTRADOR
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If RsVta.EOF Then
        MsgBox "No se encontró el Folio de Venta especificado !!", vbInformation, "Facturación"
        txtFolioVenta.Tag = ""
        Subtotal = 0: ImporteIva = 0
        Limpiar
    Else
        Limpiar
        If RsVta!Apartado = 1 And RsVta!Pagado = 0 Then
            MsgBox "El Folio de Venta no ha sido Pagado en su totalidad. !!", vbInformation, "Facturación"
            txtFolioVenta.Tag = ""
            Subtotal = 0: ImporteIva = 0
            Limpiar
            GoTo Salir
        End If
    
        If RsVta!Facturado = 1 Then
            If MsgBox("El Folio de Venta ya ha sido facturado. ¿Desea Volver a Facturar?", vbYesNo, "Facturación") = vbNo Then
                txtFolioVenta.Tag = ""
                Subtotal = 0: ImporteIva = 0
                Limpiar
                GoTo Salir
            End If
        ElseIf RsVta!Facturado = 2 Then
            MsgBox "El Folio de Venta ya se facturó en Factura del Día. No se puede refacturar. !!", vbInformation, "Facturación"
            txtFolioVenta.Tag = ""
            Subtotal = 0: ImporteIva = 0
            Limpiar
            GoTo Salir
        End If
        
        Buscar RsVta!IDCliente
        txtFolioVenta.text = RsVta!Folio
        txtFolioVenta.Tag = RsVta!Folio
        
        IvaFacturacionVentas = IIf(Regresa_Valor_BD("IvaVentas") > 0, Regresa_Valor_BD("IvaVentas"), Regresa_Valor_BD("IvaFacturacionVentas")) / 100
        
        Subtotal = 0: ImporteIva = 0
        Do While Not RsVta.EOF
            
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
                .CellText(.Rows, 5) = Format(((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format(((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
            End With
            
            Subtotal = Subtotal + (((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)) * RsVta!Cantidad)
            ImporteIva = ImporteIva + (((RsVta!PrecioVta - RsVta!Costo) - ((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas))) * RsVta!Cantidad)
            
            RsVta.MoveNext
        Loop
        
        txtSubTotal.text = Subtotal
        txtIva.text = ImporteIva 'SubTotal * (Regresa_Valor_BD("IvaVentas") / 100)
        
    End If
Salir:
    RsVta.Close
    Set RsVta = Nothing
    
End Sub

'07-NOV-2011 ::: Buscar los movimientos de la Venta del Día
Private Sub BuscarFacturaDia(ByVal Fecha As Date)
    
    '***1- OBTENER LOS INTERESES A FACTURAR DE LA VENTA DEL DIA ***
    '      INTERESES - ALMACENAJE - SEGURO - MORATORIOS
    Dim RsEmp As New ADODB.Recordset
    Dim Sql As String
    Dim ImpIntereses As Double, ImpAlmacenaje As Double, ImpSeguro As Double, ImpMoratorios As Double, ImpPerdida As Double, ImpOtros As Double
    Dim Subtotal As Double, EmpImporteIva As Double
    Dim StFacturado As String
    Dim IvaFacturacionVentas As Double
    Dim RsVta As New ADODB.Recordset
    Dim VtaMosImporteIva As Double, VtaAparImporteIva As Double, VtaBillImporteIva As Double, VtaMayImporteIva As Double
    On Error GoTo Error

    StFacturado = "Facturado=0"
    Subtotal = 0
    
    If VerificarFacturaDia(Fecha) = True Then
        If MsgBox("La Factura del Día " & CStr(Fecha) & " ya ha sido realizada. ¿Desea Volver a Refacturar?", vbYesNo + vbExclamation, "Facturación") = vbNo Then
            Limpiar
            Exit Sub
        Else
            'Asignar el filtro para tomar los docuemntos sin facturar o facturados en factura del dia previa
            StFacturado = "(Facturado=0 OR Facturado=2)"
            
            Buscar CInt(txtNombre.Tag)
            
        End If
    End If
    
    
    
    '*** OBTENER INTERESES DE REFRENDOS Y DESEMPEÑOS ***
    Sql = "SELECT IDSucursal, Sum(Intereses) as ImpIntereses, Sum(ImporteAlmacenaje) as ImpAlmacenaje, Sum(ImporteSeguro) as ImpSeguro, Sum(ImporteMoratorios) as ImpMoratorios, Sum(ImportePerdida) as ImpPerdida, Sum(ImporteOtros) as ImpOtros, Sum(ImporteIva) as ImpIVA " & _
          "FROM empeno " & _
          "WHERE DAY(fechamovimiento) = " & Day(Fecha) & " and MONTH(fechamovimiento)=" & Month(Fecha) & " and YEAR(fechamovimiento)=" & Year(Fecha) & " AND " & _
                "IDSucursal =" & frmMDI.IDSucursal & " AND cancelado=0 AND " & StFacturado & " AND (DESTINO=2 OR Destino=3) " & _
          "GROUP BY IDSucursal"
    
    RsEmp.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    ImpIntereses = 0: ImpAlmacenaje = 0: ImpSeguro = 0: ImpMoratorios = 0: ImpPerdida = 0: ImpOtros = 0: EmpImporteIva = 0
    If Not RsEmp.EOF Then
        
        ImpIntereses = RsEmp!ImpIntereses + RsEmp!ImpAlmacenaje + RsEmp!ImpSeguro + RsEmp!ImpMoratorios + RsEmp!ImpOtros
''''''' ImpAlmacenaje = RsEmp!ImpAlmacenaje
''''''' ImpSeguro = RsEmp!ImpSeguro
''''''' ImpMoratorios = RsEmp!ImpMoratorios
''''''' ImpPerdida = RsEmp!ImpPerdida
''''''' ImpOtros = RsEmp!ImpOtros
        EmpImporteIva = RsEmp!ImpIVA
        'Do While Not RsEmp.EOF
        'Loop
    End If
    RsEmp.Close
    Set RsEmp = Nothing
    
    With grdFactura
        If ImpIntereses <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "INT": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            '.CellText(.Rows, 4) = "INTERESES"
            .CellText(.Rows, 4) = "INTEXT11/INTERES 11% EXTENSIONES DEL DIA " & Format(Fecha, "YYYY MM DD") & "/RE1X-08919 RE1X-08936"
            .CellText(.Rows, 5) = Format(ImpIntereses, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpIntereses, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpAlmacenaje <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "ALM": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "ALMACENAJE"
            .CellText(.Rows, 5) = Format(ImpAlmacenaje, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpAlmacenaje, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpSeguro <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "SEG": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "SEGURO"
            .CellText(.Rows, 5) = Format(ImpSeguro, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpSeguro, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpMoratorios <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "MOR": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "MORATORIOS / RECARGOS"
            .CellText(.Rows, 5) = Format(ImpMoratorios, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpMoratorios, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
    End With
    
    Subtotal = ImpIntereses + ImpAlmacenaje + ImpSeguro + ImpMoratorios
    
    
    
    '*** OBTENER LAS VENTAS MOSTRADOR (DESGLOCE POR ARTICULO) ***
    
    IvaFacturacionVentas = IIf(Regresa_Valor_BD("IvaVentas") > 0, Regresa_Valor_BD("IvaVentas"), Regresa_Valor_BD("IvaFacturacionVentas")) / 100
    VtaMosImporteIva = 0
    
'''''''    sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, " & _
'''''''                 "dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, " & _
'''''''                 "dv.ImporteIva, v.TipoVenta " & _
'''''''          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
'''''''          "WHERE DAY(v.fecha) = " & Day(Fecha) & " AND MONTH(v.fecha) = " & Month(Fecha) & " AND YEAR(v.fecha) = " & Year(Fecha) & " AND " & _
'''''''          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND " & StFacturado & " AND v.apartado=0 AND v.TipoVenta=" & VENTAMOSTRADOR
    
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Descuento,1 AS Cantidad,v.Pagado,v.Facturado, " & _
                 "dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta," & _
                 "Round (((dv.Precio -(dv.Precio*(v.Descuento/100))) - dv.Costo) / (1 + " & IvaFacturacionVentas & ") ,5) AS Intereses , " & _
                 "0 AS Almacenaje, " & "0 AS Seguro, " & "0 AS Moratorios, " & "0 AS GtosVenta, " & _
                 "Round (((dv.Precio -(dv.Precio*(v.Descuento/100))) - dv.Costo) - ( ((dv.Precio -(dv.Precio*(v.Descuento/100))) - dv.Costo) / (1 + " & IvaFacturacionVentas & ") ),5) AS ImporteIva," & _
                 "v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE DAY(v.fecha) = " & Day(Fecha) & " AND MONTH(v.fecha) = " & Month(Fecha) & " AND YEAR(v.fecha) = " & Year(Fecha) & " AND " & _
          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND " & StFacturado & " AND v.apartado=0 AND v.TipoVenta =" & VENTAMOSTRADOR & " "
    
    Clipboard.Clear
    Clipboard.SetText Sql
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsVta.EOF Then
        Do While Not RsVta.EOF
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio & "-" & RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
'''''''         .CellText(.Rows, 5) = Format(((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''         .CellText(.Rows, 6) = Format(((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 5) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
            End With
'''''''     SubTotal = SubTotal + (((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)) * RsVta!Cantidad)
''''''      VtaMosImporteIva = VtaMosImporteIva + (((RsVta!PrecioVta - RsVta!Costo) - ((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas))) * RsVta!Cantidad)
            Subtotal = Subtotal + ((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta) * RsVta!Cantidad)
            VtaMosImporteIva = VtaMosImporteIva + (RsVta!ImporteIva)
            RsVta.MoveNext
        Loop
    End If
    RsVta.Close
    Set RsVta = Nothing
    
    
    
'''''''    '*** OBTENER LAS VENTAS MAYORISTA EN LA FECHA (DESGLOCE POR ARTICULO) ***
'''''''    sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Descuento,1 AS Cantidad,v.Pagado,v.Facturado, " & _
'''''''                 "dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta," & _
'''''''                 "dv.Intereses , " & _
'''''''                 "dv.Almacenaje, " & "dv.Seguro, " & "dv.Moratorios, " & "dv.GtosVenta, " & _
'''''''                 "dv.ImporteIva," & _
'''''''                 "v.TipoVenta " & _
'''''''          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
'''''''          "WHERE DAY(v.fecha) = " & Day(Fecha) & " AND MONTH(v.fecha) = " & Month(Fecha) & " AND YEAR(v.fecha) = " & Year(Fecha) & " AND " & _
'''''''          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND " & StFacturado & " AND v.apartado=0 AND v.TipoVenta =" & VENTAMAYORISTA & " "
'''''''
'''''''    Clipboard.Clear
'''''''    Clipboard.SetText sql
'''''''
'''''''    RsVta.Open sql, dbDatos, adOpenForwardOnly, adLockReadOnly
'''''''    If Not RsVta.EOF Then
'''''''        Do While Not RsVta.EOF
'''''''            With grdFactura
'''''''                .AddRow
'''''''                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''                .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio & "-" & RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
'''''''                .CellText(.Rows, 5) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                .CellText(.Rows, 6) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                .CellText(.Rows, 7) = RsVta!ID
'''''''            End With
'''''''            SubTotal = SubTotal + ((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta) * RsVta!Cantidad)
'''''''            VtaMosImporteIva = VtaMosImporteIva + (ImporteIva)
'''''''            RsVta.MoveNext
'''''''        Loop
'''''''    End If
'''''''    RsVta.Close
'''''''    Set RsVta = Nothing
    
    
    
    '*** OBTENER LAS VENTAS POR APARTADO PAGADAS EN LA FECHA (DESGLOCE POR ARTICULO)
    VtaAparImporteIva = 0
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.FechaMovimiento,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.ImporteIva, v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE DAY(v.fechamovimiento) = " & Day(Fecha) & " AND MONTH(v.fechamovimiento) = " & Month(Fecha) & " AND YEAR(v.fechamovimiento) = " & Year(Fecha) & " AND " & _
          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND " & StFacturado & " AND v.TipoVenta =" & VENTAMOSTRADOR & " AND v.pagado=1 AND v.apartado=1  "
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsVta.EOF Then
        Do While Not RsVta.EOF
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio & "-" & RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
                .CellText(.Rows, 5) = Format(((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format(((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
            End With
            Subtotal = Subtotal + (((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)) * RsVta!Cantidad)
            VtaAparImporteIva = VtaAparImporteIva + (((RsVta!PrecioVta - RsVta!Costo) - ((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas))) * RsVta!Cantidad)
            RsVta.MoveNext
        Loop
    End If
    RsVta.Close
    Set RsVta = Nothing
    
    
    
    '*** OBTENER LAS VENTAS CLIENTE O BILLETE EN LA FECHA
    VtaBillImporteIva = 0
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, " & _
                 "dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.Intereses, dv.Almacenaje, dv.Seguro, dv.Moratorios, dv.GtosVenta, dv.ImporteIva, v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE DAY(v.fecha) = " & Day(Fecha) & " AND MONTH(v.fecha) = " & Month(Fecha) & " AND YEAR(v.fecha) = " & Year(Fecha) & " AND " & _
                "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND " & StFacturado & " and v.apartado=0 AND v.TipoVenta=" & VENTACLIENTE
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsVta.EOF Then
        Do While Not RsVta.EOF
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio & "-" & RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
'''''''         .CellText(.Rows, 5) = Format((RsVta!PrecioVta - RsVta!Costo), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''         .CellText(.Rows, 6) = Format((RsVta!PrecioVta - RsVta!Costo), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 5) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
                
'''''''                '*** Intereses del la venta ***
'''''''                If RsVta!Intereses <> 0 Then
'''''''                    .AddRow
'''''''                    .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 3) = "INT": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 4) = "INTERESES VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio
'''''''                    .CellText(.Rows, 5) = Format(RsVta!Intereses, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 6) = Format(RsVta!Intereses, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                End If
'''''''
'''''''                If RsVta!Almacenaje <> 0 Then
'''''''                    .AddRow
'''''''                    .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 3) = "ALM": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 4) = "ALMACENAJE VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio
'''''''                    .CellText(.Rows, 5) = Format(RsVta!Almacenaje, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 6) = Format(RsVta!Almacenaje, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                End If
'''''''
'''''''                If RsVta!Seguro <> 0 Then
'''''''                    .AddRow
'''''''                    .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 3) = "SEG": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 4) = "SEGURO VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio
'''''''                    .CellText(.Rows, 5) = Format(RsVta!Seguro, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 6) = Format(RsVta!Seguro, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                End If
'''''''
'''''''                If RsVta!Moratorios <> 0 Then
'''''''                    .AddRow
'''''''                    .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 3) = "MOR": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 4) = "MORATORIOS / RECARGOS VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio
'''''''                    .CellText(.Rows, 5) = Format(RsVta!Moratorios, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 6) = Format(RsVta!Moratorios, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                End If
'''''''
'''''''                If RsVta!GTOSVenta <> 0 Then
'''''''                    .AddRow
'''''''                    .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 3) = "GTO": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 4) = "GASTOS DE VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio
'''''''                    .CellText(.Rows, 5) = Format(RsVta!GastosVenta, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                    .CellText(.Rows, 6) = Format(RsVta!GastosVenta, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''                End If
            End With
            Subtotal = Subtotal + (RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta)
            VtaBillImporteIva = VtaBillImporteIva + (RsVta!ImporteIva)
            RsVta.MoveNext
        Loop
    End If
    RsVta.Close
    Set RsVta = Nothing
    
    txtSubTotal.text = Subtotal
    txtIva.text = EmpImporteIva + VtaMosImporteIva + VtaAparImporteIva + VtaBillImporteIva 'SubTotal * (Regresa_Valor_BD("IvaVentas") / 100)
    
Exit Sub

Error:
    MsgBox "Error al cargar Refrendos/Desempeños y Ventas del Día.", vbCritical, "Error"
    Limpiar
    Maneja_Error Err
End Sub

Private Function ObtenerDescripcionArticulo(ByVal CodigoArt As String) As String
    Dim RsArt As New ADODB.Recordset
    
    On Error GoTo Error
    
    ObtenerDescripcionArticulo = ""
    RsArt.Open "SELECT Marca,Modelo,Observaciones FROM detallesentradainventario WHERE Codigo='" & CodigoArt & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsArt.EOF Then
        ObtenerDescripcionArticulo = IIf(RsArt!Marca <> "", "MARCA " & Trim(UCase(RsArt!Marca)), "") & IIf(RsArt!Modelo <> "", " MODELO " & Trim(UCase(RsArt!Modelo)), "") & IIf(RsArt!Observaciones <> "", " " & Trim(UCase(RsArt!Observaciones)), "")
    End If
    RsArt.Close
    Set RsArt = Nothing
    
Exit Function
    
Error:
    Err.Clear
    If RsArt.State = 1 Then RsArt.Close: Set RsArt = Nothing
    ObtenerDescripcionArticulo = ""
End Function

Private Function ObtenerDescripVenta(ByVal TipoVenta As Integer) As String
    On Error GoTo Error

    Select Case TipoVenta
        Case 0
            ObtenerDescripVenta = "MOSTRADOR"
        Case 1
            ObtenerDescripVenta = "CLIENTE"
        Case 2
            ObtenerDescripVenta = "MAYORISTA"
    End Select
    
Exit Function
Error:
    ObtenerDescripVenta = ""
End Function

Private Function VerificarFacturaDia(ByVal Fecha As Date) As Boolean
    Dim RsFac As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo Error
    
    VerificarFacturaDia = False
    strSql = "SELECT Folio,Cliente FROM facturas WHERE DAY(fecha) = " & Day(Fecha) & " and MONTH(fecha)=" & Month(Fecha) & " and YEAR(fecha)=" & Year(Fecha) & " AND TipoFactura = 0"
    RsFac.Open strSql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsFac.EOF Then
        VerificarFacturaDia = True
        txtNombre.Tag = RsFac!Cliente
    End If
    RsFac.Close
    Set RsFac = Nothing
    
Exit Function
Error:
    Maneja_Error Err
    VerificarFacturaDia = False
    Set RsFac = Nothing
End Function



Private Sub MarcarDocumentosFacturados(ByVal Folio As Long, ByVal TipoFactura As Integer)
    Dim Fecha As Date
    
    '*** Valor de Estatus de Facturado ****
    '1: Facturado por solicitud del cliente
    '2: Facturado por Factura del Dia
    '**************************************
    
    On Error GoTo Error
    
    Select Case TipoFactura
    
        Case 0
            '*** VENTA DEL DIA ***
            Fecha = CDate(lblFecha.Caption)
            
            'Marcar Facturado los refrendos/desempeños de la Venta del Dia
            dbDatos.Execute "UPDATE empeno SET Facturado = 2 " & _
                            "WHERE DAY(fechamovimiento) = " & Day(Fecha) & " and MONTH(fechamovimiento)=" & Month(Fecha) & " and YEAR(fechamovimiento)=" & Year(Fecha) & " AND " & _
                            "IDSucursal =" & frmMDI.IDSucursal & " AND cancelado=0 AND facturado=0 AND (DESTINO=2 OR Destino=3) "

            'Marcar Facturado las Ventas a Cliente Contado para factura Venta del Dia
            dbDatos.Execute "UPDATE ventas SET Facturado = 2 " & _
                            "WHERE DAY(fecha) = " & Day(Fecha) & " AND MONTH(fecha) = " & Month(Fecha) & " AND YEAR(fecha) = " & Year(Fecha) & " AND " & _
                            "IDSucursal=" & frmMDI.IDSucursal & " AND Cancelado=0 AND Facturado=0 and apartado=0"
   
            'Marcar Facturado las Ventas a Cliente por Anticipo Pagados para factura del dia
            dbDatos.Execute "UPDATE ventas SET Facturado = 2 " & _
                            "Where DAY(fechamovimiento) = " & Day(Fecha) & " AND MONTH(fechamovimiento) = " & Month(Fecha) & " AND YEAR(fechamovimiento) = " & Year(Fecha) & " AND " & _
                            "IDSucursal=" & frmMDI.IDSucursal & " AND Cancelado=0 AND Facturado=0 and (pagado=1 OR apartado=1)"
      
        Case 1
            '*** REFRENDOS ***
            dbDatos.Execute "UPDATE empeno SET Facturado = 1 " & _
                            "WHERE FolioNota = " & Folio & " AND " & _
                            "IDSucursal =" & frmMDI.IDSucursal & " AND cancelado=0 AND facturado=0 "
          
        Case 2
            '*** DESEMPEÑOS ***
            dbDatos.Execute "UPDATE empeno SET Facturado = 1 " & _
                            "WHERE FolioNota = " & Folio & " AND " & _
                            "IDSucursal =" & frmMDI.IDSucursal & " AND cancelado=0 AND facturado=0 "
                
        Case 3
            '*** VENTA A CLIENTE ***
            dbDatos.Execute "UPDATE ventas SET Facturado = 1 " & _
                            "WHERE Folio=" & Folio & " AND " & _
                            "IDSucursal=" & frmMDI.IDSucursal & " AND Cancelado=0 AND TipoVenta=" & VENTAMOSTRADOR
        Case 4
        
            '*** VENTA BILLETE *** 15-DIC-2011
            dbDatos.Execute "UPDATE ventas SET Facturado = 1 " & _
                            "WHERE Folio=" & Folio & " AND " & _
                            "IDSucursal=" & frmMDI.IDSucursal & " AND Cancelado=0 AND TipoVenta=" & VENTACLIENTE
        Case 5
        
            '*** VENTA MAYORISTA *** 16-DIC-2011
            dbDatos.Execute "UPDATE ventas SET Facturado = 1 " & _
                            "WHERE Folio=" & Folio & " AND " & _
                            "IDSucursal=" & frmMDI.IDSucursal & " AND Cancelado=0 AND TipoVenta=" & VENTAMAYORISTA
        
    End Select


Exit Sub

Error:
    MsgBox "Ocurrio un error al marcar Documentos Facturados. Tipo Factura " & TipoFactura, vbCritical, "Error"
    Maneja_Error Err

End Sub

Private Function GeneraTxtFactura(ByVal Folio As Long, ByVal Fecha As Date) As String
    Dim sNombreArch As String
    
    sNombreArch = "FAC-" & Format(Folio, "000000") & "-" & CStr(Format(Day(Fecha), "00")) & CStr(Format(Month(Fecha), "00")) & CStr(Format(Year(Fecha), "00")) & ".txt"
    
    CrearArchivoFactura sNombreArch
    
    GeneraTxtFactura = App.Path & "\Facturas\" & sNombreArch
    
End Function

'CREACION DEL ARCHIVO DE TEXTO PARA LA FACTURA
Private Sub CrearArchivoFactura(Archivo)
    Dim strArchivo As String
    Dim iArchivo As Long
    
    'strArchivo = "C:\Polizas\PI" & Poliza & ".txt"
    strArchivo = App.Path & "\Facturas\" & Archivo
   
    If Dir(App.Path & "\Facturas", vbDirectory) = "" Then MkDir App.Path & "\Facturas"
   
    iArchivo = FreeFile
    Open strArchivo For Output Access Write As #iArchivo
    
    Crear_Datos_Comprobante iArchivo, frmMDI.IDSucursal
    
    Crear_Datos_Emisor iArchivo, frmMDI.IDSucursal
    
    Crear_Datos_Receptor iArchivo, txtNombre.Tag
    
    Crear_Datos_Detalle iArchivo
    
    Crear_datos_Sumario iArchivo
    
    Close #iArchivo

End Sub

Private Sub Crear_Datos_Comprobante(iArchivo As Long, ByVal NumSuc As Long)
    Dim Linea As String
    Dim RsSuc As New ADODB.Recordset
    Dim Sql As String
    
    Linea = "[Comprobante]": Print #iArchivo, Linea
    Linea = "Serie=" & Trim(lblSerieFacturacion.Caption): Print #iArchivo, Linea
    Linea = "Folio=" & Trim(txtFolio.text): Print #iArchivo, Linea
''''Linea = "Fecha=" & Format(Now, "YYYY/MM/DD HH:MM:SS"): Print #iArchivo, Linea
    Linea = "Fecha=" & Format(DateAdd("n", -5, Now), "YYYY/MM/DD HH:MM:SS"): Print #iArchivo, Linea
    Linea = "FormaDePago=Pago en una sola exhibicion": Print #iArchivo, Linea
    Linea = "CondicionesDePago=Contado": Print #iArchivo, Linea
    Linea = "Subtotal=" & Format(txtSubTotal, "##########0.00"): Print #iArchivo, Linea
    Linea = "Descuento=0": Print #iArchivo, Linea
    Linea = "MotivoDescuento=": Print #iArchivo, Linea
    Linea = "TipoCambio=1.0": Print #iArchivo, Linea
    Linea = "Moneda=MXP": Print #iArchivo, Linea
    Linea = "Total=" & Format(lblTotal.text, "##########0.00"): Print #iArchivo, Linea
    Linea = "MetodoDePago=Efectivo": Print #iArchivo, Linea
    Linea = "TipoDeComprobante=ingreso": Print #iArchivo, Linea
    Linea = "NoAprobacion=1": Print #iArchivo, Linea
    Linea = "AnoAprobacion=2011": Print #iArchivo, Linea
    Linea = "Observaciones=" & Trim(txtNotas.text): Print #iArchivo, Linea
    Sql = "SELECT * FROM sucursales WHERE Clave=" & NumSuc
    RsSuc.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsSuc.EOF Then
        With RsSuc
            Linea = "LugarExpedicion=" & !Ciudad_Exp & ", " & !Estado_Exp: Print #iArchivo, Linea
        End With
    End If
    
    Linea = "": Print #iArchivo, Linea
    RsSuc.Close
    Set RsSuc = Nothing
End Sub


Private Sub Crear_Datos_Emisor(iArchivo As Long, ByVal NumSuc As Long)
    Dim Linea As String
    Dim RsSuc As New ADODB.Recordset
    Dim Sql As String
    
    Sql = "SELECT * FROM sucursales WHERE Clave=" & NumSuc
    RsSuc.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsSuc.EOF Then
        With RsSuc
            Linea = "[Emisor]": Print #iArchivo, Linea
            Linea = "Rfc=" & !RFC: Print #iArchivo, Linea
            Linea = "Nombre=" & !RazonSocial: Print #iArchivo, Linea
            Linea = "Calle=" & !Direccion: Print #iArchivo, Linea
            Linea = "NoExterior=" & !NumExt: Print #iArchivo, Linea
            Linea = "NoInterior=" & !NumInt: Print #iArchivo, Linea
            Linea = "Colonia=" & !Colonia: Print #iArchivo, Linea
            Linea = "Localidad=" & !Ciudad: Print #iArchivo, Linea
            Linea = "Referencia=" & !Referencia: Print #iArchivo, Linea
            Linea = "Municipio=" & !Municipio: Print #iArchivo, Linea
            Linea = "Estado=" & !Estado: Print #iArchivo, Linea
            Linea = "Pais=" & !Pais: Print #iArchivo, Linea
            Linea = "CodigoPostal=" & !CP: Print #iArchivo, Linea
            Linea = "Email=" & !Email: Print #iArchivo, Linea
            Linea = "RegimenFiscal=" & !RegimenFiscal: Print #iArchivo, Linea
            Linea = "": Print #iArchivo, Linea
            
            Linea = "[EmisorExpedidoEn]": Print #iArchivo, Linea
            Linea = "Calle=" & !Direccion_Exp: Print #iArchivo, Linea
            Linea = "NoExterior=" & !NumExt_Exp: Print #iArchivo, Linea
            Linea = "NoInterior=" & !NumInt_Exp: Print #iArchivo, Linea
            Linea = "Colonia=" & !Colonia_Exp: Print #iArchivo, Linea
            Linea = "Localidad=" & !Ciudad_Exp: Print #iArchivo, Linea
            Linea = "Referencia=" & !Referencia_Exp: Print #iArchivo, Linea
            Linea = "Municipio=" & !Municipio_Exp: Print #iArchivo, Linea
            Linea = "Estado=" & !Estado_Exp: Print #iArchivo, Linea
            Linea = "Pais=" & !Pais_Exp: Print #iArchivo, Linea
            Linea = "CodigoPostal=" & !CP_Exp: Print #iArchivo, Linea
            Linea = "Email=" & !Email: Print #iArchivo, Linea
            Linea = "": Print #iArchivo, Linea
            
        End With
        
    End If
    RsSuc.Close
    Set RsSuc = Nothing
    
End Sub


Private Sub Crear_Datos_Receptor(iArchivo As Long, ByVal NumCliente As Long)
    Dim Linea As String
    Dim RsSuc As New ADODB.Recordset
    Dim Sql As String
    
    Sql = "SELECT * FROM clientes WHERE ID=" & NumCliente
    RsSuc.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsSuc.EOF Then
        With RsSuc
            Linea = "[Receptor]": Print #iArchivo, Linea
            Linea = "Rfc=" & !RFC_Fac: Print #iArchivo, Linea
            Linea = "Nombre=" & !RazonSocial_Fac: Print #iArchivo, Linea
            Linea = "Calle=" & !Calle_Fac: Print #iArchivo, Linea
            Linea = "NoExterior=" & !NumExt_Fac: Print #iArchivo, Linea
            Linea = "NoInterior=" & !NumInt_Fac: Print #iArchivo, Linea
            Linea = "Colonia=" & !Colonia_Fac: Print #iArchivo, Linea
            Linea = "Localidad=" & !Ciudad_Fac: Print #iArchivo, Linea
            Linea = "Referencia=" & !Referencia_Fac: Print #iArchivo, Linea
            Linea = "Municipio=" & !Municipio_Fac: Print #iArchivo, Linea
            Linea = "Estado=" & !Estado_Fac: Print #iArchivo, Linea
            Linea = "Pais=" & !Pais_Fac: Print #iArchivo, Linea
            Linea = "CodigoPostal=" & !CP_Fac: Print #iArchivo, Linea
            Linea = "Email=" & !Email_Fac: Print #iArchivo, Linea
            Linea = "": Print #iArchivo, Linea
        End With
        
    End If
    RsSuc.Close
    Set RsSuc = Nothing
End Sub


Private Sub Crear_Datos_Detalle(iArchivo As Long)
    Dim i As Integer
    Dim Linea As String
    
    With grdFactura
        If .Rows <> 0 Then
            For i = 1 To .Rows
                Linea = "[Concepto" & CStr(i) & "]": Print #iArchivo, Linea
                Linea = "Cantidad=" & .CellText(i, 1): Print #iArchivo, Linea
                Linea = "Unidad=" & .CellText(i, 2): Print #iArchivo, Linea
                Linea = "NoIdentificacion=" & .CellText(i, 3): Print #iArchivo, Linea
                Linea = "Descripcion=" & .CellText(i, 4): Print #iArchivo, Linea
                Linea = "ValorUnitario=" & Format(.CellText(i, 5), "##########0.00"): Print #iArchivo, Linea
                Linea = "Importe=" & Format(.CellText(i, 6), "##########0.00"): Print #iArchivo, Linea
                Linea = "": Print #iArchivo, Linea
            Next i
        End If
    End With

End Sub

Private Sub Crear_datos_Sumario(iArchivo As Long)
    Dim Linea As String
    
    Linea = "[Impuestos]": Print #iArchivo, Linea
    Linea = "IVARetenido=0.00": Print #iArchivo, Linea
    Linea = "ISRRetenido=0.00": Print #iArchivo, Linea
    Linea = "": Print #iArchivo, Linea
    Linea = "IEPSTrasladado=0.00": Print #iArchivo, Linea
    Linea = "IEPSTasa=0.0": Print #iArchivo, Linea
    Linea = "": Print #iArchivo, Linea
    Linea = "IVATrasladado1=" & Format(txtIva.text, "######0.00"): Print #iArchivo, Linea
    Linea = "IVATasa1=" & Format(Regresa_Valor_BD("Iva"), "##0.00"): Print #iArchivo, Linea
    Linea = "": Print #iArchivo, Linea
    Linea = "IVATrasladado2=0.00": Print #iArchivo, Linea
    Linea = "IVATasa2=0.00": Print #iArchivo, Linea
End Sub


Private Sub txtRazonSocialFac_GotFocus()
    Seleccionar_Texto txtRazonSocialFac
    Cambiar_Color True, txtRazonSocialFac
End Sub

Private Sub txtRazonSocialFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRazonSocialFac_LostFocus()
    Cambiar_Color False, txtRazonSocialFac
End Sub

Private Sub txtRFCFac_GotFocus()
    Seleccionar_Texto txtRFCFac
    Cambiar_Color True, txtRFCFac
End Sub

Private Sub txtRFCFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRFCFac_LostFocus()
    Cambiar_Color False, txtRFCFac
End Sub

Private Sub txtCalleFac_GotFocus()
    Seleccionar_Texto txtCalleFac
    Cambiar_Color True, txtCalleFac
End Sub

Private Sub txtCalleFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCalleFac_LostFocus()
    Cambiar_Color False, txtCalleFac
End Sub

Private Sub txtNumExtFac_GotFocus()
    Seleccionar_Texto txtNumExtFac
    Cambiar_Color True, txtNumExtFac
End Sub

Private Sub txtNumExtFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumExtFac_LostFocus()
    Cambiar_Color False, txtNumExtFac
End Sub

Private Sub txtNumIntFac_GotFocus()
    Seleccionar_Texto txtNumIntFac
    Cambiar_Color True, txtNumIntFac
End Sub

Private Sub txtNumIntFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumIntFac_LostFocus()
    Cambiar_Color False, txtNumIntFac
End Sub

Private Sub txtColoniaFac_GotFocus()
    Seleccionar_Texto txtColoniaFac
    Cambiar_Color True, txtColoniaFac
End Sub

Private Sub txtColoniaFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtColoniaFac_LostFocus()
    Cambiar_Color False, txtColoniaFac
End Sub

Private Sub txtCPFac_GotFocus()
    Seleccionar_Texto txtCPFac
    Cambiar_Color True, txtCPFac
End Sub

Private Sub txtCPFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCPFac_LostFocus()
    Cambiar_Color False, txtCPFac
End Sub

Private Sub txtCiudadFac_GotFocus()
    Seleccionar_Texto txtCiudadFac
    Cambiar_Color True, txtCiudadFac
End Sub

Private Sub txtCiudadFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCiudadFac_LostFocus()
    Cambiar_Color False, txtCiudadFac
End Sub

Private Sub txtEstadoFac_GotFocus()
    Seleccionar_Texto txtEstadoFac
    Cambiar_Color True, txtEstadoFac
End Sub

Private Sub txtEstadoFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEstadoFac_LostFocus()
    Cambiar_Color False, txtEstadoFac
End Sub

Private Sub txtPaisFac_GotFocus()
    Seleccionar_Texto txtPaisFac
    Cambiar_Color True, txtPaisFac
End Sub

Private Sub txtPaisFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPaisFac_LostFocus()
    Cambiar_Color False, txtPaisFac
End Sub

Private Sub txtEmailFac_GotFocus()
    Seleccionar_Texto txtEmailFac
    Cambiar_Color True, txtEmailFac
End Sub

Private Sub txtEmailFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEmailFac_LostFocus()
    Cambiar_Color False, txtEmailFac
End Sub

Private Sub txtMunicipioFac_GotFocus()
    Seleccionar_Texto txtMunicipioFac
    Cambiar_Color True, txtMunicipioFac
End Sub

Private Sub txtMunicipioFac_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMunicipioFac_LostFocus()
    Cambiar_Color False, txtMunicipioFac
End Sub

'Funcion para validar los datos de Emision de Factura de la Sucursal
Private Function ValidarEmisor(ByVal NumSuc As Integer) As Boolean
    Dim RsSuc As New ADODB.Recordset
    Dim sMensaje As String
    Dim Valor As Boolean
    
    On Error GoTo Error
    
    Valor = True
    
    RsSuc.Open "SELECT * FROM sucursales WHERE Clave=" & NumSuc & " AND Activa=1", dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsSuc.EOF Then
        With RsSuc
            If IsNull(!RazonSocial) Or !RazonSocial = "" Then sMensaje = "la Razon Social": Valor = False: GoTo Salida
            If IsNull(!RFC) Or !RFC = "" Then sMensaje = "el RFC": Valor = False: GoTo Salida
            If IsNull(!Direccion) Or !Direccion = "" Then sMensaje = "la Calle del Domicilio": Valor = False: GoTo Salida
            If IsNull(!NumExt) Or !NumExt = "" Then sMensaje = "el Numero Exterior del Domicilio": Valor = False: GoTo Salida
            If IsNull(!Colonia) Or !Colonia = "" Then sMensaje = "la Colonia del Domicilio": Valor = False: GoTo Salida
            If IsNull(!Ciudad) Or !Ciudad = "" Then sMensaje = "la Ciudad del Domicilio": Valor = False: GoTo Salida
            If IsNull(!Municipio) Or !Municipio = "" Then sMensaje = "el Municipio del Domicilio": Valor = False: GoTo Salida
            If IsNull(!Estado) Or !Estado = "" Then sMensaje = "el Estado del Domicilio": Valor = False: GoTo Salida
            If IsNull(!Pais) Or !Pais = "" Then sMensaje = "el País": Valor = False: GoTo Salida
            If IsNull(!CP) Or !CP = "" Then sMensaje = "el Código Postal": Valor = False: GoTo Salida
            
            If IsNull(!Direccion_Exp) Or !Direccion_Exp = "" Then sMensaje = "la Calle del Domicilio de Expedición": Valor = False: GoTo Salida
            If IsNull(!NumExt_Exp) Or !NumExt_Exp = "" Then sMensaje = "el Numero Exterior del Domicilio de Expedición": Valor = False: GoTo Salida
            If IsNull(!Colonia_Exp) Or !Colonia_Exp = "" Then sMensaje = "la Colonia del Domicilio de Expedición": Valor = False: GoTo Salida
            If IsNull(!Ciudad_Exp) Or !Ciudad_Exp = "" Then sMensaje = "la Ciudad del Domicilio de Expedición": Valor = False: GoTo Salida
            If IsNull(!Municipio_Exp) Or !Municipio_Exp = "" Then sMensaje = "el Municipio del Domicilio de Expedición": Valor = False: GoTo Salida
            If IsNull(!Estado_Exp) Or !Estado_Exp = "" Then sMensaje = "el Estado del Domicilio de Expedición": Valor = False: GoTo Salida
            If IsNull(!Pais_Exp) Or !Pais_Exp = "" Then sMensaje = "el País de Domicilio de Expedición": Valor = False: GoTo Salida
            If IsNull(!CP_Exp) Or !CP_Exp = "" Then sMensaje = "el Código Postal de Domicilio de Expedición": Valor = False: GoTo Salida
        End With
    Else
        MsgBox "No existen datos completos de la Sucursal Emisora de la Factura.", vbCritical, "Facturación"
        Valor = False
    End If
    ValidarEmisor = Valor
    RsSuc.Close
    Set RsSuc = Nothing
    
Exit Function
       
Salida:
    MsgBox "Ingrese " & sMensaje & " de la Sucursal Emisora de la Factura.", vbInformation, "Facturacion"
    ValidarEmisor = Valor
    RsSuc.Close
    Set RsSuc = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    ValidarEmisor = False
    If RsSuc.State = 1 Then RsSuc.Close: Set RsSuc = Nothing
End Function


'15-DIC-2011 ::: Buscar la Venta a Cliente (Billete) a Facturar
Private Sub BuscarVentaBillete(ByVal Folio As Long)
    
    '***2- OBTENER LA VENTA BILLETE O CLIENTE (DESGLOCE POR ARTICULO) ***
    Dim RsVta As New ADODB.Recordset
    Dim Sql As String
    Dim Subtotal As Double
    Dim ImporteIva As Double
    
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.Intereses, dv.Almacenaje, dv.Seguro, dv.Moratorios, dv.GtosVenta, dv.ImporteIva, dv.Costo " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE v.Folio=" & Folio & " AND " & "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND v.TipoVenta=" & VENTACLIENTE
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If RsVta.EOF Then
        MsgBox "No se encontró el Contrato de la Venta especificado !!", vbInformation, "Facturación"
        txtFolioBillete.Tag = ""
        Subtotal = 0
        Limpiar
    Else
        Limpiar
        If RsVta!Apartado = 1 And RsVta!Pagado = 0 Then
            MsgBox "El Contrato de la Venta no ha sido Pagado en su totalidad. !!", vbInformation, "Facturación"
            txtFolioBillete.Tag = ""
            Subtotal = 0
            Limpiar
            GoTo Salir
        End If
    
        If RsVta!Facturado = 1 Then
            If MsgBox("El Contrato de Venta Cliente ya ha sido facturado. ¿Desea Volver a Facturar?", vbYesNo, "Facturación") = vbNo Then
                txtFolioBillete.Tag = ""
                Subtotal = 0
                Limpiar
                GoTo Salir
            End If
        ElseIf RsVta!Facturado = 2 Then
            MsgBox "El Contrato de la Venta Cliente ya se facturó en Factura del Día. No se puede refacturar. !!", vbInformation, "Facturación"
            txtFolioBillete.Tag = ""
            Subtotal = 0
            Limpiar
            GoTo Salir
        End If
        
        Buscar RsVta!IDCliente
        txtFolioBillete.text = RsVta!Folio
        txtFolioBillete.Tag = RsVta!Folio
        
        Subtotal = 0: ImporteIva = 0
        Do While Not RsVta.EOF
            
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
                .CellText(.Rows, 5) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
            End With
            
            Subtotal = Subtotal + (RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios)
            ImporteIva = ImporteIva + (RsVta!ImporteIva)
            
            RsVta.MoveNext
        Loop
        
        txtSubTotal.text = Subtotal
        txtIva.text = ImporteIva
        
    End If
Salir:
    RsVta.Close
    Set RsVta = Nothing
    
End Sub

Private Sub txtFolioBillete_Change()
    txtFolioBillete.Tag = ""
End Sub

Private Sub txtFolioBillete_GotFocus()
    Seleccionar_Texto txtFolioBillete
    Cambiar_Color True, txtFolioBillete
End Sub

Private Sub txtFolioBillete_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtFolioBillete.text) <> "" Then
            BuscarVentaBillete txtFolioBillete
        Else
            MsgBox "Introduzca el número de Contrato de Venta Billete que desea Facturar !!", vbCritical, "Facturación"
        End If
    End If
End Sub

Private Sub txtFolioBillete_LostFocus()
    Cambiar_Color False, txtFolioBillete
End Sub


Private Sub txtFolioVtaMay_Change()
    txtFolioVenta.Tag = ""
End Sub

Private Sub txtFolioVtaMay_GotFocus()
    Seleccionar_Texto txtFolioVtaMay
    Cambiar_Color True, txtFolioVtaMay
End Sub

Private Sub txtFolioVtaMay_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        If Trim(txtFolioVtaMay.text) <> "" Then
            BuscarVentaMay txtFolioVtaMay
        Else
            MsgBox "Introduzca el número de Folio Venta Mayorista que desea Facturar !!", vbCritical, "Facturación"
        End If
    End If
End Sub

Private Sub txtFolioVtaMay_LostFocus()
    Cambiar_Color False, txtFolioVtaMay
End Sub

'16-DIC-2011 ::: Buscar la Venta Mayorista a Facturar
Private Sub BuscarVentaMay(ByVal Folio As Long)
    
    '***2- OBTENER LAS VENTAS DEL DIA (DESGLOCE POR ARTICULO) ***
    Dim RsVta As New ADODB.Recordset
    Dim Sql As String
    Dim Subtotal As Double
    Dim ImporteIva As Double
    
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.Intereses, dv.Almacenaje, dv.Seguro, dv.Moratorios, dv.ImporteIva " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE v.Folio=" & Folio & " AND " & "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND v.TipoVenta=" & VENTAMAYORISTA
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If RsVta.EOF Then
        MsgBox "No se encontró el Folio de Venta Mayorista especificado !!", vbInformation, "Facturación"
        txtFolioVtaMay.Tag = ""
        Subtotal = 0: ImporteIva = 0
        Limpiar
    Else
        Limpiar
        If RsVta!Apartado = 1 And RsVta!Pagado = 0 Then
            MsgBox "El Folio de Venta Mayorista no ha sido Pagado en su totalidad. !!", vbInformation, "Facturación"
            txtFolioVtaMay.Tag = ""
            Subtotal = 0: ImporteIva = 0
            Limpiar
            GoTo Salir
        End If
    
        If RsVta!Facturado = 1 Then
            If MsgBox("El Folio de Venta Mayorista ya ha sido facturado. ¿Desea Volver a Facturar?", vbYesNo, "Facturación") = vbNo Then
                txtFolioVtaMay.Tag = ""
                Subtotal = 0: ImporteIva = 0
                Limpiar
                GoTo Salir
            End If
        ElseIf RsVta!Facturado = 2 Then
            MsgBox "El Folio de Venta Mayorista ya se facturó en Factura del Día. No se puede refacturar. !!", vbInformation, "Facturación"
            txtFolioVtaMay.Tag = ""
            Subtotal = 0: ImporteIva = 0
            Limpiar
            GoTo Salir
        End If
        
        'Buscar_Cliente RsVta!IDCliente
        txtFolioVtaMay.text = RsVta!Folio
        txtFolioVtaMay.Tag = RsVta!Folio
        
        Subtotal = 0: ImporteIva = 0
        Do While Not RsVta.EOF
            
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
'''''''         .CellText(.Rows, 5) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''         .CellText(.Rows, 6) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 5) = Format((RsVta!PrecioVta), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format((RsVta!PrecioVta), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
            End With
            
'''''''     SubTotal = SubTotal + ((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios) * RsVta!Cantidad)
'''''''     ImporteIva = ImporteIva + (RsVta!ImporteIva)
            Subtotal = Subtotal + ((RsVta!PrecioVta) * RsVta!Cantidad)
            ImporteIva = ImporteIva + (RsVta!ImporteIva)
            
            
            
            RsVta.MoveNext
        Loop
        
        txtSubTotal.text = Subtotal
        txtIva.text = ImporteIva 'SubTotal * (Regresa_Valor_BD("IvaVentas") / 100)
        
    End If
Salir:
    RsVta.Close
    Set RsVta = Nothing
    
End Sub

Private Sub Crear_CFDi(Archivo As String)
   On Error GoTo Error
   Dim Res As Long
   Dim cmd As String
   Dim Salir As Boolean
   Dim Error As String
   Dim xml As String
   
   cmd = """" & Path & "\CFDi\ShellCFDi.exe""" & " -CFDi """ & Archivo & """  """ & Path & "\CFDi\Configuracion.txt"""
   
   Res = Ejecutar_Shell(cmd)
  
  
  While Not Salir
      DoEvents
      Error = Find_Err(Archivo)
      xml = Find_xml(Archivo)
      Salir = (Error <> "" Or xml <> "")
  Wend
       
'  If Error <> "" Then
'      MsgBox Regresa_Valor("Error", "Mensaje", "", Error), vbOKOnly Or vbCritical
'  End If
  If Error <> "" Then
      MsgBox Regresa_Valor_XML("Error", "Mensaje", Replace(Archivo, ".txt", ".err"), Error), vbOKOnly Or vbCritical
  End If
  
  If xml <> "" Then
      MsgBox "Factura Realizada Correctamente", vbOKOnly Or vbInformation
       ShellExecute Me.hWnd, "open", Replace(Archivo, ".txt", ".pdf"), vbNullString, "C:\", SW_SHOWNORMAL
  End If
       
Error:
   Maneja_Error Err
End Sub

Private Function Regresa_Valor_XML(Seccion As String, Key As String, Archivo As String, Default As String) As String
Dim Cadena As String, Lon As Integer
   
    Cadena = String(255, 0)
   
    Lon = GetPrivateProfileString(Seccion, Key, Default, Cadena, 255, Archivo)
    Cadena = Left$(Cadena, Lon)
    Regresa_Valor_XML = Cadena
    
End Function

Private Function Find_xml(Archivo As String) As String
   On Error GoTo Error
    
   Find_xml = Dir(Replace(Archivo, ".txt", ".pdf"))
    
Error:
   Maneja_Error Err
   
End Function

Private Function Find_Err(Archivo As String) As String
   On Error GoTo Error
   
   Find_Err = Dir(Replace(Archivo, ".txt", ".err"))
   
Error:
   Maneja_Error Err
End Function


Private Function Ejecutar_Shell(Archivo As String) As Long
   On Error GoTo Error
   Dim handle_process As Long
   Dim id_process As Long
   Dim lp_ExitCode As Long
   
   id_process = Shell(Archivo, vbNormalFocus)
   Debug.Print id_process
   
   Do
      Call GetExitCodeProcess(handle_process, lp_ExitCode)
      DoEvents
   
   Loop While lp_ExitCode = STATUS_PENDING
   
   Ejecutar_Shell = lp_ExitCode
   
Error:
   Maneja_Error Err
End Function

Private Sub Buscar_Refrendos_Desempenos_Dia(Fecha As Date, Destino As Integer, Opcion As String)
   On Error GoTo Error
   Dim ImpIntereses As Double, ImpAlmacenaje As Double, ImpSeguro As Double, ImpMoratorios As Double, ImpPerdida As Double, ImpOtros As Double
   Dim EmpImporteIva As Currency
   Dim rc As New ADODB.Recordset
   Dim Subtotal As Currency
   Dim Sql As String
   
   Sql = "SELECT IDSucursal, Sum(Intereses) as ImpIntereses, Sum(ImporteAlmacenaje) as ImpAlmacenaje, Sum(ImporteSeguro) as ImpSeguro, Sum(ImporteMoratorios) as ImpMoratorios, Sum(ImportePerdida) as ImpPerdida, Sum(ImporteOtros) as ImpOtros, Sum(ImporteIva) as ImpIVA FROM empeno " & _
          "WHERE DAY(fechamovimiento) = " & Day(Fecha) & " and MONTH(fechamovimiento)=" & Month(Fecha) & " and YEAR(fechamovimiento)=" & Year(Fecha) & " AND " & _
          "IDSucursal =" & frmMDI.IDSucursal & " AND cancelado=0 AND DESTINO=" & Destino & " " & _
          "GROUP BY IDSucursal"
   
   rc.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
   
   If Not rc.EOF Then
      ImpIntereses = rc!ImpIntereses + rc!ImpAlmacenaje + rc!ImpSeguro + rc!ImpMoratorios + rc!ImpOtros
      EmpImporteIva = rc!ImpIVA
   End If
   
   With grdFactura
        If ImpIntereses <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "INT": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "INTERESES " & Opcion
            .CellText(.Rows, 5) = Format(ImpIntereses, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpIntereses, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpAlmacenaje <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "ALM": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "ALMACENAJE  " & Opcion
            .CellText(.Rows, 5) = Format(ImpAlmacenaje, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpAlmacenaje, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpSeguro <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "SEG": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "SEGURO  " & Opcion
            .CellText(.Rows, 5) = Format(ImpSeguro, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpSeguro, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpMoratorios <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "MOR": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "MORATORIOS / RECARGOS  " & Opcion
            .CellText(.Rows, 5) = Format(ImpMoratorios, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpMoratorios, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
    End With
    
    Subtotal = ImpIntereses + ImpAlmacenaje + ImpSeguro + ImpMoratorios
    txtSubTotal.text = Subtotal
    txtIva.text = EmpImporteIva
   
   rc.Close
Error:
   Maneja_Error Err
   
   Set rc = Nothing
End Sub

Private Sub Buscar_Ventas_Dia(Fecha As Date)
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String
    Dim Subtotal As Currency
    Dim VtaMosImporteIva As Currency
    Dim VtaAparImporteIva  As Currency
    Dim VtaBillImporteIva As Currency
    Dim IvaFacturacionVentas As Double
    
    IvaFacturacionVentas = IIf(Regresa_Valor_BD("IvaVentas") > 0, Regresa_Valor_BD("IvaVentas"), Regresa_Valor_BD("IvaFacturacionVentas")) / 100 '
    
   'OBTENER LAS VENTAS MOSTRADOR EN LA FECHA
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.ImporteIva, v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE DAY(v.fecha) = " & Day(Fecha) & " AND MONTH(v.fecha) = " & Month(Fecha) & " AND YEAR(v.fecha) = " & Year(Fecha) & " AND " & _
          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND v.apartado=0 AND v.TipoVenta=" & VENTAMOSTRADOR & " AND v.Facturado=0"
    
    rc.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rc.EOF
        With grdFactura
            .AddRow
            .CellText(.Rows, 1) = rc!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = rc!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(rc!TipoVenta) & " " & rc!Folio & "-" & rc!Articulo & " " & ObtenerDescripcionArticulo(rc!Codigo)
            .CellText(.Rows, 5) = Format(((rc!PrecioVta - rc!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(((rc!PrecioVta - rc!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 7) = rc!ID
        End With
        Subtotal = Subtotal + (((rc!PrecioVta - rc!Costo) / (1 + IvaFacturacionVentas)) * rc!Cantidad)
        VtaMosImporteIva = VtaMosImporteIva + (((rc!PrecioVta - rc!Costo) - ((rc!PrecioVta - rc!Costo) / (1 + IvaFacturacionVentas))) * rc!Cantidad)
        rc.MoveNext
    Wend
    rc.Close
    
    
    '*** OBTENER LAS VENTAS POR APARTADO PAGADAS EN LA FECHA (DESGLOCE POR ARTICULO)
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.FechaMovimiento,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.ImporteIva, v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE DAY(v.fechamovimiento) = " & Day(Fecha) & " AND MONTH(v.fechamovimiento) = " & Month(Fecha) & " AND YEAR(v.fechamovimiento) = " & Year(Fecha) & " AND " & _
          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND v.apartado=1 AND v.TipoVenta =" & VENTAMOSTRADOR & " AND v.pagado=1 " & " AND v.Facturado=0"
    
    rc.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rc.EOF
        With grdFactura
            .AddRow
            .CellText(.Rows, 1) = rc!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = rc!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(rc!TipoVenta) & " " & rc!Folio & "-" & rc!Articulo & " " & ObtenerDescripcionArticulo(rc!Codigo)
            .CellText(.Rows, 5) = Format(((rc!PrecioVta - rc!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(((rc!PrecioVta - rc!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 7) = rc!ID
        End With
        Subtotal = Subtotal + (((rc!PrecioVta - rc!Costo) / (1 + IvaFacturacionVentas)) * rc!Cantidad)
        VtaAparImporteIva = VtaAparImporteIva + (((rc!PrecioVta - rc!Costo) - ((rc!PrecioVta - rc!Costo) / (1 + IvaFacturacionVentas))) * rc!Cantidad)
        rc.MoveNext
    Wend
    rc.Close
    
    
    '*** OBTENER LAS VENTAS CLIENTE O BILLETE EN LA FECHA
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.Intereses, dv.Almacenaje, dv.Seguro, dv.Moratorios, dv.GtosVenta, dv.ImporteIva, dv.Costo, v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE DAY(v.fecha) = " & Day(Fecha) & " AND MONTH(v.fecha) = " & Month(Fecha) & " AND YEAR(v.fecha) = " & Year(Fecha) & " AND " & _
          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND v.apartado=0 AND v.TipoVenta=" & VENTACLIENTE & " AND v.Facturado=0"
    
    rc.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rc.EOF
       With grdFactura
           .AddRow
           .CellText(.Rows, 1) = rc!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
           .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
           .CellText(.Rows, 3) = rc!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
           .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(rc!TipoVenta) & " " & rc!Folio & "-" & rc!Articulo & " " & ObtenerDescripcionArticulo(rc!Codigo)
           .CellText(.Rows, 5) = Format((rc!Intereses + rc!Almacenaje + rc!Seguro + rc!Moratorios), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
           .CellText(.Rows, 6) = Format((rc!Intereses + rc!Almacenaje + rc!Seguro + rc!Moratorios), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
           .CellText(.Rows, 7) = rc!ID
           
'''''''           '*** Intereses del la venta ***
'''''''           If rc!Intereses <> 0 Then
'''''''               .AddRow
'''''''               .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 3) = "INT": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 4) = "INTERESES VENTA " & ObtenerDescripVenta(rc!TipoVenta) & " " & rc!Folio
'''''''               .CellText(.Rows, 5) = Format(rc!Intereses, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 6) = Format(rc!Intereses, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''           End If
'''''''
'''''''           If rc!Almacenaje <> 0 Then
'''''''               .AddRow
'''''''               .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 3) = "ALM": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 4) = "ALMACENAJE VENTA " & ObtenerDescripVenta(rc!TipoVenta) & " " & rc!Folio
'''''''               .CellText(.Rows, 5) = Format(rc!Almacenaje, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 6) = Format(rc!Almacenaje, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''           End If
'''''''
'''''''           If rc!Seguro <> 0 Then
'''''''               .AddRow
'''''''               .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 3) = "SEG": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 4) = "SEGURO VENTA " & ObtenerDescripVenta(rc!TipoVenta) & " " & rc!Folio
'''''''               .CellText(.Rows, 5) = Format(rc!Seguro, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 6) = Format(rc!Seguro, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''           End If
'''''''
'''''''           If rc!Moratorios <> 0 Then
'''''''               .AddRow
'''''''               .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 3) = "MOR": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 4) = "MORATORIOS / RECARGOS VENTA " & ObtenerDescripVenta(rc!TipoVenta) & " " & rc!Folio
'''''''               .CellText(.Rows, 5) = Format(rc!Moratorios, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 6) = Format(rc!Moratorios, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''           End If
'''''''
'''''''           If rc!GTOSVenta <> 0 Then
'''''''               .AddRow
'''''''               .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 3) = "GTO": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 4) = "GASTOS DE VENTA " & ObtenerDescripVenta(rc!TipoVenta) & " " & rc!Folio
'''''''               .CellText(.Rows, 5) = Format(rc!GastosVenta, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''               .CellText(.Rows, 6) = Format(rc!GastosVenta, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
'''''''           End If
       End With
       Subtotal = Subtotal + ((rc!Intereses + rc!Almacenaje + rc!Seguro + rc!Moratorios) * rc!Cantidad)
       VtaBillImporteIva = VtaBillImporteIva + (rc!ImporteIva)
       rc.MoveNext
    Wend
    rc.Close
    
    txtSubTotal.text = Subtotal
    txtIva.text = VtaMosImporteIva + VtaAparImporteIva + VtaBillImporteIva  'SubTotal * (Regresa_Valor_BD("IvaVentas") / 100)
    
Error:
   Maneja_Error Err
   
   Set rc = Nothing
End Sub
Private Sub CrarCfdiRV()
Dim Linea As String
Dim obj As Genera_cfd_vb6.Libreria
Set obj = New Genera_cfd_vb6.Libreria
Dim RsSuc As New ADODB.Recordset
Dim RsCliente As New ADODB.Recordset
Dim Sql As String
Dim SqlC As String
Dim sqlFac As String
Dim i As Integer
Dim rfcEm As String
Dim rzEm As String
Dim dirEm As String

Dim elSello As String
elSello = Regresa_Valor("CFDI", "CER", App.Path)
Dim llave As String
llave = Regresa_Valor("CFDI", "KEY", App.Path)
Dim clavefiel As String
clavefiel = Regresa_Valor("CFDI", "PSW", App.Path)
Dim utimbrado As String
utimbrado = Regresa_Valor("CFDI", "UTIMBRADO", App.Path)
Dim ptimbrado As String
ptimbrado = Regresa_Valor("CFDI", "PTIMBRADO", App.Path)
Dim IDfac As String
Dim co As String
Dim sqlPar As String
 Dim Fecha As Date
        Fecha = Now()

Sql = "SELECT * FROM sucursales WHERE Clave=" & frmMDI.IDSucursal
SqlC = "SELECT * FROM clientes WHERE ID=" & txtNombre.Tag
sqlFac = "SELECT IDfactura FROM detallefactura d order by ID desc"
sqlPar = "select seriefacturas from parametros"

Dim Serie As String
If lblSerieFacturacion.Caption <> "" Then
Serie = lblSerieFacturacion.Caption
Else
Serie = ""
End If
Dim NumInt As String
If txtNumIntFac.text <> "" Then
NumInt = txtNumIntFac.text
Else
NumInt = ""
End If
'se obtienen los datos de sucursal'
RsSuc.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
If Not RsSuc.EOF Then
   With RsSuc
        Linea = !Ciudad_Exp & ", " & !Estado_Exp
        rfcEm = !RFC
        rzEm = !RazonSocial
        dirEm = !Direccion & " " & Linea
    
        Call obj.agrega_comprobante(3.2, Serie, CLng(txtFolio.text), CStr(DateTime.Now), "Pago en una sola exhibicion", "Contado", CDbl(txtSubTotal.text), 0, "", "1.00", "MXP", CDbl(lblTotal.text), "ingreso", "EFECTIVO", Linea, "")
        Call obj.agregar_emisor(!RFC, !RazonSocial, !Direccion, CStr(!NumExt), CStr(!NumInt), !Colonia, !Ciudad, !Referencia, !Municipio, !Estado, !Pais, CStr(!CP)) '
        Call obj.agregar_expedido(!Direccion_Exp, CStr(!NumExt_Exp), CStr(!NumInt_Exp), !Colonia_Exp, !Ciudad_Exp, !Referencia_Exp, !Municipio_Exp, !Estado_Exp, !Pais_Exp, CStr(!CP_Exp)) '
        Call obj.agregar_receptor(txtRFCFac.text, txtNombre.text & " " & txtApellidos.text, txtCalleFac.text, txtNumExtFac.text, NumInt, txtColoniaFac.text, txtCiudadFac.text, "", txtMunicipioFac.text, txtEstadoFac.text, txtPaisFac.text, txtCPFac.text) '
        Call obj.agregar_regimen(!RegimenFiscal) '
        Call obj.agregar_impuesto(tipoImpuesto_trasladado, "IVA", 16#, Format(txtIva.text, "######0.00")) '
    End With
    
    With grdFactura
        If .Rows <> 0 Then
            For i = 1 To .Rows
                Call obj.agregar_concepto(.CellText(i, 1), .CellText(i, 2), .CellText(i, 3), .CellText(i, 4), Format(.CellText(i, 5), "##########0.00"), Format(.CellText(i, 5), "##########0.00")) '
            Next i
        End If
    End With
    
    RsSuc.Close
    
    
If obj.mensajeDeError <> "" Then
    Call MsgBox(obj.mensajeDeError, 16, "Aviso de sistema")
Else
    Call obj.generar_sello(clavefiel, llave, elSello)
    If obj.mensajeDeError <> "" Then
        Call MsgBox(obj.mensajeDeError, 16, "Aviso de sistema")
    End If
    If obj.mensajeDeError = "" Then
        'Funcion que permite generar el xml
        '===============  Parametros  ===============
        'nombre del archivo  a generar
       
        Call obj.generar_xml_3_2(App.Path & "\Facturas\FAC-" & Format(txtFolio.text, "000000") & "-" & CStr(Format(Day(Fecha), "00")) & CStr(Format(Month(Fecha), "00")) & CStr(Format(Year(Fecha), "00")))
        'Verifica que no exista error
        If (obj.mensajeDeError = "") Then
           ' Call MsgBox("Se genero el XML correctamente", 64, "Aviso de sistema")
            co = CStr(obj.cadena_original)
            obj.Dispose
            'TIMBRADO DE XML
            Call obj.timbrar_CFD(App.Path & "\Facturas\FAC-" & Format(txtFolio.text, "000000") & "-" & CStr(Format(Day(Fecha), "00")) & CStr(Format(Month(Fecha), "00")) & CStr(Format(Year(Fecha), "00")) & ".XML", App.Path & "\Facturas\", "FAC-" & Format(txtFolio.text, "000000") & "-" & CStr(Format(Day(Fecha), "00")) & CStr(Format(Month(Fecha), "00")) & CStr(Format(Year(Fecha), "00")) & ".xml", utimbrado, ptimbrado, "http://generacfdi.com.mx/rvltimbrado/service1.asmx")
           ' Verifica si existe un error
            If (obj.mensajeDeError <> "") Then
                Call MsgBox(obj.mensajeDeError, 16, "Aviso de sistema")
            Else
            'se carga el XML timbrado a memoria para obtener su informacion
                Call obj.AgregarXML(App.Path & "\Facturas\FAC-" & Format(txtFolio.text, "000000") & "-" & CStr(Format(Day(Fecha), "00")) & CStr(Format(Month(Fecha), "00")) & CStr(Format(Year(Fecha), "00")) & ".XML")
               Call obj.ImagenQrcode(App.Path & "\Facturas\QrCode.jpg")
              
                
      'Seccion del rtf
            RsSuc.Open sqlFac, dbDatos, adOpenForwardOnly, adLockReadOnly
                If Not RsSuc.EOF Then
                    With RsSuc
                        IDfac = !IDFactura
                    End With
                End If
                Dim nombrecliente As String
                 nombrecliente = txtNombre.text & " " & txtApellidos.text
                Dim DireccionCliente  As String
                 DireccionCliente = txtCalleFac.text & " " & txtNumExtFac.text & " Interior " & txtNumIntFac.text & " Colonia " & txtColoniaFac.text & " CP " & txtCPFac.text
               Dim parametrosdeshell As String
               Dim lacoma As String
               lacoma = "^"
             parametrosdeshell = CStr(rzEm) & lacoma & rfcEm & lacoma & dirEm & lacoma & nombrecliente & lacoma & txtRFCFac.text & lacoma & DireccionCliente & lacoma & Linea & lacoma & CStr(obj.UUID) & lacoma & CStr(obj.noCertificadoSAT) & lacoma & CStr(obj.fechaTimbrado) & lacoma & Linea & " " & CStr(DateTime.Now) & lacoma & txtFolio.text & lacoma & lblSerieFacturacion.Caption & lacoma & Trim(CantidadEnLetra(CCur(lblTotal.text))) & lacoma & txtIva.text & lacoma & lblTotal.text & lacoma & CStr(obj.sello_SAT) & lacoma & co & lacoma & App.Path & "\Facturas\FAC-" & Format(txtFolio.text, "000000") & "-" & CStr(Format(Day(Fecha), "00")) & CStr(Format(Month(Fecha), "00")) & CStr(Format(Year(Fecha), "00")) & ".pdf" & lacoma & CStr(IDfac)
                Dim envio As String
                
               envio = Replace(parametrosdeshell, " ", "*")
                Call Shell(App.Path & "\Facturas\CrearPDF.exe" & " " & envio, vbNormalFocus)
                
                MsgBox ("Enviado correo electronico")
                
                    
            'Termina rtf
            'Envio por correo
                obj.Vector_cadena = App.Path & "\Facturas\FAC-" & Format(txtFolio.text, "000000") & "-" & CStr(Format(Day(Fecha), "00")) & CStr(Format(Month(Fecha), "00")) & CStr(Format(Year(Fecha), "00")) & ".pdf"
                obj.Vector_cadena = App.Path & "\Facturas\FAC-" & Format(txtFolio.text, "000000") & "-" & CStr(Format(Day(Fecha), "00")) & CStr(Format(Month(Fecha), "00")) & CStr(Format(Year(Fecha), "00")) & ".xml"
               Call obj.EnviarFactura_true_smtp("Sistema de facturacion electronica MySonda www.sigmadesarrollo.com.mx", txtEmailFac.text, "Factura emitida el dia " & CStr(Now), "aquinonez@sigmadesarrollo.com.mx", "aquinonez@sigmadesarrollo.com.mx", "6am3.0ver.AIM090202", "mail.sigmadesarrollo.com.mx", "26", False)
               If obj.mensajeDeError = "" Then
                Call MsgBox("Correo enviado correctamente", 64, "Aviso de sistema")
                Else
                Call MsgBox(obj.mensajeDeError, 16, "Aviso de sistema")
                End If
            'termina Envio por correo
              
            End If
    
        
        Else
        Call MsgBox(obj.mensajeDeError, 16, "Aviso de sistema")
        End If
    End If
        
End If
    
    
    
    
End If
   
End Sub
'07-NOV-2011 ::: Buscar los movimientos de la Venta del Día
Private Sub BuscarFacturaPlazo(ByVal FechaIni As Date, ByVal FechaFin As Date)
    
    '***1- OBTENER LOS INTERESES A FACTURAR DE LA VENTA DEL DIA ***
    '      INTERESES - ALMACENAJE - SEGURO - MORATORIOS
    Dim RsEmp As New ADODB.Recordset
    Dim Sql As String
    Dim ImpIntereses As Double, ImpAlmacenaje As Double, ImpSeguro As Double, ImpMoratorios As Double, ImpPerdida As Double, ImpOtros As Double
    Dim Subtotal As Double, EmpImporteIva As Double
    Dim StFacturado As String
    Dim IvaFacturacionVentas As Double
    Dim RsVta As New ADODB.Recordset
    Dim VtaMosImporteIva As Double, VtaAparImporteIva As Double, VtaBillImporteIva As Double, VtaMayImporteIva As Double
    On Error GoTo Error

    StFacturado = "Facturado=0"
    Subtotal = 0
    
    If VerificarFacturaPlazo(FechaIni, FechaFin) = True Then
        If MsgBox("La Factura del " & CStr(FechaIni) & " al " & CStr(FechaFin) & " ya ha sido realizada o se ha realizado una factura ", vbYesNo + vbExclamation, "Facturación") = vbNo Then
            Limpiar
            Exit Sub
        Else
            'Asignar el filtro para tomar los docuemntos sin facturar o facturados en factura del dia previa
            StFacturado = "(Facturado=0 OR Facturado=2)"
            
            Buscar CInt(txtNombre.Tag)
            
        End If
    End If
    
    
    
    '*** OBTENER INTERESES DE REFRENDOS Y DESEMPEÑOS ***
    Sql = "SELECT IDSucursal, Sum(Intereses) as ImpIntereses, Sum(ImporteAlmacenaje) as ImpAlmacenaje, Sum(ImporteSeguro) as ImpSeguro, Sum(ImporteMoratorios) as ImpMoratorios, Sum(ImportePerdida) as ImpPerdida, Sum(ImporteOtros) as ImpOtros, Sum(ImporteIva) as ImpIVA " & _
          "FROM empeno " & _
          "WHERE date(fechamovimiento) >= date('" & Format(FechaIni, "YYYY/MM/DD") & "') and date(fechamovimiento) <= date('" & Format(FechaFin, "YYYY/MM/DD") & "') AND " & _
                "IDSucursal =" & frmMDI.IDSucursal & " AND cancelado=0 AND " & StFacturado & " AND (DESTINO=2 OR Destino=3) " & _
          "GROUP BY IDSucursal"
    
    RsEmp.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    ImpIntereses = 0: ImpAlmacenaje = 0: ImpSeguro = 0: ImpMoratorios = 0: ImpPerdida = 0: ImpOtros = 0: EmpImporteIva = 0
    If Not RsEmp.EOF Then
        
        ImpIntereses = RsEmp!ImpIntereses + RsEmp!ImpAlmacenaje + RsEmp!ImpSeguro + RsEmp!ImpMoratorios + RsEmp!ImpOtros
''''''' ImpAlmacenaje = RsEmp!ImpAlmacenaje
''''''' ImpSeguro = RsEmp!ImpSeguro
''''''' ImpMoratorios = RsEmp!ImpMoratorios
''''''' ImpPerdida = RsEmp!ImpPerdida
''''''' ImpOtros = RsEmp!ImpOtros
        EmpImporteIva = RsEmp!ImpIVA
        'Do While Not RsEmp.EOF
        'Loop
    End If
    RsEmp.Close
    Set RsEmp = Nothing
    
    With grdFactura
        If ImpIntereses <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "INT": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            '.CellText(.Rows, 4) = "INTERESES"
            .CellText(.Rows, 4) = "INTEXT11/INTERES 11% " & Format(FechaIni, "YYYY/MM/DD") & "-" & Format(FechaFin, "YYYY/MM/DD") & "/RE1X-08919 RE1X-08936"
            .CellText(.Rows, 5) = Format(ImpIntereses, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpIntereses, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpAlmacenaje <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "ALM": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "ALMACENAJE"
            .CellText(.Rows, 5) = Format(ImpAlmacenaje, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpAlmacenaje, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpSeguro <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "SEG": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "SEGURO"
            .CellText(.Rows, 5) = Format(ImpSeguro, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpSeguro, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
        If ImpMoratorios <> 0 Then
            .AddRow
            .CellText(.Rows, 1) = "1": .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 2) = "NA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 3) = "MOR": .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 4) = "MORATORIOS / RECARGOS"
            .CellText(.Rows, 5) = Format(ImpMoratorios, FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellText(.Rows, 6) = Format(ImpMoratorios, FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
        End If
    End With
    
    Subtotal = ImpIntereses + ImpAlmacenaje + ImpSeguro + ImpMoratorios
    
    
    
    '*** OBTENER LAS VENTAS MOSTRADOR (DESGLOCE POR ARTICULO) ***
    
    IvaFacturacionVentas = IIf(Regresa_Valor_BD("IvaVentas") > 0, Regresa_Valor_BD("IvaVentas"), Regresa_Valor_BD("IvaFacturacionVentas")) / 100
    VtaMosImporteIva = 0
    
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Descuento,1 AS Cantidad,v.Pagado,v.Facturado, " & _
                 "dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta," & _
                 "Round (((dv.Precio -(dv.Precio*(v.Descuento/100))) - dv.Costo) / (1 + " & IvaFacturacionVentas & ") ,5) AS Intereses , " & _
                 "0 AS Almacenaje, " & "0 AS Seguro, " & "0 AS Moratorios, " & "0 AS GtosVenta, " & _
                 "Round (((dv.Precio -(dv.Precio*(v.Descuento/100))) - dv.Costo) - ( ((dv.Precio -(dv.Precio*(v.Descuento/100))) - dv.Costo) / (1 + " & IvaFacturacionVentas & ") ),5) AS ImporteIva," & _
                 "v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE Date(v.fecha) >= date('" & Format(FechaIni, "YYYY/MM/DD") & "') and Date(v.fecha) <= date('" & Format(FechaFin, "YYYY/MM/DD") & "') AND " & _
          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND " & StFacturado & " AND v.apartado=0 AND v.TipoVenta =" & VENTAMOSTRADOR & " "
    
    Clipboard.Clear
    Clipboard.SetText Sql
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsVta.EOF Then
        Do While Not RsVta.EOF
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio & "-" & RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
                .CellText(.Rows, 5) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
            End With
            Subtotal = Subtotal + ((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta) * RsVta!Cantidad)
            VtaMosImporteIva = VtaMosImporteIva + (RsVta!ImporteIva)
            RsVta.MoveNext
        Loop
    End If
    RsVta.Close
    Set RsVta = Nothing
    
    
    '*** OBTENER LAS VENTAS POR APARTADO PAGADAS EN LA FECHA (DESGLOCE POR ARTICULO)
    VtaAparImporteIva = 0
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.FechaMovimiento,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.ImporteIva, v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE date(v.fechamovimiento) = date('" & Format(FechaIni, "YYYY/MM/DD") & "') AND date(v.fechamovimiento) = date('" & Format(FechaFin, "YYYY/MM/DD") & "') AND " & _
          "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND " & StFacturado & " AND v.TipoVenta =" & VENTAMOSTRADOR & " AND v.pagado=1 AND v.apartado=1  "
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsVta.EOF Then
        Do While Not RsVta.EOF
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio & "-" & RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
                .CellText(.Rows, 5) = Format(((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format(((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
            End With
            Subtotal = Subtotal + (((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas)) * RsVta!Cantidad)
            VtaAparImporteIva = VtaAparImporteIva + (((RsVta!PrecioVta - RsVta!Costo) - ((RsVta!PrecioVta - RsVta!Costo) / (1 + IvaFacturacionVentas))) * RsVta!Cantidad)
            RsVta.MoveNext
        Loop
    End If
    RsVta.Close
    Set RsVta = Nothing
    
    
    
    '*** OBTENER LAS VENTAS CLIENTE O BILLETE EN LA FECHA
    VtaBillImporteIva = 0
    Sql = "SELECT v.ID,v.Folio,v.Fecha,v.IDSucursal,v.IDCliente,v.IVA,v.Cancelado,v.Apartado,v.Pagado,v.Facturado,v.Descuento,1 AS Cantidad, " & _
                 "dv.Codigo,dv.Articulo,dv.Costo,dv.Precio, (dv.Precio -(dv.Precio*(v.Descuento/100))) AS PrecioVta, dv.Intereses, dv.Almacenaje, dv.Seguro, dv.Moratorios, dv.GtosVenta, dv.ImporteIva, v.TipoVenta " & _
          "FROM ventas AS v INNER JOIN detallesventas AS dv ON v.ID = dv.IDVenta " & _
          "WHERE date(v.fecha) = date('" & Format(FechaIni, "YYYY/MM/DD") & "') AND date(v.fecha) = date('" & Format(FechaFin, "YYYY/MM/DD") & "') AND " & _
                "v.IDSucursal=" & frmMDI.IDSucursal & " AND v.Cancelado=0 AND " & StFacturado & " and v.apartado=0 AND v.TipoVenta=" & VENTACLIENTE
    
    RsVta.Open Sql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsVta.EOF Then
        Do While Not RsVta.EOF
            With grdFactura
                .AddRow
                .CellText(.Rows, 1) = RsVta!Cantidad: .CellTextAlign(.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 2) = "PZA": .CellTextAlign(.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 3) = RsVta!Codigo: .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 4) = "VENTA " & ObtenerDescripVenta(RsVta!TipoVenta) & " " & RsVta!Folio & "-" & RsVta!Articulo & " " & ObtenerDescripcionArticulo(RsVta!Codigo)
                .CellText(.Rows, 5) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios), FMoneda): .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 6) = Format((RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios), FMoneda): .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellText(.Rows, 7) = RsVta!ID
                
            End With
            Subtotal = Subtotal + (RsVta!Intereses + RsVta!Almacenaje + RsVta!Seguro + RsVta!Moratorios + RsVta!GTOSVenta)
            VtaBillImporteIva = VtaBillImporteIva + (RsVta!ImporteIva)
            RsVta.MoveNext
        Loop
    End If
    RsVta.Close
    Set RsVta = Nothing
    
    txtSubTotal.text = Subtotal
    txtIva.text = EmpImporteIva + VtaMosImporteIva + VtaAparImporteIva + VtaBillImporteIva 'SubTotal * (Regresa_Valor_BD("IvaVentas") / 100)
    
Exit Sub

Error:
    MsgBox "Error al cargar Refrendos/Desempeños y Ventas del Día.", vbCritical, "Error"
    Limpiar
    Maneja_Error Err
End Sub
Private Function VerificarFacturaPlazo(ByVal FechaIni As Date, ByVal FechaFin As Date) As Boolean
    Dim RsFac As New ADODB.Recordset
    Dim strSql As String
    
    On Error GoTo Error
    
    VerificarFacturaPlazo = False
    strSql = "SELECT Folio,Cliente FROM facturas WHERE date(fecha) >= date('" & Format(FechaIni, "YYYY/MM/DD") & "') and date(fecha)<=date('" & Format(FechaFin, "YYYY/MM/DD") & "') AND TipoFactura = 0"
    RsFac.Open strSql, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not RsFac.EOF Then
        VerificarFacturaPlazo = True
        txtNombre.Tag = RsFac!Cliente
    End If
    RsFac.Close
    Set RsFac = Nothing
    
Exit Function
Error:
    Maneja_Error Err
    VerificarFacturaPlazo = False
    Set RsFac = Nothing
End Function
