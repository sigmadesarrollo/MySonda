VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas de mostrador"
   ClientHeight    =   9960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11820
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVentas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   11820
   Begin VB.TextBox txtNoTarjeta 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   6120
      MaxLength       =   60
      TabIndex        =   105
      Top             =   120
      Width           =   1935
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10395
      TabIndex        =   102
      Top             =   9435
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
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmVentas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   9240
      TabIndex        =   103
      Top             =   9435
      Width           =   1035
      _ExtentX        =   1826
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
      TextColor       =   4210752
      Object.ToolTipText     =   ""
      Picture         =   "frmVentas.frx":055E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   7890
      TabIndex        =   104
      Top             =   9435
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
      Picture         =   "frmVentas.frx":0AB0
   End
   Begin vbalTabStrip6.TabControl tTab 
      Height          =   8895
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   11625
      _ExtentX        =   20505
      _ExtentY        =   15690
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTrack        =   -1  'True
      CoolTabs        =   1
      Begin VB.Frame frmApartados 
         Caption         =   "APARTADOS"
         Height          =   8385
         Left            =   45
         TabIndex        =   54
         Top             =   465
         Visible         =   0   'False
         Width           =   11535
         Begin VB.TextBox txtCodigoApa 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            MaxLength       =   13
            TabIndex        =   3
            Top             =   2400
            Width           =   2445
         End
         Begin VB.ComboBox cmbVendedor2 
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
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2400
            Width           =   6480
         End
         Begin VB.TextBox txtApellidoPaterno 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   120
            MaxLength       =   60
            TabIndex        =   1
            Top             =   1320
            Width           =   3330
         End
         Begin VB.TextBox txtApellidoMaterno 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   3600
            MaxLength       =   70
            TabIndex        =   2
            Top             =   1320
            Width           =   3450
         End
         Begin VB.TextBox txtNombre 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   120
            MaxLength       =   40
            TabIndex        =   0
            Top             =   720
            Width           =   3330
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
            Left            =   4080
            TabIndex        =   123
            Top             =   600
            Width           =   990
         End
         Begin VB.TextBox txtPuntosApartados 
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
            Left            =   3840
            TabIndex        =   78
            Top             =   7080
            Width           =   2160
         End
         Begin VB.TextBox txtNumTarjeta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   12660
            MaxLength       =   12
            TabIndex        =   100
            Top             =   3375
            Visible         =   0   'False
            Width           =   1290
         End
         Begin VB.CheckBox chkTarjeta 
            Appearance      =   0  'Flat
            Caption         =   "Tarjeta Beneficio"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   11760
            TabIndex        =   99
            Top             =   2775
            Visible         =   0   'False
            Width           =   1530
         End
         Begin VB.TextBox txtIvaApa 
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
            Left            =   13530
            TabIndex        =   75
            Text            =   "0"
            Top             =   3165
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.TextBox txtDescuentoApa 
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
            Left            =   3840
            MaxLength       =   4
            TabIndex        =   76
            Text            =   "0"
            Top             =   6000
            Width           =   2160
         End
         Begin VB.TextBox txtAbonoApa 
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
            Left            =   3840
            TabIndex        =   77
            Top             =   6360
            Width           =   2160
         End
         Begin VB.TextBox txtEfecApa 
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
            Left            =   3840
            TabIndex        =   79
            Top             =   7800
            Width           =   2160
         End
         Begin VB.TextBox txtTotalApa 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1155
            Left            =   6120
            Locked          =   -1  'True
            TabIndex        =   55
            Top             =   6960
            Width           =   5175
         End
         Begin vbAcceleratorGrid6.vbalGrid grdArticulosApa 
            Height          =   2925
            Left            =   120
            TabIndex        =   56
            Top             =   2910
            Width           =   11250
            _ExtentX        =   19844
            _ExtentY        =   5159
            RowMode         =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            Begin VB.TextBox txtPrecioo 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5880
               TabIndex        =   85
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
         End
         Begin Line3D.ucLine3D ucLine3D2 
            Height          =   135
            Left            =   120
            Top             =   1920
            Width           =   11250
            _ExtentX        =   19844
            _ExtentY        =   238
         End
         Begin VB.TextBox Text6 
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
            Left            =   6120
            TabIndex        =   57
            Text            =   "0"
            Top             =   4440
            Width           =   495
         End
         Begin VB.TextBox Text7 
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
            Left            =   5880
            TabIndex        =   58
            Text            =   "0"
            Top             =   4800
            Width           =   975
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMosCliente3 
            Height          =   255
            Left            =   3585
            TabIndex        =   83
            Top             =   705
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   450
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
         Begin DevPowerFlatBttn.FlatBttn cmdMostrar 
            Height          =   255
            Index           =   1
            Left            =   2700
            TabIndex        =   90
            Top             =   2400
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   450
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
         Begin VB.Label Label48 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Vendedor"
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
            Height          =   255
            Left            =   3600
            TabIndex        =   130
            Top             =   2160
            Width           =   6480
         End
         Begin VB.Label Label43 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFC0C0&
            Caption         =   "Código"
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
            TabIndex        =   129
            Top             =   2160
            Width           =   2445
         End
         Begin VB.Label lblDireccion 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
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
            Left            =   1800
            TabIndex        =   128
            Top             =   1680
            Width           =   9405
         End
         Begin VB.Label Label3 
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
            Left            =   120
            TabIndex        =   127
            Top             =   1680
            Width           =   1635
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   16
            Left            =   3600
            TabIndex        =   126
            Top             =   1080
            Width           =   3450
         End
         Begin VB.Label Label28 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   210
            Index           =   15
            Left            =   120
            TabIndex        =   125
            Top             =   1080
            Width           =   3330
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
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
            ForeColor       =   &H00000000&
            Height          =   210
            Left            =   120
            TabIndex        =   124
            Top             =   480
            Width           =   3330
         End
         Begin VB.Label lblTotalPagarApartados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "<TotalPagar>"
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
            Left            =   3840
            TabIndex        =   116
            Top             =   7440
            Width           =   2160
         End
         Begin VB.Label lblTotalAPagarApartados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total a Pagar:"
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
            Left            =   1920
            TabIndex        =   115
            Top             =   7440
            Width           =   1815
         End
         Begin VB.Label lblPuntosA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puntos a Utilizar:"
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
            Left            =   1920
            TabIndex        =   114
            Top             =   7080
            Width           =   1815
         End
         Begin VB.Label lblTotalPuntosApartados2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Puntos:"
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
            Left            =   1920
            TabIndex        =   113
            Top             =   6720
            Width           =   1815
         End
         Begin VB.Label lblTotalPuntosApartados 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "<PuntosApartados>"
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
            Left            =   3840
            TabIndex        =   112
            Top             =   6720
            Width           =   2160
         End
         Begin VB.Label lblNumTarjeta 
            Caption         =   "Num.Tarjeta:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   11520
            TabIndex        =   101
            Top             =   2910
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label47 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   98
            Top             =   6030
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Descuento:"
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
            Left            =   1920
            TabIndex        =   97
            Top             =   6030
            Width           =   1815
         End
         Begin VB.Label Label13 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   12630
            TabIndex        =   92
            Top             =   3180
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label Label30 
            AutoSize        =   -1  'True
            Caption         =   "Descuento:"
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
            Left            =   240
            TabIndex        =   89
            Top             =   4320
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Iva:"
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
            Left            =   11655
            TabIndex        =   88
            Top             =   1125
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label lblLeyenda2 
            AutoSize        =   -1  'True
            Caption         =   "Cambio:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   135
            TabIndex        =   87
            Top             =   6000
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label lblCambio2 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   135
            TabIndex        =   86
            Top             =   6360
            Visible         =   0   'False
            Width           =   2145
         End
         Begin VB.Label Label41 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Abono:"
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
            Left            =   1920
            TabIndex        =   74
            Top             =   6420
            Width           =   1815
         End
         Begin VB.Label Label31 
            AutoSize        =   -1  'True
            Caption         =   "Vencimiento:"
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
            Left            =   6840
            TabIndex        =   71
            Top             =   120
            Width           =   1260
         End
         Begin VB.Label lblVencimiento 
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
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   8160
            TabIndex        =   64
            Top             =   120
            Width           =   1515
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
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
            Left            =   2520
            TabIndex        =   73
            Top             =   120
            Visible         =   0   'False
            Width           =   630
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
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
            Height          =   270
            Left            =   120
            TabIndex        =   72
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label32 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal:"
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
            Left            =   6960
            TabIndex        =   70
            Top             =   4080
            Width           =   1035
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Iva:"
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
            Left            =   6840
            TabIndex        =   69
            Top             =   4440
            Width           =   420
         End
         Begin VB.Label Label34 
            AutoSize        =   -1  'True
            Caption         =   "Total:"
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
            Left            =   6960
            TabIndex        =   68
            Top             =   4800
            Width           =   630
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<SubTotal>"
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
            Left            =   8340
            TabIndex        =   67
            Top             =   4080
            Width           =   1305
         End
         Begin VB.Label Label39 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<IVA>"
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
            Left            =   8880
            TabIndex        =   66
            Top             =   4440
            Width           =   765
         End
         Begin VB.Label Label40 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
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
            Left            =   8760
            TabIndex        =   65
            Top             =   4800
            Width           =   900
         End
         Begin VB.Label lblFolioApa 
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
            Height          =   270
            Left            =   3480
            TabIndex        =   63
            Top             =   120
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label lblFechaApa 
            AutoSize        =   -1  'True
            Caption         =   "<Fecha>"
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
            Left            =   1080
            TabIndex        =   62
            Top             =   120
            Width           =   1020
         End
         Begin VB.Label Label44 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   6645
            TabIndex        =   61
            Top             =   4440
            Width           =   240
         End
         Begin VB.Label Label45 
            AutoSize        =   -1  'True
            Caption         =   "Descuento:"
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
            Left            =   4680
            TabIndex        =   60
            Top             =   4800
            Width           =   1185
         End
         Begin VB.Label Label46 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Efectivo:"
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
            Left            =   1920
            TabIndex        =   59
            Top             =   7800
            Width           =   1815
         End
      End
      Begin VB.Frame frmPagos 
         Caption         =   "PAGOS"
         Height          =   8385
         Left            =   45
         TabIndex        =   14
         Top             =   465
         Visible         =   0   'False
         Width           =   11415
         Begin VB.TextBox txtPuntosAbonos 
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
            Left            =   2925
            TabIndex        =   117
            Top             =   1320
            Width           =   1575
         End
         Begin VB.TextBox txtEfectivoApa 
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
            Left            =   2925
            TabIndex        =   33
            Top             =   5640
            Width           =   1575
         End
         Begin VB.TextBox txtNombrePago 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            Locked          =   -1  'True
            TabIndex        =   16
            Top             =   360
            Width           =   3735
         End
         Begin VB.TextBox txtPago 
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
            Left            =   2925
            TabIndex        =   32
            Top             =   3720
            Width           =   1575
         End
         Begin vbAcceleratorGrid6.vbalGrid grdAbonos 
            Height          =   2940
            Left            =   5340
            TabIndex        =   36
            Top             =   600
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   5186
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
            DefaultRowHeight=   17
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMosClave 
            Height          =   285
            Left            =   4815
            TabIndex        =   17
            Top             =   360
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   503
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
         Begin vbAcceleratorGrid6.vbalGrid grdArticulosapartados 
            Height          =   2640
            Left            =   5340
            TabIndex        =   93
            Top             =   3885
            Width           =   6015
            _ExtentX        =   10610
            _ExtentY        =   4657
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
         Begin VB.Label lblPuntosAbonosUtilizar 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puntos a Utilizar:"
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
            Left            =   0
            TabIndex        =   122
            Top             =   1320
            Width           =   1965
         End
         Begin VB.Label lblPuntosMenos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "- Puntos:"
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
            Left            =   0
            TabIndex        =   121
            Top             =   4680
            Width           =   1965
         End
         Begin VB.Label lblPuntosAbonos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "<PuntosAbonos>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2925
            TabIndex        =   120
            Top             =   4680
            Width           =   1575
         End
         Begin VB.Label lblTotalA 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total a Pagar:"
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
            Left            =   0
            TabIndex        =   119
            Top             =   5160
            Width           =   1965
         End
         Begin VB.Label lblTotalPagar 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "<TotalPagar>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2925
            TabIndex        =   118
            Top             =   5160
            Width           =   1575
         End
         Begin VB.Label Label35 
            AutoSize        =   -1  'True
            Caption         =   "Artículos apartados:"
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
            Left            =   5340
            TabIndex        =   94
            Top             =   3555
            Width           =   2130
         End
         Begin VB.Label LeyendaAbo 
            AutoSize        =   -1  'True
            Caption         =   "Su Cambio:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   120
            TabIndex        =   49
            Top             =   6180
            Visible         =   0   'False
            Width           =   1845
         End
         Begin VB.Label CambioAbo 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   2040
            TabIndex        =   48
            Top             =   6120
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Efectivo:"
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
            Left            =   1005
            TabIndex        =   47
            Top             =   5640
            Width           =   915
         End
         Begin VB.Label lblFechaPago 
            AutoSize        =   -1  'True
            Caption         =   "<Fecha>"
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
            Left            =   3480
            TabIndex        =   22
            Top             =   840
            Width           =   960
         End
         Begin VB.Label lblFolioPago 
            AutoSize        =   -1  'True
            Caption         =   "<Folio>"
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
            TabIndex        =   20
            Top             =   840
            Width           =   870
         End
         Begin VB.Label lblTotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Total>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2925
            TabIndex        =   24
            Top             =   1800
            Width           =   1575
         End
         Begin VB.Label lblAbonos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Abonos>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   315
            Left            =   2925
            TabIndex        =   26
            Top             =   2280
            Width           =   1575
         End
         Begin VB.Label lblUltimoSaldo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Saldo>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   315
            Left            =   2925
            TabIndex        =   28
            Top             =   2760
            Width           =   1575
         End
         Begin VB.Label lblFechaAbono 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Fecha>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   2925
            TabIndex        =   30
            Top             =   3240
            Width           =   1575
         End
         Begin VB.Label lblSaldo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<Saldo>"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   300
            Left            =   2925
            TabIndex        =   35
            Top             =   4200
            Width           =   1575
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Abonos realizados:"
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
            Left            =   5340
            TabIndex        =   18
            Top             =   240
            Width           =   2025
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Nuevo Saldo:"
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
            Left            =   375
            TabIndex        =   34
            Top             =   4200
            Width           =   1425
         End
         Begin VB.Label Label25 
            AutoSize        =   -1  'True
            Caption         =   "Importe Abono:"
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
            Left            =   75
            TabIndex        =   31
            Top             =   3720
            Width           =   1725
         End
         Begin VB.Label Label24 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Abono:"
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
            Left            =   330
            TabIndex        =   29
            Top             =   3240
            Width           =   1470
         End
         Begin VB.Label Label23 
            AutoSize        =   -1  'True
            Caption         =   "Ultimo Saldo:"
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
            Left            =   345
            TabIndex        =   27
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label22 
            AutoSize        =   -1  'True
            Caption         =   "Abonado:"
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
            Left            =   720
            TabIndex        =   25
            Top             =   2280
            Width           =   1050
         End
         Begin VB.Label Label21 
            AutoSize        =   -1  'True
            Caption         =   "Total Apartado:"
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
            Left            =   120
            TabIndex        =   23
            Top             =   1800
            Width           =   1680
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
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
            Left            =   2640
            TabIndex        =   21
            Top             =   840
            Width           =   690
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            Caption         =   "Folio:"
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
            Left            =   120
            TabIndex        =   19
            Top             =   840
            Width           =   600
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Nombre:"
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
            Left            =   120
            TabIndex        =   15
            Top             =   345
            Width           =   945
         End
      End
      Begin VB.Frame frmVentasMostrador 
         Caption         =   "VENTAS"
         Height          =   8385
         Left            =   45
         TabIndex        =   6
         Top             =   465
         Visible         =   0   'False
         Width           =   11535
         Begin VB.CommandButton cmdEditar2 
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
            TabIndex        =   131
            Top             =   600
            Width           =   990
         End
         Begin VB.TextBox txtCodigo 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   120
            MaxLength       =   13
            TabIndex        =   40
            Top             =   2400
            Width           =   2445
         End
         Begin VB.ComboBox cmbVendedor 
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
            Left            =   3600
            Style           =   2  'Dropdown List
            TabIndex        =   41
            Top             =   2400
            Width           =   6480
         End
         Begin VB.TextBox txtApellidoMaterno2 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   3600
            MaxLength       =   70
            TabIndex        =   39
            Top             =   1320
            Width           =   3330
         End
         Begin VB.TextBox txtNombre2 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   120
            MaxLength       =   40
            TabIndex        =   37
            Top             =   720
            Width           =   3330
         End
         Begin VB.TextBox txtApellidoPaterno2 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   120
            MaxLength       =   100
            TabIndex        =   38
            Top             =   1320
            Width           =   3330
         End
         Begin VB.TextBox txtPuntosVentas 
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
            Left            =   3840
            TabIndex        =   43
            Top             =   6360
            Width           =   1935
         End
         Begin VB.TextBox txtTotal 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   48
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   1155
            Left            =   6195
            Locked          =   -1  'True
            TabIndex        =   50
            Top             =   6240
            Width           =   5175
         End
         Begin VB.TextBox txtEfectivo 
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
            Left            =   3840
            TabIndex        =   44
            Top             =   6720
            Width           =   1935
         End
         Begin VB.TextBox txtDescuento 
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
            TabIndex        =   42
            Text            =   "0"
            Top             =   6030
            Width           =   855
         End
         Begin VB.TextBox txtIva 
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
            Left            =   9600
            TabIndex        =   12
            Text            =   "0"
            Top             =   360
            Visible         =   0   'False
            Width           =   855
         End
         Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
            Height          =   2925
            Left            =   120
            TabIndex        =   11
            Top             =   2910
            Width           =   11250
            _ExtentX        =   19844
            _ExtentY        =   5159
            RowMode         =   -1  'True
            BackgroundPictureHeight=   0
            BackgroundPictureWidth=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
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
            Begin VB.TextBox txtPrecio 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   6360
               TabIndex        =   84
               Top             =   0
               Visible         =   0   'False
               Width           =   975
            End
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMosCliente2 
            Height          =   255
            Left            =   3585
            TabIndex        =   82
            Top             =   705
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   450
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
         Begin Line3D.ucLine3D ucLine3D1 
            Height          =   135
            Left            =   120
            Top             =   1920
            Width           =   11235
            _ExtentX        =   19817
            _ExtentY        =   238
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMostrar 
            Height          =   255
            Index           =   0
            Left            =   2700
            TabIndex        =   91
            Top             =   2400
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   450
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
            Height          =   195
            Index           =   2
            Left            =   3600
            TabIndex        =   138
            Top             =   1080
            Width           =   3330
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "Vendedor:"
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
            Height          =   255
            Left            =   3600
            TabIndex        =   137
            Top             =   2160
            Width           =   6480
         End
         Begin VB.Label Label11 
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
            Left            =   120
            TabIndex        =   136
            Top             =   480
            Width           =   3330
         End
         Begin VB.Label Label6 
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
            Left            =   120
            TabIndex        =   135
            Top             =   1080
            Width           =   3330
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FFC0C0&
            Caption         =   "Código"
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
            TabIndex        =   134
            Top             =   2160
            Width           =   2460
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
            Left            =   1800
            TabIndex        =   133
            Top             =   1680
            Width           =   9405
         End
         Begin VB.Label Label49 
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
            Left            =   120
            TabIndex        =   132
            Top             =   1680
            Width           =   1635
         End
         Begin VB.Label lblTotalPuntosVentas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "<PuntosVentas>"
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
            Left            =   3840
            TabIndex        =   111
            Top             =   7080
            Width           =   1935
         End
         Begin VB.Label lblTotalPuntos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Total Puntos:"
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
            Left            =   1920
            TabIndex        =   110
            Top             =   7080
            Width           =   1815
         End
         Begin VB.Label lblPuntosVentas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Puntos a Utilizar:"
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
            Left            =   1920
            TabIndex        =   109
            Top             =   6360
            Width           =   1815
         End
         Begin VB.Label Label37 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   5850
            TabIndex        =   96
            Top             =   6030
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Descuento:"
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
            Left            =   3600
            TabIndex        =   95
            Top             =   6030
            Width           =   1185
         End
         Begin VB.Label Leyenda 
            AutoSize        =   -1  'True
            Caption         =   "Cambio:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   345
            Left            =   135
            TabIndex        =   81
            Top             =   6000
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.Label Cambio 
            AutoSize        =   -1  'True
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   480
            Left            =   135
            TabIndex        =   80
            Top             =   6360
            Visible         =   0   'False
            Width           =   2640
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Iva:"
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
            Left            =   9045
            TabIndex        =   53
            Top             =   360
            Visible         =   0   'False
            Width           =   420
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "SubTotal:"
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
            Left            =   120
            TabIndex        =   52
            Top             =   5400
            Visible         =   0   'False
            Width           =   1035
         End
         Begin VB.Label lblSubtotal 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "<SubTotal>"
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
            Left            =   1140
            TabIndex        =   51
            Top             =   5400
            Visible         =   0   'False
            Width           =   1305
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Efectivo:"
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
            Left            =   2880
            TabIndex        =   46
            Top             =   6720
            Width           =   915
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            Caption         =   "Descuento:"
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
            TabIndex        =   45
            Top             =   5160
            Visible         =   0   'False
            Width           =   1185
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            Caption         =   "%"
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
            Left            =   10530
            TabIndex        =   13
            Top             =   360
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "<Fecha>"
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
            Left            =   1080
            TabIndex        =   10
            Top             =   120
            Width           =   1020
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
            Height          =   270
            Left            =   6840
            TabIndex        =   8
            Top             =   120
            Visible         =   0   'False
            Width           =   915
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
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
            Height          =   270
            Left            =   120
            TabIndex        =   9
            Top             =   120
            Width           =   735
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
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
            Left            =   5880
            TabIndex        =   7
            Top             =   120
            Visible         =   0   'False
            Width           =   630
         End
      End
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
      Left            =   5040
      TabIndex        =   108
      Top             =   120
      Width           =   1065
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
      Left            =   8160
      TabIndex        =   107
      Top             =   120
      Width           =   1890
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
      Left            =   10200
      TabIndex        =   106
      Top             =   120
      Width           =   120
   End
End
Attribute VB_Name = "frmVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim Vencimiento As String, Valor As Integer, p_IDUsuario As Integer
'***Puntos***
Dim TarjetaPuntos As New ClienteFrecuente
Dim ClienteVta As clientes

Public Property Let IDUsuario(Valor As Integer)
    p_IDUsuario = Valor
End Property

Public Property Get IDUsuario() As Integer
    IDUsuario = p_IDUsuario
End Property

Private Sub cmbVendedor_GotFocus()
    Cambiar_Color True, cmbVendedor
End Sub

Private Sub cmbVendedor_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbVendedor_LostFocus()
    Cambiar_Color False, cmbVendedor
End Sub

Private Sub cmbVendedor2_GotFocus()
    Cambiar_Color True, cmbVendedor2
End Sub

Private Sub cmbVendedor2_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbVendedor2_LostFocus()
    Cambiar_Color False, cmbVendedor2
End Sub

Private Sub cmdAceptar_Click()

    Dim Efectivo As Double, Total As Double
    '***Puntos***
    Dim crPuntos As Currency
    
    Select Case tTab.SelectedTab

        Case 1

            If Validar_Grabar_Datos_Ventas Then
                
                If CDbl(txtEfectivo.text) = 0 Or Trim(txtEfectivo.text) = "" Then
                    
                    MsgBox "Favor de poner el efectivo del cliente !!", vbInformation, "Ventas de Mostrador"
                    txtEfectivo.SetFocus
                
                Else
                    
                    Efectivo = txtEfectivo.text
                    Total = txtTotal.text

                    If Efectivo >= Total Then
                        
                        Leyenda.Visible = True
                        Cambio.Visible = True
                        Cambio.Caption = Format(CCur((CCur(txtEfectivo.text) - CCur(txtTotal.text))), "##,###0.00")
                        Grabar_Datos_Venta
                        Abrir_Cajon
                        lblFolio.Caption = Regresa_Movimiento(False, "FolioVentas")
                        txtIva.text = Regresa_Valor_BD("IvaVentas")
                    
                    Else
                        
                        MsgBox "Monto insuficiente !!", vbCritical, "Ventas de Mostrador"
                        txtEfectivo.SetFocus
                    End If
                    
                End If
                
            End If

        Case 2

            If Validar_Grabar_Datos_Apartados Then
            
                Efectivo = txtEfecApa.text
                Total = txtAbonoApa.text
                
                '***Puntos***
                crPuntos = Val(lblTotalPuntosApartados.Tag)
                
                If Total <= 0 Then MsgBox "El monto del abono no puede ser menor o igual a 0 !!", vbCritical, "Apartados": txtAbonoApa.SetFocus: Exit Sub
                If Efectivo <= 0 Then MsgBox "El monto del efectivo no puede ser menor o igual a 0 !!", vbCritical, "Apartados": txtEfecApa.SetFocus: Exit Sub
            
                '***Puntos***
                If Total <= (Efectivo + crPuntos) Then
                    
                    lblLeyenda2.Visible = True
                    lblLeyenda2.Caption = "Cambio:"
                    lblCambio2.Visible = True
                    
                    '***Puntos***
                    lblCambio2.Caption = Format(CCur(CCur(txtEfecApa.text) + CCur(crPuntos) - CCur(txtAbonoApa.text)), "##,###0.00")
                    
                    Grabar_Datos_Apartado
                    Abrir_Cajon
                    lblFolioApa.Caption = Regresa_Movimiento(False, "FolioVentas")
                    txtIvaApa.text = Regresa_Valor_BD("IvaVentas")
                
                Else
                    
                    MsgBox "Monto insuficiente !!", vbCritical, "Apartados"
                    txtEfecApa.SetFocus
                
                End If
            
            End If

        Case 3

            If Validar_Grabar_Abonos Then
                
                Efectivo = txtEfectivoApa
                Total = txtPago.text
                
                '***Puntos***
                crPuntos = CDbl(lblPuntosAbonos.Tag)
                
                If Total <= 0 Then MsgBox "El monto del pago no puede ser menor o igual a 0 !!", vbCritical, "Abonos": txtPago.SetFocus: Exit Sub
                If Efectivo <= 0 Then MsgBox "El monto del efectivo no puede ser menor o igual a 0 !!", vbCritical, "Abonos": txtEfectivoApa.SetFocus: Exit Sub

                '***Puntos***
                'If Total <= Efectivo Then
                 If Total <= (Efectivo + crPuntos) Then
                    
                    Grabar_Abonos
                    LeyendaAbo.Visible = True
                    LeyendaAbo.Caption = "Cambio:"
                    CambioAbo.Visible = True
                    
                    '***Puntos***
                    CambioAbo.Caption = Format(CCur(txtEfectivoApa.text) + CCur(crPuntos) - CCur(txtPago.text), FMoneda)
                    
                    Abrir_Cajon
                    txtPago.text = ""
                    lblSaldo.Caption = ""
                    txtEfectivo.text = ""
                    lblFolioApa.Caption = Regresa_Movimiento(False, "FolioVentas")
                
                Else
                    
                    MsgBox "Efectivo insuficiente !!", vbCritical, "Abonos"
                    txtEfectivoApa.SetFocus
                
                End If
            End If

    End Select

End Sub

'Validamos las ventas
Private Function Validar_Grabar_Datos_Ventas() As Boolean
    
    Validar_Grabar_Datos_Ventas = True
    
    '-------------------------------------------------------------------------
    If ClienteVta.Valida = False Then
        MsgBox "Datos incompletos del Cliente, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Grabar_Datos_Ventas = False
        txtNombre2.SetFocus
        Exit Function
    End If
    
    If Trim(txtNombre2.text) = "" Then
        MsgBox "Introduzca el nombre del cliente !!", vbCritical, "Ventas de Mostrador"
        Validar_Grabar_Datos_Ventas = False
        txtNombre2.SetFocus
        Exit Function
    End If

    If Trim(txtApellidoPaterno2.text) = "" Then
        MsgBox "Introduzca los apellidos del cliente !!", vbCritical, "Ventas de Mostrador"
        Validar_Grabar_Datos_Ventas = False
        txtApellidoPaterno2.SetFocus
        Exit Function
    End If
    
    'si no tiene apellido
    If Trim(txtApellidoMaterno2.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Grabar_Datos_Ventas = False
        txtApellidoMaterno2.SetFocus
        Exit Function
    End If
    
    If Not ClienteVta.Valida Then
        MsgBox "Datos requeridos del Cliente incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Grabar_Datos_Ventas = False
        cmdEditar2_Click
        Exit Function
    End If
    
    '-------------------------------------------------------------------------
    
    
'    If Trim(txtNombre2.text) = "" Then
'        MsgBox "Introduzca el nombre del cliente !!", vbCritical, "Ventas de Mostrador"
'        Validar_Grabar_Datos_Ventas = False
'        txtNombre2.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtApellidos2.text) = "" Then
'        MsgBox "Introduzca los apellidos del cliente !!", vbCritical, "Ventas de Mostrador"
'        Validar_Grabar_Datos_Ventas = False
'        txtApellidos2.SetFocus
'        Exit Function
'    End If
'
    If cmbVendedor.ListIndex = -1 Then
        MsgBox "Seleccione el vendedor !!", vbCritical, "Ventas de Mostrador"
        Validar_Grabar_Datos_Ventas = False
        cmbVendedor.SetFocus
        Exit Function
    End If
    
    If grdArticulos.Rows <= 0 Then
        
        MsgBox "Favor de agregar las prendas de la venta !!", vbInformation, "Ventas de Mostrador"
        Validar_Grabar_Datos_Ventas = False
        txtCodigo.SetFocus
        Exit Function
    End If
         
End Function

'Validamos los apartados
Private Function Validar_Grabar_Datos_Apartados() As Boolean
     
     '--------------------------------------------------------------------------
     If ClienteVta.Valida = False Then
        MsgBox "Datos incompletos del Cliente, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Grabar_Datos_Apartados = False
        txtNombre.SetFocus
        Exit Function
    End If
      
    'si no tiene nombre
    If Trim(txtNombre.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Grabar_Datos_Apartados = False
        txtNombre.SetFocus
        Exit Function
    End If
  
    'si no tiene apellido
    If Trim(txtApellidoPaterno.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Grabar_Datos_Apartados = False
        txtApellidoPaterno.SetFocus
        Exit Function
    End If
    
    'si no tiene apellido
    If Trim(txtApellidoMaterno.text) = "" Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Grabar_Datos_Apartados = False
        txtApellidoMaterno.SetFocus
        Exit Function
    End If
    
    If Not ClienteVta.Valida Then
        MsgBox "Datos requeridos del Cliente incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Validar_Grabar_Datos_Apartados = False
        cmdEditar_Click
        Exit Function
    End If
    '-----------------------------------------------------------------------------
     
'    If Trim(txtNombre.text) = "" Then
'        MsgBox "Introduzca el nombre del cliente !!", vbCritical, "Apartados"
'        Validar_Grabar_Datos_Apartados = False
'        txtNombre.SetFocus
'        Exit Function
'    End If
'
'    If Trim(txtApellidos.text) = "" Then
'        MsgBox "Introduzca los apellidos del cliente !!", vbCritical, "Apartados"
'        Validar_Grabar_Datos_Apartados = False
'        txtApellidos.SetFocus
'        Exit Function
'    End If
    
    If cmbVendedor2.ListIndex = -1 Then
        MsgBox "Seleccione el vendedor !!", vbCritical, "Ventas de Apartados"
        Validar_Grabar_Datos_Apartados = False
        cmbVendedor2.SetFocus
        Exit Function
    End If
    
    If grdArticulosApa.Rows <= 0 Then
        MsgBox "Es necesario que agregue las prendas a la venta !!", vbCritical, "Apartados"
        Validar_Grabar_Datos_Apartados = False
        txtCodigoApa.SetFocus
        Exit Function
    End If

    If Trim(txtAbonoApa.text) = "" Then
        MsgBox "Introduzca el monto del abono !!", vbCritical, "Apartados"
        Validar_Grabar_Datos_Apartados = False
        txtAbonoApa.SetFocus
        Exit Function
    End If

    If Trim(txtEfecApa.text) = "" Then
        MsgBox "Introduzca el monto del efectivo !!", vbCritical, "Apartados"
        Validar_Grabar_Datos_Apartados = False
        txtEfecApa.SetFocus
        Exit Function
    End If
   
    Validar_Grabar_Datos_Apartados = True
End Function

'Validamos antes de grabar el abono
Private Function Validar_Grabar_Abonos() As Boolean
   
   Validar_Grabar_Abonos = True
   
   If Trim(txtNombrePago.text) = "" Then
      MsgBox "Seleccione el Cliente que va a Abonar !!", vbInformation, "Abonos"
      Validar_Grabar_Abonos = False
      txtNombrePago.SetFocus
      Exit Function
   End If
   
   If Trim(txtPago.text) = "" Then
      MsgBox "Introduzca el Monto del Abono !!", vbInformation, "Abonos"
      Validar_Grabar_Abonos = False
      txtPago.SetFocus
      Exit Function
   End If
   
   If Val(lblSaldo.Tag) < 0 Then
      MsgBox "El monto a pagar no puede ser mayor al saldo restante !!", vbInformation, "Abonos"
      Validar_Grabar_Abonos = False
      txtPago.SetFocus
      Exit Function
   End If
   
   If Trim(txtEfectivoApa.text) = "" Then
      MsgBox "Favor de capturar el efectivo del cliente", vbInformation, "Abonos"
      Validar_Grabar_Abonos = False
      txtEfectivoApa.SetFocus
      Exit Function
   End If

End Function

'Grabamos los datos de la venta
Private Sub Grabar_Datos_Venta()

    Dim crImporte As Double, crTotal As Double, crCosto As Double, Iva As Integer, crIva As Double, Descuento As Double, crDescuento As Double, crEfectivo As Double
    Dim IDVenta As Long, Movimiento As Long, Folio As Long, Indice As Integer, IDCliente As Long, IDVendedor As Long, Hora As String
    Dim porcentajeDescuento As Double
    
    '***Puntos***
    Dim PuntosAcumulados As Double
    
On Error GoTo Error
        
    Descuento = 0
    crDescuento = 0
    Iva = 0
    crIva = 0
    crImporte = 0
    crTotal = 0
    crCosto = 0
    IDVendedor = 0
    '***Puntos***
    PuntosAcumulados = 0
    
    If txtEfectivo.text <> "" Then crEfectivo = txtEfectivo.text Else crEfectivo = 0
    If txtDescuento.text <> "" Then Descuento = txtDescuento.text Else Descuento = 0
    Iva = Regresa_Valor_BD("IvaVentas")
    
    For Indice = 1 To grdArticulos.Rows
        crImporte = crImporte + grdArticulos.CellText(Indice, 14)
        crDescuento = crDescuento + grdArticulos.CellText(Indice, 16)
    Next Indice
    
    
    'Se Agrego
    porcentajeDescuento = (Descuento * 100) / crImporte
    
    
'    If Trim(txtNombre2.text) <> "" And Val(txtNombre2.Tag) = 0 Then
'
'        dbDatos.Execute "INSERT INTO clientes (Iniciales,Nombre,Apellido,Direccion,FecRegistro) VALUES ('" & Iniciales(Trim(txtNombre2.text), Trim(txtApellidos2.text)) & "','" & Trim(txtNombre2.text) & "','" & Trim(txtApellidos2.text) & "','" & Trim(txtDireccion2.text) & "','" & Format(Date, "YYYY/MM/DD") & "')"
'        IDCliente = SacaValor("clientes", "MAX(ID)")
'
'    ElseIf Val(txtNombre2.Tag) > 0 Then
'
'        IDCliente = Val(txtNombre2.Tag)
'    Else
'
'        IDCliente = 0
'    End If
    
    '--- MLD-MODIF.- Grabar el Cliente ---
    ClienteVta.Grabar
    IDCliente = ClienteVta.ID
    '-------------------------------------

    
    If cmbVendedor.ListIndex = -1 Or cmbVendedor.ListIndex = 0 Then
        
        IDVendedor = 0
    Else
        
        IDVendedor = cmbVendedor.ItemData(cmbVendedor.ListIndex)
    End If
    
    'Saco el Folio
    Folio = Regresa_Movimiento(False, "FolioVentas")
    Regresa_Movimiento True, "FolioVentas"
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Tomo la Hora
    Hora = Time
    
    'Grabo la Venta
    dbDatos.Execute "INSERT INTO ventas (Fecha,Folio,IVA,Descuento,Total,PC,IDCliente,IDUsuario,IDSucursal,IDUsuarioDesc,IDVendedor,TipoVenta,Efectivo,DescuentoEfectivo) VALUES ('" & _
                   Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & ConvMoneda(Iva) & "," & ConvMoneda(porcentajeDescuento) & "," & ConvMoneda(crImporte) & ",'" & _
                   NombrePc & "'," & IDCliente & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & frmVentas.IDUsuario & "," & IDVendedor & "," & VENTAMOSTRADOR & "," & ConvMoneda(crEfectivo) & "," & ConvMoneda(crDescuento) & ")"
      
    'Tomo el ID de la venta
    IDVenta = SacaValor("ventas", "MAX(ID)")
    
    'Grabo el Detalle de Ventas
    For Indice = 1 To grdArticulos.Rows
                        
        'Grabo en la tabla de Detalles Ventas
        dbDatos.Execute "INSERT INTO detallesventas (IDVenta,Codigo,Articulo,Kilates,Peso,Costo,Precio,IDArticulo,ImporteDescuento) VALUES (" & _
                       IDVenta & ",'" & grdArticulos.CellText(Indice, 1) & "','" & grdArticulos.CellText(Indice, 2) & "'," & _
                       grdArticulos.CellItemData(Indice, 3) & "," & ConvMoneda(grdArticulos.CellText(Indice, 6)) & "," & ConvMoneda(grdArticulos.CellText(Indice, 7)) & "," & ConvMoneda(grdArticulos.CellText(Indice, 14)) & "," & grdArticulos.CellText(Indice, 5) & "," & grdArticulos.CellText(Indice, 16) & ")"
                        
        dbDatos.Execute "UPDATE detallesentradainventario SET cantidad=cantidad-1,TipoSalida=" & SALIDAVENTA & " WHERE ID=" & grdArticulos.CellItemData(Indice, 1)
        dbDatos.Execute "UPDATE empeno SET Destino=" & D_VENTA & " WHERE ID=" & SacaValor("detallesentradainventario", "IDEmpeno", " WHERE ID=" & grdArticulos.CellItemData(Indice, 1))
           
        'Tomo el costo
        crCosto = crCosto + grdArticulos.CellText(Indice, 7)
   
    Next Indice
    
    'Desgloso el Iva
      crIva = crImporte * (Iva / 100)
    crImporte = crImporte - (crImporte * (porcentajeDescuento / 100))
  
    
    'Grabo el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','110101'," & ConvMoneda(crImporte + crIva) & "," & TIPO_CARGO & ",0,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabo el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','620201'," & ConvMoneda(crCosto) & "," & TIPO_CARGO & ",0,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabo el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','620450'," & ConvMoneda(crImporte) & "," & TIPO_ABONO & ",0,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

    'Grabo el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','620350'," & ConvMoneda(crCosto) & "," & TIPO_ABONO & ",0,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    If crIva > 0 Then
        
        'Grabo el abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','120150'," & ConvMoneda(crIva) & "," & TIPO_ABONO & ",0,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    End If
    
    '***Puntos***
    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(Val(txtNombre2.Tag)) Then
    
        dbDatos.Execute "UPDATE ventas SET saldopuntosanterior = " & TarjetaPuntos.CuentaFrecuente.Puntos & " WHERE ID = " & IDVenta
        
        PuntosAcumulados = TarjetaPuntos.Acumula_Puntos(Ventas, frmMDI.IDUsuario, Val(Replace(txtTotal.text, ",", "")), Folio)
        MsgBox "Puntos acumulados por la venta: " & PuntosAcumulados, vbOKOnly Or vbInformation
        
        dbDatos.Execute "UPDATE ventas SET puntosacumulados = " & PuntosAcumulados & " WHERE ID = " & IDVenta
        
        If Val(txtPuntosVentas.text) > 0 Then
                   
            'Grabo el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','905501'," & _
                ConvMoneda(Val(lblTotalPuntosVentas.Tag)) & "," & TIPO_CARGO & ",0,'Redencion Puntos Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            'Grabo el abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','110150'," & _
                ConvMoneda(Val(lblTotalPuntosVentas.Tag)) & "," & TIPO_ABONO & ",0,'Redencion Puntos Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
            TarjetaPuntos.Redimir_Puntos Ventas, Val(txtPuntosVentas.text), crImporte + crIva, frmMDI.IDUsuario, Folio
            
            dbDatos.Execute "UPDATE ventas SET descuentoxpuntos = " & ConvMoneda(Val(lblTotalPuntosVentas.Tag)) & " WHERE ID = " & IDVenta
            
            dbDatos.Execute "UPDATE ventas SET puntosusados = " & Val(txtPuntosVentas.text) & " WHERE ID = " & IDVenta
        End If
        
        dbDatos.Execute "UPDATE ventas SET SaldoPuntosActual = " & TarjetaPuntos.CuentaFrecuente.Puntos - Val(txtPuntosVentas.text) & ",IDTarjeta=" & TarjetaPuntos.CuentaFrecuente.IDCuenta & " WHERE ID = " & IDVenta
        
    End If
    
    If MsgBox("Desea imprimir recibo ??", vbQuestion + vbYesNo + vbDefaultButton1, "Ventas de mostrador") = vbYes Then
        Imprimir_Recibo_Venta Folio
    End If
   
    Limpiar_Ventas
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub Grabar_Datos_Apartado()

    Dim rcClienteID As ADODB.Recordset
    Dim crTotal As Double, crCosto As Double, crImporte As Double, crAbono As Double, Iva As Double, crIva As Double, Descuento As Double, crDescuento As Double, crEfectivo As Double
    Dim IDVenta As Long, Folio As Long, Movimiento As Long, Indice As Integer, IDCliente As Long, IDVendedor  As Long, FechaHora As String, Hora As String
    Dim porcentajeDescuento As Double
    '***Puntos***
    Dim crPuntos As Currency, PuntosAcumulados As Double
    
On Error GoTo Error
    
    crAbono = 0
    Descuento = 0
    crDescuento = 0
    Iva = 0
    crIva = 0
    crImporte = 0
    crTotal = 0
    crCosto = 0
    IDVendedor = 0
    
    '***Puntos***
    crPuntos = Val(lblTotalPuntosApartados.Tag)
    
    If Trim(txtEfecApa.text) <> "" Then crEfectivo = txtEfecApa.text Else crEfectivo = 0
    If Trim(txtDescuentoApa.text) <> "" Then Descuento = txtDescuentoApa.text Else Descuento = 0
    Iva = Regresa_Valor_BD("IvaVentas")
    If Trim(txtAbonoApa.text) <> "" Then crAbono = txtAbonoApa.text Else crAbono = 0
    
    For Indice = 1 To grdArticulosApa.Rows
        crImporte = crImporte + grdArticulosApa.CellText(Indice, 14)
        crCosto = crCosto + grdArticulosApa.CellText(Indice, 7)
        crDescuento = crDescuento + grdArticulosApa.CellText(Indice, 16)
    Next Indice
    
     porcentajeDescuento = (Descuento * 100) / crImporte
    
    '--- MLD-MODIF.- Grabar el Cliente ---
    ClienteVta.Grabar
    IDCliente = ClienteVta.ID
    '-------------------------------------
    
'    If Trim(txtNombre.text) <> "" And Val(txtNombre.Tag) = 0 Then
'
'        dbDatos.Execute "INSERT INTO clientes (Iniciales,Nombre,Apellido,Direccion,FecRegistro) VALUES ('" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "' ,'" & Trim(txtNombre.text) & "','" & Trim(txtApellidos.text) & "','" & Trim(txtDireccion.text) & "','" & Format(Date, "YYYY/MM/DD") & "')"
'        Set rcClienteID = dbDatos.Execute("SELECT MAX(ID) AS IDD FROM clientes")
'        IDCliente = rcClienteID!idd
'
'    ElseIf Val(txtNombre.Tag) > 0 Then
'
'        IDCliente = Val(txtNombre.Tag)
'
'    Else
'
'        IDCliente = 0
'    End If
    
    If cmbVendedor2.ListIndex = -1 Or cmbVendedor2.ListIndex = 0 Then
        
        IDVendedor = 0
    Else
        
        IDVendedor = cmbVendedor2.ItemData(cmbVendedor2.ListIndex)
    End If
    
    'Saco el Folio
    Folio = Regresa_Movimiento(False, "FolioVentas")
    Regresa_Movimiento True, "FolioVentas"
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Tomo la Fecha y la hora
    FechaHora = Now
    
    'Grabo la Venta
    dbDatos.Execute "INSERT INTO ventas(Fecha,Vencimiento,Folio,IVA,Descuento,Total,Apartado,PC,IDCliente,IDUsuario,IDSucursal,IDUsuarioDesc,IDVendedor,TipoVenta,Efectivo,DescuentoEfectivo) VALUES ('" & _
                      Format(FechaHora, "YYYY/MM/DD HH:MM:SS") & "','" & Format(lblVencimiento.Caption, "YYYY/MM/DD") & "'," & Folio & "," & ConvMoneda(Iva) & "," & ConvMoneda(porcentajeDescuento) & "," & ConvMoneda(crImporte) & ",1,'" & NombrePc & "'," & IDCliente & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & frmVentas.IDUsuario & "," & IDVendedor & "," & VENTAMOSTRADOR & "," & ConvMoneda(crEfectivo) & "," & ConvMoneda(crDescuento) & ")"
    
    'Tomo el ID de la venta
    IDVenta = SacaValor("ventas", "MAX(ID)")
    
    For Indice = 1 To grdArticulosApa.Rows
                     
        dbDatos.Execute "INSERT INTO detallesventas (IDVenta,Codigo,Articulo,Kilates,Peso,Costo,Precio,IDArticulo,Intereses,Almacenaje,Seguro,ImporteDescuento) VALUES (" & _
                       IDVenta & ",'" & grdArticulosApa.CellText(Indice, 1) & "','" & grdArticulosApa.CellText(Indice, 2) & "'," & _
                       grdArticulosApa.CellItemData(Indice, 3) & "," & ConvMoneda(grdArticulosApa.CellText(Indice, 6)) & "," & ConvMoneda(grdArticulosApa.CellText(Indice, 7)) & "," & ConvMoneda(grdArticulosApa.CellText(Indice, 14)) & "," & grdArticulosApa.CellText(Indice, 5) & ",0,0,0," & ConvMoneda(grdArticulosApa.CellText(Indice, 16)) & ")"
                        
        dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=Cantidad-1,TipoSalida=" & SALIDAVENTA & " WHERE ID=" & grdArticulosApa.CellItemData(Indice, 1)
        dbDatos.Execute "UPDATE empeno SET Destino=" & D_VENTA & " WHERE ID=" & SacaValor("detallesentradainventario", "IDEmpeno", " WHERE ID=" & grdArticulosApa.CellItemData(Indice, 1))
   
    Next Indice
        
    'Grabo el abono
    dbDatos.Execute "INSERT INTO abonos (IDVenta,Fecha,Movimiento,Importe,PC,IDUsuario,IDSucursal) VALUES (" & _
                    IDVenta & ",'" & Format(FechaHora, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & ConvMoneda(crAbono) & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

    'Desgloso el Iva
     crIva = crImporte * (Iva / 100)
    crImporte = crImporte - (crImporte * (porcentajeDescuento / 100))
   
    
    'Grabo el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(FechaHora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'AP03','620501'," & ConvMoneda((crImporte + crIva) - crAbono) & "," & TIPO_CARGO & ",1,'Apartado','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    'Grabo el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(FechaHora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'AP03','620201'," & ConvMoneda(crCosto) & "," & TIPO_CARGO & ",1,'Apartado','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                                                  
    'Grabo el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(FechaHora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'AP03','620450'," & ConvMoneda(crImporte) & "," & TIPO_ABONO & ",1,'Apartado','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabo el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(FechaHora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'AP03','620350'," & ConvMoneda(crCosto) & "," & TIPO_ABONO & ",1,'Apartado','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    If crIva > 0 Then
        
        'Grabo el abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(FechaHora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'AP03','120150'," & ConvMoneda(crIva) & "," & TIPO_ABONO & ",1,'Apartado','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    End If
    
    'Grabo el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(FechaHora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'AP03','110101'," & ConvMoneda(crAbono) & "," & TIPO_CARGO & ",1,'Apartado','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                             
    '***Puntos***
    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(IDCliente) Then
        
        dbDatos.Execute "UPDATE ventas SET saldopuntosanterior = " & TarjetaPuntos.CuentaFrecuente.Puntos & " WHERE ID = " & IDVenta
        
        PuntosAcumulados = TarjetaPuntos.Acumula_Puntos(Apartados, frmMDI.IDUsuario, ConvMoneda(crAbono - crPuntos), Folio)
        MsgBox "Puntos Acumulados por el Apartado: " & PuntosAcumulados, vbOKOnly Or vbInformation
        
        dbDatos.Execute "UPDATE ventas SET puntosacumulados = " & PuntosAcumulados & " WHERE ID = " & IDVenta
        
        If crPuntos > 0 Then
        
            'Grabo el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(FechaHora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'AP03','905501'," & _
                        ConvMoneda(crPuntos) & "," & TIPO_CARGO & ",1,'Redencion Puntos Apartado','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
            'Grabo el abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(FechaHora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'AP03','110150'," & _
                        ConvMoneda(crPuntos) & "," & TIPO_ABONO & ",1,'Redencion Puntos Apartado','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                        
            'TarjetaPuntos.Redimir_Puntos Apartados, Val(txtPuntosApartados.text), crImporte + crIva, frmMDI.IDUsuario, Folio
            TarjetaPuntos.Redimir_Puntos Apartados, Val(txtPuntosApartados.text), CCur(crAbono), frmMDI.IDUsuario, Folio
            
            dbDatos.Execute "UPDATE ventas SET descuentoxpuntos = " & ConvMoneda(crPuntos) & " WHERE ID = " & IDVenta
            
            dbDatos.Execute "UPDATE ventas SET puntosusados = " & Val(txtPuntosApartados.text) & " WHERE ID = " & IDVenta
        End If
    
        dbDatos.Execute "UPDATE ventas SET SaldoPuntosActual = " & TarjetaPuntos.CuentaFrecuente.Puntos - Val(txtPuntosApartados.text) & ",IDTarjeta=" & TarjetaPuntos.CuentaFrecuente.IDCuenta & " WHERE ID = " & IDVenta
    
    End If
                                             
    If MsgBox("Desea imprimir recibo ??", vbQuestion + vbYesNo + vbDefaultButton2, "Apartados") = vbYes Then
        Imprimir_Recibo_Apartado Folio
    End If
   
    Limpiar_Apartados
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcClienteID = Nothing
End Sub

'Limpiamos los campos de ventas
Private Sub Limpiar_Ventas()
    lblFecha.Caption = Format(Date, "DD/MMM/YYYY")
    grdArticulos.Clear
    txtDescuento.text = "0"
    txtIva.text = "0"
    lblSubtotal.Caption = "$0.00"
    txtTotal = "0.00"
    txtEfectivo = ""
    txtCodigo.text = ""
    'MLD-MODIF.- --------
    ClienteVta.Limpiar
    txtNombre2.text = ""
    txtNombre2.Tag = ""
    txtApellidoPaterno2.text = ""
    txtApellidoMaterno2.text = ""
    lblDireccion2.Caption = ""
    '--------------------
'    txtNombre2.text = ""
'    txtNombre2.Tag = ""
'    txtApellidos2.text = ""
'    txtDireccion2.text = ""
    Leyenda.Visible = False
    Cambio.Visible = False
    cmbVendedor.ListIndex = -1
    
    '***Puntos***
    txtPuntosVentas.text = ""
    lblTotalPuntosVentas.Tag = ""
    lblTotalPuntosVentas.Caption = ""
    
    Limpiar_Tarjeta
    
End Sub

'Limpiamos los campos de ventas
Private Sub Limpiar_Apartados()
    
    lblFechaApa.Caption = Format(Date, "DD/MMM/YYYY")
    grdArticulosApa.Clear
    txtTotalApa = "0.00"
    txtCodigoApa.text = ""
    txtAbonoApa.text = ""
    'MLD-MODIF.- --------
    ClienteVta.Limpiar
    txtNombre.text = ""
    txtNombre.Tag = ""
    txtApellidoPaterno.text = ""
    txtApellidoMaterno.text = ""
    lblDireccion.Caption = ""
    '--------------------
'    txtNombre.text = ""
'    txtNombre.Tag = ""
'    txtApellidos.text = ""
'    txtDireccion.text = ""
    txtEfecApa.text = ""
    chkTarjeta.Value = 0
    txtDescuentoApa.text = "0"
    txtIvaApa.text = "0"
    lblLeyenda2.Visible = False
    lblCambio2.Visible = False
    
    '***Puntos***
    lblTotalPuntosApartados.Caption = ""
    lblTotalPuntosApartados.Tag = ""
    
    lblTotalPagarApartados.Caption = ""
    lblTotalPagarApartados.Tag = ""
    
    txtPuntosApartados.text = ""
    txtPuntosApartados.Tag = ""
    
    Limpiar_Tarjeta
End Sub

'Calculamos los totales de venta
Private Sub Calcular_Totales()

    Dim i As Integer, Total As Double, Iva As Double, Descuento As Double, IDUsuario As Integer
    '***Puntos***
    Dim Puntos As Currency, ImporteDescuento As Double
    Dim totalT As Double
    Dim porcentajeDescuento As Double

    
    Puntos = Val(lblTotalPuntosVentas.Tag)
    
    Total = 0
    Descuento = 0
    Iva = 0
totalT = 0
    If Val(txtDescuento.text) > 0 And Trim(txtDescuento.text) <> "" Then Descuento = txtDescuento.text Else Descuento = 0
    If Descuento > Val(Regresa_Valor_BD("DescuentoVentas")) Then
        
        frmPasswords.ConexSuc = 0
        frmPasswords.PrecioVitrina = 0
        frmPasswords.Cancel = 0
        frmPasswords.Ventas = 0
        frmPasswords.ModificaCorte = 0
        frmPasswords.HacerCorte = 0
        frmPasswords.InteresDesempeño = 0
        frmPasswords.InteresRefrendo = 0
        frmPasswords.ModificaPrecio = 0
        frmPasswords.RecalculoPrecios = 0
        frmPasswords.AutorizaPrestamo = 0
        frmPasswords.Vencido = 0
        frmPasswords.CancelaCierre = 0
        frmPasswords.DescuentoVentas = 1
        
        If frmPasswords.Password(GERENTE, 1) = False Then
            txtDescuento.text = "0"
            Exit Sub
        End If
        
    End If
    
    'Se agrego
    For i = 1 To grdArticulos.Rows
    totalT = totalT + grdArticulos.CellText(i, 14)
    Next i
    
    
    'Se agrego
    If Descuento > 0 And totalT > 0 Then
    porcentajeDescuento = ((Descuento * 100) / totalT)
    End If
    

    For i = 1 To grdArticulos.Rows
        Total = Total + grdArticulos.CellText(i, 14)
        'Esta es la real
'        ImporteDescuento = grdArticulos.CellText(i, 14) * Descuento / 100
 ImporteDescuento = grdArticulos.CellText(i, 14) * porcentajeDescuento / 100
        grdArticulos.CellText(i, 16) = ImporteDescuento
    Next i
    
    Iva = Regresa_Valor_BD("IvaVentas")
    'Esta es la Real
'    Total = Total - (Total * (Descuento / 100))
 
    Total = Total * (1 + (Iva / 100))
    Total = Total - Descuento
    
    'txtTotal.text = Format(Redondeo(CCur(Total)), FMoneda)
    
    '***Puntos***
    txtTotal.text = Format(Redondeo(CCur(Total - Puntos)), FMoneda)
    
End Sub

Private Sub Calcular_Totales_Apa(Optional Autorizado As Boolean = False)

    Dim i As Integer, Total As Double, Iva As Double, Descuento As Double, Enganche As Double, IDUsuario As Integer
    Dim ImporteDescuento As Double
    Dim totalT As Double
    Dim porcentajeDescuento As Double
    totalT = 0
    porcentajeDescuento = 0
    Total = 0
    If Val(txtDescuentoApa.text) > 0 And Trim(txtDescuentoApa.text) <> "" Then Descuento = txtDescuentoApa.text Else Descuento = 0
    'Se agrego
     For i = 1 To grdArticulosApa.Rows
    totalT = totalT + grdArticulosApa.CellText(i, 14)
    Next i
    
    'Se agrego
    If Descuento > 0 And totalT > 0 Then
    porcentajeDescuento = ((Descuento * 100) / totalT)
    End If
    
    
    For i = 1 To grdArticulosApa.Rows
        Total = Total + grdArticulosApa.CellText(i, 14)
        'se cambio por el de abajo
'        ImporteDescuento = grdArticulosApa.CellText(i, 14) * Descuento / 100
         ImporteDescuento = grdArticulosApa.CellText(i, 14) * porcentajeDescuento / 100
        grdArticulosApa.CellText(i, 16) = ImporteDescuento
    Next i
    
    'Se comento
    
'    If Val(txtDescuentoApa.text) > 0 Or Trim(txtDescuentoApa.text) <> "" Then Descuento = txtDescuentoApa.text
    If Val(txtDescuentoApa.text) > Val(Regresa_Valor_BD("DescuentoVentas")) Then
        frmPasswords.ConexSuc = 0
        frmPasswords.PrecioVitrina = 0
        frmPasswords.Cancel = 0
        frmPasswords.Ventas = 0
        frmPasswords.ModificaCorte = 0
        frmPasswords.HacerCorte = 0
        frmPasswords.InteresDesempeño = 0
        frmPasswords.InteresRefrendo = 0
        frmPasswords.ModificaPrecio = 0
        frmPasswords.RecalculoPrecios = 0
        frmPasswords.AutorizaPrestamo = 0
        frmPasswords.CancelaCierre = 0
        frmPasswords.DescuentoVentas = 1

        If frmPasswords.Password(GERENTE, 1) = False Then
            txtDescuentoApa.text = "0"
            Descuento = 0
            Exit Sub
        End If
    End If
            
    Iva = Regresa_Valor_BD("IvaVentas")
    'se agrego
'    Total = Total - (Total * (Descuento / 100))
Total = Total + (Total * (Iva / 100))
'Se comento
' Total = Total - (Total * (porcentajeDescuento / 100))
 Total = Total - Descuento
    
    Enganche = Regresa_Valor_BD("EngancheApartados") / 100
    txtAbonoApa.text = Format(Redondeo(Total * Enganche), FMoneda)
    txtTotalApa.text = Format(Redondeo(CCur(Total)), FMoneda)
End Sub

Private Sub cmdImprimir_Click()
Dim Folio As Long

    Folio = frmReimpresionrecibos.ReImprimir("ventas", "Folio", " WHERE TipoVenta=" & VENTAMOSTRADOR & " AND Folio=")
    If Folio > 0 Then
        
        If Val(SacaValor("ventas", "Apartado", " WHERE TipoVenta=" & VENTAMOSTRADOR & " AND Folio=" & Folio)) = 0 Then
            
            Imprimir_Recibo_Venta Folio
        Else
            
            Imprimir_Recibo_Apartado Folio
        End If
    
    ElseIf Folio = 0 Then
        
        MsgBox "No se encontró el folio especificado !!", vbInformation, "Ventas de mostrador"
    End If
End Sub

Private Sub cmdMosClave_Click()
    frmMostrarClientesVentas.Ver Me, txtNombrePago
    LeyendaAbo.Visible = False
    CambioAbo.Visible = False
    LeyendaAbo.Caption = ""
    CambioAbo.Caption = ""
    txtPago.text = ""
    txtEfectivoApa = ""
    lblSaldo.Caption = ""
End Sub

Private Sub cmdMosCliente2_Click()
    'frmMostrarclienteventas.Ver Me, txtNombre2, True, 0
    'frmMostrarclienteventas.Show 1
    frmMostrarCliente.Ver Me, txtNombre2, True, 0
End Sub

Private Sub cmdMosCliente3_Click()
    'frmMostrarclienteventas.Ver Me, txtNombre, True, 0
    'frmMostrarclienteventas.Show 1
    frmMostrarCliente.Ver Me, txtNombre, True, 0
End Sub

Private Sub cmdMostrar_Click(Index As Integer)

    Select Case Index
    Case 0
        
        frmMuestraarticulos.Ver Me, txtCodigo, True, 1
    Case 1
        
        frmMuestraarticulos.Ver Me, txtCodigoApa, True, 1, 1
    End Select
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

'Inicializamos la forma
Private Sub Inicializar()

    Screen.MousePointer = vbHourglass
    
    '-------- MLD-MODIF. ---------
    Set ClienteVta = New clientes
    ClienteVta.FechaExpiracion = "1900-01-01"
    ClienteVta.FechaNacimiento = "1900-01-01"
    ClienteVta.FechaAltaRazonSocial = "1900-01-01"
    '-----------------------------
    
    Set TarjetaPuntos.Conexion = dbDatos
    
    frmVentasMostrador.BorderStyle = 0
    frmPagos.BorderStyle = 0
    frmApartados.BorderStyle = 0
    
    Limpiar_Pago
    
    '***Puntos***
    Limpiar_Apartados
    
    Crear_Tabs
    Crear_Encabezados
    
    cmbVendedor.AddItem ""
    Cargar_Combos "CONCAT(Nombre,' ',Apellidos)", "vendedores", cmbVendedor, , , False
    cmbVendedor2.AddItem ""
    Cargar_Combos "CONCAT(Nombre,' ',Apellidos)", "vendedores", cmbVendedor2, , , False
    
    txtIva.text = Regresa_Valor_BD("IvaVentas")
    txtIvaApa.text = Regresa_Valor_BD("IvaVentas")
    
    lblFecha.Caption = Format(Date, "DD/MMM/YYYY")
    lblFechaAbono.Caption = Format(Date, "DD/MMM/YYYY")
    lblFechaApa.Caption = Format(Date, "DD/MMM/YYYY")
    
    lblVencimiento.Caption = Format(DateAdd("M", Regresa_Valor_BD("VenApartados"), Date), "DD/MMM/YYYY")
    
    lblFolio.Caption = Regresa_Movimiento(False, "FolioVentas")
    lblFolioApa.Caption = Regresa_Movimiento(False, "FolioVentas")
    
    '***Puntos***
    lblTotalPuntosVentas.Caption = ""
    
    If frmMDI.Com.PortOpen = True Then
        frmMDI.Com.PortOpen = False
    End If
    
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados del grid
Private Sub Crear_Encabezados()

    'Creamos los encabezados de los articulos
    With grdArticulos
        .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 117, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Artículo", ecgHdrTextALignLeft, , 450, , , , , , , CCLSortString
        .AddColumn "K3", "Kilates", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
        .AddColumn "K4", "Precio", ecgHdrTextALignRight, , 95, , , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "K5", "IDPrenda", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Peso", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Costo", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "K8", "Préstamo", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K9", "Interes", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K10", "Almacenaje", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K11", "Seguro", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K12", "Iva", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K13", "TipoEntrada", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K14", "Precio", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K15", "Existencia", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "K16", "Descuento", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
    End With
   
    With grdArticulosApa
        .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 117, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Artículo", ecgHdrTextALignLeft, , 450, , , , , , , CCLSortString
        .AddColumn "K3", "Kilates", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
        .AddColumn "K4", "Precio", ecgHdrTextALignRight, , 95, , , , , FMoneda, , CCLSortNumeric
            
        .AddColumn "K5", "IDPrenda", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Peso", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Costo", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "K8", "Préstamo", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K9", "Interes", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K10", "Almacenaje", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K11", "Seguro", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K12", "Iva", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K13", "TipoEntrada", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K14", "Precio", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K15", "Existencia", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "K16", "Descuento", ecgHdrTextALignRight, , 95, False, , , , FMoneda, , CCLSortNumeric
    End With
   
   'Creamos los encabezados de los detalles de los articulos
   With grdAbonos
      .AddColumn "K1", "Fecha", ecgHdrTextALignCentre, , 150, , , , , "DD/MMM/YYYY HH:MM:SS AM/PM", , CCLSortDate
      .AddColumn "K2", "Abono", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
      .AddColumn "K3", "Saldo", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
      .AddColumn "K4", "", ecgHdrTextALignLeft, , 70, , , , , FMoneda, , CCLSortNumeric
   End With
   
   'Grid de los articulos
   With grdArticulosapartados
      .AddColumn "K1", "Artículo", ecgHdrTextALignLeft, , 390, , , , , , , CCLSortString
   End With
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

'Creamos las pestañas de los tabs
Private Sub Crear_Tabs()
    
    With tTab
        .AddTab "Ventas Mostrador", , , "K1"
        .AddTab "Apartados", , , "K2"
        .AddTab "Abonos", , , "K3"
    End With

End Sub

Private Sub grdArticulos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    
    If KeyCode = vbKeyDelete Then
    
        If grdArticulos.SelectedRow > 0 Then
            
            If MsgBox("Desea eliminar la prenda seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Ventas mostrador") = vbYes Then
                
                grdArticulos.RemoveRow (grdArticulos.SelectedRow)
                Calcular_Totales
                txtCodigo.SetFocus
            
            End If
      
      End If
      
   End If

End Sub

Private Sub grdArticulos_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
'''''Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
'''''Dim sText As String, obj As Object
'''''
'''''   If lCol = 1 Or lCol = 2 Or lCol = 3 Then txtPrecio.Visible = False: Exit Sub
'''''
'''''    frmPasswords.ConexSuc = 0
'''''    frmPasswords.DescuentoVentas = 0
'''''    frmPasswords.PrecioVitrina = 0
'''''    frmPasswords.Cancel = 0
'''''    frmPasswords.Ventas = 0
'''''    frmPasswords.ModificaCorte = 0
'''''    frmPasswords.HacerCorte = 0
'''''    frmPasswords.InteresDesempeño = 0
'''''    frmPasswords.InteresRefrendo = 0
'''''    frmPasswords.AutorizaPrestamo = 0
'''''    frmPasswords.ModificaPrecio = 1
'''''
'''''    If frmPasswords.Password(GERENTE, 1) = False Then Exit Sub
'''''
'''''   txtPrecio.Visible = True
'''''   Set obj = txtPrecio
'''''
'''''   grdArticulos.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
'''''
'''''   If Not IsMissing(grdArticulos.CellText(lRow, lCol)) Then
'''''      sText = grdArticulos.CellText(lRow, lCol)
'''''   Else
'''''      sText = ""
'''''   End If
'''''
'''''    obj.Alignment = vbRightJustify
'''''
'''''    If (iKeyAscii > 13) Then
'''''       sText = Chr$(iKeyAscii) & sText
'''''       obj.text = sText
'''''       obj.SelStart = 1
'''''       obj.SelLength = Len(sText)
'''''    Else
'''''       obj.text = sText
'''''       obj.SelStart = 0
'''''       obj.SelLength = Len(sText)
'''''    End If
'''''
'''''    'Set txtPrecio.Font = grdArticulos.CellFont(lRow, lCol)
'''''    If grdArticulos.CellBackColor(lRow, lCol) = -1 Then
'''''       txtPrecio.BackColor = grdArticulos.BackColor
'''''    Else
'''''       grdArticulos.BackColor = grdArticulos.CellBackColor(lRow, lCol)
'''''    End If
'''''
'''''    obj.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
'''''
'''''   obj.Visible = True
'''''   obj.ZOrder
'''''   obj.SetFocus
End Sub

Private Sub grdArticulosApa_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If KeyCode = vbKeyDelete Then
        
        If grdArticulosApa.SelectedRow > 0 Then
            
            If MsgBox("Desea eliminar la prenda seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Apartados") = vbYes Then
                
                grdArticulosApa.RemoveRow (grdArticulosApa.SelectedRow)
                Calcular_Totales_Apa
                txtCodigoApa.SetFocus
            
            End If
        
        End If
    
    End If

End Sub

Private Sub grdArticulosApa_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
'''''Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
'''''Dim sText As String, obj As Object
'''''
'''''   If lCol = 1 Or lCol = 2 Or lCol = 3 Then txtPrecioo.Visible = False: Exit Sub
'''''
'''''    frmPasswords.ConexSuc = 0
'''''    frmPasswords.DescuentoVentas = 0
'''''    frmPasswords.PrecioVitrina = 0
'''''    frmPasswords.Cancel = 0
'''''    frmPasswords.Ventas = 0
'''''    frmPasswords.ModificaCorte = 0
'''''    frmPasswords.HacerCorte = 0
'''''    frmPasswords.InteresDesempeño = 0
'''''    frmPasswords.InteresRefrendo = 0
'''''    frmPasswords.AutorizaPrestamo = 0
'''''    frmPasswords.ModificaPrecio = 1
'''''
'''''    If frmPasswords.Password(GERENTE, 1) = False Then Exit Sub
'''''
'''''
'''''   txtPrecioo.Visible = True
'''''   Set obj = txtPrecioo
'''''
'''''   grdArticulosApa.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
'''''
'''''   If Not IsMissing(grdArticulosApa.CellText(lRow, lCol)) Then
'''''      sText = grdArticulosApa.CellText(lRow, lCol)
'''''   Else
'''''      sText = ""
'''''   End If
'''''
'''''    obj.Alignment = vbRightJustify
'''''
'''''    If (iKeyAscii > 13) Then
'''''       sText = Chr$(iKeyAscii) & sText
'''''       obj.text = sText
'''''       obj.SelStart = 1
'''''       obj.SelLength = Len(sText)
'''''    Else
'''''       obj.text = sText
'''''       obj.SelStart = 0
'''''       obj.SelLength = Len(sText)
'''''    End If
'''''
'''''    'Set txtPrecio.Font = grdArticulos.CellFont(lRow, lCol)
'''''    If grdArticulosApa.CellBackColor(lRow, lCol) = -1 Then
'''''       txtPrecioo.BackColor = grdArticulosApa.BackColor
'''''    Else
'''''       grdArticulosApa.BackColor = grdArticulosApa.CellBackColor(lRow, lCol)
'''''    End If
'''''
'''''    obj.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
'''''
'''''    obj.Visible = True
'''''    obj.ZOrder
'''''    obj.SetFocus
End Sub

Private Sub tTab_TabClick(ByVal lTab As Long)

    '***Puntos***
    lblNoTarjeta.Visible = False
    txtNoTarjeta.Visible = False
    lblPuntosAcumulados1.Visible = False
    lblPuntosAcumulados.Visible = False

    Select Case lTab

        Case 1
        
            '***Puntos***
            lblNoTarjeta.Visible = True
            txtNoTarjeta.Visible = True
            lblPuntosAcumulados1.Visible = True
            lblPuntosAcumulados.Visible = True
            
            frmVentasMostrador.Visible = True
            frmPagos.Visible = False
            frmApartados.Visible = False
            LeyendaAbo.Visible = False
            CambioAbo.Visible = False
            CambioAbo.Caption = ""
            Leyenda.Visible = False
            Cambio.Visible = False
            Cambio.Caption = ""
            lblFolio.Caption = Regresa_Movimiento(False, "FolioVentas")
            Limpiar_Ventas
            txtIva.text = Regresa_Valor_BD("IvaVentas")

        Case 2
        
            '***Puntos***
            lblNoTarjeta.Visible = True
            txtNoTarjeta.Visible = True
            lblPuntosAcumulados1.Visible = True
            lblPuntosAcumulados.Visible = True
            
            frmVentasMostrador.Visible = False
            frmPagos.Visible = False
            frmApartados.Visible = True
            LeyendaAbo.Visible = False
            CambioAbo.Visible = False
            CambioAbo.Caption = ""
            Leyenda.Visible = False
            Cambio.Visible = False
            Cambio.Caption = ""
            lblFolioApa.Caption = Regresa_Movimiento(False, "FolioVentas")
            Limpiar_Apartados
            txtIvaApa.text = Regresa_Valor_BD("IvaVentas")

        Case 3
            
            '***Puntos***
            lblNoTarjeta.Visible = True
            txtNoTarjeta.Visible = True
            lblPuntosAcumulados1.Visible = True
            lblPuntosAcumulados.Visible = True
            
            txtEfectivoApa.text = ""
            frmPagos.Visible = True
            frmVentasMostrador.Visible = False
            frmApartados.Visible = False
            Leyenda.Visible = False
            Cambio.Visible = False
            Cambio.Caption = ""
            LeyendaAbo.Visible = False
            CambioAbo.Visible = False
            CambioAbo.Caption = ""
            Limpiar_Pago
    End Select

    '***Puntos***
    Limpiar_Tarjeta

End Sub

Private Sub txtAbonoApa_Change()
    '***Puntos***
    Total_Pagar_Apartados
End Sub

Private Sub txtAbonoApa_GotFocus()
    Cambiar_Color True, txtAbonoApa
    Seleccionar_Texto txtAbonoApa
End Sub

Private Sub txtAbonoApa_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtAbonoApa.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAbonoApa_LostFocus()
Dim Total As Double, i As Integer, Enganche As Double, Abono As Double, AbonoActual As Double
    
    If Val(txtAbonoApa.text) > 0 Or (Trim(txtAbonoApa.text) <> "" And Trim(txtAbonoApa.text) <> ".") Then
        
        AbonoActual = txtAbonoApa.text
    Else
    
        AbonoActual = 0
    End If
    
    If Val(txtTotalApa.text) > 0 Or (Trim(txtTotalApa.text) <> "" And Trim(txtTotalApa.text) <> ".") Then
        
        Total = txtTotalApa.text
    Else
        
        Total = 0
    End If
    
    
    Enganche = Regresa_Valor_BD("EngancheApartados") / 100
    Abono = Total * Enganche
    
    If CCur(AbonoActual) < CCur(Abono) Then
        
        frmPasswords.ConexSuc = 0
        frmPasswords.DescuentoVentas = 0
        frmPasswords.PrecioVitrina = 0
        frmPasswords.ModificaPrecio = 0
        frmPasswords.ModificaCorte = 0
        frmPasswords.InteresDesempeño = 0
        frmPasswords.InteresRefrendo = 0
        frmPasswords.HacerCorte = 0
        frmPasswords.RecalculoPrecios = 0
        frmPasswords.AutorizaPrestamo = 0
        frmPasswords.Cancel = 0
        frmPasswords.Vencido = 0
        frmPasswords.CancelaCierre = 0
        frmPasswords.Ventas = 1
        
        If frmPasswords.Password(GERENTE, 1) Then
            
            txtAbonoApa.text = AbonoActual
        Else
            
            txtAbonoApa.text = Abono
        End If
    
    Else
        
        txtAbonoApa.text = AbonoActual
    End If
    
    txtAbonoApa.text = Format(txtAbonoApa.text, FMoneda)
    Cambiar_Color False, txtAbonoApa
End Sub


Private Sub txtCodigo_GotFocus()
    Seleccionar_Texto txtCodigo
    Cambiar_Color True, txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        
        If Len(txtCodigo.text) = 13 Then
        
            MuestraDatos2 Trim(txtCodigo.text), grdArticulos, txtCodigo, 0
        
        End If
        
    End If
    
End Sub

Private Sub txtCodigo_LostFocus()
    Cambiar_Color False, txtCodigo
End Sub

Private Sub txtCodigoApa_GotFocus()
    Seleccionar_Texto txtCodigoApa
    Cambiar_Color True, txtCodigoApa
End Sub

Private Sub txtCodigoApa_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
        
        If Len(txtCodigoApa.text) = 13 Then
            
            MuestraDatos2 Trim(txtCodigoApa.text), grdArticulosApa, txtCodigoApa, 1
        
        End If
    
    End If
    
End Sub

Private Sub txtCodigoApa_LostFocus()
    Cambiar_Color False, txtCodigoApa
End Sub

Private Sub txtDescuento_Change()
    Calcular_Totales
End Sub

Private Sub txtDescuento_GotFocus()
    Seleccionar_Texto txtDescuento
    Cambiar_Color True, txtDescuento
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescuento_LostFocus()
    Cambiar_Color False, txtDescuento
End Sub

Private Sub txtDescuentoapa_Change()
    Calcular_Totales_Apa
    
'''''Dim Total As Double, Descuento As Double, Iva As Double, i As Integer, Enganche As Double
'''''
'''''    Total = 0
'''''    For i = 1 To grdArticulosApa.Rows
'''''
'''''        Total = Total + grdArticulosApa.CellItemData(i, 4)
'''''    Next i
'''''
'''''    Descuento = Val(txtDescuentoApa.text) / 100
'''''
'''''    If Val(txtDescuentoApa.text) > 5 And CBool(chkTarjeta.Value) = False Then
'''''        frmPasswords.ConexSuc = 0
'''''        frmPasswords.PrecioVitrina = 0
'''''        frmPasswords.Cancel = 0
'''''        frmPasswords.Ventas = 0
'''''        frmPasswords.ModificaCorte = 0
'''''        frmPasswords.HacerCorte = 0
'''''        frmPasswords.InteresDesempeño = 0
'''''        frmPasswords.InteresRefrendo = 0
'''''        frmPasswords.ModificaPrecio = 0
'''''        frmPasswords.RecalculoPrecios = 0
'''''        frmPasswords.AutorizaPrestamo = 0
'''''        frmPasswords.DescuentoVentas = 1
'''''
'''''        If frmPasswords.Password(GERENTE, 1) = False Then
'''''            txtDescuentoApa.text = "0"
'''''            Descuento = 0
'''''            Exit Sub
'''''        End If
'''''    End If
'''''
'''''    Iva = Regresa_Valor_BD("IvaVentas")
'''''    Total = Total - (Total * Descuento)
'''''    Total = Total + (Total * (Iva / 100))
'''''    Enganche = Regresa_Valor_BD("EngancheApartados") / 100
'''''
'''''    txtAbonoApa.text = Format(Redondeo(Total * Enganche), FMoneda)
'''''    txtTotalApa.text = Format(Redondeo(CCur(Total)), FMoneda)
End Sub

Private Sub txtDescuentoapa_GotFocus()
    Seleccionar_Texto txtDescuentoApa
    Cambiar_Color True, txtDescuentoApa
End Sub

Private Sub txtDescuentoapa_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescuentoapa_LostFocus()
    Cambiar_Color False, txtDescuentoApa
End Sub


Private Sub txtEfecApa_GotFocus()
    Cambiar_Color True, txtEfecApa
    Seleccionar_Texto txtEfecApa
End Sub

Private Sub txtEfecApa_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtEfecApa.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEfecApa_LostFocus()
    txtEfecApa.text = Format(txtEfecApa.text, FMoneda)
    Cambiar_Color False, txtEfecApa
End Sub

Private Sub txtEfectivo_GotFocus()
    Seleccionar_Texto txtEfectivo
    Cambiar_Color True, txtEfectivo
End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtEfectivo.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEfectivo_LostFocus()
    txtEfectivo.text = Format(txtEfectivo.text, FMoneda)
    Cambiar_Color False, txtEfectivo
End Sub

Private Sub txtEfectivoApa_GotFocus()
    Seleccionar_Texto txtEfectivoApa
    Cambiar_Color True, txtEfectivoApa
End Sub

Private Sub txtEfectivoApa_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtEfectivoApa.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
End Sub

Private Sub txtEfectivoApa_LostFocus()
    txtEfectivoApa.text = Format(txtEfectivoApa.text, FMoneda)
    Cambiar_Color False, txtEfectivoApa
End Sub

Private Sub txtIva_Change()
    Calcular_Totales
End Sub

Private Sub txtIva_GotFocus()
    Seleccionar_Texto txtIva
    Cambiar_Color True, txtIva
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtIva_LostFocus()
    Cambiar_Color False, txtIva
End Sub

Private Sub txtIvaapa_GotFocus()
    Seleccionar_Texto txtIvaApa
    Cambiar_Color True, txtIvaApa
End Sub

Private Sub txtIvaapa_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtIvaapa_LostFocus()
    Cambiar_Color False, txtIvaApa
End Sub


Private Sub txtNombrePago_GotFocus()
    Seleccionar_Texto txtNombrePago
    Cambiar_Color True, txtNombrePago
End Sub

Private Sub txtNombrePago_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNombrePago_LostFocus()
    Cambiar_Color False, txtNombrePago
End Sub
'***Puntos***
Private Sub txtNoTarjeta_GotFocus()
    Seleccionar_Texto txtNoTarjeta
   Cambiar_Color True, txtNoTarjeta
End Sub
'***Puntos***
Private Sub txtNoTarjeta_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
   'Pasar_Foco KeyAscii
   If KeyAscii = vbKeyReturn Then
      If TarjetaPuntos.CuentaFrecuente.FindCuentaByFolio(txtNoTarjeta.text) Then
         lblPuntosAcumulados.Caption = TarjetaPuntos.CuentaFrecuente.Puntos
         'Buscar_Cliente_Ventas TarjetaPuntos.CuentaFrecuente.IDCliente
         Buscar TarjetaPuntos.CuentaFrecuente.IDCliente
      Else
         lblPuntosAcumulados.Caption = "0"
         Seleccionar_Texto txtNoTarjeta
         MsgBox "No se encuentra la tarjeta de cliente frecuente", vbOKOnly Or vbInformation
      End If
   End If
End Sub
'***Puntos***
Private Sub txtNoTarjeta_LostFocus()
    Cambiar_Color False, txtNoTarjeta
End Sub

Private Sub txtPago_Change()

    Dim crSaldo As Currency, crAbono As Currency
    
    '***Puntos***
    Dim crPuntos As Currency
    
    crSaldo = 0
    crAbono = 0
    
    '***Puntos***
    crPuntos = Val(lblPuntosAbonos.Tag)
    
    If Val(lblUltimoSaldo.Tag) > 0 Or Trim(lblUltimoSaldo.Tag) <> "" Then
        
        crSaldo = lblUltimoSaldo.Tag
    End If
    
    If Val(txtPago.text) > 0 Or (Trim(txtPago.text) <> "" And Trim(txtPago.text) <> "." And Trim(txtPago.text) <> "-") Then
        crAbono = txtPago.text
    End If
   
     '***Puntos***
    lblTotalPagar.Tag = crAbono - crPuntos
    lblTotalPagar.Caption = Format(crAbono - crPuntos, FMoneda)
   
    lblSaldo.Caption = Format(crSaldo - crAbono, FMoneda)
    lblSaldo.Tag = crSaldo - crAbono
   
End Sub

Private Sub txtPago_GotFocus()
    Seleccionar_Texto txtPago
    Cambiar_Color True, txtPago
End Sub

Private Sub txtPago_KeyPress(KeyAscii As Integer)
    KeyAscii = IIf(KeyAscii = 46 And InStr(1, Trim(txtPago.text), ".") > 0, 0, Solo_Numeros(KeyAscii, 1))
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPago_LostFocus()
    txtPago.text = Format(txtPago.text, FMoneda)
    Cambiar_Color False, txtPago
End Sub

'Buscamos el cliente para los pagos
Public Sub Buscar_Cliente(ID As Long)

    Dim rcCliente As New ADODB.Recordset
    Dim rcAbonos As New ADODB.Recordset
    Dim crTotal As Double, crAbonos As Double
Dim totalIva As Double

On Error GoTo Error

    rcCliente.Open "SELECT * FROM ventas WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
   
''''    If Date > DateAdd("D", Regresa_Valor_BD("DiasGraciaApa"), rcCliente!Vencimiento) Then
''''        MsgBox "El cliente ha sobrepasado su fecha límite de pago !!", vbInformation, "Abonos"
''''        rcCliente.Close
''''        Exit Sub
''''    End If

    With rcCliente
        Vencimiento = !Vencimiento
        txtNombrePago.Tag = !ID
        lblFechaPago.Caption = Format(!Fecha, "DD/MMM/YY")
        lblFolioPago.Caption = !Folio
        
        '***Puntos***
        lblFolioPago.Tag = !IDCliente
        'Se agrego
        totalIva = (!Total * (!Iva / 100))
        'Se comento
'        lblTotal.Caption = Format((!Total - (!Total * (!Descuento / 100))) * (1 + (!Iva / 100)), FMoneda)
         lblTotal.Caption = Format((!Total - (!Total * (!Descuento / 100))) + totalIva, FMoneda)
         crTotal = !Total
        
        '***Puntos***
        If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(!IDCliente) Then
            Mostrar_Datos_Puntos
        Else
            TarjetaPuntos.CuentaFrecuente.Clear
            txtNoTarjeta.text = ""
            lblPuntosAcumulados.Caption = ""
            If MsgBox("El Cliente no cuenta con tarjeta de cliente frecuente" & vbCrLf & "Desea asignarle una tarjeta?", vbYesNoCancel Or vbQuestion) = vbYes Then
                TarjetaPuntos.ShowAsignarTarjeta ID, frmMDI.IDUsuario
            End If
        End If
        
    End With
    rcCliente.Close
          
    grdAbonos.Clear
    grdAbonos.Redraw = False
    rcAbonos.Open "SELECT ID,Cancelado,Fecha,Importe FROM abonos WHERE IDVenta=" & ID & " ORDER BY Fecha", dbDatos, adOpenForwardOnly, adLockOptimistic
    With rcAbonos
        While Not .EOF
            grdAbonos.AddRow
            crAbonos = crAbonos + IIf(!Cancelado = 0, !Importe, 0)
            grdAbonos.CellText(grdAbonos.Rows, 1) = !Fecha
            grdAbonos.CellTextAlign(grdAbonos.Rows, 1) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdAbonos.CellText(grdAbonos.Rows, 2) = Format(!Importe, FMoneda)
            grdAbonos.CellTextAlign(grdAbonos.Rows, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdAbonos.CellText(grdAbonos.Rows, 3) = Format(CCur(lblTotal.Caption) - crAbonos, FMoneda)
            grdAbonos.CellTextAlign(grdAbonos.Rows, 3) = DT_RIGHT Or DT_WORD_ELLIPSIS
            If !Cancelado = 1 Then
                grdAbonos.CellText(grdAbonos.Rows, 4) = IIf(!Cancelado = 1, "CANCEL.", "")
                Colorea grdAbonos, grdAbonos.Rows, RGB(244, 119, 66)
            End If
        .MoveNext
        Wend
    End With
    rcAbonos.Close
    grdAbonos.Redraw = True
        
    grdArticulosapartados.Clear
    grdArticulosapartados.Redraw = False
    rcAbonos.Open "SELECT Articulo FROM detallesventas WHERE IDVenta=" & ID & " ORDER BY Articulo", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcAbonos.BOF And Not rcAbonos.EOF Then
        rcAbonos.MoveFirst
        While Not rcAbonos.EOF
            grdArticulosapartados.AddRow
            grdArticulosapartados.CellText(grdArticulosapartados.Rows, 1) = rcAbonos!Articulo
            grdArticulosapartados.CellTextAlign(grdArticulosapartados.Rows, 1) = DT_LEFT
        rcAbonos.MoveNext
        Wend
    End If
    rcAbonos.Close
    grdArticulosapartados.Redraw = True
    
    lblAbonos.Caption = Format(crAbonos, FMoneda)
    lblUltimoSaldo.Caption = Format(CCur(lblTotal.Caption) - crAbonos, FMoneda)
    lblUltimoSaldo.Tag = CCur(lblTotal.Caption) - crAbonos
   
    lblFechaAbono.Caption = Format(Date, "DD/MMM/YY")
    lblSaldo.Caption = Format((CCur(lblTotal.Caption) - crAbonos) - Val(txtPago.text), FMoneda)
    lblSaldo.Tag = (CCur(lblTotal.Caption) - crAbonos) - Val(txtPago.text)
    
    Set rcCliente = Nothing
    Set rcAbonos = Nothing
    Exit Sub
    
Error:
   Maneja_Error Err
   Set rcCliente = Nothing
   Set rcAbonos = Nothing
End Sub

'Grabamos los abonos
Private Sub Grabar_Abonos()

    Dim Movimiento As Long, crImporte As Double, Hora As String
    '***Puntos***
    Dim crPuntos As Currency, PuntosAcumulados As Double
    Dim IDAbono As Long
    
On Error GoTo Error
    
    crImporte = 0
    
    '***Puntos***
    crPuntos = Val(lblPuntosAbonos.Tag)
    
    If Val(txtPago.text) > 0 Or Trim(txtPago.text) <> "" Then
        crImporte = txtPago.text
    End If
                                
    If Val(lblSaldo.Tag) <= 0 Then dbDatos.Execute "UPDATE ventas SET Pagado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & Val(txtNombrePago.Tag)
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Tomo la hora
    Hora = Time
    
    'Grabo el Abono
    dbDatos.Execute "INSERT INTO abonos (IDVenta,Fecha,Movimiento,Importe,PC,IDUsuario,IDSucursal) VALUES (" & _
                                txtNombrePago.Tag & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & ConvMoneda(crImporte) & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    '***Puntos***
    'Tomo el ID de la venta
    IDAbono = SacaValor("abonos", "MAX(ID)")
    
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                              "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Val(lblFolioPago.Caption) & ",'AB05','110101'," & ConvMoneda(crImporte) & "," & TIPO_CARGO & ",0,'Abonos','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                              "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Val(lblFolioPago.Caption) & ",'AB05','620550'," & ConvMoneda(crImporte) & "," & TIPO_ABONO & ",0,'Abonos','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    '***Puntos***
    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(lblFolioPago.Tag) Then
        
        dbDatos.Execute "UPDATE abonos SET saldopuntosanterior = " & TarjetaPuntos.CuentaFrecuente.Puntos & " WHERE ID = " & IDAbono
        
        PuntosAcumulados = TarjetaPuntos.Acumula_Puntos(Abonos, frmMDI.IDUsuario, (crImporte - crPuntos), Val(lblFolioPago.Caption))
        MsgBox "Puntos Acumulados por el abono: " & PuntosAcumulados, vbOKOnly Or vbInformation
        
        dbDatos.Execute "UPDATE abonos SET puntosacumulados = " & PuntosAcumulados & " WHERE ID = " & IDAbono
        
        If crPuntos > 0 Then
            'Grabamos el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Val(lblFolioPago.Caption) & "," & _
                "'AB05','905501'," & ConvMoneda(crPuntos) & "," & TIPO_CARGO & ",0,'Redencion Puntos Abonos','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
            'Grabamos el abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Val(lblFolioPago.Caption) & "," & _
                "'AB05','110150'," & ConvMoneda(crPuntos) & "," & TIPO_ABONO & ",0,'Redencion Puntos Abonos','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
            TarjetaPuntos.Redimir_Puntos Abonos, Val(txtPuntosAbonos.text), ConvMoneda(crImporte), frmMDI.IDUsuario, Val(lblFolioPago.Caption)
                
            dbDatos.Execute "UPDATE abonos SET descuentoxpuntos = " & ConvMoneda(crPuntos) & " WHERE ID = " & IDAbono
            
            dbDatos.Execute "UPDATE abonos SET puntosusados = " & Val(txtPuntosAbonos.text) & " WHERE ID = " & IDAbono
        End If
        
        dbDatos.Execute "UPDATE abonos SET saldopuntosactual = " & TarjetaPuntos.CuentaFrecuente.Puntos - Val(txtPuntosAbonos.text) & ",IDTarjeta=" & TarjetaPuntos.CuentaFrecuente.IDCuenta & " WHERE ID = " & IDAbono
        
    End If
    
    If MsgBox("Desea imprimir recibo ??", vbQuestion + vbYesNo + vbDefaultButton1, "Abonos") = vbYes Then
        '***Puntos***
        Imprimir_Recibo_Abono Val(lblFolioPago.Caption), CDbl(lblAbonos.Caption), crImporte, CDbl(lblUltimoSaldo.Caption), IDAbono
    End If
    
    Limpiar_Pago
    Exit Sub
    
Error:
   Maneja_Error Err
End Sub

'Lipiamos los campos del pago
Private Sub Limpiar_Pago()
    
    grdAbonos.Clear
    grdArticulosapartados.Clear
    txtNombrePago.text = ""
    lblFechaPago.Caption = ""
    lblFolioPago.Caption = ""
    lblTotal.Caption = ""
    lblAbonos.Caption = ""
    lblUltimoSaldo.Caption = ""
    lblFechaAbono.Caption = ""
    lblSaldo.Caption = ""
    txtEfecApa.text = ""
    
    '***Puntos***
    lblPuntosAbonos.Caption = ""
    lblPuntosAbonos.Tag = 0
    
    txtPuntosAbonos.text = ""
    txtPuntosAbonos.Tag = 0
    
    lblTotalPagar.Caption = ""
    lblTotalPagar.Tag = 0
    
    Limpiar_Tarjeta
    
End Sub

Private Sub Imprimir_Recibo_Abono(Folio As Long, Abonado As Double, Abono As Double, UltimoSaldo As Double, IDAbono As Long)

    Dim ImprDefault As Boolean
    
    '***Puntos***
    Dim rcAbonos As New ADODB.Recordset
    Dim DescuentoXPuntos As Double, SaldoPuntosAnterior As Double, PuntosUsados As Double, PuntosAcumulados As Double, SaldoPuntosActual As Double

On Error GoTo Error
    
    '***Puntos***
    'SaldosPuntos
    rcAbonos.Open "SELECT * FROM abonos WHERE ID = " & IDAbono, dbDatos, adOpenStatic, adLockOptimistic
    
        DescuentoXPuntos = rcAbonos!DescuentoXPuntos
        SaldoPuntosAnterior = rcAbonos!SaldoPuntosAnterior
        PuntosUsados = rcAbonos!PuntosUsados
        PuntosAcumulados = rcAbonos!PuntosAcumulados
        SaldoPuntosActual = rcAbonos!SaldoPuntosActual
    
    rcAbonos.Close
    Set rcAbonos = Nothing
    
    'ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
    ImprDefault = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\NotaAbono.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{ventas.Folio}=" & Folio
        
        '***Puntos***
        .Formulas(0) = "Abonado=" & ConvMoneda(Abonado) & ""
        .Formulas(1) = "AbonoPuntos=" & ConvMoneda(DescuentoXPuntos) & ""
        .Formulas(2) = "Abono=" & ConvMoneda(Abono - DescuentoXPuntos) & ""
        .Formulas(3) = "Saldo=" & ConvMoneda(UltimoSaldo - Abono) & ""
        
        .Formulas(4) = "SaldoPuntosAnterior=" & ConvMoneda(SaldoPuntosAnterior) & ""
        .Formulas(5) = "PuntosUsados=" & ConvMoneda(PuntosUsados) & ""
        .Formulas(6) = "PuntosAcumulados=" & ConvMoneda(PuntosAcumulados) & ""
        .Formulas(7) = "SaldoPuntosActual=" & ConvMoneda(SaldoPuntosActual) & ""
        
        .Formulas(8) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(9) = "Usuario='" & Trim(UCase(frmMDI.Usuario)) & "'"
        .Formulas(10) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        
        .WindowState = crptMaximized
        .Destination = crptToPrinter
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                
        .WindowTitle = "Recibo abono"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Public Sub Imprimir_Recibo_Apartado(Folio As Long)
Dim ImprDefault As Boolean
Dim crAbono As Double

On Error GoTo Error
    
    'ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
    ImprDefault = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
    
    crAbono = 0
    If Val(txtAbonoApa.text) > 0 Or Trim(txtAbonoApa.text) <> "" Then
        
        crAbono = CDbl(txtAbonoApa.text)
    End If
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\NotaApartado.rpt"
        .Formulas(0) = "Folio=" & Folio & ""
        .Formulas(1) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(2) = "Usuario='" & Trim(UCase(frmMDI.Usuario)) & "'"
        .Formulas(3) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .WindowState = crptMaximized
        .Destination = crptToPrinter
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                        
        .WindowTitle = "Nota de apartado"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Public Sub Imprimir_Recibo_Venta(Folio As Long)
Dim ImprDefault As Boolean

On Error GoTo Error
    
    'ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
    ImprDefault = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\NotaVenta.rpt"
        .SelectionFormula = "{ventas.Folio}=" & Folio & " AND {ventas.TipoVenta}=" & VENTAMOSTRADOR & ""
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .WindowState = crptMaximized
        .Destination = crptToPrinter
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
        
        .WindowTitle = "Nota de venta"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Sub MuestraDatos(ID As Long, Grid As vbalGrid, txt As TextBox, Pestana As Integer)
Dim Iva As Double, crIva As Double
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error
              
    rcConsulta.Open "SELECT d.ID,d.Codigo,d.Descripcion,d.Cantidad,d.Kilates,d.Peso,d.Costo,d.TipoEntrada,d.PrecioVitrina " _
                    & "FROM detallesentradainventario d WHERE d.ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic

    If Not rcConsulta.BOF And Not rcConsulta.EOF Then

        With Grid
                    
            If VerificaPrenda(ID, Grid) Then
            
                'Saco el Iva
                Iva = Regresa_Valor_BD("IvaVentas")
                crIva = rcConsulta!PrecioVitrina * (Iva / 100)
    
                .AddRow
                .CellText(.Rows, 1) = rcConsulta!Codigo
                .CellItemData(.Rows, 1) = rcConsulta!ID
                .CellText(.Rows, 2) = rcConsulta!Descripcion
                .CellText(.Rows, 3) = SacaKilates(rcConsulta!Kilates)
                .CellItemData(.Rows, 3) = rcConsulta!Kilates
                .CellTextAlign(.Rows, 3) = DT_CENTER
                .CellText(.Rows, 4) = Redondeo(rcConsulta!PrecioVitrina + crIva)
                .CellItemData(.Rows, 4) = rcConsulta!PrecioVitrina
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                .CellText(.Rows, 5) = rcConsulta!ID
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                .CellText(.Rows, 6) = rcConsulta!Peso
                .CellTextAlign(.Rows, 6) = DT_RIGHT
                .CellText(.Rows, 7) = rcConsulta!Costo
                .CellItemData(.Rows, 7) = Iva
                .CellTextAlign(.Rows, 7) = DT_RIGHT
                
                .CellText(.Rows, 14) = rcConsulta!PrecioVitrina
                .CellText(.Rows, 15) = rcConsulta!Cantidad
                
            Else
                
                MsgBox "No se pueden agregar más prenda de las que existen en el inventario !!", vbCritical, "Ventas"
            End If
            
        End With

    End If

    rcConsulta.Close
    Set rcConsulta = Nothing
    If Pestana = 0 Then Calcular_Totales Else Calcular_Totales_Apa
    txt.SetFocus
    Exit Sub

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub txtPrecio_GotFocus()
    Seleccionar_Texto txtPrecio
    Cambiar_Color True, txtPrecio
End Sub

Private Sub txtPrecio_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then txtPrecio.Visible = False
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
Dim i As Integer, Total As Double

    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then
        
        grdArticulos.CellText(grdArticulos.SelectedRow, 4) = IIf(Trim(txtPrecio.text) = "", 0, txtPrecio.text)
        txtPrecio.Visible = False
        
        Total = 0
        For i = 1 To grdArticulos.Rows
            
            Total = Total + grdArticulos.CellText(i, 4)
        Next i
        
        Calcular_Totales
    
    End If
    
End Sub

Private Sub txtPrecio_LostFocus()
    Cambiar_Color False, txtPrecio
End Sub

Private Sub txtPrecioo_GotFocus()
    Seleccionar_Texto txtPrecioo
    Cambiar_Color True, txtPrecioo
End Sub

Private Sub txtPrecioo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then txtPrecioo.Visible = False
End Sub

Private Sub txtPrecioo_KeyPress(KeyAscii As Integer)
Dim i As Integer, Total As Double

    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then
        
        grdArticulosApa.CellText(grdArticulosApa.SelectedRow, 4) = IIf(Trim(txtPrecioo.text) = "", 0, txtPrecioo.text)
        txtPrecioo.Visible = False
        
        Total = 0
        For i = 1 To grdArticulosApa.Rows
            
            Total = Total + grdArticulosApa.CellText(i, 4)
        Next i
        
        Calcular_Totales_Apa
        
    End If
End Sub

Private Sub txtPrecioo_LostFocus()
    Cambiar_Color False, txtPrecioo
End Sub

Function RegresaIva() As Double
Dim i As Integer, Total As Double, Descuento As Double, descuentoo As String

    For i = 1 To grdArticulos.Rows - 1
        
        Total = Total + grdArticulos.CellItemData(i, 4)
    Next i
    
    If txtDescuento.text = "" Then Descuento = 0 Else Descuento = txtDescuento.text
    RegresaIva = (CDbl(txtTotal) + Descuento) - Total
End Function

Sub MuestraDatos2(Codigo As String, Grid As vbalGrid, txt As TextBox, Pestana As Integer)
Dim Iva As Double, crIva As Double
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error
              
    rcConsulta.Open "SELECT d.ID,d.Codigo,d.Descripcion,d.Cantidad,d.Kilates,d.Peso,d.Costo,d.TipoEntrada,d.PrecioVitrina " _
                    & "FROM detallesentradainventario d WHERE d.Cantidad>0 AND d.Codigo='" & Trim(Codigo) & "'", dbDatos, adOpenForwardOnly, adLockOptimistic

    If Not rcConsulta.BOF And Not rcConsulta.EOF Then

        With Grid
            
            If VerificaPrenda(rcConsulta!ID, Grid) Then
            
                'Saco el Iva
                Iva = Regresa_Valor_BD("IvaVentas")
                crIva = rcConsulta!PrecioVitrina * (Iva / 100)
    
                .AddRow
                .CellText(.Rows, 1) = rcConsulta!Codigo
                .CellItemData(.Rows, 1) = rcConsulta!ID
                .CellText(.Rows, 2) = rcConsulta!Descripcion
                .CellText(.Rows, 3) = SacaKilates(rcConsulta!Kilates)
                .CellItemData(.Rows, 3) = rcConsulta!Kilates
                .CellTextAlign(.Rows, 3) = DT_CENTER
                .CellText(.Rows, 4) = Redondeo(rcConsulta!PrecioVitrina + crIva)
                .CellItemData(.Rows, 4) = rcConsulta!PrecioVitrina
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                .CellText(.Rows, 5) = rcConsulta!ID
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                .CellText(.Rows, 6) = rcConsulta!Peso
                .CellTextAlign(.Rows, 6) = DT_RIGHT
                .CellText(.Rows, 7) = rcConsulta!Costo
                .CellItemData(.Rows, 7) = Iva
                .CellTextAlign(.Rows, 7) = DT_RIGHT
                
                .CellText(.Rows, 14) = rcConsulta!PrecioVitrina
                .CellText(.Rows, 15) = rcConsulta!Cantidad
            
            Else
                
                MsgBox "No se pueden agregar más prenda de las que existen en el inventario !!", vbCritical, "Ventas"
            End If
            
        End With

    End If

    rcConsulta.Close
    Set rcConsulta = Nothing
    If Pestana = 0 Then Calcular_Totales Else Calcular_Totales_Apa
    txt.SetFocus
    Exit Sub

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

'''''Function MuestraDatosApa2(Codigo As String)
'''''Dim rcConsulta As New ADODB.Recordset
'''''
'''''On Error GoTo error
'''''
'''''rcConsulta.Open "select ID,Codigo,Descripcion,Kilates,PrecioVitrina,Cantidad from detallesentradainventario where Codigo='" & Trim(Codigo) & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
'''''If Not rcConsulta.BOF And Not rcConsulta.EOF And rcConsulta!Cantidad > 0 Then
'''''    With grdArticulosApa
'''''        .AddRow
'''''        .CellText(.Rows, 1) = rcConsulta!Codigo
'''''        .CellItemData(.Rows, 1) = rcConsulta!ID
'''''        .CellText(.Rows, 2) = rcConsulta!Descripcion
'''''        .CellText(.Rows, 3) = SacaKilates(rcConsulta!Kilates)
'''''        .CellItemData(.Rows, 3) = rcConsulta!Kilates
'''''        .CellTextAlign(.Rows, 3) = DT_CENTER
'''''        .CellText(.Rows, 4) = rcConsulta!PrecioVitrina
'''''        .CellTextAlign(.Rows, 4) = DT_RIGHT
'''''    End With
'''''Else
'''''    MsgBox "No se encontró el código del artículo específicado !!", vbInformation, "Apartados"
'''''    txtCodigoApa.SetFocus
'''''End If
'''''rcConsulta.Close
'''''
'''''Calcular_Totales_Apa
'''''
'''''error:
'''''    Maneja_Error Err
'''''    Set rcConsulta = Nothing
'''''End Function

Function VerificaPrenda(IDArticulo As Long, Grid As vbalGrid) As Boolean
Dim i As Integer, x As Integer, Existencia As Integer, Cantidad As Integer
    
    Cantidad = 1
    VerificaPrenda = True
    
    For i = 1 To Grid.Rows
        
        Existencia = CInt(Grid.CellText(i, 15))
        
        For x = 1 To Grid.Rows
            
            If Grid.CellItemData(x, 1) = IDArticulo Then
                
                Cantidad = Cantidad + 1
                If Existencia < Cantidad Then VerificaPrenda = False: Exit For
            End If
        Next x
        
    Next i
    
End Function

'Public Sub Buscar_Cliente_Ventas(ID As Long)
'    On Error GoTo Error
'    Dim rc As New ADODB.Recordset
'
'    rc.Open "SELECT * FROM Clientes WHERE ID=" & ID, dbDatos, adOpenKeyset, adLockOptimistic
'
'    If Not rc.EOF Then
'        Select Case tTab.SelectedTab
'            Case 1
'                txtNombre2.Tag = rc!ID
'                txtNombre2.text = rc!Nombre
'                txtDireccion2.text = IIf(IsNull(rc!Direccion), "", rc!Direccion)
'                txtApellidos2.text = rc!Apellido
'            Case 2
'                txtNombre.Tag = rc!ID
'                txtNombre.text = rc!Nombre
'                txtDireccion.text = IIf(IsNull(rc!Direccion), "", rc!Direccion)
'                txtApellidos.text = rc!Apellido
'        End Select
'
'        If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(rc!ID) Then
'            Mostrar_Datos_Puntos
'        Else
'            TarjetaPuntos.CuentaFrecuente.Clear
'            txtNoTarjeta.text = ""
'            lblPuntosAcumulados.Caption = ""
'
'
'            If SacaValor("tarjetaspuntos", "count(id)", " where activa = 1") > 0 Then
'
'                If MsgBox("El Cliente no cuenta con tarjeta de cliente frecuente" & vbCrLf & "Desea asignarle una tarjeta?", vbYesNoCancel Or vbQuestion) = vbYes Then
'                    TarjetaPuntos.ShowAsignarTarjeta rc!ID, frmMDI.IDUsuario
'                End If
'
'            End If
'
'        End If
'
'    End If
'
'    rc.Close
'Error:
'    Maneja_Error Err
'    Set rc = Nothing
'End Sub

Private Sub Limpiar_Tarjeta()
   txtNoTarjeta.text = ""
   lblPuntosAcumulados.Caption = "0"
   TarjetaPuntos.CuentaFrecuente.Clear
End Sub

'***Puntos***
Private Sub txtPuntosAbonos_GotFocus()
    Seleccionar_Texto txtPuntosAbonos
    Cambiar_Color True, txtPuntosAbonos
End Sub
'***Puntos***
Private Sub txtPuntosAbonos_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Solo_Numeros(KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
    
        If Val(txtPuntosAbonos.text) <= TarjetaPuntos.CuentaFrecuente.Puntos Then
            lblPuntosAbonos.Tag = TarjetaPuntos.GetImporte(Val(txtPuntosAbonos))
            lblPuntosAbonos.Caption = Format(lblPuntosAbonos.Tag, FMoneda)
        Else
            MsgBox "Los puntos a utilizar no pueden ser mayor al saldo", vbOKOnly Or vbCritical
            txtPuntosAbonos.text = ""
            lblPuntosAbonos.Tag = ""
            lblPuntosAbonos.Caption = ""
        End If
        
        txtPago_Change
    End If
    
    Pasar_Foco KeyAscii
End Sub
'***Puntos***
Private Sub txtPuntosAbonos_LostFocus()
    Cambiar_Color False, txtPuntosAbonos
End Sub
'***Puntos***
Private Sub txtPuntosApartados_GotFocus()
    Seleccionar_Texto txtPuntosApartados
    Cambiar_Color True, txtPuntosApartados
End Sub
'***Puntos***
Private Sub txtPuntosApartados_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        
        If TarjetaPuntos.CuentaFrecuente.Folio <> "" Then
        
            If Val(txtPuntosApartados.text) <= TarjetaPuntos.CuentaFrecuente.Puntos Then
        
                lblTotalPuntosApartados.Tag = TarjetaPuntos.GetImporte(txtPuntosApartados.text)
                lblTotalPuntosApartados.Caption = Format(lblTotalPuntosApartados.Tag, FMoneda)
                Total_Pagar_Apartados
            Else
                MsgBox "Los puntos a utilizar no pueden ser mayor al saldo", vbOKOnly Or vbCritical
                txtPuntosApartados.text = ""
                txtPuntosApartados.Tag = ""
                lblTotalPuntosApartados.Caption = ""
                Total_Pagar_Apartados
            End If
        Else
            txtPuntosApartados.text = ""
            txtPuntosApartados.Tag = ""
            lblTotalPuntosApartados.Caption = ""
            Total_Pagar_Apartados
        End If
    
    End If
    
    Pasar_Foco KeyAscii
    
End Sub
'***Puntos***
Private Sub txtPuntosApartados_LostFocus()
    Cambiar_Color False, txtPuntosApartados
End Sub
'***Puntos***
Private Sub txtPuntosVentas_GotFocus()
    Seleccionar_Texto txtPuntosVentas
    Cambiar_Color True, txtPuntosVentas
End Sub
'***Puntos***
Private Sub txtPuntosVentas_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Solo_Numeros(KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        
        If TarjetaPuntos.CuentaFrecuente.Folio <> "" Then
        
            If Val(txtPuntosVentas.text) <= TarjetaPuntos.CuentaFrecuente.Puntos Then
                lblTotalPuntosVentas.Tag = TarjetaPuntos.GetImporte(txtPuntosVentas.text)
                lblTotalPuntosVentas.Caption = Format(lblTotalPuntosVentas.Tag, FMoneda)
                Calcular_Totales
            Else
                MsgBox "Los puntos a utilizar no pueden ser mayor al saldo", vbOKOnly Or vbCritical
                txtPuntosVentas.text = ""
                txtPuntosVentas.Tag = ""
                lblTotalPuntosVentas.Caption = ""
                Calcular_Totales
            End If
                
        Else
        
            txtPuntosVentas.text = ""
            txtPuntosVentas.Tag = ""
            lblTotalPuntosVentas.Caption = ""
            Calcular_Totales
            
        End If
        
    End If
    
    Pasar_Foco KeyAscii
End Sub
'***Puntos***
Private Sub txtPuntosVentas_LostFocus()
    Cambiar_Color False, txtPuntosVentas
End Sub
'***Puntos***
Private Sub Mostrar_Datos_Puntos()
    On Error GoTo Error
    
    txtNoTarjeta.text = TarjetaPuntos.CuentaFrecuente.Folio
    lblPuntosAcumulados.Caption = TarjetaPuntos.CuentaFrecuente.Puntos
    
Error:
    Maneja_Error Err
End Sub

'***Puntos***
Private Sub Total_Pagar_Apartados()
    
On Error GoTo Error
    
    Dim crPuntos As Currency
    Dim crAbono As Currency
    
    crPuntos = Val(lblTotalPuntosApartados.Tag)
    crAbono = Val(Replace(txtAbonoApa.text, ",", ""))
    
    lblTotalPagarApartados.Tag = crAbono - crPuntos
    lblTotalPagarApartados.Caption = Format(crAbono - crPuntos, FMoneda)
    
Error:
    Maneja_Error Err
    
End Sub


Private Sub txtNombre_GotFocus()
   Seleccionar_Texto txtNombre
   Cambiar_Color True, txtNombre
   Leyenda.Caption = ""
   Cambio.Caption = ""
   Leyenda.Visible = False
   Cambio.Visible = False
   If txtTotalApa.text = "0.00" Then
      Leyenda.Visible = False
      Cambio.Caption = ""
      Cambio.Visible = False
   End If
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNombre_LostFocus()
    Cambiar_Color False, txtNombre
    ClienteVta.Nombre = txtNombre.text
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
End Sub
'------------------------------------------------------

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
    ClienteVta.ApellidoPaterno = txtApellidoPaterno.text
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
    ClienteVta.ApellidoPaterno = txtApellidoPaterno2.text
End Sub

'-------------------------------------------------------

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
    ClienteVta.ApellidoMaterno = txtApellidoMaterno.text
    If Trim(txtNombre.text) <> "" And Trim(txtApellidoPaterno.text) <> "" And Trim(txtApellidoMaterno.text) <> "" And Val(txtNombre.Tag) <= 0 Then Mostrar_Seleccionar_Cliente Trim(txtNombre.text), Trim(txtApellidoPaterno.text), Trim(txtApellidoMaterno.text), Me
    If Val(txtNombre.Tag) = 0 Then
        ClienteVta.Nombre = Trim(txtNombre.text)
        ClienteVta.ApellidoPaterno = Trim(txtApellidoPaterno.text)
        ClienteVta.ApellidoMaterno = Trim(txtApellidoMaterno.text)
        frmClientes.Mostrar ClienteVta
        lblDireccion.Caption = ClienteVta.Direccion & IIf(ClienteVta.NoExterior <> "", " #" & ClienteVta.NoExterior, "") & IIf(ClienteVta.NoInterior <> "", " INT." & ClienteVta.NoInterior, "") & " COL." & ClienteVta.Colonia & " C.P." & ClienteVta.CodigoPostal
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
    ClienteVta.ApellidoMaterno = txtApellidoMaterno2.text
    If Trim(txtNombre2.text) <> "" And Trim(txtApellidoPaterno2.text) <> "" And Trim(txtApellidoMaterno2.text) <> "" And Val(txtNombre2.Tag) <= 0 Then Mostrar_Seleccionar_Cliente Trim(txtNombre2.text), Trim(txtApellidoPaterno2.text), Trim(txtApellidoMaterno2.text), Me
    If Val(txtNombre2.Tag) = 0 Then
        ClienteVta.Nombre = Trim(txtNombre2.text)
        ClienteVta.ApellidoPaterno = Trim(txtApellidoPaterno2.text)
        ClienteVta.ApellidoMaterno = Trim(txtApellidoMaterno2.text)
        frmClientes.Mostrar ClienteVta
        lblDireccion2.Caption = ClienteVta.Direccion & IIf(ClienteVta.NoExterior <> "", " #" & ClienteVta.NoExterior, "") & IIf(ClienteVta.NoInterior <> "", " INT." & ClienteVta.NoInterior, "") & " COL." & ClienteVta.Colonia & " C.P." & ClienteVta.CodigoPostal
    End If
End Sub

'--------------------------------------------------------------------------

Private Sub cmdEditar_Click()
    frmClientes.Mostrar ClienteVta
    txtNombre.text = ClienteVta.Nombre
    txtNombre.Tag = ClienteVta.ID
    txtApellidoPaterno.text = ClienteVta.ApellidoPaterno
    txtApellidoMaterno.text = ClienteVta.ApellidoMaterno
    If ClienteVta.ID = 0 Then
        lblDireccion.Caption = ""
    Else
        lblDireccion.Caption = ClienteVta.Direccion & IIf(ClienteVta.NoExterior <> "", " #" & ClienteVta.NoExterior, "") & IIf(ClienteVta.NoInterior <> "", " INT." & ClienteVta.NoInterior, "") & " COL." & ClienteVta.Colonia & " C.P." & ClienteVta.CodigoPostal
                               
    End If
End Sub

Private Sub cmdEditar2_Click()
    frmClientes.Mostrar ClienteVta
    txtNombre2.text = ClienteVta.Nombre
    txtNombre2.Tag = ClienteVta.ID
    txtApellidoPaterno2.text = ClienteVta.ApellidoPaterno
    txtApellidoMaterno2.text = ClienteVta.ApellidoMaterno
    If ClienteVta.ID = 0 Then
        lblDireccion.Caption = ""
    Else
        lblDireccion2.Caption = ClienteVta.Direccion & IIf(ClienteVta.NoExterior <> "", " #" & ClienteVta.NoExterior, "") & IIf(ClienteVta.NoInterior <> "", " INT." & ClienteVta.NoInterior, "") & " COL." & ClienteVta.Colonia & " C.P." & ClienteVta.CodigoPostal
    End If
End Sub

'-------------------------------------------------------------------------
'MLD-MODIF.- Buscamos el id cliente
Public Sub Buscar(ID As Long)
On Error GoTo Error

    ClienteVta.Buscar ID
    
    If Not ClienteVta.Valida Then
        frmClientes.Mostrar ClienteVta
    End If
    
    Select Case tTab.SelectedTab
        Case 1
            txtNombre2.text = ClienteVta.Nombre
            txtNombre2.Tag = ClienteVta.ID
            txtApellidoPaterno2.text = ClienteVta.ApellidoPaterno
            txtApellidoMaterno2.text = ClienteVta.ApellidoMaterno
            cmdEditar2.Visible = True
            lblDireccion2.Caption = ClienteVta.Direccion & IIf(ClienteVta.NoExterior <> "", " #" & ClienteVta.NoExterior, "") & IIf(ClienteVta.NoInterior <> "", " INT." & ClienteVta.NoInterior, "") & " COL." & ClienteVta.Colonia & " C.P." & ClienteVta.CodigoPostal
        Case 2
            txtNombre.text = ClienteVta.Nombre
            txtNombre.Tag = ClienteVta.ID
            txtApellidoPaterno.text = ClienteVta.ApellidoPaterno
            txtApellidoMaterno.text = ClienteVta.ApellidoMaterno
            cmdEditar.Visible = True
            lblDireccion.Caption = ClienteVta.Direccion & IIf(ClienteVta.NoExterior <> "", " #" & ClienteVta.NoExterior, "") & IIf(ClienteVta.NoInterior <> "", " INT." & ClienteVta.NoInterior, "") & " COL." & ClienteVta.Colonia & " C.P." & ClienteVta.CodigoPostal
    End Select
    
    '***Puntos***
    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(ID) Then
        Mostrar_Datos_Puntos
    Else
        TarjetaPuntos.CuentaFrecuente.Clear
        txtNoTarjeta.text = ""
        lblPuntosAcumulados.Caption = ""

        If SacaValor("tarjetaspuntos", "count(id)", " where activa = 1") > 0 Then

            If MsgBox("El Cliente no cuenta con tarjeta de cliente frecuente" & vbCrLf & "Desea asignarle una tarjeta?", vbYesNoCancel Or vbQuestion) = vbYes Then
                TarjetaPuntos.ShowAsignarTarjeta ClienteVta.ID, frmMDI.IDUsuario
            End If

        End If

    End If
    
    
Exit Sub
    
Error:
    Maneja_Error Err
End Sub



