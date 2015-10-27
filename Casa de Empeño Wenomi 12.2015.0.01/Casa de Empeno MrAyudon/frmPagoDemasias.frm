VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmPagoDemasias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pago de demasias"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12675
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPagoDemasias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   12675
   Begin VB.Frame frmEmpeño 
      Caption         =   "Empeno"
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   12375
      Begin VB.TextBox txtNombre 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   840
         MaxLength       =   20
         TabIndex        =   12
         Top             =   480
         Width           =   3135
      End
      Begin VB.TextBox txtMunicipio 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3735
         MaxLength       =   30
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox txtColonia 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   405
         MaxLength       =   30
         TabIndex        =   10
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox txtEstado 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   735
         MaxLength       =   30
         TabIndex        =   9
         Top             =   1545
         Width           =   1935
      End
      Begin VB.TextBox txtTelefono 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3735
         MaxLength       =   20
         TabIndex        =   8
         Top             =   1545
         Width           =   1335
      End
      Begin VB.TextBox txtIdentificacion 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   6540
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1545
         Width           =   1650
      End
      Begin VB.TextBox txtBuscar 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   6
         Top             =   120
         Width           =   3375
      End
      Begin VB.TextBox txtResponsable 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1275
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1905
         Width           =   6915
      End
      Begin VB.TextBox txtCP 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   6615
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1200
         Width           =   1575
      End
      Begin VB.TextBox txtDireccion 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   975
         MaxLength       =   70
         TabIndex        =   2
         Top             =   840
         Width           =   7215
      End
      Begin VB.TextBox txtApellidos 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   5055
         MaxLength       =   60
         TabIndex        =   1
         Top             =   480
         Width           =   3135
      End
      Begin vbAcceleratorGrid6.vbalGrid grdEmpeños 
         Height          =   3315
         Left            =   0
         TabIndex        =   5
         Top             =   2280
         Width           =   12330
         _ExtentX        =   21749
         _ExtentY        =   5847
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
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   135
         Left            =   -30
         Top             =   2160
         Width           =   12330
         _ExtentX        =   21749
         _ExtentY        =   238
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
         Height          =   225
         Left            =   4095
         TabIndex        =   13
         Top             =   105
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
         Left            =   0
         Top             =   4680
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   238
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
         Left            =   4095
         TabIndex        =   25
         Top             =   450
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
         Left            =   15
         TabIndex        =   24
         Top             =   450
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
         Left            =   15
         TabIndex        =   23
         Top             =   810
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
         Left            =   2775
         TabIndex        =   22
         Top             =   1170
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
         Left            =   15
         TabIndex        =   21
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
         Left            =   0
         TabIndex        =   20
         Top             =   120
         Width           =   660
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
         Left            =   15
         TabIndex        =   19
         Top             =   1530
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
         Left            =   2775
         TabIndex        =   18
         Top             =   1530
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
         Left            =   5175
         TabIndex        =   17
         Top             =   1530
         Width           =   1305
      End
      Begin VB.Label Label88 
         AutoSize        =   -1  'True
         Caption         =   "Responsable:"
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
         Left            =   0
         TabIndex        =   16
         Top             =   1875
         Width           =   1215
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
         Left            =   6255
         TabIndex        =   15
         Top             =   1170
         Width           =   300
      End
      Begin VB.Label txtPrestamo 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   9690
         TabIndex        =   14
         Top             =   855
         Width           =   75
      End
   End
End
Attribute VB_Name = "frmPagoDemasias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
