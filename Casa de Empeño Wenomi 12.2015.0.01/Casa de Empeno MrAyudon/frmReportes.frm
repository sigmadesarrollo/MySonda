VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   8655
   ClientLeft      =   375
   ClientTop       =   840
   ClientWidth     =   12540
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   12540
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   255
      Left            =   0
      TabIndex        =   38
      Top             =   8400
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtros Extra"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   8280
      TabIndex        =   32
      Top             =   1080
      Width           =   4215
      Begin VB.CheckBox chkNoPagados 
         Appearance      =   0  'Flat
         Caption         =   "Activos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   315
         Width           =   1095
      End
      Begin VB.CheckBox chkCancelados 
         Appearance      =   0  'Flat
         Caption         =   "Cancelados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1335
         TabIndex        =   36
         Top             =   315
         Width           =   1455
      End
      Begin VB.CheckBox chkPagados 
         Appearance      =   0  'Flat
         Caption         =   "Pagados"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2880
         TabIndex        =   35
         Top             =   315
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ordenado por:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   8280
      TabIndex        =   31
      Top             =   30
      Width           =   4215
      Begin VB.OptionButton opFolio 
         Appearance      =   0  'Flat
         Caption         =   "Contrato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1560
         TabIndex        =   16
         Top             =   360
         Width           =   1365
      End
      Begin VB.OptionButton opFecha 
         Appearance      =   0  'Flat
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton opCliente 
         Appearance      =   0  'Flat
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3000
         TabIndex        =   17
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton opImporte 
         Appearance      =   0  'Flat
         Caption         =   "Importes"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   18
         Top             =   660
         Width           =   1335
      End
      Begin VB.OptionButton opFechaMovimiento 
         Appearance      =   0  'Flat
         Caption         =   "Fecha Movimiento"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   1560
         TabIndex        =   19
         Top             =   660
         Width           =   2295
      End
   End
   Begin VB.Frame FRACriterios 
      Caption         =   "Criterios del Reporte"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2835
      Left            =   15
      TabIndex        =   21
      Top             =   30
      Width           =   8250
      Begin VB.Frame Frame3 
         Caption         =   "Resumen/Detalle:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   6000
         TabIndex        =   39
         Top             =   1920
         Width           =   2055
         Begin VB.OptionButton optSi 
            Appearance      =   0  'Flat
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   41
            Top             =   330
            Value           =   -1  'True
            Width           =   615
         End
         Begin VB.OptionButton optNo 
            Appearance      =   0  'Flat
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   40
            Top             =   330
            Width           =   615
         End
      End
      Begin VB.ComboBox cmbTipoReporte 
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
         ItemData        =   "frmReportes.frx":000C
         Left            =   1560
         List            =   "frmReportes.frx":0019
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   2400
      End
      Begin VB.TextBox txtCliente 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   7
         Top             =   2085
         Width           =   3720
      End
      Begin VB.TextBox txtImportes 
         Alignment       =   1  'Right Justify
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
         Left            =   6480
         TabIndex        =   14
         Top             =   1500
         Width           =   1575
      End
      Begin VB.TextBox txtEmpleado 
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
         Left            =   1560
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   9
         Top             =   2445
         Width           =   3720
      End
      Begin VB.ComboBox cmbDestino 
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
         ItemData        =   "frmReportes.frx":0037
         Left            =   5760
         List            =   "frmReportes.frx":0044
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1080
         Width           =   2400
      End
      Begin VB.ComboBox cmbOrigen 
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
         ItemData        =   "frmReportes.frx":0069
         Left            =   5760
         List            =   "frmReportes.frx":006B
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   645
         Width           =   2400
      End
      Begin VB.TextBox txtFolioFin 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         TabIndex        =   6
         Top             =   1725
         Width           =   1935
      End
      Begin VB.TextBox txtFolioIni 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         TabIndex        =   5
         Top             =   1365
         Width           =   1935
      End
      Begin VB.ComboBox cmbTipo 
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
         ItemData        =   "frmReportes.frx":006D
         Left            =   5760
         List            =   "frmReportes.frx":006F
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   240
         Width           =   2400
      End
      Begin VB.TextBox txtFechaFin 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   1005
         Width           =   1455
      End
      Begin VB.TextBox txtFechaIni 
         Alignment       =   1  'Right Justify
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
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   645
         Width           =   1455
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosFechaFin 
         Height          =   300
         Left            =   3030
         TabIndex        =   4
         Top             =   1020
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
         Picture         =   "frmReportes.frx":0071
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosFechaIni 
         Height          =   300
         Left            =   3030
         TabIndex        =   2
         Top             =   645
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
         Picture         =   "frmReportes.frx":0186
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosEmpleado 
         Height          =   285
         Left            =   5280
         TabIndex        =   10
         Top             =   2445
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
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
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
         Height          =   285
         Left            =   5280
         TabIndex        =   8
         Top             =   2085
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
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
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reporte:"
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
         TabIndex        =   34
         Top             =   300
         Width           =   870
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente:"
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
         TabIndex        =   33
         Top             =   2085
         Width           =   735
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importes mayores a:"
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
         TabIndex        =   30
         Top             =   1500
         Width           =   2040
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Empleado:"
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
         TabIndex        =   29
         Top             =   2445
         Width           =   990
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Destino:"
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
         TabIndex        =   28
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Origen:"
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
         TabIndex        =   27
         Top             =   720
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folio Inicial:"
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
         TabIndex        =   26
         Top             =   1365
         Width           =   1110
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Folio Final:"
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
         TabIndex        =   25
         Top             =   1725
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Categoría:"
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
         TabIndex        =   24
         Top             =   300
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Final:"
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
         TabIndex        =   23
         Top             =   1005
         Width           =   1110
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha Inicial:"
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
         TabIndex        =   22
         Top             =   645
         Width           =   1245
      End
   End
   Begin vbAcceleratorGrid6.vbalGrid grdReportes 
      Height          =   5430
      Left            =   0
      TabIndex        =   20
      Top             =   2880
      Width           =   12495
      _ExtentX        =   22040
      _ExtentY        =   9578
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
   Begin vbalIml6.vbalImageList lstIcons 
      Left            =   8520
      Top             =   1920
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   5740
      Images          =   "frmReportes.frx":029B
      Version         =   131072
      KeyCount        =   5
      Keys            =   "ÿÿÿÿ"
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   11325
      TabIndex        =   42
      Top             =   1920
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
      Picture         =   "frmReportes.frx":1927
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   10125
      TabIndex        =   43
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   " &Imprimir"
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
      Picture         =   "frmReportes.frx":1E79
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   10125
      TabIndex        =   44
      Top             =   1920
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
      Picture         =   "frmReportes.frx":23CB
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   11325
      TabIndex        =   45
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "        &Limpiar"
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
      Picture         =   "frmReportes.frx":2750
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub chkNoPagados_Click()
    If chkNoPagados.Value And chkPagados.Value Then
        chkPagados.Value = 0
    End If
End Sub

Private Sub chkPagados_Click()
    If chkPagados.Value And chkNoPagados.Value Then
        chkNoPagados.Value = 0
    End If
End Sub

Private Sub cmbDestino_GotFocus()
    Cambiar_Color True, cmbDestino
End Sub

Private Sub cmbDestino_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbDestino_LostFocus()
    Cambiar_Color False, cmbDestino
End Sub

Private Sub cmbOrigen_GotFocus()
    Cambiar_Color True, cmbOrigen
End Sub

Private Sub cmbOrigen_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbOrigen_LostFocus()
    Cambiar_Color False, cmbOrigen
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

Private Sub cmbTipoReporte_Click()
    
    If cmbTipoReporte.text <> "EMPEÑOS" Then
        cmbOrigen.Enabled = False
        cmbDestino.Enabled = False
    Else
        cmbOrigen.Enabled = True
        cmbDestino.Enabled = True
    End If
    
    CrearEncabezados cmbTipoReporte.text
End Sub

Private Sub cmbTipoReporte_GotFocus()
    Cambiar_Color True, cmbTipoReporte
End Sub

Private Sub cmbTipoReporte_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoReporte_LostFocus()
    Cambiar_Color False, cmbTipoReporte
End Sub

Private Sub cmdBuscar_Click()
    
    If cmbTipoReporte.ListIndex = -1 Then
        MsgBox "Indique el tipo de reporte a visualizar !!!", vbInformation, "Búsqueda Avazanda"
        cmbTipoReporte.SetFocus
        Exit Sub
    End If
    
    llenarGrid cmbTipoReporte.text
    
End Sub

Private Sub cmdImprimir_Click()
    If grdReportes.Rows > 0 Then Exportar_Excel
End Sub

Private Sub cmdLimpiar_Click()
Dim ctrl As Object

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is ComboBox Then
            ctrl.ListIndex = -1
        ElseIf TypeOf ctrl Is TextBox Then
            ctrl.text = ""
        ElseIf TypeOf ctrl Is CheckBox Then
            ctrl.Value = 0
        End If
    Next
    
    PBar.Value = 0
    
End Sub

Private Sub cmdMosCliente_Click()
    frmMostrarCliente.Ver Me, txtCliente, True
End Sub

Private Sub cmdMosEmpleado_Click()
    frmMostrarUsuarios.Ver Me, txtEmpleado, True, False
End Sub

Private Sub cmdMosFechaFin_Click()
    txtFechaFin.text = frmCalendario.Fecha(txtFechaFin.text, 1)
End Sub

Private Sub cmdMosFechaIni_Click()
    txtFechaIni.text = frmCalendario.Fecha(txtFechaIni.text, 1)
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    llenarCombos
    PBar.Value = 0
End Sub

Private Sub grdReportes_Click(ByVal lRow As Long, ByVal lCol As Long)
    
    If lCol = 0 And lRow = 0 Then Exit Sub
    
    If lCol = 1 And lRow > 0 And grdReportes.CellIcon(lRow, 1) = 3 Then
        
        grdReportes.CellIcon(lRow, 1) = 4
        MuestraOculta grdReportes.CellItemData(lRow, 1), True
    
    ElseIf lCol = 1 And lRow > 0 And grdReportes.CellIcon(lRow, 1) = 4 Then
        
        grdReportes.CellIcon(lRow, 1) = 3
        MuestraOculta grdReportes.CellItemData(lRow, 1), False
        
    End If
    
End Sub

Private Sub MuestraOculta(ID As Long, Opcion As Boolean)
Dim i As Long

    For i = 1 To grdReportes.Rows

        If grdReportes.CellItemData(i, IIf(cmbTipoReporte.text = "EMPEÑOS", 4, 3)) = ID Then
            
            grdReportes.RowVisible(i) = Opcion
        End If

    Next i

End Sub

Private Sub txtCliente_GotFocus()
    Cambiar_Color True, txtCliente
    Seleccionar_Texto txtCliente
End Sub

Private Sub txtCliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCliente_LostFocus()
    Cambiar_Color False, txtCliente
End Sub

Private Sub txtEmpleado_GotFocus()
    Cambiar_Color True, txtEmpleado
End Sub

Private Sub txtEmpleado_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEmpleado_LostFocus()
    Cambiar_Color False, txtEmpleado
End Sub

Private Sub txtFechaFin_GotFocus()
    Cambiar_Color True, txtFechaFin
    Seleccionar_Texto txtFechaFin
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaFin_LostFocus()
    Cambiar_Color False, txtFechaFin
End Sub

Private Sub txtFechaIni_GotFocus()
    Cambiar_Color True, txtFechaIni
    Seleccionar_Texto txtFechaIni
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaIni_LostFocus()
    Cambiar_Color False, txtFechaIni
End Sub

Private Sub txtFolioFin_GotFocus()
    Cambiar_Color True, txtFolioFin
    Seleccionar_Texto txtFolioFin
End Sub

Private Sub txtFolioFin_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = Solo_Numeros(KeyAscii)
End Sub

Private Sub txtFolioFin_LostFocus()
    Cambiar_Color False, txtFolioFin
End Sub

Private Sub txtFolioIni_GotFocus()
    Cambiar_Color True, txtFolioIni
    Seleccionar_Texto txtFolioIni
End Sub

Private Sub txtFolioIni_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFolioIni_LostFocus()
    Cambiar_Color False, txtFolioIni
End Sub

Private Sub txtImportes_GotFocus()
    Cambiar_Color True, txtImportes
    Seleccionar_Texto txtImportes
End Sub

Private Sub txtImportes_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtImportes_LostFocus()
    Cambiar_Color False, txtImportes
End Sub

Public Sub Buscar_Cliente(ID As Long)

    txtCliente.text = SacaValor("Clientes", "Concat(Nombre,' ',Apellido)", " WHERE ID = " & ID)
    txtCliente.Tag = ID

End Sub

Public Sub Muestra_Privilegios(ID As Long)

    txtEmpleado.text = SacaValor("Usuarios", "Nombre", " WHERE ID = " & ID)
    txtEmpleado.Tag = ID

End Sub

Private Sub CrearEncabezados(Reporte As String)
Dim i As Integer
On Error GoTo Error
Err.Clear
    
    With grdReportes
    
        'Borro Encabezados y Datos
        .Clear True
        .ImageList = lstIcons
    
        If Reporte = "EMPEÑOS" Then
            .AddColumn "K1", "Contrato", ecgHdrTextALignRight, , 60, , , , , , , CCLSortNumeric
            .AddColumn "K2", "Fecha", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortDate
            .AddColumn "K3", "Vencimiento", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortDate
            .AddColumn "K4", "Cliente", ecgHdrTextALignLeft, , 280, , , , , , , CCLSortString
            .AddColumn "K5", "Empleado", ecgHdrTextALignLeft, , 120, , , , , , , CCLSortString
            .AddColumn "K6", "Categoría", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
            .AddColumn "K7", "Origen", ecgHdrTextALignLeft, , 110, , , , , , , CCLSortString
            .AddColumn "K8", "Destino", ecgHdrTextALignLeft, , 110, , , , , , , CCLSortString
            .AddColumn "K9", "Préstamo", ecgHdrTextALignRight, , 70, , , , , , , CCLSortNumeric
            .AddColumn "K10", "Avalúo", ecgHdrTextALignRight, , 70, , , , , , , CCLSortNumeric
            .AddColumn "K11", "Tasa", ecgHdrTextALignLeft, , 130, , , , , , , CCLSortString
            .AddColumn "K12", "Intereses", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
            .AddColumn "K13", "Iva", ecgHdrTextALignLeft, , 70, False, , , , , , CCLSortString
        End If
        
        If Reporte = "VENTAS" Or Reporte = "COMPRAS" Then
            .ClearItems
            .ClearSelection
            .AddColumn "K1", "Folio", ecgHdrTextALignRight, , 60, , , , , , , CCLSortNumeric
            .AddColumn "K2", "Fecha", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortDate
            .AddColumn "K3", "Cliente", ecgHdrTextALignLeft, , 280, , , , , , , CCLSortString
            .AddColumn "K4", "Empleado", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortString
            .AddColumn "K5", "Categoría", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
            
            If Reporte = "VENTAS" Then
                .AddColumn "K6", "SubTotal", ecgHdrTextALignRight, , 70, , , , , , , CCLSortNumeric
                .AddColumn "K7", "I.T.B.M.", ecgHdrTextALignRight, , 70, , , , , , , CCLSortNumeric
                .AddColumn "K8", "Descuento", ecgHdrTextALignRight, , 70, , , , , , , CCLSortNumeric
                .AddColumn "K9", "Total", ecgHdrTextALignRight, , 70, , , , , , , CCLSortNumeric
                .AddColumn "K10", "Estatus", ecgHdrTextALignRight, , 70, , , , , , , CCLSortString
            Else
                .AddColumn "K6", "Total", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
                .AddColumn "K7", "Estatus", ecgHdrTextALignRight, , 70, , , , , , , CCLSortString
            End If
            
        End If
        
    End With
        
Error:
    Maneja_Error Err
End Sub

Private Sub llenarCombos()
    
    'Tipo
    cmbTipo.Clear
    cmbTipo.AddItem "(TODOS)"
    Cargar_Combos "Descripcion", "Tipo", cmbTipo, , , False
    
    'Origen
    cmbOrigen.Clear
    cmbOrigen.AddItem ""
    cmbOrigen.ItemData(cmbOrigen.NewIndex) = 0
    cmbOrigen.AddItem "EMPEÑOS"
    cmbOrigen.ItemData(cmbOrigen.NewIndex) = OD_EMPENO
    cmbOrigen.AddItem "REFRENDOS"
    cmbOrigen.ItemData(cmbOrigen.NewIndex) = OD_REFRENDO
    
    'Destino
    cmbDestino.Clear
    cmbDestino.AddItem ""
    cmbDestino.ItemData(cmbDestino.NewIndex) = 0
    cmbDestino.AddItem "PAGO INTERÉS"
    cmbDestino.ItemData(cmbDestino.NewIndex) = OD_REFRENDO
    cmbDestino.AddItem "RETIROS"
    cmbDestino.ItemData(cmbDestino.NewIndex) = D_DESEMPEÑO
    cmbDestino.AddItem "FUNDICIÓN"
    cmbDestino.ItemData(cmbDestino.NewIndex) = D_ALMONEDA

End Sub

Private Sub llenarGrid(Reporte As String)
Dim crSubtotal As Double, crDescuento As Double, crIva As Double, columna As Long, Fila As Long, i As Long
Dim sPrestamo As Double, sAvaluo As Double, sIntereses As Double, sDescuento As Double, sIva As Double, sSubTotal As Double, sTotal As Double
On Error GoTo Error
Err.Clear

    Screen.MousePointer = vbHourglass

    Set rcConsulta = New ADODB.Recordset
    rcConsulta.CursorLocation = adUseClient
    
    If Reporte = "EMPEÑOS" Then
        
        rcConsulta.Open "select distinct e.id,e.numcontrato,e.folioorigen,e.foliodestino,e.cancelado,e.fecha,e.vencimiento,concat(c.nombre,' ',c.apellido) as Cliente,u.nombre as " & _
                        "empleado,t.descripcion as categoria,e.origen,e.destino,e.prestamo,e.avaluo,concat(e.tipointeres,' ',e.tipotasa) " & _
                        "as tasa,e.intereses+e.importealmacenaje+e.importeseguro+e.importemoratorios+e.importeperdida+e.importeotros " & _
                        "as interesess,e.importeiva from empeno e inner join clientes c on c.id = e.idcliente inner join usuarios u on u.id = e.idusuario " & _
                        "inner join detallesempeno d on d.idempeno = e.id inner join tipo t on t.id = d.tipo where " & Condiciones(cmbTipoReporte.text), dbDatos, adOpenForwardOnly, adLockReadOnly
            
        'Barra de Progreso
        If rcConsulta.RecordCount = 0 Then
            MsgBox "No se encontraron registros !!", vbInformation, "Búsqueda Avanzada"
            GoTo Error
        End If
        i = 0
        PBar.Min = 0
        PBar.Max = rcConsulta.RecordCount
        
        With grdReportes
        
            .Redraw = False
            .Clear
            
            While Not rcConsulta.EOF
                
                DoEvents
                .AddRow
                .CellDetails .Rows, 1, rcConsulta!NumContrato, DT_RIGHT Or DT_WORD_ELLIPSIS, , , , , , , rcConsulta!ID
                .CellIcon(.Rows, 1) = 3
                .CellDetails .Rows, 2, Format(rcConsulta!Fecha, "DD/MM/YYYY"), DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 3, Format(rcConsulta!Vencimiento, "DD/MM/YYYY"), DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 4, rcConsulta!Cliente, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 5, rcConsulta!Empleado, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 6, rcConsulta!Categoria, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 7, OD_Origen(rcConsulta!Origen) & IIf(Val(rcConsulta!cancelado) = 1 And rcConsulta!Destino = 0, "/Anulado", "/" & rcConsulta!FolioOrigen), DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 8, OD_Origen(rcConsulta!Destino) & IIf(Val(rcConsulta!cancelado) = 1 And rcConsulta!Destino >= 0, "/Anulado", IIf(rcConsulta!Destino = D_VENTA Or rcConsulta!Destino = OD_REFRENDO, "/" & rcConsulta!foliodestino, ""))
                .CellDetails .Rows, 9, Format(rcConsulta!Prestamo, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 10, Format(rcConsulta!Avaluo, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 11, rcConsulta!Tasa, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 12, Format(rcConsulta!interesess, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 13, Format(rcConsulta!ImporteIva, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                
                'Hago las Sumatorias
                sPrestamo = sPrestamo + rcConsulta!Prestamo
                sAvaluo = sAvaluo + rcConsulta!Avaluo
                sIntereses = sIntereses + rcConsulta!interesess
                sIva = sIva + rcConsulta!ImporteIva
                
                'Pongo el Fondo
                i = i + 1
                Poner_Colores grdReportes, .Rows, i
                
                'Pongo los Detalles
                detalles Reporte, rcConsulta!ID
                
                PBar.Value = i
                rcConsulta.MoveNext
            Wend
        
            'Pongo los Totales
            DoEvents
            .AddRow
            .CellDetails .Rows, 1, i, DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
            .CellDetails .Rows, 9, Format(sPrestamo, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
            .CellDetails .Rows, 10, Format(sAvaluo, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
            .CellDetails .Rows, 12, Format(sIntereses, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
            .CellDetails .Rows, 13, Format(sIva, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
        
        End With
    
    End If
    
    If Reporte = "VENTAS" Then
    
        rcConsulta.Open "select distinct v.id,v.folio,v.apartado,v.pagado,v.cancelado,v.fecha,concat(c.nombre,' ',c.apellido) as cliente,u.nombre as empleado,t.descripcion as categoria,v.total,v.descuento,v.iva " & _
            "from ventas v inner join detallesventas dv on v.id = dv.idventa left join detallesempeno de on de.codigo = dv.codigo " & _
            "left join empeno e on e.id = de.idempeno inner join clientes c on c.id = v.idcliente inner join usuarios u on u.id = v.idusuario " & _
            "left join tipo t on t.id = de.tipo where " & Condiciones(cmbTipoReporte.text), dbDatos, adOpenKeyset, adLockReadOnly

    End If

    If Reporte = "COMPRAS" Then
        
        rcConsulta.Open "select distinct c.id,c.folio,c.cancelado,c.fecha,concat(cl.nombre,' ',cl.apellido) as cliente,u.nombre as empleado,c.total, t.descripcion as categoria " & _
            "from compras c inner join detallescompras dc on c.id = dc.idcompra inner join clientes cl on cl.id = c.idcliente inner join usuarios u on u.id " & _
            "= c.idusuario inner join tipo t on t.id = dc.tipo where " & Condiciones(cmbTipoReporte.text), dbDatos, adOpenKeyset, adLockReadOnly
            
    End If

    'Lleno Ventas y Compras
    If Reporte <> "EMPEÑOS" Then
        
        'Barra de Progreso
        If rcConsulta.RecordCount = 0 Then
            MsgBox "No se encontraron registros !!", vbInformation, "Búsqueda Avanzada"
            GoTo Error
        End If
        i = 0
        PBar.Min = 0
        PBar.Max = rcConsulta.RecordCount
        
        With grdReportes
            
            .Redraw = False
            .Clear
            
            While Not rcConsulta.EOF
                
                DoEvents
                .AddRow
                .CellDetails .Rows, 1, rcConsulta!Folio, DT_RIGHT Or DT_WORD_ELLIPSIS, , , , , , , rcConsulta!ID
                .CellIcon(.Rows, 1) = 3
                .CellDetails .Rows, 2, Format(rcConsulta!Fecha, "DD/MM/YYYY"), DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 3, rcConsulta!Cliente, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 4, rcConsulta!Empleado, DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 5, rcConsulta!Categoria, DT_LEFT Or DT_WORD_ELLIPSIS
                
                If Reporte = "COMPRAS" Then
                    
                    .CellDetails .Rows, 6, Format(rcConsulta!Total, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                    .CellDetails .Rows, 7, IIf(Val(rcConsulta!cancelado) = 1, "Cancelado", "Pagado"), DT_LEFT Or DT_WORD_ELLIPSIS
                    
                    'Hago la Sumatoria
                    sTotal = sTotal + rcConsulta!Total
                Else
                    
                    'Hago los cálculos
                    crIva = rcConsulta!Total * (rcConsulta!Iva / 100)
                    crDescuento = rcConsulta!Total * (rcConsulta!Descuento / 100)
                    crSubtotal = (rcConsulta!Total + crIva) - crDescuento
                    
                    .CellDetails .Rows, 6, Format(crSubtotal, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                    .CellDetails .Rows, 7, Format(crIva, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                    .CellDetails .Rows, 8, Format(crDescuento, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                    .CellDetails .Rows, 9, Format(rcConsulta!Total, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                    .CellDetails .Rows, 10, IIf(Val(rcConsulta!cancelado) = 1, "Cancelado", IIf(Val(rcConsulta!apartado) = 0 Or (Val(rcConsulta!apartado) = 1 And Val(rcConsulta!Pagado) = 1), "Pagado", "Activo")), DT_LEFT Or DT_WORD_ELLIPSIS
                    
                    'Hago la Sumatoria
                    sSubTotal = sSubTotal + crSubtotal
                    sIva = sIva + crIva
                    sDescuento = sDescuento + crDescuento
                    sTotal = sTotal + rcConsulta!Total
                    
                End If
                
                'Pongo el Fondo
                i = i + 1
                Poner_Colores grdReportes, .Rows, i

                'Pongo los Detalles
                detalles Reporte, rcConsulta!ID
                
                PBar.Value = i
                rcConsulta.MoveNext
            Wend
        
            'Pongo los Totales
            DoEvents
            .AddRow
            .CellDetails .Rows, 1, i, DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
            .CellDetails .Rows, 6, IIf(Reporte = "COMPRAS", Format(sTotal, FMoneda), Format(sSubTotal, FMoneda)), DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
            
            If Reporte = "VENTAS" Then
                .CellDetails .Rows, 7, Format(sIva, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
                .CellDetails .Rows, 8, Format(sDescuento, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
                .CellDetails .Rows, 9, Format(sTotal, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS, , , &HC0&
            End If
            
        End With
        
    End If
    
    grdReportes.Redraw = True

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
    PBar.Value = 0
    Screen.MousePointer = vbDefault
End Sub

Public Function Condiciones(Reporte As String) As String
Dim tRep As Integer

    Select Case Reporte
    
       Case "EMPEÑOS"
            tRep = 1
            Condiciones = " 1 = 1"
        Case "VENTAS"
            tRep = 2
            Condiciones = " 1 = 1"
        Case "COMPRAS"
            tRep = 3
            Condiciones = " 1 = 1"

    
    End Select
    
    'Condicion Inicial
    'Condiciones = " e.tipointeres='EMPEÑOS' "
    
    'Rango de Fechas
    If Len(txtFechaIni.text) > 0 Then Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.fecha", IIf(tRep = 2, "v.fecha", "c.fecha")) & ">='" & Format(txtFechaIni.text, "YYYY/MM/DD") & "'"
    If Len(txtFechaFin.text) > 0 Then Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.fecha", IIf(tRep = 2, "v.fecha", "c.fecha")) & "<='" & Format(txtFechaFin.text, "YYYY/MM/DD") & "'"

    'Rango de Folios
    If Len(txtFolioIni.text) > 0 Then Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.numcontrato", IIf(tRep = 2, "v.folio", "c.folio")) & ">=" & txtFolioIni.text
    If Len(txtFolioFin.text) > 0 Then Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.numcontrato", IIf(tRep = 2, "v.folio", "c.folio")) & "<=" & txtFolioFin.text
    
    'Cliente Especifico
    If Len(txtCliente.text) > 0 Then Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.idcliente", IIf(tRep = 2, "v.idcliente", "c.idcliente")) & " = " & txtCliente.Tag
    
    'Empleado Específico
    If Len(txtEmpleado.text) > 0 Then Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.idusuario", IIf(tRep = 2, "v.idusuario", "c.idusuario")) & " = " & txtEmpleado.Tag
    
    'Categoría (Tipo)
    If cmbTipo.ListIndex > 0 Then Condiciones = Condiciones & " and " & IIf(tRep = 1, "d.tipo", IIf(tRep = 2, "de.tipo", "dc.tipo")) & " = " & cmbTipo.ItemData(cmbTipo.ListIndex)

    'Origen
    If cmbOrigen.ListIndex > 0 Then Condiciones = Condiciones & IIf(tRep = 1, " and e.origen = " & cmbOrigen.ItemData(cmbOrigen.ListIndex), "")

    'Destino
    If cmbDestino.ListIndex > 0 Then Condiciones = Condiciones & IIf(tRep = 1, " and e.destino = " & cmbDestino.ItemData(cmbDestino.ListIndex), "")
    
    'Importes
    If Len(txtImportes.text) > 0 Then Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.prestamo", IIf(tRep = 2, "v.total", "c.total")) & " > " & txtImportes.text

    'Filtros Extra
    If chkPagados.Value And tRep <> 3 Then Condiciones = Condiciones & " And " & IIf(tRep = 1, "e.pagado = 1", IIf(tRep = 2, "v.apartado = 0 or (v.apartado = 1 and v.pagado = 1)", "c.pagado = 1"))
    
    If chkNoPagados.Value And tRep <> 3 Then Condiciones = Condiciones & " And " & IIf(tRep = 1, "e.pagado = 0", IIf(tRep = 2, "v.apartado = 1 and v.pagado = 0", "c.pagado = 0"))
    
    If chkCancelados.Value Then
        Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.cancelado = 1", IIf(tRep = 2, "v.cancelado = 1", "c.cancelado = 1"))
    Else
        Condiciones = Condiciones & " and " & IIf(tRep = 1, "e.cancelado = 0", IIf(tRep = 2, "v.cancelado = 0", "c.cancelado = 0"))
    End If

    'Orden
    If opFecha.Value Then
        Condiciones = Condiciones & " ORDER BY " & IIf(tRep = 1, "e.fecha", IIf(tRep = 2, "v.fecha", "c.fecha"))
    ElseIf opFolio.Value Then
        Condiciones = Condiciones & " ORDER BY " & IIf(tRep = 1, "e.numcontrato", IIf(tRep = 2, "v.folio", "c.folio"))
    ElseIf opCliente.Value Then
        Condiciones = Condiciones & " ORDER BY cliente"
    ElseIf opImporte.Value Then
        Condiciones = Condiciones & " ORDER BY " & IIf(tRep = 1, "e.prestamo", IIf(tRep = 2, "v.total", "c.total"))
    ElseIf opFechaMovimiento.Value Then
        Condiciones = Condiciones & " ORDER BY " & IIf(tRep = 1, "e.fechamovimiento", IIf(tRep = 2, "v.fechamovimiento", "c.fechamovimiento"))
    End If

End Function

Private Sub Exportar_Excel()
Dim Excel As Object, i As Integer, Col As Integer, Y As Integer, str As String, detalles As Boolean, Pos As Long, Ban As Boolean
On Error GoTo Error
Err.Clear

    If MsgBox(" Desea imprimir los detalles ??", vbQuestion + vbYesNo + vbDefaultButton1, "Reporte Detallado") = vbYes Then
        
        detalles = True
    Else
        
        detalles = False
    End If

    Screen.MousePointer = vbHourglass
    DoEvents
    
    'Creo la Referencia al Excel
    Set Excel = CreateObject("Excel.application")
    
    With Excel
        
        'Barra de Progreso
        PBar.Min = 0
        PBar.Max = grdReportes.Rows
        
        'Agrego un Nuevo Libro
        .Workbooks.Add
        
        'Creo los Encabezados
        For i = 1 To grdReportes.Columns
            DoEvents
            .Cells(1, i).Formula = grdReportes.ColumnHeader("K" & i)
        Next i
        
        Pos = 1
        
        For i = 1 To grdReportes.Rows
            DoEvents
            
            For Y = 1 To grdReportes.Columns
                
                'Omitir los Detalles
                If detalles = False And Y = 1 Then
                    If grdReportes.CellItemData(i, IIf(cmbTipoReporte.text = "EMPEÑOS", 4, 3)) > 0 Then
                        Ban = True
                        Exit For
                    End If
                End If
                
                .Cells(Pos + 1, Y).Formula = grdReportes.CellText(i, Y)
                
            Next Y
        
            If Ban Then
                Ban = False
            Else
                Pos = Pos + 1
            End If
            
            PBar.Value = i
        Next i

        ' autoajustar las columnas
        .Columns("A:A").EntireColumn.AutoFit
        .Columns("B:B").EntireColumn.AutoFit
        .Columns("C:C").EntireColumn.AutoFit
        .Columns("D:D").EntireColumn.AutoFit
        .Columns("E:E").EntireColumn.AutoFit
        .Columns("F:F").EntireColumn.AutoFit
        .Columns("G:G").EntireColumn.AutoFit
        .Columns("H:H").EntireColumn.AutoFit
        .Columns("I:I").EntireColumn.AutoFit
        .Columns("J:J").EntireColumn.AutoFit
        .Columns("K:K").EntireColumn.AutoFit
        .Columns("L:L").EntireColumn.AutoFit
        .Columns("M:M").EntireColumn.AutoFit
    
        .Range("A1:M1").Select
        With .Selection.Font
            .Name = "Arial"
            .FontStyle = "Negrita"
            .Size = 10
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
        End With
        
        str = "M" & grdReportes.Rows + 1
        '.ActiveSheet.Range("A1", str).HorizontalAlignment = xlHAlignLeft
        .ActiveSheet.Range("A1", str).HorizontalAlignment = -4131
        .Selection.Interior.ColorIndex = 35
        
        'Hago Visible la Referencia
        .Visible = True

    End With
Error:
    Set Excel = Nothing
    Maneja_Error Err
    PBar.Value = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub detalles(Reporte As String, ID As Long)
Dim RSPrendas As New ADODB.Recordset
On Error GoTo Error

    Screen.MousePointer = vbHourglass

    Select Case Reporte
    
        Case "EMPEÑOS"
            
            RSPrendas.Open "SELECT d.IDEmpeno as ID,d.Cantidad,concat(d.Articulo,' ',d.Observaciones) as Prenda,d.Peso,d.Estado,kilatajes.Descripcion as kilataje," & _
                "d.Prestamo,d.Avaluo,d.origen,d.destino FROM detallesempeno d LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE " & _
                "d.IDEmpeno=" & ID, dbDatos, adOpenKeyset, adLockReadOnly
        
        Case "COMPRAS"
            
            RSPrendas.Open "Select d.idcompra as ID,d.cantidad,concat(d.descripcion,' ',d.observaciones) as Prenda,k.descripcion as kilataje,d.peso," & _
                "d.costo,d.precio,d.estado from detallescompras d inner join kilatajes k on k.clave = d.kilates where d.idcompra = " & ID, dbDatos, adOpenKeyset, adLockReadOnly
            
        Case "VENTAS"
        
            RSPrendas.Open "Select distinct d.idventa as ID,di.cantidad,concat(di.descripcion,' ',di.observaciones) as Prenda, " & _
                "k.descripcion as kilataje,di.Peso , di.costo, di.Precio, di.Estado from detallesventas d inner join detallesentradainventario " & _
                "di on di.codigo = d.codigo inner join kilatajes k on k.clave = d.kilates where d.idventa = " & ID, dbDatos, adOpenKeyset, adLockReadOnly
    
    End Select

    'Lleno el Detalle
    With grdReportes
        
        While Not RSPrendas.EOF
        
            .AddRow
            .CellDetails .Rows, IIf(Reporte = "EMPEÑOS", 4, 3), RSPrendas!Cantidad & " " & RSPrendas!Prenda & " " & RSPrendas!Kilataje & " " & RSPrendas!Peso & "  Grms.; E: " & RSPrendas!Estado, DT_LEFT Or DT_WORD_ELLIPSIS, , , , , , , RSPrendas!ID
            
            If Reporte = "EMPEÑOS" Then
                '.CellDetails .Rows, 7, OD_Origen(Val(RSPrendas!Origen)), DT_LEFT Or DT_WORD_ELLIPSIS
                '.CellDetails .Rows, 8, OD_Origen(Val(RSPrendas!Destino)), DT_LEFT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 9, Format(RSPrendas!Prestamo, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
                .CellDetails .Rows, 10, Format(RSPrendas!Avaluo, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
            Else
                .CellDetails .Rows, 6, Format(RSPrendas!Precio, FMoneda), DT_RIGHT Or DT_WORD_ELLIPSIS
            End If
            
            SombreaGrid grdReportes, 239, 239, 239, 255, 255, 255, grdReportes.Rows
            
            .RowVisible(.Rows) = False
            RSPrendas.MoveNext
        Wend
    End With
    
Error:
    RSPrendas.Close
    Set RSPrendas = Nothing
    Screen.MousePointer = vbDefault
End Sub
