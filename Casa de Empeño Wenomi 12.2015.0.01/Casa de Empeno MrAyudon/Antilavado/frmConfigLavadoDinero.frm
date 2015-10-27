VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TabCtl32.Ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form frmConfigLavadoDinero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros"
   ClientHeight    =   7935
   ClientLeft      =   2640
   ClientTop       =   2370
   ClientWidth     =   10140
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfigLavadoDinero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   10140
   Tag             =   "Avaluoprestamo"
   Begin VB.Frame Frame7 
      Caption         =   "Aseguradora"
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
      Height          =   1485
      Left            =   4440
      TabIndex        =   97
      Top             =   8760
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox txtAseguradora 
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
         Left            =   1575
         MaxLength       =   50
         TabIndex        =   100
         Tag             =   "Aseguradora"
         Top             =   1080
         Width           =   2280
      End
      Begin VB.TextBox txtFechaexpedicion 
         Alignment       =   2  'Center
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
         Left            =   2760
         Locked          =   -1  'True
         TabIndex        =   99
         Tag             =   "FechaExpedicion"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txtPolizano 
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
         Left            =   2760
         MaxLength       =   12
         TabIndex        =   98
         Tag             =   "PolizaSeguro"
         Top             =   360
         Width           =   1095
      End
      Begin DevPowerFlatBttn.FlatBttn cmdExpedicion 
         Height          =   300
         Left            =   3885
         TabIndex        =   101
         Top             =   720
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
         Picture         =   "frmConfigLavadoDinero.frx":2832
      End
      Begin VB.Label Label35 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aseguradora:"
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
         Left            =   45
         TabIndex        =   104
         Top             =   1087
         Width           =   1500
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha de expedición:"
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
         Left            =   45
         TabIndex        =   103
         Top             =   720
         Width           =   2340
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Póliza No."
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
         Left            =   45
         TabIndex        =   102
         Top             =   360
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Precios Compra/Venta Oro"
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
      Height          =   2535
      Left            =   480
      TabIndex        =   78
      Top             =   8640
      Visible         =   0   'False
      Width           =   3330
      Begin VB.TextBox txt8K 
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
         Left            =   735
         MaxLength       =   6
         TabIndex        =   90
         Tag             =   "9K"
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txt10K 
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
         Left            =   735
         MaxLength       =   6
         TabIndex        =   89
         Tag             =   "10K"
         Top             =   705
         Width           =   1095
      End
      Begin VB.TextBox txt14K 
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
         Left            =   735
         MaxLength       =   6
         TabIndex        =   88
         Tag             =   "14K"
         Top             =   1065
         Width           =   1095
      End
      Begin VB.TextBox txt18K 
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
         Left            =   735
         MaxLength       =   6
         TabIndex        =   87
         Tag             =   "18K"
         Top             =   1425
         Width           =   1095
      End
      Begin VB.TextBox txtVen8K 
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
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   86
         Tag             =   "Venta9K"
         Top             =   345
         Width           =   1095
      End
      Begin VB.TextBox txtVen18K 
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
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   85
         Tag             =   "Venta18K"
         Top             =   1425
         Width           =   1095
      End
      Begin VB.TextBox txtVen14K 
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
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   84
         Tag             =   "Venta14K"
         Top             =   1065
         Width           =   1095
      End
      Begin VB.TextBox txtVen10K 
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
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   83
         Tag             =   "Venta10K"
         Top             =   705
         Width           =   1095
      End
      Begin VB.TextBox txtVenta21K 
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
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   82
         Tag             =   "Venta21K"
         Top             =   1785
         Width           =   1095
      End
      Begin VB.TextBox txt21K 
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
         Left            =   735
         MaxLength       =   6
         TabIndex        =   81
         Tag             =   "21K"
         Top             =   1785
         Width           =   1095
      End
      Begin VB.TextBox txtVenta24K 
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
         Left            =   1935
         MaxLength       =   6
         TabIndex        =   80
         Tag             =   "Venta24K"
         Top             =   2145
         Width           =   1095
      End
      Begin VB.TextBox txt24K 
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
         Left            =   735
         MaxLength       =   6
         TabIndex        =   79
         Tag             =   "24K"
         Top             =   2145
         Width           =   1095
      End
      Begin VB.Label Label38 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "9K:"
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
         Left            =   300
         TabIndex        =   96
         Top             =   345
         Width           =   375
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10K:"
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
         Left            =   150
         TabIndex        =   95
         Top             =   705
         Width           =   525
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "14K:"
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
         Left            =   150
         TabIndex        =   94
         Top             =   1065
         Width           =   525
      End
      Begin VB.Label Label53 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "18K:"
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
         Left            =   150
         TabIndex        =   93
         Top             =   1425
         Width           =   525
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "21K:"
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
         Left            =   150
         TabIndex        =   92
         Top             =   1785
         Width           =   525
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "24K:"
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
         Left            =   150
         TabIndex        =   91
         Top             =   2145
         Width           =   525
      End
   End
   Begin VB.Frame frmPlata 
      Caption         =   "Precios Compra / Venta Plata"
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
      Height          =   4215
      Left            =   0
      TabIndex        =   47
      Top             =   8640
      Width           =   3735
      Begin VB.TextBox txtVenta999K 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   65
         Tag             =   "Venta999K"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txt999K 
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
         Left            =   945
         MaxLength       =   6
         TabIndex        =   64
         Tag             =   "999K"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.TextBox txtVenta925K 
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
         Left            =   2145
         MaxLength       =   6
         TabIndex        =   63
         Tag             =   "Venta925K"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txt925K 
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
         Left            =   945
         MaxLength       =   6
         TabIndex        =   62
         Tag             =   "925K"
         Top             =   3240
         Width           =   1095
      End
      Begin VB.TextBox txtVenta900K 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   61
         Tag             =   "Venta900K"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txt900K 
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
         Left            =   945
         MaxLength       =   6
         TabIndex        =   60
         Tag             =   "900K"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox txt800K 
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
         Left            =   945
         MaxLength       =   6
         TabIndex        =   59
         Tag             =   "800K"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txtVenta800K 
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
         Left            =   2145
         MaxLength       =   6
         TabIndex        =   58
         Tag             =   "Venta800K"
         Top             =   2520
         Width           =   1095
      End
      Begin VB.TextBox txt720K 
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
         Left            =   960
         MaxLength       =   6
         TabIndex        =   57
         Tag             =   "720K"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtVenta720K 
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
         Left            =   2145
         MaxLength       =   6
         TabIndex        =   56
         Tag             =   "Venta720K"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.TextBox txtVenta300K 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   55
         Tag             =   "Venta300K"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtVenta400K 
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
         Left            =   2160
         MaxLength       =   6
         TabIndex        =   54
         Tag             =   "Venta400K"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txtVenta500K 
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
         Left            =   2145
         MaxLength       =   6
         TabIndex        =   53
         Tag             =   "Venta500K"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txtVenta100K 
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
         Left            =   2145
         MaxLength       =   6
         TabIndex        =   52
         Tag             =   "Venta100K"
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox txt500K 
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
         Left            =   945
         MaxLength       =   6
         TabIndex        =   51
         Tag             =   "500K"
         Top             =   1800
         Width           =   1095
      End
      Begin VB.TextBox txt400K 
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
         Left            =   960
         MaxLength       =   6
         TabIndex        =   50
         Tag             =   "400K"
         Top             =   1440
         Width           =   1095
      End
      Begin VB.TextBox txt300K 
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
         Left            =   945
         MaxLength       =   6
         TabIndex        =   49
         Tag             =   "300K"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txt100K 
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
         Left            =   945
         MaxLength       =   6
         TabIndex        =   48
         Tag             =   "100K"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lbl999 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".999:"
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
         Left            =   165
         TabIndex        =   77
         Top             =   3600
         Width           =   600
      End
      Begin VB.Label lbl925 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".925:"
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
         Left            =   165
         TabIndex        =   76
         Top             =   3240
         Width           =   600
      End
      Begin VB.Label lbl800 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".900:"
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
         Left            =   165
         TabIndex        =   75
         Top             =   2880
         Width           =   600
      End
      Begin VB.Label lbl24K 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".800:"
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
         Left            =   165
         TabIndex        =   74
         Top             =   2520
         Width           =   600
      End
      Begin VB.Label lbl22K 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".720:"
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
         Left            =   165
         TabIndex        =   73
         Top             =   2160
         Width           =   600
      End
      Begin VB.Label lbl18K 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".500:"
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
         Left            =   165
         TabIndex        =   72
         Top             =   1800
         Width           =   600
      End
      Begin VB.Label lbl14K 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".400:"
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
         Left            =   165
         TabIndex        =   71
         Top             =   1440
         Width           =   600
      End
      Begin VB.Label lbl10K 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".300:"
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
         Left            =   165
         TabIndex        =   70
         Top             =   1080
         Width           =   600
      End
      Begin VB.Label lblK 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   ".100:"
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
         Left            =   165
         TabIndex        =   69
         Top             =   720
         Width           =   600
      End
      Begin VB.Label lblVENTA 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "VENTA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   2280
         TabIndex        =   68
         Top             =   375
         Width           =   825
      End
      Begin VB.Label lblCOMPRA 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Caption         =   "COMPRA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Index           =   2
         Left            =   945
         TabIndex        =   67
         Top             =   375
         Width           =   1095
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00808080&
         BackStyle       =   0  'Transparent
         Height          =   300
         Index           =   1
         Left            =   945
         TabIndex        =   66
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame frmPreciosPlataValuacion 
      Caption         =   "% Precios Plata Valuación"
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
      Height          =   1140
      Left            =   4440
      TabIndex        =   38
      Top             =   8760
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox txtPlataM 
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
         Left            =   2640
         TabIndex        =   42
         Tag             =   "CalidadPlataM"
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txtPlataR 
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
         Left            =   720
         TabIndex        =   41
         Tag             =   "CalidadPlataR"
         Top             =   660
         Width           =   1335
      End
      Begin VB.TextBox txtPlataEx 
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
         Left            =   720
         TabIndex        =   40
         Tag             =   "CalidadPlataEx"
         Top             =   300
         Width           =   1335
      End
      Begin VB.TextBox txtPlataB 
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
         Left            =   2640
         TabIndex        =   39
         Tag             =   "CalidadPlataB"
         Top             =   300
         Width           =   1335
      End
      Begin VB.Label lblPlataEx 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EX:"
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
         Left            =   240
         TabIndex        =   46
         Top             =   307
         Width           =   360
      End
      Begin VB.Label lblPlataB 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
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
         Left            =   2250
         TabIndex        =   45
         Top             =   307
         Width           =   225
      End
      Begin VB.Label lblPlataR 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
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
         Left            =   240
         TabIndex        =   44
         Top             =   667
         Width           =   240
      End
      Begin VB.Label lblPlataM 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M:"
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
         Left            =   2250
         TabIndex        =   43
         Top             =   667
         Width           =   270
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "% Precios Oro Valuación"
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
      Height          =   1095
      Left            =   4200
      TabIndex        =   24
      Top             =   9960
      Visible         =   0   'False
      Width           =   4245
      Begin VB.TextBox txtM 
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
         Left            =   3360
         TabIndex        =   28
         Tag             =   "CalidadM"
         Top             =   9360
         Width           =   1095
      End
      Begin VB.TextBox txtR 
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
         Left            =   1440
         TabIndex        =   27
         Tag             =   "CalidadR"
         Top             =   9360
         Width           =   1095
      End
      Begin VB.TextBox txtEx 
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
         Left            =   1440
         TabIndex        =   26
         Tag             =   "CalidadEx"
         Top             =   9000
         Width           =   1095
      End
      Begin VB.TextBox txtB 
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
         Left            =   3360
         TabIndex        =   25
         Tag             =   "CalidadB"
         Top             =   9000
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   5
         Left            =   4500
         TabIndex        =   36
         Top             =   9000
         Width           =   270
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   5
         Left            =   4500
         TabIndex        =   35
         Top             =   9360
         Width           =   270
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   4
         Left            =   2580
         TabIndex        =   34
         Top             =   9000
         Width           =   270
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "%"
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
         Index           =   4
         Left            =   2580
         TabIndex        =   33
         Top             =   9360
         Width           =   270
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EX:"
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
         Left            =   960
         TabIndex        =   32
         Top             =   9000
         Width           =   360
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "B:"
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
         Left            =   3015
         TabIndex        =   31
         Top             =   9000
         Width           =   225
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "R:"
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
         TabIndex        =   30
         Top             =   9360
         Width           =   240
      End
      Begin VB.Label Label44 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "M:"
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
         Left            =   2970
         TabIndex        =   29
         Top             =   9360
         Width           =   270
      End
   End
   Begin VB.TextBox txtPrecioAutos 
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
      Left            =   2520
      MaxLength       =   10
      TabIndex        =   23
      Tag             =   "PrecioAutos"
      Top             =   10320
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Parametros"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7155
      Left            =   120
      TabIndex        =   18
      Top             =   0
      Width           =   9885
      Begin TabDlg.SSTab SSTab1 
         Height          =   6645
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   11721
         _Version        =   393216
         Style           =   1
         Tabs            =   1
         TabHeight       =   520
         TabMaxWidth     =   3528
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "LEY DE LAVADO DE DINERO"
         TabPicture(0)   =   "frmConfigLavadoDinero.frx":2947
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblActividadDe"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Label1"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Label2"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "lblTipoDe"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "Label3"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame11"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "cmbActividadVulnerable"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtNumConstancia"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "cmdDatosDe"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).Control(9)=   "cmbTipoMoneda"
         Tab(0).Control(9).Enabled=   0   'False
         Tab(0).Control(10)=   "cmbTipoGiro"
         Tab(0).Control(10).Enabled=   0   'False
         Tab(0).Control(11)=   "txtRutaArchivoXML"
         Tab(0).Control(11).Enabled=   0   'False
         Tab(0).Control(12)=   "cdlgDir"
         Tab(0).Control(12).Enabled=   0   'False
         Tab(0).Control(13)=   "Command1"
         Tab(0).Control(13).Enabled=   0   'False
         Tab(0).ControlCount=   14
         Begin VB.CommandButton Command1 
            Caption         =   ". . ."
            Height          =   360
            Left            =   8880
            TabIndex        =   6
            Top             =   3960
            Width           =   510
         End
         Begin MSComDlg.CommonDialog cdlgDir 
            Left            =   8880
            Top             =   3120
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin VB.TextBox txtRutaArchivoXML 
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
            Left            =   390
            MaxLength       =   250
            TabIndex        =   5
            Tag             =   "RutaArchivosXML"
            Top             =   3960
            Width           =   8415
         End
         Begin VB.ComboBox cmbTipoGiro 
            Height          =   360
            ItemData        =   "frmConfigLavadoDinero.frx":2963
            Left            =   360
            List            =   "frmConfigLavadoDinero.frx":2965
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Tag             =   "IdTipoGiroMercantil"
            Top             =   1680
            Width           =   9045
         End
         Begin VB.ComboBox cmbTipoMoneda 
            Height          =   360
            ItemData        =   "frmConfigLavadoDinero.frx":2967
            Left            =   360
            List            =   "frmConfigLavadoDinero.frx":2969
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Tag             =   "IdTipoMonedaLocal"
            Top             =   2520
            Width           =   9045
         End
         Begin VB.CommandButton cmdDatosDe 
            Caption         =   "Datos de la Sucursal"
            Height          =   360
            Left            =   6480
            TabIndex        =   4
            Top             =   3120
            Width           =   2190
         End
         Begin VB.TextBox txtNumConstancia 
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
            Left            =   3030
            MaxLength       =   30
            TabIndex        =   3
            Tag             =   "NumConstancia"
            Top             =   3120
            Width           =   3015
         End
         Begin VB.ComboBox cmbActividadVulnerable 
            Height          =   360
            ItemData        =   "frmConfigLavadoDinero.frx":296B
            Left            =   360
            List            =   "frmConfigLavadoDinero.frx":296D
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Tag             =   "IDActividadVulnerable"
            Top             =   840
            Width           =   9045
         End
         Begin VB.Frame Frame11 
            Caption         =   "Variables atípicas"
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
            Height          =   1695
            Left            =   240
            TabIndex        =   105
            Top             =   4560
            Width           =   9180
            Begin VB.TextBox txtImporteVSMPrestamos 
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
               Left            =   3525
               MaxLength       =   10
               TabIndex        =   9
               Tag             =   "ImporteVSMPrestamos"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtImporteUdi 
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
               Left            =   7590
               TabIndex        =   11
               Tag             =   "ImporteUdi"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtPrestamoCheque 
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
               Left            =   3510
               MaxLength       =   10
               TabIndex        =   7
               Tag             =   "PrestamoCheque"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtCompraCheque 
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
               Left            =   3510
               MaxLength       =   10
               TabIndex        =   8
               Tag             =   "CompraCheque"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtImporteSalario 
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
               Left            =   7590
               TabIndex        =   10
               Tag             =   "ImporteSalario"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Importe Prestamos (V.S.M) :"
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
               Left            =   240
               TabIndex        =   115
               Top             =   1080
               Width           =   3195
            End
            Begin VB.Label Label71 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor UDI:"
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
               Left            =   6330
               TabIndex        =   109
               Top             =   720
               Width           =   1170
            End
            Begin VB.Label Label68 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Importe préstamo cheque:"
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
               Left            =   480
               TabIndex        =   108
               Top             =   360
               Width           =   2940
            End
            Begin VB.Label Label69 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Importe compra cheque:"
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
               Left            =   705
               TabIndex        =   107
               Top             =   720
               Width           =   2715
            End
            Begin VB.Label Label70 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Valor Salario Mínimo:"
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
               Left            =   5085
               TabIndex        =   106
               Top             =   360
               Width           =   2415
            End
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ruta de Generación de Avisos:"
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
            Left            =   390
            TabIndex        =   114
            Top             =   3600
            Width           =   3390
         End
         Begin VB.Label lblTipoDe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Giro Mercantil:"
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
            Left            =   360
            TabIndex        =   113
            Top             =   1320
            Width           =   1680
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Moneda Local:"
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
            Left            =   360
            TabIndex        =   112
            Top             =   2160
            Width           =   1605
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Número de Constancia:"
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
            Left            =   360
            TabIndex        =   111
            Top             =   3120
            Width           =   2580
         End
         Begin VB.Label lblActividadDe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Actividad de la Empresa:"
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
            Left            =   360
            TabIndex        =   110
            Top             =   480
            Width           =   2940
         End
      End
      Begin VB.TextBox txtFolio 
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
         Left            =   2190
         MaxLength       =   8
         TabIndex        =   15
         Top             =   11235
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
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
         Left            =   2190
         Locked          =   -1  'True
         MaxLength       =   8
         TabIndex        =   14
         Tag             =   "Datos"
         Top             =   8355
         Visible         =   0   'False
         Width           =   1095
      End
      Begin DevPowerFlatBttn.FlatBttn cmdDesde 
         Height          =   300
         Left            =   3270
         TabIndex        =   19
         Top             =   8355
         Visible         =   0   'False
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
         Picture         =   "frmConfigLavadoDinero.frx":296F
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
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
         Left            =   1455
         TabIndex        =   21
         Top             =   11235
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Datos a partir de:"
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
         Left            =   135
         TabIndex        =   20
         Top             =   8355
         Visible         =   0   'False
         Width           =   1950
      End
   End
   Begin VB.TextBox txtNoSucursal 
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
      Left            =   240
      MaxLength       =   2
      TabIndex        =   16
      Top             =   15240
      Visible         =   0   'False
      Width           =   855
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5115
      TabIndex        =   13
      Top             =   7260
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
      Picture         =   "frmConfigLavadoDinero.frx":2A84
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3915
      TabIndex        =   12
      Top             =   7260
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
      Picture         =   "frmConfigLavadoDinero.frx":2FD6
   End
   Begin VB.Label Label57 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Precio Automóviles:"
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
      Left            =   240
      TabIndex        =   37
      Top             =   10320
      Visible         =   0   'False
      Width           =   2235
   End
   Begin VB.Label Label8 
      Caption         =   "Meses de Vencimiento de Apartado:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      TabIndex        =   17
      Top             =   15720
      Visible         =   0   'False
      Width           =   2715
   End
End
Attribute VB_Name = "frmConfigLavadoDinero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 05/04/02
' Modulo frmConfiguracion - frmConfiguracion.frm
' Ultima Modificacion - 05/04/02
''Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez

'////////////////////////////////////////////////////////////////

Option Explicit



Private Sub cmdAceptar_Click()

    If Valida Then
        
        Grabar_Configuracion
        MsgBox "Configuración guardada con éxito !!", vbInformation, "Parámetros"
    End If

End Sub

'Grabamos la configuracion de los parametros
Private Sub Grabar_Configuracion()
Dim txt As Object, Sql As String, Caracter As String, Tasa As Double, Vencimiento As Integer, VencimientoAlmoneda As Integer, crPrestamoLimite As Double, strValor As String

    
    dbDatos.Execute "UPDATE parametros SET " & _
                    "IDActividadVulnerable = " & Val(cmbActividadVulnerable.ItemData(cmbActividadVulnerable.ListIndex)) & _
                    ",IdTipoGiroMercantil=" & Val(cmbTipoGiro.ItemData(cmbTipoGiro.ListIndex)) & _
                    ",IdTipoMonedaLocal=" & Val(cmbTipoMoneda.ItemData(cmbTipoMoneda.ListIndex)) & _
                    ",PrestamoCheque=" & Val(txtPrestamoCheque) & _
                    ",CompraCheque=" & Val(txtCompraCheque) & _
                    ",ImporteSalario=" & Val(txtImporteSalario) & _
                    ",ImporteUdi=" & Val(txtImporteUdi) & _
                    ",ImporteVSMPrestamos=" & Val(txtImporteVSMPrestamos) & _
                    ",NumConstancia='" & Trim(txtNumConstancia) & "'" & _
                    ",RutaArchivosXML='" & Replace(Trim(txtRutaArchivoXML.text), "\", "\\", 1) & "'"
                    
                    
Error:
    Maneja_Error Err
    
End Sub

Private Sub cmdDatosDe_Click()
    frmCatsucursales.Show
    BringWindowToTop frmCatsucursales.hWnd
End Sub

Private Sub cmdExpedicion_Click()
    txtFechaexpedicion.text = frmCalendario.Fecha(txtFechaexpedicion.text)
End Sub


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()


 With cdlgDir
        .Flags = cdlOFNPathMustExist
        .Flags = .Flags Or cdlOFNHideReadOnly
        .Flags = .Flags Or cdlOFNNoChangeDir
        .Flags = .Flags Or cdlOFNExplorer
        .Flags = .Flags Or cdlOFNNoValidate
        .FileName = "*.*"   'Dummy File
        .CancelError = True
        '
        On Error Resume Next
        .Action = 1
        If Err = 0 Then
            txtRutaArchivoXML.text = Left(.FileName, Len(.FileName) - 4)
        Else
            MsgBox "Seleccione un Directorio."
            Exit Sub
        End If
    End With

'
'    cdlgDir.InitDir = App.Path
'
'    cdlgDir.Filter = "*.*"
'    cdlgDir.FilterIndex = 1
'    cdlgDir.flags = cdlOFNAllowMultiselect + cdlOFNExplorer
'    cdlgDir.flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
'    cdlgDir.CancelError = True
'    On Error Resume Next
'    cdlgDir.ShowOpen
'    If Err Then
'        'MsgBox "Select Folder"
'        Exit Sub
'    End If
'
'
'    txtRutaArchivoXML.text = cdlgDir.FileName
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

'inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    CentrarForm Me, frmMDI
    Frame2.BorderStyle = 0
    
    Cargar_Combos "Descripcion", "mld_actividad_vulnerable", cmbActividadVulnerable, , , False
    Cargar_Combos "Descripcion", "mld_giro_mercantil", cmbTipoGiro, , "Descripcion", False
    Cargar_Combos "Moneda", "mld_tipo_monedas", cmbTipoMoneda, " WHERE Estatus=1", , False
    
    Cargar_Configuracion
    Screen.MousePointer = vbDefault
End Sub

'Leemos los datos del archivo ini
Private Sub Cargar_Configuracion()
Dim txt As Object, i As Integer
Dim rc As New ADODB.Recordset

On Error GoTo Error
    
    rc.Open "SELECT * FROM parametros", dbDatos, adOpenForwardOnly, adLockOptimistic
    
    If Not rc.EOF Then
        
       cmbActividadVulnerable.ListIndex = ComboInformacion(cmbActividadVulnerable, rc.Fields(cmbActividadVulnerable.Tag))
       cmbTipoGiro.ListIndex = ComboInformacion(cmbTipoGiro, rc.Fields(cmbTipoGiro.Tag))
       cmbTipoMoneda.ListIndex = ComboInformacion(cmbTipoMoneda, rc.Fields(cmbTipoMoneda.Tag))
        
       txtPrestamoCheque.text = rc.Fields(txtPrestamoCheque.Tag)
       txtCompraCheque.text = rc.Fields(txtCompraCheque.Tag)
       txtImporteSalario.text = rc.Fields(txtImporteSalario.Tag)
       txtImporteUdi.text = rc.Fields(txtImporteUdi.Tag)
       txtNumConstancia.text = rc.Fields(txtNumConstancia.Tag)
       
       txtRutaArchivoXML.text = rc.Fields(txtRutaArchivoXML.Tag)
       txtImporteVSMPrestamos.text = rc.Fields(txtImporteVSMPrestamos.Tag)
       
    End If
    rc.Close
    Set rc = Nothing
Exit Sub
    
Error:
    Maneja_Error Err
    Set rc = Nothing
End Sub



Private Sub txtCompraCheque_GotFocus()
    Seleccionar_Texto txtCompraCheque
    Cambiar_Color True, txtCompraCheque
End Sub

Private Sub txtCompraCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCompraCheque_LostFocus()
    Cambiar_Color False, txtCompraCheque
End Sub


Private Sub txtImporteSalario_GotFocus()
    Seleccionar_Texto txtImporteSalario
    Cambiar_Color True, txtImporteSalario
End Sub

Private Sub txtImporteSalario_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtImporteSalario_LostFocus()
    Cambiar_Color False, txtImporteSalario
End Sub

Private Sub txtImporteUdi_GotFocus()
    Seleccionar_Texto txtImporteUdi
    Cambiar_Color True, txtImporteUdi
End Sub

Private Sub txtImporteUdi_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtImporteUdi_LostFocus()
    Cambiar_Color False, txtImporteUdi
End Sub





Function Valida() As Boolean
    
    Valida = True
    
    If cmbActividadVulnerable.ListIndex = -1 Then
        MsgBox "Seleccione la Actividad de la Empresa !!", vbCritical
        Valida = False
        cmbActividadVulnerable.SetFocus
        Exit Function
    End If
    
    If cmbTipoGiro.ListIndex = -1 Then
        MsgBox "Seleccione el Giro Mercantil de la Empresa !!", vbCritical
        Valida = False
        cmbTipoGiro.SetFocus
        Exit Function
    End If
    
    If cmbTipoMoneda.ListIndex = -1 Then
        MsgBox "Seleccione el Tipo de Moneda !!", vbCritical
        Valida = False
        cmbTipoMoneda.SetFocus
        Exit Function
    End If
    
    
    If txtPrestamoCheque.text = "" Then
        MsgBox "Introduzca el Importe de Prestamo por Cheque !!", vbCritical
        Valida = False
        txtPrestamoCheque.SetFocus
        Exit Function
    End If

    If txtCompraCheque.text = "" Then
        MsgBox "Introduzca el Importe de Compra por Cheque !!", vbCritical
        Valida = False
        txtCompraCheque.SetFocus
        Exit Function
    End If

    If txtImporteSalario.text = "" Then
        MsgBox "Introduzca el Importe de Salario !!", vbCritical
        Valida = False
        txtImporteSalario.SetFocus
        Exit Function
    End If

    If txtImporteUdi.text = "" Then
        MsgBox "Introduzca el Importe de Valor de UDI !!", vbCritical
        Valida = False
        txtImporteUdi.SetFocus
        Exit Function
    End If
    
    If txtNumConstancia.text = "" Then
        MsgBox "Introduzca el Número de Constancia !!", vbCritical
        Valida = False
        txtNumConstancia.SetFocus
        Exit Function
    End If
    
    
    
End Function

Private Sub Limpiar(Contededor As String, Optional x As Integer = 0)
Dim ctrl As Control
  
    For Each ctrl In Controls
        
        On Error Resume Next

        If ctrl.Container.Caption = Contededor Then
            If TypeOf ctrl Is TextBox Then ctrl.text = ""
            If TypeOf ctrl Is ComboBox Then ctrl.ListIndex = -1
            On Error Resume Next
            ctrl.Tag = ""
        End If

    Next

End Sub


Private Sub txtNumConstancia_GotFocus()
    Seleccionar_Texto txtNumConstancia
    Cambiar_Color True, txtNumConstancia
End Sub

Private Sub txtNumConstancia_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumConstancia_LostFocus()
    Cambiar_Color False, txtNumConstancia
End Sub

