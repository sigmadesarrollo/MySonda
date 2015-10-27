VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmDivisas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compra/Venta de Divisas"
   ClientHeight    =   8310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10035
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDivisas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8310
   ScaleWidth      =   10035
   Begin VB.Frame Frame1 
      Caption         =   "Divisas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7635
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   9855
      Begin VB.TextBox txtEfectivo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   7200
         TabIndex        =   74
         Top             =   3420
         Width           =   2535
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Left            =   0
         Top             =   90
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D4 
         Height          =   30
         Index           =   0
         Left            =   1035
         Top             =   90
         Width           =   8775
         _ExtentX        =   15478
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D3 
         Height          =   7500
         Left            =   9780
         Top             =   90
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   13229
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   7545
         Left            =   0
         Top             =   90
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   13309
         Orientation     =   0
         LineWidth       =   2
      End
      Begin VB.TextBox txtTotalDivisas 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   870
         Left            =   3855
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "0.00"
         Top             =   4530
         Width           =   5550
      End
      Begin VB.TextBox txtMonedaNacional 
         Alignment       =   2  'Center
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   3855
         Locked          =   -1  'True
         TabIndex        =   62
         Text            =   "0.00"
         Top             =   5880
         Width           =   5550
      End
      Begin VB.TextBox txtDepositoCheque 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3885
         TabIndex        =   24
         Top             =   7530
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtTraspaso 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3765
         TabIndex        =   25
         Top             =   7725
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtEmpresa 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   975
         MaxLength       =   250
         TabIndex        =   9
         Top             =   2025
         Width           =   6480
      End
      Begin VB.TextBox txtNotas 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   705
         MaxLength       =   200
         TabIndex        =   12
         Top             =   2685
         Width           =   6750
      End
      Begin VB.OptionButton opExtranjero 
         Appearance      =   0  'Flat
         Caption         =   "Extranjero"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6315
         TabIndex        =   53
         Top             =   2340
         Width           =   1140
      End
      Begin VB.OptionButton opMexicano 
         Appearance      =   0  'Flat
         Caption         =   "Mexicano"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5115
         TabIndex        =   52
         Top             =   2340
         Value           =   -1  'True
         Width           =   1140
      End
      Begin VB.TextBox txtRfc 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   3420
         MaxLength       =   30
         TabIndex        =   11
         Top             =   2340
         Width           =   1515
      End
      Begin DevPowerFlatBttn.FlatBttn cmdModificaTipoCambio 
         Height          =   255
         Left            =   7680
         TabIndex        =   50
         Top             =   240
         Visible         =   0   'False
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   450
         AlignCaption    =   4
         AlignPicture    =   2
         AutoSize        =   0   'False
         Caption         =   "..."
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
      Begin VB.TextBox txtDivisa 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   525
         Left            =   120
         TabIndex        =   49
         Top             =   3420
         Width           =   4575
      End
      Begin VB.Frame Frame2 
         Caption         =   "OPERACIÓN"
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
         Height          =   2355
         Left            =   7560
         TabIndex        =   44
         Top             =   480
         Width           =   2070
         Begin VB.OptionButton optVenta 
            Appearance      =   0  'Flat
            Caption         =   "Venta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   420
            Left            =   150
            TabIndex        =   46
            Top             =   1320
            Width           =   1785
         End
         Begin VB.OptionButton optCompra 
            Appearance      =   0  'Flat
            Caption         =   "Compra"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   420
            Left            =   150
            TabIndex        =   45
            Top             =   450
            Width           =   1785
         End
         Begin VB.Label lblVenta 
            Alignment       =   2  'Center
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   60
            TabIndex        =   72
            Top             =   1755
            Width           =   1965
         End
         Begin VB.Label lblCompra 
            Alignment       =   2  'Center
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   60
            TabIndex        =   71
            Top             =   900
            Width           =   1965
         End
      End
      Begin VB.TextBox txtCantidad 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0.00"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2058
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   5280
         TabIndex        =   14
         Top             =   3420
         Width           =   1815
      End
      Begin VB.TextBox txtTipoCambio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   21
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   9600
         TabIndex        =   13
         Top             =   3480
         Visible         =   0   'False
         Width           =   2730
      End
      Begin VB.TextBox txtCp 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   465
         MaxLength       =   5
         TabIndex        =   6
         Top             =   1710
         Width           =   1215
      End
      Begin VB.TextBox txtBuscacliente 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   795
         Locked          =   -1  'True
         MaxLength       =   60
         TabIndex        =   0
         Top             =   285
         Width           =   3375
      End
      Begin VB.TextBox txtIdentificacion 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1380
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2355
         Width           =   1515
      End
      Begin VB.TextBox txtTelefono 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   5520
         MaxLength       =   12
         TabIndex        =   8
         Top             =   1725
         Width           =   1935
      End
      Begin VB.TextBox txtEstado 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2520
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1710
         Width           =   2055
      End
      Begin VB.TextBox txtColonia 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   480
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1350
         Width           =   2775
      End
      Begin VB.TextBox txtMunicipio 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   4440
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1350
         Width           =   3015
      End
      Begin VB.TextBox txtDireccion 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   1065
         MaxLength       =   30
         TabIndex        =   3
         Top             =   990
         Width           =   6375
      End
      Begin VB.TextBox txtNombre 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   855
         MaxLength       =   20
         TabIndex        =   1
         Top             =   630
         Width           =   2775
      End
      Begin VB.TextBox txtApellidos 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   2
         Top             =   630
         Width           =   2880
      End
      Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
         Height          =   225
         Left            =   4200
         TabIndex        =   27
         Top             =   270
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   397
         AlignCaption    =   4
         AutoSize        =   0   'False
         Caption         =   "..."
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
      Begin DevPowerFlatBttn.FlatBttn cmdMosDivisa 
         Height          =   345
         Left            =   4725
         TabIndex        =   48
         Top             =   3420
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   609
         AlignCaption    =   4
         AutoSize        =   0   'False
         Caption         =   "..."
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
      Begin Line3D.ucLine3D ucLine3D4 
         Height          =   30
         Index           =   1
         Left            =   0
         Top             =   7605
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D4 
         Height          =   30
         Index           =   2
         Left            =   0
         Top             =   2985
         Width           =   9810
         _ExtentX        =   17304
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin VB.Frame Frame3 
         Caption         =   "BILLETES"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3690
         Left            =   120
         TabIndex        =   58
         Top             =   3945
         Width           =   3450
         Begin VB.TextBox txtBiUno 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   21
            Tag             =   "1"
            Text            =   "0"
            Top             =   3240
            Width           =   1395
         End
         Begin VB.TextBox txtBiDos 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1320
            MaxLength       =   1
            TabIndex        =   20
            Tag             =   "2"
            Text            =   "0"
            Top             =   2745
            Width           =   1395
         End
         Begin VB.TextBox txtBiCinco 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1320
            TabIndex        =   19
            Tag             =   "5"
            Text            =   "0"
            Top             =   2250
            Width           =   1395
         End
         Begin VB.TextBox txtBiDiez 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1320
            TabIndex        =   18
            Tag             =   "10"
            Text            =   "0"
            Top             =   1755
            Width           =   1395
         End
         Begin VB.TextBox txtBiCien 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1320
            TabIndex        =   15
            Tag             =   "100"
            Text            =   "0"
            Top             =   255
            Width           =   1395
         End
         Begin VB.TextBox txtBiCincuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1320
            TabIndex        =   16
            Tag             =   "50"
            Text            =   "0"
            Top             =   750
            Width           =   1395
         End
         Begin VB.TextBox txtBiVeinte 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   1320
            TabIndex        =   17
            Tag             =   "20"
            Text            =   "0"
            Top             =   1245
            Width           =   1395
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            TabIndex        =   70
            Top             =   3240
            Width           =   915
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            Caption         =   "2"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            TabIndex        =   69
            Top             =   2745
            Width           =   915
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "5"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            TabIndex        =   68
            Top             =   2250
            Width           =   915
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "10"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            TabIndex        =   67
            Top             =   1755
            Width           =   915
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "100"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            TabIndex        =   61
            Top             =   255
            Width           =   915
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "50"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            TabIndex        =   60
            Top             =   750
            Width           =   915
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "20"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   360
            TabIndex        =   59
            Top             =   1245
            Width           =   915
         End
      End
      Begin VB.Label lblCambio 
         AutoSize        =   -1  'True
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5280
         TabIndex        =   76
         Top             =   6960
         Visible         =   0   'False
         Width           =   945
      End
      Begin VB.Label Label2 
         Caption         =   "Efectivo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7200
         TabIndex        =   75
         Top             =   3060
         Width           =   1440
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
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
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   240
         TabIndex        =   73
         Top             =   -15
         Width           =   735
      End
      Begin VB.Label lblLeyendaCambio 
         AutoSize        =   -1  'True
         Caption         =   "CAMBIO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3855
         TabIndex        =   66
         Top             =   6960
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label lblLeyenda 
         AutoSize        =   -1  'True
         Caption         =   "MONEDA NACIONAL:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   3855
         TabIndex        =   64
         Top             =   5400
         Width           =   3765
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "TRASPASO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   57
         Top             =   7725
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "CHEQUE/DEPOSITO:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   56
         Top             =   7275
         Visible         =   0   'False
         Width           =   3255
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Empresa:"
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
         Left            =   120
         TabIndex        =   55
         Top             =   2010
         Width           =   780
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Notas:"
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
         Left            =   120
         TabIndex        =   54
         Top             =   2670
         Width           =   525
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "RFC:"
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
         Left            =   3000
         TabIndex        =   51
         Top             =   2340
         Width           =   360
      End
      Begin VB.Label Label4 
         Caption         =   "Divisa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   47
         Top             =   3060
         Width           =   1095
      End
      Begin VB.Label lblTipoventa 
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
         Left            =   8400
         TabIndex        =   43
         Top             =   2160
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "DIVISAS:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   3855
         TabIndex        =   42
         Top             =   4050
         Width           =   1650
      End
      Begin VB.Label Label9 
         Caption         =   "Cantidad:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5280
         TabIndex        =   41
         Top             =   3060
         Width           =   1440
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de cambio:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9600
         TabIndex        =   40
         Top             =   3120
         Visible         =   0   'False
         Width           =   2565
      End
      Begin VB.Label lblFolio 
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
         Height          =   375
         Left            =   5880
         TabIndex        =   39
         Top             =   240
         Visible         =   0   'False
         Width           =   1650
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4920
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label92 
         AutoSize        =   -1  'True
         Caption         =   "CP:"
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
         Left            =   120
         TabIndex        =   37
         Top             =   1710
         Width           =   255
      End
      Begin VB.Label Label115 
         AutoSize        =   -1  'True
         Caption         =   "Identificación:"
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
         Left            =   120
         TabIndex        =   36
         Top             =   2340
         Width           =   1200
      End
      Begin VB.Label Label116 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono:"
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
         Left            =   4680
         TabIndex        =   35
         Top             =   1710
         Width           =   780
      End
      Begin VB.Label Label117 
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
         Left            =   1800
         TabIndex        =   34
         Top             =   1710
         Width           =   615
      End
      Begin VB.Label Label118 
         AutoSize        =   -1  'True
         Caption         =   "Buscar:"
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
         Left            =   120
         TabIndex        =   33
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label119 
         AutoSize        =   -1  'True
         Caption         =   "Col:"
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
         Left            =   120
         TabIndex        =   32
         Top             =   1350
         Width           =   300
      End
      Begin VB.Label Label120 
         AutoSize        =   -1  'True
         Caption         =   "Municipio:"
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
         Left            =   3480
         TabIndex        =   31
         Top             =   1350
         Width           =   840
      End
      Begin VB.Label Label121 
         AutoSize        =   -1  'True
         Caption         =   "Dirección:"
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
         Left            =   120
         TabIndex        =   30
         Top             =   990
         Width           =   825
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos:"
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
         Left            =   3720
         TabIndex        =   28
         Top             =   630
         Width           =   810
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8790
      TabIndex        =   23
      Top             =   7830
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
      Picture         =   "frmDivisas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   7560
      TabIndex        =   22
      Top             =   7830
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
      Picture         =   "frmDivisas.frx":055E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCotizar 
      Height          =   375
      Left            =   5760
      TabIndex        =   65
      Top             =   7830
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "    Calcula Divisas"
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
      Picture         =   "frmDivisas.frx":0AB0
      PictureDisabled =   "frmDivisas.frx":0CB0
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   4410
      TabIndex        =   77
      Top             =   7830
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
      Picture         =   "frmDivisas.frx":0E0A
   End
End
Attribute VB_Name = "frmDivisas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
Dim Cantidad As Integer, TotalDivisas As Integer, Maximo As Double, ID As Integer, IDCliente As Long, Folio As Long, TipoDivisa As Integer, Cambio As Double, Tipo As Integer
Dim Total As Double, Movimiento As Long, Operacion As String, TipoEntrada As String, crEfectivo As Double, Hora As String

On Error GoTo Error
    
    Maximo = SacaValor("monedas", "Maximo", " WHERE Clave=" & Val(txtDivisa.Tag))
    Cantidad = IIf(Trim(txtCantidad.text) = "", 0, txtCantidad.text)
    TotalDivisas = IIf(Trim(txtTotalDivisas.text) = "", 0, txtTotalDivisas)
    IDCliente = Val(txtNombre.Tag)
    
    If IDCliente > 0 Then
        
        If Requeridos(True) Then
        
            Actualizar_Cliente IDCliente
        Else
            
            Exit Sub
        End If

    Else

        If Cantidad >= Maximo And Maximo > 0 Then
            
            If Requeridos(True) Then
                
                IDCliente = Grabar_Cliente
            Else
                
                Exit Sub
            End If
        
        End If
    
    End If

    If Requeridos(False) Then
                
        crEfectivo = 0
        TipoEntrada = 0
        Total = 0
        
        TipoDivisa = Val(txtDivisa.Tag)
        Cambio = txtTipoCambio.text
        Tipo = IIf(optCompra.Value, 0, 1)
        Total = txtTotalDivisas.text
        
        If Cantidad <> TotalDivisas Then MsgBox "Verifique su arqueo de divisas !!", vbCritical, "Compra/venta de divisas": Exit Sub
        
        'Tomo el Importe
        crEfectivo = Cantidad * Cambio
        
        'Saco el Folio
        Folio = Regresa_Movimiento(False, "FolioDivisas")
        Regresa_Movimiento True, "FolioDivisas"
        
        'Saco el Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
        
        'Tomo la Hora
        Hora = Time
        
        lblFolio.Caption = Folio
        Label1.Visible = True
        lblFolio.Visible = True
                                
        'Tabla de divisas
        dbDatos.Execute "INSERT INTO divisas (IDCliente,Folio,Fecha,IDDivisa,Importe,Cantidad,Tipo,TipoEntrada,Notas,Efectivo,ChequeDeposito,Traspaso,IDUsuario,IDSucursal,PC) VALUES (" & _
                        IDCliente & "," & Folio & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & TipoDivisa & "," & ConvMoneda(Cambio) & "," & Cantidad & "," & Tipo & "," & TipoEntrada & ",'" & Trim(txtNotas.text) & "'," & ConvMoneda(crEfectivo) & ",0,0," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & NombrePc & "')"
                
        'Grabamos el cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optCompra.Value, "CD01", "VD50") & "','" & IIf(optCompra.Value, "710301", "710350") & "'," & ConvMoneda(crEfectivo) & "," & IIf(optCompra, TIPO_CARGO, TIPO_ABONO) & ",1,'Divisas','" & Nombre_Pc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                                        
        
        'Grabamos el cargo a efectivo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optCompra.Value, "CD50", "VD01") & "','" & IIf(optCompra.Value, "110150", "110101") & "'," & ConvMoneda(crEfectivo) & "," & IIf(optCompra.Value, TIPO_ABONO, TIPO_CARGO) & ",1,'Divisas','" & Nombre_Pc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
        
'''        'Grabamos abono 199450
'''        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
'''                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optCompra.Value, "CD50", "VD01") & "','" & IIf(optCompra.Value, "199450", "199401") & "'," & ConvMoneda(crEfectivo) & "," & IIf(optCompra.Value, TIPO_ABONO, TIPO_CARGO) & ",1,'Divisas','" & Nombre_Pc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                                    
        'Muevo las Existencias de Divisas****************
        'Grabamos el cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,IDDivisa) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optCompra.Value, "CD01", "VD50") & "','" & IIf(optCompra.Value, "710301", "710350") & "'," & Cantidad & "," & IIf(optCompra, TIPO_CARGO, TIPO_ABONO) & ",2,'Divisas','" & Nombre_Pc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & TipoDivisa & ")"
                                                        
        
        'Grabamos el cargo a efectivo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,IDDivisa) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optCompra.Value, "CD50", "VD01") & "','" & IIf(optCompra.Value, "999401", "999450") & "'," & Cantidad & "," & IIf(optCompra.Value, TIPO_ABONO, TIPO_CARGO) & ",2,'Divisas','" & Nombre_Pc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & TipoDivisa & ")"
        
        '**********************************************

        'Tipo de cambio
        GrabaTipoCambio Tipo
        
        If MsgBox("Desea imprimir recibo ??", vbQuestion + vbYesNo + vbDefaultButton1, "Compra/Venta de Divisas") = vbYes Then
            
            ImprimeTicket Folio
        End If
        
        'Saco el Cambio
        frmCambio.Mostrar (crEfectivo)
        
        Limpiar "Divisas"
        Limpiar "BILLETES"
        optCompra.Value = True
        TipoCambio Val(SacaValor("monedas", "ID", " WHERE Defoult=1"))
        
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub cmdCotizar_Click()
    frmCalculaDlls.Show
    BringWindowToTop frmCalculaDlls.hWnd
End Sub

Private Sub cmdImprimir_Click()
Dim Folio As Long

    Folio = frmReimpresionrecibos.ReImprimir("divisas", "Folio", " WHERE Folio=")
    If Folio > 0 Then
        
        ImprimeTicket Folio
    
    ElseIf Folio = 0 Then
        
        MsgBox "No se encontró el folio especificado !!", vbInformation, "Compra/Venta de Divisas"
    End If

End Sub

Private Sub cmdMosCliente_Click()
    frmMostrarCliente.Ver Me, txtBuscacliente, True, 0
End Sub

Private Sub cmdMosDivisa_Click()
    frmMuestraDivisas.Posicion Me, Me.txtDivisa
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Frame1.BorderStyle = 0
    TipoCambio Val(SacaValor("monedas", "Clave", " WHERE Defoult=1"))
    optCompra.Value = True
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
End Sub

Public Sub Buscar_Cliente(ID As Long)
Dim rcClientes As New ADODB.Recordset

On Error GoTo Error
   
    rcClientes.Open "SELECT * FROM Clientes WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
    With rcClientes
        txtNombre.text = !Nombre
        txtNombre.Tag = ID
        txtApellidos.text = !Apellido
        txtDireccion.text = !Direccion
        txtColonia.text = IIf(IsNull(!Colonia), "", !Colonia)
        txtMunicipio.text = IIf(IsNull(!Municipio), "", !Municipio)
        txtEstado.text = IIf(IsNull(!Estado), "", !Estado)
        txtTelefono.text = IIf(IsNull(!Tel), "", !Tel)
        txtIdentificacion.text = IIf(IsNull(!Identificacion), "", !Identificacion)
        txtCP.text = IIf(IsNull(!CP), "", !CP)
        txtRfc.text = IIf(IsNull(!RFC), "", !RFC)
        txtEmpresa.text = IIf(IsNull(!Empresa), "", !Empresa)
        Select Case !Nacionalidad
        Case 1
            opMexicano.Value = True
        
        Case 2
            opExtranjero.Value = True
            
        End Select
    End With

    rcClientes.Close
    Set rcClientes = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcClientes = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub optCompra_Click()
    
    optVenta.BackColor = &H8000000F
    optVenta.ForeColor = &H404040
    lblVENTA.ForeColor = &H404040
    
    optCompra.BackColor = &HC000&
    optCompra.ForeColor = &HFFFFFF
    lblCOMPRA.ForeColor = &HC000&
    
    Label3.ForeColor = &HC000&
    txtTotalDivisas.ForeColor = &HC000&
    lblLeyenda.ForeColor = &HC000&
    txtMonedaNacional.ForeColor = &HC000&
    txtCantidad.ForeColor = &HC000&
    txtEfectivo.ForeColor = &HC000&
    txtDivisa.ForeColor = &HC000&
    
    Label4.ForeColor = &HC000&
    Label9.ForeColor = &HC000&
    Label2.ForeColor = &HC000&
    txtBiCien.ForeColor = &HC000&
    txtBiCincuenta.ForeColor = &HC000&
    txtBiVeinte.ForeColor = &HC000&
    txtBiDiez.ForeColor = &HC000&
    txtBiCinco.ForeColor = &HC000&
    txtBiDos.ForeColor = &HC000&
    txtBiUno.ForeColor = &HC000&
    
    Label16.ForeColor = &HC000&
    Label15.ForeColor = &HC000&
    Label10.ForeColor = &HC000&
    Label14.ForeColor = &HC000&
    Label19.ForeColor = &HC000&
    Label21.ForeColor = &HC000&
    Label22.ForeColor = &HC000&
    
    Limpiar "BILLETES"
    txtCantidad.text = "0"
    txtEfectivo.text = "0.00"
    txtTipoCambio.text = Format(lblCOMPRA.Caption, FMoneda)
    txtEfectivo.Locked = True
    Exit Sub
    
End Sub

Private Sub optVenta_Click()
    
    optCompra.BackColor = &H8000000F
    optCompra.ForeColor = &H404040
    lblCOMPRA.ForeColor = &H404040
    
    optVenta.BackColor = &HC00000
    optVenta.ForeColor = &HFFFFFF
    lblVENTA.ForeColor = &HC00000
    
    Label3.ForeColor = &HC00000
    txtTotalDivisas.ForeColor = &HC00000
    lblLeyenda.ForeColor = &HC00000
    txtMonedaNacional.ForeColor = &HC00000
    txtCantidad.ForeColor = &HC00000
    txtEfectivo.ForeColor = &HC00000
    txtDivisa.ForeColor = &HC00000
    
    Label4.ForeColor = &HC00000
    Label9.ForeColor = &HC00000
    Label2.ForeColor = &HC00000
    txtBiCien.ForeColor = &HC00000
    txtBiCincuenta.ForeColor = &HC00000
    txtBiVeinte.ForeColor = &HC00000
    txtBiDiez.ForeColor = &HC00000
    txtBiCinco.ForeColor = &HC00000
    txtBiDos.ForeColor = &HC00000
    txtBiUno.ForeColor = &HC00000
    
    Label16.ForeColor = &HC00000
    Label15.ForeColor = &HC00000
    Label10.ForeColor = &HC00000
    Label14.ForeColor = &HC00000
    Label19.ForeColor = &HC00000
    Label21.ForeColor = &HC00000
    Label22.ForeColor = &HC00000
    
    Limpiar "BILLETES"
    txtCantidad.text = "0"
    txtEfectivo.text = "0.00"
    txtTipoCambio.text = Format(lblVENTA.Caption, FMoneda)
    txtEfectivo.Locked = False
    Exit Sub
    
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

Private Sub txtBiCinco_Change()
    Calcula_Total txtBiCinco, True
End Sub

Private Sub txtBiCinco_GotFocus()
    Seleccionar_Texto txtBiCinco
    Cambiar_Color True, txtBiCinco
End Sub

Private Sub txtBiCinco_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBiCinco_LostFocus()
    Cambiar_Color False, txtBiCinco
End Sub

Private Sub txtBiDiez_Change()
    Calcula_Total txtBiDiez, True
End Sub

Private Sub txtBiDiez_GotFocus()
    Seleccionar_Texto txtBiDiez
    Cambiar_Color True, txtBiDiez
End Sub

Private Sub txtBiDiez_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBiDiez_LostFocus()
    Cambiar_Color False, txtBiDiez
End Sub

Private Sub txtBiDos_Change()
    Calcula_Total txtBiDos, True
End Sub

Private Sub txtBiDos_GotFocus()
    Seleccionar_Texto txtBiDos
    Cambiar_Color True, txtBiDos
End Sub

Private Sub txtBiDos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBiDos_LostFocus()
    Cambiar_Color False, txtBiDos
End Sub

Private Sub txtBiUno_Change()
    Calcula_Total txtBiUno, True
End Sub

Private Sub txtBiUno_GotFocus()
    Seleccionar_Texto txtBiUno
    Cambiar_Color True, txtBiUno
End Sub

Private Sub txtBiUno_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBiUno_LostFocus()
    Cambiar_Color False, txtBiUno
End Sub

Private Sub txtBuscacliente_GotFocus()
    Seleccionar_Texto txtBuscacliente
    Cambiar_Color True, txtBuscacliente
End Sub

Private Sub txtBuscacliente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBuscacliente_LostFocus()
    Cambiar_Color False, txtBuscacliente
End Sub

Private Sub txtDivisa_GotFocus()
    Seleccionar_Texto txtDivisa
    Cambiar_Color True, txtDivisa
End Sub

Private Sub txtDivisa_LostFocus()
    Cambiar_Color False, txtDivisa
End Sub

Private Sub txtEfectivo_GotFocus()
    Seleccionar_Texto txtEfectivo
    Cambiar_Color True, txtEfectivo
End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then Calcula_Divisas
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEfectivo_LostFocus()
    txtEfectivo.text = Format(txtEfectivo.text, FMoneda)
    Cambiar_Color False, txtEfectivo
End Sub

Private Sub txtTipoCambio_Change()
    Calcula_Total
End Sub

Private Sub txtTipoCambio_GotFocus()
    Seleccionar_Texto txtTipoCambio
    Cambiar_Color True, txtTipoCambio
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTipoCambio_LostFocus()
    txtTipoCambio.text = Format(txtTipoCambio.text, FMoneda)
    Cambiar_Color False, txtTipoCambio
    txtTipoCambio.Enabled = False
End Sub

Private Sub txtCantidad_GotFocus()
    Seleccionar_Texto txtCantidad
    Cambiar_Color True, txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Dim Cantidad As Integer, TipoCambio As Double

    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        
        Cantidad = 0
        TipoCambio = 0
        
        If Val(txtCantidad.text) > 0 Or Trim(txtCantidad.text) <> "" Then
            
            Cantidad = txtCantidad.text
        End If
        
        If Val(txtTipoCambio.text) > 0 Or Trim(txtTipoCambio.text) <> "" Then
            
            TipoCambio = txtTipoCambio.text
        End If
        
        txtEfectivo.text = Format(Cantidad * TipoCambio, FMoneda)
        
        Calcula_Total
    
    End If
    
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCantidad_LostFocus()
    Cambiar_Color False, txtCantidad
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

Private Sub txtCP_GotFocus()
    Seleccionar_Texto txtCP
    Cambiar_Color True, txtCP
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCP_LostFocus()
    Cambiar_Color False, txtCP
End Sub

Private Sub txtDepositoCheque_GotFocus()
    Seleccionar_Texto txtDepositoCheque
    Cambiar_Color True, txtDepositoCheque
End Sub

Private Sub txtDepositoCheque_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDepositoCheque_LostFocus()
    Cambiar_Color False, txtDepositoCheque
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

Private Sub txtDivisa_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtMonedaNacional_GotFocus()
    Seleccionar_Texto txtMonedaNacional
    Cambiar_Color True, txtMonedaNacional
End Sub

Private Sub txtMonedaNacional_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn And optVenta.Value Then Calcula_Divisas
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMonedaNacional_LostFocus()
    Cambiar_Color False, txtMonedaNacional
End Sub

Private Sub txtEmpresa_GotFocus()
    Seleccionar_Texto txtEmpresa
    Cambiar_Color True, txtEmpresa
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEmpresa_LostFocus()
    Cambiar_Color False, txtEmpresa
End Sub

Private Sub txtEstado_GotFocus()
    Seleccionar_Texto txtEstado
    Cambiar_Color True, txtEstado
End Sub

Private Sub txtEstado_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
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

Private Sub txtRfc_GotFocus()
    Seleccionar_Texto txtRfc
    Cambiar_Color True, txtRfc
End Sub

Private Sub txtRfc_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtRfc_LostFocus()
    Cambiar_Color False, txtRfc
End Sub

Private Sub txtTelefono_GotFocus()
    Seleccionar_Texto txtTelefono
    Cambiar_Color True, txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTelefono_LostFocus()
    Cambiar_Color False, txtTelefono
End Sub

Public Sub TipoCambio(ID As Integer)
Dim rcConsulta As New ADODB.Recordset
Dim IDDivisa As Long

On Error GoTo Error
    
    lblCOMPRA.Caption = Format(0, FMoneda)
    lblVENTA.Caption = Format(0, FMoneda)

    rcConsulta.Open "SELECT MAX(ID) AS Maximo FROM cotizaciones WHERE IDMoneda=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF And Not IsNull(rcConsulta!Maximo) Then
        
        lblCOMPRA.Caption = Format(SacaValor("cotizaciones", "Compra", " WHERE ID=" & rcConsulta!Maximo), FMoneda)
        lblVENTA.Caption = Format(SacaValor("cotizaciones", "Venta", " WHERE ID=" & rcConsulta!Maximo), FMoneda)
        If optCompra.Value Then optCompra_Click Else optVenta_Click
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    
    txtDivisa.text = SacaValor("monedas", "Descripcion", " WHERE Clave=" & ID)
    txtDivisa.Tag = ID
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Function Requeridos(Mayor As Boolean) As Boolean
Dim Cambio As Double, Cantidad As Integer

    Requeridos = True

    If Trim(txtNombre.text) = "" And Mayor = True Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Requeridos = False
        txtNombre.SetFocus
        Exit Function
    End If

    'si no tiene apellido
    If Trim(txtApellidos.text) = "" And Mayor = True Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Requeridos = False
        txtApellidos.SetFocus
        Exit Function
    End If

    'si no tiene direccion
    If Trim(txtDireccion.text) = "" And Mayor = True Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Requeridos = False
        txtDireccion.SetFocus
        Exit Function
    End If

    'si no tiene estado
    If Trim(txtEstado.text) = "" And Mayor = True Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Requeridos = False
        txtEstado.SetFocus
        Exit Function
    End If

    'si no tiene colonia
    If Trim(txtColonia.text) = "" And Mayor = True Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Requeridos = False
        txtColonia.SetFocus
        Exit Function
    End If

    'si no tiene municipio
    If Trim(txtMunicipio.text) = "" And Mayor = True Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Requeridos = False
        txtMunicipio.SetFocus
        Exit Function
    End If

    'si no tiene cp
    If Trim(txtCP.text) = "" And Mayor = True Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Requeridos = False
        txtCP.SetFocus
        Exit Function
    End If

    'si no identificacion
    If Trim(txtIdentificacion.text) = "" And Mayor = True Then
        MsgBox "Datos incompletos, favor de llenar completamente los datos", vbCritical + vbOKOnly
        Requeridos = False
        txtIdentificacion.SetFocus
        Exit Function
    End If

    If txtDivisa.Tag = "" Then
        MsgBox "Seleccione el tipo de Divisa !!", vbInformation, "Compra/Venta de Divisas"
        Requeridos = False
        txtDivisa.SetFocus
        Exit Function
    End If

    If txtTipoCambio.text <> "" Then Cambio = txtTipoCambio.text
    If txtTipoCambio.text = "" Or Cambio <= 0 Then
        MsgBox "Introduzca el Tipo de Cambio !!", vbInformation, "Compra/Venta de Divisas"
        Requeridos = False
        txtTipoCambio.SetFocus
        Exit Function
    End If

    If txtCantidad.text <> "" Then Cantidad = txtTipoCambio.text
    If txtCantidad.text = "" Or Cantidad = 0 Then
        MsgBox "Introduzca la cantidad de Divisas !!", vbInformation, "Compra/Venta de Divisas"
        Requeridos = False
        txtCantidad.SetFocus
    End If

End Function

Private Function Grabar_Cliente() As Long
    
On Error GoTo Error
         
    dbDatos.Execute "INSERT INTO clientes (Nombre,Apellido,Iniciales,Direccion,Colonia,Municipio,Estado,Tel,Identificacion,CP,Rfc,Nacionalidad,Empresa,FecRegistro) VALUES " & "('" & _
                    Trim(txtNombre.text) & "','" & Trim(txtApellidos.text) & "','" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "','" & txtDireccion.text & "','" & txtColonia.text & "','" & txtMunicipio.text & "','" & txtEstado.text & "','" & txtTelefono.text & "','" & txtIdentificacion.text & "','" & txtCP.text & "','" & Trim(txtRfc.text) & "'," & IIf(opMexicano.Value, 1, 2) & ",'" & Trim(txtEmpresa.text) & "','" & Format(Date, "YYYY/MM/DD") & "')"
    
    Grabar_Cliente = SacaValor("clientes", "MAX(ID)")
    Exit Function
    
Error:
    Maneja_Error Err
End Function

'Actualizamos los datos del cliente
Private Sub Actualizar_Cliente(ID As Long)

On Error GoTo Error

    dbDatos.Execute "UPDATE Clientes SET Iniciales='" & Iniciales(Trim(txtNombre.text), Trim(txtApellidos.text)) & "',Direccion='" & txtDireccion.text & "',Colonia='" & txtColonia.text & "',Municipio='" & txtMunicipio.text & "'," & "Estado='" & txtEstado.text & "',Tel='" & txtTelefono.text & "',Identificacion='" & txtIdentificacion.text & "',CP='" & txtCP.text & "',RFC='" & Trim(txtRfc.text) & "',Nacionalidad=" & IIf(opMexicano.Value, 1, 2) & ",Empresa='" & Trim(txtEmpresa.text) & "' WHERE ID = " & ID
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Sub ImprimeTicket(Folio As Long)
Dim ImprDefault As Boolean

On Error GoTo Error
    
    'Checo si hay impresora por default
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))

    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .ReportFileName = Path & "\Reportes\NotaDivisas.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{divisas.TipoEntrada}=0 AND {divisas.Folio}=" & Folio
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
        
        .WindowTitle = "Recibo"
        .WindowState = crptMaximized
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub txtTotalDivisas_GotFocus()
    Seleccionar_Texto txtTotalDivisas
    Cambiar_Color True, txtTotalDivisas
End Sub

Private Sub txtTotalDivisas_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTotalDivisas_LostFocus()
    Cambiar_Color False, txtTotalDivisas
End Sub

Private Sub txtTraspaso_GotFocus()
    Seleccionar_Texto txtTraspaso
    Cambiar_Color True, txtTraspaso
End Sub

Private Sub txtTraspaso_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTraspaso_LostFocus()
    Cambiar_Color False, txtTraspaso
End Sub

Private Sub GrabaTipoCambio(Tipo As Integer)

On Error GoTo Error

    If Tipo = 1 Then
        
        dbDatos.Execute "UPDATE monedas SET Venta=" & ConvMoneda(txtTipoCambio.text) & " WHERE Clave=" & Val(txtDivisa.Tag)
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub txtBicien_Change()
    Calcula_Total txtBiCien, True
End Sub

Private Sub txtBicien_GotFocus()
    Seleccionar_Texto txtBiCien
    Cambiar_Color True, txtBiCien
End Sub

Private Sub txtBicien_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBicien_LostFocus()
    Cambiar_Color False, txtBiCien
End Sub

Private Sub txtBicincuenta_Change()
    Calcula_Total txtBiCincuenta, True
End Sub

Private Sub txtBicincuenta_GotFocus()
    Seleccionar_Texto txtBiCincuenta
    Cambiar_Color True, txtBiCincuenta
End Sub

Private Sub txtBicincuenta_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBicincuenta_LostFocus()
    Cambiar_Color False, txtBiCincuenta
End Sub

Private Sub txtBiveinte_Change()
    Calcula_Total txtBiVeinte, True
End Sub

Private Sub txtBiveinte_GotFocus()
    Seleccionar_Texto txtBiVeinte
    Cambiar_Color True, txtBiVeinte
End Sub

Private Sub txtBiveinte_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBiveinte_LostFocus()
    Cambiar_Color False, txtBiVeinte
End Sub

Sub Calcula_Total(Optional text As TextBox, Optional Bandera As Boolean = False, Optional TotalDivisas As Integer = 0, Optional IsVenta As Boolean = False, Optional crImporteTotal As Double = 0)
Dim txt As Object, crTotal As Double, TipoCambio As Double, Divisas As Double, Cantidad As Integer, crEfectivo As Double

    crTotal = 0
    TipoCambio = 0
    Divisas = 0
    crEfectivo = 0
    
    If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
        
        crEfectivo = CDbl(txtEfectivo.text)
    Else
        
        crEfectivo = 0
    End If
    
    If Val(txtTipoCambio.text) > 0 Or Trim(txtTipoCambio.text) <> "" Then
        
        TipoCambio = CDbl(txtTipoCambio.text)
    Else
        
        TipoCambio = 0
    End If
    
    If Val(txtCantidad.text) > 0 Or Trim(txtCantidad.text) <> "" Then
        
        Cantidad = txtCantidad.text
    Else
        
        Cantidad = 0
    End If
    
    If (Divisas > Cantidad) And Trim(txtCantidad.text) <> "" Then
        
        MsgBox "Verifique la cantidad de divisas !!", vbCritical, "Compra/Venta de Divisas"
        If Bandera Then text.text = "0"
        Seleccionar_Texto text
        Exit Sub
    
    End If
         
    txtTotalDivisas.text = Format(Arqueo_Divisas, FMoneda)
    txtMonedaNacional.text = "$" & Format(Cantidad * TipoCambio, FMoneda)
        
End Sub

Sub Calcula_Divisas()
Dim TipoCambio As Double, crEfectivo As Double, Cantidad As Integer
    
    Cantidad = 0
    TipoCambio = 0
    crEfectivo = 0
    
    If Val(txtTipoCambio.text) > 0 Or Trim(txtTipoCambio.text) <> "" Then
        
        TipoCambio = CDbl(txtTipoCambio.text)
    End If
    
    If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
        
        crEfectivo = CDbl(txtEfectivo.text)
    End If
    
    Cantidad = Int(crEfectivo / TipoCambio)
    txtCantidad.text = Cantidad
    txtMonedaNacional.text = "$" & Format(Cantidad * TipoCambio, FMoneda)
    Saca_Cambio
End Sub

Function Arqueo_Divisas() As Double
Dim Divisas As Integer, txt As Object

    For Each txt In Me.Controls
    
        If TypeOf txt Is TextBox And txt.Tag <> "" And txt.Name <> "txtDivisa" And txt.Name <> "txtNombre" Then
            
            Divisas = Divisas + (Val(txt.Tag) * IIf(txt.text = "", 0, txt.text))
        
        End If

    Next
    
    Arqueo_Divisas = Divisas
    
End Function

Sub Saca_Cambio()
Dim crEfectivo As Double, crTotal As Double
    
    lblLeyendaCambio.Visible = False
    lblCambio.Visible = False
    
    If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
        
        crEfectivo = CDbl(txtEfectivo.text)
    Else
        
        crEfectivo = 0
    End If
    
    If Val(txtMonedaNacional.text) > 0 Or Trim(txtMonedaNacional.text) <> "" Then
        
        crTotal = CDbl(txtMonedaNacional.text)
    Else
        
        crTotal = 0
    End If
    
    If (crEfectivo - crTotal) > 0 Then
        lblCambio.Caption = "$" & Format(crEfectivo - crTotal, FMoneda)
        lblLeyendaCambio.Visible = True
        lblCambio.Visible = True
    End If
    
End Sub

Private Sub Limpiar(Contededor As String)
Dim ctrl As Control
  
    For Each ctrl In Controls
        
        On Error Resume Next

        If ctrl.Container.Caption = Contededor Then
            
            If TypeOf ctrl Is TextBox And Contededor <> "BILLETES" Then
            
                ctrl.text = ""
                ctrl.Tag = ""
            Else
                
                ctrl.text = "0"
            End If
            
            If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
            On Error Resume Next
            
        End If

    Next

End Sub
