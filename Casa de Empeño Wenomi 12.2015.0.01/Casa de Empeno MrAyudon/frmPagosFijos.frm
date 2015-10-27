VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmPagosFijos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos Fijos"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12945
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPagosFijos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   12945
   Begin VB.TextBox txtEfectivo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4425
      TabIndex        =   44
      Text            =   "0.00"
      Top             =   6750
      Width           =   4095
   End
   Begin VB.TextBox txtCambio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Left            =   8685
      Locked          =   -1  'True
      TabIndex        =   45
      Text            =   "0.00"
      Top             =   6750
      Width           =   4095
   End
   Begin VB.TextBox txtTotalGeneral 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   675
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   43
      Text            =   "0.00"
      Top             =   6750
      Width           =   4095
   End
   Begin Line3D.ucLine3D ucLine3D11 
      Height          =   1815
      Left            =   7545
      Top             =   9675
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3201
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D10 
      Height          =   30
      Left            =   180
      Top             =   9315
      Visible         =   0   'False
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D9 
      Height          =   30
      Index           =   0
      Left            =   195
      Top             =   8250
      Visible         =   0   'False
      Width           =   2670
      _ExtentX        =   4710
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D8 
      Height          =   30
      Left            =   4860
      Top             =   10035
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D7 
      Height          =   1815
      Left            =   6195
      Top             =   9690
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3201
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D6 
      Height          =   30
      Left            =   4875
      Top             =   9675
      Visible         =   0   'False
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D5 
      Height          =   1815
      Left            =   4830
      Top             =   9675
      Visible         =   0   'False
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3201
      Orientation     =   0
      LineWidth       =   2
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCalendario 
      Height          =   3975
      Left            =   15
      TabIndex        =   3
      Top             =   2340
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   7011
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
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
   Begin VB.Frame Frame1 
      Caption         =   "DATOS CONTRATO"
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
      Height          =   2220
      Left            =   45
      TabIndex        =   1
      Top             =   90
      Width           =   11910
      Begin VB.TextBox txtContrato 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   975
      End
      Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
         Height          =   375
         Left            =   2115
         TabIndex        =   22
         Top             =   300
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
         Picture         =   "frmPagosFijos.frx":000C
      End
      Begin Line3D.ucLine3D ucLine3D4 
         Height          =   30
         Left            =   9405
         Top             =   2130
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   30
         Left            =   9420
         Top             =   120
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   2040
         Index           =   0
         Left            =   9405
         Top             =   120
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   3598
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   2040
         Index           =   1
         Left            =   11640
         Top             =   120
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   3598
         Orientation     =   0
         LineWidth       =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Préstamo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   11
         Left            =   10005
         TabIndex        =   40
         Top             =   1410
         Width           =   1350
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   10
         Left            =   10095
         TabIndex        =   39
         Top             =   90
         Width           =   840
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   9495
         TabIndex        =   38
         Top             =   375
         Width           =   2055
      End
      Begin VB.Label lblPrestamo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   9495
         TabIndex        =   37
         Top             =   1725
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Vencimiento"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Index           =   12
         Left            =   9645
         TabIndex        =   36
         Top             =   735
         Width           =   1785
      End
      Begin VB.Label lblVencimiento 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   375
         Left            =   9495
         TabIndex        =   35
         Top             =   1035
         Width           =   2055
      End
      Begin VB.Label lblIdentificacion 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7005
         TabIndex        =   21
         Top             =   1800
         Width           =   1875
      End
      Begin VB.Label Label1 
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
         Index           =   9
         Left            =   5640
         TabIndex        =   20
         Top             =   1800
         Width           =   1305
      End
      Begin VB.Label lblTelefono 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   3840
         TabIndex        =   19
         Top             =   1800
         Width           =   1680
      End
      Begin VB.Label Label1 
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
         Index           =   8
         Left            =   2955
         TabIndex        =   18
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1080
         TabIndex        =   17
         Top             =   1800
         Width           =   1785
      End
      Begin VB.Label Label1 
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
         Index           =   7
         Left            =   120
         TabIndex        =   16
         Top             =   1800
         Width           =   690
      End
      Begin VB.Label lblCP 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   7410
         TabIndex        =   15
         Top             =   1440
         Width           =   1470
      End
      Begin VB.Label Label1 
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
         Index           =   6
         Left            =   7080
         TabIndex        =   14
         Top             =   1440
         Width           =   300
      End
      Begin VB.Label Label1 
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
         Index           =   5
         Left            =   3600
         TabIndex        =   13
         Top             =   1440
         Width           =   930
      End
      Begin VB.Label lblMunicipio 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   4560
         TabIndex        =   12
         Top             =   1440
         Width           =   2385
      End
      Begin VB.Label lblColonia 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1080
         TabIndex        =   11
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Colonia:"
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
         Index           =   4
         Left            =   120
         TabIndex        =   10
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblDireccion 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1080
         TabIndex        =   9
         Top             =   1080
         Width           =   7800
      End
      Begin VB.Label Label1 
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
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblApellidos 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   5400
         TabIndex        =   7
         Top             =   720
         Width           =   3480
      End
      Begin VB.Label Label1 
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
         Index           =   2
         Left            =   4440
         TabIndex        =   6
         Top             =   720
         Width           =   885
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   1080
         TabIndex        =   5
         Top             =   720
         Width           =   3240
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
         Index           =   1
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   900
      End
      Begin VB.Label Label1 
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
         Index           =   0
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   765
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Height          =   2040
         Left            =   9435
         TabIndex        =   41
         Top             =   135
         Width           =   2205
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   11670
      TabIndex        =   23
      Top             =   7605
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
      Picture         =   "frmPagosFijos.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   10530
      TabIndex        =   47
      Top             =   7605
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
      Picture         =   "frmPagosFijos.frx":08E3
   End
   Begin Line3D.ucLine3D ucLine3D9 
      Height          =   30
      Index           =   1
      Left            =   4830
      Top             =   10770
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D9 
      Height          =   30
      Index           =   2
      Left            =   195
      Top             =   8970
      Visible         =   0   'False
      Width           =   2685
      _ExtentX        =   4736
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   9240
      TabIndex        =   42
      Top             =   7605
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
      Picture         =   "frmPagosFijos.frx":0E35
   End
   Begin vbalIml6.vbalImageList lstIcons 
      Left            =   8520
      Top             =   7485
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   2296
      Images          =   "frmPagosFijos.frx":1387
      Version         =   131072
      KeyCount        =   2
      Keys            =   "ÿ"
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "TOTAL:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   120
      TabIndex        =   49
      Top             =   6360
      Width           =   1095
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "EFECTIVO:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4380
      TabIndex        =   48
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label4 
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
      Left            =   8640
      TabIndex        =   46
      Top             =   6360
      Width           =   1335
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Height          =   255
      Left            =   1590
      TabIndex        =   34
      Top             =   9030
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pago:"
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
      Left            =   240
      TabIndex        =   33
      Top             =   9045
      Visible         =   0   'False
      Width           =   525
   End
   Begin VB.Label lblMoratorios 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6270
      TabIndex        =   32
      Top             =   10830
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Moratorios:"
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
      Left            =   240
      TabIndex        =   31
      Top             =   8685
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblAmortizacion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6270
      TabIndex        =   30
      Top             =   10110
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amortización:"
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
      Left            =   4920
      TabIndex        =   29
      Top             =   10125
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label lblIntereses 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6270
      TabIndex        =   28
      Top             =   10470
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label lblNumPago 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   6270
      TabIndex        =   27
      Top             =   9735
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Intereses:"
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
      Left            =   240
      TabIndex        =   26
      Top             =   8325
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num. Pago:"
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
      Left            =   4920
      TabIndex        =   24
      Top             =   9750
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label Label3 
      BackColor       =   &H00808080&
      Height          =   1785
      Left            =   4890
      TabIndex        =   25
      Top             =   9705
      Visible         =   0   'False
      Width           =   2655
   End
End
Attribute VB_Name = "frmPagosFijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
Dim i As Integer, Contrato As Long, Folio As Long, Iva As Double, crAmortizacion As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double, Movimiento As Long, strIniciales As String, FolioRecibo As Long
Dim Ban As Boolean, PagosRealizados As Integer, IDUsuario As Long, UltimoPago As Boolean, NumPagos As Integer, VencimientoPago As Date, UltimoVencimiento As Date, Vencimiento As Date, crEfectivo As Double, crPrestamoPrenda As Double, crTotAmortizacion As Double, Bonificacion As Double, Hora As String
Dim rcPrendas As New ADODB.Recordset

On Error GoTo Error

    Screen.MousePointer = vbHourglass
    Ban = False
    PagosRealizados = 0
    NumPagos = 0
    crEfectivo = 0
    UltimoPago = False
    
    With grdCalendario
                        
        'Checo si algún pago esta marcado
        For i = 1 To grdCalendario.Rows
            
            If .CellIcon(i, 1) > 0 Then Ban = True: Exit For
        
        Next i
        
        If Ban Then
            
            'Checo que no este salteado algún pago
            If ValidaPagos = False Then Screen.MousePointer = vbDefault: Exit Sub
            
            'Tomo el Efectivo
            crEfectivo = CDbl(txtEfectivo.text)
            
            'Tomo el Número de Contrato
            Contrato = Val(lblNombre.Tag)
            
            'Tomo el Número de Folio
            Folio = Val(lblCP.Tag)
            
            'Tomo el IVA
            Iva = Val(lblDireccion.Tag) / 100
            
            'Tomo el Usuario
            IDUsuario = frmMDI.IDUsuario
            
            'Saco el Folio del Recibo
            FolioRecibo = Regresa_Movimiento(False, "FolioNotas")
            Regresa_Movimiento True, "FolioNotas"
            
            'Tomo el Numero de Pagos Realizados
            PagosRealizados = Val(SacaValor("pagosfijos", "SUM(pagado)", " WHERE Cancelado=0 AND IDEmpeno=" & Val(lblApellidos.Tag)))
            
            'Saco las Iniciales
            strIniciales = Iniciales(Trim(lblNombre.Caption), Trim(lblApellidos.Caption))
            
            'Tomo la fecha de vencimiento del último pago
            UltimoVencimiento = CDate(.CellText(.Rows, 2))
            
            For i = 1 To .Rows
                
                If .CellIcon(i, 1) > 0 Then
                    
                    'Saco el Movimiento
                    Movimiento = Regresa_Movimiento(False)
                    Regresa_Movimiento True
                    
                    'Tomo la Hora
                    Hora = Time
                    
                    Bonificacion = 0
                    crAmortizacion = 0
                    crIntereses = 0
                    crAlmacenaje = 0
                    crSeguro = 0
                    crMoratorios = 0
                    crIva = 0
                                                                      
                    'Checo si tiene Bonificación
                    Bonificacion = CDbl(.CellText(i, 14)) / 100
                    
                    'Tomo el Vencimiento del Pago
                    VencimientoPago = CDate(.CellText(i, 2))
                    
                    'Tomo la Amortización
                    crAmortizacion = CDbl(.CellText(i, 9))
                    
                    'Tomo los Intereses
                    crIntereses = Redondeo(CDbl(.CellText(i, 4)))
                    crIntereses = Redondeo(crIntereses - (crIntereses * Bonificacion))
                    
                    'Tomo el Almacenaje
                    crAlmacenaje = Redondeo(CDbl(.CellText(i, 5)))
                    crAlmacenaje = Redondeo(crAlmacenaje - (crAlmacenaje * Bonificacion))
                    
                    'Tomo el Seguro
                    crSeguro = Redondeo(CDbl(.CellText(i, 6)))
                    crSeguro = Redondeo(crSeguro - (crSeguro * Bonificacion))
                    
                    'Tomo los Moratorios
                    crMoratorios = Redondeo(CDbl(.CellText(i, 8)) / (1 + Iva))
                    
                    'Saco el Iva
                    crIva = Redondeo(CDbl(.CellText(i, 7)))
                    crIva = Redondeo(crIva - (crIva * Bonificacion))
                                                
                    'Marco como pagado el pago
                    dbDatos.Execute "UPDATE pagosfijos SET Pagado=1,FolioRecibo=" & FolioRecibo & ",Movimiento=" & Movimiento & ",FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Moratorios=" & ConvMoneda(.CellText(i, 8)) & ",Bonificacion=" & ConvMoneda(Bonificacion * 100) & ",IDUsuario=" & IDUsuario & ",Efectivo=" & ConvMoneda(crEfectivo) & " WHERE ID=" & Val(.CellItemData(i, 1))
                            
                    'Grabo el Monto de la Amortización
                    'Grabamos el cargo
                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crAmortizacion) & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                    'Grabamos el abono
                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','201750'," & ConvMoneda(crAmortizacion) & "," & TIPO_ABONO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    '*********************************
                    
                    If crIntereses > 0 Then
                        
                        'Grabamos el cargo
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crIntereses) & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                        'Grabamos el abono
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','520450'," & ConvMoneda(crIntereses) & "," & TIPO_ABONO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                    End If
                    
                    If crAlmacenaje > 0 Then
                        
                        'Grabamos el cargo
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crAlmacenaje) & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                        'Grabamos el abono
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','670350'," & ConvMoneda(crAlmacenaje) & "," & TIPO_ABONO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                    End If
                    
                    If crSeguro > 0 Then
                        
                        'Grabamos el cargo
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crSeguro) & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                        'Grabamos el abono
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','680350'," & ConvMoneda(crSeguro) & "," & TIPO_ABONO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                    End If
                    
                    If crMoratorios > 0 Then
                        
                        'Grabamos el cargo
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crMoratorios) & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                        'Grabamos el abono
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','690350'," & ConvMoneda(crMoratorios) & "," & TIPO_ABONO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                    End If
                    
                    If crIva > 0 Then
                        
                        'Grabamos el cargo
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crIva) & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                        'Grabamos el abono
                        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','120150'," & ConvMoneda(crIva) & "," & TIPO_ABONO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                    End If
            
'''                    'Grabamos abono 199401
'''                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'''                                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','199401'," & ConvMoneda(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva + crAmortizacion) & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
'''                    'Grabamos abono 110101
'''                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'''                                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pagos Fijos'," & Movimiento & "," & Folio & ",'" & strIniciales & "','110101'," & ConvMoneda(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva + crAmortizacion) & "," & TIPO_CARGO & "," & SERIE_C & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                    crTotAmortizacion = crTotAmortizacion + crAmortizacion
                    PagosRealizados = PagosRealizados + 1
                    NumPagos = NumPagos + 1
                End If
                
            Next i
                        
            'Se recalcula el préstamo
            rcPrendas.Open "SELECT d.ID AS IDPrenda,d.Prestamo,empeno.Prestamo AS PrestamoContrato FROM detallesempeno d INNER JOIN empeno ON d.IDEmpeno=empeno.ID WHERE d.IDEmpeno=" & Val(lblApellidos.Tag) & " ORDER BY d.ID", dbDatos, adOpenForwardOnly, adLockOptimistic
            While Not rcPrendas.EOF
                
                crPrestamoPrenda = (rcPrendas!Prestamo * 100) / rcPrendas!PrestamoContrato
                crPrestamoPrenda = Redondeo((rcPrendas!PrestamoContrato - crTotAmortizacion) * (crPrestamoPrenda / 100))
                
                dbDatos.Execute "UPDATE detallesempeno SET Prestamo=" & ConvMoneda(crPrestamoPrenda) & " WHERE ID=" & rcPrendas!IDPrenda
            rcPrendas.MoveNext
            Wend
            rcPrendas.Close
                                    
            'Actualizo la nueva Fecha de Vencimiento
            Vencimiento = DateAdd("M", 2, VencimientoPago)
            If Vencimiento > UltimoVencimiento Then Vencimiento = UltimoVencimiento
            dbDatos.Execute "UPDATE empeno SET Vencimiento='" & Format(Vencimiento, "YYYY/MM/DD") & "',Prestamo=Prestamo-" & ConvMoneda(crTotAmortizacion) & " WHERE ID=" & Val(lblApellidos.Tag)
                        
            'Checo si ya se pagaron todos los recibos
            If Val(lblColonia.Tag) <= PagosRealizados Then
            
                UltimoPago = True
                crAmortizacion = 0
                crIntereses = 0
                crAlmacenaje = 0
                crSeguro = 0
                crMoratorios = 0
                crIva = 0
                
                crAmortizacion = SacaValor("pagosfijos", "SUM(Amortizacion)", " WHERE IDEmpeno=" & Val(lblApellidos.Tag))
                crIntereses = SacaValor("pagosfijos", "SUM(Interes)", " WHERE IDEmpeno=" & Val(lblApellidos.Tag)) / (1 + Iva)
                crAlmacenaje = SacaValor("pagosfijos", "SUM(Almacenaje)", " WHERE IDEmpeno=" & Val(lblApellidos.Tag)) / (1 + Iva)
                crSeguro = SacaValor("pagosfijos", "SUM(Seguro)", " WHERE IDEmpeno=" & Val(lblApellidos.Tag)) / (1 + Iva)
                crMoratorios = SacaValor("pagosfijos", "SUM(Moratorios)", " WHERE IDEmpeno=" & Val(lblApellidos.Tag)) / (1 + Iva)
                crIva = (crIntereses + crAlmacenaje + crSeguro + crMoratorios) * (Iva)
                
                dbDatos.Execute "UPDATE empeno SET PC='" & NombrePc & "',Pago=" & ConvMoneda(crAmortizacion) & ",Intereses=" & ConvMoneda(crIntereses) & ",Destino=" & D_DESEMPEÑO & ",FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Pagado=1,ImporteIva=" & ConvMoneda(crIva) & ",ImporteAlmacenaje=" & ConvMoneda(crAlmacenaje) & ",ImporteSeguro=" & ConvMoneda(crSeguro) & ",ImporteMoratorios=" & ConvMoneda(crMoratorios) & ",FolioNota=" & FolioRecibo & " WHERE ID=" & Val(lblApellidos.Tag)
            End If
            
            'Imprimo el Recibo
            Imprimir_Nota Val(lblNombre.Tag), FolioRecibo, UltimoPago

        End If
        
    End With
        
    'Limpio la forma y vuelvo a cargar el calendario de pagos
    MuestraPagos Val(lblApellidos.Tag)
    LimpiaPagos
    txtTotalGeneral.text = "0.00"
    txtEfectivo.text = "0.00"
    txtContrato.SetFocus
    Screen.MousePointer = vbDefault
    
Error:
    Maneja_Error Err
    
End Sub

Private Sub cmdBuscar_Click()
Dim Contrato As Long

    If Trim(txtContrato.text) = "" Then
        
        MsgBox "Introduzca el número de contrato que desea consultar !!", vbInformation, "Pagos Fijos"
    
    Else
        
        Contrato = Val(txtContrato.text)
        LimpiaPagos
        grdCalendario.Clear
        Limpiar "DATOS CONTRATO"
        txtContrato.text = Contrato
        Buscar Val(txtContrato.text)
    
    End If
    
End Sub

Private Sub cmdImprimir_Click()
    If grdCalendario.Rows > 0 And grdCalendario.SelectedRow > 0 Then
        
        If Val(grdCalendario.CellItemData(grdCalendario.SelectedRow, 3)) > 0 Then
        
            'Imprimo el Recibo
            Imprimir_Nota Val(lblNombre.Tag), grdCalendario.CellItemData(grdCalendario.SelectedRow, 3), False
        Else
            
            MsgBox "Seleccione un pago que ya se encuentre registrado !!", vbInformation, "Pagos Fijos"
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CrearEncabezado
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Sub CrearEncabezado()

    With grdCalendario
        .ImageList = lstIcons
        .AddColumn "C1", "Num. Pago", ecgHdrTextALignCentre, , 69, , , , , , , CCLSortNumeric
        .AddColumn "C2", "Vencimiento", ecgHdrTextALignCentre, , 70, , , , , "DD/MMM/YY", , CCLSortDate
        .AddColumn "C3", "Interés", ecgHdrTextALignRight, , 75, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C4", "Intereses", ecgHdrTextALignRight, , 70, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C5", "Almacenaje", ecgHdrTextALignRight, , 70, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C6", "Seguro", ecgHdrTextALignRight, , 70, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C7", "IVA", ecgHdrTextALignRight, , 60, True, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C8", "Moratorios", ecgHdrTextALignRight, , 69, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C9", "Amortización", ecgHdrTextALignRight, , 74, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C10", "Pago", ecgHdrTextALignRight, , 73, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C11", "Saldo", ecgHdrTextALignRight, , 70, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C12", "Status", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
        .AddColumn "C13", "Fecha Mov.", ecgHdrTextALignCentre, , 80, , , , , "DD/MMM/YY", , CCLSortDate
        .AddColumn "C14", "Bonificación", ecgHdrTextALignCentre, , 69, False, , , , FMoneda, , CCLSortDate
        .AddColumn "C15", "Bonificación", ecgHdrTextALignCentre, , 70, False, , , , FMoneda, , CCLSortDate
        
        
    End With

End Sub

Sub Buscar(Contrato As Long)
Dim rcConsulta As New ADODB.Recordset

    rcConsulta.Open "SELECT empeno.ID,empeno.Fecha,empeno.Vencimiento,empeno.Prestamo,empeno.PrestamoInicial,empeno.NumContrato,empeno.Folio,empeno.IVA,empeno.VenPeriodo,empeno.TipoTasa,clientes.Nombre,clientes.Apellido,clientes.Direccion,clientes.Colonia,clientes.Municipio,clientes.Estado,clientes.Tel,clientes.CP,clientes.Identificacion " _
                    & "FROM empeno INNER JOIN clientes ON empeno.IDCliente=clientes.ID WHERE empeno.Cancelado=0 AND empeno.TipoInteres='FIJA' AND empeno.Destino<>4 AND empeno.NumContrato=" & Contrato, dbDatos, adOpenForwardOnly, adLockOptimistic
    
    With rcConsulta
                    
        If Not .BOF And Not .EOF Then
        
            lblNombre.Caption = !Nombre
            lblNombre.Tag = !NumContrato
            lblApellidos.Caption = !Apellido
            lblApellidos.Tag = !ID
            lblDireccion.Caption = !Direccion
            lblDireccion.Tag = !Iva
            lblColonia.Caption = !Colonia
            lblColonia.Tag = !VenPeriodo * NumPeriodos(!TipoTasa)
            lblMunicipio.Caption = !Municipio
            lblEstado.Caption = !Estado
            lblTelefono.Caption = !Tel
            lblIdentificacion.Caption = !Identificacion
            lblCP.Caption = !CP
            lblCP.Tag = !Folio
            lblFecha.Caption = Format(!Fecha, "DD/MMM/YYYY")
            lblVencimiento.Caption = Format(!Vencimiento, "DD/MMM/YYYY")
            lblPrestamo.Caption = Format(!PrestamoInicial, FMoneda)
            
            'Muestro los recibos
            MuestraPagos !ID
        Else
            
            MsgBox "No se encontró el contrato especificado !!", vbInformation, "Pagos Fijos"
            txtContrato.SetFocus
        End If
    
    End With
    
    rcConsulta.Close
    Set rcConsulta = Nothing
End Sub

Function MuestraPagos(IDContrato As Long)
Dim Interes As Double, Amortizacion As Double, Saldo As Double, crMoratorios As Double, i As Integer, Meses As Integer
Dim rcConsulta As New ADODB.Recordset

    rcConsulta.Open "SELECT pagosfijos.*,pagosfijos.Iva as IvaPF,empeno.VenPeriodo,empeno.TipoTasa,(empeno.IVA/100) AS IVA FROM pagosfijos LEFT JOIN empeno ON pagosfijos.IDEmpeno=empeno.ID WHERE pagosfijos.IDEmpeno=" & IDContrato & " AND pagosfijos.Cancelado=0 ORDER BY pagosfijos.NumPago,pagosfijos.ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        Select Case rcConsulta!TipoTasa
        Case "MENSUAL"
            
            Meses = 1
        Case "QUINCENAL"
            
            Meses = 1
        Case "SEMANAL"
            
            Meses = 1
        End Select
                
        With grdCalendario
            
            .Redraw = False
            .Clear
            txtTotalGeneral = Format(0, FMoneda)
            While Not rcConsulta.EOF
            
                DoEvents
            
                .AddRow
                .CellText(.Rows, 1) = rcConsulta!NumPago & " de " & (rcConsulta!VenPeriodo * Meses)
                .CellIcon(.Rows, 1) = IIf(rcConsulta!Pagado = 0, lstIcons.ItemIndex(1), -1)
                .CellItemData(.Rows, 1) = rcConsulta!ID
                .CellTextAlign(.Rows, 1) = DT_CENTER
                
                .CellText(.Rows, 2) = rcConsulta!Vencimiento
                .CellItemData(.Rows, 2) = rcConsulta!Pagado
                .CellTextAlign(.Rows, 2) = DT_CENTER
                
                .CellText(.Rows, 3) = (rcConsulta!Interes + rcConsulta!Almacenaje + rcConsulta!Seguro + rcConsulta!IvaPF)
                .CellItemData(.Rows, 3) = rcConsulta!FolioRecibo
                .CellTextAlign(.Rows, 3) = DT_RIGHT
                
                .CellText(.Rows, 4) = Redondeo(rcConsulta!Interes - (rcConsulta!Interes * (rcConsulta!Bonificacion / 100)))
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                
                .CellText(.Rows, 5) = Redondeo(rcConsulta!Almacenaje - (rcConsulta!Almacenaje * (rcConsulta!Bonificacion / 100)))
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                
                .CellText(.Rows, 6) = Redondeo(rcConsulta!Seguro - (rcConsulta!Seguro * (rcConsulta!Bonificacion / 100)))
                .CellTextAlign(.Rows, 6) = DT_RIGHT
                
                .CellText(.Rows, 7) = Redondeo(rcConsulta!IvaPF - (rcConsulta!IvaPF * (rcConsulta!Bonificacion / 100)))
                .CellTextAlign(.Rows, 7) = DT_RIGHT
                
                If rcConsulta!Pagado = 1 Then
                
                    crMoratorios = Redondeo(rcConsulta!Moratorios)
                Else
                
                    crMoratorios = Moratorios(rcConsulta!ID, rcConsulta!Pago)
                End If
                .CellText(.Rows, 8) = crMoratorios
                .CellTextAlign(.Rows, 8) = DT_RIGHT
                
                .CellText(.Rows, 9) = rcConsulta!Amortizacion
                .CellTextAlign(.Rows, 9) = DT_RIGHT
                
                .CellText(.Rows, 10) = Redondeo(rcConsulta!Pago + crMoratorios)
                .CellTextAlign(.Rows, 10) = DT_RIGHT
                
                .CellText(.Rows, 11) = rcConsulta!Saldo
                .CellTextAlign(.Rows, 11) = DT_RIGHT
                
                .CellText(.Rows, 12) = IIf(rcConsulta!cancelado = 1, "Cancelado", IIf(rcConsulta!Pagado = 0, "Pendiente", "Liquidado"))
                .CellTextAlign(.Rows, 12) = DT_LEFT
                
                .CellText(.Rows, 13) = IIf(IsNull(rcConsulta!FechaMovimiento), "", rcConsulta!FechaMovimiento)
                .CellTextAlign(.Rows, 13) = DT_CENTER
                
                .CellText(.Rows, 14) = 0
                .CellTextAlign(.Rows, 14) = DT_RIGHT
                
                .CellText(.Rows, 15) = Redondeo((Redondeo(rcConsulta!Interes / (1 + rcConsulta!Iva)) + Redondeo(rcConsulta!Almacenaje / (1 + rcConsulta!Iva)) + Redondeo(rcConsulta!Seguro / (1 + rcConsulta!Iva))) * (rcConsulta!Bonificacion / 100))
                .CellTextAlign(.Rows, 15) = DT_RIGHT
                
               
                
            Poner_Colores grdCalendario, .Rows
            rcConsulta.MoveNext
            Wend
            .Redraw = True
            
        End With

    End If
    
    rcConsulta.Close
    Set rcConsulta = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdCalendario_Click(ByVal lRow As Long, ByVal lCol As Long)
Dim Bonificacion As Double

    If lCol = 1 And lRow > 0 Then
        
        If grdCalendario.SelectedRow > 0 Then
            
            If grdCalendario.CellItemData(lRow, 2) = 0 Then
                
                If grdCalendario.CellIcon(lRow, lCol) = lstIcons.ItemIndex(1) Then
                    
                    Bonificacion = 0
                    If CDate(grdCalendario.CellText(grdCalendario.SelectedRow, 2)) > Date Then
                        
                        Bonificacion = Regresa_Valor_BD("DescuentoPagosFijos")
                    End If
                    
                    grdCalendario.CellText(grdCalendario.SelectedRow, 14) = Bonificacion
                    grdCalendario.CellIcon(lRow, lCol) = lstIcons.ItemIndex(2)
                    SacaTotales
                Else
                    
                    grdCalendario.CellText(grdCalendario.SelectedRow, 14) = 0
                    grdCalendario.CellIcon(lRow, lCol) = lstIcons.ItemIndex(1)
                    SacaTotales
                End If
            
            End If
            
        End If
    
    End If
    
End Sub

Function Moratorios(IDRecibo As Long, ImportePago As Double) As Double
Dim Dias As Integer, Tasa As Double, crIntereses As Double, i As Integer
Dim rcConsulta As New ADODB.Recordset
    
    Dias = 0
    crIntereses = 0
    Tasa = 0
    
    rcConsulta.Open "SELECT empeno.Prestamo,((empeno.Operacion/30)*(1 + (empeno.Iva/100)))/100 AS Moratorio,pagosfijos.Vencimiento FROM empeno INNER JOIN pagosfijos ON empeno.ID=pagosfijos.IDEmpeno WHERE pagosfijos.ID=" & IDRecibo, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        Tasa = rcConsulta!Moratorio
        Dias = DateDiff("D", CDate(rcConsulta!Vencimiento), Date)
        If Dias > 0 Then

            For i = 1 To Dias
                
                crIntereses = crIntereses + (ImportePago * Tasa)
            Next i
       
        End If
        
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing

    Moratorios = Redondeo(CCur(crIntereses))
End Function

Sub LimpiaPagos()
    lblNumPago.Caption = ""
    lblNumPago.Tag = ""
    lblAmortizacion.Caption = ""
    lblAmortizacion.Tag = ""
    lblIntereses.Caption = ""
    lblMoratorios.Caption = ""
    lblTotal.Caption = ""
End Sub

Private Sub txtContrato_GotFocus()
    Cambiar_Color True, txtContrato
End Sub

Private Sub txtContrato_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then cmdBuscar_Click
    Pasar_Foco KeyAscii
End Sub

Private Sub txtContrato_LostFocus()
    Cambiar_Color False, txtContrato
End Sub

Private Sub Limpiar(Contededor As String)
Dim ctrl As Control
  
    For Each ctrl In Controls
        
        On Error Resume Next

        If ctrl.Container.Caption = Contededor Then
            If TypeOf ctrl Is TextBox And ctrl.Name <> "NotaRef" And ctrl.Name <> "FechaCap" And ctrl.Name <> "VencimientoCap" Then ctrl.text = ""
            If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = "": ctrl.Tag = ""
            If TypeOf ctrl Is ComboBox And ctrl.Name <> "cmbTasas" And ctrl.Name <> "cmbPlazos" Then ctrl.ListIndex = -1
            On Error Resume Next
            ctrl.Tag = ""
        End If

    Next

End Sub

Sub Imprimir_Nota(Contrato As Long, FolioNota As Long, UltimoPago As Boolean)
Dim rcAux As New ADODB.Recordset
Dim Descripcion As String, NumPagos As Integer, NextVencimiento As String, ImprDefault As Boolean

On Error GoTo Error
    
    'Checo si hay impresora por default
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
    
    rcAux.Open "SELECT empeno.ID AS IDEmpeno FROM empeno LEFT JOIN pagosfijos ON empeno.ID=pagosfijos.IDEmpeno WHERE empeno.NumContrato=" & Contrato & " AND pagosfijos.FolioRecibo=" & FolioNota, dbDatos, adOpenForwardOnly, adLockReadOnly
        
        'Saco el Numero de Pago
        NumPagos = SacaValor("pagosfijos", "COUNT(IDEmpeno)", " WHERE IDEmpeno=" & rcAux!IDEmpeno)
        
        'Próximo Vencimiento
        NextVencimiento = SacaValor("pagosfijos", "Vencimiento", " WHERE ID=" & Val(SacaValor("pagosfijos", "MIN(ID)", " WHERE IDEmpeno=" & rcAux!IDEmpeno & " AND Pagado=0")))
        If Trim(NextVencimiento) <> "" Then NextVencimiento = Format(CDate(NextVencimiento), "DD-MMM-YYYY") Else NextVencimiento = "CONTRATO LIQUIDADO"
    
    rcAux.Close
    Set rcAux = Nothing
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\NotaPagoFijo.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{pagosfijos.FolioRecibo}=" & FolioNota & ""
        .Formulas(0) = "NumPagos=" & NumPagos & ""
        .Formulas(1) = "UltimoPago=" & IIf(UltimoPago, 1, 0) & ""
        .Formulas(2) = "ProximoVencimiento='" & NextVencimiento & "'"
        .Formulas(3) = "Notas='" & Trim(Regresa_Valor_BD("Notas")) & "'"
        .Formulas(4) = "Enajenacion=" & Trim(Regresa_Valor_BD("DiasEnajenacion")) & ""
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
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcAux = Nothing
End Sub

Sub SacaTotales()
Dim i As Integer, crIntereses As Double, crMoratorios As Double, crAmortizacion As Double, Bonificacion As Double
Dim crImporteInteres As Double, crBonificacion As Double

    LimpiaPagos
    For i = 1 To grdCalendario.Rows
            
        crBonificacion = IIf(grdCalendario.CellItemData(i, 2) = 1, grdCalendario.CellText(i, 15), 0)
        If grdCalendario.CellItemData(i, 2) = 0 And grdCalendario.CellIcon(i, 1) > 0 Then
                
            crBonificacion = 0
            Bonificacion = grdCalendario.CellText(i, 14) / 100
            If Bonificacion > 0 Then
                
                crImporteInteres = CDbl(grdCalendario.CellText(i, 3)) '''''/ (1 + (Val(lblDireccion.Tag) / 100)))
                crBonificacion = Redondeo(crImporteInteres * Bonificacion)
                crImporteInteres = Redondeo((crImporteInteres - crBonificacion)) '''''* (1 + (Val(lblDireccion.Tag) / 100)))
            Else
                
                crImporteInteres = CDbl(grdCalendario.CellText(i, 3))
            End If
            
            crIntereses = crIntereses + crImporteInteres
            crMoratorios = crMoratorios + CDbl(grdCalendario.CellText(i, 8))
            crAmortizacion = crAmortizacion + CDbl(grdCalendario.CellText(i, 9))
            grdCalendario.CellText(i, 15) = crBonificacion
        End If
            
        grdCalendario.CellText(i, 15) = crBonificacion
    Next i

    txtTotalGeneral.text = Format(crIntereses + crMoratorios + crAmortizacion, FMoneda)
End Sub

Function ValidaPagos() As Boolean
Dim i As Integer, x As Integer, PagoMarcado As Long, crEfectivo As Double, crTotal As Double
    
    ValidaPagos = True
    crEfectivo = 0
    crTotal = 0
    With grdCalendario
        
        For i = 1 To .Rows
            
            If .CellIcon(i, 1) > 0 Then
                
                PagoMarcado = i
                
            End If
            
        Next i
        
        For x = (PagoMarcado - 1) To 1 Step -1
            
            If .CellIcon(x, 1) = 0 Then
                
                MsgBox "Es necesario que liquide los pagos anteriores al seleccionado !!", vbInformation, "Pagos Fijos"
                ValidaPagos = False
                Exit Function
                
            End If
            
        Next x
        
    End With
    
    'Tomo el Efectivo
    If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
        
        crEfectivo = CDbl(txtEfectivo.text)
    End If
    
    'Tomo el Total
    If Val(txtTotalGeneral.text) > 0 Or Trim(txtTotalGeneral.text) <> "" Then
        
        crTotal = txtTotalGeneral.text
    End If
    
    'Valido si es correcto el efectivo
    If crTotal > crEfectivo Then
        MsgBox "El monto del efectivo es insuficiente para cubrir el importe a pagar !!", vbInformation, "Pagos Fijos"
        ValidaPagos = False
        txtEfectivo.SetFocus
        Exit Function
    End If
    
End Function

Private Sub txtEfectivo_Change()
    Saca_Cambio
End Sub

Private Sub txtEfectivo_GotFocus()
    Seleccionar_Texto txtEfectivo
    Cambiar_Color True, txtEfectivo
End Sub

Private Sub txtEfectivo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEfectivo_LostFocus()
    txtEfectivo.text = Format(txtEfectivo.text, FMoneda)
    Cambiar_Color False, txtEfectivo
End Sub

Private Sub txtCambio_GotFocus()
    Seleccionar_Texto txtCambio
    Cambiar_Color True, txtCambio
End Sub

Private Sub txtCambio_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCambio_LostFocus()
    Cambiar_Color False, txtCambio
End Sub

Private Sub txtTotalGeneral_Change()
    Saca_Cambio
End Sub

Private Sub txtTotalGeneral_GotFocus()
    Seleccionar_Texto txtTotalGeneral
    Cambiar_Color True, txtTotalGeneral
End Sub

Private Sub txtTotalGeneral_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTotalGeneral_LostFocus()
    Cambiar_Color False, txtTotalGeneral
End Sub

Sub Saca_Cambio()
Dim crTotal As Double, crEfectivo As Double
    
    If Val(txtTotalGeneral.text) > 0 Or Trim(txtTotalGeneral.text) <> "" Then
        
        crTotal = CDbl(txtTotalGeneral.text)
    Else
        
        crTotal = 0
    End If
    
    If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
        
        crEfectivo = CDbl(txtEfectivo.text)
    Else
        
        crEfectivo = 0
    End If
    
    txtCambio.text = Format(crEfectivo - crTotal, FMoneda)
End Sub

Function NumPeriodos(TipoTasa As String) As Integer
    
    NumPeriodos = 0
    Select Case Trim(TipoTasa)
    Case "MENSUAL"
        
        NumPeriodos = 1
    Case "QUINCENAL"
        
        NumPeriodos = 2
    Case "SEMANAL"
        
        NumPeriodos = 4
    End Select
    
End Function
