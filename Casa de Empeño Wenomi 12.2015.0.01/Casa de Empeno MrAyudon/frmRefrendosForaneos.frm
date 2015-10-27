VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{1781610F-46E8-4DD3-922D-8DEF1A9DA567}#28.0#0"; "Credencial.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.2#0"; "vbalSGrid6.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.Form frmRefrendosForaneos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Refrendos entre Sucursales"
   ClientHeight    =   10590
   ClientLeft      =   405
   ClientTop       =   870
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRefrendosForaneos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10590
   ScaleWidth      =   10500
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   18240
      ScaleHeight     =   1545
      ScaleWidth      =   2265
      TabIndex        =   91
      TabStop         =   0   'False
      Top             =   10200
      Width           =   2295
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBorrar 
      Height          =   375
      Left            =   20520
      TabIndex        =   92
      Top             =   10800
      Visible         =   0   'False
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
      Picture         =   "frmRefrendosForaneos.frx":000C
      PictureDisabled =   "frmRefrendosForaneos.frx":055E
   End
   Begin vbalIml6.vbalImageList lstIcons 
      Left            =   20520
      Top             =   10200
      _ExtentX        =   953
      _ExtentY        =   953
   End
   Begin vbalDTab6.vbalDTabControl tTab 
      Height          =   10335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   18230
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowCloseButton =   0   'False
      Begin VB.Frame FrameAutomatico 
         BorderStyle     =   0  'None
         Height          =   9735
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Visible         =   0   'False
         Width           =   9855
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   13
            Left            =   8280
            Top             =   8760
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin VB.Frame Frame1 
            Caption         =   "Parámetros de Búsqueda"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   7215
            Begin VB.CheckBox chkAutomovil 
               Appearance      =   0  'Flat
               Caption         =   "Automóvil"
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
               Height          =   285
               Left            =   3000
               TabIndex        =   10
               Top             =   1000
               Width           =   1575
            End
            Begin VB.TextBox txtFolioRefrendo 
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
               Left            =   1320
               TabIndex        =   6
               Top             =   1020
               Width           =   1500
            End
            Begin VB.ComboBox cmbSucursales 
               Appearance      =   0  'Flat
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
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Tag             =   "Sucursales"
               Top             =   465
               Width           =   4500
            End
            Begin VB.TextBox txtCedula 
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
               Left            =   1080
               TabIndex        =   8
               Top             =   1380
               Visible         =   0   'False
               Width           =   1860
            End
            Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
               Height          =   375
               Left            =   4680
               TabIndex        =   9
               Top             =   960
               Visible         =   0   'False
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
               Picture         =   "frmRefrendosForaneos.frx":1130
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contrato:"
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
               Left            =   240
               TabIndex        =   5
               Top             =   1005
               Width           =   945
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sucursal:"
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
               Index           =   0
               Left            =   240
               TabIndex        =   3
               Top             =   525
               Width           =   900
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cedula:"
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
               Index           =   1
               Left            =   240
               TabIndex        =   7
               Top             =   1365
               Visible         =   0   'False
               Width           =   720
            End
         End
         Begin VB.Frame Frame3 
            Height          =   1905
            Left            =   4740
            TabIndex        =   17
            Top             =   1920
            Width           =   4935
            Begin VB.TextBox txtSubTotal 
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
               Left            =   2895
               Locked          =   -1  'True
               TabIndex        =   26
               Top             =   1410
               Width           =   1815
            End
            Begin VB.TextBox txtDescuento 
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
               Left            =   2880
               TabIndex        =   28
               Top             =   2370
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox txtAbono 
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
               Left            =   2895
               TabIndex        =   24
               Top             =   1050
               Width           =   1815
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Intereses: "
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
               Index           =   0
               Left            =   615
               TabIndex        =   21
               Top             =   690
               Width           =   1080
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Abono a Capital:"
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
               Index           =   1
               Left            =   615
               TabIndex        =   23
               Top             =   1050
               Width           =   1605
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Prestamo: "
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
               Index           =   2
               Left            =   615
               TabIndex        =   19
               Top             =   330
               Width           =   1065
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Subtotal: "
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
               Index           =   3
               Left            =   615
               TabIndex        =   25
               Top             =   1410
               Width           =   960
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descuento: "
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
               Index           =   4
               Left            =   615
               TabIndex        =   27
               Top             =   2370
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label Label5 
               BackColor       =   &H00FFD4B3&
               Height          =   2235
               Left            =   0
               TabIndex        =   18
               Top             =   90
               Width           =   465
            End
            Begin VB.Label lblPrestamo 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
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
               Left            =   2895
               TabIndex        =   20
               Top             =   330
               Width           =   1815
            End
            Begin VB.Label lblIntereses 
               Alignment       =   1  'Right Justify
               BackColor       =   &H80000005&
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
               Left            =   2895
               TabIndex        =   22
               Top             =   690
               Width           =   1815
            End
         End
         Begin VB.Frame Frame2 
            Height          =   1725
            Left            =   120
            TabIndex        =   35
            Top             =   6960
            Width           =   9615
            Begin VB.TextBox txtEfectivo 
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
               Height          =   360
               Left            =   6240
               TabIndex        =   41
               Text            =   "0.00"
               Top             =   720
               Width           =   3255
            End
            Begin VB.TextBox txtCambio 
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
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   6240
               Locked          =   -1  'True
               TabIndex        =   43
               Text            =   "0.00"
               Top             =   1200
               Width           =   3255
            End
            Begin VB.TextBox txtTotalGeneral 
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
               Height          =   360
               Left            =   6240
               Locked          =   -1  'True
               TabIndex        =   39
               Text            =   "0.00"
               Top             =   240
               Width           =   3255
            End
            Begin VB.ComboBox cmbPeriodosPagos 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   435
               Left            =   120
               Style           =   2  'Dropdown List
               TabIndex        =   37
               Top             =   720
               Visible         =   0   'False
               Width           =   3615
            End
            Begin VB.Label Label4 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "CAMBIO:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   315
               Left            =   4935
               TabIndex        =   42
               Top             =   1200
               Width           =   1155
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL PAGADO:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   3990
               TabIndex        =   40
               Top             =   720
               Width           =   2100
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "TOTAL A PAGAR:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Left            =   3945
               TabIndex        =   38
               Top             =   240
               Width           =   2145
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "PERIODOS A CANCELAR"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   120
               TabIndex        =   36
               Top             =   360
               Visible         =   0   'False
               Width           =   3060
            End
         End
         Begin VB.Frame Frame4 
            Height          =   1695
            Left            =   7380
            TabIndex        =   11
            Top             =   135
            Width           =   2295
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
               Height          =   345
               Left            =   120
               TabIndex        =   15
               Tag             =   "1"
               Top             =   1200
               Width           =   2070
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Vencimiento"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   12
               Left            =   315
               TabIndex        =   14
               Top             =   840
               Width           =   1590
            End
            Begin VB.Label lblFecha 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Left            =   120
               TabIndex        =   13
               Tag             =   "1"
               Top             =   480
               Width           =   2070
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "Fecha"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   10
               Left            =   690
               TabIndex        =   12
               Top             =   120
               Width           =   750
            End
         End
         Begin Credencial.usCredencial CCliente 
            Height          =   2295
            Left            =   120
            TabIndex        =   16
            Top             =   1920
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   4048
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   6
            AlingHeader     =   262144
            AlingBody       =   0
            BodyIndent      =   10
            HeaderIndent    =   5
            HeaderText      =   "Datos del Cliente"
            HeaderBackColor =   16766131
            HeightHeader    =   20
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   39
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   13603685
            BackColor       =   16777215
         End
         Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
            Height          =   375
            Left            =   8700
            TabIndex        =   30
            Top             =   4080
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
            Picture         =   "frmRefrendosForaneos.frx":14B5
            PictureDisabled =   "frmRefrendosForaneos.frx":181F
         End
         Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
            Height          =   375
            Left            =   7620
            TabIndex        =   29
            Top             =   4080
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
            Picture         =   "frmRefrendosForaneos.frx":1979
         End
         Begin TabDlg.SSTab SSTab1 
            Height          =   2580
            Left            =   180
            TabIndex        =   31
            Top             =   4305
            Width           =   9525
            _ExtentX        =   16801
            _ExtentY        =   4551
            _Version        =   393216
            Style           =   1
            Tabs            =   2
            TabsPerRow      =   2
            TabHeight       =   520
            TabCaption(0)   =   "Contratos"
            TabPicture(0)   =   "frmRefrendosForaneos.frx":1A7D
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "grdRefrendos"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Prendas"
            TabPicture(1)   =   "frmRefrendosForaneos.frx":1A99
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "grdPrendas"
            Tab(1).ControlCount=   1
            Begin vbAcceleratorGrid6.vbalGrid grdRefrendos 
               Height          =   2160
               Left            =   0
               TabIndex        =   32
               Top             =   360
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   3810
               RowMode         =   -1  'True
               GridLines       =   -1  'True
               BackgroundPictureHeight=   0
               BackgroundPictureWidth=   0
               BackColor       =   16777215
               GridLineColor   =   12632256
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
               DefaultRowHeight=   25
               Begin VB.TextBox txtEdit 
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
                  Left            =   840
                  TabIndex        =   33
                  Top             =   0
                  Visible         =   0   'False
                  Width           =   1455
               End
            End
            Begin vbAcceleratorGrid6.vbalGrid grdPrendas 
               Height          =   2160
               Left            =   -75000
               TabIndex        =   34
               Top             =   375
               Width           =   9495
               _ExtentX        =   16748
               _ExtentY        =   3810
               RowMode         =   -1  'True
               GridLines       =   -1  'True
               BackgroundPictureHeight=   0
               BackgroundPictureWidth=   0
               BackColor       =   16777215
               GridLineColor   =   12632256
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
               DefaultRowHeight=   25
            End
         End
         Begin DevPowerFlatBttn.FlatBttn cmdReImprimir 
            Height          =   375
            Left            =   6120
            TabIndex        =   95
            Top             =   9240
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
            Picture         =   "frmRefrendosForaneos.frx":1AB5
         End
         Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
            Height          =   375
            Left            =   7440
            TabIndex        =   96
            Top             =   9240
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
            Picture         =   "frmRefrendosForaneos.frx":2007
         End
         Begin DevPowerFlatBttn.FlatBttn cmdSalir 
            Height          =   375
            Left            =   8640
            TabIndex        =   97
            Top             =   9240
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
            Picture         =   "frmRefrendosForaneos.frx":2559
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   16
            Left            =   2520
            Top             =   8760
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   8
            Left            =   120
            Top             =   8760
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   20
            Left            =   120
            Top             =   8760
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D10 
            Height          =   30
            Index           =   23
            Left            =   120
            Top             =   9120
            Width           =   9645
            _ExtentX        =   17013
            _ExtentY        =   53
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   14
            Left            =   3720
            Top             =   8760
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin VB.TextBox txtFechaTentativa 
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
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   98
            Top             =   8805
            Width           =   1215
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   10
            Left            =   4200
            Top             =   8760
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin DevPowerFlatBttn.FlatBttn cmdFechaTentativa 
            Height          =   420
            Left            =   3840
            TabIndex        =   101
            Top             =   8760
            Width           =   315
            _ExtentX        =   556
            _ExtentY        =   741
            AutoSize        =   0   'False
            Caption         =   ""
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
            Picture         =   "frmRefrendosForaneos.frx":2AAB
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   15
            Left            =   6000
            Top             =   8760
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   9
            Left            =   6480
            Top             =   8760
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin DevPowerFlatBttn.FlatBttn cmdTicketFechaTentativa 
            Height          =   375
            Left            =   6120
            TabIndex        =   104
            Top             =   8760
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   661
            AlignCaption    =   3
            AlignPicture    =   2
            AutoSize        =   0   'False
            Caption         =   ""
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
            Picture         =   "frmRefrendosForaneos.frx":328D
         End
         Begin Line3D.ucLine3D ucLine3D30 
            Height          =   360
            Index           =   0
            Left            =   9730
            Top             =   8760
            Width           =   30
            _ExtentX        =   53
            _ExtentY        =   635
            Orientation     =   0
            LineWidth       =   2
         End
         Begin VB.TextBox txtCambiarPlan 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8280
            Locked          =   -1  'True
            TabIndex        =   106
            Top             =   8760
            Width           =   1455
         End
         Begin VB.Label lblDescuento 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "<Descuento>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   210
            Left            =   4800
            TabIndex        =   108
            Top             =   3840
            Width           =   1155
         End
         Begin VB.Label lblPagoTentativo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
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
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   4920
            TabIndex        =   107
            Top             =   8820
            Width           =   900
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
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
            Height          =   360
            Index           =   7
            Left            =   6480
            TabIndex        =   105
            Top             =   8760
            Width           =   1800
         End
         Begin VB.Label lblPago 
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
            Index           =   42
            Left            =   4320
            TabIndex        =   102
            Top             =   8820
            Width           =   645
         End
         Begin VB.Label lblFechaTentativa 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
            BackStyle       =   0  'Transparent
            Caption         =   " Fecha tentativa de pago: "
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
            Height          =   240
            Index           =   9
            Left            =   120
            TabIndex        =   99
            Top             =   8820
            Width           =   2355
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
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
            Height          =   360
            Index           =   5
            Left            =   120
            TabIndex        =   100
            Top             =   8760
            Width           =   2400
         End
         Begin VB.Label Label55 
            AutoSize        =   -1  'True
            BackColor       =   &H00404040&
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
            Height          =   360
            Index           =   6
            Left            =   4200
            TabIndex        =   103
            Top             =   8760
            Width           =   1785
         End
      End
   End
   Begin VB.Frame FrameManual 
      BorderStyle     =   0  'None
      Height          =   8535
      Left            =   120
      TabIndex        =   44
      Top             =   720
      Width           =   9855
      Begin VB.Frame Frame8 
         Caption         =   "Datos de las Prendas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3495
         Left            =   120
         TabIndex        =   77
         Top             =   5040
         Width           =   9600
         Begin VB.TextBox txtValor 
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
            Left            =   7560
            TabIndex        =   93
            Top             =   1440
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtKT 
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
            Left            =   6360
            TabIndex        =   83
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtPesoPiedra 
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
            Left            =   6960
            TabIndex        =   85
            Top             =   600
            Width           =   495
         End
         Begin VB.TextBox txtPeso 
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
            Left            =   7560
            TabIndex        =   87
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox txtDescripcion 
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
            Height          =   525
            Left            =   840
            MultiLine       =   -1  'True
            TabIndex        =   81
            Top             =   600
            Width           =   5415
         End
         Begin VB.TextBox txtCantidad 
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
            Left            =   120
            TabIndex        =   79
            Top             =   600
            Width           =   615
         End
         Begin vbAcceleratorSGrid6.vbalGrid grdPrendasManual 
            Height          =   2175
            Left            =   120
            TabIndex        =   89
            Top             =   1200
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   3836
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
            BorderStyle     =   2
            DisableIcons    =   -1  'True
         End
         Begin DevPowerFlatBttn.FlatBttn cmdAgregarManual 
            Height          =   855
            Left            =   8640
            TabIndex        =   88
            Top             =   240
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   1508
            AlignCaption    =   4
            AlignPicture    =   2
            AutoSize        =   0   'False
            Caption         =   "&Agregar"
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
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Valor:"
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
            Index           =   8
            Left            =   7920
            TabIndex        =   94
            Top             =   1200
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Peso:"
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
            Index           =   13
            Left            =   7560
            TabIndex        =   86
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "P. P.:"
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
            Index           =   11
            Left            =   6960
            TabIndex        =   84
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "KT:"
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
            Index           =   9
            Left            =   6480
            TabIndex        =   82
            Top             =   360
            Width           =   300
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Descripcion:"
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
            Index           =   7
            Left            =   840
            TabIndex        =   80
            Top             =   360
            Width           =   1185
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cant.:"
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
            Index           =   6
            Left            =   120
            TabIndex        =   78
            Top             =   360
            Width           =   585
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Datos del Contrato"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4740
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   9615
         Begin VB.TextBox txtCodigoSucursal 
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
            Left            =   1800
            TabIndex        =   63
            Top             =   2760
            Width           =   1500
         End
         Begin VB.TextBox txtTasa 
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
            Left            =   3360
            TabIndex        =   51
            Top             =   660
            Width           =   1260
         End
         Begin VB.ComboBox cmbPeriodosManual 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   4920
            Style           =   2  'Dropdown List
            TabIndex        =   61
            Top             =   2040
            Width           =   3615
         End
         Begin VB.Frame Frame7 
            Height          =   1875
            Left            =   4920
            TabIndex        =   64
            Top             =   2760
            Width           =   4575
            Begin VB.TextBox txtPrestamoManual 
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
               Left            =   2760
               TabIndex        =   67
               Top             =   165
               Width           =   1575
            End
            Begin VB.TextBox txtInteresesManual 
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
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   69
               Top             =   525
               Width           =   1575
            End
            Begin VB.TextBox txtAbonoManual 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               Enabled         =   0   'False
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
               Left            =   2760
               TabIndex        =   71
               Top             =   930
               Width           =   1560
            End
            Begin VB.TextBox Text4 
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
               Left            =   2880
               TabIndex        =   75
               Top             =   2370
               Visible         =   0   'False
               Width           =   1815
            End
            Begin VB.TextBox txtSubTotalManual 
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
               Left            =   2760
               Locked          =   -1  'True
               TabIndex        =   73
               Top             =   1290
               Width           =   1560
            End
            Begin VB.Label Label7 
               BackColor       =   &H00FFD4B3&
               Height          =   2235
               Left            =   0
               TabIndex        =   65
               Top             =   90
               Width           =   465
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Descuento: "
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
               Index           =   9
               Left            =   615
               TabIndex        =   74
               Top             =   2370
               Visible         =   0   'False
               Width           =   1170
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Total Pagar: "
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
               Index           =   8
               Left            =   615
               TabIndex        =   72
               Top             =   1290
               Width           =   1245
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Préstamo: "
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
               Index           =   7
               Left            =   615
               TabIndex        =   66
               Top             =   210
               Width           =   1065
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Abono a Capital: "
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
               Index           =   6
               Left            =   615
               TabIndex        =   70
               Top             =   930
               Width           =   1665
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Intereses Generados: "
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
               Index           =   5
               Left            =   615
               TabIndex        =   68
               Top             =   570
               Width           =   2175
            End
         End
         Begin VB.Frame Frame6 
            Height          =   1335
            Left            =   4935
            TabIndex        =   52
            Top             =   240
            Width           =   4560
            Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
               Height          =   360
               Index           =   0
               Left            =   3630
               TabIndex        =   55
               Top             =   360
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   635
               AlignCaption    =   4
               AlignPicture    =   4
               AutoSize        =   0   'False
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
               Picture         =   "frmRefrendosForaneos.frx":37DF
            End
            Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
               Height          =   360
               Index           =   1
               Left            =   3630
               TabIndex        =   58
               Top             =   825
               Width           =   360
               _ExtentX        =   635
               _ExtentY        =   635
               AlignCaption    =   4
               AlignPicture    =   4
               AutoSize        =   0   'False
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
               Picture         =   "frmRefrendosForaneos.frx":38F4
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Fecha:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   5
               Left            =   960
               TabIndex        =   53
               Top             =   360
               Width           =   840
            End
            Begin VB.Label lblFechaManual 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Left            =   1920
               TabIndex        =   54
               Tag             =   "1"
               Top             =   360
               Width           =   1695
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vencimiento:"
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   12.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000000&
               Height          =   315
               Index           =   4
               Left            =   120
               TabIndex        =   56
               Top             =   840
               Width           =   1680
            End
            Begin VB.Label lblVencimientoManual 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
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
               Left            =   1920
               TabIndex        =   57
               Tag             =   "1"
               Top             =   840
               Width           =   1695
            End
         End
         Begin VB.TextBox txtNumContratoManual 
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
            Left            =   1800
            TabIndex        =   47
            Top             =   240
            Width           =   1620
         End
         Begin VB.TextBox txtCedulaCliente 
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
            Left            =   945
            TabIndex        =   49
            Top             =   660
            Width           =   2340
         End
         Begin Credencial.usCredencial cClienteManual 
            Height          =   1635
            Left            =   105
            TabIndex        =   59
            Top             =   1020
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2884
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   6
            AlingHeader     =   262144
            AlingBody       =   0
            BodyIndent      =   10
            HeaderIndent    =   5
            HeaderText      =   "Datos del Cliente"
            HeaderBackColor =   16766131
            HeightHeader    =   20
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   39
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   13603685
            BackColor       =   16777215
         End
         Begin Credencial.usCredencial cSucursal 
            Height          =   1515
            Left            =   120
            TabIndex        =   76
            Top             =   3120
            Width           =   4575
            _ExtentX        =   8070
            _ExtentY        =   2672
            BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BorderStyle     =   6
            AlingHeader     =   262144
            AlingBody       =   0
            BodyIndent      =   10
            HeaderIndent    =   5
            HeaderText      =   "Datos de la Sucursal"
            HeaderBackColor =   16766131
            HeightHeader    =   20
            SidePicture     =   -1  'True
            SideBackColor   =   15000804
            WidthSide       =   39
            SidePicture     =   -1  'True
            HeaderBorderBackColor=   13603685
            BackColor       =   16777215
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Codigo Sucursal:"
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
            Index           =   15
            Left            =   120
            TabIndex        =   62
            Top             =   2760
            Width           =   1605
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Tasa Ret:"
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
            Index           =   14
            Left            =   3720
            TabIndex        =   50
            Top             =   375
            Width           =   930
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PERIODOS A CANCELAR"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   4920
            TabIndex        =   60
            Top             =   1680
            Width           =   3060
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Num. Contrato:"
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
            Index           =   3
            Left            =   105
            TabIndex        =   46
            Top             =   240
            Width           =   1470
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Cedula:"
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
            Index           =   2
            Left            =   105
            TabIndex        =   48
            Top             =   705
            Width           =   720
         End
      End
   End
   Begin VB.Label lblCliente 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Armando Garzón Serrano"
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
      Left            =   18240
      TabIndex        =   90
      Top             =   9840
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmRefrendosForaneos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Dim crIntereses As Double, _
    crAlmacenaje As Double, _
    crSeguro As Double, _
    crGastosAdmon As Double, _
    crImporteDescuento As Double, _
    crMoratorios As Double, _
    crRedondeo As Double, _
    crIva As Double, _
    crPerdida As Double, _
    crMinimo As Double

Private m_NumContrato As Long
Private m_Sucursal As String
Private m_Clave As Long
Private m_TipoTasa As String
Private m_Periodo As Integer
Private m_Tasa As Double
Private m_GastosAdmon As Double
Private m_Almacenaje As Double
Private m_Seguro As Double
Private m_Sucursales() As DatosSucursal

Public Property Let NumeroContrato(Valor As Long)
    m_NumContrato = Valor
End Property

Public Property Let sSucursal(Valor As String)
    m_Sucursal = Valor
End Property

Public Property Let ClaveSucursal(Valor As Long)
    m_Clave = Valor
End Property

Public Property Get TipoTasa() As String
    TipoTasa = m_TipoTasa
End Property

Public Property Let TipoTasa(Valor As String)
    m_TipoTasa = Valor
End Property

Public Property Get Periodo() As Integer
    Periodo = m_Periodo
End Property

Public Property Let Periodo(Valor As Integer)
    m_Periodo = Valor
End Property

Public Property Get Tasa() As Double
    Tasa = m_Tasa
End Property

Public Property Let Tasa(Valor As Double)
    m_Tasa = Valor
End Property

Public Property Get GastosAdmon() As Double
    GastosAdmon = m_GastosAdmon
End Property

Public Property Let GastosAdmon(Valor As Double)
    m_GastosAdmon = Valor
End Property

Public Property Get Almacenaje() As Double
    Almacenaje = m_Almacenaje
End Property

Public Property Let Almacenaje(Valor As Double)
    m_Almacenaje = Valor
End Property

Public Property Get Seguro() As Double
    Seguro = m_Seguro
End Property

Public Property Let Seguro(Valor As Double)
    m_Seguro = Valor
End Property

Private Sub cmbSucursales_DropDown()
    Cambiar_Color True, cmbSucursales
End Sub

Private Sub cmbSucursales_GotFocus()
    Cambiar_Color True, cmbSucursales
End Sub

Private Sub cmbSucursales_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbSucursales_LostFocus()
    Cambiar_Color False, cmbSucursales
End Sub

Private Sub cmdAceptar_Click()
    cmdAceptar.Enabled = False
    
    Procesar_Renovacion_EnLinea
    
    cmdAceptar.Enabled = True
End Sub

Private Sub Procesar_Renovacion_EnLinea()
    Dim crEfectivo As Double
    
    If Len(Trim(txtFolioRefrendo.text)) = 0 Or grdRefrendos.Rows = 0 Then
        MsgBox "Indique el Número de Contrato !!!", vbInformation, Me.Caption
        txtFolioRefrendo.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Estan correctos los datos ?? Seguro que desea realizar el REFRENDO !!!", vbQuestion + vbYesNo + vbDefaultButton1, "Refrendo Foraneo") = vbYes Then
    
        crEfectivo = IIf(Len(Trim(txtEfectivo.text)) = 0, 0, txtEfectivo.text)
    
        If crEfectivo < CDbl(txtTotalGeneral.text) Then
            MsgBox "El Total Pagado debe ser mayor que el Total a Pagar !!", vbInformation, Me.Caption
            txtEfectivo.SetFocus
            Exit Sub
        End If
    
        GrabarRefrendo
    End If

End Sub

Private Sub cmdAgregar_Click()
Dim crAbono As Double, crDescuento As Double
On Error GoTo Error

    If Len(Trim(lblPrestamo.Caption)) = 0 Then
        MsgBox "Seleccione primero la Boleta que desea refrendar !!", vbInformation, "Pago Interés Foráneos"
        txtFolioRefrendo.SetFocus
        Exit Sub
    End If
    
    If grdRefrendos.Rows = 1 Then
        MsgBox "Sólo se puede Refrendar una boleta !!!", vbInformation, "Refrendo Foráneo"
        txtFolioRefrendo.SetFocus
        Exit Sub
    End If
    
    crAbono = IIf(txtAbono.text = "", 0, txtAbono.text)
    crDescuento = IIf(txtDescuento.text = "", 0, txtDescuento.text)
    
    CargarDatosGrid Foraneo.Folio, crIntereses, crAlmacenaje, crSeguro, crMoratorios, crIva, crMinimo, crPerdida, crAbono, crDescuento, crGastosAdmon, crRedondeo, crImporteDescuento
    
    Me.txtTotalGeneral.text = txtSubTotal.text
    
    lblPrestamo.Caption = ""
    lblIntereses.Caption = ""
    txtAbono.text = ""
    txtDescuento.text = ""
    txtSubTotal.text = ""
    
    grdRefrendos_Click 1, 1
    txtEfectivo.SetFocus
    
Error:
    Maneja_Error Err
End Sub

Private Sub cmdAgregarManual_Click()
    
    If Validar_Agregar_Prenda Then
    
        With grdPrendasManual
            .AddRow
            .CellDetails .Rows, 1, txtCantidad.text, DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellDetails .Rows, 2, txtDescripcion.text, DT_LEFT Or DT_WORD_ELLIPSIS
            .CellDetails .Rows, 3, txtPeso.text, DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellDetails .Rows, 4, txtPesoPiedra.text, DT_RIGHT Or DT_WORD_ELLIPSIS
            .CellDetails .Rows, 5, txtKT.text, DT_RIGHT Or DT_WORD_ELLIPSIS
            '.CellDetails .Rows, 6, txtValor.text, DT_RIGHT Or DT_WORD_ELLIPSIS
            
            ReDim Preserve DetallesEmpenoForaneo(.Rows)
            With DetallesEmpenoForaneo(.Rows)
             .Articulo = txtDescripcion.text
             .Cantidad = Val(txtCantidad.text)
             .Peso = Val(ConvMoneda(txtPeso.text))
             .PesoPiedras = Val(ConvMoneda(txtPesoPiedra.text))
             '.Kilates = txtKT.text
             .Estado = txtKT.text
             '.Prestamo = txtValor.text
            End With
            
            .AutoHeightRow .Rows
        
        End With
        
        txtCantidad.text = ""
        txtDescripcion.text = ""
        txtKT.text = ""
        txtPesoPiedra.text = ""
        txtPeso.text = ""
            
        txtCantidad.SetFocus
    
    End If
    
End Sub

Private Function Validar_Agregar_Prenda() As Boolean
    On Error GoTo Error
    Validar_Agregar_Prenda = False
    
    If Trim(txtCantidad.text) = "" Then
        MsgBox "Favor de digitar la cantidad", vbCritical Or vbOKOnly
        Validar_Agregar_Prenda = False
    ElseIf Trim(txtDescripcion.text) = "" Then
        MsgBox "Favor de digitar la descripcion", vbCritical Or vbOKOnly
        Validar_Agregar_Prenda = False
    End If
        
    Validar_Agregar_Prenda = True
Error:
    Maneja_Error Err
End Function


Private Sub cmdBuscar_Click()
    Buscar_Contrato
End Sub

Private Sub Buscar_Contrato()
    Limpiar
    If Trim(txtFolioRefrendo.text) <> "" Then BuscarEmpeno txtFolioRefrendo.text, Date
End Sub

Private Sub cmdFechaTentativa_Click()
    PagoFechaTentativa False
End Sub

Private Sub PagoFechaTentativa(Optional ByVal Imprimir As Boolean = False)
    
    Dim FechaComercializacion As Date, FechaTentativa As Date
    Dim Moratorios As Double, Refrendo As Double
    Dim ImprDefault As Boolean, i As Integer

On Error GoTo Error
    
    If Foraneo.NumContrato = 0 Or IsNull(Foraneo.NumContrato) = True Then Exit Sub
        
    If Imprimir = True And txtFechaTentativa.text <> "" Then
        FechaTentativa = txtFechaTentativa.text
    Else
        FechaTentativa = Format(frmCalendario.Fecha(Date), "YYYY/MM/DD")
    End If
    
    Moratorios = 0: Refrendo = 0

    If FechaTentativa < Foraneo.Vencimiento Then
        MsgBox "La fecha propuesta es menor a la fecha de vencimiento ", vbCritical, "Empeño"
        Exit Sub
    End If

    'Moratorios hasta la Fecha Tentativa
    Moratorios = Redondeo(GeneraMoratoriosFechaTentativa(Foraneo.Prestamo, Foraneo.Vencimiento, FechaTentativa))
    
    'Fecha Comercializacion
    FechaComercializacion = DateAdd("d", IIf(Foraneo.TipoTasa = "MENSUAL", Regresa_Valor_BD("DiasComercializacion"), Regresa_Valor_BD("DiasComercializacion15")), Foraneo.Vencimiento)
    
    If FechaTentativa <= FechaComercializacion Then
        If Weekday(FechaTentativa) = 5 Then Moratorios = 0
    End If

    'Redondeo
    crRedondeo = Redondeo(CCur(crIntereses + crGastosAdmon + crAlmacenaje + crSeguro + crIva)) + Redondeo(CCur(Moratorios))
            
    Refrendo = crIntereses + crGastosAdmon + crAlmacenaje + crSeguro + crIva + crRedondeo + Moratorios
    
    If Imprimir Then
    
        For i = 1 To 2
            'Imprimir Nota
            ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
        
            With frmMDI.Cr
                .Reset
                .DiscardSavedData = True
                .WindowShowPrintSetupBtn = True
                .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
                .ReportFileName = Path & "\Reportes\NotaFechaTentativa.rpt"
                .SelectionFormula = "{sucursales.Activa} = 1"
                .Formulas(0) = "NumContrato='" & CStr(Foraneo.NumContrato) & "'"
                .Formulas(1) = "Cliente='" & Trim(UCase(ClienteForaneo.Nombre)) & " " & Trim(UCase(ClienteForaneo.Apellido)) & "'"
                .Formulas(2) = "Fecha='" & FechaTentativa & "'"
                .Formulas(3) = "Prestamo='" & Format(Foraneo.Prestamo, FMonedaSigno) & "'"
                .Formulas(4) = "Refrendo='" & Format(Refrendo, FMonedaSigno) & "'"
                .Formulas(5) = "Pago='" & Format(Foraneo.Prestamo + Refrendo, FMonedaSigno) & "'"
                .Formulas(6) = "Caja='" & Trim(UCase(NombrePc)) & "'"
                .Formulas(7) = "Usuario='" & SacaValor("usuarios", "Nombre", " WHERE ID=" & frmMDI.IDUsuario) & "'"
                .Formulas(8) = "Gerente='" & Regresa_Valor_BD("Gerente") & "'"
                .WindowState = crptMaximized
                .Destination = crptToWindow
            
                If ImprDefault Then
                    .PrinterName = strNombreImp
                    .PrinterDriver = strDriverImp
                    .PrinterPort = strPuertoImp
                    .Destination = crptToPrinter
                End If
            
                .WindowTitle = "Fecha Tentativa"
                .Action = 1
            End With
        Next
    Else
        lblPagoTentativo = Format(CStr(Refrendo), FMonedaSigno)
        txtFechaTentativa.text = Format(FechaTentativa, "YYYY/MM/DD")
    End If

Error:
    Maneja_Error Err
End Sub

Private Sub cmdLimpiar_Click()
    Dim ctrl As Object

    For Each ctrl In Me.Controls
        If TypeOf ctrl Is TextBox Then
            ctrl.text = ""
        End If
        If TypeOf ctrl Is ComboBox Then
            If ctrl.Tag = "Sucursales" Then
                ctrl.ListIndex = -1
            Else
                ctrl.Clear
            End If
        End If
    Next ctrl
    
    lblIntereses.Caption = ""
    lblPrestamo.Caption = ""
    lblFecha.Caption = ""
    lblVencimiento.Caption = ""
    txtSubTotal.text = "0.00"
    grdRefrendos.Clear
    grdPrendas.Clear
    Me.CCliente.Clear
    cmbPeriodosPagos.Enabled = True
    chkAutomovil.Value = 0
    lblDescuento.Caption = "<DESCUENTO>"
    
End Sub

Private Sub Cargar_Periodos_Manuales(Meses As Integer, Fecha As Date)
    Dim Indice As Integer
    
    cmbPeriodosManual.Clear
    For Indice = 1 To Meses
        cmbPeriodosManual.AddItem "PERIODO " & CStr(Indice) & " - " & Format(DateAdd("M", Indice, Fecha), "DD/MM/YYYY")
    Next Indice
    
End Sub

Private Function Calcular_Intereses_Manual(Periodo As Integer, Prestamo As Currency, Tasa As Double) As Currency
    On Error GoTo Error
    Dim Intereses As Currency
    
    
    Intereses = (Periodo * Prestamo) * (Tasa / 100)
    Calcular_Intereses_Manual = Intereses
    
    txtSubTotalManual.text = Calcular_Subtotal_Manual(Intereses, CCur(CDbl(Val(txtAbonoManual.text))))
    
Error:
    Maneja_Error Err
End Function

Private Function Calcular_Subtotal_Manual(Intereses As Currency, Abono As Currency) As Currency
    On Error GoTo Error
    Dim Subtotal As Currency
    
    Subtotal = Intereses + Abono
    
    Calcular_Subtotal_Manual = Subtotal
Error:
    Maneja_Error Err
End Function

Private Sub cmdReImprimir_Click()
    Dim Rs As New ADODB.Recordset
    Dim Pago As Currency, crImporteTotal As Double
    Dim Serie As Integer, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double, crPrestamo As Double, crGastosAdmon As Double, crRedondeo As Double
    Dim FolioNota As Long

On Error GoTo Error

    'Verificar la existencia de contrato
    If Foraneo.NumContrato = 0 Or IsNull(Foraneo.NumContrato) = True Or cmbSucursales.ListIndex = -1 Then
        Exit Sub
    Else

        Rs.Open "SELECT * FROM remesas WHERE numcontrato = " & Foraneo.NumContrato & " AND sucursalorigen = " & cmbSucursales.ItemData(cmbSucursales.ListIndex) & " AND Date(Fecha) = '" & Format(Now, "YYYY/MM/DD") & "'", dbDatos, adOpenStatic, adLockOptimistic
        With Rs

            If Rs.EOF Then
                MsgBox "No se encontró el contrato especificado !!", vbInformation, "Refrendo Foraneo"
            Else
                Imprimir_Nota !FolioNota, CDbl(!Abono + !Importe + !importeAlmacenaje + !importeSeguro + !ImporteMoratorios + !ImporteIva + !ImporteGastosAdmon + !ImporteRedondeo), CDbl(!Abono), CCur(!Importe + !importeAlmacenaje + !importeSeguro + !ImporteMoratorios + !ImporteIva + !ImporteGastosAdmon + !ImporteRedondeo), CCur(!Efectivo), !IDUsuario, True
            End If

            .Close
        End With
    End If

    Set Rs = Nothing
    Exit Sub
Error:
    Set Rs = Nothing
    Maneja_Error Err
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdTicketFechaTentativa_Click()
    PagoFechaTentativa True
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Crear_Tabs()
    Dim c As cTab
    With tTab
        Set c = .Tabs.Add("K1", , "Refrendo")
        c.Panel = FrameAutomatico
    End With
End Sub

Private Sub Inicializar()
    On Error GoTo Error
    
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    
    Crear_Tabs
    Cargar_Sucursales
    CrearEncabezados
    
    Set rcConsulta = New ADODB.Recordset
    Me.CCliente.AlingHeader = DT_WORD_ELLIPSIS Or DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    Me.CCliente.AlingBody = DT_WORD_ELLIPSIS Or DT_LEFT Or DT_WORDBREAK
    
    
    If cmbSucursales.ListCount > 0 Then
        If m_NumContrato > 0 Then
            cmbSucursales.text = m_Sucursal
            txtFolioRefrendo.text = m_NumContrato
            BuscarEmpeno m_NumContrato
        End If
    End If
        
Error:
    Maneja_Error Err
    
End Sub

Private Sub CrearEncabezados()
    With grdRefrendos
        .AddColumn "K1", "Contrato", ecgHdrTextALignLeft, , 85, , , , , , , CCLSortString
        .AddColumn "K2", "Valor", ecgHdrTextALignRight, , 107, , , , , FMoneda, , CCLSortString
        .AddColumn "K3", "Abono", ecgHdrTextALignRight, , 103, , , , , FMoneda, , CCLSortString
        .AddColumn "K4", "Interés", ecgHdrTextALignRight, , 107, , , , , FMoneda, , CCLSortString
      
        .AddColumn "K5", "Interés", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K6", "Almacenaje", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K7", "Seguro", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K8", "Iva", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K9", "Almoneda", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K10", "Importe Perdida", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K11", "Moratorios", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        
        
        .AddColumn "K12", "Otros Cobros", ecgHdrTextALignRight, , 107, False, , , , FMoneda, , CCLSortString
        .AddColumn "K13", "Iva", ecgHdrTextALignRight, , 107, False, , , , FMoneda, , CCLSortString
        .AddColumn "K14", "Descuento", ecgHdrTextALignRight, , 107, False, , , , FMoneda, , CCLSortString
        
        .AddColumn "K15", "GastosAdmon", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        .AddColumn "K16", "Redondeo", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
        
        .AddColumn "K17", "ImporteDescuento", ecgHdrTextALignRight, , 0, False, , , , , , CCLSortString
    End With
    
    With grdPrendas
        .AddColumn "K1", "Contrato", ecgHdrTextALignLeft, , 85, False, , , , , , CCLSortString
        .AddColumn "K2", "Tipo", ecgHdrTextALignLeft, , 80, , , , , , , CCLSortString
        .AddColumn "K3", "Cant.", ecgHdrTextALignRight, , 49, , , , , , , CCLSortNumeric
        .AddColumn "K4", "Prenda", ecgHdrTextALignLeft, , 212, , , , , , , CCLSortNumeric
        .AddColumn "K5", "Peso", ecgHdrTextALignRight, , 58, , , , , , , CCLSortNumeric
        .AddColumn "K6", "Peso P.", ecgHdrTextALignRight, , 58, , , , , , , CCLSortNumeric
        .AddColumn "K7", "Kílates", ecgHdrTextALignRight, , 50, , , , , , , CCLSortString
        .AddColumn "K8", "Avalúo", ecgHdrTextALignRight, , 90, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "K9", "Valor", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    m_NumContrato = 0
    Quitar_Flat Fl()
End Sub

Private Sub grdRefrendos_Click(ByVal lRow As Long, ByVal lCol As Long)
Dim i As Integer, iCount As Integer, columna As Integer

    If lRow > 0 And lCol > 0 Then
        grdPrendas.Clear

        If chkAutomovil.Value = 1 Then
            iCount = UBound(DetallesEmpenoFA)
        Else
            iCount = UBound(DetallesEmpenoForaneo)
        End If
    
        For i = 0 To iCount
        
            grdPrendas.AddRow
            
            If chkAutomovil.Value = 1 Then
                With DetallesEmpenoFA(i)
                    grdPrendas.CellDetails grdPrendas.Rows, 1, Foraneo.NumContrato, DT_CENTER Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 2, .TipoDesc, DT_LEFT Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 3, "1", DT_CENTER Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 4, .MarcayModelo & " " & .Observaciones, DT_LEFT Or DT_WORD_ELLIPSIS Or DT_WORDBREAK
                    grdPrendas.CellDetails grdPrendas.Rows, 5, "", DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 6, "", DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 7, "", DT_CENTER Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 8, Foraneo.Avaluo, DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 9, Foraneo.Prestamo, DT_RIGHT Or DT_WORD_ELLIPSIS
                    
                    grdPrendas.AutoHeightRow grdPrendas.Rows
                End With
            Else
            
                With DetallesEmpenoForaneo(i)
                    grdPrendas.CellDetails grdPrendas.Rows, 1, Foraneo.NumContrato, DT_CENTER Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 2, SacaValor("Tipo", "Descripcion", " WHERE IDTabla = " & .Tipo), DT_LEFT Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 3, .Cantidad, DT_CENTER Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 4, .Articulo & ". " & .Observaciones, DT_LEFT Or DT_WORD_ELLIPSIS Or DT_WORDBREAK
                    grdPrendas.CellDetails grdPrendas.Rows, 5, .Peso, DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 6, .PesoPiedras, DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 7, SacaValor("Kilatajes", "Descripcion", " Where IDTabla = " & .Kilates), DT_CENTER Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 8, .Avaluo, DT_RIGHT Or DT_WORD_ELLIPSIS
                    grdPrendas.CellDetails grdPrendas.Rows, 9, .Prestamo, DT_RIGHT Or DT_WORD_ELLIPSIS
                    
                    grdPrendas.AutoHeightRow grdPrendas.Rows
                End With
            End If
            
            For columna = 1 To grdPrendas.Columns Step 2
                grdPrendas.CellBackColor(grdPrendas.Rows, columna) = RGB(242, 254, 255)
            Next columna
        Next i
    
    End If
End Sub

Private Sub txtAbono_Change()

    On Local Error Resume Next
    
    Dim crIntereses As Double, crAbono As Double, crDescuento As Double

        crIntereses = IIf(Me.lblIntereses.Caption = "", 0, Me.lblIntereses.Caption)
        crAbono = IIf(Me.txtAbono.text = "", 0, Me.txtAbono.text)
        crDescuento = IIf(Me.txtDescuento.text = "", 0, Me.txtDescuento.text)
    
        txtSubTotal.text = Format((crIntereses + crAbono) - crDescuento, FMoneda)
    
End Sub

Private Sub txtAbono_GotFocus()
    Seleccionar_Texto txtAbono
    Cambiar_Color True, txtAbono
End Sub

Private Sub txtAbono_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAbono_LostFocus()
    Cambiar_Color False, txtAbono
    txtAbono.text = Format(txtAbono.text, FMoneda)
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

Private Sub txtCedula_GotFocus()
    Seleccionar_Texto txtCedula
    Cambiar_Color True, txtCedula
End Sub

Private Sub txtCodigoSucursal_GotFocus()
    Seleccionar_Texto txtCodigoSucursal
    Cambiar_Color True, txtCodigoSucursal
End Sub

Private Sub txtCodigoSucursal_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Buscar_Sucursal txtCodigoSucursal.text
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCodigoSucursal_LostFocus()
    Cambiar_Color False, txtCodigoSucursal
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Texto txtDescripcion
    Cambiar_Color True, txtDescripcion
End Sub

Private Sub txtDescripcion_LostFocus()
    Cambiar_Color False, txtDescripcion
End Sub

Private Sub txtDescuento_Change()
    txtAbono_Change
End Sub

Private Sub txtDescuento_GotFocus()
    Seleccionar_Texto txtDescuento
    Cambiar_Color True, txtDescuento
End Sub

Private Sub txtDescuento_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescuento_LostFocus()
    Cambiar_Color False, txtDescuento
End Sub

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

Private Sub txtFolioRefrendo_GotFocus()
    Seleccionar_Texto txtFolioRefrendo
    Cambiar_Color True, txtFolioRefrendo
End Sub

Private Sub txtFolioRefrendo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 And Trim(txtFolioRefrendo.text) <> "" And cmbSucursales.ListIndex > -1 Then
        If VerificaContratoDuplicado(txtFolioRefrendo.text, grdRefrendos, 2) = False Then
            Buscar_Contrato
        Else
            txtFolioRefrendo.text = ""
        End If
    End If
    
    Pasar_Foco KeyAscii
    KeyAscii = Solo_Numeros(KeyAscii)
End Sub

Private Sub txtFolioRefrendo_LostFocus()
    Cambiar_Color False, txtFolioRefrendo
End Sub

Private Sub txtKT_GotFocus()
    Seleccionar_Texto txtKT
    Cambiar_Color True, txtKT
End Sub

Private Sub txtKT_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtKT_LostFocus()
    Cambiar_Color False, txtKT
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

Private Sub txtSubtotal_Change()
    Saca_Cambio
End Sub

Private Sub txtSubtotal_GotFocus()
    Seleccionar_Texto txtSubTotal
    Cambiar_Color True, txtSubTotal
End Sub

Private Sub txtSubtotal_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtSubtotal_LostFocus()
    Cambiar_Color False, txtSubTotal
End Sub

Private Sub Saca_Cambio()
Dim crTotal As Double, crEfectivo As Double
    
    If Val(txtTotalGeneral.text) > 0 Or Trim(txtTotalGeneral.text) <> "" Then
        
        crTotal = CDbl(txtTotalGeneral.text)
    Else
        
        crTotal = 0
    End If
    
    If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
        
        If IsNumeric(txtEfectivo.text) Then
            crEfectivo = CDbl(txtEfectivo.text)
        Else
            txtEfectivo.text = 0
            crEfectivo = 0
        End If
    Else
        crEfectivo = 0
    End If
    
    txtCambio.text = Format(crEfectivo - crTotal, FMoneda)
End Sub

Private Sub BuscarEmpeno(Folio As Long, Optional Conectar As Boolean = True)
On Error GoTo Error

    Screen.MousePointer = vbHourglass
    cmdAceptar.Enabled = True
    
    
    If grdRefrendos.Rows = 1 Then
        MsgBox "Sólo se puede Refrendar una boleta !!!", vbInformation, "Pago Interés Foráneo"
        txtFolioRefrendo.SetFocus
        
        Screen.MousePointer = vbDefault
        Exit Sub
    End If
    
    'Hago Conexión al Web Service
    If Conectar Then
        If LlenarEmpenoForaneo(Folio, cmbSucursales.ItemData(cmbSucursales.ListIndex), IIf(chkAutomovil.Value = 0, 1, 2)) = False Then
            MsgBox "El Contrato Especificado no Existe !!!", vbInformation, "Renovaciones Foráneas"
            txtFolioRefrendo.SetFocus
            
            Screen.MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    If Foraneo.Destino <> 0 Then
        MsgBox "Contrato " & IIf(Foraneo.Destino = 4, "Enajenado", IIf(Foraneo.Destino = 5, "Vendido", "Pagado")), vbInformation, "Renovaciones Foráneas"
        cmdAceptar.Enabled = False
    End If
    
    If Foraneo.Bloqueado = 1 Then
        MsgBox "Contrato BLOQUEADO: " & Foraneo.MotivoBloqueo, vbInformation, "Renovaciones Foráneas"
        cmdAceptar.Enabled = False
    End If
    
    'Muestro los Datos del Cliente en la Ficha
    With ClienteForaneo
        Me.CCliente.Clear
        Me.CCliente.Add ""
        Me.CCliente.Add "<bold>" & .Nombre & " " & .Apellido & "</bold>"
        Me.CCliente.Add vbCrLf & .Direccion & vbCrLf & .Colonia & vbCrLf & .Ciudad & ", " & .Estado & " CP " & .CP
    End With
    
    'Muestro los Datos del Contrato
    If Foraneo.Origen = OD_EMPENO Then
      lblFecha.Caption = Format(Foraneo.Fecha, "DD/MM/YYYY")
    ElseIf Foraneo.Origen = OD_REFRENDO Then
      lblFecha.Caption = Format(Foraneo.FechaPagoParcial, "DD/MM/YYYY")
    End If
    
    lblVencimiento.Caption = Format(Foraneo.Vencimiento, "DD/MM/YYYY")
    lblPrestamo.Caption = Format(Foraneo.Prestamo, FMoneda)
    
    CalcularRefrendo
    
Error:
    Maneja_Error Err
    Screen.MousePointer = vbDefault
End Sub

Private Sub CalcularRefrendo()

    Dim crInteresesDescuento As Double, crGastosAdmonDescuento As Double, crAlmacenajeDescuento As Double, crSeguroDescuento As Double, crMoratoriosDescuento As Double
    Dim DescuentoIntereses As Double

On Error GoTo Error
    
    'Calculo los Intereses
    crIntereses = Redondeo(GeneraInteresesForaneos(Foraneo.Prestamo, Foraneo.Tasa, Foraneo.Vencimiento, Foraneo.Fecha))
    crAlmacenaje = Redondeo(GeneraInteresesForaneos(Foraneo.Prestamo, Foraneo.Almacenaje, Foraneo.Vencimiento, Foraneo.Fecha))
    crSeguro = Redondeo(GeneraInteresesForaneos(Foraneo.Prestamo, Foraneo.Seguro, Foraneo.Vencimiento, Foraneo.Fecha))
    crGastosAdmon = Redondeo(GeneraInteresesForaneos(Foraneo.Prestamo, Foraneo.GastosAdmon, Foraneo.Vencimiento, Foraneo.Fecha))
    If Date > Foraneo.Vencimiento Then crMoratorios = Redondeo(Foraneo.Prestamo * (Regresa_Valor_BD("Operacion") / 100))
    
    crImporteDescuento = 0
    
    crIva = Redondeo((crIntereses + crAlmacenaje + crSeguro + crGastosAdmon) * (Foraneo.Iva / 100))
    
    crRedondeo = 0 'Redondeo(CCur(crIntereses + crGastosAdmon + crAlmacenaje + crSeguro + crIva)) + Redondeo(CCur(crMoratorios))
    
    lblIntereses.Caption = Format(crIntereses + crAlmacenaje + crSeguro + crGastosAdmon + crMoratorios + crIva + crRedondeo, FMoneda)
    
    txtAbono.text = "0.00"
    txtDescuento.text = "0.00"
    txtSubTotal.text = lblIntereses.Caption

Error:
    Maneja_Error Err
End Sub

Private Sub CargarDatosGrid(Folio As Long, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double, _
    crMinimo As Double, crPerdida As Double, crAbono As Double, crDescuento As Double, crGastosAdmon As Double, crRedondeo As Double, crImporteDescuento)

Dim columna As Integer

   'Cargo los Datos en el Grid
    grdRefrendos.Redraw = False
    grdRefrendos.AddRow
    grdRefrendos.CellText(1, 1) = Foraneo.NumContrato
    grdRefrendos.CellIcon(1, 1) = 3
    grdRefrendos.CellItemData(1, 1) = Foraneo.ID
    grdRefrendos.CellTextAlign(1, 1) = DT_CENTER Or DT_WORD_ELLIPSIS
    grdRefrendos.CellText(1, 2) = Foraneo.Prestamo
    grdRefrendos.CellItemData(1, 2) = Foraneo.ID
    grdRefrendos.CellTextAlign(1, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellText(1, 3) = crAbono
    grdRefrendos.CellItemData(1, 3) = Foraneo.Serie
    grdRefrendos.CellTextAlign(1, 3) = DT_RIGHT
    grdRefrendos.CellText(1, 4) = crIntereses + crAlmacenaje + crGastosAdmon + crSeguro + crMoratorios + crRedondeo + crIva
    grdRefrendos.CellItemData(1, 4) = crPerdida
    grdRefrendos.CellTextAlign(1, 4) = DT_RIGHT Or DT_WORD_ELLIPSIS
    
    grdRefrendos.CellText(1, 12) = (crAlmacenaje + crSeguro + crMoratorios + crPerdida)
    grdRefrendos.CellTextAlign(1, 12) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellText(1, 13) = crIva
    grdRefrendos.CellTextAlign(1, 13) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellText(1, 14) = crDescuento
    grdRefrendos.CellTextAlign(1, 14) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdRefrendos.CellItemData(1, 14) = cmbSucursales.ItemData(cmbSucursales.ListIndex)
    
    grdRefrendos.CellText(1, 5) = crIntereses
    grdRefrendos.CellText(1, 6) = crAlmacenaje
    grdRefrendos.CellText(1, 7) = crSeguro
    grdRefrendos.CellText(1, 8) = crIva
    'Si es contrato de Almoneda
    grdRefrendos.CellText(1, 9) = 0 'ContratoAlmoneda
    grdRefrendos.CellText(1, 10) = crPerdida
    grdRefrendos.CellText(1, 11) = crMoratorios
    grdRefrendos.CellText(1, 15) = crGastosAdmon
    grdRefrendos.CellText(1, 16) = crRedondeo
    grdRefrendos.CellText(1, 17) = crImporteDescuento

    For columna = 1 To grdRefrendos.Columns Step 2
       grdRefrendos.CellBackColor(1, columna) = RGB(242, 254, 255)
    Next columna
    
    grdRefrendos.Redraw = True
End Sub

Function VerificaContratoDuplicado(NumContrato As Long, Grid As vbalGrid, Indice As Integer) As Boolean
Dim i As Integer, Bandera As Boolean
    
    VerificaContratoDuplicado = False
    
    For i = 1 To Grid.Rows - 1
        If NumContrato = Grid.CellText(i, 1) Then VerificaContratoDuplicado = True: Exit For
    Next i
    
    If VerificaContratoDuplicado Then MsgBox "Contrato duplicado, el contrato ya esta listo para el movimiento !!", vbInformation, IIf(Indice = 1, "Retiro", "Pago Interés")
End Function

Private Sub GrabarRefrendo()

On Error GoTo Error
    
    Dim Pago As Currency, crImporteTotal As Double
    Dim Serie As Integer, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double, crPrestamo As Double, _
        crAdeudo As Double, crImporteAlmoneda As Double, crGastosAdmon As Double, crRedondeo As Double, crImporteDescuento As Double
    Dim Hora As String, IDEmpeno As Long, strSql As String, Vencimiento As Date, FolioNota As Long, Movimiento As Long, FolioNuevo As Long
    Dim i As Integer, IDEmpenoAnterior As Long, FolioAnterior As Long, crPrestamoPrenda As Double, x As Integer

    Dim PeriodosPagados As Long
    Dim FechaPagoParcial As Date
    Dim FolioConsecutivo As Long
    Dim IDUsuario As Long
    Dim IDRemesas As Long
    Dim FolioPagos As Long
    
    IDUsuario = frmMDI.IDUsuario

    Screen.MousePointer = vbHourglass
    
    For i = 1 To grdRefrendos.Rows
      'verificamos si el refrendo ya fue realizado anteriormente
      If Validar_Refrendo(Foraneo.ID, grdRefrendos.CellItemData(i, 14)) Then
         MsgBox "El Refrendo a este contrato ya ha sido realizando anteriormente, favor de verificar", vbOKOnly Or vbCritical
         Exit Sub
      End If
      
        
        'Checo si tiene Abono
        If Val(grdRefrendos.CellText(i, 3)) > 0 Or Trim(grdRefrendos.CellText(i, 3)) <> "" Then
            Pago = CDbl(grdRefrendos.CellText(i, 3))
                
            'Verifico el abono que sea mayor o igual al configurado en parámetros
            If Pago > 0 Then
                If Pago < Val(Regresa_Valor_BD("AbonoMinimo")) Then
                    MsgBox "El importe del abono es menor al autorizado !!", vbInformation, "Pago Interés"
                    Screen.MousePointer = vbDefault
                    Exit Sub
                End If
            End If
        End If
    
        'Tomo los Valores
        crImporteTotal = CDbl(txtTotalGeneral.text)
        Serie = grdRefrendos.CellItemData(i, 3)
        crIntereses = CDbl(grdRefrendos.CellText(i, 5))
        crAlmacenaje = CDbl(grdRefrendos.CellText(i, 6))
        crSeguro = CDbl(grdRefrendos.CellText(i, 7))
        crMoratorios = CDbl(grdRefrendos.CellText(i, 11))
        crIva = CDbl(grdRefrendos.CellText(i, 8))
        crGastosAdmon = CDbl(grdRefrendos.CellText(i, 15))
        crRedondeo = CDbl(grdRefrendos.CellText(i, 16))
        crImporteDescuento = CDbl(grdRefrendos.CellText(i, 17))
        crPrestamo = CDbl(grdRefrendos.CellText(i, 2))
        crAdeudo = crPrestamo - Pago
        crImporteAlmoneda = 0

        'Tomo el Nuevo Folio
        FolioNuevo = Regresa_NumContrato(False, SERIE_C)
        Regresa_NumContrato True, SERIE_C

        'Saco el Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True

        'Folio Notas
        FolioNota = Regresa_Movimiento(False, "FolioRenovacionesForaneas")
        Regresa_Movimiento True, "FolioRenovacionesForaneas"

        'Saco la nueva fecha de vencimiento
        FechaPagoParcial = Foraneo.FechaPagoParcial
        Vencimiento = Foraneo.Vencimiento
        
        'Grabo en Remesas
        PeriodosPagados = cmbPeriodosPagos.ListIndex + 1
        
        dbDatos.Execute "INSERT INTO Remesas (Fecha,SucursalOrigen,sSucursalOrigen,SucursalDestino,sSucursalDestino,IDContrato,NumContrato,IDTablaCliente,Importe,Movimiento,Perdida,Abono,FechaPagar,PeriodosPagados,PC,IDUsuario,Consecutivo,FolioNota,ImporteAlmacenaje,ImporteSeguro,ImporteGastosAdmon,ImporteIva,ImporteMoratorios,ImporteRedondeo,Efectivo,Serie,ImporteDescuento) VALUES (" & _
            "'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & grdRefrendos.CellItemData(i, 14) & ",'" & cmbSucursales.text & "'," & Sucursal.Clave & ",'" & Sucursal.NombreComercial & "'," & Foraneo.ID & "," & Foraneo.NumContrato & "," & Foraneo.IDTablaCliente & "," & _
            ConvMoneda(crIntereses) & "," & Movimiento & "," & Foraneo.Perdida & "," & ConvMoneda(Pago) & ",'" & Format(FechaPagoParcial, "YYYY/MM/DD") & "'," & _
            PeriodosPagados & ",'" & Nombre_Pc & "'," & IDUsuario & "," & FolioConsecutivo & "," & FolioNota & "," & ConvMoneda(crAlmacenaje) & "," & ConvMoneda(crSeguro) & "," & ConvMoneda(crGastosAdmon) & "," & ConvMoneda(crIva) & "," & ConvMoneda(crMoratorios) & "," & ConvMoneda(crRedondeo) & "," & ConvMoneda(txtEfectivo.text) & "," & Serie & "," & ConvMoneda(crImporteDescuento) & ")"
        
        IDRemesas = SacaValor("Remesas", "MAX(ID)")
    
        FolioAnterior = Foraneo.Folio
        
        '***** Grabamos en Axiliar *****
        Hora = Now
        
        'Grabamos el cargo entrada a caja
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,FechaModificacion,TablaMovimiento,IDMovimiento) VALUES " & _
            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & FolioAnterior & ",'RF01','110101'," & ConvMoneda(crImporteTotal) & "," & _
            TIPO_CARGO & ",1,'Pago Refrendo Foraneo','" & NombrePc & "'," & IDUsuario & "," & frmMDI.IDSucursal & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "','Remesas'," & IDRemesas & ")"
                        
        'Grabamos el abono de suministro
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,FechaModificacion,TablaMovimiento,IDMovimiento) VALUES " & _
            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & FolioAnterior & ",'RF01','310250'," & ConvMoneda(Pago + crIntereses + crAlmacenaje + crSeguro + crIva + crGastosAdmon) & "," & _
            TIPO_ABONO & ",1,'Pago Refrendo Foraneo','" & NombrePc & "'," & IDUsuario & "," & frmMDI.IDSucursal & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "','Remesas'," & IDRemesas & ")"
        
        'Cuenta Moratorios
        If crMoratorios > 0 Then
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal,FechaModificacion,TablaMovimiento,IDMovimiento) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pago Refrendo Foraneo'," & Movimiento & "," & FolioAnterior & ",'RF01','690301'," & ConvMoneda(crMoratorios) & "," & TIPO_CARGO & ",1,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "','Remesas'," & IDRemesas & ")"
        End If

        'Cuenta Redondeo
        If crRedondeo > 0 Then
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal,FechaModificacion,TablaMovimiento,IDMovimiento) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Pago Refrendo Foraneo'," & Movimiento & "," & FolioAnterior & ",'RF01','650301'," & ConvMoneda(crRedondeo) & "," & TIPO_CARGO & ",1,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "','Remesas'," & IDRemesas & ")"
        End If
        
    Next i
    
    Imprimir_Nota FolioNota, CDbl(txtTotalGeneral.text), CDbl(Pago), CCur(crIntereses + crAlmacenaje + crSeguro + crGastosAdmon + crIva + crMoratorios + crRedondeo), CCur(txtEfectivo.text), IDUsuario
    
    MsgBox "EL Refrendo ha sido grabado !!!", vbInformation, "Refrendos Foráneos"
Error:
    Maneja_Error Err
    Screen.MousePointer = vbDefault
    
    cmdLimpiar_Click
End Sub

Private Sub Cargar_Sucursales()

    Dim str As String, sucursales() As String, SucAux() As String, i As Integer
    Dim sTab As cTab

On Error GoTo Error

    Screen.MousePointer = vbHourglass
    'Set sTab = tTab.Tabs.Item("K2")
    'sTab.Enabled = False
    
    'Hago la Conexión al Web Service
    Set WSoap = New SoapClient30
    WSoap.MSSoapInit WServidor & "?wsdl"

    'Obtengo las Sucursales
    str = WSoap.GetSucursales("mrayudon", "montepio", WBaseDatos, WPuerto, WRutaServidor)
    sucursales = Split(str, "~")

    'Verifico si el cliente existe
    If Len(Trim(str)) = 0 Then Exit Sub

    cmbSucursales.Clear
    
    If UBound(sucursales) = 0 Then
        SucAux = Split(sucursales(i), "|")
        cmbSucursales.AddItem getValueArray(SucAux(), "NombreComercial")
        cmbSucursales.ItemData(cmbSucursales.NewIndex) = getValueArray(SucAux(), "Clave", True)
        
        ReDim m_Sucursales(1)
        
        With m_Sucursales(0)
            .RazonSocial = getValueArray(SucAux(), "RazonSocial")
            .NombreComercial = getValueArray(SucAux(), "NombreComercial")
            .Clave = getValueArray(SucAux(), "Clave", True)
            .Direccion = getValueArray(SucAux(), "Direccion")
            .RFC = getValueArray(SucAux(), "RFC")
        End With
        
        
    Else
        For i = 0 To UBound(sucursales) - 1
            SucAux = Split(sucursales(i), "|")
            cmbSucursales.AddItem getValueArray(SucAux(), "NombreComercial")
            cmbSucursales.ItemData(cmbSucursales.NewIndex) = getValueArray(SucAux(), "Clave", True)
            
            ReDim Preserve m_Sucursales(i + 1)
        
            With m_Sucursales(i)
                .RazonSocial = getValueArray(SucAux(), "RazonSocial")
                .NombreComercial = getValueArray(SucAux(), "NombreComercial")
                .Clave = getValueArray(SucAux(), "Clave", True)
                .Direccion = getValueArray(SucAux(), "Direccion")
                .RFC = getValueArray(SucAux(), "RFC")
            End With
        Next i
    End If
    
Error:
    'If Err.Number <> 0 Then sTab.Enabled = True
    Maneja_Error Err
    Screen.MousePointer = vbDefault
End Sub

Public Sub Imprimir_Nota(FolioPagos As Long, TotalPagar As Currency, Abono As Currency, Interes As Currency, Efectivo As Currency, IDUsuarioMov As Long, Optional Reimpresion As Boolean = False)

    Dim ImprDefault As Boolean, i As Integer
    Dim Impresora As Printer
    Dim sArticulos As String, Vencimiento As String
    Dim CambioPlan As Boolean

On Error GoTo Error
    
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        
        .ReportFileName = Path & "\Reportes\NotaForaneo.rpt"
        
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{sucursales.Clave}=" & cmbSucursales.ItemData(cmbSucursales.ListIndex)
          
        .Formulas(1) = "NumContrato=" & Foraneo.NumContrato
        
        If Foraneo.Migrada = 1 Then
            .Formulas(2) = "FolioOriginal" & Foraneo.FolioOriginal & "'"
        Else
            .Formulas(2) = "FolioOriginal" & Foraneo.NumContrato & "'"
        End If
        
        .Formulas(3) = "Fecha='" & Format(Now, "YYYY/MM/DD") & "'"
        .Formulas(4) = "Cliente='" & ClienteForaneo.Nombre & " " & ClienteForaneo.Apellido & "'"
        
        If Foraneo.Serie <> SERIE_B Then
        
            For i = 0 To UBound(DetallesEmpenoForaneo)
                If Foraneo.Migrada = 1 Then
                    sArticulos = sArticulos & DetallesEmpenoForaneo(i).Articulo & " / "
                Else
                    If DetallesEmpenoForaneo(i).Tipo = 1 Then
                        sArticulos = sArticulos & DetallesEmpenoForaneo(i).Articulo & " " & _
                         DetallesEmpenoForaneo(i).Peso & " grs. " & _
                        SacaValor("kilatajes", "descripcion", " WHERE clave = " & DetallesEmpenoForaneo(i).Kilates) & " " & _
                        IIf(DetallesEmpenoForaneo(i).PesoPiedras = 0, "s/p ", "c/p ") & _
                        DetallesEmpenoForaneo(i).Observaciones & " / "
                    Else
                        sArticulos = sArticulos & DetallesEmpenoForaneo(i).Articulo & " " & _
                        DetallesEmpenoForaneo(i).Marca & " " & _
                        DetallesEmpenoForaneo(i).Modelo & " " & _
                        DetallesEmpenoForaneo(i).Serie & " " & _
                        DetallesEmpenoForaneo(i).Observaciones & " / "
                    End If
                End If
            Next i
        Else
            For i = 0 To UBound(DetallesEmpenoFA)
                sArticulos = sArticulos & DetallesEmpenoFA(i).MarcayModelo & " " & DetallesEmpenoFA(i).Observaciones
            Next
        End If

        'sArticulos = sArticulos & "''"

        .Formulas(5) = "Descripcion='" & Mid(sArticulos, 1, Len(sArticulos) - 3) & "'"
        
        .Formulas(6) = "Abono=" & Abono
        .Formulas(7) = "Importe=" & Interes
        .Formulas(8) = "Efectivo=" & Efectivo
        
        'Vencimiento
        If Foraneo.TipoTasa = "MENSUAL" Then
            Vencimiento = Format(DateAdd("d", 30, Now), "YYYY/MM/DD")
        ElseIf Foraneo.TipoTasa = "QUINCENAL" Then
            Vencimiento = Format(DateAdd("d", 15, Now), "YYYY/MM/DD")
        Else
            Vencimiento = Format(DateAdd("d", 7, Now), "YYYY/MM/DD")
        End If
        
        .Formulas(9) = "Prestamo=" & (Foraneo.Prestamo - IIf(Reimpresion = True, 0, Abono))
        .Formulas(10) = "LeyendaPres='" & IIf(Abono > 0, "NUEVO PRÉSTAMO:", "PRÉSTAMO:") & "'"
        
        .Formulas(11) = "JuevesPromo=''"
        
        .Formulas(12) = "TasaInteres = '" & Regresa_Valor_BD("TasaTipica") & "%'"
        .Formulas(13) = "Almacenaje = '" & Regresa_Valor_BD("Almacenaje") & "%'"
        .Formulas(14) = "Seguro = '" & Regresa_Valor_BD("Seguro") & "%'"
        .Formulas(15) = "GastosAdmon = '" & Regresa_Valor_BD("GastosAdmon") & "%'"

        .Formulas(16) = "Iva = '" & Foraneo.Iva & "%'"
        .Formulas(17) = "Vencimiento='" & Vencimiento & "'"
        
        If Foraneo.TipoTasa = "MENSUAL" Then
            .Formulas(18) = "FechaComercializacion='" & Format(DateAdd("d", 10, Vencimiento), "YYYY/MM/DD") & "'"
        Else
            .Formulas(18) = "FechaComercializacion='" & Format(DateAdd("d", 5, Vencimiento), "YYYY/MM/DD") & "'"
        End If

        'Calculo los Intereses
        crIntereses = Redondeo(GeneraInteresesForaneos(Foraneo.Prestamo - IIf(Reimpresion = True, 0, Abono), Foraneo.Tasa, CDate(Vencimiento), CDate(Format(Now, "YYYY/MM/DD"))))
        crAlmacenaje = Redondeo(GeneraInteresesForaneos(Foraneo.Prestamo - IIf(Reimpresion = True, 0, Abono), Foraneo.Almacenaje, CDate(Vencimiento), CDate(Format(Now, "YYYY/MM/DD"))))
        crSeguro = Redondeo(GeneraInteresesForaneos(Foraneo.Prestamo - IIf(Reimpresion = True, 0, Abono), Foraneo.Seguro, CDate(Vencimiento), CDate(Format(Now, "YYYY/MM/DD"))))
        crGastosAdmon = Redondeo(GeneraInteresesForaneos(Foraneo.Prestamo - IIf(Reimpresion = True, 0, Abono), Foraneo.GastosAdmon, CDate(Vencimiento), CDate(Format(Now, "YYYY/MM/DD"))))
        crIva = Redondeo((crIntereses + crAlmacenaje + crSeguro + crGastosAdmon) * (Foraneo.Iva / 100))
        crRedondeo = 0 'Redondeo_Centavos(crIntereses + crGastosAdmon + crAlmacenaje + crSeguro + crIva)
    
        lblIntereses.Caption = Format(crIntereses + crAlmacenaje + crSeguro + crGastosAdmon + crIva + crRedondeo, FMoneda)
    
        .Formulas(19) = "Refrendo=" & (crIntereses + crAlmacenaje + crSeguro + crGastosAdmon + crIva)
        .Formulas(20) = "Desempeno=" & ((Foraneo.Prestamo - IIf(Reimpresion = True, 0, Abono)) + crIntereses + crAlmacenaje + crSeguro + crGastosAdmon + crIva)
        
        .Formulas(21) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(22) = "Usuario='" & SacaValor("usuarios", "Nombre", " WHERE ID=" & IDUsuarioMov) & "'"
        .Formulas(23) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .Formulas(24) = "Folio='*" & Foraneo.NumContrato & "*'"
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToWindow
        End If
          
        .WindowTitle = "Recibo"
        '.WindowState = crptMaximized
        .Action = 1
    End With
    
Error:
    Maneja_Error Err
End Sub

Private Sub Limpiar()
    lblFecha.Caption = ""
    lblVencimiento.Caption = ""
    lblPrestamo.Caption = ""
    lblIntereses.Caption = ""
    txtAbono.text = ""
    txtSubTotal.text = ""
    CCliente.Clear
    grdRefrendos.Clear
    grdPrendas.Clear
    cmbPeriodosPagos.Clear
    txtTotalGeneral.text = ""
    txtEfectivo.text = ""
    txtCambio.text = ""
    lblDescuento.Caption = "<DESCUENTO>"
End Sub

Private Sub txtTasa_Change()
    If txtTasa.text <> "" And txtPrestamoManual.text <> "" Then
        txtInteresesManual.text = Calcular_Intereses_Manual(cmbPeriodosManual.ListIndex + 1, CCur(CDbl(txtPrestamoManual.text)), CDbl(txtTasa.text))
    Else
        txtInteresesManual.text = ""
    End If
End Sub

Private Sub txtTasa_GotFocus()
    Seleccionar_Texto txtTasa
    Cambiar_Color True, txtTasa
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTasa_LostFocus()
    Cambiar_Color False, txtTasa
End Sub

Private Sub Buscar_Sucursal(Clave As Integer)
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    
    rc.Open "SELECT * FROM Sucursales WHERE Clave=" & Clave, dbDatos, adOpenDynamic, adLockOptimistic
    cSucursal.Clear
    If Not rc.EOF Then
        With cSucursal
            Sucursal.NombreComercial = rc!NombreComercial
            Sucursal.RFC = rc!RFC
            Sucursal.Telefono = rc!Telefono
        
            .Tag = rc!Clave
            .Add "<bold>" & rc!NombreComercial & "</bold>"
            .Add rc!Direccion
            .Add "Tel.:" & rc!Telefono
            .Add "Ced.:" & rc!RFC
        End With
    Else
        cSucursal.Tag = ""
        MsgBox "La sucursal no se encuentra", vbOKOnly Or vbCritical
    End If
    
    rc.Close
Error:
    Maneja_Error Err
    
    Set rc = Nothing
End Sub

Private Sub txtValor_GotFocus()
    Seleccionar_Texto txtValor
    Cambiar_Color True, txtValor
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtValor_LostFocus()
    Cambiar_Color False, txtValor
End Sub

'verifica si el refrendo actual ya fue realizado anteriormente
Private Function Validar_Refrendo(IDEmpeno As Long, Sucursal As Long) As Boolean
   On Error GoTo Error
   Dim Realizado As Boolean
   Realizado = (Val(SacaValor("Remesas", "ID", " WHERE Cancelado=0 AND IDContrato=" & IDEmpeno & " AND SucursalOrigen=" & Sucursal)) > 0)
   Validar_Refrendo = Realizado
Error:
   Maneja_Error Err
End Function

Function GeneraInteresesForaneos(ByVal Prestamo As Double, ByVal Tasa As Double, ByVal Vencimiento As Date, ByVal Fecha As Date) As Double
    Dim DiasTrans As Integer, crIntereses As Double
    Dim rcParametros As New ADODB.Recordset
    Dim DiasGracia As Integer

On Error GoTo Error
    
    DiasGracia = Val(Regresa_Valor_BD("DiasGracia"))
    DiasTrans = DateDiff("D", Fecha, Date)
    
    If DiasTrans = 0 Then DiasTrans = 1
                                                                                
    If DiasGracia > 0 Then
    
        If (Date > DateValue(Vencimiento)) And (Date <= DateAdd("D", DiasGracia, DateValue(Vencimiento))) Then
            DiasTrans = DateDiff("d", Fecha, Vencimiento)
        Else
            DiasTrans = DateDiff("d", Fecha, Date)
        End If
    End If
    '****
    crIntereses = Redondeo(Prestamo * (Tasa / 30) * DiasTrans) '!Periodo
    
    GeneraInteresesForaneos = crIntereses
    
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcParametros = Nothing
End Function

