VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmRepCartera 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte Cartera"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepCartera.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   8895
   Begin VB.TextBox txtFechaFin 
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
      TabIndex        =   23
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtFechaIni 
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
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   240
      Width           =   1455
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   0
      Left            =   120
      Top             =   1560
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   3135
      Index           =   0
      Left            =   120
      Top             =   1560
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   5530
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   3150
      Index           =   1
      Left            =   6315
      Top             =   1560
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   5556
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   3135
      Index           =   2
      Left            =   8400
      Top             =   1560
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   5530
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   30
      Index           =   0
      Left            =   120
      Top             =   3900
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   1
      Left            =   120
      Top             =   1950
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   2
      Left            =   120
      Top             =   2340
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   3
      Left            =   120
      Top             =   2730
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   4
      Left            =   120
      Top             =   3120
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   5
      Left            =   120
      Top             =   3510
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   3120
      Index           =   3
      Left            =   795
      Top             =   1590
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   5503
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   30
      Index           =   1
      Left            =   120
      Top             =   4290
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   30
      Index           =   2
      Left            =   120
      Top             =   4680
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   1
      Left            =   7440
      TabIndex        =   24
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
      Picture         =   "frmRepCartera.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   7440
      TabIndex        =   25
      Top             =   240
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
      Picture         =   "frmRepCartera.frx":0121
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   7785
      TabIndex        =   26
      Top             =   210
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
      Picture         =   "frmRepCartera.frx":0236
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7500
      TabIndex        =   29
      Top             =   4875
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
      Picture         =   "frmRepCartera.frx":05BB
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   6
      Left            =   120
      Top             =   240
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1185
      Index           =   4
      Left            =   120
      Top             =   240
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2090
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1200
      Index           =   5
      Left            =   2115
      Top             =   240
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2117
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1200
      Index           =   6
      Left            =   4080
      Top             =   240
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2117
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   7
      Left            =   120
      Top             =   630
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   8
      Left            =   120
      Top             =   1020
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Index           =   9
      Left            =   120
      Top             =   1410
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   6360
      TabIndex        =   38
      Top             =   4875
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   " &Aceptar"
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
      Picture         =   "frmRepCartera.frx":0B0D
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caja general:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   285
      TabIndex        =   35
      Top             =   330
      Width           =   1650
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Inventario:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   285
      TabIndex        =   34
      Top             =   1110
      Width           =   1365
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Custodia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   285
      TabIndex        =   33
      Top             =   720
      Width           =   1155
   End
   Begin VB.Label lblCajaGral 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2160
      TabIndex        =   32
      Top             =   330
      Width           =   1815
   End
   Begin VB.Label lblInventario 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2160
      TabIndex        =   31
      Top             =   1110
      Width           =   1815
   End
   Begin VB.Label lblCustodia 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   2160
      TabIndex        =   30
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha final:"
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
      Left            =   4200
      TabIndex        =   28
      Top             =   720
      Width           =   1410
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Fecha inicial:"
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
      Left            =   4200
      TabIndex        =   27
      Top             =   240
      Width           =   1590
   End
   Begin VB.Label lblAlmonedaNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   285
      TabIndex        =   21
      Top             =   4365
      Width           =   360
   End
   Begin VB.Label lblAlmoneda 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   6480
      TabIndex        =   20
      Top             =   4365
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contratos pasados a almoneda:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   960
      TabIndex        =   19
      Top             =   4365
      Width           =   3855
   End
   Begin VB.Label lblRefrendoNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   285
      TabIndex        =   18
      Top             =   3600
      Width           =   360
   End
   Begin VB.Label lblAboRefNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   285
      TabIndex        =   17
      Top             =   3210
      Width           =   360
   End
   Begin VB.Label lblLiquiFijaNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   285
      TabIndex        =   16
      Top             =   2820
      Width           =   360
   End
   Begin VB.Label lblFijaNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   285
      TabIndex        =   15
      Top             =   2430
      Width           =   360
   End
   Begin VB.Label lblLiquiTradNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   285
      TabIndex        =   14
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label lblTradicionalNum 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   285
      TabIndex        =   13
      Top             =   1650
      Width           =   360
   End
   Begin VB.Label lblRefrendo 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   6480
      TabIndex        =   12
      Top             =   3600
      Width           =   1815
   End
   Begin VB.Label lblAboRef 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   6480
      TabIndex        =   11
      Top             =   3210
      Width           =   1815
   End
   Begin VB.Label lblLiquiFija 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   6480
      TabIndex        =   10
      Top             =   2820
      Width           =   1815
   End
   Begin VB.Label lblLiquiTrad 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   6480
      TabIndex        =   9
      Top             =   2040
      Width           =   1815
   End
   Begin VB.Label lblFija 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   6480
      TabIndex        =   8
      Top             =   2430
      Width           =   1815
   End
   Begin VB.Label lblTradicional 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
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
      Left            =   6480
      TabIndex        =   7
      Top             =   1650
      Width           =   1815
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Refrendos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   960
      TabIndex        =   5
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Abono a capital:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   3210
      Width           =   1995
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liquidación contratos pagos fijos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   3
      Top             =   2820
      Width           =   4095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Liquidación contratos tradicionales:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   2
      Top             =   2040
      Width           =   4350
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contratos nuevos pagos fijos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   1
      Top             =   2430
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contratos nuevos tradicional:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   1650
      Width           =   3600
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Height          =   3090
      Left            =   150
      TabIndex        =   6
      Top             =   1590
      Width           =   6165
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Height          =   1140
      Left            =   150
      TabIndex        =   36
      Top             =   270
      Width           =   1965
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Height          =   1170
      Left            =   2160
      TabIndex        =   37
      Top             =   240
      Width           =   1920
   End
End
Attribute VB_Name = "frmRepCartera"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Dim acTradicional As Double, acLiqTradicional As Double, acAboRef As Double, acFija As Double, acLiqFija As Double, acRef As Double, acAlmoneda As Double

Private Sub cmdAceptar_Click()
    Imprimir
End Sub

Private Sub cmdBuscar_Click()
    
    If Trim(txtFechaIni.text) <> "" And Trim(txtFechaFin.text) <> "" Then
        
        CargarMontos CDate(txtFechaIni.text), CDate(txtFechaFin.text)
    End If
    
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
    
    If Index = 0 Then
        
        txtFechaIni.text = frmCalendario.Fecha(txtFechaIni.text, 1)
    Else
        
        txtFechaFin.text = frmCalendario.Fecha(txtFechaFin.text, 1)
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Poner_Flat Fl, Me.Controls, Me
    txtFechaIni.text = Format(Date, "DD/MMM/YYYY")
    txtFechaFin.text = Format(Date, "DD/MMM/YYYY")
    LimpiaMontos
    CentrarForm Me, frmMDI
End Sub

Sub CargarMontos(FechaIni As Date, FechaFin As Date)
Dim rcMovimientos As New ADODB.Recordset

On Error GoTo Error
    
    rcMovimientos.Open "SELECT * FROM auxiliar WHERE Fecha>='" & Format(FechaIni, "YYYY/MM/DD") & "' AND Fecha<='" & Format(FechaFin, "YYYY/MM/DD") & "' ORDER BY ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    LimpiaMontos

    With rcMovimientos
    
        While Not .EOF
            
'''            If !Cuenta = "199450" And !Concepto = "Empeño" And (!Serie = 1 Or !Serie = 2) Then
'''                acTradicional = acTradicional + !Importe
'''                lblTradicional.Tag = Val(lblTradicional.Tag) + 1
            If !Cuenta = "110150" And !Concepto = "Empeño" And (!Serie = 1 Or !Serie = 2) Then
                acTradicional = acTradicional + !Importe
                lblTradicional.Tag = Val(lblTradicional.Tag) + 1
                
            ElseIf !Cuenta = "201750" And !Concepto = "Desempeño" And (!Serie = 1 Or !Serie = 2) Then
                acLiqTradicional = acLiqTradicional + !Importe
                lblLiquiTrad.Tag = Val(lblLiquiTrad.Tag) + 1
            
'''            ElseIf !Cuenta = "199450" And !Concepto = "Empeño" And !Serie = 3 Then
'''                acFija = acFija + !Importe
'''                lblFija.Tag = Val(lblFija.Tag) + 1
            ElseIf !Cuenta = "110150" And !Concepto = "Empeño" And !Serie = 3 Then
                acFija = acFija + !Importe
                lblFija.Tag = Val(lblFija.Tag) + 1
            
            ElseIf !Cuenta = "201750" And !Concepto = "Desempeño" And !Serie = 3 Then
                acLiqFija = acLiqFija + !Importe
                lblLiquiFija.Tag = Val(lblLiquiFija.Tag) + 1
            
            ElseIf !Cuenta = "201750" And !Concepto = "Abono Refrendo" And (!Serie = 1 Or !Serie = 2) Then
                acAboRef = acAboRef + !Importe
                lblAboRef.Tag = Val(lblAboRef.Tag) + 1
                                                                                    
            ElseIf !Cuenta = "201701" And !Concepto = "Refrendo" Then
                acRef = acRef + !Importe
                lblRefrendo.Tag = Val(lblRefrendo.Tag) + 1
            
            ElseIf !Cuenta = "201750" And (!Concepto = "Almoneda") Then
                acAlmoneda = acAlmoneda + !Importe
                lblAlmoneda.Tag = Val(lblAlmoneda.Tag) + 1
                
            End If

        .MoveNext
        Wend
        rcMovimientos.Close
        Set rcMovimientos = Nothing
        
        lblCajaGral.Caption = Format(Regresa_Saldo("110901", "110950", " AND Fecha<'" & Format(FechaIni, "YYYY/MM/DD") & "'"), FMoneda)
        lblCustodia.Caption = Format(Regresa_Saldo("201701", "201750", " AND Fecha<'" & Format(FechaIni, "YYYY/MM/DD") & "'"), FMoneda)
        lblInventario = Format(Regresa_Saldo("620301", "620350", " AND Fecha<'" & Format(FechaIni, "YYYY/MM/DD") & "'"), FMoneda)
        
        lblTradicional.Caption = Format(acTradicional, FMoneda)
        lblTradicionalNum.Caption = "(" & lblTradicional.Tag & ")"
        lblLiquiTrad.Caption = Format(acLiqTradicional, FMoneda)
        lblLiquiTradNum.Caption = "(" & lblLiquiTrad.Tag & ")"
        lblFija.Caption = Format(acFija, FMoneda)
        lblFijaNum.Caption = "(" & lblFija.Tag & ")"
        lblLiquiFija.Caption = Format(acLiqFija, FMoneda)
        lblLiquiFijaNum.Caption = "(" & lblLiquiFija.Tag & ")"
        lblAboRef.Caption = Format(acAboRef, FMoneda)
        lblAboRefNum.Caption = "(" & lblAboRef.Tag & ")"
        lblRefrendo.Caption = Format(acRef, FMoneda)
        lblRefrendoNum.Caption = "(" & lblRefrendo.Tag & ")"
        lblAlmoneda.Caption = Format(acAlmoneda, FMoneda)
        lblAlmonedaNum.Caption = "(" & lblAlmoneda.Tag & ")"
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcMovimientos = Nothing
End Sub

Sub LimpiaMontos()
    acTradicional = 0
    acLiqTradicional = 0
    acAboRef = 0
    acFija = 0
    acLiqFija = 0
    acRef = 0
    acAlmoneda = 0
    
    lblCajaGral.Caption = "0.00"
    lblCustodia.Caption = "0.00"
    lblInventario.Caption = "0.00"
    lblTradicional.Caption = "0.00"
    lblTradicional.Tag = 0
    lblLiquiTrad.Caption = "0.00"
    lblLiquiTrad.Tag = 0
    lblFija.Caption = "0.00"
    lblFija.Tag = 0
    lblLiquiFija.Caption = "0.00"
    lblLiquiFija.Tag = 0
    lblAboRef.Caption = "0.00"
    lblAboRef.Tag = 0
    lblRefrendo.Caption = "0.00"
    lblRefrendo.Tag = 0
    lblAlmoneda.Caption = "0.00"
    lblAlmoneda.Tag = 0
End Sub

Private Sub txtFechaFin_GotFocus()
    Seleccionar_Texto txtFechaFin
    Cambiar_Color True, txtFechaFin
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaFin_LostFocus()
    Cambiar_Color False, txtFechaFin
End Sub

Private Sub txtFechaIni_GotFocus()
    Seleccionar_Texto txtFechaIni
    Cambiar_Color True, txtFechaIni
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaIni_LostFocus()
    Cambiar_Color False, txtFechaIni
End Sub

Function Regresa_Saldo(Cargo As String, Abono As String, Optional strCondicion As String = "") As Double
Dim rcAux As New ADODB.Recordset
Dim crCargo As Double, crAbono As Double

    With rcAux
                    
        .Open "SELECT Sum(Importe) AS Cargo FROM auxiliar WHERE Cuenta='" & Trim(Cargo) & "'" & strCondicion, dbDatos, adOpenForwardOnly, adLockOptimistic
            crCargo = IIf(IsNull(!Cargo), 0, !Cargo)
        .Close

        .Open "SELECT Sum(Importe) AS total FROM auxiliar WHERE Cuenta='" & Trim(Abono) & "'" & strCondicion, dbDatos, adOpenForwardOnly, adLockOptimistic
            crAbono = IIf(IsNull(!Total), 0, !Total)
        .Close
        Set rcAux = Nothing
    End With
    
    Regresa_Saldo = crCargo - crAbono
End Function

Sub Imprimir()

    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .ReportFileName = Path & "\Reportes\RepCartera.rpt"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Leyenda='De la fecha " & Format(txtFechaIni.text, "dd/mmm/yyyy") & " a " & Format(txtFechaFin.text, "dd/mmm/yyyy") & "'"
        
        .Formulas(3) = "CajaGral=" & ConvMoneda(lblCajaGral.Caption) & ""
        .Formulas(4) = "Custodia=" & ConvMoneda(lblCustodia.Caption) & ""
        .Formulas(5) = "Vitrina=" & ConvMoneda(lblInventario.Caption) & ""
        
        .Formulas(6) = "ConNuevosTrad=" & ConvMoneda(lblTradicional.Caption) & ""
        .Formulas(7) = "ConNuevosTradNum='" & lblTradicionalNum.Caption & "'"

        .Formulas(8) = "LiqConTrad=" & ConvMoneda(lblLiquiTrad.Caption) & ""
        .Formulas(9) = "LiqConTradNum='" & lblLiquiTradNum.Caption & "'"

        .Formulas(10) = "ConNuevosPagos=" & ConvMoneda(lblFija.Caption) & ""
        .Formulas(11) = "ConNuevosPagosNum='" & lblFijaNum.Caption & "'"

        .Formulas(12) = "LiqConPagos=" & ConvMoneda(lblLiquiFija.Caption) & ""
        .Formulas(13) = "LiqConPagosNum='" & lblLiquiFijaNum.Caption & "'"

        .Formulas(14) = "AbonosCapital=" & ConvMoneda(lblAboRef.Caption) & ""
        .Formulas(15) = "AbonosCapitalNum='" & lblAboRefNum.Caption & "'"

        .Formulas(16) = "Refrendos=" & ConvMoneda(lblRefrendo.Caption) & ""
        .Formulas(17) = "RefrendosNum='" & lblRefrendoNum.Caption & "'"

        .Formulas(18) = "ConAlmoneda=" & ConvMoneda(lblAlmoneda.Caption) & ""
        .Formulas(19) = "ConAlmonedaNum='" & lblAlmonedaNum.Caption & "'"
        
        .WindowTitle = "Reporte de cartera"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
        
End Sub
