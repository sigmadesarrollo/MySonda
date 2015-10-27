VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmConfiguracionPrecio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios Oro"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5070
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfiguracionPrecio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3735
   ScaleWidth      =   5070
   Begin VB.TextBox txtTipoCambio 
      Alignment       =   1  'Right Justify
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
      Height          =   225
      Left            =   1800
      TabIndex        =   7
      Top             =   600
      Width           =   1215
   End
   Begin Line3D.ucLine3D ucLine3D14 
      Height          =   45
      Left            =   1080
      Top             =   930
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   79
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D13 
      Height          =   45
      Left            =   135
      Top             =   3120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   79
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D12 
      Height          =   1950
      Left            =   105
      Top             =   1200
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   3440
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D11 
      Height          =   30
      Left            =   120
      Top             =   2835
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D10 
      Height          =   30
      Left            =   120
      Top             =   2550
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D9 
      Height          =   30
      Left            =   135
      Top             =   2265
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D8 
      Height          =   30
      Index           =   0
      Left            =   120
      Top             =   1980
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D7 
      Height          =   30
      Left            =   120
      Top             =   1440
      Width           =   4785
      _ExtentX        =   8440
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D6 
      Height          =   2205
      Left            =   4920
      Top             =   930
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3889
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D5 
      Height          =   1950
      Left            =   3915
      Top             =   1200
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3440
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D4 
      Height          =   1920
      Left            =   2985
      Top             =   1215
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3387
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   1935
      Left            =   2040
      Top             =   1200
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3413
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Left            =   120
      Top             =   1200
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   2205
      Left            =   1065
      Top             =   930
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3889
      Orientation     =   0
      LineWidth       =   2
   End
   Begin VB.TextBox txtCentenario 
      Alignment       =   1  'Right Justify
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
      Height          =   225
      Left            =   1800
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3825
      TabIndex        =   42
      Top             =   3240
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
      Picture         =   "frmConfiguracionPrecio.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdGuardar 
      Height          =   375
      Left            =   2625
      TabIndex        =   43
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Guardar"
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
      Picture         =   "frmConfiguracionPrecio.frx":055E
   End
   Begin Line3D.ucLine3D ucLine3D8 
      Height          =   30
      Index           =   1
      Left            =   120
      Top             =   1710
      Width           =   4800
      _ExtentX        =   8467
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin VB.Label lblPrecioBase 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Index           =   5
      Left            =   5040
      TabIndex        =   49
      Top             =   2280
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   3930
      TabIndex        =   48
      Tag             =   "22,4"
      Top             =   1740
      Width           =   975
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   3015
      TabIndex        =   47
      Tag             =   "22,3"
      Top             =   1740
      Width           =   900
   End
   Begin VB.Label lblB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   2040
      TabIndex        =   46
      Tag             =   "22,2"
      Top             =   1740
      Width           =   945
   End
   Begin VB.Label lblKilates 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "22"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   5
      Left            =   487
      TabIndex        =   45
      Tag             =   ".900"
      Top             =   1755
      Width           =   240
   End
   Begin VB.Label lblEX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   44
      Tag             =   "22,1"
      Top             =   1740
      Width           =   960
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de cambio:"
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
      Top             =   600
      Width           =   1500
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "CALIDAD VALOR AVALÚO"
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
      Height          =   255
      Left            =   1095
      TabIndex        =   40
      Top             =   960
      Width           =   3840
   End
   Begin VB.Label lblKilates 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   4
      Left            =   547
      TabIndex        =   5
      Tag             =   ".333"
      Top             =   2887
      Width           =   120
   End
   Begin VB.Label lblKilates 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   3
      Left            =   487
      TabIndex        =   4
      Tag             =   ".417"
      Top             =   2602
      Width           =   240
   End
   Begin VB.Label lblKilates 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "14"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   2
      Left            =   487
      TabIndex        =   3
      Tag             =   ".583"
      Top             =   2317
      Width           =   240
   End
   Begin VB.Label lblKilates 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   1
      Left            =   487
      TabIndex        =   2
      Tag             =   ".700"
      Top             =   2032
      Width           =   240
   End
   Begin VB.Label lblKilates 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "24"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   210
      Index           =   0
      Left            =   487
      TabIndex        =   1
      Tag             =   "1"
      Top             =   1462
      Width           =   240
   End
   Begin VB.Label lblM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   3930
      TabIndex        =   38
      Tag             =   "14,4"
      Top             =   2865
      Width           =   975
   End
   Begin VB.Label lblM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   3930
      TabIndex        =   37
      Tag             =   "1,4"
      Top             =   2580
      Width           =   975
   End
   Begin VB.Label lblM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   3930
      TabIndex        =   36
      Tag             =   "2,4"
      Top             =   2295
      Width           =   975
   End
   Begin VB.Label lblM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   3930
      TabIndex        =   35
      Tag             =   "3,4"
      Top             =   2010
      Width           =   975
   End
   Begin VB.Label lblM 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   3930
      TabIndex        =   34
      Tag             =   "21,4"
      Top             =   1440
      Width           =   975
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2970
      TabIndex        =   33
      Tag             =   "14,3"
      Top             =   2865
      Width           =   945
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   2970
      TabIndex        =   32
      Tag             =   "1,3"
      Top             =   2580
      Width           =   945
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2970
      TabIndex        =   31
      Tag             =   "2,3"
      Top             =   2295
      Width           =   945
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2970
      TabIndex        =   30
      Tag             =   "3,3"
      Top             =   2010
      Width           =   945
   End
   Begin VB.Label lblR 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   3015
      TabIndex        =   29
      Tag             =   "21,3"
      Top             =   1440
      Width           =   900
   End
   Begin VB.Label lblB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   2040
      TabIndex        =   28
      Tag             =   "14,2"
      Top             =   2865
      Width           =   945
   End
   Begin VB.Label lblB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   27
      Tag             =   "1,2"
      Top             =   2580
      Width           =   945
   End
   Begin VB.Label lblB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   2040
      TabIndex        =   26
      Tag             =   "2,2"
      Top             =   2295
      Width           =   945
   End
   Begin VB.Label lblB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   25
      Tag             =   "3,2"
      Top             =   2010
      Width           =   945
   End
   Begin VB.Label lblB 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   2040
      TabIndex        =   24
      Tag             =   "21,2"
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label lblPrecioBase 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Index           =   4
      Left            =   5025
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPrecioBase 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Index           =   1
      Left            =   5025
      TabIndex        =   22
      Top             =   1320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPrecioBase 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Index           =   2
      Left            =   5025
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPrecioBase 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Index           =   3
      Left            =   5025
      TabIndex        =   20
      Top             =   1800
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblPrecioBase 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Index           =   0
      Left            =   5025
      TabIndex        =   19
      Top             =   1080
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label lblEX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   18
      Tag             =   "14,1"
      Top             =   2865
      Width           =   960
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "KILATES"
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
      Height          =   210
      Left            =   150
      TabIndex        =   17
      Top             =   1215
      Width           =   915
   End
   Begin VB.Label lblEX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   16
      Tag             =   "1,1"
      Top             =   2580
      Width           =   960
   End
   Begin VB.Label lblEX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   15
      Tag             =   "2,1"
      Top             =   2295
      Width           =   960
   End
   Begin VB.Label lblEX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   14
      Tag             =   "3,1"
      Top             =   2010
      Width           =   960
   End
   Begin VB.Label lblEX 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   13
      Tag             =   "21,1"
      Top             =   1440
      Width           =   960
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "M"
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
      Height          =   210
      Left            =   3960
      TabIndex        =   12
      Top             =   1215
      Width           =   930
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "R"
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
      Height          =   210
      Left            =   3015
      TabIndex        =   11
      Top             =   1215
      Width           =   930
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "B"
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
      Height          =   210
      Left            =   2070
      TabIndex        =   10
      Top             =   1215
      Width           =   930
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "EX"
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
      Height          =   210
      Left            =   1110
      TabIndex        =   9
      Top             =   1215
      Width           =   930
   End
   Begin VB.Label lblCentenario 
      Caption         =   "Label7"
      Height          =   255
      Left            =   3120
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Onza Troy Dlls."
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
      TabIndex        =   0
      Top             =   240
      Width           =   1425
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000013&
      Height          =   1710
      Left            =   150
      TabIndex        =   39
      Top             =   1395
      Width           =   915
   End
End
Attribute VB_Name = "frmConfiguracionPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdGuardar_Click()
Dim Centenario As Double, TipoCambio As Double
    
    If Val(txtCentenario.text) > 0 Or (Trim(txtCentenario.text) <> "" And Trim(txtCentenario.text) <> ".") Then
        
        Centenario = CDbl(txtCentenario.text)
    Else
        
        Centenario = 0
    End If
    
    If Val(txtTipoCambio.text) > 0 Or (Trim(txtTipoCambio.text) <> "" And Trim(txtTipoCambio.text) <> ".") Then
        
        TipoCambio = CDbl(txtTipoCambio.text)
    Else
    
        TipoCambio = 0
    End If
    
    dbDatos.Execute "UPDATE parametros SET Centenario=" & Centenario & ",TipoCambioOnza=" & TipoCambio
    ActualizaValores lblEX
    ActualizaValores lblB
    ActualizaValores lblR
    ActualizaValores lblM
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtCentenario_Change()
    SacaBase
End Sub

Private Sub txtCentenario_GotFocus()
    Seleccionar_Texto txtCentenario
    Cambiar_Color True, txtCentenario
End Sub

Private Sub txtCentenario_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCentenario_LostFocus()
    Cambiar_Color False, txtCentenario
End Sub

Sub Inicializar()
    txtCentenario.text = Regresa_Valor_BD("Centenario")
    txtTipoCambio.text = Regresa_Valor_BD("TipoCambioOnza")
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Function SacaBase()
Dim i As Integer, Centenario As Double, TipoCambio As Double, PorEx As Double, PorB As Double, PorR As Double, PorM As Double

    PorEx = Regresa_Valor_BD("CalidadEx") / 100
    PorB = Regresa_Valor_BD("CalidadB") / 100
    PorR = Regresa_Valor_BD("CalidadR") / 100
    PorM = Regresa_Valor_BD("CalidadM") / 100
    
    If Val(txtCentenario) > 0 Or (Trim(txtCentenario.text) <> "" And Trim(txtCentenario.text) <> ".") Then
        
        Centenario = CDbl(txtCentenario.text)
    Else
        
        Centenario = 0
    End If
    
    If Val(txtTipoCambio.text) > 0 Or (Trim(txtTipoCambio.text) <> "" And Trim(txtTipoCambio.text) <> ".") Then
        
        TipoCambio = CDbl(txtTipoCambio.text)
    Else
        
        TipoCambio = 0
    End If
    
    lblCentenario.Caption = Format((Centenario / 31.1) * TipoCambio, FMoneda)
    
    For i = 0 To 5
        lblPrecioBase(i).Caption = Round(CDbl(lblKilates(i).Tag) * CDbl(lblCentenario.Caption), 2)
        lblEX(i).Caption = Format(Round(CDbl(lblPrecioBase(i).Caption) * PorEx, 2), FMoneda)
        lblB(i).Caption = Format(Round(CDbl(lblPrecioBase(i).Caption) * PorB, 2), FMoneda)
        lblR(i).Caption = Format(Round(CDbl(lblPrecioBase(i).Caption) * PorR, 2), FMoneda)
        lblM(i).Caption = Format(Round(CDbl(lblPrecioBase(i).Caption) * PorM, 2), FMoneda)
    Next i

End Function

Function ActualizaValores(Etiqueta As Object)
    Dim i As Integer, IDKilataje As Integer, IDHechura As Integer

    For i = 0 To 5
        IDKilataje = Mid(Etiqueta(i).Tag, 1, (InStr(1, Etiqueta(i).Tag, ",")) - 1)
        IDHechura = Mid(Etiqueta(i).Tag, (InStr(1, Etiqueta(i).Tag, ",")) + 1, Len(Etiqueta(i).Tag) - InStr(1, Etiqueta(i).Tag, ",") + 1)
    
        dbDatos.Execute "UPDATE precioskilataje SET Precio=" & CDbl(Etiqueta(i).Caption) & " WHERE IDTipo=1 AND IDKilataje=" & IDKilataje & " AND IDHechura=" & IDHechura
    Next i

End Function

Private Sub txtTipoCambio_Change()
    SacaBase
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
    Cambiar_Color False, txtTipoCambio
End Sub
