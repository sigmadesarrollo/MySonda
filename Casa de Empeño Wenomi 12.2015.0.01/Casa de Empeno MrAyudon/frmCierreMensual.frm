VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmCierreMensual 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de sucursal"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   Icon            =   "frmCierreMensual.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   11460
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
      Left            =   1800
      TabIndex        =   0
      Top             =   195
      Width           =   1335
   End
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
      Left            =   5220
      TabIndex        =   1
      Top             =   195
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   8160
      Left            =   30
      TabIndex        =   2
      Top             =   795
      Width           =   11415
      Begin Line3D.ucLine3D ucLine3D4 
         Height          =   810
         Left            =   5640
         Top             =   165
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1429
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D3 
         Height          =   4515
         Left            =   3480
         Top             =   1530
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   7964
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   825
         Left            =   165
         Top             =   165
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1455
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   4920
         Index           =   0
         Left            =   165
         Top             =   1140
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   8678
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   3
         Left            =   180
         Top             =   1530
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   4
         Left            =   180
         Top             =   1905
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   5
         Left            =   180
         Top             =   2280
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   6
         Left            =   180
         Top             =   2655
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   7
         Left            =   180
         Top             =   3030
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   9
         Left            =   180
         Top             =   3780
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   10
         Left            =   180
         Top             =   4155
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   11
         Left            =   180
         Top             =   4530
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   12
         Left            =   180
         Top             =   4905
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   4905
         Index           =   13
         Left            =   5640
         Top             =   1140
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   8652
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   17
         Left            =   165
         Top             =   6030
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   1
         Left            =   165
         Top             =   165
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   18
         Left            =   165
         Top             =   570
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   8
         Left            =   180
         Top             =   3405
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   2
         Left            =   165
         Top             =   1140
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   780
         Index           =   20
         Left            =   165
         Top             =   8520
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1376
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   21
         Left            =   180
         Top             =   8520
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   22
         Left            =   180
         Top             =   8895
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   23
         Left            =   180
         Top             =   9270
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   780
         Index           =   24
         Left            =   5640
         Top             =   8520
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1376
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D5 
         Height          =   750
         Index           =   1
         Left            =   3480
         Top             =   8520
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1323
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   25
         Left            =   5820
         Top             =   1140
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   4155
         Index           =   26
         Left            =   5820
         Top             =   1140
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   7329
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   27
         Left            =   5820
         Top             =   1530
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   28
         Left            =   5820
         Top             =   1905
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   29
         Left            =   5820
         Top             =   2280
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   30
         Left            =   5820
         Top             =   2655
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   31
         Left            =   5820
         Top             =   3030
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   32
         Left            =   5820
         Top             =   3405
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   33
         Left            =   5820
         Top             =   3780
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   3735
         Index           =   34
         Left            =   9120
         Top             =   1560
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   6588
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   35
         Left            =   5820
         Top             =   4155
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   39
         Left            =   5820
         Top             =   4905
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   405
         Index           =   42
         Left            =   5820
         Top             =   5640
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   714
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   43
         Left            =   5820
         Top             =   5640
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   44
         Left            =   11295
         Top             =   5640
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   45
         Left            =   5820
         Top             =   6030
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   4155
         Index           =   19
         Left            =   11295
         Top             =   1140
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   7329
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   46
         Left            =   180
         Top             =   5280
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   47
         Left            =   9120
         Top             =   5640
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1140
         Index           =   48
         Left            =   5820
         Top             =   8460
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2011
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   49
         Left            =   5835
         Top             =   8460
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   50
         Left            =   5835
         Top             =   8835
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   51
         Left            =   5835
         Top             =   9210
         Width           =   5460
         _ExtentX        =   9631
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1155
         Index           =   52
         Left            =   11295
         Top             =   8460
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2037
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D5 
         Height          =   1125
         Index           =   2
         Left            =   9135
         Top             =   8475
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1984
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   53
         Left            =   5820
         Top             =   9600
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   41
         Left            =   165
         Top             =   5655
         Width           =   5490
         _ExtentX        =   9684
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   54
         Left            =   5820
         Top             =   4530
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   55
         Left            =   165
         Top             =   6225
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1875
         Index           =   56
         Left            =   165
         Top             =   6225
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   3307
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1875
         Index           =   57
         Left            =   5640
         Top             =   6225
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   3307
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   58
         Left            =   165
         Top             =   6585
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   59
         Left            =   165
         Top             =   6960
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   60
         Left            =   165
         Top             =   7335
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   61
         Left            =   165
         Top             =   7710
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   62
         Left            =   165
         Top             =   8085
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   14
         Left            =   5820
         Top             =   6225
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1875
         Index           =   15
         Left            =   5820
         Top             =   6240
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   3307
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1875
         Index           =   16
         Left            =   11295
         Top             =   6240
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   3307
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   36
         Left            =   5820
         Top             =   6585
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   37
         Left            =   5820
         Top             =   6960
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   38
         Left            =   5820
         Top             =   7335
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   40
         Left            =   5820
         Top             =   7710
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   63
         Left            =   5820
         Top             =   8085
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D6 
         Height          =   1515
         Index           =   0
         Left            =   3480
         Top             =   6600
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2672
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D7 
         Height          =   1515
         Left            =   9120
         Top             =   6600
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2672
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D8 
         Height          =   825
         Left            =   3480
         Top             =   165
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   1455
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   64
         Left            =   165
         Top             =   975
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   65
         Left            =   5820
         Top             =   5280
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin VB.Label lblRedencionPuntos 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<RedencionPuntos>"
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
         Left            =   9120
         TabIndex        =   107
         Top             =   4590
         Width           =   2055
      End
      Begin VB.Label lblRedencion 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Redencion Puntos:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   5940
         TabIndex        =   106
         Top             =   4590
         Width           =   2445
      End
      Begin VB.Label lblCortes 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "<Cortes>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   285
         Left            =   9870
         TabIndex        =   105
         Top             =   420
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lblDotaciones 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "<Dotaciones>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   3540
         TabIndex        =   104
         Top             =   630
         Width           =   2040
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DOTACIONES A CAJA:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   270
         TabIndex        =   103
         Top             =   615
         Width           =   3075
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEPOSITARÍA"
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
         Index           =   5
         Left            =   7560
         TabIndex        =   99
         Top             =   6240
         Width           =   1995
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Inicial:"
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
         Index           =   11
         Left            =   5940
         TabIndex        =   97
         Top             =   6645
         Width           =   1590
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entradas:"
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
         Index           =   10
         Left            =   5940
         TabIndex        =   96
         Top             =   7020
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salidas:"
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
         Left            =   5940
         TabIndex        =   95
         Top             =   7395
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Final:"
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
         Left            =   5940
         TabIndex        =   94
         Top             =   7770
         Width           =   1425
      End
      Begin VB.Label lblSaldoInicialDep 
         Alignment       =   1  'Right Justify
         Caption         =   "<Saldo Inicial>"
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
         Left            =   9120
         TabIndex        =   93
         Top             =   6660
         Width           =   2055
      End
      Begin VB.Label lblEntradasDep 
         Alignment       =   1  'Right Justify
         Caption         =   "<Entradas Dep.>"
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
         Left            =   9120
         TabIndex        =   92
         Top             =   7035
         Width           =   2055
      End
      Begin VB.Label lblSalidasDep 
         Alignment       =   1  'Right Justify
         Caption         =   "<Salidas Dep.>"
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
         Left            =   9120
         TabIndex        =   91
         Top             =   7410
         Width           =   2055
      End
      Begin VB.Label lblSaldoFinalDep 
         Alignment       =   1  'Right Justify
         Caption         =   "<Saldo Final>"
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
         Left            =   9120
         TabIndex        =   90
         Top             =   7785
         Width           =   2055
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL ENTRADAS:"
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
         Index           =   1
         Left            =   255
         TabIndex        =   89
         Top             =   5700
         Width           =   2670
      End
      Begin VB.Label lblSaldoFinal 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "<Saldo Final>"
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
         Left            =   4050
         TabIndex        =   88
         Top             =   7770
         Width           =   1515
      End
      Begin VB.Label lblSalidas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "<Ventas Divisas>"
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
         Left            =   3690
         TabIndex        =   87
         Top             =   7395
         Width           =   1875
      End
      Begin VB.Label lblEntradas 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "<Compras Divisa>"
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
         Left            =   3555
         TabIndex        =   86
         Top             =   7020
         Width           =   2010
      End
      Begin VB.Label lblSaldoInicial 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "<Saldo Inicial>"
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
         TabIndex        =   85
         Top             =   6645
         Width           =   1965
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Final:"
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
         Left            =   270
         TabIndex        =   84
         Top             =   7755
         Width           =   1425
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Salidas:"
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
         Left            =   270
         TabIndex        =   83
         Top             =   7380
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Entradas:"
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
         Left            =   270
         TabIndex        =   82
         Top             =   7005
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Saldo Inicial:"
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
         Left            =   270
         TabIndex        =   81
         Top             =   6630
         Width           =   1590
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Height          =   1500
         Index           =   2
         Left            =   195
         TabIndex        =   80
         Top             =   6585
         Width           =   3300
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DIVISAS"
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
         Index           =   4
         Left            =   240
         TabIndex        =   79
         Top             =   6255
         Width           =   5415
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   165
         TabIndex        =   78
         Top             =   6225
         Width           =   5475
      End
      Begin VB.Label lblRetiros 
         Alignment       =   1  'Right Justify
         Caption         =   "<Retiros>"
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
         Left            =   9120
         TabIndex        =   77
         Top             =   1590
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retiros a bóveda:"
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
         Left            =   5940
         TabIndex        =   76
         Top             =   1575
         Width           =   2160
      End
      Begin VB.Label lblAportaciones 
         Alignment       =   1  'Right Justify
         Caption         =   "<Aportaciones>"
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
         Left            =   3510
         TabIndex        =   75
         Top             =   1590
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Aportaciones a bóveda:"
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
         Left            =   255
         TabIndex        =   74
         Top             =   1575
         Width           =   2880
      End
      Begin VB.Label lblDivisasCanceladas 
         Alignment       =   1  'Right Justify
         Caption         =   "<Divisas Canc.>"
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
         Left            =   9135
         TabIndex        =   73
         Top             =   9255
         Width           =   2055
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Divisas Canceladas:"
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
         Left            =   5910
         TabIndex        =   72
         Top             =   9240
         Width           =   2400
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contratos Cancelados:"
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
         Left            =   5910
         TabIndex        =   70
         Top             =   8895
         Width           =   2745
      End
      Begin VB.Label lblContratosCancelados 
         Alignment       =   1  'Right Justify
         Caption         =   "<Contratos Canc.>"
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
         Left            =   9135
         TabIndex        =   69
         Top             =   8895
         Width           =   2055
      End
      Begin VB.Label lblContratosAlmoneda 
         Alignment       =   1  'Right Justify
         Caption         =   "<Contratos Alm.>"
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
         Left            =   9135
         TabIndex        =   68
         Top             =   8520
         Width           =   2055
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Contratos Almoneda:"
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
         Left            =   5910
         TabIndex        =   67
         Top             =   8520
         Width           =   2595
      End
      Begin VB.Label lblAjusteSalida 
         Alignment       =   1  'Right Justify
         Caption         =   "<Ajustes Salida>"
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
         Left            =   9120
         TabIndex        =   66
         Top             =   4215
         Width           =   2055
      End
      Begin VB.Label lblAjusteEntrada 
         Alignment       =   1  'Right Justify
         Caption         =   "<Ajustes Entrada>"
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
         Left            =   3510
         TabIndex        =   65
         Top             =   5340
         Width           =   2055
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ajustes:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   255
         TabIndex        =   63
         Top             =   5355
         Width           =   1005
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ajustes:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   5940
         TabIndex        =   62
         Top             =   4215
         Width           =   1005
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "SALIDAS"
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
         Height          =   375
         Index           =   1
         Left            =   5820
         TabIndex        =   61
         Top             =   1170
         Width           =   5475
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO FINAL:"
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
         Index           =   3
         Left            =   5940
         TabIndex        =   43
         Top             =   5685
         Width           =   2025
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "<Total>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9120
         TabIndex        =   59
         Top             =   5700
         Width           =   2055
      End
      Begin VB.Label lblEmpeño 
         Alignment       =   1  'Right Justify
         Caption         =   "<Empeño>"
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
         Left            =   9120
         TabIndex        =   58
         Top             =   1965
         Width           =   2055
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Préstamos nuevos:"
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
         Left            =   5940
         TabIndex        =   57
         Top             =   1965
         Width           =   2310
      End
      Begin VB.Label lblGastos 
         Alignment       =   1  'Right Justify
         Caption         =   "<Gastos>"
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
         Left            =   9120
         TabIndex        =   56
         Top             =   3465
         Width           =   2055
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos varios:"
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
         Left            =   5940
         TabIndex        =   55
         Top             =   3465
         Width           =   1725
      End
      Begin VB.Label lblSalida 
         Alignment       =   1  'Right Justify
         Caption         =   "<TotSalida>"
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
         Left            =   9120
         TabIndex        =   54
         Top             =   4965
         Width           =   2055
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compras varios:"
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
         Left            =   5940
         TabIndex        =   53
         Top             =   2715
         Width           =   1965
      End
      Begin VB.Label lblCompraVarios 
         Alignment       =   1  'Right Justify
         Caption         =   "<Compra varios>"
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
         Left            =   9120
         TabIndex        =   52
         Top             =   2715
         Width           =   2055
      End
      Begin VB.Label lblComDivisas 
         Alignment       =   1  'Right Justify
         Caption         =   "<Compra Divisas>"
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
         Left            =   9120
         TabIndex        =   51
         Top             =   2340
         Width           =   2055
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compra divisas:"
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
         Left            =   5940
         TabIndex        =   50
         Top             =   2340
         Width           =   1935
      End
      Begin VB.Label lblIvaCompras 
         Alignment       =   1  'Right Justify
         Caption         =   "<Iva>"
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
         Left            =   9120
         TabIndex        =   49
         Top             =   3090
         Width           =   2055
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A. compras:"
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
         Left            =   5940
         TabIndex        =   48
         Top             =   3090
         Width           =   1860
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Demasías:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   285
         Left            =   5940
         TabIndex        =   47
         Top             =   3825
         Width           =   1260
      End
      Begin VB.Label lblDemasias 
         Alignment       =   1  'Right Justify
         Caption         =   "<Demasías>"
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
         Left            =   9120
         TabIndex        =   46
         Top             =   3825
         Width           =   2055
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "TOTAL SALIDAS:"
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
         Index           =   0
         Left            =   5940
         TabIndex        =   45
         Top             =   4965
         Width           =   2385
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   5835
         TabIndex        =   44
         Top             =   5670
         Width           =   3300
      End
      Begin VB.Label Label11 
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
         Left            =   255
         TabIndex        =   41
         Top             =   8580
         Width           =   1335
      End
      Begin VB.Label lblRefrendo 
         Alignment       =   1  'Right Justify
         Caption         =   "<Refrendo>"
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
         TabIndex        =   40
         Top             =   8580
         Width           =   2055
      End
      Begin VB.Label lblDescuento 
         Alignment       =   1  'Right Justify
         Caption         =   "<Descuento>"
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
         TabIndex        =   39
         Top             =   8955
         Width           =   2055
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descuento ventas:"
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
         Left            =   255
         TabIndex        =   38
         Top             =   8955
         Width           =   2250
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ventas mostrador:"
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
         Left            =   255
         TabIndex        =   37
         Top             =   3465
         Width           =   2265
      End
      Begin VB.Label lblDesempeño 
         Alignment       =   1  'Right Justify
         Caption         =   "<Desempeño>"
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
         Left            =   3510
         TabIndex        =   34
         Top             =   1965
         Width           =   2055
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Capital recuperado:"
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
         Left            =   255
         TabIndex        =   33
         Top             =   1965
         Width           =   2415
      End
      Begin VB.Label lblAboRefrendo 
         Alignment       =   1  'Right Justify
         Caption         =   "<Abono Refrendo>"
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
         Left            =   3510
         TabIndex        =   32
         Top             =   2340
         Width           =   2055
      End
      Begin VB.Label lblVentas 
         Alignment       =   1  'Right Justify
         Caption         =   "<Ventas>"
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
         Left            =   3510
         TabIndex        =   31
         Top             =   3465
         Width           =   2055
      End
      Begin VB.Label lblAboApartados 
         Alignment       =   1  'Right Justify
         Caption         =   "<Abono Aparta>"
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
         Left            =   3510
         TabIndex        =   30
         Top             =   4215
         Width           =   2055
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Abonos a apartados:"
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
         Left            =   255
         TabIndex        =   29
         Top             =   4215
         Width           =   2520
      End
      Begin VB.Label lblIntereses 
         Alignment       =   1  'Right Justify
         Caption         =   "<Intereses>"
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
         Left            =   3510
         TabIndex        =   28
         Top             =   2715
         Width           =   2055
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Intereses:"
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
         Left            =   255
         TabIndex        =   27
         Top             =   2715
         Width           =   1245
      End
      Begin VB.Label lblEntrada 
         Alignment       =   1  'Right Justify
         Caption         =   "<TotEntrada>"
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
         Left            =   3510
         TabIndex        =   26
         Top             =   5715
         Width           =   2055
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venta divisas:"
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
         Left            =   255
         TabIndex        =   25
         Top             =   4590
         Width           =   1710
      End
      Begin VB.Label lblVenDivisas 
         Alignment       =   1  'Right Justify
         Caption         =   "<Venta Divisas>"
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
         Left            =   3510
         TabIndex        =   24
         Top             =   4590
         Width           =   2055
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Otros cobros:"
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
         Left            =   255
         TabIndex        =   23
         Top             =   4965
         Width           =   1620
      End
      Begin VB.Label lblOtrosCobros 
         Alignment       =   1  'Right Justify
         Caption         =   "<Otros cobros>"
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
         Left            =   3510
         TabIndex        =   22
         Top             =   4965
         Width           =   2055
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A. Intereses:"
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
         Left            =   255
         TabIndex        =   21
         Top             =   3090
         Width           =   1995
      End
      Begin VB.Label Label5 
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
         Index           =   1
         Left            =   255
         TabIndex        =   20
         Top             =   2340
         Width           =   1995
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A. Ventas:"
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
         Left            =   255
         TabIndex        =   19
         Top             =   3825
         Width           =   1680
      End
      Begin VB.Label lblIvaIntereses 
         Alignment       =   1  'Right Justify
         Caption         =   "<Iva Intereses>"
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
         Left            =   3510
         TabIndex        =   18
         Top             =   3090
         Width           =   2055
      End
      Begin VB.Label lblIvaVentas 
         Alignment       =   1  'Right Justify
         Caption         =   "<Iva Ventas>"
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
         Left            =   3510
         TabIndex        =   17
         Top             =   3825
         Width           =   2055
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         Caption         =   "ENTRADAS"
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
         Height          =   375
         Index           =   0
         Left            =   165
         TabIndex        =   16
         Top             =   1170
         Width           =   5475
      End
      Begin VB.Label lblBoveda 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "<Boveda>"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   4080
         TabIndex        =   15
         Top             =   240
         Width           =   1500
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO INICIAL:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   300
         Left            =   270
         TabIndex        =   14
         Top             =   225
         Width           =   2310
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Height          =   4500
         Index           =   0
         Left            =   195
         TabIndex        =   35
         Top             =   1560
         Width           =   3300
      End
      Begin VB.Label Label12 
         BackColor       =   &H00808080&
         Height          =   795
         Left            =   180
         TabIndex        =   36
         Top             =   195
         Width           =   3315
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   780
         Index           =   1
         Left            =   195
         TabIndex        =   42
         Top             =   8520
         Width           =   3300
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Height          =   3750
         Index           =   1
         Left            =   5820
         TabIndex        =   60
         Top             =   1545
         Width           =   3300
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Height          =   195
         Index           =   0
         Left            =   5820
         TabIndex        =   64
         Top             =   4500
         Width           =   45
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   1140
         Index           =   2
         Left            =   5835
         TabIndex        =   71
         Top             =   8490
         Width           =   3300
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   5820
         TabIndex        =   100
         Top             =   6240
         Width           =   5475
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Height          =   1500
         Index           =   3
         Left            =   5820
         TabIndex        =   98
         Top             =   6600
         Width           =   3300
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   1
      Left            =   6615
      TabIndex        =   8
      Top             =   195
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
      Picture         =   "frmCierreMensual.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   3195
      TabIndex        =   9
      Top             =   195
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
      Picture         =   "frmCierreMensual.frx":0121
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   375
      Left            =   6990
      TabIndex        =   12
      Top             =   150
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
      Picture         =   "frmCierreMensual.frx":0236
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10290
      TabIndex        =   101
      Top             =   8985
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
      Picture         =   "frmCierreMensual.frx":05BB
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   9045
      TabIndex        =   102
      Top             =   8985
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "        &Aceptar"
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
      Picture         =   "frmCierreMensual.frx":0B0D
   End
   Begin VB.Label lblCajero 
      AutoSize        =   -1  'True
      Caption         =   "<Cajero>"
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
      Left            =   9285
      TabIndex        =   13
      Top             =   435
      Width           =   1050
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
      Left            =   120
      TabIndex        =   11
      Top             =   195
      Width           =   1590
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
      Left            =   3720
      TabIndex        =   10
      Top             =   195
      Width           =   1410
   End
   Begin VB.Label lblFecha 
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
      Left            =   4920
      TabIndex        =   7
      Top             =   960
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblCaja 
      AutoSize        =   -1  'True
      Caption         =   "<Caja>"
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
      Left            =   9285
      TabIndex        =   6
      Top             =   75
      Width           =   825
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   4080
      TabIndex        =   5
      Top             =   960
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cajero:"
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
      Left            =   8340
      TabIndex        =   4
      Top             =   435
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Caja:"
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
      Left            =   8340
      TabIndex        =   3
      Top             =   75
      Width           =   645
   End
End
Attribute VB_Name = "frmCierreMensual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim ccrSaldoBov As Double, acDotaciones As Double, ccrSaldoDep As Double, acApor As Double, acReti As Double, acDesc As Double, acCam As Double, acAbo As Double, acVen As Double, acEmp As Double, acDev As Double, acDes As Double, acRef As Double, acAboRef As Double, acInt As Double, acTot As Double, acGas As Double, acPres As Double, acPresO As Double, acPresP As Double, acAjusteEntrada As Double
Dim aInventario As Double, sInventario As Double, Efectivo As Double, acDolares As Double, ccDolares As Double, acIvaInt As Double, acIvaVen As Double, ccIva As Double, ccOtros As Double, acDemasia As Double, acAjusteSalida As Double, acCortes As Double, acDepositaria As Double
'***Puntos***
Dim acRedencion As Double

Private Sub cmdAceptar_Click()

    If lblDesempeño.Caption = "" Then
        
        MsgBox "Favor de buscar primero los importes con los rangos de fecha seleccionados", vbInformation, "Cierre de sucursal"
    Else

        Imprimir
    End If

End Sub

Private Sub cmdBuscar_Click()

    If Trim(txtFechaIni.text) = "" Or Trim(txtFechaFin.text) = "" Then
        
        MsgBox "Favor de seleccionar el rango de fechas !!", vbInformation, "Cierre Mensual"
    
    Else
        
        Screen.MousePointer = vbHourglass
        
        Cargar_Montos
        
        Screen.MousePointer = vbDefault
    
    End If

End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
    
    If Index = 0 Then
        
        txtFechaIni.text = frmCalendario.Fecha(txtFechaIni.text)
    Else
        
        txtFechaFin.text = frmCalendario.Fecha(txtFechaFin.text)
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()

    Screen.MousePointer = vbHourglass
    
    LimpiaMontos
    lblFecha.Caption = Format(Date, "DD/MM/YY")
    lblCaja.Caption = NombrePc
    lblCajero.Caption = frmMDI.Usuario
    Poner_Flat Fl, Me.Controls, Me
    txtFechaIni = Format(Date, "DD/MM/YYYY")
    txtFechaFin = Format(Date, "DD/MM/YYYY")
    CentrarForm frmCierreMensual, frmMDI
    Me.Top = 0
    Screen.MousePointer = vbDefault

End Sub

Private Sub Cargar_Montos()
Dim rcVentanilla As New ADODB.Recordset
Dim Cargo As Double, Abono As Double, Entradas As Long, Salidas As Long
Dim SaldoInicialEmpenos As Double, EntradasEmpeno As Double, SalidasEmpeno As Double

    With rcVentanilla
        
        'Inicializo las variables
        LimpiaMontos
        
        'Saco el Saldo de Bóveda ******************************
        Cargo = 0
        Abono = 0
        .Open "SELECT Sum(Importe) AS Cargo FROM auxiliar WHERE Cuenta='110901' AND Fecha<'" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            Cargo = IIf(IsNull(!Cargo), 0, !Cargo)
        .Close
        
        .Open "SELECT Sum(Importe) AS total FROM auxiliar WHERE Cuenta='110950' AND Fecha<'" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            Abono = IIf(IsNull(!Total), 0, !Total)
        .Close
        ccrSaldoBov = Cargo - Abono
        '******************************************************
        
        'Saco el Saldo de Depositaría ******************************
        .Open "SELECT SUM(Prestamo) AS SaldoInicial FROM empeno WHERE Cancelado=0 AND DATE(Fecha)<='" & Format(DateAdd("D", -1, CDate(txtFechaIni.text)), "YYYY/MM/DD") & "' AND (DATE(FechaMovimiento)>'" & Format(DateAdd("D", -1, CDate(txtFechaFin.text)), "YYYY/MM/DD") & "' OR FechaMovimiento IS NULL)", dbDatos, adOpenForwardOnly, adLockReadOnly
            SaldoInicialEmpenos = IIf(IsNull(!SaldoInicial), 0, !SaldoInicial)
        .Close
        
        .Open "SELECT Sum(Prestamo) AS Cargo FROM empeno WHERE Cancelado=0 AND Origen=" & OD_EMPENO & " AND DATE(Fecha)>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND DATE(Fecha)<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            EntradasEmpeno = IIf(IsNull(!Cargo), 0, !Cargo)
        .Close
        
        .Open "SELECT Sum(Prestamo) AS Total FROM empeno WHERE Cancelado=0 AND (Destino=" & D_DESEMPEÑO & " OR Destino=" & D_ALMONEDA & ") AND DATE(FechaMovimiento)>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND DATE(FechaMovimiento)<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            SalidasEmpeno = IIf(IsNull(!Total), 0, !Total)
        .Close
        '******************************************************

        .Open "SELECT * FROM auxiliar WHERE Fecha>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND Fecha<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "' ORDER BY ID", dbDatos, adOpenForwardOnly, adLockReadOnly
        
        While Not .EOF
            
            If !Cuenta = "110950" And !Iniciales = "DO50" Then
                acDotaciones = acDotaciones + !Importe
                lblDotaciones.Tag = Val(lblDotaciones.Tag) + 1
            
            ElseIf !Cuenta = "110901" And !Iniciales = "CV01" Then
                acCortes = acCortes + !Importe
                lblCortes.Tag = Val(lblCortes.Tag) + 1
        
            ElseIf !Cuenta = "110901" And !Iniciales <> "CV01" And !Iniciales <> "RE01" Then
                acApor = acApor + !Importe
                lblAportaciones.Tag = Val(lblAportaciones.Tag) + 1
                
            ElseIf !Cuenta = "110950" And !Iniciales <> "DO50" Then
                acReti = acReti + !Importe
                lblRetiros.Tag = Val(lblRetiros.Tag) + 1
                
            ElseIf !Cuenta = "201701" And !Concepto = "Empeño" Then
                acEmp = acEmp + !Importe
                lblEmpeño.Tag = Val(lblEmpeño.Tag) + 1
            
            ElseIf !Cuenta = "201750" And (!Concepto = "Desempeño" Or !Concepto = "Pagos Fijos") Then
                acDes = acDes + !Importe
                lblDesempeño.Tag = Val(lblDesempeño.Tag) + 1
            
            ElseIf !Cuenta = "201750" And !Concepto = "Almoneda" Then
                acDepositaria = acDepositaria + !Importe
                
            ElseIf !Cuenta = "201701" And !Concepto = "Refrendo" Then
                acRef = acRef + !Importe
                lblRefrendo.Tag = Val(lblRefrendo.Tag) + 1
            
            ElseIf !Cuenta = "201750" And !Concepto = "Abono Refrendo" Then
                acAboRef = acAboRef + !Importe
                lblAboRefrendo.Tag = Val(lblAboRefrendo.Tag) + 1
            
            ElseIf (!Cuenta = "520450" Or !Cuenta = "670350" Or !Cuenta = "680350" Or !Cuenta = "690350") And (!Concepto = "Refrendo" Or !Concepto = "Desempeño" Or !Concepto = "Pagos Fijos") Then
                acInt = acInt + !Importe
                lblIntereses.Tag = Val(lblIntereses.Tag) + 1
            
            ElseIf !Cuenta = "620450" And !Iniciales = "VT03" And !Concepto = "Ventas" Then
                acVen = acVen + !Importe
                lblVentas.Tag = Val(lblVentas.Tag) + 1
                
            ElseIf !Cuenta = "110101" And (!Iniciales = "AP03" Or !Iniciales = "AB05") And (!Concepto = "Apartado" Or !Concepto = "Abonos") Then
                acAbo = acAbo + !Importe
                lblAboApartados.Tag = Val(lblAboApartados.Tag) + 1
                            
'''            ElseIf !Cuenta = "199450" And !Iniciales = "GA50" Then
'''                acGas = acGas + !Importe
'''                lblGastos.Tag = Val(lblGastos.Tag) + 1
            ElseIf !Cuenta = "110150" And !Iniciales = "GA50" Then
                acGas = acGas + !Importe
                lblGastos.Tag = Val(lblGastos.Tag) + 1
                                                               
            ElseIf !Cuenta = "620601" And !Iniciales = "DC06" And !Serie = 0 Then
                acDesc = acDesc + !Importe
                
            ElseIf !Cuenta = "620301" And !Iniciales = "EN01" And !Serie = 0 Then
                aInventario = aInventario + !Importe
                lblCompraVarios.Tag = Val(lblCompraVarios.Tag) + 1
                
            ElseIf !Cuenta = "620350" And !Iniciales = "EN50" And !Serie = 0 Then
                sInventario = sInventario + !Importe
                
            ElseIf !Cuenta = "710350" And !Iniciales = "VD50" And !Serie = 1 Then
                acDolares = acDolares + !Importe
                lblComDivisas.Tag = Val(lblComDivisas.Tag) + 1
                
            ElseIf !Cuenta = "710301" And !Iniciales = "CD01" And !Serie = 1 Then
                ccDolares = ccDolares + !Importe
                lblVenDivisas.Tag = Val(lblVenDivisas.Tag) + 1
                
            ElseIf !Cuenta = "120150" And (!Concepto = "Refrendo" Or !Concepto = "Desempeño" Or !Concepto = "Pagos Fijos") And !Concepto <> "Apartado" Then
                acIvaInt = acIvaInt + !Importe
                lblIvaIntereses.Tag = Val(lblIvaIntereses.Tag) + 1
            
            ElseIf !Cuenta = "120150" And !Concepto = "Ventas" Then
                acIvaVen = acIvaVen + !Importe
                lblIvaVentas.Tag = Val(lblIvaVentas.Tag) + 1
                
            ElseIf !Cuenta = "120101" Then
                ccIva = ccIva + !Importe
                lblIvaCompras.Tag = Val(lblIvaCompras.Tag) + 1
                
            ElseIf (!Cuenta = "530150" Or !Cuenta = "120150") And !Concepto = "Boleta perdida" Then
                ccOtros = ccOtros + !Importe
                lblOtrosCobros.Tag = Val(lblOtrosCobros.Tag) + IIf(!Cuenta = "120150", 0, 1)
            
            ElseIf !Cuenta = "151350" And !Iniciales = "PR50" And !Concepto = "Ajuste" Then
                acAjusteEntrada = acAjusteEntrada + !Importe
                lblAjusteEntrada.Tag = Val(lblAjusteEntrada.Tag) + 1
                
            ElseIf !Cuenta = "151301" And !Iniciales = "PR01" And !Concepto = "Ajuste" Then
                acAjusteSalida = acAjusteSalida + !Importe
                lblAjusteSalida.Tag = Val(lblAjusteSalida.Tag) + 1
                
            ElseIf !Cuenta = "650201" Then
                acDemasia = acDemasia + !Importe
                lblDemasias.Tag = Val(lblDemasias.Tag) + 1
                
            '***Puntos***
            ElseIf !Cuenta = "905501" Then

                acRedencion = acRedencion + !Importe
                lblRedencionPuntos.Tag = Val(lblRedencionPuntos.Tag) + 1
            End If

        .MoveNext
        Wend
        rcVentanilla.Close
        Set rcVentanilla = Nothing
        
        'ENTRADAS**********************************
        lblDotaciones.Caption = Format(acDotaciones, FMoneda)
        lblAportaciones.Caption = Format(acApor, FMoneda)
        lblBoveda.Caption = Format(ccrSaldoBov, FMoneda)
        lblDesempeño.Caption = Format(acDes, FMoneda)
        lblAboRefrendo.Caption = Format(acAboRef, FMoneda)
        lblIntereses.Caption = Format(acInt, FMoneda)
        lblIvaIntereses.Caption = Format(acIvaInt, FMoneda)
        lblVentas.Caption = Format(acVen - acDesc, FMoneda)
        lblIvaVentas.Caption = Format(acIvaVen, FMoneda)
        lblAboApartados.Caption = Format(acAbo, FMoneda)
        lblVenDivisas.Caption = Format(acDolares, FMoneda)
        lblOtrosCobros.Caption = Format(ccOtros, FMoneda)
        lblAjusteEntrada.Caption = Format(acAjusteEntrada, FMoneda)
        lblEntrada.Caption = Format(acApor + acDes + acAboRef + acInt + acIvaInt + (acVen - acDesc) + acIvaVen + acAbo + acDolares + ccOtros + acAjusteEntrada, FMoneda)
        
        lblRefrendo.Caption = Format(acRef, FMoneda)
        lblDescuento.Caption = Format(acDesc, FMoneda)
        '*********************************************
        
        'SALIDAS**************************************
        lblCortes.Caption = Format(acCortes, FMoneda)
        lblRetiros.Caption = Format(acReti, FMoneda)
        lblEmpeño.Caption = Format(acEmp, FMoneda)
        lblComDivisas.Caption = Format(ccDolares, FMoneda)
        lblCompraVarios.Caption = Format(aInventario, FMoneda)
        lblIvaCompras.Caption = Format(ccIva, FMoneda)
        lblGastos.Caption = Format(acGas, FMoneda)
        lblDemasias.Caption = Format(acDemasia, FMoneda)
        
        '***Puntos***
        lblRedencionPuntos.Caption = Format(acRedencion, FMoneda)
        
        lblAjusteSalida.Caption = Format(acAjusteSalida, FMoneda)
        lblSalida.Caption = Format(acReti + acEmp + ccDolares + aInventario + ccIva + acGas + acDemasia + acAjusteSalida + acRedencion, FMoneda)
        '*********************************************
        
        'Pongo el Total
        lblTotal.Caption = Format(((ccrSaldoBov) + (CDbl(lblEntrada.Caption))) - CDbl(lblSalida.Caption), FMoneda)
        
        'Pongo el Saldo Inicial de Divisas----------------------------------------
        .Open "SELECT SUM(a.Importe) AS Cargo FROM auxiliar a WHERE Cuenta='910901' AND Fecha<'" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            Cargo = IIf(IsNull(!Cargo), 0, !Cargo)
        .Close
    
        .Open "SELECT SUM(a.Importe) AS Abono FROM auxiliar a WHERE Cuenta='910950' AND Fecha<'" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            Abono = IIf(IsNull(!Abono), 0, !Abono)
        .Close
        
        'Entradas
        .Open "SELECT SUM(a.Importe) AS Cargo FROM auxiliar a WHERE Cuenta='910901' AND Fecha>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND Fecha<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            Entradas = IIf(IsNull(!Cargo), 0, !Cargo)
        .Close
        
        'Salidas
        .Open "SELECT SUM(a.Importe) AS Abono FROM auxiliar a WHERE Cuenta='910950' AND Fecha>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND Fecha<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
            Salidas = IIf(IsNull(!Abono), 0, !Abono)
        .Close
        Set rcVentanilla = Nothing
        
        lblSaldoInicial.Caption = Format(Cargo - Abono, "###,###,###,###0")
        lblSaldoInicial.Tag = Cargo - Abono
        
        lblEntradas.Caption = Format(Entradas, "###,###,###,###0")
        lblEntradas.Tag = Entradas
        lblSalidas.Caption = Format(Salidas, "###,###,###,###0")
        lblSalidas.Tag = Salidas
        
        lblSaldoFinal.Caption = Format((Cargo - Abono) + (Entradas - Salidas), "###,###,###,###0")
        lblSaldoFinal.Tag = (Cargo - Abono) + (Entradas - Salidas)
        '-------------------------------------------------------------------
        
        'Pongo el Saldo de Depositaría---------------------------------------
        lblSaldoInicialDep.Caption = Format(SaldoInicialEmpenos, FMoneda)
        lblEntradasDep = Format(EntradasEmpeno, FMoneda)
        lblSalidasDep = Format(SalidasEmpeno + acAboRef, FMoneda)
        lblSaldoFinalDep.Caption = Format(SaldoInicialEmpenos + EntradasEmpeno - SalidasEmpeno - acAboRef, FMoneda)
        
'''''        lblContratosAlmoneda.Caption = Val(SacaValor("empeno", "COUNT(ID)", " WHERE Almoneda=1 AND DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "'"))
'''''        lblContratosCancelados.Caption = Val(SacaValor("cancelaciones", "COUNT(ID)", " WHERE TipoMovimiento=1 AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "'"))
'''''        lblDivisasCanceladas = Val(SacaValor("cancelaciones", "COUNT(ID)", " WHERE TipoMovimiento=4 AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "'"))
    
    End With
    Set rcVentanilla = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcVentanilla = Nothing
End Sub

Private Sub Imprimir()
Dim rcGastos As New ADODB.Recordset
Dim Encabezado As String, TotalEmpenos As Long, TotalEmpenosPagados As Long

On Error GoTo Error
    
    
    Encabezado = "De la fecha " & Format(txtFechaIni.text, "dd/mmm/yyyy") & " a " & Format(txtFechaFin.text, "dd/mmm/yyyy") & ""
           
    'Imprimimos el reporte de corte de caja
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\CierreSucursal.rpt"
        
        'ENTRADAS ***************************************************
        .Formulas(0) = "SaldoInicial=" & ConvMoneda(lblBoveda.Caption) & ""
        .Formulas(1) = "AportacionesBov=" & ConvMoneda(lblAportaciones.Caption) & ""
        .Formulas(2) = "AportacionesBovNum='(" & Val(lblAportaciones.Tag) & ")'"
        .Formulas(3) = "Desempeño=" & ConvMoneda(lblDesempeño.Caption) & ""
        .Formulas(4) = "DesempeñoNum='(" & Val(lblDesempeño.Tag) & ")'"
        .Formulas(5) = "AboRefrendo=" & ConvMoneda(lblAboRefrendo.Caption) & ""
        .Formulas(6) = "AboRefrendoNum='(" & Val(lblAboRefrendo.Tag) & ")'"
        .Formulas(7) = "Intereses=" & ConvMoneda(lblIntereses.Caption) & ""
        .Formulas(8) = "IvaIntereses=" & ConvMoneda(lblIvaIntereses.Caption) & ""
        .Formulas(9) = "VentasMostrador=" & ConvMoneda(lblVentas.Caption) & ""
        .Formulas(10) = "VentasMostradorNum='(" & Val(lblVentas.Tag) & ")'"
        .Formulas(11) = "IvaVentas=" & ConvMoneda(lblIvaVentas.Caption) & ""
        .Formulas(12) = "AboApartados=" & ConvMoneda(lblAboApartados.Caption) & ""
        .Formulas(13) = "AboApartadosNum='(" & Val(lblAboApartados.Tag) & ")'"
        .Formulas(14) = "VentaDivisas=" & ConvMoneda(lblVenDivisas.Caption) & ""
        .Formulas(15) = "VentaDivisasNum='(" & Val(lblVenDivisas.Tag) & ")'"
        .Formulas(16) = "OtrosCobros=" & ConvMoneda(lblOtrosCobros.Caption) & ""
        .Formulas(17) = "OtrosCobrosNum='(" & Val(lblOtrosCobros.Tag) & ")'"
        .Formulas(18) = "AjustesEntrada=" & ConvMoneda(lblAjusteEntrada.Caption) & ""
        .Formulas(19) = "AjustesEntradaNum='(" & Val(lblAjusteEntrada.Tag) & ")'"
        .Formulas(20) = "TotEntrada=" & ConvMoneda(lblEntrada.Caption) & ""
        .Formulas(21) = "Refrendo=" & ConvMoneda(lblRefrendo.Caption) & ""
        .Formulas(22) = "RefrendoNum='(" & Val(lblRefrendo.Tag) & ")'"
        .Formulas(23) = "Descuento=" & ConvMoneda(lblDescuento.Caption) & ""
        '************************************************************
        
        'SALIDAS*****************************************************
        .Formulas(24) = "RetirosBov=" & ConvMoneda(lblRetiros.Caption) & ""
        .Formulas(25) = "RetirosBovNum='(" & Val(lblRetiros.Tag) & ")'"
        .Formulas(26) = "Empeño=" & ConvMoneda(lblEmpeño.Caption) & ""
        .Formulas(27) = "EmpeñoNum='(" & Val(lblEmpeño.Tag) & ")'"
        .Formulas(28) = "CompraDivisas=" & ConvMoneda(lblComDivisas.Caption) & ""
        .Formulas(29) = "CompraDivisasNum='(" & Val(lblComDivisas.Tag) & ")'"
        .Formulas(30) = "CompraVarios=" & ConvMoneda(lblCompraVarios.Caption) & ""
        .Formulas(31) = "CompraVariosNum='(" & Val(lblCompraVarios.Tag) & ")'"
        .Formulas(32) = "IvaCompras=" & ConvMoneda(lblIvaCompras.Caption) & ""
        .Formulas(33) = "Gastos=" & ConvMoneda(lblGastos.Caption) & ""
        .Formulas(34) = "GastosNum='(" & Val(lblGastos.Tag) & ")'"
        .Formulas(35) = "Demasias=" & ConvMoneda(lblDemasias.Caption) & ""
        .Formulas(36) = "DemasiasNum='(" & Val(lblDemasias.Tag) & ")'"
        
        '***Puntos***
        .Formulas(37) = "RedencionPuntos=" & ConvMoneda(lblRedencionPuntos.Caption) & ""
        .Formulas(38) = "RedencionPuntosNum='(" & Val(lblRedencionPuntos.Tag) & ")'"
        
        .Formulas(39) = "AjustesSalida=" & ConvMoneda(lblAjusteSalida.Caption) & ""
        .Formulas(40) = "AjustesSalidaNum='(" & Val(lblAjusteSalida.Tag) & ")'"
        
        .Formulas(41) = "TotSalida=" & ConvMoneda(lblSalida.Caption) & ""
        '**************************************************************
        
        .Formulas(42) = "SaldoFinal=" & ConvMoneda(lblTotal.Caption) & ""
        .Formulas(43) = "ContratosAlm=" & Val(lblContratosAlmoneda.Caption) & ""
        .Formulas(44) = "ContratosCanc=" & Val(lblContratosCancelados.Caption) & ""
        .Formulas(45) = "DivisasCanc=" & Val(lblDivisasCanceladas.Caption) & ""

        .Formulas(46) = "Caja='" & NombrePc & "'"
        .Formulas(47) = "Cajero='" & frmMDI.Usuario & "'"
        .Formulas(48) = "Encabezado='De la fecha " & Format(CDate(txtFechaIni.text), "DD/MMM/YYYY") & " a " & Format(CDate(txtFechaFin.text), "DD/MMM/YYYY") & "'"
        .Formulas(49) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(50) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        
        'DIVISAS*******************************************************
        .Formulas(51) = "SaldoInicialDivVen=" & Val(lblSaldoInicial.Tag) & ""
        .Formulas(52) = "ComprasDivVen=" & Val(lblEntradas.Tag) & ""
        .Formulas(53) = "VentasDivVen=" & Val(lblSalidas.Tag) & ""
        .Formulas(54) = "SaldoFinalDivVen=" & Val(lblSaldoFinal.Tag) & ""
        
        .Formulas(55) = "SaldoInicialDivDo=" & Val(lblSaldoInicial.Tag) & ""
        .Formulas(56) = "ComprasDivDo=" & Val(lblEntradas.Tag) & ""
        .Formulas(57) = "VentasDivDo=" & Val(lblSalidas.Tag) & ""
        .Formulas(58) = "SaldoFinalDivDo=" & Val(lblSaldoFinal.Tag) & ""
        '**************************************************************
        
        'DEPOSITARIA*******************************************************
        .Formulas(59) = "SaldoInicialDep=" & ConvMoneda(lblSaldoInicialDep.Caption) & ""
        .Formulas(60) = "EntradasDep=" & ConvMoneda(lblEntradasDep.Caption) & ""
        .Formulas(61) = "SalidasDep=" & ConvMoneda(lblSalidasDep.Caption) & ""
        .Formulas(62) = "SaldoFinalDep=" & ConvMoneda(lblSaldoFinalDep.Caption) & ""
        '**************************************************************
        
        .WindowTitle = "Cierre de sucursal"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
    
    'Imprimo la Factura por Intereses
    If MsgBox("Esta lista la impresora para mandar a imprimir la factura de Intereses ??", vbQuestion + vbYesNo + vbDefaultButton1, "Factura Refrendos") = vbYes Then
            
        Imprimir_Factura
    
    End If
                    
    'Imprimo la Factura de Ventas
    If MsgBox("Esta lista la impresora para mandar a imprimir la factura de Ventas ??", vbQuestion + vbYesNo + vbDefaultButton1, "Factura Desempeños") = vbYes Then
            
        Imprimir_Factura 1
    
    End If
    
    'Imprimo los Gastos
    rcGastos.Open "SELECT ID FROM gastos WHERE Fecha>='" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "' AND Fecha<='" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcGastos.BOF And Not rcGastos.EOF Then

        With frmMDI.Cr
            .Reset
            .DiscardSavedData = True
            .WindowShowPrintSetupBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\Gastos.rpt"
            .SelectionFormula = "{gastos.Fecha}>=date('" & Format(CDate(txtFechaIni.text), "YYYY/MM/DD") & "') AND {gastos.Fecha}<=date('" & Format(CDate(txtFechaFin.text), "YYYY/MM/DD") & "')"
            .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
            .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(2) = "Leyenda='De la fecha " & Format(CDate(txtFechaIni.text), "DD/MMM/YYYY") & " a " & Format(CDate(txtFechaFin.text), "DD/MMM/YYYY") & "'"
            .WindowTitle = "Reporte de Gastos"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With

    End If
    rcGastos.Close
    Set rcGastos = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcGastos = Nothing
End Sub

Sub LimpiaMontos()
        
    'INGRESOS
    acDotaciones = 0
    acApor = 0
    acDes = 0
    acDepositaria = 0
    acAboRef = 0
    acInt = 0
    acIvaInt = 0
    acVen = 0
'''''    acDesc = 0
    acIvaVen = 0
    acAbo = 0
    acDolares = 0
    ccOtros = 0
    acAjusteEntrada = 0
    acRef = 0
    acDesc = 0
    
    lblDotaciones.Caption = Format(acDotaciones, FMoneda)
    lblDotaciones.Tag = ""
    lblAportaciones.Caption = Format(acApor, FMoneda)
    lblAportaciones.Tag = ""
    lblBoveda.Caption = Format(0, FMoneda)
    lblBoveda.Tag = ""
    lblDesempeño.Caption = Format(acDes, FMoneda)
    lblDesempeño.Tag = ""
    lblAboRefrendo.Caption = Format(acAboRef, FMoneda)
    lblAboRefrendo.Tag = ""
    lblIntereses.Caption = Format(acInt, FMoneda)
    lblIntereses.Tag = ""
    lblIvaIntereses.Caption = Format(acIvaInt, FMoneda)
    lblIvaIntereses.Tag = ""
    lblVentas.Caption = Format(acVen, FMoneda)
    lblVentas.Tag = ""
    lblIvaVentas.Caption = Format(acIvaVen, FMoneda)
    lblIvaVentas.Tag = ""
    lblAboApartados.Caption = Format(acAbo, FMoneda)
    lblAboApartados.Tag = ""
    lblVenDivisas.Caption = Format(acDolares, FMoneda)
    lblVenDivisas.Tag = ""
    lblOtrosCobros.Caption = Format(ccOtros, FMoneda)
    lblOtrosCobros.Tag = ""
    lblAjusteEntrada.Caption = Format(acAjusteEntrada, FMoneda)
    lblAjusteEntrada.Tag = ""
    
    lblEntrada.Caption = Format(0, FMoneda)
    lblEntrada.Tag = ""
    lblRefrendo.Caption = Format(0, FMoneda)
    lblRefrendo.Tag = ""
    lblDescuento.Caption = Format(0, FMoneda)
    lblDescuento.Tag = ""
    
    'EGRESOS
    acCortes = 0
    acReti = 0
    acEmp = 0
    ccDolares = 0
    aInventario = 0
    ccIva = 0
    acGas = 0
    acDemasia = 0
    
    '***Puntos***
    acRedencion = 0
    
    acAjusteSalida = 0
    
    lblCortes.Caption = Format(acCortes, FMoneda)
    lblRetiros.Caption = Format(acReti, FMoneda)
    lblRetiros.Tag = ""
    lblEmpeño.Caption = Format(acEmp, FMoneda)
    lblEmpeño.Tag = ""
    lblComDivisas.Caption = Format(ccDolares, FMoneda)
    lblComDivisas.Tag = ""
    lblCompraVarios.Caption = Format(aInventario, FMoneda)
    lblCompraVarios.Tag = ""
    lblIvaCompras.Caption = Format(ccIva, FMoneda)
    lblIvaCompras.Tag = ""
    lblGastos.Caption = Format(acGas, FMoneda)
    lblGastos.Tag = ""
    lblDemasias.Caption = Format(acDemasia, FMoneda)
    lblDemasias.Tag = ""
    
    '***Puntos***
    lblRedencionPuntos.Caption = Format(acRedencion, FMoneda)
    lblRedencionPuntos.Tag = ""
    
    lblAjusteSalida.Caption = Format(acAjusteSalida, FMoneda)
    lblAjusteSalida.Tag = ""
    lblSalida.Caption = Format(0, FMoneda)
    lblSalida.Tag = ""
    lblTotal.Caption = Format(0, FMoneda)
    lblTotal.Tag = ""
    
    'Divisas
    lblSaldoInicial.Caption = "0"
    lblEntradas.Caption = "0"
    lblSalidas.Caption = "0"
    lblSaldoFinal.Caption = "0"
        
    'Depositaría
    lblSaldoInicialDep.Caption = Format(0, FMoneda)
    lblEntradasDep.Caption = Format(0, FMoneda)
    lblSalidasDep.Caption = Format(0, FMoneda)
    lblSaldoFinalDep.Caption = Format(0, FMoneda)
    
    lblContratosAlmoneda.Caption = "0"
    lblContratosCancelados.Caption = "0"
    lblDivisasCanceladas.Caption = "0"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtFechaFin_GotFocus()
    Seleccionar_Texto txtFechaFin
    Cambiar_Color True, txtFechaFin
End Sub

Private Sub txtFechaFin_LostFocus()
    Cambiar_Color False, txtFechaFin
End Sub

Private Sub txtFechaIni_GotFocus()
    Seleccionar_Texto txtFechaIni
    Cambiar_Color True, txtFechaIni
End Sub

Private Sub txtFechaIni_LostFocus()
    Cambiar_Color False, txtFechaIni
End Sub

Public Sub Imprimir_Factura(Optional TipoFactura As Integer = 0)
Dim rcFactura As New ADODB.Recordset
Dim crImporteIva As Double, Iva As Double, Copia As Integer
Dim Impresora As Printer

On Error GoTo Error
    
    rcFactura.Open "SELECT NumRegistros,ImporteTotal FROM " & IIf(TipoFactura = 0, "vwfacturadiaria", "vwfacturaventas") & " WHERE Fecha>='" & Format(txtFechaIni.text, "YYYY/MM/DD") & "' AND Fecha<='" & Format(txtFechaFin.text, "YYYY/MM/DD") & "' ORDER BY Fecha", dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not rcFactura.BOF And Not rcFactura.EOF Then
        
        Iva = Regresa_Valor_BD(IIf(TipoFactura = 0, "IVA", "IVAVentas")) / 100
        Set Impresora = Printer
        With Impresora
        
            For Copia = 1 To 2
                
                If MsgBox("Esta lista la factura para imprimir la copia No. " & Copia & " ??", vbQuestion + vbYesNo + vbDefaultButton1) = vbYes Then
                    
                    .ScaleMode = vbMillimeters
                    .Font = "Verdana"
                    .FontBold = False
                    .FontSize = 10
                    
                    'Imprimo la fecha
                    .CurrentX = Regresa_Valor("FACTURA", "FechaX", 0)
                    .CurrentY = Regresa_Valor("FACTURA", "FechaY", 0)
                    Impresora.Print Format(Date, "DD/MMM/YYYY")
                    
                    'Imprimo el nombre del cliente
                    .CurrentX = Regresa_Valor("FACTURA", "ClienteX", 0)
                    .CurrentY = Regresa_Valor("FACTURA", "ClienteY", 0)
                    Impresora.Print "PUBLICO EN GENERAL"
                    
                    .FontSize = 9
                    'Imprimo la Descripcion
                    .CurrentX = Regresa_Valor("FACTURA", "DescripcionX", 0)
                    .CurrentY = Regresa_Valor("FACTURA", "PartidaY", 0)
                    Impresora.Print IIf(TipoFactura = 0, "INTERESES ", "VENTAS ") & IIf(CDate(txtFechaIni.text) = CDate(txtFechaFin.text), "DEL DÍA " & Format(CDate(txtFechaIni.text), "DD/MMM/YYYY"), "DEL " & Format(CDate(txtFechaIni.text), "DD/MMM/YYYY") & " AL " & Format(CDate(txtFechaFin.text), "DD/MMM/YYYY"))
                                
                    'Imprimo el Precio
                    .CurrentX = Regresa_Valor("FACTURA", "ImporteX", 0)
                    .CurrentY = Regresa_Valor("FACTURA", "PartidaY", 0)
                    Impresora.Print RegresaEspacios(rcFactura!ImporteTotal, 25)
                                                            
                    .FontSize = 10
                                
                    'Imprimo el Subtotal
                    .CurrentX = Regresa_Valor("FACTURA", "SubTotalX", 0)
                    .CurrentY = Regresa_Valor("FACTURA", "SubTotalY", 0)
                    Impresora.Print RegresaEspacios(rcFactura!ImporteTotal, 25)
                    
                    'Saco el Iva
                    crImporteIva = Redondeo(rcFactura!ImporteTotal * Iva)
                    
                    'Imprimo el Iva Total
                    .CurrentX = Regresa_Valor("FACTURA", "IvaX", 0)
                    .CurrentY = Regresa_Valor("FACTURA", "IvaY", 0)
                    Impresora.Print RegresaEspacios(crImporteIva, 25)
                                            
                    'Imprimo el Total Gral.
                    .CurrentX = Regresa_Valor("FACTURA", "TotalX", 0)
                    .CurrentY = Regresa_Valor("FACTURA", "TotalY", 0)
                    Impresora.Print RegresaEspacios(rcFactura!ImporteTotal + crImporteIva, 25)
                    
                    'Imprimo la Cantidad en Letras
                    .CurrentX = Regresa_Valor("FACTURA", "CantidadLetraX", 0)
                    .CurrentY = Regresa_Valor("FACTURA", "CantidadLetraY", 0)
                    Impresora.Print CantidadEnLetra(CCur(rcFactura!ImporteTotal + crImporteIva))
                    
                .EndDoc
                End If
                
            Next Copia
            
        End With
    
    End If
    
    rcFactura.Close
    Set rcFactura = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub
