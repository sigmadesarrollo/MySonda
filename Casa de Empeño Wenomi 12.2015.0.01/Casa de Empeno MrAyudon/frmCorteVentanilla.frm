VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmCorteVentanilla 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cierre de Caja"
   ClientHeight    =   10965
   ClientLeft      =   45
   ClientTop       =   720
   ClientWidth     =   11400
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCorteVentanilla.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10965
   ScaleWidth      =   11400
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   10200
      Left            =   0
      TabIndex        =   5
      Top             =   480
      Width           =   11385
      Begin Line3D.ucLine3D ucLine3D5 
         Height          =   30
         Left            =   120
         Top             =   9000
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   3
      End
      Begin Line3D.ucLine3D ucLine3D4 
         Height          =   1335
         Left            =   5520
         Top             =   7680
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2355
         Orientation     =   0
         LineWidth       =   3
      End
      Begin Line3D.ucLine3D ucLine3D3 
         Height          =   30
         Left            =   240
         Top             =   8520
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   53
         LineWidth       =   3
      End
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   30
         Left            =   120
         Top             =   8160
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   3
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   375
         Index           =   59
         Left            =   2985
         Top             =   240
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   661
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   7530
         Index           =   0
         Left            =   120
         Top             =   240
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   13282
         Orientation     =   0
         LineWidth       =   2
      End
      Begin VB.TextBox txtEfectivo 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   8655
         MaxLength       =   14
         TabIndex        =   6
         Top             =   7035
         Width           =   2220
      End
      Begin DevPowerFlatBttn.FlatBttn cmdModificaCorte 
         Height          =   300
         Left            =   10890
         TabIndex        =   49
         Top             =   7035
         Width           =   225
         _ExtentX        =   397
         _ExtentY        =   529
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
         Height          =   30
         Index           =   1
         Left            =   120
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   2
         Left            =   150
         Top             =   615
         Width           =   5370
         _ExtentX        =   9472
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   3
         Left            =   135
         Top             =   990
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   4
         Left            =   135
         Top             =   1365
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   5
         Left            =   135
         Top             =   1740
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   6
         Left            =   135
         Top             =   2115
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   7
         Left            =   135
         Top             =   2490
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   8
         Left            =   135
         Top             =   3240
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   9
         Left            =   135
         Top             =   3615
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   10
         Left            =   135
         Top             =   3990
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   11
         Left            =   135
         Top             =   4740
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   12
         Left            =   120
         Top             =   7365
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   7530
         Index           =   13
         Left            =   5520
         Top             =   240
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   13282
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   405
         Index           =   14
         Left            =   120
         Top             =   9120
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   714
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   15
         Left            =   120
         Top             =   9120
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   390
         Index           =   16
         Left            =   5520
         Top             =   9120
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   688
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   17
         Left            =   120
         Top             =   9480
         Width           =   5430
         _ExtentX        =   9578
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   6765
         Index           =   18
         Left            =   2985
         Top             =   1005
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   11933
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   20
         Left            =   120
         Top             =   9600
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   21
         Left            =   120
         Top             =   9600
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   22
         Left            =   120
         Top             =   7740
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   23
         Left            =   120
         Top             =   9960
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   24
         Left            =   5280
         Top             =   9600
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   25
         Left            =   5760
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   5640
         Index           =   26
         Left            =   5760
         Top             =   240
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   9948
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   27
         Left            =   5760
         Top             =   615
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   28
         Left            =   5760
         Top             =   990
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   29
         Left            =   5760
         Top             =   1365
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   30
         Left            =   5760
         Top             =   1740
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   31
         Left            =   5760
         Top             =   2115
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   32
         Left            =   5760
         Top             =   2490
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   33
         Left            =   5760
         Top             =   2865
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   4875
         Index           =   34
         Left            =   8595
         Top             =   1005
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   8599
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   35
         Left            =   5760
         Top             =   3240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   405
         Index           =   36
         Left            =   5760
         Top             =   6015
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   714
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   37
         Left            =   5760
         Top             =   6015
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   38
         Left            =   11145
         Top             =   6015
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   39
         Left            =   5760
         Top             =   6405
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   40
         Left            =   8595
         Top             =   6015
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   5625
         Index           =   41
         Left            =   11145
         Top             =   240
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   9922
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   405
         Index           =   42
         Left            =   5760
         Top             =   6555
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   714
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   43
         Left            =   5760
         Top             =   6555
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   44
         Left            =   11145
         Top             =   6555
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   45
         Left            =   5760
         Top             =   6945
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   405
         Index           =   47
         Left            =   5760
         Top             =   6975
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   714
         Orientation     =   0
         ShadowColor     =   16777215
         LigthColor      =   16777215
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   48
         Left            =   5760
         Top             =   6975
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         ShadowColor     =   16777215
         LigthColor      =   16777215
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   49
         Left            =   11145
         Top             =   6960
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         ShadowColor     =   16777215
         LigthColor      =   16777215
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   50
         Left            =   5760
         Top             =   7365
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         ShadowColor     =   16777215
         LigthColor      =   16777215
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   51
         Left            =   8595
         Top             =   6960
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         ShadowColor     =   16777215
         LigthColor      =   16777215
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   405
         Index           =   52
         Left            =   5760
         Top             =   7425
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   714
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   53
         Left            =   5760
         Top             =   7425
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   54
         Left            =   11145
         Top             =   7425
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   55
         Left            =   5760
         Top             =   7815
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   56
         Left            =   8595
         Top             =   7425
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   58
         Left            =   135
         Top             =   2865
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   60
         Left            =   135
         Top             =   4365
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   61
         Left            =   135
         Top             =   5115
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   62
         Left            =   135
         Top             =   5490
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   63
         Left            =   135
         Top             =   5865
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   64
         Left            =   135
         Top             =   6240
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   19
         Left            =   120
         Top             =   6615
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   65
         Left            =   120
         Top             =   6990
         Width           =   5400
         _ExtentX        =   9525
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   46
         Left            =   5760
         Top             =   3615
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   57
         Left            =   5760
         Top             =   3990
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   66
         Left            =   5760
         Top             =   4365
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   67
         Left            =   5760
         Top             =   4740
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   68
         Left            =   5760
         Top             =   5115
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   69
         Left            =   5760
         Top             =   5865
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   375
         Index           =   70
         Left            =   8595
         Top             =   240
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   661
         Orientation     =   0
         LineWidth       =   2
      End
      Begin DevPowerFlatBttn.FlatBttn cmdSalir 
         Cancel          =   -1  'True
         Height          =   375
         Left            =   9840
         TabIndex        =   93
         Top             =   8880
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
         Picture         =   "frmCorteVentanilla.frx":000C
      End
      Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
         Height          =   375
         Left            =   8640
         TabIndex        =   94
         Top             =   8880
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
         Picture         =   "frmCorteVentanilla.frx":055E
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   390
         Index           =   71
         Left            =   3000
         Top             =   9120
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   688
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   420
         Index           =   72
         Left            =   8595
         Top             =   6555
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   741
         Orientation     =   0
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   30
         Index           =   73
         Left            =   5760
         Top             =   5490
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   53
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   390
         Index           =   74
         Left            =   2985
         Top             =   9600
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   688
         Orientation     =   0
         LineWidth       =   2
      End
      Begin VB.Label lblIVARenta 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "<IVA Renta GYS>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   105
         Top             =   8640
         Width           =   2505
      End
      Begin VB.Label lblRentaPoliza 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "<Renta Poliza>"
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
         Left            =   3000
         TabIndex        =   104
         Top             =   8160
         Width           =   2490
      End
      Begin VB.Label lblRentaGPS 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<Renta GPS>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   103
         Top             =   7800
         Width           =   2535
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Caption         =   "I.V.A Renta GPS y Seguro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   360
         Left            =   120
         TabIndex        =   102
         Top             =   8640
         Width           =   2835
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Renta Póliza Seguro"
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
         Height          =   375
         Left            =   120
         TabIndex        =   101
         Top             =   8180
         Width           =   2835
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Renta GPS"
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
         Height          =   345
         Left            =   120
         TabIndex        =   100
         Top             =   7800
         Width           =   2835
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "* Inform. Refrendos:"
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
         Left            =   240
         TabIndex        =   17
         Top             =   9600
         Width           =   2535
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   98
         Top             =   9600
         Width           =   2835
      End
      Begin VB.Label Label27 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
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
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   97
         Top             =   9120
         Width           =   2835
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
         Left            =   8595
         TabIndex        =   96
         Top             =   5540
         Width           =   2340
      End
      Begin VB.Label lblRedencionPuntos1 
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
         Left            =   5835
         TabIndex        =   95
         Top             =   5540
         Width           =   2340
      End
      Begin VB.Label Label33 
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "Devoluciones:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   11
         Left            =   5775
         TabIndex        =   92
         Top             =   4410
         Width           =   2850
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "COMPRAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   10
         Left            =   5775
         TabIndex        =   91
         Top             =   3285
         Width           =   2850
      End
      Begin VB.Label lblEmpeñoFijo 
         Alignment       =   1  'Right Justify
         Caption         =   "<Empeñ Fijos>"
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
         Left            =   8595
         TabIndex        =   90
         Top             =   2175
         Width           =   2340
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "DIVISAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   9
         Left            =   5775
         TabIndex        =   89
         Top             =   2535
         Width           =   2850
      End
      Begin VB.Label Label7 
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
         Left            =   5835
         TabIndex        =   88
         Top             =   2145
         Width           =   2310
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "PAGOS FIJOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   8
         Left            =   5775
         TabIndex        =   87
         Top             =   1785
         Width           =   2850
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "TRADICIONALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   7
         Left            =   5775
         TabIndex        =   86
         Top             =   1035
         Width           =   2850
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "OTRAS ENTRADAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   6
         Left            =   120
         TabIndex        =   85
         Top             =   7050
         Width           =   2850
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "DIVISAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   5
         Left            =   135
         TabIndex        =   84
         Top             =   6285
         Width           =   2850
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "VENTAS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   4
         Left            =   135
         TabIndex        =   83
         Top             =   4770
         Width           =   2850
      End
      Begin VB.Label lblIvaInteresesFijo 
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
         Left            =   3000
         TabIndex        =   82
         Top             =   4425
         Width           =   2340
      End
      Begin VB.Label lblMoratorios 
         Alignment       =   1  'Right Justify
         Caption         =   "<Moratorios>"
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
         Left            =   3000
         TabIndex        =   81
         Top             =   4050
         Width           =   2340
      End
      Begin VB.Label lblInteresesFijo 
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
         Left            =   3000
         TabIndex        =   80
         Top             =   3675
         Width           =   2340
      End
      Begin VB.Label lblDesempeñoFijo 
         Alignment       =   1  'Right Justify
         Caption         =   "<Capital Fijo>"
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
         Left            =   3000
         TabIndex        =   79
         Top             =   3300
         Width           =   2340
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
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
         Left            =   195
         TabIndex        =   78
         Top             =   4410
         Width           =   765
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moratorios:"
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
         Left            =   195
         TabIndex        =   77
         Top             =   4035
         Width           =   1410
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
         Index           =   3
         Left            =   195
         TabIndex        =   76
         Top             =   3660
         Width           =   1245
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
         Index           =   2
         Left            =   195
         TabIndex        =   75
         Top             =   3300
         Width           =   2415
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "PAGOS FIJOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   3
         Left            =   135
         TabIndex        =   74
         Top             =   2895
         Width           =   2850
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
         BackStyle       =   0  'Transparent
         Caption         =   "TRADICIONALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Index           =   2
         Left            =   135
         TabIndex        =   73
         Top             =   1035
         Width           =   2850
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AJUSTE:"
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
         Left            =   5880
         TabIndex        =   71
         Top             =   7455
         Width           =   1185
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "EFECTIVO CAJA:"
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
         Left            =   5880
         TabIndex        =   69
         Top             =   7035
         Width           =   2325
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "SALDO CAJA:"
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
         Left            =   5880
         TabIndex        =   68
         Top             =   6615
         Width           =   1860
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   3
         Left            =   5760
         TabIndex        =   67
         Top             =   6585
         Width           =   2835
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
         Index           =   2
         Left            =   5955
         TabIndex        =   66
         Top             =   6075
         Width           =   2385
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   2
         Left            =   5760
         TabIndex        =   65
         Top             =   6045
         Width           =   2835
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         Left            =   120
         TabIndex        =   64
         Top             =   630
         Width           =   5415
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         BackColor       =   &H00404040&
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
         Left            =   5760
         TabIndex        =   63
         Top             =   630
         Width           =   5415
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
         Left            =   3000
         TabIndex        =   61
         Top             =   5550
         Width           =   2340
      End
      Begin VB.Label lblIvaInteresesTrad 
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
         Left            =   3000
         TabIndex        =   60
         Top             =   2550
         Width           =   2340
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
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
         Left            =   195
         TabIndex        =   59
         Top             =   5550
         Width           =   765
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
         Left            =   195
         TabIndex        =   58
         Top             =   1785
         Width           =   1995
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
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
         Left            =   195
         TabIndex        =   57
         Top             =   2520
         Width           =   765
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
         Left            =   8595
         TabIndex        =   55
         Top             =   5160
         Width           =   2340
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
         Left            =   5835
         TabIndex        =   54
         Top             =   5160
         Width           =   1260
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
         Left            =   3000
         TabIndex        =   53
         Top             =   7425
         Width           =   2340
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
         Index           =   0
         Left            =   195
         TabIndex        =   52
         Top             =   7425
         Width           =   1620
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "I.V.A.:"
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
         Left            =   5835
         TabIndex        =   51
         Top             =   4035
         Width           =   765
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
         Left            =   8595
         TabIndex        =   50
         Top             =   4050
         Width           =   2340
      End
      Begin VB.Label Label28 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Compra:"
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
         Left            =   5835
         TabIndex        =   48
         Top             =   2925
         Width           =   1035
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
         Left            =   8595
         TabIndex        =   47
         Top             =   2925
         Width           =   2340
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
         Left            =   3000
         TabIndex        =   46
         Top             =   6675
         Width           =   2340
      End
      Begin VB.Label Label26 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Venta:"
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
         Left            =   195
         TabIndex        =   45
         Top             =   6675
         Width           =   810
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
         Left            =   8595
         TabIndex        =   44
         Top             =   3675
         Width           =   2340
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
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
         Left            =   5835
         TabIndex        =   43
         Top             =   3660
         Width           =   1965
      End
      Begin VB.Label Label25 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Prestamos:"
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
         Left            =   6600
         TabIndex        =   41
         Top             =   10320
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblPrestamosDia 
         Alignment       =   1  'Right Justify
         Caption         =   "<Prestamos>"
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
         Left            =   8160
         TabIndex        =   40
         Top             =   10200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Venta de Joyería:"
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
         Left            =   6240
         TabIndex        =   39
         Top             =   10680
         Visible         =   0   'False
         Width           =   2145
      End
      Begin VB.Label lblVentaJoyeria 
         Alignment       =   1  'Right Justify
         Caption         =   "<Venta Joyería>"
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
         Left            =   8280
         TabIndex        =   38
         Top             =   10680
         Visible         =   0   'False
         Width           =   1935
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
         Left            =   240
         TabIndex        =   37
         Top             =   9720
         Width           =   2250
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
         Left            =   8880
         TabIndex        =   36
         Top             =   10200
         Width           =   2055
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
         Left            =   8595
         TabIndex        =   35
         Top             =   6075
         Width           =   2340
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
         Left            =   3120
         TabIndex        =   34
         Top             =   9120
         Width           =   2340
      End
      Begin VB.Label lblDevoluciones 
         Alignment       =   1  'Right Justify
         Caption         =   "<Devoluciones>"
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
         Left            =   8640
         TabIndex        =   33
         Top             =   4440
         Width           =   2295
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Devoluciones:"
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
         Left            =   4320
         TabIndex        =   32
         Top             =   10200
         Width           =   1695
      End
      Begin VB.Label lblPrestamos 
         Alignment       =   1  'Right Justify
         Caption         =   "<Prestamos>"
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
         Left            =   7920
         TabIndex        =   31
         Top             =   10200
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pago de ajustes:"
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
         Left            =   6000
         TabIndex        =   30
         Top             =   10200
         Visible         =   0   'False
         Width           =   2040
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
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
         Left            =   5835
         TabIndex        =   29
         Top             =   4785
         Width           =   1725
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
         Left            =   8595
         TabIndex        =   28
         Top             =   4785
         Width           =   2340
      End
      Begin VB.Label lblCambios 
         Alignment       =   1  'Right Justify
         Caption         =   "<Cambios>"
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
         Left            =   8400
         TabIndex        =   27
         Top             =   9480
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cambios Ventas:"
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
         Left            =   6240
         TabIndex        =   26
         Top             =   9480
         Visible         =   0   'False
         Width           =   2040
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
         Left            =   195
         TabIndex        =   25
         Top             =   2145
         Width           =   1245
      End
      Begin VB.Label lblInteresesTrad 
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
         Left            =   3000
         TabIndex        =   24
         Top             =   2175
         Width           =   2340
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
         Left            =   195
         TabIndex        =   23
         Top             =   5940
         Width           =   2520
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
         Left            =   3000
         TabIndex        =   22
         Top             =   5925
         Width           =   2340
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
         Left            =   195
         TabIndex        =   21
         Top             =   5160
         Width           =   2265
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
         Left            =   3000
         TabIndex        =   20
         Top             =   5175
         Width           =   2340
      End
      Begin VB.Label lblAboRefrendoTrad 
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
         Left            =   3000
         TabIndex        =   19
         Top             =   1800
         Width           =   2340
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
         Left            =   3240
         TabIndex        =   18
         Top             =   9600
         Width           =   2055
      End
      Begin VB.Label lblRetiro 
         Alignment       =   1  'Right Justify
         Caption         =   "<Retiro>"
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
         Left            =   8595
         TabIndex        =   16
         Top             =   300
         Width           =   2340
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Retiros a caja:"
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
         Left            =   5835
         TabIndex        =   15
         Top             =   270
         Width           =   1785
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dotaciones a caja:"
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
         Left            =   195
         TabIndex        =   14
         Top             =   270
         Width           =   2250
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
         Left            =   195
         TabIndex        =   13
         Top             =   1410
         Width           =   2415
      End
      Begin VB.Label Label6 
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
         Left            =   5835
         TabIndex        =   12
         Top             =   1410
         Width           =   2310
      End
      Begin VB.Label lblBoveda 
         Alignment       =   1  'Right Justify
         Caption         =   "<Bóveda>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3000
         TabIndex        =   11
         Top             =   300
         Width           =   2340
      End
      Begin VB.Label lblDesempeñoTrad 
         Alignment       =   1  'Right Justify
         Caption         =   "<Capital Tradicional>"
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
         Left            =   3000
         TabIndex        =   10
         Top             =   1425
         Width           =   2340
      End
      Begin VB.Label lblEmpeñoTrad 
         Alignment       =   1  'Right Justify
         Caption         =   "<Empeñ Tradicional>"
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
         Left            =   8595
         TabIndex        =   9
         Top             =   1425
         Width           =   2340
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "<Total>"
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
         Left            =   8595
         TabIndex        =   8
         Top             =   6615
         Width           =   2340
      End
      Begin VB.Label lblAjuste 
         Alignment       =   1  'Right Justify
         Caption         =   "<Ajuste>"
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
         Left            =   8595
         TabIndex        =   7
         Top             =   7500
         Width           =   2340
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Height          =   7485
         Index           =   0
         Left            =   120
         TabIndex        =   56
         Top             =   270
         Width           =   2835
      End
      Begin VB.Label Label32 
         BackColor       =   &H00000000&
         Height          =   5600
         Index           =   1
         Left            =   5760
         TabIndex        =   62
         Top             =   270
         Width           =   2835
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   4
         Left            =   5760
         TabIndex        =   70
         Top             =   7005
         Width           =   2835
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Height          =   375
         Index           =   5
         Left            =   5760
         TabIndex        =   72
         Top             =   7455
         Width           =   2835
      End
   End
   Begin VB.Label Label32 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   7485
      Index           =   2
      Left            =   0
      TabIndex        =   99
      Top             =   0
      Width           =   2835
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
      Left            =   4320
      TabIndex        =   42
      Top             =   120
      Width           =   1050
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   645
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
      Left            =   3360
      TabIndex        =   3
      Top             =   120
      Width           =   900
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
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   795
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
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   825
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
      Left            =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   960
   End
End
Attribute VB_Name = "frmCorteVentanilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.I. Jorge Gabriel Colio Ramos
' Mazatlan, Sin. 15/08/2002
' Modulo frmCorteCajaVnetanilla - frmCorteCajaVnetanilla.frm
' Ultima Modificacion - 16/08/2002
'Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit

Dim acDesc, acDesFij As Double, acCam As Double, acAbo As Double, acVen As Double, acEmp As Double, acEmpFij As Double, acDev As Double, acDes As Double, acRef As Double, acAboRef As Double, acInt As Double, acIntFij As Double, acMoratorios As Double, acBov As Double, acTot As Double, acRet As Double, acGas As Double, acPres As Double, acPresO As Double, acPresP As Double
Dim aInventario As Double, sInventario As Double, Efectivo As Double, acDolares As Double, ccDolares As Double, acIvaInt As Double, acIvaIntFij As Double, acIvaVen As Double, ccIva As Double, ccOtros As Double, acDemasia As Double
Dim acRentaGPS As Double, acRentaPoliza As Double, acIVARenta As Double
'***Puntos***
Dim acRedencion As Double
Dim Fl() As cFlatControl

Private Sub Inicializar()
    
    Screen.MousePointer = vbHourglass
    lblFecha.Caption = Format(Date, "DD/MMM/YYYY")
    lblCaja.Caption = NombrePc
    Cargar_Montos

    If lblTotal.Caption = "0" Then
        
        cmdAceptar.Enabled = False
    
    Else
        
        cmdAceptar.Enabled = True
    
    End If

    lblCajero.Caption = frmMDI.Usuario
    lblCajero.Tag = frmMDI.IDUsuario
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub Cargar_Montos()
Dim rcVentanilla As New ADODB.Recordset

On Error GoTo Error

    '''''rcVentanilla.Open "SELECT * FROM auxiliar WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "' AND Corte=0 ORDER BY ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    rcVentanilla.Open "SELECT * FROM auxiliar WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "' ORDER BY ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    LimpiaMontos

    With rcVentanilla
    
        While Not .EOF

'''            If !Cuenta = "199401" And !Iniciales = "DO01" Then
'''                acBov = acBov + !Importe
'''                lblBoveda.Tag = Val(lblBoveda.Tag) + 1
'''
'''            ElseIf !Cuenta = "199450" And !Iniciales = "RE50" Then
'''                acRet = acRet + !Importe
'''                lblRetiro.Tag = Val(lblRetiro.Tag) + 1
'''
'''            ElseIf !Cuenta = "199450" And !Concepto = "Empeño" And (!Serie = 1 Or !Serie = 2) Then
'''                acEmp = acEmp + !Importe
'''                lblEmpeñoTrad.Tag = Val(lblEmpeñoTrad.Tag) + 1
'''
'''            ElseIf !Cuenta = "199450" And !Concepto = "Empeño" And !Serie = 3 Then
'''                acEmpFij = acEmpFij + !Importe
'''                lblEmpeñoFijo.Tag = Val(lblEmpeñoFijo.Tag) + 1

            If !Cuenta = "110101" And !Iniciales = "DO01" Then
                acBov = acBov + !Importe
                lblBoveda.Tag = Val(lblBoveda.Tag) + 1
            
            ElseIf !Cuenta = "110150" And !Iniciales = "RE50" Then
                acRet = acRet + !Importe
                lblRetiro.Tag = Val(lblRetiro.Tag) + 1
            
            ElseIf !Cuenta = "110150" And !Concepto = "Empeño" And !Serie <> 3 Then
                acEmp = acEmp + !Importe
                lblEmpeñoTrad.Tag = Val(lblEmpeñoTrad.Tag) + 1
            
            ElseIf !Cuenta = "110150" And !Concepto = "Empeño" And !Serie = 3 Then
                acEmpFij = acEmpFij + !Importe
                lblEmpeñoFijo.Tag = Val(lblEmpeñoFijo.Tag) + 1
                
            ElseIf !Cuenta = "201750" And !Concepto = "Desempeño" And !Serie <> 3 Then
                acDes = acDes + !Importe
                lblDesempeñoTrad.Tag = Val(lblDesempeñoTrad.Tag) + 1
            
            ElseIf !Cuenta = "201750" And !Concepto = "Pagos Fijos" Then
                acDesFij = acDesFij + !Importe
                lblDesempeñoFijo.Tag = Val(lblDesempeñoFijo.Tag) + 1
            
            ElseIf !Cuenta = "201701" And !Concepto = "Refrendo" Then
                acRef = acRef + !Importe
                lblRefrendo.Tag = Val(lblRefrendo.Tag) + 1
            
            ElseIf !Cuenta = "201750" And !Concepto = "Abono Refrendo" Then
                acAboRef = acAboRef + !Importe
                lblAboRefrendoTrad.Tag = Val(lblAboRefrendoTrad.Tag) + 1
            
            ElseIf (!Cuenta = "520450" Or !Cuenta = "670350" Or !Cuenta = "680350" Or !Cuenta = "690350") And (!Concepto = "Refrendo" Or !Concepto = "Desempeño") Then
                acInt = acInt + !Importe
                lblInteresesTrad.Tag = Val(lblInteresesTrad.Tag) + 1
            
            ElseIf (!Cuenta = "520450" Or !Cuenta = "670350" Or !Cuenta = "680350") And !Concepto = "Pagos Fijos" Then
                acIntFij = acIntFij + !Importe
                lblInteresesFijo.Tag = Val(lblInteresesFijo.Tag) + 1
            
            ElseIf !Cuenta = "690350" And !Concepto = "Pagos Fijos" Then
                acMoratorios = acMoratorios + !Importe
                lblMoratorios.Tag = Val(lblMoratorios.Tag) + 1
                
            ElseIf !Cuenta = "620450" And !Iniciales = "VT03" And !Concepto = "Ventas" Then
                acVen = acVen + !Importe
                lblVentas.Tag = Val(lblVentas.Tag) + 1
                
            ElseIf !Cuenta = "110101" And (!Iniciales = "AP03" Or !Iniciales = "AB05") And (!Concepto = "Apartado" Or !Concepto = "Abonos") Then
                acAbo = acAbo + !Importe
                lblAboApartados.Tag = Val(lblAboApartados.Tag) + 1
                
'''            ElseIf !Cuenta = "199401" And !Iniciales = "CM01" Then
'''                acCam = acCam + !Importe
'''                lblCambios.Tag = Val(lblCambios.Tag) + 1
'''
'''            ElseIf !Cuenta = "199450" And !Iniciales = "GA50" Then
'''                acGas = acGas + !Importe
'''                lblGastos.Tag = Val(lblGastos.Tag) + 1
            ElseIf !Cuenta = "110101" And !Iniciales = "CM01" Then
                acCam = acCam + !Importe
                lblCambios.Tag = Val(lblCambios.Tag) + 1
            
            ElseIf !Cuenta = "110150" And !Iniciales = "GA50" Then
                acGas = acGas + !Importe
                lblGastos.Tag = Val(lblGastos.Tag) + 1
                
'''''            ElseIf !Cuenta = "199401" And !Iniciales = "AJ01" Then
'''''                acPres = acPres + !Importe
'''''                lblPrestamos.Tag = Val(lblPrestamos.Tag) + 1
                
            ElseIf (!Cuenta = "110150" And !Iniciales = "DEV50") Or (!Cuenta = "620501" And !Iniciales = "CA01" And !Concepto = "Abonos Cancelados") Then
                acDev = acDev + !Importe
                lblDevoluciones.Tag = Val(lblDevoluciones.Tag) + 1
                
'''''            ElseIf !Cuenta = "151301" And !Iniciales = "PR01" Then
'''''                acPresO = acPresO + !Importe
'''''
'''''            ElseIf !Cuenta = "151350" And !Iniciales = "PR50" Then
'''''                acPresP = acPresP + !Importe
'''''                lblPrestamosDia.Tag = Val(lblPrestamosDia.Tag) + 1
                
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
                
            ElseIf !Cuenta = "120150" And (!Concepto = "Refrendo" Or !Concepto = "Desempeño") And !Concepto <> "Apartado" Then
                acIvaInt = acIvaInt + !Importe
                lblIvaInteresesTrad.Tag = Val(lblIvaInteresesTrad.Tag) + 1
            ElseIf !Cuenta = "120150" And (!Concepto = "Refrendo Renta GPS" Or !Concepto = "Desempeño Renta GPS") And !Concepto <> "Apartado" Then  'se agrega para IVA de GPS
                acIVARenta = acIVARenta + !Importe
                lblIVARenta.Tag = Val(lblIVARenta.Tag) + 1
                
            ElseIf !Cuenta = "818150" And (!Concepto = "Refrendo Renta GPS" Or !Concepto = "Desempeño Renta GPS") Then   'se agrega para Cargo por renta de GPS
                acRentaGPS = acRentaGPS + !Importe
                lblRentaGPS.Tag = Val(lblRentaGPS.Tag) + 1
            ElseIf !Cuenta = "828250" And (!Concepto = "Refrendo Renta GPS" Or !Concepto = "Desempeño Renta GPS") Then   'se agrega para Cargo RentaPoliza
                acRentaPoliza = acRentaPoliza + !Importe
                lblRentaPoliza.Tag = Val(lblRentaPoliza.Tag) + 1
                
            ElseIf !Cuenta = "120150" And !Concepto = "Pagos Fijos" And !Concepto <> "Apartado" Then
                acIvaIntFij = acIvaIntFij + !Importe
                lblIvaInteresesFijo.Tag = Val(lblIvaInteresesFijo.Tag) + 1
                
            ElseIf !Cuenta = "120150" And !Concepto = "Ventas" Then
                acIvaVen = acIvaVen + !Importe
                lblIvaVentas.Tag = Val(lblIvaVentas.Tag) + 1
                
            ElseIf !Cuenta = "120101" And !Iniciales <> "DEV01" And !Concepto <> "Apartado Cancelado" Then
                ccIva = ccIva + !Importe
                lblIvaCompras.Tag = Val(lblIvaCompras.Tag) + 1
                
            ElseIf (!Cuenta = "530150" Or !Cuenta = "120150") And !Concepto = "Boleta perdida" Then
                ccOtros = ccOtros + !Importe
                lblOtrosCobros.Tag = Val(lblOtrosCobros.Tag) + 1
            
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
        lblBoveda.Caption = Format(acBov, FMoneda)
        lblDesempeñoTrad.Caption = Format(acDes, FMoneda)
        lblAboRefrendoTrad.Caption = Format(acAboRef, FMoneda)
        lblInteresesTrad.Caption = Format(acInt, FMoneda)
        lblIvaInteresesTrad.Caption = Format(acIvaInt, FMoneda)
        
        lblDesempeñoFijo.Caption = Format(acDesFij, FMoneda)
        lblInteresesFijo.Caption = Format(acIntFij, FMoneda)
        lblMoratorios.Caption = Format(acMoratorios, FMoneda)
        lblIvaInteresesFijo.Caption = Format(acIvaIntFij, FMoneda)
        
        lblVentas.Caption = Format(acVen, FMoneda)
        lblIvaVentas.Caption = Format(acIvaVen, FMoneda)
        lblAboApartados.Caption = Format(acAbo, FMoneda)
        lblVenDivisas.Caption = Format(acDolares, FMoneda)
        lblOtrosCobros.Caption = Format(ccOtros, FMoneda)
        lblRentaGPS.Caption = Format(acRentaGPS, FMoneda)
        lblRentaPoliza.Caption = Format(acRentaPoliza, FMoneda)
        lblIVARenta.Caption = Format(acIVARenta, FMoneda)
        lblEntrada.Caption = Format(acBov + acDes + acAboRef + acInt + acIvaInt + acDesFij + acIntFij + acMoratorios + acIvaIntFij + acVen + acIvaVen + acAbo + acDolares + ccOtros + acRentaGPS + acRentaPoliza + acIVARenta, FMoneda)
        
        lblRefrendo.Caption = Format(acRef, FMoneda)
        lblDescuento.Caption = Format(acDesc, FMoneda)
        '*********************************************
        
        'SALIDAS**************************************
        lblRetiro.Caption = Format(acRet, FMoneda)
        lblEmpeñoTrad.Caption = Format(acEmp, FMoneda)
        lblEmpeñoFijo.Caption = Format(acEmpFij, FMoneda)
        lblComDivisas.Caption = Format(ccDolares, FMoneda)
        lblCompraVarios.Caption = Format(aInventario, FMoneda)
        lblIvaCompras.Caption = Format(ccIva, FMoneda)
        lblGastos.Caption = Format(acGas, FMoneda)
        lblDemasias.Caption = Format(acDemasia, FMoneda)
        lblDevoluciones.Caption = Format(acDev, FMoneda)
        '***Puntos***
        lblRedencionPuntos.Caption = Format(acRedencion, FMoneda)
        lblSalida.Caption = Format(acRet + acEmp + acEmpFij + ccDolares + aInventario + ccIva + acGas + acDemasia + acRedencion + acDev, FMoneda)
        '*********************************************
   
        lblTotal.Caption = Format(CDbl(lblEntrada.Caption) - CDbl(lblSalida.Caption), FMoneda)
        lblAjuste.Caption = Format(CDbl(lblTotal.Caption), FMoneda)
    
    End With
    Exit Sub
    
Error:
    Set rcVentanilla = Nothing
End Sub

Private Sub Imprimir_Corte()

On Error GoTo Error

    'Imprimimos el reporte de corte de caja
    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\CierreCaja.rpt"
        .Formulas(0) = "Boveda=" & ConvMoneda(lblBoveda.Caption) & ""
        .Formulas(1) = "MovBoveda='(" & Val(lblBoveda.Tag) & ")'"
        .Formulas(2) = "Desempeño=" & ConvMoneda(lblDesempeñoTrad.Caption) & ""
        .Formulas(3) = "MovDesempeño='(" & Val(lblDesempeñoTrad.Tag) & ")'"
        .Formulas(4) = "AboRefrendo=" & ConvMoneda(lblAboRefrendoTrad.Caption) & ""
        .Formulas(5) = "MovAboRefrendo='(" & Val(lblAboRefrendoTrad.Tag) & ")'"
        .Formulas(6) = "Intereses=" & ConvMoneda(lblInteresesTrad.Caption) & ""
        .Formulas(7) = "IvaIntereses=" & ConvMoneda(lblIvaInteresesTrad.Caption) & ""
        .Formulas(8) = "DesempeñoFijo=" & ConvMoneda(lblDesempeñoFijo.Caption) & ""
        .Formulas(9) = "MovDesempeñoFijo='(" & Val(lblDesempeñoFijo.Tag) & ")'"
        .Formulas(10) = "InteresesFijo=" & ConvMoneda(lblInteresesFijo.Caption) & ""
        .Formulas(11) = "Moratorios=" & ConvMoneda(lblMoratorios.Caption) & ""
        .Formulas(12) = "IvaInteresesFijo=" & ConvMoneda(lblIvaInteresesFijo.Caption) & ""
        .Formulas(13) = "Ventas=" & ConvMoneda(lblVentas.Caption) & ""
        .Formulas(14) = "MovVentas='(" & Val(lblVentas.Tag) & ")'"
        .Formulas(15) = "IvaVentas=" & ConvMoneda(lblIvaVentas.Caption) & ""
        .Formulas(16) = "AboApartados=" & ConvMoneda(lblAboApartados.Caption) & ""
        .Formulas(17) = "MovAboApartados='(" & Val(lblAboApartados.Tag) & ")'"
        .Formulas(18) = "VentaDivisas=" & ConvMoneda(lblVenDivisas.Caption) & ""
        .Formulas(19) = "MovVentaDivisas='(" & Val(lblVenDivisas.Tag) & ")'"
        .Formulas(20) = "OtrosCobros=" & ConvMoneda(lblOtrosCobros.Caption) & ""
        .Formulas(21) = "MovOtrosCobros='(" & Val(lblOtrosCobros.Tag) & ")'"
        .Formulas(22) = "TotEntrada=" & ConvMoneda(lblEntrada.Caption) & ""
        .Formulas(23) = "Refrendo=" & ConvMoneda(lblRefrendo.Caption) & ""
        .Formulas(24) = "MovRefrendo='(" & Val(lblRefrendo.Tag) & ")'"
        .Formulas(25) = "Descuento=" & ConvMoneda(lblDescuento.Caption) & ""
        
        .Formulas(26) = "Retiro=" & ConvMoneda(lblRetiro.Caption) & ""
        .Formulas(27) = "MovRetiro='(" & Val(lblRetiro.Tag) & ")'"
        .Formulas(28) = "Empeño=" & ConvMoneda(lblEmpeñoTrad.Caption) & ""
        .Formulas(29) = "MovEmpeño='(" & Val(lblEmpeñoTrad.Tag) & ")'"
        .Formulas(30) = "EmpeñoFijo=" & ConvMoneda(lblEmpeñoFijo.Caption) & ""
        .Formulas(31) = "MovEmpeñoFijo='(" & Val(lblEmpeñoFijo.Tag) & ")'"
        .Formulas(32) = "CompraDivisas=" & ConvMoneda(lblComDivisas.Caption) & ""
        .Formulas(33) = "MovCompraDivisas='(" & Val(lblComDivisas.Tag) & ")'"
        .Formulas(34) = "CompraVarios=" & ConvMoneda(lblCompraVarios.Caption) & ""
        .Formulas(35) = "MovCompraVarios='(" & Val(lblCompraVarios.Tag) & ")'"
        .Formulas(36) = "IvaCompras=" & ConvMoneda(lblIvaCompras.Caption) & ""
        .Formulas(37) = "Gastos=" & ConvMoneda(lblGastos.Caption) & ""
        .Formulas(38) = "MovGastos='(" & Val(lblGastos.Tag) & ")'"
        .Formulas(39) = "Demasias=" & ConvMoneda(lblDemasias.Caption) & ""
        .Formulas(40) = "MovDemasias='(" & Val(lblDemasias.Tag) & ")'"
        
        '***Puntos***
        .Formulas(41) = "RedencionPuntos=" & ConvMoneda(lblRedencionPuntos.Caption) & ""
        .Formulas(42) = "MovRedencionPuntos='(" & Val(lblRedencionPuntos.Tag) & ")'"
        
        .Formulas(43) = "TotSalida=" & ConvMoneda(lblSalida.Caption) & ""
        
        .Formulas(44) = "Total=" & ConvMoneda(lblTotal.Caption) & ""
        .Formulas(45) = "Efectivo=" & ConvMoneda(txtEfectivo.text) & ""
        .Formulas(46) = "Ajuste=" & ConvMoneda(lblAjuste.Caption) & ""
        
        .Formulas(47) = "Caja='" & NombrePc & "'"
        .Formulas(48) = "Cajero='" & frmMDI.Usuario & "'"
        .Formulas(49) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(50) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(51) = "Usuario='" & SacaValor("usuarios", "Nombre", " WHERE ID=" & Val(frmMDI.IDUsuario)) & "'"
        .Formulas(52) = "Gerente='" & Trim(Regresa_Valor_BD("Gerente")) & "'"
        .Formulas(53) = "Devoluciones=" & ConvMoneda(lblDevoluciones.Caption) & ""
        .Formulas(54) = "MovDevoluciones='(" & Val(lblDevoluciones.Tag) & ")'"
        .Formulas(55) = "InfRefrendos=" & ConvMoneda(lblRefrendo.Caption)
        .Formulas(56) = "MovInfRefrendos='(" & Val(lblRefrendo.Tag) & ")'"
        .Formulas(57) = "RentaGPS=" & ConvMoneda(lblRentaGPS.Caption) & ""
        .Formulas(58) = "RentaSeguro=" & ConvMoneda(lblRentaPoliza.Caption) & ""
        .Formulas(59) = "IvaRentaGPS=" & ConvMoneda(lblIVARenta.Caption) & ""
        
        .WindowTitle = "Cierre de caja"
        .SelectionFormula = "{Empeno.PC}='" & NombrePc & "' AND {Empeno.Origen}=" & OD_EMPENO & " AND {Empeno.Captura}=0 AND {Empeno.Fecha}=date('" & Format(Date, "YYYY,MM,DD") & "')"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub cmdAceptar_Click()
    
    'Checo si hay un ajuste
    If Val(lblAjuste.Tag) <> 0 And Val(txtEfectivo.Tag) = 0 Then
        
        If MsgBox("Se realizará un ajuste por la cantidad de " & lblAjuste.Caption & vbCrLf & "Desea continuar ??", vbQuestion + vbYesNo + vbDefaultButton2, "Cierre de Caja") = vbNo Then Exit Sub
    
    End If
    
    If Val(txtEfectivo.Tag) = 0 Then
        
        If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
            
            Cargos_Abonos
            Marcar
            txtEfectivo.Tag = 1
            Reporte_Auxiliar True, Date, Date
            Sleep 1000
            
        Else
            
            MsgBox "Favor de poner el efectivo de caja !!", vbCritical, "Cierre de Caja"
            Exit Sub
        End If
        
    End If
    
    Imprimir_Corte
    
End Sub

Private Sub cmdModificaCorte_Click()
    
    frmPasswords.ConexSuc = 0
    frmPasswords.DescuentoVentas = 0
    frmPasswords.PrecioVitrina = 0
    frmPasswords.Cancel = 0
    frmPasswords.Ventas = 0
    frmPasswords.HacerCorte = 0
    frmPasswords.ModificaPrecio = 0
    frmPasswords.InteresDesempeño = 0
    frmPasswords.InteresRefrendo = 0
    frmPasswords.RecalculoPrecios = 0
    frmPasswords.AutorizaPrestamo = 0
    frmPasswords.Vencido = 0
    frmPasswords.CancelaCierre = 0
    frmPasswords.ModificaCorte = 1

    If frmPasswords.Password(GERENTE, 1) Then
    
        txtEfectivo.Enabled = True
        txtEfectivo.SetFocus
    
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl()
End Sub



Private Sub txtEfectivo_Change()
Dim crAjuste As Double, Efectivo As Double

    If Trim(txtEfectivo.text) <> "" Or Val(txtEfectivo.text) > 0 Then
        
        Efectivo = txtEfectivo.text
        
    Else
        Efectivo = 0
    
    End If
  
    crAjuste = CCur(lblTotal.Caption) - Efectivo
    lblAjuste.Caption = Format(Format(crAjuste, "Currency"), FMoneda)
    lblAjuste.Tag = crAjuste

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
    Cambiar_Color False, txtEfectivo
    txtEfectivo.text = Format(txtEfectivo.text, FMoneda)
    txtEfectivo.Enabled = False
End Sub

Private Sub Cargos_Abonos()
Dim crImporte As Double, crAjuste As Double, Folio As Long, Movimiento As Long, Hora As String
    
    If Val(lblAjuste.Tag) > 0 Or Trim(lblAjuste.Tag) <> "" Then
        
        crAjuste = CDbl(lblAjuste.Tag)
    Else
    
        crAjuste = 0
    End If
    
    'Tomo la Hora
    Hora = Time
    
    'Saca el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Saca el Folio
    Folio = Regresa_Movimiento(False, "FolioBoveda")
    Regresa_Movimiento True, "FolioBoveda"
    
    If crAjuste <> 0 Then
            
    
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & ",0,'PR01','" & IIf(crAjuste < 0, "110101", "151301") & "'," & ConvMoneda(IIf(crAjuste < 0, -1 * crAjuste, crAjuste)) & "," & TIPO_CARGO & ",0,'Ajuste','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                  
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & ",0,'PR50','" & IIf(crAjuste < 0, "151350", "110150") & "'," & ConvMoneda(IIf(crAjuste < 0, -1 * crAjuste, crAjuste)) & "," & TIPO_ABONO & ",0,'Ajuste','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                      
    End If
            
    If Val(txtEfectivo.text) > 0 Or Trim(txtEfectivo.text) <> "" Then
        
        crImporte = txtEfectivo.text
    
    Else
        
        crImporte = 0
    End If
    
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " _
                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'CV01','110901'," & ConvMoneda(crImporte) & "," & TIPO_CARGO & ",0,'Corte de Caja','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                  
'''    'Grabamos el abono
'''    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " _
'''                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'CV50','199450'," & ConvMoneda(crImporte) & "," & TIPO_ABONO & ",0,'Corte de Caja','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " _
                    & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'CV50','110150'," & ConvMoneda(crImporte) & "," & TIPO_ABONO & ",0,'Corte de Caja','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                  
End Sub

Private Sub Marcar()

    dbDatos.Execute "UPDATE auxiliar SET Corte=1 WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "'"
   
    dbDatos.Execute "UPDATE empeno SET Corte=1 WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND Corte=0 AND PC='" & NombrePc & "'"

End Sub

Private Sub Limpiar()

    txtEfectivo.text = ""
    txtEfectivo.Tag = ""
    
    lblBoveda.Caption = "0.00"
    lblEmpeñoTrad.Caption = "0.00"
    lblDesempeñoTrad.Caption = "0.00"
    lblRefrendo.Caption = "0.00"
    lblAboRefrendoTrad.Caption = "0.00"
    lblInteresesTrad.Caption = "0.00"
    lblVentas.Caption = "0.00"
    lblAboApartados.Caption = "0.00"
    lblCambios.Caption = "0.00"
    lblTotal.Caption = "0.00"
    lblAjuste.Caption = "0.00"
    lblRetiro.Caption = "0.00"
    lblPrestamos.Caption = "0.00"
    lblGastos.Caption = "0.00"
    lblDevoluciones.Caption = "0.00"
    lblVenDivisas.Caption = "0.00"
    lblComDivisas.Caption = "0.00"
    cmdAceptar.Enabled = False
    lblCompraVarios.Caption = "0.00"
    lblIvaInteresesTrad.Caption = "0.00"
    lblIvaVentas.Caption = "0.00"
    lblOtrosCobros.Caption = "0.00"
    lblDevoluciones.Caption = "0.00"

End Sub

'Sacamos el reporte auxiliar
Private Sub Reporte_Auxiliar(Optional Opcion As Boolean = False, Optional FechaIni As String, Optional FechaFin As String)
Dim rcAuxiliar As New ADODB.Recordset
Dim crCargo As Currency, crAbono As Currency
Dim crSaldo As Currency, strCuenta As String
   
    On Error GoTo Error

    dbReportes.Execute "DELETE FROM cortecuentas"

    If Opcion Then
    
        rcAuxiliar.Open "SELECT auxiliar.*,cuentas.Mayor,cuentas.Concepto as ConceptoCuenta FROM auxiliar,cuentas WHERE cuentas.Cuenta=auxiliar.Cuenta AND Fecha BETWEEN '" & Format(FechaIni, "YYYY/MM/DD") & "' AND '" & Format(FechaFin, "YYYY/MM/DD") & "' AND Corte=0 AND PC='" & NombrePc & "' ORDER BY cuentas.Mayor,Fecha,auxiliar.ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    
    Else
        
        rcAuxiliar.Open "SELECT auxiliar.*,cuentas.Mayor,cuentas.Concepto as ConceptoCuenta FROM auxiliar,cuentas  WHERE cuentas.Cuenta=auxiliar.Cuenta AND Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND Corte=0 AND PC='" & NombrePc & "' ORDER BY cuentas.Mayor,Fecha,auxiliar.ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    
    End If
   
    With rcAuxiliar
    
        While Not .EOF
            
            crCargo = 0
            crAbono = 0

            If strCuenta <> !Mayor Then
                
                crSaldo = 0
                strCuenta = !Mayor
            
            End If

            If Right(!Cuenta, 2) = "01" Then
                
                crCargo = !Importe
                crSaldo = crSaldo + crCargo
            
            Else
                
                crAbono = !Importe
                crSaldo = crSaldo - crAbono
            
            End If
         
            dbReportes.Execute "INSERT INTO cortecuentas (Cuenta,Descripcion,Fecha,Concepto,Folio,Cargo,Abono,Saldo) VALUES " & "('" & !Mayor & "','" & ![ConceptoCuenta] & "','" & Format(!Fecha, "YYYY/MM/DD") & "','" & ![Concepto] & "'," & !Folio & "," & ConvMoneda(crCargo) & "," & ConvMoneda(crAbono) & "," & ConvMoneda(crSaldo) & ")"
        
        .MoveNext
        Wend
        
    End With
    rcAuxiliar.Close
    Set rcAuxiliar = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcAuxiliar = Nothing
End Sub

Private Sub Checar_Auxiliar()
Dim rcEmpeño As New ADODB.Recordset
Dim rcAuxiliar As New ADODB.Recordset

'checamos los empeños del dia
rcEmpeño.Open "SELECT * FROM Empeno WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND folio=folioorigen ORDER BY ID", dbDatos, adOpenDynamic, adLockOptimistic
With rcEmpeño
   While Not .EOF
      rcAuxiliar.Open "SELECT * FROM Auxiliar WHERE Folio=" & !Folio & " AND Serie=" & !Serie & " ORDER BY ID", dbDatos, adOpenDynamic, adLockOptimistic
      If rcAuxiliar.RecordCount <> 3 Then
         If rcAuxiliar!Cuenta = "201701" Then
            rcAuxiliar.MoveNext
         Else
            dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "','Empeno',0," & !Folio & ",'" & !Iniciales & "','201701'," & !Prestamo & "," & TIPO_CARGO & "," & !Serie & ",'" & NombrePc & "')"
            rcAuxiliar.MoveNext
         End If
         If rcAuxiliar!Cuenta = "110150" Then
            rcAuxiliar.MoveNext
         Else
            dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "','Empeno',0," & !Folio & ",'" & !Iniciales & "','110150'," & !Prestamo & "," & TIPO_ABONO & "," & !Serie & ",'" & NombrePc & "')"
            rcAuxiliar.MoveNext
         End If
'''         If rcAuxiliar!Cuenta = "199450" Then
'''            rcAuxiliar.MoveNext
'''         Else
'''            dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC) VALUES " & _
'''                  "('" & Format(Date, "YYYY/MM/DD") & "','Empeno',0," & !Folio & ",'" & !Iniciales & "','199450'," & !Prestamo & "," & TIPO_ABONO & "," & !Serie & ",'" & NombrePc & "')"
'''            rcAuxiliar.MoveNext
'''         End If
      End If
      rcAuxiliar.Close
   Wend
   .Close
End With

End Sub

Sub LimpiaMontos()
    
    txtEfectivo.text = ""
    txtEfectivo.Tag = ""
    
    'INGRESOS
    acBov = 0
    acDes = 0
    acAboRef = 0
    acInt = 0
    acIvaInt = 0
    acDesFij = 0
    acIntFij = 0
    acMoratorios = 0
    acIvaIntFij = 0
    
    acVen = 0
    acDesc = 0
    acIvaVen = 0
    acAbo = 0
    acDolares = 0
    ccOtros = 0
    acRef = 0
    acDesc = 0
                
    acRentaGPS = 0
    acRentaPoliza = 0
    acIVARenta = 0
                
    lblBoveda.Caption = Format(0, FMoneda)
    lblBoveda.Tag = ""
    
    lblDesempeñoTrad.Caption = Format(0, FMoneda)
    lblDesempeñoTrad.Tag = ""
    lblAboRefrendoTrad.Caption = Format(0, FMoneda)
    lblAboRefrendoTrad.Tag = ""
    lblInteresesTrad.Caption = Format(0, FMoneda)
    lblInteresesTrad.Tag = ""
    lblIvaInteresesTrad.Caption = Format(0, FMoneda)
    lblIvaInteresesTrad.Tag = ""
    
    lblDesempeñoFijo.Caption = Format(0, FMoneda)
    lblDesempeñoFijo.Tag = ""
    lblInteresesFijo.Caption = Format(0, FMoneda)
    lblInteresesFijo.Tag = ""
    lblMoratorios.Caption = Format(0, FMoneda)
    lblMoratorios.Tag = ""
    lblIvaInteresesFijo.Caption = ""
    lblIvaInteresesFijo.Tag = ""
    
    lblVentas.Caption = Format(0, FMoneda)
    lblVentas.Tag = ""
    lblIvaVentas.Caption = Format(0, FMoneda)
    lblIvaVentas.Tag = ""
    lblAboApartados.Caption = Format(0, FMoneda)
    lblAboApartados.Tag = ""
    lblVenDivisas.Caption = Format(0, FMoneda)
    lblVenDivisas.Tag = ""
    lblOtrosCobros.Caption = Format(0, FMoneda)
    lblOtrosCobros.Tag = ""
    lblEntrada.Caption = Format(0, FMoneda)
    lblEntrada.Tag = ""
    lblRefrendo.Caption = Format(0, FMoneda)
    lblRefrendo.Tag = ""
    lblDescuento.Caption = Format(0, FMoneda)
    lblDescuento.Tag = ""
    lblRentaGPS.Caption = Format(0, FMoneda)
    lblRentaGPS.Tag = ""
    lblRentaPoliza.Caption = Format(0, FMoneda)
    lblRentaPoliza.Tag = ""
    lblIVARenta.Caption = Format(0, FMoneda)
    lblIVARenta.Tag = ""
    
    'EGRESOS
    acRet = 0
    acEmp = 0
    acEmpFij = 0
    ccDolares = 0
    aInventario = 0
    ccIva = 0
    acGas = 0
    acDemasia = 0
    acDev = 0
    '***Puntos***
    acRedencion = 0
    
    lblRetiro.Caption = Format(0, FMoneda)
    lblRetiro.Tag = ""
    lblEmpeñoTrad.Caption = Format(0, FMoneda)
    lblEmpeñoTrad.Tag = ""
    lblEmpeñoFijo.Caption = Format(0, FMoneda)
    lblEmpeñoFijo.Tag = ""
    lblComDivisas.Caption = Format(0, FMoneda)
    lblComDivisas.Tag = ""
    lblCompraVarios.Caption = Format(0, FMoneda)
    lblCompraVarios.Tag = ""
    lblIvaCompras.Caption = Format(0, FMoneda)
    lblIvaCompras.Tag = ""
    lblGastos.Caption = Format(0, FMoneda)
    lblGastos.Tag = ""
    lblDemasias.Caption = Format(0, FMoneda)
    lblDemasias.Tag = ""
    lblSalida.Caption = Format(0, FMoneda)
    lblSalida.Tag = ""
    lblTotal.Caption = Format(0, FMoneda)
    lblTotal.Tag = ""
    lblAjuste.Caption = Format(0, FMoneda)
    lblAjuste.Tag = ""
    lblDevoluciones.Caption = Format(0, FMoneda)
    lblDevoluciones.Tag = ""
    
    '***Puntos***
    lblRedencionPuntos.Caption = Format(0, FMoneda)
    lblRedencionPuntos.Tag = ""
    
End Sub

Function VerificaEmpenos() As Boolean
Dim rcConsulta As New ADODB.Recordset
Dim NumEmpenos As Long

On Error GoTo Error
    
    VerificaEmpenos = True
    NumEmpenos = SacaValor("empeno", "COUNT(ID)", " WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "' AND Cancelado=0 AND Origen=" & OD_EMPENO)
    
    If NumEmpenos > 0 Then
        
        rcConsulta.Open "SELECT Folio FROM Empeno WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "' AND Verificado=0 AND Cancelado=0 AND Origen=" & OD_EMPENO, dbDatos, adOpenForwardOnly, adLockReadOnly
        If Not rcConsulta.BOF And Not rcConsulta.EOF Then
            
            VerificaEmpenos = False
        End If
        rcConsulta.Close
    
    End If
    Set rcConsulta = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function
