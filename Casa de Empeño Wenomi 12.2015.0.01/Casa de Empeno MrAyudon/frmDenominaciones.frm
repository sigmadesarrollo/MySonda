VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmDenominaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arqueo"
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7575
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDenominaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   7575
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "ARQUEO"
      Height          =   3495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   7590
      Begin VB.Frame Frame2 
         Caption         =   "MONEDAS"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3360
         Left            =   3840
         TabIndex        =   17
         Top             =   60
         Width           =   3705
         Begin VB.TextBox txtMoCincuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   11
            Tag             =   ".50"
            Text            =   "0"
            Top             =   2175
            Width           =   885
         End
         Begin VB.TextBox txtMopeso 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   10
            Tag             =   "1"
            Text            =   "0"
            Top             =   1815
            Width           =   885
         End
         Begin VB.TextBox txtMocinco 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   8
            Tag             =   "5"
            Text            =   "0"
            Top             =   1095
            Width           =   885
         End
         Begin VB.TextBox txtModiez 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   7
            Tag             =   "10"
            Text            =   "0"
            Top             =   735
            Width           =   885
         End
         Begin VB.TextBox txtMoveinte 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   6
            Tag             =   "20"
            Text            =   "0"
            Top             =   375
            Width           =   885
         End
         Begin VB.TextBox txtModospesos 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   9
            Tag             =   "2"
            Text            =   "0"
            Top             =   1455
            Width           =   885
         End
         Begin VB.TextBox txtMorralla 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1260
            MaxLength       =   5
            TabIndex        =   12
            Tag             =   ".01"
            Text            =   "0"
            Top             =   2535
            Width           =   885
         End
         Begin VB.Label lblTotalMonedas 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3030
            TabIndex        =   48
            Top             =   3000
            Width           =   465
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   ".50"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   705
            TabIndex        =   32
            Top             =   2190
            Width           =   300
         End
         Begin VB.Label Label12 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   585
            TabIndex        =   31
            Top             =   1830
            Width           =   420
         End
         Begin VB.Label Label13 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "5.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   585
            TabIndex        =   30
            Top             =   1110
            Width           =   420
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "10.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   465
            TabIndex        =   29
            Top             =   750
            Width           =   540
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "20.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   465
            TabIndex        =   28
            Top             =   390
            Width           =   540
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "2.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   585
            TabIndex        =   27
            Top             =   1470
            Width           =   420
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Centavos"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   105
            TabIndex        =   26
            Top             =   2520
            Width           =   1035
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TOTAL MONEDAS:"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   60
            TabIndex        =   25
            Top             =   3000
            Width           =   1905
         End
         Begin VB.Label lblMoveinte 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3075
            TabIndex        =   24
            Top             =   390
            Width           =   420
         End
         Begin VB.Label lblMocincuenta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3075
            TabIndex        =   23
            Top             =   2190
            Width           =   420
         End
         Begin VB.Label lblMopeso 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3075
            TabIndex        =   22
            Top             =   1830
            Width           =   420
         End
         Begin VB.Label lblModospesos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3075
            TabIndex        =   21
            Top             =   1470
            Width           =   420
         End
         Begin VB.Label lblMocinco 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3075
            TabIndex        =   20
            Top             =   1110
            Width           =   420
         End
         Begin VB.Label lblModiez 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3075
            TabIndex        =   19
            Top             =   750
            Width           =   420
         End
         Begin VB.Label lblMorralla 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3075
            TabIndex        =   18
            Top             =   2520
            Width           =   420
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "BILLETES"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   3360
         Left            =   45
         TabIndex        =   33
         Top             =   60
         Width           =   3705
         Begin VB.TextBox txtBimil 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   0
            Tag             =   "1000"
            Text            =   "0"
            Top             =   375
            Width           =   885
         End
         Begin VB.TextBox txtBiQuinientos 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   1
            Tag             =   "500"
            Text            =   "0"
            Top             =   735
            Width           =   885
         End
         Begin VB.TextBox txtBiDoscientos 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   2
            Tag             =   "200"
            Text            =   "0"
            Top             =   1095
            Width           =   885
         End
         Begin VB.TextBox txtBiCien 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   3
            Tag             =   "100"
            Text            =   "0"
            Top             =   1455
            Width           =   885
         End
         Begin VB.TextBox txtBiCincuenta 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   4
            Tag             =   "50"
            Text            =   "0"
            Top             =   1815
            Width           =   885
         End
         Begin VB.TextBox txtBiVeinte 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   1320
            MaxLength       =   5
            TabIndex        =   5
            Tag             =   "20"
            Text            =   "0"
            Top             =   2175
            Width           =   885
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "1,000.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   225
            TabIndex        =   47
            Top             =   390
            Width           =   840
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "500.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   405
            TabIndex        =   46
            Top             =   750
            Width           =   660
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "200.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   405
            TabIndex        =   45
            Top             =   1110
            Width           =   660
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "100.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   405
            TabIndex        =   44
            Top             =   1470
            Width           =   660
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "50.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   525
            TabIndex        =   43
            Top             =   1830
            Width           =   540
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "20.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   525
            TabIndex        =   42
            Top             =   2190
            Width           =   540
         End
         Begin VB.Label lblBiMil 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3135
            TabIndex        =   41
            Top             =   390
            Width           =   420
         End
         Begin VB.Label lblBiCincuenta 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3135
            TabIndex        =   40
            Top             =   1830
            Width           =   420
         End
         Begin VB.Label lblBiCien 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3135
            TabIndex        =   39
            Top             =   1470
            Width           =   420
         End
         Begin VB.Label lblBiDoscientos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3135
            TabIndex        =   38
            Top             =   1110
            Width           =   420
         End
         Begin VB.Label lblBiQuinientos 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3135
            TabIndex        =   37
            Top             =   750
            Width           =   420
         End
         Begin VB.Label lblBiVeinte 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   3135
            TabIndex        =   36
            Top             =   2190
            Width           =   420
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "TOTAL BILLETES:"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Left            =   60
            TabIndex        =   35
            Top             =   3000
            Width           =   1620
         End
         Begin VB.Label lblTotalBilletes 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Century Gothic"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   3090
            TabIndex        =   34
            Top             =   3000
            Width           =   465
         End
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   5280
      TabIndex        =   13
      Top             =   3525
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
      Picture         =   "frmDenominaciones.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6420
      TabIndex        =   14
      Top             =   3525
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
      Picture         =   "frmDenominaciones.frx":055E
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "TOTAL ARQUEO:"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   90
      TabIndex        =   49
      Top             =   3540
      Width           =   2160
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   2520
      TabIndex        =   15
      Top             =   3540
      Width           =   630
   End
End
Attribute VB_Name = "frmDenominaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim crTotal As Double, ImprimeReporte As Boolean, Salir As Integer

Private Sub cmdAceptar_Click()
Dim Centavos As Double
    
    If Val(lblTotal.Caption) > 0 Or Trim(lblTotal.Caption) <> "" Then
            
        crTotal = lblTotal.Caption
    Else
        
        crTotal = 0
    End If
    
    If Val(txtMorralla.text) > 0 Or (Trim(txtMorralla.text) <> "" And Trim(txtMorralla.text) <> ".") Then
    
        Centavos = CDbl(txtMorralla.text)
    Else
    
        Centavos = 0
    End If
        
    If ImprimeReporte And (crTotal > 0 Or Centavos > 0) Then
            
        If MsgBox("Desea imprimir el arqueo ??", vbQuestion + vbYesNo + vbDefaultButton1, "Arqueo") = vbYes Then
    
            With frmMDI.Cr
                .Reset
                .DiscardSavedData = True
                .ReportFileName = Path & "\Reportes\Arqueo.rpt"
                .Formulas(0) = "1000=" & txtBimil.text & ""
                .Formulas(1) = "Imp1000=" & ConvMoneda(lblBiMil.Caption) & ""
                .Formulas(2) = "500=" & txtBiQuinientos.text & ""
                .Formulas(3) = "Imp500=" & ConvMoneda(lblBiQuinientos.Caption) & ""
                .Formulas(4) = "200=" & txtBiDoscientos.text & ""
                .Formulas(5) = "Imp200=" & ConvMoneda(lblBiDoscientos.Caption) & ""
                .Formulas(6) = "100=" & txtBiCien.text & ""
                .Formulas(7) = "Imp100=" & ConvMoneda(lblBiCien.Caption) & ""
                .Formulas(8) = "50=" & txtBiCincuenta.text & ""
                .Formulas(9) = "Imp50=" & ConvMoneda(lblBiCincuenta.Caption) & ""
                .Formulas(10) = "20=" & txtBiVeinte.text & ""
                .Formulas(11) = "Imp20=" & ConvMoneda(lblBiVeinte.Caption) & ""
                .Formulas(12) = "20P=" & txtMoveinte.text & ""
                .Formulas(13) = "Imp20P=" & ConvMoneda(lblMoveinte.Caption) & ""
                .Formulas(14) = "10P=" & txtModiez.text & ""
                .Formulas(15) = "Imp10P=" & ConvMoneda(lblModiez.Caption) & ""
                .Formulas(16) = "5P=" & txtMocinco.text & ""
                .Formulas(17) = "Imp5P=" & ConvMoneda(lblMocinco.Caption) & ""
                .Formulas(18) = "2P=" & txtModospesos.text & ""
                .Formulas(19) = "Imp2P=" & ConvMoneda(lblModospesos.Caption) & ""
                .Formulas(20) = "1P=" & txtMopeso.text & ""
                .Formulas(21) = "Imp1P=" & ConvMoneda(lblMopeso.Caption) & ""
                .Formulas(22) = "50C=" & txtMoCincuenta.text & ""
                .Formulas(23) = "Imp50C=" & ConvMoneda(lblMocincuenta.Caption) & ""
                .Formulas(24) = "Morralla=" & txtMorralla.text & ""
                .Formulas(25) = "ImpMorralla=" & ConvMoneda(lblMorralla.Caption) & ""
                .Formulas(26) = "TotBilletes=" & ConvMoneda(lblTotalBilletes.Caption) & ""
                .Formulas(27) = "TotMonedas=" & ConvMoneda(lblTotalMonedas.Caption) & ""
                
                .Formulas(28) = "Total=" & ConvMoneda(lblTotal.Caption) & ""
                .Formulas(29) = "RazonSocial='" & Sucursal.RazonSocial & "'"
                .Formulas(30) = "Sucursal='SUCURSAL " & Sucursal.NombreComercial & "'"
                .Formulas(31) = "Direccion='" & Sucursal.Direccion & ", " & Sucursal.Ciudad & " " & Sucursal.Estado & "'"
                .Formulas(32) = "Telefono='TEL. " & Sucursal.Telefono & "'"
                .Formulas(33) = "Rfc='RFC: " & Sucursal.RFC & "'"
                
                .WindowShowPrintSetupBtn = True
                .WindowTitle = "Arqueo"
                .WindowState = crptMaximized
                .Destination = crptToWindow
                .Action = 1
            End With
    
        End If
    
    End If
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Salir = 1
    Unload Me
End Sub

Private Sub Form_Load()
    Salir = 0
    txtMoCincuenta.Tag = Separador & "50"
    txtMorralla.Tag = Separador & "01"
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    If UnloadMode = 0 Or Salir = 1 Then
        crTotal = -1
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub lblTotalBilletes_Change()
    TotalGral
End Sub

Private Sub lblTotalMonedas_Change()
    TotalGral
End Sub

Private Sub txtBicien_Change()
    CalculaImporte txtBiCien, lblBiCien
    CalculaTotal "BILLETES"
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
    CalculaImporte txtBiCincuenta, lblBiCincuenta
    CalculaTotal "BILLETES"
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

Private Sub txtBidoscientos_Change()
    CalculaImporte txtBiDoscientos, lblBiDoscientos
    CalculaTotal "BILLETES"
End Sub

Private Sub txtBidoscientos_GotFocus()
    Seleccionar_Texto txtBiDoscientos
    Cambiar_Color True, txtBiDoscientos
End Sub

Private Sub txtBidoscientos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBidoscientos_LostFocus()
    Cambiar_Color False, txtBiDoscientos
End Sub

Private Sub txtBimil_Change()
    CalculaImporte txtBimil, lblBiMil
    CalculaTotal "BILLETES"
End Sub

Private Sub txtBimil_GotFocus()
    Seleccionar_Texto txtBimil
    Cambiar_Color True, txtBimil
End Sub

Private Sub txtBimil_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBimil_LostFocus()
    Cambiar_Color False, txtBimil
End Sub

Private Sub txtBiquinientos_Change()
    CalculaImporte txtBiQuinientos, lblBiQuinientos
    CalculaTotal "BILLETES"
End Sub

Private Sub txtBiquinientos_GotFocus()
    Seleccionar_Texto txtBiQuinientos
    Cambiar_Color True, txtBiQuinientos
End Sub

Private Sub txtBiquinientos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBiquinientos_LostFocus()
    Cambiar_Color False, txtBiQuinientos
End Sub

Private Sub txtBiveinte_Change()
    CalculaImporte txtBiVeinte, lblBiVeinte
    CalculaTotal "BILLETES"
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

Private Sub txtMocinco_Change()
    CalculaImporte txtMocinco, lblMocinco
    CalculaTotal "MONEDAS"
End Sub

Private Sub txtMocinco_GotFocus()
    Seleccionar_Texto txtMocinco
    Cambiar_Color True, txtMocinco
End Sub

Private Sub txtMocinco_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMocinco_LostFocus()
    Cambiar_Color False, txtMocinco
End Sub

Private Sub txtMocincuenta_Change()
    CalculaImporte txtMoCincuenta, lblMocincuenta
    CalculaTotal "MONEDAS"
End Sub

Private Sub txtMocincuenta_GotFocus()
    Seleccionar_Texto txtMoCincuenta
    Cambiar_Color True, txtMoCincuenta
End Sub

Private Sub txtMocincuenta_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMocincuenta_LostFocus()
    Cambiar_Color False, txtMoCincuenta
End Sub

Private Sub txtModiez_Change()
    CalculaImporte txtModiez, lblModiez
    CalculaTotal "MONEDAS"
End Sub

Private Sub txtModiez_GotFocus()
    Seleccionar_Texto txtModiez
    Cambiar_Color True, txtModiez
End Sub

Private Sub txtModiez_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtModiez_LostFocus()
    Cambiar_Color False, txtModiez
End Sub

Private Sub txtModospesos_Change()
    CalculaImporte txtModospesos, lblModospesos
    CalculaTotal "MONEDAS"
End Sub

Private Sub txtModospesos_GotFocus()
    Seleccionar_Texto txtModospesos
    Cambiar_Color True, txtModospesos
End Sub

Private Sub txtModospesos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtModospesos_LostFocus()
    Cambiar_Color False, txtModospesos
End Sub

Private Sub txtMopeso_Change()
    CalculaImporte txtMopeso, lblMopeso
    CalculaTotal "MONEDAS"
End Sub

Private Sub txtMopeso_GotFocus()
    Seleccionar_Texto txtMopeso
    Cambiar_Color True, txtMopeso
End Sub

Private Sub txtMopeso_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMopeso_LostFocus()
    Cambiar_Color False, txtMopeso
End Sub

Private Sub txtMorralla_Change()
    CalculaImporte txtMorralla, lblMorralla
    CalculaTotal "MONEDAS"
End Sub

Private Sub txtMorralla_GotFocus()
    Seleccionar_Texto txtMorralla
    Cambiar_Color True, txtMorralla
End Sub

Private Sub txtMorralla_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMorralla_LostFocus()
    Cambiar_Color False, txtMorralla
End Sub

Private Sub txtMoveinte_Change()
    CalculaImporte txtMoveinte, lblMoveinte
    CalculaTotal "MONEDAS"
End Sub

Private Sub txtMoveinte_GotFocus()
    Seleccionar_Texto txtMoveinte
    Cambiar_Color True, txtMoveinte
End Sub

Private Sub txtMoveinte_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMoveinte_LostFocus()
    Cambiar_Color False, txtMoveinte
End Sub

Sub CalculaTotal(Contenedor As String)
Dim txt As Object, Etiqueta As Label, Cantidad As Long, crDenominacion As Double

    crTotal = 0
    For Each txt In Me.Controls
        
        If txt.Container.Caption = Contenedor Then
            
            If TypeOf txt Is TextBox Then
                
                Cantidad = 0
                crDenominacion = CDbl(txt.Tag)
                If Val(txt.text) > 0 Or Trim(txt.text) <> "" Then
                    
                    Cantidad = txt.text
                End If
                
                crTotal = crTotal + (Cantidad * crDenominacion)
            
            End If
        
        End If
    
    Next
    
    Set Etiqueta = IIf(Contenedor = "BILLETES", lblTotalBilletes, lblTotalMonedas)
    Etiqueta.Caption = Format(crTotal, FMoneda)
End Sub

Sub CalculaImporte(caja As TextBox, Etiqueta As Label)
Dim Cantidad As Long, crDenominacion As Double
    
    Cantidad = 0
    crDenominacion = CDbl(caja.Tag)
    If Val(caja.text) > 0 Or Trim(caja.text) <> "" Then
        
        Cantidad = caja.text
    End If
    
    Etiqueta.Caption = Format(Cantidad * crDenominacion, FMoneda)
End Sub

Sub TotalGral()
Dim crBilletes As Double, crMonedas As Double
    
    crBilletes = lblTotalBilletes.Caption
    crMonedas = lblTotalMonedas.Caption
    
    lblTotal = Format(crBilletes + crMonedas, FMoneda)
End Sub

Public Function Arqueo(Optional Imprimir As Boolean = False) As Double
   crTotal = 0
   ImprimeReporte = Imprimir
   Me.Show vbModal
   Arqueo = crTotal
End Function
