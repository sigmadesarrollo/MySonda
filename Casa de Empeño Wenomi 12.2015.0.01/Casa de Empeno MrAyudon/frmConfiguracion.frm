VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmConfiguracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parámetros"
   ClientHeight    =   5730
   ClientLeft      =   2100
   ClientTop       =   1995
   ClientWidth     =   9360
   BeginProperty Font 
      Name            =   "Century Gothic"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfiguracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5730
   ScaleWidth      =   9360
   Tag             =   "Avaluoprestamo"
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
      Height          =   4935
      Left            =   120
      TabIndex        =   58
      Top             =   120
      Width           =   9075
      Begin TabDlg.SSTab SSTab1 
         Height          =   4785
         Left            =   120
         TabIndex        =   62
         Top             =   120
         Width           =   8745
         _ExtentX        =   15425
         _ExtentY        =   8440
         _Version        =   393216
         Style           =   1
         Tabs            =   4
         TabsPerRow      =   4
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
         TabCaption(0)   =   "PERIODOS/TASAS"
         TabPicture(0)   =   "frmConfiguracion.frx":2832
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "Frame4"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "Frame6"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "Frame7"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "FraGPS"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         TabCaption(1)   =   "OTRAS VARIABLES"
         TabPicture(1)   =   "frmConfiguracion.frx":284E
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame5"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "PRECIOS ORO"
         TabPicture(2)   =   "frmConfiguracion.frx":286A
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "txtPrecioAutos"
         Tab(2).Control(1)=   "Frame3"
         Tab(2).Control(2)=   "Frame8"
         Tab(2).Control(3)=   "cmdRecalcular"
         Tab(2).Control(4)=   "Label57"
         Tab(2).ControlCount=   5
         TabCaption(3)   =   "SEMAFORO"
         TabPicture(3)   =   "frmConfiguracion.frx":2886
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "lbl5"
         Tab(3).Control(1)=   "FramSemaforo"
         Tab(3).Control(2)=   "FramPrestamo"
         Tab(3).ControlCount=   3
         Begin VB.Frame FraGPS 
            Caption         =   "GPS"
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
            Height          =   855
            Left            =   120
            TabIndex        =   173
            Top             =   3360
            Width           =   4215
            Begin VB.TextBox txtCostoGPS 
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
               MaxLength       =   6
               TabIndex        =   174
               Tag             =   "RentaGPS"
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Costo por renta GPS"
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
               Index           =   1
               Left            =   90
               TabIndex        =   175
               Top             =   360
               Width           =   2250
            End
         End
         Begin VB.Frame FramPrestamo 
            Caption         =   "% Prestamo"
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
            Height          =   1935
            Left            =   -74880
            TabIndex        =   159
            Top             =   2040
            Width           =   2775
            Begin VB.TextBox txtSemafRojo 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1080
               TabIndex        =   165
               Tag             =   "PrestamoRojo"
               Top             =   1320
               Width           =   975
            End
            Begin VB.TextBox txtSemafAmarillo 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1080
               TabIndex        =   164
               Tag             =   "PrestamoAmarillo"
               Top             =   840
               Width           =   975
            End
            Begin VB.TextBox txtSemafVerde 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1080
               TabIndex        =   163
               Tag             =   "PrestamoVerde"
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lbl8 
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
               Left            =   2160
               TabIndex        =   169
               Top             =   1320
               Width           =   270
            End
            Begin VB.Label lbl7 
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
               Left            =   2160
               TabIndex        =   168
               Top             =   840
               Width           =   270
            End
            Begin VB.Label lbl6 
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
               Left            =   2160
               TabIndex        =   167
               Top             =   360
               Width           =   270
            End
            Begin VB.Label lblRojo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Rojo"
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
               TabIndex        =   162
               Top             =   1320
               Width           =   510
            End
            Begin VB.Label lblAmarillo 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Amarillo"
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
               Left            =   120
               TabIndex        =   161
               Top             =   840
               Width           =   945
            End
            Begin VB.Label lblVerde 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Verde"
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
               Left            =   120
               TabIndex        =   160
               Top             =   360
               Width           =   660
            End
         End
         Begin VB.Frame FramSemaforo 
            Caption         =   "% Enajenados"
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
            Left            =   -74880
            TabIndex        =   155
            Top             =   600
            Width           =   2775
            Begin VB.TextBox txtPorEnajenados 
               Alignment       =   1  'Right Justify
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   1200
               TabIndex        =   157
               Tag             =   "PorEnajenados"
               Top             =   465
               Width           =   855
            End
            Begin VB.Label lblporEnajenado 
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
               Left            =   2160
               TabIndex        =   158
               Top             =   480
               Width           =   270
            End
            Begin VB.Label lblEnajenado 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Enajenado"
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
               TabIndex        =   156
               Top             =   480
               Width           =   1005
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
            Left            =   -69000
            MaxLength       =   10
            TabIndex        =   146
            Tag             =   "PrecioAutos"
            Top             =   1800
            Width           =   1095
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
            Height          =   2910
            Left            =   -74880
            TabIndex        =   136
            Top             =   480
            Width           =   3360
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
               TabIndex        =   38
               Tag             =   "8K"
               Top             =   705
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
               TabIndex        =   40
               Tag             =   "10K"
               Top             =   1065
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
               TabIndex        =   42
               Tag             =   "14K"
               Top             =   1425
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
               TabIndex        =   44
               Tag             =   "18K"
               Top             =   1785
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
               TabIndex        =   39
               Tag             =   "Venta8K"
               Top             =   705
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
               TabIndex        =   45
               Tag             =   "Venta18K"
               Top             =   1785
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
               TabIndex        =   43
               Tag             =   "Venta14K"
               Top             =   1425
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
               TabIndex        =   41
               Tag             =   "Venta10K"
               Top             =   1065
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
               TabIndex        =   47
               Tag             =   "Venta22K"
               Top             =   2145
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
               TabIndex        =   46
               Tag             =   "22K"
               Top             =   2145
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
               TabIndex        =   49
               Tag             =   "Venta24K"
               Top             =   2505
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
               TabIndex        =   48
               Tag             =   "24K"
               Top             =   2505
               Width           =   1095
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
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
               Index           =   0
               Left            =   735
               TabIndex        =   145
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label51 
               AutoSize        =   -1  'True
               BackColor       =   &H00808080&
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
               Index           =   1
               Left            =   2070
               TabIndex        =   144
               Top             =   360
               Width           =   825
            End
            Begin VB.Label Label49 
               BackColor       =   &H00808080&
               Height          =   300
               Left            =   735
               TabIndex        =   143
               Top             =   345
               Width           =   2295
            End
            Begin VB.Label Label38 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "8K:"
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
               TabIndex        =   142
               Top             =   705
               Width           =   375
            End
            Begin VB.Label Label20 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   141
               Top             =   1065
               Width           =   525
            End
            Begin VB.Label Label21 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   140
               Top             =   1425
               Width           =   525
            End
            Begin VB.Label Label53 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   139
               Top             =   1785
               Width           =   525
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "22K:"
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
               TabIndex        =   138
               Top             =   2145
               Width           =   525
            End
            Begin VB.Label Label22 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   137
               Top             =   2505
               Width           =   525
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
            Left            =   -71280
            TabIndex        =   126
            Top             =   480
            Width           =   4245
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
               Left            =   2640
               TabIndex        =   51
               Tag             =   "CalidadB"
               Top             =   360
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
               Left            =   720
               TabIndex        =   50
               Tag             =   "CalidadEx"
               Top             =   360
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
               Left            =   720
               TabIndex        =   52
               Tag             =   "CalidadR"
               Top             =   720
               Width           =   1095
            End
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
               Left            =   2640
               TabIndex        =   53
               Tag             =   "CalidadM"
               Top             =   720
               Width           =   1095
            End
            Begin VB.Label Label44 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   134
               Top             =   727
               Width           =   270
            End
            Begin VB.Label Label42 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   360
               TabIndex        =   133
               Top             =   727
               Width           =   240
            End
            Begin VB.Label Label27 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   2295
               TabIndex        =   132
               Top             =   367
               Width           =   225
            End
            Begin VB.Label Label26 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   131
               Top             =   367
               Width           =   360
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   1860
               TabIndex        =   130
               Top             =   727
               Width           =   270
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   1860
               TabIndex        =   129
               Top             =   367
               Width           =   270
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   3780
               TabIndex        =   128
               Top             =   727
               Width           =   270
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   3780
               TabIndex        =   127
               Top             =   367
               Width           =   270
            End
         End
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
            TabIndex        =   109
            Top             =   3000
            Width           =   4245
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
               TabIndex        =   16
               Tag             =   "PolizaSeguro"
               Top             =   360
               Width           =   1095
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
               TabIndex        =   17
               Tag             =   "FechaExpedicion"
               Top             =   720
               Width           =   1095
            End
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
               TabIndex        =   18
               Tag             =   "Aseguradora"
               Top             =   1080
               Width           =   2280
            End
            Begin DevPowerFlatBttn.FlatBttn cmdExpedicion 
               Height          =   300
               Left            =   3840
               TabIndex        =   110
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
               Picture         =   "frmConfiguracion.frx":28A2
            End
            Begin VB.Label Label33 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   113
               Top             =   360
               Width           =   1140
            End
            Begin VB.Label Label34 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   112
               Top             =   720
               Width           =   2340
            End
            Begin VB.Label Label35 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               TabIndex        =   111
               Top             =   1080
               Width           =   1500
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Apartados"
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
            Left            =   4440
            TabIndex        =   100
            Top             =   360
            Width           =   4245
            Begin VB.TextBox txtVApartado 
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
               MaxLength       =   2
               TabIndex        =   11
               Tag             =   "VenApartados"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtIvafacturacion 
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
               MaxLength       =   2
               TabIndex        =   13
               Tag             =   "IvaVentas"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtEngancheapa 
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
               MaxLength       =   2
               TabIndex        =   12
               Tag             =   "EngancheApartados"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtDiasGraciaApa 
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
               MaxLength       =   2
               TabIndex        =   15
               Tag             =   "DiasGraciaApa"
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtDescuentoAutorizado 
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
               TabIndex        =   14
               Tag             =   "DescuentoVentas"
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label32 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vencimiento Apartado:"
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
               Left            =   120
               TabIndex        =   108
               Top             =   367
               Width           =   2535
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "I.V.A./Ventas:"
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
               Left            =   1050
               TabIndex        =   107
               Top             =   1087
               Width           =   1605
            End
            Begin VB.Label Label36 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Enganche apartado:"
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
               Left            =   450
               TabIndex        =   106
               Top             =   727
               Width           =   2205
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   0
               Left            =   3900
               TabIndex        =   105
               Top             =   1087
               Width           =   270
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   0
               Left            =   3900
               TabIndex        =   104
               Top             =   727
               Width           =   270
            End
            Begin VB.Label Label23 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Dias de Gracia:"
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
               TabIndex        =   103
               Top             =   1807
               Width           =   1695
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   0
               Left            =   3900
               TabIndex        =   102
               Top             =   1447
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label48 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Descuento Autorizado:"
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
               Left            =   135
               TabIndex        =   101
               Top             =   1447
               Width           =   2520
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Tasas Aplicadas"
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
            Height          =   2955
            Left            =   120
            TabIndex        =   83
            Top             =   360
            Width           =   4245
            Begin VB.TextBox txtAlmAnual 
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
               MaxLength       =   6
               TabIndex        =   10
               Tag             =   "AlmAnual"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtIntAnual 
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
               Left            =   2745
               MaxLength       =   6
               TabIndex        =   9
               Tag             =   "IntAnual"
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtPrestamoAvaluoDiamante 
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
               Left            =   2745
               MaxLength       =   8
               TabIndex        =   8
               Tag             =   "PrestamoAvaluoDiamante"
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtAlmacenaje 
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
               Left            =   2745
               MaxLength       =   5
               TabIndex        =   0
               Tag             =   "Almacenaje"
               Top             =   3600
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtSeguro 
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
               Left            =   2745
               MaxLength       =   5
               TabIndex        =   1
               Tag             =   "Seguro"
               Top             =   3960
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtOperacion 
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
               Left            =   2745
               MaxLength       =   6
               TabIndex        =   3
               Tag             =   "Operacion"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtVenta 
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
               Left            =   2745
               MaxLength       =   6
               TabIndex        =   2
               Tag             =   "GtosVenta"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtIVA 
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
               Left            =   2745
               MaxLength       =   5
               TabIndex        =   4
               Tag             =   "IVA"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtPrestamoAvaluo 
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
               Left            =   2745
               MaxLength       =   5
               TabIndex        =   5
               Tag             =   "PrestamoAvaluo"
               Top             =   3600
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtPrestamoAvaluoAutos 
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
               Left            =   2745
               MaxLength       =   5
               TabIndex        =   6
               Tag             =   "PrestamoAvaluoAutos"
               Top             =   3960
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtPrestamoAvaluoElec 
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
               MaxLength       =   5
               TabIndex        =   7
               Tag             =   "PrestamoAvaluoElec"
               Top             =   1440
               Width           =   1095
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   8
               Left            =   3900
               TabIndex        =   152
               Top             =   2535
               Width           =   270
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   7
               Left            =   3900
               TabIndex        =   151
               Top             =   2160
               Width           =   270
            End
            Begin VB.Label Label65 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tasa Almacenaje Anual:"
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
               Left            =   60
               TabIndex        =   150
               Top             =   2520
               Width           =   2640
            End
            Begin VB.Label Label59 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Tasa Interés Anual:"
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
               Left            =   525
               TabIndex        =   149
               Top             =   2160
               Width           =   2175
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   6
               Left            =   3900
               TabIndex        =   125
               Top             =   1800
               Width           =   270
            End
            Begin VB.Label Label56 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Préstamo/Avalúo Dia.:"
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
               TabIndex        =   124
               Top             =   1800
               Width           =   2550
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   3
               Left            =   3900
               TabIndex        =   123
               Top             =   3960
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   3
               Left            =   3900
               TabIndex        =   122
               Top             =   1440
               Width           =   270
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   2
               Left            =   3900
               TabIndex        =   121
               Top             =   3600
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   2
               Left            =   3900
               TabIndex        =   120
               Top             =   720
               Width           =   270
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   2
               Left            =   3900
               TabIndex        =   119
               Top             =   1080
               Width           =   270
            End
            Begin VB.Label Label45 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   1
               Left            =   3900
               TabIndex        =   118
               Top             =   360
               Width           =   270
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   1
               Left            =   3900
               TabIndex        =   117
               Top             =   3600
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Index           =   1
               Left            =   3900
               TabIndex        =   116
               Top             =   3960
               Visible         =   0   'False
               Width           =   270
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Almacenaje/Préstamo:"
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
               Left            =   135
               TabIndex        =   99
               Top             =   3600
               Visible         =   0   'False
               Width           =   2565
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Seguro/Préstamo:"
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
               Left            =   630
               TabIndex        =   98
               Top             =   3960
               Visible         =   0   'False
               Width           =   2070
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Gastos de Operación:"
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
               Left            =   315
               TabIndex        =   97
               Top             =   720
               Width           =   2385
            End
            Begin VB.Label Label11 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Gastos de venta:"
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
               Left            =   840
               TabIndex        =   96
               Top             =   360
               Width           =   1860
            End
            Begin VB.Label Label18 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "I.V.A./Intereses:"
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
               Left            =   780
               TabIndex        =   95
               Top             =   1080
               Width           =   1920
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   4245
               TabIndex        =   94
               Top             =   3180
               Width           =   270
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   4245
               TabIndex        =   93
               Top             =   3540
               Width           =   270
            End
            Begin VB.Label Label4 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   4245
               TabIndex        =   92
               Top             =   3900
               Width           =   270
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   4245
               TabIndex        =   91
               Top             =   4620
               Width           =   270
            End
            Begin VB.Label Label30 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   4245
               TabIndex        =   90
               Top             =   4260
               Width           =   270
            End
            Begin VB.Label Label43 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   4245
               TabIndex        =   89
               Top             =   4980
               Width           =   270
            End
            Begin VB.Label Label54 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Préstamo/Avalúo:"
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
               Left            =   660
               TabIndex        =   88
               Top             =   3600
               Visible         =   0   'False
               Width           =   2040
            End
            Begin VB.Label Label63 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Préstamo/Avalúo Auto:"
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
               Left            =   90
               TabIndex        =   87
               Top             =   3960
               Visible         =   0   'False
               Width           =   2610
            End
            Begin VB.Label Label64 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   4245
               TabIndex        =   86
               Top             =   5340
               Width           =   270
            End
            Begin VB.Label Label14 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Préstamo/Avalúo Elec.:"
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
               Left            =   60
               TabIndex        =   85
               Top             =   1440
               Width           =   2640
            End
            Begin VB.Label Label28 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
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
               Left            =   4245
               TabIndex        =   84
               Top             =   5700
               Width           =   270
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Otras Variables"
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
            Height          =   4275
            Left            =   -74880
            TabIndex        =   63
            Top             =   360
            Width           =   8535
            Begin VB.TextBox txtDiasgraciaparaautos 
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
               Left            =   3090
               TabIndex        =   177
               Tag             =   "DiasGraciaAuto"
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox txtBonificacion 
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
               Left            =   7050
               TabIndex        =   26
               Tag             =   "DescuentoPagosFijos"
               Top             =   6720
               Width           =   975
            End
            Begin VB.TextBox TxtHorario 
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
               Height          =   600
               Left            =   2400
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   37
               Tag             =   "HorarioSucursal"
               Top             =   3600
               Width           =   5865
            End
            Begin VB.TextBox txtCodProfeco 
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
               Left            =   2400
               MaxLength       =   50
               TabIndex        =   29
               Tag             =   "codprofeco"
               Top             =   4680
               Visible         =   0   'False
               Width           =   1800
            End
            Begin VB.TextBox txtCat 
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
               Left            =   2850
               MaxLength       =   6
               TabIndex        =   27
               Tag             =   "Cat"
               Top             =   6360
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtEnajenacion 
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
               Left            =   3090
               MaxLength       =   2
               TabIndex        =   21
               Tag             =   "DiasEnajenacion"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtDiasGracia 
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
               Left            =   3090
               MaxLength       =   2
               TabIndex        =   22
               Tag             =   "DiasGracia"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtBoletaExtraviada 
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
               Left            =   3090
               MaxLength       =   8
               TabIndex        =   24
               Tag             =   "ImportePerdida"
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtAutorizacion 
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
               Left            =   3240
               MaxLength       =   8
               TabIndex        =   65
               Tag             =   "ImporteAutorizacion"
               Top             =   7575
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtPagoMinimo 
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
               Left            =   3090
               MaxLength       =   8
               TabIndex        =   25
               Tag             =   "PagoMinimo"
               Top             =   2520
               Width           =   1095
            End
            Begin VB.TextBox txtNegociacion 
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
               Left            =   3090
               MaxLength       =   5
               TabIndex        =   23
               Tag             =   "Negociacion"
               Top             =   1800
               Width           =   1095
            End
            Begin VB.TextBox txtGerente 
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
               Left            =   1080
               MaxLength       =   80
               TabIndex        =   28
               Tag             =   "Gerente"
               Top             =   2880
               Width           =   3120
            End
            Begin VB.TextBox txtNotas 
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
               Height          =   570
               Left            =   5280
               MaxLength       =   250
               MultiLine       =   -1  'True
               TabIndex        =   36
               Tag             =   "Notas"
               Top             =   2520
               Width           =   3015
            End
            Begin VB.TextBox txtLimiteInferior 
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
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   32
               Tag             =   "LimiteInferior"
               Top             =   1080
               Width           =   1095
            End
            Begin VB.TextBox txtLimiteSuperior 
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
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   33
               Tag             =   "LimiteSuperior"
               Top             =   1440
               Width           =   1095
            End
            Begin VB.TextBox txtVenAlmoneda 
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
               Left            =   3255
               MaxLength       =   5
               TabIndex        =   64
               Tag             =   "VenAlmoneda"
               Top             =   7230
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtLimite1 
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
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   30
               Tag             =   "Limite1"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtLimite2 
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
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   31
               Tag             =   "Limite2"
               Top             =   720
               Width           =   1095
            End
            Begin VB.TextBox txtMinimoAbonar 
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
               Left            =   3090
               MaxLength       =   8
               TabIndex        =   20
               Tag             =   "AbonoMinimo"
               Top             =   360
               Width           =   1095
            End
            Begin VB.TextBox txtDiasPenaliza 
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
               Left            =   2880
               MaxLength       =   5
               TabIndex        =   19
               Tag             =   "DiasPenaliza"
               Top             =   6720
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtLimiteSuperiorAutos 
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
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   35
               Tag             =   "LimiteSuperiorAutos"
               Top             =   2160
               Width           =   1095
            End
            Begin VB.TextBox txtLimiteInferiorAutos 
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
               Left            =   7200
               MaxLength       =   10
               TabIndex        =   34
               Tag             =   "LimiteInferiorAutos"
               Top             =   1800
               Width           =   1095
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dias de Gracia para Autos:"
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
               Left            =   120
               TabIndex        =   176
               Top             =   1440
               Width           =   2955
            End
            Begin VB.Label lblFechaAdhesion 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Fecha Adhesion:"
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
               Left            =   120
               TabIndex        =   172
               Top             =   4680
               Visible         =   0   'False
               Width           =   2130
            End
            Begin VB.Label lblporBonificacion 
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
               Left            =   8040
               TabIndex        =   171
               Top             =   6720
               Width           =   270
            End
            Begin VB.Label Label68 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bonificación pagos fijos:"
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
               Left            =   4260
               TabIndex        =   170
               Top             =   6720
               Width           =   2730
            End
            Begin VB.Label Label67 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Horario Sucursal:"
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
               Left            =   120
               TabIndex        =   154
               Top             =   3600
               Width           =   2055
            End
            Begin VB.Label Label66 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Contrato Adhesion:"
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
               Left            =   120
               TabIndex        =   153
               Top             =   4680
               Visible         =   0   'False
               Width           =   2130
            End
            Begin VB.Label Label40 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "CAT:"
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
               Left            =   2280
               TabIndex        =   148
               Top             =   6360
               Visible         =   0   'False
               Width           =   510
            End
            Begin VB.Label Label29 
               Alignment       =   1  'Right Justify
               Caption         =   "Dias de enajenación:"
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
               Left            =   30
               TabIndex        =   82
               Top             =   720
               Width           =   3015
            End
            Begin VB.Label Label47 
               Alignment       =   1  'Right Justify
               Caption         =   "Dias de Gracia:"
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
               Left            =   30
               TabIndex        =   81
               Top             =   1080
               Width           =   3015
            End
            Begin VB.Label Label37 
               Alignment       =   1  'Right Justify
               Caption         =   "Importe boleta extraviada:"
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
               Left            =   30
               TabIndex        =   80
               Top             =   2160
               Width           =   3015
            End
            Begin VB.Label Label31 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Importe autorización:"
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
               Left            =   720
               TabIndex        =   79
               Top             =   7560
               Visible         =   0   'False
               Width           =   2430
            End
            Begin VB.Label Label24 
               Alignment       =   1  'Right Justify
               Caption         =   "Importe mínimo a cobrar:"
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
               Left            =   30
               TabIndex        =   78
               Top             =   2520
               Width           =   3015
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Margen de negociación:"
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
               Left            =   30
               TabIndex        =   77
               Top             =   1800
               Width           =   3015
            End
            Begin VB.Label Label25 
               AutoSize        =   -1  'True
               Caption         =   "Gerente:"
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
               TabIndex        =   76
               Top             =   2880
               Width           =   975
            End
            Begin VB.Label Label12 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Notas:"
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
               Left            =   4440
               TabIndex        =   75
               Top             =   2520
               Width           =   735
            End
            Begin VB.Label Label55 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "P. Límite inferior:"
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
               Left            =   5100
               TabIndex        =   74
               Top             =   1080
               Width           =   1980
            End
            Begin VB.Label Label58 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "P. Límite superior:"
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
               Left            =   5010
               TabIndex        =   73
               Top             =   1440
               Width           =   2070
            End
            Begin VB.Label Label60 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Vencimiento Almoneda:"
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
               Left            =   570
               TabIndex        =   72
               Top             =   7230
               Visible         =   0   'False
               Width           =   2610
            End
            Begin VB.Label Label61 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Límite contrato 1:"
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
               TabIndex        =   71
               Top             =   367
               Width           =   1995
            End
            Begin VB.Label Label62 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Límite contrato 2:"
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
               TabIndex        =   70
               Top             =   727
               Width           =   1995
            End
            Begin VB.Label Label39 
               Alignment       =   1  'Right Justify
               Caption         =   "Importe mínimo abono:"
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
               Left            =   0
               TabIndex        =   69
               Top             =   360
               Width           =   3015
            End
            Begin VB.Label Label41 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "Días mínimos a cobrar:"
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
               Left            =   255
               TabIndex        =   68
               Top             =   6720
               Visible         =   0   'False
               Width           =   2550
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "P. Límite superior Auto:"
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
               Left            =   4440
               TabIndex        =   67
               Top             =   2160
               Width           =   2640
            End
            Begin VB.Label Label52 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               Caption         =   "P. Límite inferior Auto:"
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
               Left            =   4530
               TabIndex        =   66
               Top             =   1800
               Width           =   2550
            End
         End
         Begin DevPowerFlatBttn.FlatBttn cmdRecalcular 
            Height          =   375
            Left            =   -69720
            TabIndex        =   135
            Top             =   2640
            Width           =   1665
            _ExtentX        =   2937
            _ExtentY        =   661
            AlignCaption    =   4
            AlignPicture    =   2
            AutoSize        =   0   'False
            Caption         =   "     &Recalcular precios"
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
            Picture         =   "frmConfiguracion.frx":29B7
            PictureDisabled =   "frmConfiguracion.frx":2D21
         End
         Begin VB.Label lbl5 
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
            Left            =   -72840
            TabIndex        =   166
            Top             =   2520
            Width           =   270
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
            Left            =   -71280
            TabIndex        =   147
            Top             =   1800
            Width           =   2235
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
         TabIndex        =   55
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
         TabIndex        =   54
         Tag             =   "Datos"
         Top             =   8355
         Visible         =   0   'False
         Width           =   1095
      End
      Begin DevPowerFlatBttn.FlatBttn cmdDesde 
         Height          =   300
         Left            =   3270
         TabIndex        =   59
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
         Picture         =   "frmConfiguracion.frx":2E7B
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
         TabIndex        =   61
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
         Left            =   135
         TabIndex        =   60
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
      TabIndex        =   56
      Top             =   15240
      Visible         =   0   'False
      Width           =   855
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   8025
      TabIndex        =   114
      Top             =   5160
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
      Picture         =   "frmConfiguracion.frx":2F90
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   6885
      TabIndex        =   115
      Top             =   5160
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
      Picture         =   "frmConfiguracion.frx":34E2
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
      TabIndex        =   57
      Top             =   15720
      Visible         =   0   'False
      Width           =   2715
   End
End
Attribute VB_Name = "frmConfiguracion"
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

Dim Fl() As cFlatControl

'''''Private Sub cmbImpresoras_GotFocus()
'''''    Cambiar_Color True, cmbImpresoras
'''''End Sub
'''''
'''''Private Sub cmbImpresoras_KeyPress(KeyAscii As Integer)
'''''    Pasar_Foco KeyAscii
'''''End Sub
'''''
'''''Private Sub cmbImpresoras_LostFocus()
'''''    Cambiar_Color False, cmbImpresoras
'''''End Sub

Private Sub cmdAceptar_Click()

    If Valida Then
        
        Grabar_Configuracion
        MsgBox "Configuración guardada con éxito !!", vbInformation, "Parámetros"
    End If

End Sub

'Grabamos la configuracion de los parametros
Private Sub Grabar_Configuracion()
Dim txt As Object, Sql As String, Caracter As String, Tasa As Double, Vencimiento As Integer, VencimientoAlmoneda As Integer, crPrestamoLimite As Double, strValor As String

On Error GoTo Error

    For Each txt In Me.Controls

        If TypeOf txt Is TextBox Or TypeOf txt Is ComboBox Then
            
            If txt.Tag <> "" And TypeOf txt Is TextBox Then
                
                Caracter = ""
                If txt.Tag = "Datos" Or txt.Tag = "FechaExpedicion" Or txt.Tag = "Aseguradora" Or txt.Tag = "Gerente" Or txt.Tag = "PolizaSeguro" Or txt.Tag = "Notas" Or txt.Tag = "codprofeco" Or txt.Tag = "FechaProfeco" Or txt.Tag = "HorarioSucursal" Then Caracter = "'"
                
                If Caracter = "" Then
                    
                    strValor = ConvMoneda(txt.text)
                Else
                    
                    strValor = txt.text
                End If
                
                Sql = Sql & txt.Tag & "=" & Caracter & IIf(txt.Tag = "Datos" Or txt.Tag = "FechaExpedicion", Format(txt.text, "YYYY/MM/DD"), txt.text) & Caracter & ","
            End If
        
        End If

    Next

    Sql = Mid(Sql, 1, Len(Sql) - 1)
    dbDatos.Execute "UPDATE parametros SET " & Sql
    
Error:
    Maneja_Error Err
End Sub

Private Sub cmdExpedicion_Click()
    txtFechaexpedicion.text = frmCalendario.Fecha(txtFechaexpedicion.text)
End Sub

Private Sub cmdRecalcular_Click()
Dim rcParametros As New ADODB.Recordset
Dim PrecioVenta As Double, AvaluoDiam As Double

On Error GoTo Error
    
    frmPasswords.ConexSuc = 0
    frmPasswords.PrecioVitrina = 0
    frmPasswords.Cancel = 0
    frmPasswords.Ventas = 0
    frmPasswords.ModificaCorte = 0
    frmPasswords.HacerCorte = 0
    frmPasswords.InteresDesempeño = 0
    frmPasswords.InteresRefrendo = 0
    frmPasswords.ModificaPrecio = 0
    frmPasswords.AutorizaPrestamo = 0
    frmPasswords.DescuentoVentas = 0
    frmPasswords.Vencido = 0
    frmPasswords.CancelaCierre = 0
    frmPasswords.RecalculoPrecios = 1
    
    If frmPasswords.Password(GERENTE, 1) Then
        
        If MsgBox("Se recalcularán los precios de inventario desea continuar ??", vbExclamation + vbYesNo + vbDefaultButton2, "Parámetros") = vbYes Then
            
            'Detalles Entrada Inventario
            rcParametros.Open "SELECT kilatajes.Descripcion AS Kilataje,detallesentradainventario.Kilates FROM detallesentradainventario LEFT JOIN kilatajes ON detallesentradainventario.Kilates=kilatajes.ID WHERE detallesentradainventario.Cantidad>0 AND detallesentradainventario.Kilates>0 GROUP BY detallesentradainventario.Kilates ORDER BY detallesentradainventario.Kilates", dbDatos, adOpenForwardOnly, adLockOptimistic
            If Not rcParametros.BOF And Not rcParametros.EOF Then
                
                AvaluoDiam = CDbl(Regresa_Valor_BD("PrestamoAvaluoDiamante"))

                While Not rcParametros.EOF
                    
                    PrecioVenta = Regresa_Valor_BD("GtosVenta") / 100
                    dbDatos.Execute "UPDATE detallesentradainventario SET PrecioVitrina=(Costo * (1+" & PrecioVenta & ")) WHERE Kilates=" & rcParametros!Kilates & " AND Cantidad>0"
                
                rcParametros.MoveNext
                Wend
                
                MsgBox "Actualización aplicada con éxito !!", vbInformation, "Parámetros"
            
            End If
            rcParametros.Close
            Set rcParametros = Nothing
            
        End If
    
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcParametros = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

'inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    CentrarForm Me, frmMDI
    Frame2.BorderStyle = 0
    Cargar_Configuracion
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

'Leemos los datos del archivo ini
Private Sub Cargar_Configuracion()
Dim txt As Object, i As Integer
Dim rc As New ADODB.Recordset

On Error GoTo Error
    
    rc.Open "SELECT * FROM parametros", dbDatos, adOpenForwardOnly, adLockOptimistic
  
    For Each txt In Me.Controls

        If TypeOf txt Is TextBox Then
            
            If txt.Tag <> "" Then
                
                If txt.Tag = "Datos" Or txt.Tag = "FechaExpedicion" Then
                    
                    txt.text = Format(rc.Fields(txt.Tag) & "", "DD/MM/YY")
                Else
                    
                    txt.text = rc.Fields(txt.Tag) & ""
                End If
            
            End If
        
        End If

    Next
    
'''''    'Cargo las impresoras y pongo la impresora por Default
'''''    CargaImpresoras
'''''    cmbImpresoras.ListIndex = ComboInformacion(cmbImpresoras, 0, rc!ImpresoraDefault)

    If Trim(TxtHorario.text) = "" Then
        TxtHorario.text = "EL HORARIO DE SERVICIO AL PÚBLICO DE ESTE ESTABLECIMIENTO ES DE LUNES A VIERNES DE 9:30 A 19:00 HRS Y SABADOS DE 9:30 A 16:00 HRS."
    End If
    
    rc.Close
    Set rc = Nothing
    txtFolio.text = Regresa_NumContrato(False, 1)
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rc = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txt10K_GotFocus()
    Seleccionar_Texto txt10K
    Cambiar_Color True, txt10K
End Sub

Private Sub txt10K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txt10K_LostFocus()
    Cambiar_Color False, txt10K
End Sub

Private Sub txt14K_GotFocus()
    Seleccionar_Texto txt14K
    Cambiar_Color True, txt14K
End Sub

Private Sub txt14K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txt14K_LostFocus()
    Cambiar_Color False, txt14K
End Sub

Private Sub txt18K_GotFocus()
    Seleccionar_Texto txt18K
    Cambiar_Color True, txt18K
End Sub

Private Sub txt18K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txt18K_LostFocus()
    Cambiar_Color False, txt18K
End Sub

Private Sub txt21K_GotFocus()
    Seleccionar_Texto txt21K
    Cambiar_Color True, txt21K
End Sub

Private Sub txt21K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txt21K_LostFocus()
    Cambiar_Color False, txt21K
End Sub

Private Sub txt24K_GotFocus()
    Seleccionar_Texto txt24K
    Cambiar_Color True, txt24K
End Sub

Private Sub txt24K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txt24K_LostFocus()
    Cambiar_Color False, txt24K
End Sub

Private Sub txt8K_GotFocus()
    Seleccionar_Texto txt8K
    Cambiar_Color True, txt8K
End Sub

Private Sub txt8K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txt8K_LostFocus()
    Cambiar_Color False, txt8K
End Sub

Private Sub txtAlmacenaje_GotFocus()
    Seleccionar_Texto txtAlmacenaje
    Cambiar_Color True, txtAlmacenaje
End Sub

Private Sub txtAlmacenaje_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAlmacenaje_LostFocus()
    Cambiar_Color False, txtAlmacenaje
End Sub

Private Sub txtAlmAnual_GotFocus()
    Seleccionar_Texto txtAlmAnual
    Cambiar_Color True, txtAlmAnual
End Sub

Private Sub txtAlmAnual_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAlmAnual_LostFocus()
    Cambiar_Color False, txtAlmAnual
End Sub

Private Sub txtAseguradora_GotFocus()
    Seleccionar_Texto txtAseguradora
    Cambiar_Color True, txtAseguradora
End Sub

Private Sub txtAseguradora_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAseguradora_LostFocus()
    Cambiar_Color False, txtAseguradora
End Sub

Private Sub txtAutorizacion_GotFocus()
    Seleccionar_Texto txtAutorizacion
    Cambiar_Color True, txtAutorizacion
End Sub

Private Sub txtAutorizacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAutorizacion_LostFocus()
    Cambiar_Color False, txtAutorizacion
End Sub

Private Sub txtB_GotFocus()
    Seleccionar_Texto txtB
    Cambiar_Color True, txtB
End Sub

Private Sub txtB_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtB_LostFocus()
    Cambiar_Color False, txtB
End Sub

Private Sub txtBoletaExtraviada_GotFocus()
    Seleccionar_Texto txtBoletaExtraviada
    Cambiar_Color True, txtBoletaExtraviada
End Sub

Private Sub txtBoletaExtraviada_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBoletaExtraviada_LostFocus()
    Cambiar_Color False, txtBoletaExtraviada
End Sub

Private Sub txtBonificacion_GotFocus()
    Seleccionar_Texto txtBonificacion
    Cambiar_Color True, txtBonificacion
End Sub

Private Sub txtBonificacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBonificacion_LostFocus()
    Cambiar_Color False, txtBonificacion
End Sub

Private Sub txtCat_GotFocus()
    Seleccionar_Texto txtCat
    Cambiar_Color True, txtCat
End Sub

Private Sub txtCat_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCat_LostFocus()
    Cambiar_Color False, txtCat
End Sub

Private Sub txtCodProfeco_GotFocus()
    Seleccionar_Texto txtCodProfeco
    Cambiar_Color True, txtCodProfeco
End Sub

Private Sub txtCodProfeco_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCodProfeco_LostFocus()
    Cambiar_Color False, txtCodProfeco
End Sub

Private Sub txtDiasPenaliza_GotFocus()
    Seleccionar_Texto txtDiasPenaliza
    Cambiar_Color True, txtDiasPenaliza
End Sub

Private Sub txtDiasPenaliza_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDiasPenaliza_LostFocus()
    Cambiar_Color False, txtDiasPenaliza
End Sub

Private Sub txtIntAnual_GotFocus()
    Seleccionar_Texto txtIntAnual
    Cambiar_Color True, txtIntAnual
End Sub

Private Sub txtIntAnual_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtIntAnual_LostFocus()
    Cambiar_Color False, txtIntAnual
End Sub

Private Sub txtLimite1_GotFocus()
    Seleccionar_Texto txtLimite1
    Cambiar_Color True, txtLimite1
End Sub

Private Sub txtLimite1_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtLimite1_LostFocus()
    Cambiar_Color False, txtLimite1
End Sub

Private Sub txtLimite2_GotFocus()
    Seleccionar_Texto txtLimite2
    Cambiar_Color True, txtLimite2
End Sub

Private Sub txtLimite2_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtLimite2_LostFocus()
    Cambiar_Color False, txtLimite2
End Sub

Private Sub txtLimiteInferior_GotFocus()
    Seleccionar_Texto txtLimiteInferior
    Cambiar_Color True, txtLimiteInferior
End Sub

Private Sub txtLimiteInferior_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtLimiteInferior_LostFocus()
    Cambiar_Color False, txtLimiteInferior
End Sub

Private Sub txtLimiteInferiorAutos_GotFocus()
    Seleccionar_Texto txtLimiteInferiorAutos
    Cambiar_Color True, txtLimiteInferiorAutos
End Sub

Private Sub txtLimiteInferiorAutos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtLimiteInferiorAutos_LostFocus()
    Cambiar_Color False, txtLimiteInferiorAutos
End Sub

Private Sub txtLimiteSuperior_GotFocus()
    Seleccionar_Texto txtLimiteSuperior
    Cambiar_Color True, txtLimiteSuperior
End Sub

Private Sub txtLimiteSuperior_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtLimiteSuperior_LostFocus()
    Cambiar_Color False, txtLimiteSuperior
End Sub

Private Sub txtLimiteSuperiorAutos_GotFocus()
    Seleccionar_Texto txtLimiteSuperiorAutos
    Cambiar_Color True, txtLimiteSuperiorAutos
End Sub

Private Sub txtLimiteSuperiorAutos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtLimiteSuperiorAutos_LostFocus()
    Cambiar_Color False, txtLimiteSuperiorAutos
End Sub

Private Sub txtMinimoAbonar_GotFocus()
    Seleccionar_Texto txtMinimoAbonar
    Cambiar_Color True, txtMinimoAbonar
End Sub

Private Sub txtMinimoAbonar_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMinimoAbonar_LostFocus()
    Cambiar_Color False, txtMinimoAbonar
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

Private Sub txtOperacion_GotFocus()
    Seleccionar_Texto txtOperacion
    Cambiar_Color True, txtOperacion
End Sub

Private Sub txtOperacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtOperacion_LostFocus()
    Cambiar_Color False, txtOperacion
End Sub

Private Sub txtDescuentoAutorizado_GotFocus()
    Seleccionar_Texto txtDescuentoAutorizado
    Cambiar_Color True, txtDescuentoAutorizado
End Sub

Private Sub txtDescuentoAutorizado_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescuentoAutorizado_LostFocus()
    Cambiar_Color False, txtDescuentoAutorizado
End Sub

Private Sub txtDiasGracia_GotFocus()
    Seleccionar_Texto txtDiasGracia
    Cambiar_Color True, txtDiasGracia
End Sub

Private Sub txtDiasGracia_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDiasGracia_LostFocus()
    Cambiar_Color False, txtDiasGracia
End Sub

Private Sub txtDiasGraciaApa_GotFocus()
    Seleccionar_Texto txtDiasGraciaApa
    Cambiar_Color True, txtDiasGraciaApa
End Sub

Private Sub txtDiasGraciaApa_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDiasGraciaApa_LostFocus()
    Cambiar_Color False, txtDiasGraciaApa
End Sub

Private Sub txtEngancheapa_GotFocus()
    Seleccionar_Texto txtEngancheapa
    Cambiar_Color True, txtEngancheapa
End Sub

Private Sub txtEngancheapa_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEngancheapa_LostFocus()
    Cambiar_Color False, txtEngancheapa
End Sub

Private Sub txtEx_GotFocus()
    Seleccionar_Texto txtEx
    Cambiar_Color True, txtEx
End Sub

Private Sub txtEx_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEx_LostFocus()
    Cambiar_Color False, txtEx
End Sub

Private Sub txtFechaexpedicion_GotFocus()
    Seleccionar_Texto txtFechaexpedicion
    Cambiar_Color True, txtFechaexpedicion
End Sub

Private Sub txtFechaexpedicion_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtFechaexpedicion_LostFocus()
    Cambiar_Color False, txtFechaexpedicion
End Sub

Private Sub txtFolio_GotFocus()
    Seleccionar_Texto txtFolio
    Cambiar_Color True, txtFolio
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFolio_LostFocus()
    Cambiar_Color False, txtFolio
End Sub

Private Sub txtGerente_GotFocus()
    Seleccionar_Texto txtGerente
    Cambiar_Color True, txtGerente
End Sub

Private Sub txtGerente_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtGerente_LostFocus()
    Cambiar_Color False, txtGerente
End Sub

Private Sub txtIvafacturacion_GotFocus()
    Seleccionar_Texto txtIvafacturacion
    Cambiar_Color True, txtIvafacturacion
End Sub

Private Sub txtIvafacturacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtIvafacturacion_LostFocus()
    Cambiar_Color False, txtIvafacturacion
End Sub

Private Sub txtEnajenacion_GotFocus()
    Seleccionar_Texto txtEnajenacion
    Cambiar_Color True, txtEnajenacion
End Sub

Private Sub txtEnajenacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtEnajenacion_LostFocus()
    Cambiar_Color False, txtEnajenacion
End Sub

Private Sub txtIva_GotFocus()
    Seleccionar_Texto txtIva
    Cambiar_Color True, txtIva
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtIva_LostFocus()
    Cambiar_Color False, txtIva
End Sub

Private Sub txtM_GotFocus()
    Seleccionar_Texto txtM
    Cambiar_Color True, txtM
End Sub

Private Sub txtM_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtM_LostFocus()
    Cambiar_Color False, txtM
End Sub

Private Sub txtNegociacion_GotFocus()
    Seleccionar_Texto txtNegociacion
    Cambiar_Color True, txtNegociacion
End Sub

Private Sub txtNegociacion_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNegociacion_LostFocus()
    Cambiar_Color False, txtNegociacion
End Sub

Private Sub txtNoSucursal_GotFocus()
    Seleccionar_Texto txtNoSucursal
    Cambiar_Color True, txtNoSucursal
End Sub

Private Sub txtNoSucursal_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNoSucursal_LostFocus()
    Cambiar_Color False, txtNoSucursal
End Sub

Private Sub txtPagoMinimo_GotFocus()
    Seleccionar_Texto txtPagoMinimo
    Cambiar_Color True, txtPagoMinimo
End Sub

Private Sub txtPagoMinimo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPagoMinimo_LostFocus()
    Cambiar_Color False, txtPagoMinimo
End Sub

Private Sub txtPolizano_GotFocus()
    Seleccionar_Texto txtPolizano
    Cambiar_Color True, txtPolizano
End Sub

Private Sub txtPolizano_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPolizano_LostFocus()
    Cambiar_Color False, txtPolizano
End Sub

Private Sub txtPorEnajenados_GotFocus()
    Seleccionar_Texto txtPorEnajenados
    Cambiar_Color True, txtPorEnajenados
End Sub

Private Sub txtPorEnajenados_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPorEnajenados_LostFocus()
     Cambiar_Color False, txtPorEnajenados
End Sub

Private Sub txtPrecioAutos_GotFocus()
    Seleccionar_Texto txtPrecioAutos
    Cambiar_Color True, txtPrecioAutos
End Sub

Private Sub txtPrecioAutos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrecioAutos_LostFocus()
    Cambiar_Color False, txtPrecioAutos
End Sub

Private Sub txtPrestamoAvaluo_GotFocus()
    Seleccionar_Texto txtPrestamoAvaluo
    Cambiar_Color True, txtPrestamoAvaluo
End Sub

Private Sub txtPrestamoAvaluo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamoAvaluo_LostFocus()
    Cambiar_Color False, txtPrestamoAvaluo
End Sub

Private Sub txtPrestamoAvaluoAutos_GotFocus()
    Seleccionar_Texto txtPrestamoAvaluoAutos
    Cambiar_Color True, txtPrestamoAvaluoAutos
End Sub

Private Sub txtPrestamoAvaluoAutos_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamoAvaluoAutos_LostFocus()
    Cambiar_Color False, txtPrestamoAvaluoAutos
End Sub

Private Sub txtPrestamoAvaluoDiamante_GotFocus()
    Seleccionar_Texto txtPrestamoAvaluoDiamante
    Cambiar_Color True, txtPrestamoAvaluoDiamante
End Sub

Private Sub txtPrestamoAvaluoDiamante_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamoAvaluoDiamante_LostFocus()
    Cambiar_Color False, txtPrestamoAvaluoDiamante
End Sub

Private Sub txtPrestamoAvaluoElec_GotFocus()
    Seleccionar_Texto txtPrestamoAvaluoElec
    Cambiar_Color True, txtPrestamoAvaluoElec
End Sub

Private Sub txtPrestamoAvaluoElec_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamoAvaluoElec_LostFocus()
    Cambiar_Color False, txtPrestamoAvaluoElec
End Sub

Private Sub txtR_GotFocus()
    Seleccionar_Texto txtR
    Cambiar_Color True, txtR
End Sub

Private Sub txtR_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtR_LostFocus()
    Cambiar_Color False, txtR
End Sub

Private Sub txtSeguro_GotFocus()
    Seleccionar_Texto txtSeguro
    Cambiar_Color True, txtSeguro
End Sub

Private Sub txtSeguro_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtSeguro_LostFocus()
    Cambiar_Color False, txtSeguro
End Sub

Private Sub txtSemafAmarillo_GotFocus()
    Seleccionar_Texto txtSemafAmarillo
    Cambiar_Color True, txtSemafAmarillo
End Sub

Private Sub txtSemafAmarillo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtSemafAmarillo_LostFocus()
     Cambiar_Color False, txtSemafAmarillo
End Sub

Private Sub txtSemafRojo_GotFocus()
Seleccionar_Texto txtSemafRojo
    Cambiar_Color True, txtSemafRojo
End Sub

Private Sub txtSemafRojo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtSemafRojo_LostFocus()
    Cambiar_Color False, txtSemafRojo
End Sub

Private Sub txtSemafVerde_GotFocus()
    Seleccionar_Texto txtSemafVerde
    Cambiar_Color True, txtSemafVerde
End Sub

Private Sub txtSemafVerde_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtSemafVerde_LostFocus()
     Cambiar_Color False, txtSemafVerde
End Sub

Private Sub txtVen10K_GotFocus()
    Seleccionar_Texto txtVen10K
    Cambiar_Color True, txtVen10K
End Sub

Private Sub txtVen10K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVen10K_LostFocus()
    Cambiar_Color False, txtVen10K
End Sub

Private Sub txtVen14K_GotFocus()
    Seleccionar_Texto txtVen14K
    Cambiar_Color True, txtVen14K
End Sub

Private Sub txtVen14K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVen14K_LostFocus()
    Cambiar_Color False, txtVen14K
End Sub

Private Sub txtVen18K_GotFocus()
    Seleccionar_Texto txtVen18K
    Cambiar_Color True, txtVen18K
End Sub

Private Sub txtVen18K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVen18K_LostFocus()
    Cambiar_Color False, txtVen18K
End Sub

Private Sub txtVen8K_GotFocus()
    Seleccionar_Texto txtVen8K
    Cambiar_Color True, txtVen8K
End Sub

Private Sub txtVen8K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVen8K_LostFocus()
    Cambiar_Color False, txtVen8K
End Sub

Private Sub txtVenAlmoneda_GotFocus()
    Seleccionar_Texto txtVenAlmoneda
    Cambiar_Color True, txtVenAlmoneda
End Sub

Private Sub txtVenAlmoneda_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVenAlmoneda_LostFocus()
    Cambiar_Color False, txtVenAlmoneda
End Sub

Private Sub txtVenta_GotFocus()
    Seleccionar_Texto txtVenta
    Cambiar_Color True, txtVenta
End Sub

Private Sub txtVenta_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVenta_LostFocus()
    Cambiar_Color False, txtVenta
End Sub

Private Sub txtVApartado_GotFocus()
    Seleccionar_Texto txtVApartado
    Cambiar_Color True, txtVApartado
End Sub

Private Sub txtVApartado_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVApartado_LostFocus()
    Cambiar_Color False, txtVApartado
End Sub

Function Valida() As Boolean
    
    Valida = True

    If txtPrestamoAvaluoAutos.text = "" Then
        MsgBox "Introduzca el Porcentaje del Avalúo sobre el Préstamo para Automoviles !!", vbCritical
        Valida = False
        txtPrestamoAvaluoAutos.SetFocus
        Exit Function
    End If
    
    If txtPrestamoAvaluo.text = "" Then
        MsgBox "Introduzca el Porcentaje del Avalúo sobre el Préstamo para Joyería !!", vbCritical
        Valida = False
        txtPrestamoAvaluo.SetFocus
        Exit Function
    End If
    
    If txtAlmacenaje.text = "" Then
        MsgBox "Introduzca el Porcentaje de Almacenaje !!", vbCritical
        Valida = False
        txtAlmacenaje.SetFocus
        Exit Function
    End If

    If txtSeguro.text = "" Then
        MsgBox "Introduzca el Porcentaje de Seguro !!", vbCritical
        Valida = False
        txtSeguro.SetFocus
        Exit Function
    End If

    If txtIva.text = "" Then
        MsgBox "Introduzca el Porcentaje de IVA !!", vbCritical
        Valida = False
        txtIva.SetFocus
        Exit Function
    End If

    If txtNegociacion.text = "" Then
        MsgBox "Introduzca el Porcentaje de Negociación !!", vbCritical
        Valida = False
        txtNegociacion.SetFocus
        Exit Function
    End If

    If txt10K.text = "" Then
        MsgBox "Introduzca el Precio del Oro de 10K !!", vbCritical
        Valida = False
        txt10K.SetFocus
        Exit Function
    End If

    If txt14K.text = "" Then
        MsgBox "Introduzca el Precio del Oro de 14K !!", vbCritical
        Valida = False
        txt14K.SetFocus
        Exit Function
    End If

    If txtVenta.text = "" Then
        MsgBox "Introduzca el Porcentaje de Gastos de Venta !!", vbCritical
        Valida = False
        txtVenta.SetFocus
        Exit Function
    End If

    If txtEnajenacion.text = "" Then
        MsgBox "Introduzca el Número de Días de Enajenación !!", vbCritical
        Valida = False
        txtEnajenacion.SetFocus
        Exit Function
    End If

    If txtVApartado.text = "" Then
        MsgBox "Introduzca el Número de Meses para Ventas de Apartado !!", vbCritical
        Valida = False
        txtVApartado.SetFocus
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

Private Sub txtVenta21K_GotFocus()
    Seleccionar_Texto txtVenta21K
    Cambiar_Color True, txtVenta21K
End Sub

Private Sub txtVenta21K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVenta21K_LostFocus()
    Cambiar_Color False, txtVenta21K
End Sub

Private Sub txtVenta24K_GotFocus()
    Seleccionar_Texto txtVenta24K
    Cambiar_Color True, txtVenta24K
End Sub

Private Sub txtVenta24K_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtVenta24K_LostFocus()
    Cambiar_Color False, txtVenta24K
End Sub

'''''Function CargaImpresoras() As Boolean
'''''Dim prt As Printer
'''''
'''''    cmbImpresoras.AddItem ""
'''''    For Each prt In Printers
'''''
'''''        cmbImpresoras.AddItem prt.DeviceName
'''''    Next prt
'''''
'''''    Set prt = Nothing
'''''End Function

Private Sub txtHorario_GotFocus()
    Seleccionar_Texto TxtHorario
    Cambiar_Color True, TxtHorario
End Sub

Private Sub txtHorario_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtHorario_LostFocus()
    Cambiar_Color False, TxtHorario
End Sub

