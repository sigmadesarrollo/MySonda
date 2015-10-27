VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmInventario 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras"
   ClientHeight    =   7650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmInventario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7650
   ScaleWidth      =   11685
   Begin vbalTabStrip6.TabControl tTab 
      Height          =   7005
      Left            =   30
      TabIndex        =   4
      Top             =   105
      Width           =   11640
      _ExtentX        =   20532
      _ExtentY        =   12356
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CoolTabs        =   1
      Begin VB.Frame frmEntradas 
         Caption         =   "ENTRADAS"
         Height          =   6480
         Left            =   75
         TabIndex        =   5
         Top             =   480
         Width           =   11535
         Begin VB.TextBox txtIva 
            Alignment       =   1  'Right Justify
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   3360
            TabIndex        =   3
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtNombre 
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1200
            TabIndex        =   0
            Top             =   600
            Width           =   3615
         End
         Begin VB.TextBox txtTelefono 
            BorderStyle     =   0  'None
            Height          =   195
            Left            =   1200
            TabIndex        =   2
            Top             =   1320
            Width           =   1455
         End
         Begin VB.TextBox txtDireccion 
            BorderStyle     =   0  'None
            Height          =   210
            Left            =   1200
            TabIndex        =   1
            Top             =   960
            Width           =   4110
         End
         Begin VB.TextBox txtPeso 
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
            Left            =   7920
            MaxLength       =   3
            TabIndex        =   18
            Top             =   3120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtKilates 
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
            Left            =   6000
            MaxLength       =   2
            TabIndex        =   16
            Top             =   3120
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtPrecio 
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
            Left            =   6000
            MaxLength       =   5
            TabIndex        =   20
            Top             =   3480
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtDescripcion 
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
            Left            =   2880
            TabIndex        =   14
            Top             =   3120
            Visible         =   0   'False
            Width           =   3855
         End
         Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
            Height          =   4200
            Left            =   60
            TabIndex        =   21
            Top             =   1695
            Width           =   11430
            _ExtentX        =   20161
            _ExtentY        =   7408
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
            HeaderFlat      =   -1  'True
            BorderStyle     =   2
            ScrollBarStyle  =   2
            Editable        =   -1  'True
            DisableIcons    =   -1  'True
            Begin VB.TextBox txtSerie 
               BorderStyle     =   0  'None
               Height          =   255
               Left            =   10440
               TabIndex        =   42
               Top             =   0
               Visible         =   0   'False
               Width           =   975
            End
            Begin VB.ComboBox cmbTipo 
               Height          =   315
               Left            =   9120
               Style           =   2  'Dropdown List
               TabIndex        =   41
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtGrupo 
               BorderStyle     =   0  'None
               Height          =   210
               Left            =   840
               TabIndex        =   36
               Top             =   120
               Visible         =   0   'False
               Width           =   1140
            End
            Begin VB.TextBox txtPesoo 
               BorderStyle     =   0  'None
               Height          =   405
               Left            =   8040
               TabIndex        =   32
               Top             =   240
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.ComboBox dcbKilates 
               Appearance      =   0  'Flat
               Height          =   315
               ItemData        =   "frmInventario.frx":000C
               Left            =   5955
               List            =   "frmInventario.frx":002B
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   31
               Top             =   0
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.TextBox txtCode 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   120
               MaxLength       =   8
               TabIndex        =   30
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtArticulo 
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   1560
               MaxLength       =   100
               TabIndex        =   29
               Top             =   0
               Visible         =   0   'False
               Width           =   3255
            End
            Begin VB.TextBox txtPrecioo 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   7920
               TabIndex        =   28
               Top             =   0
               Visible         =   0   'False
               Width           =   1095
            End
            Begin VB.TextBox txtCosto 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   6885
               TabIndex        =   27
               Top             =   0
               Visible         =   0   'False
               Width           =   930
            End
            Begin VB.TextBox txtCantidad 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   0  'None
               Height          =   240
               Left            =   5160
               TabIndex        =   26
               Top             =   0
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin VB.PictureBox cmdMosClave 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3960
            MousePointer    =   1  'Arrow
            ScaleHeight     =   225
            ScaleWidth      =   315
            TabIndex        =   6
            Top             =   2040
            Visible         =   0   'False
            Width           =   375
         End
         Begin DevPowerFlatBttn.FlatBttn cmdMosCliente 
            Height          =   240
            Left            =   4845
            TabIndex        =   37
            Top             =   600
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   423
            AlignCaption    =   4
            AutoSize        =   0   'False
            Caption         =   "..."
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
         Begin VB.Label Label11 
            Caption         =   "Iva:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   45
            Top             =   1320
            Width           =   495
         End
         Begin VB.Label lblTotal 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   7635
            TabIndex        =   40
            Top             =   6105
            Width           =   525
         End
         Begin VB.Label Label20 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            Caption         =   "Total a pagar:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   270
            Left            =   6030
            TabIndex        =   39
            Top             =   6090
            Width           =   1545
         End
         Begin VB.Label Label19 
            AutoSize        =   -1  'True
            BackColor       =   &H8000000C&
            Height          =   435
            Left            =   15
            TabIndex        =   38
            Top             =   5970
            Width           =   11475
         End
         Begin VB.Label Label10 
            Caption         =   "Teléfono:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   35
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label9 
            Caption         =   "Dirección:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   34
            Top             =   960
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Nombre:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   600
            Width           =   735
         End
         Begin VB.Label lblCodigo1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   240
            TabIndex        =   25
            Top             =   4560
            Width           =   75
         End
         Begin VB.Label lblDescripcion1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   1680
            TabIndex        =   24
            Top             =   4560
            Width           =   75
         End
         Begin VB.Label lblKilates1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   6360
            TabIndex        =   23
            Top             =   4560
            Width           =   75
         End
         Begin VB.Label lblPrecio1 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   285
            Left            =   7560
            TabIndex        =   22
            Top             =   4560
            Width           =   75
         End
         Begin VB.Label lblFolio 
            AutoSize        =   -1  'True
            Caption         =   "<Folio>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   7320
            TabIndex        =   10
            Top             =   240
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblConsecutivo 
            AutoSize        =   -1  'True
            Caption         =   "<Consecutivo>"
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
            TabIndex        =   8
            Top             =   2520
            Visible         =   0   'False
            Width           =   1635
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Consecutivo:"
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
            Left            =   2160
            TabIndex        =   7
            Top             =   2520
            Visible         =   0   'False
            Width           =   1365
         End
         Begin VB.Label lblFecha 
            AutoSize        =   -1  'True
            Caption         =   "<Fecha>"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   1200
            TabIndex        =   12
            Top             =   240
            Width           =   780
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Fecha:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   585
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Folio:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   6360
            TabIndex        =   9
            Top             =   240
            Visible         =   0   'False
            Width           =   480
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Peso:"
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
            Left            =   7200
            TabIndex        =   17
            Top             =   3120
            Visible         =   0   'False
            Width           =   585
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Kilates:"
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
            Left            =   5160
            TabIndex        =   15
            Top             =   3120
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Precio:"
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
            Left            =   5160
            TabIndex        =   19
            Top             =   3480
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descripción:"
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
            Left            =   3360
            TabIndex        =   13
            Top             =   3480
            Visible         =   0   'False
            Width           =   1305
         End
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   10485
      TabIndex        =   43
      Top             =   7215
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
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmInventario.frx":0071
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   9285
      TabIndex        =   44
      Top             =   7215
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
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmInventario.frx":0102
   End
End
Attribute VB_Name = "frmInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 27/05/2002
' Modulo frmInventario - frmInventario.frm
' Ultima Modificacion - 27/05/2002
'Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit
Dim Fl() As cFlatControl
Dim Band As Boolean

Private Function Validar() As Boolean
Dim i As Integer

Validar = True

If txtNombre.Text = "" Then
    MsgBox "Introduzca el nombre del cliente !!", vbInformation, "Compra de joyería"
    Validar = False
    txtNombre.SetFocus
    Exit Function
End If

If txtDireccion.Text = "" Then
    MsgBox "Introduzca la dirección del cliente !!", vbInformation, "Compra de joyería"
    Validar = False
    txtDireccion.SetFocus
    Exit Function
End If

For i = 1 To grdArticulos.Rows
    If grdArticulos.CellText(i, 1) = "" Or grdArticulos.CellText(i, 2) = "" Or grdArticulos.CellText(i, 3) = "" Or grdArticulos.CellText(i, 4) = "" Or grdArticulos.CellText(i, 5) = "" Or grdArticulos.CellText(i, 6) = "" Or grdArticulos.CellText(i, 7) = "" Or grdArticulos.CellText(i, 8) = "" Or grdArticulos.CellText(i, 9) = "" Then
        If grdArticulos.CellText(i, 1) = "" And grdArticulos.CellText(i, 2) = "" And grdArticulos.CellText(i, 3) = "" And grdArticulos.CellText(i, 4) = "" And grdArticulos.CellText(i, 5) = "" And grdArticulos.CellText(i, 6) = "" And grdArticulos.CellText(i, 7) = "" And grdArticulos.CellText(i, 8) = "" And grdArticulos.CellText(i, 9) = "" Then GoTo 125
        
        If grdArticulos.CellText(i, 8) = "" Then MsgBox "Seleccione el tipo del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 8, 13, False: Exit Function
        If grdArticulos.CellText(i, 1) = "" Then MsgBox "Seleccione el grupo del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 1, 13, False: Exit Function
        If grdArticulos.CellText(i, 2) = "" Then MsgBox "Introduzca la descripción del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 2, 13, False: Exit Function
        If grdArticulos.CellText(i, 3) = "" Then MsgBox "Introduzca la cantidad de artículos !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 3, 13, False: Exit Function
        If grdArticulos.CellText(i, 4) = "" And grdArticulos.CellText(i, 8) = "ORO" Then MsgBox "Seleccione el kilataje del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 4, 13, False: Exit Function
        If grdArticulos.CellText(i, 5) = "" And grdArticulos.CellText(i, 8) = "ORO" Then MsgBox "Introduzca el peso del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 5, 13, False: Exit Function
        If grdArticulos.CellText(i, 6) = "" Then MsgBox "Introduzca el costo del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 6, 13, False: Exit Function
        If grdArticulos.CellText(i, 7) = "" Then MsgBox "Introduzca el precio del artículo !!", vbInformation, "Compra de joyería": Validar = False: grdArticulos.SelectedRow = i: grdArticulos_RequestEdit i, 7, 13, False: Exit Function
    End If
125:
Next i
End Function

Private Sub cmbTipo_Click()
grdArticulos.CellText(grdArticulos.SelectedRow, 8) = cmbTipo.Text
grdArticulos.CellItemData(grdArticulos.SelectedRow, 8) = cmbTipo.ItemData(cmbTipo.ListIndex)
cmbTipo.Visible = False
grdArticulos_CancelEdit
End Sub

Private Sub cmbTipo_GotFocus()
cmbTipo.BackColor = &HC0FFFF
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 9) = cmbTipo.Text
    cmbTipo.Visible = False
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
End If
End Sub

Private Sub cmbTipo_LostFocus()
cmbTipo.BackColor = vbWhite
End Sub

Private Sub cmdImprimir_Click()
Screen.MousePointer = vbHourglass
If Validar_Datos Then If Validar Then Grabar_Entradas
Screen.MousePointer = vbDefault
End Sub

Private Sub cmdMosCliente_Click()
'Unload frmMostrarclientecompra
frmMostrarclientecompra.Ver Me, txtNombre, False
'frmMostrarclientecompra.Show
Exit Sub
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub dcbKilates_Click()
Dim Cantidad As Integer, Peso As Double, Total As Double, kilataje As String, Kilates As String


If dcbKilates.ListIndex = -1 Then Exit Sub
If dcbKilates.Text <> "" Then Kilates = dcbKilates.Text Else Kilates = grdArticulos.CellText(grdArticulos.SelectedRow, 4)
kilataje = dcbKilates.ItemData(dcbKilates.ListIndex)

grdArticulos.CellText(grdArticulos.SelectedRow, 4) = dcbKilates.Text 'sacakilates(grdArticulos.CellItemData(grdArticulos.SelectedRow, 5))
grdArticulos.CellItemData(grdArticulos.SelectedRow, 4) = kilataje

'--------
    If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then Peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else Peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" Then
        Set rcTmp = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.Text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 4), dcbKilates.Text) & " as costo from parametros")
        If Not rcTmp.BOF And Not rcTmp.EOF Then
            Total = (Peso * rcTmp!costo) * Cantidad
        End If
        
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    Else
        Total = 0
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    End If
   '--------

'''If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
'''If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else peso = 0
'''
'''If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" Then
'''    Set rcTmp = dbDatos.Execute("select " & "Venta" & dcbKilates.Text & " as costo from parametros")
'''    If Not rcTmp.BOF And Not rcTmp.EOF Then
'''        total = (peso * rcTmp!costo)
'''        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = total
'''        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
'''    End If
'''Else
'''    total = 0
'''    grdArticulos.CellText(grdArticulos.SelectedRow, 6) = total
'''    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
'''End If

grdArticulos.CancelEdit
dcbKilates.Visible = False
lblTotal.Caption = Format(Regresa_Total, "##,###0.00")
Set rcTmp = Nothing
End Sub

Private Sub dcbKilates_GotFocus()
dcbKilates.BackColor = &HC0FFFF
End Sub

Private Sub dcbKilates_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then dcbKilates.Visible = False
End Sub

Private Sub dcbKilates_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 5) = dcbKilates.Text
    dcbKilates.Visible = False
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
End If
End Sub

Private Sub dcbKilates_LostFocus()
dcbKilates.BackColor = vbWhite
End Sub

Private Sub Form_Load()
Inicializar
End Sub

'Inicializamos la forma
Private Sub Inicializar()
Screen.MousePointer = vbHourglass
frmEntradas.BorderStyle = 0
lblFecha.Caption = Format(Date, "DD/MMM/YYYY")
Cargar_Combos "Descripcion", "Tipo", cmbTipo
CentrarForm Me, frmMDI
Crear_Pestañas
Crear_Encabezados
Poner_Flat Fl, Me.Controls, Me
lblFolio.Caption = Regresa_Movimiento(False, "FolioCompras")
Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados
Private Sub Crear_Encabezados()
With grdArticulos
   '.AddColumn "K1", "Código", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
   .AddColumn "K8", "Grupo", ecgHdrTextALignLeft, , 45, , , , , , , CCLSortString
   .AddColumn "K2", "Descripción", ecgHdrTextALignLeft, , 190, , , , , , , CCLSortString
   .AddColumn "K5", "Cantidad", ecgHdrTextALignRight, , 55, , , , , , , CCLSortString
   .AddColumn "K3", "Kilates", ecgHdrTextALignLeft, , 75, , , , , , , CCLSortNumeric
   .AddColumn "K7", "Peso", ecgHdrTextALignRight, , 50, , , , , "0.000", , CCLSortNumeric
   .AddColumn "K6", "Costo", ecgHdrTextALignRight, , 65, , , , , "###,###,###,##0.00", , CCLSortString
   .AddColumn "K4", "Precio", ecgHdrTextALignRight, , 75, , , , , "###,###,###,##0.00", , CCLSortNumeric
   .AddColumn "K9", "Tipo", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
   .AddColumn "K10", "Serie", ecgHdrTextALignLeft, , 75, , , , , , , CCLSortString
   .GridLines = True
   .Rows = 20
End With
End Sub

'Creamos las pestañas
Private Sub Crear_Pestañas()
With tTab
   .AddTab "Compra de joyería", , , "K1"
End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
Quitar_Flat Fl()
Unload Me
End Sub

Private Sub grdArticulos_CancelEdit()
txtCantidad.Visible = False
txtCosto.Visible = False
txtPrecioo.Visible = False
txtCode.Visible = False
txtArticulo.Visible = False
dcbKilates.Visible = False
dcbKilates.ListIndex = -1
txtPesoo.Visible = False
txtGrupo.Visible = False
cmbTipo.Visible = False
txtSerie.Visible = False
End Sub

Private Sub grdArticulos_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
Dim Cantidad As Integer, Precio As Double, Total As Double, i As Integer

If grdArticulos.SelectedRow > 0 Then
      If KeyCode = vbKeyDelete Then
       If MsgBox("Desea Eliminar Estos Articulos ??", vbQuestion + vbYesNo + vbDefaultButton2, "Dotación a Inventario") = vbYes Then
            Cantidad = 0
            Precio = 0
            Total = 0
            grdArticulos.RemoveRow grdArticulos.SelectedRow
            For i = 1 To grdArticulos.Rows
               Cantidad = grdArticulos.CellText(i, 3)
               Precio = grdArticulos.CellText(i, 6)
               Total = Total + (Cantidad * Precio)
            Next i
            'lblTotal.Caption = "$ " & Format(Total, "##,###0.00")
            'lblNumCap.Caption = grdArticulos.Rows
            grdArticulos.CancelEdit
            'txtClave.SetFocus
        End If
      End If
   End If
End Sub

Private Sub grdArticulos_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
 Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long
   Dim sText As String
   Dim obj As Object
   
   txtCantidad.Visible = False
   grdArticulos_CancelEdit
   
    If (lCol = 4 Or lCol = 5) And grdArticulos.CellText(grdArticulos.SelectedRow, 8) <> "ORO" Then Exit Sub
    'If lCol = 8 And grdArticulos.CellText(grdArticulos.SelectedRow, 8) = "METAL" Then Exit Sub
    
    Select Case lCol
    Case 1: Set obj = txtGrupo
    Case 2: Set obj = txtArticulo
    Case 3: Set obj = txtCantidad
    Case 4: Set obj = dcbKilates
    Case 5: Set obj = txtPesoo
    Case 6: Set obj = txtCosto
    Case 7: Set obj = txtPrecioo
    Case 8: Set obj = cmbTipo
    Case 9: Set obj = txtSerie
    Case Else: Exit Sub
    End Select
   
   grdArticulos.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight

   If Not IsMissing(grdArticulos.CellText(lRow, lCol)) Then
      sText = grdArticulos.CellText(lRow, lCol)
   Else
      sText = ""
   End If
   
   If lCol <> 4 And lCol <> 8 Then
         If lCol = 1 Or lCol = 2 Then
            obj.Alignment = vbLeftJustify
            grdArticulos.CellTextAlign(grdArticulos.SelectedRow, lCol) = DT_LEFT
         Else
            obj.Alignment = vbRightJustify
         End If
   
      'iKeyAscii = Solo_Numeros(iKeyAscii)
      If (iKeyAscii > 13) Then
         sText = Chr$(iKeyAscii) & sText
         obj.Text = sText
         obj.SelStart = 1
         obj.SelLength = Len(sText)
      Else
         obj.Text = sText
         obj.SelStart = 0
         obj.SelLength = Len(sText)
      End If
      
      Set txtCantidad.Font = grdArticulos.CellFont(lRow, lCol)
      If grdArticulos.CellBackColor(lRow, lCol) = -1 Then
         txtCantidad.BackColor = grdArticulos.BackColor
      Else
         txtCantidad.BackColor = grdArticulos.CellBackColor(lRow, lCol)
      End If
   End If
   
   If lCol <> 4 And lCol <> 8 Then
      obj.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
   Else
      obj.Move lLeft + 40, lTop + 25, lWidth - 60
   End If
   
   obj.Visible = True
   obj.ZOrder
   
   If lCol = 1 Then
    txtGrupo.Visible = True
    frmMostrarGrupo.Ver Me, txtGrupo, 1, 1
    'grdArticulos.CellText(grdArticulos.SelectedRow, 2) = txtGrupo.Text
    Exit Sub
   End If
   
   obj.SetFocus
End Sub

Private Sub tTab_TabClick(ByVal lTab As Long)
Select Case lTab
Case 1
    frmEntradas.Visible = True
    Limpiar
    txtNombre.SetFocus
End Select
End Sub

Private Sub txtArticulo_GotFocus()
txtArticulo.BackColor = &HC0FFFF
End Sub

Private Sub txtArticulo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtArticulo.Visible = False
End Sub

Private Sub txtArticulo_KeyPress(KeyAscii As Integer)
mayusculas KeyAscii
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 2) = txtArticulo.Text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 2) = DT_LEFT
    grdArticulos.CancelEdit
    txtArticulo.Visible = False
    grdArticulos.SetFocus
End If
End Sub

Private Sub txtArticulo_LostFocus()
txtArticulo.BackColor = vbWhite
End Sub

Private Sub txtCantidad_GotFocus()
txtCantidad.BackColor = &HC0FFFF
End Sub

Private Sub txtCantidad_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtCantidad.Visible = False
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
Dim i As Integer, Total As Double, Cantidad As Integer, Precio As Double, X As Integer, Peso As Double, kilataje As String


KeyAscii = Solo_Numeros(KeyAscii)

If KeyAscii = vbKeyReturn Then
   
   kilataje = RegresaKilates(IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 8) = "METAL", grdArticulos.CellText(grdArticulos.SelectedRow, 4), ""), grdArticulos.CellText(grdArticulos.SelectedRow, 8))

   Total = 0
   Cantidad = 0
   Precio = 0
   
   Set rcTmp = New ADODB.Recordset
   
   If txtCantidad.Text <> "" Then grdArticulos.CellText(grdArticulos.SelectedRow, 3) = txtCantidad.Text Else grdArticulos.CellText(grdArticulos.SelectedRow, 3) = "": GoTo 125
   grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 3) = DT_RIGHT
   For i = 1 To grdArticulos.Rows
        Cantidad = IIf(grdArticulos.CellText(i, 3) = "", 0, grdArticulos.CellText(i, 3))
        If grdArticulos.CellText(i, 5) <> "" Then Precio = grdArticulos.CellText(i, 7) Else Precio = 0
        Total = Total + (Cantidad * Precio)
   Next i
  
    '--------
    If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then Peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else Peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" Then
        Set rcTmp = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.Text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 4), dcbKilates.Text) & " as costo from parametros")
        If Not rcTmp.BOF And Not rcTmp.EOF Then
            Total = (Peso * rcTmp!costo) * Cantidad
        End If
        
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    Else
        Total = 0
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
    End If
   '--------
   
   
125:
   grdArticulos.CancelEdit
   txtCantidad.Visible = False
   lblTotal.Caption = Format(Regresa_Total, "##,###0.00")
   KeyAscii = 0
   grdArticulos.SetFocus
ElseIf KeyAscii = vbKeyEscape Then
   grdArticulos.CancelEdit
   KeyAscii = 0
End If

Set rcTmp = Nothing
End Sub

Private Sub txtCantidad_LostFocus()
txtCantidad.BackColor = vbWhite
End Sub

'Buscamos el grupo y lo visualizamos
Private Function Buscar_Clave(Codigo As String) As Boolean
   On Error GoTo error
   Dim rcGrupos As New ADODB.Recordset
   
   Buscar_Clave = True
   
   rcGrupos.Open "SELECT * FROM Grupos WHERE Clave='" & Codigo & "'", dbDatos, adOpenKeyset, adLockOptimistic
   
   If rcGrupos.RecordCount = 0 Then
      Buscar_Clave = False
   Else
      txtDescripcion.Text = rcGrupos!Descripcion
      'txtClave.Tag = rcGrupos!ID
   End If

   rcGrupos.Close
   
error:
   Maneja_Error Err
   
   Set rcGrupos = Nothing
      
End Function

Private Sub txtCode_GotFocus()
txtCode.BackColor = &HC0FFFF
End Sub

Private Sub txtCode_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtCode.Visible = False
End Sub

Private Sub txtCode_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 1) = txtCode.Text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 1) = DT_LEFT
    grdArticulos.CancelEdit
    txtCode.Visible = False
End If
End Sub

Private Sub txtCode_LostFocus()
txtCode.BackColor = vbWhite
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
   KeyAscii = Solo_Numeros(KeyAscii)
   If KeyAscii = vbKeyReturn Then
      'If Len(txtCodigo.Text) = 8 Then 'txtCodigo.Text = "0" & txtCodigo.Text
         'Buscar_Articulo txtCodigo.Text, grdSalidas, lblTotalSalida, txtCodigo
      'End If
   End If
   'Pasar_Foco KeyAscii
End Sub

Private Sub txtCosto_GotFocus()
txtCosto.BackColor = &HC0FFFF
End Sub

Private Sub txtCosto_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtCosto.Visible = False
End Sub

Private Sub txtCosto_KeyPress(KeyAscii As Integer)
Dim kilataje As String, Cantidad As Integer, Peso As Double, Total As Double


KeyAscii = Solo_Numeros(KeyAscii, 1)

If KeyAscii = vbKeyReturn Then
   
   kilataje = RegresaKilates(grdArticulos.CellText(grdArticulos.SelectedRow, 4))
   
   grdArticulos.CellText(grdArticulos.SelectedRow, 6) = txtCosto.Text
   grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
   grdArticulos.CancelEdit
    
    'If grdArticulos.CellText(grdArticulos.SelectedRow, 7) <> "" Then txtCosto.Text = grdArticulos.CellText(grdArticulos.SelectedRow, 7) Else txtCosto = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
    If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then Peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else Peso = 0
    
    If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Plata" Then
        Set rcTmp = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.Text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 4), dcbKilates.Text) & " as costo from parametros")
        If Not rcTmp.BOF And Not rcTmp.EOF Then
            Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 6) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 6))
            grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
        End If
    Else
        
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 6) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 6))
        grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
    End If
   
   txtCosto.Visible = False
   lblTotal.Caption = Format(Regresa_Total, "##,###0.00")

   KeyAscii = 0
   grdArticulos.SetFocus
ElseIf KeyAscii = vbKeyEscape Then
   grdArticulos.CancelEdit
   grdArticulos.SetFocus
   KeyAscii = 0
End If
Set rcTmp = Nothing
End Sub

Private Sub txtCosto_LostFocus()
txtCosto.BackColor = vbWhite
End Sub

Private Sub txtDescripcion_GotFocus()
   Seleccionar_Texto txtDescripcion
   Cambiar_Color True, txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
   KeyAscii = mayusculas(KeyAscii)
End Sub

Private Sub txtDescripcion_LostFocus()
   Cambiar_Color False, txtDescripcion
End Sub

Private Sub txtDireccion_GotFocus()
Seleccionar_Texto txtDireccion
Cambiar_Color True, txtDireccion
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtDireccion_LostFocus()
Cambiar_Color False, txtDireccion
End Sub

Private Sub txtDireccion2_KeyPress(KeyAscii As Integer)
Pasar_Foco KeyAscii
End Sub

Private Sub txtGrupo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtGrupo.Visible = False
End Sub

Public Sub txtGrupo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 1) = txtGrupo.Text
    'grdArticulos.CellText(grdArticulos.SelectedRow, 1) = genera_codigo(grdArticulos.CellText(grdArticulos.SelectedRow, 2), Format(IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 8) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 8)), "00000"))
    txtGrupo.Visible = False
End If
End Sub

Private Sub txtIva_GotFocus()
Seleccionar_Texto txtIva
Cambiar_Color True, txtIva
End Sub

Private Sub txtIva_KeyPress(KeyAscii As Integer)
KeyAscii = Solo_Numeros(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtIva_LostFocus()
Cambiar_Color False, txtIva
End Sub

Private Sub txtKilates_GotFocus()
   Seleccionar_Texto txtKilates
   Cambiar_Color True, txtKilates
End Sub

Private Sub txtKilates_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
   KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtKilates_LostFocus()
   Cambiar_Color False, txtKilates
End Sub

Private Sub txtNombre_GotFocus()
Seleccionar_Texto txtNombre
Cambiar_Color True, txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
Pasar_Foco KeyAscii
End Sub

Private Sub txtNombre_LostFocus()
Cambiar_Color False, txtNombre
End Sub

Private Sub txtNombre2_KeyPress(KeyAscii As Integer)
Pasar_Foco KeyAscii
End Sub

Private Sub txtPeso_GotFocus()
   Seleccionar_Texto txtPeso
   Cambiar_Color True, txtPeso
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
   KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtPeso_LostFocus()
   Cambiar_Color False, txtPeso
End Sub

Private Sub txtPesoo_GotFocus()
txtPesoo.BackColor = &HC0FFFF
End Sub

Private Sub txtPesoo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtPesoo.Visible = False
End Sub

Private Sub txtPesoo_KeyPress(KeyAscii As Integer)
Dim Peso As Double, Cantidad As Integer, costo As Double
Dim i As Integer

Dim Precio As Double
Dim Total As Double

KeyAscii = Solo_Numeros(KeyAscii, 1)
If KeyAscii = vbKeyReturn Then
    
    If txtPesoo.Text <> "" Then grdArticulos.CellText(grdArticulos.SelectedRow, 5) = txtPesoo.Text: Peso = txtPesoo.Text Else grdArticulos.CellText(grdArticulos.SelectedRow, 5) = "": Peso = 0
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 5) = DT_RIGHT
    
    Set rcTmp = New ADODB.Recordset
    
     '--------
     If grdArticulos.CellText(grdArticulos.SelectedRow, 3) <> "" Then Cantidad = grdArticulos.CellText(grdArticulos.SelectedRow, 3) Else Cantidad = 0
     If grdArticulos.CellText(grdArticulos.SelectedRow, 5) <> "" Then Peso = grdArticulos.CellText(grdArticulos.SelectedRow, 5) Else Peso = 0
     
     If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" Then
         Set rcTmp = dbDatos.Execute("select " & "Venta" & IIf(dcbKilates.Text = "", grdArticulos.CellText(grdArticulos.SelectedRow, 4), dcbKilates.Text) & " as costo from parametros")
         If Not rcTmp.BOF And Not rcTmp.EOF Then
             Total = (Peso * rcTmp!costo) * Cantidad
         End If
         grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
         grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
     Else
         Total = 0
         grdArticulos.CellText(grdArticulos.SelectedRow, 6) = Total
         grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 6) = DT_RIGHT
     End If
    '--------

125:
    grdArticulos.CancelEdit
    txtPesoo.Visible = False
    KeyAscii = 0
   
    grdArticulos.SetFocus
ElseIf KeyAscii = vbKeyEscape Then
   grdArticulos.CancelEdit
   KeyAscii = 0
End If

Set rcTmp = Nothing
End Sub

Private Sub txtPesoo_LostFocus()
txtPesoo.BackColor = vbWhite
End Sub

Private Sub txtPrecio_GotFocus()
   Seleccionar_Texto txtPrecio
   Cambiar_Color True, txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
   KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtPrecio_LostFocus()
   Cambiar_Color False, txtPrecio
End Sub

'Buscamos la categoria por medio de la clave
Public Sub Buscar_Grupo(Codigo As String)
   On Error GoTo error
   Dim rcGrupo As New ADODB.Recordset
   
   rcGrupo.Open "SELECT * FROM Grupos WHERE Clave='" & Codigo & "'", dbDatos, adOpenKeyset, adLockOptimistic
      
   With rcGrupo
      If .RecordCount = 0 Then
         MsgBox "El Grupo no se encuentra dado de alta", vbOKOnly + vbCritical
      Else
         'txtClave.Tag = !ID
         lblConsecutivo.Caption = Format(!consecutivo, "00000")
      End If
   End With
   
   rcGrupo.Close

error:
   Maneja_Error Err
   
   Set rcGrupo = Nothing

End Sub

'Grabamos todos los datos necesarios
Private Sub Grabar_Entradas()

If grdArticulos.Rows > 0 Then
   Screen.MousePointer = vbHourglass
   Grabar_Encabezado
   lblFolio.Caption = Regresa_Movimiento(False, "FolioCompras")
   Limpiar
   grdArticulos.GridLines = True
   grdArticulos.Rows = 20
   Screen.MousePointer = vbDefault
Else
    MsgBox "Introduzca los datos de los artículos que desea registrar !!", vbInformation, "Compra de joyería"
    txtCode.SetFocus
End If
End Sub

'imprimimos la entrada del inventario
Private Sub Imprimir_Entrada()
Dim Folio As Integer

With frmMDI.Cr
     Folio = lblFolio.Caption
    .Reset
    .DataFiles(0) = Path & "\Base De Datos\Datos.mdb"
    .DataFiles(1) = Path & "\Base De Datos\Datos.mdb"
    .Password = Chr(10) & "administrativo"
    .ReportFileName = Path & "\Reportes\EntradaInventarioo.rpt"
    .DiscardSavedData = True
    .SelectionFormula = "{DetallesEntradaInventario.folio}=" & Folio & ""
    .Formulas(0) = "Folio='" & Trim(lblFolio.Caption) & "'"
    .Destination = crptToWindow
    .WindowState = crptMaximized
    .Action = 1
End With
End Sub

Private Sub Limpiar()
grdArticulos.Clear
txtNombre.Tag = ""
txtNombre.Text = ""
txtDireccion.Text = ""
txtTelefono.Text = ""
txtIva.Text = ""
lblTotal.Caption = "0.00"
End Sub

'Grabamos el encabezado de la entrada
Private Function Grabar_Encabezado() As Long
Dim rcID As New ADODB.Recordset
Dim Folio As Long
    
On Error GoTo error

    Folio = Regresa_Movimiento(False, "FolioCompras")
    Regresa_Movimiento True, "FolioCompras"
    
    dbDatos.Execute "INSERT INTO EntradaInventario (Folio,Fecha,IDUsuario,IDSucursal,TipoEntrada) VALUES (" & Folio & ",'" & Format(lblFecha.Caption, "YYYY/MM/DD") & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & ENTRADACOMPRA & ")"
    rcID.Open "SELECT MAX(ID) AS IDD FROM EntradaInventario", dbDatos, adOpenForwardOnly, adLockOptimistic
   
    Grabar_Inventario rcID!idd, Folio
    rcID.Close
    
error:
    Maneja_Error Err
    Set rcID = Nothing
End Function

'Grabamos el inventario
Private Sub Grabar_Inventario(ID As Long, Folio As Long)
   On Error GoTo error
   Dim rcInventario As New ADODB.Recordset
   Dim Indice As Integer
   Dim Movimiento As Long
   Dim crImporte As Double
   Dim kilataje As Integer
   Dim IDCliente  As Long
   Dim rcClientes As New ADODB.Recordset
   Dim ImporteIva As Double, Iva As Integer
   Dim idd As Long, Codigo As String
    
    Iva = IIf(txtIva.Text = "", 0, txtIva.Text)
    crImporte = lblTotal.Caption - (lblTotal.Caption * (Iva / 100))
    ImporteIva = lblTotal.Caption * (Iva / 100)
   
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
   
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Folio & ",'EN01','620301'," & crImporte & "," & TIPO_CARGO & ",0,'Entrada Inventario','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Folio & ",'EN01','120101'," & ImporteIva & "," & TIPO_CARGO & ",0,'Entrada Inventario','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                  
   'Grabamos el abono
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Folio & ",'EN50','110150'," & crImporte & "," & TIPO_ABONO & ",0,'Entrada Inventario','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
             
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                  "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Folio & ",'EN50','199450'," & crImporte + ImporteIva & "," & TIPO_ABONO & ",0,'Entrada Inventario','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
           
    
    'ClienteCompras
    If txtNombre.Tag = "" Then
        dbDatos.Execute "insert into ClientesCompras(nombre,direccion,telefono)values('" & Trim(txtNombre.Text) & "','" & Trim(txtDireccion.Text) & "','" & Trim(txtTelefono.Text) & "')"
        rcClientes.Open "select max(id) as maximo from ClientesCompras", dbDatos, adOpenDynamic, adLockOptimistic
        If Not rcClientes.BOF And Not rcClientes.EOF Then IDCliente = rcClientes!maximo
        rcClientes.Close
    Else
        dbDatos.Execute "update ClientesCompras set nombre='" & Trim(txtNombre.Text) & "',direccion='" & Trim(txtDireccion.Text) & "',telefono='" & Trim(txtTelefono.Text) & "' where id=" & Val(txtNombre.Tag) & ""
        IDCliente = txtNombre.Tag
    End If
    
    'Registro la compra
    dbDatos.Execute "insert into Compras(Fecha,Folio,IDCliente,Total,Iva,IDUsuario,IDSucursal,Hora)values('" & Format(Date, "YYYY/MM/DD") & "'," & Folio & "," & IDCliente & "," & CDbl(crImporte) & "," & CDbl(ImporteIva) & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & Format(Time, "HH:MM:SS") & "')"
    Set rcTmp = dbDatos.Execute("select max(ID) as IDD from Compras")
    
    For Indice = 1 To grdArticulos.Rows
        
        If grdArticulos.CellText(Indice, 1) = "" And grdArticulos.CellText(Indice, 2) = "" And grdArticulos.CellText(Indice, 3) = "" And grdArticulos.CellText(Indice, 4) = "" And grdArticulos.CellText(Indice, 5) = "" And grdArticulos.CellText(Indice, 6) = "" And grdArticulos.CellText(Indice, 7) = "" Then GoTo 126
            
            kilataje = RegresaKilates(IIf(grdArticulos.CellText(Indice, 8) = "ORO", grdArticulos.CellText(Indice, 4), ""), grdArticulos.CellText(Indice, 8))
            
            Codigo = CreaCodigoBarras(frmMDI.IDSucursal, ENTRADACOMPRA, Trim(Folio), Indice)
            
            'DetallesEntradaInventario
            dbDatos.Execute "INSERT INTO DetallesEntradaInventario (IDEntrada,Codigo,Descripcion,Kilates,Peso,Costo,Precio,Cantidad,Tipo,Serie,SucursalOrigen,TipoEntrada) VALUES (" & _
                         ID & ",'" & Trim(Codigo) & "','" & grdArticulos.CellText(Indice, 2) & "'," & kilataje & "," & _
                         Val(grdArticulos.CellText(Indice, 5)) & "," & CDbl(grdArticulos.CellText(Indice, 6)) & "," & CDbl(grdArticulos.CellText(Indice, 7)) & "," & Val(grdArticulos.CellText(Indice, 3)) & "," & _
                         grdArticulos.CellItemData(Indice, 8) & ",'" & grdArticulos.CellText(Indice, 9) & "'," & frmMDI.IDSucursal & "," & ENTRADACOMPRA & ")"
            
            'DetallesCompra
            dbDatos.Execute "insert into DetallesCompras(IDCompra,Codigo,Descripcion,Cantidad,Kilataje,Peso,Costo,Precio,Tipo,Serie)values(" & rcTmp!idd & ",'" & Trim(Codigo) & "','" & grdArticulos.CellText(Indice, 2) & "'," & Val(grdArticulos.CellText(Indice, 3)) & "," & kilataje & "," & Val(grdArticulos.CellText(Indice, 5)) & "," & CDbl(grdArticulos.CellText(Indice, 6)) & "," & CDbl(grdArticulos.CellText(Indice, 7)) & "," & grdArticulos.CellItemData(Indice, 8) & ",'" & grdArticulos.CellText(Indice, 9) & "')"
                      
126:
      Next Indice
   
   If MsgBox("Desea imprimir nota de compra ??", vbQuestion + vbYesNo + vbDefaultButton1, "Compra de joyería") = vbYes Then
        With frmMDI.Cr
            .Reset
            .WindowShowPrintSetupBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & Regresa_Valor("MONTEPIO", "Servidor", "localhost")
            .ReportFileName = Path & "\Reportes\ComprobanteCompra.rpt"
            .SelectionFormula = "{DetallesCompras.IDCompra}=" & rcTmp!idd & ""
            .DiscardSavedData = True
            .WindowTitle = "Nota de compra"
            .Formulas(0) = "Total=" & CDbl(lblTotal.Caption) & ""
            .Formulas(1) = "Caja='" & Trim(UCase(NombrePc)) & "'"
            .Formulas(2) = "Usuario='" & Trim(UCase(frmMDI.Usuario)) & "'"
            .WindowState = crptMaximized
            .Action = 1
        End With
   End If
   
error:
   Maneja_Error Err
   
   Set rcInventario = Nothing
   Set rcClientes = Nothing
   Set rcTmp = Nothing
End Sub

'Buscamos el articulo en el inventario por el codigo para darle salida
Private Sub Buscar_Articulo_Fisico(Codigo As String, grd As vbalGrid, lbl As Label, txt As TextBox)
   On Error GoTo error
   Dim rcArticulo As New ADODB.Recordset
   Band = False
   
   rcArticulo.Open "SELECT * FROM Inventario WHERE Codigo='" & Codigo & "'", dbDatos, adOpenKeyset, adLockOptimistic
   
   If rcArticulo.RecordCount = 0 Then
      MsgBox "El Articulo no se encuentra, separarlo y dotarlo", vbOKOnly + vbCritical
      Band = True
      GoTo error
   ElseIf rcArticulo!existencia = 0 Then
      MsgBox "El Articulo tiene agotada su existencia, separarlo y dotarlo", vbOKOnly + vbCritical
      Band = True
      GoTo error
   ElseIf rcArticulo!existencia > 0 And rcArticulo!fisico < rcArticulo!existencia Then
      rcArticulo!fisico = rcArticulo!fisico + 1
      rcArticulo.Update
      'Leyenda.Caption = "REGISTRADO"
      'txtCodigoAjustes = ""
      'txtCodigoAjustes.SetFocus
   ElseIf rcArticulo!existencia > 0 And rcArticulo!fisico >= rcArticulo!existencia Then
      MsgBox ("El articulo ya llego a su maximo de existencia, separar el articulo y dotarlo")
      Band = True
   End If
   
   rcArticulo.Close

error:
   Maneja_Error Err
   
   Set rcArticulo = Nothing
   
   txt.SetFocus
   'txt_GotFocus

End Sub

Private Sub Ajustar_Inventario()
dbDatos.Execute "UPDATE Inventario SET Existencia=Fisico WHERE Existencia<>Fisico"
dbDatos.Execute "UPDATE Inventario SET Fisico=0"
End Sub

Private Sub txtPrecioo_GotFocus()
txtPrecioo.BackColor = &HC0FFFF
End Sub

Private Sub txtPrecioo_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then txtPrecioo.Visible = False
End Sub

Private Sub txtPrecioo_KeyPress(KeyAscii As Integer)
Dim i As Integer, Cantidad As Integer, Precio As Double, Total As Double

KeyAscii = Solo_Numeros(KeyAscii, 1)
If KeyAscii = vbKeyReturn Then
   
    grdArticulos.CellText(grdArticulos.SelectedRow, 7) = txtPrecioo.Text
    grdArticulos.CellTextAlign(grdArticulos.SelectedRow, 7) = DT_RIGHT
    grdArticulos.CancelEdit
        
    If grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Fino" And grdArticulos.CellText(grdArticulos.SelectedRow, 4) <> "Plata" Then
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 7) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 7))
    Else
        Total = IIf(grdArticulos.CellText(grdArticulos.SelectedRow, 7) = "", 0, grdArticulos.CellText(grdArticulos.SelectedRow, 7))
    End If

    txtCosto.Visible = False
    lblTotal.Caption = Format(Regresa_Total, "##,###0.00")
    KeyAscii = 0
    grdArticulos.SetFocus

ElseIf KeyAscii = vbKeyEscape Then
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
    KeyAscii = 0
End If
End Sub

Private Sub txtPrecioo_LostFocus()
txtPrecioo.BackColor = vbWhite
End Sub

Function REMATE(str As String) As Boolean
Dim rcBusca As New ADODB.Recordset


rcBusca.Open "select clave from grupos where clave='" & Trim(str) & "'", dbDatos, adOpenDynamic, adLockOptimistic
If Not rcBusca.BOF And Not rcBusca.EOF Then REMATE = False Else REMATE = True
rcBusca.Close
Set rcBusca = Nothing
End Function

Function remates(Folio As String, Codigo As String)
Dim rcBusca As New ADODB.Recordset
Dim i As Integer, Cantidad As Integer, Precio As Double, Total As Double
Dim ID As Long



rcBusca.Open "select id,destino from Empeno where codigoremate='" & Trim(Folio) & "'", dbDatos, adOpenDynamic, adLockOptimistic
If Not rcBusca.BOF And Not rcBusca.EOF Then
    If rcBusca!Destino = 5 Then
        ID = rcBusca!ID
        rcBusca.Close
        rcBusca.Open "select * from detallesempeño where idempeño=" & ID & " order by articulo", dbDatos, adOpenDynamic, adLockOptimistic
        If Not rcBusca.BOF And Not rcBusca.EOF Then
            rcBusca.MoveFirst
            While Not rcBusca.EOF
                grdArticulos.AddRow
                grdArticulos.CellText(grdArticulos.Rows, 1) = Codigo
                grdArticulos.CellText(grdArticulos.Rows, 2) = rcBusca!Articulo
                grdArticulos.CellText(grdArticulos.Rows, 3) = rcBusca!Cantidad
                grdArticulos.CellTextAlign(grdArticulos.Rows, 3) = DT_RIGHT
                grdArticulos.CellText(grdArticulos.Rows, 4) = rcBusca!Kilates
                grdArticulos.CellTextAlign(grdArticulos.Rows, 4) = DT_RIGHT
                grdArticulos.CellText(grdArticulos.Rows, 5) = rcBusca!Prestamo
                grdArticulos.CellTextAlign(grdArticulos.Rows, 5) = DT_RIGHT
                grdArticulos.CellText(grdArticulos.Rows, 6) = rcBusca!avaluo
                grdArticulos.CellTextAlign(grdArticulos.Rows, 6) = DT_RIGHT
                'Poner_Colores2 grdArticulos, grdArticulos.Rows, &HE0E0E0
            rcBusca.MoveNext
            Wend
            grdArticulos.CancelEdit
            
            For i = 1 To grdArticulos.Rows
                Cantidad = grdArticulos.CellText(i, 3)
                Precio = grdArticulos.CellText(i, 6)
                Total = Total + (Cantidad * Precio)
            Next i
            'lblTotal.Caption = "$ " & Format(Total, "##,###0.00")
            'lblNumCap.Caption = grdArticulos.Rows
        End If
    Else
        MsgBox "Esta Boleta todavia no ha sido Marcada como Remate !!", vbInformation, "Dotación a Inventario"
        'txtClave.SetFocus
    End If
End If
End Function

Function sacagrupo(codegrupo As String) As Integer
Dim str As String




str = Left(codegrupo, 2)

rcTmp.Open "select id from grupos where clave='" & Trim(str) & "'", dbDatos, adOpenDynamic, adLockOptimistic
If Not rcTmp.BOF And Not rcTmp.EOF Then sacagrupo = rcTmp!ID
End Function

Function mayusculas(ascii As Integer) As Integer
If (ascii >= 97) And (ascii <= 122) Then ascii = ascii - 32 Else If ascii = 39 Then ascii = 0
mayusculas = ascii
End Function

Private Function Validar_Datos() As Boolean
Dim i As Integer, X As Integer

X = 0
For i = 1 To grdArticulos.Rows
    If grdArticulos.CellText(i, 1) = "" Or grdArticulos.CellText(i, 2) = "" Or grdArticulos.CellText(i, 3) = "" Or grdArticulos.CellText(i, 4) = "" Or Trim(grdArticulos.CellText(i, 5)) = "" Or Trim(grdArticulos.CellText(i, 6)) = "" Or grdArticulos.CellText(i, 7) = "" Or grdArticulos.CellText(i, 8) = "" Or grdArticulos.CellText(i, 9) = "" Then
        If grdArticulos.CellText(i, 1) = "" And grdArticulos.CellText(i, 2) = "" And grdArticulos.CellText(i, 3) = "" And grdArticulos.CellText(i, 4) = "" And grdArticulos.CellText(i, 5) = "" And grdArticulos.CellText(i, 6) = "" And grdArticulos.CellText(i, 7) = "" And grdArticulos.CellText(i, 8) = "" And grdArticulos.CellText(i, 9) = "" Then X = X + 1
    End If
Next i

If X = grdArticulos.Rows Then MsgBox "Introduzca los datos de los artículos que desea comprar !!", vbInformation, "Dotación a Inventario": Validar_Datos = False: grdArticulos.SetFocus Else Validar_Datos = True
End Function

'Function genera_codigo(grupo As String, precio As String) As String
'
'genera_codigo = grupo & precio
'genera_codigo = genera_codigo & digitoverificador(genera_codigo)
'End Function

Function DigitoVerificador(Codigo As String) As String
Dim i As Integer, X(7) As Integer, suma As Integer, residuo As Double, vc As Integer, dc As Integer
Dim length As Integer

length = Len(Trim(Codigo))

For i = 1 To 7
    X(i) = Mid(Trim(Codigo), i, 1)
Next i

suma = 0

For i = 1 To 7
    residuo = i Mod 2
    If residuo <> 0 Then vc = 3 Else vc = 1
    suma = suma + (X(i) * vc)
Next i

dc = 10 - (suma Mod 10)
If dc = 10 Then dc = 0

DigitoVerificador = dc
End Function

Private Sub txtSerie_GotFocus()
Cambiar_Color True, txtSerie
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
KeyAscii = mayusculas(KeyAscii)
If KeyAscii = 13 Then
    grdArticulos.CellText(grdArticulos.SelectedRow, 9) = txtSerie
    txtSerie.Visible = False
    grdArticulos.CancelEdit
    grdArticulos.SetFocus
End If
End Sub

Private Sub txtSerie_LostFocus()
Cambiar_Color False, txtSerie
End Sub

Private Sub txtTelefono_GotFocus()
Cambiar_Color True, txtTelefono
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
Pasar_Foco KeyAscii
End Sub

Private Sub txtTelefono_LostFocus()
Cambiar_Color False, txtTelefono
End Sub

Function Regresa_Total() As Double
Dim Total As Double, i As Integer

Total = 0
For i = 1 To grdArticulos.Rows
    Total = Total + (grdArticulos.CellText(i, 6) * IIf(grdArticulos.CellText(i, 3) = "", 1, grdArticulos.CellText(i, 3)))
Next i
Regresa_Total = Total
End Function

Public Function Buscar(ID As Long)
On Error GoTo error

Set rcTmp = dbDatos.Execute("select * from ClientesCompras where id=" & ID & "")
If Not rcTmp.BOF And Not rcTmp.EOF Then
    txtNombre.Tag = ID
    txtNombre.Text = rcTmp!Nombre
    txtDireccion.Text = rcTmp!Direccion
    txtTelefono.Text = rcTmp!Telefono
End If

error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function
