VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmDemasias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Demasías"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDemasias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   11805
   Begin VB.Frame frmDemasias 
      Caption         =   "Demasias"
      Height          =   6735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   11775
      Begin VB.TextBox txtDemasia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   35.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   7545
         Locked          =   -1  'True
         TabIndex        =   34
         Text            =   "0.00"
         Top             =   5805
         Width           =   4155
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1650
         Index           =   0
         Left            =   8340
         Top             =   330
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2910
         Orientation     =   0
         ShadowColor     =   14737632
         LigthColor      =   14737632
         LineWidth       =   2
      End
      Begin VB.TextBox txtFolio 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   960
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin vbAcceleratorGrid6.vbalGrid grdDemasias 
         Height          =   3660
         Left            =   15
         TabIndex        =   2
         Top             =   2085
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   6456
         RowMode         =   -1  'True
         GridLines       =   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         GridLineColor   =   8421504
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
      End
      Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
         Height          =   375
         Left            =   2280
         TabIndex        =   3
         Top             =   180
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
         Picture         =   "frmDemasias.frx":000C
      End
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   60
         Index           =   0
         Left            =   8340
         Top             =   300
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   106
         ShadowColor     =   14737632
         LigthColor      =   14737632
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1650
         Index           =   1
         Left            =   11670
         Top             =   345
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2910
         Orientation     =   0
         ShadowColor     =   14737632
         LigthColor      =   14737632
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D3 
         Height          =   60
         Left            =   8340
         Top             =   1935
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   106
         ShadowColor     =   14737632
         LigthColor      =   14737632
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D1 
         Height          =   1650
         Index           =   2
         Left            =   9810
         Top             =   330
         Visible         =   0   'False
         Width           =   30
         _ExtentX        =   53
         _ExtentY        =   2910
         Orientation     =   0
         ShadowColor     =   14737632
         LigthColor      =   14737632
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   60
         Index           =   1
         Left            =   8340
         Top             =   720
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   106
         ShadowColor     =   14737632
         LigthColor      =   14737632
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   60
         Index           =   2
         Left            =   8340
         Top             =   1125
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   106
         ShadowColor     =   14737632
         LigthColor      =   14737632
         LineWidth       =   2
      End
      Begin Line3D.ucLine3D ucLine3D2 
         Height          =   60
         Index           =   3
         Left            =   8340
         Top             =   1530
         Visible         =   0   'False
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   106
         ShadowColor     =   14737632
         LigthColor      =   14737632
         LineWidth       =   2
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "IMPORTE A PAGAR:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   24
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   2820
         TabIndex        =   35
         Top             =   5940
         Width           =   4725
      End
      Begin VB.Label lblDemasia 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   9900
         TabIndex        =   33
         Top             =   1605
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label lblPrecioVenta 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   9900
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "P. VENTA:"
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
         Height          =   300
         Index           =   3
         Left            =   8415
         TabIndex        =   31
         Top             =   1200
         Visible         =   0   'False
         Width           =   1110
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEMASÍA:"
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
         Height          =   300
         Index           =   2
         Left            =   8415
         TabIndex        =   30
         Top             =   1605
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lblIntereses 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   9900
         TabIndex        =   29
         Top             =   795
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "INTERESES:"
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
         Height          =   300
         Index           =   1
         Left            =   8415
         TabIndex        =   28
         Top             =   795
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblPrestamo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   9900
         TabIndex        =   26
         Top             =   390
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRÉSTAMO:"
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
         Height          =   300
         Index           =   10
         Left            =   8415
         TabIndex        =   25
         Top             =   390
         Visible         =   0   'False
         Width           =   1350
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Height          =   1605
         Left            =   8355
         TabIndex        =   27
         Top             =   345
         Visible         =   0   'False
         Width           =   3315
      End
      Begin VB.Label Label91 
         AutoSize        =   -1  'True
         Caption         =   "Cp:"
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
         Left            =   5880
         TabIndex        =   22
         Top             =   1320
         Width           =   255
      End
      Begin VB.Label Label73 
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
         Left            =   4560
         TabIndex        =   21
         Top             =   1680
         Width           =   1200
      End
      Begin VB.Label Label72 
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
         Left            =   2280
         TabIndex        =   20
         Top             =   1680
         Width           =   780
      End
      Begin VB.Label Label71 
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
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label66 
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
         TabIndex        =   18
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label Label65 
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
         Left            =   2640
         TabIndex        =   17
         Top             =   1320
         Width           =   840
      End
      Begin VB.Label Label64 
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
         TabIndex        =   16
         Top             =   960
         Width           =   825
      End
      Begin VB.Label Label2 
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
         TabIndex        =   15
         Top             =   600
         Width           =   705
      End
      Begin VB.Label Label3 
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
         Left            =   3840
         TabIndex        =   14
         Top             =   600
         Width           =   810
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label lblApellido 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   4800
         TabIndex        =   12
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lblDireccion 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   7215
      End
      Begin VB.Label lblColonia 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   600
         TabIndex        =   10
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label lblMunicipio 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3540
         TabIndex        =   9
         Top             =   1320
         Width           =   2235
      End
      Begin VB.Label lblCp 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   6240
         TabIndex        =   8
         Top             =   1320
         Width           =   2055
      End
      Begin VB.Label lblEstado 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   750
         TabIndex        =   7
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label lblTelefono 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   3165
         TabIndex        =   6
         Top             =   1680
         Width           =   1290
      End
      Begin VB.Label lblIdentificacion 
         BackColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   5880
         TabIndex        =   5
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Contrato:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   795
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   10665
      TabIndex        =   23
      Top             =   6885
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
      Picture         =   "frmDemasias.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   9465
      TabIndex        =   24
      Top             =   6885
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
      Picture         =   "frmDemasias.frx":08E3
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   8235
      TabIndex        =   36
      Top             =   6885
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Cancelar"
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
      Picture         =   "frmDemasias.frx":0E35
      PictureDisabled =   "frmDemasias.frx":1084
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   6930
      TabIndex        =   37
      Top             =   6885
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
      Picture         =   "frmDemasias.frx":1C56
   End
End
Attribute VB_Name = "frmDemasias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
Dim Movimiento As Long, Folio As Long, crDemasia As Double, i As Integer

    If DatosValidos Then
                   
        'Folio
        Folio = txtFolio.text
    
        'Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
    
        'Importe de la Demasia
        crDemasia = CDbl(txtDemasia.text)
    
        'Marco las prendas que ya se pagaron
        MarcaDemasiasPagadas Val(txtFolio.Tag)
    
        'Grabamos el cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "','Demasia'," & Movimiento & "," & Folio & ",'DP01','650201'," & ConvMoneda(crDemasia) & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
  
        'Grabamos el abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "','Demasia'," & Movimiento & "," & Folio & ",'DP50','110150'," & ConvMoneda(crDemasia) & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

'''        'Grabamos el abono
'''        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
'''                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "','Demasia'," & Movimiento & "," & Folio & ",'DP50','199450'," & ConvMoneda(crDemasia) & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
        'Imprimo el ticket
        Imprimir Folio, False
        
        'Limpio la ventana
        cmdCancelar_Click
        
    End If

End Sub

Private Sub cmdBuscar_Click()
    If txtFolio.text <> "" Then BuscaContrato
End Sub

Private Sub cmdCancelar_Click()
    Limpiar "Demasias"
    grdDemasias.Clear
    lblPrestamo.Caption = "0.00"
    lblIntereses.Caption = "0.00"
    lblPrecioVenta.Caption = "0.00"
    lblDemasia.Caption = "0.00"
    txtDemasia.text = Format(0, FMoneda)
    txtFolio.SetFocus
End Sub

Private Sub cmdImprimir_Click()
Dim Folio As Long

    Folio = frmReimpresionrecibos.ReImprimir("empeno", "NumContrato", " WHERE NumContrato=")
    Imprimir Folio, True
    End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    frmDemasias.BorderStyle = 0
    Crear_Encabezado
    Limpiar "frmDemasias"
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl()
End Sub

Sub Crear_Encabezado()

    With grdDemasias
        .AddColumn "C1", "Código", ecgHdrTextALignLeft, , 90, False, , , , , , CCLSortString
        .AddColumn "C2", "Cant.", ecgHdrTextALignCentre, , 40, False, , , , , , CCLSortNumeric
        .AddColumn "C3", "Artículo", ecgHdrTextALignLeft, , 270, , , , , , , CCLSortString
        .AddColumn "C4", "Peso", ecgHdrTextALignRight, , 48, , , , , "0.000", , CCLSortNumeric
        .AddColumn "C5", "Kilates", ecgHdrTextALignCentre, , 50, , , , , , , CCLSortString
        .AddColumn "C6", "Avalúo", ecgHdrTextALignRight, , 80, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C7", "Préstamo", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C8", "Interés", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortString
        .AddColumn "C9", "P. Venta", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C10", "Demasía", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C11", "Estatus", ecgHdrTextALignLeft, , 70, , , , , , , CCLSortString
        
        .AddColumn "C12", "Intereses", ecgHdrTextALignLeft, , 70, False, , , , , , CCLSortString
        .AddColumn "C13", "Almacenaje", ecgHdrTextALignLeft, , 70, False, , , , , , CCLSortString
        .AddColumn "C14", "Seguro", ecgHdrTextALignLeft, , 70, False, , , , , , CCLSortString
        .AddColumn "C15", "Moratorios", ecgHdrTextALignLeft, , 70, False, , , , , , CCLSortString
        .AddColumn "C16", "GastosVenta", ecgHdrTextALignLeft, , 70, False, , , , , , CCLSortString
        .AddColumn "C17", "IVA", ecgHdrTextALignLeft, , 70, False, , , , , , CCLSortString
    End With

End Sub

Private Sub txtFolio_GotFocus()
    Seleccionar_Texto txtFolio
    Cambiar_Color True, txtFolio
End Sub

Private Sub txtFolio_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn And txtFolio.text <> "" Then BuscaContrato
End Sub

Private Sub txtFolio_LostFocus()
    Cambiar_Color False, txtFolio
End Sub

Function MuestraPrendas(IDEmpeno As Long)

    Dim rcArticulos As New ADODB.Recordset
    Dim strDestino As String, crDemasiaTotal As Double, crInteresTotal As Double, crPrecioVentaTotal As Double, crPrecioVenta As Double, crPrestamoTotal As Double, GTOSVenta As Double
    Dim crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crGastosVenta As Double, crIva As Double

On Error GoTo Error
    
    crPrestamoTotal = 0: crInteresTotal = 0: crPrecioVentaTotal = 0
    
    rcArticulos.Open "SELECT e.ID,e.NumContrato,e.TipoTasa,e.Fecha,e.Vencimiento,e.Operacion,e.Periodo," & _
                     "e.VenPeriodo,d.IDEmpeno,d.ID AS IDPrenda,d.Codigo,d.Articulo,d.Cantidad,d.Peso," & _
                     "d.Kilates,d.Avaluo,d.Prestamo,di.TipoSalida AS Destino,d.DemasiaPagada,dv.ID AS IDPrendaVenta,ve.Fecha AS FechaVenta " & _
                     "FROM empeno e INNER JOIN detallesempeno d ON e.ID=d.IDEmpeno LEFT JOIN " & _
                     "detallesentradainventario di ON d.Codigo=di.Codigo LEFT JOIN detallesventas dv " & _
                     "ON di.ID=dv.IDArticulo LEFT JOIN ventas ve ON dv.IDVenta=ve.ID WHERE d.Destino=" & D_VENTA & " AND d.IDEmpeno=" & IDEmpeno, dbDatos, adOpenForwardOnly, adLockOptimistic

    If Not rcArticulos.BOF And Not rcArticulos.EOF Then
        
        With grdDemasias
                        
            .Redraw = False
            .Clear
            
            GTOSVenta = Regresa_Valor_BD("GtosVenta")
            
            While Not rcArticulos.EOF
                
                crIntereses = 0: crAlmacenaje = 0: crSeguro = 0: crMoratorios = 0: crGastosVenta = 0: crIva = 0: crPrecioVenta = 0
                
                If IsNull(rcArticulos!FechaVenta) Then GoTo SinVender
                                                       
                'Intereses
                'crIntereses = Redondeo(GeneraIntereses(rcArticulos!Cantidad * rcArticulos!Prestamo, rcArticulos!Cantidad * rcArticulos!Avaluo, rcArticulos!NumContrato, IIf(rcArticulos!TipoTasa = "MENSUAL", rcArticulos!Fecha, DateAdd("D", -1, rcArticulos!Fecha)), IDEmpeno, "Tasa", rcArticulos!FechaVenta, , True))
                crIntereses = GeneraIntereses(rcArticulos!ID, "Tasa")
                
                'Almacenaje
                'crAlmacenaje = Redondeo(GeneraIntereses(rcArticulos!Cantidad * rcArticulos!Prestamo, rcArticulos!Cantidad * rcArticulos!Avaluo, rcArticulos!NumContrato, IIf(rcArticulos!TipoTasa = "MENSUAL", rcArticulos!Fecha, DateAdd("D", -1, rcArticulos!Fecha)), IDEmpeno, "Almacenaje", rcArticulos!FechaVenta, , True))
                crAlmacenaje = GeneraIntereses(rcArticulos!ID, "Almacenaje")
                
                'Seguro
                'crSeguro = Redondeo(GeneraIntereses(rcArticulos!Cantidad * rcArticulos!Prestamo, rcArticulos!Cantidad * rcArticulos!Avaluo, rcArticulos!NumContrato, IIf(rcArticulos!TipoTasa = "MENSUAL", rcArticulos!Fecha, DateAdd("D", -1, rcArticulos!Fecha)), IDEmpeno, "Seguro", rcArticulos!FechaVenta, , True))
                 crSeguro = GeneraIntereses(rcArticulos!ID, "Seguro")
                                                                                      
                'Gastos de Venta
                crGastosVenta = Redondeo((rcArticulos!Cantidad * rcArticulos!Prestamo) * (GTOSVenta / 100))
                                                                                       
                'IVA
                crIva = Redondeo(Regresa_Iva(crIntereses + crAlmacenaje + crSeguro + crMoratorios + crGastosVenta, IDEmpeno))

SinVender:

                'Saco el Precio de Venta
                crPrecioVenta = SacaPrecioVenta(rcArticulos!Codigo)
                
                'Saco el Destino
                strDestino = SacaDestino(rcArticulos!Codigo)
                
                .AddRow
                .CellText(.Rows, 1) = rcArticulos!Codigo
                .CellItemData(.Rows, 1) = rcArticulos!Destino
                .CellText(.Rows, 2) = rcArticulos!Cantidad
                .CellItemData(.Rows, 2) = rcArticulos!DemasiaPagada
                .CellTextAlign(.Rows, 2) = DT_CENTER
                .CellText(.Rows, 3) = rcArticulos!Cantidad & " " & rcArticulos!Articulo
                .CellItemData(.Rows, 3) = rcArticulos!IDPrenda
                .CellText(.Rows, 4) = rcArticulos!Peso
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                .CellText(.Rows, 5) = SacaKilates(rcArticulos!Kilates)
                .CellTextAlign(.Rows, 5) = DT_CENTER
                .CellText(.Rows, 6) = rcArticulos!Avaluo
                .CellTextAlign(.Rows, 6) = DT_RIGHT
                .CellText(.Rows, 7) = rcArticulos!Prestamo
                .CellTextAlign(.Rows, 7) = DT_RIGHT
                .CellText(.Rows, 8) = crIntereses + crAlmacenaje + crSeguro + crMoratorios + crGastosVenta + crIva
                .CellTextAlign(.Rows, 8) = DT_RIGHT
                .CellText(.Rows, 9) = crPrecioVenta
                .CellTextAlign(.Rows, 9) = DT_RIGHT
                .CellText(.Rows, 10) = IIf(crPrecioVenta = 0, 0, crPrecioVenta - (rcArticulos!Prestamo + crIntereses + crAlmacenaje + crSeguro + crMoratorios + crGastosVenta + crIva))
                .CellTextAlign(.Rows, 10) = DT_RIGHT
                .CellText(.Rows, 11) = IIf(rcArticulos!DemasiaPagada = 0, strDestino, "PAGADA")
                .CellTextAlign(.Rows, 11) = DT_LEFT
                .CellText(.Rows, 12) = crIntereses
                .CellText(.Rows, 13) = crAlmacenaje
                .CellText(.Rows, 14) = crSeguro
                .CellText(.Rows, 15) = crMoratorios
                .CellText(.Rows, 16) = crGastosVenta
                .CellText(.Rows, 17) = crIva
                .CellItemData(.Rows, 17) = IIf(IsNull(rcArticulos!IDPrendaVenta), 0, rcArticulos!IDPrendaVenta)
                
                Colorea grdDemasias, .Rows, IIf(.Rows Mod 2 <> 0, Trim(RGB(242, 254, 255)), Trim(RGB(255, 255, 255)))
                
                crPrestamoTotal = crPrestamoTotal + (rcArticulos!Cantidad * rcArticulos!Prestamo)
                crInteresTotal = crInteresTotal + (crIntereses + crAlmacenaje + crSeguro + crMoratorios + crGastosVenta + crIva)
                crPrecioVentaTotal = crPrecioVentaTotal + crPrecioVenta
                crDemasiaTotal = crDemasiaTotal + IIf(crPrecioVenta = 0, 0, IIf(strDestino = "VENTA" And rcArticulos!DemasiaPagada = 1, 0, crPrecioVenta - (rcArticulos!Prestamo + crIntereses + crAlmacenaje + crSeguro + crMoratorios + crGastosVenta + crIva)))
            rcArticulos.MoveNext
            Wend
            .Redraw = True
        End With
        
        lblPrestamo.Caption = Format(crPrestamoTotal, FMoneda)
        lblIntereses.Caption = Format(crInteresTotal, FMoneda)
        lblPrecioVenta.Caption = Format(crPrecioVenta, FMoneda)
        txtDemasia.text = Format(crDemasiaTotal, FMoneda)
        lblDemasia.Caption = Format(crPrecioVenta - (crPrestamoTotal + crInteresTotal), FMoneda)
    End If
    rcArticulos.Close
    Set rcArticulos = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcArticulos = Nothing
End Function

Function SacaPrecioVenta(Codigo As String) As Double
Dim rcPrecioVenta As New ADODB.Recordset

On Error GoTo Error
    
    rcPrecioVenta.Open "SELECT (detallesventas.Precio-((ventas.Descuento/100)*detallesventas.Precio)) AS PrecioVenta FROM detallesventas INNER JOIN ventas ON detallesventas.IDVenta=ventas.ID INNER JOIN detallesentradainventario ON detallesventas.IDArticulo=detallesentradainventario.ID WHERE ventas.Cancelado=0 AND if(ventas.Apartado=1,ventas.Pagado=1,ventas.Apartado=0) AND detallesentradainventario.Codigo='" & Trim(Codigo) & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcPrecioVenta.BOF And Not rcPrecioVenta.EOF Then
        
        SacaPrecioVenta = Redondeo(rcPrecioVenta!PrecioVenta)
    Else
        
        SacaPrecioVenta = 0
    End If
    rcPrecioVenta.Close
    Set rcPrecioVenta = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcPrecioVenta = Nothing
End Function

Private Sub Limpiar(Contededor As String)
Dim ctrl As Control
  
    For Each ctrl In Controls

        If ctrl.Container.Caption = Contededor Then
            
            If TypeOf ctrl Is TextBox Then ctrl.text = "": ctrl.Tag = ""
            If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
        
        End If

    Next

End Sub

Function SacaDestino(Codigo As String) As String
Dim rcDestino As New ADODB.Recordset

On Error GoTo Error

    rcDestino.Open "SELECT TipoSalida FROM detallesentradainventario WHERE Codigo='" & Trim(Codigo) & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
    Select Case rcDestino!TipoSalida

        Case 0
        
            SacaDestino = "ALMONEDA"
        Case SALIDAVENTA
        
            SacaDestino = "VENTA"
        Case SALIDATRASPASO
        
            SacaDestino = "TRASPASO"
    End Select
    rcDestino.Close
    Set rcDestino = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcDestino = Nothing
End Function

Sub BuscaContrato()
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error
    
    rcConsulta.Open "SELECT empeno.ID,clientes.Nombre,clientes.Apellido,clientes.Direccion,clientes.Colonia,clientes.Municipio,clientes.CP,clientes.Estado,clientes.Tel,clientes.Identificacion " & _
                    "FROM empeno LEFT JOIN clientes ON empeno.IDCliente=clientes.ID WHERE empeno.NumContrato=" & Val(txtFolio.text) & " AND (empeno.Serie=" & SERIE_A & " OR empeno.Serie=" & SERIE_C & ") AND empeno.Cancelado=0 AND empeno.DemasiaPagada=0 AND (empeno.Destino=" & D_ALMONEDA & ")", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
    
        With rcConsulta
            txtFolio.Tag = !ID
            lblNombre.Caption = !Nombre
            lblApellido.Caption = !Apellido
            lblDireccion.Caption = !Direccion
            lblColonia.Caption = !Colonia
            lblMunicipio.Caption = !Municipio
            lblCP.Caption = !CP
            lblEstado.Caption = !Estado
            lblTelefono.Caption = !Tel
            lblIdentificacion.Caption = !Identificacion
            MuestraPrendas !ID
        End With

    Else
                
        MsgBox "No se encontró el contrato especificado !!", vbCritical, "Demasías"
        txtFolio.SetFocus
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Function DatosValidos() As Boolean
Dim crDemasia As Double
    
    DatosValidos = True
    crDemasia = 0
    
    If Val(txtFolio.Tag) = 0 Then
        DatosValidos = False
        Exit Function
    End If
    
    If Val(txtDemasia.text) > 0 Or Trim(txtDemasia.text) <> "" Then
        crDemasia = txtDemasia.text
    End If
    
    If crDemasia = 0 Then
        MsgBox "No existe demasía por pagar !!", vbInformation, "Demasías"
        DatosValidos = False
        Exit Function
    End If
    
End Function

Sub MarcaDemasiasPagadas(IDEmpeno As Long)
Dim i As Integer, FechaHora As Date
    
    FechaHora = Now
    For i = 1 To grdDemasias.Rows
        
        If grdDemasias.CellItemData(i, 2) = 0 And grdDemasias.CellText(i, 11) = "VENTA" Then
            
            dbDatos.Execute "UPDATE detallesempeno SET DemasiaPagada=1 WHERE ID=" & grdDemasias.CellItemData(i, 3)
            
            dbDatos.Execute "UPDATE detallesventas SET Intereses=" & ConvMoneda(grdDemasias.CellText(i, 12)) & ",Almacenaje=" & ConvMoneda(grdDemasias.CellText(i, 13)) & "," & _
                            "Seguro=" & ConvMoneda(grdDemasias.CellText(i, 14)) & ",Moratorios=" & ConvMoneda(grdDemasias.CellText(i, 15)) & ",GtosVenta=" & ConvMoneda(grdDemasias.CellText(i, 16)) & _
                            ",ImporteIva=" & ConvMoneda(grdDemasias.CellText(i, 17)) & ",FechaDemasia='" & Format(FechaHora, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & CLng(grdDemasias.CellItemData(i, 17))
        End If
        
    Next i
    
End Sub

Sub Imprimir(Contrato As Long, Reimpresion As Boolean)
Dim ImprDefault As Boolean
Dim rcNota As New ADODB.Recordset

On Error GoTo Error
    
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
        
    rcNota.Open "SELECT " & IIf(Reimpresion, "DISTINCT", "MAX") & "(ve.Folio) AS FolioVenta FROM detallesentradainventario de INNER JOIN empeno e ON de.IDEmpeno=e.ID LEFT JOIN detallesventas dv ON de.ID=dv.IDArticulo " & _
                "INNER JOIN ventas ve ON dv.IDVenta=ve.ID WHERE e.NumContrato=" & Contrato & " ORDER BY ve.ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    While Not rcNota.EOF
        
        With frmMDI.Cr
            .Reset
            .DiscardSavedData = True
            .WindowShowPrintSetupBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\NotaDemasia.rpt"
            .SelectionFormula = "{ventas.Folio}=" & rcNota!FolioVenta & " AND {ventas.TipoVenta}=" & VENTAMOSTRADOR
            .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
            .Formulas(1) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
            .WindowState = crptMaximized
            .Destination = crptToWindow
            
            'La mando a la impresora por default
            If ImprDefault Then
                .PrinterName = strNombreImp
                .PrinterDriver = strDriverImp
                .PrinterPort = strPuertoImp
                .Destination = crptToPrinter
            End If
            
            .WindowTitle = "Nota Demasía"
            .Action = 1
        End With
    
    rcNota.MoveNext
    Wend
    rcNota.Close
    Set rcNota = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcNota = Nothing
End Sub
