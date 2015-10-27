VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMigracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Migracion"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMigracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   6375
   Begin VB.CheckBox ChckSoloAuxiliar 
      Caption         =   "Migrar solo Auxiliar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2400
      TabIndex        =   23
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CheckBox chAuxiliar 
      Caption         =   "Migrar Auxiliar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   22
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CheckBox chMyBD 
      Caption         =   "Migrar BD Completa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   21
      Top             =   1800
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.TextBox txtLog 
      Appearance      =   0  'Flat
      Height          =   2775
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   3600
      Width           =   6015
   End
   Begin VB.TextBox txtBD 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1440
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   420
      Width           =   4770
   End
   Begin DevPowerFlatBttn.FlatBttn cmdProcesar 
      Default         =   -1  'True
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   6480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "  &Procesar"
      Enabled         =   0   'False
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
      Picture         =   "frmMigracion.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   5070
      TabIndex        =   1
      Top             =   6480
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
      Picture         =   "frmMigracion.frx":0376
   End
   Begin MSComctlLib.ProgressBar ProgressBar 
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3120
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   360
      Width           =   915
      _ExtentX        =   1614
      _ExtentY        =   450
      AlignCaption    =   4
      AutoSize        =   0   'False
      Caption         =   "Conectar"
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
   Begin MSComDlg.CommonDialog CD 
      Left            =   5760
      Top             =   600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblContratosEnajenados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contratos Enajenados"
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
      Left            =   3600
      TabIndex        =   20
      Top             =   1080
      Width           =   2025
   End
   Begin VB.Label lblEnajenado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   3240
      TabIndex        =   19
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contratos Desempeñados"
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
      Left            =   600
      TabIndex        =   18
      Top             =   2160
      Width           =   2370
   End
   Begin VB.Label lblDesemp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   240
      TabIndex        =   17
      Top             =   2040
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tarjetas de Puntos"
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
      Left            =   600
      TabIndex        =   16
      Top             =   1440
      Width           =   1740
   End
   Begin VB.Label lblTarjetas 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   240
      TabIndex        =   15
      Top             =   1320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblVitrina 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   3240
      TabIndex        =   13
      Top             =   1320
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblContratos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   240
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblClientes 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Wingdings 3"
         Size            =   18
         Charset         =   2
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   420
      Left            =   240
      TabIndex        =   11
      Top             =   960
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Label lblRegistro 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<Registro>"
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
      Left            =   120
      TabIndex        =   10
      Top             =   2880
      Width           =   5280
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Clientes"
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
      Left            =   600
      TabIndex        =   9
      Top             =   1080
      Width           =   720
   End
   Begin VB.Label lblEstado 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estado: NO Conectado"
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
      Left            =   120
      TabIndex        =   8
      Tag             =   "0"
      Top             =   720
      Width           =   2505
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Seleccione la base de datos"
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
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2505
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vitrina/Apartados"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   1440
      Width           =   1665
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Contratos"
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
      Left            =   600
      TabIndex        =   2
      Top             =   1800
      Width           =   930
   End
End
Attribute VB_Name = "frmMigracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim dbCnn As ADODB.Connection
Dim IDOtros As Long, IDOro As Integer, IDElectronicos As Integer, IDLibreta As Integer, IDPlata As Long, IDAuto As Long
Dim ID8K As Long, ID10k As Long, ID14k As Long, ID18k As Long, ID22K As Long, ID24K As Long
Dim ID720 As Long, ID825 As Long, ID925 As Long, ID999 As Long
Dim CodigoArticulo As Long
Dim IDMarcaGeneral As Long, IDFamiliaGeneral As Long, IDPrendaGeneral As Long, IDGrupoGeneral As Long, IDOroGeneral As Long, IDPlataGeneral As Long
Dim CostoInventario As Currency, CostoInventarioOro As Currency, CompraOro As Currency, CompraMisc As Currency, CompraPlata As Currency, CompraAuto As Currency
Dim Suc_Int(1 To 12, 1 To 6) As Double
Dim Suc_Plazo(1 To 12, 1 To 3) As Integer

Private Sub cmdProcesar_Click()
   Migracion
End Sub

Private Sub Form_Load()
   Inicializar
End Sub

Private Sub Form_Unload(Cancel As Integer)
 If Not (dbCnn Is Nothing) Then
   If dbCnn.State = 1 Then dbCnn.Close
 End If
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
   
    lblRegistro.Caption = ""
    CentrarForm Me, frmMDI
    'Tablas de Intereses y Plazos
    'Sucursal          , Interes Mensual  , Seguro           , Almacenaje       , IVa               , CAT                   , Plazo 1             , Plazo 3
    Suc_Int(1, 1) = 101: Suc_Int(1, 2) = 5: Suc_Int(1, 3) = 2: Suc_Int(1, 4) = 2: Suc_Int(1, 5) = 16: Suc_Int(1, 6) = 108: Suc_Plazo(1, 1) = 90: Suc_Plazo(1, 3) = 30
    Suc_Int(2, 1) = 102: Suc_Int(2, 2) = 5: Suc_Int(2, 3) = 2.5: Suc_Int(2, 4) = 2.5: Suc_Int(2, 5) = 16: Suc_Int(2, 6) = 120: Suc_Plazo(2, 1) = 90: Suc_Plazo(2, 3) = 30
    Suc_Int(3, 1) = 103: Suc_Int(3, 2) = 5: Suc_Int(3, 3) = 2: Suc_Int(3, 4) = 2: Suc_Int(3, 5) = 16: Suc_Int(3, 6) = 108: Suc_Plazo(3, 1) = 90: Suc_Plazo(3, 3) = 30
    Suc_Int(4, 1) = 104: Suc_Int(4, 2) = 4: Suc_Int(4, 3) = 2.3: Suc_Int(4, 4) = 2.3: Suc_Int(4, 5) = 16: Suc_Int(4, 6) = 103: Suc_Plazo(4, 1) = 90: Suc_Plazo(4, 3) = 30
    Suc_Int(5, 1) = 105: Suc_Int(5, 2) = 4: Suc_Int(5, 3) = 2.3: Suc_Int(5, 4) = 2.3: Suc_Int(5, 5) = 16: Suc_Int(5, 6) = 103: Suc_Plazo(5, 1) = 90: Suc_Plazo(5, 3) = 30
    Suc_Int(6, 1) = 106: Suc_Int(6, 2) = 5: Suc_Int(6, 3) = 2: Suc_Int(6, 4) = 2: Suc_Int(6, 5) = 16: Suc_Int(6, 6) = 108: Suc_Plazo(6, 1) = 90: Suc_Plazo(6, 3) = 30
    Suc_Int(7, 1) = 108: Suc_Int(7, 2) = 5: Suc_Int(7, 3) = 2.5: Suc_Int(7, 4) = 2.5: Suc_Int(7, 5) = 16: Suc_Int(7, 6) = 120: Suc_Plazo(7, 1) = 90: Suc_Plazo(7, 3) = 30
    Suc_Int(8, 1) = 109: Suc_Int(8, 2) = 5: Suc_Int(8, 3) = 2: Suc_Int(8, 4) = 2: Suc_Int(8, 5) = 16: Suc_Int(8, 6) = 108: Suc_Plazo(8, 1) = 90: Suc_Plazo(8, 3) = 30
    Suc_Int(9, 1) = 110: Suc_Int(9, 2) = 5: Suc_Int(9, 3) = 2: Suc_Int(9, 4) = 2: Suc_Int(9, 5) = 16: Suc_Int(9, 6) = 108: Suc_Plazo(9, 1) = 90: Suc_Plazo(9, 3) = 30
    Suc_Int(10, 1) = 111: Suc_Int(10, 2) = 3: Suc_Int(10, 3) = 2: Suc_Int(10, 4) = 2: Suc_Int(10, 5) = 16: Suc_Int(10, 6) = 77: Suc_Plazo(10, 1) = 90: Suc_Plazo(10, 3) = 30
    Suc_Int(11, 1) = 112: Suc_Int(11, 2) = 5: Suc_Int(11, 3) = 2.5: Suc_Int(11, 4) = 2.5: Suc_Int(11, 5) = 16: Suc_Int(11, 6) = 120: Suc_Plazo(11, 1) = 90: Suc_Plazo(11, 3) = 30
    Suc_Int(12, 1) = 114: Suc_Int(12, 2) = 5: Suc_Int(12, 3) = 2: Suc_Int(12, 4) = 2: Suc_Int(12, 5) = 16: Suc_Int(12, 6) = 108: Suc_Plazo(12, 1) = 90: Suc_Plazo(12, 3) = 30
 
   'Inicializo Variables
    ID8K = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '8K'"))
    ID10k = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '10K'"))
    ID14k = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '14K'"))
    ID18k = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '18K'"))
    ID22K = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '22K'"))
    ID24K = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '24K'"))
    ID720 = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '.720'"))
    ID825 = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '.825'"))
    ID925 = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '.925'"))
    ID999 = Val(SacaValor("Kilatajes", "ID", " Where Descripcion = '.999'"))
    IDOro = Val(SacaValor("Tipo", "ID", " Where Descripcion = 'ORO'"))
    IDPlata = Val(SacaValor("Tipo", "ID", " Where Descripcion = 'PLATA'"))
    IDElectronicos = Val(SacaValor("Tipo", "ID", " Where Descripcion = 'ELECTRONICOS'"))
    IDAuto = 999
    If IDElectronicos = 0 Then IDElectronicos = Val(SacaValor("Tipo", "ID", " Where Descripcion = 'ELECTRODOMESTICO'"))
    If IDElectronicos = 0 Then IDElectronicos = Val(SacaValor("Tipo", "ID", " Where Descripcion = 'MISCELANEOS'"))


    IDOroGeneral = Val(SacaValor("TipoPrenda", "ID", "WHERE IDTipo=" & IDOro & " AND Descripcion='GENERAL'"))
    If IDOroGeneral = 0 Then
       dbDatos.Execute "INSERT INTO TipoPrenda (IDTipo,Descripcion) VALUES (" & IDOro & ",'GENERAL')"
       IDOroGeneral = Val(SacaValor("TipoPrenda", "MAX(ID)"))
    End If

    IDPlataGeneral = Val(SacaValor("TipoPrenda", "ID", "WHERE IDTipo=" & IDPlata & " AND Descripcion='GENERAL'"))
    If IDPlataGeneral = 0 Then
       dbDatos.Execute "INSERT INTO TipoPrenda (IDTipo,Descripcion) VALUES (" & IDPlata & ",'GENERAL')"
       IDPlataGeneral = Val(SacaValor("TipoPrenda", "MAX(ID)"))
    End If

    IDMarcaGeneral = Val(SacaValor("Marcas", "ID", "WHERE Descripcion='GENERAL'"))
    If IDMarcaGeneral = 0 Then
        dbDatos.Execute "INSERT INTO Marcas (Descripcion) VALUES ('GENERAL')"
        IDMarcaGeneral = SacaValor("Marcas", "MAX(ID)")
    End If

    IDFamiliaGeneral = Val(SacaValor("TipoPrenda", "ID", "WHERE Descripcion='GENERAL' AND IDTipo=" & IDElectronicos))
    If IDFamiliaGeneral = 0 Then
        dbDatos.Execute "INSERT INTO TipoPrenda (IDTipo,Descripcion) VALUES (" & IDElectronicos & ",'GENERAL')"
        IDFamiliaGeneral = SacaValor("TipoPrenda", "MAX(ID)")
    End If

    IDPrendaGeneral = Val(SacaValor("PrendaSelec", "ID", "WHERE IDMarca=" & IDMarcaGeneral & " AND IDTipo=" & IDElectronicos & " AND IDFamilia=" & IDFamiliaGeneral))
    If IDPrendaGeneral = 0 Then
        dbDatos.Execute "INSERT INTO PrendaSelec (IDMarca,IDTipo,IDFamilia) VALUES (" & _
                        IDMarcaGeneral & "," & IDElectronicos & "," & IDFamiliaGeneral & ")"
        IDPrendaGeneral = SacaValor("PrendaSelec", "MAX(ID)")
    End If


    IDLibreta = Val(SacaValor("Tipo", "ID", " Where Descripcion = 'LIBRETAS DE NAVIDAD'"))
    IDOtros = Val(SacaValor("Tipo", "ID", " Where Descripcion = 'OTROS'"))
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdBuscar_Click()
       CD.Filter = "Microsoft Access (*.mdb)|*.mdb"
       CD.ShowOpen
       
       If CD.FileName <> "" Then txtBD.text = CD.FileName
   
       'Hago la Conexión
       If Trim(txtBD.text) <> "" Then
           If CONEXION(txtBD.text) Then
               lblEstado.Caption = "Estado: CONECTADO"
               lblEstado.Tag = 1
               cmdProcesar.Enabled = True
           Else
               lblEstado.Caption = "Estado: DESCONECTADO"
               lblEstado.Tag = 0
           End If
       End If
End Sub

Private Function CONEXION(RutaBD As String) As Boolean
On Error GoTo Error
    
    CONEXION = False
    'Variable para la ruta de la aplicacion
 
   Set dbCnn = New ADODB.Connection
   dbCnn.Open CONEXION_ & RutaBD & Usuario_
    If dbCnn.State = 1 Then CONEXION = True
    
Error:
    Maneja_Error Err
End Function

Private Function Saca_Valor(Tabla As String, Campo As String, Optional Condicion As String = "") As String
Dim rcValor As New ADODB.Recordset

On Error GoTo Error
    
    rcValor.Open "SELECT " & Campo & " AS Valor FROM " & Tabla & " " & Condicion, dbCnn, adOpenForwardOnly, adLockReadOnly
    If Not rcValor.BOF And Not rcValor.EOF And Not IsNull(rcValor.Fields("Valor")) Then
        
        Saca_Valor = rcValor.Fields("Valor")
    Else
        
        Saca_Valor = ""
    End If
    rcValor.Close
    Set rcValor = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcValor = Nothing
End Function

Private Sub cmdSalir_Click()
   Unload Me
End Sub

'*****************************************************************************************************************************
'*****Funcion para separa nombres de la cadena********************************************************************
'*****************************************************************************************************************************
Function ObtieneNombres(ByVal Espacios As Integer, ByVal Cadena As String) As String
   Dim i As Integer
   Dim num As Integer
   Dim intLongitud As Integer
   Dim strNombres As String
   Dim strLetras As String
   
   'Obtiene el nombres cuando hay un nombre con los apellidos
   If Espacios = 0 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   Exit For
   End If
   strLetras = Mid(Cadena, i, 1)
   strNombres = strNombres & strLetras
   Next
   End If
   
   
   If Espacios = 1 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   Exit For
   End If
   strLetras = Mid(Cadena, i, 1)
   strNombres = strNombres & strLetras
   Next
   End If
   
   If Espacios = 2 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   Exit For
   End If
   strLetras = Mid(Cadena, i, 1)
   strNombres = strNombres & strLetras
   Next
   End If
   
   If Espacios = 3 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   If num = 2 Then Exit For
   End If
   strLetras = Mid(Cadena, i, 1)
   strNombres = strNombres & strLetras
   Next
   End If
   
   If Espacios >= 4 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   If num = 3 Then Exit For
   End If
   strLetras = Mid(Cadena, i, 1)
   strNombres = strNombres & strLetras
   Next
   End If
   
   ObtieneNombres = Trim(strNombres)

End Function

'*****************************************************************************************************************************
'*****Funcion para separa apellidos de la cadena*******************************************************************
'*****************************************************************************************************************************
Function ObtieneApellidos(ByVal Espacios As Integer, ByVal Cadena As String) As String
   Dim i As Integer
   Dim num As Integer
   Dim intLongitud As Integer
   Dim strApellidos As String
   Dim strLetras As String
   
   'Obtiene el nombres cuando hay un nombre con los apellidos
   If Espacios = 1 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   End If
   If num >= 1 Then
   strLetras = Mid(Cadena, i, 1)
   strApellidos = strApellidos & strLetras
   If num = 3 Then Exit For
   End If
   Next
   End If
   
   If Espacios = 2 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   End If
   If num >= 1 Then
   strLetras = Mid(Cadena, i, 1)
   strApellidos = strApellidos & strLetras
   If num = 3 Then Exit For
   End If
   Next
   End If
   
   If Espacios = 3 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   End If
   If num >= 2 Then
   strLetras = Mid(Cadena, i, 1)
   strApellidos = strApellidos & strLetras
   End If
   If num = 4 Then Exit For
   Next
   End If
   
   If Espacios >= 4 Then
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   End If
   If num >= 3 Then
   strLetras = Mid(Cadena, i, 1)
   strApellidos = strApellidos & strLetras
   End If
   If num = 5 Then Exit For
   Next
   End If
   
   ObtieneApellidos = Trim(strApellidos)

End Function

'**************************************************************************************************
'*****Funcion para contar espacios**********************************************************
'**************************************************************************************************
Function ContarEspacios(ByVal Cadena As String) As Integer
   Dim i As Integer
   Dim num As Integer
   Dim strLetras As String
   
   For i = 1 To Len(Cadena)
   If InStr(Mid(Cadena, i, 1), " ") <> 0 Then
   num = num + 1
   End If
   strLetras = Mid(Cadena, i, 3)
   If strLetras = "Y/O" Then
   num = num - 1
   Exit For
   End If
   Next
   ContarEspacios = num
End Function

Private Function SacaValorAccess(Tabla As String, Campo As String, Optional Condicion As String = "") As String
   On Error GoTo Error
   Dim rcValor As New ADODB.Recordset
    
    rcValor.Open "SELECT " & Campo & " AS Valor FROM " & Tabla & " " & Condicion, dbCnn, adOpenForwardOnly, adLockReadOnly
    If Not rcValor.BOF And Not rcValor.EOF And Not IsNull(rcValor.Fields("Valor")) Then
        SacaValorAccess = rcValor.Fields("Valor")
    Else
        SacaValorAccess = ""
    End If
    rcValor.Close
    Set rcValor = Nothing
    
Error:
    Maneja_Error Err
    Set rcValor = Nothing
End Function

'************************************************************************************************************************************************
'* Migraciones
'************************************************************************************************************************************************

Private Sub Migracion()
    On Error GoTo Error
    Screen.MousePointer = vbHourglass
    
    If ChckSoloAuxiliar Then
        Migrar_Auxiliar
    Else
    
        If chMyBD Then
            txtLog.text = ""
            Migrar_Abonos
            If chAuxiliar Then
                Migrar_Auxiliar
            End If
            Migrar_Bancos
            Migrar_Boveda
            Migrar_CierreDiario
            Migrar_Usuarios
            Migrar_Clientes
            Migrar_ClientesCompras
            Migrar_Compras
            Migrar_DetallesCompras
            Migrar_Contratos
            Migrar_Vitrina
            Migrar_DetEntra
            Migrar_Apartados
            Migrar_SalidaInventario
            Migrar_DetallesSalida
        Else
        End If
    
    End If

   
    MsgBox "Informacion migrada correctamente", vbOKOnly Or vbInformation
    ProgressBar.Value = 0
    lblRegistro.Caption = ""
    cmdProcesar.Enabled = False
     If Not (dbCnn Is Nothing) Then
    If dbCnn.State = 1 Then dbCnn.Close
     End If
   
Error:
    Maneja_Error Err
    
    Screen.MousePointer = vbDefault
End Sub
Private Sub Migrar_Compras()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String
    Dim Total As Integer
  
    Sql = "SELECT * FROM Compras order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Compras Registro " & ProgressBar.Value & "/" & ProgressBar.Max
        
        Total = IIf(SacaValorAccess("DetallesCompras", "Sum(Costo)", "WHERE IDCompra= " & rc!ID) = "", 0, SacaValorAccess("DetallesCompras", "Sum(Costo)", "WHERE IDCompra= " & rc!ID))
        
        dbDatos.Execute "INSERT INTO Compras(ID, Cancelado, Fecha, Folio, IDCliente, Total, Iva, IDSucursal, FechaMovimiento) VALUES (" & _
                        rc!ID & ",0,'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "'," & rc!Folio & "," & rc!IDCliente & "," & Total & ",0,101,'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "')"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_ClientesCompras()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String
  
    Sql = "SELECT * FROM ClientesCompras order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Clientes Compras Registro " & ProgressBar.Value & "/" & ProgressBar.Max
      
        dbDatos.Execute "INSERT INTO clientescompras(Id, Nombre, Direccion, Telefono) VALUES (" & _
                        rc!ID & ",'" & rc!Nombre & "','" & rc!Direccion & "','" & rc!Telefono & "')"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_DetallesSalida()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql, DescripKilates As String
    Dim rcDescripcio As ADODB.Recordset
    Dim rcidOro As ADODB.Recordset
    Dim rcidTipoOro As ADODB.Recordset
    Dim rcTipo As ADODB.Recordset
    Dim idKil, Idks, idkt, T As Integer
  
    Sql = "SELECT * FROM DetallesSalida order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Detalles Salida Registro " & ProgressBar.Value & "/" & ProgressBar.Max
            '/////////////////////
            'Consulta Descripcion
            DescripKilates = SacaValorAccess("Kilatajes", "Descripcion", "WHERE ID= " & rc!Kilates)
            '///////////////////////////'Consulta id
            Idks = 0
            idkt = 0
            If DescripKilates = "" Then
            
            
            Else
                Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                If Not rcidOro.EOF Then
                
                Else
                   idKil = Val(SacaValor("kilatajes", "MAX(ID)"))
                   T = SacaValorAccess("Kilatajes", "IDTipo", "WHERE Descripcion= '" & DescripKilates & "'")
                   dbDatos.Execute "insert into kilatajes (ID, Clave, Descripcion,IDTipo)" & _
                   "Values(" & idKil + 1 & "," & idKil + 1 & ",'" & DescripKilates & "'," & T & ")"
                   Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                End If
                
               Idks = rcidOro!ID
               idkt = rcidOro!IDTabla
            
            End If
           'Consulta id tipo
           If rc!Tipo = "DOCUMENTO" Then
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='DOCUMENTOS'")
           ElseIf rc!Tipo = "HERRAMIENTAS" Then
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='HERRAMIENTA'")
           Else
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='" & rc!Tipo & "'")
           End If
            '////////////
      
        dbDatos.Execute "INSERT INTO DetallesSalida(ID, IDSalidaInventario, IDArticulo, Codigo, Descripcion, Kilates, IDTablaKilates, Costo, Peso, Tipo, IDTablaTipo, Precio, Serie, IDEmpeno) VALUES (" & _
                        rc!ID & "," & rc!IDSalidaInventario & "," & rc!IDArticulo & ",'" & rc!Codigo & "','" & rc!Descripcion & "'," & Idks & "," & idkt & "," & rc!Costo & "," & rc!Peso & "," & rcTipo!ID & "," & rcTipo!IDTabla & _
                        "," & rc!Precio & ",'" & IIf(IsNull(rc!Serie), "", rc!Serie) & "'," & rc!IDEmpeño & " )"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_SalidaInventario()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String

    Sql = "SELECT * FROM SalidaInventario order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic

    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If

    While Not rc.EOF
        DoEvents

        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Salida Inventario Registro " & ProgressBar.Value & "/" & ProgressBar.Max

        dbDatos.Execute "INSERT INTO SalidaInventario(ID, Fecha, Pagado, Folio, TipoSalida, IDUsuario, IDSucursal) VALUES (" & _
                        rc!ID & ",'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "',0," & rc!Folio & ",0,1,101)"
       rc.MoveNext
   Wend
   rc.Close

Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_DetallesCompras()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql, DescripKilates As String
    Dim rcDescripcio As ADODB.Recordset
    Dim rcidOro As ADODB.Recordset
    Dim rcidTipoOro As ADODB.Recordset
    Dim rcTipo As ADODB.Recordset
    Dim idKil, Idks, idkt, T As Integer
  
    Sql = "SELECT * FROM DetallesCompras order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Detalles Compras Registro " & ProgressBar.Value & "/" & ProgressBar.Max
         '/////////////////////

            'Consulta Descripcion
            DescripKilates = SacaValorAccess("Kilatajes", "Descripcion", "WHERE ID= " & rc!Kilates)
            '/////////////////////

            'Consulta Descripcion
            DescripKilates = SacaValorAccess("Kilatajes", "Descripcion", "WHERE ID= " & rc!Kilates)
            '///////////////////////////'Consulta id
            Idks = 0
            idkt = 0
            If DescripKilates = "" Then
            
            
            Else
                Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                If Not rcidOro.EOF Then
                
                Else
                   idKil = Val(SacaValor("kilatajes", "MAX(ID)"))
                   T = SacaValorAccess("Kilatajes", "IDTipo", "WHERE Descripcion= '" & DescripKilates & "'")
                   dbDatos.Execute "insert into kilatajes (ID, Clave, Descripcion,IDTipo)" & _
                   "Values(" & idKil + 1 & "," & idKil + 1 & ",'" & DescripKilates & "'," & T & ")"
                   Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                End If
                
               Idks = rcidOro!ID
               idkt = rcidOro!IDTabla
            
            End If
           'Consulta id tipo
           If rc!Tipo = "DOCUMENTO" Then
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='DOCUMENTOS'")
           ElseIf rc!Tipo = "HERRAMIENTAS" Then
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='HERRAMIENTA'")
           Else
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='" & rc!Tipo & "'")
           End If
            '////////////
        dbDatos.Execute "INSERT INTO detallescompras(ID, IDCompra, Tipo, IDTablaTipo, Codigo, Descripcion, Kilates, IDTablaKilates, Cantidad, Peso, Costo, Precio ) VALUES (" & _
                        rc!ID & "," & rc!IDCompra & "," & rcTipo!ID & "," & rcTipo!IDTabla & ",'" & rc!Codigo & "','" & rc!Descripcion & "'," & Idks & "," & idkt & _
                        "," & rc!Cantidad & "," & rc!Peso & "," & rc!Costo & "," & rc!Precio & ")"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_CierreDiario()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String
  
    Sql = "SELECT * FROM CierreDiario order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Cierre Diario Registro " & ProgressBar.Value & "/" & ProgressBar.Max
      
        dbDatos.Execute "INSERT INTO CierreDiario(ID, Fecha, Sucursal, Cajero, Saldo, Debe, Haber, Efectivo, Ajuste, IDUsuario) VALUES (" & _
                        rc!ID & ",'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "','" & rc!Sucursal & "','" & rc!Cajero & "'," & rc!Saldo & _
                        "," & rc!Debe & "," & rc!Haber & "," & rc!Efectivo & "," & rc!Ajuste & ",1)"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_Boveda()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String
  
    Sql = "SELECT * FROM Boveda order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Boveda Registro " & ProgressBar.Value & "/" & ProgressBar.Max
      
        dbDatos.Execute "INSERT INTO Boveda( ID, Fecha, Folio, Cancelado, FechaMovimiento, Deposito, Concepto, Importe, IDUsuario, IDSucursal ) VALUES (" & _
                                rc!ID & ",'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "'," & rc!Folio & "," & _
                                 "0,'" & IIf(IsNull(rc!Fecha), Format(Now, "YYYY/MM/DD HH:MM:SS"), Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS")) & "'," & IIf(rc!Deposito, 1, 0) & ",'" & rc!Concepto & "'," & rc!Importe & ",1,101)"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_Bancos()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String
  
    Sql = "SELECT * FROM Bancos order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Bancos Registro " & ProgressBar.Value & "/" & ProgressBar.Max
      
        dbDatos.Execute "INSERT INTO Bancos(ID, Fecha, Folio, Cancelado, FechaMovimiento, Deposito, Concepto, Importe, IDUsuario, IDSucursal ) VALUES (" & _
                                rc!ID & ",'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "'," & rc!Folio & "," & _
                                "0,'" & IIf(IsNull(rc!Fecha), Format(Now, "YYYY/MM/DD HH:MM:SS"), Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS")) & "'," & IIf(rc!Deposito, 1, 0) & ",'" & rc!Concepto & "'," & rc!Importe & ",1,101)"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_Abonos()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String
  
    Sql = "SELECT * FROM Abonos order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Abonos Registro " & ProgressBar.Value & "/" & ProgressBar.Max
      
        dbDatos.Execute "INSERT INTO Abonos(ID, IDVenta, Fecha, Importe, Cancelado, FechaMovimiento, PC, IDUsuario, IDSucursal) VALUES (" & _
                                rc!ID & "," & rc!IDVenta & ",'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "'," & _
                                rc!Abono & ",0," & "'" & IIf(IsNull(rc!Fecha), Format(Now, "YYYY/MM/DD HH:MM:SS"), Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS")) & "','" & rc!PC & "',1,101)"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_Auxiliar()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Sql As String
  
    Sql = "SELECT * FROM auxiliar order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Auxiliar Registro " & ProgressBar.Value & "/" & ProgressBar.Max
      
        dbDatos.Execute "INSERT INTO Auxiliar(ID, Fecha, Hora, Movimiento, Concepto, Folio, Iniciales, Cuenta, Importe, Tipo, Serie, PC, Corte, IDUsuario, IDSucursal ) VALUES (" & _
                                rc!ID & ",'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "','" & IIf(IsNull(rc!Fecha), Format(Now, "HH:MM:SS"), Format(rc!Fecha, "HH:MM:SS")) & "'," & rc!Movimiento & ",'" & _
                                rc!Concepto & "'," & rc!Folio & "," & "'" & rc!Iniciales & "','" & _
                                rc!Cuenta & "'," & rc!Importe & "," & rc!Tipo & "," & rc!Serie & ",'" & rc!PC & "'," & rc!Corte & ",1,101)"
                          
       rc.MoveNext
   Wend
   rc.Close
   
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_Usuarios()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim strNombre As String
    Dim strApellido As String
    Dim strApePaterno As String
    Dim strApeMaterno As String
    Dim strIniciales As String
    Dim Espacios As Long, IDCliente As Long, IDCuenta As Long
    Dim Sql As String
    Dim IDNacionalidad As Long
    Dim Identificacion As String, NumeroIdentificacion As String
    Dim FecNac As String
    
    lblClientes.Visible = True
    
    
    Sql = "SELECT * From Usuarios order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Usuarios " & ProgressBar.Value & "/" & ProgressBar.Max
       
                                   
        dbDatos.Execute "insert into usuarios (ID, Nombre, Usuario, Contraseña) VALUES (" & rc!ID & ",'" & rc!Nombre & "','" & rc!Usuario & "','" & rc!contraseña & "')"
    
     rc.MoveNext
    Wend
   
    rc.Close
       
    Sql = "SELECT * From Parametros order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Parametros " & ProgressBar.Value & "/" & ProgressBar.Max
                
         
         dbDatos.Execute "INSERT INTO sucursales (ID,Clave,NombreSucursal,RazonSocial, NombreComercial, RFC, Direccion, Ciudad, Estado, Telefono, Cp, Email) Values (" & _
                        rc!ID & ",101,'" & rc!Nomcomercial & "','" & rc!RazonSocial & "','" & rc!Nomcomercial & "','" & rc!RFC & "','" & rc!Direccion & " " & rc!Colonia & "','" & rc!Ciudad & "','" & rc!Estado & "','" & rc!Telefono & "'," & rc!CP & ",'" & rc!CorreoElectronico & "')"
     rc.MoveNext
    Wend
   
    rc.Close
    Sql = "SELECT * From Folios order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
    dbDatos.Execute "Delete from Folios"
    While Not rc.EOF
        DoEvents

        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Folios " & ProgressBar.Value & "/" & ProgressBar.Max
        
       

         dbDatos.Execute "INSERT INTO Folios ( ID,  Folio, Serie ) Values (" & _
                        rc!ID & "," & rc!Folio & "," & rc!Serie & ")"
     rc.MoveNext
    Wend

    rc.Close
    
    Sql = "SELECT * From Movimientos order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If

    While Not rc.EOF
        DoEvents

        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Folios " & ProgressBar.Value & "/" & ProgressBar.Max
        
        dbDatos.Execute "Delete from Movimientos"

         dbDatos.Execute "INSERT INTO Movimientos ( ID,  Movimiento, FolioBancos, FolioGastos, FolioVentas, FolioDepositos, FolioTransferencias, FolioCompras, FolioSalidaInventario, FolioAjustes, FolioBoveda, FolioDivisas, FolioNotas, FolioAutorizacion, FolioInventario, FolioTraspasos, FolioBovedaDivisas, FolioAvisosLavado, FolioDevolucion, FolioPasesInventario, FolioReImpresiones, FolioRenovacionesForaneas, FolioFacturas ) Values (" & _
                        rc!ID & "," & rc!Movimiento & "," & rc!FolioBancos & "," & rc!FolioGastos & "," & rc!FolioVentas & "," & rc!FolioDepositos & "," & rc!FolioTransferencias & ",1," & rc!FolioSalidaInventario & "," & rc!FolioAjustes & "," & rc!FolioBoveda & ",1," & rc!FolioNotas & ",1,1,1,1,1,1,1,1,1,1)"
     rc.MoveNext
    Wend

    rc.Close
Error:
   Maneja_Error Err
End Sub

Private Sub Migrar_Clientes()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim strNombre As String
    Dim strApellido As String
    Dim strApePaterno As String
    Dim strApeMaterno As String
    Dim strIniciales As String
    Dim Espacios As Long, IDCliente As Long, IDCuenta As Long
    Dim Sql As String
    Dim IDNacionalidad As Long
    Dim Identificacion As String, NumeroIdentificacion As String
    Dim FecNac As String
    
    lblClientes.Visible = True

    Sql = "SELECT * FROM Clientes order by ID"
             
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
    
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Clientes Registro " & ProgressBar.Value & "/" & ProgressBar.Max
      
        'If Val(SacaValor("Clientes", "ID", "WHERE concat(Nombre,' ',Apellido)='" & Replace(Replace(Trim(QuitarEspacios(rc!Nomb)), "'", ""), "\", "") & "'")) = 0 Then
            If (Trim(QuitarEspacios(rc!Nombre & " " & rc!Apellido)) & "") <> "" Then
                'Separo el Nombre y el Apellido del Cliente
                Espacios = ContarEspacios(Trim(QuitarEspacios(rc!Nombre & " " & rc!Apellido)))
                'Los pongo al reves por que asi vienen en Excel
                strNombre = Trim(Replace(Replace(ObtieneNombres(Espacios, Trim(QuitarEspacios(rc!Nombre & " " & rc!Apellido))), "'", ""), "\", ""))  'Replace(Replace(Trim(rc!Nombre & " " & rc!Apellido), "'", ""), "\", "")
                'strApellido = Replace(Replace(ObtieneApellidos(Espacios, Trim(rc!Nombre)), "'", ""), "\", "") 'Replace(Replace(Trim(rc!ApellidoPaterno & " " & rc!ApellidoMaterno), "'", ""), "\", "")
                Espacios = ContarEspacios(Trim(rc!Apellido))
                strApellido = QuitarEspacios(Trim$(Mid$(QuitarEspacios(rc!Apellido), IIf(Espacios = 0, 1, Espacios))))
                strIniciales = Replace(Iniciales(strNombre, strApellido), "\", "")
                'strNombre = Replace(Replace(Trim(rc!Nombre), "'", ""), "\", "")
                IDNacionalidad = GetIDNacionalidad("Mexicana")
                'Identificacion =
                If IsNull(rc!Identificacion) Then
                    Identificacion = ""
                Else
                    Identificacion = rc!Identificacion
                End If '"Credencial para votar"
                If IsNull(rc!NumeroIdentificacion) Then
                    NumeroIdentificacion = ""
                Else
                    NumeroIdentificacion = Quitar_Letras(rc!NumeroIdentificacion)
                End If
                
                If InStr(1, rc!Identificacion, "IFE") > 0 Then
                    Identificacion = "CREDENCIAL PARA VOTAR"
                End If
                
                Espacios = ContarEspacios(Trim(strApellido))
                'Los pongo al reves por que asi vienen en Excel
                strApePaterno = Trim(Replace(Replace(ObtieneNombres(Espacios, Trim(strApellido)), "'", ""), "\", ""))
                strApeMaterno = Trim$(Mid$(strApellido, Len(strApePaterno) + 1))
                
                
                If IsNull(rc!FecNac) Then
                Else
                    FecNac = rc!FecNac
                End If
                
                Dim FechaAltaRazonSocial As Date
                
             Dim Sexo As Integer
            If IsNull(rc!Sexo) Then
            Sexo = 0
            Else
            Sexo = rc!Sexo
            End If

             dbDatos.Execute "INSERT INTO Clientes(ID, Nombre, Apellido, Iniciales, Identificacion, NumeroIdentificacion, FecRegistro, Tel, Direccion, Municipio, Estado, IDUsuario, Colonia, FecNac, ApellidoPaterno, ApellidoMaterno, " & _
                                                    "IDMedio, Boletas, Notas, CP, Rfc, Sexo, IDSucursal) VALUES (" & _
                                                    rc!ID & ",'" & strNombre & "','" & strApellido & "','" & Replace$(strIniciales, "'", "") & "','" & Identificacion & "','" & _
                                                    NumeroIdentificacion & "','" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & "'" & Replace$(Trim(rc!Tel) & "", "'", "") & "','" & _
                                                    Mid(rc!Direccion, 1, 120) & "','" & rc!Municipio & "','" & rc!Estado & "',1, '" & rc!Colonia & "','" & Format(FecNac, "YYYY/MM/DD HH:MM:SS") & "','" & strApePaterno & "','" & strApeMaterno & "'," & _
                                                    rc!IDMedio & "," & rc!Boletas & ",'" & IIf(IsNull(rc!Notas), " ", rc!Notas) & "'," & IIf(IsNull(rc!CP), 0, rc!CP) & ",'" & rc!RFC & "'," & rc!Sexo & ",101)"
            End If
       'End If
        rc.MoveNext
   Wend
   
   rc.Close
   
Error:
   Maneja_Error Err
End Sub

Private Sub Migrar_Clientes_Tarjetas()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim strNombre As String
    Dim strApellido As String
    Dim strApePaterno As String
    Dim strApeMaterno As String
    Dim strIniciales As String
    Dim Espacios As Long, IDCliente As Long, IDCuenta As Long
    Dim Sql As String
    Dim IDNacionalidad As Long
    Dim Identificacion As String, NumeroIdentificacion As String
    Dim FecNac As String
    
    lblTarjetas.Visible = True
    
    Sql = "SELECT * FROM Clientes WHERE NumTarjeta<>''"
    
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
   
    If Not rc.EOF Then
       ProgressBar.Max = rc.RecordCount
       ProgressBar.Min = 0
       ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Tarjetas de Puntos Registro " & ProgressBar.Value & "/" & ProgressBar.Max
        
        'If Left(rc!Nombre, 4) = "IRMA" Then
        '    Sql = Sql
        'End If
        
        If Val(SacaValor("Clientes", "ID", "WHERE concat(Nombre,' ',Apellido)='" & Replace(Replace(Trim(QuitarEspacios(rc!Nombre)), "'", ""), "\", "") & "'")) = 0 Then
            If (Trim(QuitarEspacios(rc!Nombre)) & "") <> "" Then
                'Separo el Nombre y el Apellido del Cliente
                Espacios = ContarEspacios(Trim(QuitarEspacios(rc!Nombre)))
                'Los pongo al reves por que asi vienen en Excel
                strNombre = Trim(Replace(Replace(ObtieneNombres(Espacios, Trim(QuitarEspacios(rc!Nombre))), "'", ""), "\", ""))  'Replace(Replace(Trim(rc!Nombre), "'", ""), "\", "")
                'strApellido = Replace(Replace(ObtieneApellidos(Espacios, Trim(rc!Nombre)), "'", ""), "\", "") 'Replace(Replace(Trim(rc!ApellidoPaterno & " " & rc!ApellidoMaterno), "'", ""), "\", "")
                strApellido = QuitarEspacios(Trim$(Mid$(QuitarEspacios(rc!Nombre), Len(strNombre) + 1)))
                strIniciales = Replace(Iniciales(strNombre, strApellido), "\", "")
                'strNombre = Replace(Replace(Trim(rc!Nombre), "'", ""), "\", "")
                IDNacionalidad = GetIDNacionalidad("Mexicana")
                Identificacion = "Credencial para votar"
                NumeroIdentificacion = SacaValorAccess("Boletas", "Identificacion", "WHERE Nombre='" & Replace(Replace(Trim(rc!Nombre), "'", "\'"), "\", "\\") & "'")
                'If InStr(1, rc!Identificacion, "IFE") > 0 Then
                '    Identificacion = "CREDENCIAL PARA VOTAR"
                'End If
                
                Espacios = ContarEspacios(Trim(strApellido))
                'Los pongo al reves por que asi vienen en Excel
                strApePaterno = Trim(Replace(Replace(ObtieneNombres(Espacios, Trim(strApellido)), "'", ""), "\", ""))
                strApeMaterno = Trim$(Mid$(strApellido, Len(strApePaterno) + 1))
                
                
                FecNac = "NULL"
                
                dbDatos.Execute "INSERT INTO Clientes (Nombre, Apellido, Iniciales, Identificacion, NumeroIdentificacion, FecRegistro, Tel, Direccion,Municipio,Estado,IDUsuario,IDNacionalidad,Colonia,FecNac,ApellidoPaterno,ApellidoMaterno) VALUES (" & _
                                "'" & strNombre & "','" & strApellido & "','" & Replace$(strIniciales, "'", "") & "','" & Identificacion & "','" & _
                                NumeroIdentificacion & "','" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & "'" & Replace$(Trim(rc!Telefonos) & "", "'", "") & "','" & _
                                Mid(rc!Direccion, 1, 70) & "','" & rc!Ciudad & "','" & rc!Estado & "'," & frmMDI.IDUsuario & "," & IDNacionalidad & ",'" & rc!Colonia & "'," & FecNac & ",'" & strApePaterno & "','" & strApeMaterno & "')"
                                
                IDCliente = Val(SacaValor("Clientes", "MAX(ID)", ""))
                
                If Val(Trim$(rc!IDTarjeta & "")) > 0 Then
                    dbDatos.Execute "INSERT INTO asignaciontarjetas (Fecha, NumeroTarjeta, IDTarjeta, IDCliente, IDUsuario, PC, Puntos) VALUES (" & _
                                    "'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "','" & Trim$(rc!IDTarjeta & "") & "',1," & IDCliente & "," & _
                                    frmMDI.IDUsuario & ",'" & NombrePc & "'," & rc!Puntos & ")"
                    
                    If rc!Puntos > 0 Then
                        IDCuenta = Val(SacaValor("asignaciontarjetas", "MAX(ID)", ""))
                        
                        dbDatos.Execute "INSERT INTO MovimientosPuntos (Fecha,IDTarjeta,TipoMovimiento,Concepto,Folio,Cargo,Abono,Importe,PC,IDUsuario) VALUES ('" & _
                              Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & IDCuenta & ",2,'Saldo Importacion',0," & rc!Puntos & ",0,0,'" & Nombre_Pc & "'," & frmMDI.IDUsuario & ")"
                    End If
                End If
            End If
        End If
        rc.MoveNext
   Wend
   
   rc.Close
   
Error:
   Maneja_Error Err
End Sub

Private Function GetIDNacionalidad(Nacionalidad As String) As Long
   On Error GoTo Error
   Dim ID As Long
   
   If Nacionalidad = "" Then
      ID = 0
   Else
   
      ID = Val(SacaValor("Nacionalidad", "ID", " WHERE Nacionalidad='" & Nacionalidad & "'"))
      
      If ID = 0 Then
         dbDatos.Execute "INSERT INTO Nacionalidad(Nacionalidad) VALUES ('" & Nacionalidad & "')"
         ID = Val(SacaValor("Nacionalidad", "ID", " WHERE Nacionalidad='" & Nacionalidad & "'"))
      End If
      
   End If
   
   GetIDNacionalidad = ID
Error:
   Maneja_Error Err
End Function

Private Sub Migrar_Contratos()
    On Error GoTo Error
    Dim FechaOriginal As Date, FechaContrato As Date, Vencimiento As Date
    Dim Interes As Double, Seguro As Double, Almacenaje As Double, Iva As Double, CAT As Double
    Dim Prestamo As Currency, Avaluo As Currency, PrestamoInicial As Currency
    Dim Origen As Integer, PSuc As Integer, Plazos As Integer, Almoneda, IDEntrada, TipoEntrada, Destino As Integer, Pagado, ExiIDC As Integer
    Dim IDCliente As Long, DiasAcumulados As Long, IDPrenda As Long, IDEmpeno As Long
    Dim FechaUltimoPago As Date, FechaAlmoneda, FechaMovimiento As String
    Dim Bodega As String, Zona As String, Sql As String, NumContrato As String, sTipoInteres As String, sTipoTasa As String
    
    
    Dim rc As New ADODB.Recordset
    lblContratos.Visible = True
    
    CompraOro = 0
    CompraPlata = 0
    CompraMisc = 0
    CompraAuto = 0
    If chMyBD Then
        Sql = "SELECT B.* FROM Empeño as B ORDER BY B.ID"
    Else
        Sql = "SELECT B.* FROM empeno as B WHERE  Cancelado=0 AND Pagado=0 ORDER BY B.Fecha, B.NumContrato, B.Folio"
    End If

    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockOptimistic
   
    If Not rc.EOF Then
       ProgressBar.Value = 0
       ProgressBar.Value = 0
       ProgressBar.Max = rc.RecordCount
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Contratos Registro " & ProgressBar.Value & "/" & ProgressBar.Max
        If chMyBD Then
                    
                Origen = rc!Origen
                FechaOriginal = SacaValorAccess("Empeño", "min(Fecha)", "WHERE NumContrato= " & rc!NumContrato & " And Serie =" & rc!Serie)
                
                PrestamoInicial = rc!PrestamoInicial
                FechaContrato = rc!Fecha
                If IsNull(rc!FechaMovimiento) Then
                    'FechaUltimoPago = ""
                Else
                    FechaUltimoPago = rc!FechaMovimiento
                End If
                
                Vencimiento = rc!Vencimiento
                Prestamo = rc!Prestamo
                Avaluo = rc!Avaluo
                PSuc = Suc_Posicion(frmMDI.IDSucursal)
                Interes = rc!Tasa ' Suc_Int(PSuc, 2)
                Seguro = rc!Seguro 'Suc_Int(PSuc, 3)
                Almacenaje = rc!GTOAlmacenaje 'Suc_Int(PSuc, 4)
                Iva = rc!Iva 'Suc_Int(PSuc, 5)
                CAT = 0 'rc!CAT 'Suc_Int(PSuc, 6)
                NumContrato = CInt(rc!NumContrato)
                If PrestamoInicial = 0 Then PrestamoInicial = Prestamo
                Plazos = rc!VenPeriodo
                If Plazos > 3 Then Plazos = 3
                sTipoInteres = "TRADICIONAL" 'Trim(Replace(Replace(Trim(QuitarEspacios(rc!TipoInteres)), "'", ""), "\", ""))
                sTipoTasa = Trim(Replace(Replace(Trim(QuitarEspacios(rc!TipoInteres)), "'", ""), "\", ""))
               
                If IsNull(rc!FechaMovimiento) Then
                     FechaMovimiento = "NULL"
                Else
                     FechaMovimiento = rc!FechaMovimiento
                End If
                               
                 IDEntrada = SacaValorAccess("DetallesEntradaInventario", "IDEntrada", "WHERE IDEmpeño= " & rc!ID)
                
                If IDEntrada = "" Then
                    Almoneda = 0
                    Pagado = Val(rc!Pagado)
                    FechaAlmoneda = "NULL"
                    Destino = rc!Destino
                
                Else
                    Almoneda = 1
                    Pagado = Val(rc!Pagado)
                    FechaAlmoneda = SacaValorAccess("EntradaInventario", "Fecha", "WHERE ID= " & IDEntrada)
                    Destino = 4
                End If
               
                IDCliente = rc!IDCliente 'Val(SacaValor("Clientes", "ID", "WHERE concat(Nombre,' ',Apellido)='" & Replace(Replace(Trim(QuitarEspacios(rc!Responsable)), "'", ""), "\", "") & "'"))
                ExiIDC = IIf(SacaValor("Clientes", "ID", "WHERE ID= " & rc!IDCliente) = "", 0, SacaValor("Clientes", "ID", "WHERE ID= " & rc!IDCliente))
                If ExiIDC = 0 Then
                          dbDatos.Execute "INSERT INTO Clientes(ID, Nombre, Apellido, Iniciales, Direccion, Municipio, IDUsuario, Colonia, IDSucursal) VALUES (" & _
                                            rc!IDCliente & ",'" & rc!Nombre & "','" & rc!Apellidos & "','" & rc!Iniciales & "','" & rc!Domicilio & "','" & rc!Municipio & "',1, '" & rc!Colonia & "',101)"
                End If
                'If IDCliente > 0 Then
                IDPrenda = IDOro
                    
                dbDatos.Execute "INSERT INTO empeno (ID,Fecha,FechaOriginal,IDTipoPrenda,Movimiento,NumContrato,Folio,Prestamo,Avaluo,Origen,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Tasa," & _
                              "Almacenaje,Seguro,Iva,Cat,Venperiodo,Periodo,Tipointeres,TipoTasa,IDSucursal,IDUsuario,PrestamoInicial,Destino,Almoneda,Pagado,FechaAlmoneda,FolioDestino,FechaMovimiento," & _
                              "Cancelado,Pago,Intereses,Operacion,Comision,ImporteMoratorios,NumBolsa,FolioNota,Efectivo) VALUES " & _
                              "(" & rc!ID & ",'" & Format(FechaContrato, "YYYY/MM/DD") & "','" & Format(FechaOriginal, "YYYY/MM/DD") & "'," & IDPrenda & "," & rc!Movimiento & "," & Val(rc!NumContrato) & "," & _
                              Val(rc!Folio) & "," & Val(ConvMoneda(Prestamo)) & "," & Val(ConvMoneda(rc!Avaluo)) & "," & Origen & ",'" & Format(Vencimiento, "YYYY/MM/DD") & "'," & _
                              rc!FolioOrigen & "," & rc!Serie & ",'" & rc!PC & "'," & rc!IDCliente & "," & Interes & "," & Almacenaje & "," & Seguro & "," & Iva & "," & CAT & "," & rc!VenPeriodo & "," & rc!Periodo & _
                              ",'" & sTipoInteres & "' , '" & sTipoTasa & "' ,101,1," & Val(ConvMoneda(PrestamoInicial)) & ",'" & Destino & "'," & _
                              Almoneda & "," & IIf(rc!Pagado, 1, 0) & "," & IIf(FechaAlmoneda = "NULL", FechaAlmoneda, "'" & Format(FechaAlmoneda, "YYYY/MM/DD") & "'") & "," & Val(rc!FolioDestino) & " , " & IIf(FechaMovimiento = "NULL", FechaMovimiento, "'" & Format(FechaMovimiento, "YYYY/MM/DD HH:MM:SS") & "'") & _
                              "," & IIf(rc!Cancelado, 1, 0) & "," & Val(rc!Pago) & "," & Val(rc!Intereses) & "," & Val(rc!Operacion) & "," & Val(rc!GTOVenta) & "," & Val(rc!Moratorios) & "," & Val(rc!Bolsa) & "," & rc!FolioNota & "," & Val(rc!Efectivo) & ")"
                    
                    IDEmpeno = Val(SacaValor("Empeno", "MAX(ID)"))
                    Avaluo = 0
                    PrestamoInicial = Prestamo
                    'Migrar_Prendas IDEmpeno, Val(NumContrato), IDPrenda, rc!Folio, Avaluo, PrestamoInicial, IIf(rc!Serie = 1, True, False)
                    Migrar_Prendas IIf(Destino = 4, 5, 0), IDEmpeno, rc!ID, Val(NumContrato), IDPrenda, rc!Folio, Avaluo, PrestamoInicial, IIf(rc!Serie = 2, True, False)
                   ' If Avaluo = 0 Then Avaluo = rc!Avaluo
                    Avaluo = rc!Avaluo

        End If
        rc.MoveNext
    Wend
   
   rc.Close
   
Error:
   Maneja_Error Err
   Set rc = Nothing
End Sub

Private Sub Migrar_Prendas(ByVal Destino As Long, ByVal IDEmpeno As Long, ByVal IDEmpeno_ As Long, ByVal NumContrato As Long, ByRef IDPrenda As Long, ByVal Folio As String, ByRef Avaluo As Currency, ByRef Prestamo As Currency, ByVal Auto As Boolean)
    On Error GoTo Error
    Dim Tipo As Integer, Partida As Integer
    Dim CodigoPrenda As String, Marca As String, Modelo As String, Categoria As String, Serie As String, Color As String, Tamano As String
    Dim TipoPrenda As Long, Kilates As Long
    Dim Peso As Double, PesoPiedra As Double, PesoTotal As Double, Porcentaje As Double
    Dim Precio As Currency, Avaluo2 As Currency, PrestamoDet As Currency
    Dim rc As New ADODB.Recordset
    Dim Sql, DescripKilates As String
    Dim rcDescripcio As ADODB.Recordset
    Dim rcidOro As ADODB.Recordset
    Dim rcidTipoOro As ADODB.Recordset
    Dim rcTipo As ADODB.Recordset
    Dim idKil, Idks, idkt, T As Integer
  
    
    If Not Auto Then
        Sql = "SELECT Count(ID) as Art, sum(Avaluo) as Prestamo FROM DetallesEmpeño WHERE IDEmpeño=" & IDEmpeno_ & " Group By IDEmpeño"
              
        rc.Open Sql, dbCnn, adOpenForwardOnly, adLockOptimistic
        If Not rc.EOF Then
            PrestamoDet = rc!Prestamo
        End If
        rc.Close
    Else
        PrestamoDet = Prestamo
    End If
    
    If PrestamoDet > 0 Then
    Porcentaje = (PrestamoDet - Prestamo) / PrestamoDet
    End If
    Sql = "SELECT Det.* FROM DetallesEmpeño" & IIf(Auto, "Autos", "") & " as Det WHERE Det.IDEmpeño=" & IDEmpeno_
    
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockOptimistic
    'rc.Open Sql, dbCnn, adOpenForwardOnly, adLockOptimistic
    Partida = 1
    While Not rc.EOF
        If Not Auto Then
            CodigoPrenda = CreaCodigoBarras(Trim(Format(frmMDI.IDSucursal, "000")), Format(ENTRADAEMPENO, "00"), Trim(NumContrato), Partida)
            '/////////////////////

            'Consulta Descripcion
            DescripKilates = SacaValorAccess("Kilatajes", "Descripcion", "WHERE ID= " & rc!Kilates)
            '///////////////////////////'Consulta id
            Idks = 0
            idkt = 0
            If DescripKilates = "" Then
            
            
            Else
                Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                If Not rcidOro.EOF Then
                
                Else
                   idKil = Val(SacaValor("kilatajes", "MAX(ID)"))
                   T = SacaValorAccess("Kilatajes", "IDTipo", "WHERE Descripcion= '" & DescripKilates & "'")
                   dbDatos.Execute "insert into kilatajes (ID, Clave, Descripcion,IDTipo)" & _
                   "Values(" & idKil + 1 & "," & idKil + 1 & ",'" & DescripKilates & "'," & T & ")"
                   Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                End If
                
               Idks = rcidOro!ID
               idkt = rcidOro!IDTabla
            
            End If
           'Consulta id tipo
           If rc!Tipo = "DOCUMENTO" Or rc!Tipo = "DEOCUMENTOS" Then
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='DOCUMENTOS'")
           ElseIf rc!Tipo = "HERRAMIENTAS" Then
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='HERRAMIENTA'")
           Else
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='" & rc!Tipo & "'")
           End If
            '////////////
            
            If rc!Kilates <> 0 Then
                
                Peso = Val(rc!Peso & "")
                PesoTotal = Val(rc!Peso & "")
            Else
                IDPrenda = IDElectronicos
                Peso = 0
                PesoPiedra = 0
                PesoTotal = 0
'                Marca = SacaValorAccess("detallesempeno", "Marca", "WHERE IDEmpeno=" & IDEmpeno_)
'                Modelo = SacaValorAccess("detallesempeno", "Modelo", "WHERE IDEmpeno=" & IDEmpeno_)
                'Serie = rc!Serie
                Categoria = ""
            End If
         
'            If Not rcidOro.EOF Then
'                Kilates = rcidOro!ID
'            Else
'                Kilates = 0
'            End If
            
            If rc!Avaluo > 0 Then
                Precio = rc!Avaluo - (rc!Avaluo * Porcentaje)
            End If
            Avaluo2 = rc!Avaluo
            
            If Avaluo2 = 0 Then Avaluo2 = Precio
            dbDatos.Execute "INSERT INTO detallesempeno (ID, IDEmpeno, Tipo, IDTablaTipo, Cantidad, Codigo, Articulo, Peso, Estado, Kilates, IDTablaKilates, Avaluo, Prestamo, Origen, Destino,Almoneda ) VALUES (" & _
                          rc!ID & "," & IDEmpeno & "," & rcTipo!ID & "," & rcTipo!IDTabla & "," & rc!Cantidad & ",'" & CodigoPrenda & "','" & Replace(Trim(rc!Articulo & ""), "'", "") & "'," & ConvMoneda(Peso) & ",'" & rc!Estado & "'," & Idks & "," & idkt & "," & _
                           ConvMoneda(Avaluo2) & "," & ConvMoneda(Precio) & ",1," & Destino & "," & IIf(Destino = 5, 1, 0) & ")"
            Avaluo = Avaluo + Avaluo2
            Partida = Partida + 1
        Else
            Dim marcaModelo As String
            IDPrenda = IDAuto
            'Marca = Trim(SacaValorAccess("detallesempenoautos", "Marca", "WHERE IDEmpeno=" & IDEmpeno_))
            'Modelo = Trim(SacaValorAccess("detallesempenoautos", "Modelo", "WHERE IDEmpeno=" & IDEmpeno_))
            marcaModelo = Trim(SacaValorAccess("DetallesEmpeñoAutos", "marcaymodelo", "WHERE IDEmpeño=" & IDEmpeno_))
        
            dbDatos.Execute "INSERT INTO detallesempenoautos (IDEmpeno,MarcayModelo,Marca,Modelo,Año,Color,Placas,Factura,Agencia,NumTarjetacircu,NumMotor," & _
                            "SerieChasis,VIN,RePuVe,Kms,Gas,Aseguradora,Poliza,FechaVenci,Tipo,Observaciones) VALUES (" & _
                            IDEmpeno & ",'" & marcaModelo & "','" & Marca & "','" & Modelo & "'," & Val(rc!Año) & ",'" & rc!Color & "','" & _
                            rc!Placas & "','" & rc!Factura & "','" & rc!Agencia & "','" & rc!NumTarjetaCircu & "','" & rc!NumMotor & "','" & rc!SerieChasis & _
                            "','','','" & rc!Kms & "','" & rc!Gas & "','" & rc!Aseguradora & "','" & rc!Poliza & "',NULL,'','')"
        End If
        rc.MoveNext
    Wend
    Prestamo = PrestamoDet
    rc.Close
    Exit Sub
    
Error:
   Maneja_Error Err
End Sub
Private Sub Migrar_DetEntra()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim rcA As New ADODB.Recordset
    Dim Tipo As Integer
    Dim CodigoPrenda As String
    Dim Peso As Double, PesoPiedra As Double, PrecioVitrina As Double
    Dim Marca As String, Modelo As String, Serie As String, Color As String, Tamano As String, Observaciones As String, Sql, DescripKilates As String
    Dim Kilates As Long, TipoPrenda As Long, IDInventario As Long, NumContrato As Long, FolioInventario As Long
    Dim rcDescripcio As ADODB.Recordset
    Dim rcidOro As ADODB.Recordset
    Dim rcidTipoOro As ADODB.Recordset
    Dim rcTipo As ADODB.Recordset
    Dim idKil, Idks, idkt, T As Integer
    Dim T_E, T_S, C_T_S, C_T_E As Integer
    
    lblVitrina.Visible = True
   
   Sql = "SELECT * FROM DetallesEntradaInventario order by ID asc"

   rc.CursorLocation = adUseClient
   rc.Open Sql, dbCnn, adOpenDynamic, adLockOptimistic

    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If

   
   While Not rc.EOF
    DoEvents
   
      
        CostoInventarioOro = 0
        CostoInventario = 0

        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Inventario Detalle Registro " & ProgressBar.Value & "/" & ProgressBar.Max
            '/////////////////////

            'Consulta Descripcion
            DescripKilates = SacaValorAccess("Kilatajes", "Descripcion", "WHERE ID= " & rc!Kilates)
            '///////////////////////////'Consulta id
            Idks = 0
            idkt = 0
            If DescripKilates = "" Then
            
            
            Else
                Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                If Not rcidOro.EOF Then
                
                Else
                   idKil = Val(SacaValor("kilatajes", "MAX(ID)"))
                   T = SacaValorAccess("Kilatajes", "IDTipo", "WHERE Descripcion= '" & DescripKilates & "'")
                   dbDatos.Execute "insert into kilatajes (ID, Clave, Descripcion,IDTipo)" & _
                   "Values(" & idKil + 1 & "," & idKil + 1 & ",'" & DescripKilates & "'," & T & ")"
                   Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                End If
                
               Idks = rcidOro!ID
               idkt = rcidOro!IDTabla
            
            End If
           'Consulta id tipo
           If rc!Tipo = "DOCUMENTO" Then
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='DOCUMENTOS'")
           ElseIf rc!Tipo = "HERRAMIENTAS" Then
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='HERRAMIENTA'")
           ElseIf rc!Tipo = "AUTOMOVIL" Then
                Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='" & rc!Tipo & "'")
                If Not rcTipo.EOF Then
                Else
                'idKil = Val(SacaValor("kilatajes", "MAX(ID)"))
                'T = SacaValorAccess("Kilatajes", "IDTipo", "WHERE Descripcion= '" & DescripKilates & "'")
                dbDatos.Execute "insert into Tipo (IDTabla, Descripcion, Kilataje, Peso, Ordenamiento, IdTipoGarantia, IdTipoBienes, IdTipoUnidad, Actualizar )" & _
                "Values(" & Val(SacaValor("Tipo", "MAX(IDTabla)")) + 1 & ",'" & rc!Tipo & "',0,0,1,13,0,0,0 )"
                Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='" & rc!Tipo & "'")
                End If
           Else
            Set rcTipo = dbDatos.Execute("SELECT ID, IDTabla FROM tipo WHERE Descripcion='" & rc!Tipo & "'")
           End If
            '////////////
            Peso = Val(rc!Peso)
             PesoPiedra = 0 'Val(rc!PesoPiedras)
            'Kilates = ID10k
           
            If rc!Kilates > 0 Then
                Tipo = rcTipo!ID
                TipoPrenda = IDOroGeneral
            Else
               Tipo = rcTipo!ID 'IDElectronicos
               TipoPrenda = IDPrendaGeneral
               Kilates = 0
            End If
               
            Dim Contrato As Long
            Dim Partida As Integer
           
'            Contrato = Val(rc!Original & "")
'            Partida = 1
'            PrecioVitrina = Val(rc!Precio & "") / 1.16 'Redondear(Val(rc!Precio & "") / 1.16, 0)
            T_E = SacaValorAccess("EntradaInventario", "TipoEntrada", "WHERE ID= " & rc!IDEntrada)
            
            If T_E = 1 Then
               T_E = 5
            End If
            
            
            
           If SacaValorAccess("DetallesVentasCA", "IDArticulo", "WHERE IDArticulo= " & rc!ID) <> "" Then
              T_S = 1
           ElseIf SacaValorAccess("DetallesSalida", "IDArticulo", "WHERE IDArticulo= " & rc!ID) <> "" Then
              T_S = 4
           End If
            
            CodigoPrenda = rc!Codigo 'CreaCodigoBarras(Trim(Format(frmMDI.IDSucursal, "000")), Format(ENTRADAMIGRACION, "00"), CStr(Contrato), Partida)
            'NumContrato = Contrato
            dbDatos.Execute "INSERT INTO detallesentradainventario (ID,Tipo,TipoPrenda,Cantidad,IDEntrada,IDEmpeno,Codigo,Descripcion,Costo,Precio,Kilates,Peso,PesoPiedras,TipoEntrada,PrecioVitrina,tipoSalida) VALUES (" & _
                                                                    rc!ID & "," & rcTipo!ID & "," & TipoPrenda & "," & rc!Cantidad & "," & rc!IDEntrada & "," & rc!IDEmpeño & ",'" & CodigoPrenda & "','" & Replace(Replace(Trim(rc!Descripcion & ""), "'", ""), "\", "") & "'," & _
                                                                    ConvMoneda(Val(rc!Costo & "")) & "," & ConvMoneda(Val(rc!Precio & "")) & "," & _
                                                                    Idks & "," & Peso & "," & PesoPiedra & "," & T_E & "," & ConvMoneda(Val(rc!Precio & "")) & "," & T_S & ")"
      rc.MoveNext
 Wend
    'rc.Close
Error:
    Maneja_Error Err
    
    Set rc = Nothing
End Sub
Private Sub Migrar_Vitrina()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim rcA As New ADODB.Recordset
    Dim Tipo As Integer
    Dim CodigoPrenda As String
    Dim Peso As Double, PesoPiedra As Double, PrecioVitrina As Double
    Dim Marca As String, Modelo As String, Serie As String, Color As String, Tamano As String, Observaciones As String, Sql As String
    Dim Kilates As Long, TipoPrenda As Long, IDInventario As Long, NumContrato As Long, FolioInventario As Long
    Dim T_E As Integer
   
   lblVitrina.Visible = True
   
   Sql = "SELECT * FROM entradainventario order by ID asc"

   rc.CursorLocation = adUseClient
   rc.Open Sql, dbCnn, adOpenDynamic, adLockOptimistic
   
   If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If

   
   While Not rc.EOF
    DoEvents
   
    If rc.RecordCount <> 0 Then
        
        CostoInventarioOro = 0
        CostoInventario = 0

         ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Inventario Registro " & ProgressBar.Value & "/" & ProgressBar.Max

        'Saco el Folio
        
        FolioInventario = Regresa_Movimiento(False, "FolioInventario")
        Regresa_Movimiento True, "FolioInventario"
        
        If rc!TipoEntrada = 1 Then
            T_E = 4
        Else
            T_E = rc!TipoEntrada
        End If
        dbDatos.Execute "INSERT INTO entradainventario (ID, Fecha,Folio,TipoEntrada,IDUsuario,IDSucursal) VALUES (" & _
                        rc!ID & ",'" & Format(rc!Fecha, "YYYY/MM/DD HH:MM:SS") & "'," & rc!Folio & "," & T_E & ", 1,101)"
        
     End If
      rc.MoveNext
 Wend
    'rc.Close
Error:
    Maneja_Error Err
    
    Set rc = Nothing
End Sub

Private Sub Migrar_Apartados()
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim rcArt As New ADODB.Recordset
    Dim Sql As String
    Dim FolioApartados As Long
    Dim FechaHora As Date
    Dim IDVenta As Long
    Dim IDCliente As Long
    Dim Vencimiento As Date
    Dim Vencimiento_ As String
    Dim Fecha As Date
    Dim crImporte As Currency
    Dim crEfectivo As Currency
    Dim Movimiento As Long
    Dim IDAbono As Long
    Dim IDVendedor As Long
    Dim IDUsuario As Long
    Dim CuentaInventarioVenta, DescripKilates As String
    Dim CuentaInventarioApartados As String
    Dim rcDescripcio As ADODB.Recordset
    Dim rcidOro As ADODB.Recordset
    Dim rcidTipoOro As ADODB.Recordset
    Dim rcTipo As ADODB.Recordset
    Dim idKil, Idks, idkt, T, ExiIDC As Integer
    
    Sql = "SELECT * FROM VentasCA order by ID"
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockOptimistic
   
    If Not rc.EOF Then
        ProgressBar.Max = rc.RecordCount
        ProgressBar.Min = 0
        ProgressBar.Value = 0
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Apartados y Ventas Registro " & ProgressBar.Value & "/" & ProgressBar.Max
        
'        FolioApartados = Regresa_Movimiento(False, "FolioVentas")
'        Regresa_Movimiento True, "FolioVentas"
        
        Fecha = rc!Fecha
         Vencimiento_ = ""
        If IsNull(rc!Vencimiento) Then
            Vencimiento_ = "NULL"
        Else
          Vencimiento = rc!Vencimiento
        End If
        
        FechaHora = Now
        IDVendedor = 1 'frmMDI.IDUsuario
        IDUsuario = 1 'frmMDI.IDUsuario
        IDCliente = rc!IDCliente 'Val(SacaValor("Clientes", "ID", "WHERE concat(Nombre,' ',Apellido)='" & Replace(Replace(Trim(QuitarEspacios(rc!Comprador)), "'", ""), "\", "") & "'"))
        crImporte = rc!Total ' / (1.16)
        crEfectivo = rc!Total
        
       ' Movimiento = Val(SacaValor("detallesventas", "Count(ID)")) + 1
      
'        If IDCliente = 0 Then IDCliente = Clientes_Agregar_Apartado(rc!Comprador)
'
'        If IDCliente > 0 Then
        
         ExiIDC = IIf(SacaValor("Clientes", "ID", "WHERE ID= " & rc!IDCliente) = "", 0, SacaValor("Clientes", "ID", "WHERE ID= " & rc!IDCliente))
         If ExiIDC = 0 Then
            If rc!IDCliente > 0 Then
                  dbDatos.Execute "INSERT INTO Clientes(ID, Nombre, Apellido, Direccion,  IDSucursal) VALUES (" & _
                                   rc!IDCliente & ",'" & rc!Nombre & "','" & rc!Apellido & "','" & rc!Direccion & "',101)"
            End If
        End If
      
            dbDatos.Execute "INSERT INTO ventas(Fecha,Vencimiento,Folio,IVA,Descuento,Cancelado,Total,Apartado,PC,IDCliente,IDUsuario,IDSucursal,IDVendedor,Efectivo,Pagado,ID) VALUES ('" & _
                             Format(Fecha, "YYYY/MM/DD HH:MM:SS") & "'," & IIf(Vencimiento_ = "NULL", Vencimiento_, "'" & Format(Vencimiento, "YYYY/MM/DD") & "'") & "," & rc!Folio & "," & rc!Iva & "," & rc!Descuento & "," & IIf(rc!Cancelado, 1, 0) & "," & _
                            ConvMoneda(crImporte) & "," & IIf(rc!Apartado, 1, 0) & ",'" & rc!PC & "'," & IDCliente & ",1,101," & IDVendedor & "," & _
                            ConvMoneda(crEfectivo) & "," & IIf(rc!Pagado, 1, 0) & "," & rc!ID & ")"
                            
            'Tomo el ID de la venta
            IDVenta = rc!ID 'SacaValor("ventas", "MAX(ID)")
            
            rcArt.Open "SELECT * FROM DetallesVentasCA WHERE IDVenta=" & rc!ID, dbCnn, adOpenDynamic, adLockOptimistic
            
            While Not rcArt.EOF
            '/////////////////////

            'Consulta Descripcion
            DescripKilates = SacaValorAccess("Kilatajes", "Descripcion", "WHERE ID= " & rcArt!Kilates)
            '///////////////////////////'Consulta id
            Idks = 0
            idkt = 0
            If DescripKilates = "" Then
            
            
            Else
                Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                If Not rcidOro.EOF Then
                
                Else
                   idKil = Val(SacaValor("kilatajes", "MAX(ID)"))
                   T = SacaValorAccess("Kilatajes", "IDTipo", "WHERE Descripcion= '" & DescripKilates & "'")
                   dbDatos.Execute "insert into kilatajes (ID, Clave, Descripcion,IDTipo)" & _
                   "Values(" & idKil + 1 & "," & idKil + 1 & ",'" & DescripKilates & "'," & T & ")"
                   Set rcidOro = dbDatos.Execute("SELECT ID, IDTabla FROM kilatajes WHERE Descripcion='" & DescripKilates & "'")
                End If
                
               Idks = rcidOro!ID
               idkt = rcidOro!IDTabla
            
            End If
 
            '////////////

                dbDatos.Execute "INSERT INTO detallesventas (ID,IDVenta,Codigo,Articulo,Kilates,Peso,Costo,Precio,Intereses,Almacenaje,Seguro,IDArticulo) VALUES (" & _
                                rcArt!ID & "," & IDVenta & ",'" & rcArt!Codigo & "','" & rcArt!Articulo & "'," & _
                                 Idks & "," & rcArt!Peso & "," & rcArt!Costo & "," & _
                                 rcArt!Precio & "," & rcArt!Intereses & "," & rcArt!Almacenaje & "," & rcArt!Seguro & "," & rcArt!IDArticulo & ")"
                rcArt.MoveNext
              Wend

            rcArt.Close

    
        rc.MoveNext
   Wend
         
'         dbDatos.Execute "delete from folios"
'
'         dbDatos.Execute "delete from movimientos"
'
'         dbDatos.Execute "insert into folios(ID, Folio, Serie) SELECT ID, Folio, Serie FROM basedatos_old.folios"
'
'         dbDatos.Execute "insert into movimientos(ID, Movimiento, FolioBancos, FolioGastos, FolioVentas, FolioDepositos, FolioTransferencias, FolioCompras, FolioSalidaInventario, FolioAjustes, FolioBoveda, FolioDivisas, FolioNotas," & _
'                         "Fecha, FolioAutorizacion, FolioInventario, FolioTraspasos, FolioBovedaDivisas, FolioReImpresiones, FolioAvisosLavado )" & _
'                         "SELECT ID, Movimiento, FolioBancos, FolioGastos, FolioVentas, FolioDepositos, FolioTransferencias, FolioCompras, FolioSalidaInventario, FolioAjustes, FolioBoveda, FolioDivisas, FolioNotas," & _
'                         "Fecha, FolioAutorizacion, FolioInventario, FolioTraspasos, FolioBovedaDivisas, FolioReImpresiones, FolioAvisosLavado FROM basedatos_old.movimientos"
'
'
Error:
   Maneja_Error Err
End Sub

Private Function Quitar_Letras(ByRef Dato As String) As String
    Dim Cadena As String
    Dim x As Integer
    
    Cadena = ""
    For x = 1 To Len(Dato)
        If IsNumeric(Mid$(Dato, x, 1)) Then
            Cadena = Cadena & Mid$(Dato, x, 1)
        End If
    Next x
    Quitar_Letras = Cadena
End Function

Private Function Suc_Posicion(ByRef Sucursal As Integer) As Integer
    Dim x As Integer
    
    Suc_Posicion = 0
    For x = 1 To 12
        If Suc_Int(x, 1) = Sucursal Then
            Suc_Posicion = x
            Exit For
        End If
    Next x
End Function

Private Function QuitarEspacios(ByVal sString As String) As String
    Do Until InStr(1, sString, "  ") = 0
        sString = Replace(sString, "  ", " ")
        DoEvents
    Loop
    QuitarEspacios = sString
End Function

Private Function Clientes_Agregar_Apartado(Nombre As String) As Long
    'On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim strNombre As String
    Dim strApellido As String
    Dim strApePaterno As String
    Dim strApeMaterno As String
    Dim strIniciales As String
    Dim Espacios As Long, IDCliente As Long, IDCuenta As Long
    Dim Sql As String
    Dim IDNacionalidad As Long
    Dim Identificacion As String, NumeroIdentificacion As String
    Dim FecNac As String
    
    'Separo el Nombre y el Apellido del Cliente
    Espacios = ContarEspacios(Trim(QuitarEspacios(Nombre)))
    'Los pongo al reves por que asi vienen en Excel
    strNombre = Trim(Replace(Replace(ObtieneNombres(Espacios, Trim(QuitarEspacios(Nombre))), "'", ""), "\", ""))
    strApellido = QuitarEspacios(Trim$(Mid$(QuitarEspacios(Nombre), Len(strNombre) + 1)))
    strIniciales = Replace(Iniciales(strNombre, strApellido), "\", "")
    IDNacionalidad = GetIDNacionalidad("Mexicana")
    Identificacion = "Credencial para votar"
    NumeroIdentificacion = ""
    
    Espacios = ContarEspacios(Trim(strApellido))
    strApePaterno = Trim(Replace(Replace(ObtieneNombres(Espacios, Trim(strApellido)), "'", ""), "\", ""))
    strApeMaterno = Trim$(Mid$(strApellido, Len(strApePaterno) + 1))
    
    FecNac = "NULL"
    
    dbDatos.Execute "INSERT INTO Clientes (Nombre, Apellido, Iniciales, Identificacion, NumeroIdentificacion, FecRegistro, Tel, Direccion,Municipio,Estado,IDUsuario,IDNacionalidad,Colonia,FecNac,ApellidoPaterno,ApellidoMaterno) VALUES (" & _
                    "'" & strNombre & "','" & strApellido & "','" & Replace$(strIniciales, "'", "") & "','" & Identificacion & "','" & _
                    NumeroIdentificacion & "','" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & "'','','',''," & frmMDI.IDUsuario & "," & _
                    IDNacionalidad & ",''," & FecNac & ",'" & strApePaterno & "','" & strApeMaterno & "')"
                    
    IDCliente = Val(SacaValor("Clientes", "MAX(ID)", ""))
   
    Clientes_Agregar_Apartado = IDCliente
Error:
    Maneja_Error Err
End Function

Private Function Clientes_Agregar_Desempeño(Folio As Long) As Long
    'On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim strNombre As String
    Dim strApellido As String
    Dim strApePaterno As String
    Dim strApeMaterno As String
    Dim strIniciales As String
    Dim Espacios As Long, IDCliente As Long, IDCuenta As Long
    Dim Sql As String
    Dim IDNacionalidad As Long
    Dim Identificacion As String, NumeroIdentificacion As String
    Dim FecNac As String
    
    Sql = "SELECT DISTINCT B.Nombre as Nom,B.Identificacion,B.FechaPago,C.* FROM Boletas as B LEFT JOIN Clientes as C ON C.Nombre=B.Nombre WHERE Folio=" & Folio
    
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockBatchOptimistic
   
   
    'Separo el Nombre y el Apellido del Cliente
    Espacios = ContarEspacios(Trim(QuitarEspacios(rc!Nom)))
    'Los pongo al reves por que asi vienen en Excel
    strNombre = Trim(Replace(Replace(ObtieneNombres(Espacios, Trim(QuitarEspacios(rc!Nom))), "'", ""), "\", ""))  'Replace(Replace(Trim(rc!Nombre), "'", ""), "\", "")
    'strApellido = Replace(Replace(ObtieneApellidos(Espacios, Trim(rc!Nombre)), "'", ""), "\", "") 'Replace(Replace(Trim(rc!ApellidoPaterno & " " & rc!ApellidoMaterno), "'", ""), "\", "")
    strApellido = QuitarEspacios(Trim$(Mid$(QuitarEspacios(rc!Nom), Len(strNombre) + 1)))
    strIniciales = Replace(Iniciales(strNombre, strApellido), "\", "")
    'strNombre = Replace(Replace(Trim(rc!Nombre), "'", ""), "\", "")
    IDNacionalidad = GetIDNacionalidad("Mexicana")
    Identificacion = "Credencial para votar"
    NumeroIdentificacion = Quitar_Letras(rc!Identificacion)
    
    Espacios = ContarEspacios(Trim(strApellido))
    'Los pongo al reves por que asi vienen en Excel
    strApePaterno = Trim(Replace(Replace(ObtieneNombres(Espacios, Trim(strApellido)), "'", ""), "\", ""))
    strApeMaterno = Trim$(Mid$(strApellido, Len(strApePaterno) + 1))
    
    
    FecNac = "NULL"
    
    dbDatos.Execute "INSERT INTO Clientes (Nombre, Apellido, Iniciales, Identificacion, NumeroIdentificacion, FecRegistro, Tel, Direccion,Municipio,Estado,IDUsuario,IDNacionalidad,Colonia,FecNac,ApellidoPaterno,ApellidoMaterno) VALUES (" & _
                    "'" & strNombre & "','" & strApellido & "','" & Replace$(strIniciales, "'", "") & "','" & Identificacion & "','" & _
                    NumeroIdentificacion & "','" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & "'" & Replace$(Trim(rc!Telefonos & "") & "", "'", "") & "','" & _
                    Mid(rc!Direccion & "", 1, 70) & "','" & rc!Ciudad & "','" & rc!Estado & "'," & frmMDI.IDUsuario & "," & IDNacionalidad & ",'" & rc!Colonia & "'," & FecNac & ",'" & strApePaterno & "','" & strApeMaterno & "')"
                    
    IDCliente = Val(SacaValor("Clientes", "MAX(ID)", ""))
    
    If Val(Trim$(rc!IDTarjeta & "")) > 0 Then
        dbDatos.Execute "INSERT INTO asignaciontarjetas (Fecha, NumeroTarjeta, IDTarjeta, IDCliente, IDUsuario, PC, Puntos) VALUES (" & _
                        "'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "','" & Trim$(rc!IDTarjeta & "") & "',1," & IDCliente & "," & _
                        frmMDI.IDUsuario & ",'" & NombrePc & "'," & rc!Puntos & ")"
        
        If rc!Puntos > 0 Then
            IDCuenta = Val(SacaValor("asignaciontarjetas", "MAX(ID)", ""))
            
            dbDatos.Execute "INSERT INTO MovimientosPuntos (Fecha,IDTarjeta,TipoMovimiento,Concepto,Folio,Cargo,Abono,Importe,PC,IDUsuario) VALUES ('" & _
                  Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & IDCuenta & ",2,'Saldo Importacion',0," & rc!Puntos & ",0,0,'" & Nombre_Pc & "'," & frmMDI.IDUsuario & ")"
        End If
    End If
   
    rc.Close
    Clientes_Agregar_Desempeño = IDCliente
Error:
   Maneja_Error Err
End Function

Private Sub Migrar_Contratos_Desempeñados()
    'On Error GoTo Error
    Dim FechaOriginal As Date, FechaContrato As Date, Vencimiento As Date, FechaMov As Date
    Dim Interes As Double, Seguro As Double, Almacenaje As Double, Iva As Double, CAT As Double
    Dim Prestamo As Currency, Avaluo As Currency, PrestamoInicial As Currency
    Dim Origen As Integer, PSuc As Integer, Plazos As Integer, Almoneda As Integer, Pagado As Integer, Destino As Integer
    Dim IDCliente As Long, DiasAcumulados As Long, IDPrenda As Long, IDEmpeno As Long
    Dim FechaUltimoPago As Date, FechaAlmoneda As String
    Dim Bodega As String, Zona As String, Sql As String, NumContrato As String
    
    Dim rc As New ADODB.Recordset
    
    'Set rc = Nothing
    'Set rc = New ADODB.Recordset
   
    lblDesemp.Visible = True
    
    CompraOro = 0
    CompraPlata = 0
    CompraMisc = 0
    CompraAuto = 0
    
    Sql = "SELECT B.* TPago FROM empeno as B order by B.ID"
    
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockOptimistic
   
    If Not rc.EOF Then
       ProgressBar.Value = 0
       ProgressBar.Value = 0
       ProgressBar.Max = rc.RecordCount
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Contratos Registro " & ProgressBar.Value & "/" & ProgressBar.Max
           
        'If Val(SacaValorAccess("BoletasDetalle", "ID", "WHERE Folio=" & rc!FolioOriginal)) <> 0 Then
        If Val(SacaValorAccess("detallesempeno", "ID", "WHERE idEmpeno=" & rc!ID)) <> 0 Or Val(SacaValorAccess("detallesempeno", "id", "WHERE idEmpeno=" & rc!ID)) <> 0 Then
            If rc!Folio = rc!FolioOrigen Then
                Origen = OD_EMPENO
            Else
                Origen = OD_REFRENDO
            End If
            Destino = D_DESEMPEÑO
            
            FechaOriginal = rc!FechaOriginal ' IIf(SacaValorAccess("Boletas", "Emision", "WHERE Folio=" & rc!FolioOriginal) = "", rc!Emision, SacaValorAccess("Boletas", "Emision", "WHERE Folio=" & rc!FolioOriginal))
            PrestamoInicial = rc!PrestamoInicial 'Val(SacaValorAccess("Boletas", "Capital", "WHERE Folio=" & rc!FolioOriginal))
            FechaContrato = rc!Fecha
            FechaUltimoPago = rc!FechaMovimiento
            Vencimiento = rc!Vencimiento
            Prestamo = rc!Prestamo
            Avaluo = rc!Avaluo
            PSuc = Suc_Posicion(frmMDI.IDSucursal)
            Interes = rc!Tasa 'Suc_Int(PSuc, 2)
            Seguro = rc!Seguro 'Suc_Int(PSuc, 3)
            Almacenaje = rc!Almacenaje 'Suc_Int(PSuc, 4)
            Iva = rc!Iva ' Suc_Int(PSuc, 5)
            CAT = rc!CAT ' Suc_Int(PSuc, 6)
            NumContrato = CInt(rc!NumContrato)
            If PrestamoInicial = 0 Then PrestamoInicial = Prestamo
            Plazos = rc!VenPeriodo
            If Plazos > 3 Then Plazos = 3
            
            Almoneda = 0
            Pagado = 1
            FechaAlmoneda = "NULL"
            
            IDCliente = Val(SacaValor("Clientes", "ID", "WHERE concat(Nombre,' ',Apellido)='" & Replace(Replace(Trim(QuitarEspacios(rc!Responsable)), "'", ""), "\", "") & "'"))
            
            If IDCliente = 0 Then IDCliente = Clientes_Agregar_Desempeño(rc!Folio)
            
            'If IDCliente > 0 Then
                IDPrenda = IDOro
                
                dbDatos.Execute "INSERT INTO empeno (Fecha,FechaOriginal,FechaMovimiento,IDTipoPrenda,Movimiento,NumContrato,Folio,Prestamo,Avaluo,Origen,Destino,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Tasa," & _
                              "Almacenaje,Seguro,Iva,Cat,Venperiodo,Periodo,Tipointeres,TipoTasa,IDSucursal,IDUsuario,PrestamoInicial,Caja,Cajon,Almoneda,Pagado,FechaAlmoneda,Pago,Intereses,ImporteIva,ImporteMoratorios) VALUES " & _
                              "('" & Format(FechaContrato, "YYYY/MM/DD") & "','" & Format(FechaOriginal, "YYYY/MM/DD") & "','" & Format(FechaUltimoPago, "YYYY/MM/DD") & "'," & IDPrenda & "," & ProgressBar.Value & "," & Val(rc!Folio) & "," & _
                              NumContrato & "," & Val(ConvMoneda(Prestamo)) & "," & Val(ConvMoneda(rc!Avaluo)) & "," & Origen & "," & Destino & ",'" & Format(Vencimiento, "YYYY/MM/DD") & "'," & _
                              NumContrato & "," & SERIE_A & ",'" & NombrePc & "'," & IDCliente & "," & Interes & "," & Almacenaje & "," & Seguro & "," & Iva & "," & CAT & "," & Suc_Plazo(PSuc, Plazos) & ",1" & _
                              ",'TRADICIONAL','DIARIA'," & frmMDI.IDSucursal & "," & frmMDI.IDUsuario & "," & Val(ConvMoneda(PrestamoInicial)) & ",'',''," & Almoneda & "," & Pagado & "," & FechaAlmoneda & "," & Val(ConvMoneda(Prestamo)) & "," & rc!Comision & "," & rc!Iva & "," & rc!Recargos & ")"
                
                IDEmpeno = Val(SacaValor("Empeno", "MAX(ID)"))
                Avaluo = 0
                PrestamoInicial = Prestamo
                'Migrar_Prendas IDEmpeno, rc!ID, Val(NumContrato), IDPrenda, rc!FolioOriginal, Avaluo, PrestamoInicial, IIf(rc!Categoria = 7, True, False)
                'If Avaluo = 0 Then Avaluo = rc!Avaluo
                Avaluo = rc!Valor
                
                dbDatos.Execute "UPDATE empeno SET Serie=" & IIf(rc!Categoria = 7, SERIE_B, IIf(IDPrenda = IDElectronicos, SERIE_D, SERIE_A)) & ", IDTipoPrenda=" & IIf(rc!Categoria = 7, 0, IDPrenda) & ", Avaluo=" & IIf(rc!Categoria = 7, rc!Avaluo, Avaluo) & ",PrestamoInicial=" & PrestamoInicial & " WHERE ID=" & IDEmpeno
                
'            Else
'                IDCliente = IDCliente
'            End If
        End If
      
        rc.MoveNext
    Wend
    
    dbDatos.Execute "UPDATE Folios SET Folio=" & Val(SacaValor("Empeno", "MAX(NumContrato)")) + 1 & " WHERE Serie=" & SERIE_A
    dbDatos.Execute "UPDATE Folios SET Folio=" & Val(SacaValor("Empeno", "MAX(NumContrato)")) + 1 & " WHERE Serie=" & SERIE_B
   
    rc.Close
   
Error:
   Maneja_Error Err
   Set rc = Nothing
End Sub

Private Sub Migrar_Contratos_Enajenados()
    'On Error GoTo Error
    Dim FechaOriginal As Date, FechaContrato As Date, Vencimiento As Date, FechaMov As Date
    Dim Interes As Double, Seguro As Double, Almacenaje As Double, Iva As Double, CAT As Double
    Dim Prestamo As Currency, Avaluo As Currency, PrestamoInicial As Currency
    Dim Origen As Integer, PSuc As Integer, Plazos As Integer, Almoneda As Integer, Pagado As Integer, Destino As Integer
    Dim IDCliente As Long, DiasAcumulados As Long, IDPrenda As Long, IDEmpeno As Long
    Dim FechaUltimoPago As Date, FechaAlmoneda As Date
    Dim Bodega As String, Zona As String, Sql As String, NumContrato As String
    
    Dim rc As New ADODB.Recordset
    
    'Set rc = Nothing
    'Set rc = New ADODB.Recordset
   
    lblEnajenado.Visible = True
    
    CompraOro = 0
    CompraPlata = 0
    CompraMisc = 0
    CompraAuto = 0
    
    Sql = "SELECT B.* FROM Boletas as B where B.FechaEnajenacion <> #01/01/2100# order by B.Folio"
    
    rc.CursorLocation = adUseClient
    rc.Open Sql, dbCnn, adOpenDynamic, adLockOptimistic
   
    If Not rc.EOF Then
       ProgressBar.Value = 0
       ProgressBar.Value = 0
       ProgressBar.Max = rc.RecordCount
    End If
   
    While Not rc.EOF
        DoEvents
        
        ProgressBar.Value = ProgressBar.Value + 1
        lblRegistro.Caption = "Importando Contratos Registro " & ProgressBar.Value & "/" & ProgressBar.Max
           
        'If Val(SacaValorAccess("BoletasDetalle", "ID", "WHERE Folio=" & rc!FolioOriginal)) <> 0 Then
        If Val(SacaValorAccess("BoletasDetalle", "ID", "WHERE Folio=" & rc!FolioOriginal)) <> 0 Or Val(SacaValorAccess("BoletasDetalleAutos", "Folio", "WHERE Folio=" & rc!FolioOriginal)) <> 0 Then
            If rc!Folio = rc!FolioOriginal Then
                Origen = OD_EMPENO
            Else
                Origen = OD_REFRENDO
            End If
            Destino = D_ALMONEDA
            
            FechaOriginal = IIf(SacaValorAccess("Boletas", "Emision", "WHERE Folio=" & rc!FolioOriginal) = "", rc!Emision, SacaValorAccess("Boletas", "Emision", "WHERE Folio=" & rc!FolioOriginal))
            PrestamoInicial = Val(SacaValorAccess("Boletas", "Capital", "WHERE Folio=" & rc!FolioOriginal))
            FechaContrato = rc!Emision
            FechaUltimoPago = rc!FechaEnajenacion
            Vencimiento = rc!Vencimiento
            Prestamo = rc!Capital
            Avaluo = rc!Valor
            PSuc = Suc_Posicion(frmMDI.IDSucursal)
            Interes = Suc_Int(PSuc, 2)
            Seguro = Suc_Int(PSuc, 3)
            Almacenaje = Suc_Int(PSuc, 4)
            Iva = Suc_Int(PSuc, 5)
            CAT = Suc_Int(PSuc, 6)
            NumContrato = CInt(rc!Folio)
            If PrestamoInicial = 0 Then PrestamoInicial = Prestamo
            Plazos = rc!plazo
            If Plazos > 3 Then Plazos = 3
            
            Almoneda = 1
            Pagado = 1
            FechaAlmoneda = rc!FechaEnajenacion
            
            IDCliente = Val(SacaValor("Clientes", "ID", "WHERE concat(Nombre,' ',Apellido)='" & Replace(Replace(Trim(QuitarEspacios(rc!Nombre)), "'", ""), "\", "") & "'"))
            
            If IDCliente = 0 Then IDCliente = Clientes_Agregar_Desempeño(rc!Folio)
            
            'If IDCliente > 0 Then
                IDPrenda = IDOro
                
                dbDatos.Execute "INSERT INTO empeno (Fecha,FechaOriginal,FechaMovimiento,IDTipoPrenda,Movimiento,NumContrato,Folio,Prestamo,Avaluo,Origen,Destino,Vencimiento,FolioOrigen,Serie,PC,IDCliente,Tasa," & _
                              "Almacenaje,Seguro,Iva,Cat,Venperiodo,Periodo,Tipointeres,TipoTasa,IDSucursal,IDUsuario,PrestamoInicial,Caja,Cajon,Almoneda,Pagado,FechaAlmoneda,Pago,Intereses,ImporteIva,ImporteMoratorios) VALUES " & _
                              "('" & Format(FechaContrato, "YYYY/MM/DD") & "','" & Format(FechaOriginal, "YYYY/MM/DD") & "','" & Format(FechaUltimoPago, "YYYY/MM/DD") & "'," & IDPrenda & "," & ProgressBar.Value & "," & Val(rc!Folio) & "," & _
                              NumContrato & "," & Val(ConvMoneda(Prestamo)) & "," & Val(ConvMoneda(rc!Avaluo)) & "," & Origen & "," & Destino & ",'" & Format(Vencimiento, "YYYY/MM/DD") & "'," & _
                              NumContrato & "," & SERIE_A & ",'" & NombrePc & "'," & IDCliente & "," & Interes & "," & Almacenaje & "," & Seguro & "," & Iva & "," & CAT & "," & Suc_Plazo(PSuc, Plazos) & ",1" & _
                              ",'TRADICIONAL','DIARIA'," & frmMDI.IDSucursal & "," & frmMDI.IDUsuario & "," & Val(ConvMoneda(PrestamoInicial)) & ",'',''," & Almoneda & "," & Pagado & ",'" & Format(FechaAlmoneda, "YYYY/MM/DD") & "'," & Val(ConvMoneda(Prestamo)) & "," & rc!Comision & "," & rc!Iva & "," & rc!Recargos & ")"
                
                IDEmpeno = Val(SacaValor("Empeno", "MAX(ID)"))
                Avaluo = 0
                PrestamoInicial = Prestamo
                'Migrar_Prendas IDEmpeno, Val(NumContrato), IDPrenda, rc!FolioOriginal, Avaluo, PrestamoInicial, IIf(rc!Categoria = 7, True, False)
                'If Avaluo = 0 Then Avaluo = rc!Avaluo
                Avaluo = rc!Valor
                
                dbDatos.Execute "UPDATE empeno SET Serie=" & IIf(rc!Categoria = 7, SERIE_B, IIf(IDPrenda = IDElectronicos, SERIE_D, SERIE_A)) & ", IDTipoPrenda=" & IIf(rc!Categoria = 7, 0, IDPrenda) & ", Avaluo=" & IIf(rc!Categoria = 7, rc!Avaluo, Avaluo) & ",PrestamoInicial=" & PrestamoInicial & " WHERE ID=" & IDEmpeno
                
'            Else
'                IDCliente = IDCliente
'            End If
        End If
      
        rc.MoveNext
    Wend
    
    dbDatos.Execute "UPDATE Folios SET Folio=" & Val(SacaValor("Empeno", "MAX(NumContrato)")) + 1 & " WHERE Serie=" & SERIE_A
    dbDatos.Execute "UPDATE Folios SET Folio=" & Val(SacaValor("Empeno", "MAX(NumContrato)")) + 1 & " WHERE Serie=" & SERIE_B
   
    rc.Close
   
Error:
   Maneja_Error Err
   Set rc = Nothing
End Sub


