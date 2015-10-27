VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCapturaImagenBiometrico 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fotografía"
   ClientHeight    =   4470
   ClientLeft      =   3405
   ClientTop       =   2730
   ClientWidth     =   6345
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCapturaImagenBiometrico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4470
   ScaleWidth      =   6345
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      FillColor       =   &H8000000F&
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
      Height          =   2460
      Left            =   3240
      ScaleHeight     =   2430
      ScaleWidth      =   2910
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2940
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      FillColor       =   &H8000000F&
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
      Height          =   2460
      Left            =   150
      ScaleHeight     =   2430
      ScaleWidth      =   2910
      TabIndex        =   0
      Top             =   480
      Width           =   2940
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCerrar 
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3660
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
      Picture         =   "frmCapturaImagenBiometrico.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdIniciar 
      Height          =   375
      Left            =   150
      TabIndex        =   2
      Top             =   3060
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Capturar Foto con Cámara Web"
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
      Picture         =   "frmCapturaImagenBiometrico.frx":055E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdIniciar2 
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3060
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Capturar Foto con Cámara Web"
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
      Picture         =   "frmCapturaImagenBiometrico.frx":0A33
   End
   Begin DevPowerFlatBttn.FlatBttn cmdHuellaDig 
      Height          =   375
      Left            =   4080
      TabIndex        =   7
      Top             =   3660
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "Huella Digital"
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
      Picture         =   "frmCapturaImagenBiometrico.frx":0F08
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Identificación"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   3240
      TabIndex        =   6
      Top             =   120
      Width           =   2940
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H8000000D&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   390
      Left            =   150
      TabIndex        =   5
      Top             =   120
      Width           =   2940
   End
   Begin VB.Image ImgIFE 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2460
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Image ImgFoto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2460
      Left            =   150
      Stretch         =   -1  'True
      Top             =   480
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.Menu mnuPropiedades 
      Caption         =   "&Propiedades"
      Enabled         =   0   'False
      Begin VB.Menu for_res 
         Caption         =   "Resolución"
      End
      Begin VB.Menu for_col 
         Caption         =   "Colores"
      End
      Begin VB.Menu compre 
         Caption         =   "Compresión"
      End
   End
End
Attribute VB_Name = "frmCapturaImagenBiometrico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hwdc As Long, hwdc2 As Long
Dim startcap As Boolean, Ini As Boolean, Direccion As String, Nombre As String, IDCte As Long

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Public Function Ver(IDCliente As Long, Cliente As String, Opcion As Integer)

    Screen.MousePointer = vbHourglass
    
    Inicializar
'    cmdGuardar.Enabled = False
    Nombre = Cliente
    IDCte = IDCliente
    
    If Dir(Path & "\Fotos\" & Cliente & ".jpg") <> "" Then
        Picture1.Visible = False
        'ImgFoto.Height = 3660
        'ImgFoto.Width = 4860
        ImgFoto.Picture = LoadPicture(Path & "\Fotos\" & Cliente & ".jpg")
        ImgFoto.Visible = True
    End If
    
    If Dir(Path & "\Fotos\" & Cliente & "-CRED.jpg") <> "" Then
        Picture2.Visible = False
        'ImgFoto.Height = 3660
        'ImgFoto.Width = 4860
        ImgIFE.Picture = LoadPicture(Path & "\Fotos\" & Cliente & "-CRED.jpg")
        ImgIFE.Visible = True
    End If
    
    Screen.MousePointer = vbDefault
    
    'If Opcion = 3 Or Opcion = 4 Then cmdIniciar.Visible = False: cmdTomar.Visible = False: cmdGuardar.Visible = False
    'cmdTomar.Enabled = False
    Me.Show vbModal
End Function

Sub Inicializar()
    'Height = 5130
    'Width = 5025
    Ini = False
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
    
    
    'Set Biometric = New Biometrico.FingerPrint
    'Set Biometric.Conexion = dbDatos
    'Biometric.Initialize
    
    
End Sub

Private Sub cmdHuellaDig_Click()
    'Biometric.Captura IDCte
        
            
    'If Biometric.Valida(IDCte) = True Then
    '    MsgBox "Cliente Correcto"
    'End If
        
End Sub

Private Sub cmdIniciar_Click()
    frmCapturaImagen.Ver Nombre, 1
    
    If Dir(Path & "\Fotos\" & Nombre & ".jpg") <> "" Then
        Picture1.Visible = False
        ImgFoto.Picture = LoadPicture(Path & "\Fotos\" & Nombre & ".jpg")
        ImgFoto.Visible = True
    End If
    
End Sub

Private Sub cmdIniciar2_Click()
    frmCapturaImagen.Ver CStr(Nombre & "-CRED"), 1
    
    If Dir(Path & "\Fotos\" & Nombre & "-CRED.jpg") <> "" Then
        Picture2.Visible = False
        ImgIFE.Picture = LoadPicture(Path & "\Fotos\" & Nombre & "-CRED.jpg")
        ImgIFE.Visible = True
    End If
    
End Sub
