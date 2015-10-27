VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCapturaImagen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Fotografía"
   ClientHeight    =   4350
   ClientLeft      =   3405
   ClientTop       =   2730
   ClientWidth     =   4935
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCapturaImagen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   4935
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
      Height          =   3660
      Left            =   30
      ScaleHeight     =   3630
      ScaleWidth      =   4830
      TabIndex        =   0
      Top             =   120
      Width           =   4860
   End
   Begin DevPowerFlatBttn.FlatBttn cmdTomar 
      Height          =   375
      Left            =   1275
      TabIndex        =   1
      Top             =   3900
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Tomar"
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
      Object.ToolTipText     =   ""
      Picture         =   "frmCapturaImagen.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCerrar 
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3900
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
      Picture         =   "frmCapturaImagen.frx":00BB
   End
   Begin DevPowerFlatBttn.FlatBttn cmdIniciar 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   3900
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Iniciar"
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
      Picture         =   "frmCapturaImagen.frx":060D
   End
   Begin DevPowerFlatBttn.FlatBttn cmdGuardar 
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   3900
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
      Picture         =   "frmCapturaImagen.frx":0AE2
   End
   Begin VB.Image ImgFoto 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   3660
      Left            =   30
      Stretch         =   -1  'True
      Top             =   120
      Visible         =   0   'False
      Width           =   4860
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
Attribute VB_Name = "frmCapturaImagen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim hwdc As Long
Dim startcap As Boolean
Dim Ini As Boolean, Direccion As String, Nombre As String

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdGuardar_Click()
Dim x As String

x = Path & "\Fotos\" & Nombre & ".jpg"
If Dir(x) <> "" Then
    If MsgBox("La foto del cliente ya se encuentra registrada" & Chr(13) & "Desea remplazarla ?? ", vbYesNo + vbQuestion, "Tomar Foto") = vbYes Then
        Kill x
    Else
        Exit Sub
    End If
End If

Direccion = Path & "\Fotos\" & Nombre
SendMessage hwdc, WM_CAP_GET_FRAME, 0&, 0&
SendMessage hwdc, WM_CAP_COPY, 0&, 0&
Picture1.Picture = ImgFoto.Picture
   
    Call SavePicture(ImgFoto.Picture, Direccion & ".jpg")  ' el directorio  y el nombre de la imagen que quieres ir a poner
    DoEvents
    
    MsgBox "Fotografia guardada con éxito !!", vbInformation, "Tomar Foto"
'''''Parar

Picture1.Visible = False
End Sub

Private Sub cmdIniciar_Click()
Dim temp As Long

If Ini = True Then
    Parar
    cmdIniciar.Enabled = True
    cmdIniciar.Caption = "Iniciar"
    Ini = False
End If

ImgFoto.Visible = False
cmdGuardar.Enabled = False
Me.Picture1.Visible = True
hwdc = capCreateCaptureWindow("RCR Soluciones", ws_child Or ws_visible, 0, 0, 320, 240, Picture1.hWnd, 0)

'Dixanta Vision System
If (hwdc <> 0) Then
    temp = SendMessage(hwdc, wm_cap_driver_connect, 0, 0)
    temp = SendMessage(hwdc, wm_cap_set_preview, 1, 0)
    temp = SendMessage(hwdc, WM_CAP_SET_PREVIEWRATE, 30, 0)
    startcap = True
    cmdTomar.Enabled = True
    mnuPropiedades.Enabled = True
    cmdIniciar.Enabled = False
    'mensaje.Visible = False
Else
    MsgBox "Webcam no encontrada !!", vbInformation, "Tomar foto"
End If
End Sub

Private Sub cmdTomar_Click()

On Error GoTo Error

    SendMessage hwdc, WM_CAP_GET_FRAME, 0&, 0&
    SendMessage hwdc, WM_CAP_COPY, 0&, 0&
    ImgFoto.Picture = Clipboard.GetData
    ImgFoto.Height = 3660
    ImgFoto.Width = 4860
    
    Parar
    ImgFoto.Visible = True
    cmdGuardar.Enabled = True
    Picture1.Visible = False
    Exit Sub
    
Error: Err
    Maneja_Error Err
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Parar
    DestroyWindow hwdc
End Sub

Private Sub for_res_Click()
Dim temp As Long
 If startcap = True Then
  temp = SendMessage(hwdc, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
End If
End Sub

Private Sub for_col_Click()
Dim temp As Long
 If startcap = True Then
  temp = SendMessage(hwdc, WM_VIDEOFORMAT_COLOR, 0&, 0&)
End If
End Sub

Private Sub compre_Click()
Dim temp As Long
 If startcap = True Then
  temp = SendMessage(hwdc, WM_VIDEOFORMAT_COMPRESION, 0&, 0&)
End If
End Sub

Private Sub salir_Click()
Unload Me
End Sub

Sub Parar()
Dim temp As Long

If startcap = True Then
    temp = SendMessage(hwdc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
    startcap = False
    cmdIniciar.Enabled = True
    cmdTomar.Enabled = False
    mnuPropiedades.Enabled = False
End If
End Sub

Public Function Ver(Cliente As String, Opcion As Integer)

    Screen.MousePointer = vbHourglass
    
    Inicializar
    cmdGuardar.Enabled = False
    Nombre = Cliente
    
    If Dir(Path & "\Fotos\" & Cliente & ".jpg") <> "" Then
        Picture1.Visible = False
        ImgFoto.Height = 3660
        ImgFoto.Width = 4860
        ImgFoto.Picture = LoadPicture(Path & "\Fotos\" & Cliente & ".jpg")
        ImgFoto.Visible = True
    End If
    
    Screen.MousePointer = vbDefault
    
    If Opcion = 3 Or Opcion = 4 Then cmdIniciar.Visible = False: cmdTomar.Visible = False: cmdGuardar.Visible = False
    cmdTomar.Enabled = False
    Me.Show vbModal
End Function

Sub Inicializar()
    Height = 5130
    Width = 5025
    Ini = False
    Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 2
End Sub
