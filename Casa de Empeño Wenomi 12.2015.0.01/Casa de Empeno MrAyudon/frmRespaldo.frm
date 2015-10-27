VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRespaldo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Respaldo"
   ClientHeight    =   1425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7800
   Icon            =   "frmRespaldo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1425
   ScaleWidth      =   7800
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Height          =   735
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1296
      AutoSize        =   0   'False
      Caption         =   "Destino"
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
      Picture         =   "frmRespaldo.frx":08CA
   End
   Begin DevPowerFlatBttn.FlatBttn cmdRespaldar 
      Height          =   375
      Left            =   5400
      TabIndex        =   0
      Top             =   960
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Respaldar"
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
      Picture         =   "frmRespaldo.frx":60BC
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCerrar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6570
      TabIndex        =   1
      Top             =   960
      Width           =   1125
      _ExtentX        =   1984
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
      Picture         =   "frmRespaldo.frx":614C
   End
   Begin VB.Label lblDestino 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   375
      Left            =   1200
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Label Label1 
      Caption         =   "Destino:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmRespaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Objeto para manejar los Shell Objects
Private oShell As Shell32.Shell

Dim Fl() As cFlatControl
Dim base As Object
Dim origen As String
Dim Destino As String

Private Sub cmdBuscar_Click()
Dim tFolder As Shell32.Folder
Dim s As String

On Error GoTo cancelado
    Set tFolder = oShell.BrowseForFolder(0, "Seleccione el Destino del Respaldo...", 0, "")
    Destino = tFolder.ParentFolder.ParseName(tFolder.Title).Path
    lblDestino.Caption = tFolder.ParentFolder.ParseName(tFolder.Title).Path
    Label1.Visible = True
    lblDestino.Visible = True
    Exit Sub
cancelado:
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdRespaldar_Click()

On Error GoTo errores

If lblDestino.Caption = "" Then
    MsgBox "Seleccione la dirección donde desea realizar el respaldo !!", vbInformation, "Respaldo"
    Exit Sub
Else
    origen = Path & "\Base De Datos\Datos.mdb"
    Destino = Destino & "\"
    Set base = CreateObject("Scripting.FileSystemObject")
    base.CopyFile origen, Destino, False
    MsgBox "El respaldo fue realizado con éxito en la dirección:" & Chr(13) & "" & Trim(lblDestino.Caption) & " ", vbInformation, "Respaldo"
    Label1.Caption = ""
    Label1.Visible = False
    lblDestino.Visible = False
    Exit Sub

errores:
    If Err.Number = 58 Then
        If MsgBox("El archivo ya existe !!" & Chr(13) & "Desea remplazarlo ??", vbYesNo + vbQuestion, "Respaldo de Información") = vbYes Then
            base.CopyFile origen, Destino, True
            MsgBox "El respaldo fue realizado con éxito en la dirección:" & Chr(13) & "" & lblDestino.Caption & " ", vbInformation, "Respaldo"
            Label1.Caption = ""
            Label1.Visible = False
            lblDestino.Visible = False
        End If
    Exit Sub
    End If
End If
End Sub

Private Sub Drive1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
' Crear el objeto Shell
Set oShell = New Shell32.Shell
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 5
End Sub
