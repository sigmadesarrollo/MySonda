VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRangoFechas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Rango de fechas"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRangoFechas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1335
   ScaleWidth      =   5025
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   840
      Width           =   1440
   End
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
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   1440
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   1
      Left            =   3240
      TabIndex        =   5
      Top             =   840
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
      Picture         =   "frmRangoFechas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   3240
      TabIndex        =   2
      Top             =   240
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
      Picture         =   "frmRangoFechas.frx":0121
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3720
      TabIndex        =   6
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "       &Imprimir"
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
      Picture         =   "frmRangoFechas.frx":0236
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   795
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
      Picture         =   "frmRangoFechas.frx":0788
   End
   Begin VB.Label Label1 
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
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1410
   End
   Begin VB.Label Label4 
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
      TabIndex        =   0
      Top             =   240
      Width           =   1590
   End
End
Attribute VB_Name = "frmRangoFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim Fecha1 As String, Fecha2 As String

Private Sub cmdAceptar_Click()

    If Validar Then
        
        Fecha1 = txtFechaIni.text
        Fecha2 = txtFechaFin.text
        Unload Me
    End If

End Sub

Private Function Validar() As Boolean
    
    Validar = True
   
    If Not IsDate(txtFechaIni.text) Then
        Me.Hide
        MsgBox "Favor de poner una fecha correcta", vbOKOnly + vbCritical
        Me.Show
        txtFechaIni.SetFocus
        Validar = False
        Exit Function
    End If
   
    If Not IsDate(txtFechaFin.text) Then
        Me.Hide
        MsgBox "Favor de poner una fecha correcta", vbOKOnly + vbCritical
        Me.Show
        txtFechaFin.text = ""
        txtFechaFin.SetFocus
        Validar = False
        Exit Function
    End If
   
    If CDate(txtFechaIni.text) > CDate(txtFechaFin.text) Then
        Me.Hide
        MsgBox "La fecha inicial no puede ser mayor a la fecha final", vbOKOnly + vbCritical
        Me.Show
        txtFechaIni.text = ""
        txtFechaIni.SetFocus
        Validar = False
        Exit Function
    End If
      
End Function

Private Sub cmdMosFecha_Click(Index As Integer)

    If Index = 0 Then
        
        txtFechaIni.text = frmCalendario.Fecha(txtFechaIni.text, 1)
    Else
        
        txtFechaFin.text = frmCalendario.Fecha(txtFechaFin.text, 1)
    End If

End Sub

Private Sub cmdSalir_Click()
    Fecha1 = ""
    Fecha2 = ""
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    txtFechaIni.text = Format(Date, "DD/MMM/YYYY")
    txtFechaFin.text = Format(Date, "DD/MMM/YYYY")
    frmMDI.FechaIni = "": frmMDI.FechaFin = ""
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtFechaFin_GotFocus()
    Seleccionar_Texto txtFechaFin
    Cambiar_Color True, txtFechaFin
End Sub

Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaFin_LostFocus()
    Cambiar_Color False, txtFechaFin
End Sub

Private Sub txtFechaIni_GotFocus()
    Seleccionar_Texto txtFechaIni
    Cambiar_Color True, txtFechaIni
End Sub

Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFechaIni_LostFocus()
    Cambiar_Color False, txtFechaIni
End Sub

Public Sub Fechas(ByRef FechaIni As String, ByRef FechaFin As String)
    Me.Show vbModal
    FechaIni = Fecha1
    FechaFin = Fecha2
End Sub
