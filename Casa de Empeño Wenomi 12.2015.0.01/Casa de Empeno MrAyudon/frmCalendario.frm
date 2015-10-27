VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{3D425BEC-988C-11D5-B192-000102ACE780}#1.0#0"; "CalendarVB.ocx"
Begin VB.Form frmCalendario 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Calendario"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   255
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ctrCalendarVB.CalendarVB Calendario 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5318
      CurrentPeriodbackColor=   -2147483643
      CurrentPeriodForeColor=   -2147483630
      DayHeaderBackColor=   -2147483633
      DayHeaderForeColor=   -2147483630
      ActiveDayForeColor=   158700
      FlatLineColor   =   12632256
      PrePeriodBackColor=   -2147483648
      PrePeriodforeColor=   -2147483632
      PostPeriodBackColor=   -2147483648
      PostPeriodforeColor=   -2147483632
      DateTipFormat   =   "Dddd dd Mmmm  yyyy"
      ActiveDayFontBold=   -1  'True
      ActiveDayFontItalic=   0   'False
      ActiveDayFontSize=   8.25
      ActiveDayFontName=   "Tahoma"
      DayHeaderFontBold=   -1  'True
      DayHeaderFontItalic=   0   'False
      DayHeaderFontSize=   8.25
      DayHeaderFontName=   "Tahoma"
      DaysFontBold    =   0   'False
      DaysFontItalic  =   0   'False
      DaysFontSize    =   8.25
      DaysFontName    =   "Tahoma"
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   1
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "  &Aceptar"
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
      Picture         =   "frmCalendario.frx":0000
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   3240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Cancelar"
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
      Picture         =   "frmCalendario.frx":036A
   End
End
Attribute VB_Name = "frmCalendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strDia As String
Private m_c() As cFlatControl

Private Sub Calendario_DblClick(ByVal DateValue As String, ByVal PeriodValue As String, ByVal Row As Integer, ByVal Col As Integer)
  cmdAceptar_Click
End Sub

Private Sub cmdSalir_Click()
  strDia = ""
   Unload Me
End Sub

Private Sub Form_Load()
   Inicializar
End Sub

'inicializamos la forma
Private Sub Inicializar()
   ReDim m_c(0 To 1) As cFlatControl
   Set m_c(0) = New cFlatControl
   m_c(0).hWndAttach Calendario.hWndPeriod, Calendario.hwnd, True
   Set m_c(1) = New cFlatControl
   m_c(1).hWndAttach Calendario.hWndYear, Calendario.hwnd, True
   Calendario.YearBegin = 1900
   Calendario.YearEnd = 2100
End Sub

'Funcion publica que regresa la fecha seleccionada
Public Function Fecha(dia As String, Optional X As Integer) As String
On Error GoTo error
    
    Calendario.DateValue = IIf(dia = "", Date, dia)
    strDia = dia
    Me.Show vbModal
    If X = 0 Then Fecha = Format(strDia, "DD/MM/YYYY") Else Fecha = Format(strDia, "DD/MMM/YYYY")
    Exit Function
    
error:
  Unload Me
End Function

Private Sub cmdAceptar_Click()
On Error GoTo error
  
  strDia = Calendario.DateValue
  Unload Me

error:
    Maneja_Error Err
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Dim Indice As Integer
   For Indice = LBound(m_c) To UBound(m_c)
      Set m_c(Indice) = Nothing
   Next Indice
End Sub
