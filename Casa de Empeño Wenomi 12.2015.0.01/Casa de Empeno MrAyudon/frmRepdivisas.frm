VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRepdivisas 
   Caption         =   "Compra / venta de dólares"
   ClientHeight    =   1380
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4845
   Icon            =   "frmRepdivisas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1380
   ScaleWidth      =   4845
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1500
   End
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
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   1500
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   720
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
      Picture         =   "frmRepdivisas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   1
      Left            =   3240
      TabIndex        =   3
      Top             =   735
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
      Picture         =   "frmRepdivisas.frx":009D
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   3240
      TabIndex        =   4
      Top             =   120
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
      Picture         =   "frmRepdivisas.frx":01B2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   120
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
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmRepdivisas.frx":02C7
      PictureDisabled =   "frmRepdivisas.frx":0631
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "De la fecha:"
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
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "A la fecha:"
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
      TabIndex        =   6
      Top             =   720
      Width           =   1305
   End
End
Attribute VB_Name = "frmRepdivisas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Dim fecha1 As String
Dim fecha2 As String

Private Sub cmdAceptar_Click()
   If Validar Then
      fecha1 = txtFechaIni.Text
      fecha2 = txtFechaFin.Text
      muestrareporte CDate(fecha1), CDate(fecha2)
   End If
End Sub

Private Function Validar() As Boolean
   Validar = True
   
   If Not IsDate(txtFechaIni.Text) Then
      Me.Hide
      MsgBox "Favor de poner una fecha correcta", vbOKOnly + vbCritical
      Me.Show
      txtFechaIni.SetFocus
      Validar = False
      Exit Function
   End If
   
   
   If Not IsDate(txtFechaFin.Text) Then
      Me.Hide
      MsgBox "Favor de poner una fecha correcta", vbOKOnly + vbCritical
      Me.Show
      txtFechaFin.SetFocus
      Validar = False
      Exit Function
   End If
   
   If CDate(txtFechaIni.Text) > CDate(txtFechaFin.Text) Then
      Me.Hide
      MsgBox "La fecha inicial no puede ser mayor a la fecha final", vbOKOnly + vbCritical
      Me.Show
      txtFechaIni.SetFocus
      Validar = False
      Exit Function
   End If
      
End Function


Private Sub cmdMosFecha_Click(Index As Integer)
   If Index = 0 Then
      txtFechaIni.Text = frmCalendario.Fecha(txtFechaIni.Text, 1)
   Else
      txtFechaFin.Text = frmCalendario.Fecha(txtFechaFin.Text, 1)
   End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
inicializar
End Sub

Private Sub inicializar()
Me.Height = 1890
Me.Width = 4965
txtFechaIni.Text = Format(Date, "dd/mmm/yyyy")
txtFechaFin.Text = Format(Date, "dd/mmm/yyyy")
Screen.MousePointer = vbHourglass
Poner_Flat Fl, Me.Controls, Me
Screen.MousePointer = vbDefault
Move (Screen.Width - Width) \ 2, (Screen.Height - Height) \ 5
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
   FechaIni = fecha1
   FechaFin = fecha2
End Sub

Function muestrareporte(fechainicial As Date, fechafinal As Date)
Dim Campo As String

dbReportes.Execute "delete * from compraventa"
Set rcConsulta = dbDatos.Execute("select * from divisas where fecha>=#" & Format(fechainicial, "mm/dd/yyyy") & "# and fecha<=#" & Format(fechafinal, "mm/dd/yyyy") & "# and interno_externo=0 order by id")
If Not rcConsulta.BOF And Not rcConsulta.EOF Then
    rcConsulta.MoveFirst
    While Not rcConsulta.EOF
        Campo = IIf(rcConsulta!tipo = 0, "compra", "venta")
        dbReportes.Execute "insert into compraventa (iddivisa,cotizacion," & Campo & ",interno)values(" & rcConsulta!divisa & "," & rcConsulta!importe & "," & rcConsulta!Cantidad & ",0)"
    rcConsulta.MoveNext
    Wend
End If

Set rcConsulta = dbDatos.Execute("select * from divisas where fecha>=#" & Format(fechainicial, "mm/dd/yyyy") & "# and fecha<=#" & Format(fechafinal, "mm/dd/yyyy") & "# and interno_externo=1 order by id")
If Not rcConsulta.BOF And Not rcConsulta.EOF Then
    rcConsulta.MoveFirst
    While Not rcConsulta.EOF
        Campo = IIf(rcConsulta!tipo = 0, "compra", "venta")
        dbReportes.Execute "insert into compraventa (iddivisa,cotizacion," & Campo & ",interno)values(" & rcConsulta!divisa & "," & rcConsulta!importe & "," & rcConsulta!Cantidad & ",1)"
    rcConsulta.MoveNext
    Wend
End If


With frmMDI.Cr
    .Reset
    .DiscardSavedData = True
    .ReportFileName = Path & "\Reportes\Divisas.rpt"
    .DataFiles(0) = Path & "\Base De Datos\Datos.mdb"
    .DataFiles(1) = Path & "\Base De Datos\Datos.mdb"
    .DataFiles(2) = Path & "\Base De Datos\Reportes.mdb"
    .password = Chr(10) & "administrativo"
    .Formulas(0) = "Titulo='" & Trim(Regresa_Valor_BD("Razonsocial")) & "'"
    .Formulas(1) = "Subtitulo='SUCURSAL: " & Trim(Regresa_Valor_BD("Nomcomercial")) & "'"
    .WindowShowPrintSetupBtn = True
    .DiscardSavedData = True
    .WindowTitle = "Reporte de compra/venta de divisas"
    .WindowState = crptMaximized
    .Action = 1
End With

Set rcConsulta = Nothing
End Function
