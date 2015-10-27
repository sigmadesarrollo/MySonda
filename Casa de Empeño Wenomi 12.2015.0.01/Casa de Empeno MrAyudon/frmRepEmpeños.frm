VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRepEmpeños 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de empeños"
   ClientHeight    =   1380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepEmpeños.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1380
   ScaleWidth      =   5340
   Begin VB.CheckBox chkDesempeños 
      Appearance      =   0  'Flat
      Caption         =   "Desempeños"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1455
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
      Left            =   2160
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   1455
   End
   Begin VB.CheckBox chkRefrendos 
      Appearance      =   0  'Flat
      Caption         =   "Refrendos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CheckBox chkEmpeños 
      Appearance      =   0  'Flat
      Caption         =   "&Empeños"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   1
      Left            =   3660
      TabIndex        =   4
      Top             =   600
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
      Picture         =   "frmRepEmpeños.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   1605
      TabIndex        =   5
      Top             =   600
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
      Picture         =   "frmRepEmpeños.frx":0121
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4080
      TabIndex        =   9
      Top             =   885
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
      Picture         =   "frmRepEmpeños.frx":0236
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   4080
      TabIndex        =   10
      Top             =   360
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   " &Imprimir"
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
      Picture         =   "frmRepEmpeños.frx":0788
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
      TabIndex        =   7
      Top             =   240
      Width           =   1590
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
      Left            =   2160
      TabIndex        =   6
      Top             =   240
      Width           =   1410
   End
End
Attribute VB_Name = "frmRepEmpeños"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
    Imprimir
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)
    
    If Index = 0 Then
        txtFechaIni.text = frmCalendario.Fecha(txtFechaIni.text)
    
    Else
        txtFechaFin.text = frmCalendario.Fecha(txtFechaFin.text)
    
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
    txtFechaIni.text = Format(Date, "DD/MM/YYYY")
    txtFechaFin.text = Format(Date, "DD/MM/YYYY")
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

Private Sub Imprimir()

On Error GoTo Error
    
    If Trim(txtFechaIni.text) = "" Or Trim(txtFechaFin.text) = "" Then Exit Sub
    
    If chkEmpeños.Value = vbChecked Then
        With frmMDI.Cr
            .Reset
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\Empeno.rpt"
            .SelectionFormula = "{empeno.Fecha} >= date('" & Format(txtFechaIni.text, "YYYY,MM,DD") & "') AND {empeno.Fecha}<= date('" & Format(txtFechaFin.text, "YYYY,MM,DD") & "') AND {empeno.Origen}=1"
            .Formulas(0) = "Leyenda='REPORTE DE EMPEÑOS'"
            .Formulas(1) = "Encabezado='" & Sucursal.RazonSocial & "'"
            .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(3) = "Caja='" & NombrePc & "'"
            .Formulas(4) = "Cajero='" & frmMDI.Usuario & "'"
            .Formulas(5) = "Subencabezado='De la fecha " & Format(txtFechaIni.text, "dd/mmm/yyyy") & " a " & Format(txtFechaFin.text, "dd/mmm/yyyy") & "'"
            .WindowShowPrintSetupBtn = True
            .WindowShowExportBtn = True
            .DiscardSavedData = True
            .WindowTitle = "Reporte de empeños"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
    End If
   
    If chkRefrendos.Value = vbChecked Then
        With frmMDI.Cr
            .Reset
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\Refrendos.rpt"
            .SelectionFormula = "{empeno.FechaMovimiento} >= date('" & Format(txtFechaIni.text, "YYYY,MM,DD") & "') AND {empeno.FechaMovimiento}<= date('" & Format(txtFechaFin.text, "YYYY,MM,DD") & "') AND ({empeno.Destino}=2  OR {empeno.Destino}=5) AND {empeno.Cancelado}=0"
            .Formulas(0) = "Leyenda='REPORTE DE REFRENDOS'"
            .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(2) = "Encabezado='" & Sucursal.RazonSocial & "'"
            .Formulas(3) = "Caja='" & NombrePc & "'"
            .Formulas(4) = "Cajero='" & frmMDI.Usuario & "'"
            .Formulas(5) = "Subencabezado='De la fecha " & Format(txtFechaIni.text, "dd/mmm/yyyy") & " a " & Format(txtFechaFin.text, "dd/mmm/yyyy") & "'"
            .DiscardSavedData = True
            .WindowShowPrintSetupBtn = True
            .WindowShowExportBtn = True
            .WindowTitle = "Reporte de refrendos"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
    End If

    If chkDesempeños.Value = vbChecked Then
        With frmMDI.Cr
            .Reset
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\Desempeno.rpt"
            .SelectionFormula = "{empeno.FechaMovimiento} >= date('" & Format(txtFechaIni.text, "YYYY,MM,DD") & "') AND {empeno.FechaMovimiento}<= date('" & Format(txtFechaFin.text, "YYYY,MM,DD") & "') AND {empeno.Destino}=3 AND {empeno.Cancelado}=0"
            .Formulas(0) = "Leyenda='REPORTE DE DESEMPEÑOS'"
            .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(2) = "Encabezado='" & Sucursal.RazonSocial & "'"
            .Formulas(3) = "Caja='" & NombrePc & "'"
            .Formulas(4) = "Cajero='" & frmMDI.Usuario & "'"
            .Formulas(5) = "Subencabezado='De la fecha " & Format(txtFechaIni.text, "dd/mmm/yyyy") & " a " & Format(txtFechaFin.text, "dd/mmm/yyyy") & "'"
            .DiscardSavedData = True
            .WindowShowExportBtn = True
            .WindowShowPrintSetupBtn = True
            .WindowTitle = "Reporte de desempeños"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With
    End If

Error:
    Maneja_Error Err
    
End Sub
