VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmRepInventarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de depositaría"
   ClientHeight    =   1470
   ClientLeft      =   10755
   ClientTop       =   780
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmRepInventarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1470
   ScaleWidth      =   3795
   Begin VB.ComboBox cmbTipo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmRepInventarios.frx":000C
      Left            =   120
      List            =   "frmRepInventarios.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   390
      Width           =   2295
   End
   Begin VB.OptionButton opApartados 
      Appearance      =   0  'Flat
      Caption         =   "&Apartados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1695
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   2760
      Visible         =   0   'False
      Width           =   1320
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
      Top             =   2760
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.OptionButton opTipos 
      Appearance      =   0  'Flat
      Caption         =   "Empeños"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton opJoyeria 
      Appearance      =   0  'Flat
      Caption         =   "Vitrina"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   1
      Left            =   3480
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
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
      Picture         =   "frmRepInventarios.frx":0010
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Index           =   0
      Left            =   1560
      TabIndex        =   5
      Top             =   2760
      Visible         =   0   'False
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
      Picture         =   "frmRepInventarios.frx":0125
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2550
      TabIndex        =   10
      Top             =   840
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
      Picture         =   "frmRepInventarios.frx":023A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   2550
      TabIndex        =   11
      Top             =   360
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
      Picture         =   "frmRepInventarios.frx":078C
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
      Left            =   240
      TabIndex        =   7
      Top             =   2400
      Visible         =   0   'False
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
      Top             =   2400
      Visible         =   0   'False
      Width           =   1410
   End
End
Attribute VB_Name = "frmRepInventarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbTipo_DropDown()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_GotFocus()
    Seleccionar_Texto cmbTipo
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
End Sub

Private Sub cmdAceptar_Click()
    Imprimir
End Sub

Private Sub cmdMosFecha_Click(Index As Integer)

    If Index = 0 Then
    
        txtFechaIni.text = frmCalendario.Fecha(txtFechaIni.text)
    Else
    
        txtFechaFin.text = frmCalendario.Fecha(txtFechaFin.text)
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub opApartados_Click()
    cmbTipo.Enabled = False
    txtFechaIni.Enabled = True
    txtFechaFin.Enabled = True
    cmdMosFecha(0).Enabled = True
    cmdMosFecha(1).Enabled = True
End Sub

Private Sub opJoyeria_Click()
    cmbTipo.Enabled = False
    txtFechaIni.text = ""
    txtFechaFin.text = ""
    txtFechaIni.Enabled = False
    txtFechaFin.Enabled = False
    cmdMosFecha(0).Enabled = False
    cmdMosFecha(1).Enabled = False
End Sub

Private Sub opTipos_Click()
    cmbTipo.Enabled = True
    txtFechaIni.Enabled = True
    txtFechaFin.Enabled = True
    cmdMosFecha(0).Enabled = True
    cmdMosFecha(1).Enabled = True
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

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    cmbTipo.AddItem "(TODOS)"
    Cargar_Combos "Descripcion", "tipo", cmbTipo, , , False
    cmbTipo.AddItem "AUTOS"
    cmbTipo.ListIndex = 0
    Screen.MousePointer = vbDefault
End Sub

Private Sub Imprimir()
Dim Filtro As String, Filtra As String, Serie  As Integer

On Error GoTo Error

    If opTipos.Value Then
    
        If cmbTipo.text = "(TODOS)" Then
                
            Serie = 0
            Filtro = ""
        ElseIf cmbTipo.text = "AUTOS" Then
            
            Serie = SERIE_B
            Filtro = "{empeno.Serie}=" & SERIE_B
        Else
        
            Serie = 0
            Filtro = "{detallesempeno.tipo}=" & cmbTipo.ItemData(cmbTipo.ListIndex)
        End If
    
    End If
      
    If opTipos.Value Then

        With frmMDI.Cr
            .Reset
            .DiscardSavedData = True
            .WindowShowPrintSetupBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\RepDepositaria2.rpt"
            If Filtro <> "" Then
                .SelectionFormula = Filtro
            End If
            .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
            .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(2) = "Leyenda='" & IIf(txtFechaIni.text = "" Or txtFechaFin.text = "", "De la fecha: " & Format(Date, "dd/mmm/yyyy") & "", "De la fecha: " & Format(txtFechaIni.text, "dd/mmm/yyyy") & " a " & Format(txtFechaFin.text, "dd/mmm/yyyy") & "") & "'"
            .Formulas(3) = "Serie=" & Serie
            .WindowTitle = "Reporte depositaria"
            .WindowState = crptMaximized
            .Destination = crptToWindow
            .Action = 1
        End With

    ElseIf opApartados.Value Then

        With frmMDI.Cr
                       
            .Reset
            .WindowShowPrintSetupBtn = True
            .WindowShowExportBtn = True
            .DiscardSavedData = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\RepApartados.rpt"
            .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
            .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(2) = "Encabezado='De la fecha: " & Format(Date, "DD/MMM/YYYY") & "'"
            .Formulas(3) = "Todos=1"
            .WindowTitle = "Reporte ventas de apartado"
            .WindowState = crptMaximized
            .Action = 1
        End With

    ElseIf opJoyeria.Value Then

        With frmMDI.Cr
            .Reset
            .WindowShowPrintSetupBtn = True
            .WindowShowExportBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\Existencias.rpt"
            .SelectionFormula = "({detallesentradainventario.TipoEntrada}=" & ENTRADAALMONEDA & " OR {detallesentradainventario.TipoEntrada}=" & ENTRADACOMPRA & " OR {detallesentradainventario.TipoEntrada}=" & ENTRADADOTACION & ") AND {detallesentradainventario.Cantidad}>0"
            .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
            .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(2) = "Encabezado='De la fecha: " & Format(Date, "dd/mmm/yyyy") & "'"
            .WindowTitle = "Reporte de existencias"
            .DiscardSavedData = True
            .WindowState = crptMaximized
            .Destination = crptToWindow
            .Action = 1
        End With

    End If
   
Error:
    Maneja_Error Err

End Sub

