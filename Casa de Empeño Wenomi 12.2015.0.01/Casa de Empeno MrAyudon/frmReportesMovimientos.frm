VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmReportesMovimientos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de auditoria"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReportesMovimientos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7320
   ScaleWidth      =   8520
   Begin vbAcceleratorGrid6.vbalGrid grdMovimientos 
      Height          =   5535
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   9763
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
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
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   17
   End
   Begin VB.TextBox txtCajero 
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
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   5175
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   6120
      TabIndex        =   7
      Top             =   6840
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
      Picture         =   "frmReportesMovimientos.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7320
      TabIndex        =   8
      Top             =   6840
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
      Picture         =   "frmReportesMovimientos.frx":055E
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   6240
      TabIndex        =   2
      Top             =   120
      Width           =   795
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "<Fecha>"
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
      Left            =   7080
      TabIndex        =   3
      Top             =   120
      Width           =   960
   End
   Begin VB.Label lblSucursal 
      AutoSize        =   -1  'True
      Caption         =   "<Sucursal>"
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
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cajero:"
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
      TabIndex        =   4
      Top             =   600
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal:"
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
      Top             =   120
      Width           =   1110
   End
End
Attribute VB_Name = "frmReportesMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 17/07/02
' Modulo frmReportesMovimientos - frmReportesMovimientos.frm
' Ultima Modificacion - 17/07/02
''Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit
Dim Fl() As cFlatControl

Private Sub cmdImprimir_Click()

    'imprimimos el reporte de diario
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .ReportFileName = Path & "\Reportes\Movimientos.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{Auxiliar.Fecha}=date(" & Format(Date, "YYYY,MM,DD") & ")"
        .Formulas(0) = "Sucursal='" & lblSucursal.Caption & "'"
        .Formulas(1) = "Cajero='" & txtCajero.text & "'"
        .Formulas(2) = "Encabezado='" & Sucursal.RazonSocial & "'"
        .Formulas(3) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(4) = "Sucursal='" & Sucursal.NombreComercial & "'"
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "Reporte de audítoria"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

'inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    lblSucursal.Caption = Sucursal.NombreComercial
    lblFecha.Caption = Format(Date, "DD/MMM/YY")
    Crear_Encabezados
    Cargar_Movimientos
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados
Private Sub Crear_Encabezados()

    With grdMovimientos
        .AddColumn "K1", "Movimiento", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Folio", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
        .AddColumn "K3", "Iniciales", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K4", "Cuenta", ecgHdrTextALignLeft, , 85, , , , , , , CCLSortNumeric
        .AddColumn "K5", "Importe", ecgHdrTextALignRight, , 95, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Caja", ecgHdrTextALignLeft, , 95, , , , , , , CCLSortString
    End With

End Sub

'Cargamos los movimientos del dia
Private Sub Cargar_Movimientos()
Dim rcMovimientos As New ADODB.Recordset

On Error GoTo error

    Screen.MousePointer = vbHourglass

    rcMovimientos.Open "SELECT * FROM auxiliar WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY ID", dbDatos, adOpenForwardOnly, adLockOptimistic
   
    With rcMovimientos
        
        While Not .EOF
            grdMovimientos.AddRow
            grdMovimientos.CellText(grdMovimientos.Rows, 1) = !Movimiento
            grdMovimientos.CellTextAlign(grdMovimientos.Rows, 1) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdMovimientos.CellText(grdMovimientos.Rows, 2) = !Folio
            grdMovimientos.CellTextAlign(grdMovimientos.Rows, 2) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdMovimientos.CellText(grdMovimientos.Rows, 3) = !Iniciales
            grdMovimientos.CellTextAlign(grdMovimientos.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdMovimientos.CellText(grdMovimientos.Rows, 4) = !Cuenta
            grdMovimientos.CellTextAlign(grdMovimientos.Rows, 4) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdMovimientos.CellText(grdMovimientos.Rows, 5) = Format(!Importe, "Currency")
            grdMovimientos.CellTextAlign(grdMovimientos.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdMovimientos.CellText(grdMovimientos.Rows, 6) = !PC
            grdMovimientos.CellTextAlign(grdMovimientos.Rows, 6) = DT_LEFT Or DT_WORD_ELLIPSIS
        .MoveNext
        Wend
    
    End With
    rcMovimientos.Close
    Set rcMovimientos = Nothing
    Screen.MousePointer = vbDefault
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcMovimientos = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtCajero_GotFocus()
    Seleccionar_Texto txtCajero
    Cambiar_Color True, txtCajero
End Sub

Private Sub txtCajero_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCajero_LostFocus()
    Cambiar_Color False, txtCajero
End Sub
