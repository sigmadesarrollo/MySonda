VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmReporteFinanciero 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Activos"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3240
   Icon            =   "frmReporteFinanciero.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   3240
   Begin vbAcceleratorGrid6.vbalGrid grdDivisas 
      Height          =   1005
      Left            =   105
      TabIndex        =   14
      Top             =   5160
      Visible         =   0   'False
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   1773
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   -2147483638
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Header          =   0   'False
      HeaderButtons   =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   12
      Top             =   2400
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
      Picture         =   "frmReporteFinanciero.frx":0442
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   840
      TabIndex        =   13
      Top             =   2400
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "       &Aceptar"
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
      Picture         =   "frmReporteFinanciero.frx":0994
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Divisas:"
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
      TabIndex        =   15
      Top             =   4800
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Caption         =   "<Total>"
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
      TabIndex        =   11
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Line Line2 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   1560
      X2              =   3360
      Y1              =   6960
      Y2              =   6960
   End
   Begin VB.Line Line1 
      BorderStyle     =   3  'Dot
      Visible         =   0   'False
      X1              =   1560
      X2              =   3360
      Y1              =   6840
      Y2              =   6840
   End
   Begin VB.Label lblFaltante 
      Alignment       =   1  'Right Justify
      Caption         =   "<Faltante>"
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
      TabIndex        =   10
      Top             =   6240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Apartados:"
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
      TabIndex        =   9
      Top             =   1920
      Width           =   1350
   End
   Begin VB.Label lblApartados 
      Alignment       =   1  'Right Justify
      Caption         =   "<Apartados>"
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
      Left            =   1605
      TabIndex        =   8
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Almoneda:"
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
      Top             =   1320
      Width           =   1320
   End
   Begin VB.Label lblAlmoneda 
      Alignment       =   1  'Right Justify
      Caption         =   "<Joyeria>"
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
      Left            =   1605
      TabIndex        =   6
      Top             =   1320
      Width           =   1575
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Depositaría:"
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
      TabIndex        =   5
      Top             =   720
      Width           =   1500
   End
   Begin VB.Label lblEmpeño 
      Alignment       =   1  'Right Justify
      Caption         =   "<Empeño>"
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
      Left            =   1605
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Bancos:"
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
      Top             =   4920
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblBancos 
      Alignment       =   1  'Right Justify
      Caption         =   "<Bancos>"
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
      Left            =   1605
      TabIndex        =   2
      Top             =   4920
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Bóveda:"
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
      TabIndex        =   1
      Top             =   240
      Width           =   990
   End
   Begin VB.Label lblBoveda 
      Alignment       =   1  'Right Justify
      Caption         =   "<Boveda>"
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
      Left            =   1605
      TabIndex        =   0
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "frmReporteFinanciero"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAceptar_Click()
    Imprimir
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
Dim ctrl As Control

    Screen.MousePointer = vbHourglass
    

    For Each ctrl In Controls

        If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = "0.00"
    
    Next
    
    'Creo los encabezados del Grid de Divisas
    grdDivisas.AddColumn "C1", "", ecgHdrTextALignLeft, , 120, , , , , , , CCLSortString
    grdDivisas.AddColumn "C2", "", ecgHdrTextALignLeft, , 75, , , , , "###,###,###,###0", , CCLSortString
    
    Cargar_Montos
    Poner_Totales
    CentrarForm frmReporteFinanciero, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Private Sub Poner_Totales()
    lblTotal.Caption = Format(CCur(IIf(lblBoveda.Caption = "", 0, lblBoveda.Caption)) + CCur(IIf(lblBancos.Caption = "", 0, lblBancos.Caption)) + CCur(IIf(lblEmpeño.Caption = "", 0, lblEmpeño.Caption)) + CCur(IIf(lblAlmoneda.Caption = "", 0, lblAlmoneda.Caption)) + CCur(IIf(lblApartados.Caption = "", 0, lblApartados.Caption)) + CCur(IIf(lblFaltante.Caption = "", 0, lblFaltante.Caption)), "###,###,###,###0.00")
End Sub

Private Sub Cargar_Montos()
Dim rcBD As New ADODB.Recordset
Dim rcAux As New ADODB.Recordset
Dim Cargo As Currency, Abono As Currency

    DoEvents
    
    'Boveda**************************************************
    Cargo = 0: Abono = 0
    rcBD.Open "SELECT Sum(Importe) AS Cargo FROM auxiliar WHERE Cuenta='110901'", dbDatos, adOpenForwardOnly, adLockOptimistic
        Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
    rcBD.Close

    rcBD.Open "SELECT Sum(Importe) AS total FROM auxiliar WHERE Cuenta='110950'", dbDatos, adOpenForwardOnly, adLockOptimistic
        Abono = IIf(IsNull(rcBD!Total), 0, rcBD!Total)
    rcBD.Close
    
    lblBoveda.Caption = Format(Cargo - Abono, FMoneda)
    lblBoveda.Tag = Val(Cargo - Abono)
    '**********************************************************
    
    'Bancos***********************************
    Cargo = 0: Abono = 0
    rcBD.Open "SELECT Sum(Importe) AS Cargo FROM auxiliar WHERE Cuenta='210101'", dbDatos, adOpenForwardOnly, adLockOptimistic
        Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
    rcBD.Close

    rcBD.Open "SELECT Sum(Importe) AS total FROM auxiliar WHERE Cuenta='210150'", dbDatos, adOpenForwardOnly, adLockOptimistic
        Abono = IIf(IsNull(rcBD!Total), 0, rcBD!Total)
    rcBD.Close

    lblBancos.Caption = Format(Cargo - Abono, FMoneda)
    lblBancos.Tag = Cargo - Abono
    '****************************************
            
    'Empeños***************************************************
    Cargo = 0
    rcBD.Open "SELECT Sum(Prestamo) AS Cargo FROM empeno WHERE Cancelado=0 AND Destino=0", dbDatos, adOpenForwardOnly, adLockOptimistic
        Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
    rcBD.Close
    
    lblEmpeño.Caption = Format(Cargo, FMoneda)
    lblEmpeño.Tag = Cargo
    '***********************************************************
    
    'Almoneda************************************************
    Cargo = 0
    rcBD.Open "SELECT Sum(Costo) AS Cargo FROM detallesentradainventario WHERE Cantidad>0", dbDatos, adOpenForwardOnly, adLockOptimistic
        Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
    rcBD.Close

    lblAlmoneda.Caption = Format(Cargo, FMoneda)
    lblAlmoneda.Tag = Cargo
    '************************************************************
        
    'Apartados***************************************************
    Cargo = 0: Abono = 0
    rcBD.Open "SELECT Sum(Importe) AS Cargo FROM auxiliar WHERE Cuenta='620501'", dbDatos, adOpenForwardOnly, adLockOptimistic
        Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
    rcBD.Close

    rcBD.Open "SELECT Sum(Importe) AS Total FROM auxiliar WHERE Cuenta='620550'", dbDatos, adOpenForwardOnly, adLockOptimistic
        Abono = IIf(IsNull(rcBD!Total), 0, rcBD!Total)
    rcBD.Close
    
    lblApartados.Caption = Format(Cargo - Abono, FMoneda)
    lblApartados.Tag = Val(Cargo - Abono)
    '***************************************************************
    
'''''    'Divisas*********************************************************
'''''    Cargo = 0
'''''    Abono = 0
'''''    rcAux.Open "SELECT DISTINCT a.IDDivisa,m.Descripcion AS Divisa FROM auxiliar a INNER JOIN monedas m ON a.IDDivisa=m.Clave WHERE a.Cuenta='910901' OR a.Cuenta='910950'", dbDatos, adOpenForwardOnly, adLockReadOnly
'''''    While Not rcAux.EOF
'''''
'''''        rcBD.Open "SELECT SUM(a.Importe) AS Cargo FROM auxiliar a WHERE Cuenta='910901' AND a.IDDivisa=" & rcAux!IDDivisa, dbDatos, adOpenForwardOnly, adLockOptimistic
'''''            Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
'''''        rcBD.Close
'''''
'''''        rcBD.Open "SELECT SUM(a.Importe) AS Abono FROM auxiliar a WHERE Cuenta='910950' AND a.IDDivisa=" & rcAux!IDDivisa, dbDatos, adOpenForwardOnly, adLockOptimistic
'''''            Abono = IIf(IsNull(rcBD!Abono), 0, rcBD!Abono)
'''''        rcBD.Close
'''''
'''''        grdDivisas.AddRow
'''''        grdDivisas.CellText(grdDivisas.Rows, 1) = rcAux!divisa
'''''        grdDivisas.CellItemData(grdDivisas.Rows, 1) = rcAux!IDDivisa
'''''        grdDivisas.CellText(grdDivisas.Rows, 2) = Cargo - Abono
'''''        grdDivisas.CellTextAlign(grdDivisas.Rows, 2) = DT_RIGHT
'''''
'''''    rcAux.MoveNext
'''''    Wend
'''''    rcAux.Close
'''''    '******************************************************************
    
    Set rcBD = Nothing
    Set rcAux = Nothing
End Sub

Private Sub Imprimir()

On Error GoTo error

    'Imprimimos el reporte de corte de caja
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\Financiero.rpt"
        .Formulas(1) = "Boveda=" & ConvMoneda(lblBoveda.Caption) & ""
        .Formulas(2) = "Bancos=" & ConvMoneda(lblBancos.Caption) & ""
        .Formulas(3) = "Empeño=" & ConvMoneda(lblEmpeño.Caption) & ""
        .Formulas(4) = "Joyeria=" & ConvMoneda(lblAlmoneda.Caption) & ""
        .Formulas(5) = "Apartados=" & ConvMoneda(lblApartados.Caption) & ""
        .Formulas(6) = "Faltante=" & ConvMoneda(lblFaltante.Caption) & ""
        .Formulas(7) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(8) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(10) = "Cajero='" & frmMDI.Usuario & "'"
        .WindowTitle = "Activos"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
    Exit Sub
    
error:
    Maneja_Error Err
End Sub
