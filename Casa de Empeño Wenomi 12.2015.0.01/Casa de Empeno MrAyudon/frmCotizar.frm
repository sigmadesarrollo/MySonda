VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmCotizar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotizar Empeño"
   ClientHeight    =   5625
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCotizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   7515
   Begin VB.ComboBox cmbPlazos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCotizar.frx":000C
      Left            =   3525
      List            =   "frmCotizar.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   555
      Width           =   870
   End
   Begin VB.ComboBox cmbPeriodo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCotizar.frx":0010
      Left            =   1890
      List            =   "frmCotizar.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   555
      Width           =   1455
   End
   Begin Line3D.ucLine3D ucLine3D7 
      Height          =   30
      Index           =   0
      Left            =   120
      Top             =   1245
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D6 
      Height          =   1530
      Index           =   0
      Left            =   1800
      Top             =   120
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2699
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D4 
      Height          =   30
      Left            =   120
      Top             =   900
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   30
      Left            =   120
      Top             =   525
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Left            =   120
      Top             =   120
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1515
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2672
      Orientation     =   0
      LineWidth       =   2
   End
   Begin vbAcceleratorGrid6.vbalGrid grdIntereses 
      Height          =   3840
      Left            =   0
      TabIndex        =   4
      Top             =   1560
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   6773
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
   End
   Begin VB.ComboBox cmbTipoInteres 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmCotizar.frx":0014
      Left            =   195
      List            =   "frmCotizar.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   555
      Width           =   1560
   End
   Begin VB.TextBox txtPrestamo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   218
      TabIndex        =   1
      Top             =   1320
      Width           =   1515
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   1200
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
      Picture         =   "frmCotizar.frx":0018
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCalcular 
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   1200
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "    &Cotizar"
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
      Picture         =   "frmCotizar.frx":056A
      PictureDisabled =   "frmCotizar.frx":076A
   End
   Begin Line3D.ucLine3D ucLine3D6 
      Height          =   1530
      Index           =   1
      Left            =   3450
      Top             =   120
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2699
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1485
      Index           =   1
      Left            =   4470
      Top             =   135
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2619
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D7 
      Height          =   30
      Index           =   1
      Left            =   120
      Top             =   1620
      Width           =   4380
      _ExtentX        =   7726
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin VB.Label lblTasa 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0.00%"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   270
      Left            =   3563
      TabIndex        =   15
      Top             =   1297
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tasa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   3450
      TabIndex        =   14
      Top             =   960
      Width           =   1020
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Plazo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   3450
      TabIndex        =   11
      Top             =   180
      Width           =   1020
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Periodo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   1950
      TabIndex        =   10
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label lblAvaluo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
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
      Height          =   225
      Left            =   1875
      TabIndex        =   7
      Top             =   1320
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Préstamo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   480
      TabIndex        =   5
      Top             =   960
      Width           =   930
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Interés"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   308
      TabIndex        =   3
      Top             =   180
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Avalúo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   675
   End
   Begin VB.Label Label4 
      BackColor       =   &H00404040&
      Height          =   1470
      Left            =   150
      TabIndex        =   6
      Top             =   150
      Width           =   4335
   End
End
Attribute VB_Name = "frmCotizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim Ban As Boolean, IsCliente As Boolean

Private Sub cmbPeriodo_Click()
    
    If cmbPeriodo.ListIndex > -1 Then
        
        Cargar_Combos "DISTINCT plazos.Descripcion", "configuraciontasas INNER JOIN plazos ON plazos.ID=configuraciontasas.IDPlazo", cmbPlazos, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND configuraciontasas.IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex), "plazos.Descripcion", , "plazos.ID"
    
    End If
    
    If cmbPlazos.ListCount > 0 Then cmbPlazos.ListIndex = 0
End Sub

Private Sub cmbPeriodo_GotFocus()
    Cambiar_Color True, cmbPeriodo
End Sub

Private Sub cmbPeriodo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPeriodo_LostFocus()
    Cambiar_Color False, cmbPeriodo
End Sub

Private Sub cmbPlazos_Click()
    SacaTasa
End Sub

Private Sub cmbPlazos_GotFocus()
    Cambiar_Color True, cmbPlazos
End Sub

Private Sub cmbPlazos_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPlazos_LostFocus()
    Cambiar_Color False, cmbPlazos
End Sub

Private Sub cmbTipoInteres_Click()
    
    If cmbTipoInteres.ListIndex > -1 Then
        
        cmbPeriodo.Clear
        cmbPlazos.Clear
        Cargar_Combos "DISTINCT tipoperiodo.Descripcion", "configuraciontasas INNER JOIN tipoperiodo ON tipoperiodo.ID=configuraciontasas.IDTipoPeriodo", cmbPeriodo, " WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex), "tipoperiodo.Ordenamiento", , "tipoperiodo.ID"
        
        Encabezado cmbTipoInteres.ListIndex
    End If
        
    If cmbPeriodo.ListCount > 0 Then cmbPeriodo.ListIndex = 0
End Sub

Private Sub cmbTipoInteres_GotFocus()
    Cambiar_Color True, cmbTipoInteres
End Sub

Private Sub cmbTipoInteres_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoInteres_LostFocus()
    Cambiar_Color False, cmbTipoInteres
End Sub

Private Sub cmdCalcular_Click()

    If txtPrestamo.text = "" Then MsgBox "Introduzca el monto del préstamo !!", vbCritical, "Cotizar Empeño": txtPrestamo.SetFocus:   Exit Sub
    
    MuestraInteres CDbl(lblAvaluo.Caption), CDbl(txtPrestamo.text), Date
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Deactivate()
    If Ban Then Unload Me
End Sub

Private Sub Form_Load()
    Ban = False
    IsCliente = False
    Cargar_Combos "Descripcion", "tipointeres", cmbTipoInteres, " WHERE Serie=" & SERIE_A, "Ordenamiento"
    cmbTipoInteres.ListIndex = 0
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Function MuestraInteres(Avaluo As Double, Prestamo As Double, Fecha As Date)
Dim Tasa As Double, Almacenaje As Double, Seguro As Double, Vencimiento As Integer, Periodo   As Integer, Meses As Integer
Dim rcConsulta As New ADODB.Recordset

On Error GoTo error

    'Tasa
    Tasa = CDbl(Mid(lblTasa.Caption, 1, Len(lblTasa.Caption) - 1)) / 100
    
    'Saco la Tasa ***********************************************************************
    rcConsulta.Open "SELECT tipoperiodo.Periodo,plazos.Descripcion AS VenPeriodo,Almacenaje,Seguro FROM configuraciontasas INNER JOIN tipoperiodo ON configuraciontasas.IDTipoPeriodo=tipoperiodo.ID INNER JOIN plazos ON configuraciontasas.IDPlazo=plazos.ID " _
                                    & "WHERE configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND configuraciontasas.IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex) & " AND configuraciontasas.IDPlazo=" & cmbPlazos.ItemData(cmbPlazos.ListIndex) _
                                    & " ORDER BY tipoperiodo.Ordenamiento,plazos.Descripcion", dbDatos, adOpenForwardOnly, adLockOptimistic

    Almacenaje = rcConsulta!Almacenaje / 100
    Seguro = rcConsulta!Seguro / 100
    Periodo = rcConsulta!Periodo
    Vencimiento = rcConsulta!VenPeriodo
    
    rcConsulta.Close
    '*************************************************************************************
    
    Select Case cmbPeriodo.text
    Case "MENSUAL"
        
        Meses = 1
    Case "QUINCENAL"
        
        Meses = 2
    Case "SEMANAL"
        
        Meses = 4
    
    Case "DIARIA"
    
        Meses = 30
    End Select
        
    'Saco los intereses
    If cmbTipoInteres.text = "TRADICIONAL" Or cmbTipoInteres.text = "COMPLETO" Or cmbTipoInteres.text = "TRAD COMPRA" Then
    
    TasaTradicional Date, Prestamo, Tasa, Almacenaje, Seguro, Regresa_Valor_BD("Iva") / 100, Periodo, Vencimiento, Meses
    Else
    
    TasaFija Prestamo, Tasa, Almacenaje, Seguro, Vencimiento, Periodo, Date
    End If
    
    Set rcConsulta = Nothing
    Exit Function
    
error:
    Maneja_Error Err
    Set rcTmp = Nothing
    Set rcConsulta = Nothing
End Function

Sub Encabezado(TipoInteres As Integer)
        
    With grdIntereses
        .Redraw = False
        .Clear True
        .AddColumn "K1", "Vencimiento", ecgHdrTextALignCentre, , 75, , , , , "DD/MMM/YYYY", , CCLSortDate
        .AddColumn "K2", "Interes", ecgHdrTextALignRight, , 78, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K3", "Almacenaje", ecgHdrTextALignRight, , 78, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K4", "Iva", ecgHdrTextALignRight, , 78, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K5", IIf(TipoInteres = 0, "Refrendo", "Pago Fijo"), ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", IIf(TipoInteres = 0, "Desempeño", "Saldo"), ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .Redraw = True
    End With

End Sub

Private Sub txtPrestamo_Change()
Dim crPrestamo As Double

On Error GoTo error

    If Val(txtPrestamo.text) = 0 Or Trim(txtPrestamo.text) = "" Then
        
        crPrestamo = 0
    Else
        
        crPrestamo = txtPrestamo.text
    End If
        
    lblAvaluo = Format(Redondeo((crPrestamo * 100) / Regresa_Valor_BD("PrestamoAvaluo")), FMoneda)
    
    'Checo la Tasa
    MuestraTasa crPrestamo, IsCliente

error:
    Maneja_Error Err
End Sub

Private Sub txtPrestamo_GotFocus()
    Seleccionar_Texto txtPrestamo
    Cambiar_Color True, txtPrestamo
End Sub

Private Sub txtPrestamo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrestamo_LostFocus()
    Cambiar_Color False, txtPrestamo
End Sub

Sub SacaTasa()
Dim rcConsulta As New ADODB.Recordset
Dim crPrestamo As Double

On Error GoTo error

    'Tasa
    rcConsulta.Open "SELECT tipoperiodo.Periodo,plazos.Descripcion AS Vencimiento " _
                    & "FROM configuraciontasas INNER JOIN plazos ON configuraciontasas.IDPlazo=plazos.ID INNER JOIN tipoperiodo ON configuraciontasas.IDTipoPeriodo=tipoperiodo.ID WHERE " _
                    & "configuraciontasas.IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND configuraciontasas.IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex) & " AND configuraciontasas.IDPlazo=" & cmbPlazos.ItemData(cmbPlazos.ListIndex), dbDatos, adOpenForwardOnly, adLockReadOnly
    
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        If Val(txtPrestamo.text) = 0 Or Trim(txtPrestamo.text) = "" Then
            
            crPrestamo = 0
        Else
            
            crPrestamo = CDbl(txtPrestamo.text)
        End If
        
        MuestraTasa crPrestamo, IsCliente
    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub

error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Function MuestraTasa(crPrestamo As Double, ExisteCliente As Boolean) As Double
Dim rcTasas As New ADODB.Recordset
Dim TasaTipica As Double, TasaPromocion As Double, TasaPreferencial As Double, LimiteInferior As Double, LimiteSuperior As Double
    
    TasaTipica = 0
    TasaPromocion = 0
    TasaPreferencial = 0
    LimiteInferior = 0
    LimiteSuperior = 0
    
    LimiteInferior = Regresa_Valor_BD("LimiteInferior")
    LimiteSuperior = Regresa_Valor_BD("LimiteSuperior")
    
    rcTasas.Open "SELECT TasaTipica,TasaPromocion,TasaPreferencial FROM configuraciontasas WHERE IDTipoInteres=" & cmbTipoInteres.ItemData(cmbTipoInteres.ListIndex) & " AND IDTipoPeriodo=" & cmbPeriodo.ItemData(cmbPeriodo.ListIndex) & " AND IDPlazo=" & cmbPlazos.ItemData(cmbPlazos.ListIndex), dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTasas.BOF And Not rcTasas.EOF Then
        
        TasaTipica = rcTasas!TasaTipica
        TasaPromocion = rcTasas!TasaPromocion
        TasaPreferencial = rcTasas!TasaPreferencial
    End If
    rcTasas.Close
    Set rcTasas = Nothing
    
    If ExisteCliente = False Then
        
        If crPrestamo >= LimiteSuperior Then
            
            MuestraTasa = TasaPreferencial
            
        ElseIf crPrestamo >= LimiteInferior Then
            
            MuestraTasa = TasaPromocion
        
        Else
            
            MuestraTasa = TasaTipica
        End If
    
    Else
        
        If crPrestamo >= LimiteInferior Then
            
            MuestraTasa = TasaPreferencial
        Else
            
            MuestraTasa = TasaPromocion
        End If
        
    End If
    
    lblTasa.Caption = Format(MuestraTasa, "0.00") & "%"
End Function

Sub TasaTradicional(Fecha As Date, crPrestamo As Double, Tasa As Double, Almacenaje As Double, Seguro As Double, Iva As Double, Periodo As Integer, Vencimiento As Integer, Division As Integer)
Dim i As Integer, crAlmacenaje As Double, crSeguro As Double, crInteres As Double
Dim FechaOriginal As Date, NumDias As Integer, crIva As Double

    grdIntereses.Redraw = False
    grdIntereses.Clear
    
    NumDias = 0
    FechaOriginal = Fecha
    For i = 1 To Vencimiento

        'Vencimiento
        If Periodo = 30 Then
            
            Fecha = DateAdd("M", i, FechaOriginal)
            NumDias = NumDias + Regresa_Ultimo_Dia_Mes(DateAdd("M", i - 1, FechaOriginal))
        Else
            
            NumDias = Periodo * i
            Fecha = DateAdd("D", Periodo, Fecha)
        End If
                                                
        'Almacenaje
        crAlmacenaje = Redondeo(crPrestamo * ((Almacenaje / Periodo) * NumDias))
  
        'Seguro
        crSeguro = Redondeo(crPrestamo * ((Seguro / Periodo) * NumDias))
        
        'Interes
        crInteres = Redondeo(crPrestamo * ((Tasa / Periodo) * NumDias))
        
        'Iva
        crIva = Redondeo((crInteres + crAlmacenaje + crSeguro) * Iva)

        With grdIntereses
            .AddRow
            .CellText(.Rows, 1) = Fecha
            .CellTextAlign(.Rows, 1) = DT_CENTER
            .CellText(.Rows, 2) = crInteres
            .CellTextAlign(.Rows, 2) = DT_RIGHT
            .CellText(.Rows, 3) = crAlmacenaje
            .CellTextAlign(.Rows, 3) = DT_RIGHT
            .CellText(.Rows, 4) = crIva
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            .CellText(.Rows, 5) = crInteres + crAlmacenaje + crSeguro + crIva
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            .CellText(.Rows, 6) = crPrestamo + crInteres + crAlmacenaje + crSeguro + crIva
            .CellTextAlign(.Rows, 6) = DT_RIGHT
        End With

    Next i

    SombreaGrid grdIntereses, 226, 220, 197, 238, 234, 221
    grdIntereses.Redraw = True

End Sub

Sub TasaFija(crPrestamo As Double, Tasa As Double, Almacenaje As Double, Seguro As Double, plazo As Integer, Periodo As Integer, Fecha As Date)
Dim SaldoInsoluto As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double
Dim Vencimiento As Date, i As Integer, crSaldo As Double, crImporteTotal As Double, crPagoFijo As Double, crAmortizacion As Double, strIntervalo As String
    
    crPrestamo = crPrestamo
    SaldoInsoluto = crPrestamo
    crImporteTotal = Redondeo(Pmt((Tasa + Almacenaje + Seguro), plazo, -crPrestamo, 0, 0), 1) * plazo
    crPagoFijo = Redondeo(Pmt((Tasa + Almacenaje + Seguro), plazo, -crPrestamo, 0, 0), 2)
    crSaldo = crImporteTotal
    strIntervalo = "D"
    Vencimiento = DateAdd("D", IIf(strIntervalo = "D", -1, 0), Fecha)
    
    If Periodo = 30 Then
        Periodo = 1
        strIntervalo = "M"
    End If
    
    With grdIntereses
    
        .Redraw = False
        .Clear
        
        For i = 1 To plazo
            
            Vencimiento = DateAdd(strIntervalo, Periodo, Vencimiento)
            crImporteTotal = crImporteTotal
            SaldoInsoluto = IIf(i = 1, crPrestamo, SaldoInsoluto - crAmortizacion)
            
            crIntereses = Redondeo(SaldoInsoluto * Tasa)
            crAlmacenaje = Redondeo(SaldoInsoluto * Almacenaje)
            crSeguro = Redondeo(SaldoInsoluto * Seguro)
            
            crAmortizacion = crPagoFijo - (crIntereses + crAlmacenaje + crSeguro)
            crSaldo = crSaldo - (crIntereses + crAlmacenaje + crSeguro + crAmortizacion)
                                                  
            .AddRow
            .CellText(.Rows, 1) = Vencimiento
            .CellTextAlign(.Rows, 1) = DT_CENTER
            .CellText(.Rows, 2) = crIntereses
            .CellTextAlign(.Rows, 2) = DT_RIGHT
            .CellText(.Rows, 3) = crAlmacenaje
            .CellTextAlign(.Rows, 3) = DT_RIGHT
            .CellText(.Rows, 4) = crSeguro
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            .CellText(.Rows, 5) = crPagoFijo
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            .CellText(.Rows, 6) = crSaldo
            .CellTextAlign(.Rows, 6) = DT_RIGHT
            
        Next i
        
        'Sombreo el Grid
        SombreaGrid grdIntereses, 226, 220, 197, 238, 234, 221
        
        'Cargo los totales en la última linea
        .AddRow
        .CellText(.Rows, 5) = crPagoFijo * plazo
        .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
        
        For i = 1 To .Columns
            .CellBackColor(.Rows, i) = RGB(223, 208, 102)
            .CellForeColor(.Rows, i) = &HFF0000
        Next i
            
        grdIntereses.Redraw = True
    
    End With
    
End Sub

Public Sub Cotizacion(crPrestamo As Double, crAvaluo As Double, TipoInteres As Integer, TipoPeriodo As Integer, TipoPlazo As Integer, Cliente As Boolean)
    
    lblAvaluo.Caption = Format(crAvaluo, FMoneda)
    Ban = True
    IsCliente = Cliente
    txtPrestamo.text = Format(crPrestamo, FMoneda)
    cmbTipoInteres.ListIndex = ComboInformacion(cmbTipoInteres, CLng(TipoInteres))
    cmbPeriodo.ListIndex = ComboInformacion(cmbPeriodo, CLng(TipoPeriodo))
    cmbPlazos.ListIndex = ComboInformacion(cmbPlazos, CLng(TipoPlazo))
    
    MuestraInteres crAvaluo, crPrestamo, Date
    
End Sub
