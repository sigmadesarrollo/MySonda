VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCierreDiario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Balance"
   ClientHeight    =   4905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7245
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCierreDiario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4905
   ScaleWidth      =   7245
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   7095
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nvo. Saldo:"
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
         TabIndex        =   24
         Top             =   2880
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.Label lblNvoSaldo 
         Alignment       =   1  'Right Justify
         Caption         =   "<Total>"
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
         Left            =   1800
         TabIndex        =   23
         Top             =   2880
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblSalida 
         Alignment       =   1  'Right Justify
         Caption         =   "<Total>"
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
         Left            =   4920
         TabIndex        =   22
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblEntrada 
         Alignment       =   1  'Right Justify
         Caption         =   "<Total>"
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
         Left            =   1920
         TabIndex        =   21
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Line Line7 
         BorderStyle     =   3  'Dot
         X1              =   6840
         X2              =   5400
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line6 
         BorderStyle     =   3  'Dot
         X1              =   6840
         X2              =   5400
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line5 
         BorderStyle     =   3  'Dot
         X1              =   3240
         X2              =   1800
         Y1              =   1800
         Y2              =   1800
      End
      Begin VB.Line Line4 
         BorderStyle     =   3  'Dot
         X1              =   3240
         X2              =   1800
         Y1              =   1845
         Y2              =   1845
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   7080
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cheques:"
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
         Left            =   135
         TabIndex        =   20
         Top             =   1440
         Width           =   1125
      End
      Begin VB.Label lblCheques 
         Alignment       =   1  'Right Justify
         Caption         =   "<Cheques>"
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
         Left            =   1920
         TabIndex        =   19
         Top             =   1440
         Width           =   1335
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Deposito:"
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
         Left            =   3855
         TabIndex        =   18
         Top             =   1440
         Width           =   1170
      End
      Begin VB.Label lblDeposito 
         Alignment       =   1  'Right Justify
         Caption         =   "<Deposito>"
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
         Left            =   5040
         TabIndex        =   17
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Caption         =   "<Total>"
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
         Left            =   4680
         TabIndex        =   16
         Top             =   2880
         Width           =   2175
      End
      Begin VB.Label lblHaber 
         Alignment       =   1  'Right Justify
         Caption         =   "<Haber>"
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
         Left            =   4920
         TabIndex        =   14
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label lblDebe 
         Alignment       =   1  'Right Justify
         Caption         =   "<Debe>"
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
         TabIndex        =   12
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label lblSaldoAnterior 
         Alignment       =   1  'Right Justify
         Caption         =   "<Saldo Anterior>"
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
         Left            =   2760
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   3840
         TabIndex        =   15
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Haber:"
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
         Left            =   3840
         TabIndex        =   13
         Top             =   840
         Width           =   825
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Debe:"
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
         TabIndex        =   11
         Top             =   840
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Saldo anterior:"
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
         Left            =   750
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1830
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6075
      TabIndex        =   26
      Top             =   4440
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
      Picture         =   "frmCierreDiario.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   4875
      TabIndex        =   27
      Top             =   4440
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
      Picture         =   "frmCierreDiario.frx":055E
   End
   Begin VB.Label lblCajero 
      AutoSize        =   -1  'True
      Caption         =   "<Cajero>"
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
      TabIndex        =   25
      Top             =   600
      Width           =   1050
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
      Left            =   5520
      TabIndex        =   4
      Top             =   120
      Width           =   1200
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
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Verifique bien el efectivo en caja antes de correr este proceso."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "El cierre diario es un proceso que no se puede efectuar mas de una vez."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AVISO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label Label3 
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
      Left            =   4680
      TabIndex        =   3
      Top             =   120
      Width           =   795
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
      TabIndex        =   2
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
   Begin VB.Shape Shape1 
      FillColor       =   &H0080FFFF&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   120
      Top             =   1440
      Visible         =   0   'False
      Width           =   2535
   End
End
Attribute VB_Name = "frmCierreDiario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Implements ISubclass

Private Sub cmdAceptar_Click()
    Realizar_Cierre
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Limpiar
    AttachMessage Me, Me.hWnd, WM_CONFIGURACION
    lblFecha.Caption = Format(Date, "DD/MMM/YYYY")
    lblSucursal.Caption = Sucursal.NombreComercial
    Cargar_Montos
    lblCajero.Caption = frmMDI.Usuario
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DetachMessage Me, Me.hWnd, WM_CONFIGURACION
End Sub

Private Property Let ISubclass_MsgResponse(ByVal RHS As SSubTimer6.EMsgResponse)
'
End Property

Private Property Get ISubclass_MsgResponse() As SSubTimer6.EMsgResponse
    ISubclass_MsgResponse = emrPreprocess
End Property

Private Function ISubclass_WindowProc(ByVal hWnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Select Case iMsg
    
    Case WM_CONFIGURACION:  'volvemos a cargar el nombre de la sucursal
        
        lblSucursal.Caption = Regresa_Valor("MONTEPIO", "Sucursal", "")
    End Select
    
End Function

Private Sub Limpiar()
Dim ctrl As Control
  
    For Each ctrl In Controls
        
        On Error Resume Next
        If TypeOf ctrl Is TextBox Then ctrl.text = ""
        If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
        On Error Resume Next
        ctrl.Tag = ""
    
    Next
  
End Sub

'Cargamos los montos para cada cuenta
Private Sub Cargar_Montos()
Dim rcMonto As New ADODB.Recordset
Dim crTotal As Currency, Importe1, Importe2, Importe3, Importe4, Dot As Currency, Fecha As String
    
    On Error GoTo Error
    
    Fecha = DateAdd("D", -1, Format(Date, "DD/MM/YY"))
  
    'Ponemos el saldo anterior
    rcMonto.Open "SELECT Saldo,Fecha FROM saldos ORDER BY Fecha DESC", dbDatos, adOpenStatic, adLockOptimistic
  
    If rcMonto.BOF Or rcMonto.EOF Then GoTo 125
  
    If rcMonto!Fecha = Date Then
        
        rcMonto.MoveNext
    End If

125:
   
    'Ponemos la cantidad del saldo anterior
    If Not rcMonto.EOF Then
        lblSaldoAnterior.Caption = Format(Format(rcMonto!Saldo & "", "Currency"), "###,###,###,##0.00")
        lblSaldoAnterior.Tag = Val(rcMonto!Saldo & "")
        crTotal = Val(rcMonto!Saldo & "")
    
    Else
        lblSaldoAnterior.Caption = Format(Format(0, "Currency"), "###,###,###,##0.00")
        lblSaldoAnterior.Tag = 0
        crTotal = 0
    End If
    rcMonto.Close
  
    'Ponemos la cantidad del debe
    rcMonto.Open "SELECT SUM(Importe)AS Total FROM auxiliar WHERE cuenta='110101' AND Fecha='" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenStatic, adLockOptimistic
        lblDebe.Caption = Format(Format(IIf(IsNull(rcMonto!Total), 0, rcMonto!Total), "Currency"), "###,###,###,##0.00")
        lblDebe.Tag = IIf(IsNull(rcMonto!Total), 0, rcMonto!Total)
    rcMonto.Close
  
    'Ponemos la cantidad del haber
    rcMonto.Open "SELECT SUM(Importe)AS Total FROM auxiliar WHERE cuenta='110150' AND Fecha='" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenStatic, adLockOptimistic
        lblHaber.Caption = Format(Format(IIf(IsNull(rcMonto!Total), 0, rcMonto!Total), "Currency"), "###,###,###,##0.00")
        lblHaber.Tag = IIf(IsNull(rcMonto!Total), 0, rcMonto!Total)
    rcMonto.Close
  
    'Depositos
    rcMonto.Open "SELECT SUM(Importe)AS Total FROM auxiliar WHERE cuenta='110950' AND Iniciales='BA50' AND Fecha='" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenStatic, adLockOptimistic
        lblDeposito.Caption = Format(Format(IIf(IsNull(rcMonto!Total), 0, rcMonto!Total), "Currency"), "###,###,###,##0.00")
        lblDeposito.Tag = IIf(IsNull(rcMonto!Total), 0, rcMonto!Total)
    rcMonto.Close
  
    'Cheques
    rcMonto.Open "SELECT SUM(Importe)AS Total FROM auxiliar WHERE cuenta='110901' AND Iniciales='BA01' AND Fecha='" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenStatic, adLockOptimistic
        lblCheques.Caption = Format(Format(IIf(IsNull(rcMonto!Total), 0, rcMonto!Total), "Currency"), "###,###,###,##0.00")
        lblCheques.Tag = IIf(IsNull(rcMonto!Total), 0, rcMonto!Total)
    rcMonto.Close
  
    lblEntrada.Caption = Format(Format(CCur(lblDebe.Caption) + CCur(lblCheques.Caption), "Currency"), "###,###,###,##0.00")
    lblSalida.Caption = Format(Format(CCur(lblHaber.Caption) + CCur(lblDeposito.Caption), "Currency"), "###,###,###,##0.00")
  
    crTotal = CCur(lblEntrada.Caption) - CCur(lblSalida.Caption)
  
    If lblHaber.Caption = "" Then lblHaber.Caption = "0.00"
  
    lblTotal.Caption = Format(Format(crTotal, "Currency"), "###,###,###,##0.00")
    lblTotal.Tag = crTotal
  
    lblNvoSaldo.Caption = Format(CCur(lblSaldoAnterior.Caption) + CCur(lblTotal.Caption), "###,###,##0.00")
  
Error:
    Maneja_Error Err
    Set rcMonto = Nothing
End Sub

'Realizamos el cierre diario
Private Sub Realizar_Cierre()
Dim crAjuste As Currency, Movimiento As Long, Bandera As Boolean
Dim rcCuentas As New ADODB.Recordset

    Screen.MousePointer = vbHourglass

    rcCuentas.Open "SELECT * FROM cuentas ORDER BY Cuenta", dbDatos, adOpenForwardOnly, adLockOptimistic
    
    Grabar_Datos
    Realizar_Reporte lblSucursal.Caption, frmMDI.Usuario
    Realizar_Cuentas
    Bandera = True
    
    With rcCuentas
    
        While Not .EOF
            
            Realizar_Diario lblSucursal.Caption, frmMDI.Usuario, !Mayor, !Cuenta, !Descripcion, Bandera
            Bandera = False
        .MoveNext
        Wend
  
    End With
  
    Sleep 1000
  
    Imprimir_Reportes
    
    rcCuentas.Close
    Set rcCuentas = Nothing

    Screen.MousePointer = vbDefault
End Sub

'Grabamos los datos
Private Sub Grabar_Datos()
Dim rcCD As New ADODB.Recordset
Dim Efectivo As Currency, Saldo As Currency, Debe As Currency
Dim Haber As Currency, Ajuste As Currency, crSaldo As Currency, crTotal As Currency
  
    rcCD.Open "SELECT * FROM cierrediario WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenDynamic, adLockOptimistic
    If rcCD.BOF = True And rcCD.EOF = True Then
        
        crSaldo = Format(CCur(Val(lblTotal.Tag)), "###########0.00")
  
        Efectivo = CCur(lblTotal.Caption) + CCur(lblSaldoAnterior.Caption)
        Saldo = CCur(Val(lblSaldoAnterior.Tag))
        Debe = CCur(Val(lblDebe.Tag))
        Haber = CCur(Val(lblHaber.Tag))
        crTotal = crSaldo - Ajuste
  
        dbDatos.Execute "INSERT INTO cierrediario (Fecha,Cajero,Saldo,Debe,Haber,Efectivo) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & frmMDI.Usuario & "'," & ConvMoneda(Saldo) & "," & ConvMoneda(Debe) & "," & ConvMoneda(Haber) & "," & ConvMoneda(Efectivo) & ")"
    End If
    rcCD.Close
    Set rcCD = Nothing
    
End Sub

'imprimimos el reporte
Private Sub Imprimir_Reportes()

On Error GoTo Error
  
    'Imprimimos el reporte de corte de caja
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\Balance.rpt"
        .Formulas(0) = "SaldoAnterior=" & ConvMoneda(lblSaldoAnterior.Caption) & ""
        .Formulas(1) = "Debe=" & ConvMoneda(lblDebe.Caption) & ""
        .Formulas(2) = "Haber=" & ConvMoneda(lblHaber.Caption) & ""
        .Formulas(3) = "Cheques=" & ConvMoneda(lblCheques.Caption) & ""
        .Formulas(4) = "Depositos=" & ConvMoneda(lblDeposito.Caption) & ""
        .Formulas(5) = "TotEntrada=" & ConvMoneda(lblEntrada.Caption) & ""
        .Formulas(6) = "TotSalida=" & ConvMoneda(lblSalida.Caption) & ""
        .Formulas(7) = "Total=" & ConvMoneda(lblTotal.Caption) & ""
        .Formulas(8) = "Efectivo=" & ConvMoneda(lblTotal.Caption) & ""
        .Formulas(9) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(10) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(11) = "Sucursal='" & Sucursal.NombreComercial & "'"
        .Formulas(12) = "Cajero='" & Trim(frmMDI.Usuario) & "'"
        .WindowTitle = "Balance"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub
