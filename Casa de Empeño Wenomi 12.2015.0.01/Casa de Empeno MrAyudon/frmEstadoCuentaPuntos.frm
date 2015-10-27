VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{1781610F-46E8-4DD3-922D-8DEF1A9DA567}#28.0#0"; "Credencial.ocx"
Begin VB.Form frmEstadoCuentaPuntos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Estado de Cuenta Puntos"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6525
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEstadoCuentaPuntos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4170
   ScaleWidth      =   6525
   Begin VB.TextBox txtNoTarjeta 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   1200
      MaxLength       =   60
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
   Begin Credencial.usCredencial DatosCliente 
      Height          =   2775
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   4895
      BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   6
      AlingHeader     =   0
      AlingBody       =   0
      BodyIndent      =   0
      HeaderIndent    =   0
      HeaderText      =   " Datos del cliente"
      HeaderBackColor =   16766131
      HeightHeader    =   22
      SidePicture     =   -1  'True
      SideBackColor   =   15000804
      WidthSide       =   25
      SidePicture     =   -1  'True
      HeaderBorderBackColor=   16744576
      BackColor       =   16777215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   3600
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Imprimir"
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
      Picture         =   "frmEstadoCuentaPuntos.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   3600
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
      Picture         =   "frmEstadoCuentaPuntos.frx":055E
   End
   Begin VB.Label lblPuntosAcumulados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   5280
      TabIndex        =   3
      Top             =   240
      Width           =   960
   End
   Begin VB.Label labePuntosAcumulados 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Puntos Acumulados:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3240
      TabIndex        =   2
      Top             =   240
      Width           =   1890
   End
   Begin VB.Label lblNoTarjeta 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No. Tarjeta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   1
      Top             =   240
      Width           =   1065
   End
End
Attribute VB_Name = "frmEstadoCuentaPuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TarjetaPuntos As New ClienteFrecuente
Dim FechaIni As String, FechaFin As String

Private Sub cmdImprimir_Click()
    
    Dim rsEstadoCuenta As New ADODB.Recordset
    Dim Sql As String, Saldo As Double
    
On Error GoTo error
    
    frmRangoFechas.Caption = "Estado de Cuenta"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    dbReportes.Execute "DELETE FROM estadocuentapuntos"
    
    rsEstadoCuenta.Open "SELECT SUM(cargo) AS Cargo FROM movimientospuntos WHERE IDTarjeta = " & TarjetaPuntos.CuentaFrecuente.IDCuenta & " AND DATE(fecha) < '" & Format(FechaIni, "YYYY-MM-DD") & "'", dbDatos, adOpenStatic, adLockOptimistic
        Saldo = IIf(IsNull(rsEstadoCuenta!Cargo), 0, rsEstadoCuenta!Cargo)
    rsEstadoCuenta.Close
    
    rsEstadoCuenta.Open "SELECT SUM(abono) AS Abono FROM movimientospuntos WHERE IDTarjeta = " & TarjetaPuntos.CuentaFrecuente.IDCuenta & " AND DATE(fecha) < '" & Format(FechaIni, "YYYY-MM-DD") & "'", dbDatos, adOpenStatic, adLockOptimistic
        Saldo = Saldo - IIf(IsNull(rsEstadoCuenta!Abono), 0, rsEstadoCuenta!Abono)
    rsEstadoCuenta.Close
    
    dbReportes.Execute "INSERT INTO estadocuentapuntos(Fecha,Folio,Movimiento,Saldo,IDCliente) VALUE('" & Format(FechaIni, "YYYY-MM-DD") & "',0,'SALDO INICIAL'," & Saldo & "," & TarjetaPuntos.CuentaFrecuente.IDCliente & ")"
    
    Sql = "SELECT *,CASE TipoMovimiento WHEN 0 THEN 'EMPEÑO' WHEN 1 THEN 'EMPEÑO AUTO' WHEN 2 THEN 'REFRENDO' WHEN 3 THEN 'REFRENDOEXT' WHEN 4 THEN 'DESEMPEÑO' WHEN 5 THEN 'VENTA' WHEN 6 THEN 'APARTADO' WHEN 7 THEN 'ABONO' WHEN 8 THEN 'EMPEÑO CANCELADO' WHEN 9 THEN 'EMPEÑO AUTO CANCELADO' WHEN 10 THEN 'REFRENDO CANCELADO' WHEN 11 THEN 'DESEMPEÑO CANCELADO' WHEN 12 THEN 'VENTA CANCELADO' WHEN 13 THEN 'APARTADO CANCELADO' WHEN 14 THEN 'ABONO CANCELADO' END AS Movimiento " & _
        "FROM movimientospuntos WHERE DATE(Fecha) >= '" & Format(FechaIni, "YYYY-MM-DD") & "' AND DATE(Fecha) <= '" & Format(FechaFin, "YYYY-MM-DD") & "' AND IDTarjeta = " & TarjetaPuntos.CuentaFrecuente.IDCuenta
        
    rsEstadoCuenta.Open Sql, dbDatos, adOpenStatic, adLockOptimistic
    
    With rsEstadoCuenta
        Do While Not .EOF
        
            Saldo = Saldo + !Cargo - !Abono
        
            dbReportes.Execute "INSERT INTO estadocuentapuntos(Fecha,Folio,Movimiento,Cargo,Abono,Saldo,IDCliente) VALUE('" & Format(!Fecha, "YYYY-MM-DD") & "'," & !Folio & ",'" & !Movimiento & "'," & !Cargo & "," & !Abono & _
                "," & Saldo & "," & TarjetaPuntos.CuentaFrecuente.IDCliente & ")"
        
            .MoveNext
        Loop
    End With
        
    rsEstadoCuenta.Close
    Set rsEstadoCuenta = Nothing
    
    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .ReportFileName = Path & "\Reportes\EstadoCuentaPuntos.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
    
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "PERIODO " & Format(FechaIni, "DD/MM/YYYY") & " A " & Format(FechaFin, "DD/MM/YYYY") & "'"
        .Formulas(3) = "NoCuenta='" & TarjetaPuntos.CuentaFrecuente.Folio & "'"
        
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Estado de Cuenta Puntos"
        .Action = 1
    End With
    
    Exit Sub
    
error:
    Set rsEstadoCuenta = Nothing
    Maneja_Error error
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
    
End Sub

Private Sub txtNoTarjeta_GotFocus()
    Seleccionar_Texto txtNoTarjeta
    Cambiar_Color True, txtNoTarjeta
End Sub

Private Sub txtNoTarjeta_KeyPress(KeyAscii As Integer)

    KeyAscii = Solo_Numeros(KeyAscii)
    
    If KeyAscii = vbKeyReturn Then
        
        Limpiar
        
        If TarjetaPuntos.CuentaFrecuente.FindCuentaByFolio(txtNoTarjeta.text) Then
            lblPuntosAcumulados.Caption = TarjetaPuntos.CuentaFrecuente.Puntos
            Buscar_Cliente TarjetaPuntos.CuentaFrecuente.IDCliente
        Else
            lblPuntosAcumulados.Caption = "0"
            Seleccionar_Texto txtNoTarjeta
            MsgBox "No se encuentra la tarjeta de cliente frecuente", vbOKOnly Or vbInformation
        End If
      
    End If
   
End Sub

Private Sub txtNoTarjeta_LostFocus()
    Cambiar_Color False, txtNoTarjeta
End Sub

Public Sub Buscar_Cliente(ID As Long)

    Dim rcClientes As New ADODB.Recordset

On Error GoTo error

    rcClientes.Open "SELECT * FROM clientes WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
    
    With rcClientes
    
        DatosCliente.Tag = !ID
        DatosCliente.Add "<bold> " & !Nombre & " " & !Apellido & "</bold>"
        DatosCliente.Add " " & !Direccion & " " & !Colonia & vbCrLf & _
            " " & !Municipio & ", " & !Estado & " C.P. " & !CP & vbCrLf
                
        DatosCliente.Add " NO. TARJETA: " & TarjetaPuntos.CuentaFrecuente.Folio & vbCrLf & _
            " FECHA TARJETA: " & TarjetaPuntos.CuentaFrecuente.FechaTarjeta
    
    End With
    
    rcClientes.Close
    Set rcClientes = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcClientes = Nothing
End Sub

Private Sub Limpiar()
    DatosCliente.Clear
    TarjetaPuntos.CuentaFrecuente.Clear
End Sub
