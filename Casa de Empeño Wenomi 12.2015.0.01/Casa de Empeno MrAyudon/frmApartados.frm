VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmApartados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apartados Vigentes"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13320
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmApartados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8040
   ScaleWidth      =   13320
   Begin vbAcceleratorGrid6.vbalGrid grdApartados 
      Height          =   7515
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   13256
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
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   12150
      TabIndex        =   1
      Top             =   7590
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmApartados.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   10950
      TabIndex        =   2
      Top             =   7590
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmApartados.frx":055E
   End
End
Attribute VB_Name = "frmApartados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdImprimir_Click()
    Imprimir
End Sub

Private Sub Imprimir()

On Error GoTo error

    If grdApartados.Rows > 0 Then
        
        With frmMDI.Cr
            
            .Reset
            .WindowShowPrintSetupBtn = True
            .DiscardSavedData = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\RepApartadosVig.rpt"
            .SelectionFormula = "DATEADD('D'," & Val(Regresa_Valor_BD("DiasGraciaApa") + 1) & ",{ventas.Vencimiento})>date('" & Format(Date, "YYYY/MM/DD") & "') AND {ventas.Pagado}=0 AND {ventas.Cancelado}=0 AND {ventas.Apartado}=1"
            .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
            .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(2) = "Totall=" & grdApartados.CellText(grdApartados.Rows, 5) & ""
            .Formulas(3) = "Saldoo=" & grdApartados.CellText(grdApartados.Rows, 7) & ""
            .Destination = crptToWindow
            .WindowTitle = "Apartados Vigentes"
            .WindowState = crptMaximized
            .Action = 1
        End With
    
    End If
    Exit Sub
    
error:
   Maneja_Error Err
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
    Crear_Encabezados
    Cargar_Apartados
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados
Private Sub Crear_Encabezados()

    With grdApartados
                
        .AddColumn "K3", "Fecha", ecgHdrTextALignCentre, , 120, , , , , "DD/MM/YY HH:MM:SS am/pm", , CCLSortDate
        .AddColumn "K1", "Folio", ecgHdrTextALignCentre, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Cliente", ecgHdrTextALignLeft, , 225, , , , , , , CCLSortString
        .AddColumn "K4", "Vencimiento", ecgHdrTextALignCentre, , 72, , , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K5", "Total", ecgHdrTextALignRight, , 81, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Abonos", ecgHdrTextALignRight, , 81, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Saldo", ecgHdrTextALignRight, , 81, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "Vendedor", ecgHdrTextALignLeft, , 140, , , , , , , CCLSortString
    
    End With

End Sub

'Cargamos los apartados vencidos
Private Sub Cargar_Apartados()
Dim rcApartado As New ADODB.Recordset
Dim DiasGracia As Integer, crAbonos As Double

On Error GoTo error
        
    DiasGracia = Val(Regresa_Valor_BD("DiasGraciaApa"))
    rcApartado.Open "SELECT ventas.ID,ventas.Folio,ventas.Fecha,ADDDATE(ventas.Vencimiento,INTERVAL " & DiasGracia & " DAY) AS Vencida,ventas.Vencimiento,ventas.Total,ventas.Descuento,ventas.Iva,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente,CONCAT(vendedores.Nombre,' ',vendedores.Apellidos) AS Vendedor " _
                    & "FROM ventas LEFT JOIN clientes ON ventas.IDCliente=clientes.ID LEFT JOIN vendedores ON ventas.IDVendedor=vendedores.ID WHERE ventas.Pagado=0 AND ventas.Apartado=1 AND ventas.Cancelado=0 AND DATE_FORMAT(ADDDATE(Vencimiento,INTERVAL " & DiasGracia & " DAY),'%Y%/%m%/%d')>='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY Fecha,Folio", dbDatos, adOpenForwardOnly, adLockReadOnly
       
    grdApartados.Redraw = False
    grdApartados.Clear
    With rcApartado
        
        While Not .EOF
            
            crAbonos = Regresa_Abonos(!ID)
            grdApartados.AddRow
            grdApartados.CellText(grdApartados.Rows, 1) = !Fecha
            grdApartados.CellTextAlign(grdApartados.Rows, 1) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdApartados.CellText(grdApartados.Rows, 2) = !Folio
            grdApartados.CellItemData(grdApartados.Rows, 2) = !ID
            grdApartados.CellTextAlign(grdApartados.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdApartados.CellText(grdApartados.Rows, 3) = !Cliente
            grdApartados.CellTextAlign(grdApartados.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdApartados.CellText(grdApartados.Rows, 4) = !Vencimiento
            grdApartados.CellTextAlign(grdApartados.Rows, 4) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdApartados.CellText(grdApartados.Rows, 5) = (!Total - (!Total * (!Descuento / 100))) * (1 + (!Iva / 100))
            grdApartados.CellTextAlign(grdApartados.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdApartados.CellText(grdApartados.Rows, 6) = crAbonos
            grdApartados.CellTextAlign(grdApartados.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdApartados.CellText(grdApartados.Rows, 7) = (!Total - (!Total * (!Descuento / 100))) - crAbonos
            grdApartados.CellTextAlign(grdApartados.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdApartados.CellText(grdApartados.Rows, 8) = !Vendedor
            
        .MoveNext
        Wend
        
    End With
    rcApartado.Close
    Set rcApartado = Nothing
    
    If grdApartados.Rows > 0 Then
        
        grdApartados.AddRow
        Poner_Totales
    End If
   
    grdApartados.Redraw = True
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcApartado = Nothing
End Sub

Private Sub Poner_Totales()
Dim TotalAbo As Currency, Abonos As Currency, Saldo As Currency
Dim Renglon As Integer, Total As Integer, Columna As Integer

On Error GoTo error

    'Hago la sumatoria de los totales (TotalAbo, Abonos, Saldos) desde el renglon 1 hasta el numero de renglones del GRID
    For Renglon = 1 To grdApartados.Rows - 1
        
        TotalAbo = TotalAbo + CCur(grdApartados.CellText(Renglon, 5))
        Abonos = Abonos + CCur(grdApartados.CellText(Renglon, 6))
        Saldo = Saldo + CCur(grdApartados.CellText(Renglon, 7))
        Total = Total + 1
    
    Next Renglon
         
    'En la ultima linea del GRID cargo los totales (TotalAbo, Abonos, Saldos) y cambio el color de la linea
    grdApartados.CellText(grdApartados.Rows, 7) = Saldo
    grdApartados.CellTextAlign(grdApartados.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdApartados.CellText(grdApartados.Rows, 5) = TotalAbo
    grdApartados.CellTextAlign(grdApartados.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdApartados.CellText(grdApartados.Rows, 6) = Abonos
    grdApartados.CellTextAlign(grdApartados.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdApartados.CellText(grdApartados.Rows, 2) = Total
    grdApartados.CellTextAlign(grdApartados.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS

    For Columna = 1 To grdApartados.Columns
        
        grdApartados.CellBackColor(grdApartados.Rows, Columna) = RGB(223, 208, 102)
    
    Next Columna
    Exit Sub
    
error:
    Maneja_Error Err
End Sub
