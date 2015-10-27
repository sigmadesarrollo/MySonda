VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{ADD24EDC-ADC1-11D2-95D1-F7A835DD4948}#3.0#0"; "nslock15vb5.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#12.1#0"; "Codejock.CommandBars.v12.1.1.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm frmMDI 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Casa de Empeño Mr. Ayudón"
   ClientHeight    =   9015
   ClientLeft      =   60
   ClientTop       =   -945
   ClientWidth     =   9825
   Icon            =   "frmMDI.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMDI.frx":BB08
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog Dlgopen 
      Left            =   4080
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Crystal.CrystalReport Cr 
      Left            =   3120
      Top             =   1560
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
   Begin nslock15vb5.ActiveLock ActiveLock2 
      Left            =   2520
      Tag             =   "15"
      Top             =   1560
      _ExtentX        =   847
      _ExtentY        =   820
      Password        =   "SondaMrAyudon"
      SoftwareName    =   "Sonda"
      LiberationKeyLength=   16
      SoftwareCodeLength=   16
   End
   Begin MSComctlLib.ProgressBar Bar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   8760
      Visible         =   0   'False
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSCommLib.MSComm Com 
      Left            =   1920
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Timer tmrHora 
      Interval        =   500
      Left            =   3600
      Top             =   1560
   End
   Begin vbalIml6.vbalImageList img 
      Left            =   720
      Top             =   1560
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   3444
      Images          =   "frmMDI.frx":1DA6B
      Version         =   131072
      KeyCount        =   3
      Keys            =   "26ÿÿ"
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   1320
      Top             =   1560
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1E7FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1EB19
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":1EE33
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":20B3D
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":20E57
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":213F1
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMDI.frx":2198B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager Imagenes 
      Left            =   5520
      Top             =   1560
      _Version        =   786433
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMDI.frx":21B7B
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   5040
      Top             =   1560
      _Version        =   786433
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMDI.frx":43E43
   End
   Begin XtremeCommandBars.CommandBars CommandBars 
      Left            =   4560
      Top             =   1560
      _Version        =   786433
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   6
   End
   Begin VB.Menu mnuEmpeños 
      Caption         =   "&Empeños"
      Begin VB.Menu mnuEmpeñoss 
         Caption         =   "Empeños"
      End
      Begin VB.Menu mnuPagoEmpeños 
         Caption         =   "Pago de Empeños"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuDesempeños 
         Caption         =   "Desempeños"
      End
      Begin VB.Menu mnuRefrendos 
         Caption         =   "Refrendos"
      End
      Begin VB.Menu mnuRefrendosForaneos 
         Caption         =   "Refrendos Foraneos"
      End
      Begin VB.Menu mnuPagosFijoss 
         Caption         =   "Pagos Fijos"
      End
      Begin VB.Menu mnuCotizaciones 
         Caption         =   "Cotizar Empeño"
      End
      Begin VB.Menu mnuCambioPlan 
         Caption         =   "Cambiar Plan"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBuscarBoletas 
         Caption         =   "Búsqueda de Contratos"
      End
      Begin VB.Menu mnuRegUbicacion 
         Caption         =   "Registrar ubicación"
      End
      Begin VB.Menu mnuReportesEmpeños 
         Caption         =   "Reportes"
         Begin VB.Menu mnuRepEmpeños 
            Caption         =   "Empeños"
         End
         Begin VB.Menu mnuRepEmpeVencidos 
            Caption         =   "Contratos Vencidos"
         End
         Begin VB.Menu mnuReporteAlmoneda 
            Caption         =   "Contratos Almoneda"
         End
      End
      Begin VB.Menu mnuLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSalir 
         Caption         =   "Salir"
         Shortcut        =   ^{F4}
      End
   End
   Begin VB.Menu mnuCierreS 
      Caption         =   "&Cierres"
      Begin VB.Menu mnuCierreCaja 
         Caption         =   "Cierre de Caja"
      End
      Begin VB.Menu mnuCierreDivisas 
         Caption         =   "Cierre de Divisas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRepCartera 
         Caption         =   "Cartera"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBalance 
         Caption         =   "Balance"
      End
      Begin VB.Menu mnuActivos 
         Caption         =   "Activos"
      End
      Begin VB.Menu mnuCierresSucursal 
         Caption         =   "Cierre de Sucursal"
      End
   End
   Begin VB.Menu mnuVentas 
      Caption         =   "&Ventas"
      Begin VB.Menu mnuMostrador 
         Caption         =   "Mostrador"
      End
      Begin VB.Menu mnuVentasClientes 
         Caption         =   "Clientes"
      End
      Begin VB.Menu mnuApartados 
         Caption         =   "Apartados"
      End
      Begin VB.Menu mnuAbonos 
         Caption         =   "Abonos"
      End
      Begin VB.Menu mnuPagoDemasias 
         Caption         =   "Demasías"
      End
      Begin VB.Menu mnuDevolucionVta 
         Caption         =   "Devolucion Ventas"
      End
      Begin VB.Menu mnuApartadosVencidos 
         Caption         =   "Apartados Vencidos"
      End
      Begin VB.Menu mnuReporteVentas 
         Caption         =   "Reportes"
         Begin VB.Menu mnuRepAnaVentas 
            Caption         =   "Análisis Ventas"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRepVentasCon 
            Caption         =   "Ventas Mostrador"
         End
         Begin VB.Menu mnuRepVentaCliente 
            Caption         =   "Ventas Cliente"
         End
         Begin VB.Menu mnuRepVentasApa 
            Caption         =   "Ventas Apartado"
         End
         Begin VB.Menu mnuMuestraApartados 
            Caption         =   "Apartados Vigentes"
         End
         Begin VB.Menu mnuAparRem 
            Caption         =   "Apartados Rematados"
         End
         Begin VB.Menu mnuRepUtilidadVentas 
            Caption         =   "Utilidad Ventas"
         End
         Begin VB.Menu mnuRepComiVen 
            Caption         =   "Comisión Ventas"
         End
      End
   End
   Begin VB.Menu mnuInventario 
      Caption         =   "&Inventario"
      Begin VB.Menu mnuEntradasInventario 
         Caption         =   "Entrada a Inventario"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuExistencias 
         Caption         =   "Existencias"
      End
      Begin VB.Menu mnuInventarioFund 
         Caption         =   "Inventario Fundición"
      End
      Begin VB.Menu mnuCompraJoyeria 
         Caption         =   "Compras/Dotaciónes"
      End
      Begin VB.Menu mnuSalidasInventario 
         Caption         =   "Salida de Inventario"
      End
      Begin VB.Menu mnuInvenFisico 
         Caption         =   "Inventario Físico"
      End
      Begin VB.Menu mnuDeslotificacion 
         Caption         =   "Deslotificación"
      End
      Begin VB.Menu mnuEtiquetas 
         Caption         =   "Impresión Etiquetas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuEtiAlmoneda 
         Caption         =   "Impresión Etiquetas"
      End
      Begin VB.Menu mnuTrasInventario 
         Caption         =   "Traspaso de Inventario"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuReporteInventario 
         Caption         =   "Reportes"
         Begin VB.Menu mnuCompras 
            Caption         =   "Compras"
         End
         Begin VB.Menu mnuRepEntradasInven 
            Caption         =   "Dotaciones"
         End
         Begin VB.Menu mnuRepSalInventario 
            Caption         =   "Salidas de Inventario"
         End
         Begin VB.Menu mnuRepTraspasos 
            Caption         =   "Traspasos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRepAntiguedad 
            Caption         =   "Antiguedad"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuReEnvejecimiento 
            Caption         =   "Envejecimiento"
         End
         Begin VB.Menu mnuRepEnvejecimientoPeriodo 
            Caption         =   "Envejecimiento por periodo"
         End
      End
   End
   Begin VB.Menu mnuDivisass 
      Caption         =   "&Divisas"
      Visible         =   0   'False
      Begin VB.Menu mnuCatDivisas 
         Caption         =   "Catálogo"
      End
      Begin VB.Menu mnuCotizacionDivisas 
         Caption         =   "Cotización diaria"
      End
      Begin VB.Menu mnuCompraVenta 
         Caption         =   "Compra/Venta"
      End
      Begin VB.Menu mnuRepDivisas 
         Caption         =   "Reportes"
         Begin VB.Menu mnuRepComVenDiv 
            Caption         =   "Compra/Venta"
         End
         Begin VB.Menu mnuRepExistenciasDiv 
            Caption         =   "Existencias"
         End
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuGrupRepEmpenos 
         Caption         =   "Empeños"
         Begin VB.Menu mnuRepEmpeñoss 
            Caption         =   "Empeños"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuRepHistorico 
            Caption         =   "Histórico"
         End
         Begin VB.Menu mnuRepDetallado 
            Caption         =   "Detallado"
         End
         Begin VB.Menu mnuRepInventario 
            Caption         =   "Depositaría"
         End
         Begin VB.Menu mnuRepVencidos 
            Caption         =   "Contratos Vencidos"
         End
         Begin VB.Menu mnuRepAlmoneda 
            Caption         =   "Contratos Almoneda"
         End
         Begin VB.Menu mnuCierreRepEmpeños 
            Caption         =   "Empeños"
         End
         Begin VB.Menu mnuRepCieDesempeño 
            Caption         =   "Desempeños"
         End
         Begin VB.Menu mnuRepCieRefrendos 
            Caption         =   "Refrendos"
         End
         Begin VB.Menu mnuRepPagosFijos 
            Caption         =   "Pagos Fijos"
         End
      End
      Begin VB.Menu mnuGrupRepFinancieros 
         Caption         =   "Financieros"
         Begin VB.Menu mnuRepAuxiliares 
            Caption         =   "Auxiliares"
         End
         Begin VB.Menu mnuRepContable 
            Caption         =   "Contable"
         End
         Begin VB.Menu mnuRepAuditoria 
            Caption         =   "Auditoría"
         End
         Begin VB.Menu mnuRepGastos 
            Caption         =   "Gastos"
         End
         Begin VB.Menu mnuRepIngresos 
            Caption         =   "Ingresos"
         End
      End
      Begin VB.Menu mnuGrupRepMonitoreo 
         Caption         =   "Monitoreo"
         Begin VB.Menu mnuRepOperaciones 
            Caption         =   "Movimientos por Horario"
         End
         Begin VB.Menu mnuRepAutorizaciones 
            Caption         =   "Autorizaciones"
         End
         Begin VB.Menu mnuRepPartidasBoveda 
            Caption         =   "Partidas en Bóveda"
         End
         Begin VB.Menu mnuRepAseguradora 
            Caption         =   "Aseguradora"
         End
         Begin VB.Menu mnuRepCancelaciones 
            Caption         =   "Cancelaciones"
         End
      End
      Begin VB.Menu mnuGrupRepPrestPromedio 
         Caption         =   "Contratos promedio"
         Begin VB.Menu mnuRepEmpeMes 
            Caption         =   "Empeños"
         End
         Begin VB.Menu mnuRepDesMes 
            Caption         =   "Desempeños"
         End
         Begin VB.Menu mnuRepRefMes 
            Caption         =   "Refrendos"
         End
      End
      Begin VB.Menu mnuRepGraficos 
         Caption         =   "Gráficos"
         Begin VB.Menu mnuRepEmpenosTipoTasa 
            Caption         =   "Contratos tipo tasa"
         End
         Begin VB.Menu mnuRepEmpenosVencidos 
            Caption         =   "Contratos Vencidos"
         End
         Begin VB.Menu mnuRepPrestamoStatus 
            Caption         =   "Contratos por status"
         End
         Begin VB.Menu mnuRepPrestamosMes 
            Caption         =   "Préstamos por mes"
         End
         Begin VB.Menu mnuRepMedios 
            Caption         =   "Medios difusión"
         End
      End
      Begin VB.Menu mnuRepClienteFrecuente 
         Caption         =   "Cliente Frecuente"
         Begin VB.Menu mnuRepClienteFrecuenteMovimientos 
            Caption         =   "Movimientos"
         End
         Begin VB.Menu mnuRepClienteFrecuentaEdoCuenta 
            Caption         =   "Estado de Cuenta"
         End
      End
   End
   Begin VB.Menu mnuMovimientos 
      Caption         =   "&Movimientos"
      Begin VB.Menu mnuCaja 
         Caption         =   "Caja General"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuCajaGralDivisas 
         Caption         =   "Caja General Divisas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBancos 
         Caption         =   "Bancos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuBoveda 
         Caption         =   "Bóveda/Caja"
      End
      Begin VB.Menu mnuMovDivisas 
         Caption         =   "Bóveda Divisas"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuTraspasos 
         Caption         =   "Traspasos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRegGastos 
         Caption         =   "Gastos"
      End
      Begin VB.Menu mnuCargosAbonos 
         Caption         =   "Cargos abonos"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuRemates 
         Caption         =   "Pasar a Almoneda"
      End
      Begin VB.Menu mnuCancelaciones 
         Caption         =   "Cancelar Movimientos"
      End
      Begin VB.Menu mnuConexInter 
         Caption         =   "Conexión Intersucursal"
      End
      Begin VB.Menu mnuContabilidad 
         Caption         =   "Contabilidad"
         Begin VB.Menu mnuRelCuentas 
            Caption         =   "Relacionar Cuentas"
         End
         Begin VB.Menu mnuCrearPolizas 
            Caption         =   "Crear Pólizas"
         End
         Begin VB.Menu mnuPolizaDia 
            Caption         =   "Póliza del día"
         End
      End
   End
   Begin VB.Menu mnuAdmonLavadoDinero 
      Caption         =   "&Módulo Antilavado"
      Begin VB.Menu mnuLLDParametros 
         Caption         =   "Configuración de Parametros"
      End
      Begin VB.Menu mnuLLDMovAtipicos 
         Caption         =   "Movimientos Atípicos"
      End
      Begin VB.Menu mnuLLDExpClientes 
         Caption         =   "Expedientes de Clientes"
      End
      Begin VB.Menu mnuLLDRepMensual 
         Caption         =   "Reporte Mensual Pormenorizado"
      End
   End
   Begin VB.Menu mnuUtilerias 
      Caption         =   "&Utilerias"
      Begin VB.Menu mnuConfiguracion 
         Caption         =   "Configuración"
         Begin VB.Menu mnuParametros 
            Caption         =   "Parámetros"
         End
         Begin VB.Menu mnuConfTasas 
            Caption         =   "Tasas"
         End
         Begin VB.Menu mnuPreciosOro 
            Caption         =   "Precios Oro/Plata"
         End
         Begin VB.Menu mnuPreciosDiamante 
            Caption         =   "Precios Diamante"
         End
         Begin VB.Menu mnuUsuarios 
            Caption         =   "Usuarios"
         End
         Begin VB.Menu mnuSucursales 
            Caption         =   "Sucursales"
         End
         Begin VB.Menu mnuClienteFrecuente 
            Caption         =   "Cliente Frecuente"
            Begin VB.Menu mnuClienteFrecuenteParametros 
               Caption         =   "Parametros"
            End
            Begin VB.Menu mnuClienteFrecuenteTiposTarjeta 
               Caption         =   "Tipos Tarjeta"
            End
            Begin VB.Menu mnuReasignarT 
               Caption         =   "Reasignar Tarjeta"
            End
         End
         Begin VB.Menu mnuImpresoras 
            Caption         =   "Impresoras"
         End
         Begin VB.Menu mnuPromociones 
            Caption         =   "Promociones"
         End
      End
      Begin VB.Menu mnuGeneraAutoriza 
         Caption         =   "Generar autorización"
      End
      Begin VB.Menu mnuFacturacion 
         Caption         =   "Facturación"
      End
      Begin VB.Menu mnuCapturaBoletas 
         Caption         =   "Captura de contratos"
      End
      Begin VB.Menu mnuRegSoftware 
         Caption         =   "Registrar Software"
      End
      Begin VB.Menu mnuMensajesBol 
         Caption         =   "Mensajes contratos"
      End
      Begin VB.Menu mnuRepInfo 
         Caption         =   "Respaldar información"
      End
      Begin VB.Menu mnuConfImpresiones 
         Caption         =   "Impresiones"
      End
      Begin VB.Menu mnuCalculadora 
         Caption         =   "Calculadora"
      End
      Begin VB.Menu mnuCatalogos 
         Caption         =   "Catálogos"
         Begin VB.Menu mnuClientes 
            Caption         =   "Clientes"
         End
         Begin VB.Menu mnuCatVendedores 
            Caption         =   "Vendedores"
         End
         Begin VB.Menu mnuConceptos 
            Caption         =   "Conceptos"
         End
         Begin VB.Menu mnuMedios 
            Caption         =   "Medios"
         End
         Begin VB.Menu mnuCuentaGastos 
            Caption         =   "Cuentas gastos"
         End
         Begin VB.Menu mnuTipos 
            Caption         =   "Tipos"
            Visible         =   0   'False
         End
         Begin VB.Menu mnuCatTiposPrenda 
            Caption         =   "Prendas Joyería"
         End
         Begin VB.Menu mnuCatElectronicos 
            Caption         =   "Prendas Varios"
            Begin VB.Menu mnuFamilias 
               Caption         =   "Familia"
            End
            Begin VB.Menu mnuCatMarcas 
               Caption         =   "Marcas"
            End
            Begin VB.Menu mnuCatPrendasVarios 
               Caption         =   "Prendas"
            End
         End
      End
      Begin VB.Menu mnuMigracion 
         Caption         =   "Migracion"
      End
   End
   Begin VB.Menu mnuAcercaDe 
      Caption         =   "&Acerca de..."
      Begin VB.Menu mnuSistema 
         Caption         =   "Sistema"
      End
      Begin VB.Menu mnuManualUsuario 
         Caption         =   "Manual de Usuario"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema de casa de Empeno montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 02/04/02
' Modulo Principal frmMDI
' Ultima Modificacion - 06/08/02
' Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit
Dim m_Aparatos As Boolean
Dim m_Tipo As String
Dim m_Fecha1 As String
Dim m_Fecha2 As String
Dim m_Usuario As String
Dim m_IDUsuario As Long
Dim p_IDSucursal As Integer
Public FechaIni As String, FechaFin As String
Public PaseAlmoneda As Boolean
Dim TiempoActualizacion As Long

Public Property Let IDSucursal(Valor As Integer)
    p_IDSucursal = Valor
End Property

Public Property Get IDSucursal() As Integer
    IDSucursal = p_IDSucursal
End Property

Public Property Let Usuario(Valor As String)
    m_Usuario = Valor
End Property

Public Property Let IDUsuario(Valor As Long)
    m_IDUsuario = Valor
End Property

Public Property Get Usuario() As String
    Usuario = m_Usuario
End Property

Public Property Get IDUsuario() As Long
    IDUsuario = m_IDUsuario
End Property

Public Property Let Fecha1(Valor As String)
    m_Fecha1 = Valor
End Property

Public Property Let Fecha2(Valor As String)
    m_Fecha2 = Valor
End Property

'Realizamos el reporte pasando los totales del mes
Public Sub Realizar_Reporte_Mensual(Sucursal As String, Cajero As String)
Dim rcMonto As New ADODB.Recordset
Dim crSaldo As Currency, crDebe As Currency, crHaber As Currency
  
    dbReportes.Execute "DELETE FROM CorteCajaVentanilla"
  
    'Ponemos el saldo anterior
    rcMonto.Open "SELECT Saldo FROM Saldos WHERE MONTH(Fecha) < " & Format(Date, "MM") & " AND YEAR(Fecha)<=  " & Format(Date, "YYYY") & " ORDER BY Fecha DESC", dbDatos, adOpenStatic, adLockOptimistic
    'Ponemos la cantidad del debe
    crSaldo = Val(rcMonto!Saldo & "")
    rcMonto.Close
  
    rcMonto.Open "SELECT SUM(Importe)AS Total FROM Auxiliar WHERE Cuenta='110101' AND MONTH(Fecha)=" & Format(Date, "MM") & " AND YEAR(Fecha)=" & Format(Date, "YYYY"), dbDatos, adOpenStatic, adLockOptimistic
    'Ponemos la cantidad del debe
    crDebe = Val(rcMonto!Total & "")
    rcMonto.Close
  
    'Ponemos la cantidad del haber
    rcMonto.Open "SELECT SUM(Importe)AS Total FROM Auxiliar WHERE Cuenta='110150' AND MONTH(Fecha)=" & Format(Date, "MM") & "AND YEAR(Fecha)=" & Format(Date, "YYYY"), dbDatos, adOpenStatic, adLockOptimistic
    'Ponemos la cantidad del debe
    crHaber = Val(rcMonto!Total & "")
    rcMonto.Close
  
    dbReportes.Execute "INSERT INTO CorteCajaVentanilla (Sucursal,Cajero,Saldo,Debe,Haber) VALUES " & _
                        "('" & Sucursal & "','" & Cajero & "'," & crSaldo & "," & crDebe & "," & crHaber & ")"
  
Error:
    Maneja_Error Err
    Set rcMonto = Nothing
End Sub

'ponemos las cuentas en el reporte y las separamos del mes
Private Sub Realizar_Cuentas_Mensual()
Dim rcAuxiliar As New ADODB.Recordset
  
    dbReportes.Execute "DELETE FROM CorteCuentas"
    
    rcAuxiliar.Open "SELECT PC,Cuenta,Importe,Concepto,Folio,Movimiento FROM Auxiliar WHERE MONTH(Fecha)=" & Format(Date, "MM") & " AND YEAR(Fecha)=" & Format(Date, "YYYY"), dbDatos, adOpenForwardOnly, adLockOptimistic
    With rcAuxiliar
        
        While Not .EOF

            DoEvents
            dbReportes.Execute "INSERT INTO CorteCuentas (PC,Cuenta,Descripcion,Folio,Movimientos,Cargo,Abono) VALUES " & "('" & !PC & "','" & !Cuenta & "','" & !Concepto & "'," & !Folio & "," & !Movimiento & "," & IIf(Right(!Cuenta, 2) = "01", !Importe, 0) & "," & IIf(Right(!Cuenta, 2) = "50", !Importe, 0) & ")"
        .MoveNext
        Wend
    
    End With
    rcAuxiliar.Close
  
Error:
    Maneja_Error Err
    Set rcAuxiliar = Nothing
End Sub

'Creamos el reporte Mensual
Private Sub Realizar_Mensual(Sucursal As String, Cajero As String, Mayor As String, Cuenta As String, Leyenda As String, Optional Opcion As Boolean = False)
Dim rcAuxiliar As New ADODB.Recordset
Dim lFolio1 As Long, lFolio2 As Long, lFolio3 As Long
Dim crImporte1 As Currency, crImporte2 As Currency, crImporte3 As Currency
  
  If Opcion Then dbReportes.Execute "DELETE * FROM Diario"
  
  rcAuxiliar.Open "SELECT * FROM Auxiliar WHERE Cuenta='" & Cuenta & "' AND MONTH(Fecha)=" & Format(Date, "MM") & " AND YEAR(Fecha)=" & Format(Date, "YYYY") & " ORDER BY Folio", dbDatos, adOpenForwardOnly, adLockOptimistic
  
  With rcAuxiliar
    While Not .EOF
      DoEvents
      lFolio1 = 0
      lFolio2 = 0
      lFolio2 = 0
      crImporte1 = 0
      crImporte2 = 0
      crImporte3 = 0
    
    
      lFolio1 = !Folio
      crImporte1 = !Importe
      .MoveNext
      If Not .EOF Then
        lFolio2 = !Folio
        crImporte2 = !Importe
        .MoveNext
      End If
      If Not .EOF Then
        lFolio3 = !Folio
        crImporte3 = !Importe
        .MoveNext
      End If
      dbReportes.Execute "INSERT INTO Diario (Sucursal,Cajero,Cuenta,Leyenda,Importe1,Folio1,Importe2,Folio2,Importe3,Folio3) VALUES " & _
                         "('" & Sucursal & "','" & Cajero & "','" & Cuenta & "','" & Leyenda & "'," & crImporte1 & "," & lFolio1 & "," & crImporte2 & "," & lFolio2 & "," & crImporte3 & "," & lFolio3 & ")"
      
    Wend
    .Close
  End With
    
End Sub

Private Sub MDIForm_Activate()
Dim Remate As Integer
Dim FechaAlmoneda As Date
'PaseAlmoneda = True
On Error GoTo Error

    If PaseAlmoneda = False Then
        FechaAlmoneda = Regresa_Valor_BD("FechaAlmoneda")
        
        Remate = Val(SacaValor("rematediario", "Status", " WHERE DATE_FORMAT(rematediario.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "'"))
        If Remate = 0 And Format(Date, "YYYY-MM-DD") = Format(FechaAlmoneda, "YYYY-MM-DD") Then

            If Val(SacaValor("empeno", "COUNT(empeno.ID)", " WHERE (empeno.Serie=" & SERIE_A & " OR empeno.Serie=" & SERIE_C & " OR empeno.Serie=" & SERIE_B & ") AND empeno.Cancelado=0 AND DATE_FORMAT(Vencimiento,'%Y%/%m%/%d') <='" & Format(FechaAlmoneda, "YYYY/MM/DD") & "' AND empeno.Destino=0 AND empeno.Pagado=0")) > 0 Then

                If MsgBox("Se realizará el pase a Destino AAutomático desea continuar ??", vbQuestion + vbYesNo + vbDefaultButton1, "Pase a Destino") = vbYes Then

                    Pase_Automatico_Almoneda
                    dbDatos.Execute "INSERT INTO rematediario (Fecha,Status) VALUES('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "',1)"

                    PaseAlmoneda = True
                Else

                    End
                End If

            Else

                dbDatos.Execute "INSERT INTO rematediario (Fecha,Status) VALUES('" & _
                                Format(Now, "YYYY/MM/DD HH:MM:SS") & "',1)"
                PaseAlmoneda = True

            End If

        Else

            PaseAlmoneda = True
        End If

    End If
    Exit Sub

Error:
    Maneja_Error Err
End Sub

Private Sub mnuAbonos_Click()
    frmVentas.Show
    frmVentas.tTab.SelectTab "K3"
    frmVentas.frmPagos.Visible = True
    frmVentas.frmApartados.Visible = False
    frmVentas.frmVentasMostrador.Visible = False
    BringWindowToTop frmVentas.hWnd
End Sub

Private Sub mnuActivos_Click()
    frmReporteFinanciero.Show
    BringWindowToTop frmReporteFinanciero.hWnd
End Sub

Private Sub mnuAparRem_Click()
    
    frmRangoFechas.Caption = "Reporte Apartados Rematados"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
       
    With Cr
       .Reset
       .WindowShowPrintSetupBtn = True
       .WindowShowExportBtn = True
       .DiscardSavedData = True
       .ReportFileName = Path & "\Reportes\RepApartadosVencidos.rpt"
       .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
       .SelectionFormula = "{ventas.FechaMovimiento}>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {ventas.FechaMovimiento}<=date('" & Format(FechaFin, "YYYY,MM,DD" & "'") & ")"
       .Formulas(0) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
       .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
       .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
       .WindowTitle = "Reporte apartados rematados"
       .WindowState = crptMaximized
       .Destination = crptToWindow
       .Action = 1
    End With
End Sub

Private Sub mnuApartados_Click()
    frmVentas.Show
    frmVentas.tTab.SelectTab "K2"
    frmVentas.frmPagos.Visible = False
    frmVentas.frmApartados.Visible = True
    frmVentas.frmVentasMostrador.Visible = False
    BringWindowToTop frmVentas.hWnd
End Sub

Private Sub mnuApartadosVencidos_Click()
    frmApartadosVencidos.Show
    BringWindowToTop frmApartadosVencidos.hWnd
End Sub

Private Sub mnuBalance_Click()
    frmCierreDiario.Show
    BringWindowToTop frmCierreDiario.hWnd
End Sub

Private Sub mnuBancos_Click()
    frmMovimiento.Show
    BringWindowToTop frmMovimiento.hWnd
End Sub

Private Sub mnuBoveda_Click()
    frmMovimiento.Show
    BringWindowToTop frmMovimiento.hWnd
End Sub

Private Sub mnuBuscarBoletas_Click()
    frmBusqueda.Show
    BringWindowToTop frmBusqueda.hWnd
End Sub

Private Sub mnuCaja_Click()
Dim crImporte As Double
    
    crImporte = 0
    crImporte = frmDenominaciones.Arqueo
    If crImporte > 0 Then
        
        frmBoveda.txtImporte.text = Format(crImporte, FMoneda)
        BringWindowToTop frmBoveda.hWnd
    
    End If

End Sub

Private Sub mnuCajaGralDivisas_Click()
    frmBovedaDivisas.Show
    BringWindowToTop frmBovedaDivisas.hWnd
End Sub

Private Sub mnuCalculadora_Click()
    Shell ("C:\windows\System32\calc.exe")
End Sub

Private Sub mnuCambioPlan_Click()
    frmCambioPlan.Show
    BringWindowToTop frmCambioPlan.hWnd
End Sub

Private Sub mnuCancelaciones_Click()
    frmPasswords.ConexSuc = 0
    frmPasswords.DescuentoVentas = 0
    frmPasswords.PrecioVitrina = 0
    frmPasswords.Ventas = 0
    frmPasswords.ModificaPrecio = 0
    frmPasswords.ModificaCorte = 0
    frmPasswords.InteresDesempeño = 0
    frmPasswords.InteresRefrendo = 0
    frmPasswords.HacerCorte = 0
    frmPasswords.RecalculoPrecios = 0
    frmPasswords.AutorizaPrestamo = 0
    frmPasswords.Vencido = 0
    frmPasswords.CancelaCierre = 0
    frmPasswords.Cancel = 1

    If frmPasswords.Password(CANCELACION, 1) Then
                
        frmCancelaciones.Show
        BringWindowToTop frmCancelaciones.hWnd
    End If
    
End Sub

Private Sub mnuCapturaBoletas_Click()
    frmCapboletas.Show
    BringWindowToTop frmCapboletas.hWnd
End Sub

Private Sub mnuCatDivisas_Click()
    frmCatdivisas.Show
    BringWindowToTop frmCatdivisas.hWnd
End Sub

Private Sub mnuCatMarcas_Click()
    frmCatMarcas.Show
    BringWindowToTop frmCatMarcas.hWnd
End Sub

Private Sub mnuCatPrendasVarios_Click()
    frmCatTipoPrendaOtros.Show
    BringWindowToTop frmCatTipoPrendaOtros.hWnd
End Sub

Private Sub mnuCatTiposPrenda_Click()
    frmCatTipoPrenda.Show
    BringWindowToTop frmCatTipoPrenda.hWnd
End Sub

Private Sub mnuCatVendedores_Click()
    frmCatVendedores.Show
    BringWindowToTop frmCatVendedores.hWnd
End Sub

Private Sub mnuCierreCaja_Click()
Dim rcCorte As New ADODB.Recordset
Dim crImporte As Double
    
    If MsgBox("Desea finalizar el dia y cerrar la caja ??", vbQuestion + vbYesNo + vbDefaultButton1, "Cierre de Caja") = vbNo Then
        
        frmCorteVentanilla.ucLine3D1(52).Visible = False
        frmCorteVentanilla.ucLine3D1(53).Visible = False
        frmCorteVentanilla.ucLine3D1(54).Visible = False
        frmCorteVentanilla.ucLine3D1(55).Visible = False
        frmCorteVentanilla.ucLine3D1(56).Visible = False
        frmCorteVentanilla.Label27(5).Visible = False
        frmCorteVentanilla.Label13(5).Visible = False
        frmCorteVentanilla.lblAjuste.Visible = False
        
        frmCorteVentanilla.cmdAceptar.Visible = False
        frmCorteVentanilla.txtEfectivo.text = Format(0, FMoneda)
        BringWindowToTop frmCorteVentanilla.hWnd
    Else
        
        rcCorte.Open "SELECT ID,Importe FROM auxiliar WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND Concepto='Corte de Caja' AND PC='" & NombrePc & "' AND (Iniciales='CV01' OR Iniciales='CV50')", dbDatos, adOpenForwardOnly, adLockOptimistic
        If rcCorte.BOF = True Or rcCorte.EOF = True Then
            
            frmPasswords.ConexSuc = 0
            frmPasswords.DescuentoVentas = 0
            frmPasswords.PrecioVitrina = 0
            frmPasswords.Cancel = 0
            frmPasswords.Ventas = 0
            frmPasswords.ModificaPrecio = 0
            frmPasswords.ModificaCorte = 0
            frmPasswords.InteresDesempeño = 0
            frmPasswords.InteresRefrendo = 0
            frmPasswords.RecalculoPrecios = 0
            frmPasswords.AutorizaPrestamo = 0
            frmPasswords.Vencido = 0
            frmPasswords.CancelaCierre = 0
            frmPasswords.HacerCorte = 1
                
            If frmPasswords.Password(GERENTE, 1) Then
                
                crImporte = 0
                crImporte = frmDenominaciones.Arqueo(True)
                If crImporte > 0 Then
                    
                    frmCorteVentanilla.txtEfectivo.text = Format(crImporte, FMoneda)
                    BringWindowToTop frmCorteVentanilla.hWnd
                
                End If
                
            End If
        
        Else
            
            frmCorteVentanilla.txtEfectivo.text = Format(rcCorte!Importe, FMoneda)
            frmCorteVentanilla.txtEfectivo.Tag = rcCorte!ID
            frmCorteVentanilla.cmdModificaCorte.Visible = False
            BringWindowToTop frmCorteVentanilla.hWnd
            
        End If
        rcCorte.Close
        Set rcCorte = Nothing
    
    End If
End Sub

Private Sub mnuCierreDivisas_Click()
    frmCorteDivisas.Show
    BringWindowToTop frmCorteDivisas.hWnd
End Sub

Private Sub mnuCierreRepEmpeños_Click()

    frmRangoFechas.Caption = "Reporte Empeños"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .ReportFileName = Path & "\Reportes\RepEmpenos.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.Origen}=" & OD_EMPENO & " AND {empeno.Fecha}>= date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {empeno.Fecha}<= date('" & Format(FechaFin, "YYYY,MM,DD") & "')"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Reporte de Empeños"
        .Action = 1
    End With

End Sub

Private Sub mnuCierresSucursal_Click()
    frmCierreMensual.Show
    BringWindowToTop frmCierreMensual.hWnd
End Sub

Private Sub mnuClienteFrecuenteParametros_Click()
    '***Puntos***
    Dim Puntos As New ClienteFrecuente
    Set Puntos.CONEXION = dbDatos
    Puntos.ShowParametros Me.hWnd
End Sub

Private Sub mnuClienteFrecuenteTiposTarjeta_Click()
    '***Puntos***
    Dim Puntos As New ClienteFrecuente
    Set Puntos.CONEXION = dbDatos
    Puntos.ShowTiposTarjetas Me.hWnd
End Sub

Private Sub mnuClientes_Click()
    frmClientes.Show
    BringWindowToTop frmClientes.hWnd
End Sub

Private Sub mnuCompraJoyeria_Click()
    frmCompras.Show
    BringWindowToTop frmCompras.hWnd
End Sub

Private Sub mnuCompras_Click()

    frmRangoFechas.Caption = "Reporte de compras"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepCompras.rpt"
        .SelectionFormula = "{Compras.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') and {Compras.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowTitle = "Reporte de compras"
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With

End Sub

Private Sub mnuCompraVenta_Click()
    frmDivisas.Show
    BringWindowToTop frmDivisas.hWnd
End Sub

Private Sub mnuConceptos_Click()
    frmConceptos.Show
    BringWindowToTop frmConceptos.hWnd
End Sub

Private Sub mnuConexInter_Click()
        
    frmPasswords.DescuentoVentas = 0
    frmPasswords.PrecioVitrina = 0
    frmPasswords.Ventas = 0
    frmPasswords.ModificaPrecio = 0
    frmPasswords.ModificaCorte = 0
    frmPasswords.InteresDesempeño = 0
    frmPasswords.InteresRefrendo = 0
    frmPasswords.HacerCorte = 0
    frmPasswords.RecalculoPrecios = 0
    frmPasswords.AutorizaPrestamo = 0
    frmPasswords.Cancel = 0
    frmPasswords.Vencido = 0
    frmPasswords.CancelaCierre = 0
    frmPasswords.ConexSuc = 1
    
    If frmPasswords.Password(CANCELACION, 1) Then
    
        frmConexionSucursal.Show
        BringWindowToTop frmConexionSucursal.hWnd
    End If
End Sub

Private Sub mnuConfImpresiones_Click()
    frmConfiguracionINI.Show
    BringWindowToTop frmConfiguracionINI.hWnd
End Sub

Private Sub mnuConfTasas_Click()
    frmConfiguracionTasas.Show
    BringWindowToTop frmConfiguracionTasas.hWnd
End Sub

Private Sub mnuCotizacionDivisas_Click()
    frmCotizacion.Show
    BringWindowToTop frmCotizacion.hWnd
End Sub

Private Sub mnuCotizaciones_Click()
    frmCotizar.Show
    BringWindowToTop frmCotizar.hWnd
End Sub

Private Sub mnuCrearPolizas_Click()
    frmCrearPoliza.Show
    BringWindowToTop frmCrearPoliza.hWnd
End Sub

Private Sub mnuCuentaGastos_Click()
    frmCuentas.Show
    BringWindowToTop frmCuentas.hWnd
End Sub

Private Sub mnuDesempeños_Click()
    frmEmpeño.Show
    frmEmpeño.TPestañas.SelectTab "K2"
    frmEmpeño.frmDesempeño.Visible = True
    frmEmpeño.frmEmpeño.Visible = False
    frmEmpeño.frmRefrendos.Visible = False
    
    '***Puntos***
    frmEmpeño.lblNoTarjeta.Visible = False
    frmEmpeño.txtNoTarjeta.Visible = False
    frmEmpeño.lblPuntosAcumulados1.Visible = False
    frmEmpeño.lblPuntosAcumulados.Visible = False
    
    BringWindowToTop frmEmpeño.hWnd
    frmEmpeño.txtFolioDesempeño.SetFocus
End Sub

Private Sub mnuDeslotificacion_Click()
    frmLotes.Show
    BringWindowToTop frmLotes.hWnd
End Sub

Private Sub mnuEmpeñoAutos_Click()
    frmEmpeño.Show
    frmEmpeño.TPestañas.SelectTab "K4"
    frmEmpeño.frmAutomoviles.Visible = True
    frmEmpeño.frmEmpeño.Visible = False
    frmEmpeño.frmRefrendos.Visible = False
    frmEmpeño.frmDesempeño.Visible = False
    frmEmpeño.lblAlmacenaje2.Caption = Regresa_Valor_BD("Almacenaje") & "%"
    frmEmpeño.lblSeguro2.Caption = Regresa_Valor_BD("Seguro") & "%"
    frmEmpeño.lblFecha(4).Caption = Format(Date, "DD/MMM/YY")
    BringWindowToTop frmEmpeño.hWnd
End Sub

Private Sub mnuDevolucionVta_Click()
    frmGarantias.Show
    BringWindowToTop frmGarantias.hWnd
End Sub

Private Sub mnuEmpeñoss_Click()
    frmEmpeño.Show
    frmEmpeño.TPestañas.SelectTab "K1"
    frmEmpeño.frmEmpeño.Visible = True
    frmEmpeño.frmRefrendos.Visible = False
    frmEmpeño.frmDesempeño.Visible = False
    BringWindowToTop frmEmpeño.hWnd
End Sub

Private Sub mnuEtiAlmoneda_Click()

    frmRangoFechas.Caption = "Etiquetas Inventario"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    frmEtiquetasAlmoneda.FechaIni = FechaIni
    frmEtiquetasAlmoneda.FechaFin = FechaFin
    frmEtiquetasAlmoneda.Show
    BringWindowToTop frmEtiquetasAlmoneda.hWnd
End Sub

Private Sub mnuEtiquetas_Click()
    frmEtiquetas.Show
    BringWindowToTop frmEtiquetas.hWnd
End Sub

Private Sub mnuExistencias_Click()
    frmExistencias.Show
    BringWindowToTop frmExistencias.hWnd
End Sub

Private Sub mnuFacturacion_Click()
    frmFacturacion.Show
    BringWindowToTop frmFacturacion.hWnd
End Sub

Private Sub mnuFamilias_Click()
    frmCatFamilias.Show
    BringWindowToTop frmCatFamilias.hWnd
End Sub

Private Sub mnuGeneraAutoriza_Click()
    frmGeneraAutorizaciones.Show
    BringWindowToTop frmGeneraAutorizaciones.hWnd
End Sub

Private Sub mnuImpresoras_Click()
    frmImpresoras.Show
    BringWindowToTop frmImpresoras.hWnd
End Sub

Private Sub mnuInvenFisico_Click()

    Screen.MousePointer = vbHourglass
    
    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\InventarioFisico.rpt"
        .SelectionFormula = "({detallesentradainventario.TipoEntrada}=" & ENTRADAALMONEDA & " OR {detallesentradainventario.TipoEntrada}=" & ENTRADACOMPRA & " OR {detallesentradainventario.TipoEntrada}=" & ENTRADADOTACION & ") AND {detallesentradainventario.Cantidad}>0"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado=''"
        .DiscardSavedData = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .WindowTitle = "Reporte de inventario físico"
        .Action = 1
    End With
    
    Screen.MousePointer = vbDefault

End Sub

Private Sub mnuInventarioFund_Click()
    
    frmRangoFechas.Caption = "Inventario Fundición"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    frmInventarioFund.CargarInventario CDate(FechaIni), CDate(FechaFin)
    BringWindowToTop frmInventarioFund.hWnd
End Sub

Private Sub mnuLLDExpClientes_Click()
    frmSelecRangoCliente.Show
    BringWindowToTop frmSelecRangoCliente.hWnd
End Sub

Private Sub mnuLLDMovAtipicos_Click()
    frmMovimientosAtipicos.Show
    BringWindowToTop frmMovimientosAtipicos.hWnd
End Sub

Private Sub mnuLLDParametros_Click()
    frmConfigLavadoDinero.Show
    BringWindowToTop frmConfigLavadoDinero.hWnd
End Sub

Private Sub mnuLLDRepMensual_Click()
    frmRepRegistro.Show
    BringWindowToTop frmRepRegistro.hWnd
End Sub

Private Sub mnuManualUsuario_Click()
    Shell "hh.exe " & Path & "\ayuda.chm", vbMaximizedFocus
End Sub

Private Sub mnuMedios_Click()
    frmCatmedios.Show
    BringWindowToTop frmCatmedios.hWnd
End Sub

Private Sub mnuMensajesBol_Click()
    frmMensajes.Show
    BringWindowToTop frmMensajes.hWnd
End Sub

Private Sub mnuMigracion_Click()
    frmMigracion.Show
    BringWindowToTop frmMigracion.hWnd
End Sub

Private Sub mnuMostrador_Click()
    frmVentas.Show
    frmVentas.tTab.SelectTab "K1"
    frmVentas.frmPagos.Visible = False
    frmVentas.frmApartados.Visible = False
    frmVentas.frmVentasMostrador.Visible = True
    BringWindowToTop frmVentas.hWnd
End Sub

Private Sub mnuMovDivisas_Click()
    frmMovidolares.Show
    BringWindowToTop frmMovidolares.hWnd
End Sub

Private Sub mnuMuestraApartados_Click()
    frmApartados.Show
    BringWindowToTop frmApartados.hWnd
End Sub

Private Sub mnuPagoDemasias_Click()
    frmDemasias.Show
    BringWindowToTop frmDemasias.hWnd
End Sub

Private Sub mnuPagoEmpeños_Click()
    frmVerificaEmpeño.Show
    BringWindowToTop frmVerificaEmpeño.hWnd
End Sub

Private Sub mnuPagosFijoss_Click()
    frmPagosFijos.Show
    BringWindowToTop frmPagosFijos.hWnd
End Sub

Private Sub mnuParametros_Click()
    frmConfiguracion.Show
    BringWindowToTop frmConfiguracion.hWnd
End Sub

Private Sub mnuPolizaDia_Click()
    frmPolizaDia.Show
    BringWindowToTop frmPolizaDia.hWnd
End Sub

Private Sub mnuPreciosDiamante_Click()
    frmPreciosDiamante.Show
    BringWindowToTop frmPreciosDiamante.hWnd
End Sub

Private Sub mnuPreciosOro_Click()
    'frmConfiguracionPrecio.Show
    'BringWindowToTop frmConfiguracionPrecio.hWnd
    frmPrecios.Show
    BringWindowToTop frmPrecios.hWnd
End Sub

Private Sub mnuPromociones_Click()
    frmCatPromociones.Show
    BringWindowToTop frmCatPromociones.hWnd
End Sub

Private Sub mnuReasignarT_Click()
    frmReasignarTarjeta.Show
    BringWindowToTop frmReasignarTarjeta.hWnd
End Sub

Private Sub mnuReEnvejecimiento_Click()
    frmRepenvejecimiento.Show
    BringWindowToTop frmRepenvejecimiento.hWnd
End Sub

Private Sub mnuRefrendos_Click()
    frmEmpeño.Show
    frmEmpeño.TPestañas.SelectTab "K3"
    frmEmpeño.frmRefrendos.Visible = True
    frmEmpeño.frmDesempeño.Visible = False
    frmEmpeño.frmEmpeño.Visible = False
    
    '***Puntos***
    frmEmpeño.lblNoTarjeta.Visible = False
    frmEmpeño.txtNoTarjeta.Visible = False
    frmEmpeño.lblPuntosAcumulados1.Visible = False
    frmEmpeño.lblPuntosAcumulados.Visible = False
    
    BringWindowToTop frmEmpeño.hWnd
    frmEmpeño.txtFolioRefrendo.SetFocus
End Sub

Private Sub mnuRefrendosForaneos_Click()
 frmRefrendosForaneos.Show
 BringWindowToTop frmRefrendosForaneos.hWnd
End Sub

Private Sub mnuRegGastos_Click()
    frmGastos.Show
    BringWindowToTop frmGastos.hWnd
End Sub

Private Sub mnuRegSoftware_Click()
    frmRegistrar.Salir = 1
    frmRegistrar.Show 1
End Sub

Private Sub mnuRegUbicacion_Click()
    frmUbicacion.Show
    BringWindowToTop frmUbicacion.hWnd
End Sub

Private Sub mnuRelCuentas_Click()
    frmRelacionarCuentas.Show
    BringWindowToTop frmRelacionarCuentas.hWnd
End Sub

Private Sub mnuRemates_Click()
Dim IDPrenda As Integer
    
    IDPrenda = frmTiposPrenda.Mostrar
    If IDPrenda = -2 Then Exit Sub
    frmRemates.MuestraPrendas IDPrenda
    BringWindowToTop frmRemates.hWnd
End Sub

Private Sub mnuRepAlmoneda_Click()

    frmRangoFechas.Caption = "Contratos a Almoneda"
    frmRangoFechas.Fechas FechaIni, FechaFin

    If Trim(FechaIni) = "" Or Trim(FechaFin) = "" Then Exit Sub

    With frmMDI.Cr
        
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\ContratosAlmoneda.rpt"
        .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY,MM,DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY,MM,DD") & "') AND {Empeno.Almoneda}=1"
        .Formulas(0) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(1) = "Encabezado='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Leyenda='De la fecha " & Format(CDate(FechaIni), "dd/mmm/yyyy") & " a " & Format(CDate(FechaFin), "dd/mmm/yyyy") & "'"
        
        .SubreportToChange = "Resumen"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY,MM,DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY,MM,DD") & "') AND {articulos.Kilates}>0 AND {articulos.Destino}=" & D_VENTA
                
        .SubreportToChange = "ResumenFundicion"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY,MM,DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY,MM,DD") & "') AND {articulos.Kilates}>0 AND {articulos.Destino}=" & D_FUNDICION

        .WindowTitle = "Contratos a Almoneda"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With

End Sub

Private Sub mnuRepAnaVentas_Click()

    frmRangoFechas.Caption = "Reporte Análisis de Ventas"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
   
    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepAnalisisVentas.rpt"
        .SelectionFormula = "{ventas.Fecha}>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {ventas.Fecha}<=date('" & Format(FechaFin, "YYYY,MM,DD") & "') AND {ventas.Cancelado}=0 AND {ventas.Apartado}=1"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowTitle = "Reporte Análisis de Ventas"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With

End Sub

Private Sub mnuRepAseguradora_Click()

    frmRangoFechas.Caption = "Reporte Aseguradora"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\Aseguradora.rpt"
        .SelectionFormula = "{empeno.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "')" & " AND {empeno.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')" & " AND {empeno.Cancelado}=0 AND {empeno.Pagado}=0 AND {empeno.Destino}=0"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='De " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Formulas(3) = "Usuario='" & frmMDI.Usuario & "'"
        .WindowTitle = "Reporte Aseguradora"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With

End Sub

Private Sub mnuRepAuditoria_Click()
    frmReportesMovimientos.Show
    BringWindowToTop frmReportesMovimientos.hWnd
End Sub

Private Sub mnuRepAutorizaciones_Click()

    frmRangoFechas.Caption = "Autorizaciones"
    frmRangoFechas.Fechas FechaIni, FechaFin
    If Trim(FechaIni) = "" Or Trim(FechaFin) = "" Then Exit Sub
    
    frmAutorizaciones.Ver CDate(FechaIni), CDate(FechaFin)
End Sub

Private Sub mnuRepAuxiliares_Click()
    frmRepAuxiliar.Show
    BringWindowToTop frmRepAuxiliar.hWnd
End Sub

Private Sub mnuRepCancelaciones_Click()
    
    frmRangoFechas.Caption = "Reporte de Cancelaciones"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepCancelaciones.rpt"
        .SelectionFormula = "{cancelaciones.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "')" & " And {cancelaciones.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='De " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowTitle = "Reporte de Cancelaciones"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
End Sub

Private Sub mnuRepCartera_Click()
    frmRepCartera.Show
    BringWindowToTop frmRepCartera.hWnd
End Sub

Private Sub mnuRepCieDesempeño_Click()

    frmRangoFechas.Caption = "Reporte Desempeños"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .ReportFileName = Path & "\Reportes\RepDesempenos.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.Destino}=" & D_DESEMPEÑO & " AND {empeno.FechaMovimiento}>= date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {empeno.FechaMovimiento}<= date('" & Format(FechaFin, "YYYY,MM,DD") & "')"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Reporte de Desempeños"
        .Action = 1
    End With

End Sub

Private Sub mnuRepCieRefrendos_Click()

    frmRangoFechas.Caption = "Reporte Refrendos"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .ReportFileName = Path & "\Reportes\RepRefrendos.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        '.SelectionFormula = "({empeno.Destino}=" & OD_REFRENDO & " OR {empeno.Destino}=5) AND {empeno.FechaMovimiento}>= date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {empeno.FechaMovimiento}<= date('" & Format(FechaFin, "YYYY,MM,DD") & "')"
        .SelectionFormula = "({empeno.Destino}=" & OD_REFRENDO & ") AND {empeno.FechaMovimiento}>= date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {empeno.FechaMovimiento}<= date('" & Format(FechaFin, "YYYY,MM,DD") & "')"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Reporte de Refrendos"
        .Action = 1
    End With

End Sub

Private Sub mnuRepClienteFrecuentaEdoCuenta_Click()
    '***Puntos***
    frmEstadoCuentaPuntos.Show
    BringWindowToTop frmEstadoCuentaPuntos.hWnd
End Sub

Private Sub mnuRepClienteFrecuenteMovimientos_Click()
    '***Puntos***
    frmRangoFechas.Caption = "Movimientos Puntos"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .ReportFileName = Path & "\Reportes\RepMovimientosPuntos.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{movimientospuntos.Fecha}>= date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {movimientospuntos.Fecha}<= date('" & Format(FechaFin, "YYYY,MM,DD") & "')"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Formulas(4) = "ValorPunto=" & Regresa_Valor_BD("PuntosTarjeta")
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Reporte de Movimiento Puntos"
        .Action = 1
    End With
End Sub

Private Sub mnuRepComiVen_Click()
Dim rcConsulta As New ADODB.Recordset

    frmRangoFechas.Caption = "Reporte de Comisiones"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    dbReportes.Execute "DELETE FROM repcomisiones"
    rcConsulta.Open "spRepComisiones ('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "','" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')", dbDatos, adOpenForwardOnly
    While Not rcConsulta.EOF
        
        dbReportes.Execute "INSERT INTO repcomisiones (Abonos,IDVenta,Folio,Fecha,IVA,Descuento,Vencimiento,Total,Pagado,Cliente,Vendedor) VALUES (" & _
                            rcConsulta!Abonos & "," & rcConsulta!IDVenta & "," & rcConsulta!Folio & ",'" & Format(rcConsulta!FechaAbono, "YYYY/MM/DD HH:MM:SS") & "'," & _
                            rcConsulta!Iva & "," & rcConsulta!Descuento & ",'" & Format(rcConsulta!Vencimiento, "YYYY/MM/DD") & "'," & rcConsulta!Total & "," & rcConsulta!Pagado & ",'" & rcConsulta!Cliente & "','" & rcConsulta!Vendedor & "')"
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    Set rcConsulta = Nothing
    
    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .ReportFileName = Path & "\Reportes\RepComisiones.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Formulas(3) = "FechaIni='" & FechaIni & "'"
        .Formulas(4) = "FechaFin='" & FechaFin & "'"
        
        .SubreportToChange = "Apartados"
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{repcomisiones.Fecha}>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {repcomisiones.Fecha}<= date('" & Format(FechaFin, "YYYY,MM,DD") & "')"
        
        .SubreportToChange = "Resumen"
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .Formulas(0) = "FechaIni='" & FechaIni & "'"
        .Formulas(1) = "FechaFin='" & FechaFin & "'"
        
        .WindowTitle = "Reporte de Comisiones"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    
End Sub

Private Sub mnuRepComVenDiv_Click()

On Error GoTo Error
    
    frmRangoFechas.Caption = "Compra/Venta Divisas"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\Divisas.rpt"
        .SelectionFormula = "{divisas.TipoEntrada}=0 AND {divisas.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY,MM,DD") & "') AND {divisas.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY,MM,DD") & "')"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Leyenda='De la fecha " & Format(CDate(FechaIni), "DD/MMM/YYYY") & " a " & Format(CDate(FechaFin), "DD/MMM/YYYY") & "'"
        .WindowTitle = "Reporte Compra/Venta Divisas"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub mnuRepContable_Click()
Dim rcCuentas As New ADODB.Recordset
Dim Bandera As Boolean

Screen.MousePointer = vbHourglass
    
  rcCuentas.Open "SELECT * FROM cuentas ORDER BY Cuenta", dbDatos, adOpenForwardOnly, adLockOptimistic
  
  Bandera = True
  With rcCuentas
      While Not .EOF
          Realizar_Diario Sucursal.NombreComercial, frmMDI.Usuario, !Mayor, !Cuenta, !Descripcion, Bandera
          Bandera = False
          .MoveNext
      Wend
  End With

  Sleep 1000

  'Imprimimos el reporte de diario
  With Cr
      .Reset
      .WindowShowPrintSetupBtn = True
      .WindowShowExportBtn = True
      .ReportFileName = Path & "\Reportes\Diario.rpt"
      .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
      .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
      .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
      .Formulas(2) = "Cajero='" & frmMDI.Usuario & "'"
      .Formulas(3) = "Subleyenda='De la fecha " & Format(Date, "dd/mmm/yyyy") & "'"
      .DiscardSavedData = True
      .WindowState = crptMaximized
      .Destination = crptToWindow
      .WindowTitle = "Reporte contable"
      .Action = 1
  End With

  rcCuentas.Close
  Set rcCuentas = Nothing
  Screen.MousePointer = vbDefault

End Sub

Private Sub mnuRepDesMes_Click()
    
On Error GoTo Error

    frmRangoFechas.Caption = "Reporte Desempeños por Mes"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
                
    SacaReporte FechaIni, FechaFin, 3
    Sleep 1000

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\DesempenosMes.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .WindowTitle = "Contratos Promedio Desempeños"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub mnuRepDetallado_Click()
    frmReportes.Show
    BringWindowToTop frmReportes.hWnd
End Sub

Private Sub mnuRepEmpeMes_Click()
Dim TipoPrenda As Long

On Error GoTo Error

    frmRangoFechas.Caption = "Reporte Empeños por Mes"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    TipoPrenda = 0
    TipoPrenda = frmTiposPrenda.Mostrar
    
    SacaReporte FechaIni, FechaFin, 1, TipoPrenda
    Sleep 1000

    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\EmpenosMes.rpt"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .WindowTitle = "Contratos Promedio Empeños"
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
    
End Sub

Private Sub mnuRepEmpenosTipoTasa_Click()

    frmRangoFechas.Caption = "Contratos por tipo interés"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\EmpenosTipoTasa.rpt"
        .SelectionFormula = "{empeno.Fecha} >= date('" & Format(CDate(FechaIni), "YYYY,MM,DD") & "') AND {empeno.Fecha}<= date('" & Format(CDate(FechaFin), "YYYY,MM,DD") & "') AND {empeno.Origen}=1 AND {empeno.Cancelado}=0"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Contratos por tipo interés"
        .Action = 1
    End With
    
End Sub

Private Sub mnuRepEmpenosVencidos_Click()

    frmRangoFechas.Caption = "Contratos vencidos"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\EmpenosVencidos.rpt"
        .SelectionFormula = "DATEADD('D'," & Val(Regresa_Valor_BD("DiasEnajenacion") + 1) & ",{empeno.Vencimiento})>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "')" & " AND DATEADD('D'," & Val(Regresa_Valor_BD("DiasEnajenacion") + 1) & ",{empeno.Vencimiento})<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')" & " AND {empeno.Cancelado}=0 AND {empeno.Pagado}=0 AND {empeno.Destino}=0"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Contratos vencidos"
        .Action = 1
    End With
    
End Sub

Private Sub mnuRepEmpeños_Click()
    frmRepEmpeños.Show
    BringWindowToTop frmRepEmpeños.hWnd
End Sub

Private Sub mnuRepEmpeñoss_Click()
    frmRepEmpeños.Show
    BringWindowToTop frmRepEmpeños.hWnd
End Sub

Private Sub mnuRepEmpeVencidos_Click()
'Dim IDTipoPrenda As Integer, FechaMov As String
'Dim rcConsulta As New ADODB.Recordset
'
'On Error GoTo Error
'
'    frmRangoFechas.Caption = "Reporte de contratos vencidos"
'    frmRangoFechas.Fechas FechaIni, FechaFin
'
'    IDTipoPrenda = frmTiposPrenda.Mostrar
'    If IDTipoPrenda = -2 Then Exit Sub
'    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
'
'    With rcConsulta
'
'        .Open "spRepVencidos('" & Format(FechaIni, "YYYY/MM/DD") & "','" & Format(FechaFin, "YYYY/MM/DD") & "'," & Val(Regresa_Valor_BD("DiasEnajenacion")) & "," & IIf(IDTipoPrenda = 0, 1, IIf(IDTipoPrenda = -1, 3, 2)) & "," & IDTipoPrenda & ")", dbDatos, adOpenForwardOnly, adLockReadOnly
'
'        dbReportes.Execute "DELETE FROM repvencidos"
'        While Not .EOF
'
'            If !Serie = SERIE_C Then
'
'                FechaMov = SacaValor("pagosfijos", "MAX(FechaMovimiento)", " WHERE IDEmpeno=" & !ID & " AND Cancelado=0 AND Pagado=1")
'                FechaMov = IIf(Trim(FechaMov) <> "", "'" & Format(FechaMov, "YYYY/MM/DD") & "'", "NULL")
'            Else
'
'                FechaMov = "NULL"
'            End If
'
'            dbReportes.Execute "INSERT INTO repvencidos (IDEmpeno,NumContrato,Fecha,Vencimiento,Cliente,Avaluo,Prestamo,Serie,TipoInteres,TipoTasa,FechaMovimiento,Tel) VALUES (" & _
'                                !ID & "," & !NumContrato & ",'" & Format(!Fecha, "YYYY/MM/DD HH:MM:SS") & "','" & Format(!Vencimiento, "YYYY/MM/DD") & "','" & !Cliente & "'," & !Avaluo & "," & !Prestamo & "," & !Serie & ",'" & !TipoInteres & "','" & !TipoTasa & "'," & FechaMov & ",'" & !Tel & "')"
'
'        .MoveNext
'        Wend
'        .Close
'        Set rcConsulta = Nothing
'
'    End With
'
'    With Cr
'        .Reset
'        .DiscardSavedData = True
'        .WindowShowPrintSetupBtn = True
'        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'        .ReportFileName = Path & "\Reportes\ContratosVencidos.rpt"
'        .Formulas(0) = "Encabezado='" & Sucursal.RazonSocial & "'"
'        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
'        .Formulas(2) = "Leyenda='De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
'        .WindowTitle = "Reporte de contratos vencidos"
'        .WindowState = crptMaximized
'        .Destination = crptToWindow
'        .Action = 1
'    End With
'    Exit Sub

Dim IDTipoPrenda As Integer, FechaMov As String
Dim rcConsulta As New ADODB.Recordset
Dim rcDetalle As New ADODB.Recordset
Dim diasEnajenacion As Integer
Dim FechaComercializacion As Date

On Error GoTo Error

    frmRangoFechas.Caption = "Reporte de contratos vencidos"
    frmRangoFechas.Fechas FechaIni, FechaFin
    
    IDTipoPrenda = frmTiposPrenda.Mostrar
    If IDTipoPrenda = -2 Then Exit Sub
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    With rcConsulta
    
        '.Open "CALL spRepVencidos('" & Format(FechaIni, "YYYY/MM/DD") & "','" & Format(FechaFin, "YYYY/MM/DD") & "'," & Val(Regresa_Valor_BD("DiasEnajenacion")) & "," & IIf(IDTipoPrenda = 0, 1, IIf(IDTipoPrenda = -1, 3, 2)) & "," & IDTipoPrenda & ")", dbDatos, adOpenForwardOnly, adLockReadOnly
        .Open "spRepVencidos('" & Format(FechaIni, "YYYY/MM/DD") & "','" & Format(FechaFin, "YYYY/MM/DD") & "'," & Val(0) & "," & IIf(IDTipoPrenda = 0, 1, IIf(IDTipoPrenda = -1, 3, 2)) & "," & IDTipoPrenda & ")", dbDatos, adOpenForwardOnly, adLockReadOnly
        dbReportes.Execute "DELETE FROM repvencidos"
        dbReportes.Execute "DELETE FROM repvencidosdetalle"
        diasEnajenacion = SacaValor("parametros", "diasEnajenacion")
        
        While Not .EOF
            
            If !Serie = SERIE_C Then
                FechaMov = SacaValor("pagosfijos", "MAX(FechaMovimiento)", " WHERE IDEmpeno=" & !ID & " AND Cancelado=0 AND Pagado=1")
                FechaMov = IIf(Trim(FechaMov) <> "", "'" & Format(FechaMov, "YYYY/MM/DD") & "'", "NULL")
            Else
                FechaMov = "NULL"
            End If
            
            
            FechaComercializacion = DateAdd("d", diasEnajenacion, Format(!Vencimiento, "YYYY/MM/DD"))
            dbReportes.Execute "INSERT INTO repvencidos (IDEmpeno,NumContrato,Fecha,Vencimiento,Cliente,Avaluo,Prestamo,Serie,TipoInteres,TipoTasa,FechaMovimiento,Tel,Celular,fechaComercializacion) VALUES (" & _
                !ID & "," & !NumContrato & ",'" & Format(!Fecha, "YYYY/MM/DD HH:MM:SS") & "','" & Format(!Vencimiento, "YYYY/MM/DD") & "','" & !Cliente & "'," & !Avaluo & "," & !Prestamo & "," & !Serie & ",'" & !TipoInteres & "','" & !TipoTasa & "'," & FechaMov & ",'" & !Tel & "','" & !Celular & "','" & Format(FechaComercializacion, "YYYY/MM/DD") & "')"
        

             rcDetalle.Open "select de.articulo,de.peso,k.descripcion,de.prestamo,de.marca,de.modelo from detallesempeno de left join kilatajes k on de.kilates=k.ID where IDEmpeno=" & !ID, dbDatos, adOpenForwardOnly, adLockOptimistic

             While Not rcDetalle.EOF And Not rcDetalle.BOF

             dbReportes.Execute "INSERT INTO repvencidosdetalle (IDEmpeno,articulo,peso,kilates,prestamo,marca,modelo) VALUES (" & !ID & ",'" & rcDetalle!Articulo & "','" & rcDetalle!Peso & "','" & rcDetalle!Descripcion & "'," & rcDetalle!Prestamo & ",'" & rcDetalle!Marca & "','" & rcDetalle!Modelo & "')"

             rcDetalle.MoveNext

             Wend
             
             rcDetalle.Close
        
        
            .MoveNext
        
        Wend
        
        .Close
        Set rcConsulta = Nothing
    
    End With
    
    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\ContratosVencidos.rpt"
        .Formulas(0) = "Encabezado='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Leyenda='De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowTitle = "Reporte de contratos vencidos"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub mnuRepEntradasInven_Click()

    frmRangoFechas.Caption = "Reporte Dotaciones"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepDotacionFechas.rpt"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .SelectionFormula = "{entradainventario.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND {entradainventario.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .WindowTitle = "Reporte Dotaciones"
        .Action = 1
    End With

End Sub

Private Sub mnuRepEnvejecimientoPeriodo_Click()

    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\EnvejecimientoPeriodo.rpt"
        .SelectionFormula = "{DetallesEntradaInventario.Cantidad} > 0 AND {DetallesEntradaInventario.TipoEntrada} <> " & D_FUNDICION & " AND {DetallesEntradaInventario.TipoEntrada} <> " & D_OTRO
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .WindowState = crptMaximized
        .WindowTitle = "Reporte de envejecimiento"
        .Action = 1
    End With
    
End Sub

Private Sub mnuRepExistenciasDiv_Click()
    
    Screen.MousePointer = vbHourglass
    
    'Saco las existencias de divisas
    ExistenciasDivisas
    
    Sleep 1500
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\ExistenciaDivisas.rpt"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .WindowTitle = "Reporte existencias divisas"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRepGastos_Click()

    frmRangoFechas.Caption = "Reporte de gastos"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\Gastos.rpt"
        .SelectionFormula = "{Gastos.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') and {Gastos.Fecha}<=date('" & Format(FechaFin, "YYYY/MM/DD") & "')"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Leyenda='De la fecha " & Format(CDate(FechaIni), "DD/MMM/YYYY") & " a " & Format(CDate(FechaFin), "DD/MMM/YYYY") & "'"
        .WindowTitle = "Reporte de gastos"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With

End Sub

Private Sub mnuRepHistorico_Click()
    frmHistorico.Show
    BringWindowToTop frmHistorico.hWnd
End Sub

Private Sub mnuRepInfo_Click()
Dim lnArch As Integer, psNomArch As String, sDireccion As String

On Error GoTo Error
    
    Dlgopen.CancelError = True
    Dlgopen.Filter = "*.sql"
    Dlgopen.ShowSave
    
    sDireccion = Dlgopen.FileName

    psNomArch = App.Path & "\Respaldo.bat"
    lnArch = FreeFile
    Open psNomArch For Output As #lnArch
    
        Print #lnArch, "@echo off"
        Print #lnArch, "title Respaldo de informacion"
        Print #lnArch, "Set path_mysqldump=""" & App.Path & """"
        Print #lnArch, "Set path_backups=""" & sDireccion & """"
        Print #lnArch, "Set User=""" & USERBD & """"
        Print #lnArch, "Set Password=""" & PWDBD & """"
        Print #lnArch, "Set host=" & Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost"))
        
        Print #lnArch, "echo Respaldando Informacion..."
        Print #lnArch, "%path_mysqldump%\mysqldump --user=%user% --password=%password% -h %host% --port=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & " --databases basedatos --single-transaction > %path_backups%_%date:~6,4%%date:~3,2%%date:~0,2%_%time:~0,2%%time:~3,2%.sql"
        Print #lnArch, "erase /q %0"
    Close lnArch
    
    Shell App.Path & "\Respaldo.bat", vbNormalFocus
        
    Exit Sub

Error:
    If Err = 32755 Then Exit Sub
    Maneja_Error Err
End Sub

Private Sub mnuRepIngresos_Click()

    frmRangoFechas.Caption = "Reporte de ingresos"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    
    RepIngresos CDate(FechaIni), CDate(FechaFin)
    
    Sleep 1000
    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\RepIngresos.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowTitle = "Reporte de Ingresos"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuRepInventario_Click()
    frmRepInventarios.Show
    BringWindowToTop frmRepInventarios.hWnd
End Sub

Private Sub mnuRepMedios_Click()

    frmRangoFechas.Caption = "Reporte de medios"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepMedios.rpt"
        .SelectionFormula = "{empeno.Fecha}>= date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {empeno.Fecha}<= date('" & Format(FechaFin, "YYYY,MM,DD") & "')"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "SubTitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowTitle = "Reporte de medios"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    
    End With
    
End Sub

Private Sub mnuRepOperaciones_Click()
Dim i As Integer, x As Integer, RangoFechas As String, FechaIni As String, FechaFin As String
Dim rcTmp As New ADODB.Recordset

    frmRangoFechas.Caption = "Operaciones por Horario"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    Screen.MousePointer = vbHourglass
    RangoFechas = " AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')>='" & Format(CDate(FechaIni), "YYYY/MM/DD") & "' AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')<='" & Format(CDate(FechaFin), "YYYY/MM/DD") & "'"
    
    dbReportes.Execute "DELETE FROM Horarios"
    dbReportes.Execute "INSERT INTO Horarios (Hora,Clave) VALUES ('00:00 A 08:00',1)"
    For i = 8 To 23
        
        dbReportes.Execute "INSERT INTO Horarios (Hora,Clave) VALUES ('" & Format(i, "00") & ":00" & " A " & Format(i + 1, "00") & ":00" & "'," & i - 6 & ")"
    Next i

    'Empeños
    rcTmp.Open "SELECT COUNT(ID) AS Empeños FROM empeno WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(0, "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(8 & ":00", "HH:MM:SS") & "' AND Origen=1" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Empeños) Then
        
        dbReportes.Execute "UPDATE horarios SET Empeos=" & rcTmp!Empeños & " WHERE Clave=1"
    End If
    rcTmp.Close
    
    x = 2
    For i = 8 To 23
        rcTmp.Open "SELECT COUNT(ID) AS Empeños FROM empeno WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(i & ":00", "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(IIf(i = 23, "23:59:59", i + 1 & ":00"), "HH:MM:SS") & "' AND Origen=1" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Empeños) Then
            
            dbReportes.Execute "UPDATE horarios SET Empeos=" & rcTmp!Empeños & " WHERE Clave=" & x
        End If
        rcTmp.Close
    x = x + 1
    Next i
    
    RangoFechas = " AND DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')>='" & Format(CDate(FechaIni), "YYYY/MM/DD") & "' AND DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')<='" & Format(CDate(FechaFin), "YYYY/MM/DD") & "'"
    'Refrendos
    rcTmp.Open "SELECT COUNT(ID) AS Refrendos FROM empeno WHERE DATE_FORMAT(FechaMovimiento,'%H%:%i%:%s')>'" & Format(0, "HH:MM:SS") & "' AND DATE_FORMAT(FechaMovimiento,'%H%:%i%:%s')<='" & Format(8 & ":00", "HH:MM:SS") & "' AND Destino=2" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!refrendos) Then
        
        dbReportes.Execute "UPDATE horarios SET Refrendos=" & rcTmp!refrendos & " WHERE Clave=1"
    End If
    rcTmp.Close
    
    x = 2
    For i = 8 To 23
        rcTmp.Open "SELECT COUNT(ID) AS Refrendos FROM empeno WHERE DATE_FORMAT(FechaMovimiento,'%H%:%i%:%s')>'" & Format(i & ":00", "HH:MM:SS") & "' AND DATE_FORMAT(FechaMovimiento,'%H%:%i%:%s')<='" & Format(IIf(i = 23, "23:59:59", i + 1 & ":00"), "HH:MM:SS") & "' AND Destino=2" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!refrendos) Then
            
            dbReportes.Execute "UPDATE horarios SET Refrendos=" & rcTmp!refrendos & " WHERE Clave=" & x
        End If
        rcTmp.Close
    x = x + 1
    Next i

    'Desempeños
    rcTmp.Open "SELECT COUNT(ID) AS Desempeños FROM empeno WHERE DATE_FORMAT(FechaMovimiento,'%H%:%i%:%s')>'" & Format(0, "HH:MM:SS") & "' AND DATE_FORMAT(FechaMovimiento,'%H%:%i%:%s')<='" & Format(8 & ":00", "HH:MM:SS") & "' AND Destino=3" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!desempeños) Then
        
        dbReportes.Execute "UPDATE horarios SET Desempeos=" & rcTmp!desempeños & " WHERE Clave=1"
    End If
    rcTmp.Close
    
    x = 2
    For i = 8 To 23
        rcTmp.Open "SELECT COUNT(ID) AS Desempeños FROM empeno WHERE DATE_FORMAT(FechaMovimiento,'%H%:%i%:%s')>'" & Format(i & ":00", "HH:MM:SS") & "' AND DATE_FORMAT(FechaMovimiento,'%H%:%i%:%s')<='" & Format(IIf(i = 23, "23:59:59", i + 1 & ":00"), "HH:MM:SS") & "' AND Destino=3" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!desempeños) Then
            
            dbReportes.Execute "UPDATE horarios SET Desempeos=" & rcTmp!desempeños & " WHERE Clave=" & x
        End If
        rcTmp.Close
    x = x + 1
    Next i
    
    RangoFechas = " AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')>='" & Format(CDate(FechaIni), "YYYY/MM/DD") & "' AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')<='" & Format(CDate(FechaFin), "YYYY/MM/DD") & "'"
    'Ventas
    rcTmp.Open "SELECT COUNT(ID) AS Ventas FROM ventas WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(0, "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(8 & ":00", "HH:MM:SS") & "' AND Apartado=0 AND Cancelado=0" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Ventas) Then
        
        dbReportes.Execute "UPDATE horarios SET Ventas=" & rcTmp!Ventas & " WHERE Clave=1"
    End If
    rcTmp.Close
    
    x = 2
    For i = 8 To 23
        rcTmp.Open "SELECT COUNT(ID) AS Ventas FROM ventas WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(i & ":00", "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(IIf(i = 23, "23:59:59", i + 1 & ":00"), "HH:MM:SS") & "' AND Apartado=0 AND Cancelado=0" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Ventas) Then
            
            dbReportes.Execute "UPDATE horarios SET Ventas=" & rcTmp!Ventas & " WHERE Clave=" & x
        End If
        rcTmp.Close
    x = x + 1
    Next i
    
    'Apartados
    rcTmp.Open "SELECT COUNT(ID) AS Apartados FROM ventas WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(0, "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(8 & ":00", "HH:MM:SS") & "' AND Apartado=1 AND Cancelado=0" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Apartados) Then
        
        dbReportes.Execute "UPDATE horarios SET Apartados=" & rcTmp!Apartados & " WHERE Clave=1"
    End If
    rcTmp.Close
    
    x = 2
    For i = 8 To 23
        rcTmp.Open "SELECT COUNT(ID) AS Apartados FROM ventas WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(i & ":00", "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(IIf(i = 23, "23:59:59", i + 1 & ":00"), "HH:MM:SS") & "' AND Apartado=1 AND Cancelado=0" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Apartados) Then
            
            dbReportes.Execute "UPDATE horarios SET Apartados=" & rcTmp!Apartados & " WHERE Clave=" & x
        End If
        rcTmp.Close
    x = x + 1
    Next i
    
    'Compra Divisas
    rcTmp.Open "SELECT COUNT(ID) AS ComDivisas FROM divisas WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(0, "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(8 & ":00", "HH:MM:SS") & "' AND Tipo=0 AND Cancelado=0 AND TipoEntrada=0" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!ComDivisas) Then
        
        dbReportes.Execute "UPDATE horarios SET ComDiv=" & rcTmp!ComDivisas & " WHERE Clave=1"
    End If
    rcTmp.Close
    
    x = 2
    For i = 8 To 23
        rcTmp.Open "SELECT COUNT(ID) AS ComDivisas FROM divisas WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(i & ":00", "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(IIf(i = 23, "23:59:59", i + 1 & ":00"), "HH:MM:SS") & "' AND Tipo=0 AND Cancelado=0 AND TipoEntrada=0" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!ComDivisas) Then
            
            dbReportes.Execute "UPDATE horarios SET ComDiv=" & rcTmp!ComDivisas & " WHERE Clave=" & x
        End If
        rcTmp.Close
    x = x + 1
    Next i
    
    'Venta Divisas
    rcTmp.Open "SELECT COUNT(ID) AS VenDivisas FROM divisas WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(0, "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(8 & ":00", "HH:MM:SS") & "' AND Tipo=1 AND Cancelado=0 AND TipoEntrada=0" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!VenDivisas) Then
        
        dbReportes.Execute "UPDATE horarios SET VenDiv=" & rcTmp!VenDivisas & " WHERE Clave=1"
    End If
    rcTmp.Close
    
    x = 2
    For i = 8 To 23
        rcTmp.Open "SELECT COUNT(ID) AS VenDivisas FROM divisas WHERE DATE_FORMAT(Fecha,'%H%:%i%:%s')>'" & Format(i & ":00", "HH:MM:SS") & "' AND DATE_FORMAT(Fecha,'%H%:%i%:%s')<='" & Format(IIf(i = 23, "23:59:59", i + 1 & ":00"), "HH:MM:SS") & "' AND Tipo=1 AND Cancelado=0 AND TipoEntrada=0" & RangoFechas, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!VenDivisas) Then
            
            dbReportes.Execute "UPDATE horarios SET VenDiv=" & rcTmp!VenDivisas & " WHERE Clave=" & x
        End If
        rcTmp.Close
    x = x + 1
    Next i
    
    Sleep 800
    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\OperacionesHorario.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "Del " & Format(FechaIni, "DD/MMM/YYYY") & " al " & Format(FechaFin, "DD/MMM/YYYY") & "'"
        .WindowTitle = "Operaciones por Horario"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    
    Set rcTmp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Private Sub mnuReporteAlmoneda_Click()

'''    frmRangoFechas.Caption = "Contratos a Almoneda"
'''    frmRangoFechas.Fechas FechaIni, FechaFin
'''
'''    If Trim(FechaIni) = "" Or Trim(FechaFin) = "" Then Exit Sub
'''
'''    With frmMDI.Cr
'''
'''        .Reset
'''        .DiscardSavedData = True
'''        .WindowShowPrintSetupBtn = True
'''        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'''        .ReportFileName = Path & "\Reportes\ContratosAlmoneda.rpt"
'''
'''.SQLQuery = " SELECT detallesempeno1.`ID`, detallesempeno1.`Codigo`, detallesempeno1.`Cantidad`, detallesempeno1.`Articulo`, detallesempeno1.`Peso`, detallesempeno1.`Avaluo`, detallesempeno1.`Prestamo`, detallesempeno1.`Destino`, detallesempeno1.`Observaciones`, " & _
'''    " empeno1.`ID`, empeno1.`Fecha`, empeno1.`NumContrato`, empeno1.`Folio`, empeno1.`Prestamo`, empeno1.`Avaluo`, empeno1.`FechaMovimiento`, empeno1.`Serie`, " & _
'''    " kilatajes1.`Descripcion`, clientes1.`Nombre`, clientes1.`Apellido`, " & _
'''   " detallesempenoautos1.`MarcayModelo`, detallesempenoautos1.`Año`, detallesempenoautos1.`Placas`, detallesempenoautos1.`Factura`, detallesempenoautos1.`Agencia`, detallesempenoautos1.`NumTarjetacircu`, detallesempenoautos1.`NumMotor` " & _
'''" From { oj (((`BaseDatos`.`detallesempeno` detallesempeno1 LEFT JOIN `BaseDatos`.`kilatajes` kilatajes1 ON detallesempeno1.`Kilates` = kilatajes1.`Clave`) INNER JOIN `BaseDatos`.`empeno` empeno1 ON detallesempeno1.`IDEmpeno` = empeno1.`ID`) LEFT JOIN `BaseDatos`.`detallesempenoautos` detallesempenoautos1 ON empeno1.`ID` = detallesempenoautos1.`IDEmpeno`) INNER JOIN `BaseDatos`.`clientes` clientes1 ON empeno1.`IDCliente` = clientes1.`ID`} " & _
'''" WHERE empeno1.FechaAlmoneda>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND empeno1.FechaAlmoneda<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "') AND Empeno1.Almoneda=1 " & _
'''" Order By  empeno1.`NumContrato` ASC "
''''        .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "') AND {Empeno.Almoneda}=1"
'''        .Formulas(0) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
'''        .Formulas(1) = "Encabezado='" & Sucursal.RazonSocial & "'"
'''        .Formulas(2) = "Leyenda='De la fecha " & Format(CDate(FechaIni), "dd/mmm/yyyy") & " a " & Format(CDate(FechaFin), "dd/mmm/yyyy") & "'"
'''
'''        .SubreportToChange = "Resumen"
'''        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
''''        .SQLQuery = "SELECT articulos1.`Cantidad`, articulos1.`Peso`, articulos1.`Kilates`, articulos1.`Prestamo`, " & _
''''   " kilatajes1.`Descripcion` " & _
''''" From " & _
''''   " `BaseReportes`.`articulos` articulos1 INNER JOIN `BaseDatos`.`empeno` empeno1 ON articulos1.`IDEmpeno` = empeno1.`ID` INNER JOIN `BaseDatos`.`kilatajes` kilatajes1 ON articulos1.`Kilates` = kilatajes1.`Clave` " & _
''''   " WHERE empeno1.FechaAlmoneda>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND empeno1.FechaAlmoneda<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "') AND articulos1.Kilates>0 AND articulos1.Destino=" & D_VENTA & _
''''" Order By " & _
''''   " articulos1.`Kilates` ASC"
'''        .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "') AND {articulos.Kilates}>0 AND {articulos.Destino}=" & D_VENTA
'''' .SelectionFormula = "{articulos.Kilates}>0 AND {articulos.Destino}=" & D_VENTA
'''
''''
'''        .SubreportToChange = "ResumenFundicion"
'''        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
''''       .SQLQuery = "SELECT articulos1.`Cantidad`, articulos1.`Peso`, articulos1.`Kilates`, articulos1.`Prestamo`, " &
''''    " kilatajes1.`Descripcion` " & _
''''" From `BaseReportes`.`articulos` articulos1 INNER JOIN `BaseDatos`.`kilatajes` kilatajes1 ON articulos1.`Kilates` = kilatajes1.`Clave` " & _
''''" Where empeno1.FechaAlmoneda>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND empeno1.FechaAlmoneda<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "') AND articulos1.Kilates>0 AND articulos1.Destino=" & D_FUNDICION & _
''''" Order By " & _
''''    " articulos1.`Kilates` ASC"
'''       .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "') AND {articulos.Kilates}>0 AND {articulos.Destino}=" & D_FUNDICION
''''.SelectionFormula = "{articulos.Kilates}>0 AND {articulos.Destino}=" & D_FUNDICION
'''
'''
'''        .WindowTitle = "Contratos a Almoneda"
'''        .Destination = crptToWindow
'''        .WindowState = crptMaximized
'''        .Action = 1
'''    End With
    frmRangoFechas.Caption = "Contratos a Almoneda"
    frmRangoFechas.Fechas FechaIni, FechaFin

    If Trim(FechaIni) = "" Or Trim(FechaFin) = "" Then Exit Sub

    With frmMDI.Cr
        
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\ContratosAlmoneda.rpt"
        .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY,MM,DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY,MM,DD") & "') AND {Empeno.Almoneda}=1"
        .Formulas(0) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(1) = "Encabezado='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Leyenda='De la fecha " & Format(CDate(FechaIni), "dd/mmm/yyyy") & " a " & Format(CDate(FechaFin), "dd/mmm/yyyy") & "'"
        
        .SubreportToChange = "Resumen"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY,MM,DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY,MM,DD") & "') AND {articulos.Kilates}>0 AND {articulos.Destino}=" & D_VENTA
                
        .SubreportToChange = "ResumenFundicion"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.FechaAlmoneda}>=date('" & Format(CDate(FechaIni), "YYYY,MM,DD") & "') AND {empeno.FechaAlmoneda}<=date('" & Format(CDate(FechaFin), "YYYY,MM,DD") & "') AND {articulos.Kilates}>0 AND {articulos.Destino}=" & D_FUNDICION

        .WindowTitle = "Contratos a Almoneda"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .Action = 1
    End With
End Sub

Private Sub mnuRepPagosFijos_Click()
    
    frmRangoFechas.Caption = "Reporte Pagos Fijos"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .ReportFileName = Path & "\Reportes\RepPagosFijos.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{pagosfijos.FechaMovimiento}>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {pagosfijos.FechaMovimiento}<=date('" & Format(FechaFin, "YYYY,MM,DD") & "')"
        .Formulas(1) = "Encabezado='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "SubEncabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Reporte Pagos Fijos"
        .Action = 1
    End With

End Sub

Private Sub mnuRepPartidasBoveda_Click()

    frmRangoFechas.Caption = "Reporte de Partidas en Bóveda"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    With Cr
        .Reset
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\PartidasBoveda.rpt"
        .SelectionFormula = "{Empeno.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "')" & " And {Empeno.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')" & " and {Empeno.Cancelado}=0 and {Empeno.Pagado}=0 and {Empeno.Destino}=0 And {Empeno.Caja}<>''"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='De " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowShowPrintSetupBtn = True
        .WindowTitle = "Reporte de Partidas en Bóveda"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
End Sub

Private Sub mnuRepPrestamosMes_Click()

    frmRangoFechas.Caption = "Contratos por mes"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\PrestamoPorMes.rpt"
        .SelectionFormula = "{empeno.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND {empeno.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Contratos por mes"
        .Action = 1
    End With
    
End Sub

Private Sub mnuRepPrestamoStatus_Click()

    frmRangoFechas.Caption = "Contratos por status"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\ContratosStatus.rpt"
        .SelectionFormula = "{empeno.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND {empeno.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "') AND {empeno.Cancelado}=0"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Encabezado='" & "De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Contratos por status"
        .Action = 1
    End With
    
End Sub

Private Sub mnuRepRefMes_Click()
    
On Error GoTo Error

    frmRangoFechas.Caption = "Reporte Refrendos por Mes"
    frmRangoFechas.Fechas FechaIni, FechaFin
   
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

    SacaReporte FechaIni, FechaFin, 2
    Sleep 1000

    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RefrendosMes.rpt"
        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .WindowTitle = "Contratos Promedio Refrendos"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
    
End Sub

Private Sub mnuRepSalInventario_Click()

'    frmRangoFechas.Caption = "Salidas de inventario"
'    frmRangoFechas.Fechas FechaIni, FechaFin
'
'    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub

'    With Cr
'        .Reset
'        .DiscardSavedData = True
'        .WindowShowPrintSetupBtn = True
'        .WindowShowExportBtn = True
'        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'        .ReportFileName = Path & "\Reportes\RepSalidaInventarioFechas.rpt"
'        .SelectionFormula = "{salidainventario.TipoSalida}=0 AND {salidainventario.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND {salidainventario.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')"
'        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
'        .Formulas(2) = "Subtitulo='SUCURSAL:" & Sucursal.NombreComercial & "'"
'        .Formulas(3) = "Leyenda='" & "De la fecha " & Format(FechaIni, "DD/MMM/YYYY") & " a " & Format(FechaFin, "DD/MMM/YYYY") & "'"
'        .Destination = crptToWindow
'        .WindowState = crptMaximized
'        .WindowTitle = "Salida de inventario"
'        .Action = 1
'    End With
    frmRangoFechas.Caption = "Salidas de inventario"
    frmRangoFechas.Fechas FechaIni, FechaFin

    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepSalidaInventarioFechas.rpt"
        .SelectionFormula = "{salidainventario.TipoSalida}=0 AND {salidainventario.Fecha}>=date('" & Format(CDate(FechaIni), "YYYY/MM/DD") & "') AND {salidainventario.Fecha}<=date('" & Format(CDate(FechaFin), "YYYY/MM/DD") & "')"
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL:" & Sucursal.NombreComercial & "'"
        .Formulas(3) = "Leyenda='" & "De la fecha " & Format(FechaIni, "DD/MMM/YYYY") & " a " & Format(FechaFin, "DD/MMM/YYYY") & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Salida de inventario"
        .Action = 1
    End With
    
End Sub

Private Sub mnuRepTraspasos_Click()

    frmRangoFechas.Caption = "Reporte de traspasos"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
       
    With Cr
       .Reset
       .WindowShowPrintSetupBtn = True
       .WindowShowExportBtn = True
       .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
       .ReportFileName = Path & "\Reportes\TraspasosFechas.rpt"
       .SelectionFormula = "{Traspasos.Fecha} >= date(" & Format(CDate(FechaIni), "YYYY,MM,DD") & ") AND {Traspasos.Fecha} <= date(" & Format(CDate(FechaFin), "YYYY,MM,DD") & ")"
       .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
       .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
       .Formulas(2) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
       .WindowTitle = "Reporte de traspasos"
       .DiscardSavedData = True
       .WindowState = crptMaximized
       .Action = 1
    End With

End Sub

Private Sub mnuRepUtilidadVentas_Click()

    frmRangoFechas.Caption = "Reporte de utilidad"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
       
    With Cr
       .Reset
       .WindowShowPrintSetupBtn = True
       .WindowShowExportBtn = True
       .DiscardSavedData = True
       .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
       .ReportFileName = Path & "\Reportes\RepUtilidadVentas.rpt"
       .SelectionFormula = "{ventas.Fecha}>=date(" & Format(CDate(FechaIni), "YYYY,MM,DD") & ") AND {ventas.Fecha}<=date(" & Format(CDate(FechaFin), "YYYY,MM,DD") & ") AND {ventas.Cancelado}=0 AND {ventas.Apartado}=0"
       .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
       .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
       .Formulas(2) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
       .WindowTitle = "Reporte Utilidad Ventas"
       .WindowState = crptMaximized
       .Action = 1
    End With

End Sub

Private Sub mnuRepVencidos_Click()

Dim IDTipoPrenda As Integer, FechaMov As String
Dim rcConsulta As New ADODB.Recordset
Dim rcDetalle As New ADODB.Recordset
Dim diasEnajenacion As Integer, LaSerie As Integer
Dim FechaComercializacion As Date

On Error GoTo Error

    frmRangoFechas.Caption = "Reporte de contratos vencidos"
    frmRangoFechas.Fechas FechaIni, FechaFin
    
    IDTipoPrenda = frmTiposPrenda.Mostrar
    If IDTipoPrenda = -2 Then Exit Sub
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
    
    With rcConsulta
    
        '.Open "CALL spRepVencidos('" & Format(FechaIni, "YYYY/MM/DD") & "','" & Format(FechaFin, "YYYY/MM/DD") & "'," & Val(Regresa_Valor_BD("DiasEnajenacion")) & "," & IIf(IDTipoPrenda = 0, 1, IIf(IDTipoPrenda = -1, 3, 2)) & "," & IDTipoPrenda & ")", dbDatos, adOpenForwardOnly, adLockReadOnly
        .Open "spRepVencidos('" & Format(FechaIni, "YYYY/MM/DD") & "','" & Format(FechaFin, "YYYY/MM/DD") & "',0," & IIf(IDTipoPrenda = 0, 1, IIf(IDTipoPrenda = -1, 3, 2)) & "," & IDTipoPrenda & ")", dbDatos, adOpenForwardOnly, adLockReadOnly
        dbReportes.Execute "DELETE FROM repvencidos"
        dbReportes.Execute "DELETE FROM repvencidosdetalle"
        diasEnajenacion = SacaValor("parametros", "diasEnajenacion")
        
        While Not .EOF
            
            If !Serie = SERIE_C Then
                FechaMov = SacaValor("pagosfijos", "MAX(FechaMovimiento)", " WHERE IDEmpeno=" & !ID & " AND Cancelado=0 AND Pagado=1")
                FechaMov = IIf(Trim(FechaMov) <> "", "'" & Format(FechaMov, "YYYY/MM/DD") & "'", "NULL")
            Else
                FechaMov = "NULL"
            End If
            
            
            FechaComercializacion = DateAdd("d", diasEnajenacion, Format(!Vencimiento, "YYYY/MM/DD"))
            dbReportes.Execute "INSERT INTO repvencidos (IDEmpeno,NumContrato,Fecha,Vencimiento,Cliente,Avaluo,Prestamo,Serie,TipoInteres,TipoTasa,FechaMovimiento,Tel,Celular,fechaComercializacion) VALUES (" & _
                !ID & "," & !NumContrato & ",'" & Format(!Fecha, "YYYY/MM/DD HH:MM:SS") & "','" & Format(!Vencimiento, "YYYY/MM/DD") & "','" & !Cliente & "'," & !Avaluo & "," & !Prestamo & "," & !Serie & ",'" & !TipoInteres & "','" & !TipoTasa & "'," & FechaMov & ",'" & !Tel & "','" & !Celular & "','" & Format(FechaComercializacion, "YYYY/MM/DD") & "')"
        
             LaSerie = !Serie
             If LaSerie = 2 Then
             
              'strDescripcion = "MARCA Y MODELO: " & !MarcayModelo & ", PLACAS: " & !Placas & ", AÑO: " & !Año & ", COLOR: " & !Color & ", SERIE CHASIS: " & !SerieChasis & ", NUM. MOTOR: " & !NumMotor & ", TARJETA CIRC.: " & !NumTarjetaCircu

                 rcDetalle.Open "select '0' as descripcion,'0' as peso,MarcayModelo,Placas , Año ,Color ,SerieChasis  ,NumMotor , NumTarjetaCircu, '' as articulo,e.prestamo,de.marca,de.modelo from empeno as e inner join detallesempenoautos de on de.IDEmpeno = e.ID where de.IDEmpeno=" & !ID, dbDatos, adOpenForwardOnly, adLockOptimistic

             Else
                 rcDetalle.Open "select de.articulo,de.peso,k.descripcion,de.prestamo,de.marca,de.modelo from detallesempeno de left join kilatajes k on de.kilates=k.ID where IDEmpeno=" & !ID, dbDatos, adOpenForwardOnly, adLockOptimistic

             End If
            
             While Not rcDetalle.EOF And Not rcDetalle.BOF

                Dim descripciones As String
                If LaSerie = 2 Then
                    descripciones = "MARCA Y MODELO " & rcDetalle!MarcayModelo & ", PLACAS " & rcDetalle!Placas & ", AÑO " & rcDetalle!Año & ", COLOR " & rcDetalle!Color & ", SERIE CHASIS " & rcDetalle!SerieChasis & ", NUM. MOTOR " & rcDetalle!NumMotor & ", TARJETA CIRC" & rcDetalle!NumTarjetaCircu & " "
                Else
                    descripciones = rcDetalle!Articulo
                End If
             
                 dbReportes.Execute "INSERT INTO repvencidosdetalle (IDEmpeno,articulo,peso,kilates,prestamo,marca,modelo) VALUES (" & !ID & ",'" & Trim(descripciones) & "','" & rcDetalle!Peso & "','" & rcDetalle!Descripcion & "'," & rcDetalle!Prestamo & ",'" & rcDetalle!Marca & "','" & rcDetalle!Modelo & "')"
    
             rcDetalle.MoveNext

             Wend
             
             rcDetalle.Close
        
        
            .MoveNext
        
        Wend
        
        .Close
        Set rcConsulta = Nothing
    
    End With
    
    With Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\ContratosVencidos.rpt"
        .Formulas(0) = "Encabezado='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Leyenda='De la fecha " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .WindowTitle = "Reporte de contratos vencidos"
        .WindowState = crptMaximized
        .Destination = crptToWindow
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub mnuRepVentaCliente_Click()
    
    frmRangoFechas.Caption = "Reporte de ventas"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
       
    With Cr
       .Reset
       .WindowShowPrintSetupBtn = True
       .WindowShowExportBtn = True
       .DiscardSavedData = True
       .ReportFileName = Path & "\Reportes\RepVentasBillete.rpt"
       .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
       .SelectionFormula = "{ventas.TipoVenta}=1 AND {ventas.Fecha}>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {ventas.Fecha}<=date('" & Format(FechaFin, "YYYY,MM,DD" & "'") & ") AND {ventas.Cancelado}=0 AND {ventas.Apartado}=0"
       .Formulas(0) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
       .WindowTitle = "Reporte ventas billete"
       .WindowState = crptMaximized
       .Destination = crptToWindow
       .Action = 1
    End With
    
End Sub

Private Sub mnuRepVentasApa_Click()

    frmRangoFechas.Caption = "Reporte de ventas apartado"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
       
    With Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepApartados.rpt"
'        .SelectionFormula = "date({vwrepapartados.Fecha})>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND date({vwrepapartados.Fecha})<=date('" & Format(FechaFin, "YYYY,MM,DD" & "'") & ") AND {vwrepapartados.Cancelado}=0 AND {vwrepapartados.Pagado}=0"
         .SelectionFormula = "date({vwrepapartados.Fecha})>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND date({vwrepapartados.Fecha})<=date('" & Format(FechaFin, "YYYY,MM,DD" & "'") & ")  AND {vwrepapartados.Pagado}=0"

        .Formulas(0) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(1) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Formulas(2) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
        .Formulas(3) = "FechaIni='" & FechaIni & "'"
        .Formulas(4) = "FechaFin='" & FechaFin & "'"
        .WindowTitle = "Reporte ventas de apartado"
        .WindowState = crptMaximized
        .Action = 1
    End With
End Sub

Private Sub mnuRepVentasCon_Click()

    frmRangoFechas.Caption = "Reporte de ventas"
    frmRangoFechas.Fechas FechaIni, FechaFin
       
    If (FechaIni = "" And FechaFin = "") Or (FechaIni = "" Or FechaFin = "") Then Exit Sub
       
    With Cr
'       .Reset
'       .WindowShowPrintSetupBtn = True
'       .WindowShowExportBtn = True
'       .DiscardSavedData = True
'       .ReportFileName = Path & "\Reportes\RepVentas.rpt"
'       .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'       '''.SelectionFormula = "{ventas.TipoVenta}=0 AND {ventas.Fecha}>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {ventas.Fecha}<=date('" & Format(FechaFin, "YYYY,MM,DD" & "'") & ") AND {ventas.Cancelado}=0 AND {ventas.Apartado}=0"
'       .Formulas(0) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
'       .WindowTitle = "Reporte ventas mostrador"
'       .WindowState = crptMaximized
'       .Destination = crptToWindow
'       .Action = 1
               .Reset
       .WindowShowPrintSetupBtn = True
       .WindowShowExportBtn = True
       .DiscardSavedData = True
       .ReportFileName = Path & "\Reportes\RepVentas.rpt"
       .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'''''       .SelectionFormula = "{ventas.TipoVenta}=0 AND {ventas.Fecha}>=date('" & Format(FechaIni, "YYYY,MM,DD") & "') AND {ventas.Fecha}<=date('" & Format(FechaFin, "YYYY,MM,DD" & "'") & ") AND {ventas.Cancelado}=0"
       .Formulas(0) = "Encabezado='" & "Del " & Format(FechaIni, "dd/mmm/yyyy") & " a " & Format(FechaFin, "dd/mmm/yyyy") & "'"
       .Formulas(1) = "FechaIni='" & Format(FechaIni, "YYYY/MM/DD") & "'"
       .Formulas(2) = "FechaFin='" & Format(FechaFin, "YYYY/MM/DD") & "'"
       .WindowTitle = "Reporte ventas mostrador"
       .WindowState = crptMaximized
       .Destination = crptToWindow
       .Action = 1

    End With
End Sub

Private Sub mnuSalidasInventario_Click()
    frmSalidaInventario.Show
    BringWindowToTop frmSalidaInventario.hWnd
End Sub

Private Sub mnuSalir_Click()
    Unload Me
End Sub

Private Sub mnuSistema_Click()
    ShellAbout Me.hWnd, "MySonda Versión: " & App.Major & "." & App.Minor & "." & App.Revision, "Programado por: Juan A. Gómez Vázquez" & vbCrLf & "Lider de Proyecto: Ing. Ricardo Suárez", Me.Icon
End Sub

Private Sub mnuSucursales_Click()
    frmCatsucursales.Show
    BringWindowToTop frmCatsucursales.hWnd
End Sub

Private Sub mnuTipos_Click()
    frmCattipos.Show
    BringWindowToTop frmCattipos.hWnd
End Sub

Private Sub mnuTrasInventario_Click()
    frmTraspasos.Show
    BringWindowToTop frmTraspasos.hWnd
End Sub

Private Sub mnuTraspasos_Click()
    frmTraspasos.Show
    BringWindowToTop frmTraspasos.hWnd
End Sub

Private Sub mnuUsuarios_Click()
    frmAltaUsuarios.Show
    BringWindowToTop frmAltaUsuarios.hWnd
End Sub

Private Sub mnuVentasClientes_Click()
    frmVentaCliente.Show
    BringWindowToTop frmVentaCliente.hWnd
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComCtlLib.Button)

On Error GoTo Error

    Select Case Button.Key
    
    Case "Salir"
        
        Unload Me
          
    Case "Buscar"
        
        frmBusqueda.Show
        BringWindowToTop frmBusqueda.hWnd
        
    Case "Empeño"
    
        frmEmpeño.Show
        frmEmpeño.TPestañas.SelectTab "K1"
        frmEmpeño.frmEmpeño.Visible = True
        frmEmpeño.frmRefrendos.Visible = False
        frmEmpeño.frmDesempeño.Visible = False
        BringWindowToTop frmEmpeño.hWnd
    
    Case "Cierre"
    
        frmCierreDiario.Show
        BringWindowToTop frmCierreDiario.hWnd
        
    Case "Venta"
    
        frmVentas.Show
        frmVentas.tTab.SelectTab "K1"
        frmVentas.frmPagos.Visible = False
        frmVentas.frmApartados.Visible = False
        frmVentas.frmVentasMostrador.Visible = True
        BringWindowToTop frmVentas.hWnd
                
    Case "divisas"
        
        frmDivisas.Show
        BringWindowToTop frmDivisas.hWnd
        
    End Select
    Exit Sub
    
Error:
    Maneja_Error Err
    
End Sub

Private Sub MDIForm_Load()

On Error GoTo Error
    
    Inicializar
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

'Inicializamos el moulo
Private Sub Inicializar()
        
    TiempoActualizacion = 0

    Set Toolbar = frmMDI.CommandBars.Add("Standar", xtpBarLeft)
    Set stBar = frmMDI.CommandBars.StatusBar
    
    Set PaneHora = stBar.AddPane(1)
    Set PaneSucursal = stBar.AddPane(2)
    
    CreateToolBars
    CreateStatusBar
    tmrHora_Timer
    PaseAlmoneda = False
End Sub

Private Sub CommandBars_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    
    Select Case Control.ID
    Case eToolBar.Buscar
        
        frmBusqueda.Show
        BringWindowToTop frmBusqueda.hWnd
    
    Case eToolBar.Empeno
        
        frmEmpeño.Show
        frmEmpeño.TPestañas.SelectTab "K1"
        frmEmpeño.frmEmpeño.Visible = True
        frmEmpeño.frmRefrendos.Visible = False
        frmEmpeño.frmDesempeño.Visible = False
        BringWindowToTop frmEmpeño.hWnd
    
    Case eToolBar.Cierre
        
        frmCierreDiario.Show
        BringWindowToTop frmCierreDiario.hWnd
    
    Case eToolBar.Venta
        
        frmVentas.Show
        frmVentas.tTab.SelectTab "K1"
        frmVentas.frmPagos.Visible = False
        frmVentas.frmApartados.Visible = False
        frmVentas.frmVentasMostrador.Visible = True
        BringWindowToTop frmVentas.hWnd
    
    Case eToolBar.Divisas
        
        frmDivisas.Show
        BringWindowToTop frmDivisas.hWnd
    
    Case eToolBar.Salir
        
        Unload Me
    End Select

End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    
    If MsgBox("Esta seguro que desea salir del sistema ??", vbQuestion + vbYesNo + vbDefaultButton2, "Casa de Empeños para ti") = vbYes Then
        
        'Modifico el ODBC de Datos
        ModificaODBC "BaseDatos", Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")), "BaseDatos"
        
        'Modifico el ODBC de Reportes
        ModificaODBC "BaseReportes", Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")), "BaseReportes"
        
        Cancel = False
        dbDatos.Close
        dbReportes.Close
        Set dbDatos = Nothing
        Set dbReportes = Nothing
        End
    
    Else
        
        Cancel = True
    
    End If

End Sub

Private Sub tmrHora_Timer()
    PaneHora.text = UCase(Regresar_Fecha(Date) & " - " & Format(Time, "HH:MM:SS AM/PM"))
End Sub

'Sacamos el reporte auxiliar
Private Sub Reporte_Auxiliar(Optional Opcion As Boolean = False, Optional FechaIni As String, Optional FechaFin As String)
   On Error GoTo Error
   Dim rcAuxiliar As New ADODB.Recordset
   Dim crCargo As Currency
   Dim crAbono As Currency
   Dim crSaldo As Currency
   Dim strCuenta As String
   
   dbReportes.Execute "DELETE FROM CorteCuentas"
   If Opcion Then
      rcAuxiliar.Open "SELECT Auxiliar.*,Cuentas.Mayor,Cuentas.Concepto FROM Auxiliar,Cuentas  WHERE Cuentas.Cuenta=Auxiliar.Cuenta AND Fecha BETWEEN '" & Format(FechaIni, "YYYY/MM/DD") & "' AND '" & Format(FechaFin, "YYYY/MM/DD") & "'  ORDER BY Cuentas.Mayor,Fecha", dbDatos, adOpenForwardOnly, adLockOptimistic   '  WHERE Fecha=#" & Format(Date, "MM/DD/YY") & "#", dbDatos, adOpenForwardOnly, adLockOptimistic
   Else
      rcAuxiliar.Open "SELECT Auxiliar.*,Cuentas.Mayor,Cuentas.Concepto FROM Auxiliar,Cuentas  WHERE Cuentas.Cuenta=Auxiliar.Cuenta AND Fecha='" & Format(Date, "YYYY/MM/DD") & "'  ORDER BY Cuentas.Mayor,Fecha", dbDatos, adOpenForwardOnly, adLockOptimistic '  WHERE Fecha=#" & Format(Date, "MM/DD/YY") & "#", dbDatos, adOpenForwardOnly, adLockOptimistic
   End If
   
   With rcAuxiliar
      While Not .EOF
         crCargo = 0
         crAbono = 0
         If strCuenta <> !Mayor Then
            crSaldo = 0
            strCuenta = !Mayor
         End If
         If Right(!Cuenta, 2) = "01" Then
            crCargo = !Importe
            crSaldo = crSaldo + crCargo
         Else
            crAbono = !Importe
            crSaldo = crSaldo - crAbono
         End If
         
         dbReportes.Execute "INSERT INTO CorteCuentas (Cuenta,Descripcion,Fecha,Concepto,Folio,Cargo,Abono,Saldo) VALUES " & _
                                           "('" & !Mayor & "','" & ![Cuentas.Concepto] & "','" & Format(!Fecha, "YYYY/MM/DD") & "','" & ![Auxiliar.Concepto] & "'," & !Folio & "," & crCargo & "," & crAbono & "," & crSaldo & ")"
         .MoveNext
      Wend
   End With
   
   rcAuxiliar.Close

Error:
   Maneja_Error Err
   
   Set rcAuxiliar = Nothing

End Sub

Sub RepIngresos(Fecha1 As Date, Fecha2 As Date)
Dim rcTmp As New ADODB.Recordset
Dim rcConsulta As New ADODB.Recordset
Dim crImporteInteres As Double, crImporteIva As Double, Iva As Double

On Error GoTo Error

    DoEvents
    dbReportes.Execute "DELETE FROM repingresos"

    'Intereses contratos tradicionales
    rcConsulta.Open "SELECT empeno.FechaMovimiento AS Fecha,(SUM(empeno.Intereses)+SUM(empeno.ImporteAlmacenaje)+SUM(empeno.ImporteSeguro)+SUM(empeno.ImporteMoratorios)) AS Intereses,SUM(empeno.ImportePerdida) AS ImportePerdida,SUM(empeno.ImporteIva) AS Iva FROM empeno WHERE empeno.Cancelado=0 AND DATE_FORMAT(empeno.FechaMovimiento,'%Y%/%m%/%d')>='" & Format(Fecha1, "YYYY/MM/DD") & "' AND DATE_FORMAT(empeno.FechaMovimiento,'%Y%/%m%/%d')<='" & Format(Fecha2, "YYYY/MM/DD") & "' GROUP BY DATE_FORMAT(empeno.FechaMovimiento,'%Y%/%m%/%d') ORDER BY empeno.Fechamovimiento", dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rcConsulta.EOF
        
        dbReportes.Execute "INSERT INTO repingresos (Fecha,Intereses,Iva,OtrosIngre) VALUES ('" & _
                            Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'," & rcConsulta!Intereses & "," & rcConsulta!Iva & "," & rcConsulta!ImportePerdida & ")"
    
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    
    'Intereses contratos pagos fijos
    rcConsulta.Open "SELECT pagosfijos.FechaMovimiento AS Fecha,(SUM(pagosfijos.Interes)+SUM(pagosfijos.Almacenaje)+SUM(pagosfijos.Seguro)+SUM(pagosfijos.Moratorios)) AS Intereses FROM pagosfijos WHERE pagosfijos.Cancelado=0 AND DATE_FORMAT(pagosfijos.FechaMovimiento,'%Y%/%m%/%d')>='" & Format(Fecha1, "YYYY/MM/DD") & "' AND DATE_FORMAT(pagosfijos.FechaMovimiento,'%Y%/%m%/%d')<='" & Format(Fecha2, "YYYY/MM/DD") & "' GROUP BY DATE_FORMAT(pagosfijos.FechaMovimiento,'%Y%/%m%/%d') ORDER BY pagosfijos.Fechamovimiento", dbDatos, adOpenForwardOnly, adLockReadOnly
    Iva = Regresa_Valor_BD("IVA") / 100
    While Not rcConsulta.EOF
    
        crImporteInteres = Redondeo(rcConsulta!Intereses / (1 + Iva))
        crImporteIva = Redondeo(rcConsulta!Intereses - crImporteInteres)
        
        rcTmp.Open "SELECT Fecha FROM repingresos WHERE Fecha='" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'", dbReportes, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF Then
            
            dbReportes.Execute "UPDATE repingresos SET Intereses=Intereses+" & crImporteInteres & ",Iva=Iva+ " & crImporteIva & " WHERE Fecha='" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'"
        Else
            
            dbReportes.Execute "INSERT INTO repingresos (Fecha,Intereses,Iva,OtrosIngre) VALUES ('" & _
                            Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'," & crImporteInteres & "," & crImporteIva & ",0)"
        End If
        rcTmp.Close
    
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    
    'Ventas
    rcConsulta.Open "SELECT ventas.Fecha,SUM((ventas.Total - (ventas.Total * (ventas.Descuento/100)))) AS Ventas,SUM(((ventas.Total - (ventas.Total * (ventas.Descuento/100))) * (ventas.IVA/100))) AS ImporteIva FROM ventas WHERE DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d')>='" & Format(Fecha1, "YYYY/MM/DD") & "' AND DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d')<='" & Format(Fecha2, "YYYY/MM/DD") & "' AND ventas.Cancelado=0 AND ventas.Apartado=0 AND ventas.TipoVenta=" & VENTAMOSTRADOR & " GROUP BY DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d') ORDER BY ventas.Fecha", dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rcConsulta.EOF

        rcTmp.Open "SELECT Fecha FROM repingresos WHERE Fecha='" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'", dbReportes, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF Then
            
            dbReportes.Execute "UPDATE repingresos SET Ventas=Ventas+" & rcConsulta!Ventas & ",Iva=Iva+ " & rcConsulta!ImporteIva & " WHERE Fecha='" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'"
        Else
            
            dbReportes.Execute "INSERT INTO repingresos (Fecha,Ventas,Iva) VALUES ('" & _
                                Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'," & rcConsulta!Ventas & "," & rcConsulta!ImporteIva & ")"
        End If
        rcTmp.Close
        
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    
    'Ventas Cliente
    rcConsulta.Open "SELECT ventas.Fecha,SUM(ventas.Total) As Total,SUM(dv.ImporteIva) AS Iva FROM ventas INNER JOIN detallesventas dv ON ventas.ID=dv.IDVenta WHERE DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d')>='" & Format(Fecha1, "YYYY/MM/DD") & "' AND DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d')<='" & Format(Fecha2, "YYYY/MM/DD") & "' AND ventas.Cancelado=0 AND ventas.Apartado=0 AND ventas.TipoVenta=" & VENTACLIENTE & " GROUP BY DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d') ORDER BY ventas.Fecha", dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rcConsulta.EOF

        rcTmp.Open "SELECT Fecha FROM repingresos WHERE Fecha='" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'", dbReportes, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF Then
            
            dbReportes.Execute "UPDATE repingresos SET Ventas=Ventas+" & rcConsulta!Total & ",Iva=Iva+" & rcConsulta!Iva & " WHERE Fecha='" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'"
        Else
            
            dbReportes.Execute "INSERT INTO repingresos (Fecha,Ventas,Iva) VALUES ('" & _
                                Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'," & rcConsulta!Total & "," & rcConsulta!Iva & ")"
        End If
        rcTmp.Close
        
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    
    'Abonos
    rcConsulta.Open "SELECT abonos.Fecha,SUM(abonos.Importe) AS Importe FROM abonos LEFT JOIN ventas ON abonos.IDVenta=ventas.ID WHERE abonos.Cancelado=0 AND DATE_FORMAT(abonos.Fecha,'%Y%/%m%/%d')>=date('" & Format(Fecha1, "YYYY/MM/DD") & "') AND DATE_FORMAT(abonos.Fecha,'%Y%/%m%/%d')<=date('" & Format(Fecha2, "YYYY/MM/DD") & "') GROUP BY DATE_FORMAT(abonos.Fecha,'%Y%/%m%/%d') ORDER BY abonos.Fecha,abonos.ID", dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rcConsulta.EOF
    
        rcTmp.Open "SELECT Fecha FROM repingresos WHERE Fecha=date('" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "')", dbReportes, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF Then
            
            dbReportes.Execute "UPDATE repingresos SET Apartados=Apartados+" & rcConsulta!Importe & " WHERE Fecha=date('" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "')"
        Else
            
            dbReportes.Execute "INSERT INTO repingresos (Fecha,Apartados) VALUES ('" & Format(rcConsulta!Fecha, "YYYY/MM/DD") & "'," & rcConsulta!Importe & ")"
        End If
        rcTmp.Close
        
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    Set rcConsulta = Nothing
    Set rcTmp = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
    Set rcTmp = Nothing
End Sub

Sub ExistenciasDivisas()
Dim rcConsulta As New ADODB.Recordset
Dim rcDivisas As New ADODB.Recordset
Dim TipoCambioCompra As Double, TipoCambio As String
Dim Cargo As Long, Abono As Long
Dim rcAux As New ADODB.Recordset
Dim rcBD As New ADODB.Recordset
Dim Compras As Long, Ventas As Long
On Error GoTo Error
    
    dbReportes.Execute "DELETE FROM existencia_divisas"
    
    rcAux.Open "SELECT DISTINCT a.IDDivisa,m.Descripcion AS Divisa FROM auxiliar a INNER JOIN monedas m ON a.IDDivisa=m.Clave WHERE a.Cuenta='910901' OR a.Cuenta='910950'", dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rcAux.EOF
                                
        TipoCambio = ""
        TipoCambioCompra = 0
        Cargo = 0: Abono = 0
        TipoCambio = SacaValor("divisas", "SUM(IMPORTE)/COUNT(ID)", " WHERE Cancelado=0 AND Tipo=0 AND IDDivisa=" & rcAux!IDDivisa)
                 
        'Saldo Inicial***********
        rcBD.Open "SELECT SUM(a.Importe) AS Cargo FROM auxiliar a WHERE Fecha<'" & Format(Date, "YYYY/MM/DD") & "' AND Cuenta='910901' AND a.IDDivisa=" & rcAux!IDDivisa, dbDatos, adOpenForwardOnly, adLockOptimistic
            Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
        rcBD.Close
    
        rcBD.Open "SELECT SUM(a.Importe) AS Abono FROM auxiliar a WHERE Fecha<'" & Format(Date, "YYYY/MM/DD") & "' AND Cuenta='910950' AND a.IDDivisa=" & rcAux!IDDivisa, dbDatos, adOpenForwardOnly, adLockOptimistic
            Abono = IIf(IsNull(rcBD!Abono), 0, rcBD!Abono)
        rcBD.Close
        '********************
        
        Compras = 0: Ventas = 0
        'Compras
        rcBD.Open "SELECT SUM(a.Importe) AS Cargo FROM auxiliar a WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND Cuenta='710301' AND Serie=2 AND a.IDDivisa=" & rcAux!IDDivisa, dbDatos, adOpenForwardOnly, adLockOptimistic
            Compras = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
        rcBD.Close
        
        'Ventas
        rcBD.Open "SELECT SUM(a.Importe) AS Abono FROM auxiliar a WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND Cuenta='710350' AND Serie=2 AND a.IDDivisa=" & rcAux!IDDivisa, dbDatos, adOpenForwardOnly, adLockOptimistic
            Ventas = IIf(IsNull(rcBD!Abono), 0, rcBD!Abono)
        rcBD.Close
        
        dbReportes.Execute "INSERT into existencia_divisas(Divisa,EntradaInicial,TipoCambio,Entrada,Salida) VALUES (" & _
                                rcAux!IDDivisa & "," & Cargo - Abono & "," & TipoCambio & "," & Compras & "," & Ventas & ")"
    
    rcAux.MoveNext
    Wend
    rcAux.Close
    Set rcBD = Nothing
    Set rcAux = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcBD = Nothing
    Set rcAux = Nothing
End Sub
