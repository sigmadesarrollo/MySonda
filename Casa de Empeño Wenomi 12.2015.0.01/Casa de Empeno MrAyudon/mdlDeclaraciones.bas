Attribute VB_Name = "mdlDeclaraciones"
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 02/04/02
' Modulo mdlDeclaraciones - mdlDeclaraciones.bas
' Ultima Modificacion - 05/04/02
' Modificacion para mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit

'Funciones api para tomar el separador decimal
Public Separador As String
Public Const LOCALE_SDECIMAL = &HE
Public Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Public Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Public Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'Declaraciones de la api de windows
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetTempFileName Lib "kernel32" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long

Public Declare Function GetPrivateProfileSectionNames Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" (ByVal lpszReturnBuffer As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hWnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As rect) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const ODBC_CONFIG_SYS_DSN As Long = 5
Public Declare Function SQLConfigDataSource Lib "ODBCCP32.DLL" (ByVal hwndParent As Long, ByVal fRequest As Long, ByVal lpszDriver As String, ByVal lpszAttributes As String) As Long
Public Declare Function ShellExecuteEx Lib "shell32.dll" (SEI As SHELLEXECUTEINFO) As Long

'FACTURACION
Public Const SW_SHOWNORMAL = 1
Public Const STATUS_PENDING = &H103&
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long


' La versión simple
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Const SEE_MASK_INVOKEIDLIST = &HC
Public Const SEE_MASK_NOCLOSEPROCESS = &H40
Public Const SEE_MASK_FLAG_NO_UI = &H400
Public Const SEE_MASK_DEFAULT = &H0

Public Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Type DatosSucursal
    Clave As Integer
    RazonSocial As String
    NombreComercial As String
    RFC As String
    Direccion As String
    Ciudad As String
    Estado As String
    Telefono As String
    CP As String
End Type

Public stBar As XtremeCommandBars.IStatusBar
Public Toolbar As CommandBar
Public PaneHora As StatusBarPane
Public PaneSucursal As StatusBarPane

'Public Enum eToolBar
'    Buscar = 10000
'    Empeno
'    Cierre
'    Venta
'    Divisas
'    Salir
'End Enum

Public Enum eToolBar
    Buscar = 1001
    Empeno = 1007
    Cierre
    Venta = 1004
    Divisas
    Salir = 1010
End Enum


Public Enum eStatusBar
    ID_INDICATOR_CAPS = 59137
    ID_INDICATOR_NUM = 59138
    ID_INDICATOR_SCRL = 59139
End Enum

'Constantes para los mensajes
Private Const WM_USER = &H400
Public Const MAX_COMPUTERNAME_LENGTH As Long = 31
Public Const WM_CONFIGURACION = WM_USER + 1

'Constantes para utilizar con las opciones del sistema
'Constantes para las series
Public Const SERIE_A = 1
Public Const SERIE_B = 2 'Autos
Public Const SERIE_C = 3
Public Const SERIE_D = 4 'Electronicos

'Cargo o Abono
Public Const TIPO_CARGO = 1
Public Const TIPO_ABONO = 2

'Constantes para el origen
Public Const OD_EMPENO = 1
Public Const OD_REFRENDO = 2
Public Const D_DESEMPEÑO = 3
Public Const D_ALMONEDA = 4
Public Const D_VENTA = 5
Public Const D_FUNDICION = 6
Public Const D_OTRO = 7
Public Const D_CENTRAL = 8



Public Const CANCELACION = 2
Public Const GERENTE = 4

Public Cajero As String
Public IDCajero As String

'////////roger
'para la cadena de conexion a la base de datos
Public Const CONEXION_ = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="

'para la cadena de conexion que contiene el usuario y el password
Public Const Usuario_ = ";Jet OLEDB:Database Password=administrativo"
'////////
'Variables para la conexion con la base de datos
Public dbDatos As New ADODB.Connection
Public dbReportes As New ADODB.Connection
'Variables para la conexion con la base de datos old
Public dbDatos_old As New ADODB.Connection
Public dbReportes_old As New ADODB.Connection

'Variables recordset
Public rcConsulta As ADODB.Recordset
Public rcTmp As ADODB.Recordset

'Variable para la ruta de la aplicacion
Public Path As String

'Variable para el servidor de la base de datos
Public sServidor As String

'Variable para manejar los datos de la sucursal
Public Sucursal As DatosSucursal

'Para la cadena de conexion a la base de datos
Public Const USERBD = "mrayudon"
Public Const PWDBD = "montepio"
Public Const cCONEXION = "Provider=MSDASQL;DRIVER={MySQL ODBC 3.51 Driver}; Server="
Public Const cDB = "; database=BaseDatos; UID=" & USERBD & "; pwd=" & PWDBD & "; Option=" & 1 + 2 + 8 + 32 + 2048 + 16384
Public Const cDBR = "; database=BaseReportes; UID=" & USERBD & "; pwd=" & PWDBD & "; Option=" & 1 + 2 + 8 + 32 + 2048 + 16384
'Para la cadena de conexion a la base de datos que se va a migrar 15/06/2015
Public Const cDB_old = "; database=BaseDatos_old; UID=" & USERBD & "; pwd=" & PWDBD & "; Option=" & 1 + 2 + 8 + 32 + 2048 + 16384
Public Const cDBR_old = "; database=BaseReportes_old; UID=" & USERBD & "; pwd=" & PWDBD & "; Option=" & 1 + 2 + 8 + 32 + 2048 + 16384

'Nombre de la Maquina
Public NombrePc As String

'Estructura de un cuadro de region
Public Type rect
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

'Constantes de la api de windows
Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_WININICHANGE = &H1A
Public Const CB_FINDSTRING = &H14C
Public Const CB_ERR = (-1)
Public Const CB_SETCURSEL = &H14E


'Variables para la camara
Global Const ws_child As Long = &H40000000
Global Const ws_visible As Long = &H10000000
Global Const wm_cap_driver_connect = WM_USER + 10
Global Const wm_cap_set_preview = WM_USER + 50
Global Const WM_CAP_SET_PREVIEWRATE = WM_USER + 52
Global Const WM_CAP_DRIVER_DISCONNECT As Long = WM_USER + 11
Global Const WM_CAP_DLG_VIDEOFORMAT As Long = WM_USER + 41
Global Const WM_CAP_COPY As Long = 1054
Global Const WM_CAP_GET_FRAME = 1084
Global Const WM_VIDEOFORMAT_COLOR As Long = WM_USER + 42
Global Const WM_VIDEOFORMAT_COMPRESION As Long = WM_USER + 46
Global Const WM_VIDEOFORMAT_PREVEIW As Long = WM_USER + 52
Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal a As String, ByVal b As Long, ByVal c As Integer, ByVal d As Integer, ByVal e As Integer, ByVal F As Integer, ByVal g As Long, ByVal h As Integer) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Boolean

'Formato de Moneda
Public Const FMoneda = "###,###,###,###0.00"
Public Const FMonedaSigno = "$ ###,###,###,###0.00"

'Tipos de Entrada
Public Const ENTRADAEMPENO = 1
Public Const ENTRADADOTACION = 2
Public Const ENTRADACOMPRA = 3
Public Const ENTRADAALMONEDA = 4
Public Const ENTRADAMIGRACION = 5
Public Const ENTRADATRASPASO = 7
'Tipos de Salida
Public Const SALIDAVENTA = 1
Public Const SALIDATRASPASO = 2
Public Const SALIDAVENTAFUNDICION = 3
Public Const SALIDAINVENTARIO = 4
Public Const SALIDAREEMPENO = 5
Public Const SALIDAVENTAPIGNORANTE = 6
Public Const SALIDAAPARTADO = 7

'Tipos de Venta
Public Const VENTAMOSTRADOR = 0
Public Const VENTACLIENTE = 1
Public Const VENTAMAYORISTA = 2

'Tipo de Autorizaciones
Public Const AUTORIZACIONLIMITE1 = 1
Public Const AUTORIZACIONLIMITE2 = 2

'Datos para la impresora de tickets
Public strNombreImp As String
Public strDriverImp As String
Public strPuertoImp As String

Public Enum eTipoImpresora
    Contratos
    Tickets
    EtiquetasEmpeno
    EtiquetasAlmoneda
End Enum

'Conexión al Web Service
Public WSoap As SoapClient30
Public WServidor As String
Public WRutaServidor As String
Public WPuerto As Integer
Public WBaseDatos As String

Public Type EmpenoForaneo
    ID As Long
    Cancelado As Integer
    Fecha As Date
    FechaOriginal As Date
    Movimiento As Long
    NumContrato As Long
    Folio As Long
    Prestamo As Double
    PrestamoInicial As Double
    Avaluo As Double
    Origen As Integer
    Destino As Integer
    Vencimiento As Date
    FolioOrigen As Long
    FolioDestino As Long
    FolioOriginal As Long
    FechaMovimiento As Date
    FechaPagoParcial As Date
    IDUsuarioMov As Long
    Serie As Integer
    Pagado As Integer
    PC As String
    Corte As Integer
    Perdida As Integer
    IDCliente As Long
    IDTablaCliente As Long
    Responsable As String
    Beneficiario As String
    Valuador As String
    Notas As String
    Tasa As Double
    Almacenaje As Double
    Seguro As Double
    Operacion As Double
    Comision As Double
    GastosAdmon As Double
    Iva As Double
    CAT As Double
    Periodo As Long
    VenPeriodo As Long
    Almoneda As Integer
    FechaAlmoneda As Date
    VenAlmoneda As Integer
    TipoInteres As String
    TipoTasa As String
    IDSucursal As Long
    IDUsuario As Long
    Pago As Double
    Intereses As Double
    importeAlmacenaje As Double
    importeSeguro As Double
    ImporteMoratorios As Double
    ImportePerdida As Double
    Descuento As Double
    ImporteIva As Double
    AutTasa As Long
    ChequeReferencia As String
    ImporteOtros As Double
    DemasiaPagada As Long
    IDAutorizacion As Long
    Captura As Integer
    NumBolsa As String
    Verificado As Integer
    IDUsuarioAutoriza As Long
    TipoAutoriza As Long
    ubicacion As String
    FolioNota As Long
    Efectivo As Double
    caja As String
    Cajon As String
    Fila As String
    Promocion As Long
    SaldoPuntosAnteriorEmp As Double
    PuntosAcumuladosEmp As Double
    SaldoPuntosActualEmp As Double
    IDTarjetaEmp As Long
    DescuentoXPuntos As Double
    SaldoPuntosAnterior As Double
    PuntosUsados As Double
    PuntosAcumulados As Double
    SaldoPuntosActual As Double
    IDTarjeta As Long
    IDCotitular As Long
    Cheque As Integer
    SalarioMin As Integer
    ValorSalarioMin As Double
    ValorUDI As Double
    UltDigitosTarj As String
    IdTipoOperacion As Long
    ClaveTipoOperacion As Long
    IdInstrumentoMonetario As Long
    IdTipoMoneda As Long
    IDTipoAlerta As Long
    DescTipoAlerta As String
    IDTipoPrenda As Long
    IDEmpenoOrigen As Long
    IDEmpenoDestino As Long
    NumRefrendos As Long
    Bloqueado As Integer
    MotivoBloqueo As String
    Migrada As Integer
End Type

'Declaro la Estructura para el Cliente Foráneo
Public Type ClienteForaneo
    ID As Long
    Iniciales As String
    Nombre As String
    Apellido As String
    Direccion As String
    Colonia As String
    Ciudad As String
    Estado As String
    CP As Long
    FecNac As Date
    Identificacion As String
    NumeroIdentificacion As String
    Telefono As String
End Type

'Declaro la Estructura para DetallesEmpeno
Public Type DetallesEmpenoForaneo
    IDEmpeno As Long
    Codigo As String
    Tipo As Integer
    IDTablaTipo As Long
    Cantidad As Integer
    Articulo As String
    Peso As Double
    Kilates As Integer
    IDTablaKilates As Long
    Avaluo As Double
    Prestamo As Double
    Origen As Integer
    CantidadPiedras As Double
    PesoPiedras As Double
    CantidadDiamantes As Integer
    Puntos As Double
    PrestamoDiamante As Double
    Observaciones As String
    TipoPrenda As Integer
    Estado As String
    Marca As String
    Modelo As String
    Serie As String
    Color As String
    Tamano As String
End Type

'Declaro la Estructura para DetallesEmpeno Autos
Public Type DetallesEmpenoForaneoA
    IDEmpeno As Long
    MarcayModelo As String
    Año As Long
    Color As String
    Placas As String
    Factura As String
    Agencia As String
    NumTarjetaCircu As String
    NumMotor As String
    SerieChasis As String
    Kms As String
    Gas As String
    Aseguradora As String
    Poliza As String
    FechaVenci As Date
    Tipo As String
    Observaciones As String
    TipoMovil As Integer
    TipoDesc As String
End Type

Public Type ArticulosForaneos
    NombreComercial As String
    Clave As Integer
    Codigo As String
    Articulo As String
    Modelo As String
    Marca As String
    Peso As Double
    Estado As String
    Kilates As String
    PrecioVitrina As Currency
End Type

Public Type Apartado
    ID As Long
    Fecha As Date
    IDCliente As Long
    IDTablaCliente As Long
    Iva As Double
    Vencimiento As Date
    Total As Currency
End Type

Public Foraneo As EmpenoForaneo
Public ForaneoCuenta() As EmpenoForaneo
Public ClienteForaneo As ClienteForaneo
Public DetallesEmpenoForaneo() As DetallesEmpenoForaneo
Public DetallesEmpenoFA() As DetallesEmpenoForaneoA


