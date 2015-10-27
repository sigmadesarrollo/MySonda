VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmConexionSucursal 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conexión Sucursales"
   ClientHeight    =   4035
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   Icon            =   "frmConexionSucursal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   5325
   Begin vbAcceleratorGrid6.vbalGrid grdSucursales 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   6800
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
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4140
      TabIndex        =   1
      Top             =   720
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
      Picture         =   "frmConexionSucursal.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdConectar 
      Height          =   375
      Left            =   4140
      TabIndex        =   2
      Top             =   240
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Conectar"
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
      TextColor       =   4210752
      Object.ToolTipText     =   ""
      Picture         =   "frmConexionSucursal.frx":055E
   End
End
Attribute VB_Name = "frmConexionSucursal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim pIDUsuario As Integer, pUsuario As String, pClaveSucursal As Integer, pStrSucursal As String, pStrRazonSucursal As String

Public Property Let IDUsuario(Valor As Integer)
    pIDUsuario = Valor
End Property

Public Property Get IDUsuario() As Integer
    IDUsuario = pIDUsuario
End Property

Public Property Let Usuario(Valor As String)
    pUsuario = Valor
End Property

Public Property Get Usuario() As String
    Usuario = pUsuario
End Property

Private Sub cmdConectar_Click()
Dim strIP As String, strSucursal As String

On Error GoTo error

    If grdSucursales.SelectedRow > 0 Then
    
        If MsgBox("Desea conectarse a la sucursal de: " & grdSucursales.CellText(grdSucursales.SelectedRow, 1) & " ??", vbQuestion + vbYesNo + vbDefaultButton2, "Conexión Sucursales") = vbYes Then
            
            Screen.MousePointer = vbHourglass
            
            'Tomo los datos de la nueva sucursal
            DatosSucursal False, grdSucursales.CellText(grdSucursales.SelectedRow, 3)
            
            sServidor = grdSucursales.CellText(grdSucursales.SelectedRow, 2)
            strSucursal = grdSucursales.CellText(grdSucursales.SelectedRow, 1)
            
            dbDatos.Close
            dbDatos.Open cCONEXION & sServidor & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDB
            
            dbReportes.Close
            dbReportes.Open cCONEXION & sServidor & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDBR
            
            'Modifico los ODBC
            ModificaODBC "BaseDatos", sServidor, "BaseDatos"
            ModificaODBC "BaseReportes", sServidor, "BaseReportes"
            
            'Cargo las Sucursales de nuevo
            Cargar_Datos
            
            frmMDI.IDUsuario = Me.IDUsuario
            frmMDI.Usuario = Me.Usuario
                        
            PaneSucursal.text = "SUCURSAL: " & strSucursal
            
            MsgBox "Se ha conectado a la sucursal de: " & strSucursal, vbInformation, "Conexión Sucursales"
            Screen.MousePointer = vbDefault
        End If
    
    Else
        
        MsgBox "Seleccione la sucursal a la que desea conectarse !!", vbInformation, "Conexión Sucursales"
    End If
    Exit Sub
    
error:
    Maneja_Error Err
    If dbDatos.State = adStateOpen Then dbDatos.Close
    If dbReportes.State = adStateOpen Then dbReportes.Close
    
    dbDatos.Open cCONEXION & Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")) & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDB
    dbReportes.Open cCONEXION & Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")) & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDBR
    
    'Modifico los ODBC
    ModificaODBC "BaseDatos", Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")), "BaseDatos"
    ModificaODBC "BaseReportes", Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")), "BaseReportes"
        
    'Regreso los datos de la sucursal por default
    DatosSucursal True
    
    PaneSucursal.text = "SUCURSAL: " & Sucursal.NombreComercial
    
    MsgBox "No es posible conectarse en este momento a la sucursal seleccionada !!", vbInformation, "Conexión Sucursales"
    Screen.MousePointer = vbDefault
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Crear_Encabezados
    Cargar_Datos
    CentrarForm Me, frmMDI
End Sub

Sub Crear_Encabezados()
    With grdSucursales
        .AddColumn "C1", "Sucursal", ecgHdrTextALignLeft, , 250, , , , , , , CCLSortString
        .AddColumn "C2", "IP", ecgHdrTextALignLeft, , 50, False, , , , , , CCLSortString
        .AddColumn "C3", "Clave", ecgHdrTextALignLeft, , 50, False, , , , , , CCLSortString
    End With
End Sub

Sub Cargar_Datos()
Dim rcSucursales As New ADODB.Recordset

On Error GoTo error

    With rcSucursales
        
        grdSucursales.Redraw = False
        grdSucursales.Clear
        .Open "SELECT ID,Clave,NombreComercial,IP FROM sucursales WHERE Activa=0 ORDER BY NombreComercial", dbDatos, adOpenForwardOnly, adLockReadOnly
        While Not .EOF
            
            grdSucursales.AddRow
            grdSucursales.CellText(grdSucursales.Rows, 1) = !NombreComercial
            grdSucursales.CellItemData(grdSucursales.Rows, 1) = !ID
            grdSucursales.CellText(grdSucursales.Rows, 2) = !Ip
            grdSucursales.CellText(grdSucursales.Rows, 3) = !Clave
            Colorea grdSucursales, grdSucursales.Rows, IIf(grdSucursales.Rows Mod 2 <> 0, Trim(RGB(242, 254, 255)), Trim(RGB(255, 255, 255)))
        .MoveNext
        Wend
        .Close
        Set rcSucursales = Nothing
        
        grdSucursales.Redraw = True
    End With
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcSucursales = Nothing
End Sub

Private Sub grdSucursales_DblClick(ByVal lRow As Long, ByVal lCol As Long)
Dim strIP As String, strSucursal As String

On Error GoTo error

    If grdSucursales.SelectedRow > 0 Then
    
        If MsgBox("Desea conectarse a la sucursal de: " & grdSucursales.CellText(grdSucursales.SelectedRow, 1) & " ??", vbQuestion + vbYesNo + vbDefaultButton2, "Conexión Sucursales") = vbYes Then
            
            Screen.MousePointer = vbHourglass
            
            'Tomo los datos de la nueva sucursal
            DatosSucursal False, grdSucursales.CellText(grdSucursales.SelectedRow, 3)
            
            sServidor = grdSucursales.CellText(grdSucursales.SelectedRow, 2)
            strSucursal = grdSucursales.CellText(grdSucursales.SelectedRow, 1)
            
            dbDatos.Close
            dbDatos.Open cCONEXION & sServidor & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDB
            
            dbReportes.Close
            dbReportes.Open cCONEXION & sServidor & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDBR
            
            'Modifico los ODBC
            ModificaODBC "BaseDatos", sServidor, "BaseDatos"
            ModificaODBC "BaseReportes", sServidor, "BaseReportes"
            
            'Cargo las Sucursales de nuevo
            Cargar_Datos
            
            frmMDI.IDUsuario = Me.IDUsuario
            frmMDI.Usuario = Me.Usuario
                        
            PaneSucursal.text = "SUCURSAL: " & strSucursal
            
            MsgBox "Se ha conectado a la sucursal de: " & strSucursal, vbInformation, "Conexión Sucursales"
            Screen.MousePointer = vbDefault
        End If
    
    Else
        
        MsgBox "Seleccione la sucursal a la que desea conectarse !!", vbInformation, "Conexión Sucursales"
    End If
    Exit Sub
    
error:
    Maneja_Error Err
    If dbDatos.State = adStateOpen Then dbDatos.Close
    If dbReportes.State = adStateOpen Then dbReportes.Close
    
    dbDatos.Open cCONEXION & Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")) & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDB
    dbReportes.Open cCONEXION & Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")) & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDBR
    
    'Modifico los ODBC
    ModificaODBC "BaseDatos", Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")), "BaseDatos"
    ModificaODBC "BaseReportes", Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost")), "BaseReportes"
        
    'Regreso los datos de la sucursal por default
    DatosSucursal True
    
    PaneSucursal.text = "SUCURSAL: " & Sucursal.NombreComercial
    MsgBox "No es posible conectarse en este momento a la sucursal seleccionada !!", vbInformation, "Conexión Sucursales"
    Screen.MousePointer = vbDefault
End Sub
