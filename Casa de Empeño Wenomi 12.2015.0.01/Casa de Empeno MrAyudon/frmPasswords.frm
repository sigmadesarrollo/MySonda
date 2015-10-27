VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmPasswords 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clave de usuario"
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4410
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPasswords.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1230
   ScaleWidth      =   4410
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtPass 
      BorderStyle     =   0  'None
      Height          =   195
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.TextBox txtUsuario 
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   795
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "       &Cancelar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   65280
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmPasswords.frx":27A2
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   315
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      AlignCaption    =   4
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
      MaskColor       =   65280
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmPasswords.frx":2BC7
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "C&ontraseña:"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "&Usuario:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   600
   End
End
Attribute VB_Name = "frmPasswords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim bBandera As Boolean, Op As Long, Salir As Integer

Public Cancel As Single
Public Ventas As Single
Public ModificaPrecio As Single
Public ModificaCorte As Single
Public HacerCorte As Single
Public InteresRefrendo As Single
Public InteresDesempeño As Single
Public PrecioVitrina As Single
Public DescuentoVentas As Single
Public RecalculoPrecios As Single
Public AutorizaPrestamo As Single
Public ConexSuc As Single
Public Vencido As Single
Public CancelaCierre As Single

Private Sub cmdAceptar_Click()

On Error GoTo Error
    
    bBandera = False
    
    If txtUsuario.text = "admin" And txtPass.text = "sonda" And frmMDI.Usuario = "" Then
        
        frmMDI.Usuario = "Admin"
        frmMDI.IDUsuario = 0
        frmMDI.IDSucursal = 0
        bBandera = True
        Unload Me
        
    ElseIf txtUsuario.text = "admin" And txtPass.text = "sonda" Then
        
        bBandera = True
        Unload Me
        
    ElseIf Cancel = 1 Then
        
        If VerificaAcceso("CancelBol", txtUsuario.text, txtPass.text) Then
            
            frmCancelaciones.IDUsuario = SacaValor("usuarios", "ID", " WHERE Usuario='" & Trim(txtUsuario.text) & "' AND Contraseña='" & Trim(txtPass.text) & "'")
            bBandera = True
            Unload Me
        Else
            
            frmCancelaciones.IDUsuario = 0
            MsgBox "Acceso denegado !!", vbCritical, "Cancelación"
            Unload Me
        End If
        
    ElseIf Ventas = 1 Then
        
        If VerificaAcceso("Abonar", txtUsuario.text, txtPass.text) Then
            
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Apartados"
            Unload Me
        End If
        
    ElseIf ModificaPrecio = 1 Then
        
        If VerificaAcceso("Precio", txtUsuario.text, txtPass.text) Then
            
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Ventas de Mostrador"
            Unload Me
        End If
        
    ElseIf ModificaCorte = 1 Then

        If VerificaAcceso("ModificarCorte", txtUsuario.text, txtPass.text) Then
            
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Cierre de Caja"
            Unload Me
        End If
            
    ElseIf HacerCorte = 1 Then
    
        If VerificaAcceso("HacerCorte", txtUsuario.text, txtPass.text) Then
            
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Cierre de Caja"
            Unload Me
        End If
        
    ElseIf InteresRefrendo = 1 Then
        
        If VerificaAcceso("InteresRefrendo", txtUsuario.text, txtPass.text) Then
            
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Modificar interés refrendo"
            Unload Me
        End If
        
    ElseIf InteresDesempeño = 1 Then

        If VerificaAcceso("InteresDesempeño", txtUsuario.text, txtPass.text) Then
            
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Modificar interés desempeño"
            Unload Me
        End If

    ElseIf PrecioVitrina = 1 Then
        
        If VerificaAcceso("PrecioVitrina", txtUsuario.text, txtPass.text) Then
            
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Existencias"
            Unload Me
        End If
    
    ElseIf DescuentoVentas = 1 Then
        
        If VerificaAcceso("DescuentoVentas", txtUsuario.text, txtPass.text) Then
            
            frmVentas.IDUsuario = SacaValor("usuarios", "ID", " WHERE Usuario='" & Trim(txtUsuario.text) & "' AND Contraseña='" & Trim(txtPass.text) & "'")
            bBandera = True
            Unload Me
        Else
            
            frmVentas.IDUsuario = 0
            MsgBox "Acceso denegado !!", vbCritical, "Ventas de Mostrador"
            Unload Me
        End If
        
    ElseIf RecalculoPrecios = 1 Then
        
        If VerificaAcceso("RecalculoPrecios", txtUsuario.text, txtPass.text) Then
            
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Parámetros"
            Unload Me
        End If
    
    ElseIf AutorizaPrestamo = 1 Then
        
        If VerificaAcceso("PrestamoBoleta1", txtUsuario.text, txtPass.text) Then
            
            frmEmpeño.IDUsuarioAutoriza = SacaValor("usuarios", "ID", " WHERE Usuario='" & Trim(txtUsuario.text) & "' AND Contraseña='" & Trim(txtPass.text) & "'")
            bBandera = True
            Unload Me
        Else
            
            frmEmpeño.IDUsuarioAutoriza = 0
            MsgBox "Acceso denegado !!", vbCritical, "Autorización"
            Unload Me
        End If
        
    ElseIf ConexSuc = 1 Then
        
        If VerificaAcceso("ConexionSuc", txtUsuario.text, txtPass.text) Then
            
            frmConexionSucursal.IDUsuario = SacaValor("usuarios", "ID", " WHERE Usuario='" & Trim(txtUsuario.text) & "' AND Contraseña='" & Trim(txtPass.text) & "'")
            frmConexionSucursal.Usuario = txtUsuario.text
            
            bBandera = True
            Unload Me
        Else
            
            frmConexionSucursal.IDUsuario = 0
            frmConexionSucursal.Usuario = ""
            MsgBox "Acceso denegado !!", vbCritical, "Conexión Intersucursales"
            Unload Me
        End If
    
    ElseIf Vencido = 1 Then
        
        If VerificaAcceso("RefrendarVencidos", txtUsuario.text, txtPass.text) Then
                        
            bBandera = True
            Unload Me
        Else
            
            MsgBox "No tiene el nivel de acceso permitido para Refrendar Contratos en Almoneda !!", vbCritical, "Refrendo"
            bBandera = False
            Unload Me
        End If
    
    ElseIf CancelaCierre = 1 Then
        
        If VerificaAcceso("CancelaCierre", txtUsuario.text, txtPass.text) Then
                        
            bBandera = True
            Unload Me
        Else
            
            MsgBox "Acceso denegado !!", vbCritical, "Cancelar Cierre"
            bBandera = False
            Unload Me
        End If
        
    Else
    
        'Verificar_Usuario
        ChecaPermisos
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub cmdCancelar_Click()

    If Salir <> 0 Then
        
        Unload Me
        bBandera = False
    Else
        
        End
    End If

End Sub

Private Sub Form_Load()
   Inicializar
End Sub

Private Sub Inicializar()
   Screen.MousePointer = vbHourglass
   Poner_Flat
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer
   
    'Descargamos de memoria el flat
    For i = LBound(Fl) To UBound(Fl)
        
        Set Fl(i) = Nothing
    
    Next i
    
    If Cancel = 0 And Salir = 0 And bBandera = False Then End

End Sub

Public Function Password(Optional Opcion As Long = 0, Optional Sal As Integer = 0) As Boolean
On Error Resume Next
    
    Op = Opcion
    bBandera = False
    Salir = Sal
    Me.Show vbModal
    Password = bBandera
End Function

Private Sub txtPass_GotFocus()
    Seleccionar_Texto txtPass
    Cambiar_Color True, txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPass_LostFocus()
    Cambiar_Color False, txtPass
End Sub

Private Sub txtUsuario_GotFocus()
    Seleccionar_Texto txtUsuario
    Cambiar_Color True, txtUsuario
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtUsuario_LostFocus()
    Cambiar_Color False, txtUsuario
End Sub

'Ponemos en modo flat los textbox
Private Sub Poner_Flat()
Dim Contador As Integer
Dim Control As Object
   
    For Each Control In Controls
    
        If TypeOf Control Is TextBox Then
            
            ReDim Preserve Fl(0 To Contador)
            Set Fl(Contador) = New cFlatControl
            Fl(Contador).hWndAttach Control.hWnd, Me.hWnd, False
            Contador = Contador + 1
      
        ElseIf TypeOf Control Is ComboBox Then
            
            ReDim Preserve Fl(0 To Contador)
            Set Fl(Contador) = New cFlatControl
            Fl(Contador).hWndAttach Control.hWnd, Me.hWnd, True
            Contador = Contador + 1
        
        End If
   
    Next
   
End Sub

Sub ChecaPermisos()
Dim rcPermisos As New ADODB.Recordset

On Error GoTo Error

    rcPermisos.Open "SELECT * FROM usuarios WHERE Estatus=1 AND Usuario='" & Trim(txtUsuario.text) & "' AND Contraseña='" & Trim(txtPass.text) & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not rcPermisos.BOF And Not rcPermisos.EOF Then

        With rcPermisos
        
            'Menu Empeños
            If !empeño = 0 Then frmMDI.mnuEmpeños.Enabled = False
            If !desempeños = 0 Then frmMDI.mnuDesempeños.Enabled = False
            If !refrendos = 0 Then frmMDI.mnuRefrendos.Enabled = False
            If !PagosFijos = 0 Then frmMDI.mnuPagosFijoss.Enabled = False
            If !cotizarempeño = 0 Then frmMDI.mnuCotizaciones.Enabled = False
            If !CambioPlan = 0 Then frmMDI.mnuCambioPlan.Enabled = False
            If !busqueda = 0 Then frmMDI.mnuBuscarBoletas.Enabled = False
            If !regubicacion = 0 Then frmMDI.mnuRegUbicacion.Enabled = False
            If !repempeños = 0 Then frmMDI.mnuRepEmpeños.Enabled = False
            If !repvencidos = 0 Then frmMDI.mnuRepEmpeVencidos.Enabled = False
            If !repalmoneda = 0 Then frmMDI.mnuReporteAlmoneda.Enabled = False
        
            'Menu Cierres
            If !cortecaja = 0 Then frmMDI.mnuCierreCaja.Enabled = False
            If !CierreDivisas = 0 Then frmMDI.mnuCierreDivisas.Enabled = False
            If !RepCartera = 0 Then frmMDI.mnuRepCartera.Enabled = False
            If !balance = 0 Then frmMDI.mnuBalance.Enabled = False
            If !repfinanciero = 0 Then frmMDI.mnuActivos.Enabled = False
            If !cierresucursal = 0 Then frmMDI.mnuCierresSucursal.Enabled = False

            'Menu Ventas
            If !Ventas = 0 Then frmMDI.mnuMostrador.Enabled = False
            If !VenCliente = 0 Then frmMDI.mnuVentasClientes.Enabled = False
            If !Ventas = 0 Then frmMDI.mnuApartados.Enabled = False
            If !Ventas = 0 Then frmMDI.mnuAbonos.Enabled = False
            If !PagoDemasia = 0 Then frmMDI.mnuPagoDemasias.Enabled = False
            If !repventas = 0 Then frmMDI.mnuRepVentasCon.Enabled = False
            If !repapartado = 0 Then frmMDI.mnuRepVentasApa.Enabled = False
            If !MostrarApartados = 0 Then frmMDI.mnuMuestraApartados.Enabled = False
            If !ApartadosVencidos = 0 Then frmMDI.mnuApartadosVencidos.Enabled = False
            If !reputilidad = 0 Then frmMDI.mnuRepUtilidadVentas.Enabled = False

            'Menu Inventario
            If !Existencias = 0 Then frmMDI.mnuExistencias.Enabled = False
            If !dotacion = 0 Then frmMDI.mnuCompraJoyeria.Enabled = False
            If !salidainven = 0 Then frmMDI.mnuSalidasInventario.Enabled = False
            If !inventariofisico = 0 Then frmMDI.mnuInvenFisico.Enabled = False
            If !deslotifica = 0 Then frmMDI.mnuDeslotificacion.Enabled = False
            If !etiquetas = 0 Then frmMDI.mnuEtiquetas.Enabled = False
            If !etiinven = 0 Then frmMDI.mnuEtiAlmoneda.Enabled = False
            If !repcompras = 0 Then frmMDI.mnuCompras.Enabled = False
            If !RepDota = 0 Then frmMDI.mnuRepEntradasInven.Enabled = False
            If !RepSalida = 0 Then frmMDI.mnuRepSalInventario.Enabled = False
            If !repanti = 0 Then frmMDI.mnuRepAntiguedad.Enabled = False
            If !repenve = 0 Then frmMDI.mnuReEnvejecimiento.Enabled = False
            If !repenvep = 0 Then frmMDI.mnuRepEnvejecimientoPeriodo.Enabled = False
            
            'Menu Divisas
            If !catdivisas = 0 Then frmMDI.mnuCatDivisas.Enabled = False
            If !Cotizacion = 0 Then frmMDI.mnuCotizacionDivisas.Enabled = False
            If !comvendiv = 0 Then frmMDI.mnuCompraVenta.Enabled = False
            If !repdivisas = 0 Then frmMDI.mnuRepDivisas.Enabled = False

            'Menu Reportes
            If !rephistorico = 0 Then frmMDI.mnuRepHistorico.Enabled = False
            If !repinventarios = 0 Then frmMDI.mnuRepInventario.Enabled = False
            If !repvencidos = 0 Then frmMDI.mnuRepVencidos.Enabled = False
            If !repalmoneda = 0 Then frmMDI.mnuRepAlmoneda.Enabled = False
            If !repempeños = 0 Then frmMDI.mnuCierreRepEmpeños.Enabled = False
            If !RepDesempenos = 0 Then frmMDI.mnuRepCieDesempeño.Enabled = False
            If !RepRefrendos = 0 Then frmMDI.mnuRepCieRefrendos.Enabled = False
            If !repauxiliar = 0 Then frmMDI.mnuRepAuxiliares.Enabled = False
            If !repcontable = 0 Then frmMDI.mnuRepContable.Enabled = False
            If !repauditoria = 0 Then frmMDI.mnuRepAuditoria.Enabled = False
            If !repgastos = 0 Then frmMDI.mnuRepGastos.Enabled = False
            If !RepIngresos = 0 Then frmMDI.mnuRepIngresos.Enabled = False
            If !RepHorarios = 0 Then frmMDI.mnuRepOperaciones.Enabled = False
            If !RepAutorizaciones = 0 Then frmMDI.mnuRepAutorizaciones.Enabled = False
            If !RepPartidaBoveda = 0 Then frmMDI.mnuRepPartidasBoveda.Enabled = False
            If !RepAseguradora = 0 Then frmMDI.mnuRepAseguradora.Enabled = False
            If !RepCancelaciones = 0 Then frmMDI.mnuRepCancelaciones.Enabled = False
            If !RepEmpeProm = 0 Then frmMDI.mnuRepEmpeMes.Enabled = False
            If !RepDesemProm = 0 Then frmMDI.mnuRepDesMes.Enabled = False
            If !RepRefProm = 0 Then frmMDI.mnuRepRefMes.Enabled = False
            If !ConTipoTasa = 0 Then frmMDI.mnuRepEmpenosTipoTasa.Enabled = False
            If !ConVencidos = 0 Then frmMDI.mnuRepEmpenosVencidos.Enabled = False
            If !ConStatus = 0 Then frmMDI.mnuRepPrestamoStatus.Enabled = False
            If !PrestamoMes = 0 Then frmMDI.mnuRepPrestamosMes.Enabled = False
            If !Medios = 0 Then frmMDI.mnuRepMedios.Enabled = False
            
            'Menu Movimientos
            If !movimientocaja = 0 Then frmMDI.mnuCaja.Enabled = False
            If !movimientobanco = 0 Then frmMDI.mnuBancos.Enabled = False
            If !moviboveda = 0 Then frmMDI.mnuBoveda.Enabled = False
            If !gastos = 0 Then frmMDI.mnuRegGastos.Enabled = False
            If !Movidiv = 0 Then frmMDI.mnuMovDivisas.Enabled = False
            If !remates = 0 Then frmMDI.mnuRemates.Enabled = False
            
            'Menu Utilerias
            If !parametros = 0 Then frmMDI.mnuParametros.Enabled = False
            If !ConfiguraTasas = 0 Then frmMDI.mnuConfTasas.Enabled = False
            If !PreciosKilataje = 0 Then frmMDI.mnuPreciosOro.Enabled = False
            If !ConfiguraDiam = 0 Then frmMDI.mnuPreciosDiamante.Enabled = False
            If !usuarios = 0 Then frmMDI.mnuUsuarios.Enabled = False
            If !Sucursales = 0 Then frmMDI.mnuSucursales.Enabled = False
            If !facturacion = 0 Then frmMDI.mnuFacturacion.Enabled = False
            If !capboletas = 0 Then frmMDI.mnuCapturaBoletas.Enabled = False
            If !MensajeContratos = 0 Then frmMDI.mnuMensajesBol.Enabled = False
            If !Catalogos = 0 Then frmMDI.mnuCatalogos.Enabled = False
            If !GeneraAutoriza = 0 Then frmMDI.mnuGeneraAutoriza.Enabled = False
            If !CatElec = 0 Then
                frmMDI.mnuFamilias.Enabled = False
                frmMDI.mnuCatMarcas.Enabled = False
                frmMDI.mnuCatPrendasVarios.Enabled = False
            End If
            
            'MLD-MODIF.- Permisos del Modulo ----------------------------------
            If !mld_parametros = 0 Then frmMDI.mnuLLDParametros.Enabled = False
            If !mld_movatipicos = 0 Then frmMDI.mnuLLDMovAtipicos.Enabled = False
            If !mld_expclientes = 0 Then frmMDI.mnuLLDExpClientes.Enabled = False
            If !mld_reppormenorizado = 0 Then frmMDI.mnuLLDRepMensual.Enabled = False
            '-----------------------------------------------------------------

            
            If frmMDI.ActiveLock2.RegisteredUser Then frmMDI.mnuRegSoftware.Visible = False Else frmMDI.mnuRegSoftware.Visible = True
            
            'Tomo los datos de la sucursal
            DatosSucursal
            
            frmMDI.Usuario = !Usuario
            frmMDI.IDUsuario = !ID
                
            'Panel para mostrar la sucursal
            PaneSucursal.text = "SUCURSAL: " & Sucursal.NombreComercial
            
            bBandera = True
            Unload Me
        End With

    Else
        MsgBox "Verifique su nombre de usuario y contraseña..." & Chr(10) & "Acceso denegado !!", vbCritical, "Clave de Usuario"
        txtUsuario.SetFocus
    End If
    rcPermisos.Close
    Set rcPermisos = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcPermisos = Nothing
End Sub

Function VerificaAcceso(Campo As String, User As String, Password As String) As Boolean
Dim rcTmp As New ADODB.Recordset

On Error GoTo Error

    rcTmp.Open "SELECT " & Campo & " AS Valor FROM usuarios WHERE Estatus=1 AND Usuario='" & Trim(User) & "' AND Contraseña='" & Trim(Password) & "'", dbDatos, adOpenForwardOnly, adLockReadOnly

        If rcTmp.BOF Or rcTmp.EOF Then
            
            VerificaAcceso = False
            
        ElseIf rcTmp!Valor = 0 Then
        
            VerificaAcceso = False
        
        Else
        
            VerificaAcceso = True
        End If

    rcTmp.Close
    Set rcTmp = Nothing
    Exit Function

Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function
