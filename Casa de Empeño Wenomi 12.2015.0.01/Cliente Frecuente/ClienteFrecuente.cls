VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClienteFrecuente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Error As String
Private m_CuentaFrecuente As New Cuenta
Private m_Usuario As String
Private m_Password As String

'Propiedades
Public Property Set Conexion(Value As ADODB.Connection)
   Set m_Conexion = Value
End Property

Public Property Get Error() As String
   Error = m_Error
End Property

Public Property Get CuentaFrecuente() As Cuenta
   Set CuentaFrecuente = m_CuentaFrecuente
End Property

'instala lo necesiario para que funcione el modulo
'tablas
'scripts
Public Sub Instalar(Usuario As String, Password As String)
   On Error Resume Next
   InstalarPuntos
End Sub

'acumula puntos en la cuenta
Public Function Acumula_Puntos(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long) As Long
   
    Dim Puntos As Long
   
    If m_CuentaFrecuente.IDCuenta > 0 Then   ' si hay una cuenta de puntos activa
   
        Select Case Movimiento
            Case TipoMovimiento.Empeno
                Puntos = m_CuentaFrecuente.Acumula_Empenos(Movimiento, IDUsuario, Importe, Folio, m_CuentaFrecuente.IDTarjeta, m_CuentaFrecuente.IDCuenta)
            Case TipoMovimiento.EmpenoAutos
                Puntos = m_CuentaFrecuente.Acumula_Empeno_Autos(Movimiento, IDUsuario, Importe, Folio, m_CuentaFrecuente.IDTarjeta)
             Case TipoMovimiento.Refrendo
                Puntos = m_CuentaFrecuente.Acumula_Refrendos(Movimiento, IDUsuario, Importe, Folio, m_CuentaFrecuente.IDTarjeta)
            Case TipoMovimiento.RefrendoExt
                Puntos = m_CuentaFrecuente.Acumula_Refrendos(Movimiento, IDUsuario, Importe, Folio, m_CuentaFrecuente.IDTarjeta)
             Case TipoMovimiento.Desempeno
                Puntos = m_CuentaFrecuente.Acumula_Desempeno(Movimiento, IDUsuario, Importe, Folio, m_CuentaFrecuente.IDTarjeta)
            Case TipoMovimiento.Ventas
                Puntos = m_CuentaFrecuente.Acumula_Ventas(Movimiento, IDUsuario, Importe, Folio, m_CuentaFrecuente.IDTarjeta)
            Case TipoMovimiento.Apartados
                Puntos = m_CuentaFrecuente.Acumula_Apartados(Movimiento, IDUsuario, Importe, Folio, m_CuentaFrecuente.IDTarjeta)
            Case TipoMovimiento.Abonos
                Puntos = m_CuentaFrecuente.Acumula_Abonos(Movimiento, IDUsuario, Importe, Folio, m_CuentaFrecuente.IDTarjeta)
          End Select
      
    Else
          Puntos = 0
    End If
   
    Acumula_Puntos = Puntos
End Function

Public Function GetCierrePuntos(FechaInicial As Date, FechaFinal As Date) As Long
    GetCierrePuntos = m_CuentaFrecuente.GetCierrePuntos(FechaInicial, FechaFinal)
End Function

Public Function GetCierrePuntosTotal(FechaInicial As Date, FechaFinal As Date) As Currency
    GetCierrePuntosTotal = m_CuentaFrecuente.GetCierrePuntosTotal(FechaInicial, FechaFinal)
End Function
'cancela los puntos de la cuenta
Public Function Cancela_Puntos() As Boolean

End Function

'cambiar puntos
Public Function Redimir_Puntos(Movimiento As TipoMovimiento, Puntos As Long, Importe As Currency, IDUsuario As Long, Folio As Long) As Boolean
    m_CuentaFrecuente.Redimir_Puntos Movimiento, Puntos, Importe, IDUsuario, Folio
End Function

'regresamos el importe en efectivo de puntos
Public Function GetImporte(Puntos As Long) As Currency
   Dim Importe As Currency
   
   Importe = Round(Val(Puntos / Val(SacaValor("Parametros", "PuntosTarjeta"))) * 1, 2)
   
   GetImporte = Importe
End Function

'parametros
Public Sub ShowParametros(hWndOwner As Long)
   SetParent frmParametros.hWnd, hWndOwner
   frmParametros.Show
End Sub

'tipos de tarjetas
Public Sub ShowTiposTarjetas(hWndOwner As Long)
   SetParent frmTiposTarjetas.hWnd, hWndOwner
   frmTiposTarjetas.Show
End Sub

'agregar tarjeta
Public Sub ShowAsignarTarjeta(IDCliente As Long, IDUsuario As Long)
   frmAsignarTarjeta.IDCliente = IDCliente
   frmAsignarTarjeta.IDUsuario = IDUsuario
   frmAsignarTarjeta.Show vbModal
End Sub

