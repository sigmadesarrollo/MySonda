VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Fecha As String
Private m_IDCuenta As Long
Private m_Folio As String
Private m_IDCliente As Long
Private m_Cliente As String
Private m_Puntos As Long
Private m_IDTarjeta As Long

Public Enum TipoMovimiento
   Empeno
   EmpenoAutos
   Refrendo
   RefrendoExt
   Desempeno
   Ventas
   Apartados
   Abonos
   EmpenoCancelacion
   EmpenoAutosCancelacion
   RefrendoCancelacion
   DesempenoCancelacion
   VentasCancelacion
   ApartadosCancelacion
   AbonosCancelacion
End Enum

'propiedades
Public Property Get FechaTarjeta() As String
    FechaTarjeta = m_Fecha
End Property

Public Property Get IDCuenta() As Long
   IDCuenta = m_IDCuenta
End Property

Public Property Get Folio() As String
   Folio = m_Folio
End Property

Public Property Get IDCliente() As Long
   IDCliente = m_IDCliente
End Property

Public Property Get Cliente() As String
   Cliente = m_Cliente
End Property

Public Property Get Puntos() As Long
   Puntos = m_Puntos
End Property

Public Property Get IDTarjeta() As Long
   IDTarjeta = m_IDTarjeta
End Property


'procedimientos
Public Sub Clear()
    m_Fecha = ""
    m_IDCuenta = 0
    m_Folio = 0
    m_IDCliente = 0
    m_Cliente = ""
    m_Puntos = 0
    m_IDTarjeta = 0
End Sub

'buscamos la cuenta por el folio
Public Function FindCuentaByFolio(Folio As String) As Boolean
   On Error GoTo Error
   Dim Sql As String
   Dim rc As New ADODB.Recordset
   
   Sql = "SELECT * FROM AsignacionTarjetas WHERE NumeroTarjeta='" & Folio & "'"
   
   rc.Open Sql, m_Conexion, adOpenForwardOnly, adLockOptimistic
   
   If rc.EOF Then
      Clear
      FindCuentaByFolio = False
   Else
      Poner_Datos rc
      FindCuentaByFolio = True
   End If
   
   rc.Close
Error:
   Maneja_Error Err
   
   Set rc = Nothing
End Function

'buscamos la cuenta por el id del cliente
Public Function FindCuentaByIDCliente(IDCliente As Long) As Boolean
On Error GoTo Error
   Dim Sql As String
   Dim rc As New ADODB.Recordset
   
   Sql = "SELECT * FROM AsignacionTarjetas WHERE IDCliente=" & IDCliente
   
   rc.Open Sql, m_Conexion, adOpenForwardOnly, adLockOptimistic
   
   If rc.EOF Then
      Clear
      FindCuentaByIDCliente = False
   Else
      Poner_Datos rc
      FindCuentaByIDCliente = True
   End If
   
   rc.Close
Error:
   Maneja_Error Err
   
   Set rc = Nothing
End Function

Public Function GetCierrePuntos(FechaInicial As Date, FechaFinal As Date) As Long
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Puntos As Long
    
    rc.Open "SELECT SUM(Abono) AS Total FROM MovimientosPuntos WHERE DATE(Fecha)>='" & Format(FechaInicial, "YYYY/MM/DD") & "' AND DATE(Fecha)<='" & Format(FechaFinal, "YYYY/MM/DD") & "'", m_Conexion, adOpenDynamic, adLockOptimistic
    
    If Not rc.EOF Then
        Puntos = rc!Total
    Else
        Puntos = 0
    End If
    
    GetCierrePuntos = Puntos
    
    rc.Close
Error:
    Maneja_Error Err
    
    Set rc = Nothing
End Function

Public Function GetCierrePuntosTotal(FechaInicial As Date, FechaFinal As Date) As Currency
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim Puntos As Currency
    
    rc.Open "SELECT SUM(Importe) AS Total FROM Auxiliar WHERE Cuenta='905501' AND DATE(Fecha)>='" & Format(FechaInicial, "YYYY/MM/DD") & "' AND DATE(Fecha)<='" & Format(FechaFinal, "YYYY/MM/DD") & "'", m_Conexion, adOpenDynamic, adLockOptimistic
    
    If Not rc.EOF Then
        Puntos = rc!Total
    Else
        Puntos = 0
    End If
    
    GetCierrePuntosTotal = Puntos
    
    rc.Close
Error:
    Maneja_Error Err
    
    Set rc = Nothing
End Function

Private Sub Poner_Datos(rc As ADODB.Recordset)
   On Error GoTo Error
   
   With rc
        m_Fecha = Format(!Fecha, "DD-MMM-YYYY")
        m_IDCuenta = !ID
        m_Folio = !NumeroTarjeta
        m_IDCliente = !IDCliente
        m_Cliente = ""
        m_Puntos = !Puntos
        m_IDTarjeta = !IDTarjeta
   End With
   
Error:
   Maneja_Error Err
End Sub

Public Function Acumula_Empenos(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long, IDTarjeta As Long, IDCuenta As Long) As Long
   On Error GoTo Error
   Dim Porcentaje As Double
   Dim Puntos As Long
   
   Porcentaje = GetPorcentaje(Movimiento, IDTarjeta)
   Puntos = CInt(Importe * Porcentaje)
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   
   m_Puntos = m_Puntos + Puntos
   
   Acumula_Empenos = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Empeno_Autos(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   Dim Porcentaje As Double
   Dim Puntos As Long
   
   Porcentaje = GetPorcentaje(Movimiento, IDTarjeta)
   Puntos = Importe * Porcentaje
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   
   m_Puntos = m_Puntos + Puntos
   
   Acumula_Empeno_Autos = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Refrendos(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   Dim Porcentaje As Double
   Dim Puntos As Long
   
   Porcentaje = GetPorcentaje(Movimiento, IDTarjeta)
   Puntos = Importe * Porcentaje
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   
   m_Puntos = m_Puntos + Puntos
   
   Acumula_Refrendos = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_RefrendosExt(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   
On Error GoTo Error
   
   Dim Porcentaje As Double
   Dim Puntos As Long
   
   Porcentaje = GetPorcentaje(Movimiento, IDTarjeta)
   Puntos = Importe * Porcentaje
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   
   m_Puntos = m_Puntos + Puntos
   
   Acumula_RefrendosExt = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Refrendos_Cancelacion(Movimiento As TipoMovimiento, IDUsuario As Long, Puntos As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   
   On Error GoTo Error
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   
   m_Puntos = m_Puntos + Puntos
   
   Acumula_Refrendos_Cancelacion = Puntos
   
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Desempeno(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   Dim Porcentaje As Double
   Dim Puntos As Long
   
   Porcentaje = GetPorcentaje(Movimiento, IDTarjeta)
   Puntos = Importe * Porcentaje
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   m_Puntos = m_Puntos + Puntos
   
   
   Acumula_Desempeno = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Desempeno_Cancelacion(Movimiento As TipoMovimiento, IDUsuario As Long, Puntos As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   
   m_Puntos = m_Puntos + Puntos
   
   
   Acumula_Desempeno_Cancelacion = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Ventas(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   Dim Porcentaje As Double
   Dim Puntos As Long
   
   Porcentaje = GetPorcentaje(Movimiento, IDTarjeta)
   Puntos = Importe * Porcentaje
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   m_Puntos = m_Puntos + Puntos
   
   
   Acumula_Ventas = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Ventas_Cancelacion(Movimiento As TipoMovimiento, IDUsuario As Long, Puntos As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   m_Puntos = m_Puntos + Puntos
   
   Acumula_Ventas_Cancelacion = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Apartados(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   Dim Porcentaje As Double
   Dim Puntos As Long
   
   Porcentaje = GetPorcentaje(Movimiento, IDTarjeta)
   Puntos = Importe * Porcentaje
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   m_Puntos = m_Puntos + Puntos
   
   
   Acumula_Apartados = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Apartados_Cancelacion(Movimiento As TipoMovimiento, IDUsuario As Long, Puntos As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   m_Puntos = m_Puntos + Puntos
   
   
   Acumula_Apartados_Cancelacion = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Abonos(Movimiento As TipoMovimiento, IDUsuario As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   Dim Porcentaje As Double
   Dim Puntos As Long
   
   Porcentaje = GetPorcentaje(Movimiento, IDTarjeta)
   Puntos = Importe * Porcentaje
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   m_Puntos = m_Puntos + Puntos
   
   Acumula_Abonos = Puntos
Error:
   Maneja_Error Err
End Function

Public Function Acumula_Abonos_Cancelacion(Movimiento As TipoMovimiento, IDUsuario As Long, Puntos As Long, Importe As Currency, Folio As Long, IDTarjeta As Long) As Long
   On Error GoTo Error
   
   Grabar_Movimiento Movimiento, Puntos, Importe, IDUsuario, Folio, IDCuenta
   m_Puntos = m_Puntos + Puntos
   
   Acumula_Abonos_Cancelacion = Puntos
Error:
   Maneja_Error Err
End Function

Public Sub Redimir_Puntos(Movimiento As TipoMovimiento, Puntos As Long, Importe As Currency, IDUsuario As Long, Folio As Long)
    On Error GoTo Error
    Dim Sql As String
    
    'actualizamos los puntos en la tarjeta
   Grabar_Movimiento Movimiento, (Puntos * -1), Importe, IDUsuario, Folio, IDCuenta
    
Error:
    Maneja_Error Err
End Sub


Private Function Grabar_Movimiento(Movimiento As TipoMovimiento, Puntos As Long, Importe As Currency, IDUsuario As Long, Folio As Long, IDCuenta As Long) As Boolean
   
   On Error GoTo Error
   
   Dim Sql As String
   Dim Cargo As Long
   Dim Abono As Long
   
   Cargo = 0
   Abono = 0
   
   'actualizamos los puntos en la tarjeta
   If Puntos > 0 Then
     Sql = "UPDATE AsignacionTarjetas SET Puntos=Puntos+" & Puntos & " WHERE ID=" & IDCuenta
     Cargo = Puntos
   Else
     Sql = "UPDATE AsignacionTarjetas SET Puntos=Puntos-" & (Puntos * -1) & " WHERE ID=" & IDCuenta
     Abono = Puntos * -1
   End If
   
   m_Conexion.Execute Sql
   
   Sql = "INSERT INTO MovimientosPuntos (Fecha,IDTarjeta,TipoMovimiento,Concepto,Folio,Cargo,Abono,Importe,PC,IDUsuario) VALUES ('" & _
         Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & IDCuenta & "," & Movimiento & ",''," & Folio & "," & Cargo & "," & Abono & "," & Importe & ",'" & Nombre_Pc & "'," & IDUsuario & ")"
        
   m_Conexion.Execute Sql
   
Error:
   Maneja_Error Err
End Function

Private Function GetPorcentaje(Movimiento As TipoMovimiento, IDTarjeta As Long) As Double
   On Error GoTo Error
   Dim Campo As String
   Dim Porcentaje As Double
   
    Select Case Movimiento
        Case TipoMovimiento.Empeno
             Campo = "pEmpeno"
        Case TipoMovimiento.EmpenoAutos
            Campo = "pEmpenoAutos"
        Case TipoMovimiento.Refrendo
             Campo = "pRefrendo"
        Case TipoMovimiento.RefrendoExt
             Campo = "pRefrendoExt"
            Case TipoMovimiento.Refrendo
            Campo = "pRefrendo"
        Case TipoMovimiento.Desempeno
            Campo = "pDesempeno"
        Case TipoMovimiento.Ventas
             Campo = "pVentas"
        Case TipoMovimiento.Apartados
             Campo = "pApartados"
        Case TipoMovimiento.Abonos
             Campo = "pAbonos"
    End Select
      
    Porcentaje = SacaValor("TarjetasPuntos", Campo, " WHERE ID=" & IDTarjeta)
    GetPorcentaje = (Porcentaje / 100)
Error:
   Maneja_Error Err
End Function



