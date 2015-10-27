Attribute VB_Name = "mdlDeclaraciones"
Option Explicit
'
'Public Enum TipoMovimiento
'   Empeno
'   EmpenoAutos
'   Refrendo
'   Desempeno
'   Ventas
'   Apartados
'   Abonos
'End Enum

Public Enum TipoParametro
   PorcentajeEmpeno
   Porcentajerefrendo
   Porcentajedesempeno
   Porcentajeventa
   ValorPuntos
End Enum

Public Const MAX_COMPUTERNAME_LENGTH As Long = 31


Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Public m_Conexion As New ADODB.Connection
