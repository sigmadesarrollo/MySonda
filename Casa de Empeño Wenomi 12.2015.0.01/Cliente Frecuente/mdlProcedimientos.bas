Attribute VB_Name = "mdlProcedimientos"
Option Explicit


Public Sub InstalarPuntos()
   On Error Resume Next
   Dim Sql As String
   
   
   Sql = "CREATE TABLE `basedatos`.`TarjetasPuntos` (" & _
         "`ID` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT," & _
         "`TipoTarjeta` VARCHAR(60) NOT NULL," & _
         "`pEmpeno` DOUBLE(15,5) NOT NULL DEFAULT '0.00000'," & _
         "`pEmpenoAutos` DOUBLE(15,5) NOT NULL DEFAULT '0.00000'," & _
         "`pRefrendo` DOUBLE(15,5) NOT NULL DEFAULT '0.00000'," & _
         "`pRefrendoExt` DOUBLE(15,5) NOT NULL DEFAULT '0.00000'," & _
         "`pDesempeno` DOUBLE(15,5) NOT NULL DEFAULT '0.00000'," & _
         "`pVentas` DOUBLE(15,5) NOT NULL DEFAULT '0.00000'," & _
         "`pApartados` DOUBLE(15,5) NOT NULL DEFAULT '0.00000'," & _
         "`pAbonos` DOUBLE(15,5) NOT NULL DEFAULT '0.00000'," & _
         "`FechaCreacion` DATETIME NOT NULL," & _
         "`Activa` INT NOT NULL DEFAULT 1," & _
         "PRIMARY KEY (`ID`)" & _
         ")" & _
         "ENGINE = MyISAM;"
   
   m_Conexion.Execute Sql
   
   Sql = "ALTER TABLE `basedatos`.`parametros` ADD COLUMN `PuntosTarjeta` INTEGER UNSIGNED DEFAULT 0;"
   
   m_Conexion.Execute Sql
   
   Sql = "CREATE TABLE `basedatos`.`AsignacionTarjetas` (" & _
         "`ID` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT," & _
         "`Fecha` DATETIME NOT NULL DEFAULT 0," & _
         "`NumeroTarjeta` VARCHAR(60) DEFAULT NULL," & _
         "`IDTarjeta` INT UNSIGNED NOT NULL DEFAULT 0," & _
         "`IDCliente` INT UNSIGNED NOT NULL DEFAULT 0," & _
         "`IDUsuario` INT UNSIGNED NOT NULL DEFAULT 0," & _
         "`PC` VARCHAR(60) DEFAULT NULL," & _
         "`Puntos` INT UNSIGNED NOT NULL DEFAULT 0," & _
         "PRIMARY KEY (`ID`)" & _
         ")" & _
         "ENGINE = MyISAM;"


   m_Conexion.Execute Sql
   
   
   Sql = "CREATE TABLE `basedatos`.`MovimientosPuntos` (" & _
         "`ID` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT," & _
         "`Fecha` DATETIME NOT NULL," & _
         "`IDTarjeta` INTEGER UNSIGNED NOT NULL DEFAULT 0," & _
         "`TipoMovimiento` INTEGER UNSIGNED NOT NULL DEFAULT 0," & _
         "`Concepto` VARCHAR(80) NOT NULL," & _
         "`Folio` INTEGER UNSIGNED NOT NULL DEFAULT 0," & _
         "`Cargo` DECIMAL(14,4) NOT NULL DEFAULT 0," & _
         "`Abono` DECIMAL(14,4) NOT NULL DEFAULT 0," & _
         "`Importe` DECIMAL(14,4) NOT NULL DEFAULT 0," & _
         "`PC` VARCHAR(45) NOT NULL," & _
         "`IDUsuario` INTEGER UNSIGNED NOT NULL DEFAULT 0," & _
         "PRIMARY KEY (`ID`)" & _
         ")" & _
         "ENGINE = MyISAM;"
         
   m_Conexion.Execute Sql
End Sub

'Cambiamos el color del control cuando se posiciona en el
'Si opcion es true cambiamos el color de lo contrario lo ponemos blanco
Public Sub Cambiar_Color(Valor As Boolean, obj As Object)

On Error Resume Next
  
    If Valor Then
        
        obj.BackColor = RGB(250, 248, 180)
    Else
        
        obj.BackColor = vbWindowBackground 'RGB(255, 255, 255)
    End If
End Sub


'Remarcamos cuando el texto seleccionado
Public Sub Seleccionar_Texto(obj As Object)

On Error Resume Next
  
    obj.SelStart = 0
    obj.SelLength = Len(obj.Text)

End Sub

'Mensaje estandar para todos los errores
Public Sub Maneja_Error(Error As ErrObject)
    
    If Error.Number <> 0 And Error <> 383 Then
        
        MsgBox Error.Source & " " & Error.Description, vbOKOnly + vbCritical
    
    End If

End Sub

'Verificamos si la tecla presionada fue enter
'y pasamos el foco al siguiente control
Public Sub Pasar_Foco(ByRef Codigo As Integer)

    If Codigo = vbKeyReturn Then
        
        Codigo = 0
        SendKeys "{Tab}"
    
    End If

End Sub


'Nos permite aceptar solo numeros a una cantidad con decimales
Public Function Solo_Numeros(Codigo As Integer, Optional Opcion As Integer = 0) As Integer
    
    If (Codigo >= vbKey0 And Codigo <= vbKey9) Or Codigo = vbKeyBack Or Codigo = vbKeyReturn Or (Opcion = 1 And (Codigo = Asc("-"))) Then
        
        Solo_Numeros = Codigo
        
    Else
        
        Solo_Numeros = 0
    
    End If
End Function


Function SacaValor(Tabla As String, Campo As String, Optional Condicion As String = "") As String
Dim rcValor As New ADODB.Recordset

On Error GoTo Error
    
    rcValor.Open "SELECT " & Campo & " AS Valor FROM " & Tabla & Condicion, m_Conexion, adOpenForwardOnly, adLockReadOnly
    If Not rcValor.BOF And Not rcValor.EOF And Not IsNull(rcValor.Fields("Valor")) Then
        SacaValor = rcValor.Fields("Valor")
    Else
        
        SacaValor = ""
    End If
    
    rcValor.Close
    
Error:
    Maneja_Error Err
    Set rcValor = Nothing
End Function

'Regresamos el nombre de la computadora
Public Function Nombre_Pc() As String
Dim dwLen As Long, strString As String
   
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    dwLen = 31 + 1
    strString = String(dwLen, "X")
    
    'Get the computer name
    GetComputerName strString, dwLen
    
    'get only the actual data
    strString = Left(strString, dwLen)
    
    'Show the computer name
    Nombre_Pc = strString
    
End Function
