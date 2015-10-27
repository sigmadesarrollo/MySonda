Attribute VB_Name = "mdlProcedimientos"
Option Explicit

'////////////////////////////////////////////////////////////////
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 02/04/02
' Modulo mdlProcedimientos - mdlProcedimientos.bas
' Ultima Modificacion - 15/08/02 - L.I. Jorge Gabriel Colio Ramos
' Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

'Regresamos el total de abonos de la venta
Public Function Regresa_Abonos(ID As Long) As Currency
Dim rcAbonos As New ADODB.Recordset
   
On Error GoTo Error
    
    Regresa_Abonos = 0
   
    rcAbonos.Open "SELECT SUM(Importe) AS Total FROM abonos WHERE IDVenta=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
   
    If Not rcAbonos.EOF Then Regresa_Abonos = IIf(IsNull(rcAbonos!Total), 0, rcAbonos!Total)
   
    rcAbonos.Close

Error:
    Maneja_Error Err
    Set rcAbonos = Nothing
End Function

Public Function OD_Origen(Valor As Long) As String
Dim Cadena As String
   
    Select Case Valor

    Case OD_EMPENO
        Cadena = "Empeño"

    Case OD_REFRENDO
        Cadena = "Refrendo"

    Case D_DESEMPEÑO
        Cadena = "Desempeño"

    Case D_ALMONEDA
        Cadena = "Almoneda"

    Case D_VENTA
        Cadena = "Venta"

    Case D_FUNDICION
        Cadena = "Fundición"

    Case D_OTRO
        Cadena = "Otro"
    End Select
   
    OD_Origen = Cadena
End Function

 Public Sub Main()
    Inicializar
End Sub

Private Sub Inicializar()
 
'***Puntos***
Dim Puntos As New ClienteFrecuente

On Error GoTo Error
        
    If ChecaCandado Then
'            frmSplash.Show
        Path = Regresa_Valor("Montepio", "Path", App.Path)
        sServidor = Trim(Regresa_Valor("MONTEPIO", "Servidor", "localhost"))
        
        dbDatos.Open cCONEXION & sServidor & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDB
        dbReportes.Open cCONEXION & sServidor & "; PORT=" & Regresa_Valor("MONTEPIO", "Puerto", "3306") & cDBR
        Cargar_WebService
'        Actualiza_Sistema
        'MsgBox "Revisa Version"
        If Regresa_Valor("Software", "Version", "") <> App.Major & "." & App.Minor & "." & App.Revision Then
            'frmSplash.lblMensaje.Caption = "Actualizando base de datos..."
            DoEvents
            'MsgBox "Crea Campos Inicia"
            Crear_Campos
            'MsgBox "Crea Campos Finaliza"
            
            Graba_Valor "Software", "Version", App.Major & "." & App.Minor & "." & App.Revision
        End If
                        
        'MsgBox "Revisa Movimientos Inicia"
        Checar_Movimientos
        'MsgBox "Revisa Movimientos Finaliza"
                        
        'MLD-MODIF +++ CREACION DE CAMPOS PARA MODULO DE LAVADO DE DINERO +++
        'MsgBox "Revisa Antilavado Inicia"
        Crear_Campos_LavadoDinero
        'MsgBox "Revisa Antilavado Finaliza"
                                        
'        Forzar_Actualizacion
'         Unload frmSplash
        '***Puntos***
        'MsgBox "Puntos Inicia"
        Set Puntos.CONEXION = dbDatos
        Puntos.Instalar "mrayudon", "montepio"
        'MsgBox "Puntos Finaliza"
                        
        Separador = Obtener_Separador_Decimal
        If frmPasswords.Password Then
            NombrePc = Nombre_Pc
            frmMDI.Caption = "Casa de Empeño - " & "Usuario Actual: " & SacaValor("Usuarios", "Nombre", " where ID=" & frmMDI.IDUsuario) & " - Ver." & App.Major & "." & App.Minor & "." & App.Revision
            frmMDI.Show
        Else
            frmMDI.tmrHora.Interval = 0
            Set dbDatos = Nothing
            Set dbReportes = Nothing
        End If
   
    Else

        frmMDI.tmrHora.Interval = 0
        Set dbDatos = Nothing
        Set dbReportes = Nothing

        MsgBox "El periodo de prueba de MySonda ha expirado... !!" & Chr(13) & "Introduzca una clave de activación válida !!", vbExclamation, "Activación MySonda"
        frmRegistrar.Show vbModal
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
    End
End Sub

'Creo los campos
Private Sub Crear_Campos()

On Error Resume Next
''''''' 2014-09-17
dbReportes.Execute "CREATE TABLE `basereportes`.`repVencidosDetalle` ( " & _
  " `ID` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT, " & _
  " `IDEmpeno` INTEGER UNSIGNED, " & _
  " `articulo` VARCHAR(100), " & _
  " `peso` VARCHAR(45), " & _
  " `kilates` VARCHAR(45), " & _
  " `prestamo` DOUBLE(15,2), " & _
  " `marca` VARCHAR(45), " & _
  " `modelo` VARCHAR(45), " & _
  " PRIMARY KEY (`ID`) " & _
" ) " & _
" ENGINE = InnoDB; "

dbReportes.Execute "ALTER TABLE `basereportes`.`repvencidos` MODIFY COLUMN `Celular` VARCHAR(80) CHARACTER SET latin1 COLLATE latin1_swedish_ci, " & _
 " ADD COLUMN `fechaComercializacion` DATE AFTER `Celular`;"
                
    dbDatos.Execute "ALTER TABLE clientes ADD COLUMN NumeroIdentificacion VARCHAR(30) AFTER Identificacion"
    dbDatos.Execute "ALTER TABLE clientes MODIFY COLUMN Tel VARCHAR(50)"
    
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN Cat Double(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN IntAnual Double(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN AlmAnual Double(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN PorEnajenados Double(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN PrestamoVerde Double(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN PrestamoAmarillo Double(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN PrestamoRojo Double(15,5) Default 0"
    
    dbDatos.Execute "ALTER TABLE pagosfijos ADD COLUMN Iva DOUBLE(15,5) DEFAULT 0 AFTER Seguro"
    
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN Beneficiario VARCHAR(90) AFTER Responsable"
    dbDatos.Execute "ALTER TABLE auxiliar ADD COLUMN IDDivisa INT(10) Default 0"
    dbDatos.Execute "ALTER TABLE auxiliar ADD COLUMN Hora Time AFTER Fecha"
    dbDatos.Execute "ALTER TABLE movimientos ADD COLUMN FolioBovedaDivisas INT(10) Default 1"
    
    dbReportes.Execute "CREATE TABLE cortedivisas(ID INTEGER UNSIGNED NOT NULL AUTO_INCREMENT,IDDivisa INTEGER(11) DEFAULT 0,Dotacion INTEGER(11) DEFAULT 0,Retiro INTEGER(11) DEFAULT 0,Compras INTEGER(11) DEFAULT 0,Ventas INTEGER(11) DEFAULT 0,PC VARCHAR(30), PRIMARY KEY (ID)) ENGINE = MyISAM"
    dbDatos.Execute "ALTER TABLE configuraciontasas ADD COLUMN PorPrestamo DOUBLE(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE configuraciontasas ADD COLUMN Cat DOUBLE(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE configuraciontasas ADD COLUMN Almacenaje DOUBLE(15,5) Default 0"
    dbDatos.Execute "ALTER TABLE configuraciontasas ADD COLUMN Seguro DOUBLE(15,5) Default 0"
    
    dbDatos.Execute "ALTER TABLE autorizaciones MODIFY Fecha DATETIME"
    
    dbDatos.Execute "ALTER TABLE salidainventario MODIFY Fecha DATETIME"
    dbDatos.Execute "ALTER TABLE salidainventario ADD COLUMN TipoSalida INTEGER Default 0 AFTER Folio"
    dbDatos.Execute "ALTER TABLE salidainventario ADD COLUMN Pagado INTEGER Default 0 AFTER Fecha"
    dbDatos.Execute "ALTER TABLE detallessalida ADD COLUMN Observaciones VARCHAR(250)"
    
    dbDatos.Execute "ALTER TABLE auxiliar MODIFY Hora TIME AFTER Fecha"
    dbDatos.Execute "ALTER TABLE usuarios ADD COLUMN GeneraAutoriza INT(1) Default 0"
    dbDatos.Execute "ALTER TABLE usuarios ADD COLUMN CatElec INT(1) Default 0"
    dbDatos.Execute "ALTER TABLE usuarios ADD COLUMN RefrendarVencidos INT(1) Default 0"
    dbDatos.Execute "ALTER TABLE usuarios ADD COLUMN CancelaCierre INT(1) Default 0"
    
    dbReportes.Execute "ALTER TABLE opcionpagos ADD COLUMN PC VARCHAR(30)"
    dbReportes.Execute "CREATE TABLE repvencidos (IDEmpeno INTEGER DEFAULT 0,NumContrato INT(10) DEFAULT 0,Fecha DATETIME NOT NULL,Vencimiento DATE NOT NULL,Cliente VARCHAR(90),Avaluo DOUBLE(15,5) DEFAULT 0,Prestamo DOUBLE(15,5) DEFAULT 0,Serie INT(10) DEFAULT 0,TipoInteres VARCHAR(20),TipoTasa VARCHAR(20),FechaMovimiento DATE Default NULL,PRIMARY KEY (IDEmpeno)) ENGINE = MyISAM"
    dbReportes.Execute "ALTER TABLE repvencidos ADD COLUMN Tel VARCHAR(80)"
    
    dbDatos.Execute "ALTER TABLE compras ADD COLUMN FechaMovimiento DATETIME Default NULL"
    dbDatos.Execute "ALTER TABLE cierrediario MODIFY Sucursal VARCHAR(60) IDSucursal INT(10) Default 0"
    
    dbDatos.Execute "CREATE TABLE rematediario (ID INT(10) UNSIGNED NOT NULL AUTO_INCREMENT, Fecha DATETIME DEFAULT NULL, Status INT(1) DEFAULT 0, PRIMARY KEY (ID)) ENGINE=MyISAM"

    'Vistas y procedimientos
    dbDatos.Execute "CREATE OR REPLACE VIEW vwapartadosrematados AS SELECT SUM(abonos.Importe) AS Abonos,ventas.ID,ventas.Fecha,ventas.FechaMovimiento,ventas.Folio," _
                    & "ventas.IVA,ventas.Vencimiento,ventas.Total,ventas.Descuento,ventas.Pagado,ventas.Cancelado,ventas.OrigenCancelacion,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente " _
                    & "FROM ventas LEFT JOIN abonos ON ventas.ID=abonos.IDVenta LEFT JOIN clientes ON ventas.IDCliente=clientes.ID WHERE ventas.OrigenCancelacion=2 AND ventas.Apartado=1 AND ventas.Cancelado=1 AND abonos.Cancelado=1 GROUP BY ventas.ID"
    
    dbDatos.Execute "CREATE OR REPLACE VIEW vwdetallesempeno AS SELECT detallesempeno.ID,detallesempeno.IDEmpeno,detallesempeno.Cantidad,detallesempeno.Tipo,detallesempeno.Articulo," _
                    & "detallesempeno.Peso AS PesoTotal,detallesempeno.PesoPiedras,(detallesempeno.Peso-detallesempeno.PesoPiedras) AS PesoReal,detallesempeno.Prestamo,detallesempeno.Avaluo,detallesempeno.Observaciones,detallesempeno.Estado,detallesempeno.Marca,detallesempeno.Modelo,detallesempeno.Serie," _
                    & "detallesempeno.Tamano,detallesempeno.Color,tipo.Descripcion AS Tipo_DESC,kilatajes.Descripcion AS Kil_DESC FROM detallesempeno LEFT JOIN tipo ON detallesempeno.Tipo=tipo.ID LEFT JOIN kilatajes ON detallesempeno.Kilates=kilatajes.Clave"
    
    dbDatos.Execute "CREATE OR REPLACE VIEW vwpagosfijos AS SELECT pagosfijos.IDEmpeno,MAX(pagosfijos.FechaMovimiento) AS FechaMovimiento FROM pagosfijos " _
                    & "WHERE pagosfijos.Pagado=1 AND pagosfijos.Cancelado=0 GROUP BY pagosfijos.IDEmpeno ORDER BY pagosfijos.IDEmpeno,pagosfijos.ID"
    
    dbDatos.Execute "CREATE OR REPLACE VIEW vwrepapartados AS SELECT SUM(abonos.Importe) AS Abonos,ventas.ID,ventas.Fecha,ventas.Folio,ventas.IVA,ventas.Vencimiento,ventas.Total,ventas.Descuento," _
                    & "ventas.Pagado,ventas.Cancelado,ventas.OrigenCancelacion,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente FROM ventas LEFT JOIN abonos ON ventas.ID=abonos.IDVenta LEFT JOIN " _
                    & "clientes ON ventas.IDCliente=clientes.ID WHERE ventas.Apartado=1 AND abonos.Cancelado=0 GROUP BY ventas.ID"
    
    dbDatos.Execute "CREATE OR REPLACE VIEW vwfacturadiaria AS SELECT COUNT(a.ID) AS NumRegistros,a.Fecha,SUM(a.Importe) AS ImporteTotal FROM auxiliar a WHERE a.Importe>0 AND (a.Cuenta='520450' OR a.Cuenta='670350' OR a.Cuenta='680350' OR a.Cuenta='690350') GROUP BY a.Fecha ORDER BY a.Fecha"
    
    dbDatos.Execute "CREATE OR REPLACE VIEW vwfacturaventas AS SELECT COUNT(a.ID) AS NumRegistros,a.Fecha,SUM(a.Importe) AS ImporteTotal FROM auxiliar a WHERE a.Importe>0 AND a.Cuenta='620450' AND Concepto='Ventas' GROUP BY a.Fecha ORDER BY a.Fecha"
    
    dbDatos.Execute "CREATE PROCEDURE spRepComisiones(IN FechaIni DATE, IN FechaFin DATE) BEGIN " _
                    & "SELECT SUM(abonos.Importe) AS Abonos,ventas.ID AS IDVenta,ventas.Folio,abonos.Fecha AS FechaAbono, " _
                    & "ventas.IVA,ventas.Descuento,ventas.Vencimiento,ventas.Total,ventas.Pagado,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente, " _
                    & "CONCAT(vendedores.Nombre,' ',vendedores.Apellidos) AS Vendedor FROM ventas left join abonos ON " _
                    & "ventas.ID=abonos.IDVenta INNER JOIN clientes ON ventas.IDCliente=clientes.ID LEFT JOIN vendedores " _
                    & "ON ventas.IDVendedor=vendedores.ID WHERE DATE_FORMAT(abonos.Fecha,'%Y%-%m%-%d') BETWEEN FechaIni AND FechaFin " _
                    & "AND abonos.Cancelado=0 AND ventas.Cancelado=0 AND ventas.Apartado=1 GROUP BY ventas.ID; End"

'    dbDatos.Execute "CREATE PROCEDURE spRepVencidos(IN FechaIni DATE, IN FechaFin DATE, IN DiasEnajena INTEGER, IN TipoContrato INTEGER, IN TipoPrenda INTEGER)" _
'                    & "BEGIN SELECT DISTINCT e.ID,e.NumContrato,e.Fecha,e.Vencimiento,CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,c.Tel,e.Avaluo,e.Serie," _
'                    & "e.Prestamo,e.TipoInteres,e.TipoTasa FROM empeno e INNER JOIN clientes c ON e.IDCliente=c.ID " _
'                    & "LEFT JOIN detallesempeno de ON e.ID=de.IDEmpeno WHERE " _
'                    & "if(TipoContrato=1,(e.Serie=1 OR e.Serie=2 OR e.Serie=3),if(TipoContrato=3,e.Serie=2,de.Tipo=TipoPrenda)) " _
'                    & "AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0,DiasEnajena) DAY),'%Y%/%m%/%d') " _
'                    & "BETWEEN FechaIni AND FechaFin AND e.Cancelado=0 AND e.Destino=0 AND e.Pagado=0 ORDER BY NumContrato; END"
'*************
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN CodProfeco VARCHAR(80) DEFAULT 'EN TRÁMITE' AFTER AlmAnual"
    dbDatos.Execute "ALTER TABLE sucursales ADD COLUMN HorarioSucursal VARCHAR(250)"
    dbDatos.Execute "ALTER TABLE sucursales ADD COLUMN DomicilioAclaraciones VARCHAR(200) AFTER Cp"
    dbDatos.Execute "ALTER TABLE sucursales ADD COLUMN TelefonoAclaraciones VARCHAR(50) AFTER DomicilioAclaraciones"
    dbDatos.Execute "ALTER TABLE sucursales ADD COLUMN CorreoAclaraciones VARCHAR(50) AFTER TelefonoAclaraciones"
    dbDatos.Execute "ALTER TABLE sucursales ADD COLUMN ContratoRegistrado VARCHAR(50) AFTER CorreoAclaraciones"
    dbDatos.Execute "ALTER TABLE sucursales ADD COLUMN FechaContratoRegistrado DATE AFTER ContratoRegistrado"
    
    dbDatos.Execute "ALTER TABLE sucursales ADD COLUMN NombreSucursal VARCHAR(100) AFTER Clave"
'*************
    dbDatos.Execute "UPDATE parametros SET DescuentoPagosFijos = 0"
    
    '***Puntos***
    dbReportes.Execute "CREATE TABLE estadocuentapuntos (" & _
        "ID INT(11) UNSIGNED NOT NULL AUTO_INCREMENT," & _
        "Fecha DATE NOT NULL," & _
        "Folio INT(11) DEFAULT 0," & _
        "Movimiento VARCHAR(45)," & _
        "Cargo DOUBLE(15,5) DEFAULT '0.00000'," & _
        "Abono DOUBLE(15,5) DEFAULT '0.00000'," & _
        "Saldo DOUBLE(15,5) DEFAULT '0.00000'," & _
        "IDCliente INT(10) DEFAULT 0," & _
        "PRIMARY KEY (ID)) ENGINE = MyISAM"

    dbDatos.Execute "ALTER TABLE ventas ADD COLUMN DescuentoEfectivo DOUBLE(15,5) DEFAULT '0.00000'"
    dbDatos.Execute "ALTER TABLE ventas ADD COLUMN DescuentoXPuntos DOUBLE(15,5) DEFAULT '0.00000' AFTER DescuentoEfectivo"
    dbDatos.Execute "ALTER TABLE ventas ADD COLUMN SaldoPuntosAnterior DOUBLE(15,5) DEFAULT '0.00000' AFTER DescuentoXPuntos"
    dbDatos.Execute "ALTER TABLE ventas ADD COLUMN PuntosUsados DOUBLE(15,5) DEFAULT '0.00000' AFTER SaldoPuntosAnterior"
    dbDatos.Execute "ALTER TABLE ventas ADD COLUMN PuntosAcumulados DOUBLE(15,5) DEFAULT '0.00000' AFTER PuntosUsados"
    dbDatos.Execute "ALTER TABLE ventas ADD COLUMN SaldoPuntosActual DOUBLE(15,5) DEFAULT '0.00000' AFTER PuntosAcumulados"
    dbDatos.Execute "ALTER TABLE ventas ADD COLUMN IDTarjeta INT(10) DEFAULT 0 AFTER SaldoPuntosActual"

    dbDatos.Execute "ALTER TABLE detallesventas ADD COLUMN ImporteDescuento DOUBLE(15,5) DEFAULT '0.00000'"

    dbDatos.Execute "ALTER TABLE abonos ADD COLUMN DescuentoXPuntos DOUBLE(15,5) DEFAULT '0.00000'"
    dbDatos.Execute "ALTER TABLE abonos ADD COLUMN SaldoPuntosAnterior DOUBLE(15,5) DEFAULT '0.00000' AFTER DescuentoXPuntos"
    dbDatos.Execute "ALTER TABLE abonos ADD COLUMN PuntosUsados DOUBLE(15,5) DEFAULT '0.00000' AFTER SaldoPuntosAnterior"
    dbDatos.Execute "ALTER TABLE abonos ADD COLUMN PuntosAcumulados DOUBLE(15,5) DEFAULT '0.00000' AFTER PuntosUsados"
    dbDatos.Execute "ALTER TABLE abonos ADD COLUMN SaldoPuntosActual DOUBLE(15,5) DEFAULT '0.00000' AFTER PuntosAcumulados"
    dbDatos.Execute "ALTER TABLE abonos ADD COLUMN IDTarjeta INT(10) DEFAULT 0 AFTER SaldoPuntosActual"
    
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN SaldoPuntosAnteriorEmp DOUBLE(15,5) DEFAULT '0.00000'"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN PuntosAcumuladosEmp DOUBLE(15,5) DEFAULT '0.00000' AFTER SaldoPuntosAnteriorEmp"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN SaldoPuntosActualEmp DOUBLE(15,5) DEFAULT '0.00000' AFTER PuntosAcumuladosEmp"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN IDTarjetaEmp INT(10) DEFAULT 0 AFTER SaldoPuntosActualEmp"
    
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN DescuentoXPuntos DOUBLE(15,5) DEFAULT '0.00000' AFTER IDTarjetaEmp"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN SaldoPuntosAnterior DOUBLE(15,5) DEFAULT '0.00000' AFTER DescuentoXPuntos"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN PuntosUsados DOUBLE(15,5) DEFAULT '0.00000' AFTER SaldoPuntosAnterior"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN PuntosAcumulados DOUBLE(15,5) DEFAULT '0.00000' AFTER PuntosUsados"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN SaldoPuntosActual DOUBLE(15,5) DEFAULT '0.00000' AFTER PuntosAcumulados"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN IDTarjeta INT(10) DEFAULT 0 AFTER SaldoPuntosActual"
    
    dbDatos.Execute "ALTER TABLE clientes ADD COLUMN Celular VARCHAR(50) AFTER Tel"
    dbDatos.Execute "ALTER TABLE clientes ADD COLUMN CorreoElectronico VARCHAR(100) AFTER Celular"
    
    dbDatos.Execute "ALTER TABLE sucursales ADD COLUMN CodProfeco VARCHAR(50) AFTER ContratoRegistrado"
    
    dbReportes.Execute "ALTER TABLE repvencidos ADD COLUMN Celular VARCHAR(80) AFTER Tel"
    
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN HorarioSucursal VARCHAR(250)"
    
    dbDatos.Execute "ALTER TABLE parametros ADD COLUMN Version VARCHAR(20) DEFAULT NULL"
    
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN FechaOriginal DATETIME DEFAULT NULL AFTER Fecha;"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN IDTipoPrenda INT(10) DEFAULT 0;"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN IDEmpenoOrigen INT(11) DEFAULT 0;"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN IDEmpenoDestino INT(11) DEFAULT 0;"
    dbDatos.Execute "ALTER TABLE empeno ADD COLUMN Cat DOUBLE(15,5) DEFAULT '0.00000' AFTER Iva;"
    dbDatos.Execute "ALTER TABLE detallesentradainventario ADD COLUMN IDDetallesEmpeno INT(11) DEFAULT 0;"
    
    dbDatos.Execute "ALTER TABLE movimientos ADD COLUMN FolioDevolucion INT(10) Default 1"
    dbDatos.Execute "ALTER TABLE movimientos ADD COLUMN FolioPasesInventario INT(10) Default 1"
    dbDatos.Execute "ALTER TABLE ventas ADD COLUMN Devolucion TINYINT(1) Default 0"
    dbDatos.Execute "ALTER TABLE detallesventas ADD COLUMN Devolucion TINYINT(1) Default 0"
    
'''''    dbDatos.Execute "DROP TRIGGER `garantiasdevoluciones_insert`;"
'''''    dbDatos.Execute "DROP TRIGGER `garantiasdevoluciones_update`;"
'''''    dbDatos.Execute "DROP TRIGGER `pasesinventarios_insert`;"
'''''    dbDatos.Execute "DROP TRIGGER `pasesinventarios_update`;"
    
'''''    dbDatos.Execute "DELIMITER $$ CREATE DEFINER=`root`@`localhost` TRIGGER `garantiasdevoluciones_insert` BEFORE INSERT ON `garantiasdevoluciones` FOR EACH ROW BEGIN " & _
'''''                    "  DECLARE IDSucursal       INT; " & _
'''''                    "  DECLARE NewID            BIGINT(21); " & _
'''''                    "  DECLARE Actualiza        TINYINT(1) DEFAULT 0; " & _
'''''                    "  SELECT AUTO_INCREMENT INTO NewID FROM information_schema.TABLES WHERE TABLE_SCHEMA = database() AND TABLE_NAME = 'garantiasdevoluciones'; " & _
'''''                    "  SELECT Clave INTO IDSucursal FROM Sucursales WHERE Activa = 1 LIMIT 1; " & _
'''''                    "  IF NEW.ACTUALIZAR = 1 THEN " & _
'''''                    "    SET Actualiza = 0; " & _
'''''                    "  Else " & _
'''''                    "    SET Actualiza = 1; " & _
'''''                    "  END IF; " & _
'''''                    "  If Actualiza = 1 Then " & _
'''''                    "    INSERT INTO Actualizaciones (IDOrigen, Tabla, Sucursal) VALUES (NewID, 'garantiasdevoluciones', IDSucursal); " & _
'''''                    "  END IF; " & _
'''''                    "  SET New.Actualizar = 0; " & _
'''''                    "END $$ DELIMITER ;"
'''''    dbDatos.Execute "CREATE DEFINER=`root`@`localhost` TRIGGER `garantiasdevoluciones_update` BEFORE UPDATE ON `garantiasdevoluciones` FOR EACH ROW BEGIN " & _
'''''                    "  DECLARE IDSucursal INT; " & _
'''''                    "  DECLARE IDExiste   INT; " & _
'''''                    "  DECLARE Actualiza  TINYINT(1) DEFAULT 0; " & _
'''''                    "  SELECT Clave INTO IDSucursal FROM Sucursales WHERE Activa = 1 LIMIT 1; " & _
'''''                    "  SELECT count(IDOrigen) INTO IDExiste FROM Actualizaciones WHERE Tabla = 'garantiasdevoluciones' AND IDOrigen = New.ID; " & _
'''''                    "  IF NEW.ACTUALIZAR = 1 THEN " & _
'''''                    "    SET Actualiza = 0; " & _
'''''                    "  Else " & _
'''''                    "    SET Actualiza = 1; " & _
'''''                    "  END IF; " & _
'''''                    "  If Actualiza = 1 Then " & _
'''''                    "    If IDExiste = 0 Or IDExiste Is Null Then " & _
'''''                    "      INSERT INTO Actualizaciones (IDOrigen, Tabla, Sucursal, IDTabla) VALUES (Old.ID, 'garantiasdevoluciones', IDSucursal, New.IDTabla); " & _
'''''                    "    END IF; " & _
'''''                    "  END IF; " & _
'''''                    "  SET New.Actualizar = 0; " & _
'''''                    "END $$ DELIMITER ;"
'''''    dbDatos.Execute "CREATE DEFINER=`root`@`localhost` TRIGGER `pasesinventarios_insert` BEFORE INSERT ON `pasesinventarios` FOR EACH ROW BEGIN " & _
'''''                    "  DECLARE IDSucursal       INT; " & _
'''''                    "  DECLARE NewID            BIGINT(21); " & _
'''''                    "  DECLARE Actualiza        TINYINT(1) DEFAULT 0; " & _
'''''                    "  SELECT AUTO_INCREMENT INTO NewID FROM information_schema.TABLES WHERE TABLE_SCHEMA = database() AND TABLE_NAME = 'pasesinventarios'; " & _
'''''                    "  SELECT Clave INTO IDSucursal FROM Sucursales WHERE Activa = 1 LIMIT 1; " & _
'''''                    "  IF NEW.ACTUALIZAR = 1 THEN " & _
'''''                    "    SET Actualiza = 0; " & _
'''''                    "  Else " & _
'''''                    "    SET Actualiza = 1; " & _
'''''                    "  END IF; " & _
'''''                    "  If Actualiza = 1 Then " & _
'''''                    "    INSERT INTO Actualizaciones (IDOrigen, Tabla, Sucursal) VALUES (NewID, 'pasesinventarios', IDSucursal); " & _
'''''                    "  END IF; " & _
'''''                    "  SET New.Actualizar = 0; " & _
'''''                    "END"
'''''    dbDatos.Execute "CREATE DEFINER=`root`@`localhost` TRIGGER `pasesinventarios_update` BEFORE UPDATE ON `pasesinventarios` FOR EACH ROW BEGIN " & _
'''''                    "  DECLARE IDSucursal INT; " & _
'''''                    "  DECLARE IDExiste   INT; " & _
'''''                    "  DECLARE Actualiza  TINYINT(1) DEFAULT 0; " & _
'''''                    "  SELECT Clave INTO IDSucursal FROM Sucursales WHERE Activa = 1 LIMIT 1; " & _
'''''                    "  SELECT count(IDOrigen) INTO IDExiste FROM Actualizaciones WHERE Tabla = 'pasesinventarios' AND IDOrigen = New.ID; " & _
'''''                    "  IF NEW.ACTUALIZAR = 1 THEN " & _
'''''                    "    SET Actualiza = 0; " & _
'''''                    "  Else " & _
'''''                    "    SET Actualiza = 1; " & _
'''''                    "  END IF; " & _
'''''                    "  If Actualiza = 1 Then " & _
'''''                    "    If IDExiste = 0 Or IDExiste Is Null Then " & _
'''''                    "      INSERT INTO Actualizaciones (IDOrigen, Tabla, Sucursal, IDTabla) VALUES (Old.ID, 'pasesinventarios', IDSucursal, New.IDTabla); " & _
'''''                    "    END IF; " & _
'''''                    "  END IF; " & _
'''''                    "  SET New.Actualizar = 0; " & _
'''''                    "END"

    dbDatos.Execute "ALTER TABLE `promociones` CHANGE COLUMN `IDTabla` `IDTabla` INT(10) UNSIGNED NOT NULL DEFAULT '0' AFTER `Id`," & _
                    "CHANGE COLUMN `Descripcion` `Descripcion` VARCHAR(50) NOT NULL DEFAULT '' AFTER `IDTabla`," & _
                    "CHANGE COLUMN `Periodo` `Tipo` ENUM('P','D','A') NOT NULL DEFAULT 'D' AFTER `Descripcion`," & _
                    "CHANGE COLUMN `TasaInteres` `PorcentajeDescuento` DOUBLE(15,5) NOT NULL DEFAULT '0' AFTER `Tipo`," & _
                    "CHANGE COLUMN `PorcBonificacion1` `DiasDescuento` INT(10) UNSIGNED NOT NULL DEFAULT '0' AFTER `PorcentajeDescuento`," & _
                    "CHANGE COLUMN `DiasGracia` `Activa` TINYINT(1) NOT NULL DEFAULT '1' AFTER `DiasDescuento`," & _
                    "DROP COLUMN `ClavePromocion`, DROP COLUMN `TipoInteres`, DROP COLUMN `PorcBonificacion2`," & _
                    "DROP COLUMN `PorcBonificacion3`, DROP COLUMN `PorcBonificacion4`, DROP COLUMN `NoTrabajo`,DROP COLUMN `TipoPeriodo`;"
    
    If Val(SacaValor("promociones", "Id", " Where Id=1")) = 0 Then
        dbDatos.Execute "INSERT INTO `promociones` (`IDTabla`, `Descripcion`, `Tipo`, `PorcentajeDescuento`, `DiasDescuento`, `Activa`, `Actualizar`) VALUES (0, '15 Dias Descuento', 'D', 0.00000, 15, 1, 0);"
    End If
    If Val(SacaValor("promociones", "Id", " Where Id=2")) = 0 Then
        dbDatos.Execute "INSERT INTO `promociones` (`IDTabla`, `Descripcion`, `Tipo`, `PorcentajeDescuento`, `DiasDescuento`, `Activa`, `Actualizar`) VALUES (0, 'CAMBIATE', 'D', 0.00000, 15, 1, 0);"
    End If
    dbDatos.Execute "ALTER TABLE `movimientos` ADD COLUMN `FolioReImpresiones` INT(10) NULL DEFAULT '1' AFTER `FolioPasesInventario`;"
    
    dbDatos.Execute "DROP PROCEDURE IF EXISTS `spRepVencidos`;"
    dbDatos.Execute "CREATE PROCEDURE `spRepVencidos`(IN `FechaIni` DATE, IN `FechaFin` DATE, IN `DiasEnajena` INTEGER, IN `TipoContrato` INTEGER, IN `TipoPrenda` INTEGER) " & _
                    "BEGIN SELECT DISTINCT e.ID,e.NumContrato,e.Fecha,e.Vencimiento,CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,e.Avaluo,e.Serie,e.Prestamo,e.TipoInteres,e.TipoTasa,c.Tel,c.Celular " & _
                    "FROM empeno e INNER JOIN clientes c ON e.IDCliente=c.ID LEFT JOIN detallesempeno de ON e.ID=de.IDEmpeno " & _
                    "WHERE if(TipoContrato=1,(e.Serie=1 OR e.Serie=2 OR e.Serie=3),if(TipoContrato=3,e.Serie=2,de.Tipo=TipoPrenda)) AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0,DiasEnajena) DAY),'%Y%/%m%/%d') BETWEEN FechaIni AND FechaFin AND e.Cancelado=0 AND e.Destino=0 AND e.Pagado=0 ORDER BY NumContrato; END"
End Sub

'Checamos los movimientos
Private Sub Checar_Movimientos()

On Error Resume Next

    dbDatos.Execute "UPDATE Movimientos SET Fecha='" & Format(Date, "YYYY/MM/DD") & "',Movimiento=1 WHERE Fecha <> '" & Format(Date, "YYYY/MM/DD") & "'"
End Sub

Public Function Regresa_Semaforo(ID As Long) As String

    Dim rcSemaforo As New ADODB.Recordset
    Dim Pagados As Integer, Activos As Integer, Almoneda As Integer, porEnajenados As Double, TotalMovimientos As Double
    
     rcSemaforo.Open "SELECT COUNT(ID) AS Pagados FROM empeno WHERE cancelado = 0 AND periodo <> 1 AND (destino = 3 or destino = 2) AND IDCliente=" & ID, dbDatos, adOpenStatic, adLockOptimistic
    If Not rcSemaforo.EOF Then
        Pagados = rcSemaforo!Pagados
    End If
        rcSemaforo.Close
            
    rcSemaforo.Open "SELECT COUNT(ID) AS Activos FROM empeno WHERE cancelado = 0 AND periodo <> 1 AND destino = 0 AND IDCliente=" & ID, dbDatos, adOpenStatic, adLockOptimistic
    If Not rcSemaforo.EOF Then
        Activos = rcSemaforo!Activos
    End If
        rcSemaforo.Close
            
    rcSemaforo.Open "SELECT COUNT(ID) AS Almoneda FROM empeno WHERE cancelado = 0 AND periodo <> 1 AND destino = 4 AND IDCliente=" & ID, dbDatos, adOpenStatic, adLockOptimistic
    If Not rcSemaforo.EOF Then
        Almoneda = rcSemaforo!Almoneda
    End If
        rcSemaforo.Close
        
        Set rcSemaforo = Nothing
    
  porEnajenados = SacaValor("parametros", "porEnajenados", "")
      
   If Almoneda = 0 Then
        Regresa_Semaforo = "Verde"
    ElseIf (Almoneda * 100) / (Pagados + Activos + Almoneda) >= porEnajenados Then
        Regresa_Semaforo = "Rojo"
    Else
        Regresa_Semaforo = "Amarillo"
    End If
    
    TotalMovimientos = Pagados + Activos + Almoneda
    
    If TotalMovimientos = 0 Then TotalMovimientos = 1
            
    frmEmpeño.lblInfoSemaforo.Caption = Redondeo((Activos * 100) / TotalMovimientos) & "% Activos " & _
                                    Redondeo((Pagados * 100) / TotalMovimientos) & "% Movimientos " & _
                                    Redondeo((Almoneda * 100) / TotalMovimientos) & "% Almoneda"


End Function


'Regresamos el Numero de Contrato segun la serie
Public Function Regresa_NumContrato(Opcion As Boolean, Serie As Integer) As Long
Dim rcFolios As New ADODB.Recordset
Dim Folio As Long

    rcFolios.Open "SELECT Folio FROM Folios WHERE Serie=" & Serie, dbDatos, adOpenForwardOnly, adLockOptimistic

    With rcFolios
    
        Folio = !Folio
    
        If Opcion Then
            Folio = Folio + 1
            dbDatos.Execute "UPDATE Folios SET Folio=Folio+1 WHERE Serie=" & Serie
        End If
  
        Regresa_NumContrato = Folio
    End With
  
    rcFolios.Close
    Set rcFolios = Nothing
End Function

''''''Regresamos el Folio segun la serie
'''''Public Function Regresa_Folio(Opcion As Boolean, Serie As Integer) As Long
'''''Dim rcFolios As New ADODB.Recordset
'''''Dim Folio As Long
'''''
'''''    rcFolios.Open "SELECT Folio FROM Folios WHERE Serie=" & Serie, dbDatos, adOpenDynamic, adLockOptimistic
'''''    With rcFolios
'''''
'''''        Folio = !Folio
'''''
'''''        If Opcion Then
'''''
'''''            Folio = Folio + 1
'''''            dbDatos.Execute "UPDATE Folios SET Folio=Folio+1 WHERE Serie=" & Serie
'''''
'''''        End If
'''''
'''''        Regresa_Folio = Folio
'''''    End With
'''''
'''''    rcFolios.Close
'''''    Set rcFolios = Nothing
'''''End Function

'Regresamos el movimiento
Public Function Regresa_Movimiento(Opcion As Boolean, Optional Campo As String = "Movimiento") As Long
Dim rcMovimientos As New ADODB.Recordset
Dim Movimiento As Long

    rcMovimientos.Open "SELECT " & Campo & " AS Campo FROM movimientos", dbDatos, adOpenForwardOnly, adLockOptimistic
    With rcMovimientos
        
        Movimiento = !Campo
        
        If Opcion Then
            
            Movimiento = Movimiento + 1
            dbDatos.Execute "UPDATE movimientos SET " & Campo & "=" & Movimiento
        
        End If
    
        Regresa_Movimiento = Movimiento
    End With

    rcMovimientos.Close
    Set rcMovimientos = Nothing
End Function

'regresamos el valor de la clave del archivo ini
Public Function Regresa_Valor(Seccion As String, Key As String, Default As String) As String
Dim Cadena As String, Lon As Integer
   
    Cadena = String(2555, 0)
   
    Lon = GetPrivateProfileString(Seccion, Key, Default, Cadena, 2555, App.Path & "\Configuracion.Ini")
    Cadena = Left$(Cadena, Lon)
    Regresa_Valor = Cadena
    
End Function

'''''Public Function Regresa_Valor_Exp(Seccion As String, Key As String, Default As String, Dir As String) As String
'''''Dim Cadena As String, Lon As Integer
'''''
'''''    Cadena = String(255, 0)
'''''
'''''    Lon = GetPrivateProfileString(Seccion, Key, Default, Cadena, 255, Dir)
'''''    Cadena = Left$(Cadena, Lon)
'''''    Regresa_Valor_Exp = Cadena
'''''
'''''End Function

''''''Grabamos datos en el fichero .ini
'''''Public Sub Graba_Valor(Seccion As String, Key As String, Valor As String)
'''''    WritePrivateProfileString Seccion, Key, Valor, App.Path & "\Configuracion.Ini"
'''''End Sub
'''''
'''''Public Sub Graba_Valor_Exp(Seccion As String, Key As String, Valor As String)
'''''    WritePrivateProfileString Seccion, Key, Valor, App.Path & "\Exportar\Ex" & Regresa_Valor("MONTEPIO", "NoSucursal", "") & "-" & Format(Date, "YYYYMMDD") & ".txt"
'''''End Sub

'Regresamos el parametro de la base de datos
Public Function Regresa_Valor_BD(Campo As String) As String
Dim rc As New ADODB.Recordset

On Error GoTo Error

    rc.Open "SELECT " & Campo & " FROM parametros", dbDatos, adOpenForwardOnly, adLockReadOnly
      
        Regresa_Valor_BD = IIf(IsNull(rc.Fields(Campo)), "", rc.Fields(Campo))
    
    rc.Close
    Set rc = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rc = Nothing
End Function

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
    obj.SelLength = Len(obj.text)

End Sub

'Convertimos las letras a Mayusculas
Public Function Mayusculas(Codigo As Integer) As Integer
    
    Mayusculas = Asc(UCase(Chr(Codigo)))

End Function

'Posicionamos las ventanas desplegables sobre el objeto donde apareceran
Public Sub Position(frmThis As Form, objThis As Object)
Dim tR As rect

    GetWindowRect objThis.hWnd, tR
    frmThis.Move tR.Left * Screen.TwipsPerPixelX, (tR.Bottom + 1) * Screen.TwipsPerPixelY

End Sub

'nos regresa la fecha completa
Public Function Regresar_Fecha(Fecha As Date) As String
    
    Regresar_Fecha = Format(Fecha, "DDDD") & " " & Format(Fecha, "DD") & " de " & Format(Fecha, "MMMM") & " del " & Format(Fecha, "YYYY")

End Function

'Centramos la forma
Public Sub CentrarForm(F As Form, m As MDIForm)
    
    F.Top = m.Top + (m.ScaleHeight - F.Height) / 2
    F.Left = m.Left + (m.ScaleWidth - F.Width) / 2

End Sub

'Nos permite aceptar solo numeros a una cantidad con decimales
Public Function Solo_Numeros(Codigo As Integer, Optional Opcion As Integer = 0) As Integer
    
    If (Codigo >= vbKey0 And Codigo <= vbKey9) Or Codigo = vbKeyBack Or Codigo = vbKeyReturn Or (Opcion = 1 And (Codigo = Asc(Separador) Or Codigo = Asc("-"))) Then
        
        Solo_Numeros = Codigo
        
    Else
        
        Solo_Numeros = 0
    
    End If
End Function

'Ordenamos el grid
Public Sub Ordenar_Grid(ByVal lCol As Long, ByRef grdGrid As vbalGrid, Asc As Integer, Des As Integer)
Dim iCol As Long
      
    With grdGrid.SortObject
        .Clear
        .SortColumn(1) = lCol
        If (grdGrid.ColumnSortOrder(lCol) = CCLOrderNone) Or (grdGrid.ColumnSortOrder(lCol) = CCLOrderDescending) Then
            .SortOrder(1) = CCLOrderAscending
        Else
            .SortOrder(1) = CCLOrderDescending
        End If
        grdGrid.ColumnSortOrder(lCol) = .SortOrder(1)
        .SortType(1) = grdGrid.ColumnSortType(lCol)
      
        'Place ascending/descending icon:
        For iCol = 1 To grdGrid.Columns
            If (iCol <> lCol) Then
                If grdGrid.ColumnImage(iCol) = Des Or grdGrid.ColumnImage(iCol) = Asc Then
                    grdGrid.ColumnImage(iCol) = -1
                End If
            ElseIf grdGrid.ColumnHeader(iCol) <> "" Then
                grdGrid.ColumnImageOnRight(iCol) = True
                If (.SortOrder(1) = CCLOrderAscending) Then
                    grdGrid.ColumnImage(iCol) = Asc
                Else
                    grdGrid.ColumnImage(iCol) = Des
                End If
            End If
        Next iCol
      
    End With
   
    Screen.MousePointer = vbHourglass
    On Error Resume Next
    grdGrid.Sort
    Screen.MousePointer = vbDefault
End Sub

'Ponemos los controles en modo flat
Public Sub Poner_Flat(ByRef Fl() As cFlatControl, Controles As Object, forma As Form)
'''''''Dim Contador As Integer
'''''''Dim Control As Object
'''''''
'''''''   For Each Control In Controles
'''''''
'''''''      If TypeOf Control Is TextBox Or TypeOf Control Is MaskEdBox Then
'''''''
'''''''         ReDim Preserve Fl(0 To Contador)
'''''''         Set Fl(Contador) = New cFlatControl
'''''''         Fl(Contador).hWndAttach Control.hWnd, forma.hWnd, False
'''''''         Contador = Contador + 1
'''''''
'''''''      ElseIf TypeOf Control Is ComboBox Then
'''''''
'''''''         ReDim Preserve Fl(0 To Contador)
'''''''         Set Fl(Contador) = New cFlatControl
'''''''         Fl(Contador).hWndAttach Control.hWnd, forma.hWnd, True
'''''''         Contador = Contador + 1
'''''''
'''''''      End If
'''''''
'''''''   Next
End Sub

Public Sub Quitar_Flat(ByRef Fl() As cFlatControl)
''''''Dim i As Integer
''''''
''''''    'Descargamos de memoria el flat
''''''    For i = LBound(Fl) To UBound(Fl)
''''''
''''''      Set Fl(i) = Nothing
''''''
''''''    Next i
''''''
End Sub

''''''Mandamos el mensaje a todas las ventanas
'''''Public Sub Mandar_Mensaje(frm As Object, Msg As Long, wParam As Long, lParam As Variant)
'''''Dim tForm As Object
'''''
'''''   For Each tForm In Forms
'''''
'''''        If tForm.Name <> frm.Name Then SendMessage tForm.hwnd, Msg, wParam, ByVal CLng(lParam)
'''''
'''''   Next
'''''
'''''End Sub

'Regresemos los intereses
'Public Function Regresa_Intereses(ID As Long, Fecha As Date, Prestamo As Double, Avaluo As Double, Folio As Long, Vencimiento As Date, TipoInteres As String, Optional VENTACLIENTE As Boolean = False) As Double
'Dim crIntereses As Double
'
'    Select Case TipoInteres
'    Case "MENSUAL"
'
'        crIntereses = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Tasa", Vencimiento)
'
'    Case "QUINCENAL", "SEMANAL"
'
'        crIntereses = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Tasa", Vencimiento)
'
'    Case "DIARIA"
'
'        crIntereses = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Tasa", Vencimiento)
'    End Select
'
'    Regresa_Intereses = crIntereses
'End Function

'Regresamos el ultimo dia del mes
Public Function Regresa_Ultimo_Dia_Mes(Fecha As Date) As Integer
Dim dia As Integer
   
    dia = Day(Fecha)
   
    Fecha = Fecha - (dia - 1)
   
    Fecha = DateAdd("M", 1, Fecha)

    Fecha = Fecha - 1
   
    Regresa_Ultimo_Dia_Mes = Day(Fecha)

End Function

Public Function Regresa_Mes_Ultimo_Dia(Fecha As Date) As Date
Dim dia As Integer
   
    dia = Day(Fecha)
   
    Fecha = Fecha - (dia - 1)
   
    Fecha = DateAdd("M", 1, Fecha)

    Fecha = Fecha - 1
   
    Regresa_Mes_Ultimo_Dia = Fecha

End Function

Public Function Regresa_Dias(Fecha As Date) As Long
    Regresa_Dias = DateDiff("D", Fecha, Date)
End Function

'Ponemos los colores intercalados a los renglones de las listas
Public Sub Poner_Colores(grdObj As Object, Renglon As Long, Optional Opcion As Long = 0)
Dim columna As Integer
    
    For columna = 1 To grdObj.Columns
        
        grdObj.CellBackColor(Renglon, columna) = IIf(IIf(Opcion = 0, Renglon, Opcion) Mod 2 = 0, RGB(226, 220, 197), RGB(238, 234, 221))
    
    Next columna

End Sub

Public Sub Colorea(grdObj As Object, Renglon As Long, Color As String)
Dim columna As Integer

    For columna = 1 To grdObj.Columns
        
        grdObj.CellBackColor(Renglon, columna) = Color
    
    Next columna
    
End Sub

'Realizamos el reporte pasando los totales
Public Sub Realizar_Reporte(Sucursal As String, Cajero As String)
Dim rcMonto As New ADODB.Recordset
Dim crSaldo As Currency, crDebe As Currency, crHaber As Currency
  
On Error GoTo Error

    dbReportes.Execute "DELETE FROM cortecajaventanilla"
  
    'Ponemos el saldo anterior
    rcMonto.Open "SELECT Saldo FROM saldos WHERE Fecha <'" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "' ORDER BY Fecha DESC", dbDatos, adOpenDynamic, adLockOptimistic
    
    'Ponemos la cantidad del debe
    If Not rcMonto.BOF And Not rcMonto.EOF Then
    
        crSaldo = Val(rcMonto!Saldo & "")
    Else
    
        crSaldo = 0
    End If
    rcMonto.Close
  
    rcMonto.Open "SELECT SUM(Importe)AS Total FROM auxiliar WHERE Cuenta='110101' OR Cuenta='110901' AND Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "'", dbDatos, adOpenStatic, adLockOptimistic
    'Ponemos la cantidad del debe
    crDebe = Val(rcMonto!Total & "")
    rcMonto.Close
  
    'Ponemos la cantidad del haber
    rcMonto.Open "SELECT SUM(Importe)AS Total FROM auxiliar WHERE Cuenta='110150' AND Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND PC='" & NombrePc & "'", dbDatos, adOpenStatic, adLockOptimistic
  
    'Ponemos la cantidad del debe
    crHaber = Val(rcMonto!Total & "")
    rcMonto.Close
  
    dbReportes.Execute "INSERT INTO cortecajaventanilla (Sucursal,Cajero,Saldo,Debe,Haber) VALUES " & _
                        "('" & Sucursal & "','" & Cajero & "'," & ConvMoneda(crSaldo) & "," & ConvMoneda(crDebe) & "," & ConvMoneda(crHaber) & ")"
    
    Set rcMonto = Nothing
    Exit Sub
  
Error:
    Maneja_Error Err
    Set rcMonto = Nothing
End Sub

'ponemos las cuentas en el reporte y las separamos
Public Sub Realizar_Cuentas()
Dim rcAuxiliar As New ADODB.Recordset

On Error GoTo Error

    dbReportes.Execute "DELETE FROM cortecuentas"
    
    rcAuxiliar.Open "SELECT Cuenta,Importe,Concepto,Folio,Movimiento,Iniciales,Serie FROM auxiliar WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' AND Cuenta<>'999401' AND Cuenta<>'999450' AND Cuenta<>'910901' AND Cuenta<>'910950'", dbDatos, adOpenForwardOnly, adLockOptimistic
    With rcAuxiliar
        
        While Not .EOF
            
            dbReportes.Execute "INSERT INTO cortecuentas (Cuenta,Descripcion,Folio,Movimientos,Cargo,Abono,Concepto,PC,Serie) VALUES " & _
                                "('" & !Cuenta & "','" & !Concepto & "','" & !Folio & "','" & !Movimiento & "'," & IIf(Right(!Cuenta, 2) = "01", ConvMoneda(!Importe), 0) & "," & IIf(Right(!Cuenta, 2) = "50", ConvMoneda(!Importe), 0) & ",'" & !Iniciales & "','" & Mid(!Cuenta, 1, 4) & "00" & "'," & !Serie & ")"
        .MoveNext
        Wend
    
    End With
    rcAuxiliar.Close
    Set rcAuxiliar = Nothing
    
    dbReportes.Execute "DELETE FROM cortecuentas WHERE (Concepto<>'CV50' AND Concepto<>'RE50' AND Cuenta='110150') OR (Concepto<>'DO01' AND Cuenta='110101')"
'''    dbReportes.Execute "DELETE FROM cortecuentas WHERE (Concepto<>'CV50' AND Concepto<>'RE50' AND Cuenta='199450') OR (Concepto<>'DO01' AND Cuenta='199401')"
    dbReportes.Execute "DELETE FROM cortecuentas WHERE (Descripcion='Refrendo' AND Cuenta='201701') OR (Descripcion='Refrendo' AND Cuenta='201750')"
    dbReportes.Execute "DELETE FROM cortecuentas WHERE (Descripcion='Dotacion Divisas' AND Cuenta='200950') OR (Descripcion='Retiro Divisas' AND Cuenta='200901')"
    dbReportes.Execute "DELETE FROM cortecuentas WHERE Serie=2 AND Descripcion<>'Empeño'"
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcAuxiliar = Nothing
End Sub

'Creamos el reporte diario
Public Sub Realizar_Diario(Sucursal As String, Cajero As String, Mayor As String, Cuenta As String, Leyenda As String, Optional Opcion As Boolean = False)
Dim rcAuxiliar As New ADODB.Recordset
Dim lFolio1 As Long, lFolio2 As Long, lFolio3 As Long
Dim crImporte1 As Currency, crImporte2 As Currency, crImporte3 As Currency
  
    If Opcion Then dbReportes.Execute "DELETE FROM diario"
  
    rcAuxiliar.Open "SELECT * FROM auxiliar WHERE Cuenta='" & Cuenta & "' AND Fecha='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY Folio", dbDatos, adOpenForwardOnly, adLockOptimistic
  
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
        
            dbReportes.Execute "INSERT INTO diario (Sucursal,Cajero,Cuenta,Leyenda,Importe1,Folio1,Importe2,Folio2,Importe3,Folio3) VALUES " & _
                                  "('" & Sucursal & "','" & Cajero & "','" & Cuenta & "','" & Leyenda & "'," & ConvMoneda(crImporte1) & "," & lFolio1 & "," & ConvMoneda(crImporte2) & "," & lFolio2 & "," & ConvMoneda(crImporte3) & "," & lFolio3 & ")"
      
        Wend
        
    End With
  
    rcAuxiliar.Close
    Set rcAuxiliar = Nothing
End Sub

'Regresamos el nombre de la computadora
Public Function Nombre_Pc() As String
Dim dwLen As Long, strString As String
   
    'Create a buffer
    dwLen = MAX_COMPUTERNAME_LENGTH + 1
    strString = String(dwLen, "X")
    
    'Get the computer name
    GetComputerName strString, dwLen
    
    'get only the actual data
    strString = Left(strString, dwLen)
    
    'Show the computer name
    Nombre_Pc = strString
    
End Function

Public Sub Cargar_Combos(Campo As String, Tabla As String, Combo As ComboBox, Optional Condicion As String = "", Optional CampoOrdenamiento As String = "", Optional Limpiar As Boolean = True, Optional CampoClave As String = "ID")
Dim rcTipos As New ADODB.Recordset
   
On Error GoTo Error

    rcTipos.Open "SELECT " & Campo & " AS Valor," & CampoClave & " AS IDD FROM " & Tabla & IIf(Condicion = "", "", Condicion) & IIf(CampoOrdenamiento = "", "", " ORDER BY " & CampoOrdenamiento), dbDatos, adOpenForwardOnly, adLockReadOnly
    
    If Limpiar Then Combo.Clear

    With rcTipos
        
        While Not .EOF
            
            Combo.AddItem UCase(.Fields("Valor"))
            Combo.ItemData(Combo.NewIndex) = !idd
            
        .MoveNext
        Wend
    
    End With

    rcTipos.Close
    Set rcTipos = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTipos = Nothing
End Sub

'''''Function ActualizaInformacion(NomComercial As String, RazonSocial As String, RFC As String, Direccion As String, Colonia As String, Ciudad As String, Estado As String, Telefono As String, CP As String)
'''''
'''''    dbDatos.Execute "UPDATE sucursales SET NombreComercial='" & NomComercial & "',RazonSocial='" & RazonSocial & "',RFC='" & RFC & "',Direccion='" & Direccion & "',Ciudad='" & Ciudad & "',Estado='" & Estado & "',Telefono='" & Telefono & "',CP=" & CP & " where Activa=1"
'''''
'''''End Function

Function Regresa_Valor_Empeno(Campo As String, ID As Long) As String
Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error

    rcConsulta.Open "SELECT " & Campo & " FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockReadOnly

    If Not rcConsulta.BOF And Not rcConsulta.EOF And Not IsNull(rcConsulta.Fields(0)) Then Regresa_Valor_Empeno = rcConsulta.Fields(0) Else Regresa_Valor_Empeno = 0

    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Function

Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function

Function CantidadEnLetra(tyCantidad As Currency) As String
Dim lyCantidad As Currency, lyCentavos As Currency, lnDigito As Byte, lnPrimerDigito As Byte, lnSegundoDigito As Byte, lnTercerDigito As Byte, lcBloque As String, lnNumeroBloques As Byte, lnBloqueCero
Dim launidades, ladecenas, lacentenas As Variant
Dim i As Long
    tyCantidad = Round(tyCantidad, 2)
    lyCantidad = Int(tyCantidad)
    lyCentavos = (tyCantidad - lyCantidad) * 100
    launidades = Array("UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", "VEINTE", "VEINTIUN", "VEINTIDOS", "VEINTITRES", "VEINTICUATRO", "VEINTICINCO", "VEINTISEIS", "VEINTISIETE", "VEINTIOCHO", "VEINTINUEVE")
    ladecenas = Array("DIEZ", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA")
    lacentenas = Array("CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS")
    lnNumeroBloques = 1
    Do
        lnPrimerDigito = 0
        lnSegundoDigito = 0
        lnTercerDigito = 0
        lcBloque = ""
        lnBloqueCero = 0
        For i = 1 To 3
            lnDigito = lyCantidad Mod 10
            If lnDigito <> 0 Then
                Select Case i
                Case 1
                    lcBloque = " " & launidades(lnDigito - 1)
                    lnPrimerDigito = lnDigito
                Case 2
                    If lnDigito <= 2 Then
                        lcBloque = " " & launidades((lnDigito * 10) + lnPrimerDigito - 1)
                    Else
                        lcBloque = " " & ladecenas(lnDigito - 1) & IIf(lnPrimerDigito <> 0, " Y", Null) & lcBloque
                    End If
                    lnSegundoDigito = lnDigito
                Case 3
                    lcBloque = " " & IIf(lnDigito = 1 And lnPrimerDigito = 0 And lnSegundoDigito = 0, "CIEN", lacentenas(lnDigito - 1)) & lcBloque
                    lnTercerDigito = lnDigito
                End Select
            Else
                lnBloqueCero = lnBloqueCero + 1
            End If
            lyCantidad = Int(lyCantidad / 10)
            If lyCantidad = 0 Then
                Exit For
            End If
        Next i
        Select Case lnNumeroBloques
        Case 1
            CantidadEnLetra = lcBloque
        Case 2
            CantidadEnLetra = lcBloque & IIf(lnBloqueCero = 3, Null, " MIL") & CantidadEnLetra
        Case 3
            CantidadEnLetra = lcBloque & IIf(lnPrimerDigito = 1 And lnSegundoDigito = 0 And lnTercerDigito = 0, " MILLON", " MILLONES") & CantidadEnLetra
        End Select
        lnNumeroBloques = lnNumeroBloques + 1
    Loop Until lyCantidad = 0
    CantidadEnLetra = "" & CantidadEnLetra & IIf(tyCantidad > 1, " PESOS ", " PESO ") & Format(str(lyCentavos), "00") & "/100 M.N."
End Function

Function RegresaEspacios(crImporte As Double, Espacios As Integer, Optional Signo As Boolean = False) As String
Dim strCadena As String, Lon As Integer

    strCadena = Format(crImporte, FMoneda)
    If Signo Then strCadena = "$" & strCadena

    Lon = Len(Trim(strCadena))
    strCadena = String(Espacios - Lon, " ") & strCadena
    
    RegresaEspacios = strCadena
End Function

Function SacaKilates(Kilates As Integer) As String
Dim rcKilataje As New ADODB.Recordset

On Error GoTo Error

    rcKilataje.Open "SELECT descripcion FROM kilatajes WHERE Clave=" & Kilates, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcKilataje.BOF And Not rcKilataje.EOF Then
        
        SacaKilates = rcKilataje!Descripcion
    
    Else
        
        SacaKilates = ""
    
    End If
    rcKilataje.Close
    Set rcKilataje = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcKilataje = Nothing
End Function

Function RegresaKilates(Kilates As String, Optional Tipo As String) As Integer
Dim rcKilataje As New ADODB.Recordset

On Error GoTo Error

    RegresaKilates = 0
    rcKilataje.Open "SELECT clave FROM kilatajes WHERE descripcion='" & Trim(Kilates) & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcKilataje.BOF And Not rcKilataje.EOF Then

        RegresaKilates = rcKilataje!Clave
    End If
    rcKilataje.Close
    Set rcKilataje = Nothing
    Exit Function

Error:
    Maneja_Error Err
    Set rcKilataje = Nothing
End Function

Public Sub SombreaGrid(Grilla As vbalGrid, Rojo As Long, Verde As Long, Azul As Long, Optional Rojoo As Long = 255, Optional Verdee As Long = 255, Optional Azull As Long = 255, Optional Opcion As Long = 0)
Dim i As Long, x As Integer

    With Grilla

        For i = IIf(Opcion = 0, 1, Opcion) To IIf(Opcion = 0, .Rows, Opcion)
            
            For x = 1 To .Columns

                If Opcion = 0 Then
                    
                    .CellBackColor(i, x) = IIf(i Mod 2 <> 0, RGB(Rojo, Verde, Azul), RGB(Rojoo, Verdee, Azull))
                
                Else
                    
                    .CellBackColor(i, x) = IIf((Opcion Mod 2) = 0, RGB(Rojo, Verde, Azul), RGB(Rojoo, Verdee, Azull))
                
                End If

            Next x
        
        Next i

    End With

End Sub

Public Function ComboInformacion(Combo As ComboBox, Clave As Integer, Optional strDescripcion As String = "") As Integer
Dim i As Integer

On Error GoTo Error

    For i = 0 To Combo.ListCount
        
        If strDescripcion = "" Then
            
            If Clave = Combo.ItemData(i) Then ComboInformacion = i: Exit Function
        
        Else
        
            If strDescripcion = Combo.List(i) Then ComboInformacion = i: Exit Function
        End If
        
    Next i

Error:
    ComboInformacion = -1
End Function

Public Function Calcula_Edad(Fecha As Date) As Integer

    Calcula_Edad = Format(Date, "YYYY") - Format(Fecha, "YYYY")
    
    If Date < DateAdd("YYYY", Calcula_Edad, Fecha) Then
        
        Calcula_Edad = Calcula_Edad - 1
    End If
    
End Function

Function CreaDigitoVerificador(Codigo As String) As String
Dim lcRet As String, lnI As Integer, lnCheckSum As Double, lnAux As Double, x As Double

    lcRet = Trim(Codigo)
    lnCheckSum = 0
    
    For lnI = 1 To 12
        x = lnI Mod 2
        If x = 0 Then
            lnCheckSum = lnCheckSum + Val(Mid(lcRet, lnI, 1)) * 3
        Else
            lnCheckSum = lnCheckSum + Val(Mid(lcRet, lnI, 1)) * 1
        End If
    Next lnI
    
    lnAux = lnCheckSum Mod 10
    lcRet = lcRet + Trim(str(IIf(lnAux = 0, 0, 10 - lnAux)))
    CreaDigitoVerificador = lcRet
End Function

Public Function Regresa_Sucursal(Campo As String) As String
Dim rc As New ADODB.Recordset
   
On Error GoTo Error

    rc.Open "SELECT " & Campo & " FROM sucursales WHERE Activa=1", dbDatos, adOpenForwardOnly, adLockReadOnly
      
        Regresa_Sucursal = rc.Fields(Campo)
    
    rc.Close
    Set rc = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rc = Nothing
End Function

Public Function CreaCodigoBarras(Sucursal As String, TipoEntrada As Integer, Boleta As String, Partida As Integer) As String
Dim i As Integer

    For i = Len(Boleta) To 5
        
        Boleta = "0" & Boleta
    
    Next i

    CreaCodigoBarras = Sucursal & TipoEntrada & Boleta & Format(Partida, "00")
    CreaCodigoBarras = CreaDigitoVerificador(CreaCodigoBarras)

End Function

''Regresemos el importe del Almacenaje
'Public Function Regresa_Almacenaje(ID As Long, Fecha As Date, Prestamo As Double, Avaluo As Double, Folio As Long, Vencimiento As Date, TipoInteres As String, Optional VENTACLIENTE As Boolean = False) As Double
'Dim crAlmacenaje As Double
'
'    Select Case TipoInteres
'    Case "MENSUAL"
'
'        crAlmacenaje = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Almacenaje", Vencimiento)
'
'    Case "QUINCENAL", "SEMANAL"
'
'        crAlmacenaje = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Almacenaje", Vencimiento)
'
'    Case "DIARIA"
'
'        crAlmacenaje = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Almacenaje", Vencimiento)
'    End Select
'
'    Regresa_Almacenaje = crAlmacenaje
'End Function

'Regresemos el importe del Seguro
'Public Function Regresa_Seguro(ID As Long, Fecha As Date, Prestamo As Double, Avaluo As Double, Folio As Long, Vencimiento As Date, TipoInteres As String, Optional VENTACLIENTE As Boolean = False) As Double
'Dim crSeguro As Double
'
'    Select Case TipoInteres
'    Case "MENSUAL"
'
'        crSeguro = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Seguro", Vencimiento)
'
'    Case "QUINCENAL", "SEMANAL"
'
'        crSeguro = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Seguro", Vencimiento)
'
'    Case "DIARIA"
'
'        crSeguro = GeneraIntereses(Prestamo, Avaluo, Folio, Fecha, ID, "Seguro", Vencimiento)
'    End Select
'
'
'    Regresa_Seguro = crSeguro
'End Function

Function Regresa_Iva(Importe As Double, ID As Long) As Double
Dim Iva As Double

    Iva = Regresa_Valor_Empeno("Iva", ID) / 100

    Regresa_Iva = Importe * Iva
End Function

'Funcion para abrir el cajon mediante el puerto serial(COM1)
Public Sub Abrir_Cajon()

'On Error GoTo error
'
'    'Cierra el puerto para permitir nuevos parametros
'    If frmMDI.Com.PortOpen Then
'        frmMDI.Com.PortOpen = False
'    End If
'
'    'Puerto que sera usado
'    frmMDI.Com.CommPort = 1
'
'    'Baudios, paridad, datos, detener
'    frmMDI.Com.Settings = "9600,N,8,1"
'
'    'Activa el puerto COM
'    frmMDI.Com.PortOpen = True
'
'    'Texto de salida para el puerto
'    frmMDI.Com.Output = "U"
'
'    Exit Sub
'
'error:
'    Maneja_Error Err
End Sub

Function SacaValor(Tabla As String, Campo As String, Optional Condicion As String = "") As String
Dim rcValor As New ADODB.Recordset

On Error GoTo Error
    
    rcValor.Open "SELECT " & Campo & " AS Valor FROM " & Tabla & " " & Condicion, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not rcValor.BOF And Not rcValor.EOF And Not IsNull(rcValor.Fields("Valor")) Then
        
        SacaValor = rcValor.Fields("Valor")
    Else
        
        SacaValor = ""
    End If
    rcValor.Close
    Set rcValor = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcValor = Nothing
End Function

'Function GeneraIntereses(crPrestamo As Double, crAvaluo As Double, Folio As Long, Fecha As Date, ID As Long, TipoTasa As String, Vencimiento As Date, Optional VENTACLIENTE As Boolean = False, Optional Demasia As Boolean = False) As Double
'
'    Dim crInteres As Double, i As Integer, PeriodosPromocion As Double, FechaOriginal As Date, DiasGracia As Integer, NumDias As Integer, crImportePromocion As Double
'    Dim rcParametros As New ADODB.Recordset
'
'On Error GoTo error
'
'    FechaOriginal = Fecha
'    DiasGracia = Regresa_Valor_BD("DiasGracia")
'
'    With rcParametros
'
'        .Open "SELECT " & TipoTasa & " / 100 AS Tasa,Periodo,VenPeriodo,TipoInteres,Promocion FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
'        If Not .BOF And Not .EOF Then
'
'            If VENTACLIENTE = False And !Promocion > 0 Then
'
'                crImportePromocion = ChecaPromocion(ID, DiasGracia, TipoTasa)
'            End If
'
'            For i = 1 To IIf(VENTACLIENTE, 500, !VenPeriodo)
'
'                'Vencimiento
'                If !Periodo = 30 Then
'
'                    Fecha = DateAdd("M", i, FechaOriginal)
'                Else
'
'                    Fecha = DateAdd("D", !Periodo, Fecha)
'                End If
'
'                'Tasa que se va a consultar
'                crInteres = crPrestamo * (!Tasa * i)
'
'                GeneraIntereses = (crInteres - crImportePromocion)
'                If Demasia Then
'
'                    If Fecha >= Vencimiento Then GoTo DiasEnajenacion
'                Else
'
'                    If Date <= DateAdd("D", DiasGracia, Fecha) Then GoTo DiasEnajenacion
'                End If
'
'            Next i
'
'        End If
'        .Close
'        Set rcParametros = Nothing
'
'    End With
'
'DiasEnajenacion:
'    Exit Function
'
'error:
'    Maneja_Error Err
'    Set rcParametros = Nothing
'End Function

'Function GeneraIntereses(ByVal ID As Long, ByVal TipoTasa As String) As Double
'    Dim DiasTrans As Integer, crIntereses As Double
'    Dim rcParametros As New ADODB.Recordset
'    Dim rcPromocion As New ADODB.Recordset
'    Dim DiasGracia As Integer, DiasDescuento As Integer, PorcDescuento As Double
'    Dim lDiasGracia As Integer
'On Error GoTo Error
'
'    With rcParametros
'
'        .Open "SELECT " & TipoTasa & " / 100 AS Tasa,Periodo,VenPeriodo,Fecha,Prestamo,Vencimiento,Promocion,Origen FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
'
'        If Not .BOF And Not .EOF Then
'            DiasGracia = Val(Regresa_Valor_BD("DiasGracia"))
'            DiasTrans = DateDiff("D", !Fecha, Date)
'
'            DiasDescuento = 0
'            PorcDescuento = 0
'            If DiasTrans = 0 Then DiasTrans = 1
'            If !Promocion <> 0 Then
'                If !Origen = 1 Then
'                    rcPromocion.Open "SELECT Tipo,PorcentajeDescuento,DiasDescuento FROM Promociones WHERE ID=" & !Promocion, dbDatos, adOpenForwardOnly, adLockOptimistic
'                    If Not rcPromocion.BOF And Not rcPromocion.EOF Then
'                        Select Case rcPromocion!Tipo
'                            Case "D"
'                                DiasDescuento = IIf(DiasTrans > rcPromocion!DiasDescuento, rcPromocion!DiasDescuento, 0)
'                            Case "P"
'                                PorcDescuento = rcPromocion!porcentajeDescuento / 100
'                            Case "A"
'                                DiasDescuento = IIf(DiasTrans > rcPromocion!DiasDescuento, rcPromocion!DiasDescuento, 0)
'                                PorcDescuento = rcPromocion!porcentajeDescuento / 100
'                        End Select
'                    End If
'                End If
'            End If
'
'''''            If DiasGracia > 0 Then
'''''
'''''                If (Date > DateValue(!Vencimiento)) And (Date <= DateAdd("D", DiasGracia, DateValue(!Vencimiento))) Then
'''''                    DiasTrans = DateDiff("d", !Fecha, !Vencimiento)
'''''                Else
'''''                    DiasTrans = DateDiff("d", !Fecha, Date)
'''''                End If
'''''            End If
'            '****
'            Dim rc As New ADODB.Recordset
'            Set rc = dbDatos.Execute("SELECT * FROM empeno WHERE ID=" & ID)
'            lDiasGracia = 0
'
'
'            If !Periodo <> 1 And rc!Serie = 2 Then
'                        Dim rcParam As New ADODB.Recordset
'                        Dim diasTranscurridos As Integer
'                        Set rcParam = dbDatos.Execute("SELECT * FROM parametros WHERE ID=1")
'
'                        If rc!Serie = 2 Then
'                            lDiasGracia = rcParam!DiasGraciaAuto
'                        Else
'                            lDiasGracia = rcParam!DiasGracia
'                        End If
'                        'Tomo los días transcurridos
'                        diasTranscurridos = DateDiff("d", !Fecha, Date)
'                        'Asigno interes según el periodo
'                        If diasTranscurridos <= (!Periodo + lDiasGracia) Then
'                            crIntereses = Redondeo((((!Prestamo * !Tasa))) * 1)
'                        ElseIf diasTranscurridos > (!Periodo + lDiasGracia) And diasTranscurridos <= ((!Periodo * 2) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 2)
'                        ElseIf diasTranscurridos > ((!Periodo * 2) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 3) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 3)
'                        ElseIf diasTranscurridos > ((!Periodo * 3) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 4) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 4)
'                        ElseIf diasTranscurridos > ((!Periodo * 4) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 5) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 5)
'                        ElseIf diasTranscurridos > ((!Periodo * 5) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 6) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 6)
'                        ElseIf diasTranscurridos > ((!Periodo * 6) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 7) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 7)
'                        ElseIf diasTranscurridos > ((!Periodo * 7) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 8) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 8)
'                        ElseIf diasTranscurridos > ((!Periodo * 8) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 9) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 9)
'                        ElseIf diasTranscurridos > ((!Periodo * 9) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 10) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 10)
'                        ElseIf diasTranscurridos > ((!Periodo * 10) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 11) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 11)
'                        ElseIf diasTranscurridos > ((!Periodo * 11) + lDiasGracia) And diasTranscurridos <= ((!Periodo * 12) + lDiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 12)
'                    End If
'            Else
'                crIntereses = Redondeo(!Prestamo * (!Tasa / 30) * (DiasTrans - DiasDescuento)) '!Periodo
'                crIntereses = crIntereses - (crIntereses * PorcDescuento)
'            End If
'
'        End If
'
'        .Close
'
'        Set rcParametros = Nothing
'
'    End With
'
'    GeneraIntereses = crIntereses
'
'    Exit Function
'
'Error:
'    Maneja_Error Err
'    Set rcParametros = Nothing
'End Function
Public Function Regresa_Intereses_Plazo(ID As Long, Fecha As Date, Prestamo As Double, Avaluo As Double, Folio As Long, Vencimiento As Date, TipoInteres As String) As Double
Dim Intereses As Double
   
    Intereses = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Tasa", Vencimiento)
    Regresa_Intereses_Plazo = Intereses
End Function
'Regreso el Almacenaje
Public Function Regresa_Almacenaje_Plazo(ID As Long, Fecha As Date, Prestamo As Double, Avaluo As Double, Folio As Long, Vencimiento As Date, TipoInteres As String) As Double
Dim crAlmacenaje As Double
    
    Select Case TipoInteres
    Case "Mensual"
        
        crAlmacenaje = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Almacenaje", Vencimiento)
        
    Case "Quincenal"
        
        crAlmacenaje = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Almacenaje", Vencimiento)
            
    Case "Semanal"
    
        crAlmacenaje = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Almacenaje", Vencimiento)
                
    Case "Diaria"
        
        crAlmacenaje = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Almacenaje", Vencimiento)
        
    End Select
    
    Regresa_Almacenaje_Plazo = crAlmacenaje
End Function
'Regreso el Seguro
Public Function Regresa_Seguro_Plazo(ID As Long, Fecha As Date, Prestamo As Double, Avaluo As Double, Folio As Long, Vencimiento As Date, TipoInteres As String) As Double
Dim crSeguro As Double
    
    Select Case TipoInteres
    Case "Mensual"
            
        crSeguro = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Seguro", Vencimiento)
                            
    Case "Quincenal"
            
        crSeguro = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Seguro", Vencimiento)
                        
    Case "Semanal"
        
        crSeguro = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Seguro", Vencimiento)
        
    Case "Diaria"
        
        crSeguro = GeneraInteresesPlazo(Prestamo, Avaluo, Folio, Fecha, ID, "Seguro", Vencimiento)
        
    End Select
    
    Regresa_Seguro_Plazo = crSeguro
End Function
Function GeneraInteresesPlazo(Prestamo As Double, Avaluo As Double, Folio As Long, Fecha As Date, ID As Long, TipoTasa As String, Vencimiento As Date) As Double
    Dim Interes As Double, i As Integer, Dias As Integer, GTOSVenta As Double, FechaOriginal As Date, FechaAux As Date
    Dim InteresPeriodo As Double, DiasGracia As Integer, Promocion As Integer
    Dim rcParametros As New ADODB.Recordset
    Dim DiasTrans As Integer
    On Error GoTo Error

    FechaOriginal = Fecha
    'DiasGracia = Regresa_Valor_BD("DiasGracia")
    InteresPeriodo = 0

    rcParametros.Open "SELECT " & TipoTasa & " / 100 AS Tasa,Periodo,VenPeriodo,TipoTasa,Serie FROM Empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockReadOnly
    If rcParametros!Serie = 1 Then
        DiasGracia = Regresa_Valor_BD("DiasGracia")
    Else
        DiasGracia = Regresa_Valor_BD("DiasGraciaAutos")
    End If
    If Not rcParametros.BOF And Not rcParametros.EOF Then

        'Vencimiento
        If rcParametros!Periodo = 30 Then

            Fecha = DateAdd("M", 1, FechaOriginal)
        Else

            Fecha = DateAdd("D", 1, FechaOriginal)
        End If

        If Date <= DateAdd("D", DiasGracia, Fecha) Then GoTo diasEnajenacion
        
        DiasTrans = DateDiff("D", Fecha, Date)
        DiasTrans = IIf(DiasTrans < 0, 0, DiasTrans)
       
        FechaOriginal = Fecha
        For i = 1 To DiasTrans

            'Vencimiento
                Fecha = DateAdd("D", i, FechaOriginal)
                Interes = Prestamo * (rcParametros!Tasa / rcParametros!Periodo) * i
            
            GeneraInteresesPlazo = Interes

            If Date <= DateAdd("D", IIf(i = DateDiff("D", FechaOriginal, Vencimiento), DiasGracia, 0), Fecha) Then GoTo diasEnajenacion

        Next i

    End If

    rcParametros.Close
    Set rcParametros = Nothing

diasEnajenacion:
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcParametros = Nothing
End Function
Function GeneraIntereses(ByVal ID As Long, ByVal TipoTasa As String, Optional DiasMinimo As Integer = 1) As Double
    Dim DiasTrans, MesTrancurridos As Integer, crIntereses As Double, InteresPeriodo As Double
    Dim rcParametros As New ADODB.Recordset
    Dim rcPromocion As New ADODB.Recordset
    Dim DiasGracia As Integer, DiasDescuento As Integer, PorcDescuento As Double

On Error GoTo Error
    
    With rcParametros
        
        .Open "SELECT " & TipoTasa & " / 100 AS Tasa,Periodo,VenPeriodo,Fecha,Prestamo,Vencimiento,Promocion,Origen FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
        
        If Not .BOF And Not .EOF Then
            DiasGracia = Val(Regresa_Valor_BD("DiasGracia"))
            DiasTrans = DateDiff("D", !Fecha, Date)
            
            DiasDescuento = 0
            PorcDescuento = 0
            If DiasTrans = 0 Then DiasTrans = 1
            If !Promocion <> 0 Then
                If !Origen = 1 Then
                    rcPromocion.Open "SELECT Tipo,PorcentajeDescuento,DiasDescuento FROM Promociones WHERE ID=" & !Promocion, dbDatos, adOpenForwardOnly, adLockOptimistic
                    If Not rcPromocion.BOF And Not rcPromocion.EOF Then
                        Select Case rcPromocion!Tipo
                            Case "D"
                                DiasDescuento = IIf(DiasTrans > rcPromocion!DiasDescuento, rcPromocion!DiasDescuento, 0)
                            Case "P"
                                PorcDescuento = rcPromocion!porcentajeDescuento / 100
                            Case "A"
                                DiasDescuento = IIf(DiasTrans > rcPromocion!DiasDescuento, rcPromocion!DiasDescuento, 0)
                                PorcDescuento = rcPromocion!porcentajeDescuento / 100
                        End Select
                    End If
                End If
            End If
                                                                                        
''''            If DiasGracia > 0 Then
''''
''''                If (Date > DateValue(!Vencimiento)) And (Date <= DateAdd("D", DiasGracia, DateValue(!Vencimiento))) Then
''''                    DiasTrans = DateDiff("d", !Fecha, !Vencimiento)
''''                Else
''''                    DiasTrans = DateDiff("d", !Fecha, Date)
''''                End If
''''            End If
            '****
            Dim rc As New ADODB.Recordset
            Set rc = dbDatos.Execute("SELECT * FROM empeno WHERE ID=" & ID)
            
            If !Periodo <> 1 And rc!Serie = 2 Then
                        Dim rcParam As New ADODB.Recordset
                        Dim diasTranscurridos As Integer
                        Set rcParam = dbDatos.Execute("SELECT * FROM parametros WHERE ID=1")
                                                
                        'Tomo los días transcurridos
                        diasTranscurridos = DateDiff("d", !Vencimiento, Date)
                        MesTrancurridos = DateDiff("M", !Fecha, Date)
                        
                        
                        If Date <= DateAdd("D", rcParam!DiasGraciaAutos, DateValue(!Vencimiento)) Then
                            crIntereses = (!Prestamo * !Tasa) * IIf(MesTrancurridos = 1 Or MesTrancurridos = 0, 1, MesTrancurridos)
                        
                        ElseIf Date >= DateAdd("D", rcParam!DiasGraciaAutos, DateValue(!Vencimiento)) Then
                            crIntereses = (!Prestamo * !Tasa) + (!Prestamo * (!Tasa / !Periodo) * diasTranscurridos)
                        End If
'                        crIntereses = !Prestamo * (!Tasa / !Periodo)
'                        InteresPeriodo = (!Prestamo * !Tasa)
                        
                        'crIntereses = crIntereses * diasTranscurridos
                        
                        'Asigno interes según el periodo  Prestamo * (rcParametros!Tasa / rcParametros!Periodo) * i
'                        If diasTranscurridos <= (!Periodo + rcParam!DiasGracia) Then
'                            crIntereses = Redondeo((((!Prestamo * !Tasa))) * 1)
'                        ElseIf diasTranscurridos > (!Periodo + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 2) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 2)
'                        ElseIf diasTranscurridos > ((!Periodo * 2) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 3) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 3)
'                        ElseIf diasTranscurridos > ((!Periodo * 3) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 4) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 4)
'                        ElseIf diasTranscurridos > ((!Periodo * 4) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 5) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 5)
'                        ElseIf diasTranscurridos > ((!Periodo * 5) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 6) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 6)
'                        ElseIf diasTranscurridos > ((!Periodo * 6) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 7) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 7)
'                        ElseIf diasTranscurridos > ((!Periodo * 7) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 8) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 8)
'                        ElseIf diasTranscurridos > ((!Periodo * 8) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 9) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 9)
'                        ElseIf diasTranscurridos > ((!Periodo * 9) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 10) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 10)
'                        ElseIf diasTranscurridos > ((!Periodo * 10) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 11) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 11)
'                        ElseIf diasTranscurridos > ((!Periodo * 11) + rcParam!DiasGracia) And diasTranscurridos <= ((!Periodo * 12) + rcParam!DiasGracia) Then crIntereses = Redondeo((((!Prestamo * !Tasa))) * 12)
'                        End If
            Else
                'crIntereses = Redondeo(!Prestamo * (!Tasa / 30) * (DiasTrans - DiasDescuento)) '!Periodo
               '//////////////////roger
'                DiasTrans = DateDiff("D", !Fecha, Date)
'
'                'Checo los Dias Minimos
'                If DiasMinimo > DiasTrans Then DiasTrans = DiasMinimo
'
'                'Checo si no ha pasado la fecha de vencimiento
''                If Date > !Vencimiento Then
''                    DiasTrans = DateDiff("D", !Fecha, !Vencimiento)
''                End If
'
'                crIntereses = !Prestamo * (!Tasa / !Periodo) * (DiasTrans - DiasDescuento) '!Periodo
                '////////////////////////
               'Tomo los días transcurridos
                        diasTranscurridos = DateDiff("d", !Vencimiento, Date)
                        MesTrancurridos = DateDiff("M", !Fecha, Date)
                        
                        
                        If Date <= DateAdd("D", DiasGracia, DateValue(!Vencimiento)) Then
                            crIntereses = (!Prestamo * !Tasa) * IIf(MesTrancurridos = 1 Or MesTrancurridos = 0, 1, MesTrancurridos)
                        
                        ElseIf Date >= DateAdd("D", DiasGracia, DateValue(!Vencimiento)) Then
                            crIntereses = (!Prestamo * !Tasa) + (!Prestamo * (!Tasa / !Periodo) * diasTranscurridos)
                        End If
                crIntereses = crIntereses - (crIntereses * PorcDescuento)
            End If
            
        End If
        
        .Close
        
        Set rcParametros = Nothing
        
    End With
    
    GeneraIntereses = crIntereses
    
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcParametros = Nothing
End Function

Function GeneraInteresesPeriodoCompleto(ByVal ID As Long, ByVal TipoTasa As String, Optional ByVal VENTACLIENTE As Boolean = False, Optional ByVal Demasia As Boolean = False) As Double

    Dim crInteres As Double, i As Integer, FechaOriginal As Date, DiasGracia As Integer, Fecha As Date
    Dim rcParametros As New ADODB.Recordset

On Error GoTo Error
    
    With rcParametros
        
        .Open "SELECT " & TipoTasa & " / 100 AS Tasa,Periodo,Vencimiento,Fecha,Prestamo,Serie FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
        
        If Not .BOF And Not .EOF Then
                    
            Fecha = !Fecha: FechaOriginal = !Fecha
            If !Serie = 2 Then
                'dias de gracia de autos
                DiasGracia = Regresa_Valor_BD("DiasGraciaAuto")
            Else
                DiasGracia = Regresa_Valor_BD("DiasGracia")
            End If
            
            
            For i = 1 To 500
              
                If !Periodo = 30 Then
                    Fecha = DateAdd("M", i, FechaOriginal)
                Else
                    Fecha = DateAdd("D", !Periodo, Fecha)
                End If
                                                
                crInteres = !Prestamo * (!Tasa * i)
                
                If Demasia Then
                    If Fecha >= !Vencimiento Then GoTo Salir
                Else
                    If Date <= DateAdd("D", DiasGracia, Fecha) Then GoTo Salir
                End If
                
            Next i
        
        End If
        .Close
        Set rcParametros = Nothing
        
    End With
    
Salir:
    GeneraInteresesPeriodoCompleto = crInteres
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcParametros = Nothing
End Function


Function GeneraMoratorios(crPrestamo As Double, Moratorios As Double, Vencimiento As Date, Serie As Integer, crInteresDiario As Double) As Double
Dim Dias As Integer, crIntereses As Double, i As Integer

    Dias = 0: crIntereses = 0
    If DateAdd("D", Val(Regresa_Valor_BD("DiasGracia")), Vencimiento) < Date Then
        
        Dias = DateDiff("D", Vencimiento, Date)
        crIntereses = crPrestamo * Moratorios
    End If
    
    If Dias > 0 Then crIntereses = crIntereses + (Dias * crInteresDiario)
    GeneraMoratorios = crIntereses
End Function

Function GeneraMoratoriosFechaTentativa(crPrestamo As Double, Vencimiento As Date, FechaTentativa As Date) As Double
    
    Dim Dias As Integer, crIntereses As Double, i As Integer, Moratorios As Double
    
    Moratorios = CDbl(Regresa_Valor_BD("Operacion")) / 100
    
    Dias = 0: crIntereses = 0
    Dias = DateDiff("D", Vencimiento, FechaTentativa)
    
    If Dias > 0 Then crIntereses = (crPrestamo * Moratorios) * Dias
    
    GeneraMoratoriosFechaTentativa = crIntereses
End Function

Function Redondeo(crImporte As Currency, Optional crPauta As Double = 0.5) As Double
Dim Punto As Integer, crValor As Double, Lenght As Integer
    
    Punto = InStr(crImporte, Separador)
    
    If Punto > 0 Then
        
        crValor = Mid(crImporte, Punto, Len(Trim(crImporte)) - (Punto - 1))
        crValor = Left(crValor, 4)
    
        If Trim(Mid(crValor, 4, 1)) = "" Then
        
            GoTo 300
        ElseIf Mid(crValor, 4, 1) = 5 Then
        
            crImporte = Left(crImporte, Punto + 1) & "5"
        ElseIf Mid(crValor, 4, 1) < 5 Then
        
            crImporte = Left(crImporte, Punto + 1) & "0"
        Else
        
            crImporte = Left(crImporte, Punto + 1) + 0.1
        End If
        
    End If

300:
    Redondeo = IIf(Punto > 0, Mid(Trim(crImporte), 1, Punto + 2), crImporte)
End Function

Public Function Redondear(ByVal Numero As Double, ByVal Decimales As Double) As Double
      Redondear = Int(Numero * 10 ^ Decimales + 1 / 2) / 10 ^ Decimales
End Function

Sub SacaReporte(Fecha1 As String, Fecha2 As String, Opcion As Integer, Optional TipoPrenda As Long = 0)
Dim Fecha As String, Criterio As String, strPrenda As String, crIntereses As Double, crIvaIntereses As Double
Dim rcTmp As New ADODB.Recordset

    dbReportes.Execute "DELETE FROM Reportes"

    Select Case Opcion
    Case 1
    
        Criterio = " AND empeno.Origen=" & OD_EMPENO
        Fecha = " AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')>='" & Format(CDate(Fecha1), "YYYY/MM/DD") & "' AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')<='" & Format(CDate(Fecha2), "YYYY/MM/DD") & "'"
    Case 2
        
        Criterio = " AND empeno.Destino=" & OD_REFRENDO
        Fecha = " AND DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')>='" & Format(CDate(Fecha1), "YYYY/MM/DD") & "' AND DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')<='" & Format(CDate(Fecha2), "YYYY/MM/DD") & "'"
    Case 3
        
        Criterio = " AND empeno.Destino=" & D_DESEMPEÑO
        Fecha = " AND DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')>='" & Format(CDate(Fecha1), "YYYY/MM/DD") & "' AND DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')<='" & Format(CDate(Fecha2), "YYYY/MM/DD") & "'"
    End Select
        
    
    Select Case Opcion
    Case 1
        
        'Tipo de Prenda
        Select Case TipoPrenda
        Case 0, -1
             
            strPrenda = ""
        Case Else
            
            strPrenda = " AND detallesempeno.Tipo=" & TipoPrenda
        End Select
    
        'Contratos
        With rcTmp
            
            .Open "SELECT empeno.Fecha,COUNT(DISTINCT(empeno.ID)) AS Contratos,detallesempeno.Tipo AS TipoPrenda,SUM(detallesempeno.Cantidad) AS Prendas,SUM(detallesempeno.Peso) AS Peso,SUM(detallesempeno.Prestamo) AS Prestamo,SUM(detallesempeno.Avaluo) AS Avaluo FROM empeno INNER JOIN detallesempeno ON empeno.ID=detallesempeno.IDEmpeno " _
                & "WHERE empeno.Cancelado=0" & Criterio & Fecha & strPrenda & " GROUP BY DATE_FORMAT(Fecha,'%Y%/%m%/%d'),detallesempeno.Tipo ORDER BY DATE_FORMAT(Fecha,'%Y%/%m%/%d')", dbDatos, adOpenForwardOnly, adLockOptimistic
            
            If Not .BOF And Not .EOF Then
                    
                While Not .EOF
                    
                    dbReportes.Execute "INSERT INTO reportes (Dia,Contratos,TipoPrenda,Prendas,Peso,Avaluo,Prestamo) VALUES ('" & _
                                        Format(!Fecha, "YYYY/MM/DD") & "'," & !Contratos & "," & !TipoPrenda & "," & !Prendas & "," & !Peso & "," & !Avaluo & "," & !Prestamo & ")"
                    
                .MoveNext
                Wend
            
            End If
            
            .Close
            
        End With
    
    Case 2, 3
                      
        With rcTmp

            .Open "SELECT FechaMovimiento,COUNT(ID) AS NumContratos,SUM(Prestamo) AS Prestamo,(SUM(Intereses)+SUM(ImporteAlmacenaje)+SUM(ImporteSeguro)+SUM(ImporteMoratorios)) AS Intereses,SUM(ImporteIva) AS ImporteIva,SUM(Pago) AS Abono " _
                & "FROM empeno WHERE empeno.Cancelado=0" & Criterio & Fecha & " GROUP BY DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d') ORDER BY DATE_FORMAT(FechaMovimiento,'%Y%/%m%/%d')", dbDatos, adOpenForwardOnly, adLockOptimistic
            
            If Not .BOF And Not .EOF Then
            
                While Not rcTmp.EOF
                
                    crIntereses = IIf(IsNull(rcTmp!Intereses), 0, rcTmp!Intereses)
                    crIvaIntereses = IIf(IsNull(rcTmp!ImporteIva), 0, rcTmp!ImporteIva)
                    
                    dbReportes.Execute "INSERT INTO reportes (Dia,Contratos,Prestamo,Intereses,IvaIntereses,Abono) VALUES ('" & _
                                        Format(!FechaMovimiento, "YYYY/MM/DD") & "'," & !NumContratos & "," & !Prestamo & "," & !Intereses & "," & !ImporteIva & "," & !Abono & ")"
                
                .MoveNext
                Wend
                
            End If
            
            .Close
            
        End With
    
    End Select
    Set rcTmp = Nothing
End Sub

Function Iniciales(Nombre As String, Apellidos As String) As String
Dim Cadena As String
   
    Cadena = Mid(Nombre, 1, 1)

    If InStr(1, Nombre, " ") <> 0 Then Cadena = Cadena & Mid(Nombre, InStr(1, Nombre, " ") + 1, 1)
   
    Cadena = Cadena & Mid(Apellidos, 1, 1)

    If InStr(1, Apellidos, " ") <> 0 Then Cadena = Cadena & Mid(Apellidos, InStr(1, Apellidos, " ") + 1, 1)
      
    Iniciales = Cadena
End Function

Public Sub GeneraPagos(IDEmpeno As Long, crPrestamo As Double, Tasa As Double, Almacenaje As Double, Seguro As Double, plazo As Integer, Periodo As Integer, Fecha As Date)

Dim SaldoInsoluto As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double
Dim Vencimiento As Date, i As Integer, crSaldo As Double, crImporteTotal As Double, crPagoFijo As Double, crAmortizacion As Double, strIntervalo As String
Dim Iva As Double, crIva As Double

    Iva = CDbl(Regresa_Valor_BD("IVA")) / 100
    
    Tasa = Tasa / 100
    Almacenaje = Almacenaje / 100
    Seguro = Seguro / 100
    
    Iva = (Tasa + Almacenaje + Seguro) * Iva
    
        
    crPrestamo = crPrestamo
    SaldoInsoluto = crPrestamo
    crImporteTotal = Redondeo(Pmt((Tasa + Almacenaje + Seguro + Iva), plazo, -crPrestamo, 0, 0), 1) * plazo
    crPagoFijo = Redondeo(Pmt((Tasa + Almacenaje + Seguro + Iva), plazo, -crPrestamo, 0, 0), 2)
    crSaldo = crImporteTotal
    strIntervalo = "D"
    
    If Periodo = 30 Then
        Periodo = 1
        strIntervalo = "M"
    End If
    
    Vencimiento = DateAdd("D", IIf(strIntervalo = "D", -1, 0), Fecha)
    
    For i = 1 To plazo
        
        Vencimiento = DateAdd(strIntervalo, Periodo, Vencimiento)
        crImporteTotal = crImporteTotal
        SaldoInsoluto = IIf(i = 1, crPrestamo, SaldoInsoluto - crAmortizacion)
        
        crIntereses = Redondeo(SaldoInsoluto * Tasa)
        crAlmacenaje = Redondeo(SaldoInsoluto * Almacenaje)
        crSeguro = Redondeo(SaldoInsoluto * Seguro)
        crIva = Redondeo(SaldoInsoluto * Iva)
        
        crAmortizacion = crPagoFijo - (crIntereses + crAlmacenaje + crSeguro + crIva)
        crSaldo = crSaldo - (crIntereses + crAlmacenaje + crSeguro + crAmortizacion + crIva)
                
        dbDatos.Execute "INSERT INTO pagosfijos (IDEmpeno,NumPago,Vencimiento,Pago,Interes,Almacenaje,Seguro,Iva,Amortizacion,Saldo,Pagado) VALUES (" & _
                        IDEmpeno & "," & i & ",'" & Format(Vencimiento, "YYYY/MM/DD") & "'," & ConvMoneda(crPagoFijo) & "," & ConvMoneda(crIntereses) & "," & ConvMoneda(crAlmacenaje) & "," & ConvMoneda(crSeguro) & "," & crIva & "," & ConvMoneda(crAmortizacion) & "," & ConvMoneda(crSaldo) & ",0)"
                                
    Next i
    
End Sub

'Public Function OpcionesPago(crPrestamo As Double, crAvaluo As Double, Fecha As Date, ID As Long, TipoInteres As String, Optional Autos As Boolean = False) As Double
'
'    Dim crAlmacenaje As Double, crSeguro As Double, crInteres As Double, crIva As Double, crTotal As Double, i As Integer, FechaIni As String, PeriodosPromocion As Double, FechaOriginal As Date, NumDias As Integer
'    Dim rcParametros As New ADODB.Recordset
'
'On Error GoTo Error
'
'    dbReportes.Execute "DELETE FROM opcionpagos WHERE PC='" & NombrePc & "'"
'
'    FechaOriginal = Fecha
'
'    rcParametros.Open "SELECT Almacenaje / 100 AS Almacenaje,Seguro / 100 AS Seguro,Iva / 100 AS Iva,Tasa / 100 AS Tasa,Operacion / 100 AS Operacion,Periodo,VenPeriodo,Tipointeres,TipoTasa,Promocion FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
'
'    PeriodosPromocion = IIf(rcParametros!Periodo = 30 And rcParametros!Promocion = 15, rcParametros!Promocion / rcParametros!Periodo, Int(rcParametros!Promocion / rcParametros!Periodo))
'
'    'For i = 1 To IIf(Autos, 12, IIf(rcParametros!VenPeriodo < 3, 3, rcParametros!VenPeriodo))
'
'    'For i = 1 To rcParametros!VenPeriodo / 30
'        i = rcParametros!VenPeriodo / 30
'
'
'
'
'        crInteres = 0: crAlmacenaje = 0: crSeguro = 0: crIva = 0
'
'
'
'        'If i > rcParametros!VenPeriodo Then GoTo MeteVencimiento
'
'        'Vencimiento
'        FechaIni = Fecha
'
'          If rcParametros!TipoTasa = "MENSUAL" Then
'                i = rcParametros!VenPeriodo
'                 Fecha = DateAdd("D", rcParametros!VenPeriodo * 30, Fecha)
'          Else
'                 Fecha = DateAdd("D", rcParametros!VenPeriodo, Fecha)
'          End If
'
'        'If i > 1 Then FechaIni = DateAdd("D", 1, FechaIni)
'
'''''        If rcParametros!TipoInteres = "TRADICIONAL" Then
'''''
'''''''''            If rcParametros!Periodo = 30 Then
'''''''''
'''''''''                Fecha = DateAdd("M", i, FechaOriginal)
'''''''''                NumDias = NumDias + Regresa_Ultimo_Dia_Mes(DateAdd("M", i - 1, FechaOriginal))
'''''''''            Else
'''''
'''''                NumDias = 30 * i '(rcParametros!VenPeriodo / rcParametros!Periodo)
'''''                Fecha = DateAdd("D", (rcParametros!VenPeriodo / rcParametros!Periodo), Fecha) '
'''''''''            End If
'''''
'''''            crInteres = Redondeo(crPrestamo * ((rcParametros!Tasa / 30) * NumDias))
'''''            crAlmacenaje = Redondeo(crPrestamo * ((rcParametros!Almacenaje / 30) * NumDias))
'''''            crSeguro = Redondeo(crPrestamo * ((rcParametros!Seguro / 30) * NumDias))
'''''            crIva = Redondeo((crInteres + crAlmacenaje + crSeguro) * rcParametros!Iva)
'''''
'''''        Else
'
'            crInteres = Redondeo(crPrestamo * rcParametros!Tasa * i)
'            crAlmacenaje = Redondeo(crPrestamo * rcParametros!Almacenaje * i)
'            crSeguro = Redondeo(crPrestamo * rcParametros!Seguro * i)
'            crIva = Redondeo((crInteres + crAlmacenaje + crSeguro) * rcParametros!Iva)
'
'            'Fecha = DateAdd("D", rcParametros!Periodo, Fecha)
'
'
'''''        End If
'
'        'Tomo el Importe
'        OpcionesPago = Redondeo(crAlmacenaje + crSeguro + crInteres + crIva)
'
'        'Meto los Datos
''''''        If rcParametros!TipoTasa = "DIARIA" And (i <> 1 And i < rcParametros!VenPeriodo) Then GoTo TasaDiaria
'
'MeteVencimiento:
'
'        dbReportes.Execute "INSERT INTO opcionpagos(Vencimiento,Almacenaje,Seguro,Interes,Refrendo,IDEmpeno,Prestamo,TipoInteres,FechaIni,ImporteIva,PC) VALUES ('" & _
'                            Format(Fecha, "YYYY/MM/DD") & "'," & ConvMoneda(crAlmacenaje) & "," & ConvMoneda(crSeguro) & "," & ConvMoneda(crInteres) & "," & ConvMoneda(OpcionesPago) & "," & ID & "," & ConvMoneda(crPrestamo) & ",'" & TipoInteres & "','" & Format(FechaIni, "YYYY/MM/DD") & "'," & ConvMoneda(crIva) & ",'" & NombrePc & "')"
'TasaDiaria:
'    'Next i
'
'    rcParametros.Close
'    Set rcParametros = Nothing
'    Exit Function
'
'Error:
'    Maneja_Error Err
'    Set rcParametros = Nothing
'End Function
Public Function OpcionesPago(crPrestamo As Double, crAvaluo As Double, Fecha As Date, ID As Long, TipoInteres As String, Optional Autos As Boolean = False) As Double
Dim crAlmacenaje As Double, crSeguro As Double, crInteres As Double, crIva As Double, crTotal As Double, i As Integer, FechaIni As String, PeriodosPromocion As Double, FechaOriginal As Date
Dim rcParametros As New ADODB.Recordset

On Error GoTo Error

    dbReportes.Execute "DELETE FROM opcionpagos WHERE PC='" & NombrePc & "'"
    
    FechaOriginal = Fecha
    rcParametros.Open "SELECT Almacenaje / 100 AS Almacenaje,Seguro / 100 AS Seguro,Iva / 100 AS Iva,Tasa / 100 AS Tasa,Operacion / 100 AS Operacion,Periodo,VenPeriodo,Tipointeres,TipoTasa,Promocion FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
    PeriodosPromocion = IIf(rcParametros!Periodo = 30 And rcParametros!Promocion = 15, rcParametros!Promocion / rcParametros!Periodo, Int(rcParametros!Promocion / rcParametros!Periodo))
    For i = 1 To rcParametros!VenPeriodo
        
        crInteres = 0: crAlmacenaje = 0: crSeguro = 0: crIva = 0
        If rcParametros!Periodo = 30 Then
            
            Fecha = DateAdd("M", i, FechaOriginal)
        Else
            
            Fecha = DateAdd("D", rcParametros!Periodo * i, FechaOriginal)
        End If
        
        'Interes
        crInteres = Redondeo(crPrestamo * ((rcParametros!Tasa) * i))
        
        'Almacenaje
        crAlmacenaje = Redondeo(crPrestamo * (((rcParametros!Almacenaje / 30) * rcParametros!Periodo)))  'Redondeo(crPrestamo * ((rcParametros!Almacenaje) * i))
        
        'Seguro
        crSeguro = Redondeo(crPrestamo * (((rcParametros!Seguro / 30) * rcParametros!Periodo)))  '0
        
        'Iva
        crIva = Redondeo((crInteres + crAlmacenaje + crSeguro) * rcParametros!Iva)
        
        'Tomo el Importe
        OpcionesPago = Redondeo(crAlmacenaje + crSeguro + crInteres + crIva)
        
        'Meto los Datos
        If rcParametros!TipoTasa = "DIARIA" And (i <> 1 And i < rcParametros!VenPeriodo) Then GoTo TasaDiaria

MeteVencimiento:
        
        dbReportes.Execute "INSERT INTO opcionpagos(Vencimiento,Almacenaje,Seguro,Interes,Refrendo,IDEmpeno,Prestamo,TipoInteres,ImporteIva,PC) VALUES ('" & _
                            Format(Fecha, "YYYY/MM/DD") & "'," & ConvMoneda(crAlmacenaje) & "," & ConvMoneda(crSeguro) & "," & ConvMoneda(crInteres) & "," & ConvMoneda(OpcionesPago) & "," & ID & "," & ConvMoneda(crPrestamo) & ",'" & TipoInteres & "'," & ConvMoneda(crIva) & ",'" & NombrePc & "')"
TasaDiaria:
    Next i
    rcParametros.Close
    Set rcParametros = Nothing
    Exit Function

Error:
    Maneja_Error Err
    Set rcParametros = Nothing
End Function

Public Function OpcionesPagoAutos(crPrestamo As Double, crAvaluo As Double, Fecha As Date, ID As Long, TipoInteres As String, Optional Autos As Boolean = False) As Double
Dim crAlmacenaje As Double, crSeguro As Double, crInteres As Double, crIva As Double, i As Integer, FechaIni As String, FechaOriginal As Date
Dim rcParametros As New ADODB.Recordset

On Error GoTo Error

    dbReportes.Execute "DELETE FROM opcionpagos WHERE PC='" & NombrePc & "'"
    FechaOriginal = Fecha
    rcParametros.Open "SELECT Almacenaje / 100 AS Almacenaje,Seguro / 100 AS Seguro,Iva / 100 AS Iva,Tasa / 100 AS Tasa,Operacion / 100 AS Operacion,Periodo,VenPeriodo,Tipointeres,TipoTasa,Promocion FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
    For i = 1 To rcParametros!VenPeriodo
        
        crInteres = 0: crAlmacenaje = 0: crSeguro = 0: crIva = 0
        
        If i > rcParametros!VenPeriodo Then GoTo MeteVencimiento
        
        'Vencimiento
        FechaIni = Fecha
        If i > 1 Then FechaIni = DateAdd("D", 1, FechaIni)
        If rcParametros!Periodo = 30 Then
            
            Fecha = DateAdd("M", i, FechaOriginal)
        Else
            
            Fecha = DateAdd("D", rcParametros!Periodo, Fecha)
        End If
        
        'Interes
        crInteres = Redondeo(crPrestamo * (rcParametros!Tasa * i))
        
        'Almacenaje
        crAlmacenaje = Redondeo(crPrestamo * (rcParametros!Almacenaje * i))
        
        'Seguro
        crSeguro = Redondeo(crPrestamo * (rcParametros!Seguro * i))
        
        'Iva
        crIva = Redondeo((crInteres + crAlmacenaje + crSeguro) * rcParametros!Iva)
        
        'Tomo el Importe
        OpcionesPagoAutos = crAlmacenaje + crSeguro + crInteres + crIva
        
        'Meto los Datos
        If rcParametros!TipoTasa = "DIARIA" And (i <> 1 And i < rcParametros!VenPeriodo) Then GoTo TasaDiaria

MeteVencimiento:
        
        dbReportes.Execute "INSERT INTO opcionpagos(Vencimiento,Almacenaje,Seguro,Interes,Refrendo,IDEmpeno,Prestamo,TipoInteres,FechaIni,ImporteIva,PC) VALUES ('" & _
                            Format(Fecha, "YYYY/MM/DD") & "'," & ConvMoneda(crAlmacenaje) & "," & ConvMoneda(crSeguro) & "," & ConvMoneda(crInteres) & "," & ConvMoneda(OpcionesPagoAutos) & "," & ID & "," & ConvMoneda(crPrestamo) & ",'" & TipoInteres & "','" & Format(FechaIni, "YYYY/MM/DD") & "'," & ConvMoneda(crIva) & ",'" & NombrePc & "')"
TasaDiaria:
    Next i
    
    rcParametros.Close
    Set rcParametros = Nothing
    Exit Function

Error:
    Maneja_Error Err
    Set rcParametros = Nothing
End Function

Function LocalizaImpresora(strImpresora As String) As Boolean
Dim prt As Printer
    
    LocalizaImpresora = False
    
    strNombreImp = ""
    strDriverImp = ""
    strPuertoImp = ""
    If strImpresora <> "" Then
        
        For Each prt In Printers
    
            If InStr(1, UCase(prt.DeviceName), UCase(strImpresora)) > 0 Then
                strNombreImp = prt.DeviceName
                strDriverImp = prt.DriverName
                strPuertoImp = prt.Port
                LocalizaImpresora = True
                Exit For
            End If
    
        Next prt
    
    End If
    
    Set prt = Nothing
End Function

Sub ModificaODBC(strODBC As String, strServidor As String, strDatos As String)
Dim dl As Long, sAttributes As String, sDriver As String, sDsnName As String

    ' Establecemos los atributos necesarios
    sDsnName = strODBC
    sDriver = "MySQL ODBC 3.51 Driver"

    ' Los pares de cadenas acabarán en valor Null
    sAttributes = "DSN=" & sDsnName & Chr(0)
    sAttributes = sAttributes & "Server=" & strServidor & Chr$(0)
    sAttributes = sAttributes & "User=" & USERBD & Chr$(0)
    sAttributes = sAttributes & "Password=" & PWDBD & Chr$(0)
    sAttributes = sAttributes & "Database=" & strDatos & Chr(0)

    ' Modificamos el origen de datos de usuario especificado
    dl = SQLConfigDataSource(0&, ODBC_CONFIG_SYS_DSN, sDriver, sAttributes)
    
End Sub

Function ChecaPromocion(IDEmpeno As Long, DiasGracia As Integer, Tasa As String, Optional DiasTrans As Integer = 0, Optional Pestana As Boolean = False) As Double
Dim rcEmpeno As New ADODB.Recordset
Dim Fecha As Date, crImporte As Double, Dias As Integer, Descuento As Double

    ChecaPromocion = False
    With rcEmpeno
        
        .Open "SELECT Fecha,Prestamo,Vencimiento,TipoTasa,VenPeriodo,Periodo,Promocion,((Tasa+Almacenaje+Seguro) * (1+(IVA/100))) AS TasaGlobal,(" & Tasa & "/100) AS Tasa FROM empeno WHERE ID=" & IDEmpeno, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not .BOF And Not .EOF Then
            
            crImporte = 0
            Fecha = Format(DateAdd("D", IIf(!TipoTasa = "MENSUAL", 0, -1), !Fecha), "DD/MM/YYYY")
            
            Select Case !Promocion
            Case 15, 30
                
'''''                Dias = DateDiff("D", !Vencimiento, Date)
'''''                If Dias > 0 Then crImporte = (!Prestamo * (!Tasa / !Periodo)) * Dias
                            
            Case 2, 3, 4, 5
                
                Select Case !Promocion
                Case 2 '5% de Descuento
                    
                    Descuento = 0.05
                Case 3 '10% de Descuento
                    
                    Descuento = 0.1
                Case 4 '15% de Descuento
                    
                    Descuento = 0.15
                Case 5 '20% de Descuento
                    
                    Descuento = 0.2
                
                End Select
                
                Fecha = DateAdd(IIf(!TipoTasa = "MENSUAL", "M", "D"), IIf(!TipoTasa = "MENSUAL", 1, IIf(!TipoTasa = "QUINCENAL", 15 * 2, 7 * 4)), Fecha)
                Dias = IIf(DiasTrans > 0, DiasTrans, DateDiff("D", !Fecha, Fecha))
                crImporte = (!Prestamo * ((!Tasa / !Periodo) * Dias)) * Descuento
                        
            Case 20, 50
                
                crImporte = ((!Tasa * 100) / !TasaGlobal) * !Promocion
            End Select
        
        End If
        
        ChecaPromocion = crImporte
        .Close
        Set rcEmpeno = Nothing
        
    End With
    
End Function

'Imprimimos la boleta
Public Sub Imprimir_Boleta_CR(ID As Long, Optional Reimpresion As Boolean = False)

    Dim i As Integer, crIntereses As Double, Contrato As String, CAT As Double
    Dim Meses As Integer, ImpresoraTickets As Boolean, NextVencimiento As String
    Dim rcConsulta As New ADODB.Recordset

On Error GoTo Error

    dbReportes.Execute "DELETE FROM articulos"
    
    rcConsulta.Open "SELECT Tipo,Articulo,Cantidad,Kilates,Peso,Estado,Avaluo,Prestamo,Observaciones,PesoPiedras FROM detallesempeno WHERE IDEmpeno=" & ID, dbDatos, adOpenForwardOnly, adLockReadOnly
    While Not rcConsulta.EOF
        
        dbReportes.Execute "INSERT INTO articulos (IDEmpeno,Tipo,Articulo,Cantidad,Kilates,Peso,Claridad,Avaluo,Prestamo,Observaciones,PesoPiedra) VALUES (" & _
                            ID & "," & rcConsulta!Tipo & ",'" & rcConsulta!Articulo & "'," & rcConsulta!Cantidad & "," & rcConsulta!Kilates & "," & rcConsulta!Peso & ",'" & rcConsulta!Estado & "'," & rcConsulta!Avaluo & "," & rcConsulta!Prestamo & ",'" & rcConsulta!Observaciones & "'," & rcConsulta!PesoPiedras & ")"
    
    i = i + 1
    rcConsulta.MoveNext
    Wend
    rcConsulta.Close
    Set rcConsulta = Nothing
    
    For i = i To 7 '8
            
        dbReportes.Execute "INSERT INTO articulos (IDEmpeno) VALUES (" & ID & ")"
    Next i
                       
    'Leo los datos del empeño
    rcConsulta.Open "SELECT ID,Prestamo,Tasa,Iva,Almacenaje,Seguro,Periodo,Avaluo,Folio,Fecha,NumContrato,Serie,TipoInteres,TipoTasa,Vencimiento,VenPeriodo FROM empeno WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockReadOnly
    
    
    ' PAGOS FIJOS
     If rcConsulta!TipoInteres = "FIJA" Then
                  
        Select Case rcConsulta!TipoTasa
        Case "MENSUAL"
              
            Meses = 1
        Case "QUINCENAL"
              
            Meses = 1
        Case "SEMANAL"
              
            Meses = 1
        End Select
        
        If Reimpresion = False Then
            
            'Intereses
            GeneraPagos rcConsulta!ID, rcConsulta!Prestamo, rcConsulta!Tasa, rcConsulta!Almacenaje, rcConsulta!Seguro, rcConsulta!VenPeriodo * Meses, rcConsulta!Periodo, rcConsulta!Fecha
            Sleep 500
        
        End If
        
        'Próximo Vencimiento
        NextVencimiento = SacaValor("pagosfijos", "Vencimiento", " WHERE ID = " & Val(SacaValor("pagosfijos", "MIN(ID)", " WHERE IDEmpeno = " & rcConsulta!ID)))
        
        ImpresoraTickets = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
        
        'Imprimo el Calendario de pagos
        With frmMDI.Cr
            .Reset
            .DiscardSavedData = True
            .WindowShowPrintSetupBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\TicketPagosFijos.rpt"
            .SelectionFormula = "{empeno.ID}=" & rcConsulta!ID & ""
            .Formulas(0) = "NumPagos=" & rcConsulta!VenPeriodo * Meses & ""
            .Formulas(1) = "Enajenacion=" & Regresa_Valor_BD("DiasEnajenacion") & ""
            .Formulas(2) = "Notas='" & Trim(Regresa_Valor_BD("Notas")) & "'"
            .Formulas(3) = "ProximoVencimiento='" & Format(CDate(NextVencimiento), "DD-MMM-YYYY") & "'"
            .Destination = crptToWindow
            
            'La mando a la impresora por default
            If ImpresoraTickets Then
                .PrinterName = strNombreImp
                .PrinterDriver = strDriverImp
                .PrinterPort = strPuertoImp
                .Destination = crptToPrinter
            End If
        
            .WindowTitle = "Calendario Pagos"
            .WindowState = crptMaximized
            .Action = 1
        End With
        
        crIntereses = SacaValor("pagosfijos", "Sum(Pago)", " WHERE IDEmpeno = " & rcConsulta!ID) - rcConsulta!Prestamo
        
    Else
    
    'Opciones de Pago
        crIntereses = OpcionesPago(rcConsulta!Prestamo, rcConsulta!Avaluo, rcConsulta!Fecha, ID, rcConsulta!TipoTasa)
    
    End If
    'Saco el CAT
    CAT = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Cat", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie=" & IIf(rcConsulta!Serie = SERIE_A Or rcConsulta!Serie = SERIE_C, SERIE_A, rcConsulta!Serie) & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & rcConsulta!VenPeriodo))

    Contrato = rcConsulta!NumContrato
    
    For i = 1 To 6 - Len(Contrato)
        Contrato = "0" & Contrato
    Next i
    
    '::: 31-OCT-2011 ::: Rellenar la tabla de opcion pago
    Dim rc As New ADODB.Recordset
    Dim Ren As Byte
    rc.Open "SELECT COUNT(ID) As Ren FROM opcionpagos", dbReportes, adOpenForwardOnly, adLockReadOnly
    If Not rc.EOF Then
        Ren = rc.Fields(0)
    End If
    rc.Close
    Set rc = Nothing
    For i = Ren To 5
        dbReportes.Execute "INSERT INTO opcionpagos (IDEmpeno,PC) VALUES (" & ID & ",'" & NombrePc & "')"
    Next i
    '::::
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\Boleta.rpt"
        .SelectionFormula = "{empeno.ID}=" & ID
        .Formulas(0) = "CodigoBarras='*" & Contrato & "*'"
        .Formulas(1) = "GastosVenta=" & Regresa_Valor_BD("GtosVenta") & ""
        .Formulas(2) = "Cat=" & CAT
        .Formulas(3) = "ImporteRefrendo=" & crIntereses
        .Formulas(4) = "CantidadLetra='" & CantidadEnLetra(rcConsulta!Prestamo) & "'"
        .Formulas(5) = "FechaComercializacion='" & Format(DateAdd("D", Regresa_Valor_BD("DiasEnajenacion") + 1, rcConsulta!Vencimiento), "DD/MMM/YYYY") & "'"
        .Formulas(6) = "FechaFiniquito='" & Format(DateAdd("D", Regresa_Valor_BD("DiasGracia"), rcConsulta!Vencimiento), "DD/MMM/YYYY") & "'"
        .Formulas(7) = "RazonSocial='" & Sucursal.RazonSocial & "'"
        .Formulas(8) = "DireccionSuc='" & Sucursal.Direccion & " " & Sucursal.Ciudad & " " & Sucursal.Estado & "'"
        .Formulas(9) = "RfcSuc='" & Sucursal.RFC & "'"
        .Formulas(10) = "Penalizacion=" & Regresa_Valor_BD("Operacion") & ""
        .Formulas(11) = "Reposicion=" & Regresa_Valor_BD("ImportePerdida") & ""
        
        .Formulas(12) = "Horario='" & Regresa_Valor_BD("HorarioSucursal") & "'"
        
        If rcConsulta!TipoInteres = "FIJA" Then .Formulas(13) = "PagoFijo= 1"
        
        .SubreportToChange = "OpcionPago2"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .DiscardSavedData = True
        
        .SubreportToChange = "OpcionPagos"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{opcionpagos.PC}='" & NombrePc & "'"
        .DiscardSavedData = True
        
        .SubreportToChange = "Articulos"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .DiscardSavedData = True
        
        .SubreportToChange = "detalles_articulo"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{detallesempeno.IDEmpeno}=" & ID
        .DiscardSavedData = True
        
        .WindowTitle = "Contrato"
        .WindowState = crptMaximized
        .Action = 1
    End With
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub
Public Sub Imprimir_Boleta_CR_Caidas(ID As Long, Optional Reimpresion As Boolean = False, Optional Etiqueta As Boolean = False, Optional Auto As Boolean = False)

    Dim ImpresoraEtiquetas As Printer
    Dim ImpresoraContratos As Printer
    Dim rcConsulta As New ADODB.Recordset
    Dim rcAux As New ADODB.Recordset
    Dim rc As New ADODB.Recordset
    Dim i As Integer, x As Integer, crIntereses As Double, Contrato As String, strDescripcion, DescPrendaY As Double, PosicionY As Double, PesoTotal As Double
    Dim TotalAvaluo As Double, TotalPrestamo As Double, TotalPeso As Double, Serie As Integer, NumRefrendos As Integer, Tipo As String
    Dim Meses As Integer, ImpresoraTickets As Boolean, NextVencimiento As String, DiasGracia As Integer, diasEnajenacion As Integer
    Dim Impresiones As Integer
    Dim CAT As Double
    

On Error GoTo Error
        
    rcConsulta.Open "SELECT e.Fecha,e.Prestamo,e.Avaluo,e.NumContrato,e.Folio,e.Responsable,e.Beneficiario,e.Serie,e.TipoInteres,e.TipoTasa,e.Tasa,e.Almacenaje,e.Iva,e.VenPeriodo,e.Periodo,e.Vencimiento,e.Valuador,e.NumBolsa,c.Id as IdCliente,c.IDTipoIdent " & _
                    ",e.NumIdentBeneficiario, CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,c.Identificacion,c.NumeroIdentificacion,c.Direccion,c.Tel,c.Email  " & _
                    "FROM empeno e INNER JOIN clientes c ON e.IDCliente = c.ID WHERE e.ID = " & ID, dbDatos, adOpenForwardOnly, adLockReadOnly

    Set rc = dbDatos.Execute("SELECT * FROM clientes WHERE ID=" & rcConsulta!IDCliente)
    
    'Si es contrato de pagos fijos
    If rcConsulta!TipoInteres = "FIJA" Then
                  
        Select Case rcConsulta!TipoTasa
        Case "MENSUAL"
              
            Meses = 1
        Case "QUINCENAL"
              
            Meses = 2
        Case "SEMANAL"
              
            Meses = 4
        End Select
        
        If Reimpresion = False Then
            
            'Intereses
            GeneraPagos rcConsulta!ID, rcConsulta!Prestamo, rcConsulta!Tasa * (1 + (rcConsulta!Iva / 100)), rcConsulta!Almacenaje * (1 + (rcConsulta!Iva / 100)), rcConsulta!Seguro * (1 + (rcConsulta!Iva / 100)), rcConsulta!VenPeriodo * Meses, rcConsulta!Periodo, rcConsulta!Fecha
            Sleep 500
        
        End If
        
        'Próximo Vencimiento
        NextVencimiento = SacaValor("pagosfijos", "Vencimiento", " WHERE ID = " & Val(SacaValor("pagosfijos", "MIN(ID)", " WHERE IDEmpeno = " & rcConsulta!ID)))
        
        ImpresoraTickets = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
        
        'Imprimo el Calendario de pagos
        With frmMDI.Cr
            .Reset
            .DiscardSavedData = True
            .WindowShowPrintSetupBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\TicketPagosFijos.rpt"
            .SelectionFormula = "{empeno.ID}=" & rcConsulta!ID & ""
            .Formulas(0) = "NumPagos=" & rcConsulta!VenPeriodo * Meses & ""
            .Formulas(1) = "Enajenacion=" & Regresa_Valor_BD("DiasEnajenacion") & ""
            .Formulas(2) = "Notas='" & Trim(Regresa_Valor_BD("Notas")) & "'"
            .Formulas(3) = "ProximoVencimiento='" & Format(CDate(NextVencimiento), "DD-MMM-YYYY") & "'"
            .Destination = crptToWindow
            
            'La mando a la impresora por default
            If ImpresoraTickets Then
                .PrinterName = strNombreImp
                .PrinterDriver = strDriverImp
                .PrinterPort = strPuertoImp
                .Destination = crptToPrinter
            End If
        
            .WindowTitle = "Calendario Pagos"
            .WindowState = crptMaximized
            .Action = 1
        End With
        
    Else
        
        'Opciones de Pago
        crIntereses = OpcionesPago(rcConsulta!Prestamo, rcConsulta!Avaluo, rcConsulta!Fecha, ID, rcConsulta!TipoTasa)

    End If
                    
    'Saco los Dias de Gracia
    DiasGracia = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres = ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo = tp.ID INNER JOIN plazos p ON ct.IDPlazo = p.ID", "DGracia", " WHERE ti.Descripcion = '" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo)))
                                                                 
    'Saco los dias de Enajenación
    diasEnajenacion = Regresa_Valor_BD("DiasEnajenacion")
    
    'Tomo el Número de Contrato
    Contrato = rcConsulta!NumContrato
    
    For i = 1 To 6 - Len(Contrato)
        Contrato = "0" & Contrato
    Next i
                
    Regresa_Impresora Contratos, ImpresoraContratos

    For Impresiones = 1 To IIf(Reimpresion = True, 1, 2)
        
        With ImpresoraContratos
        
            .ScaleMode = vbMillimeters
            .FontBold = True
            .Font = "Arial Narrow"
            .FontSize = 20
            
            'Número de Contrato
            .CurrentX = Regresa_Valor("CONTRATO", "NumContratoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "NumContratoY", 0)
            ImpresoraContratos.Print rcConsulta!NumContrato
            
            .FontBold = False
            .FontSize = 7
            
            'fecha de contrato
            .CurrentX = Regresa_Valor("CONTRATO", "fechaContratoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "fechaContratoY", 0)
            ImpresoraContratos.Print rcConsulta!Fecha
            
            'NumRefrendos
            NumRefrendos = NumeroRefrendos(rcConsulta!NumContrato, rcConsulta!Serie)
            .CurrentX = Regresa_Valor("CONTRATO", "NumRefrendosX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "NumRefrendosY", 0)
            ImpresoraContratos.Print IIf(NumRefrendos > 0, "R / " & NumRefrendos, "")
            
            'Imprimo el Cliente
            .CurrentX = Regresa_Valor("CONTRATO", "ClienteX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ClienteY", 0)
            ImpresoraContratos.Print rcConsulta!Cliente
            
            'Imprimo el Cliente Email
            .CurrentX = Regresa_Valor("CONTRATO", "EmailuX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "EmailuY", 0)
            ImpresoraContratos.Print rcConsulta!Email
            
            'Imprimo el Cliente
            .CurrentX = Regresa_Valor("CONTRATO", "ClienteDesempeñoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ClienteDesempeñoY", 0)
            ImpresoraContratos.Print rcConsulta!Cliente
            
            'Imprimo el Cliente
            .CurrentX = Regresa_Valor("CONTRATO", "ClienteConsumidorX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ClienteConsumidorY", 0)
            ImpresoraContratos.Print rcConsulta!Cliente
            
            'Identificacion
            .CurrentX = Regresa_Valor("CONTRATO", "IdentificacionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "IdentificacionY", 0)
            If IsNull(rcConsulta!Identificacion) = False Then
                ImpresoraContratos.Print UCase(rcConsulta!Identificacion)
            Else
                ImpresoraContratos.Print UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE Id=" & rcConsulta!IDTipoIdent))
            End If
            
            'Número Identificacion
            .CurrentX = Regresa_Valor("CONTRATO", "NumIdentificacionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "NumIdentificacionY", 0)
            ImpresoraContratos.Print rcConsulta!NumeroIdentificacion
            
            'Direccion
            .CurrentX = Regresa_Valor("CONTRATO", "DireccionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DireccionY", 0)
            'IIf(IsNull(), 0, )
            ImpresoraContratos.Print rcConsulta!Direccion & " " & IIf(IsNull(rc!NoExterior), 0, rc!NoExterior) & " " & IIf(IsNull(rc!NoInterior), 0, rc!NoInterior) & " COL: " & IIf(IsNull(rc!Colonia), 0, rc!Colonia) & " CP: " & IIf(IsNull(rc!CP), 0, rc!CP) & " " & IIf(IsNull(rc!Municipio), 0, rc!Municipio) & " " & IIf(IsNull(rc!Estado), 0, rc!Estado)
            
             'Tel cliente
            .CurrentX = Regresa_Valor("CONTRATO", "telClienteX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "telClienteY", 0)
            ImpresoraContratos.Print rcConsulta!Tel
            
            'Cotitular
            .CurrentX = Regresa_Valor("CONTRATO", "CotitularX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "CotitularY", 0)
            ImpresoraContratos.Print rcConsulta!Responsable
            
            'Beneficiario
            .CurrentX = Regresa_Valor("CONTRATO", "BeneficiarioX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "BeneficiarioY", 0)
            ImpresoraContratos.Print rcConsulta!Beneficiario
            
            'Beneficiario
''            .CurrentX = Regresa_Valor("CONTRATO", "NumIdentBeneficiarioX", 0)
''            .CurrentY = Regresa_Valor("CONTRATO", "NumIdentBeneficiarioY", 0)
''            ImpresoraContratos.Print "N° IDENT: " & rcConsulta!NumIdentBeneficiario
            
           CAT = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Cat", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo)))
            'CAT
            .CurrentX = Regresa_Valor("CONTRATO", "CATX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "CATY", 0)
            ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Cat", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
            
            'Tasa de Interes Anual
            .CurrentX = Regresa_Valor("CONTRATO", "TasaInteresAnualX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "TasaInteresAnualY", 0)
            ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
            
            'costo total mensual
            .CurrentX = Regresa_Valor("CONTRATO", "costoMensualX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "costoMensualY", 0)
            ImpresoraContratos.Print Format(CAT / 12, FMoneda)
            
             'costo total diario
            .CurrentX = Regresa_Valor("CONTRATO", "costoDiarioX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "costoDiarioY", 0)
            ImpresoraContratos.Print Format((CAT / 12) / 30, FMoneda)
            
            'Monto del Prestamo
            .CurrentX = Regresa_Valor("CONTRATO", "PrestamoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "PrestamoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Prestamo, FMoneda)
                            
            'Monto Total a Pagar
            .CurrentX = Regresa_Valor("CONTRATO", "MontoTotalPagarX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MontoTotalPagarY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Prestamo + crIntereses, FMoneda)
            
            'Almacenaje CAT
            .CurrentX = Regresa_Valor("CONTRATO", "AlmacenajeCATX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "AlmacenajeCATY", 0)
            ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Almacenaje", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))) * IIf(rcConsulta!TipoTasa = "MENSUAL", 1, IIf(rcConsulta!TipoTasa = "QUINCENAL", 2, IIf(rcConsulta!TipoTasa = "SEMANAL", 4, 30))), "0.00")
    
            'Comercialización
            .CurrentX = Regresa_Valor("CONTRATO", "ComercializacionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ComercializacionY", 0)
            ImpresoraContratos.Print Format(Regresa_Valor_BD("GtosVenta"), "0.00")
    
            'ReposicionContrato
            .CurrentX = Regresa_Valor("CONTRATO", "ReposicionContratoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ReposicionContratoY", 0)
            ImpresoraContratos.Print Format(Regresa_Valor_BD("ImportePerdida"), FMoneda)
    
            'Desempeño Extemporaneo
            .CurrentX = Regresa_Valor("CONTRATO", "DesempenoExtemporaneoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DesempenoExtemporaneoY", 0)
            ImpresoraContratos.Print Format(Regresa_Valor_BD("Operacion"), "0.00")
    
            'PlazoPrestamo
            .CurrentX = Regresa_Valor("CONTRATO", "PlazoPrestamoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "PlazoPrestamoY", 0)
            ImpresoraContratos.Print rcConsulta!VenPeriodo & " " & IIf(rcConsulta!TipoTasa = "MENSUAL" And rcConsulta!VenPeriodo > 1, "MESES", IIf(rcConsulta!TipoTasa = "MENSUAL" And rcConsulta!VenPeriodo = 1, "MES", IIf(rcConsulta!TipoTasa = "QUINCENAL", "QUINCENAS", IIf(rcConsulta!TipoTasa = "SEMANAL", "SEMANAS", "DIAS"))))
            
            'Opciones de pago
            
         .FontSize = 7
            PosicionY = Regresa_Valor("CONTRATO", "OpcionPagoY", 0)
            If rcConsulta!TipoInteres = "FIJA" Then
    
                .CurrentX = Regresa_Valor("CONTRATO", "OpcionPagoX", 0)
                .CurrentY = PosicionY
                ImpresoraContratos.Print "CONTRATO DE PAGOS FIJOS. IMPORTES A PAGAR EN EL TICKET ANEXO A ESTE CONTRATO"
    
            Else
                
                i = 1
                rcAux.Open "SELECT * FROM opcionpagos WHERE IDEmpeno=" & ID & " AND PC='" & NombrePc & "' ORDER BY ID", dbReportes, adOpenForwardOnly, adLockReadOnly
                Do While Not rcAux.EOF
    
                    'Opcion Numero
                    .CurrentX = Regresa_Valor("CONTRATO", "OpcionPagoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print i
    
                    'ImporteMutuo
                    .CurrentX = Regresa_Valor("CONTRATO", "ImporteMutuoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Prestamo, 15, True)
    
                    'MontoIntereses
                    .CurrentX = Regresa_Valor("CONTRATO", "MontoInteresesX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Interes, 15, True)
    
                    'MontoAlmacenaje
                    .CurrentX = Regresa_Valor("CONTRATO", "MontoAlmacenajeX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Almacenaje, 15, True)
    
                    'ImporteIva
                    .CurrentX = Regresa_Valor("CONTRATO", "ImporteIvaX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!ImporteIva, 15, True)
    
                    'Por Refrendo
                    .CurrentX = Regresa_Valor("CONTRATO", "PagoRefrendoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Interes + rcAux!Almacenaje + rcAux!ImporteIva, 15, True)
    
                    'Por Desempeno
                    .CurrentX = Regresa_Valor("CONTRATO", "PagoDesempenoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Prestamo + rcAux!Interes + rcAux!Almacenaje + rcAux!ImporteIva, 15, True)
    
                    'Vencimiento
                    .CurrentX = Regresa_Valor("CONTRATO", "PagoVencimientoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print Format(rcAux!Vencimiento, "DD/MMM/YYYY")
    
                i = i + 1
                PosicionY = PosicionY + 4
                rcAux.MoveNext
                Loop
            
            End If
            rcAux.Close
            
'************************************************************************************************************************************************************************
            'Imprimo la descripción de las prendas
            .FontSize = 7
            
            If rcConsulta!Serie = 1 Then
                
                PesoTotal = 0
                TotalAvaluo = 0
                TotalPrestamo = 0
                
                DescPrendaY = Regresa_Valor("CONTRATO", "DescripcionPrendasY", 0)
                
            
                rcAux.Open "SELECT tipo.Descripcion AS Desc_Tipo,d.Cantidad,d.Articulo AS Des_Prenda,d.CantidadPiedras,d.PesoPiedras,kilatajes.Descripcion AS Desc_Kilates,d.Peso,d.Estado AS Desc_Estado,d.Avaluo,d.Prestamo,d.Observaciones,d.Marca,d.Modelo,d.Serie,d.Tamano,d.Color " & _
                        "FROM detallesempeno d LEFT JOIN tipo ON d.Tipo=tipo.ID LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE d.IDEmpeno=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
                
                
                While Not rcAux.EOF
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "DescripcionPrendasX", 0)
                    .CurrentY = DescPrendaY
                    ImpresoraContratos.Print rcAux!Desc_Tipo
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "PiedraPesoX", 0)
                    .CurrentY = DescPrendaY
                    ImpresoraContratos.Print IIf(rcAux!PesoPiedras <= 0, "", rcAux!PesoPiedras)
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "KilatajeX", 0)
                    .CurrentY = DescPrendaY
                    ImpresoraContratos.Print IIf(IsNull(rcAux!Desc_Kilates), "", rcAux!Desc_Kilates)
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "PesoX", 0)
                    .CurrentY = DescPrendaY
                    ImpresoraContratos.Print IIf(rcAux!Peso <= 0, "", Format(rcAux!Peso, "0.00"))
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "CalidadX", 0)
                    .CurrentY = DescPrendaY
                    ImpresoraContratos.Print IIf(rcAux!Desc_Estado <= 0, "", rcAux!Desc_Estado)
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "AvaluoPrendaX", 0)
                    .CurrentY = DescPrendaY
                    ImpresoraContratos.Print "$" & Format(rcAux!Avaluo, FMoneda)
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "PrestamoPrendaX", 0)
                    .CurrentY = DescPrendaY
                    ImpresoraContratos.Print "$" & Format(rcAux!Prestamo, FMoneda)
                    
                    
                                'PorcenPrestamoAvaluo
                    .CurrentX = Regresa_Valor("CONTRATO", "PorcenPrestamoAvaluoUX", 0)
                    .CurrentY = Regresa_Valor("CONTRATO", "PorcenPrestamoAvaluoUY", 0)
                    ImpresoraContratos.Print Format(Round((rcAux!Prestamo * 100) / rcAux!Avaluo, 1), "0.00")
                    
                    strDescripcion = ""
                    x = 0
                    For i = 1 To Len(rcAux!Des_Prenda & " " & rcAux!Observaciones & IIf(IsNull(rcAux!Marca) Or Trim(rcAux!Marca) = "", "", " MARCA: " & rcAux!Marca) & IIf(IsNull(rcAux!Modelo) Or Trim(rcAux!Modelo) = "", "", " MODELO: " & rcAux!Modelo) & IIf(IsNull(rcAux!Serie) Or Trim(rcAux!Serie) = "", "", " SERIE: " & rcAux!Serie) & IIf(IsNull(rcAux!Tamano) Or Trim(rcAux!Tamano) = "", "", " TAMAÑO: " & rcAux!Tamano) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color)) Step 40
                        
                        .CurrentX = Regresa_Valor("CONTRATO", "ReferenciasX", 0)
                        .CurrentY = DescPrendaY + (2.5 * x)
                        
                        strDescripcion = LTrim(Mid(rcAux!Des_Prenda & " " & rcAux!Observaciones & IIf(IsNull(rcAux!Marca) Or Trim(rcAux!Marca) = "", "", " MARCA: " & rcAux!Marca) & IIf(IsNull(rcAux!Modelo) Or Trim(rcAux!Modelo) = "", "", " MODELO: " & rcAux!Modelo) & IIf(IsNull(rcAux!Serie) Or Trim(rcAux!Serie) = "", "", " SERIE: " & rcAux!Serie) & IIf(IsNull(rcAux!Tamano) Or Trim(rcAux!Tamano) = "", "", " TAMAÑO: " & rcAux!Tamano) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color), i * 1, 40 + IIf(Mid(strDescripcion, i, 1) = " ", 1, 0)))
                        ImpresoraContratos.Print strDescripcion
                        x = x + 1
                    Next i
                                                       
                TotalAvaluo = TotalAvaluo + rcAux!Avaluo
                TotalPrestamo = TotalPrestamo + rcAux!Prestamo
                DescPrendaY = DescPrendaY + (2.5 * (x - 1)) + 2.5
                PesoTotal = PesoTotal + (rcAux!Peso * rcAux!Cantidad)
                rcAux.MoveNext
                Wend
                rcAux.Close
                
                'Total Peso
                .CurrentX = Regresa_Valor("CONTRATO", "TotalPesoX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "TotalPesoY", 0)
                ImpresoraContratos.Print Format(PesoTotal, "0.00")
                
                'Total Avaluo Prendas
                .CurrentX = Regresa_Valor("CONTRATO", "TotalAvaluoPrendaX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "TotalAvaluoPrendaY", 0)
                ImpresoraContratos.Print Format(TotalAvaluo, FMonedaSigno)
                
                'Total Prestamo Prendas
                .CurrentX = Regresa_Valor("CONTRATO", "TotalPrestamoPrendaX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "TotalPrestamoPrendaY", 0)
                ImpresoraContratos.Print Format(TotalPrestamo, FMonedaSigno)
    
            Else
                
                DescPrendaY = Regresa_Valor("CONTRATO", "DescripcionPrendasY", 0)

                rcAux.Open "SELECT MarcayModelo,Año,Color,Placas,Observaciones FROM detallesempenoautos WHERE IDEmpeno=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
                
                .CurrentX = Regresa_Valor("CONTRATO", "DescripcionPrendasX", 0)
                .CurrentY = DescPrendaY
                ImpresoraContratos.Print "AUTO"
                
                .CurrentX = Regresa_Valor("CONTRATO", "AvaluoPrendaX", 0)
                .CurrentY = DescPrendaY
                ImpresoraContratos.Print "$" & Format(rcConsulta!Avaluo, FMoneda)
                
                .CurrentX = Regresa_Valor("CONTRATO", "PrestamoPrendaX", 0)
                .CurrentY = DescPrendaY
                ImpresoraContratos.Print "$" & Format(rcConsulta!Prestamo, FMoneda)
                
                strDescripcion = ""
                x = 0
                For i = 1 To Len(IIf(IsNull(rcAux!MarcayModelo) Or Trim(rcAux!MarcayModelo) = "", "", rcAux!MarcayModelo) & _
                                 IIf(IsNull(rcAux!Año) Or Trim(rcAux!Año) = "", "", " " & rcAux!Año) & _
                                 IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & _
                                 IIf(IsNull(rcAux!Placas) Or Trim(rcAux!Placas) = "", "", " PLACAS: " & rcAux!Placas) & _
                                 IIf(IsNull(rcAux!Observaciones) Or Trim(rcAux!Observaciones) = "", "", " OBSERVACIONES: " & rcAux!Observaciones)) Step 40
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "ReferenciasX", 0)
                    .CurrentY = DescPrendaY + (2.5 * x)
                    
                    strDescripcion = LTrim(Mid(IIf(IsNull(rcAux!MarcayModelo) Or Trim(rcAux!MarcayModelo) = "", "", rcAux!MarcayModelo) & _
                                               IIf(IsNull(rcAux!Año) Or Trim(rcAux!Año) = "", "", " " & rcAux!Año) & _
                                               IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & _
                                               IIf(IsNull(rcAux!Placas) Or Trim(rcAux!Placas) = "", "", " PLACAS: " & rcAux!Placas) & _
                                               IIf(IsNull(rcAux!Observaciones) Or Trim(rcAux!Observaciones) = "", "", " OBSERVACIONES: " & rcAux!Observaciones), i * 1, 40 + IIf(Mid(strDescripcion, i, 1) = " ", 1, 0)))
                    
                    ImpresoraContratos.Print strDescripcion
                    x = x + 1
                Next i
                
                rcAux.Close
                
            End If
'************************************************************************************************************************************************************************
    
            .FontSize = 7
                    
            'MontoAvaluo
            .CurrentX = Regresa_Valor("CONTRATO", "MontoAvaluoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MontoAvaluoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Avaluo, FMoneda)
            
            'MontoAvaluoLetraX
            .CurrentX = Regresa_Valor("CONTRATO", "MontoAvaluoLetraX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MontoAvaluoLetraY", 0)
            ImpresoraContratos.Print CantidadEnLetra(rcConsulta!Avaluo)
                    
            'PorcenPrestamoAvaluo
            .CurrentX = Regresa_Valor("CONTRATO", "PorcenPrestamoAvaluoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "PorcenPrestamoAvaluoY", 0)
            ImpresoraContratos.Print Format(Round((rcConsulta!Prestamo * 100) / rcConsulta!Avaluo, 1), "0.00")
            
            'FechaLimiteRefrendo
            .CurrentX = Regresa_Valor("CONTRATO", "FechaLimiteRefrendoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "FechaLimiteRefrendoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Vencimiento, "DD/MMM/YYYY")
            
            'FechaLimiteFiniquito
            .CurrentX = Regresa_Valor("CONTRATO", "FechaLimiteFiniquitoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "FechaLimiteFiniquitoY", 0)
            ImpresoraContratos.Print Format(DateAdd("D", DiasGracia, rcConsulta!Vencimiento), "DD/MMM/YYYY")
            
            'FechaComercializacion
            .CurrentX = Regresa_Valor("CONTRATO", "FechaComercializacionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "FechaComercializacionY", 0)
            ImpresoraContratos.Print Format(DateAdd("D", diasEnajenacion, rcConsulta!Vencimiento), "DD/MMM/YYYY")
            
            'Tasa de Iva
            .CurrentX = Regresa_Valor("CONTRATO", "TasaIvaX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "TasaIvaY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Iva, "0.00")
            
            .FontSize = 7
            
            'Horario
            .CurrentX = Regresa_Valor("CONTRATO", "HorarioX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "HorarioY", 0)
            ImpresoraContratos.Print Regresa_Valor_BD("HorarioSucursal")
            
            'DiaFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "DiaFirmasX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DiaFirmasY", 0)
            ImpresoraContratos.Print Day(rcConsulta!Fecha)
            
            'MesFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "MesFirmasX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MesFirmasY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Fecha, "MMMM")
            
            'YearFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "YearFirmasX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "YearFirmasY", 0)
            ImpresoraContratos.Print Year(rcConsulta!Fecha)
            
            
            
            'Valuador
            .CurrentX = Regresa_Valor("CONTRATO", "ValuadorX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ValuadorY", 0)
            ImpresoraContratos.Print rcConsulta!Valuador
            
            '---------
             'DiaFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "DiaFirmasDesempeñoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DiaFirmasDesempeñoY", 0)
            ImpresoraContratos.Print Day(rcConsulta!Fecha)
            
            'MesFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "MesFirmasDesempeñoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MesFirmasDesempeñoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Fecha, "MMMM")
            
            'YearFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "YearFirmasDesempeñoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "YearFirmasDesempeñoY", 0)
            ImpresoraContratos.Print Year(rcConsulta!Fecha)
            
            
            'MLD.- MODIF -----------------------------------------------------------------
            
            Dim RsCte As New ADODB.Recordset
            Dim SqlCte As String
            
            .FontBold = False
            .FontSize = 6.5
            
            SqlCte = "SELECT c.Nombre,c.ApellidoPaterno,c.ApellidoMaterno,c.Apellido,c.FecNac,if(p.Descripcion is Null ,'MEXICO',p.Descripcion) AS PaisNacimiento,if(n.Descripcion is null, 'MEXICO',n.Descripcion) AS PaisNacionalidad," & _
                     "c.Direccion,c.NoExterior,c.NoInterior,c.Colonia,c.Municipio,c.Estado,c.Tel,c.NumeroIdentificacion,c.CP,c.Email,c.Rfc,c.Curp,o.Descripcion AS Ocupacion,i.Descripcion AS TipoIdentificacion, i.Dependencia as Expide " & _
                     "FROM clientes AS c Left Join mld_paises AS p ON c.IdPaisNacimiento = p.Id Left Join mld_tipo_identificaciones AS i ON c.IdTipoIdent = i.Id Left Join mld_paises AS n ON c.IdPaisNacionalidad = n.Id Left Join mld_actividades_economicas AS o ON c.IdOcupacion = o.Id " & _
                     "WHERE c.Id=" & rcConsulta!IDCliente
            
            RsCte.Open SqlCte, dbDatos, adOpenForwardOnly, adLockOptimistic
            If Not RsCte.EOF Then
            
                'Expediente Contrato
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ContratoX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ContratoY", 0)
                ImpresoraContratos.Print rcConsulta!NumContrato
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ApellidoPX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ApellidoPY", 0)
                ImpresoraContratos.Print IIf(Trim(RsCte!ApellidoPaterno) = "", RsCte!Apellido, RsCte!ApellidoPaterno)
                
                'Expediente ApellidoM
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ApellidoMX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ApellidoMY", 0)
                ImpresoraContratos.Print RsCte!ApellidoMaterno
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NombreX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NombreY", 0)
                ImpresoraContratos.Print RsCte!Nombre
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "FechaNacX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "FechaNacY", 0)
                ImpresoraContratos.Print RsCte!FecNac
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "PaisNacX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "PaisNacY", 0)
                ImpresoraContratos.Print RsCte!PaisNacimiento
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NacionalidadX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NacionalidadY", 0)
                ImpresoraContratos.Print RsCte!PaisNacionalidad
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "CalleX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "CalleY", 0)
                ImpresoraContratos.Print RsCte!Direccion
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumExtX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumExtY", 0)
                ImpresoraContratos.Print RsCte!NoExterior
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumIntX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumIntY", 0)
                ImpresoraContratos.Print RsCte!NoInterior
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ColoniaX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ColoniaY", 0)
                ImpresoraContratos.Print RsCte!Colonia
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "MunicipioX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "MunicipioY", 0)
                ImpresoraContratos.Print RsCte!Municipio
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "EstadoX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "EstadoY", 0)
                ImpresoraContratos.Print RsCte!Estado
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "CpX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "CpY", 0)
                ImpresoraContratos.Print RsCte!CP
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "TelefonoX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "TelefonoY", 0)
                ImpresoraContratos.Print RsCte!Tel
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "EmailX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "EmailY", 0)
                ImpresoraContratos.Print RsCte!Email
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "RFCX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "RFCY", 0)
                ImpresoraContratos.Print IIf(IsNull(RsCte!RFC), "", RsCte!RFC)
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "CURPX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "CURPY", 0)
                ImpresoraContratos.Print IIf(Trim(RsCte!Curp) = "", "", RsCte!Curp)
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "TipoIdentX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "TipoIdentY", 0)
                If IsNull(RsCte!tipoidentificacion) = True Then
                    ImpresoraContratos.Print UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE RegDefault=1"))
                Else
                    ImpresoraContratos.Print UCase(RsCte!tipoidentificacion)
                End If
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumIdentX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumIdentY", 0)
                ImpresoraContratos.Print RsCte!NumeroIdentificacion
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ExpideX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ExpideY", 0)
                ImpresoraContratos.Print UCase(IIf(IsNull(RsCte!expide), SacaValor("mld_tipo_identificaciones", "Dependencia", " WHERE RegDefault=1"), RsCte!expide))
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "OcupacionX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "OcupacionY", 0)
                ImpresoraContratos.Print IIf(IsNull(RsCte!ocupacion), "", RsCte!ocupacion)
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "CiudadSucX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "CiudadSucY", 0)
                ImpresoraContratos.Print SacaValor("sucursales", "Ciudad", " WHERE Activa=1") 'RsCte!CiudadSucursal
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "DiaX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "DiaY", 0)
                ImpresoraContratos.Print Day(rcConsulta!Fecha)
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "MesX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "MesY", 0)
                ImpresoraContratos.Print UCase(Format(rcConsulta!Fecha, "MMMM"))
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "AnoX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "AnoY", 0)
                ImpresoraContratos.Print Year(rcConsulta!Fecha)
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "FirmaValuadorX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "FirmaValuadorY", 0)
                ImpresoraContratos.Print UCase(rcConsulta!Valuador)
                
                'Expediente ApellidoP
                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "FirmaClienteX", 0)
                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "FirmaClienteY", 0)
                ImpresoraContratos.Print rcConsulta!Cliente
                
            End If
            RsCte.Close
            Set RsCte = Nothing
            
'***********Segudo Suaje******************************************************************************************************************************************
'''            .Font = "3 of 9 Barcode"
'''            .FontSize = 20
'''
'''            'SuajeCodigo
'''            .CurrentX = Regresa_Valor("CONTRATO_SUAJE", "SuajeCodigoX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO_SUAJE", "SuajeCodigoY", 0)
'''            ImpresoraContratos.Print "*" & Contrato & "*"
'''
'''            .Font = "Arial Narrow"
'''            .FontBold = True
'''            .FontSize = 16
'''
'''            'SuajeContrato
'''            .CurrentX = Regresa_Valor("CONTRATO_SUAJE", "SuajeContratoX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO_SUAJE", "SuajeContratoY", 0)
'''            ImpresoraContratos.Print rcConsulta!NumContrato
'''
'''            .FontBold = False
'''            .FontSize = 8
'''
'''            'SuajeCliente
'''            .CurrentX = Regresa_Valor("CONTRATO_SUAJE", "SuajeClienteX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO_SUAJE", "SuajeClienteY", 0)
'''            ImpresoraContratos.Print "CLIENTE: " & rcConsulta!Cliente
'''
'''            'SuajeVencimiento
'''            .CurrentX = Regresa_Valor("CONTRATO_SUAJE", "SuajeVencimientoX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO_SUAJE", "SuajeVencimientoY", 0)
'''            ImpresoraContratos.Print "VENCIMIENTO: " & Format(rcConsulta!Vencimiento, "DD/MMM/YYYY")
'''
'''            'SuajePrestamo
'''            .CurrentX = Regresa_Valor("CONTRATO_SUAJE", "SuajePrestamoX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO_SUAJE", "SuajePrestamoY", 0)
'''            ImpresoraContratos.Print "PRÉSTAMO: " & Format(rcConsulta!Prestamo, FMonedaSigno)
'''
'''            'SuajeTotalPeso
'''            .CurrentX = Regresa_Valor("CONTRATO_SUAJE", "SuajeTotalPesoX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO_SUAJE", "SuajeTotalPesoY", 0)
'''            ImpresoraContratos.Print "PESO: " & PesoTotal & " GRMS."
'''
'''            'SuajeBolsa
'''            .CurrentX = Regresa_Valor("CONTRATO_SUAJE", "SuajeBolsaX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO_SUAJE", "SuajeBolsaY", 0)
'''            ImpresoraContratos.Print "BOLSA: " & rcConsulta!NumBolsa
'***********Segudo Suaje Fin**************************************************************************************************************************************
            
            .EndDoc
        End With
    Next
    
    rcConsulta.Close
    Set rcConsulta = Nothing
    
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub
'Public Sub Imprimir_Boleta_CR_Caidas(ID As Long, Optional Reimpresion As Boolean = False, Optional Etiqueta As Boolean = False, Optional Auto As Boolean = False, Optional Refrendo As Boolean = False)
'
'    Dim ImpresoraEtiquetas As Printer
'    Dim ImpresoraContratos As Printer
'    Dim rcConsulta As New ADODB.Recordset
'    Dim rcConsultaCo As New ADODB.Recordset
'    Dim rcAux As New ADODB.Recordset
'    Dim i As Integer, x As Integer, crIntereses As Double, Contrato As String, strDescripcion, DescPrendaY As Double, PosicionY As Double, PesoTotal As Double
'    Dim TotalAvaluo As Double, TotalPrestamo As Double, TotalPeso As Double, Serie As Integer, NumRefrendos As Integer, Tipo As String
'    Dim Meses As Integer, ImpresoraTickets As Boolean, NextVencimiento As String, DiasGracia As Integer, diasEnajenacion As Integer
'    Dim Impresiones As Integer
'
'
'On Error GoTo 0 'error
'
''''    rcConsulta.Open "SELECT e.Fecha,e.Prestamo,e.Avaluo,e.NumContrato,e.Folio,e.Responsable,e.Beneficiario,e.Notas,e.Serie,e.TipoInteres,e.TipoTasa,e.Tasa,e.Almacenaje,e.Iva,e.VenPeriodo,e.Periodo,e.Vencimiento,e.Valuador,e.NumBolsa,c.Id as IdCliente,c.IDTipoIdent,e.Cat,e.IntAnual,e.AlmAnual " & _
''''                    ",e.NumIdentBeneficiario,e.IDcotitular, CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,c.Tel as TelCliente ,c.Identificacion,c.NumeroIdentificacion,CONCAT(c.Direccion,' COL.: ',c.Colonia,' ',c.Municipio,' ',c.Estado) AS Direccion,Mercadeo " & _
''''                    "FROM empeno e INNER JOIN clientes c ON e.IDCliente = c.ID WHERE e.ID = " & ID, dbDatos, adOpenForwardOnly, adLockReadOnly
'    rcConsulta.Open "SELECT e.Fecha,e.Prestamo,e.Avaluo,e.NumContrato,e.Folio,e.Responsable,e.Beneficiario,e.Notas,e.Serie,e.TipoInteres,e.TipoTasa,e.Tasa,e.Almacenaje,e.Seguro,e.Iva,e.VenPeriodo,e.Periodo,e.Vencimiento,e.Valuador,e.NumBolsa,c.Id as IdCliente,c.IDTipoIdent,e.Cat " & _
'                    ",e.IDcotitular, CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,c.Tel as TelCliente ,c.Identificacion,c.NumeroIdentificacion,CONCAT(c.Direccion,' COL.: ',c.Colonia,' ',c.Municipio,' ',c.Estado) AS Direccion, c.Email " & _
'                    "FROM empeno e INNER JOIN clientes c ON e.IDCliente = c.ID WHERE e.ID = " & ID, dbDatos, adOpenForwardOnly, adLockReadOnly
'
'    If rcConsulta!IDCotitular <> 0 Then
'         rcConsultaCo.Open "SELECT CONCAT(Nombre,' ',Apellido) AS Cotitular,Identificacion ,NumeroIdentificacion AS IdentificacionCo,CONCAT(Direccion,' COL.: ',Colonia,' ',Municipio,' ',Estado) AS DireccionCo " & _
'                    "FROM clientes WHERE ID = " & rcConsulta!IDCotitular, dbDatos, adOpenForwardOnly, adLockReadOnly
'    End If
'
'
'    'Si es contrato de pagos fijos
'    If rcConsulta!TipoInteres = "FIJA" Then
'
''        Select Case rcConsulta!TipoTasa
''        Case "MENSUAL"
''
'            Meses = 1
''        Case "QUINCENAL"
''
''            Meses = 2
''        Case "SEMANAL"
''
''            Meses = 4
''        End Select
'
'        If Reimpresion = False Then
'
'            'Intereses
'            'GeneraPagos ID, rcConsulta!Prestamo, rcConsulta!Tasa * (1 + (rcConsulta!Iva / 100)), rcConsulta!Almacenaje * (1 + (rcConsulta!Iva / 100)), rcConsulta!Seguro * (1 + (rcConsulta!Iva / 100)), rcConsulta!VenPeriodo * Meses, rcConsulta!Periodo, rcConsulta!Fecha
'            GeneraPagos ID, rcConsulta!Prestamo, rcConsulta!Tasa, rcConsulta!Almacenaje, rcConsulta!Seguro, rcConsulta!VenPeriodo * Meses, rcConsulta!Periodo, rcConsulta!Fecha
'            Sleep 500
'
'        End If
'
'        'Próximo Vencimiento
'        'NextVencimiento = SacaValor("pagosfijos", "Vencimiento", " WHERE ID = " & Val(SacaValor("pagosfijos", "MIN(ID)", " WHERE IDEmpeno = " & rcConsulta!ID)))
'        NextVencimiento = SacaValor("pagosfijos", "Vencimiento", " WHERE ID = " & Val(SacaValor("pagosfijos", "MIN(ID)", " WHERE IDEmpeno = " & ID)))
'
'        ImpresoraTickets = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
'
'        'Imprimo el Calendario de pagos
'        With frmMDI.Cr
'            .Reset
'            .DiscardSavedData = True
'            .WindowShowPrintSetupBtn = True
'            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'            .ReportFileName = Path & "\Reportes\TicketPagosFijos.rpt"
'            .SelectionFormula = "{empeno.ID}=" & ID & ""
'            .Formulas(0) = "NumPagos=" & rcConsulta!VenPeriodo * Meses & ""
'            .Formulas(1) = "Enajenacion=" & Regresa_Valor_BD("DiasEnajenacion") & ""
'            .Formulas(2) = "Notas='" & Trim(Regresa_Valor_BD("Notas")) & "'"
'            .Formulas(3) = "ProximoVencimiento='" & Format(CDate(NextVencimiento), "DD-MMM-YYYY") & "'"
'            .Destination = crptToWindow
'
'            'La mando a la impresora por default
'            If ImpresoraTickets Then
'                .PrinterName = strNombreImp
'                .PrinterDriver = strDriverImp
'                .PrinterPort = strPuertoImp
'                .Destination = crptToPrinter
'            End If
'
'            .WindowTitle = "Calendario Pagos"
'            .WindowState = crptMaximized
'            .Action = 1
'        End With
'
'    Else
'
'        'Opciones de Pago
'        crIntereses = OpcionesPago(rcConsulta!Prestamo, rcConsulta!Avaluo, rcConsulta!Fecha, ID, rcConsulta!TipoTasa)
'        crIntereses = IIf(crIntereses < Redondeo(Regresa_Valor_BD("PagoMinimo") * (1 + (Regresa_Valor_BD("IVA") / 100))), Redondeo(Regresa_Valor_BD("PagoMinimo") * (1 + (Regresa_Valor_BD("IVA") / 100))), crIntereses)
'    End If
'
'    'Saco los Dias de Gracia
'    'DiasGracia = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres = ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo = tp.ID INNER JOIN plazos p ON ct.IDPlazo = p.ID", "DGracia", " WHERE ti.Descripcion = '" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo)))
'    DiasGracia = Val(Regresa_Valor_BD("DiasGracia"))
'
'    'Saco los dias de Enajenación
'    diasEnajenacion = Regresa_Valor_BD("DiasEnajenacion")
'
'    'Tomo el Número de Contrato
'    Contrato = rcConsulta!NumContrato
'
'    For i = 1 To 6 - Len(Contrato)
'        Contrato = "0" & Contrato
'    Next i
'
'    Regresa_Impresora Contratos, ImpresoraContratos
'
'    For Impresiones = 1 To IIf(Reimpresion = True, 1, 2)
'
'        With ImpresoraContratos
'
'            .ScaleMode = vbMillimeters
''           .FontBold = True
'            .Font = "Arial Narrow"
'            .FontSize = 18
'
'            'Número de Contrato
'            .CurrentX = Regresa_Valor("CONTRATO", "NumContratoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "NumContratoY", 0)
'            ImpresoraContratos.Print rcConsulta!NumContrato
'
'            .FontBold = False
'            .FontSize = 14
'
''            'Número de Bolsa
''            .CurrentX = Regresa_Valor("CONTRATO", "NumBolsaX", 0)
''            .CurrentY = Regresa_Valor("CONTRATO", "NumBolsaY", 0)
''            ImpresoraContratos.Print "Bolsa: " & rcConsulta!NumContrato
''
'            'Codigo de Barras
'
'            .Font = "3 of 9 Barcode"
'            .FontSize = 24
'
'            .CurrentX = Regresa_Valor("CONTRATO", "CodigoBarrasX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "CodigoBarrasY", 0)
'            ImpresoraContratos.Print rcConsulta!NumContrato
'
'            'Fecha de Contrato
'            .Font = "Arial Narrow"
'            .FontSize = Regresa_Valor("CONTRATO", "SucursalEncFS", 0)
'             '10
'
'            'Nombre de la sucursal
'            .CurrentX = Regresa_Valor("CONTRATO", "SucursalEncX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SucursalEncY", 0)
'            ImpresoraContratos.Print Sucursal.NombreComercial
'
'            .FontSize = 8
'            'Fecha del Contrato
'            .CurrentX = Regresa_Valor("CONTRATO", "FechaContratoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "FechaContratoY", 0)
'            ImpresoraContratos.Print Format(rcConsulta!Fecha, "DD/MMM/YYYY")
'
'            'Fecha del Vencimiento encabezado
'            .CurrentX = Regresa_Valor("CONTRATO", "FechaVencimientoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "FechaVencimientoY", 0)
'            ImpresoraContratos.Print Format(rcConsulta!Vencimiento, "DD/MMM/YYYY")
'
'            'Tipo de Interes encabezado
'            .CurrentX = Regresa_Valor("CONTRATO", "TipoIntX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "TipoIntY", 0)
'            'ImpresoraContratos.Print rcConsulta!TipoInteres
'
''            'NumRefrendos
''            NumRefrendos = NumeroRefrendos(rcConsulta!NumContrato, rcConsulta!Serie)
''            .CurrentX = Regresa_Valor("CONTRATO", "NumRefrendosX", 0)
''            .CurrentY = Regresa_Valor("CONTRATO", "NumRefrendosY", 0)
''            ImpresoraContratos.Print IIf(NumRefrendos > 0, "Núm Refrendos: " & NumRefrendos, "")
'
'            .Font.Size = 8
'
'            'Imprimo la sucursal
'            .CurrentX = Regresa_Valor("CONTRATO", "RazonSocialSucursalX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "RazonSocialSucursalY", 0)
'            ImpresoraContratos.Print Sucursal.RazonSocial
'
'            'Imprimo el domicilio sucursal
'            .CurrentX = Regresa_Valor("CONTRATO", "DireccionSucursalX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "DireccionSucursalY", 0)
'            ImpresoraContratos.Print Sucursal.Direccion & " " & Sucursal.Ciudad & " " & Sucursal.Estado
'
'            'Telefono Sucursal
'            .CurrentX = Regresa_Valor("CONTRATO", "TelefonoSucursalX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "TelefonoSucursalY", 0)
'            ImpresoraContratos.Print Sucursal.Telefono
'
'            'Imprimo el RFC
'            .CurrentX = Regresa_Valor("CONTRATO", "RFCSucursalX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "RFCSucursalY", 0)
'            ImpresoraContratos.Print Sucursal.RFC
'
'            .CurrentX = Regresa_Valor("CONTRATO", "EmailSucursalX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "EmailSucursalY", 0)
'            ImpresoraContratos.Print SacaValor("sucursales", "CorreoAclaraciones", " Where Clave=" & Sucursal.Clave)
'            'CorreoAclaraciones
'
'            'Imprimo el Cliente
'            .CurrentX = Regresa_Valor("CONTRATO", "ClienteX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "ClienteY", 0)
'            ImpresoraContratos.Print rcConsulta!Cliente
'
'            'Imprimo el Telefono del Cliente
'            .CurrentX = Regresa_Valor("CONTRATO", "TelClienteX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "TelClienteY", 0)
'            ImpresoraContratos.Print "Tel:" & rcConsulta!TelCliente
'
'            'Imprimo el Correo del Cliente
'            .CurrentX = Regresa_Valor("CONTRATO", "CorreoClienteX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "CorreoClienteY", 0)
'            If IsNull(rcConsulta!Email) Then
'            ImpresoraContratos.Print ""
'            Else
'            ImpresoraContratos.Print rcConsulta!Email
'            End If
'
'            'Identificacion
'            .CurrentX = Regresa_Valor("CONTRATO", "IdentificacionX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "IdentificacionY", 0)
'            If IsNull(rcConsulta!Identificacion) = False Then
'                ImpresoraContratos.Print UCase(rcConsulta!Identificacion)
'            Else
'                ImpresoraContratos.Print UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE Id=" & rcConsulta!IDTipoIdent))
'            End If
'
'            'Número Identificacion
'            .CurrentX = Regresa_Valor("CONTRATO", "NumIdentificacionX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "NumIdentificacionY", 0)
'            ImpresoraContratos.Print rcConsulta!NumeroIdentificacion
'
'            'Direccion
'            .CurrentX = Regresa_Valor("CONTRATO", "DireccionX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "DireccionY", 0)
'            ImpresoraContratos.Print rcConsulta!Direccion
'
'            If rcConsulta!IDCotitular <> 0 Then
'                If Not rcConsultaCo.EOF And Not rcConsultaCo.BOF Then
'
'                'Cotitular
'                .CurrentX = Regresa_Valor("CONTRATO", "CotitularX", 0)
'                .CurrentY = Regresa_Valor("CONTRATO", "CotitularY", 0)
'                ImpresoraContratos.Print rcConsultaCo!Cotitular
'
'
'                'Cotitular Direccion
'                '.CurrentX = Regresa_Valor("CONTRATO", "IdentCotitularX", 0)
'                '.CurrentY = Regresa_Valor("CONTRATO", "IdentCotitularY", 0)
''                ImpresoraContratos.Print rcConsultaCo!IdentificacionCo
'
'                'Cotitular Direccion
'                .CurrentX = Regresa_Valor("CONTRATO", "DirCotitularX", 0)
'                .CurrentY = Regresa_Valor("CONTRATO", "DirCotitularY", 0)
'                ImpresoraContratos.Print rcConsultaCo!DireccionCo
'                End If
'
'            End If
'
'            'Beneficiario
'            .CurrentX = Regresa_Valor("CONTRATO", "BeneficiarioX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "BeneficiarioY", 0)
'            ImpresoraContratos.Print rcConsulta!Beneficiario & ""
'
'            'Identificación Beneficiario
'            '.CurrentX = Regresa_Valor("CONTRATO", "NumIdentBeneficiarioX", 0)
'            '.CurrentY = Regresa_Valor("CONTRATO", "NumIdentBeneficiarioY", 0)
'            '*****************ImpresoraContratos.Print rcConsulta!NumIdentBeneficiario
'
'            .Font.Size = 8 '10
'
'
'            'CAT
'            .CurrentX = Regresa_Valor("CONTRATO", "CATX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "CATY", 0)
''            ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Cat", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
'            ImpresoraContratos.Print Format(rcConsulta!CAT, "0.00")
'
'            'CMT
'            .CurrentX = Regresa_Valor("CONTRATO", "CMTX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "CMTY", 0)
''            ImpresoraContratos.Print Format((Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))) / 12), "0.00")
'            'ImpresoraContratos.Print Format((CDbl(rcConsulta!CAT) / 12), "0.00")
'
'            'CDT
'            .CurrentX = Regresa_Valor("CONTRATO", "CDTX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "CDTY", 0)
''            ImpresoraContratos.Print Format((Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))) / 365), "0.00")
'            'ImpresoraContratos.Print Format((CDbl(rcConsulta!CAT) / 365), "0.00")
'
'            'Tasa de Interes Anual
'            .CurrentX = Regresa_Valor("CONTRATO", "TasaInteresAnualX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "TasaInteresAnualY", 0)
'            'ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
'            'ImpresoraContratos.Print Format(rcConsulta!IntAnual, "0.00")
'            ImpresoraContratos.Print Format(Val(Regresa_Valor_BD("IntAnual")), "0.00")
'
'            'Tasa de Interes Mensual
'            .CurrentX = Regresa_Valor("CONTRATO", "TasaInteresMensualX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "TasaInteresMensualY", 0)
'            'ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
'            'ImpresoraContratos.Print Format(rcConsulta!IntAnual, "0.00")
'            ImpresoraContratos.Print Format(rcConsulta!Tasa + rcConsulta!Almacenaje + rcConsulta!Seguro, "0.00")
'
'            'Tasa de Interes Diario
'            .CurrentX = Regresa_Valor("CONTRATO", "TasaInteresDiariaX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "TasaInteresDiariaY", 0)
'            'ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
'            'ImpresoraContratos.Print Format(rcConsulta!IntAnual, "0.00")
'            ImpresoraContratos.Print Format((rcConsulta!Tasa + rcConsulta!Almacenaje + rcConsulta!Seguro) / 30, "0.00")
'
'
'            'Monto del Prestamo
'            .CurrentX = Regresa_Valor("CONTRATO", "PrestamoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "PrestamoY", 0)
'            ImpresoraContratos.Print Format(rcConsulta!Prestamo, FMoneda)
'
'            'Monto Total a Pagar
'            .CurrentX = Regresa_Valor("CONTRATO", "MontoTotalPagarX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "MontoTotalPagarY", 0)
'            ImpresoraContratos.Print Format(rcConsulta!Prestamo + crIntereses, FMoneda)
'
'            'Almacenaje CAT
'            .CurrentX = Regresa_Valor("CONTRATO", "AlmacenajeX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "AlmacenajeY", 0)
''            ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Almacenaje", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))) * IIf(rcConsulta!TipoTasa = "MENSUAL", 1, IIf(rcConsulta!TipoTasa = "QUINCENAL", 2, IIf(rcConsulta!TipoTasa = "SEMANAL", 4, 30))), "0.00")
'            ImpresoraContratos.Print Format(rcConsulta!Almacenaje, "0.00")
'
'            'Comercialización
'            .CurrentX = Regresa_Valor("CONTRATO", "ComercializacionX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "ComercializacionY", 0)
'            ImpresoraContratos.Print Format(Regresa_Valor_BD("GtosVenta"), "0.00")
'
'            'ReposicionContrato
'            .CurrentX = Regresa_Valor("CONTRATO", "ReposicionContratoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "ReposicionContratoY", 0)
'            ImpresoraContratos.Print Format(Regresa_Valor_BD("ImportePerdida"), FMoneda)
'
'            'Desempeño Extemporaneo
'            .CurrentX = Regresa_Valor("CONTRATO", "DesempenoExtemporaneoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "DesempenoExtemporaneoY", 0)
'            ImpresoraContratos.Print Format(Regresa_Valor_BD("Operacion"), "0.00")
'
'            'PlazoPrestamo
'            .CurrentX = Regresa_Valor("CONTRATO", "PlazoPrestamoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "PlazoPrestamoY", 0)
'            ImpresoraContratos.Print rcConsulta!VenPeriodo & " " & IIf(rcConsulta!TipoTasa = "MENSUAL" And rcConsulta!VenPeriodo > 1, "MESES", IIf(rcConsulta!TipoTasa = "MENSUAL" And rcConsulta!VenPeriodo = 1, "MES", IIf(rcConsulta!TipoTasa = "QUINCENAL", "QUINCENAS", IIf(rcConsulta!TipoTasa = "SEMANAL", "SEMANAS", "DIAS"))))
'
'
''            'Etiqueta Notas
''            .CurrentX = Regresa_Valor("CONTRATO", "NotasX", 0)
''            .CurrentY = Regresa_Valor("CONTRATO", "NotasY", 0)
''            ImpresoraContratos.Print rcConsulta!Notas
'
'            'Campos Opciones de Pago
'
'''''            .Font.Bold = True
'''''            .Font.Size = 6
'''''
'''''            .CurrentX = Regresa_Valor("CONTRATO", "EtiquetaImporteX", 0)
'''''            .CurrentY = Regresa_Valor("CONTRATO", "EtiquetaImporteY", 0)
'''''            ImpresoraContratos.Print "Imp. del Mutuo"
'''''
'''''            .CurrentX = Regresa_Valor("CONTRATO", "EtiquetaInteresesX", 0)
'''''            .CurrentY = Regresa_Valor("CONTRATO", "EtiquetaInteresesY", 0)
'''''            ImpresoraContratos.Print "Intereses"
'''''
'''''            .CurrentX = Regresa_Valor("CONTRATO", "EtiquetaAlmacenajeX", 0)
'''''            .CurrentY = Regresa_Valor("CONTRATO", "EtiquetaAlmacenajeY", 0)
'''''            ImpresoraContratos.Print "Almacenaje"
'''''
'''''            .CurrentX = Regresa_Valor("CONTRATO", "EtiquetaIVAX", 0)
'''''            .CurrentY = Regresa_Valor("CONTRATO", "EtiquetaIVAY", 0)
'''''            ImpresoraContratos.Print "IVA"
'
'            'Opciones de pago
'
'            .Font.Bold = False
'            .Font.Size = 8 '10
'
'            PosicionY = Regresa_Valor("CONTRATO", "OpcionPagoY", 0)
'            If rcConsulta!TipoInteres = "FIJA" Then
'
'                .CurrentX = Regresa_Valor("CONTRATO", "OpcionPagoX", 0)
'                .CurrentY = PosicionY
'                ImpresoraContratos.Print "CONTRATO DE PAGOS FIJOS. IMPORTES A PAGAR EN EL TICKET ANEXO A ESTE CONTRATO"
'
'            Else
'
'                i = 1
'                rcAux.Open "SELECT * FROM opcionpagos WHERE IDEmpeno=" & ID & " AND PC='" & NombrePc & "' ORDER BY ID", dbReportes, adOpenForwardOnly, adLockReadOnly
'                Do While Not rcAux.EOF
'
'                    'Opcion Numero
'                    .CurrentX = Regresa_Valor("CONTRATO", "OpcionPagoX", 0)
'                    .CurrentY = PosicionY
'                    ImpresoraContratos.Print i
'
'                    'ImporteMutuo
'                    .CurrentX = Regresa_Valor("CONTRATO", "ImporteMutuoX", 0)
'                    .CurrentY = PosicionY
'                    ImpresoraContratos.Print RegresaEspacios(rcAux!Prestamo, 15, True)
'
'                    'MontoIntereses
'                    .CurrentX = Regresa_Valor("CONTRATO", "MontoInteresesX", 0)
'                    .CurrentY = PosicionY
'                    ImpresoraContratos.Print RegresaEspacios(rcAux!Interes, 15, True)
'
'                    'MontoAlmacenaje
'                    .CurrentX = Regresa_Valor("CONTRATO", "MontoAlmacenajeX", 0)
'                    .CurrentY = PosicionY
'                    ImpresoraContratos.Print RegresaEspacios(rcAux!Almacenaje + rcAux!Seguro, 15, True)
'
'                    'MontoAlmacenaje
'                    .CurrentX = Regresa_Valor("CONTRATO", "MontoSeguroX", 0)
'                    .CurrentY = PosicionY
'                    'ImpresoraContratos.Print RegresaEspacios(rcAux!Seguro, 15, True)
'
'                    'ImporteIva
'                    .CurrentX = Regresa_Valor("CONTRATO", "ImporteIvaX", 0)
'                    .CurrentY = PosicionY
'                    ImpresoraContratos.Print RegresaEspacios(rcAux!ImporteIva, 15, True)
'
'                    'Por Refrendo
'                    .CurrentX = Regresa_Valor("CONTRATO", "PagoRefrendoX", 0)
'                    .CurrentY = PosicionY
'''''                    ImpresoraContratos.Print RegresaEspacios(rcAux!Interes + rcAux!Almacenaje + rcAux!ImporteIva, 15, True)
'                    ImpresoraContratos.Print RegresaEspacios(rcAux!Interes + rcAux!Almacenaje + rcAux!Seguro + rcAux!ImporteIva, 15, True)
'
'                    'Por Desempeno
'                    .CurrentX = Regresa_Valor("CONTRATO", "PagoDesempenoX", 0)
'                    .CurrentY = PosicionY
''''                    ImpresoraContratos.Print RegresaEspacios(rcAux!Prestamo + rcAux!Interes + rcAux!Almacenaje + rcAux!ImporteIva, 15, True)
'                    ImpresoraContratos.Print RegresaEspacios(rcAux!Prestamo + rcAux!Interes + rcAux!Almacenaje + rcAux!Seguro + rcAux!ImporteIva, 15, True)
'
'                    'Vencimiento
'                    .CurrentX = Regresa_Valor("CONTRATO", "PagoVencimientoX", 0)
'                    .CurrentY = PosicionY
'                    ImpresoraContratos.Print Format(rcAux!Vencimiento, "DD/MMM/YYYY") 'Format(rcAux!FechaIni, "DD/MMM/YYYY") & " al " &
'
'                i = i + 1
'                PosicionY = PosicionY + 4
'                rcAux.MoveNext
'                Loop
'                rcAux.Close
'
'            End If
'
'            'Opcion de Mercadeo
'            '*****************If rcConsulta!Mercadeo = 1 Then
'            '*****************    .CurrentX = Regresa_Valor("CONTRATO", "MercadeoSiX", 0)
'            '*****************    .CurrentY = Regresa_Valor("CONTRATO", "MercadeoSiY", 0)
'            '*****************    ImpresoraContratos.Print "X"
'            '*****************Else
'                .CurrentX = Regresa_Valor("CONTRATO", "MercadeoNoX", 0)
'                .CurrentY = Regresa_Valor("CONTRATO", "MercadeoNoY", 0)
'                ImpresoraContratos.Print "X"
'            '*****************End If
'
''************************************************************************************************************************************************************************
''Imprimo la descripción de las prendas
'            .FontSize = 6
'
'            If rcConsulta!Serie = 1 Or rcConsulta!Serie = 4 Then
'
'                PesoTotal = 0
'                TotalAvaluo = 0
'                TotalPrestamo = 0
'
'                DescPrendaY = Regresa_Valor("CONTRATO", "DescripcionPrendasY", 0)
'
'
'                rcAux.Open "SELECT tipo.Descripcion AS Desc_Tipo,d.Cantidad,d.Articulo AS Des_Prenda,d.CantidadPiedras,d.PesoPiedras,kilatajes.Descripcion AS Desc_Kilates,d.Peso,d.Estado AS Desc_Estado,d.Avaluo,d.Prestamo,d.Observaciones,d.Marca,d.Modelo,d.Serie,d.Tamano,d.Color " & _
'                        "FROM detallesempeno d LEFT JOIN tipo ON d.Tipo=tipo.ID LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE d.IDEmpeno=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
'
'
'                While Not rcAux.EOF
'
'                    .CurrentX = Regresa_Valor("CONTRATO", "DescripcionPrendasX", 0)
'                    .CurrentY = DescPrendaY
'                    ImpresoraContratos.Print rcAux!Desc_Tipo
'
'                    .CurrentX = Regresa_Valor("CONTRATO", "PiedraPesoX", 0)
'                    .CurrentY = DescPrendaY
'                    'ImpresoraContratos.Print IIf(rcAux!PesoPiedras <= 0, "", rcAux!PesoPiedras)
'
'                    .CurrentX = Regresa_Valor("CONTRATO", "KilatajeX", 0)
'                    .CurrentY = DescPrendaY
'                    'ImpresoraContratos.Print IIf(IsNull(rcAux!Desc_Kilates), "", rcAux!Desc_Kilates)
'
'                    .CurrentX = Regresa_Valor("CONTRATO", "PesoX", 0)
'                    .CurrentY = DescPrendaY
'                    'ImpresoraContratos.Print IIf(rcAux!Peso <= 0, "", Format(rcAux!Peso, "0.00"))
'
'                    .CurrentX = Regresa_Valor("CONTRATO", "CalidadX", 0)
'                    .CurrentY = DescPrendaY
'                    'ImpresoraContratos.Print IIf(rcAux!Desc_Estado <= 0, "", rcAux!Desc_Estado)
'
'                    .Font.Bold = True
'                    .FontSize = 7
'
'                    .CurrentX = Regresa_Valor("CONTRATO", "AvaluoPrendaX", 0)
'                    .CurrentY = DescPrendaY
'                    ImpresoraContratos.Print "$" & Format(rcAux!Avaluo, FMoneda)
'
'                    .CurrentX = Regresa_Valor("CONTRATO", "PrestamoPrendaX", 0)
'                    .CurrentY = DescPrendaY
'                    ImpresoraContratos.Print "$" & Format(rcAux!Prestamo, FMoneda)
'
'                    .Font.Bold = False
'                    .FontSize = 6
'
'                    strDescripcion = ""
'                    x = 0
'                    'For i = 1 To Len(rcAux!Des_Prenda & " " & IIf(IsNull(rcAux!Desc_Kilates), "", " Kilataje: " & rcAux!Desc_Kilates) & IIf(IsNull(rcAux!Peso) Or rcAux!Peso = 0, "", " Peso: " & rcAux!Peso) & IIf(IsNull(rcAux!Desc_Estado), "", " Estado: " & rcAux!Desc_Estado) & IIf(IsNull(rcAux!Marca) Or Trim(rcAux!Marca) = "", "", " MARCA: " & rcAux!Marca) & IIf(IsNull(rcAux!Modelo) Or Trim(rcAux!Modelo) = "", "", " MODELO: " & rcAux!Modelo) & IIf(IsNull(rcAux!Serie) Or Trim(rcAux!Serie) = "", "", " SERIE: " & rcAux!Serie) & IIf(IsNull(rcAux!Tamano) Or Trim(rcAux!Tamano) = "", "", " TAMAÑO: " & rcAux!Tamano) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & " " & rcAux!Observaciones) Step 40
'                    For i = 1 To Len(IIf(IsNull(rcAux!Desc_Kilates), "", rcAux!Desc_Kilates) & IIf(IsNull(rcAux!Peso) Or rcAux!Peso = 0, "", " Peso " & rcAux!Peso & " gramos") & IIf(IsNull(rcAux!Marca) Or Trim(rcAux!Marca) = "", "", " MARCA: " & rcAux!Marca) & IIf(IsNull(rcAux!Modelo) Or Trim(rcAux!Modelo) = "", "", " MODELO: " & rcAux!Modelo) & IIf(IsNull(rcAux!Serie) Or Trim(rcAux!Serie) = "", "", " SERIE: " & rcAux!Serie) & IIf(IsNull(rcAux!Tamano) Or Trim(rcAux!Tamano) = "", "", " TAMAÑO: " & rcAux!Tamano) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color)) Step 40
'
'                        .CurrentX = Regresa_Valor("CONTRATO", "ReferenciasX", 0)
'                        .CurrentY = DescPrendaY + (2.5 * x)
'
'                        'strDescripcion = LTrim(Mid(rcAux!Des_Prenda & " " & IIf(IsNull(rcAux!Desc_Kilates), "", " Kilataje: " & rcAux!Desc_Kilates) & IIf(IsNull(rcAux!Peso) Or rcAux!Peso = 0, "", " Peso " & rcAux!Peso & " gramos") & IIf(IsNull(rcAux!Desc_Estado), "", " Estado: " & rcAux!Desc_Estado) & IIf(IsNull(rcAux!Marca) Or Trim(rcAux!Marca) = "", "", " MARCA: " & rcAux!Marca) & IIf(IsNull(rcAux!Modelo) Or Trim(rcAux!Modelo) = "", "", " MODELO: " & rcAux!Modelo) & IIf(IsNull(rcAux!Serie) Or Trim(rcAux!Serie) = "", "", " SERIE: " & rcAux!Serie) & IIf(IsNull(rcAux!Tamano) Or Trim(rcAux!Tamano) = "", "", " TAMAÑO: " & rcAux!Tamano) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & " " & rcAux!Observaciones, i * 1, 40 + IIf(Mid(strDescripcion, i, 1) = " ", 1, 0)))
'                        strDescripcion = LTrim(Mid(IIf(IsNull(rcAux!Des_Prenda), "", rcAux!Des_Prenda) & IIf(IsNull(rcAux!Desc_Kilates), "", rcAux!Desc_Kilates) & IIf(IsNull(rcAux!Peso) Or rcAux!Peso = 0, "", " Peso " & rcAux!Peso & " gramos") & IIf(IsNull(rcAux!Marca) Or Trim(rcAux!Marca) = "", "", " MARCA: " & rcAux!Marca) & IIf(IsNull(rcAux!Modelo) Or Trim(rcAux!Modelo) = "", "", " MODELO: " & rcAux!Modelo) & IIf(IsNull(rcAux!Serie) Or Trim(rcAux!Serie) = "", "", " SERIE: " & rcAux!Serie) & IIf(IsNull(rcAux!Tamano) Or Trim(rcAux!Tamano) = "", "", " TAMAÑO: " & rcAux!Tamano) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & IIf(IsNull(rcAux!Observaciones), "", rcAux!Observaciones), i * 1, 40 + IIf(Mid(strDescripcion, i, 1) = " ", 1, 0)))
'                        ImpresoraContratos.Print strDescripcion
'                        x = x + 1
'                    Next i
'
'                    .CurrentX = Regresa_Valor("CONTRATO", "PorcentajeX", 0)
'                    .CurrentY = DescPrendaY
'                    ImpresoraContratos.Print Format((rcAux!Prestamo / rcAux!Avaluo) * 100, "0.00") & "%"
'
'                    TotalAvaluo = TotalAvaluo + rcAux!Avaluo
'                    TotalPrestamo = TotalPrestamo + rcAux!Prestamo
'                    DescPrendaY = DescPrendaY + (2.5 * (x - 1)) + 2.5
'                    PesoTotal = PesoTotal + (rcAux!Peso * rcAux!Cantidad)
'                    rcAux.MoveNext
'                Wend
'                rcAux.Close
'
'                'Total Peso
'                .CurrentX = Regresa_Valor("CONTRATO", "TotalPesoX", 0)
'                .CurrentY = Regresa_Valor("CONTRATO", "TotalPesoY", 0)
'                ImpresoraContratos.Print Format(PesoTotal, "0.00")
'
'                'Total Avaluo Prendas
'                .CurrentX = Regresa_Valor("CONTRATO", "TotalAvaluoPrendaX", 0)
'                .CurrentY = Regresa_Valor("CONTRATO", "TotalAvaluoPrendaY", 0)
'                ImpresoraContratos.Print Format(TotalAvaluo, FMonedaSigno)
'
'                'Total Prestamo Prendas
'                .CurrentX = Regresa_Valor("CONTRATO", "TotalPrestamoPrendaX", 0)
'                .CurrentY = Regresa_Valor("CONTRATO", "TotalPrestamoPrendaY", 0)
'                ImpresoraContratos.Print Format(TotalPrestamo, FMonedaSigno)
'
'            End If
''************************************************************************************************************************************************************************
'
'            '.FontSize = 10
'             .Font.Bold = False
'            'MontoAvaluo
'            .CurrentX = Regresa_Valor("CONTRATO", "MontoAvaluoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "MontoAvaluoY", 0)
'            ImpresoraContratos.Print Format(rcConsulta!Avaluo, FMoneda)
'
'            'MontoAvaluoLetraX
'            .CurrentX = Regresa_Valor("CONTRATO", "MontoAvaluoLetraX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "MontoAvaluoLetraY", 0)
'            ImpresoraContratos.Print CantidadEnLetra(rcConsulta!Avaluo)
'
'            'PorcenPrestamoAvaluo
'            .CurrentX = Regresa_Valor("CONTRATO", "PorcenPrestamoAvaluoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "PorcenPrestamoAvaluoY", 0)
'            ImpresoraContratos.Print Format(Round((rcConsulta!Prestamo * 100) / rcConsulta!Avaluo, 1), "0.00")
'
'            'FechaLimiteRefrendo
'            '.CurrentX = Regresa_Valor("CONTRATO", "FechaLimiteRefrendoX", 0)
'            '.CurrentY = Regresa_Valor("CONTRATO", "FechaLimiteRefrendoY", 0)
'            'ImpresoraContratos.Print Format(rcConsulta!Vencimiento, "DD/MMM/YYYY")
'
'            'FechaLimiteFiniquito
'            .CurrentX = Regresa_Valor("CONTRATO", "FechaLimiteFiniquitoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "FechaLimiteFiniquitoY", 0)
'            ImpresoraContratos.Print Format(DateAdd("D", DiasGracia, rcConsulta!Vencimiento), "DD/MMM/YYYY")
'
'
'            'FechaComercializacion
'            .CurrentX = Regresa_Valor("CONTRATO", "FechaComercializacionX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "FechaComercializacionY", 0)
'            ImpresoraContratos.Print Format(DateAdd("D", diasEnajenacion, rcConsulta!Vencimiento), "DD/MMM/YYYY")
'
'
'            'IVA
'            .CurrentX = Regresa_Valor("CONTRATO", "IVAX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "IVAY", 0)
'            ImpresoraContratos.Print Regresa_Valor_BD("IVA")
'
'            'Datos Profeco
'            .Font.Size = 8
''''            'Razon Social
''''            .CurrentX = Regresa_Valor("CONTRATO", "RazonSocialX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "RazonSocialY", 0)
''''            ImpresoraContratos.Print Sucursal.RazonSocial
'
''''            'Direccion
''''            .CurrentX = Regresa_Valor("CONTRATO", "DireccionSucX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "DireccionSucY", 0)
''''            ImpresoraContratos.Print Sucursal.Direccion
''''
''''            'Telefono
''''            .CurrentX = Regresa_Valor("CONTRATO", "TelefonoX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "TelefonoY", 0)
''''            ImpresoraContratos.Print Sucursal.Telefono
''''
''''            'Email
''''            .CurrentX = Regresa_Valor("CONTRATO", "EmailX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "EmailY", 0)
''''            ImpresoraContratos.Print Sucursal.Email
'
'            'Número Registro Profeco
'            .CurrentX = Regresa_Valor("CONTRATO", "NumProfecoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "NumProfecoY", 0)
'            ImpresoraContratos.Print SacaValor("sucursales", "ContratoRegistrado", " Where Clave=" & Sucursal.Clave)
'
'            'Fecha Registro Profeco
'            .CurrentX = Regresa_Valor("CONTRATO", "FechaProfecoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "FechaProfecoY", 0)
'            ImpresoraContratos.Print SacaValor("sucursales", "FechaContratoRegistrado", " Where Clave=" & Sucursal.Clave)
'
''''            'Etiqueta Tramite
''''            .CurrentX = Regresa_Valor("CONTRATO", "TramiteSucX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "TramiteSucY", 0)
''''            ImpresoraContratos.Print "EN TRAMITE"
'
'            'Responsable
'            .CurrentX = Regresa_Valor("CONTRATO", "ResponsableX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "ResponsableY", 0)
'            ImpresoraContratos.Print Sucursal.RazonSocial
'
'             .Font.Size = 6
'             'Responsable
'            .CurrentX = Regresa_Valor("CONTRATO", "ClienteDesX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "ClienteDesY", 0)
'            ImpresoraContratos.Print rcConsulta!Cliente
'
''''            'Tasa de Iva
''''            .CurrentX = Regresa_Valor("CONTRATO", "TasaIvaX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "TasaIvaY", 0)
''''            ImpresoraContratos.Print Format(rcConsulta!Iva, "0.00")
'
'            .FontSize = 8 '10
''            .Font.Bold = True
'            'Horario
''            .CurrentX = Regresa_Valor("CONTRATO", "HorarioX", 0)
''            .CurrentY = Regresa_Valor("CONTRATO", "HorarioY", 0)
''            ImpresoraContratos.Print Regresa_Valor_BD("Horario")
'            .Font.Bold = False
'
'            .CurrentX = Regresa_Valor("CONTRATO", "NotaX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "NotaY", 0)
'            ImpresoraContratos.Print Regresa_Valor_BD("Notas") ' MisParametros.Notas
'
'            'DiaFirmas
'            .CurrentX = Regresa_Valor("CONTRATO", "DiaFirmasX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "DiaFirmasY", 0)
'            ImpresoraContratos.Print Day(rcConsulta!Fecha)
'
'            'MesFirmas
'            .CurrentX = Regresa_Valor("CONTRATO", "MesFirmasX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "MesFirmasY", 0)
'            ImpresoraContratos.Print Month(rcConsulta!Fecha)
'
'            'YearFirmas
'            .CurrentX = Regresa_Valor("CONTRATO", "YearFirmasX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "YearFirmasY", 0)
'            ImpresoraContratos.Print Year(rcConsulta!Fecha)
'
'
'            'Consumidor
'            .CurrentX = Regresa_Valor("CONTRATO", "ClienteFirmaX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "ClienteFirmaY", 0)
'            ImpresoraContratos.Print rcConsulta!Cliente
'
'
'            'Consumidor
'            '.CurrentX = Regresa_Valor("CONTRATO", "ConsumidorX", 0)
'            '.CurrentY = Regresa_Valor("CONTRATO", "ConsumidorY", 0)
'            'ImpresoraContratos.Print rcConsulta!Cliente
'
'            .Font.Size = 8
'            'Valuador
'            .CurrentX = Regresa_Valor("CONTRATO", "ValuadorX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "ValuadorY", 0)
'            ImpresoraContratos.Print rcConsulta!Valuador
'
''            .FontSize = 8
'            'Horario
'            .CurrentX = Regresa_Valor("CONTRATO", "HorarioX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "HorarioY", 0)
'            ImpresoraContratos.Print Regresa_Valor_BD("HorarioSucursal")
'
'
'            'Imprimo el domicilio sucursal 2
'            .CurrentX = Regresa_Valor("CONTRATO", "DireccionSucursal2X", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "DireccionSucursal2Y", 0)
'            ImpresoraContratos.Print Sucursal.Direccion & " " & Sucursal.Ciudad & " " & Sucursal.Estado
'
'            'Telefono Sucursal 2
'            .CurrentX = Regresa_Valor("CONTRATO", "TelefonoSucursal2X", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "TelefonoSucursal2Y", 0)
'            ImpresoraContratos.Print Sucursal.Telefono
'
'            'Correo Sucursal 2
'            .CurrentX = Regresa_Valor("CONTRATO", "EmailSucursal2X", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "EmailSucursal2Y", 0)
'            ImpresoraContratos.Print SacaValor("sucursales", "CorreoAclaraciones", " Where Clave=" & Sucursal.Clave)
'
'            'Pagina Internet
'            .CurrentX = Regresa_Valor("CONTRATO", "PaginaInternetX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "PaginaInternetY", 0)
'            ImpresoraContratos.Print "www.mrayudon.com"
'
'            'Avaluo Letra
'            '.CurrentX = Regresa_Valor("CONTRATO", "AvaluoLetraX", 0)
'            '.CurrentY = Regresa_Valor("CONTRATO", "AvaluoLetraY", 0)
'            'ImpresoraContratos.Print CantidadEnLetra(rcConsulta!Avaluo)
'
'            .Font = "3 of 9 Barcode"
'            .FontSize = 24
'
'            'SuajeCodigo
'            .CurrentX = Regresa_Valor("CONTRATO", "SuajeCodigoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SuajeCodigoY", 0)
'            ImpresoraContratos.Print rcConsulta!NumContrato
'
'            .Font = "Arial Narrow"
'            .FontBold = False
'            .FontSize = 10
'
'            'SuajeContrato
'            .CurrentX = Regresa_Valor("CONTRATO", "SuajeContratoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SuajeContratoY", 0)
'            ImpresoraContratos.Print rcConsulta!NumContrato
'
'
'            .FontSize = 8
'            'SuajeCliente
'            .CurrentX = Regresa_Valor("CONTRATO", "SuajeClienteX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SuajeClienteY", 0)
'            ImpresoraContratos.Print rcConsulta!Cliente
'
'            'SuajePrestamo
'            .CurrentX = Regresa_Valor("CONTRATO", "SuajePrestamoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SuajePrestamoY", 0)
'            ImpresoraContratos.Print Format(rcConsulta!Prestamo, FMonedaSigno)
'
''''            'SuajeTotalPeso
''''            .CurrentX = Regresa_Valor("CONTRATO", "SuajeTotalPesoX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "SuajeTotalPesoY", 0)
''''            ImpresoraContratos.Print "PESO: " & PesoTotal & " GRMS."
'
'            'SuajePrenda
'            .FontSize = 6
'            DescPrendaY = Regresa_Valor("CONTRATO", "SuajePrendaY", 0)
'
'            rcAux.Open "SELECT tipo.Descripcion AS Desc_Tipo,d.Cantidad,d.Articulo AS Des_Prenda,d.CantidadPiedras,d.PesoPiedras,kilatajes.Descripcion AS Desc_Kilates,d.Peso,d.Estado AS Desc_Estado,d.Avaluo,d.Prestamo,d.Observaciones,d.Marca,d.Modelo,d.Serie,d.Tamano,d.Color " & _
'                       "FROM detallesempeno d LEFT JOIN tipo ON d.Tipo=tipo.ID LEFT JOIN kilatajes ON d.Kilates=kilatajes.Clave WHERE d.IDEmpeno=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
'
'            While Not rcAux.EOF
'            strDescripcion = ""
'                    x = 0
'                    For i = 1 To Len(rcAux!Des_Prenda & " " & IIf(IsNull(rcAux!Desc_Kilates), "", " Kilataje: " & rcAux!Desc_Kilates) & IIf(IsNull(rcAux!Peso) Or rcAux!Peso = 0, "", " Peso: " & rcAux!Peso) & IIf(IsNull(rcAux!Desc_Estado), "", " Estado: " & rcAux!Desc_Estado) & IIf(IsNull(rcAux!Marca) Or Trim(rcAux!Marca) = "", "", " MARCA: " & rcAux!Marca) & IIf(IsNull(rcAux!Modelo) Or Trim(rcAux!Modelo) = "", "", " MODELO: " & rcAux!Modelo) & IIf(IsNull(rcAux!Serie) Or Trim(rcAux!Serie) = "", "", " SERIE: " & rcAux!Serie) & IIf(IsNull(rcAux!Tamano) Or Trim(rcAux!Tamano) = "", "", " TAMAÑO: " & rcAux!Tamano) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & " " & rcAux!Observaciones) Step 80
'
'                        .CurrentX = Regresa_Valor("CONTRATO", "SuajePrendaX", 0)
'                        .CurrentY = DescPrendaY + (2.5 * x)
'
'                        strDescripcion = LTrim(Mid(rcAux!Des_Prenda & " " & IIf(IsNull(rcAux!Desc_Kilates), "", " Kilataje: " & rcAux!Desc_Kilates) & IIf(IsNull(rcAux!Peso) Or rcAux!Peso = 0, "", " Peso: " & rcAux!Peso) & IIf(IsNull(rcAux!Desc_Estado), "", " Estado: " & rcAux!Desc_Estado) & IIf(IsNull(rcAux!Marca) Or Trim(rcAux!Marca) = "", "", " MARCA: " & rcAux!Marca) & IIf(IsNull(rcAux!Modelo) Or Trim(rcAux!Modelo) = "", "", " MODELO: " & rcAux!Modelo) & IIf(IsNull(rcAux!Serie) Or Trim(rcAux!Serie) = "", "", " SERIE: " & rcAux!Serie) & IIf(IsNull(rcAux!Tamano) Or Trim(rcAux!Tamano) = "", "", " TAMAÑO: " & rcAux!Tamano) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & " " & rcAux!Observaciones, i * 1, 80 + IIf(Mid(strDescripcion, i, 1) = " ", 1, 0)))
'                        ImpresoraContratos.Print strDescripcion
'                        x = x + 1
'                    Next i
'            rcAux.MoveNext
'            DescPrendaY = DescPrendaY + (2.5 * (x - 1)) + 2.5
'            Wend
'            rcAux.Close
'
'
'            'SuajePrestamo Letra
'            .CurrentX = Regresa_Valor("CONTRATO", "SuajePrestamoLetraX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SuajePrestamoLetraY", 0)
'            ImpresoraContratos.Print CantidadEnLetra(rcConsulta!Prestamo)
'
'            'SuajeFecha
'            .CurrentX = Regresa_Valor("CONTRATO", "SuajeFechaX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SuajeFechaY", 0)
'            ImpresoraContratos.Print Format(rcConsulta!Fecha, "DD/MMM/YYYY")
'
''            'SuajeBolsa
''            .CurrentX = Regresa_Valor("CONTRATO", "SuajeBolsaX", 0)
''            .CurrentY = Regresa_Valor("CONTRATO", "SuajeBolsaY", 0)
''            ImpresoraContratos.Print "BOLSA: " & rcConsulta!NumBolsa
'            .FontSize = Regresa_Valor("CONTRATO", "SucursalEnc2FS", 0)
'            'Nombre de la sucursal2
'            .CurrentX = Regresa_Valor("CONTRATO", "SucursalEnc2X", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SucursalEnc2Y", 0)
'            ImpresoraContratos.Print Sucursal.NombreComercial
'              .FontSize = 6
'            .EndDoc
'
'    End With
'
'    Next
'
'    If rcConsulta!IDCotitular <> 0 Then
'    rcConsultaCo.Close
'    Set rcConsultaCo = Nothing
'    End If
'
'    rcConsulta.Close
'    Set rcConsulta = Nothing
'
'
'    Exit Sub
'
'Error:
'    Maneja_Error Err
'    Set rcConsulta = Nothing
'End Sub
Function NumeroRefrendos(NumContrato As Long, Serie As Integer) As Integer

    Dim rcAux As New ADODB.Recordset

    rcAux.Open "SELECT COUNT(ID) AS Numero FROM empeno WHERE Cancelado=0 AND NumContrato = " & NumContrato & " AND Destino = 2 AND Serie = " & Serie, dbDatos, adOpenForwardOnly, adLockOptimistic
    
    If Not rcAux.BOF And Not rcAux.EOF Then
        NumeroRefrendos = rcAux!Numero
    Else
        NumeroRefrendos = 0
    End If
    
    rcAux.Close
    Set rcAux = Nothing
    
End Function

Public Sub Imprimir_Boleta_CR_Caidas_Autos(ID As Long, Optional Reimpresion As Boolean = False, Optional Etiqueta As Boolean = False, Optional Auto As Boolean = False, Optional Refrendo As Boolean = False)

    Dim ImpresoraEtiquetas As Printer
    Dim ImpresoraContratos As Printer
    Dim rcConsulta As New ADODB.Recordset
    Dim rcConsultaCo As New ADODB.Recordset
    Dim rcAux As New ADODB.Recordset
    Dim i As Integer, x As Integer, crIntereses As Double, Contrato As String, strDescripcion, DescPrendaY As Double, PosicionY As Double, PesoTotal As Double
    Dim TotalAvaluo As Double, TotalPrestamo As Double, TotalPeso As Double, Serie As Integer, NumRefrendos As Integer, Tipo As String
    Dim Meses As Integer, ImpresoraTickets As Boolean, NextVencimiento As String, DiasGracia As Integer, diasEnajenacion As Integer
    Dim Impresiones As Integer
 

On Error GoTo Error
        
'''    rcConsulta.Open "SELECT e.Fecha,e.Prestamo,e.Avaluo,e.NumContrato,e.Folio,e.Responsable,e.Beneficiario,e.Notas,e.Serie,e.TipoInteres,e.TipoTasa,e.Tasa,e.Almacenaje,e.Iva,e.VenPeriodo,e.Periodo,e.Vencimiento,e.Valuador,e.NumBolsa,c.Id as IdCliente,c.IDTipoIdent " & _
'''                    ",e.NumIdentBeneficiario, CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,c.Identificacion,c.NumeroIdentificacion,CONCAT(c.Direccion,' COL.: ',c.Colonia,' ',c.Municipio,' ',c.Estado) AS Direccion " & _
'''                    "FROM empeno e INNER JOIN clientes c ON e.IDCliente = c.ID WHERE e.ID = " & ID, dbDatos, adOpenForwardOnly, adLockReadOnly

    rcConsulta.Open "SELECT e.Fecha,e.Prestamo,e.Avaluo,e.NumContrato,e.Folio,e.Responsable,e.Beneficiario,e.Notas,e.Serie,e.TipoInteres,e.TipoTasa,e.Tasa,e.Almacenaje,e.Iva,e.seguro,e.VenPeriodo,e.Periodo,e.Vencimiento,e.Valuador,e.NumBolsa,c.Id as IdCliente,c.IDTipoIdent,e.Cat " & _
                    ",e.IDcotitular, CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,c.Tel as TelCliente ,c.Identificacion,c.NumeroIdentificacion,CONCAT(c.Direccion,' COL.: ',c.Colonia,' ',c.Municipio,' ',c.Estado) AS Direccion, c.Email " & _
                    "FROM empeno e INNER JOIN clientes c ON e.IDCliente = c.ID WHERE e.ID = " & ID, dbDatos, adOpenForwardOnly, adLockReadOnly
    
    If rcConsulta!IDCotitular <> 0 Then
    rcConsultaCo.Open "SELECT CONCAT(Nombre,' ',Apellido) AS Cotitular,Identificacion ,NumeroIdentificacion AS IdentificacionCo,CONCAT(Direccion,' COL.: ',Colonia,' ',Municipio,' ',Estado) AS DireccionCo " & _
                    "FROM clientes WHERE ID = " & rcConsulta!IDCotitular, dbDatos, adOpenForwardOnly, adLockReadOnly
    End If
    
    'Si es contrato de pagos fijos
    If rcConsulta!TipoInteres = "FIJA" Then
                  
        Select Case rcConsulta!TipoTasa
        Case "MENSUAL"
              
            Meses = 1
       Case "QUINCENAL"

            Meses = 2
        Case "SEMANAL"

            Meses = 4
        End Select
        
        If Reimpresion = False Then
            
            'Intereses
            'GeneraPagos ID, rcConsulta!Prestamo, rcConsulta!Tasa * (1 + (rcConsulta!Iva / 100)), rcConsulta!Almacenaje * (1 + (rcConsulta!Iva / 100)), rcConsulta!Seguro * (1 + (rcConsulta!Iva / 100)), rcConsulta!VenPeriodo * Meses, rcConsulta!Periodo, rcConsulta!Fecha
            GeneraPagos ID, rcConsulta!Prestamo, rcConsulta!Tasa, rcConsulta!Almacenaje, rcConsulta!Seguro, rcConsulta!VenPeriodo * Meses, rcConsulta!Periodo, rcConsulta!Fecha
            Sleep 500
        
        End If
        
        'Próximo Vencimiento
        NextVencimiento = SacaValor("pagosfijos", "Vencimiento", " WHERE ID = " & Val(SacaValor("pagosfijos", "MIN(ID)", " WHERE IDEmpeno = " & ID)))
        
        ImpresoraTickets = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
        
        'Imprimo el Calendario de pagos
        With frmMDI.Cr
            .Reset
            .DiscardSavedData = True
            .WindowShowPrintSetupBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\TicketPagosFijos.rpt"
            .SelectionFormula = "{empeno.ID}=" & ID & ""
            .Formulas(0) = "NumPagos=" & rcConsulta!VenPeriodo * Meses & ""
            .Formulas(1) = "Enajenacion=" & Regresa_Valor_BD("DiasEnajenacion") & ""
            .Formulas(2) = "Notas='" & Trim(Regresa_Valor_BD("Notas")) & "'"
            .Formulas(3) = "ProximoVencimiento='" & Format(CDate(NextVencimiento), "DD-MMM-YYYY") & "'"
            .Destination = crptToWindow
            
            'La mando a la impresora por default
            If ImpresoraTickets Then
                .PrinterName = strNombreImp
                .PrinterDriver = strDriverImp
                .PrinterPort = strPuertoImp
                .Destination = crptToPrinter
            End If
        
            .WindowTitle = "Calendario Pagos"
            .WindowState = crptMaximized
            .Action = 1
        End With
        
    Else
        
        'Opciones de Pago
        'OpcionesPago
        'OpcionesPagoAutos
        crIntereses = OpcionesPago(rcConsulta!Prestamo, rcConsulta!Avaluo, rcConsulta!Fecha, ID, rcConsulta!TipoTasa)

    End If
                    
    'Saco los Dias de Gracia
    'DiasGracia = Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres = ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo = tp.ID INNER JOIN plazos p ON ct.IDPlazo = p.ID", "DGracia", " WHERE ti.Descripcion = '" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo)))
    DiasGracia = Val(Regresa_Valor_BD("DiasGracia"))
                                                                 
    'Saco los dias de Enajenación
    diasEnajenacion = Regresa_Valor_BD("DiasEnajenacion")
    
    'Tomo el Número de Contrato
    Contrato = rcConsulta!NumContrato
    
    For i = 1 To 6 - Len(Contrato)
        Contrato = "0" & Contrato
    Next i
                
    Regresa_Impresora Contratos, ImpresoraContratos

    For Impresiones = 1 To IIf(Reimpresion = True, 2, 2)
        
        With ImpresoraContratos
        
            .ScaleMode = vbMillimeters
'           .FontBold = True
            .Font = "Arial Narrow"
            .FontSize = 18
            
            'Número de Contrato
            .CurrentX = Regresa_Valor("CONTRATO", "NumContratoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "NumContratoY", 0)
            ImpresoraContratos.Print rcConsulta!NumContrato
            
            .FontBold = False
            .FontSize = 14
            
'            'Número de Bolsa
'            .CurrentX = Regresa_Valor("CONTRATO", "NumBolsaX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "NumBolsaY", 0)
'            ImpresoraContratos.Print "Bolsa: " & rcConsulta!NumContrato
'
            'Codigo de Barras
            
            .Font = "3 of 9 Barcode"
            .FontSize = 24
            
            .CurrentX = Regresa_Valor("CONTRATO", "CodigoBarrasX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "CodigoBarrasY", 0)
            ImpresoraContratos.Print rcConsulta!NumContrato
            
            'Fecha de Sucursal
            .Font = "Arial Narrow"
            .FontSize = Regresa_Valor("CONTRATO", "SucursalEncFS", 0)
            
            'Nombre de la sucursal
            .CurrentX = Regresa_Valor("CONTRATO", "SucursalEncX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "SucursalEncY", 0)
            ImpresoraContratos.Print Sucursal.NombreComercial
            
            'Fecha de Contrato
            .Font = "Arial Narrow"
            .FontSize = 8 '10
            'Fecha del Contrato
            .CurrentX = Regresa_Valor("CONTRATO", "FechaContratoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "FechaContratoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Fecha, "DD/MMM/YYYY")
            
            'Fecha del Vencimiento encabezado
            .CurrentX = Regresa_Valor("CONTRATO", "FechaVencimientoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "FechaVencimientoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Vencimiento, "DD/MMM/YYYY")
            
            'Tipo de Interes encabezado
            '.CurrentX = Regresa_Valor("CONTRATO", "TipoIntX", 0)
            '.CurrentY = Regresa_Valor("CONTRATO", "TipoIntY", 0)
            'ImpresoraContratos.Print rcConsulta!TipoInteres
            
'            'NumRefrendos
'            NumRefrendos = NumeroRefrendos(rcConsulta!NumContrato, rcConsulta!Serie)
'            .CurrentX = Regresa_Valor("CONTRATO", "NumRefrendosX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "NumRefrendosY", 0)
'            ImpresoraContratos.Print IIf(NumRefrendos > 0, "Núm Refrendos: " & NumRefrendos, "")
            
            .Font.Size = 8
    
            'Imprimo la sucursal
            .CurrentX = Regresa_Valor("CONTRATO", "RazonSocialSucursalX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "RazonSocialSucursalY", 0)
            ImpresoraContratos.Print Sucursal.RazonSocial

            'Imprimo el domicilio sucursal
            .CurrentX = Regresa_Valor("CONTRATO", "DireccionSucursalX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DireccionSucursalY", 0)
            ImpresoraContratos.Print Sucursal.Direccion & " " & Sucursal.Ciudad & " " & Sucursal.Estado

            'Telefono Sucursal
            .CurrentX = Regresa_Valor("CONTRATO", "TelefonoSucursalX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "TelefonoSucursalY", 0)
            ImpresoraContratos.Print Sucursal.Telefono

            'Imprimo el RFC
            .CurrentX = Regresa_Valor("CONTRATO", "RFCSucursalX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "RFCSucursalY", 0)
            ImpresoraContratos.Print Sucursal.RFC
            
            .CurrentX = Regresa_Valor("CONTRATO", "EmailSucursalX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "EmailSucursalY", 0)
            ImpresoraContratos.Print SacaValor("sucursales", "CorreoAclaraciones", " Where Clave=" & Sucursal.Clave)
            'CorreoAclaraciones
            
'''            'Imprimo la sucursal
'''            .CurrentX = Regresa_Valor("CONTRATO", "SucursalX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "SucursalY", 0)
'''            ImpresoraContratos.Print Sucursal.RazonSocial
'''
'''            'Imprimo el domicilio sucursal
'''            .CurrentX = Regresa_Valor("CONTRATO", "DireccionSucursalX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "DireccionSucursalY", 0)
'''            ImpresoraContratos.Print Sucursal.Direccion & " " & Sucursal.Ciudad & " " & Sucursal.Estado
'''
'''            'Telefono Sucursal
'''            .CurrentX = Regresa_Valor("CONTRATO", "TelefonoSucursalX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "TelefonoSucursalY", 0)
'''            ImpresoraContratos.Print Sucursal.Telefono
'''
'''            'Imprimo el RFC
'''            .CurrentX = Regresa_Valor("CONTRATO", "RFCX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "RFCY", 0)
'''            ImpresoraContratos.Print Sucursal.RFC
            
            'Imprimo el Cliente
            .CurrentX = Regresa_Valor("CONTRATO", "ClienteX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ClienteY", 0)
            ImpresoraContratos.Print rcConsulta!Cliente
            
'''            'Identificacion
'''            .CurrentX = Regresa_Valor("CONTRATO", "IdentificacionX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "IdentificacionY", 0)
'''            If IsNull(rcConsulta!identificacion) = False Then
'''                ImpresoraContratos.Print UCase(rcConsulta!identificacion)
'''            Else
'''                ImpresoraContratos.Print UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE Id=" & rcConsulta!IDTipoIdent))
'''            End If

            'Imprimo el Telefono del Cliente
            .CurrentX = Regresa_Valor("CONTRATO", "TelClienteX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "TelClienteY", 0)
            ImpresoraContratos.Print "Tel:" & rcConsulta!TelCliente
                   'Imprimo el Correo del Cliente
            .CurrentX = Regresa_Valor("CONTRATO", "CorreoClienteX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "CorreoClienteY", 0)
            If IsNull(rcConsulta!Email) Then
            ImpresoraContratos.Print ""
            Else
            ImpresoraContratos.Print rcConsulta!Email
            End If
            
            'Identificacion
            .CurrentX = Regresa_Valor("CONTRATO", "IdentificacionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "IdentificacionY", 0)
            If IsNull(rcConsulta!Identificacion) = False Then
                ImpresoraContratos.Print UCase(rcConsulta!Identificacion)
            Else
                ImpresoraContratos.Print UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE Id=" & rcConsulta!IDTipoIdent))
            End If
            
            'Número Identificacion
            .CurrentX = Regresa_Valor("CONTRATO", "NumIdentificacionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "NumIdentificacionY", 0)
            ImpresoraContratos.Print rcConsulta!NumeroIdentificacion
            
            'Direccion
            .CurrentX = Regresa_Valor("CONTRATO", "DireccionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DireccionY", 0)
            ImpresoraContratos.Print rcConsulta!Direccion
            
            If rcConsulta!IDCotitular <> 0 Then
            
                If Not rcConsultaCo.EOF And Not rcConsultaCo.BOF Then
                
                'Cotitular
                .CurrentX = Regresa_Valor("CONTRATO", "CotitularX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "CotitularY", 0)
                ImpresoraContratos.Print rcConsultaCo!Cotitular
                
                            
                'Cotitular Direccion
                '.CurrentX = Regresa_Valor("CONTRATO", "IdentCotitularX", 0)
               ' .CurrentY = Regresa_Valor("CONTRATO", "IdentCotitularY", 0)
                'ImpresoraContratos.Print rcConsultaCo!IdentificacionCo
                
                'Cotitular Direccion
                .CurrentX = Regresa_Valor("CONTRATO", "DirCotitularX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "DirCotitularY", 0)
                ImpresoraContratos.Print rcConsultaCo!DireccionCo
                End If
            
            End If
            
            'Beneficiario
            .CurrentX = Regresa_Valor("CONTRATO", "BeneficiarioX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "BeneficiarioY", 0)
            ImpresoraContratos.Print rcConsulta!Beneficiario
            
            'Identificación Beneficiario
'            .CurrentX = Regresa_Valor("CONTRATO", "NumIdentBeneficiarioX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "NumIdentBeneficiarioY", 0)
'            ImpresoraContratos.Print rcConsulta!NumIdentBeneficiario
                 
                 
            .Font.Size = 8 '10
            
            
            'CAT
            .CurrentX = Regresa_Valor("CONTRATO", "CATX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "CATY", 0)
'            ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Cat", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
            ImpresoraContratos.Print Format(rcConsulta!CAT, "0.00")
            
            'CMT
            '.CurrentX = Regresa_Valor("CONTRATO", "CMTX", 0)
            '.CurrentY = Regresa_Valor("CONTRATO", "CMTY", 0)
'            ImpresoraContratos.Print Format((Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))) / 12), "0.00")
            'ImpresoraContratos.Print Format((CDbl(rcConsulta!CAT) / 12), "0.00")
            
            'CDT
            '.CurrentX = Regresa_Valor("CONTRATO", "CDTX", 0)
            '.CurrentY = Regresa_Valor("CONTRATO", "CDTY", 0)
'            ImpresoraContratos.Print Format((Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))) / 365), "0.00")
            'ImpresoraContratos.Print Format((CDbl(rcConsulta!CAT) / 365), "0.00")
            
            'Tasa de Interes Anual
            .CurrentX = Regresa_Valor("CONTRATO", "TasaInteresAnualX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "TasaInteresAnualY", 0)
            'ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
            'ImpresoraContratos.Print Format(rcConsulta!IntAnual, "0.00")
            ImpresoraContratos.Print Format(Val(Regresa_Valor_BD("IntAnual")), "0.00")
            
            'Tasa de Interes Mensual
            .CurrentX = Regresa_Valor("CONTRATO", "TasaInteresMensualX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "TasaInteresMensualY", 0)
            'ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
            'ImpresoraContratos.Print Format(rcConsulta!IntAnual, "0.00")
            ImpresoraContratos.Print Format(rcConsulta!Tasa + rcConsulta!Almacenaje + rcConsulta!Seguro, "0.00")
            
            'Tasa de Interes Diario
            .CurrentX = Regresa_Valor("CONTRATO", "TasaInteresDiariaX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "TasaInteresDiariaY", 0)
            'ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "IntAnual", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))), "0.00")
            'ImpresoraContratos.Print Format(rcConsulta!IntAnual, "0.00")
            ImpresoraContratos.Print Format((rcConsulta!Tasa + rcConsulta!Almacenaje + rcConsulta!Seguro) / 30, "0.00")
            
            
            
            
            
            
            
            
            'Monto del Prestamo
            .CurrentX = Regresa_Valor("CONTRATO", "PrestamoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "PrestamoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Prestamo, FMoneda)
                            
            'Monto Total a Pagar
            .CurrentX = Regresa_Valor("CONTRATO", "MontoTotalPagarX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MontoTotalPagarY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Prestamo + crIntereses, FMoneda)
'
'            'Almacenaje CAT
'            .CurrentX = Regresa_Valor("CONTRATO", "AlmacenajeCATX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "AlmacenajeCATY", 0)
''            ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Almacenaje", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))) * IIf(rcConsulta!TipoTasa = "MENSUAL", 1, IIf(rcConsulta!TipoTasa = "QUINCENAL", 2, IIf(rcConsulta!TipoTasa = "SEMANAL", 4, 30))), "0.00")
'            '************** ImpresoraContratos.Print Format(rcConsulta!AlmAnual, "0.00")
    
           'Almacenaje CAT
            .CurrentX = Regresa_Valor("CONTRATO", "AlmacenajeX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "AlmacenajeY", 0)
'            ImpresoraContratos.Print Format(Val(SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "Almacenaje", " WHERE ti.Descripcion='" & rcConsulta!TipoInteres & "' AND ti.Serie = " & rcConsulta!Serie & " AND tp.Descripcion='" & rcConsulta!TipoTasa & "' AND p.Descripcion=" & Val(rcConsulta!VenPeriodo))) * IIf(rcConsulta!TipoTasa = "MENSUAL", 1, IIf(rcConsulta!TipoTasa = "QUINCENAL", 2, IIf(rcConsulta!TipoTasa = "SEMANAL", 4, 30))), "0.00")
            ImpresoraContratos.Print Format(rcConsulta!Almacenaje, "0.00")
    
    
    
            'Comercialización
            .CurrentX = Regresa_Valor("CONTRATO", "ComercializacionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ComercializacionY", 0)
            ImpresoraContratos.Print Format(Regresa_Valor_BD("GtosVenta"), "0.00")
    
            'ReposicionContrato
            .CurrentX = Regresa_Valor("CONTRATO", "ReposicionContratoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ReposicionContratoY", 0)
            ImpresoraContratos.Print Format(Regresa_Valor_BD("ImportePerdida"), FMoneda)
    
            'Desempeño Extemporaneo
            .CurrentX = Regresa_Valor("CONTRATO", "DesempenoExtemporaneoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DesempenoExtemporaneoY", 0)
            ImpresoraContratos.Print Format(Regresa_Valor_BD("Operacion"), "0.00")
    
            'PlazoPrestamo
            .CurrentX = Regresa_Valor("CONTRATO", "PlazoPrestamoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "PlazoPrestamoY", 0)
            ImpresoraContratos.Print rcConsulta!VenPeriodo & " " & IIf(rcConsulta!TipoTasa = "MENSUAL" And rcConsulta!VenPeriodo > 1, "MESES", IIf(rcConsulta!TipoTasa = "MENSUAL" And rcConsulta!VenPeriodo = 1, "MES", IIf(rcConsulta!TipoTasa = "QUINCENAL", "QUINCENAS", IIf(rcConsulta!TipoTasa = "SEMANAL", "SEMANAS", "DIAS"))))
            
            
'            'Etiqueta Notas
'            .CurrentX = Regresa_Valor("CONTRATO", "NotasX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "NotasY", 0)
'            ImpresoraContratos.Print rcConsulta!Notas
            
            'Campos Opciones de Pago
            
''''            .Font.Bold = True
''''            .Font.Size = 6
''''
''''            .CurrentX = Regresa_Valor("CONTRATO", "EtiquetaImporteX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "EtiquetaImporteY", 0)
''''            ImpresoraContratos.Print "Imp. del Mutuo"
''''
''''            .CurrentX = Regresa_Valor("CONTRATO", "EtiquetaInteresesX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "EtiquetaInteresesY", 0)
''''            ImpresoraContratos.Print "Intereses"
''''
''''            .CurrentX = Regresa_Valor("CONTRATO", "EtiquetaAlmacenajeX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "EtiquetaAlmacenajeY", 0)
''''            ImpresoraContratos.Print "Almacenaje"
''''
''''            .CurrentX = Regresa_Valor("CONTRATO", "EtiquetaIVAX", 0)
''''            .CurrentY = Regresa_Valor("CONTRATO", "EtiquetaIVAY", 0)
''''            ImpresoraContratos.Print "IVA"
            
            'Opciones de pago
            
            .Font.Bold = False
            .Font.Size = 6 '10
            
            PosicionY = Regresa_Valor("CONTRATO", "OpcionPagoY", 0)
            If rcConsulta!TipoInteres = "FIJA" Then
    
                .CurrentX = Regresa_Valor("CONTRATO", "OpcionPagoX", 0)
                .CurrentY = PosicionY
                ImpresoraContratos.Print "CONTRATO DE PAGOS FIJOS. IMPORTES A PAGAR EN EL TICKET ANEXO A ESTE CONTRATO"
    
            Else
                
                i = 1
                rcAux.Open "SELECT * FROM opcionpagos WHERE IDEmpeno=" & ID & " AND PC='" & NombrePc & "' ORDER BY ID", dbReportes, adOpenForwardOnly, adLockReadOnly
                Do While Not rcAux.EOF
    
                    'Opcion Numero
                    .CurrentX = Regresa_Valor("CONTRATO", "OpcionPagoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print i
    
                    'ImporteMutuo
                    .CurrentX = Regresa_Valor("CONTRATO", "ImporteMutuoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Prestamo, 15, True)
    
                    'MontoIntereses
                    .CurrentX = Regresa_Valor("CONTRATO", "MontoInteresesX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Interes, 15, True)
    
                    'MontoAlmacenaje
                    .CurrentX = Regresa_Valor("CONTRATO", "MontoAlmacenajeX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Almacenaje, 15, True)
    
                    'ImporteIva
                    .CurrentX = Regresa_Valor("CONTRATO", "ImporteIvaX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!ImporteIva, 15, True)
    
                    'Por Refrendo
                    .CurrentX = Regresa_Valor("CONTRATO", "PagoRefrendoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Interes + rcAux!Almacenaje + rcAux!ImporteIva, 15, True)
    
                    'Por Desempeno
                    .CurrentX = Regresa_Valor("CONTRATO", "PagoDesempenoX", 0)
                    .CurrentY = PosicionY
                    ImpresoraContratos.Print RegresaEspacios(rcAux!Prestamo + rcAux!Interes + rcAux!Almacenaje + rcAux!ImporteIva, 15, True)
    
                    'Vencimiento
                    .CurrentX = Regresa_Valor("CONTRATO", "PagoVencimientoX", 0)
                    .CurrentY = PosicionY
                    
                    'stFechaVencimiento = Format(rcAux!FechaIni, "DD/MMM/YYYY") & " al " & Format(rcAux!Vencimiento, "DD/MMM/YYYY")
                    ImpresoraContratos.Print Format(rcAux!Vencimiento, "DD/MMM/YYYY")
                i = i + 1
                PosicionY = PosicionY + 4
                rcAux.MoveNext
                Loop
                rcAux.Close
            
            End If
            
            'Opcion de Mercadeo
'                If rcConsulta!Mercadeo = 1 Then
'                    .CurrentX = Regresa_Valor("CONTRATO", "MercadeoSiX", 0)
'                    .CurrentY = Regresa_Valor("CONTRATO", "MercadeoSiY", 0)
'                    ImpresoraContratos.Print "X"
'                Else
                    .CurrentX = Regresa_Valor("CONTRATO", "MercadeoNoX", 0)
                    .CurrentY = Regresa_Valor("CONTRATO", "MercadeoNoY", 0)
                    ImpresoraContratos.Print "X"
'                End If
'************************************************************************************************************************************************************************
'Imprimo la descripción de las prendas
            .FontSize = 6
    
'                DescPrendaY = Regresa_Valor("CONTRATO", "DescripcionPrendasY", 0)

                rcAux.Open "SELECT MarcayModelo,Año,Color,Placas,NumMotor,SerieChasis,Poliza,Gas,Kms,Factura,Observaciones FROM detallesempenoautos WHERE IDEmpeno=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic

                .CurrentX = Regresa_Valor("CONTRATO", "ModeloX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "ModeloY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!MarcayModelo) Or Trim(rcAux!MarcayModelo) = "", "", rcAux!MarcayModelo)

                .CurrentX = Regresa_Valor("CONTRATO", "AñoX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "AñoY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!Año) Or Trim(rcAux!Año) = "", "", rcAux!Año)

                .CurrentX = Regresa_Valor("CONTRATO", "PlacaX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "PlacaY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!Placas) Or Trim(rcAux!Placas) = "", "", rcAux!Placas)
                
                .CurrentX = Regresa_Valor("CONTRATO", "ColorX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "ColorY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", rcAux!Color)
                
                .CurrentX = Regresa_Valor("CONTRATO", "NumMotorX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "NumMotorY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!NumMotor) Or Trim(rcAux!NumMotor) = "", "", rcAux!NumMotor)
                
                .CurrentX = Regresa_Valor("CONTRATO", "SerieX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "SerieY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!SerieChasis) Or Trim(rcAux!SerieChasis) = "", "", rcAux!SerieChasis)
                
                .CurrentX = Regresa_Valor("CONTRATO", "PolizaX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "PolizaY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!Poliza) Or Trim(rcAux!Poliza) = "", "", rcAux!Poliza)
                
                .CurrentX = Regresa_Valor("CONTRATO", "GasX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "GasY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!Gas) Or Trim(rcAux!Gas) = "", "", rcAux!Gas)
                
                .CurrentX = Regresa_Valor("CONTRATO", "KmsX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "KmsY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!Kms) Or Trim(rcAux!Kms) = "", "", rcAux!Kms)
                
                .CurrentX = Regresa_Valor("CONTRATO", "FacturaX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "FacturaY", 0)
                ImpresoraContratos.Print IIf(IsNull(rcAux!Factura) Or Trim(rcAux!Factura) = "", "", rcAux!Factura)
                
                .CurrentX = Regresa_Valor("CONTRATO", "ObservacionesX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "ObservacionesY", 0)
                
                Dim ObservacionesY As Double
                ObservacionesY = Regresa_Valor("CONTRATO", "ObservacionesY", 0)
                
                If IsNull(rcAux!Observaciones) Or Trim(rcAux!Observaciones) = "" Then
                   ImpresoraContratos.Print IIf(IsNull(rcAux!Observaciones) Or Trim(rcAux!Observaciones) = "", "", rcAux!Observaciones)
                Else
                  If Len(rcAux!Observaciones) > 50 Then
                    'Dim ob1 As String
                    'ob1 = Mid$(rcAux!Observaciones, 1, 65)
                    'ImpresoraContratos.Print ob1
                    'ObservacionesY = ObservacionesY + 2
                    '.CurrentX = Regresa_Valor("CONTRATO", "ObservacionesX", 0)
                    '.CurrentY = ObservacionesY
                    'ob1 = Replace$(rcAux!Observaciones, ob1, "")
                    'ImpresoraContratos.Print ob1
                    
                    Dim strDescripciones2 As String
                   
                    x = 0
                    For i = 1 To Len(rcAux!Observaciones) Step 50
                        
                        .CurrentX = Regresa_Valor("CONTRATO", "ObservacionesX", 0)
                        .CurrentY = ObservacionesY + (2.5 * x)
                        
                        strDescripciones2 = LTrim(Mid(rcAux!Observaciones, i * 1, 50 + IIf(Mid(strDescripciones2, i, 1) = " ", 1, 0)))
                        ImpresoraContratos.Print strDescripciones2
                        x = x + 1
                    Next i
                    x = 0
                  Else
                    ImpresoraContratos.Print IIf(IsNull(rcAux!Observaciones) Or Trim(rcAux!Observaciones) = "", "", rcAux!Observaciones)

                  End If
                End If
               ' ImpresoraContratos.Print IIf(IsNull(rcAux!Observaciones) Or Trim(rcAux!Observaciones) = "", "", rcAux!Observaciones)
                rcAux.Close



                    .CurrentX = Regresa_Valor("CONTRATO", "AvaluoPrendaX", 0)
                    .CurrentY = Regresa_Valor("CONTRATO", "AvaluoPrendaY", 0)
                    ImpresoraContratos.Print "$" & Format(rcConsulta!Avaluo, FMoneda)
                    
                    .CurrentX = Regresa_Valor("CONTRATO", "PrestamoPrendaX", 0)
                    .CurrentY = Regresa_Valor("CONTRATO", "PrestamoPrendaY", 0)
                    ImpresoraContratos.Print "$" & Format(rcConsulta!Prestamo, FMoneda)
           
           
             .CurrentX = Regresa_Valor("CONTRATO", "TotalAvaluoPrendaX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "TotalAvaluoPrendaY", 0)
                ImpresoraContratos.Print Format(rcConsulta!Avaluo, FMonedaSigno)
                
                'Total Prestamo Prendas
                .CurrentX = Regresa_Valor("CONTRATO", "TotalPrestamoPrendaX", 0)
                .CurrentY = Regresa_Valor("CONTRATO", "TotalPrestamoPrendaY", 0)
                ImpresoraContratos.Print Format(rcConsulta!Prestamo, FMonedaSigno)
'************************************************************************************************************************************************************************
    
            .FontSize = 6
             .Font.Bold = False
            'MontoAvaluo
            .CurrentX = Regresa_Valor("CONTRATO", "MontoAvaluoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MontoAvaluoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Avaluo, FMoneda)
            
            'MontoAvaluoLetraX
            .CurrentX = Regresa_Valor("CONTRATO", "MontoAvaluoLetraX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MontoAvaluoLetraY", 0)
            ImpresoraContratos.Print CantidadEnLetra(rcConsulta!Avaluo)
                    
            'PorcenPrestamoAvaluo
            .CurrentX = Regresa_Valor("CONTRATO", "PorcenPrestamoAvaluoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "PorcenPrestamoAvaluoY", 0)
            ImpresoraContratos.Print Format(Round((rcConsulta!Prestamo * 100) / rcConsulta!Avaluo, 1), "0.00")
            
            'FechaLimiteRefrendo
'            .CurrentX = Regresa_Valor("CONTRATO", "FechaLimiteRefrendoX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "FechaLimiteRefrendoY", 0)
'            ImpresoraContratos.Print Format(rcConsulta!Vencimiento, "DD/MMM/YYYY")
            
            'FechaLimiteFiniquito
            .CurrentX = Regresa_Valor("CONTRATO", "FechaLimiteFiniquitoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "FechaLimiteFiniquitoY", 0)
            ImpresoraContratos.Print Format(DateAdd("D", DiasGracia, rcConsulta!Vencimiento), "DD/MMM/YYYY")
        
            
            'FechaComercializacion
            .CurrentX = Regresa_Valor("CONTRATO", "FechaComercializacionX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "FechaComercializacionY", 0)
            ImpresoraContratos.Print Format(DateAdd("D", diasEnajenacion, rcConsulta!Vencimiento), "DD/MMM/YYYY")
            
            
      
            
            
            'IVA
            .CurrentX = Regresa_Valor("CONTRATO", "IVAX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "IVAY", 0)
            ImpresoraContratos.Print Regresa_Valor_BD("IVA")

            'Datos Profeco
            .Font.Size = 6
'''            'Razon Social
'''            .CurrentX = Regresa_Valor("CONTRATO", "RazonSocialX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "RazonSocialY", 0)
'''            ImpresoraContratos.Print Sucursal.RazonSocial
            
'''            'Direccion
'''            .CurrentX = Regresa_Valor("CONTRATO", "DireccionSucX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "DireccionSucY", 0)
'''            ImpresoraContratos.Print Sucursal.Direccion
'''
'''            'Telefono
'''            .CurrentX = Regresa_Valor("CONTRATO", "TelefonoX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "TelefonoY", 0)
'''            ImpresoraContratos.Print Sucursal.Telefono
'''
'''            'Email
'''            .CurrentX = Regresa_Valor("CONTRATO", "EmailX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "EmailY", 0)
'''            ImpresoraContratos.Print Sucursal.Email
            
            'Número Registro Profeco
            .CurrentX = Regresa_Valor("CONTRATO", "NumProfecoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "NumProfecoY", 0)
          ImpresoraContratos.Print SacaValor("sucursales", "ContratoRegistrado", " Where Clave=" & Sucursal.Clave)
            
            'Fecha Registro Profeco
            .CurrentX = Regresa_Valor("CONTRATO", "FechaProfecoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "FechaProfecoY", 0)
          ImpresoraContratos.Print SacaValor("sucursales", "FechaContratoRegistrado", " Where Clave=" & Sucursal.Clave)
            
'''            'Etiqueta Tramite
'''            .CurrentX = Regresa_Valor("CONTRATO", "TramiteSucX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "TramiteSucY", 0)
'''            ImpresoraContratos.Print "EN TRAMITE"
            
            'Responsable
            .CurrentX = Regresa_Valor("CONTRATO", "ResponsableX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ResponsableY", 0)
            ImpresoraContratos.Print Sucursal.RazonSocial
            
            
'''            'Tasa de Iva
'''            .CurrentX = Regresa_Valor("CONTRATO", "TasaIvaX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "TasaIvaY", 0)
'''            ImpresoraContratos.Print Format(rcConsulta!Iva, "0.00")
            
           ' .FontSize = 10
'            .Font.Bold = True
            'Horario
'            .CurrentX = Regresa_Valor("CONTRATO", "HorarioX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "HorarioY", 0)
'            ImpresoraContratos.Print Regresa_Valor_BD("Horario")
           
            
            .Font.Size = 6
             'Responsable
            .CurrentX = Regresa_Valor("CONTRATO", "ClienteDesX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ClienteDesY", 0)
            ImpresoraContratos.Print rcConsulta!Cliente
            
'''            'Tasa de Iva
'''            .CurrentX = Regresa_Valor("CONTRATO", "TasaIvaX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "TasaIvaY", 0)
'''            ImpresoraContratos.Print Format(rcConsulta!Iva, "0.00")
            
            .FontSize = 8 '10
'            .Font.Bold = True
            'Horario
'            .CurrentX = Regresa_Valor("CONTRATO", "HorarioX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "HorarioY", 0)
'            ImpresoraContratos.Print Regresa_Valor_BD("Horario")
            .Font.Bold = False
            
            .CurrentX = Regresa_Valor("CONTRATO", "NotaX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "NotaY", 0)
            ImpresoraContratos.Print Regresa_Valor_BD("Notas")
            
            
            
            'DiaFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "DiaFirmasX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DiaFirmasY", 0)
            ImpresoraContratos.Print Day(rcConsulta!Fecha)
            
            'MesFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "MesFirmasX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "MesFirmasY", 0)
            ImpresoraContratos.Print Month(rcConsulta!Fecha)
            
            'YearFirmas
            .CurrentX = Regresa_Valor("CONTRATO", "YearFirmasX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "YearFirmasY", 0)
            ImpresoraContratos.Print Year(rcConsulta!Fecha)
            
             'Consumidor
            .CurrentX = Regresa_Valor("CONTRATO", "ClienteFirmaX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ClienteFirmaY", 0)
            ImpresoraContratos.Print rcConsulta!Cliente
            
            
            'Consumidor
            '.CurrentX = Regresa_Valor("CONTRATO", "ConsumidorX", 0)
            '.CurrentY = Regresa_Valor("CONTRATO", "ConsumidorY", 0)
            'ImpresoraContratos.Print rcConsulta!Cliente
            
            .Font.Size = 8
            'Valuador
            .CurrentX = Regresa_Valor("CONTRATO", "ValuadorX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "ValuadorY", 0)
            ImpresoraContratos.Print rcConsulta!Valuador
            
'            .FontSize = 8
            'Horario
            .CurrentX = Regresa_Valor("CONTRATO", "HorarioX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "HorarioY", 0)
            ImpresoraContratos.Print Regresa_Valor_BD("HorarioSucursal")
            
            
            'Imprimo el domicilio sucursal 2
            .CurrentX = Regresa_Valor("CONTRATO", "DireccionSucursal2X", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "DireccionSucursal2Y", 0)
            ImpresoraContratos.Print Sucursal.Direccion & " " & Sucursal.Ciudad & " " & Sucursal.Estado

            'Telefono Sucursal 2
            .CurrentX = Regresa_Valor("CONTRATO", "TelefonoSucursal2X", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "TelefonoSucursal2Y", 0)
            ImpresoraContratos.Print Sucursal.Telefono
            
            'Correo Sucursal 2
            .CurrentX = Regresa_Valor("CONTRATO", "EmailSucursal2X", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "EmailSucursal2Y", 0)
            ImpresoraContratos.Print SacaValor("sucursales", "CorreoAclaraciones", " Where Clave=" & Sucursal.Clave)
            
            'Pagina Internet
            .CurrentX = Regresa_Valor("CONTRATO", "PaginaInternetX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "PaginaInternetY", 0)
            ImpresoraContratos.Print "www.mrayudon.com"
            
            .Font = "3 of 9 Barcode"
            .FontSize = 24

            'SuajeCodigo
            .CurrentX = Regresa_Valor("CONTRATO", "SuajeCodigoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "SuajeCodigoY", 0)
            ImpresoraContratos.Print rcConsulta!NumContrato

'''            .Font = "Arial Narrow"
'''            .FontBold = True
'''            .FontSize = 16
'''
            .Font = "Arial Narrow"
            .FontBold = False
            .FontSize = 10
            
            'SuajeContrato
            .CurrentX = Regresa_Valor("CONTRATO", "SuajeContratoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "SuajeContratoY", 0)
            ImpresoraContratos.Print rcConsulta!NumContrato
            
            .FontSize = 8

            'SuajeCliente
            .CurrentX = Regresa_Valor("CONTRATO", "SuajeClienteX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "SuajeClienteY", 0)
            ImpresoraContratos.Print rcConsulta!Cliente

            'SuajePrestamo
            .CurrentX = Regresa_Valor("CONTRATO", "SuajePrestamoX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "SuajePrestamoY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Prestamo, FMonedaSigno)

'''            'SuajeTotalPeso
'''            .CurrentX = Regresa_Valor("CONTRATO", "SuajeTotalPesoX", 0)
'''            .CurrentY = Regresa_Valor("CONTRATO", "SuajeTotalPesoY", 0)
'''            ImpresoraContratos.Print "PESO: " & PesoTotal & " GRMS."

            'SuajePrenda
            DescPrendaY = Regresa_Valor("CONTRATO", "SuajePrendaY", 0)
            
            rcAux.Open "SELECT de.marcaymodelo,de.año,de.placas,de.color,de.nummotor,de.serieChasis,de.poliza,de.gas,de.kms,de.factura,de.observaciones,e.prestamo,e.avaluo " & _
                       "FROM detallesempenoautos de inner join empeno e on de.IDEmpeno=e.ID WHERE de.IDEmpeno=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
            While Not rcAux.EOF
            strDescripcion = ""
                    x = 0
                    For i = 1 To Len(rcAux!MarcayModelo & " " & IIf(IsNull(rcAux!Año) Or Trim(rcAux!Año) = "", "", " AÑO: " & rcAux!Año) & IIf(IsNull(rcAux!Placas) Or Trim(rcAux!Placas) = "", "", " PLACAS: " & rcAux!Placas) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & IIf(IsNull(rcAux!NumMotor) Or Trim(rcAux!NumMotor) = "", "", " NUM MOTOR: " & rcAux!NumMotor) & IIf(IsNull(rcAux!SerieChasis) Or Trim(rcAux!SerieChasis) = "", "", " Serie: " & rcAux!SerieChasis) & IIf(IsNull(rcAux!Poliza) Or Trim(rcAux!Poliza) = "", "", " POLIZA: " & rcAux!Poliza) & IIf(IsNull(rcAux!Gas) Or Trim(rcAux!Gas) = "", "", " GAS: " & rcAux!Gas) & IIf(IsNull(rcAux!Kms) Or Trim(rcAux!Kms) = "", "", " KMS: " & rcAux!Kms) & IIf(IsNull(rcAux!Factura) Or Trim(rcAux!Factura) = "", "", " Factura: " & rcAux!Factura) & IIf(IsNull(rcAux!Observaciones) Or Trim(rcAux!Observaciones) = "", "", " OBSERVACIONES: " & rcAux!Observaciones)) Step 80
                        
                        .CurrentX = Regresa_Valor("CONTRATO", "SuajePrendaX", 0)
                        .CurrentY = DescPrendaY + (2.5 * x)
                        
                        strDescripcion = LTrim(Mid(rcAux!MarcayModelo & " " & IIf(IsNull(rcAux!Año) Or Trim(rcAux!Año) = "", "", " AÑO: " & rcAux!Año) & IIf(IsNull(rcAux!Placas) Or Trim(rcAux!Placas) = "", "", " PLACAS: " & rcAux!Placas) & IIf(IsNull(rcAux!Color) Or Trim(rcAux!Color) = "", "", " COLOR: " & rcAux!Color) & IIf(IsNull(rcAux!NumMotor) Or Trim(rcAux!NumMotor) = "", "", " NUM MOTOR: " & rcAux!NumMotor) & IIf(IsNull(rcAux!SerieChasis) Or Trim(rcAux!SerieChasis) = "", "", " Serie: " & rcAux!SerieChasis) & IIf(IsNull(rcAux!Poliza) Or Trim(rcAux!Poliza) = "", "", " POLIZA: " & rcAux!Poliza) & IIf(IsNull(rcAux!Gas) Or Trim(rcAux!Gas) = "", "", " GAS: " & rcAux!Gas) & IIf(IsNull(rcAux!Kms) Or Trim(rcAux!Kms) = "", "", " KMS: " & rcAux!Kms) & IIf(IsNull(rcAux!Factura) Or Trim(rcAux!Factura) = "", "", " Factura: " & rcAux!Factura) & IIf(IsNull(rcAux!Observaciones) Or Trim(rcAux!Observaciones) = "", "", " OBSERVACIONES: " & rcAux!Observaciones), i * 1, 80 + IIf(Mid(strDescripcion, i, 1) = " ", 1, 0)))
                        ImpresoraContratos.Print strDescripcion
                        x = x + 1
                    Next i
            rcAux.MoveNext
            DescPrendaY = DescPrendaY + (2.5 * (x - 1)) + 2.5
            Wend
            rcAux.Close
            
            
            'SuajePrestamo Letra
            .CurrentX = Regresa_Valor("CONTRATO", "SuajePrestamoLetraX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "SuajePrestamoLetraY", 0)
            ImpresoraContratos.Print CantidadEnLetra(rcConsulta!Prestamo)
            
            'SuajeFecha
            .CurrentX = Regresa_Valor("CONTRATO", "SuajeFechaX", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "SuajeFechaY", 0)
            ImpresoraContratos.Print Format(rcConsulta!Fecha, "DD/MMM/YYYY")
            
            .FontSize = Regresa_Valor("CONTRATO", "SucursalEnc2FS", 0)
            
            .CurrentX = Regresa_Valor("CONTRATO", "SucursalEnc2X", 0)
            .CurrentY = Regresa_Valor("CONTRATO", "SucursalEnc2Y", 0)
            ImpresoraContratos.Print Sucursal.NombreComercial
            
            .FontSize = 8
'            'SuajeBolsa
'            .CurrentX = Regresa_Valor("CONTRATO", "SuajeBolsaX", 0)
'            .CurrentY = Regresa_Valor("CONTRATO", "SuajeBolsaY", 0)
'            ImpresoraContratos.Print "BOLSA: " & rcConsulta!NumBolsa
            
            
            'MLD.- MODIF -----------------------------------------------------------------
            
'            Dim RsCte As New ADODB.Recordset
'            Dim SqlCte As String
'
'            .FontBold = False
'            .FontSize = 8
'
'            SqlCte = "SELECT c.Nombre,c.ApellidoPaterno,c.ApellidoMaterno,c.Apellido,c.FecNac,if(p.Descripcion is Null ,'MEXICO',p.Descripcion) AS PaisNacimiento,if(n.Descripcion is null, 'MEXICO',n.Descripcion) AS PaisNacionalidad," & _
'                     "c.Direccion,c.NoExterior,c.NoInterior,c.Colonia,c.Municipio,c.Estado,c.Tel,c.NumeroIdentificacion,c.CP,c.Email,c.Rfc,c.Curp,o.Descripcion AS Ocupacion,i.Descripcion AS TipoIdentificacion, i.Dependencia as Expide " & _
'                     "FROM clientes AS c Left Join mld_paises AS p ON c.IdPaisNacimiento = p.Id Left Join mld_tipo_identificaciones AS i ON c.IdTipoIdent = i.Id Left Join mld_paises AS n ON c.IdPaisNacionalidad = n.Id Left Join mld_actividades_economicas AS o ON c.IdOcupacion = o.Id " & _
'                     "WHERE c.Id=" & rcConsulta!IDCliente
'
'            RsCte.Open SqlCte, dbDatos, adOpenForwardOnly, adLockOptimistic
'            If Not RsCte.EOF Then
'
'                'Expediente Contrato
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ContratoX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ContratoY", 0)
'                ImpresoraContratos.Print rcConsulta!NumContrato
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ApellidoPX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ApellidoPY", 0)
'                ImpresoraContratos.Print IIf(Trim(RsCte!ApellidoPaterno) = "", RsCte!Apellido, RsCte!ApellidoPaterno)
'
'                'Expediente ApellidoM
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ApellidoMX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ApellidoMY", 0)
'                ImpresoraContratos.Print RsCte!ApellidoMaterno
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NombreX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NombreY", 0)
'                ImpresoraContratos.Print RsCte!Nombre
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "FechaNacX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "FechaNacY", 0)
'                ImpresoraContratos.Print RsCte!FecNac
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "PaisNacX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "PaisNacY", 0)
'                ImpresoraContratos.Print RsCte!PaisNacimiento
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NacionalidadX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NacionalidadY", 0)
'                ImpresoraContratos.Print RsCte!PaisNacionalidad
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "CalleX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "CalleY", 0)
'                ImpresoraContratos.Print RsCte!Direccion
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumExtX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumExtY", 0)
'                ImpresoraContratos.Print RsCte!NoExterior
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumIntX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumIntY", 0)
'                ImpresoraContratos.Print RsCte!NoInterior
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ColoniaX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ColoniaY", 0)
'                ImpresoraContratos.Print RsCte!Colonia
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "MunicipioX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "MunicipioY", 0)
'                ImpresoraContratos.Print RsCte!Municipio
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "EstadoX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "EstadoY", 0)
'                ImpresoraContratos.Print RsCte!Estado
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "CpX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "CpY", 0)
'                ImpresoraContratos.Print RsCte!CP
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "TelefonoX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "TelefonoY", 0)
'                ImpresoraContratos.Print RsCte!Tel
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "EmailX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "EmailY", 0)
'                ImpresoraContratos.Print RsCte!Email
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "RFCX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "RFCY", 0)
'                ImpresoraContratos.Print IIf(IsNull(RsCte!RFC), "", RsCte!RFC)
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "CURPX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "CURPY", 0)
'                ImpresoraContratos.Print IIf(Trim(RsCte!Curp) = "", "", RsCte!Curp)
'
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "TipoIdentX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "TipoIdentY", 0)
'                If IsNull(RsCte!tipoidentificacion) = True Then
'                    ImpresoraContratos.Print UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE RegDefault=1"))
'                Else
'                    ImpresoraContratos.Print UCase(RsCte!tipoidentificacion)
'                End If
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumIdentX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "NumIdentY", 0)
'                ImpresoraContratos.Print RsCte!NumeroIdentificacion
'
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "ExpideX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "ExpideY", 0)
'                ImpresoraContratos.Print UCase(IIf(IsNull(RsCte!expide), SacaValor("mld_tipo_identificaciones", "Dependencia", " WHERE RegDefault=1"), RsCte!expide))
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "OcupacionX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "OcupacionY", 0)
'                ImpresoraContratos.Print IIf(IsNull(RsCte!ocupacion), "", RsCte!ocupacion)
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "CiudadSucX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "CiudadSucY", 0)
'                ImpresoraContratos.Print SacaValor("sucursales", "Ciudad", " WHERE Activa=1") 'RsCte!CiudadSucursal
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "DiaX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "DiaY", 0)
'                ImpresoraContratos.Print Day(rcConsulta!Fecha)
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "MesX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "MesY", 0)
'                ImpresoraContratos.Print UCase(Format(rcConsulta!Fecha, "MMMM"))
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "AnoX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "AnoY", 0)
'                ImpresoraContratos.Print Year(rcConsulta!Fecha)
'
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "FirmaValuadorX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "FirmaValuadorY", 0)
'                ImpresoraContratos.Print UCase(rcConsulta!valuador)
'
'                'Expediente ApellidoP
'                .CurrentX = Regresa_Valor("EXPEDIENTE_CLIENTE", "FirmaClienteX", 0)
'                .CurrentY = Regresa_Valor("EXPEDIENTE_CLIENTE", "FirmaClienteY", 0)
'                ImpresoraContratos.Print rcConsulta!Cliente
'
'            End If
'            RsCte.Close
'            Set RsCte = Nothing
'
'
            .EndDoc
        
    End With
        
    Next
    
    rcConsulta.Close
    Set rcConsulta = Nothing
    
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Public Function Regresa_Impresora(Opcion As eTipoImpresora, ByRef prnImpresora As Printer) As Printer
    Dim prn As Printer
    Dim Impresora As String
   
    Select Case Opcion
    Case eTipoImpresora.Contratos
        Impresora = Regresa_Valor("Impresoras", "ImpresoraContratos", Printer.DeviceName)
    Case eTipoImpresora.Tickets
        Impresora = Regresa_Valor("Impresoras", "ImpresoraTickets", Printer.DeviceName)
    Case eTipoImpresora.EtiquetasEmpeno
        Impresora = Regresa_Valor("Impresoras", "ImpresoraEtiquetas", Printer.DeviceName)
    Case eTipoImpresora.EtiquetasAlmoneda
        Impresora = Regresa_Valor("Impresoras", "ImpresoraEtiquetasAlmoneda", Printer.DeviceName)
    End Select
               
    For Each prn In Printers
        If prn.DeviceName = Impresora Then
            Set Printer = prn
            Set prnImpresora = Printer
            Exit Function
        End If
   Next
End Function

Function LeyendaPromocion(Opcion As Integer) As String
Dim strPromocion As String

    strPromocion = ""
    Select Case Opcion
    Case 1
        
        strPromocion = "CAMBIATE."
    
    Case 2
        
        strPromocion = "5% DE DESCUENTO."
    
    Case 3
    
        strPromocion = "10% DE DESCUENTO."
    
    Case 4
    
        strPromocion = "15% DE DESCUENTO."
    
    Case 5
    
        strPromocion = "20% DE DESCUENTO."
    
    Case 15, 30
    
        strPromocion = Opcion & " DIAS."
    
    Case 20, 40
        
        strPromocion = "$" & Format(Opcion, FMoneda) & " DESC."
    End Select
    
    LeyendaPromocion = strPromocion
End Function

Public Sub DatosSucursal(Optional Inicio As Boolean = True, Optional Clave As Integer = 0)
Dim rcSucursal As New ADODB.Recordset

    rcSucursal.Open "SELECT Clave,RazonSocial,NombreComercial,RFC,Direccion,Ciudad,Estado,Telefono,CP FROM sucursales WHERE " & IIf(Inicio, "Activa=1", "Clave=" & Clave), dbDatos, adOpenForwardOnly, adLockReadOnly
        frmMDI.IDSucursal = rcSucursal!Clave
        Sucursal.Clave = rcSucursal!Clave
        Sucursal.RazonSocial = rcSucursal!RazonSocial
        Sucursal.NombreComercial = rcSucursal!NombreComercial
        Sucursal.RFC = rcSucursal!RFC
        Sucursal.Direccion = rcSucursal!Direccion
        Sucursal.Ciudad = rcSucursal!Ciudad
        Sucursal.Estado = rcSucursal!Estado
        Sucursal.Telefono = rcSucursal!Telefono
        Sucursal.CP = rcSucursal!CP
    rcSucursal.Close
    Set rcSucursal = Nothing
            
End Sub

Function ConvMoneda(Valor) As String
'''''Dim strValor As String
'''''
'''''    If Separador = "," Then
'''''
'''''        strValor = Replace(CStr(Valor), ".", "")
'''''        strValor = Replace(CStr(strValor), ",", ".")
'''''        ConvMoneda = strValor
'''''    Else
        
        ConvMoneda = CDbl(Valor)
'''''    End If
    
End Function

'Función quie retorna la cadena con el resultado
'************************************************
Public Function Obtener_Separador_Decimal() As String
Dim Buffer As String, Ret As Long

    Buffer = String(255, " ")
    
    'Ejecutamos el Api. En el Buffer obtenermos el separador
    Ret = GetLocaleInfo(GetUserDefaultLCID, LOCALE_SDECIMAL, Buffer, 255)
    
    'Quitamos los espacios nulos
    Obtener_Separador_Decimal = Trim$(Replace$(Buffer, Chr(0), ""))

End Function

Public Sub CreateToolBars()

    'frmMDI.CommandBars.Icons = frmMDI.ImageManager1.Icons
    frmMDI.CommandBars.Icons = frmMDI.Imagenes.Icons
    Toolbar.Closeable = False
    Toolbar.Customizable = False
    
    frmMDI.CommandBars.TabWorkspace.ThemedBackColor = False
    frmMDI.CommandBars.VisualTheme = xtpThemeVisualStudio2008
    
    With Toolbar
        .SetIconSize 32, 32
        AddButton .Controls, xtpControlButton, eToolBar.Buscar, , , "Búsqueda Contratos"
        AddButton .Controls, xtpControlButton, eToolBar.Empeno, , , "Empeño"
        AddButton .Controls, xtpControlButton, eToolBar.Cierre, , , "Balance"
        AddButton .Controls, xtpControlButton, eToolBar.Venta, , , "Ventas Mostrador"
'''''        AddButton .Controls, xtpControlButton, eToolBar.Divisas, , , "Compra/Venta Divisas"
        AddButton .Controls, xtpControlButton, eToolBar.Salir, True, , "Salir del sistema"
    End With
    
End Sub

Public Sub CreateStatusBar()
        
    stBar.Visible = True
    With stBar
                                
        PaneSucursal.Style = SBPS_STRETCH
        PaneSucursal.Customizable = True
        PaneSucursal.Alignment = xtpAlignmentCenter
        PaneSucursal.BeginGroup = True
        
        .AddPane eStatusBar.ID_INDICATOR_CAPS
        .AddPane eStatusBar.ID_INDICATOR_NUM
        .AddPane eStatusBar.ID_INDICATOR_SCRL
    End With
End Sub

Public Function AddButton(Controls As CommandBarControls, ControlType As XTPControlType, ID As Long, Optional BeginGroup As Boolean = False, Optional ButtonStyle As XTPButtonStyle = xtpButtonAutomatic, Optional ToolTip As String = "") As CommandBarControl
Dim Control As CommandBarControl
    
    Set Control = Controls.Add(ControlType, ID, "")
    
    Control.ToolTipText = ToolTip
    Control.BeginGroup = BeginGroup
    Control.Style = ButtonStyle
    
    Set AddButton = Control
End Function

'regresamos los nombres de las secciones del archivo ini
Public Function Regresa_Secciones() As String
    
    Dim Cadena As String, Lon As Integer
    
    Cadena = String(255, 0)
    Lon = GetPrivateProfileSectionNames(Cadena, 255, App.Path & "\Configuracion.Ini")
    
    Cadena = Left$(Cadena, Lon)
    Regresa_Secciones = Cadena
    
End Function

'regresamos todos los valores de una seccion
Public Function Regresa_Seccion_Valores(Seccion As String) As String
    
    Dim Cadena As String, Lon As Integer
    
    Cadena = String(4096, 0)
    Lon = GetPrivateProfileSection(Seccion, Cadena, 4096, App.Path & "\Configuracion.Ini")
    
    Cadena = Left$(Cadena, Lon)
    Regresa_Seccion_Valores = Cadena
    
End Function

'Grabamos el valor de una key
Public Function Graba_Valor(Seccion As String, Key As String, Valor As String) As Boolean
    WritePrivateProfileString Seccion, Key, Valor, App.Path & "\Configuracion.Ini"
End Function

Public Sub Pase_Automatico_Almoneda()
Dim Indice As Long, Movimiento As Long, Folio As Long, Prestamo As Double, Cantidad As Integer, Serie As Integer, IDEntrada As Long, strDestino As String, strDescripcion As String, Hora As String, strCodigo As String, Kilates As Integer, crPrecio As Double, AvaluoDiam As Double, GTOSVenta As Double, GtosComer As Double, Iva As Double, DiasEnaje As Integer
Dim rcRemate As New ADODB.Recordset
Dim rcTmp As New ADODB.Recordset
Dim rcAux As New ADODB.Recordset
Dim FechaAlmoneda As Date

On Error GoTo Error
    FechaAlmoneda = Regresa_Valor_BD("FechaAlmoneda")
    DiasEnaje = Regresa_Valor_BD("DiasEnajenacion")
    rcRemate.Open "SELECT COUNT(e.ID) AS Total FROM empeno e WHERE e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL (" & DiasEnaje & ") DAY),'%Y%/%m%/%d')<'" & Format(FechaAlmoneda, "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
        If rcRemate!Total > 0 Then frmMDI.Bar.Value = 0: frmMDI.Bar.Min = 0: frmMDI.Bar.Max = rcRemate!Total Else Exit Sub
    rcRemate.Close
        
    If frmMDI.Bar.Max > 0 Then
                
        Screen.MousePointer = vbHourglass
        frmMDI.Bar.Visible = True
        
        'Checo los contratos que estan vencidos
        rcRemate.Open "SELECT DISTINCT e.ID,e.NumContrato,e.Folio,e.Fecha,e.Prestamo,e.Avaluo,e.Origen,e.Vencimiento,e.TipoInteres,e.TipoTasa,e.Serie,c.Iniciales,CONCAT(c.Apellido,' ',c.Nombre) AS Cliente " _
                    & "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL (" & Regresa_Valor_BD("DiasEnajenacion") & ") DAY),'%Y%/%m%/%d')<'" & Format(FechaAlmoneda, "YYYY/MM/DD") & "' ORDER BY e.NumContrato", dbDatos, adOpenForwardOnly, adLockReadOnly

        'Saco el Folio
        Folio = Regresa_Movimiento(False, "FolioInventario")
        Regresa_Movimiento True, "FolioInventario"
                
        'Saco el movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
        
        'Tabla Entrada Inventario
        dbDatos.Execute "INSERT INTO entradainventario(Fecha,Folio,TipoEntrada,IDUsuario,IDSucursal) VALUES " _
                        & "('" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & ENTRADAALMONEDA & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
        'Saco el ID de la Entrada
        IDEntrada = SacaValor("entradainventario", "MAX(ID)")
        
        'Tomo la Hora
        Hora = Time
        
        'Tomo los Gastos de Venta
        GTOSVenta = Regresa_Valor_BD("GtosVenta") / 100
        
        'Tomo los Gastos de Comercialización
'        GtosComer = Regresa_Valor_BD("GtosComer") / 100
        
        'Tomo el IVA
        Iva = Regresa_Valor_BD("IVA") / 100
        
        dbReportes.Execute "DELETE FROM articulos"
        While Not rcRemate.EOF
                                        
            'Marco la prenda como pasada a Remate
            If rcRemate!Serie = SERIE_B Then
                
                rcTmp.Open "SELECT de.IDEmpeno,de.MarcayModelo,de.Placas,de.Año,de.Color,de.SerieChasis,de.NumMotor,de.NumTarjetaCircu FROM detallesempenoautos de WHERE de.IDEmpeno=" & rcRemate!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
                
                'strDescripcion = "MARCA Y MODELO: " & rcRemate!MarcayModelo & ", PLACAS: " & rcRemate!Placas & ", AÑO: " & rcRemate!Año & ", COLOR: " & rcRemate!Color & ", SERIE CHASIS: " & rcRemate!SerieChasis & ", NUM. MOTOR: " & rcRemate!NumMotor & ", TARJETA CIRC.: " & rcRemate!NumTarjetaCircu
                strDescripcion = "MARCA Y MODELO: " & rcTmp!MarcayModelo & ", PLACAS: " & rcTmp!Placas & ", AÑO: " & rcTmp!Año & ", COLOR: " & rcTmp!Color & ", SERIE CHASIS: " & rcTmp!SerieChasis & ", NUM. MOTOR: " & rcTmp!NumMotor & ", TARJETA CIRC.: " & rcTmp!NumTarjetaCircu
                
                rcTmp.Close
                
                dbReportes.Execute "INSERT INTO articulos (IDEmpeno,Articulo,Avaluo,Prestamo) VALUES (" & _
                                            rcRemate!ID & ",'" & strDescripcion & "'," & rcRemate!Avaluo & "," & rcRemate!Prestamo & ")"
                
                'Marco el contrato como pasado a Almoneda
                dbDatos.Execute "UPDATE empeno SET IDEntradaInventario =" & IDEntrada & ", Pagado=1,Destino=" & D_ALMONEDA & ", FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & rcRemate!ID
                
                crPrecio = rcRemate!Prestamo * (1 + Regresa_Valor_BD("PrecioAutos") / 100)
                
                'Paso el Automovil al Inventario
                dbDatos.Execute "INSERT INTO detallesentradainventario(IDEntrada,Codigo,Tipo,Cantidad,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,IDEmpeno,SucursalOrigen,TipoEntrada,PrecioVitrina) VALUES (" & _
                                IDEntrada & ",'" & strCodigo & "',0,1,'" & strDescripcion & "',0,0," & rcRemate!Avaluo & "," & rcRemate!Prestamo & ",'','','','','','',0,''," & rcRemate!ID & "," & frmMDI.IDSucursal & "," & ENTRADAALMONEDA & "," & ConvMoneda(crPrecio) & ")"

                
                'Grabamos el cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & Folio & ",'RE01','620301'," & ConvMoneda(Prestamo) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                         
                'Grabamos el abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & Folio & ",'RE50','201750'," & ConvMoneda(Prestamo) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

            Else
                
                'Marco el contrato como marcado a Remate
                dbDatos.Execute "UPDATE empeno SET Pagado=1,Destino=" & D_ALMONEDA & ", FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',IDEntradaInventario=" & IDEntrada & ",Almoneda=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & rcRemate!ID
                
                rcTmp.Open "SELECT de.ID,de.IDEmpeno,de.Codigo,de.Cantidad,de.Articulo,de.Peso,de.Kilates,de.Estado,de.Avaluo,de.Prestamo,de.Tipo,de.Observaciones FROM detallesempeno de WHERE de.IDEmpeno=" & rcRemate!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
                With rcTmp
                    
                    While Not rcTmp.EOF
                            
                        If !Tipo = 1 Or !Tipo = 10 Then
                            
                            strDestino = "FUNDICION"
                        Else
                            
                            strDestino = "VENTA"
                        
                        End If
                        
                        dbReportes.Execute "INSERT INTO articulos (IDEmpeno,Articulo,Peso,Kilates,Avaluo,Prestamo,Cantidad,Destino,Observaciones) VALUES (" & _
                                            !IDEmpeno & ",'" & !Articulo & "'," & ConvMoneda(!Peso) & "," & Val(!Kilates) & "," & ConvMoneda(!Avaluo) & "," & ConvMoneda(!Prestamo) & "," & Val(!Cantidad) & ",'" & strDestino & "','" & !Observaciones & "')"
                        
                        dbDatos.Execute "UPDATE detallesempeno SET Almoneda=1,Destino=" & IIf(strDestino = "VENTA", D_VENTA, D_FUNDICION) & "  WHERE ID=" & !ID
                        
                        'Paso las prendas al Inventario
                        With rcAux
                            
                            .Open "SELECT e.ID,e.Fecha,e.Folio,e.Vencimiento,e.TipoInteres,d.ID AS IDPrenda,d.Articulo,d.Kilates,d.Peso,d.Avaluo,d.Cantidad,d.Prestamo,d.Avaluo,d.Tipo,d.IDEmpeno,d.Marca,d.Modelo,d.Serie,d.Color,d.Tamano,d.Codigo,d.TipoPrenda,d.Observaciones,d.Estado,d.CantidadPiedras,d.PesoPiedras,d.CantidadDiamantes,d.Puntos,d.PrestamoDiamante,k.Descripcion AS Kilataje FROM detallesempeno d LEFT JOIN kilatajes k ON d.Kilates=k.Clave INNER JOIN empeno e ON d.IDEmpeno=e.ID WHERE d.ID=" & rcTmp!ID & " ORDER BY d.Codigo", dbDatos, adOpenForwardOnly, adLockReadOnly
                            
                            'Calculo el Precio de la Prenda
                            Kilates = IIf(IsNull(!Kilates), 0, !Kilates)
                            
                            'Saco el Precio
                            crPrecio = Redondeo((!Prestamo + Redondeo(Redondeo((!Prestamo * GTOSVenta))) * (1 + Iva)))
                                                
                            'Tabla de DetalleEntradaInventario
                            dbDatos.Execute "INSERT INTO detallesentradainventario(IDEntrada,Codigo,Tipo,Cantidad,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,IDEmpeno,SucursalOrigen,TipoEntrada,PrecioVitrina,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante) VALUES (" & _
                                            IDEntrada & ",'" & !Codigo & "'," & !Tipo & "," & !Cantidad & ",'" & !Articulo & "'," & ConvMoneda(!Peso) & "," & !Kilates & "," & ConvMoneda(crPrecio) & "," & ConvMoneda(!Prestamo) & ",'" & !Estado & "','" & !Marca & "','" & !Modelo & "','" & !Serie & "','" & !Color & "','" & !Tamano & "'," & !TipoPrenda & ",'" & !Observaciones & "'," & !IDEmpeno & "," & frmMDI.IDSucursal & "," & IIf(strDestino = "VENTA", D_ALMONEDA, D_FUNDICION) & "," & ConvMoneda(crPrecio) & "," & !CantidadPiedras & "," & ConvMoneda(!PesoPiedras) & "," & !CantidadDiamantes & "," & ConvMoneda(!Puntos) & "," & ConvMoneda(!PrestamoDiamante) & ")"
                            
                            .Close
                            Set rcAux = Nothing
                        
                        End With
                       
                       'Muevo las Cuentas Contables
                        If strDestino = "VENTA" Then
                                                        
                            'Grabamos el cargo
                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE01','620301'," & ConvMoneda(!Prestamo) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                     
                            'Grabamos el abono
                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE50','201750'," & ConvMoneda(!Prestamo) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                            
                        Else
                                                        
                            'Grabamos el cargo
                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE01','310101'," & ConvMoneda(!Prestamo) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                                     
                            'Grabamos el abono
                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE50','201750'," & ConvMoneda(!Prestamo) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                        End If
                    
                    rcTmp.MoveNext
                    Wend
                
                End With
                rcTmp.Close
            
            End If
        
        frmMDI.Bar.Value = frmMDI.Bar.Value + 1
        rcRemate.MoveNext
        Wend
        rcRemate.Close
        Set rcRemate = Nothing
        
        frmMDI.Bar.Visible = False
        
        Sleep 1000
        With frmMDI.Cr
            .Reset
            .DiscardSavedData = True
            .WindowShowPrintSetupBtn = True
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .ReportFileName = Path & "\Reportes\ContratosAlmoneda.rpt"
            .SelectionFormula = "{empeno.IDEntradaInventario}=" & IDEntrada
            .Formulas(0) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
            .Formulas(1) = "Encabezado='" & Sucursal.RazonSocial & "'"
            .Formulas(2) = "Leyenda=''"
            
            .SubreportToChange = "Resumen"
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
                      .SQLQuery = " SELECT articulos1.`Cantidad`, articulos1.`Peso`, articulos1.`Kilates`, articulos1.`Prestamo`,kilatajes1.`Descripcion` " & _
    " From " & _
    " `BaseReportes`.`articulos` articulos1 INNER JOIN `BaseDatos`.`empeno` empeno1 ON articulos1.`IDEmpeno` = empeno1.`ID` INNER JOIN `BaseDatos`.`kilatajes` kilatajes1 ON articulos1.`Kilates` = kilatajes1.`Clave` " & _
    " WHERE " & _
    " articulos1.Kilates >0 AND articulos1.Destino= 'VENTA' " & _
    "Order By " & _
    "articulos1.`Kilates` ASC "
'            .SelectionFormula = "{articulos.Kilates} >0 AND {articulos.Destino}= 'VENTA' "

            .SubreportToChange = "ResumenFundicion"
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .SQLQuery = " SELECT articulos1.`Cantidad`, articulos1.`Peso`, articulos1.`Kilates`, articulos1.`Prestamo`,kilatajes1.`Descripcion` " & _
" From " & _
    " `BaseReportes`.`articulos` articulos1 INNER JOIN `BaseDatos`.`empeno` empeno1 ON articulos1.`IDEmpeno` = empeno1.`ID` INNER JOIN `BaseDatos`.`kilatajes` kilatajes1 ON articulos1.`Kilates` = kilatajes1.`Clave` " & _
    " WHERE " & _
    " articulos1.Kilates >0 AND articulos1.Destino= 'FUNDICION' " & _
    "Order By " & _
"articulos1.`Kilates` ASC "
'            .SelectionFormula = "{articulos.Kilates} >0 AND {articulos.Destino}= 'FUNDICION'"

            .WindowTitle = "Contratos a Almoneda"
            .Destination = crptToWindow
            .WindowState = crptMaximized
            .Action = 1
        End With

        Screen.MousePointer = vbDefault
        
    End If
    Set rcTmp = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcRemate = Nothing
    Set rcTmp = Nothing
    Screen.MousePointer = vbDefault
End Sub

Public Function OpenFileDialog(forma As Form) As String
Dim Archivo As String
Dim OFName As OPENFILENAME
   
   OFName.lStructSize = Len(OFName)
   'Set the parent window
   OFName.hwndOwner = forma.hWnd
   'Set the application's instance
   OFName.hInstance = App.hInstance
   'Select a filter
   OFName.lpstrFilter = "Archivos de Excel (*.xls)" + Chr$(0) + "*.xls" '+ Chr$(0) + "All Files (*.*)" + Chr$(0) + "*.*" + Chr$(0)
   'create a buffer for the file
   OFName.lpstrFile = Space$(254)
   'set the maximum length of a returned file
   OFName.nMaxFile = 255
   'Create a buffer for the file title
   OFName.lpstrFileTitle = Space$(254)
   'Set the maximum length of a returned file title
   OFName.nMaxFileTitle = 255
   'Set the initial directory
   OFName.lpstrInitialDir = App.Path  '"C:\"
   'Set the title
   OFName.lpstrTitle = "Seleccione el archivo de cuentas"
   'No flags
   OFName.flags = 0
   
   'Show the 'Open File'-dialog
   If GetOpenFileName(OFName) Then
       Archivo = Trim$(OFName.lpstrFile)
   Else
       Archivo = ""
   End If
    
    OpenFileDialog = Archivo
End Function

Function ChecaCandado() As Boolean
Dim strLicenciaDias As String, DiasExtra As Integer, DiasLicencia As Integer
    
    DiasExtra = 0
    DiasLicencia = Val(frmMDI.ActiveLock2.Tag)
    If Trim(Regresa_Valor("MONTEPIO", "Days", "")) <> "" Then
    
        strLicenciaDias = Base64Decode(Regresa_Valor("MONTEPIO", "Days", "0"))
        DiasExtra = Val(Mid(strLicenciaDias, Len(strLicenciaDias) - 2, 3))
    End If
    
    If frmMDI.ActiveLock2.RegisteredUser Then
    
        ChecaCandado = True
    
    ElseIf frmMDI.ActiveLock2.UsedDays <= (DiasLicencia + DiasExtra) Then
            
        ChecaCandado = True
    
    Else
        
        ChecaCandado = False
    End If
    
End Function

Public Function Base64Decode(ByVal base64String) As String
  'rfc1521
  '1999 Antonin Foller, Motobit Software, http://Motobit.cz
  Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
  Dim dataLength, sOut, groupBegin
  
  'remove white spaces, If any
  base64String = Replace(base64String, vbCrLf, "")
  base64String = Replace(base64String, vbTab, "")
  base64String = Replace(base64String, " ", "")
  
  'The source must consists from groups with Len of 4 chars
  dataLength = Len(base64String)
  If dataLength Mod 4 <> 0 Then
    Err.Raise 1, "Base64Decode", "Bad Base64 string."
    Exit Function
  End If

  
  ' Now decode each group:
  For groupBegin = 1 To dataLength Step 4
    Dim numDataBytes, CharCounter, thisChar, thisData, nGroup, pOut
    ' Each data group encodes up To 3 actual bytes.
    numDataBytes = 3
    nGroup = 0

    For CharCounter = 0 To 3
    
      ' Convert each character into 6 bits of data, And add it To
      ' an integer For temporary storage.  If a character is a '=', there
      ' is one fewer data byte.  (There can only be a maximum of 2 '=' In
      ' the whole string.)

      thisChar = Mid(base64String, groupBegin + CharCounter, 1)

      If thisChar = "=" Then
        numDataBytes = numDataBytes - 1
        thisData = 0
      Else
        thisData = InStr(1, Base64, thisChar, vbBinaryCompare) - 1
      End If
      If thisData = -1 Then
        Err.Raise 2, "Base64Decode", "Bad character In Base64 string."
        Exit Function
      End If

      nGroup = 64 * nGroup + thisData
    Next
    
    'Hex splits the long To 6 groups with 4 bits
    nGroup = Hex(nGroup)
    
    'Add leading zeros
    nGroup = String(6 - Len(nGroup), "0") & nGroup
    
    'Convert the 3 byte hex integer (6 chars) To 3 characters
    pOut = Chr(CByte("&H" & Mid(nGroup, 1, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 3, 2))) + _
      Chr(CByte("&H" & Mid(nGroup, 5, 2)))
    
    'add numDataBytes characters To out string
    sOut = sOut & Left(pOut, numDataBytes)
  Next

  Base64Decode = sOut
End Function

Public Sub Actualiza_Sistema()
   On Error GoTo Error

   If Verificar_Version Then
      MsgBox "Existe una nueva version del sistema disponible, favor de actualizar", vbOKOnly Or vbInformation
      Lanza_Actualizacion
    End If

Error:
    Maneja_Error Err
End Sub

Public Function Verificar_Version() As Boolean
   On Error GoTo Error
   Dim VersionBD As String
   Dim VersionSistema As String
   
   VersionBD = Regresa_Valor_BD("Version")
   VersionSistema = App.Major & "." & App.Minor & "." & App.Revision
   
   If VersionBD <> "" Then
      If VersionBD = VersionSistema Then
         Verificar_Version = False
      Else
         Verificar_Version = True
      End If
   Else
      Verificar_Version = False
   End If

Error:
   Maneja_Error Err
   
End Function

Public Sub Lanza_Actualizacion()
   Execute App.Path & "\wyUpdate.exe", frmMDI.hWnd
   End
End Sub

Public Sub Execute(FileName As String, OwnerhWnd As Long)
    Dim SEI As SHELLEXECUTEINFO
    Dim r As Long
    With SEI
        'Set the structure's size
        .cbSize = Len(SEI)
        'Seet the mask
        .fMask = SEE_MASK_DEFAULT 'SEE_MASK_NOCLOSEPROCESS Or SEE_MASK_INVOKEIDLIST Or SEE_MASK_FLAG_NO_UI
        'Set the owner window
        .hWnd = OwnerhWnd
        'Show the properties
        .lpVerb = "open"
        'Set the filename
        .lpFile = FileName
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        .nShow = 1
        .hInstApp = 0
        .lpIDList = 0
    End With
    r = ShellExecuteEx(SEI)
End Sub

'Lleno la Estructura de Empeños Foráneos
Public Function LlenarEmpenoForaneo(Folio As Long, Sucursal As Integer, Serie As Integer) As Boolean
Dim str As String, Empeno() As String, clientes() As String, DetallesEmpeno() As String, i As Integer, DEmpeno() As String
Dim lafecha As String
Dim curDate As Date
On Error GoTo Error

    Set WSoap = New SoapClient30
    WSoap.MSSoapInit WServidor & "?wsdl"

    str = WSoap.GetEmpeno("mrayudon", "montepio", WBaseDatos, WPuerto, WRutaServidor, Folio, Sucursal, Serie)
    Empeno = Split(str, "|")

    If UBound(Empeno) = -1 Then
        LlenarEmpenoForaneo = False
        Exit Function
    End If

    With Foraneo
        .NumContrato = getValueArray(Empeno(), "NumContrato", True)
        .Almacenaje = getValueArray(Empeno(), "Almacenaje", True)
        '.GastosAdmon = getValueArray(Empeno(), "GastosAdmon", True)
        .Avaluo = getValueArray(Empeno(), "Avaluo", True)
        .Beneficiario = getValueArray(Empeno(), "Beneficiario")
        .caja = getValueArray(Empeno(), "Caja")
        .Cajon = getValueArray(Empeno(), "Cajon")
        '.CedulaCotitular = getValueArray(Empeno(), "CedulaCotitular")
        .Comision = getValueArray(Empeno(), "Comision", True)
        lafecha = Trim(getValueArray(Empeno(), "Fecha"))
        
       
        'lafecha = "#" & lafecha & "#"
       ' curDate = CDate(lafecha)
                
        If Trim(getValueArray(Empeno(), "Fecha")) = "" Then
            .Fecha = Now
        ElseIf IsDate(CDate(Left$(getValueArray(Empeno(), "Fecha"), 10))) Then
            .Fecha = CDate(Left$(getValueArray(Empeno(), "Fecha"), 10))
        Else
            .Fecha = Now
        End If
        
'        If Trim(getValueArray(Empeno(), "FechaPagoParcial")) = "" Then
'            .FechaPagoParcial = Now
'        ElseIf IsDate(CDate(getValueArray(Empeno(), "FechaPagoParcial"))) Then
'            .FechaPagoParcial = CDate(getValueArray(Empeno(), "FechaPagoParcial"))
'        Else
'            .FechaPagoParcial = Now
'        End If

        .Fila = getValueArray(Empeno(), "Fila")
        .Folio = getValueArray(Empeno(), "Folio", True)
        .ID = getValueArray(Empeno(), "ID", True)
        .IDCliente = getValueArray(Empeno(), "IDCliente", True)
        .IDTablaCliente = getValueArray(Empeno(), "IDTablaCliente", True)
        .IDUsuario = getValueArray(Empeno(), "IDUsuario", True)
        .IDUsuarioAutoriza = getValueArray(Empeno(), "IDUsuarioAutoriza", True)
        .Notas = getValueArray(Empeno(), "Notas")
        .Iva = getValueArray(Empeno(), "Iva", True)
        .NumBolsa = getValueArray(Empeno(), "NumBolsa")
        .Origen = getValueArray(Empeno(), "Origen", True)
        .Operacion = getValueArray(Empeno(), "Operacion", True)
        .Perdida = getValueArray(Empeno(), "Perdida", True)
        .Periodo = getValueArray(Empeno(), "Periodo", True)
        .Prestamo = getValueArray(Empeno(), "Prestamo", True)
        .PrestamoInicial = getValueArray(Empeno(), "PrestamoInicial", True)
        .Responsable = getValueArray(Empeno(), "Responsable")
        .Seguro = getValueArray(Empeno(), "Seguro", True)
        .Serie = getValueArray(Empeno(), "Serie", True)
        .Cancelado = getValueArray(Empeno(), "Cancelado", True)
        .Destino = getValueArray(Empeno(), "Destino", True)
        .Tasa = getValueArray(Empeno(), "Tasa", True)
        .TipoAutoriza = getValueArray(Empeno(), "TipoAutoriza", True)
        .TipoInteres = getValueArray(Empeno(), "TipoInteres")
        .TipoTasa = getValueArray(Empeno(), "TipoTasa")
        .ubicacion = getValueArray(Empeno(), "Ubicacion")
        .Valuador = getValueArray(Empeno(), "Valuador")
        .VenAlmoneda = getValueArray(Empeno(), "VenAlmoneda", True)
        
        If Trim(getValueArray(Empeno(), "Vencimiento")) = "" Then
            .Vencimiento = Now
        Else
            lafecha = Left$(getValueArray(Empeno(), "Vencimiento"), 10)
            .Vencimiento = lafecha
        End If
       
        
        '.Vencimiento = getValueArray(Empeno(), "Vencimiento")
        .VenPeriodo = getValueArray(Empeno(), "VenPeriodo", True)
        '.MesesAcumulados = getValueArray(Empeno(), "MesesAcumulados", True)
        '.Migrada = getValueArray(Empeno(), "Migrada", True)
        '.FolioOriginal = getValueArray(Empeno(), "FolioOriginal", True)
        '.TipoPrenda = getValueArray(Empeno(), "TipoPrenda", True)
        '.Bloqueado = getValueArray(Empeno(), "Bloqueado", True)
        '.MotivoBloqueo = getValueArray(Empeno(), "MotivoBloqueo", True)
    End With

'    str = ""
'    str = WSoap.getclientes("root", "admin", WBaseDatos, WPuerto, WRutaServidor, Foraneo.IDTablaCliente, Sucursal)
'    clientes = Split(str, "|")
'
'    With ClienteForaneo
'        .ID = getValueArray(clientes(), "ID", True)
'        .Iniciales = getValueArray(clientes(), "Iniciales")
'        .Nombre = getValueArray(clientes(), "Nombre")
'        .Apellido = getValueArray(clientes(), "Apellido")
'        .Ciudad = getValueArray(clientes(), "Municipio")
'        .Colonia = getValueArray(clientes(), "Colonia")
'        .CP = getValueArray(clientes(), "CP", True)
'        .Direccion = getValueArray(clientes(), "Direccion")
'        .Estado = getValueArray(clientes(), "Estado")
'
'        If Trim(getValueArray(clientes(), "FecNac")) = "" Then
'            .FecNac = Now
'        ElseIf IsDate(CDate(getValueArray(clientes(), "FecNac"))) Then
'            .FecNac = CDate(getValueArray(clientes(), "FecNac"))
'        Else
'            .FecNac = Now
'        End If
'
'        .Identificacion = getValueArray(clientes(), "Identificacion")
'        .NumeroIdentificacion = getValueArray(clientes(), "NumeroIdentificacion")
'    End With
    
    Dim rc As New ADODB.Recordset
    
    rc.Open "SELECT * FROM Clientes WHERE IDTabla=" & Foraneo.IDTablaCliente, dbDatos, adOpenDynamic, adLockOptimistic
    
    If Not rc.EOF Then
        With ClienteForaneo
            .ID = rc!ID
            .Iniciales = rc!Iniciales & ""
            .Nombre = rc!Nombre & ""
            .Apellido = rc!Apellido & ""
            .Ciudad = rc!Municipio & ""
            .Colonia = rc!Colonia & ""
            .CP = Val(rc!CP & "")
            .Direccion = rc!Direccion & ""
            .Estado = rc!Estado & ""
                        
            If IsNull(rc!FecNac) Then
                .FecNac = Now
            Else
                .FecNac = rc!FecNac
            End If

            .Identificacion = rc!Identificacion & ""
            .NumeroIdentificacion = rc!NumeroIdentificacion & ""
        
        End With
    End If
    
    

    str = ""
    str = WSoap.GetDetallesEmpeno("mrayudon", "montepio", WBaseDatos, WPuerto, WRutaServidor, Foraneo.ID, Sucursal, Serie)
    
    If str = "" Then
        LlenarEmpenoForaneo = False
        Exit Function
    End If
    
    DetallesEmpeno = Split(str, "~")

    If Serie = 1 Then
    
        ReDim DetallesEmpenoForaneo(UBound(DetallesEmpeno) - 1)
    
        For i = 0 To UBound(DetallesEmpeno) - 1
            DEmpeno = Split(DetallesEmpeno(i), "|")
            With DetallesEmpenoForaneo(i)
                .Articulo = getValueArray(DEmpeno(), "Articulo")
                .Avaluo = getValueArray(DEmpeno(), "Avaluo", True)
                .Cantidad = getValueArray(DEmpeno(), "Cantidad", True)
                .CantidadDiamantes = getValueArray(DEmpeno(), "CantidadDiamantes", True)
                .CantidadPiedras = getValueArray(DEmpeno(), "CantidadPiedras", True)
                .Codigo = getValueArray(DEmpeno(), "Codigo")
                .Color = getValueArray(DEmpeno(), "Color")
                .Estado = getValueArray(DEmpeno(), "Estado")
                .IDEmpeno = getValueArray(DEmpeno(), "IDEmpeno", True)
                '.Kilates = getValueArray(DEmpeno(), "Kilates", True)
                .Kilates = getValueArray(DEmpeno(), "IDTablaKilates", True)
                .Marca = getValueArray(DEmpeno(), "Marca")
                .Modelo = getValueArray(DEmpeno(), "Modelo")
                .Observaciones = getValueArray(DEmpeno(), "Observaciones")
                .Origen = getValueArray(DEmpeno(), "Origen", True)
                .Peso = getValueArray(DEmpeno(), "Peso", True)
                .PesoPiedras = getValueArray(DEmpeno(), "PesoPiedras", True)
                .Prestamo = getValueArray(DEmpeno(), "Prestamo", True)
                .PrestamoDiamante = getValueArray(DEmpeno(), "PrestamoDiamante", True)
                .Puntos = getValueArray(DEmpeno(), "Puntos", True)
                .Serie = getValueArray(DEmpeno(), "Serie")
                .Tamano = getValueArray(DEmpeno(), "Tamano")
                '.Tipo = getValueArray(DEmpeno(), "IDTipo", True)
                .Tipo = getValueArray(DEmpeno(), "IDTablaTipo", True)
                .TipoPrenda = getValueArray(DEmpeno(), "TipoPrenda", True)
            End With
        Next i
        
    Else
    
        ReDim DetallesEmpenoFA(UBound(DetallesEmpeno) - 1)
        For i = 0 To UBound(DetallesEmpeno) - 1
            DEmpeno = Split(DetallesEmpeno(i), "|")
            
            With DetallesEmpenoFA(i)
                .IDEmpeno = getValueArray(DEmpeno(), "IDEmpeno", True)
                
                .MarcayModelo = getValueArray(DEmpeno(), "MarcayModelo")
                .Año = getValueArray(DEmpeno(), "Año", True)
                .Color = getValueArray(DEmpeno(), "Color")
                .Placas = getValueArray(DEmpeno(), "Placas")
                .Factura = getValueArray(DEmpeno(), "Factura")
                .Agencia = getValueArray(DEmpeno(), "Agencia")
                .NumTarjetaCircu = getValueArray(DEmpeno(), "NumTarjetacircu")
                .NumMotor = getValueArray(DEmpeno(), "NumMotor")
                .SerieChasis = getValueArray(DEmpeno(), "SerieChasis")
                .Kms = getValueArray(DEmpeno(), "Kms")
                .Gas = getValueArray(DEmpeno(), "Gas")
                .Aseguradora = getValueArray(DEmpeno(), "Aseguradora")
                .Poliza = getValueArray(DEmpeno(), "Poliza")
                
                If Trim(getValueArray(DEmpeno(), "FechaVenci")) = "" Then
                    .FechaVenci = Now
                ElseIf IsDate(CDate(getValueArray(DEmpeno(), "FechaVenci"))) Then
                    .FechaVenci = CDate(getValueArray(DEmpeno(), "FechaVenci"))
                Else
                    .FechaVenci = Now
                End If
                
                .Tipo = getValueArray(DEmpeno(), "Tipo", True)
                .Observaciones = getValueArray(DEmpeno(), "Observaciones", True)
                .TipoMovil = getValueArray(DEmpeno(), "TipoMovil", True)
                .TipoDesc = getValueArray(DEmpeno(), "TipoDesc", True)
            End With
        Next i
    End If

    str = ""
    'str = WSoap.GetNoRefrendos("root", "admin", WBaseDatos, WPuerto, WRutaServidor, Foraneo.NumContrato, Sucursal, Serie)
    
    If str = "" Then
        LlenarEmpenoForaneo = True
        Set WSoap = Nothing
        Exit Function
    End If

    Foraneo.NumRefrendos = Val(str)

    LlenarEmpenoForaneo = True
    Set WSoap = Nothing
    Exit Function

Error:
    Maneja_Error Err
    Set WSoap = Nothing
    LlenarEmpenoForaneo = False
End Function

Function GeneraInteresesForaneos(ByVal TipoDeTasa As String, ByVal Serie As Integer, ByVal crPrestamo As Double, ByVal crAvaluo As Double, ByVal Fecha As Date, ByVal TipoTasa As String, ByVal Vencimiento As Date, ByVal Tasa As Double, ByVal Iva As Double, ByVal Periodo As Integer, ByVal VenPeriodo As Integer, ByVal TipoInteres As String) As Double
Dim crInteres As Double, crAlmacenaje As Double, crSeguro As Double, crIva As Double, i As Integer, FechaOriginal As Date, DiasGracia As Integer

On Error GoTo Error

    FechaOriginal = Fecha
    
       'DiasGracia = SacaValor("configuraciontasas ct INNER JOIN tipoInteres ti ON ct.IDTipoInteres=ti.ID INNER JOIN tipoperiodo tp ON ct.IDTipoPeriodo=tp.ID INNER JOIN plazos p ON ct.IDPlazo=p.ID", "DiasGracia", " WHERE ti.Descripcion='" & TipoInteres & "' AND ti.Serie=" & Serie & " AND tp.Descripcion='" & TipoTasa & "' AND p.ID=" & VenPeriodo)
        DiasGracia = Regresa_Valor_BD("DiasGracia") 'SacaValor("ConfiguracionTasas", "DiasGracia", " ORDER BY ID Limit 1")
    
    For i = 1 To 500
      
        'Vencimiento
        If Periodo = 30 Then
            
            Fecha = DateAdd("M", i, FechaOriginal)
        Else
            
            Fecha = DateAdd("D", Periodo * i, FechaOriginal)
        End If
        
        'Tasa que se va a consultar
        crInteres = IIf(TipoDeTasa = "Tasa", crPrestamo, crAvaluo) * (Tasa * i)
        
        GeneraInteresesForaneos = crInteres
        
        If Vencimiento <= DateAdd("D", DiasGracia, Fecha) Then GoTo diasEnajenacion
    
    Next i
    
diasEnajenacion:
    Exit Function
    
Error:
    Maneja_Error Err
End Function

Public Sub Conectar_WS()
    'Hago la Conexión al Web Service
    Set WSoap = New SoapClient30
    WSoap.MSSoapInit WServidor & "?wsdl"
End Sub

Public Sub Cargar_WebService()
    'Conexión al Web Service
    'WServidor = "http://" & Regresa_Valor("WEBSERVICE", "WebService", "localhost:1285") '& "/MySigmaWS/MySigmaMonitor.asmx?wsdl"
    WServidor = Regresa_Valor("WEBSERVICE", "WebService", "localhost:1285") '& "/MySigmaWS/MySigmaMonitor.asmx?wsdl"
    'WServidor = "http://" & Regresa_Valor("WEBSERVICE", "WebService", "localhost:1285") & "/RemoteQuery.asmx?wsdl"
    WRutaServidor = Regresa_Valor("WEBSERVICE", "Servidor", "localhost")
    WBaseDatos = Regresa_Valor("WEBSERVICE", "BaseDeDatos", "BaseDatos")
    WPuerto = Regresa_Valor("WEBSERVICE", "Puerto", "3307")
End Sub

'Regreso el Valor de un Registro dentro del Arreglo
Public Function getValueArray(Arreglo() As String, Parametro As String, Optional Numero As Boolean = False, Optional Index As Integer = 0) As String
On Error GoTo Error
Dim i As Integer

    For i = 0 To UBound(Arreglo) - 1
        If Parametro = Mid(Arreglo(i), 1, InStr(1, Arreglo(i), "=") - 1) Then
            getValueArray = Mid(Arreglo(i), InStr(1, Arreglo(i), "=") + 1)
            Exit For
        End If
    Next i
    
    If Numero And Trim(getValueArray) = "" Then getValueArray = "0"
    If Numero And getValueArray = "True" Then getValueArray = "1"
    If Numero And getValueArray = "False" Then getValueArray = "0"

    Exit Function
Error:
    Maneja_Error Err
End Function

Private Sub Forzar_Actualizacion()
   On Error Resume Next
   
   Dim rc As New ADODB.Recordset
   
   rc.Open "SHOW Tables;", dbDatos, adOpenStatic, adLockOptimistic
   
    
   While Not rc.EOF
      'DoEvents
      
      If UCase(rc.Fields(0)) <> "ACTUALIZACIONES" Then
         On Error Resume Next
        
'         frmSplash.lblMensaje.Caption = "Forzando Actualizacion: " & rc.Fields(0)
         'Debug.Print rc.Fields(0)
         dbDatos.Execute "Update " & rc.Fields(0) & " SET ID=ID WHERE IDTabla=0;"
      
         Sleep 10
         
      End If
      
      rc.MoveNext
   Wend
   
   Err.Clear
   rc.Close
   
   'Exit Sub
   
Error:
   Maneja_Error Err
   Set rc = Nothing
End Sub

'Public Sub Pase_Automatico_Almoneda()
'Dim Indice As Long, Movimiento As Long, Folio As Long, Prestamo As Double, Cantidad As Integer, Serie As Integer, IDEntrada As Long, strDestino As String, strDescripcion As String, Hora As String, strCodigo As String, Kilates As Integer, crPrecio As Double, AvaluoDiam As Double, GTOSVenta As Double, DiasEnaje As Integer, crPrestamoGlobal As Double, Bloqueado As Integer, FechaBloqueo As String
'Dim rcRemate As New ADODB.Recordset
'Dim rcTmp As New ADODB.Recordset
'Dim rcAux As New ADODB.Recordset
'Dim Salida, Entrada  As Long
'Dim Inventario, sacaO, sacaM As Boolean
'Dim rcPresacaO As New ADODB.Recordset
'Dim rcPresacaM As New ADODB.Recordset
'Dim rcExistencia As New ADODB.Recordset
'Dim rcAuxiliar As New ADODB.Recordset
'Dim IDSacaO, IDSacaM As Long
'Dim auxiliar As Double
'Dim ClaveSucursal As Integer
'On Error GoTo Error
'Dim SacaOro As Double, SacaMiscelaneos As Double, salidaOro As Double, salidaMiscelaneos As Double
'Dim fechaAlmoneda As Date
'
'
'
'IDSacaO = 0
'IDSacaM = 0
'
'
'    fechaAlmoneda = Regresa_Valor_BD("FechaAlmoneda")
'    DiasEnaje = Regresa_Valor_BD("DiasEnajenacion")
'    ClaveSucursal = SacaValor("Sucursales", "Clave", " WHERE Activa=1")
''    rcRemate.Open "SELECT COUNT(e.ID) AS Total FROM empeno e WHERE e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0," & DiasEnaje & ") DAY),'%Y%/%m%/%d')<'" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
'    rcRemate.Open "SELECT COUNT(e.ID) AS Total FROM empeno e WHERE e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0  AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0," & DiasEnaje & ") DAY),'%Y%/%m%/%d')<'" & Format(fechaAlmoneda, "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
'        If rcRemate!Total > 0 Then frmMDI.Bar.Value = 0: frmMDI.Bar.Min = 0: frmMDI.Bar.Max = rcRemate!Total Else Exit Sub
'    rcRemate.Close
'
'    If frmMDI.Bar.Max > 0 Then
'
'        Screen.MousePointer = vbHourglass
'        frmMDI.Bar.Visible = True
'
'        'Checo los contratos que estan vencidos
''        rcRemate.Open "SELECT DISTINCT e.ID,e.NumContrato,e.Folio,e.Fecha,e.Prestamo,e.Avaluo,e.Origen,e.Vencimiento,e.TipoInteres,e.TipoTasa,e.Serie,c.Iniciales,CONCAT(c.Apellido,' ',c.Nombre) AS Cliente " _
''                    & "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0," & DiasEnaje & ") DAY),'%Y%/%m%/%d')<'" & Format(Date, "YYYY/MM/DD") & "' ORDER BY e.NumContrato", dbDatos, adOpenForwardOnly, adLockReadOnly
'                    rcRemate.Open "SELECT DISTINCT e.ID,e.NumContrato,e.Folio,e.Fecha,e.Prestamo,e.Avaluo,e.Origen,e.Vencimiento,e.TipoInteres,e.TipoTasa,e.Serie,c.Iniciales,CONCAT(c.Apellido,' ',c.Nombre) AS Cliente " _
'                    & "FROM empeno e LEFT JOIN clientes c ON e.IDCliente=c.ID WHERE e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0  AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0," & DiasEnaje & ") DAY),'%Y%/%m%/%d')<'" & Format(fechaAlmoneda, "YYYY/MM/DD") & "' ORDER BY e.NumContrato", dbDatos, adOpenForwardOnly, adLockReadOnly
'
'        'Saco el Folio
'        Folio = Regresa_Movimiento(False, "FolioInventario")
'        Regresa_Movimiento True, "FolioInventario"
'
'        'Saco el movimiento
'        Movimiento = Regresa_Movimiento(False)
'        Regresa_Movimiento True
'
'        'Tabla Entrada Inventario
'        dbDatos.Execute "INSERT INTO entradainventario(Fecha,Folio,TipoEntrada,IDUsuario,IDSucursal) VALUES " _
'                        & "('" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & ENTRADAALMONEDA & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'        'Saco el ID de la Entrada
'        IDEntrada = SacaValor("entradainventario", "MAX(ID)")
'
'        'saco los totales de prestamo y peso para la presaca
'        rcPresacaO.Open "SELECT sum(de.prestamo) as total ,sum(basedatos.GetPesoNeto(e.ID)) as totalPeso FROM empeno e inner join detallesempeno de on de.IDempeno=e.ID WHERE (e.IDTipoPrenda=1 or e.IDTipoPrenda=3) AND e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND e.sucursalOrigen=" & ClaveSucursal & " AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(TipoTasa='DIARIA',0," & DiasEnaje & ") DAY),'%Y%/%m%/%d')<'" & Format(Date, "YYYY/MM/DD") & "' ORDER BY e.NumContrato", dbDatos, adOpenForwardOnly, adLockReadOnly
'        If Not rcPresacaO.EOF Or Not rcPresacaO.BOF Then
'        If IsNull(rcPresacaO!Total) = False And IsNull(rcPresacaO!TotalPeso) = False Then
'        dbDatos.Execute "INSERT INTO presaca(Folio,Fecha,IDtipoSaca,IDTablaTipoSaca,TipoSaca,TotalSACA,GramosSaca,Status,IDUsuario)VALUES(0,'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',1,1,'ORO'," & rcPresacaO!Total & "," & rcPresacaO!TotalPeso & ",5, " & frmMDI.IDUsuario & ")"
'
'        SacaOro = rcPresacaO!Total
'        sacaO = True
'        IDSacaO = SacaValor("presaca", "MAX(ID)")
'        End If
'        End If
'
'        rcPresacaM.Open "SELECT sum(de.prestamo) as total FROM empeno e inner join detallesempeno de on de.IDempeno=e.ID WHERE e.IDTipoPrenda=2 AND e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND e.sucursalOrigen=" & ClaveSucursal & " AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0," & DiasEnaje & ") DAY),'%Y%/%m%/%d')<'" & Format(Date, "YYYY/MM/DD") & "' ORDER BY e.NumContrato", dbDatos, adOpenForwardOnly, adLockReadOnly
'        If Not rcPresacaM.EOF And Not rcPresacaM.BOF Then
'
'        If IsNull(rcPresacaM!Total) = False Then
'         dbDatos.Execute "INSERT INTO presaca(Folio,Fecha,IDtipoSaca,IDTablaTipoSaca,TipoSaca,TotalSACA,GramosSaca,Status,IDUsuario)VALUES(0,'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',2,2,'Miscelaneos'," & rcPresacaM!Total & ",0,4, " & frmMDI.IDUsuario & ")"
'         SacaMiscelaneos = rcPresacaM!Total
'         sacaM = True
'         IDSacaM = SacaValor("presaca", "MAX(ID)")
'         End If
'        End If
'
'
''        rcExistencia.Open "SELECT sum(de.prestamo) as total ,sum(de.peso)as totalPeso,de.kilates FROM empeno e INNER JOIN detallesempeno de on e.ID=de.IDEmpeno WHERE (de.tipo=1 or de.tipo=3) AND e.Destino=0 AND e.Pagado=0 AND e.Cancelado=0 AND e.sucursalOrigen=" & SacaValor("Sucursales", "Clave", " WHERE Activa=1") & " AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0," & DiasEnaje & ") DAY),'%Y%/%m%/%d')<'" & Format(Date, "YYYY/MM/DD") & "' group by de.kilates ORDER BY e.NumContrato", dbDatos, adOpenForwardOnly, adLockReadOnly
''        While Not rcExistencia.BOF And Not rcExistencia.EOF
''
''
''
''        Wend
'        'Tomo la Hora
'        Hora = Time
'
'        'Tomo los Gastos de Venta
'        GTOSVenta = Regresa_Valor_BD("GtosVenta") / 100
'
'        dbReportes.Execute "DELETE FROM articulos"
'
'        While Not rcRemate.EOF
'
'            'Marco la prenda como pasada a Remate
''            If rcRemate!Serie = SERIE_B Then
''
''                rcTmp.Open "SELECT de.IDEmpeno,de.MarcayModelo,de.Placas,de.Año,de.Color,de.SerieChasis,de.NumMotor,de.NumTarjetaCircu FROM detallesempenoautos de WHERE de.IDEmpeno=" & rcRemate!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
''
''                    strDescripcion = "MARCA Y MODELO: " & rcTmp!MarcayModelo & ", PLACAS: " & rcTmp!Placas & ", AÑO: " & rcTmp!Año & ", COLOR: " & rcTmp!Color & ", SERIE CHASIS: " & rcTmp!SerieChasis & ", NUM. MOTOR: " & rcTmp!NumMotor & ", TARJETA CIRC.: " & rcTmp!NumTarjetaCircu
''
''                rcTmp.Close
''
''                dbReportes.Execute "INSERT INTO articulos (IDEmpeno,Articulo,Avaluo,Prestamo) VALUES (" & _
''                                            rcRemate!ID & ",'" & strDescripcion & "'," & rcRemate!Avaluo & "," & rcRemate!Prestamo & ")"
''
''                'Marco el contrato como pasado a Almoneda
''                dbDatos.Execute "UPDATE empeno SET Pagado=1,Destino=" & D_ALMONEDA & ",FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & rcRemate!ID
''
''                crPrecio = rcRemate!Avaluo * (1 + Regresa_Valor_BD("PrecioAutos") / 100)
''
''                'Creo el Codigo
''                strCodigo = CreaCodigoBarras(Trim(Format(frmMDI.IDSucursal, "000")), ENTRADAALMONEDA, Trim(rcRemate!NumContrato), 1)
''
''                'Paso el Automovil al Inventario
''                dbDatos.Execute "INSERT INTO detallesentradainventario(IDEntrada,Codigo,Tipo,Cantidad,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,IDEmpeno,SucursalOrigen,TipoEntrada,PrecioVitrina) VALUES (" & _
''                                IDEntrada & ",'" & strCodigo & "',42,1,'" & strDescripcion & "',0,0," & rcRemate!Avaluo & "," & rcRemate!Prestamo & ",'','','','','','',0,''," & rcRemate!ID & "," & frmMDI.IDSucursal & "," & ENTRADAALMONEDA & "," & ConvMoneda(crPrecio) & ")"
''
''
''                'Grabamos el cargo
''                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
''                                & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE01','620301'," & ConvMoneda(Prestamo) & "," & TIPO_CARGO & "," & rcRemate!Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
''
''                'Grabamos el abono
''                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
''                                & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE50','201750'," & ConvMoneda(Prestamo) & "," & TIPO_ABONO & "," & rcRemate!Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
''
''            Else
'
''                'Marco el contrato como marcado a Remate
''                dbDatos.Execute "UPDATE empeno SET Pagado=1,Destino=" & D_ALMONEDA & ",FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & rcRemate!ID
'
'                'Tomo el Prestamo Original y checo si no se ha abonado a capital
'                crPrestamoGlobal = SacaValor("detallesempeno", "SUM(Prestamo)", " WHERE IDEmpeno=" & rcRemate!ID)
'
''''''''         rcTmp.Open "SELECT de.ID,de.IDEmpeno,de.Codigo,de.Cantidad,de.Articulo,(de.Peso-de.PesoPiedras) AS PesoTotal,de.Kilates,de.Estado,de.Avaluo,de.Prestamo,de.Tipo,de.Observaciones,es.ID AS IDEstado FROM detallesempeno de LEFT JOIN estado es ON (de.Estado=es.Estado) WHERE de.IDEmpeno=" & rcRemate!ID, dbDatos, adOpenForwardOnly, adLockReadOnly
'
'                rcTmp.Open "SELECT de.ID,de.IDEmpeno,de.Codigo,de.Cantidad,de.Articulo,(de.Peso-de.PesoPiedras) AS PesoTotal,de.Kilates,k.descripcion,de.Estado,de.Avaluo,de.Prestamo,de.Tipo,de.Observaciones,es.ID AS IDEstado " & _
'                           "FROM detallesempeno de LEFT JOIN estado es ON (de.Estado=es.Estado) LEFT JOIN kilatajes k on de.kilates=k.ID " & _
'                           "WHERE de.IDEmpeno=" & rcRemate!ID & _
'                           IIf(Regresa_Tipo_Empeno(rcRemate!ID) = 1 Or Regresa_Tipo_Empeno(rcRemate!ID) = 43, " AND de.Tipo=es.IDTipo", ""), dbDatos, adOpenForwardOnly, adLockReadOnly
'
'                With rcTmp
'
'
'
'                    While Not rcTmp.EOF
'
'                        'Destino
''                        strDestino = "VENTA"
'
'                        If !Tipo = 1 Or !Tipo = 3 Then
''                            If !Estado = "E" Or !Estado = "B" Or !Estado = "A" Then
''                                strDestino = "VENTA"
''                            Else
'                                strDestino = "Fundicion"
'                                Salida = 201750
'                                Entrada = 620309
'                                Inventario = False
''                            End If
'
''                                fundicionOro = 620309
'                            Else
'
'                            strDestino = "VENTA"
'                            Salida = 201752
'                            Entrada = 620301
'                            Inventario = True
'                        End If
'
'                        Dim crPrestamoPrenda As Double
'
'                        crPrestamoPrenda = Redondeo((((!Prestamo * 100) / crPrestamoGlobal) / 100) * rcRemate!Prestamo)
'
'                        dbReportes.Execute "INSERT INTO articulos (IDEmpeno,Articulo,Peso,Kilates,Avaluo,Prestamo,Cantidad,Destino,Observaciones,DescripcionKilates) VALUES (" & _
'                                            !IDEmpeno & ",'" & !Articulo & "'," & ConvMoneda(!PesoTotal) & "," & Val(!Kilates) & "," & ConvMoneda(!Avaluo) & "," & ConvMoneda(!Prestamo) & "," & Val(!Cantidad) & ",'" & strDestino & "','" & !Observaciones & "','" & !Descripcion & "')"
'
'                        'Marco el contrato como marcado a Remate
'                        If !Tipo = 1 Or !Tipo = 3 Then
'                dbDatos.Execute "UPDATE empeno SET Pagado=1,IDSaca=" & IDSacaO & " ,Destino=" & D_ALMONEDA & ",FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & rcRemate!ID
'                Else
'                 dbDatos.Execute "UPDATE empeno SET Pagado=1,IDSaca=" & IDSacaM & " ,Destino=" & D_ALMONEDA & ",FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & rcRemate!ID
'
'
'
'      End If
'
'       dbDatos.Execute "UPDATE detallesempeno SET Almoneda=1,Destino=" & IIf(strDestino = "VENTA", D_VENTA, D_FUNDICION) & "  WHERE ID=" & !ID
'
'                        'Paso las prendas al Inventario
'                        With rcAux
'
'                            .Open "SELECT e.ID,e.Fecha,e.Folio,e.Vencimiento,e.TipoInteres,d.ID AS IDPrenda,d.Articulo,d.Kilates,d.Peso,d.Avaluo,d.Cantidad,d.Prestamo,d.Avaluo,d.Tipo,d.IDEmpeno,d.Marca,d.Modelo,d.Serie,d.Color,d.Tamano,d.Codigo,d.TipoPrenda,d.Observaciones,d.Estado,d.CantidadPiedras,d.PesoPiedras,d.CantidadDiamantes,d.Puntos,d.PrestamoDiamante,k.Descripcion AS Kilataje FROM detallesempeno d LEFT JOIN kilatajes k ON d.Kilates=k.Clave INNER JOIN empeno e ON d.IDEmpeno=e.ID WHERE d.ID=" & rcTmp!ID & " ORDER BY d.Codigo", dbDatos, adOpenForwardOnly, adLockReadOnly
'
'
'
'                            'Saco el Precio
'                            crPrecio = Redondeo(!Avaluo * (1 + GTOSVenta))
'
'                            'Tabla de DetalleEntradaInventario
'
'                            If Inventario Then
'                            dbDatos.Execute "INSERT INTO detallesentradainventario(IDEntrada,Codigo,Tipo,Cantidad,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,IDEmpeno,SucursalOrigen,TipoEntrada,PrecioVitrina,CantidadPiedras,PesoPiedras,CantidadDiamantes,Puntos,PrestamoDiamante,fechaSaca) VALUES (" & _
'                                            IDEntrada & ",'" & !Codigo & "'," & !Tipo & "," & !Cantidad & ",'" & !Articulo & "'," & ConvMoneda(!Peso) & "," & !Kilates & "," & ConvMoneda(!Avaluo) & "," & ConvMoneda(rcAux!Prestamo) & ",'" & !Estado & "','" & !Marca & "','" & !Modelo & "','" & !Serie & "','" & !Color & "','" & !Tamano & "'," & !TipoPrenda & ",'" & !Observaciones & "'," & !IDEmpeno & "," & frmMDI.IDSucursal & "," & IIf(strDestino = "VENTA", D_VENTA, D_FUNDICION) & "," & ConvMoneda(crPrecio) & "," & !CantidadPiedras & "," & ConvMoneda(!PesoPiedras) & "," & !CantidadDiamantes & "," & ConvMoneda(!Puntos) & "," & ConvMoneda(!PrestamoDiamante) & ",'" & Format(Date, "YYYY/MM/DD") & "')"
''
'                            End If
'                            .Close
'                            Set rcAux = Nothing
'
'                        End With
'
'                        Inventario = False
'
'                       'Muevo las Cuentas Contables
''                        If strDestino = "VENTA" Then
''
''                            'Grabamos el cargo
''                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
''                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE01'," & Entrada & "," & ConvMoneda(crPrestamoPrenda) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
''
''                            'Grabamos la cuenta de saca
''
''                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
''                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE01',610501," & ConvMoneda(crPrestamoPrenda) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
''
''                            'Grabamos el abono
''                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
''                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE50'," & Salida & "," & ConvMoneda(crPrestamoPrenda) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
''
''                        Else
''
''                            'Grabamos el cargo
''                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
''                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','SACA'," & Movimiento & "," & rcRemate!NumContrato & ",'SA01'," & Entrada & "," & ConvMoneda(crPrestamoPrenda) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
''
''                                            'Grabamos el car a la cuenta de saca
''                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
''                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','SACA'," & Movimiento & "," & rcRemate!NumContrato & ",'SA01',610502," & ConvMoneda(crPrestamoPrenda) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
''
''                            'Grabamos el abono
''                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
''                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Almoneda'," & Movimiento & "," & rcRemate!NumContrato & ",'RE50'," & Salida & "," & ConvMoneda(crPrestamoPrenda) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
''
''                        End If
'
'                    rcTmp.MoveNext
'                    Wend
'
'
'
'
'                End With
'                rcTmp.Close
'
''            End If
'
'        frmMDI.Bar.Value = frmMDI.Bar.Value + 1
'        rcRemate.MoveNext
'        Wend
'
'
'
'         If SacaMiscelaneos > 0 Then
'
'                            'Grabamos el cargo
'                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','SACA',0,0,'SA01',620301," & ConvMoneda(SacaMiscelaneos) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'                            'Grabamos la cuenta de saca
'
'                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','SACA',0,0,'SA01',610501," & ConvMoneda(SacaMiscelaneos) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'                            'Grabamos el abono
'                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','SACA',0,0,'SA01',201752," & ConvMoneda(SacaMiscelaneos) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'                        End If
'
'                        If SacaOro > 0 Then
'                            'Grabamos el cargo
'                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','SACA',0,0,'SA01',620309," & ConvMoneda(SacaOro) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'                                            'Grabamos el car a la cuenta de saca
'                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','SACA',0,0,'SA01',610502," & ConvMoneda(SacaOro) & "," & TIPO_CARGO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'                            'Grabamos el abono
'                            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " _
'                                            & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','SACA',0,0,'SA01',201750," & ConvMoneda(SacaOro) & "," & TIPO_ABONO & "," & Serie & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'                        End If
''        rcExistencia.Open "SELECT sum(de.prestamo) as total ,sum(de.peso)as totalPeso,de.kilates,k.descripcion FROM empeno e INNER JOIN detallesempeno de on e.ID=de.IDEmpeno INNER JOIN kilatajes k on de.kilates=k.ID WHERE e.IDSaca=" & IDSacaO & " AND e.sucursalOrigen=" & ClaveSucursal & "  group by de.kilates ORDER BY e.NumContrato;", dbDatos, adOpenForwardOnly, adLockReadOnly
''        While Not rcExistencia.BOF And Not rcExistencia.EOF
''
''       rcAuxiliar.Open "select " & rcExistencia!Descripcion & " as costo from parametros", dbDatos, adOpenForwardOnly, adLockOptimistic
''        If Not rcAuxiliar.BOF And Not rcAuxiliar.EOF Then
''
''        dbDatos.Execute "INSERT INTO Corporativo.existenciaOro(IDKilataje,CostoGramo,Q,Sucursal,Actualizar) VALUES(" & rcExistencia!Kilates & "," & rcAuxiliar!Costo & "," & rcExistencia!TotalPeso & "," & ClaveSucursal & ",0)"
''
''        End If
''
''
''         rcAuxiliar.Close
''        rcExistencia.MoveNext
''
''        Wend
'
''
''        rcExistencia.Close
'        rcRemate.Close
'        Set rcRemate = Nothing
'
'        frmMDI.Bar.Visible = False
'
'        Sleep 1000
'        With frmMDI.Cr
'            .Reset
'            .DiscardSavedData = True
'            .WindowShowPrintSetupBtn = True
'            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'            .ReportFileName = Path & "\Reportes\ContratosAlmoneda2.rpt"
''            FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1
''            .SelectionFormula = "{empeno.FechaAlmoneda}='" & Format(Date, "YYYY-MM-DD") & "' AND {empeno.Almoneda}=1"
'             .SelectionFormula = "{empeno.IDSaca}=" & IDSacaO
'
'            .Formulas(0) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
'            .Formulas(1) = "Encabezado='" & Sucursal.RazonSocial & "'"
'            .Formulas(2) = "Leyenda=''"
'
'            .SubreportToChange = "Resumen"
'            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'            .SelectionFormula = "{articulos.Kilates} >0 AND {articulos.Destino}= 'VENTA'"
''            .SelectionFormula = "{empeno.FechaMovimiento} >= '" & Format(Date, "YYYY-MM-DD") & "' AND {empeno.FechaMovimiento} <= '" & Format(Date, "YYYY-MM-DD") & "' AND {Articulos.Kilates}<>'' AND {Articulos.Destino}= VENTA"
'            .SubreportToChange = "ResumenFundicion"
'            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'            .SelectionFormula = "{articulos.Kilates} >0 AND {articulos.Destino}='Fundicion'"
''            .SelectionFormula = "{empeno.FechaMovimiento} >= '" & Format(Date, "YYYY-MM-DD") & "' AND {empeno.FechaMovimiento} <= '" & Format(Date, "YYYY-MM-DD") & "' AND {detallesempeno.Kilates}<>'' AND {detallesempeno.Destino}=" & D_FUNDICION
'
'            .WindowTitle = "Contratos Almoneda"
'            .Destination = crptToWindow
'            .WindowState = crptMaximized
'            .Action = 1
'        End With
'
' Sleep 1000
'        With frmMDI.Cr
'            .Reset
'            .DiscardSavedData = True
'            .WindowShowPrintSetupBtn = True
'            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'            .ReportFileName = Path & "\Reportes\ContratosAlmoneda.rpt"
''            FechaAlmoneda='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Almoneda=1
''            .SelectionFormula = "{empeno.FechaAlmoneda}='" & Format(Date, "YYYY-MM-DD") & "' AND {empeno.Almoneda}=1"
''             .SelectionFormula = "{empeno.IDSaca}=" & IDSacaO
'.SelectionFormula = "{vwrematediario.IDEntrada}=" & IDEntrada
'            .Formulas(0) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
'            .Formulas(1) = "Encabezado='" & Sucursal.RazonSocial & "'"
'            .Formulas(2) = "Leyenda=''"
'
''            .SubreportToChange = "Resumen"
''            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
''            .SelectionFormula = "{articulos.Kilates} >0 AND {articulos.Destino}= 'VENTA'"
'''            .SelectionFormula = "{empeno.FechaMovimiento} >= '" & Format(Date, "YYYY-MM-DD") & "' AND {empeno.FechaMovimiento} <= '" & Format(Date, "YYYY-MM-DD") & "' AND {Articulos.Kilates}<>'' AND {Articulos.Destino}= VENTA"
''            .SubreportToChange = "ResumenFundicion"
''            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
''            .SelectionFormula = "{articulos.Kilates} >0 AND {articulos.Destino}='Fundicion'"
''            .SelectionFormula = "{empeno.FechaMovimiento} >= '" & Format(Date, "YYYY-MM-DD") & "' AND {empeno.FechaMovimiento} <= '" & Format(Date, "YYYY-MM-DD") & "' AND {detallesempeno.Kilates}<>'' AND {detallesempeno.Destino}=" & D_FUNDICION
'
'            .WindowTitle = "Contratos Almoneda"
'            .Destination = crptToWindow
'            .WindowState = crptMaximized
'            .Action = 1
'        End With
'
'
'
'        Screen.MousePointer = vbDefault
'    End If
'    Set rcTmp = Nothing
'    Exit Sub
'
'Error:
'    Maneja_Error Err
'    Set rcRemate = Nothing
'    Set rcTmp = Nothing
'    Screen.MousePointer = vbDefault
'End Sub


