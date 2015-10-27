-- MySQL Administrator dump 1.4
--
-- ------------------------------------------------------
-- Server version	5.1.53-community


/*!40101 SET @OLD_CHARACTER_SET_CLIENT=@@CHARACTER_SET_CLIENT */;
/*!40101 SET @OLD_CHARACTER_SET_RESULTS=@@CHARACTER_SET_RESULTS */;
/*!40101 SET @OLD_COLLATION_CONNECTION=@@COLLATION_CONNECTION */;
/*!40101 SET NAMES utf8 */;

/*!40014 SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0 */;
/*!40014 SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0 */;
/*!40101 SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='NO_AUTO_VALUE_ON_ZERO' */;


--
-- Create schema basedatos
--

CREATE DATABASE IF NOT EXISTS basedatos;
USE basedatos;

--
-- Temporary table structure for view `vwapartadosrematados`
--
DROP TABLE IF EXISTS `vwapartadosrematados`;
DROP VIEW IF EXISTS `vwapartadosrematados`;
CREATE TABLE `vwapartadosrematados` (
  `Abonos` double(22,5),
  `ID` int(10),
  `Fecha` datetime,
  `FechaMovimiento` datetime,
  `Folio` int(10),
  `IVA` double(15,5),
  `Vencimiento` date,
  `Total` double(15,5),
  `Descuento` double(15,5),
  `Pagado` tinyint(1),
  `Cancelado` tinyint(1),
  `OrigenCancelacion` int(2),
  `Cliente` varchar(171)
);

--
-- Temporary table structure for view `vwdetallesempeno`
--
DROP TABLE IF EXISTS `vwdetallesempeno`;
DROP VIEW IF EXISTS `vwdetallesempeno`;
CREATE TABLE `vwdetallesempeno` (
  `ID` int(10),
  `IDEmpeno` int(10),
  `Cantidad` int(10),
  `Tipo` int(10),
  `Articulo` varchar(400),
  `PesoTotal` double(15,3),
  `PesoPiedras` double(15,5),
  `PesoReal` double(22,5),
  `Prestamo` double(15,5),
  `Avaluo` double(15,5),
  `Observaciones` varchar(250),
  `Estado` varchar(50),
  `Marca` varchar(100),
  `Modelo` varchar(100),
  `Serie` varchar(50),
  `Tamano` varchar(50),
  `Color` varchar(50),
  `Tipo_DESC` varchar(50),
  `Kil_DESC` varchar(50)
);

--
-- Temporary table structure for view `vwfacturadiaria`
--
DROP TABLE IF EXISTS `vwfacturadiaria`;
DROP VIEW IF EXISTS `vwfacturadiaria`;
CREATE TABLE `vwfacturadiaria` (
  `NumRegistros` bigint(21),
  `Fecha` date,
  `ImporteTotal` double(21,4)
);

--
-- Temporary table structure for view `vwfacturaventas`
--
DROP TABLE IF EXISTS `vwfacturaventas`;
DROP VIEW IF EXISTS `vwfacturaventas`;
CREATE TABLE `vwfacturaventas` (
  `NumRegistros` bigint(21),
  `Fecha` date,
  `ImporteTotal` double(21,4)
);

--
-- Temporary table structure for view `vwpagosfijos`
--
DROP TABLE IF EXISTS `vwpagosfijos`;
DROP VIEW IF EXISTS `vwpagosfijos`;
CREATE TABLE `vwpagosfijos` (
  `IDEmpeno` int(11),
  `FechaMovimiento` datetime
);

--
-- Temporary table structure for view `vwrepapartados`
--
DROP TABLE IF EXISTS `vwrepapartados`;
DROP VIEW IF EXISTS `vwrepapartados`;
CREATE TABLE `vwrepapartados` (
  `Abonos` double(22,5),
  `ID` int(10),
  `Fecha` datetime,
  `Folio` int(10),
  `IVA` double(15,5),
  `Vencimiento` date,
  `Total` double(15,5),
  `Descuento` double(15,5),
  `Pagado` tinyint(1),
  `Cancelado` tinyint(1),
  `OrigenCancelacion` int(2),
  `Cliente` varchar(171)
);

--
-- Definition of table `abonos`
--

DROP TABLE IF EXISTS `abonos`;
CREATE TABLE `abonos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDVenta` int(10) DEFAULT '0',
  `Fecha` datetime DEFAULT NULL,
  `Movimiento` int(10) DEFAULT '0',
  `Importe` double(15,5) DEFAULT '0.00000',
  `Cancelado` tinyint(1) DEFAULT '0',
  `FechaMovimiento` datetime DEFAULT NULL,
  `PC` varchar(25) DEFAULT NULL,
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(5) DEFAULT '0',
  `DescuentoXPuntos` double(15,5) DEFAULT '0.00000',
  `SaldoPuntosAnterior` double(15,5) DEFAULT '0.00000',
  `PuntosUsados` double(15,5) DEFAULT '0.00000',
  `PuntosAcumulados` double(15,5) DEFAULT '0.00000',
  `SaldoPuntosActual` double(15,5) DEFAULT '0.00000',
  `IDTarjeta` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDVenta` (`IDVenta`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `abonos`
--

/*!40000 ALTER TABLE `abonos` DISABLE KEYS */;
/*!40000 ALTER TABLE `abonos` ENABLE KEYS */;


--
-- Definition of table `asignaciontarjetas`
--

DROP TABLE IF EXISTS `asignaciontarjetas`;
CREATE TABLE `asignaciontarjetas` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Fecha` datetime NOT NULL DEFAULT '0000-00-00 00:00:00',
  `NumeroTarjeta` varchar(60) DEFAULT NULL,
  `IDTarjeta` int(10) unsigned NOT NULL DEFAULT '0',
  `IDCliente` int(10) unsigned NOT NULL DEFAULT '0',
  `IDUsuario` int(10) unsigned NOT NULL DEFAULT '0',
  `PC` varchar(60) DEFAULT NULL,
  `Puntos` int(10) unsigned NOT NULL DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `asignaciontarjetas`
--

/*!40000 ALTER TABLE `asignaciontarjetas` DISABLE KEYS */;
/*!40000 ALTER TABLE `asignaciontarjetas` ENABLE KEYS */;


--
-- Definition of table `autorizaciones`
--

DROP TABLE IF EXISTS `autorizaciones`;
CREATE TABLE `autorizaciones` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` datetime DEFAULT NULL,
  `IDUsuario` int(10) DEFAULT '0',
  `Opcion` int(10) DEFAULT '0',
  `Status` int(1) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  `Codigo` varchar(15) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `autorizaciones`
--

/*!40000 ALTER TABLE `autorizaciones` DISABLE KEYS */;
/*!40000 ALTER TABLE `autorizaciones` ENABLE KEYS */;


--
-- Definition of table `auxiliar`
--

DROP TABLE IF EXISTS `auxiliar`;
CREATE TABLE `auxiliar` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` date DEFAULT NULL,
  `Hora` time DEFAULT NULL,
  `Movimiento` int(10) DEFAULT '0',
  `Concepto` varchar(200) DEFAULT NULL,
  `Folio` int(10) DEFAULT '0',
  `Iniciales` varchar(10) DEFAULT NULL,
  `Cuenta` varchar(15) DEFAULT NULL,
  `Importe` double(19,4) DEFAULT '0.0000',
  `Tipo` int(10) DEFAULT '0',
  `Serie` int(10) DEFAULT '0',
  `PC` varchar(25) DEFAULT NULL,
  `Corte` smallint(5) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  `IDDivisa` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `Folio` (`Folio`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `auxiliar`
--

/*!40000 ALTER TABLE `auxiliar` DISABLE KEYS */;
/*!40000 ALTER TABLE `auxiliar` ENABLE KEYS */;


--
-- Definition of table `bancos`
--

DROP TABLE IF EXISTS `bancos`;
CREATE TABLE `bancos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` datetime DEFAULT NULL,
  `Folio` int(10) DEFAULT '0',
  `Cancelado` tinyint(1) DEFAULT '0',
  `FechaMovimiento` datetime DEFAULT NULL,
  `Deposito` int(1) NOT NULL DEFAULT '0',
  `Concepto` varchar(80) NOT NULL,
  `Importe` double(15,5) DEFAULT '0.00000',
  `TipoMov` tinyint(1) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `Folio` (`Folio`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `bancos`
--

/*!40000 ALTER TABLE `bancos` DISABLE KEYS */;
/*!40000 ALTER TABLE `bancos` ENABLE KEYS */;


--
-- Definition of table `boveda`
--

DROP TABLE IF EXISTS `boveda`;
CREATE TABLE `boveda` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` datetime DEFAULT NULL,
  `Folio` int(10) DEFAULT '0',
  `Cancelado` tinyint(1) DEFAULT '0',
  `FechaMovimiento` datetime DEFAULT NULL,
  `Deposito` int(1) DEFAULT '0',
  `Concepto` varchar(80) DEFAULT NULL,
  `Importe` double(15,5) DEFAULT '0.00000',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(5) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `Folio` (`Folio`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `boveda`
--

/*!40000 ALTER TABLE `boveda` DISABLE KEYS */;
/*!40000 ALTER TABLE `boveda` ENABLE KEYS */;


--
-- Definition of table `cancelaciones`
--

DROP TABLE IF EXISTS `cancelaciones`;
CREATE TABLE `cancelaciones` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Fecha` datetime NOT NULL,
  `TipoMovimiento` int(11) DEFAULT '0',
  `Contrato` int(10) DEFAULT '0',
  `Folio` int(10) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `Descripcion` varchar(200) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cancelaciones`
--

/*!40000 ALTER TABLE `cancelaciones` DISABLE KEYS */;
/*!40000 ALTER TABLE `cancelaciones` ENABLE KEYS */;


--
-- Definition of table `cierrediario`
--

DROP TABLE IF EXISTS `cierrediario`;
CREATE TABLE `cierrediario` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` date DEFAULT NULL,
  `Sucursal` varchar(60) DEFAULT NULL,
  `Cajero` varchar(60) DEFAULT NULL,
  `Saldo` double(15,5) DEFAULT '0.00000',
  `Debe` double(15,5) DEFAULT '0.00000',
  `Haber` double(15,5) DEFAULT '0.00000',
  `Efectivo` double(15,5) DEFAULT '0.00000',
  `Ajuste` double(15,5) DEFAULT '0.00000',
  `IDUsuario` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `Fecha` (`Fecha`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cierrediario`
--

/*!40000 ALTER TABLE `cierrediario` DISABLE KEYS */;
/*!40000 ALTER TABLE `cierrediario` ENABLE KEYS */;


--
-- Definition of table `clientes`
--

DROP TABLE IF EXISTS `clientes`;
CREATE TABLE `clientes` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Iniciales` varchar(4) DEFAULT NULL,
  `Nombre` varchar(50) DEFAULT NULL,
  `Apellido` varchar(120) DEFAULT NULL,
  `ApellidoPaterno` varchar(50) DEFAULT '',
  `ApellidoMaterno` varchar(50) DEFAULT '',
  `RazonSocial` varchar(180) DEFAULT '',
  `PersonaFisica` tinyint(1) DEFAULT '1',
  `FechaAltaRazonSocial` date DEFAULT NULL,
  `Direccion` varchar(70) DEFAULT NULL,
  `NoExterior` varchar(8) DEFAULT '',
  `NoInterior` varchar(8) DEFAULT '',
  `Colonia` varchar(120) DEFAULT NULL,
  `Municipio` varchar(120) DEFAULT NULL,
  `Estado` varchar(60) DEFAULT NULL,
  `Tel` varchar(50) DEFAULT NULL,
  `Celular` varchar(50) DEFAULT NULL,
  `CorreoElectronico` varchar(100) DEFAULT NULL,
  `Identificacion` varchar(60) DEFAULT NULL,
  `NumeroIdentificacion` varchar(30) DEFAULT NULL,
  `IDMedio` int(10) DEFAULT '0',
  `Boletas` int(10) DEFAULT '0',
  `Notas` varchar(150) DEFAULT NULL,
  `CP` varchar(10) DEFAULT NULL,
  `Rfc` varchar(35) DEFAULT NULL,
  `Curp` varchar(30) DEFAULT '',
  `FecNac` date DEFAULT NULL,
  `Sexo` smallint(5) DEFAULT NULL,
  `FecRegistro` date DEFAULT NULL,
  `IDUsuario` int(10) DEFAULT NULL,
  `Caja` varchar(50) DEFAULT NULL,
  `Foto` varchar(50) DEFAULT NULL,
  `NumTarjeta` tinytext,
  `Empresa` varchar(150) DEFAULT NULL,
  `Nacionalidad` int(1) DEFAULT '0',
  `IDNacionalidad` int(10) DEFAULT '0',
  `IDEstadoCivil` int(10) DEFAULT '0',
  `Profesion` varchar(80) DEFAULT NULL,
  `CedulaCotitular` varchar(30) DEFAULT NULL,
  `IdOcupacion` int(11) DEFAULT '0',
  `IdEstadoNac` int(10) DEFAULT '0',
  `IdPaisNacimiento` int(10) DEFAULT '0',
  `IdPaisNacionalidad` int(10) DEFAULT '0',
  `IdTipoIdent` int(10) DEFAULT '0',
  `DescIdentificacionOtro` varchar(200) DEFAULT '',
  `FechaExpIdent` date DEFAULT NULL,
  `Email` varchar(50) DEFAULT '',
  `RL_Nombre` varchar(50) DEFAULT '',
  `RL_ApellidoPaterno` varchar(50) DEFAULT '',
  `RL_ApellidoMaterno` varchar(50) DEFAULT '',
  `RL_Rfc` varchar(30) DEFAULT '',
  `RL_Curp` varchar(30) DEFAULT '',
  `IdTipoAlerta` int(10) DEFAULT '0',
  `DescTipoAlerta` varchar(3000) DEFAULT '',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `Identificacion` (`Identificacion`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `clientes`
--

/*!40000 ALTER TABLE `clientes` DISABLE KEYS */;
/*!40000 ALTER TABLE `clientes` ENABLE KEYS */;


--
-- Definition of table `compras`
--

DROP TABLE IF EXISTS `compras`;
CREATE TABLE `compras` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Cancelado` tinyint(1) DEFAULT '0',
  `Fecha` datetime DEFAULT NULL,
  `Folio` int(10) DEFAULT '0',
  `IDCliente` int(10) DEFAULT '0',
  `Total` double(15,5) DEFAULT '0.00000',
  `Iva` int(10) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(5) DEFAULT '0',
  `FechaMovimiento` datetime DEFAULT NULL,
  PRIMARY KEY (`ID`) USING BTREE,
  KEY `IDCliente` (`IDCliente`),
  KEY `IDUsuario` (`IDUsuario`),
  KEY `ID` (`ID`) USING BTREE
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `compras`
--

/*!40000 ALTER TABLE `compras` DISABLE KEYS */;
/*!40000 ALTER TABLE `compras` ENABLE KEYS */;


--
-- Definition of table `conceptos`
--

DROP TABLE IF EXISTS `conceptos`;
CREATE TABLE `conceptos` (
  `Id` int(10) NOT NULL AUTO_INCREMENT,
  `Descripcion` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`Id`),
  KEY `ID` (`Id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `conceptos`
--

/*!40000 ALTER TABLE `conceptos` DISABLE KEYS */;
/*!40000 ALTER TABLE `conceptos` ENABLE KEYS */;


--
-- Definition of table `configuraciontasas`
--

DROP TABLE IF EXISTS `configuraciontasas`;
CREATE TABLE `configuraciontasas` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDTipoInteres` int(10) DEFAULT '0',
  `IDTipoPeriodo` int(10) DEFAULT '0',
  `IDPlazo` int(10) DEFAULT '0',
  `TasaTipica` double(15,5) DEFAULT '0.00000',
  `TasaPromocion` double(15,5) DEFAULT '0.00000',
  `TasaPreferencial` double(15,5) DEFAULT '0.00000',
  `PorPrestamo` double(15,5) DEFAULT '0.00000',
  `Cat` double(15,5) DEFAULT '0.00000',
  `Almacenaje` double(15,5) DEFAULT '0.00000',
  `Seguro` double(15,5) DEFAULT '0.00000',
  PRIMARY KEY (`ID`),
  KEY `Id` (`ID`),
  KEY `ID1` (`IDTipoInteres`) USING BTREE
) ENGINE=MyISAM AUTO_INCREMENT=10 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `configuraciontasas`
--

/*!40000 ALTER TABLE `configuraciontasas` DISABLE KEYS */;
INSERT INTO `configuraciontasas` (`ID`,`IDTipoInteres`,`IDTipoPeriodo`,`IDPlazo`,`TasaTipica`,`TasaPromocion`,`TasaPreferencial`,`PorPrestamo`,`Cat`,`Almacenaje`,`Seguro`) VALUES 
 (1,1,1,3,7.00000,7.00000,7.00000,83.00000,174.00000,7.50000,0.00000),
 (4,1,1,1,7.00000,7.00000,7.00000,83.00000,174.00000,7.50000,0.00000),
 (7,6,4,2,7.00000,7.00000,7.00000,83.00000,174.00000,7.50000,0.00000),
 (8,9,1,1,8.00000,8.00000,8.00000,83.00000,120.00000,2.00000,0.00000),
 (9,3,1,2,10.00000,10.00000,10.00000,100.00000,120.00000,1.00000,0.00000);
/*!40000 ALTER TABLE `configuraciontasas` ENABLE KEYS */;


--
-- Definition of table `cortediamantes`
--

DROP TABLE IF EXISTS `cortediamantes`;
CREATE TABLE `cortediamantes` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Descripcion` varchar(45) NOT NULL DEFAULT '',
  `Descuento` double(15,5) NOT NULL DEFAULT '0.00000',
  `Ordenamiento` int(2) NOT NULL DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=10 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cortediamantes`
--

/*!40000 ALTER TABLE `cortediamantes` DISABLE KEYS */;
INSERT INTO `cortediamantes` (`ID`,`Descripcion`,`Descuento`,`Ordenamiento`) VALUES 
 (1,'REDONDO',100.00000,1),
 (2,'PRINCESS',70.00000,2),
 (3,'CORAZON',85.00000,5),
 (4,'MARQUISE',70.00000,4),
 (5,'BAGUETTE',70.00000,3),
 (6,'OVALO',70.00000,6),
 (7,'PERA',70.00000,7),
 (8,'RADIANT',70.00000,8),
 (9,'ESMERALDA',70.00000,9);
/*!40000 ALTER TABLE `cortediamantes` ENABLE KEYS */;


--
-- Definition of table `cotizaciones`
--

DROP TABLE IF EXISTS `cotizaciones`;
CREATE TABLE `cotizaciones` (
  `Id` int(10) NOT NULL AUTO_INCREMENT,
  `IDMoneda` int(10) DEFAULT '0',
  `Fecha` date DEFAULT NULL,
  `Hora` time DEFAULT NULL,
  `Compra` double(15,5) DEFAULT '0.00000',
  `Venta` double(15,5) DEFAULT '0.00000',
  `Ponderado` double(15,5) DEFAULT '0.00000',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  PRIMARY KEY (`Id`),
  KEY `id` (`Id`),
  KEY `Idmoneda` (`IDMoneda`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cotizaciones`
--

/*!40000 ALTER TABLE `cotizaciones` DISABLE KEYS */;
/*!40000 ALTER TABLE `cotizaciones` ENABLE KEYS */;


--
-- Definition of table `cuentas`
--

DROP TABLE IF EXISTS `cuentas`;
CREATE TABLE `cuentas` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Mayor` varchar(10) DEFAULT NULL,
  `Concepto` varchar(80) DEFAULT NULL,
  `Cuenta` varchar(10) DEFAULT NULL,
  `Descripcion` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `Cuenta` (`Cuenta`),
  KEY `ID` (`ID`),
  KEY `Mayor` (`Mayor`)
) ENGINE=MyISAM AUTO_INCREMENT=58 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cuentas`
--

/*!40000 ALTER TABLE `cuentas` DISABLE KEYS */;
INSERT INTO `cuentas` (`ID`,`Mayor`,`Concepto`,`Cuenta`,`Descripcion`) VALUES 
 (1,'110100','CAJA','110101','EFECTIVO RECIBIDO EN CAJA'),
 (2,'110100','CAJA','110150','EFECTIVO PAGADO DE CAJA'),
 (3,'201700','EMPEÑO','201701','EMPEÑOS'),
 (4,'201700','EMPEÑO','201750','DESEMPEÑOS'),
 (5,'151200','DEUDORES','151201','FALTANTE DE EFECTIVO'),
 (6,'151200','DEUDORES','151250','PAGO DE FALTANTES'),
 (7,'200900','CENTRAL','200901','CONCENTRACION DE EFECTIVO'),
 (8,'200900','CENTRAL','200950','DOTACION DE EFECTIVO/JOYERIA'),
 (11,'210100','BANCOS','210101','DEPOSITOS RECIBIDOS'),
 (12,'210100','BANCOS','210150','CHEQUES PAGADOS'),
 (13,'511100','GASTOS','511101','GASTOS VARIOS'),
 (15,'520400','INTERESES','520401','INTERESES'),
 (16,'520400','INTERESES','520450','INTERESES'),
 (17,'520600','REMATES','520601','REMATE'),
 (18,'620300','INVENTARIOS','620301','DOTACION DE INVENTARIO'),
 (19,'620300','INVENTARIOS','620350','SALIDA DE INVENTARIO'),
 (22,'620500','APARTADOS','620501','CLIENTES APARTADOS'),
 (23,'620500','APARTADOS','620550','ABONO A APARTADO'),
 (24,'110900','BOVEDA','110901','ENTRADA A BOVEDA'),
 (25,'110900','BOVEDA','110950','RETIRO DE BOVEDA'),
 (28,'620600','DESCUENTO','620601','DESCUENTO VENTAS'),
 (29,'620600','DESCUENTO','620650','DEVOLUCION DESCUENTO'),
 (30,'620700','CAMBIOS','620701','SALIDAS CAMBIOS APARTADOS'),
 (31,'620700','CAMBIOS','620750','ENTRADA CAMBIOS APARTADOS'),
 (32,'110000','OTROS','110001','DEPOSITOS OTROS'),
 (33,'110000','OTROS','110050','CHEQUES OTROS'),
 (34,'520600','VALORES','520650','VALORES'),
 (35,'151300','PRESTAMOS','151301','PRESTAMOS OTORGADOS'),
 (36,'151300','PRESTAMOS','151350','PAGO DE PRESTAMOS'),
 (37,'710300','DIVISAS','710301','COMPRA DE DIVISAS'),
 (38,'710300','DIVISAS','710350','VENTA DE DIVISAS'),
 (39,'120100','IVA','120101','IVA'),
 (40,'120100','IVA','120150','IVA'),
 (41,'202000','ALMONEDA','202001','ENTRADA A INVENTARIO'),
 (42,'202000','ALMONEDA','202050','SALIDA DE INVENTARIO'),
 (43,'530100','PRODUCTOS','530101','PRODUCTOS PAGADOS'),
 (44,'530100','PRODUCTOS','530150','PRODUCTOS COBRADOS'),
 (45,'650200','DEMASIAS','650201','DEMASIA GENERADA'),
 (46,'650200','DEMASIAS','650250','DEMASIA PAGADA'),
 (47,'670300','ALMACENAJE','670350','ALMACENAJE COBRADO'),
 (48,'680300','SEGURO','680350','SEGURO COBRADO'),
 (49,'690300','MORATORIOS','690350','MORATORIOS COBRADOS'),
 (50,'620200','COSTO','620201','COSTOS'),
 (51,'620200','COSTO','620250','COSTOS'),
 (52,'620400','VENTAS','620401','VENTA REALIZADA'),
 (53,'620400','VENTAS','620450','VENTA CANCELADA'),
 (54,'999400','CAJA DIVISAS','999401','DOTACION DE DIVISAS A CAJA'),
 (55,'999400','CAJA DIVISAS','999450','RETIRO DE DIVISAS A CAJA'),
 (56,'910900','BOVEDA DIVISAS','910901','ENTRADA DE DIVISAS A BOVEDA'),
 (57,'910900','BOVEDA DIVISAS','910950','SALIDA DE DIVISAS A BOVEDA');
/*!40000 ALTER TABLE `cuentas` ENABLE KEYS */;


--
-- Definition of table `cuentasgastos`
--

DROP TABLE IF EXISTS `cuentasgastos`;
CREATE TABLE `cuentasgastos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` datetime DEFAULT NULL,
  `Cuenta` varchar(50) DEFAULT NULL,
  `Descripcion` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `id` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cuentasgastos`
--

/*!40000 ALTER TABLE `cuentasgastos` DISABLE KEYS */;
/*!40000 ALTER TABLE `cuentasgastos` ENABLE KEYS */;


--
-- Definition of table `depositos`
--

DROP TABLE IF EXISTS `depositos`;
CREATE TABLE `depositos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Folio` int(10) DEFAULT '0',
  `Fecha` datetime DEFAULT NULL,
  `Concepto` varchar(60) DEFAULT NULL,
  `Importe` double(15,5) DEFAULT '0.00000',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `Folio` (`Folio`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `depositos`
--

/*!40000 ALTER TABLE `depositos` DISABLE KEYS */;
/*!40000 ALTER TABLE `depositos` ENABLE KEYS */;


--
-- Definition of table `detallefactura`
--

DROP TABLE IF EXISTS `detallefactura`;
CREATE TABLE `detallefactura` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDFactura` int(10) DEFAULT NULL,
  `Cantidad` int(10) DEFAULT NULL,
  `Concepto` varchar(50) DEFAULT NULL,
  `Importe` double(15,5) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `Id` (`ID`),
  KEY `Idfactura` (`IDFactura`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallefactura`
--

/*!40000 ALTER TABLE `detallefactura` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallefactura` ENABLE KEYS */;


--
-- Definition of table `detallesajuste`
--

DROP TABLE IF EXISTS `detallesajuste`;
CREATE TABLE `detallesajuste` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDAjustes` int(10) DEFAULT NULL,
  `IDArticulo` int(10) DEFAULT NULL,
  `Codigo` varchar(20) DEFAULT NULL,
  `Descripcion` varchar(255) DEFAULT NULL,
  `Kilates` int(10) DEFAULT NULL,
  `Peso` double(15,5) DEFAULT NULL,
  `Precio` double(15,5) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDArticulo` (`IDArticulo`),
  KEY `IDSalidaInventario` (`IDAjustes`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallesajuste`
--

/*!40000 ALTER TABLE `detallesajuste` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallesajuste` ENABLE KEYS */;


--
-- Definition of table `detallescompras`
--

DROP TABLE IF EXISTS `detallescompras`;
CREATE TABLE `detallescompras` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `IDCompra` int(11) DEFAULT '0',
  `Tipo` int(11) DEFAULT '0',
  `Codigo` varchar(20) DEFAULT NULL,
  `Descripcion` varchar(255) DEFAULT NULL,
  `Kilates` int(11) DEFAULT '0',
  `Estado` varchar(50) DEFAULT NULL,
  `Cantidad` int(11) DEFAULT '0',
  `Peso` double(15,5) DEFAULT '0.00000',
  `Costo` double(15,5) DEFAULT '0.00000',
  `Precio` double(15,5) DEFAULT '0.00000',
  `Observaciones` varchar(250) DEFAULT NULL,
  `TipoPrenda` int(11) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallescompras`
--

/*!40000 ALTER TABLE `detallescompras` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallescompras` ENABLE KEYS */;


--
-- Definition of table `detallesempeno`
--

DROP TABLE IF EXISTS `detallesempeno`;
CREATE TABLE `detallesempeno` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDEmpeno` int(10) DEFAULT '0',
  `Codigo` varchar(15) DEFAULT NULL,
  `Tipo` int(10) DEFAULT '0',
  `Cantidad` int(10) DEFAULT '0',
  `Articulo` varchar(400) DEFAULT NULL,
  `Peso` double(15,3) DEFAULT '0.000',
  `Kilates` int(10) DEFAULT '0',
  `Avaluo` double(15,5) DEFAULT '0.00000',
  `Prestamo` double(15,5) DEFAULT '0.00000',
  `Estado` varchar(50) DEFAULT NULL,
  `Marca` varchar(100) DEFAULT NULL,
  `Modelo` varchar(100) DEFAULT NULL,
  `Serie` varchar(50) DEFAULT NULL,
  `Color` varchar(50) DEFAULT NULL,
  `Tamano` varchar(50) DEFAULT NULL,
  `Origen` smallint(5) DEFAULT '0',
  `Destino` smallint(5) DEFAULT '0',
  `TipoPrenda` int(10) DEFAULT '0',
  `CantidadPiedras` double(15,5) DEFAULT NULL,
  `PesoPiedras` double(15,5) DEFAULT NULL,
  `CantidadDiamantes` int(11) DEFAULT '0',
  `Puntos` double(15,5) DEFAULT '0.00000',
  `PrestamoDiamante` double(15,5) DEFAULT '0.00000',
  `Almoneda` int(1) DEFAULT '0',
  `Observaciones` varchar(250) DEFAULT NULL,
  `DemasiaPagada` int(1) DEFAULT '0',
  `IDTipoGarantia` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDEmpeno` (`IDEmpeno`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallesempeno`
--

/*!40000 ALTER TABLE `detallesempeno` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallesempeno` ENABLE KEYS */;


--
-- Definition of table `detallesempenoautos`
--

DROP TABLE IF EXISTS `detallesempenoautos`;
CREATE TABLE `detallesempenoautos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDEmpeno` int(10) DEFAULT '0',
  `MarcayModelo` varchar(80) DEFAULT '',
  `Marca` varchar(50) DEFAULT '',
  `Modelo` varchar(50) DEFAULT '',
  `Año` int(10) DEFAULT '0',
  `Color` varchar(50) DEFAULT NULL,
  `Placas` varchar(50) DEFAULT NULL,
  `Factura` varchar(50) DEFAULT NULL,
  `Agencia` varchar(50) DEFAULT NULL,
  `NumTarjetacircu` varchar(50) DEFAULT NULL,
  `NumMotor` varchar(50) DEFAULT NULL,
  `SerieChasis` varchar(50) DEFAULT NULL,
  `VIN` varchar(30) DEFAULT '',
  `RePuVe` varchar(30) DEFAULT '',
  `Kms` varchar(50) DEFAULT NULL,
  `Gas` varchar(50) DEFAULT NULL,
  `Aseguradora` varchar(50) DEFAULT NULL,
  `Poliza` varchar(50) DEFAULT NULL,
  `FechaVenci` date DEFAULT NULL,
  `Tipo` varchar(50) DEFAULT NULL,
  `EntregaFactura` smallint(1) DEFAULT '0',
  `TarjetaCircu` smallint(1) DEFAULT '0',
  `CopiaIfe` smallint(1) DEFAULT '0',
  `Tenencias` smallint(1) DEFAULT '0',
  `PolizaSeguro` smallint(1) DEFAULT '0',
  `CopiaLicencia` smallint(1) DEFAULT '0',
  `Importacion` smallint(1) DEFAULT '0',
  `Factu` smallint(1) DEFAULT '0',
  `Observaciones` varchar(250) DEFAULT NULL,
  `IDTipoGarantia` int(10) DEFAULT '0',
  `IDTipoBlindajeAutos` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDEmpeño` (`IDEmpeno`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallesempenoautos`
--

/*!40000 ALTER TABLE `detallesempenoautos` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallesempenoautos` ENABLE KEYS */;


--
-- Definition of table `detallesentradainventario`
--

DROP TABLE IF EXISTS `detallesentradainventario`;
CREATE TABLE `detallesentradainventario` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDEntrada` int(10) DEFAULT '0',
  `Codigo` varchar(20) DEFAULT NULL,
  `Tipo` int(10) DEFAULT '0',
  `Cantidad` int(10) DEFAULT '0',
  `Descripcion` varchar(255) DEFAULT NULL,
  `Peso` double(15,5) DEFAULT '0.00000',
  `Kilates` int(10) DEFAULT '0',
  `Precio` double(15,5) DEFAULT '0.00000',
  `Costo` double(15,5) DEFAULT '0.00000',
  `Estado` varchar(50) DEFAULT NULL,
  `Marca` varchar(100) DEFAULT NULL,
  `Modelo` varchar(100) DEFAULT NULL,
  `Serie` varchar(50) DEFAULT NULL,
  `Color` varchar(50) DEFAULT NULL,
  `Tamano` varchar(50) DEFAULT NULL,
  `TipoPrenda` int(11) DEFAULT '0',
  `Observaciones` varchar(250) DEFAULT NULL,
  `Tmp` int(10) DEFAULT '0',
  `IDEmpeno` int(10) DEFAULT '0',
  `SucursalOrigen` int(10) DEFAULT '0',
  `SucursalDestino` int(10) DEFAULT '0',
  `TipoEntrada` int(2) DEFAULT '0',
  `TipoSalida` int(2) DEFAULT '0',
  `PrecioVitrina` double(15,5) DEFAULT '0.00000',
  `CantidadPiedras` double(15,5) DEFAULT '0.00000',
  `PesoPiedras` double(15,5) DEFAULT '0.00000',
  `CantidadDiamantes` int(11) DEFAULT '0',
  `Puntos` double(15,5) DEFAULT '0.00000',
  `PrestamoDiamante` double(15,5) DEFAULT '0.00000',
  `IDDetallesEmpeno` int(11) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDEntrada` (`IDEntrada`),
  KEY `IDEmpeno` (`IDEmpeno`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallesentradainventario`
--

/*!40000 ALTER TABLE `detallesentradainventario` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallesentradainventario` ENABLE KEYS */;


--
-- Definition of table `detallessalida`
--

DROP TABLE IF EXISTS `detallessalida`;
CREATE TABLE `detallessalida` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDSalidaInventario` int(10) DEFAULT '0',
  `IDArticulo` int(10) DEFAULT '0',
  `Codigo` varchar(20) DEFAULT NULL,
  `Descripcion` varchar(255) DEFAULT NULL,
  `Kilates` int(5) DEFAULT '0',
  `Costo` double(15,5) DEFAULT '0.00000',
  `Peso` double(15,5) DEFAULT '0.00000',
  `Precio` double(15,5) DEFAULT '0.00000',
  `Tipo` int(5) DEFAULT '0',
  `Serie` varchar(50) DEFAULT NULL,
  `IDEmpeno` int(10) DEFAULT '0',
  `Observaciones` varchar(250) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDArticulo` (`IDArticulo`),
  KEY `IDSalidaInventario` (`IDSalidaInventario`),
  KEY `Idempeño` (`IDEmpeno`) USING BTREE
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallessalida`
--

/*!40000 ALTER TABLE `detallessalida` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallessalida` ENABLE KEYS */;


--
-- Definition of table `detallestraspasos`
--

DROP TABLE IF EXISTS `detallestraspasos`;
CREATE TABLE `detallestraspasos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDTraspaso` int(10) DEFAULT '0',
  `Codigo` varchar(20) DEFAULT NULL,
  `Descripcion` varchar(255) DEFAULT NULL,
  `Kilates` int(10) DEFAULT '0',
  `Peso` double(15,3) DEFAULT '0.000',
  `Precio` double(15,5) DEFAULT '0.00000',
  `Costo` double(15,5) DEFAULT '0.00000',
  `Cantidad` int(10) DEFAULT '0',
  `Tipo` int(10) DEFAULT '0',
  `Serie` varchar(50) DEFAULT NULL,
  `IDEmpeno` int(10) DEFAULT '0',
  `SucursalOrigen` int(10) DEFAULT '0',
  `SucursalDestino` int(10) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDEntrada` (`IDTraspaso`),
  KEY `Idempeño` (`IDEmpeno`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallestraspasos`
--

/*!40000 ALTER TABLE `detallestraspasos` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallestraspasos` ENABLE KEYS */;


--
-- Definition of table `detallesventas`
--

DROP TABLE IF EXISTS `detallesventas`;
CREATE TABLE `detallesventas` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `IDVenta` int(10) DEFAULT '0',
  `Codigo` varchar(15) DEFAULT NULL,
  `Articulo` varchar(255) DEFAULT NULL,
  `Kilates` int(10) DEFAULT '0',
  `Peso` double(15,5) DEFAULT '0.00000',
  `Costo` double(15,5) DEFAULT '0.00000',
  `Precio` double(15,5) DEFAULT '0.00000',
  `IDArticulo` int(11) DEFAULT '0',
  `Intereses` double(15,5) DEFAULT '0.00000',
  `Almacenaje` double(15,5) DEFAULT '0.00000',
  `Seguro` double(15,5) DEFAULT '0.00000',
  `Moratorios` double(15,5) DEFAULT '0.00000',
  `GtosVenta` double(15,5) DEFAULT '0.00000',
  `ImporteIva` double(15,5) DEFAULT '0.00000',
  `Devolucion` int(1) DEFAULT '0',
  `FechaDemasia` datetime DEFAULT NULL,
  `ImporteDescuento` double(15,5) DEFAULT '0.00000',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDVenta` (`IDVenta`),
  KEY `IDArticulo` (`IDArticulo`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `detallesventas`
--

/*!40000 ALTER TABLE `detallesventas` DISABLE KEYS */;
/*!40000 ALTER TABLE `detallesventas` ENABLE KEYS */;


--
-- Definition of table `diamantepuntos`
--

DROP TABLE IF EXISTS `diamantepuntos`;
CREATE TABLE `diamantepuntos` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Punto` varchar(15) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=19 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `diamantepuntos`
--

/*!40000 ALTER TABLE `diamantepuntos` DISABLE KEYS */;
INSERT INTO `diamantepuntos` (`ID`,`Punto`) VALUES 
 (1,'.01-.03'),
 (2,'.04-.07'),
 (3,'.08-.14'),
 (4,'.15-.17'),
 (5,'.18-.22'),
 (6,'.23-.29'),
 (7,'.30-.37'),
 (8,'.38-.45'),
 (9,'.46-.49'),
 (10,'.50-.69'),
 (11,'.70-.89'),
 (12,'.90-.99'),
 (13,'1.00-1.49'),
 (14,'1.50-1.99'),
 (15,'2.00-2.99'),
 (16,'3.00-3.99'),
 (17,'4.00-4.99'),
 (18,'5.00-5.99');
/*!40000 ALTER TABLE `diamantepuntos` ENABLE KEYS */;


--
-- Definition of table `divisas`
--

DROP TABLE IF EXISTS `divisas`;
CREATE TABLE `divisas` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Cancelado` tinyint(1) DEFAULT '0',
  `Folio` int(10) DEFAULT '0',
  `Fecha` datetime DEFAULT NULL,
  `IDDivisa` int(10) DEFAULT '0',
  `Importe` double(15,5) DEFAULT '0.00000',
  `Cantidad` int(10) DEFAULT '0',
  `Tipo` int(2) DEFAULT '0',
  `TipoEntrada` int(2) DEFAULT '0',
  `IDCliente` int(10) DEFAULT '0',
  `Efectivo` double(15,5) DEFAULT '0.00000',
  `ChequeDeposito` double(15,5) DEFAULT '0.00000',
  `Traspaso` double(15,5) DEFAULT '0.00000',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  `Notas` varchar(250) DEFAULT NULL,
  `PC` varchar(20) DEFAULT NULL,
  `Corte` int(1) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `id` (`ID`),
  KEY `Idcliente` (`IDCliente`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `divisas`
--

/*!40000 ALTER TABLE `divisas` DISABLE KEYS */;
/*!40000 ALTER TABLE `divisas` ENABLE KEYS */;


--
-- Definition of table `empeno`
--

DROP TABLE IF EXISTS `empeno`;
CREATE TABLE `empeno` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Cancelado` tinyint(1) NOT NULL DEFAULT '0',
  `Fecha` datetime DEFAULT NULL,
  `FechaOriginal` datetime DEFAULT NULL,
  `Movimiento` int(10) DEFAULT '0',
  `NumContrato` int(10) NOT NULL DEFAULT '0',
  `Folio` int(10) DEFAULT '0',
  `Prestamo` double(15,5) DEFAULT '0.00000',
  `PrestamoInicial` double(15,5) DEFAULT '0.00000',
  `Avaluo` double(15,5) DEFAULT '0.00000',
  `Origen` int(10) DEFAULT '0',
  `Destino` int(10) DEFAULT '0',
  `Vencimiento` date DEFAULT NULL,
  `FolioOrigen` int(10) DEFAULT '0',
  `FolioDestino` int(10) DEFAULT '0',
  `FechaMovimiento` datetime DEFAULT NULL,
  `IDUsuarioMov` int(10) DEFAULT '0',
  `Serie` int(10) DEFAULT '0',
  `Pagado` tinyint(1) DEFAULT '0',
  `PC` varchar(20) DEFAULT '0',
  `Corte` tinyint(1) DEFAULT '0',
  `Perdida` tinyint(1) DEFAULT '0',
  `IDCliente` int(10) DEFAULT '0',
  `Responsable` varchar(220) DEFAULT '',
  `Beneficiario` varchar(90) DEFAULT NULL,
  `Valuador` varchar(60) DEFAULT '0',
  `Notas` varchar(255) DEFAULT '0',
  `Tasa` double(15,5) DEFAULT '0.00000',
  `Almacenaje` double(15,5) DEFAULT '0.00000',
  `Seguro` double(15,5) DEFAULT '0.00000',
  `Operacion` double(15,5) DEFAULT '0.00000',
  `Comision` double(15,5) DEFAULT '0.00000',
  `Iva` double(15,5) DEFAULT '0.00000',
  `Cat` double(15,5) DEFAULT '0.00000',
  `Periodo` int(10) DEFAULT '0',
  `VenPeriodo` int(10) DEFAULT '0',
  `Almoneda` int(1) DEFAULT '0',
  `FechaAlmoneda` date DEFAULT NULL,
  `VenAlmoneda` int(10) DEFAULT '0',
  `TipoInteres` varchar(20) DEFAULT NULL,
  `TipoTasa` varchar(50) DEFAULT NULL,
  `IDSucursal` int(10) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `Pago` double(15,5) DEFAULT '0.00000',
  `Intereses` double(15,5) DEFAULT '0.00000',
  `ImporteAlmacenaje` double(15,5) DEFAULT '0.00000',
  `ImporteSeguro` double(15,5) DEFAULT '0.00000',
  `ImporteMoratorios` double(15,5) DEFAULT '0.00000',
  `ImportePerdida` double(15,5) DEFAULT '0.00000',
  `Descuento` double(15,5) DEFAULT '0.00000',
  `ImporteIva` double(15,5) DEFAULT '0.00000',
  `AutTasa` int(10) DEFAULT '0',
  `ChequeReferencia` varchar(50) DEFAULT NULL,
  `ImporteOtros` double(15,5) DEFAULT '0.00000',
  `DemasiaPagada` int(11) DEFAULT '0',
  `IDAutorizacion` int(11) DEFAULT '0',
  `Captura` int(1) DEFAULT '0',
  `NumBolsa` varchar(15) DEFAULT NULL,
  `Verificado` int(5) DEFAULT '0',
  `IDUsuarioAutoriza` int(10) DEFAULT '0',
  `TipoAutoriza` int(10) DEFAULT '0',
  `Ubicacion` varchar(250) DEFAULT NULL,
  `FolioNota` int(11) DEFAULT '0',
  `Efectivo` double(15,5) DEFAULT '0.00000',
  `Caja` varchar(250) DEFAULT NULL,
  `Cajon` varchar(250) DEFAULT NULL,
  `Fila` varchar(250) DEFAULT NULL,
  `Promocion` int(10) DEFAULT '0',
  `SaldoPuntosAnteriorEmp` double(15,5) DEFAULT '0.00000',
  `PuntosAcumuladosEmp` double(15,5) DEFAULT '0.00000',
  `SaldoPuntosActualEmp` double(15,5) DEFAULT '0.00000',
  `IDTarjetaEmp` int(10) DEFAULT '0',
  `DescuentoXPuntos` double(15,5) DEFAULT '0.00000',
  `SaldoPuntosAnterior` double(15,5) DEFAULT '0.00000',
  `PuntosUsados` double(15,5) DEFAULT '0.00000',
  `PuntosAcumulados` double(15,5) DEFAULT '0.00000',
  `SaldoPuntosActual` double(15,5) DEFAULT '0.00000',
  `IDTarjeta` int(10) DEFAULT '0',
  `IDCoTitular` int(10) DEFAULT '0',
  `Cheque` tinyint(4) DEFAULT '0',
  `SalarioMin` tinyint(4) DEFAULT '0',
  `ValorSalarioMin` double(15,5) DEFAULT '0.00000',
  `ValorUDI` double(15,6) DEFAULT '0.000000',
  `UltDigitosTarj` varchar(4) DEFAULT '',
  `IDTipoOperacion` int(10) DEFAULT '0',
  `ClaveTipoOperacion` int(10) DEFAULT '0',
  `IDInstrumentoMonetario` int(10) DEFAULT '0',
  `IDTipoMoneda` int(10) DEFAULT '0',
  `IdTipoAlerta` int(10) DEFAULT '0',
  `DescTipoAlerta` varchar(3000) DEFAULT '',
  `IdTipoPrenda` int(10) DEFAULT '0',
  `IDEmpenoOrigen` int(11) DEFAULT '0',
  `IDEmpenoDestino` int(11) DEFAULT '0',
  PRIMARY KEY (`ID`,`NumContrato`) USING BTREE,
  KEY `Folio` (`Folio`),
  KEY `FolioDestino` (`FolioDestino`),
  KEY `ID` (`ID`),
  KEY `IDCliente` (`IDCliente`),
  KEY `IDSucursal` (`IDSucursal`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `empeno`
--

/*!40000 ALTER TABLE `empeno` DISABLE KEYS */;
/*!40000 ALTER TABLE `empeno` ENABLE KEYS */;


--
-- Definition of table `entradainventario`
--

DROP TABLE IF EXISTS `entradainventario`;
CREATE TABLE `entradainventario` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` datetime DEFAULT NULL,
  `Folio` int(10) DEFAULT '0',
  `TipoEntrada` int(11) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `entradainventario`
--

/*!40000 ALTER TABLE `entradainventario` DISABLE KEYS */;
/*!40000 ALTER TABLE `entradainventario` ENABLE KEYS */;


--
-- Definition of table `estado`
--

DROP TABLE IF EXISTS `estado`;
CREATE TABLE `estado` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Estado` varchar(50) DEFAULT NULL,
  `IDTipo` int(5) DEFAULT '0',
  `Ordenamiento` int(1) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=8 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `estado`
--

/*!40000 ALTER TABLE `estado` DISABLE KEYS */;
INSERT INTO `estado` (`ID`,`Estado`,`IDTipo`,`Ordenamiento`) VALUES 
 (1,'E',1,1),
 (2,'B',1,2),
 (3,'R',1,3),
 (4,'M',1,4),
 (5,'BCO. COMER./G-H',4,1),
 (6,'LIG. COLOR/I-K',4,2),
 (7,'CON COLOR/L-M',4,3);
/*!40000 ALTER TABLE `estado` ENABLE KEYS */;


--
-- Definition of table `estadospais`
--

DROP TABLE IF EXISTS `estadospais`;
CREATE TABLE `estadospais` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Codigo` varchar(2) DEFAULT NULL,
  `descripcion` varchar(50) DEFAULT NULL,
  `Ordenamiento` int(2) DEFAULT '0',
  `Actualizar` tinyint(4) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=33 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `estadospais`
--

/*!40000 ALTER TABLE `estadospais` DISABLE KEYS */;
INSERT INTO `estadospais` (`ID`,`Codigo`,`descripcion`,`Ordenamiento`,`Actualizar`) VALUES 
 (1,'AS','AGUASCALIENTES',1,0),
 (2,'BC','BAJA CALIFORNIA',2,0),
 (3,'BS','BAJA CALIFORNIA SUR',3,0),
 (4,'CC','CAMPECHE',4,0),
 (5,'CS','CHIAPAS',5,0),
 (6,'CH','CHIHUAHUA',6,0),
 (7,'CL','COAHUILA',7,0),
 (8,'CM','COLIMA',8,0),
 (9,'DF','DISTRITO FEDERAL',9,0),
 (10,'DG','DURANGO',10,0),
 (11,'GT','GUANAJUATO',11,0),
 (12,'GR','GUERRERO',12,0),
 (13,'HG','HIDALGO',13,0),
 (14,'JC','JALISCO',14,0),
 (15,'MC','MEXICO',15,0),
 (16,'MN','MICHOACAN',16,0),
 (17,'MS','MORELOS',17,0),
 (18,'NT','NAYARIT',18,0),
 (19,'NL','NUEVO LEON',19,0),
 (20,'OC','OAXACA',20,0),
 (21,'PL','PUEBLA',21,0),
 (22,'QT','QUERETARO',22,0),
 (23,'QR','QUINTANA ROO',23,0),
 (24,'SP','SAN LUIS POTOSI',24,0),
 (25,'SL','SINALOA',25,0),
 (26,'SR','SONORA',26,0),
 (27,'TC','TABASCO',27,0),
 (28,'TS','TAMAULIPAS',28,0),
 (29,'TL','TLAXCALA',29,0),
 (30,'VZ','VERACRUZ',30,0),
 (31,'YN','YUCATAN',31,0),
 (32,'ZS','ZACATECAS',32,0);
/*!40000 ALTER TABLE `estadospais` ENABLE KEYS */;


--
-- Definition of table `facturas`
--

DROP TABLE IF EXISTS `facturas`;
CREATE TABLE `facturas` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Folio` varchar(10) DEFAULT NULL,
  `Fecha` datetime DEFAULT NULL,
  `Cliente` int(10) DEFAULT '0',
  `Subtotal` double(15,5) DEFAULT '0.00000',
  `Iva` double(15,5) DEFAULT '0.00000',
  `Total` double(15,5) DEFAULT '0.00000',
  `Notas` varchar(250) DEFAULT NULL,
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `Id` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `facturas`
--

/*!40000 ALTER TABLE `facturas` DISABLE KEYS */;
/*!40000 ALTER TABLE `facturas` ENABLE KEYS */;


--
-- Definition of table `folios`
--

DROP TABLE IF EXISTS `folios`;
CREATE TABLE `folios` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Folio` int(10) DEFAULT '0',
  `Serie` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `folios`
--

/*!40000 ALTER TABLE `folios` DISABLE KEYS */;
INSERT INTO `folios` (`ID`,`Folio`,`Serie`) VALUES 
 (1,1,1),
 (2,1,2),
 (3,1,3);
/*!40000 ALTER TABLE `folios` ENABLE KEYS */;


--
-- Definition of table `gastos`
--

DROP TABLE IF EXISTS `gastos`;
CREATE TABLE `gastos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Cancelado` tinyint(1) DEFAULT '0',
  `Fecha` datetime DEFAULT NULL,
  `Folio` int(10) DEFAULT '0',
  `Concepto` varchar(100) DEFAULT NULL,
  `Importe` double(15,5) DEFAULT '0.00000',
  `CuentaGastos` int(10) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(5) DEFAULT '0',
  `PC` varchar(45) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `Folio` (`Folio`),
  KEY `ID` (`ID`),
  KEY `Idcuentagasto` (`CuentaGastos`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `gastos`
--

/*!40000 ALTER TABLE `gastos` DISABLE KEYS */;
/*!40000 ALTER TABLE `gastos` ENABLE KEYS */;


--
-- Definition of table `identificaciones`
--

DROP TABLE IF EXISTS `identificaciones`;
CREATE TABLE `identificaciones` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Identificacion` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `identificaciones`
--

/*!40000 ALTER TABLE `identificaciones` DISABLE KEYS */;
/*!40000 ALTER TABLE `identificaciones` ENABLE KEYS */;


--
-- Definition of table `kilatajes`
--

DROP TABLE IF EXISTS `kilatajes`;
CREATE TABLE `kilatajes` (
  `ID` int(10) unsigned NOT NULL,
  `Clave` int(11) DEFAULT '0',
  `Descripcion` varchar(50) DEFAULT NULL,
  `IDTipo` int(5) DEFAULT '0',
  `Ordenamiento` int(1) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `kilatajes`
--

/*!40000 ALTER TABLE `kilatajes` DISABLE KEYS */;
INSERT INTO `kilatajes` (`ID`,`Clave`,`Descripcion`,`IDTipo`,`Ordenamiento`) VALUES 
 (1,1,'10K',1,2),
 (2,2,'14K',1,3),
 (3,3,'18K',1,4),
 (10,10,'Electronicos',0,0),
 (11,11,'Relojes',0,0),
 (12,12,'Diamantes',0,0),
 (13,13,'Otros',0,0),
 (14,14,'8K',1,1),
 (21,21,'24K',1,6),
 (18,18,'C. LIMPIO/VVS2',4,1),
 (19,19,'LIG. DEF./VS2',4,2),
 (20,20,'CON DEF./SI2',4,3),
 (22,22,'22K',1,5),
 (23,23,'.720',10,1),
 (24,24,'.825',10,2),
 (25,25,'.925',10,3),
 (26,26,'.999',10,4);
/*!40000 ALTER TABLE `kilatajes` ENABLE KEYS */;


--
-- Definition of table `marcas`
--

DROP TABLE IF EXISTS `marcas`;
CREATE TABLE `marcas` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Descripcion` varchar(100) DEFAULT NULL,
  `Ordenamiento` int(2) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=102 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `marcas`
--

/*!40000 ALTER TABLE `marcas` DISABLE KEYS */;
INSERT INTO `marcas` (`ID`,`Descripcion`,`Ordenamiento`) VALUES 
 (1,'DAEWOO',0),
 (2,'IPHONE',0),
 (3,'IPOD',0),
 (4,'LG',0),
 (5,'SONY',0),
 (6,'XBOX',0),
 (7,'VARIAS',0),
 (8,'VARIAS',0),
 (9,'TOSHIBA',0),
 (10,'APPLE',0),
 (11,'COMPAQ',0),
 (12,'GUITARRA',0),
 (13,'MERCURIO',0),
 (14,'VARIAS',0),
 (15,'PANASONIC',0),
 (16,'***',0),
 (17,'CRAFSMAN',0),
 (18,'EMACHINES',0),
 (19,'NOKIA',0),
 (20,'SAMSUNG',0),
 (21,'CASIO',0),
 (22,'LANIX',0),
 (23,'GATEWAY',0),
 (24,'ACER',0),
 (25,'DELL',0),
 (26,'WII',0),
 (27,'HP',0),
 (28,'SONY',0),
 (29,'MOTOROLA',0),
 (30,'ACER',0),
 (31,'LG',0),
 (32,'TRUPER',0),
 (33,'IPOD',0),
 (34,'ALCATEL',0),
 (35,'TOSHIBA',0),
 (36,'WILSON LONG',0),
 (37,'TECNO LINE',0),
 (38,'HTC',0),
 (39,'KRAFMAN',0),
 (40,'IMAC',0),
 (41,'STARRETT',0),
 (42,'PHILIPS',0),
 (43,'MSI',0),
 (44,'RCA',0),
 (45,'BLACBERRY',0),
 (46,'BENQ',0),
 (47,'VIZIO',0),
 (48,'SILVANIA',0),
 (49,'MERCURIO',0),
 (50,'KODAK',0),
 (51,'LENOVO',0),
 (52,'BLACK&DECKER',0),
 (53,'PIONEER',0),
 (54,'SHARP',0),
 (55,'LX-50',0),
 (56,'OLIMPUS',0),
 (57,'IPAD',0),
 (58,'FUJIFILM',0),
 (59,'MITSUI',0),
 (60,'NIKON',0),
 (61,'ASUS',0),
 (62,'HUAWEI',0),
 (63,'LENOVO',0),
 (64,'SONY ERICSSON',0),
 (65,'SANYO',0),
 (66,'PANASONIC',0),
 (67,'JVC',0),
 (68,'ATVIO',0),
 (69,'MITSUI',0),
 (70,'DAEWOO',0),
 (71,'EMERSON',0),
 (72,'MACBOOOKPRO',0),
 (73,'HITACHI',0),
 (74,'COBY',0),
 (75,'MAC',0),
 (76,'ZTE',0),
 (77,'BOSCH',0),
 (78,'DEWALT',0),
 (79,'HP',0),
 (80,'PHILIPS',0),
 (81,'BLACK&DECKER',0),
 (82,'CANON',0),
 (83,'GT',0),
 (84,'MAKITA',0),
 (85,'REDLINE',0),
 (86,'MONSTER',0),
 (87,'OSTER',0),
 (88,'SUPERMATIC',0),
 (89,'AURUS',0),
 (90,'MYTEL',0),
 (91,'PROTEUS',0),
 (92,'GE',0),
 (93,'SKIL',0),
 (94,'EASY',0),
 (95,'ACOUSTIK',0),
 (96,'ACTRON',0),
 (97,'TITAN',0),
 (98,'SPLASH',0),
 (99,'TRUPER',0),
 (100,'PIXXO',0),
 (101,'GENERAL',0);
/*!40000 ALTER TABLE `marcas` ENABLE KEYS */;


--
-- Definition of table `medios`
--

DROP TABLE IF EXISTS `medios`;
CREATE TABLE `medios` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Descripcion` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=19 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `medios`
--

/*!40000 ALTER TABLE `medios` DISABLE KEYS */;
INSERT INTO `medios` (`ID`,`Descripcion`) VALUES 
 (1,'REFERIDO'),
 (3,'VOLANTES'),
 (5,'OTROS'),
 (6,'RADIO'),
 (7,'PERIFONEO'),
 (8,'INTERNET'),
 (9,'ESPECTACULAR'),
 (11,'PASO DE LOCAL'),
 (12,'TV'),
 (13,'REVISTA'),
 (14,'PERIODICO'),
 (15,'EVENTO ESPECIAL'),
 (16,'PROMOCION'),
 (17,'YA ES CLIENTE');
/*!40000 ALTER TABLE `medios` ENABLE KEYS */;


--
-- Definition of table `mld_actividad_vulnerable`
--

DROP TABLE IF EXISTS `mld_actividad_vulnerable`;
CREATE TABLE `mld_actividad_vulnerable` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=18 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_actividad_vulnerable`
--

/*!40000 ALTER TABLE `mld_actividad_vulnerable` DISABLE KEYS */;
INSERT INTO `mld_actividad_vulnerable` (`Id`,`Clave`,`Descripcion`) VALUES 
 (1,'JYS','La venta de boletos, fichas o cualquier otro tipo de comprobante similar para la práctica de juegos con apuesta, concursos o sorteos que realicen organismos descentralizados, así como el pago del valor que representen dichos boletos, fichas o recibos o, en general, la entrega o pago de premios y la realización de cualquier operación financiera.'),
 (2,'TSC','La emisión o comercialización, habitual o profesional, de tarjetas de servicios o de crédito que no sean emitidas o comercializadas por Entidades Financieras.'),
 (3,'TPP','La emisión o comercialización, habitual o profesional de tarjetas prepagadas, vales o cupones, impresos o electrónicos, que puedan ser utilizados o canjeados para la adquisición de bienes o servicios, que no sean emitidos o comercializados por Entidades Financieras.'),
 (4,'TDR','La emisión o comercialización, habitual o profesional, de monederos electrónicos, certificados, o cupones, en los que, sin que exista un depósito previo del titular de dichos instrumentos, le sean abonados recursoso a los mismos provenientes de premios, promociones, devoluciones o derivados de recompensas comerciales y puedan ser utilizados para la adquisición de bienes o servicios en establecimeintos distintos al emisor de los referidos instrumentos o para la disposición de dinero en efectivo a través de cajeros automáticos o terminales puntos de venta o cualquier otro medio.'),
 (5,'CHV','La emisión y comercialización habitual o profesional de cheques de viajero, distinta a la realizada por las Entidades Financieras.'),
 (6,'MPC','El ofrecimiento habitual o profesional de operaciones de mutuo o de garantía o de otorgamiento de préstamos o créditos, con o sin garantía, por parte de sujetos distintos a las Entidades Financieras.'),
 (7,'INM','La prestación habitual o profesional de servicios de construcción o desarrollo de bienes inmuebles o de intermediación en la transmisión de la propiedad o constitución de derechos sobre dichos bienes, en los que se involucren operaciones de compra o venta de los propios bienes por cuenta o a favor de clientes de quienes presten dichos servicios.'),
 (8,'MJR','La comercialización o intermediación habitual o profesional de Metales Preciosos, Piedras Preciosas, joyas o relojes, en las que se involucren operaciones de compra o venta de dichos bienes.'),
 (9,'OBA','La subasta o comercialización habitual o profesional de obras de arte, en las que se involucren operaciones de compra o venta de dichos bienes.'),
 (10,'VEH','La comercialización o distribución habitual profesional de vehículos, nuevos o usados, ya sean aéreos, marítimos o terrestres.'),
 (11,'BLI','La prestación habitual o profesional de servicios de blindaje de vehículos terrestres, nuevos o usados, así como de bienes inmuebles.'),
 (12,'TCV','La prestación habitual o profesional de servicios de traslado o custodia de dinero o valores, con excepción de aquellos en los que intervenga el Banco de México y las instituciones dedicadas al depósito de valores.'),
 (13,'SPR','La prestación de servicios profesionales, de manera independiente, sin que medie relación laboral con el cliente respectivo, en aquellos casos en los que se prepare para un cliente o se lleven a cabo en nombre y representación del cliente cualquiera de las operaciones establecidas en el Artículo 17 fracción XI de la LFPIORPI.'),
 (14,'FEP','La prestación de servicios de fe pública, en los términos establecidos en el Artículo 17 fracción XII de la LFPIORPI.'),
 (15,'DON','La recepción de donativos, por parte de las asociaciones y sociedades sin fines de lucro.'),
 (16,'ADU','La prestación de servicios de comercio exterior como agente o apoderado aduanal, mediante autorización otorgada por la Secretaría de Hacienda y Crédito Público, para promover por cuenta ajena, el despacho de mercancías, en los diferentes regímenes aduaneros previstos en la Ley Aduanera, de las mercancías establecidas en el Artículo 17 fracción XIV de la LFPIORPI.'),
 (17,'ARI','La constitución de derechos personales de uso o goce de bienes inmuebles.');
/*!40000 ALTER TABLE `mld_actividad_vulnerable` ENABLE KEYS */;


--
-- Definition of table `mld_actividades_economicas`
--

DROP TABLE IF EXISTS `mld_actividades_economicas`;
CREATE TABLE `mld_actividades_economicas` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(7) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=268 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_actividades_economicas`
--

/*!40000 ALTER TABLE `mld_actividades_economicas` DISABLE KEYS */;
INSERT INTO `mld_actividades_economicas` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'0100008','AGRICULTURA',0),
 (2,'0200006','GANADERÍA',0),
 (3,'0300004','SILVICULTURA',0),
 (4,'0400002','PESCA',0),
 (5,'0500000','CAZA',0),
 (6,'1100007','EXTRACCIÓN Y BENEFICIO DE CARBÓN MINERAL Y GRAFITO',0),
 (7,'1200005','EXTRACCIÓN DE PETRÓLEO CRUDO Y GAS NATURAL',0),
 (8,'1300003','EXTRACCIÓN Y BENEFICIO DE MINERALES METÁLICOS',0),
 (9,'1311018','EXTRACCION Y BENEFICIO DE MINERAL DE HIERRO',0),
 (10,'1322015','EXTRACCION Y BENEFICIO DE MERCURIO Y ANTIMONIO',0),
 (11,'1329011','EXTRACCION Y BENEFICIO DE COBRE  PLOMO  ZINC Y OTROS MINERALES NO FERROSOS',0),
 (12,'1400001','EXTRACCIÓN DE MINERALES NO METÁLICOS, EXCEPTO SAL',0),
 (13,'1500009','EXPLOTACIÓN DE SAL',0),
 (14,'2000008','FABRICACIÓN DE ALIMENTOS',0),
 (15,'2012011','EMPACADORA DE CONSERVAS ALIMENTICIAS',0),
 (16,'2012029','EMPACADORA DE FRUTAS Y LEGUMBRES',0),
 (17,'2025014','BENEFICIO DE CAFE EXCEPTO MOLIENDA Y TOSTADO',0),
 (18,'2049022','FABRICACION DE CARNES FRIAS Y EMBUTIDOS',0),
 (19,'2100006','FABRICACIÓN Y ELABORACIÓN DE BEBIDAS (AGUA, REFRESCOS, CERVEZA, VINOS Y LICORES)',0),
 (20,'2300002','INDUSTRIA TEXTIL (FABRICACIÓN DE: HILADOS Y TEJIDOS)',0),
 (21,'2400000','FABRICACIÓN DE PRENDAS DE VESTIR Y OTROS ARTÍCULOS CONFECCIONADOS CON TEXTILES Y OTROS MATERIALES EXCEPTO CALZADO',0),
 (22,'2500008','FABRICACIÓN DE CALZADO E INDUSTRIA DEL CUERO',0),
 (23,'2600006','INDUSTRIA Y PRODUCTOS DE MADERA Y CORCHO; EXCEPTO MUEBLES',0),
 (24,'2711019','FABRICACION DE MUEBLES DE MADERA',0),
 (25,'2711027','FABRICACION DE MUEBLES DE MATERIAL SINTETICO',0),
 (26,'2800002','INDUSTRIA DEL PAPEL',0),
 (27,'2900000','INDUSTRIAS EDITORIAL, DE IMPRESIÓN Y CONEXAS',0),
 (28,'3000007','INDUSTRIA QUÍMICA',0),
 (29,'3021011','FABRICACION DE ABONOS Y FERTILIZANTES QUIMICOS',0),
 (30,'3100005','REFINACIÓN DE PETRÓLEO Y DERIVADOS DEL CARBÓN MINERAL',0),
 (31,'3111010','FABRICACION DE GASOLINA Y OTROS PRODUCTOS DERIVADOS DE LA REFINACION DE PETROLEO',0),
 (32,'3112018','FABRICACION DE PRODUCTOS PETROQUIMICOS BASICOS',0),
 (33,'3113016','FABRICACION DE ACEITES Y LUBRICANTES',0),
 (34,'3200003','FABRICACIÓN DE PRODUCTOS DE HULE Y DE PLÁSTICO',0),
 (35,'3300001','FABRICACIÓN DE PRODUCTOS DE MINERALES NO METÁLICOS; EXCEPTO DEL PETRÓLEO Y DEL CARBÓN MINERAL',0),
 (36,'3322013','FABRICACION DE CRISTALES PARA AUTOMOVIL',0),
 (37,'3331022','FABRICACION DE LADRILLOS',0),
 (38,'3341013','FABRICACION DE CEMENTO',0),
 (39,'3400009','INDUSTRIAS METÁLICAS BÁSICAS',0),
 (40,'3411022','FUNDICION DE FIERRO Y ACERO',0),
 (41,'3411030','PLANTA METALURGICA',0),
 (42,'3412012','FABRICACION DE LAMINAS DE HIERRO Y ACERO',0),
 (43,'3413010','FABRICACION DE TUBOS DE HIERRO Y ACERO',0),
 (44,'3500007','FABRICACIÓN DE PRODUCTOS METÁLICOS; EXCEPTO MAQUINARIA Y EQUIPO',0),
 (45,'3599026','FABRICACION DE CAJAS FUERTES',0),
 (46,'3600005','FABRICACIÓN, ENSAMBLE Y REPARACIÓN DE MAQUINARIA, EQUIPO Y SUS PARTES; EXCEPTO LOS ELÉCTRICOS',0),
 (47,'3700003','FABRICACIÓN Y ENSAMBLE DE MAQUINARIA, EQUIPO, APARATOS, ACCESORIOS Y ARTÍCULOS ELÉCTRICOS, ELECTRÓNICOS Y SUS PARTES',0),
 (48,'3800001','CONSTRUCCIÓN, RECONSTRUCCIÓN Y ENSAMBLE DE EQUIPO DE TRANSPORTE Y SUS PARTES',0),
 (49,'3819010','FABRICACION DE REFACCIONES Y ACCESORIOS AUTOMOTRICES',0),
 (50,'3831014','FABRICACION Y REPARACION DE BUQUES Y BARCOS',0),
 (51,'3832012','FABRICACION ENSAMBLE Y REPARACION DE AERONAVES',0),
 (52,'3933018','FABRICACION DE ARTICULOS DE QUINCALLERIA Y BISUTERIA',0),
 (53,'3997014','FABRICACION DE ARMAS',0),
 (54,'4100004','CONTRATACIÓN DE OBRAS COMPLETAS DE CONSTRUCCIÓN (CASAS, DEPARTAMENTOS, INMUEBLES, PAVIMENTACIÓN, NO RESIDENCIALES, VIAS DE COMUNICACIÓN)',0),
 (55,'4111019','CONSTRUCCION DE CASAS Y TECHOS DESARMABLES',0),
 (56,'4111027','CONSTRUCCION DE INMUEBLES',0),
 (57,'4112017','CONSTRUCCION DE EDIFICIOS PARA OFICINAS ESCUELAS HOSPITALES HOTELES Y OTROS NO RESIDENCIALES',0),
 (58,'4113015','CONSTRUCCION DE EDIFICIOS INDUSTRIALES Y PARA FINES ANALOGOS',0),
 (59,'4121018','CONSTRUCCION DE VIAS DE COMUNICACION',0),
 (60,'4199015','CONSTRUCCION DE ESTADIOS MONUMENTOS Y OTRAS OBRAS DE INGENIERIA',0),
 (61,'5012018','DISTRIBUCION DE ENERGIA ELECTRICA',0),
 (62,'6100002','COMPRAVENTA DE ALIMENTOS, BEBIDAS Y PRODUCTOS DE TABACO',0),
 (63,'6121024','COMPRAVENTA DE GANADO MAYOR EN PIE',0),
 (64,'6121032','COMPRAVENTA DE GANADO MENOR EN PIE',0),
 (65,'6131023','TIENDA DE ABARROTES Y MISCELANEA',0),
 (66,'6200000','COMPRAVENTA DE PRENDAS DE VESTIR Y OTROS ARTÍCULOS DE USO PERSONAL',0),
 (67,'1321017','EXTRACCION Y BENEFICIO DE ORO PLATA Y OTROS METALES PRECIOSOS',0),
 (68,'3921013','FABRICACION DE RELOJES',0),
 (69,'3932010','FABRICACION DE ARTICULOS DE JOYERIA',0),
 (70,'3932036','TALLADO DE PIEDRAS PRECIOSAS',0),
 (71,'6225016','COMPRAVENTA DE ARTICULOS DE PLATA',0),
 (72,'6225024','COMPRAVENTA DE JOYAS',0),
 (73,'6225032','COMPRAVENTA DE RELOJES',0),
 (74,'6999017','COMPRAVENTA DE DIAMANTES',0),
 (75,'9900916','COMPRAVENTA DE ARTICULOS DE ORO',0),
 (76,'9900917','COMPRAVENTA DE ARTICULOS DE PLATINO',0),
 (77,'9900918','COMPRAVENTA DE AGUAMARINAS,  ESMERALDAS, RUBÍES, TOPACIOS, TURQUESAS Y/O ZAFIROS',0),
 (78,'9900919','COMPRAVENTA DE PLATA, ORO O PLATINO A GRANEL',0),
 (79,'6325014','COMPRAVENTA DE ANTIGÜEDADES',0),
 (80,'8832017','GALERIAS DE ARTES GRAFICAS Y MUSEOS',0),
 (81,'9900920','COMPRAVENTA DE OBRAS DE ARTE',0),
 (82,'9900921','CASA DE SUBASTAS DE OBRAS DE ARTE, JOYAS Y/O ANTIGÜEDADES',0),
 (83,'6300008','COMPRAVENTA DE ARTÍCULOS PARA EL HOGAR (ELECTRODOMESTICOS, REFACCIONES, LOZA Y PORCELANA, ANTIGUEDADES)',0),
 (84,'6400006','COMPRAVENTA EN TIENDAS DE AUTOSERVICIO Y DE DEPARTAMENTOS ESPECIALIZADOS POR LÍNEA DE MERCANCÍAS',0),
 (85,'6500004','COMPRAVENTA DE GASES, COMBUSTIBLES Y LUBRICANTES',0),
 (86,'6513015','COMPRAVENTA DE GASOLINA Y DIESEL',0),
 (87,'6514013','COMPRAVENTA DE PETROLEO COMBUSTIBLE',0),
 (88,'6515011','COMPRAVENTA DE LUBRICANTES',0),
 (89,'6600002','COMPRAVENTA DE MATERIAS PRIMAS, MATERIALES Y AUXILIARES (ALGODÓN, CEMENTO, SANITARIOS, PIELES, FERRETERIA, MADERA, PINTURAS)',0),
 (90,'6691019','COMPRAVENTA DE FERTILIZANTES Y PLAGUICIDAS',0),
 (91,'6695011','COMPRAVENTA DE SUBSTANCIAS QUIMICAS PARA LA INDUSTRIA',0),
 (92,'6700000','COMPRAVENTA DE MAQUINARIA, EQUIPO, INSTRUMENTOS, APARATOS Y HERRAMIENTAS, SUS REFACCIONES Y ACCESORIOS',0),
 (93,'6712013','COMPRAVENTA DE ARTICULOS PARA LA EXPLOTACION DE MINAS',0),
 (94,'6811013','COMPRAVENTA DE AUTOMOVILES Y CAMIONES NUEVOS',0),
 (95,'6812011','COMPRAVENTA DE AUTOMOVILES Y CAMIONES USADOS',0),
 (96,'9900922','COMPRAVENTA DE VEHICULOS MARÍTIMOS',0),
 (97,'6819033','COMPRAVENTA DE VEHICULOS AEREOS',0),
 (98,'6813027','COMPRAVENTA DE MOTOCICLETAS Y SUS ACCESORIOS',0),
 (99,'6819017','COMPRAVENTA DE PARTES Y REFACCIONES PARA VEHICULOS TERRESTRES, AÉREOS Y MARÍTIMOS',0),
 (100,'4111051','DESARROLLADORES DE VIVIENDA',0),
 (101,'6900006','COMPRAVENTA DE BIENES INMUEBLES Y ARTÍCULOS DIVERSOS',0),
 (102,'6911053','COMPRAVENTA DE TERRENOS',0),
 (103,'8313017','SERVICIO DE CORREDORES DE BIENES RAICES',0),
 (104,'6991013','COMPRAVENTA DE ARMAS DE FUEGO',0),
 (105,'6992011','AGENCIAS DE RIFAS Y SORTEOS (QUINIELAS Y LOTERIA)',0),
 (106,'8829022','HIPODROMO',0),
 (107,'9900910','SALAS DE JUEGOS Y APUESTAS',0),
 (108,'9900911','ORGANIZACIÓN DE FERIAS REGIONALES CON APUESTAS',0),
 (109,'9900912','ORGANIZACIÓN DE CARRERAS DE CABALLOS O PELEAS DE GALLOS EN ESCENARIOS TEMPORALES',0),
 (110,'7100001','TRANSPORTE TERRESTRE',0),
 (111,'7200009','TRANSPORTE POR AGUA',0),
 (112,'7300007','TRANSPORTE AÉREO',0),
 (113,'7312010','SERVICIOS RELACIONADOS CON EL TRANSPORTE EN AERONAVES CON MATRICULA EXTRANJERA',0),
 (114,'7400005','SERVICIOS CONEXOS AL TRANSPORTE',0),
 (115,'8429038','EMPRESAS DE SEGURIDAD PRIVADA',0),
 (116,'8429046','EMPRESAS TRANSPORTADORAS DE VALORES',0),
 (117,'9900924','EMPRESAS DE CUSTODIA DE VALORES',0),
 (118,'7512016','AGENCIA DE TURISMO',0),
 (119,'7513014','AGENCIA ADUANAL',0),
 (120,'9900928','AGENTE ADUANAL',0),
 (121,'8524010','ALQUILER O RENTA DE AUTOMOVILES SIN CHOFER',0),
 (122,'7519020','ALQUILER DE LANCHAS Y VELEROS',0),
 (123,'7519038','RENTA DE VEHICULOS AEREOS',0),
 (124,'8311011','ALQUILER DE TERRENOS LOCALES Y EDIFICIOS NO RESIDENCIALES',0),
 (125,'8312019','ARRENDAMIENTO DE INMUEBLES RESIDENCIALES',0),
 (126,'7600001','COMUNICACIONES',0),
 (127,'8114019','SERVICIOS DE FONDOS Y FIDEICOMISOS DE FOMENTO ECONOMICO',0),
 (128,'8123010','INSTITUCIONES DE BANCA MÚLTIPLE',0),
 (129,'9900929','INSTITUCIONES DE LA BANCA DE DESARROLLO',0),
 (130,'8123052','SOCIEDADES DE AHORRO Y PRESTAMO',0),
 (131,'8123060','SOCIEDADES DE AHORRO Y CREDITO POPULAR',0),
 (132,'8123078','SOCIEDADES FINANCIERAS DE OBJETO LIMITADO',0),
 (133,'8123086','SOCIEDADES FINANCIERAS DE OBJETO MULTIPLE REGULADAS',0),
 (134,'8123094','SOCIEDADES FINANCIERAS DE OBJETO MULTIPLE NO REGULADAS',0),
 (135,'8131021','ALMACENES DE DEPOSITO',0),
 (136,'8132029','UNIONES DE CREDITO',0),
 (137,'8133027','COMPAÑIAS DE FIANZAS',0),
 (138,'8142010','SOCIEDADES DE INVERSION',0),
 (139,'8151029','COMPAÑIAS DE SEGUROS PRIVADAS',0),
 (140,'8200008','SERVICIOS COLATERALES A INSTITUCIONES FINANCIERAS Y DE SEGUROS',0),
 (141,'8211013','INVERSIONISTA',0),
 (142,'8211021','AGENTE DE BOLSA',0),
 (143,'8211047','CASAS DE BOLSA',0),
 (144,'8219017','AGENTE DE SEGUROS',0),
 (145,'8219025','CASA DE CAMBIO',0),
 (146,'6999992','CENTROS CAMBIARIOS',0),
 (147,'8219033','CORRESPONSAL BANCARIO',0),
 (148,'8219041','CAJA DE AHORROS',0),
 (149,'8219075','FACTORING',0),
 (150,'8511033','ARRENDADORAS FINANCIERAS',0),
 (151,'9311044','SOCIEDADES COOPERATIVAS',0),
 (152,'9900902','TRANSMISORES DE DINERO O DISPERSORES',0),
 (153,'9900903','CAMBISTAS O CENTROS CAMBIARIOS',0),
 (154,'9911018','INSTITUCIONES FINANCIERAS DEL EXTRANJERO',0),
 (155,'6999124','CREDITOS PARA ADQUISICION DE BIENES DE CONSUMO DURADERO',0),
 (156,'6999132','CREDITOS CONSUMOS PERSONALES',0),
 (157,'6999166','CREDITOS AUTOMOTRIZ',0),
 (158,'6999174','CREDITOS ADQUISICION DE BIENES MUEBLES',0),
 (159,'8219059','MONTEPIO',0),
 (160,'8219067','PRESTAMISTA',0),
 (161,'8219122','EMPRESAS DE AUTOFINANCIAMIENTO AUTOMOTRIZ',0),
 (162,'8219130','EMPRESAS DE AUTOFINANCIAMIENTO RESIDENCIAL',0),
 (163,'9900904','CASAS DE EMPEÑO',0),
 (164,'8219114','ADMINISTRADORAS DE TARJETA DE CREDITO',0),
 (165,'9900913','ADMINISTRADORAS DE TARJETA DE SERVICIOS',0),
 (166,'9505001','VENTA DE TARJETAS PREPAGADAS',0),
 (167,'9900914','ADMINISTRADORAS Y/O COMERCIALIZADORAS DE TARJETAS DE PREPAGO',0),
 (168,'9900915','COMERCIALIZADORA DE CHEQUES DE VIAJERO',0),
 (169,'8219083','EMPRESAS CONTROLADORAS FINANCIERAS',0),
 (170,'8300006','SERVICIOS RELACIONADOS CON INMUEBLES',0),
 (171,'8314015','ADMINISTRACION DE INMUEBLES',0),
 (172,'8400004','SERVICIOS PROFESIONALES Y TÉCNICOS',0),
 (173,'8412017','SERVICIOS DE BUFETES JURIDICOS',0),
 (174,'8413015','SERVICIOS DE CONTADURIA Y AUDITORIA; INCLUSO TENEDURIA DE LIBROS',0),
 (175,'8414013','SERVICIOS DE ASESORIA Y ESTUDIOS TECNICOS DE ARQUITECTURA E INGENIERIA (INCLUSO DISEÑO INDUSTRIAL)',0),
 (176,'8419013','SERVICIO DE INVESTIGACION DE MERCADO  SOLVENCIA FINANCIERA, DE PATENTES  Y MARCAS INDUSTRIALES Y OTROS SIMILARES',0),
 (177,'8424012','SERVICIOS ADMINISTRATIVOS DE TRAMITE Y COBRANZA; INCLUSO ESCRITORIOS PUBLICOS',0),
 (178,'8411019','SERVICIOS DE NOTARIAS PUBLICAS',0),
 (179,'9900925','SERVICIOS DE CORREDURÍAS PUBLICAS',0),
 (180,'9100009','SERVICIOS DE ENSEÑANZA, INVESTIGACIÓN CIENTÍFICA Y DIFUSIÓN CULTURAL',0),
 (181,'9200007','SERVICIOS MÉDICOS, DE ASISTENCIA SOCIAL Y VETERINARIOS',0),
 (182,'9221011','CENTRO DE BENEFICENCIA',0),
 (183,'9311010','ASOCIACIONES Y CONFEDERACIONES',0),
 (184,'9311028','CAMARAS DE COMERCIO',0),
 (185,'9311036','CAMARAS INDUSTRIALES',0),
 (186,'9312018','ORGANIZACIONES DE ABOGADOS MEDICOS INGENIEROS Y OTRAS ASOCIACIONES DE PROFESIONALES',0),
 (187,'9319014','ORGANIZACIONES CIVICAS',0),
 (188,'9321019','ORGANIZACIONES LABORALES Y SINDICALES',0),
 (189,'9322017','ORGANIZACIONES POLITICAS',0),
 (190,'9331018','ORGANIZACIONES RELIGIOSAS',0),
 (191,'9900926','OTRA ASOCIACIÓN CIVIL O SOCIEDAD CIVIL',0),
 (192,'9900927','OTRA INSTUTUCION DE ASISTENCIA PRIVADA, INSTITUCION DE BENEFICENCIA PRIVADA O ASOCIACIÓN DE ASISTENCIA PRIVADA',0),
 (193,'8600000','SERVICIOS DE ALOJAMIENTO TEMPORAL',0),
 (194,'8700008','PREPARACIÓN Y SERVICIO DE ALIMENTOS Y BEBIDAS',0),
 (195,'8711021','RESTAURANTE',0),
 (196,'8721012','BARES Y CANTINAS',0),
 (197,'8800006','SERVICIOS RECREATIVOS Y DE ESPARCIMIENTO',0),
 (198,'8829048','PROMOCION DE ESPECTACULOS DEPORTIVOS',0),
 (199,'8831019','CENTRO NOCTURNO',0),
 (200,'8833015','FEDERACIONES Y ASOCIACIONES DEPORTIVAS Y OTRAS CON FINES RECREATIVOS',0),
 (201,'8900004','SERVICIOS PERSONALES, PARA EL HOGAR Y DIVERSOS',0),
 (202,'9900923','SERVICIOS DE BLINDAJE DE VEHÍCULOS TERRESTRES Y/O INMUEBLES O PARTES DE ELLOS',0),
 (203,'8911019','TALLER DE REPARACION GENERAL DE AUTOMOVILES Y CAMIONES',0),
 (204,'8914013','SERVICIOS DE REPARACION DE CARROCERIAS PINTURA TAPICERIA HOJALATERIA Y CRISTALES DE AUTOMOVILES',0),
 (205,'8916019','ESTACIONAMIENTO PRIVADO PARA VEHICULOS',0),
 (206,'8916027','ESTACIONAMIENTO PUBLICO PARA VEHICULOS',0),
 (207,'8991011','QUEHACERES DEL HOGAR',0),
 (208,'9900905','ESTUDIANTE Ó MENOR DE EDAD SIN OCUPACIÓN',0),
 (209,'9900906','DESEMPLEADO',0),
 (210,'9900907','JUBILADO',0),
 (211,'9900908','PENSIONADO',0),
 (212,'9900909','AMA DE CASA',0),
 (213,'9900930','MINISTROS DE CULTO RELIGIOSO (SACERDOTE, PASTOR, MONJA, ETC)',0),
 (214,'9501009','EMPLEADO DEL SECTOR PRIVADO',0),
 (215,'9411018','GOBIERNO FEDERAL',0),
 (216,'9411026','GOBIERNO ESTATAL',0),
 (217,'9411034','GOBIERNO MUNICIPAL',0),
 (218,'9411998','EMPLEADO PUBLICO',0),
 (219,'9471012','PRESTACION DE SERVICIOS PUBLICOS Y SOCIALES',0),
 (220,'9900003','SERVICIOS DE ORGANIZACIONES INTERNACIONALES Y OTROS ORGANISMOS EXTRATERRITORIALES',0),
 (221,'9912016','CONSULADO',0),
 (222,'9912024','GOBIERNO EXTRANJERO',0),
 (223,'9800101','PRESIDENTE DE LA REPUBLICA',0),
 (224,'9800102','SECRETARIA DE ESTADO,  TITULAR DE ENTIDAD / PROCURADOR GENERAL DE LA REPÚBLICA',0),
 (225,'9800103','SUBSECRETARIA DE ESTADO',0),
 (226,'9800104','OFICIAL MAYOR',0),
 (227,'9800105','TITULAR DE ENTIDAD',0),
 (228,'9800106','JEFATURA DE UNIDAD',0),
 (229,'9800107','DIRECCIÓN GENERAL U HOMOLOGA',0),
 (230,'9800112','DIPUTADO  DEL H. CONGRESO DE LA UNIÓN',0),
 (231,'9800113','SENADOR DE LA REPUBLICA',0),
 (232,'9800114','SECRETARIA DE LA DEFENSA NACIONAL (MILITAR)',0),
 (233,'9800115','SECRETARIA DE MARINA (MARINO)',0),
 (234,'9800116','EMBAJADOR',0),
 (235,'9800117','CÓNSUL',0),
 (236,'9800200','EMPLEADO DE GOBIERNO DE ENTIDAD FEDERATIVA',0),
 (237,'9800201','GOBERNADOR CONSTITUCIONAL DEL ESTADO Ó JEFE DE GOBIERNO DEL DF',0),
 (238,'9800202','SECRETARIO DEL RAMO',0),
 (239,'9800203','PROCURADOR GENERAL DE JUSTICIA DEL ESTADO',0),
 (240,'9800204','DIPUTADO AL H. CONGRESO DEL ESTADO',0),
 (241,'9800205','SUB-PROCURADOR GENERAL DE JUSTICIA DEL ESTAD',0),
 (242,'9800206','CONTRALOR DEL H. CONGRESO DEL ESTADO',0),
 (243,'9800207','TESORERO DEL ESTADO',0),
 (244,'9800208','SUB-SECRETARIO',0),
 (245,'9800209','VOCAL EJECUTIVO',0),
 (246,'9800210','OFICIAL MAYOR DE ENTIDAD FEDERATIVA',0),
 (247,'9800211','DIRECTOR GENERAL',0),
 (248,'9800212','DIRECTOR EJECUTIVO',0),
 (249,'9800213','JEFE DE LA POLICÍA JUDICIAL DEL ESTADO',0),
 (250,'9800214','JEFE DE UNIDAD',0),
 (251,'9800300','EMPLEADO DE GOBIERNO MUNICIPAL O DELEGACIONES (D.F)',0),
 (252,'9800301','PRESIDENTE MUNICIPAL O DELEGADO',0),
 (253,'9800302','SUB-DELEGADO',0),
 (254,'9800303','REGIDOR',0),
 (255,'9800304','DIRECTOR GENERAL',0),
 (256,'9800400','EMPLEADO DEL PODER JUDICIAL',0),
 (257,'9800401','MINISTRO DE LA SUPREMA CORTE DE JUSTICIA DE LA NACIÓN',0),
 (258,'9800402','MAGISTRADO FEDERAL',0),
 (259,'9800403','JUEZ FEDERAL',0),
 (260,'9800404','EMPLEADO DEL PODER JUDICIAL FEDERAL (SECRETARIOS, ACTUARIOS, ETC)',0),
 (261,'9800405','MAGISTRADO DE ENTIDAD FEDERATIVA',0),
 (262,'9800406','JUEZ DE ENTIDAD FEDERATIVA',0),
 (263,'9800407','EMPLEADO DEL PODER JUDICIAL DE ENTIDAD FEDERATIVA (SECRETARIOS, ACTUARIOS, ETC)',0),
 (264,'9800408','JUEZ MUNICIPAL Ó DE DELEGACIÓN',0),
 (265,'9800409','EMPLEADO DEL PODER JUDICIAL DE MUNICIPIO Ó DELEGACIÓN (SECRETARIOS, ACTUARIOS, ETC)',0),
 (266,'9900900','DEPENDENCIAS DE GOBIERNO O EMPRESAS PARAESTATALES',0),
 (267,'9999999','NO APLICA',1);
/*!40000 ALTER TABLE `mld_actividades_economicas` ENABLE KEYS */;


--
-- Definition of table `mld_folio_avisos`
--

DROP TABLE IF EXISTS `mld_folio_avisos`;
CREATE TABLE `mld_folio_avisos` (
  `AnoAviso` int(4) DEFAULT NULL,
  `FolioAviso` int(14) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_folio_avisos`
--

/*!40000 ALTER TABLE `mld_folio_avisos` DISABLE KEYS */;
INSERT INTO `mld_folio_avisos` (`AnoAviso`,`FolioAviso`) VALUES 
 (2014,5);
/*!40000 ALTER TABLE `mld_folio_avisos` ENABLE KEYS */;


--
-- Definition of table `mld_giro_mercantil`
--

DROP TABLE IF EXISTS `mld_giro_mercantil`;
CREATE TABLE `mld_giro_mercantil` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(7) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=216 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_giro_mercantil`
--

/*!40000 ALTER TABLE `mld_giro_mercantil` DISABLE KEYS */;
INSERT INTO `mld_giro_mercantil` (`Id`,`Clave`,`Descripcion`) VALUES 
 (1,'0100008','AGRICULTURA'),
 (2,'0200006','GANADERÍA'),
 (3,'0300004','SILVICULTURA'),
 (4,'0400002','PESCA'),
 (5,'0500000','CAZA'),
 (6,'1100007','EXTRACCIÓN Y BENEFICIO DE CARBÓN MINERAL Y GRAFITO'),
 (7,'1200005','EXTRACCIÓN DE PETRÓLEO CRUDO Y GAS NATURAL'),
 (8,'1300003','EXTRACCIÓN Y BENEFICIO DE MINERALES METÁLICOS'),
 (9,'1311018','EXTRACCION Y BENEFICIO DE MINERAL DE HIERRO'),
 (10,'1322015','EXTRACCION Y BENEFICIO DE MERCURIO Y ANTIMONIO'),
 (11,'1329011','EXTRACCION Y BENEFICIO DE COBRE  PLOMO  ZINC Y OTROS MINERALES NO FERROSOS'),
 (12,'1400001','EXTRACCIÓN DE MINERALES NO METÁLICOS, EXCEPTO SAL'),
 (13,'1500009','EXPLOTACIÓN DE SAL'),
 (14,'2000008','FABRICACIÓN DE ALIMENTOS'),
 (15,'2012011','EMPACADORA DE CONSERVAS ALIMENTICIAS'),
 (16,'2012029','EMPACADORA DE FRUTAS Y LEGUMBRES'),
 (17,'2025014','BENEFICIO DE CAFE EXCEPTO MOLIENDA Y TOSTADO'),
 (18,'2049022','FABRICACION DE CARNES FRIAS Y EMBUTIDOS'),
 (19,'2100006','FABRICACIÓN Y ELABORACIÓN DE BEBIDAS (AGUA, REFRESCOS, CERVEZA, VINOS Y LICORES)'),
 (20,'2300002','INDUSTRIA TEXTIL (FABRICACIÓN DE: HILADOS Y TEJIDOS)'),
 (21,'2400000','FABRICACIÓN DE PRENDAS DE VESTIR Y OTROS ARTÍCULOS CONFECCIONADOS CON TEXTILES Y OTROS MATERIALES EXCEPTO CALZADO'),
 (22,'2500008','FABRICACIÓN DE CALZADO E INDUSTRIA DEL CUERO'),
 (23,'2600006','INDUSTRIA Y PRODUCTOS DE MADERA Y CORCHO; EXCEPTO MUEBLES'),
 (24,'2711019','FABRICACION DE MUEBLES DE MADERA'),
 (25,'2711027','FABRICACION DE MUEBLES DE MATERIAL SINTETICO'),
 (26,'2800002','INDUSTRIA DEL PAPEL'),
 (27,'2900000','INDUSTRIAS EDITORIAL, DE IMPRESIÓN Y CONEXAS'),
 (28,'3000007','INDUSTRIA QUÍMICA'),
 (29,'3021011','FABRICACION DE ABONOS Y FERTILIZANTES QUIMICOS'),
 (30,'3100005','REFINACIÓN DE PETRÓLEO Y DERIVADOS DEL CARBÓN MINERAL'),
 (31,'3111010','FABRICACION DE GASOLINA Y OTROS PRODUCTOS DERIVADOS DE LA REFINACION DE PETROLEO'),
 (32,'3112018','FABRICACION DE PRODUCTOS PETROQUIMICOS BASICOS'),
 (33,'3113016','FABRICACION DE ACEITES Y LUBRICANTES'),
 (34,'3200003','FABRICACIÓN DE PRODUCTOS DE HULE Y DE PLÁSTICO'),
 (35,'3300001','FABRICACIÓN DE PRODUCTOS DE MINERALES NO METÁLICOS; EXCEPTO DEL PETRÓLEO Y DEL CARBÓN MINERAL'),
 (36,'3322013','FABRICACION DE CRISTALES PARA AUTOMOVIL'),
 (37,'3331022','FABRICACION DE LADRILLOS'),
 (38,'3341013','FABRICACION DE CEMENTO'),
 (39,'3400009','INDUSTRIAS METÁLICAS BÁSICAS'),
 (40,'3411022','FUNDICION DE FIERRO Y ACERO'),
 (41,'3411030','PLANTA METALURGICA'),
 (42,'3412012','FABRICACION DE LAMINAS DE HIERRO Y ACERO'),
 (43,'3413010','FABRICACION DE TUBOS DE HIERRO Y ACERO'),
 (44,'3500007','FABRICACIÓN DE PRODUCTOS METÁLICOS; EXCEPTO MAQUINARIA Y EQUIPO'),
 (45,'3599026','FABRICACION DE CAJAS FUERTES'),
 (46,'3600005','FABRICACIÓN, ENSAMBLE Y REPARACIÓN DE MAQUINARIA, EQUIPO Y SUS PARTES; EXCEPTO LOS ELÉCTRICOS'),
 (47,'3700003','FABRICACIÓN Y ENSAMBLE DE MAQUINARIA, EQUIPO, APARATOS, ACCESORIOS Y ARTÍCULOS ELÉCTRICOS, ELECTRÓNICOS Y SUS PARTES'),
 (48,'3800001','CONSTRUCCIÓN, RECONSTRUCCIÓN Y ENSAMBLE DE EQUIPO DE TRANSPORTE Y SUS PARTES'),
 (49,'3819010','FABRICACION DE REFACCIONES Y ACCESORIOS AUTOMOTRICES'),
 (50,'3831014','FABRICACION Y REPARACION DE BUQUES Y BARCOS'),
 (51,'3832012','FABRICACION ENSAMBLE Y REPARACION DE AERONAVES'),
 (52,'3933018','FABRICACION DE ARTICULOS DE QUINCALLERIA Y BISUTERIA'),
 (53,'3997014','FABRICACION DE ARMAS'),
 (54,'4100004','CONTRATACIÓN DE OBRAS COMPLETAS DE CONSTRUCCIÓN (CASAS, DEPARTAMENTOS, INMUEBLES, PAVIMENTACIÓN, NO RESIDENCIALES, VIAS DE COMUNICACIÓN)'),
 (55,'4111019','CONSTRUCCION DE CASAS Y TECHOS DESARMABLES'),
 (56,'4111027','CONSTRUCCION DE INMUEBLES'),
 (57,'4112017','CONSTRUCCION DE EDIFICIOS PARA OFICINAS ESCUELAS HOSPITALES HOTELES Y OTROS NO RESIDENCIALES'),
 (58,'4113015','CONSTRUCCION DE EDIFICIOS INDUSTRIALES Y PARA FINES ANALOGOS'),
 (59,'4121018','CONSTRUCCION DE VIAS DE COMUNICACION'),
 (60,'4199015','CONSTRUCCION DE ESTADIOS MONUMENTOS Y OTRAS OBRAS DE INGENIERIA'),
 (61,'5012018','DISTRIBUCION DE ENERGIA ELECTRICA'),
 (62,'6100002','COMPRAVENTA DE ALIMENTOS, BEBIDAS Y PRODUCTOS DE TABACO'),
 (63,'6121024','COMPRAVENTA DE GANADO MAYOR EN PIE'),
 (64,'6121032','COMPRAVENTA DE GANADO MENOR EN PIE'),
 (65,'6131023','TIENDA DE ABARROTES Y MISCELANEA'),
 (66,'6200000','COMPRAVENTA DE PRENDAS DE VESTIR Y OTROS ARTÍCULOS DE USO PERSONAL'),
 (67,'1321017','EXTRACCION Y BENEFICIO DE ORO PLATA Y OTROS METALES PRECIOSOS'),
 (68,'3921013','FABRICACION DE RELOJES'),
 (69,'3932010','FABRICACION DE ARTICULOS DE JOYERIA'),
 (70,'3932036','TALLADO DE PIEDRAS PRECIOSAS'),
 (71,'6225016','COMPRAVENTA DE ARTICULOS DE PLATA'),
 (72,'6225024','COMPRAVENTA DE JOYAS'),
 (73,'6225032','COMPRAVENTA DE RELOJES'),
 (74,'6999017','COMPRAVENTA DE DIAMANTES'),
 (75,'9900916','COMPRAVENTA DE ARTICULOS DE ORO'),
 (76,'9900917','COMPRAVENTA DE ARTICULOS DE PLATINO'),
 (77,'9900918','COMPRAVENTA DE AGUAMARINAS,  ESMERALDAS, RUBÍES, TOPACIOS, TURQUESAS Y/O ZAFIROS'),
 (78,'9900919','COMPRAVENTA DE PLATA, ORO O PLATINO A GRANEL'),
 (79,'6325014','COMPRAVENTA DE ANTIGÜEDADES'),
 (80,'8832017','GALERIAS DE ARTES GRAFICAS Y MUSEOS'),
 (81,'9900920','COMPRAVENTA DE OBRAS DE ARTE'),
 (82,'9900921','CASA DE SUBASTAS DE OBRAS DE ARTE, JOYAS Y/O ANTIGÜEDADES'),
 (83,'6300008','COMPRAVENTA DE ARTÍCULOS PARA EL HOGAR (ELECTRODOMESTICOS, REFACCIONES, LOZA Y PORCELANA, ANTIGUEDADES)'),
 (84,'6400006','COMPRAVENTA EN TIENDAS DE AUTOSERVICIO Y DE DEPARTAMENTOS ESPECIALIZADOS POR LÍNEA DE MERCANCÍAS'),
 (85,'6500004','COMPRAVENTA DE GASES, COMBUSTIBLES Y LUBRICANTES'),
 (86,'6513015','COMPRAVENTA DE GASOLINA Y DIESEL'),
 (87,'6514013','COMPRAVENTA DE PETROLEO COMBUSTIBLE'),
 (88,'6515011','COMPRAVENTA DE LUBRICANTES'),
 (89,'6600002','COMPRAVENTA DE MATERIAS PRIMAS, MATERIALES Y AUXILIARES (ALGODÓN, CEMENTO, SANITARIOS, PIELES, FERRETERIA, MADERA, PINTURAS)'),
 (90,'6691019','COMPRAVENTA DE FERTILIZANTES Y PLAGUICIDAS'),
 (91,'6695011','COMPRAVENTA DE SUBSTANCIAS QUIMICAS PARA LA INDUSTRIA'),
 (92,'6700000','COMPRAVENTA DE MAQUINARIA, EQUIPO, INSTRUMENTOS, APARATOS Y HERRAMIENTAS, SUS REFACCIONES Y ACCESORIOS'),
 (93,'6712013','COMPRAVENTA DE ARTICULOS PARA LA EXPLOTACION DE MINAS'),
 (94,'6811013','COMPRAVENTA DE AUTOMOVILES Y CAMIONES NUEVOS'),
 (95,'6812011','COMPRAVENTA DE AUTOMOVILES Y CAMIONES USADOS'),
 (96,'9900922','COMPRAVENTA DE VEHICULOS MARÍTIMOS'),
 (97,'6819033','COMPRAVENTA DE VEHICULOS AEREOS'),
 (98,'6813027','COMPRAVENTA DE MOTOCICLETAS Y SUS ACCESORIOS'),
 (99,'6819017','COMPRAVENTA DE PARTES Y REFACCIONES PARA VEHICULOS TERRESTRES, AÉREOS Y MARÍTIMOS'),
 (100,'4111051','DESARROLLADORES DE VIVIENDA'),
 (101,'6900006','COMPRAVENTA DE BIENES INMUEBLES Y ARTÍCULOS DIVERSOS'),
 (102,'6911053','COMPRAVENTA DE TERRENOS'),
 (103,'8313017','SERVICIO DE CORREDORES DE BIENES RAICES'),
 (104,'6991013','COMPRAVENTA DE ARMAS DE FUEGO'),
 (105,'6992011','AGENCIAS DE RIFAS Y SORTEOS (QUINIELAS Y LOTERIA)'),
 (106,'8829022','HIPODROMO'),
 (107,'9900910','SALAS DE JUEGOS Y APUESTAS'),
 (108,'9900911','ORGANIZACIÓN DE FERIAS REGIONALES CON APUESTAS'),
 (109,'9900912','ORGANIZACIÓN DE CARRERAS DE CABALLOS O PELEAS DE GALLOS EN ESCENARIOS TEMPORALES'),
 (110,'7100001','TRANSPORTE TERRESTRE'),
 (111,'7200009','TRANSPORTE POR AGUA'),
 (112,'7300007','TRANSPORTE AÉREO'),
 (113,'7312010','SERVICIOS RELACIONADOS CON EL TRANSPORTE EN AERONAVES CON MATRICULA EXTRANJERA'),
 (114,'7400005','SERVICIOS CONEXOS AL TRANSPORTE'),
 (115,'8429038','EMPRESAS DE SEGURIDAD PRIVADA'),
 (116,'8429046','EMPRESAS TRANSPORTADORAS DE VALORES'),
 (117,'9900924','EMPRESAS DE CUSTODIA DE VALORES'),
 (118,'7512016','AGENCIA DE TURISMO'),
 (119,'7513014','AGENCIA ADUANAL'),
 (120,'8524010','ALQUILER O RENTA DE AUTOMOVILES SIN CHOFER'),
 (121,'7519020','ALQUILER DE LANCHAS Y VELEROS'),
 (122,'7519038','RENTA DE VEHICULOS AEREOS'),
 (123,'8311011','ALQUILER DE TERRENOS LOCALES Y EDIFICIOS NO RESIDENCIALES'),
 (124,'8312019','ARRENDAMIENTO DE INMUEBLES RESIDENCIALES'),
 (125,'7600001','COMUNICACIONES'),
 (126,'8114019','SERVICIOS DE FONDOS Y FIDEICOMISOS DE FOMENTO ECONOMICO'),
 (127,'8123010','INSTITUCIONES DE BANCA MÚLTIPLE'),
 (128,'9900929','INSTITUCIONES DE LA BANCA DE DESARROLLO'),
 (129,'8123052','SOCIEDADES DE AHORRO Y PRESTAMO'),
 (130,'8123060','SOCIEDADES DE AHORRO Y CREDITO POPULAR'),
 (131,'8123078','SOCIEDADES FINANCIERAS DE OBJETO LIMITADO'),
 (132,'8123086','SOCIEDADES FINANCIERAS DE OBJETO MULTIPLE REGULADAS'),
 (133,'8123094','SOCIEDADES FINANCIERAS DE OBJETO MULTIPLE NO REGULADAS'),
 (134,'8131021','ALMACENES DE DEPOSITO'),
 (135,'8132029','UNIONES DE CREDITO'),
 (136,'8133027','COMPAÑIAS DE FIANZAS'),
 (137,'8142010','SOCIEDADES DE INVERSION'),
 (138,'8151029','COMPAÑIAS DE SEGUROS PRIVADAS'),
 (139,'8200008','SERVICIOS COLATERALES A INSTITUCIONES FINANCIERAS Y DE SEGUROS'),
 (140,'8211013','INVERSIONISTA'),
 (141,'8211021','AGENTE DE BOLSA'),
 (142,'8211047','CASAS DE BOLSA'),
 (143,'8219017','AGENTE DE SEGUROS'),
 (144,'8219025','CASA DE CAMBIO'),
 (145,'6999992','CENTROS CAMBIARIOS'),
 (146,'8219033','CORRESPONSAL BANCARIO'),
 (147,'8219041','CAJA DE AHORROS'),
 (148,'8219075','FACTORING'),
 (149,'8511033','ARRENDADORAS FINANCIERAS'),
 (150,'9311044','SOCIEDADES COOPERATIVAS'),
 (151,'9900902','TRANSMISORES DE DINERO O DISPERSORES'),
 (152,'9900903','CAMBISTAS O CENTROS CAMBIARIOS'),
 (153,'9911018','INSTITUCIONES FINANCIERAS DEL EXTRANJERO'),
 (154,'6999124','CREDITOS PARA ADQUISICION DE BIENES DE CONSUMO DURADERO'),
 (155,'6999132','CREDITOS CONSUMOS PERSONALES'),
 (156,'6999166','CREDITOS AUTOMOTRIZ'),
 (157,'6999174','CREDITOS ADQUISICION DE BIENES MUEBLES'),
 (158,'8219059','MONTEPIO'),
 (159,'8219067','PRESTAMISTA'),
 (160,'8219122','EMPRESAS DE AUTOFINANCIAMIENTO AUTOMOTRIZ'),
 (161,'8219130','EMPRESAS DE AUTOFINANCIAMIENTO RESIDENCIAL'),
 (162,'9900904','CASAS DE EMPEÑO'),
 (163,'8219114','ADMINISTRADORAS DE TARJETA DE CREDITO'),
 (164,'9900913','ADMINISTRADORAS DE TARJETA DE SERVICIOS'),
 (165,'9505001','VENTA DE TARJETAS PREPAGADAS'),
 (166,'9900914','ADMINISTRADORAS Y/O COMERCIALIZADORAS DE TARJETAS DE PREPAGO'),
 (167,'9900915','COMERCIALIZADORA DE CHEQUES DE VIAJERO'),
 (168,'8219083','EMPRESAS CONTROLADORAS FINANCIERAS'),
 (169,'8300006','SERVICIOS RELACIONADOS CON INMUEBLES'),
 (170,'8314015','ADMINISTRACION DE INMUEBLES'),
 (171,'8400004','SERVICIOS PROFESIONALES Y TÉCNICOS'),
 (172,'8412017','SERVICIOS DE BUFETES JURIDICOS'),
 (173,'8413015','SERVICIOS DE CONTADURIA Y AUDITORIA; INCLUSO TENEDURIA DE LIBROS'),
 (174,'8414013','SERVICIOS DE ASESORIA Y ESTUDIOS TECNICOS DE ARQUITECTURA E INGENIERIA (INCLUSO DISEÑO INDUSTRIAL)'),
 (175,'8419013','SERVICIO DE INVESTIGACION DE MERCADO  SOLVENCIA FINANCIERA, DE PATENTES  Y MARCAS INDUSTRIALES Y OTROS SIMILARES'),
 (176,'8424012','SERVICIOS ADMINISTRATIVOS DE TRAMITE Y COBRANZA; INCLUSO ESCRITORIOS PUBLICOS'),
 (177,'8411019','SERVICIOS DE NOTARIAS PUBLICAS'),
 (178,'9900925','SERVICIOS DE CORREDURÍAS PUBLICAS'),
 (179,'9100009','SERVICIOS DE ENSEÑANZA, INVESTIGACIÓN CIENTÍFICA Y DIFUSIÓN CULTURAL'),
 (180,'9200007','SERVICIOS MÉDICOS, DE ASISTENCIA SOCIAL Y VETERINARIOS'),
 (181,'9221011','CENTRO DE BENEFICENCIA'),
 (182,'9311010','ASOCIACIONES Y CONFEDERACIONES'),
 (183,'9311028','CAMARAS DE COMERCIO'),
 (184,'9311036','CAMARAS INDUSTRIALES'),
 (185,'9312018','ORGANIZACIONES DE ABOGADOS MEDICOS INGENIEROS Y OTRAS ASOCIACIONES DE PROFESIONALES'),
 (186,'9319014','ORGANIZACIONES CIVICAS'),
 (187,'9321019','ORGANIZACIONES LABORALES Y SINDICALES'),
 (188,'9322017','ORGANIZACIONES POLITICAS'),
 (189,'9331018','ORGANIZACIONES RELIGIOSAS'),
 (190,'9900926','OTRA ASOCIACIÓN CIVIL O SOCIEDAD CIVIL'),
 (191,'9900927','OTRA INSTUTUCION DE ASISTENCIA PRIVADA, INSTITUCION DE BENEFICENCIA PRIVADA O ASOCIACIÓN DE ASISTENCIA PRIVADA'),
 (192,'8600000','SERVICIOS DE ALOJAMIENTO TEMPORAL'),
 (193,'8700008','PREPARACIÓN Y SERVICIO DE ALIMENTOS Y BEBIDAS'),
 (194,'8711021','RESTAURANTE'),
 (195,'8721012','BARES Y CANTINAS'),
 (196,'8800006','SERVICIOS RECREATIVOS Y DE ESPARCIMIENTO'),
 (197,'8829048','PROMOCION DE ESPECTACULOS DEPORTIVOS'),
 (198,'8831019','CENTRO NOCTURNO'),
 (199,'8833015','FEDERACIONES Y ASOCIACIONES DEPORTIVAS Y OTRAS CON FINES RECREATIVOS'),
 (200,'8900004','SERVICIOS PERSONALES, PARA EL HOGAR Y DIVERSOS'),
 (201,'9900923','SERVICIOS DE BLINDAJE DE VEHÍCULOS TERRESTRES Y/O INMUEBLES O PARTES DE ELLOS'),
 (202,'8911019','TALLER DE REPARACION GENERAL DE AUTOMOVILES Y CAMIONES'),
 (203,'8914013','SERVICIOS DE REPARACION DE CARROCERIAS PINTURA TAPICERIA HOJALATERIA Y CRISTALES DE AUTOMOVILES'),
 (204,'8916019','ESTACIONAMIENTO PRIVADO PARA VEHICULOS'),
 (205,'8916027','ESTACIONAMIENTO PUBLICO PARA VEHICULOS'),
 (206,'8991011','QUEHACERES DEL HOGAR'),
 (207,'9411018','GOBIERNO FEDERAL'),
 (208,'9411026','GOBIERNO ESTATAL'),
 (209,'9411034','GOBIERNO MUNICIPAL'),
 (210,'9471012','PRESTACION DE SERVICIOS PUBLICOS Y SOCIALES'),
 (211,'9900003','SERVICIOS DE ORGANIZACIONES INTERNACIONALES Y OTROS ORGANISMOS EXTRATERRITORIALES'),
 (212,'9912016','CONSULADO'),
 (213,'9912024','GOBIERNO EXTRANJERO'),
 (214,'9900900','DEPENDENCIAS DE GOBIERNO O EMPRESAS PARAESTATALES'),
 (215,'9999999','NO APLICA');
/*!40000 ALTER TABLE `mld_giro_mercantil` ENABLE KEYS */;


--
-- Definition of table `mld_inmuebles_figura_cliente`
--

DROP TABLE IF EXISTS `mld_inmuebles_figura_cliente`;
CREATE TABLE `mld_inmuebles_figura_cliente` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_inmuebles_figura_cliente`
--

/*!40000 ALTER TABLE `mld_inmuebles_figura_cliente` DISABLE KEYS */;
INSERT INTO `mld_inmuebles_figura_cliente` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'1','Vendedor',0),
 (2,'2','Comprador',1);
/*!40000 ALTER TABLE `mld_inmuebles_figura_cliente` ENABLE KEYS */;


--
-- Definition of table `mld_inmuebles_figura_empresa`
--

DROP TABLE IF EXISTS `mld_inmuebles_figura_empresa`;
CREATE TABLE `mld_inmuebles_figura_empresa` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_inmuebles_figura_empresa`
--

/*!40000 ALTER TABLE `mld_inmuebles_figura_empresa` DISABLE KEYS */;
INSERT INTO `mld_inmuebles_figura_empresa` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'1','Vendedor',1),
 (2,'2','Comprador',0),
 (3,'3','Intermediario',0);
/*!40000 ALTER TABLE `mld_inmuebles_figura_empresa` ENABLE KEYS */;


--
-- Definition of table `mld_inmuebles_tipo_alertas`
--

DROP TABLE IF EXISTS `mld_inmuebles_tipo_alertas`;
CREATE TABLE `mld_inmuebles_tipo_alertas` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  `ReqDesc` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=25 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_inmuebles_tipo_alertas`
--

/*!40000 ALTER TABLE `mld_inmuebles_tipo_alertas` DISABLE KEYS */;
INSERT INTO `mld_inmuebles_tipo_alertas` (`Id`,`Clave`,`Descripcion`,`RegDefault`,`ReqDesc`) VALUES 
 (1,'100  ','Sin alerta',1,0),
 (2,'501  ','El cliente o usuario se rehúsa a proporcionar documentos personales que lo identifiquen',0,0),
 (3,'502  ','El pago por el inmueble es realizado por un tercero sin relación aparente con el cliente o usuario',0,0),
 (4,'503  ','El cliente o usuario o personas relacionadas con él realizan múltiples operaciones en un periodo muy corto sin razón aparente',0,0),
 (5,'504  ','El cliente o usuario no muestra tener interés en las características de la propiedad objeto de la operación o en el precio y condiciones de la transacción',0,0),
 (6,'505  ','Se conoce un historial criminal del cliente o usuario, de algún familiar directo o persona relacionada con él',0,0),
 (7,'506  ','El cliente o usuario no quiere ser relacionado con la compra del inmueble',0,0),
 (8,'507  ','De acuerdo con la ocupación del cliente o usuario, la operación parece estar fuera de su alcance',0,0),
 (9,'508  ','De acuerdo con los ingresos declarados por el cliente o usuario, la operación parece estar fuera de su alcance',0,0),
 (10,'509  ','El cliente o usuario muestra fuerte interés en la realización de la transacción con rapidez, sin que exista causa justificada',0,0),
 (11,'510  ','El cliente o usuario pide que el pago sea dividido en partes con un breve intervalo de tiempo entre ellos',0,0),
 (12,'511  ','El cliente o usuario solicita que se realice la operación por medio de un contrato privado, donde no hay intención de registrarlo ante notario, o cuando esta intención se expresa, no se lleva a cabo finalmente',0,0),
 (13,'512  ','Transacciones sucesivas de compra y venta de la misma propiedad en un periodo corto de tiempo, con cambios injustificados del valor de la misma',0,0),
 (14,'513  ','La operación se lleva acabo a un valor de venta o compra significativamente diferente (mucho mayor o mucho menor) a partir del valor real de la propiedad o a los valores de mercado',0,0),
 (15,'514  ','Hay indicios, o certeza, que las partes no están actuando en nombre propio y están tratando de ocultar la identidad del cliente o usuario real',0,0),
 (16,'515  ','Uso de divisas en efectivo en montos elevados o de poco uso sin que la ocupación del cliente o usuario lo justifique',0,0),
 (17,'516  ','La operación se liquida por medio de una transferencia internacional sin que la ocupación, perfil o nacionalidad del cliente o usuario lo justifique',0,0),
 (18,'517  ','La operación se liquida por medio de una transferencia internacional proveniente de un país considerado como paraíso fiscal o de alto riesgo',0,0),
 (19,'518  ','El cliente o usuario insiste en liquidar pagar la operación en efectivo rebasando el umbral permitido para uso de efectivo',0,0),
 (20,'519  ','El cliente o usuario intenta sobornar o extorsionar al vendedor con  el fin de realizar la operación de forma irregular',0,0),
 (21,'520  ','El cliente o usuario solicita condiciones especiales poco usuales en la realización de la operación',0,0),
 (22,'521  ','El cliente o usuario proporcionó datos falsos  o documentos apócrifos al realizar la operación',0,0),
 (23,'522  ','Operaciones con organizaciones sin fines de lucro, cuando las características de la transacción no coinciden con los objetivos de la entidad',0,0),
 (24,'9999 ','Otra alerta',0,1);
/*!40000 ALTER TABLE `mld_inmuebles_tipo_alertas` ENABLE KEYS */;


--
-- Definition of table `mld_inmuebles_tipo_blindaje`
--

DROP TABLE IF EXISTS `mld_inmuebles_tipo_blindaje`;
CREATE TABLE `mld_inmuebles_tipo_blindaje` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_inmuebles_tipo_blindaje`
--

/*!40000 ALTER TABLE `mld_inmuebles_tipo_blindaje` DISABLE KEYS */;
INSERT INTO `mld_inmuebles_tipo_blindaje` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'1','SI esta total o parcialmente blindado',0),
 (2,'2','NO esta blindado',1);
/*!40000 ALTER TABLE `mld_inmuebles_tipo_blindaje` ENABLE KEYS */;


--
-- Definition of table `mld_inmuebles_tipo_operacion`
--

DROP TABLE IF EXISTS `mld_inmuebles_tipo_operacion`;
CREATE TABLE `mld_inmuebles_tipo_operacion` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_inmuebles_tipo_operacion`
--

/*!40000 ALTER TABLE `mld_inmuebles_tipo_operacion` DISABLE KEYS */;
INSERT INTO `mld_inmuebles_tipo_operacion` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'501','Compra Venta de Inmuebles',1);
/*!40000 ALTER TABLE `mld_inmuebles_tipo_operacion` ENABLE KEYS */;


--
-- Definition of table `mld_instr_monetarios`
--

DROP TABLE IF EXISTS `mld_instr_monetarios`;
CREATE TABLE `mld_instr_monetarios` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `Estatus` tinyint(2) DEFAULT '0',
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=15 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_instr_monetarios`
--

/*!40000 ALTER TABLE `mld_instr_monetarios` DISABLE KEYS */;
INSERT INTO `mld_instr_monetarios` (`Id`,`Clave`,`Descripcion`,`Estatus`,`RegDefault`) VALUES 
 (1,'  1  ','Efectivo',1,1),
 (2,'  2  ','Tarjeta de Crédito',1,0),
 (3,'  3  ','Tarjeta de Debito',1,0),
 (4,'  4  ','Tarjeta de Prepago',0,0),
 (5,'  5  ','Cheque Nominativo',0,0),
 (6,'  6  ','Cheque de Caja',0,0),
 (7,'  7  ','Cheques de Viajero',0,0),
 (8,'  8  ','Transferencia Interbancaria',0,0),
 (9,'  9  ','Transferencia Misma Institución',0,0),
 (10,'  10 ','Transferencia Internacional',0,0),
 (11,'  11 ','Orden de Pago',0,0),
 (12,'  12 ','Giro',0,0),
 (13,'  13 ','Oro o Platino Amonedados',0,0),
 (14,'  14 ','Plata Amonedada',0,0);
/*!40000 ALTER TABLE `mld_instr_monetarios` ENABLE KEYS */;


--
-- Definition of table `mld_metales_tipo_alertas`
--

DROP TABLE IF EXISTS `mld_metales_tipo_alertas`;
CREATE TABLE `mld_metales_tipo_alertas` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  `ReqDesc` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=29 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_metales_tipo_alertas`
--

/*!40000 ALTER TABLE `mld_metales_tipo_alertas` DISABLE KEYS */;
INSERT INTO `mld_metales_tipo_alertas` (`Id`,`Clave`,`Descripcion`,`RegDefault`,`ReqDesc`) VALUES 
 (1,'100 ','Sin alerta',1,0),
 (2,'601 ','El cliente o usuario se rehúsa a proporcionar documentos personales que lo identifiquen',0,0),
 (3,'602 ','Se conoce un historial criminal del cliente o usuario, de algún familiar directo o persona relacionada con él',0,0),
 (4,'603 ','El cliente o usuario compra o vende grandes cantidades de metales preciosos, piedras preciosas, joyas y/o relojes sin justificar su procedencia',0,0),
 (5,'604 ','El cliente o usuario o personas relacionadas con él  realizan diversas operaciones de compra o venta de grandes cantidades de metales preciosas, piedras preciosas, joyas y/o',0,0),
 (6,'605 ','El cliente o usuario realiza compras indiscriminadas de mercancía (sin importar tamaño, color, precio) de metales preciosos, piedras preciosas, joyas y/o relojes, sin que es',0,0),
 (7,'606 ','El cliente o usuario no quiere ser relacionado con la operación realizada',0,0),
 (8,'607 ','De acuerdo con la ocupación del cliente o usuario, la operación parece estar fuera de su alcance',0,0),
 (9,'608 ','De acuerdo con los ingresos declarados por el cliente o usuario, la operación parece estar fuera de su alcance',0,0),
 (10,'609 ','El cliente o usuario muestra fuerte interés en la realización de la transacción con rapidez, sin que exista causa justificada',0,0),
 (11,'610 ','Hay indicios, o certeza, que las partes no están actuando en nombre propio y están tratando de ocultar la identidad del cliente o usuario real',0,0),
 (12,'611 ','Uso de divisas en efectivo en montos elevados o de poco uso sin que la ocupación del cliente o usuario lo justifique',0,0),
 (13,'612 ','La operación se liquida por medio de una transferencia internacional sin que la ocupación, perfil o nacionalidad del cliente o usuario lo justifique',0,0),
 (14,'613 ','La operación se liquida por medio de una transferencia internacional proveniente de un país considerado como paraíso fiscal o de alto riesgo',0,0),
 (15,'614 ','El cliente o usuario insiste en liquidar pagar la operación en efectivo rebasando el umbral permitido para uso de efectivo',0,0),
 (16,'615 ','El cliente o usuario intenta sobornar o extorsionar al vendedor con  el fin de realizar la operación de forma irregular',0,0),
 (17,'616 ','El cliente o usuario solicita condiciones especiales poco usuales en la realización de la operación',0,0),
 (18,'617 ','El cliente o usuario proporcionó datos falsos  o documentos apócrifos al realizar la operación',0,0),
 (19,'618 ','Operaciones con organizaciones sin fines de lucro, cuando las características de la transacción no coinciden con los objetivos de la entidad',0,0),
 (20,'619 ','De lo que se conoce del cliente o usuario se sabe que recolecta metales preciosos (ej. oro) y después lo funde con la intención de venderlo con una calidad diferente a la de',0,0),
 (21,'620 ','Exporta e importa joyería, relojes, metales y piedras preciosas de países riesgosos sin que haya justificación aparente',0,0),
 (22,'621 ','El cliente o usuario vende pedacería de metales preciosos que podría provenir de actividades ilícitas como el robo',0,0),
 (23,'622 ','El cliente o usuario compra pedacería de metales preciosos por cantidades bajas a precios que exceden los de mercado o viceversa',0,0),
 (24,'623 ','El cliente o usuario liquida la mayoría de sus operaciones con otras divisas, sin que su actividad lo justifique',0,0),
 (25,'624 ','La operación se liquida a través de cheques de caja con la intención de ocultar el origen de los recursos',0,0),
 (26,'625 ','Para liquidar sus operaciones el cliente o usuario utiliza recursos de diferentes cuentas, instituciones financieras, denominaciones o con diversos instrumentos monetarios,',0,0),
 (27,'626 ','El cliente o usuario es socio de empresas que se dedican a la compra venta de metales y piedras preciosas que operan en estados riesgosos',0,0),
 (28,' 9999','Otra alerta',0,1);
/*!40000 ALTER TABLE `mld_metales_tipo_alertas` ENABLE KEYS */;


--
-- Definition of table `mld_metales_tipo_bienes`
--

DROP TABLE IF EXISTS `mld_metales_tipo_bienes`;
CREATE TABLE `mld_metales_tipo_bienes` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `Estatus` tinyint(2) DEFAULT '0',
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=13 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_metales_tipo_bienes`
--

/*!40000 ALTER TABLE `mld_metales_tipo_bienes` DISABLE KEYS */;
INSERT INTO `mld_metales_tipo_bienes` (`Id`,`Clave`,`Descripcion`,`Estatus`,`RegDefault`) VALUES 
 (1,'1','Metales Preciosos - Oro',1,1),
 (2,'2','Metales Preciosos - Plata',1,0),
 (3,'3','Metales Preciosos - Platino',1,0),
 (4,'4','Piedras Preciosas - Aguamarinas',0,0),
 (5,'5','Piedras Preciosas - Diamantes',0,0),
 (6,'6','Piedras Preciosas - Esmeraldas',0,0),
 (7,'7','Piedras Preciosas - Rubíes',0,0),
 (8,'8','Piedras Preciosas - Topacios',0,0),
 (9,'9','Piedras Preciosas - Turquesas',0,0),
 (10,'10','Piedras Preciosas - Zafiros',0,0),
 (11,'11','Joyas',1,0),
 (12,'12','Relojes',1,0);
/*!40000 ALTER TABLE `mld_metales_tipo_bienes` ENABLE KEYS */;


--
-- Definition of table `mld_metales_tipo_operacion`
--

DROP TABLE IF EXISTS `mld_metales_tipo_operacion`;
CREATE TABLE `mld_metales_tipo_operacion` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_metales_tipo_operacion`
--

/*!40000 ALTER TABLE `mld_metales_tipo_operacion` DISABLE KEYS */;
INSERT INTO `mld_metales_tipo_operacion` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'601','Venta',1),
 (2,'602','Compra',0);
/*!40000 ALTER TABLE `mld_metales_tipo_operacion` ENABLE KEYS */;


--
-- Definition of table `mld_metales_tipo_unidades`
--

DROP TABLE IF EXISTS `mld_metales_tipo_unidades`;
CREATE TABLE `mld_metales_tipo_unidades` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `Estatus` tinyint(2) DEFAULT '0',
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_metales_tipo_unidades`
--

/*!40000 ALTER TABLE `mld_metales_tipo_unidades` DISABLE KEYS */;
INSERT INTO `mld_metales_tipo_unidades` (`Id`,`Clave`,`Descripcion`,`Estatus`,`RegDefault`) VALUES 
 (1,'1','Pieza',1,0),
 (2,'2','Gramos',1,1),
 (3,'3','Kilates',1,0);
/*!40000 ALTER TABLE `mld_metales_tipo_unidades` ENABLE KEYS */;


--
-- Definition of table `mld_paises`
--

DROP TABLE IF EXISTS `mld_paises`;
CREATE TABLE `mld_paises` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=245 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_paises`
--

/*!40000 ALTER TABLE `mld_paises` DISABLE KEYS */;
INSERT INTO `mld_paises` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'  AF ','AFGANISTAN',0),
 (2,'  AL ','ALBANIA',0),
 (3,'  DE ','ALEMANIA',0),
 (4,'  AD ','ANDORRA',0),
 (5,'  AO ','ANGOLA',0),
 (6,'  AI ','ANGUILA',0),
 (7,'  AQ ','ANTARTIDA',0),
 (8,'  AG ','ANTIGUA Y BARBUDA',0),
 (9,'  AN ','ANTILLAS NEERLANDESAS',0),
 (10,'  SA ','ARABIA SAUDI',0),
 (11,'  DZ ','ARGELIA',0),
 (12,'  AR ','ARGENTINA',0),
 (13,'  AM ','ARMENIA',0),
 (14,'  AW ','ARUBA',0),
 (15,'  AU ','AUSTRALIA',0),
 (16,'  AT ','AUSTRIA',0),
 (17,'  AZ ','AZERBAIYAN',0),
 (18,'  BS ','BAHAMAS',0),
 (19,'  BH ','BAHREIN',0),
 (20,'  BD ','BANGLADESH',0),
 (21,'  BB ','BARBADOS',0),
 (22,'  BY ','BELARUS',0),
 (23,'  BE ','BELGICA',0),
 (24,'  BZ ','BELICE',0),
 (25,'  BJ ','BENIN',0),
 (26,'  BM ','BERMUDAS',0),
 (27,'  BT ','BHUTAN',0),
 (28,'  BO ','BOLIVIA',0),
 (29,'  BA ','BOSNIA Y HERZEGOVINA',0),
 (30,'  BW ','BOTSUANA',0),
 (31,'  BR ','BRASIL',0),
 (32,'  BN ','BRUNEI',0),
 (33,'  BG ','BULGARIA',0),
 (34,'  BF ','BURKINA FASO',0),
 (35,'  BI ','BURUNDI',0),
 (36,'  CV ','CABO VERDE',0),
 (37,'  KH ','CAMBOYA',0),
 (38,'  CM ','CAMERUN',0),
 (39,'  CA ','CANADA',0),
 (40,'  TD ','CHAD',0),
 (41,'  CZ ','CHEQUIA',0),
 (42,'  CL ','CHILE',0),
 (43,'  CN ','CHINA',0),
 (44,'  CY ','CHIPRE',0),
 (45,'  CP ','CLIPPERTON',0),
 (46,'  CO ','COLOMBIA',0),
 (47,'  KM ','COMORAS',0),
 (48,'  CG ','CONGO',0),
 (49,'  KP ','COREA DEL NORTE',0),
 (50,'  KR ','COREA DEL SUR',0),
 (51,'  CI ','COSTA DE MARFIL',0),
 (52,'  CR ','COSTA RICA',0),
 (53,'  HR ','CROACIA',0),
 (54,'  CU ','CUBA',0),
 (55,'  DK ','DINAMARCA',0),
 (56,'  DM ','DOMINICA',0),
 (57,'  EC ','ECUADOR',0),
 (58,'  EG ','EGIPTO',0),
 (59,'  SV ','EL SALVADOR',0),
 (60,'  AE ','EMIRATOS ARABES UNIDOS',0),
 (61,'  ER ','ERITREA',0),
 (62,'  SK ','ESLOVAQUIA',0),
 (63,'  SI ','ESLOVENIA',0),
 (64,'  ES ','ESPAÑA',0),
 (65,'  US ','ESTADOS UNIDOS',0),
 (66,'  EE ','ESTONIA',0),
 (67,'  ET ','ETIOPIA',0),
 (68,'  PH ','FILIPINAS',0),
 (69,'  FI ','FINLANDIA',0),
 (70,'  FJ ','FIYI',0),
 (71,'  FR ','FRANCIA',0),
 (72,'  GA ','GABON',0),
 (73,'  GM ','GAMBIA',0),
 (74,'  GE ','GEORGIA',0),
 (75,'  GS ','GEORGIA DEL SUR E ISLAS SANDWICH DEL SUR',0),
 (76,'  GH ','GHANA',0),
 (77,'  GI ','GIBRALTAR',0),
 (78,'  GD ','GRANADA',0),
 (79,'  EL ','GRECIA',0),
 (80,'  GL ','GROENLANDIA',0),
 (81,'  GP ','GUADALUPE',0),
 (82,'  GU ','GUAM',0),
 (83,'  GT ','GUATEMALA',0),
 (84,'  GF ','GUAYANA FRANCESA',0),
 (85,'  GG ','GUERNESEY',0),
 (86,'  GN ','GUINEA',0),
 (87,'  GQ ','GUINEA ECUATORIAL',0),
 (88,'  GW ','GUINEA-BISSAU',0),
 (89,'  GY ','GUYANA',0),
 (90,'  HT ','HAITI',0),
 (91,'  HN ','HONDURAS',0),
 (92,'  HK ','HONG KONG',0),
 (93,'  HU ','HUNGRIA',0),
 (94,'  IN ','INDIA',0),
 (95,'  ID ','INDONESIA',0),
 (96,'  IR ','IRAN',0),
 (97,'  IQ ','IRAQ',0),
 (98,'  IE ','IRLANDA',0),
 (99,'  BV ','ISLA BOUVET',0),
 (100,'  CX ','ISLA CHRISTMAS',0),
 (101,'  IM ','ISLA DE MAN',0),
 (102,'  NF ','ISLA NORFOLK',0),
 (103,'  IS ','ISLANDIA',0),
 (104,'  AX ','ISLAS ÅLAND',0),
 (105,'  KY ','ISLAS CAIMAN',0),
 (106,'  CC ','ISLAS COCOS',0),
 (107,'  CK ','ISLAS COOK',0),
 (108,'  FO ','ISLAS FEROE',0),
 (109,'  HM ','ISLAS HEARD Y MCDONALD',0),
 (110,'  FK ','ISLAS MALVINAS',0),
 (111,'  MP ','ISLAS MARIANAS DEL NORTE',0),
 (112,'  MH ','ISLAS MARSHALL',0),
 (113,'  UM ','ISLAS MENORES ALEJADAS DE LOS ESTADOS UNIDOS',0),
 (114,'  PN ','ISLAS PITCAIRN',0),
 (115,'  SB ','ISLAS SALOMON',0),
 (116,'  TC ','ISLAS TURCAS Y CAICOS',0),
 (117,'  VG ','ISLAS VIRGENES BRITANICAS',0),
 (118,'  VI ','ISLAS VIRGENES DE LOS ESTADOS UNIDOS',0),
 (119,'  IL ','ISRAEL',0),
 (120,'  IT ','ITALIA',0),
 (121,'  JM ','JAMAICA',0),
 (122,'  JP ','JAPON',0),
 (123,'  JE ','JERSEY',0),
 (124,'  JO ','JORDANIA',0),
 (125,'  KZ ','KAZAJSTAN',0),
 (126,'  KE ','KENIA',0),
 (127,'  KG ','KIRGUISTAN',0),
 (128,'  KI ','KIRIBATI',0),
 (129,'  KW ','KUWAIT',0),
 (130,'  LA ','LAOS',0),
 (131,'  LS ','LESOTHO',0),
 (132,'  LV ','LETONIA',0),
 (133,'  LB ','LIBANO',0),
 (134,'  LR ','LIBERIA',0),
 (135,'  LY ','LIBIA',0),
 (136,'  LI ','LIECHTENSTEIN',0),
 (137,'  LT ','LITUANIA',0),
 (138,'  LU ','LUXEMBURGO',0),
 (139,'  MO ','MACAO',0),
 (140,'  MG ','MADAGASCAR',0),
 (141,'  MY ','MALASIA',0),
 (142,'  MW ','MALAWI',0),
 (143,'  MV ','MALDIVAS',0),
 (144,'  ML ','MALI',0),
 (145,'  MT ','MALTA',0),
 (146,'  MA ','MARRUECOS',0),
 (147,'  MQ ','MARTINICA',0),
 (148,'  MU ','MAURICIO',0),
 (149,'  MR ','MAURITANIA',0),
 (150,'  YT ','MAYOTTE',0),
 (151,'  MX ','MEXICO',1),
 (152,'  FM ','MICRONESIA',0),
 (153,'  MD ','MOLDOVA',0),
 (154,'  MC ','MONACO',0),
 (155,'  MN ','MONGOLIA',0),
 (156,'  ME ','MONTENEGRO',0),
 (157,'  MS ','MONTSERRAT',0),
 (158,'  MZ ','MOZAMBIQUE',0),
 (159,'  MM ','MYANMAR',0),
 (160,'  NA ','NAMIBIA',0),
 (161,'  NR ','NAURU',0),
 (162,'  NP ','NEPAL',0),
 (163,'  NI ','NICARAGUA',0),
 (164,'  NE ','NIGER',0),
 (165,'  NG ','NIGERIA',0),
 (166,'  NU ','NIUE',0),
 (167,'  NO ','NORUEGA',0),
 (168,'  NC ','NUEVA CALEDONIA',0),
 (169,'  NZ ','NUEVA ZELANDA',0),
 (170,'  OM ','OMAN',0),
 (171,'  NL ','PAISES BAJOS',0),
 (172,'  PK ','PAKISTAN',0),
 (173,'  PW ','PALAOS',0),
 (174,'  PA ','PANAMA',0),
 (175,'  PG ','PAPUA NUEVA GUINEA',0),
 (176,'  PY ','PARAGUAY',0),
 (177,'  PE ','PERU',0),
 (178,'  PF ','POLINESIA FRANCESA',0),
 (179,'  PL ','POLONIA',0),
 (180,'  PT ','PORTUGAL',0),
 (181,'  PR ','PUERTO RICO',0),
 (182,'  QA ','QATAR',0),
 (183,'  UK ','REINO UNIDO',0),
 (184,'  CF ','REPUBLICA CENTROAFRICANA',0),
 (185,'  CD ','REPUBLICA DEMOCRATICA DEL CONGO',0),
 (186,'  DO ','REPUBLICA DOMINICANA',0),
 (187,'  RE ','REUNION',0),
 (188,'  RW ','RUANDA',0),
 (189,'  RO ','RUMANIA',0),
 (190,'  RU ','RUSIA',0),
 (191,'  EH ','SAHARA OCCIDENTAL',0),
 (192,'  WS ','SAMOA',0),
 (193,'  AS ','SAMOA AMERICANA',0),
 (194,'  KN ','SAN CRISTOBAL Y NIEVES',0),
 (195,'  SM ','SAN MARINO',0),
 (196,'  PM ','SAN PEDRO Y MIQUELON',0),
 (197,'  VC ','SAN VICENTE Y LAS GRANADINAS',0),
 (198,'  SH ','SANTA ELENA',0),
 (199,'  LC ','SANTA LUCIA',0),
 (200,'  VA ','SANTA SEDE / ESTADO DE LA CIUDAD DEL VATICANO',0),
 (201,'  ST ','SANTO TOME Y PRINCIPE',0),
 (202,'  SN ','SENEGAL',0),
 (203,'  RS ','SERBIA',0),
 (204,'  SC ','SEYCHELLES',0),
 (205,'  SL ','SIERRA LEONA',0),
 (206,'  SG ','SINGAPUR',0),
 (207,'  SY ','SIRIA',0),
 (208,'  SO ','SOMALIA',0),
 (209,'  LK ','SRI LANKA',0),
 (210,'  SZ ','SUAZILANDIA',0),
 (211,'  ZA ','SUDAFRICA',0),
 (212,'  SD ','SUDAN',0),
 (213,'  SE ','SUECIA',0),
 (214,'  CH ','SUIZA',0),
 (215,'  SR ','SURINAM',0),
 (216,'  SJ ','SVALBARD Y JAN MAYEN',0),
 (217,'  TH ','TAILANDIA',0),
 (218,'  TW ','TAIWAN',0),
 (219,'  TZ ','TANZANIA',0),
 (220,'  TJ ','TAYIKISTAN',0),
 (221,'  IO ','TERRITORIO BRITANICO DEL OCEANO INDICO',0),
 (222,'  TF ','TERRITORIOS AUSTRALES FRANCESES',0),
 (223,'  PS ','TERRITORIOS PALESTINOS',0),
 (224,'  TL ','TIMOR ORIENTAL',0),
 (225,'  TG ','TOGO',0),
 (226,'  TK ','TOKELAU',0),
 (227,'  TO ','TONGA',0),
 (228,'  TT ','TRINIDAD Y TOBAGO',0),
 (229,'  TN ','TUNEZ',0),
 (230,'  TM ','TURKMENISTAN',0),
 (231,'  TR ','TURQUIA',0),
 (232,'  TV ','TUVALU',0),
 (233,'  UA ','UCRANIA',0),
 (234,'  UG ','UGANDA',0),
 (235,'  UY ','URUGUAY',0),
 (236,'  UZ ','UZBEKISTAN',0),
 (237,'  VU ','VANUATU',0),
 (238,'  VE ','VENEZUELA',0),
 (239,'  VN ','VIETNAM',0),
 (240,'  WF ','WALLIS Y FUTUNA',0),
 (241,'  YE ','YEMEN',0),
 (242,'  DJ ','YIBUTI',0),
 (243,'  ZM ','ZAMBIA',0),
 (244,'  ZW ','ZIMBABUE',0);
/*!40000 ALTER TABLE `mld_paises` ENABLE KEYS */;


--
-- Definition of table `mld_prestamos_tipo_alertas`
--

DROP TABLE IF EXISTS `mld_prestamos_tipo_alertas`;
CREATE TABLE `mld_prestamos_tipo_alertas` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  `ReqDesc` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=18 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_prestamos_tipo_alertas`
--

/*!40000 ALTER TABLE `mld_prestamos_tipo_alertas` DISABLE KEYS */;
INSERT INTO `mld_prestamos_tipo_alertas` (`Id`,`Clave`,`Descripcion`,`RegDefault`,`ReqDesc`) VALUES 
 (1,'100','Sin alerta',1,0),
 (2,'401 ','El cliente o usuario se rehúsa a proporcionar documentos personales que lo identifiquen',0,0),
 (3,'402 ','El cliente o usuario realiza varias operaciones en un periodo corto de tiempo en las que se desconoce el origen de los objetos empeñados',0,0),
 (4,'403 ','La operación de mutuo o crédito se lleva a cabo por medio de una garantía poco usual o que no corresponde con la actividad o ingresos del cliente o usuario',0,0),
 (5,'404 ','Hay indicios, o certeza, de que los bienes empeñados son robados o provienen de una actividad ilícita',0,0),
 (6,'405 ','El cliente o usuario o personas relacionadas con él realizan múltiples operaciones en un periodo muy corto sin razón aparente',0,0),
 (7,'406 ','El cliente o usuario no muestra tener interés en las características y condiciones del crédito otorgado',0,0),
 (8,'407 ','El cliente o usuario realiza operaciones de manera periódica en las que se liquida el total del monto del préstamo otorgado en efectivo al poco tiempo de haberlo adquirido',0,0),
 (9,'408 ','Se conoce un historial criminal del cliente o usuario, de algún familiar directo o persona relacionada con él',0,0),
 (10,'409 ','El cliente o usuario no quiere ser relacionado con la operación realizada',0,0),
 (11,'410 ','El cliente o usuario muestra fuerte interés en la realización de la transacción con rapidez, sin que exista causa justificada',0,0),
 (12,'411 ','Hay indicios, o certeza, que las partes no están actuando en nombre propio y están tratando de ocultar la identidad del cliente o usuario real',0,0),
 (13,'412 ','El cliente o usuario intenta sobornar o extorsionar al vendedor con  el fin de realizar la operación de forma irregular',0,0),
 (14,'413 ','El cliente o usuario solicita condiciones especiales poco usuales en la realización de la operación',0,0),
 (15,'414 ','El cliente o usuario proporcionó datos falsos  o documentos apócrifos al realizar la operación',0,0),
 (16,'415','Operaciones con organizaciones sin fines de lucro, cuando las características de la transacción no coinciden con los objetivos de la entidad',0,0),
 (17,'9999','Otra alerta',0,1);
/*!40000 ALTER TABLE `mld_prestamos_tipo_alertas` ENABLE KEYS */;


--
-- Definition of table `mld_prestamos_tipo_garantias`
--

DROP TABLE IF EXISTS `mld_prestamos_tipo_garantias`;
CREATE TABLE `mld_prestamos_tipo_garantias` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `Estatus` tinyint(2) DEFAULT '0',
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=17 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_prestamos_tipo_garantias`
--

/*!40000 ALTER TABLE `mld_prestamos_tipo_garantias` DISABLE KEYS */;
INSERT INTO `mld_prestamos_tipo_garantias` (`Id`,`Clave`,`Descripcion`,`Estatus`,`RegDefault`) VALUES 
 (1,'  1  ','Sin garantía',0,0),
 (2,'  2  ','Inmueble',1,0),
 (3,'  3  ','Vehículo terrestre',1,0),
 (4,'  4  ','Vehículo aéreo',0,0),
 (5,'  5  ','Vehículo marítimo',1,0),
 (6,'  6  ','Priedras Preciosas',1,0),
 (7,'  7  ','Metales Preciosos',1,0),
 (8,'  8  ','Joyas o relojes',1,1),
 (9,'  9  ','Obras de arte o antigüedades',0,0),
 (10,'  10 ','Acciones o partes sociales',0,0),
 (11,'  11 ','Derechos fiduciarios',0,0),
 (12,'  12 ','Derechos de crédito',0,0),
 (13,'  13 ','Prenda con disposición',1,0),
 (14,'  14 ','Prenda sin disposición',0,0),
 (15,'  15 ','Garantía Quirografaria',0,0),
 (16,'  99 ','Otro (Especificar)',0,0);
/*!40000 ALTER TABLE `mld_prestamos_tipo_garantias` ENABLE KEYS */;


--
-- Definition of table `mld_prestamos_tipo_operacion`
--

DROP TABLE IF EXISTS `mld_prestamos_tipo_operacion`;
CREATE TABLE `mld_prestamos_tipo_operacion` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=3 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_prestamos_tipo_operacion`
--

/*!40000 ALTER TABLE `mld_prestamos_tipo_operacion` DISABLE KEYS */;
INSERT INTO `mld_prestamos_tipo_operacion` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'401 ','Otorgamiento de Mutuo, Préstamo o Crédito sin Garantía',0),
 (2,'402 ','Otorgamiento de Mutuo, Préstamo o Crédito con Garantía',1);
/*!40000 ALTER TABLE `mld_prestamos_tipo_operacion` ENABLE KEYS */;


--
-- Definition of table `mld_tipo_identificaciones`
--

DROP TABLE IF EXISTS `mld_tipo_identificaciones`;
CREATE TABLE `mld_tipo_identificaciones` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(120) DEFAULT NULL,
  `Dependencia` varchar(80) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  `Estatus` tinyint(2) DEFAULT '0',
  `ReqDesc` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=14 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_tipo_identificaciones`
--

/*!40000 ALTER TABLE `mld_tipo_identificaciones` DISABLE KEYS */;
INSERT INTO `mld_tipo_identificaciones` (`Id`,`Clave`,`Descripcion`,`Dependencia`,`RegDefault`,`Estatus`,`ReqDesc`) VALUES 
 (1,'  1  ','Credencial para votar','Instituto Federal Electoral',1,1,0),
 (2,'  2  ','Pasaporte','Secretaria de Relaciones Exteriores',0,1,0),
 (3,'  3  ','Documentación expedida por el Instituto Nacional de Migración','Instituto Nacional de Migracion',0,1,0),
 (4,'  4  ','Cédula Profesional','Secretaria de Educacion Publica',0,1,0),
 (5,'  5  ','Cartilla de Servicio Militar','Secretaria de la Defensa Nacional',0,1,0),
 (6,'  6  ','Certificado de matrícula consular','Secretaria de Relaciones Exteriores',0,1,0),
 (7,'  7  ','Tarjeta única de identificación militar','Secretaria de la Defensa Nacional',0,1,0),
 (8,'  8  ','Tarjeta de afiliación al Instituto Nacional de las Personas Adultas M','Instituto Nacional de las Personas Adultas M',0,1,0),
 (9,'  9  ','Credenciales y Carnets expedidos por el Instituto Mexicano del Seguro','Instituto Mexicano del Seguro Social',0,1,0),
 (10,'  10 ','Licencia para conducir','Secretaria de Transito y Vialidad',0,1,0),
 (11,'  11 ','Otra credencial emitida por autoridades federales',NULL,0,0,1),
 (12,'  12 ','Otra credencial emitida por autoridades estatales',NULL,0,0,1),
 (13,'  13 ','Otra credencial emitida por autoridades municipales',NULL,0,0,1);
/*!40000 ALTER TABLE `mld_tipo_identificaciones` ENABLE KEYS */;


--
-- Definition of table `mld_tipo_inmuebles`
--

DROP TABLE IF EXISTS `mld_tipo_inmuebles`;
CREATE TABLE `mld_tipo_inmuebles` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) NOT NULL,
  `Descripcion` varchar(600) NOT NULL,
  `Actualizar` tinyint(4) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=20 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_tipo_inmuebles`
--

/*!40000 ALTER TABLE `mld_tipo_inmuebles` DISABLE KEYS */;
INSERT INTO `mld_tipo_inmuebles` (`Id`,`Clave`,`Descripcion`,`Actualizar`) VALUES 
 (1,'1','Casa /Casa en condominio',0),
 (2,'2','Departamento',0),
 (3,'6','Local comercial independiente',0),
 (4,'12','Terreno urbano habitacional',0),
 (5,'3','Edificio habitacional',0),
 (6,'4','Edificio comercial',0),
 (7,'5','Edificio oficinas',0),
 (8,'7','Local en centro comercial',0),
 (9,'8','Oficina',0),
 (10,'9','Bodega comercial',0),
 (11,'10','Bodega industrial',0),
 (12,'11','Nave Industrial',0),
 (13,'13','Terreno no urbano habitacional',0),
 (14,'14','Terreno urbano comercial o industrial',0),
 (15,'15','Terreno no urbano comercial o industrial',0),
 (16,'16','Terreno ejidal',0),
 (17,'17','Rancho/Hacienda/Quinta',0),
 (18,'18','Huerta',0),
 (19,'99','Otro',0);
/*!40000 ALTER TABLE `mld_tipo_inmuebles` ENABLE KEYS */;


--
-- Definition of table `mld_tipo_inmuebles_base`
--

DROP TABLE IF EXISTS `mld_tipo_inmuebles_base`;
CREATE TABLE `mld_tipo_inmuebles_base` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=20 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_tipo_inmuebles_base`
--

/*!40000 ALTER TABLE `mld_tipo_inmuebles_base` DISABLE KEYS */;
INSERT INTO `mld_tipo_inmuebles_base` (`Id`,`Clave`,`Descripcion`) VALUES 
 (1,'  1  ','Casa /Casa en condominio'),
 (2,'  2  ','Departamento'),
 (3,'  3  ','Edificio habitacional'),
 (4,'  4  ','Edificio comercial'),
 (5,'  5  ','Edificio oficinas'),
 (6,'  6  ','Local comercial independiente'),
 (7,'  7  ','Local en centro comercial'),
 (8,'  8  ','Oficina'),
 (9,'  9  ','Bodega comercial'),
 (10,'  10 ','Bodega industrial'),
 (11,'  11 ','Nave Industrial'),
 (12,'  12 ','Terreno urbano habitacional'),
 (13,'  13 ','Terreno no urbano habitacional'),
 (14,'  14 ','Terreno urbano comercial o industrial'),
 (15,'  15 ','Terreno no urbano comercial o industrial'),
 (16,'  16 ','Terreno ejidal'),
 (17,'  17 ','Rancho/Hacienda/Quinta'),
 (18,'  18 ','Huerta'),
 (19,'  99 ','Otro');
/*!40000 ALTER TABLE `mld_tipo_inmuebles_base` ENABLE KEYS */;


--
-- Definition of table `mld_tipo_monedas`
--

DROP TABLE IF EXISTS `mld_tipo_monedas`;
CREATE TABLE `mld_tipo_monedas` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Moneda` varchar(600) DEFAULT NULL,
  `Pais` varchar(40) DEFAULT NULL,
  `Estatus` tinyint(2) DEFAULT '0',
  `MonedaDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=InnoDB AUTO_INCREMENT=185 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_tipo_monedas`
--

/*!40000 ALTER TABLE `mld_tipo_monedas` DISABLE KEYS */;
INSERT INTO `mld_tipo_monedas` (`ID`,`Clave`,`Moneda`,`Pais`,`Estatus`,`MonedaDefault`) VALUES 
 (1,'MXN  ','Peso mexicano                     ','México',1,1),
 (2,'USD  ','Dólar estadounidense              ','Estados Unidos',1,0),
 (3,'EUR  ','Euro                              ','Unión Europea',0,0),
 (4,'AED  ','Dirham de los Emiratos Árabes Unidos','Emiratos Árabes Unidos',0,0),
 (5,'AFN  ','Afgani afgano                     ','Afganistán',0,0),
 (6,'ALL  ','Lek albanés                       ','Albania',0,0),
 (7,'AMD  ','Dram armenio                      ','Armenia',0,0),
 (8,'ANG  ','Florín antillano neerlandés       ','Antillas Neerlandesas',0,0),
 (9,'AOA  ','Kwanza angoleño                   ','Angola',0,0),
 (10,'ARS  ','Peso argentino                    ','Argentina',0,0),
 (11,'AUD  ','Dólar australiano                 ','Australia',0,0),
 (12,'AWG  ','Florín arubeño                    ','Aruba',0,0),
 (13,'AZM  ','Manat azerbaiyano                 ','Azerbaiyán',0,0),
 (14,'AZN  ','Manat azerbaiyano                 ','Azerbaiyán',0,0),
 (15,'BAM  ','Marco convertible de Bosnia-Herzeg','Bosnia Y Herzegovina',0,0),
 (16,'BBD  ','Dólar de Barbados                 ','Barbados',0,0),
 (17,'BDT  ','Taka de Bangladesh                ','Bangladesh',0,0),
 (18,'BGN  ','Lev búlgaro                       ','Bulgaria',0,0),
 (19,'BHD  ','Dinar bahreiní                    ','Bahréin',0,0),
 (20,'BIF  ','Franco burundés                   ','Burundi',0,0),
 (21,'BMD  ','Dólar de Bermuda                  ','Bermudas',0,0),
 (22,'BND  ','Dólar de Brunéi                   ','Brunei',0,0),
 (23,'BOB  ','Boliviano                         ','Bolivia',0,0),
 (24,'BRL  ','Real brasileño                    ','Brasil',0,0),
 (25,'BSD  ','Dólar bahameño                    ','Bahamas',0,0),
 (26,'BTN  ','Ngultrum de Bután                 ','Bhutan',0,0),
 (27,'BWP  ','Pula de Botsuana                  ','Botsuana',0,0),
 (28,'BYR  ','Rublo bielorruso                  ','Belarus',0,0),
 (29,'BZD  ','Dólar de Belice                   ','Belice',0,0),
 (30,'CAD  ','Dólar canadiense                  ','Canadá',0,0),
 (31,'CDF  ','Franco congoleño, o congolés      ','Republica Democrática Del Congo',0,0),
 (32,'CHF  ','Franco suizo                      ','Suiza',0,0),
 (33,'CLP  ','Peso chileno                      ','Chile',0,0),
 (34,'CNY  ','Yuan chino                        ','China',0,0),
 (35,'COP  ','Peso colombiano                   ','Colombia',0,0),
 (36,'CRC  ','Colón costarricense               ','Costa Rica',0,0),
 (37,'CSD  ','Dinar serbio                      ','Serbia',0,0),
 (38,'CUC  ','Peso cubano convertible           ','Cuba',0,0),
 (39,'CUP  ','Peso cubano                       ','Cuba',0,0),
 (40,'CVE  ','Escudo caboverdiano               ','Cabo Verde',0,0),
 (41,'CZK  ','Koruna checa                      ','Chequia',0,0),
 (42,'DJF  ','Franco yibutiano                  ','Yibuti',0,0),
 (43,'DKK  ','Corona danesa                     ','Dinamarca',0,0),
 (44,'DOP  ','Peso dominicano                   ','Republica Dominicana',0,0),
 (45,'DZD  ','Dinar algerino                    ','Argelia',0,0),
 (46,'EGP  ','Libra egipcia                     ','Egipto',0,0),
 (47,'ERN  ','Nakfa eritreo                     ','Eritrea',0,0),
 (48,'ETB  ','Birr etíope                       ','Etiopia',0,0),
 (49,'FJD  ','Dólar fiyiano                     ','Fiyi',0,0),
 (50,'FKP  ','Libra malvinense                  ','Islas Malvinas',0,0),
 (51,'GBP  ','Libra esterlina (libra de Gran Bre','Reino Unido',0,0),
 (52,'GEL  ','Lari georgiano                    ','Georgia',0,0),
 (53,'GHS  ','Cedi ghanés                       ','Ghana',0,0),
 (54,'GIP  ','Libra de Gibraltar                ','Gibraltar',0,0),
 (55,'GMD  ','Dalasi gambiano                   ','Gambia',0,0),
 (56,'GNF  ','Franco guineano                   ','Guinea',0,0),
 (57,'GTQ  ','Quetzal guatemalteco              ','Guatemala',0,0),
 (58,'GYD  ','Dólar guyanés                     ','Guyana',0,0),
 (59,'HKD  ','Dólar de Hong Kong                ','Hong Kong',0,0),
 (60,'HNL  ','Lempira hondureño                 ','Honduras',0,0),
 (61,'HRK  ','Kuna croata                       ','Croacia',0,0),
 (62,'HTG  ','Gourde haitiano                   ','Haití',0,0),
 (63,'HUF  ','Forint húngaro                    ','Hungría',0,0),
 (64,'IDR  ','Rupiah indonesia                  ','Indonesia',0,0),
 (65,'ILS  ','Nuevo shéquel israelí             ','Israel',0,0),
 (66,'INR  ','Rupia india                       ','India',0,0),
 (67,'IQD  ','Dinar iraquí                      ','Iraq',0,0),
 (68,'IRR  ','Rial iraní                        ','Irán',0,0),
 (69,'ISK  ','Króna islandesa                   ','Islandia',0,0),
 (70,'JMD  ','Dólar jamaicano                   ','Jamaica',0,0),
 (71,'JOD  ','Dinar jordano                     ','Jordania',0,0),
 (72,'JPY  ','Yen japonés                       ','Japón',0,0),
 (73,'KES  ','Chelín keniata                    ','Kenia',0,0),
 (74,'KGS  ','Som kirguís (de Kirguistán)       ','Kirguistán',0,0),
 (75,'KHR  ','Riel camboyano                    ','Camboya',0,0),
 (76,'KMF  ','Franco comoriano (de Comoras)     ','Comoras',0,0),
 (77,'KPW  ','Won norcoreano                    ','Corea Del Norte',0,0),
 (78,'KRW  ','Won surcoreano                    ','Corea Del Sur',0,0),
 (79,'KWD  ','Dinar kuwaití                     ','Kuwait',0,0),
 (80,'KYD  ','Dólar caimano (de Islas Caimán)   ','Islas Caimán',0,0),
 (81,'KZT  ','Tenge kazajo                      ','Kazajstán',0,0),
 (82,'LAK  ','Kip lao                           ','Laos',0,0),
 (83,'LBP  ','Libra libanesa                    ','Líbano',0,0),
 (84,'LKR  ','Rupia de Sri Lanka                ','Sri Lanka',0,0),
 (85,'LRD  ','Dólar liberiano                   ','Liberia',0,0),
 (86,'LSL  ','Loti lesotense                    ','Lesotho',0,0),
 (87,'LTL  ','Litas lituano                     ','Lituania',0,0),
 (88,'LVL  ','Lat letón                         ','Letonia',0,0),
 (89,'LYD  ','Dinar libio                       ','Libia',0,0),
 (90,'MAD  ','Dirham marroquí                   ','Marruecos',0,0),
 (91,'MDL  ','Leu moldavo                       ','Moldova',0,0),
 (92,'MGA  ','Ariary malgache                   ','Madagascar',0,0),
 (93,'MKD  ','Denar macedonio                   ','Macedonia',0,0),
 (94,'MMK  ','Kyat birmano                      ','Myanmar',0,0),
 (95,'MNT  ','Tughrik mongol                    ','Mongolia',0,0),
 (96,'MOP  ','Pataca de Macao                   ','Macao',0,0),
 (97,'MRO  ','Ouguiya mauritana                 ','Mauritania',0,0),
 (98,'MUR  ','Rupia mauricia                    ','Mauricio',0,0),
 (99,'MVR  ','Rufiyaa maldiva                   ','Maldivas',0,0),
 (100,'MWK  ','Kwacha malauí                     ','Malawi',0,0),
 (101,'MYR  ','Ringgit malayo                    ','Malasia',0,0),
 (102,'MZN  ','Metical mozambiqueño              ','Mozambique',0,0),
 (103,'NAD  ','Dólar namibio                     ','Namibia',0,0),
 (104,'NGN  ','Naira nigeriana                   ','Nigeria',0,0),
 (105,'NIO  ','Córdoba nicaragüense              ','Nicaragua',0,0),
 (106,'NOK  ','Corona noruega                    ','Noruega',0,0),
 (107,'NPR  ','Rupia nepalesa                    ','Nepal',0,0),
 (108,'NZD  ','Dólar neozelandés                 ','Nueva Zelanda',0,0),
 (109,'OMR  ','Rial omaní                        ','Omán',0,0),
 (110,'PAB  ','Balboa panameña                   ','Panamá',0,0),
 (111,'PEN  ','Nuevo sol peruano                 ','Perú',0,0),
 (112,'PGK  ','Kina de Papúa Nueva Guinea        ','Papúa Nueva Guinea',0,0),
 (113,'PHP  ','Peso filipino                     ','Filipinas',0,0),
 (114,'PKR  ','Rupia pakistaní                   ','Pakistán',0,0),
 (115,'PLN  ','zloty polaco                      ','Polonia',0,0),
 (116,'PYG  ','Guaraní paraguayo                 ','Paraguay',0,0),
 (117,'QAR  ','Rial qatarí                       ','Qatar',0,0),
 (118,'RON  ','Leu rumano                        ','Rumania',0,0),
 (119,'RSD  ','Dinar serbio                      ','Serbia',0,0),
 (120,'RUB  ','Rublo ruso                        ','Rusia',0,0),
 (121,'RWF  ','Franco ruandés                    ','Ruanda',0,0),
 (122,'SAR  ','Riyal saudí                       ','Arabia Saudí',0,0),
 (123,'SBD  ','Dólar de las Islas Salomón        ','Islas Salomón',0,0),
 (124,'SCR  ','Rupia de Seychelles               ','Seychelles',0,0),
 (125,'SDG  ','Dinar sudanés                     ','Sudan',0,0),
 (126,'SEK  ','Corona sueca                      ','Suecia',0,0),
 (127,'SGD  ','Dólar de Singapur                 ','Singapur',0,0),
 (128,'SHP  ','Libra de Santa Helena             ','Santa Elena',0,0),
 (129,'SLL  ','Leone de Sierra Leona             ','Sierra Leona',0,0),
 (130,'SOS  ','Chelín somalí                     ','Somalia',0,0),
 (131,'SRD  ','Dólar surinamés                   ','Surinam',0,0),
 (132,'SSP  ','Libra de Sudán del Sur            ','Sudan Del Sur',0,0),
 (133,'STD  ','Dobra de Santo Tomé y Príncipe    ','Santo Tome Y Príncipe',0,0),
 (134,'SVC  ','Colón Salvadoreño                 ','El Salvador',0,0),
 (135,'SYP  ','Libra siria                       ','Siria',0,0),
 (136,'SZL  ','Lilangeni suazi                   ','Suazilandia',0,0),
 (137,'THB  ','Baht tailandés                    ','Tailandia',0,0),
 (138,'TJS  ','Somoni tayik (de Tayikistán)      ','Tayikistán',0,0),
 (139,'TMT  ','Manat turcomano                   ','Turkmenistán',0,0),
 (140,'TND  ','Dinar tunecino                    ','Túnez',0,0),
 (141,'TOP  ','Pa\'anga tongano                   ','Tonga',0,0),
 (142,'TRY  ','Lira turca                        ','Turquía',0,0),
 (143,'TTD  ','Dólar de Trinidad y Tobago        ','Trinidad Y Tobago',0,0),
 (144,'TWD  ','Dólar taiwanés                    ','Taiwán',0,0),
 (145,'TZS  ','Chelín tanzano                    ','Tanzania',0,0),
 (146,'UAH  ','Grivna ucraniana                  ','Ucrania',0,0),
 (147,'UGX  ','Chelín ugandés                    ','Uganda',0,0),
 (148,'UYU  ','Peso uruguayo                     ','Uruguay',0,0),
 (149,'UZS  ','Som uzbeko                        ','Uzbekistán',0,0),
 (150,'VEF  ','Bolívar fuerte venezolano         ','Venezuela',0,0),
 (151,'VND  ','Dong vietnamita                   ','Vietnam',0,0),
 (152,'VUV  ','Vatu vanuatense                   ','Vanuatu',0,0),
 (153,'WST  ','Tala samoana                      ','Samoa',0,0),
 (154,'YER  ','Rial yemení (de Yemen)            ','Yemen',0,0),
 (155,'ZAR  ','Rand sudafricano                  ','Sudáfrica',0,0),
 (156,'ZMK  ','Kwacha zambiano                   ','Zambia',0,0),
 (157,'ZMW  ','Kwacha zambiano                   ','Zambia',0,0),
 (158,'ZWL  ','Dólar zimbabuense                 ','Zimbabue',0,0),
 (159,'MXA  ','CENTENARIO                        ','México',0,0),
 (160,'MXB  ','AZTECA                            ','México',0,0),
 (161,'MXC  ','HIDALGO                           ','México',0,0),
 (162,'MXD  ','1/2 HIDALGO                       ','México',0,0),
 (163,'MXE  ','1/4 HIDALGO                       ','México',0,0),
 (164,'MXF  ','1/5 HIDALGO                       ','México',0,0),
 (165,'MXG  ','1 OZ LIBERTAD DE ORO              ','México',0,0),
 (166,'MXH  ','1/2 OZ LIBERTAD DE ORO            ','México',0,0),
 (167,'MXI  ','1/4 OZ LIBERTAD DE ORO            ','México',0,0),
 (168,'MXJ  ','1/10 OZ LIBERTAD DE ORO           ','México',0,0),
 (169,'MXK  ','1/20 OZ LIBERTAD DE ORO           ','México',0,0),
 (170,'MXL  ','1 OZ LIBERTAD DE PLATA            ','México',0,0),
 (171,'MXM  ','1/2 OZ LIBERTAD DE PLATA          ','México',0,0),
 (172,'MXN  ','1/4 OZ LIBERTAD DE PLATA          ','México',0,0),
 (173,'MXO  ','1/10 OZ LIBERTAD DE PLATA         ','México',0,0),
 (174,'MXP  ','1/20 OZ LIBERTAD DE PLATA         ','México',0,0),
 (175,'XAF  ','Franco CFA de África Central      ','Na',0,0),
 (176,'XAG  ','Onza de plata                     ','Na',0,0),
 (177,'XAU  ','Onza de oro                       ','Na',0,0),
 (178,'XCD  ','Dólar del Caribe Oriental         ','Na',0,0),
 (179,'XFO  ','Franco de oro (Special settlement ','Na',0,0),
 (180,'XFU  ','Franco UIC (Special settlement cur','Na',0,0),
 (181,'XOF  ','Franco CFA de África Occidental   ','Na',0,0),
 (182,'XPD  ','Onza de paladio                   ','Na',0,0),
 (183,'XPF  ','Franco CFP                        ','Na',0,0),
 (184,'XPT  ','Onza de platino                   ','Na',0,0);
/*!40000 ALTER TABLE `mld_tipo_monedas` ENABLE KEYS */;


--
-- Definition of table `mld_vehiculos_forma_pago`
--

DROP TABLE IF EXISTS `mld_vehiculos_forma_pago`;
CREATE TABLE `mld_vehiculos_forma_pago` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `Estatus` tinyint(2) DEFAULT '0',
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_vehiculos_forma_pago`
--

/*!40000 ALTER TABLE `mld_vehiculos_forma_pago` DISABLE KEYS */;
INSERT INTO `mld_vehiculos_forma_pago` (`Id`,`Clave`,`Descripcion`,`Estatus`,`RegDefault`) VALUES 
 (1,'1','Contado',1,1),
 (2,'2','Financiamiento',1,0),
 (3,'3','Dacion en pago',0,0);
/*!40000 ALTER TABLE `mld_vehiculos_forma_pago` ENABLE KEYS */;


--
-- Definition of table `mld_vehiculos_tipo_alertas`
--

DROP TABLE IF EXISTS `mld_vehiculos_tipo_alertas`;
CREATE TABLE `mld_vehiculos_tipo_alertas` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  `ReqDesc` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=23 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_vehiculos_tipo_alertas`
--

/*!40000 ALTER TABLE `mld_vehiculos_tipo_alertas` DISABLE KEYS */;
INSERT INTO `mld_vehiculos_tipo_alertas` (`Id`,`Clave`,`Descripcion`,`RegDefault`,`ReqDesc`) VALUES 
 (1,'100','Sin alerta',1,0),
 (2,'801','El cliente o usuario se rehúsa a proporcionar documentos personales que lo identifiquen',0,0),
 (3,'802','El financiamiento otorgado por el vehículo es pagado por un tercero sin relación aparente con el cliente o usuario',0,0),
 (4,'803','El cliente o usuario solicita el reembolso del pago del vehículo poco tiempo después de ser adquirido',0,0),
 (5,'804','El cliente o usuario compra el vehículo sin inspeccionarlo o revisarlo',0,0),
 (6,'805','El cliente o usuario compra múltiples vehículos en un periodo muy corto sin tener la preocupación sobre el costo, condiciones o tipo de vehículos',0,0),
 (7,'806','Se conoce un historial criminal del cliente o usuario, de algún familiar directo o persona relacionada con él',0,0),
 (8,'807','El cliente o usuario no quiere ser relacionado con la compra del vehículo',0,0),
 (9,'808','De acuerdo con la ocupación del cliente o usuario, la compra del vehículo parece estar fuera de su alcance',0,0),
 (10,'809','De acuerdo con los ingresos declarados por el cliente o usuario, la compra del vehículo parece estar fuera de su alcance',0,0),
 (11,'810','El cliente o usuario vende su vehículo a precios muy por debajo del precio de mercado',0,0),
 (12,'811','Hay indicios, o certeza, que las partes no están actuando en nombre propio y están tratando de ocultar la identidad del cliente o usuario real',0,0),
 (13,'812','Uso de divisas en efectivo en montos elevados o de poco uso sin que la ocupación del cliente o usuario lo justifique',0,0),
 (14,'813','La operación se liquida por medio de una transferencia internacional sin que la ocupación, perfil o nacionalidad del cliente o usuario lo justifique',0,0),
 (15,'814','La operación se liquida por medio de una transferencia internacional proveniente de un país considerado como paraíso fiscal o de alto riesgo',0,0),
 (16,'815','El cliente o usuario muestra fuerte interés en la realización de la transacción con rapidez, sin que exista causa justificada',0,0),
 (17,'816','El cliente o usuario insiste en liquidar pagar la operación en efectivo rebasando el umbral permitido para uso de efectivo',0,0),
 (18,'817','El cliente o usuario intenta sobornar o extorsionar al vendedor con  el fin de realizar la operación de forma irregular',0,0),
 (19,'818','El cliente o usuario solicita condiciones especiales poco usuales en la realización de la operación',0,0),
 (20,'819','El cliente o usuario proporcionó datos falsos  o documentos apócrifos al realizar la operación',0,0),
 (21,'820','Operaciones con organizaciones sin fines de lucro, cuando las características de la transacción no coinciden con los objetivos de la entidad',0,0),
 (22,'9999','Otra alerta',0,1);
/*!40000 ALTER TABLE `mld_vehiculos_tipo_alertas` ENABLE KEYS */;


--
-- Definition of table `mld_vehiculos_tipo_blindaje`
--

DROP TABLE IF EXISTS `mld_vehiculos_tipo_blindaje`;
CREATE TABLE `mld_vehiculos_tipo_blindaje` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=8 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_vehiculos_tipo_blindaje`
--

/*!40000 ALTER TABLE `mld_vehiculos_tipo_blindaje` DISABLE KEYS */;
INSERT INTO `mld_vehiculos_tipo_blindaje` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'1','Nivel A',0),
 (2,'2','Nivel B',0),
 (3,'3','Nivel B Plus',0),
 (4,'4','Nivel C',0),
 (5,'5','Nivel C Plus',0),
 (6,'6','Nivel D',0),
 (7,'7','Nivel E',0);
/*!40000 ALTER TABLE `mld_vehiculos_tipo_blindaje` ENABLE KEYS */;


--
-- Definition of table `mld_vehiculos_tipo_operacion`
--

DROP TABLE IF EXISTS `mld_vehiculos_tipo_operacion`;
CREATE TABLE `mld_vehiculos_tipo_operacion` (
  `Id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` varchar(5) DEFAULT NULL,
  `Descripcion` varchar(600) DEFAULT NULL,
  `RegDefault` tinyint(2) DEFAULT '0',
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `mld_vehiculos_tipo_operacion`
--

/*!40000 ALTER TABLE `mld_vehiculos_tipo_operacion` DISABLE KEYS */;
INSERT INTO `mld_vehiculos_tipo_operacion` (`Id`,`Clave`,`Descripcion`,`RegDefault`) VALUES 
 (1,'801','Venta de vehículo nuevo',0),
 (2,'802','Venta de vehículo usado',1),
 (3,'803','Compra de vehículo nuevo',0),
 (4,'804','Compra de vehículo usado',1);
/*!40000 ALTER TABLE `mld_vehiculos_tipo_operacion` ENABLE KEYS */;


--
-- Definition of table `monedas`
--

DROP TABLE IF EXISTS `monedas`;
CREATE TABLE `monedas` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Clave` int(10) DEFAULT '0',
  `Descripcion` varchar(50) DEFAULT NULL,
  `Compra` double(15,5) DEFAULT '0.00000',
  `Venta` double(15,5) DEFAULT '0.00000',
  `Maximo` int(10) DEFAULT '0',
  `Defoult` int(1) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `id` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `monedas`
--

/*!40000 ALTER TABLE `monedas` DISABLE KEYS */;
/*!40000 ALTER TABLE `monedas` ENABLE KEYS */;


--
-- Definition of table `movimientos`
--

DROP TABLE IF EXISTS `movimientos`;
CREATE TABLE `movimientos` (
  `ID` int(10) NOT NULL,
  `Movimiento` int(10) DEFAULT '0',
  `FolioBancos` int(10) DEFAULT '0',
  `FolioGastos` int(10) DEFAULT '0',
  `FolioVentas` int(10) DEFAULT '0',
  `FolioDepositos` int(10) DEFAULT '0',
  `FolioTransferencias` int(10) DEFAULT '0',
  `FolioCompras` int(10) DEFAULT '0',
  `FolioSalidaInventario` int(10) DEFAULT '0',
  `FolioAjustes` int(10) DEFAULT '0',
  `FolioBoveda` int(10) DEFAULT '0',
  `FolioDivisas` int(10) DEFAULT '0',
  `FolioNotas` int(10) DEFAULT '0',
  `Fecha` date DEFAULT NULL,
  `FolioAutorizacion` int(10) DEFAULT '0',
  `FolioInventario` int(10) DEFAULT '0',
  `FolioTraspasos` int(10) DEFAULT '0',
  `FolioBovedaDivisas` int(10) DEFAULT '1',
  `FolioAvisosLavado` int(10) DEFAULT '1',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `movimientos`
--

/*!40000 ALTER TABLE `movimientos` DISABLE KEYS */;
INSERT INTO `movimientos` (`ID`,`Movimiento`,`FolioBancos`,`FolioGastos`,`FolioVentas`,`FolioDepositos`,`FolioTransferencias`,`FolioCompras`,`FolioSalidaInventario`,`FolioAjustes`,`FolioBoveda`,`FolioDivisas`,`FolioNotas`,`Fecha`,`FolioAutorizacion`,`FolioInventario`,`FolioTraspasos`,`FolioBovedaDivisas`,`FolioAvisosLavado`) VALUES 
 (1,1,1,1,1,1,1,1,1,1,1,1,1,'2014-07-15',1,1,1,1,1);
/*!40000 ALTER TABLE `movimientos` ENABLE KEYS */;


--
-- Definition of table `movimientospuntos`
--

DROP TABLE IF EXISTS `movimientospuntos`;
CREATE TABLE `movimientospuntos` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Fecha` datetime NOT NULL,
  `IDTarjeta` int(10) unsigned NOT NULL DEFAULT '0',
  `TipoMovimiento` int(10) unsigned NOT NULL DEFAULT '0',
  `Concepto` varchar(80) NOT NULL,
  `Folio` int(10) unsigned NOT NULL DEFAULT '0',
  `Cargo` decimal(14,4) NOT NULL DEFAULT '0.0000',
  `Abono` decimal(14,4) NOT NULL DEFAULT '0.0000',
  `Importe` decimal(14,4) NOT NULL DEFAULT '0.0000',
  `PC` varchar(45) NOT NULL,
  `IDUsuario` int(10) unsigned NOT NULL DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `movimientospuntos`
--

/*!40000 ALTER TABLE `movimientospuntos` DISABLE KEYS */;
/*!40000 ALTER TABLE `movimientospuntos` ENABLE KEYS */;


--
-- Definition of table `nacionalidad`
--

DROP TABLE IF EXISTS `nacionalidad`;
CREATE TABLE `nacionalidad` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Nacionalidad` varchar(30) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `nacionalidad`
--

/*!40000 ALTER TABLE `nacionalidad` DISABLE KEYS */;
INSERT INTO `nacionalidad` (`ID`,`Nacionalidad`) VALUES 
 (1,'Mexicana');
/*!40000 ALTER TABLE `nacionalidad` ENABLE KEYS */;


--
-- Definition of table `ocupaciones`
--

DROP TABLE IF EXISTS `ocupaciones`;
CREATE TABLE `ocupaciones` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `descripcion` varchar(50) DEFAULT NULL,
  `estatus` tinyint(4) DEFAULT '1',
  `Ordenamiento` int(2) DEFAULT '0',
  `Actualizar` tinyint(4) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `ocupaciones`
--

/*!40000 ALTER TABLE `ocupaciones` DISABLE KEYS */;
/*!40000 ALTER TABLE `ocupaciones` ENABLE KEYS */;


--
-- Definition of table `pagosfijos`
--

DROP TABLE IF EXISTS `pagosfijos`;
CREATE TABLE `pagosfijos` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `IDEmpeno` int(11) DEFAULT '0',
  `NumPago` int(11) DEFAULT '0',
  `Cancelado` tinyint(1) DEFAULT '0',
  `Vencimiento` date DEFAULT NULL,
  `Pago` double(15,5) DEFAULT '0.00000',
  `Interes` double(15,5) DEFAULT '0.00000',
  `Almacenaje` double(15,5) DEFAULT '0.00000',
  `Seguro` double(15,5) DEFAULT '0.00000',
  `Iva` double(15,5) DEFAULT '0.00000',
  `Amortizacion` double(15,5) DEFAULT '0.00000',
  `Moratorios` double(15,5) DEFAULT '0.00000',
  `Bonificacion` double(15,5) DEFAULT '0.00000',
  `Saldo` double(15,5) DEFAULT '0.00000',
  `Pagado` int(1) DEFAULT '0',
  `Movimiento` int(10) DEFAULT '0',
  `FechaMovimiento` datetime DEFAULT NULL,
  `FolioRecibo` int(11) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `Efectivo` double(15,5) DEFAULT '0.00000',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `pagosfijos`
--

/*!40000 ALTER TABLE `pagosfijos` DISABLE KEYS */;
/*!40000 ALTER TABLE `pagosfijos` ENABLE KEYS */;


--
-- Definition of table `parametros`
--

DROP TABLE IF EXISTS `parametros`;
CREATE TABLE `parametros` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Datos` date DEFAULT NULL,
  `PrestamoAvaluo` double(15,5) DEFAULT '0.00000',
  `PrestamoAvaluoAutos` double(15,5) DEFAULT '0.00000',
  `PrestamoAvaluoElec` double(15,5) DEFAULT '0.00000',
  `Almacenaje` double(15,5) DEFAULT '0.00000',
  `Seguro` double(15,5) DEFAULT '0.00000',
  `GtosVenta` double(15,5) DEFAULT '0.00000',
  `Comision` double(15,5) DEFAULT NULL,
  `IVA` double(15,5) DEFAULT '0.00000',
  `Negociacion` double(15,5) DEFAULT '0.00000',
  `Operacion` double(15,5) DEFAULT '0.00000',
  `PagoMinimo` double(15,5) DEFAULT '0.00000',
  `DiasEnajenacion` int(10) DEFAULT '0',
  `VenApartados` int(10) DEFAULT '0',
  `EngancheApartados` int(10) DEFAULT '0',
  `IvaVentas` double(15,5) DEFAULT '0.00000',
  `DiasGracia` int(10) DEFAULT '0',
  `DiasGraciaAutos` int(10) DEFAULT '0',
  `PolizaSeguro` varchar(15) DEFAULT NULL,
  `FechaExpedicion` date DEFAULT NULL,
  `Aseguradora` varchar(50) DEFAULT NULL,
  `ImportePerdida` double(15,5) DEFAULT '0.00000',
  `Notas` varchar(250) DEFAULT NULL,
  `VenDemasia` int(5) DEFAULT '0',
  `ImporteAutorizacion` double(15,5) DEFAULT '0.00000',
  `DiasGraciaApa` int(10) DEFAULT '0',
  `Gerente` varchar(250) DEFAULT NULL,
  `CalidadEx` double(15,2) DEFAULT NULL,
  `CalidadB` double(15,2) DEFAULT NULL,
  `CalidadR` double(15,2) DEFAULT NULL,
  `CalidadM` double(15,2) DEFAULT NULL,
  `Centenario` double(15,5) DEFAULT '0.00000',
  `DescuentoVentas` double(15,5) DEFAULT '0.00000',
  `TipoCambioOnza` double(15,5) DEFAULT '0.00000',
  `PrestamoAvaluoDiamante` double(15,5) DEFAULT NULL,
  `8K` double(15,5) DEFAULT '0.00000',
  `Venta8K` double(15,5) DEFAULT '0.00000',
  `10K` double(15,5) DEFAULT '0.00000',
  `Venta10K` double(15,5) DEFAULT '0.00000',
  `14K` double(15,5) DEFAULT '0.00000',
  `Venta14K` double(15,5) DEFAULT '0.00000',
  `18K` double(15,5) DEFAULT '0.00000',
  `Venta18K` double(15,5) DEFAULT '0.00000',
  `22K` double(15,5) DEFAULT '0.00000',
  `Venta22K` double(15,5) DEFAULT '0.00000',
  `24K` double(15,5) DEFAULT '0.00000',
  `Venta24K` double(15,5) DEFAULT '0.00000',
  `LimiteInferior` double(15,5) DEFAULT '0.00000',
  `LimiteSuperior` double(15,5) DEFAULT '0.00000',
  `LimiteInferiorAutos` double(15,5) DEFAULT '0.00000',
  `LimiteSuperiorAutos` double(15,5) DEFAULT '0.00000',
  `DescuentoPagosFijos` double(15,5) DEFAULT '0.00000',
  `Limite1` double(15,5) DEFAULT '0.00000',
  `Limite2` double(15,5) DEFAULT '0.00000',
  `VenAlmoneda` int(2) DEFAULT '0',
  `AbonoMinimo` double(15,5) DEFAULT '0.00000',
  `ImpresoraDefault` varchar(60) DEFAULT NULL,
  `DiasPenaliza` int(2) DEFAULT '0',
  `PrecioAutos` double(15,5) DEFAULT '0.00000',
  `Cat` double(15,5) DEFAULT '0.00000',
  `IntAnual` double(15,5) DEFAULT '0.00000',
  `AlmAnual` double(15,5) DEFAULT '0.00000',
  `CodProfeco` varchar(80) DEFAULT 'EN TRÁMITE',
  `PorEnajenados` double(15,5) DEFAULT '0.00000',
  `PrestamoVerde` double(15,5) DEFAULT '0.00000',
  `PrestamoAmarillo` double(15,5) DEFAULT '0.00000',
  `PrestamoRojo` double(15,5) DEFAULT '0.00000',
  `HorarioSucursal` varchar(250) DEFAULT NULL,
  `PuntosTarjeta` int(10) unsigned DEFAULT '0',
  `Version` varchar(20) DEFAULT NULL,
  `IDActividadVulnerable` int(10) DEFAULT '0',
  `IdTipoGiroMercantil` int(10) DEFAULT '0',
  `IDTipoMonedaLocal` int(10) DEFAULT '0',
  `PrestamoCheque` double(15,5) DEFAULT '0.00000',
  `CompraCheque` double(15,5) DEFAULT '0.00000',
  `ImporteSalario` double(15,5) DEFAULT '0.00000',
  `ImporteUdi` double(15,6) DEFAULT '0.000000',
  `NumConstancia` varchar(30) DEFAULT '',
  `RutaArchivosXML` varchar(250) DEFAULT '',
  `ImporteVSMPrestamos` int(10) DEFAULT '1605',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `parametros`
--

/*!40000 ALTER TABLE `parametros` DISABLE KEYS */;
INSERT INTO `parametros` (`ID`,`Datos`,`PrestamoAvaluo`,`PrestamoAvaluoAutos`,`PrestamoAvaluoElec`,`Almacenaje`,`Seguro`,`GtosVenta`,`Comision`,`IVA`,`Negociacion`,`Operacion`,`PagoMinimo`,`DiasEnajenacion`,`VenApartados`,`EngancheApartados`,`IvaVentas`,`DiasGracia`,`DiasGraciaAutos`,`PolizaSeguro`,`FechaExpedicion`,`Aseguradora`,`ImportePerdida`,`Notas`,`VenDemasia`,`ImporteAutorizacion`,`DiasGraciaApa`,`Gerente`,`CalidadEx`,`CalidadB`,`CalidadR`,`CalidadM`,`Centenario`,`DescuentoVentas`,`TipoCambioOnza`,`PrestamoAvaluoDiamante`,`8K`,`Venta8K`,`10K`,`Venta10K`,`14K`,`Venta14K`,`18K`,`Venta18K`,`22K`,`Venta22K`,`24K`,`Venta24K`,`LimiteInferior`,`LimiteSuperior`,`LimiteInferiorAutos`,`LimiteSuperiorAutos`,`DescuentoPagosFijos`,`Limite1`,`Limite2`,`VenAlmoneda`,`AbonoMinimo`,`ImpresoraDefault`,`DiasPenaliza`,`PrecioAutos`,`Cat`,`IntAnual`,`AlmAnual`,`CodProfeco`,`PorEnajenados`,`PrestamoVerde`,`PrestamoAmarillo`,`PrestamoRojo`,`HorarioSucursal`,`PuntosTarjeta`,`Version`,`IDActividadVulnerable`,`IdTipoGiroMercantil`,`IDTipoMonedaLocal`,`PrestamoCheque`,`CompraCheque`,`ImporteSalario`,`ImporteUdi`,`NumConstancia`,`RutaArchivosXML`,`ImporteVSMPrestamos`) VALUES 
 (1,'2005-02-01',80.00000,60.00000,20.50000,2.50000,2.50000,135.00000,25.00000,16.00000,10.00000,6.00000,15.00000,0,1,30,16.00000,0,0,'0','2007-02-27','NO',20.00000,'EN REFRENDO Y DESEMPEÑOS EL INTERES SE COBRA POR DIA!,  MEJORANDO POR TI!!',2,2500.00000,0,'TAPIA SANCHEZ JOSE MANUEL',90.00,90.00,90.00,90.00,1480.00000,40.00000,13.00000,20.50000,54.59000,176.00000,78.97000,220.00000,128.17000,310.00000,122.67000,396.00000,149.00000,475.00000,163.56000,528.00000,3000.00000,6000.00000,100000.00000,220000.00000,0.00000,10000.00000,15000.00000,4,25.00000,'',7,100.00000,0.00000,84.00000,90.00000,'7/002608-2012',100.00000,100.00000,100.00000,100.00000,'EL HORARIO DE SERVICIO AL PÚBLICO DE ESTE ESTABLECIMIENTO ES DE LUNES A VIERNES DE 9:30 A 19:00 HRS Y SABADOS DE 9:30 A 16:00 HRS.',500,'12.2012.17',6,162,1,0.00000,0.00000,64.78000,3.111100,'00001','\\\\VBOXSVR\\proyectos\\LEY LAVADO DINERO\\VERSIONES LIBERADAS\\MR_AYUDON\\Casa de Empeno Mr_Ayudon 2014.7.1\\Casa de Empeno MrAyudon\\AvisosXML',1605);
/*!40000 ALTER TABLE `parametros` ENABLE KEYS */;


--
-- Definition of table `plazos`
--

DROP TABLE IF EXISTS `plazos`;
CREATE TABLE `plazos` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Descripcion` int(5) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=17 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `plazos`
--

/*!40000 ALTER TABLE `plazos` DISABLE KEYS */;
INSERT INTO `plazos` (`ID`,`Descripcion`) VALUES 
 (1,1),
 (2,2),
 (3,3),
 (4,4),
 (5,5),
 (6,6),
 (7,7),
 (8,8),
 (9,9),
 (10,10),
 (11,11),
 (12,12),
 (13,60),
 (14,16),
 (15,30),
 (16,90);
/*!40000 ALTER TABLE `plazos` ENABLE KEYS */;


--
-- Definition of table `precioskilataje`
--

DROP TABLE IF EXISTS `precioskilataje`;
CREATE TABLE `precioskilataje` (
  `ID` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `IDTipo` int(11) DEFAULT '0',
  `IDKilataje` int(11) DEFAULT '0',
  `IDHechura` int(5) DEFAULT '0',
  `Precio` double(15,5) DEFAULT '0.00000',
  `IDRango` int(11) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=357 DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

--
-- Dumping data for table `precioskilataje`
--

/*!40000 ALTER TABLE `precioskilataje` DISABLE KEYS */;
INSERT INTO `precioskilataje` (`ID`,`IDTipo`,`IDKilataje`,`IDHechura`,`Precio`,`IDRango`) VALUES 
 (13,1,14,1,185.41000,0),
 (14,1,14,2,185.41000,0),
 (15,1,14,3,185.41000,0),
 (16,1,14,4,185.41000,0),
 (17,1,1,1,232.18000,0),
 (18,1,1,2,232.18000,0),
 (19,1,1,3,232.18000,0),
 (20,1,1,4,232.18000,0),
 (26,1,2,1,324.60000,0),
 (27,1,2,2,324.60000,0),
 (29,1,2,3,324.60000,0),
 (30,1,2,4,324.60000,0),
 (31,1,3,1,389.74000,0),
 (32,1,3,2,389.74000,0),
 (33,1,3,3,389.74000,0),
 (34,1,3,4,389.74000,0),
 (76,1,21,4,556.78000,0),
 (75,1,21,3,556.78000,0),
 (74,1,21,2,556.78000,0),
 (73,1,21,1,556.78000,0),
 (168,1,22,1,501.10000,0),
 (169,1,22,2,501.10000,0),
 (170,1,22,3,501.10000,0),
 (171,1,22,4,501.10000,0),
 (175,4,18,5,12.00000,1),
 (179,4,18,7,6.00000,1),
 (180,4,19,5,10.00000,1),
 (178,4,18,6,8.00000,1),
 (181,4,19,6,7.00000,1),
 (182,4,19,7,4.00000,1),
 (183,4,20,5,8.00000,1),
 (184,4,20,6,5.00000,1),
 (185,4,20,7,3.00000,1),
 (186,4,18,5,13.00000,2),
 (187,4,18,6,8.00000,2),
 (188,4,18,7,6.00000,2),
 (189,4,19,5,11.00000,2),
 (190,4,19,6,7.00000,2),
 (191,4,19,7,5.00000,2),
 (192,4,20,5,8.00000,2),
 (193,4,20,6,6.00000,2),
 (194,4,20,7,4.00000,2),
 (195,4,18,5,16.00000,3),
 (196,4,18,6,11.00000,3),
 (197,4,18,7,8.00000,3),
 (198,4,19,5,13.00000,3),
 (199,4,19,6,9.00000,3),
 (200,4,19,7,6.00000,3),
 (201,4,20,5,10.00000,3),
 (202,4,20,6,7.00000,3),
 (203,4,20,7,5.00000,3),
 (204,4,18,5,22.00000,4),
 (205,4,18,6,14.00000,4),
 (206,4,18,7,9.00000,4),
 (207,4,19,5,18.00000,4),
 (208,4,19,6,11.00000,4),
 (209,4,19,7,8.00000,4),
 (210,4,20,5,12.00000,4),
 (211,4,20,6,8.00000,4),
 (212,4,20,7,5.00000,4),
 (213,4,18,5,24.00000,5),
 (214,4,18,6,16.00000,5),
 (215,4,18,7,13.00000,5),
 (216,4,19,5,20.00000,5),
 (217,4,19,6,13.00000,5),
 (242,4,19,7,11.00000,5),
 (219,4,20,5,15.00000,5),
 (220,4,20,6,10.00000,5),
 (221,4,20,7,8.00000,5),
 (246,4,19,5,30.00000,7),
 (223,4,18,6,19.00000,6),
 (224,4,18,7,15.00000,6),
 (225,4,19,5,24.00000,6),
 (226,4,19,6,16.00000,6),
 (227,4,19,7,14.00000,6),
 (228,4,20,5,17.00000,6),
 (229,4,20,6,13.00000,6),
 (230,4,20,7,10.00000,6),
 (249,4,20,5,23.00000,7),
 (248,4,19,7,15.00000,7),
 (247,4,19,6,20.00000,7),
 (243,4,18,5,38.00000,7),
 (245,4,18,7,16.00000,7),
 (244,4,18,6,25.00000,7),
 (241,4,18,5,30.00000,6),
 (250,4,20,6,16.00000,7),
 (251,4,20,7,11.00000,7),
 (252,4,18,5,43.00000,8),
 (253,4,18,6,31.00000,8),
 (254,4,18,7,21.00000,8),
 (255,4,19,5,36.00000,8),
 (256,4,19,6,25.00000,8),
 (257,4,19,7,20.00000,8),
 (258,4,20,5,28.00000,8),
 (259,4,20,6,20.00000,8),
 (260,4,20,7,15.00000,8),
 (261,4,18,5,50.00000,9),
 (262,4,18,6,31.00000,9),
 (263,4,18,7,23.00000,9),
 (264,4,19,5,41.00000,9),
 (265,4,19,6,25.00000,9),
 (266,4,19,7,21.00000,9),
 (267,4,20,5,30.00000,9),
 (268,4,20,6,21.00000,9),
 (269,4,20,7,18.00000,9),
 (270,4,18,5,60.00000,10),
 (271,4,18,6,38.00000,10),
 (272,4,18,7,30.00000,10),
 (273,4,19,5,48.00000,10),
 (274,4,19,6,31.00000,10),
 (275,4,19,7,28.00000,10),
 (276,4,20,5,35.00000,10),
 (277,4,20,6,26.00000,10),
 (278,4,20,7,23.00000,10),
 (279,4,18,5,76.00000,11),
 (280,4,18,6,48.00000,11),
 (281,4,18,7,35.00000,11),
 (282,4,19,5,68.00000,11),
 (283,4,19,6,43.00000,11),
 (284,4,19,7,33.00000,11),
 (285,4,20,5,58.00000,11),
 (286,4,20,6,36.00000,11),
 (287,4,20,7,30.00000,11),
 (288,4,18,5,98.00000,12),
 (289,4,18,6,60.00000,12),
 (290,4,18,7,46.00000,12),
 (291,4,19,5,88.00000,12),
 (292,4,19,6,53.00000,12),
 (293,4,19,7,43.00000,12),
 (294,4,20,5,71.00000,12),
 (295,4,20,6,45.00000,12),
 (296,4,20,7,38.00000,12),
 (297,4,18,5,115.00000,13),
 (298,4,18,6,78.00000,13),
 (299,4,18,7,58.00000,13),
 (300,4,19,5,103.00000,13),
 (301,4,19,6,66.00000,13),
 (302,4,19,7,50.00000,13),
 (303,4,20,5,83.00000,13),
 (304,4,20,6,56.00000,13),
 (305,4,20,7,41.00000,13),
 (306,4,18,5,143.00000,14),
 (307,4,18,6,90.00000,14),
 (308,4,18,7,68.00000,14),
 (309,4,19,5,131.00000,14),
 (310,4,19,6,80.00000,14),
 (311,4,19,7,60.00000,14),
 (312,4,20,5,101.00000,14),
 (313,4,20,6,68.00000,14),
 (314,4,20,7,50.00000,14),
 (332,4,18,7,100.00000,16),
 (316,4,18,6,111.00000,15),
 (317,4,18,7,83.00000,15),
 (318,4,19,5,175.00000,15),
 (319,4,19,6,98.00000,15),
 (320,4,19,7,66.00000,15),
 (321,4,18,5,193.00000,15),
 (331,4,18,6,143.00000,16),
 (330,4,18,5,246.00000,16),
 (329,4,20,7,56.00000,15),
 (328,4,20,6,83.00000,15),
 (327,4,20,5,131.00000,15),
 (333,4,19,5,203.00000,16),
 (334,4,19,6,118.00000,16),
 (335,4,19,7,85.00000,16),
 (336,4,20,5,153.00000,16),
 (337,4,20,6,95.00000,16),
 (338,4,20,7,66.00000,16),
 (339,4,18,5,265.00000,17),
 (340,4,18,6,151.00000,17),
 (341,4,18,7,108.00000,17),
 (342,4,19,5,213.00000,17),
 (343,4,19,6,126.00000,17),
 (344,4,19,7,95.00000,17),
 (345,4,20,5,160.00000,17),
 (346,4,20,6,101.00000,17),
 (347,4,20,7,73.00000,17),
 (348,4,18,5,333.00000,18),
 (349,4,18,6,166.00000,18),
 (350,4,18,7,115.00000,18),
 (351,4,19,5,265.00000,18),
 (352,4,19,6,143.00000,18),
 (353,4,19,7,96.00000,18),
 (354,4,20,5,181.00000,18),
 (355,4,20,6,110.00000,18),
 (356,4,20,7,81.00000,18);
/*!40000 ALTER TABLE `precioskilataje` ENABLE KEYS */;


--
-- Definition of table `prendaselec`
--

DROP TABLE IF EXISTS `prendaselec`;
CREATE TABLE `prendaselec` (
  `ID` int(11) NOT NULL AUTO_INCREMENT,
  `IDTipo` int(10) DEFAULT '0',
  `IDMarca` int(10) DEFAULT '0',
  `IDFamilia` int(10) DEFAULT '0',
  `Funciones` varchar(250) DEFAULT NULL,
  `Caracteristicas` varchar(250) DEFAULT NULL,
  `Minimo` double(15,5) DEFAULT '0.00000',
  `Maximo` double(15,5) DEFAULT '0.00000',
  `Modelo` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=514 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `prendaselec`
--

/*!40000 ALTER TABLE `prendaselec` DISABLE KEYS */;
INSERT INTO `prendaselec` (`ID`,`IDTipo`,`IDMarca`,`IDFamilia`,`Funciones`,`Caracteristicas`,`Minimo`,`Maximo`,`Modelo`) VALUES 
 (1,2,1,176,'21\"','PANTALLA PLANA CON CONTROL',500.00000,700.00000,'21\"'),
 (2,2,2,217,'MP3','16 GB',300.00000,500.00000,'16GB'),
 (3,2,3,216,'MPS','32 GB',500.00000,1000.00000,'32 GB'),
 (4,2,4,172,'***','***',600.00000,800.00000,'---'),
 (5,2,5,180,'REPRODUCTOR','CON UN JUEGO Y UN CONTROL',300.00000,500.00000,'2 SLIM'),
 (6,2,5,180,'REPRODUCTOR','1 CONTROL 1 JUEGO',900.00000,1400.00000,'TRES'),
 (7,2,6,180,'REPRODUCTOR','1 CONTROL , 1 JUEGO',800.00000,1000.00000,'360'),
 (8,2,5,180,'VIDEOJUEGO','PORTATIL CON UN JUEGO',400.00000,600.00000,'PSP'),
 (9,2,7,176,'TV','CON CONTROL',250.00000,350.00000,'14\"'),
 (10,2,7,173,'TV','CON CONTROL',300.00000,400.00000,'21\"'),
 (11,2,7,176,'TV','CON CONTROL',450.00000,700.00000,'21\"'),
 (12,2,7,173,'TV','CON CONTROL',400.00000,600.00000,'25\"'),
 (13,2,7,176,'TV','CON CONTROL',500.00000,800.00000,'25\"'),
 (14,2,7,173,'TV','CON CONTROL',600.00000,900.00000,'29\"'),
 (15,2,8,175,'LCD','CON CONTROL',700.00000,1100.00000,'20\"'),
 (16,2,8,175,'TV','CON CONTROL',1250.00000,1600.00000,'26\"'),
 (17,2,8,175,'TV','CON CONTROL',1700.00000,2500.00000,'37\"'),
 (18,2,8,175,'TV','CON CONTROL',2600.00000,3600.00000,'42\"'),
 (19,2,7,176,'TV','CON CONTROL',800.00000,1200.00000,'29\"'),
 (20,2,7,178,'REPRODUCTOR','CON CONTROL',100.00000,200.00000,'RECIENTE'),
 (21,2,7,178,'REPRODUCTOR','CON CONTROL, CABLES Y BATERIA',250.00000,400.00000,'PORTATIL'),
 (22,2,8,218,'AUDIO VIDEO','BOCINAS, WOFER, CONTROL',350.00000,500.00000,'RECIENTE'),
 (23,2,7,172,'AUDIO','CON CONTROL',150.00000,300.00000,'MICROCOMPONETE'),
 (24,2,7,172,'AUDIO MP3','2 BOCINAS, CON CONTROL',400.00000,700.00000,'2 BOCINAS'),
 (25,2,7,172,'MP3','2 BOCINAS, WOFER, CON CONTROL',500.00000,800.00000,'3 BOCINAS'),
 (26,2,7,172,'MP3','4 BOCINAS, WOFER, CONTROL',700.00000,1300.00000,'4 BOCINAS'),
 (27,2,7,169,'FOTOGRAFIA, VIDEO','CARGADOR, USB, MEMORIA',300.00000,400.00000,'8 MPX'),
 (28,2,7,169,'FOTOGRAFIA, VIDEO','CARGADOR, MEMORIA, USB',350.00000,500.00000,'9 MPX O MAS'),
 (29,2,8,170,'MP3','0',100.00000,300.00000,'RECIENTE'),
 (30,2,8,171,'NORMALES','MODELO RECIENTE',100.00000,250.00000,'RECIENTE'),
 (31,2,3,216,'MP3','CABLE USB',500.00000,600.00000,'80 GB VIDEO'),
 (32,2,3,216,'MP3','CON CABLES USB',700.00000,900.00000,'160 GB VIDEO'),
 (33,2,3,216,'MP3, VIDEO','CON CABLES USB',500.00000,700.00000,'16 GB TOUCH'),
 (34,2,7,172,'***','***',1000.00000,1500.00000,'4 BOCINAS 1 SUWOFER'),
 (35,2,3,216,'***','***',300.00000,500.00000,'8 GB'),
 (36,2,9,219,'***','***',1000.00000,2000.00000,'**'),
 (37,2,10,184,'***','***',1500.00000,3500.00000,'MACBOOK PRO'),
 (38,2,5,217,'***','***',300.00000,700.00000,'W 395'),
 (39,2,6,180,'***|','***',400.00000,500.00000,'NEGRO'),
 (40,2,5,180,'***','***',400.00000,500.00000,'XBOX NEGRO'),
 (41,2,7,169,'***','***',300.00000,500.00000,'5.1 MP'),
 (42,2,11,234,'**','**',200.00000,500.00000,'SENCILLO'),
 (43,7,12,235,'***','***',1000.00000,2000.00000,'***'),
 (44,7,13,236,'***','***',200.00000,500.00000,'***'),
 (45,2,14,237,'***','***',1000.00000,1500.00000,'***'),
 (46,2,15,238,'***','***',350.00000,1000.00000,'VARIAS'),
 (47,2,7,173,'***','***',100.00000,250.00000,'3\"'),
 (48,2,7,171,'***','***',200.00000,350.00000,'***'),
 (49,3,16,240,'***','***',900.00000,1500.00000,'***'),
 (50,2,7,169,'***','***',500.00000,800.00000,'10 MP O MAS'),
 (51,9,17,241,'***','***',1000.00000,2000.00000,'***'),
 (52,2,18,219,'***','***',1400.00000,2100.00000,'E520'),
 (53,2,19,217,'***','***',500.00000,800.00000,'5530'),
 (54,2,20,217,'***','***',450.00000,700.00000,'GTB310'),
 (55,2,4,175,'****','****',2800.00000,3500.00000,'32\"'),
 (56,2,21,173,'***','***',200.00000,400.00000,'3\"'),
 (57,2,22,219,'***','***',1000.00000,1600.00000,'NEURON LT 3G'),
 (58,2,5,172,'***','***',1800.00000,2300.00000,'HCR-GTR88 4 BOCINAS 2 BUFERS'),
 (59,2,10,242,'***','***',600.00000,1200.00000,'2 G'),
 (60,2,23,219,'***','***',1500.00000,2500.00000,'***'),
 (61,2,24,237,'***','***',1000.00000,2500.00000,'***'),
 (62,2,7,178,'***','***',350.00000,550.00000,'PORTATIL'),
 (63,2,25,219,'***','***',1500.00000,3000.00000,'***'),
 (64,2,26,180,'***','***',700.00000,1000.00000,'RECIENTE'),
 (65,2,27,237,'***','***',1500.00000,3000.00000,'TOUCH SMART 300'),
 (66,2,7,173,'***','***',200.00000,350.00000,'14\"'),
 (67,2,28,219,'***','***',1500.00000,2700.00000,'VAIO'),
 (68,2,20,217,'***','***',400.00000,900.00000,'DJ GT-7600L'),
 (69,2,29,217,'***','***',150.00000,400.00000,'A1200'),
 (70,2,30,219,'***','***',1000.00000,2500.00000,'ASPIRE'),
 (71,2,31,237,'***','***',700.00000,1500.00000,'***'),
 (72,2,10,242,'***','***',700.00000,1300.00000,'32 GB'),
 (73,2,11,219,'***','***',1500.00000,2500.00000,'***'),
 (74,2,5,238,'***','***',400.00000,750.00000,'HDD'),
 (75,9,32,243,'***','***',150.00000,400.00000,'INDUSTRIAL'),
 (76,2,5,238,'***','***',500.00000,1000.00000,'MINIDVD'),
 (77,2,33,244,'***','***',300.00000,500.00000,'NANO 16 GB'),
 (78,2,34,217,'***','***',100.00000,500.00000,'***'),
 (79,2,4,217,'***','***',300.00000,800.00000,'***'),
 (80,2,30,219,'***','***',750.00000,1500.00000,'ASPIRE 5040'),
 (81,2,35,219,'***','***',1500.00000,2500.00000,'***'),
 (82,2,4,218,'GRABA DE DISCO A MEMORIA','***',400.00000,700.00000,'GDDAM'),
 (83,7,36,236,'***','***',300.00000,600.00000,'2010'),
 (84,2,28,218,'***','***',400.00000,700.00000,'2010'),
 (85,2,6,180,'***','***',800.00000,1500.00000,'2009'),
 (86,7,37,235,'***','***',300.00000,1000.00000,'TECLADO'),
 (87,2,38,217,'***','***',300.00000,1100.00000,'***'),
 (88,2,4,217,'***','***',500.00000,900.00000,'***'),
 (89,2,5,170,'***','***',200.00000,400.00000,'EXPLOD'),
 (90,2,27,219,'***','***',1500.00000,2500.00000,'HP ROKER'),
 (91,9,39,245,'***','***',500.00000,1000.00000,'***'),
 (92,2,20,217,'***','***',500.00000,1000.00000,'***'),
 (93,2,20,175,'*****','******',3000.00000,4000.00000,'52\"'),
 (94,2,5,180,'***','***',1000.00000,1600.00000,'3'),
 (95,2,40,237,'***','***',2000.00000,3500.00000,'***'),
 (96,9,17,246,'***','***',100.00000,300.00000,'***'),
 (97,9,41,247,'***','***',400.00000,800.00000,'***'),
 (98,2,20,175,'***','***',1300.00000,2000.00000,'33\"'),
 (99,2,5,172,'***','***',700.00000,900.00000,'3 BOCINAS'),
 (100,2,7,178,'***','***',150.00000,250.00000,'USB'),
 (101,3,16,248,'***','***',1000.00000,2000.00000,'***'),
 (102,3,16,248,'***','***',1000.00000,2500.00000,'***'),
 (103,2,7,176,'**','**',500.00000,800.00000,'21\" MODERNA'),
 (104,2,8,175,'***','***',1000.00000,1800.00000,'27\" A 1 MES'),
 (105,2,28,175,'VARIOS','VARIOS',1550.00000,2200.00000,'DE 26\"|'),
 (106,2,5,178,'VARIAS','VARIAS',150.00000,200.00000,'DVD'),
 (107,2,15,169,'VARIAS','VARIAS',200.00000,300.00000,'DE 8MP'),
 (108,2,20,249,'VARIAS','VARIAS',1000.00000,1500.00000,'*'),
 (109,2,20,217,'VARIAS','VARIAS',300.00000,400.00000,'SAMSUNG'),
 (110,2,11,219,'VARIAS','VARIAS',800.00000,1500.00000,'COMPAQ MINI'),
 (111,2,31,175,'VARIAS','VARIAS',1550.00000,1950.00000,'26\"'),
 (112,2,42,178,'**','**',150.00000,200.00000,'DVD'),
 (113,2,4,175,'VARIAS','VARIAS',1150.00000,1500.00000,'22\"'),
 (114,2,27,219,'VARIAS','VARIAS',1000.00000,2000.00000,'NOTEBOOK'),
 (115,2,31,173,'VARIOS','VARIAS',450.00000,700.00000,'P/PLANA'),
 (116,2,20,173,'VARIAS','VARIAS',400.00000,600.00000,'SAMSUNG'),
 (117,2,5,249,'VARIAS','VARIAS',700.00000,900.00000,'SONY'),
 (118,2,18,237,'VARIAS','VARIAS',1500.00000,2000.00000,'DE ESCRITORIO'),
 (119,2,26,180,'VARIAS','VARIAS',500.00000,1100.00000,'WII'),
 (120,2,25,219,'VARIAS','VARIAS',500.00000,1000.00000,'AMD'),
 (121,2,43,219,'VARIAS','VARIAS',500.00000,1000.00000,'MINI LAPTOP'),
 (122,2,27,219,'VARIAS','VARIAS',500.00000,1000.00000,'MINI LAPTOP'),
 (123,2,15,173,'VARIAS','VARIAS',450.00000,600.00000,'21\"'),
 (124,2,4,178,'VARIAS','VARIAS',150.00000,200.00000,'DVD'),
 (125,2,15,249,'VARIAS','VARIAS',700.00000,900.00000,'STEREO'),
 (126,2,24,237,'VARIAS','VARIAS',1000.00000,1700.00000,'DE ESCRITORIO'),
 (127,2,5,169,'VARIAS','VARIAS',200.00000,300.00000,'7MP'),
 (128,2,11,237,'VARIAS','VARIAS',1000.00000,1700.00000,'ESCRITORIO'),
 (129,2,4,175,'VARIAS','VARIAS',3100.00000,4000.00000,'42\"'),
 (130,2,28,238,'VARIAS','VARIAS',500.00000,950.00000,'DE VIDEO'),
 (131,2,44,173,'VARIAS','VARIAS',850.00000,1100.00000,'P/PLANA'),
 (132,2,28,219,'VARIAS','VARIAS',1000.00000,1500.00000,'VAIO'),
 (133,2,20,176,'VARIAS','VARIAS',450.00000,600.00000,'DE 21\"'),
 (134,2,31,249,'VARIAS','VARIAS',700.00000,1300.00000,'STEREO'),
 (135,2,1,173,'VARIAS','VARIAS',200.00000,400.00000,'20\"'),
 (136,2,28,177,'VARIAS','VARIAS',1500.00000,2600.00000,'32\"'),
 (137,2,5,178,'VARIAS','VARIAS',150.00000,200.00000,'DVD'),
 (138,2,20,178,'VARIAS','VARIAS',150.00000,200.00000,'DVD'),
 (139,2,31,250,'VARIAS','VARIAS',500.00000,750.00000,'MINICOMPONENTE'),
 (140,2,18,219,'VARIAS','VARIAS',800.00000,1500.00000,'MINI LAPTOP'),
 (141,2,28,217,'VARIAS','VARIAS',200.00000,500.00000,'SONY ERICSSON'),
 (142,2,27,219,'VARIAS','VARIAS',1000.00000,2500.00000,'PROBOOK'),
 (143,2,15,176,'VARIAS','VARIAS',400.00000,600.00000,'21\"'),
 (144,2,31,176,'VARIAS','VARIAS',400.00000,600.00000,'21\"'),
 (145,2,27,219,'VARIAS','VARIAS',500.00000,1500.00000,'MINI LAPTOP'),
 (146,2,1,178,'VARIAS','VARIAS',300.00000,550.00000,'PORTATIL'),
 (147,2,15,250,'VARIAS','VARIAS',500.00000,850.00000,'MINICOMPONENTE'),
 (148,2,45,217,'VARIAS','VARIAS',400.00000,600.00000,'BLACBERRY'),
 (149,2,28,249,'VARIAS','VARIAS',600.00000,1300.00000,'GENEZI'),
 (150,2,42,225,'VARIAS','VARIAS',350.00000,500.00000,'PORTATIL'),
 (151,2,1,173,'VARIAS','VARIAS',250.00000,550.00000,'20\"'),
 (152,2,31,176,'VARIAS','VARIAS',250.00000,400.00000,'14\"'),
 (153,2,15,249,'VARIAS','VARIAS',900.00000,1400.00000,'STEREO'),
 (154,2,20,182,'VARIAS','VARIAS',500.00000,1000.00000,'VIDEO'),
 (155,2,33,244,'VARIAS','VARIAS',500.00000,800.00000,'8GB'),
 (156,2,28,169,'VARIAS','VARIAS',500.00000,700.00000,'12 MP'),
 (157,2,46,219,'VARIAS','VARIAS',600.00000,1200.00000,'MINI'),
 (158,2,2,242,'VARIAS','VARIAS',500.00000,950.00000,'16GB'),
 (159,2,28,180,'VARIAS','VARIAS',400.00000,500.00000,'SLIM 2'),
 (160,2,27,219,'VARIAS','VARIAS',1000.00000,3000.00000,'LAPTOP'),
 (161,2,6,180,'VARIAS','VARIAS',800.00000,1500.00000,'ELITE'),
 (162,2,15,175,'*','*',500.00000,1450.00000,'22\"'),
 (163,2,47,175,'VARIAS','VARIAS',850.00000,1050.00000,'29\"'),
 (164,2,20,249,'VARIAS','VARIAS',500.00000,1300.00000,'STEREO'),
 (165,2,28,169,'VARIAS','VARIAS',200.00000,350.00000,'8MP'),
 (166,2,48,175,'VARIAS','VARIAS',850.00000,1050.00000,'20\"'),
 (167,2,4,225,'*','*',350.00000,500.00000,'PORTATIL'),
 (168,2,20,217,'VARIAS','VARIAS',450.00000,600.00000,'TOUCH'),
 (169,2,20,175,'**','**',850.00000,1050.00000,'19\"'),
 (170,2,20,176,'VARIAS','VARIAS',150.00000,350.00000,'12\"'),
 (171,2,20,176,'VARIAS','VARIAS',150.00000,350.00000,'14\"'),
 (172,2,19,217,'VARIAS','VARIAS',300.00000,450.00000,'VARIOS'),
 (173,2,28,180,'VAIRAS','VARIAS',550.00000,750.00000,'PSP'),
 (174,2,20,176,'VARIAS','VARIAS',350.00000,700.00000,'21\"'),
 (175,2,47,175,'VARIAS','VARIAS',850.00000,1050.00000,'19\"'),
 (176,2,9,219,'VARIAS','VARIAS',1000.00000,1300.00000,'MINI LAPTOP'),
 (177,2,28,225,'VARIAS','VARIAS',400.00000,500.00000,'PORTATIL'),
 (178,2,5,250,'VARIAS','VARIAS',450.00000,550.00000,'MINICOMPONENTE'),
 (179,2,20,238,'VARIAS','VARIAS',600.00000,900.00000,'MINIDVD'),
 (180,8,49,236,'VARIAS','VARIAS',300.00000,550.00000,'KAIZER 200'),
 (181,2,4,173,'VARIAS','VARIAS',650.00000,950.00000,'29\"'),
 (182,2,19,217,'VARIAS','VARIAS',450.00000,550.00000,'E63'),
 (183,2,30,219,'VARIAS','VARIAS',500.00000,1500.00000,'MINI'),
 (184,2,50,169,'VARIAS','VARIAS',450.00000,600.00000,'12MP'),
 (185,2,20,175,'VARIAS','VARIAS',3000.00000,4500.00000,'42\"'),
 (186,2,1,176,'VARIAS','VARIAS',450.00000,600.00000,'21\"'),
 (187,2,46,219,'VARIAS','VARIAS',500.00000,1300.00000,'MINI LAPTOP'),
 (188,2,15,173,'VARIAS','VARIAS',650.00000,900.00000,'29\"'),
 (189,2,5,217,'VARIAS','VARIAS',550.00000,750.00000,'SONY ERICSSON'),
 (190,2,6,180,'VARIAS','VARIAS',800.00000,1100.00000,'360'),
 (191,2,51,237,'VARIAS','VARIAS',1500.00000,2000.00000,'ESRITORIO'),
 (192,9,52,243,'VARIAS','VARIAS',250.00000,350.00000,'VARIOS'),
 (193,2,18,237,'VARIAS','VARIAS',1000.00000,2500.00000,'LAPTOP'),
 (194,2,24,219,'VARIAS','VARIAS',1000.00000,3000.00000,'ASPIRE'),
 (195,2,43,219,'VARIAS','VARIAS',500.00000,1300.00000,'MINI LAPTOP'),
 (196,2,18,219,'VARIAS','VARIAS',1000.00000,2300.00000,'LAPTOP'),
 (197,2,16,217,'VARIAS','VARIAS',200.00000,400.00000,'***'),
 (198,2,5,250,'VARIAS','VARIAS',550.00000,750.00000,'MINICOMPONENTE'),
 (199,2,15,176,'VARIAS','VARIAS',450.00000,600.00000,'21\"'),
 (200,2,42,176,'VARIAS','VARIAS',450.00000,600.00000,'21\"'),
 (201,2,5,175,'VARIAS','VARIAS',1900.00000,2400.00000,'\"\"\"\"\"'),
 (202,2,31,225,'VARIAS','VARIAS',400.00000,600.00000,'PORTATIL'),
 (203,2,46,169,'VARIAS','VARIAS',400.00000,600.00000,'12 MP'),
 (204,2,15,178,'VARIAS','VARIAS',150.00000,200.00000,'DVD'),
 (205,2,15,177,'VARIAS','VARIAS',1000.00000,2700.00000,'42\"'),
 (206,2,1,178,'VARIAS','VARIAS',150.00000,200.00000,'DVD'),
 (207,2,5,175,'VARIAS','VARIAS',1900.00000,2400.00000,'32\"'),
 (208,2,5,182,'VARIAS','VARIAS',600.00000,900.00000,'VIDEO'),
 (209,2,20,219,'VARIAS','VARIAS',1500.00000,2500.00000,'LAPTOP'),
 (210,2,20,250,'VARIAS','VARIAS',500.00000,800.00000,'MINICOMPONENTE'),
 (211,2,20,169,'VARIAS','VARIAS',200.00000,300.00000,'10MP'),
 (212,2,5,249,'VARIAS','VARIAS',900.00000,1500.00000,'GENEZI'),
 (213,2,15,177,'VARIAS','VARIAS',2700.00000,3000.00000,'42\"'),
 (214,2,4,176,'VARIAS','VARIAS',850.00000,1000.00000,'29\"'),
 (215,2,15,182,'VARIAS','VARIAS',600.00000,900.00000,'VIDEOCAMARA'),
 (216,2,27,219,'VARIAS','VARIAS',2000.00000,3000.00000,'SDMS/PRO MMC'),
 (217,2,5,169,'VARIAS','VARIAS',300.00000,500.00000,'10 MP'),
 (218,2,20,177,'VARIAS','VARIAS',2700.00000,3600.00000,'42\"'),
 (219,2,15,176,'VARIAS','VARIAS',850.00000,1000.00000,'29\"'),
 (220,2,31,175,'VARIAS','VARIAS',700.00000,1050.00000,'19\"'),
 (221,2,31,177,'VARIAS','VARIAS',3200.00000,3900.00000,'PLASMA'),
 (222,2,15,175,'VARIAS','VARIAS',350.00000,600.00000,'21\"'),
 (223,2,20,175,'VARIAS','VARIAS',1100.00000,1350.00000,'22\"'),
 (224,2,42,173,'VARIAS','VARIAS',350.00000,450.00000,'21\"'),
 (225,2,27,219,'VARIOS','VARIOS',2000.00000,3000.00000,'NOTEBOOK'),
 (226,2,19,217,'VARIAS','VARIAS',350.00000,500.00000,'VARIOS'),
 (227,2,53,178,'VARIAS','VARIAS',150.00000,200.00000,'DVD USB'),
 (228,9,32,232,'VARIOS','VARIOS',100.00000,150.00000,'CALA- A2'),
 (229,9,32,232,'VARIOS','VARIOS',100.00000,200.00000,'VARIAS'),
 (230,2,15,176,'VARIOS','VARIOS',550.00000,700.00000,'27\"'),
 (231,2,35,176,'VARIOS','VARIOS',850.00000,1000.00000,'32\"'),
 (232,2,44,225,'VARIOS','VARIOS',250.00000,300.00000,'VARIOS'),
 (233,7,50,169,'VARIOS','VARIOS',350.00000,500.00000,'VARIOS'),
 (234,2,30,219,'VARIOS','VARIOS',1200.00000,1600.00000,'MINILAPTOP'),
 (235,2,9,219,'VARIOS','VARIOS',2000.00000,3000.00000,'VARIOS'),
 (236,2,54,176,'VARIOS','VARIOS',850.00000,1000.00000,'VARIOS'),
 (237,2,14,219,'VARIOS','VARIOS',1000.00000,13000.00000,'VARIOS'),
 (238,2,20,169,'VARIOS','VARIOS',450.00000,600.00000,'12 MP'),
 (239,2,42,249,'VARIOS','VARIOS',1100.00000,1300.00000,'VARIOS'),
 (240,2,20,219,'VARIOS','VARIOS',1000.00000,1300.00000,'MINILAPTOP'),
 (241,2,14,175,'VARIOS','VARIOS',1700.00000,2200.00000,'32\"'),
 (242,2,5,217,'VARIOS','VARIOS',750.00000,1000.00000,'SONY ERIKSON'),
 (243,2,14,176,'VAIAS','VARIAS',650.00000,850.00000,'MITSUI'),
 (244,2,20,219,'VARIOS','VARIOS',1000.00000,1500.00000,'MINILAPTOP'),
 (245,2,29,217,'VARIOS','VARIOS',750.00000,1000.00000,'MB300'),
 (246,2,28,169,'VARIOS','VARIOS',450.00000,650.00000,'14 MP'),
 (247,2,33,242,'VARIOS','VARIOS',500.00000,750.00000,'8G'),
 (248,2,2,217,'VARIOS','VARIOS',750.00000,1000.00000,'3G 16G'),
 (249,2,28,180,'VARIOS','VARIOS',550.00000,800.00000,'PSP'),
 (250,2,20,175,'VARIOS','VARIOS',1450.00000,1800.00000,'26\"'),
 (251,2,46,169,'VARIAS','VARIAS',350.00000,500.00000,'14MP'),
 (252,2,5,175,'VARIOS','VARIOS',1100.00000,1350.00000,'22\"'),
 (253,2,31,176,'VARIAS','VARIAS',2700.00000,3300.00000,'XXX'),
 (254,2,4,177,'VARIAS','VARIAS',2700.00000,3300.00000,'42\"'),
 (255,2,42,250,'VARIAS','VARIAS',650.00000,750.00000,'XXX'),
 (256,2,31,218,'VARIOS','VARIOS',450.00000,600.00000,'VARIOS'),
 (257,2,25,219,'VAIOR','VARIOS',1100.00000,1300.00000,'MINILAPTOP'),
 (258,2,16,251,'00','00',200.00000,500.00000,'00'),
 (259,2,53,250,'VARIOS','VARIOS',700.00000,900.00000,'VARIOS'),
 (260,2,53,249,'VARIOS','VARIOS',700.00000,900.00000,'VARIOS'),
 (261,2,4,219,'VARIOS','VARIOS',1200.00000,1800.00000,'LAPTOP'),
 (262,2,28,180,'VARIAS','VARIAS',1250.00000,1550.00000,'PS 3'),
 (263,2,5,180,'VARIO','VARIOS',1250.00000,1550.00000,'PS 3'),
 (264,2,14,169,'VARIOS','VARIOS',450.00000,600.00000,'12 MP'),
 (265,2,20,217,'VARIOS','VARIOS',450.00000,550.00000,'VARIOS'),
 (266,2,4,217,'VARIAS','VARIAS',200.00000,300.00000,'GW300'),
 (267,2,28,175,'VARIAS','VARIAS',2900.00000,3450.00000,'40\"'),
 (268,2,2,217,'VARIOS','VARIOS',550.00000,750.00000,'8 G'),
 (269,2,28,219,'VARIOS','VARIOS',1500.00000,3000.00000,'VAIO'),
 (270,2,20,175,'VARIOS','VARIOS',1200.00000,2000.00000,'VARIOS'),
 (271,2,20,175,'VARIOS','VARIOSD',2200.00000,3000.00000,'32\"'),
 (272,2,31,171,'VARIAS','VARIAS',200.00000,350.00000,'GRANDE'),
 (273,2,11,219,'VARIAS','VARIAS',1800.00000,3000.00000,'LAPTOP'),
 (274,2,20,169,'VARIOS','VARIOS',300.00000,600.00000,'10 MP'),
 (275,2,14,171,'VARIAS','VARIAS',200.00000,500.00000,'MABE'),
 (276,2,5,180,'VARIOS','VARIOS',850.00000,1000.00000,'PSP'),
 (277,2,27,238,'VARIOS','VARIOS',700.00000,1200.00000,'VARIOS'),
 (278,2,25,219,'VARIAS','VARIAS',1100.00000,1600.00000,'MINI LAPTOP'),
 (279,8,55,224,'00','00',2000.00000,3000.00000,'2006'),
 (280,2,5,180,'VARIOS','VARIOS',300.00000,700.00000,'PLAYSTATION'),
 (281,2,19,217,'VARIOS','VARIOS',400.00000,800.00000,'X6'),
 (282,2,7,178,'VARIOS','VARIOS',50.00000,300.00000,'VARIOS'),
 (283,2,11,237,'VARIOS','VARIOS',1000.00000,2000.00000,'VARIOS'),
 (284,2,1,225,'VARIOS','VARIOS',350.00000,500.00000,'VARIOS'),
 (285,2,14,171,'00','00',100.00000,300.00000,'00'),
 (286,2,2,217,'VARIOS','VARIOS',700.00000,1200.00000,'VARIOS'),
 (287,2,11,219,'VARIAS','VARIAS',1200.00000,1800.00000,'LAPTOP'),
 (288,2,50,169,'VARIAS','VARIAS',400.00000,550.00000,'10MP'),
 (289,2,1,250,'VARIOS','VARIOS',650.00000,900.00000,'VARIOS'),
 (290,2,31,217,'VARIAS','VARIAS',150.00000,300.00000,'GD330'),
 (291,2,31,252,'VARIAS','VARIAS',3000.00000,3700.00000,'42\"'),
 (292,2,4,175,'VARIAS','VARIAS',3700.00000,4000.00000,'50\"'),
 (293,2,56,169,'VARIAS','VARIAS',400.00000,500.00000,'10MP'),
 (294,2,57,253,'VARIAS','VARIAS',2000.00000,3200.00000,'16GB'),
 (295,2,20,217,'VARIAS','VARIAS',500.00000,1000.00000,'GALAXGT-S5670L'),
 (296,2,27,254,'VARIAS','VARIAS',700.00000,1100.00000,'V5061U'),
 (297,2,58,169,'VARIAS','VARIAS',400.00000,600.00000,'10MP'),
 (298,2,59,176,'VARIAS','VARIAS',350.00000,500.00000,'21\"'),
 (299,2,10,244,'VARIAS','VARIAS',500.00000,900.00000,'8GB'),
 (300,2,31,171,'VARIAS','VARIAS',200.00000,300.00000,'MEDIANO'),
 (301,2,6,180,'VARIAS','VARIAS',1300.00000,1900.00000,'SLIM KINECT'),
 (302,2,35,175,'VARIAS','VARIAS',2250.00000,2800.00000,'37\"'),
 (303,2,60,169,'VARIAS','VARIAS',400.00000,500.00000,'10MP'),
 (304,2,57,253,'VARIAS','VARIAS',1000.00000,3000.00000,'32GB'),
 (305,2,8,250,'VARIOS','VARIOS',500.00000,1000.00000,'VARIOS'),
 (306,2,23,219,'VARIOS','VARIOS',1000.00000,3500.00000,'VERIOS'),
 (307,2,45,217,'VARIOS','VARIOS',300.00000,1000.00000,'VARIOS'),
 (308,2,15,177,'VAEIAS','VARIAS',1500.00000,3500.00000,'PLASMA'),
 (309,2,4,217,'VARIOS','VARIOS',500.00000,1000.00000,'VARIOS'),
 (310,2,16,182,'VARIOS','VARIOS',500.00000,2000.00000,'CANON'),
 (311,2,4,249,'VARIOS','VARIOS',500.00000,1500.00000,'VARIOS'),
 (312,2,2,217,'VARIOS','VARIOS',500.00000,2000.00000,'VARIOS'),
 (313,2,10,253,'VARIOS','VARIOS',1000.00000,3500.00000,'IPAD'),
 (314,2,16,178,'VARIOS','VARIOS',100.00000,600.00000,'BLUERAY'),
 (315,2,7,177,'VARIOS','VARIOS',1000.00000,5000.00000,'VARIOS'),
 (316,2,2,217,'VARIAS','VARIAS',500.00000,2500.00000,'32G'),
 (317,2,15,175,'VARIAS','VARIAS',1900.00000,2250.00000,'32\"'),
 (318,2,3,244,'VARIOS','VARIOS',950.00000,1200.00000,'4 GENERACION'),
 (319,2,14,225,'VARIOS','VARIOS',100.00000,500.00000,'VARIOS'),
 (320,2,8,252,'VARIAS','VARIAS',1000.00000,3000.00000,'22\"'),
 (321,2,25,219,'VARIAS','VARIAS',1500.00000,2500.00000,'VARIOS'),
 (322,2,61,219,'VARIOS','VARIOS',1500.00000,2000.00000,'VARIOS'),
 (323,2,20,175,'VARIAS','VARIAS',1000.00000,3500.00000,'37\"'),
 (324,2,14,176,'VARIAS','VARIAS',800.00000,1000.00000,'29\"'),
 (325,2,4,175,'VARIAS','VARIAS',2000.00000,3000.00000,'XXX'),
 (326,2,20,175,'VARIAS','VARIAS',1500.00000,2000.00000,'26\"'),
 (327,2,6,180,'VARIOS','VARIOS',1000.00000,2000.00000,'SLIM KINECT'),
 (328,2,16,176,'VARIOS','VARIOS|',100.00000,600.00000,'VARIOS'),
 (329,2,20,217,'.','.',400.00000,800.00000,'GT-S5830L'),
 (330,2,62,217,'.','.',200.00000,250.00000,'VARIOS'),
 (331,2,63,219,'.','.',2000.00000,3000.00000,'631'),
 (332,2,64,217,'.','.',200.00000,300.00000,'W105A'),
 (333,2,27,219,'.','.',1000.00000,2000.00000,'MINILAP'),
 (334,2,45,217,'VARIOS','VARIOS',500.00000,1600.00000,'VARIOS'),
 (335,2,19,217,'.','.',1000.00000,1300.00000,'N8'),
 (336,2,28,180,'VARIOS','VARIOS',1000.00000,2000.00000,'PS 3'),
 (337,2,65,176,'.','.',400.00000,500.00000,'21\"'),
 (338,2,28,172,'.','.',500.00000,750.00000,'GENEZI'),
 (339,2,66,172,'.','.',1000.00000,1300.00000,'MASH'),
 (340,2,4,172,'.','.',700.00000,900.00000,'LG'),
 (341,2,10,217,'VARIOS','VARIOS',1000.00000,2600.00000,'IPHONE'),
 (342,2,28,169,'.','.',250.00000,450.00000,'16 MP'),
 (343,2,44,225,'.','.',300.00000,450.00000,'PD707M'),
 (344,2,4,172,'.','.',1000.00000,1500.00000,'MCV904AOU'),
 (345,2,4,178,'.','.',150.00000,200.00000,'DV692H'),
 (346,2,20,217,'.','.',1000.00000,1500.00000,'GALAXY 19000T'),
 (347,2,67,218,'.','.',400.00000,600.00000,'SPTHD5W'),
 (348,2,66,172,'.','.',100.00000,1400.00000,'SA-AK980'),
 (349,2,4,178,'.','.',150.00000,200.00000,'DV497H'),
 (350,2,31,177,'.','.',2000.00000,3000.00000,'42LK450'),
 (351,2,29,255,'.','.',1000.00000,1500.00000,'MZ604'),
 (352,2,20,182,'0','0',300.00000,700.00000,'SMX-F54BN'),
 (353,2,24,219,'.','.',1500.00000,2000.00000,'P5WE6'),
 (354,2,20,169,'.','.',200.00000,350.00000,'PL100'),
 (355,2,4,178,'.','.',100.00000,200.00000,'DV552'),
 (356,2,53,178,'.','.',100.00000,200.00000,'DEV2022K'),
 (357,2,66,172,'.','.',300.00000,600.00000,'SATM61'),
 (358,2,34,217,'..','.',400.00000,600.00000,'990'),
 (359,2,68,175,'.','.',500.00000,800.00000,'ATV20LCD'),
 (360,2,69,178,'VARIOS','VARIOD',100.00000,200.00000,'0'),
 (361,2,64,217,'.','.',200.00000,600.00000,'TXT PRO'),
 (362,2,4,217,'VARIAS','VARIAS',500.00000,1500.00000,'VARIOS'),
 (363,2,20,177,'.','.',2800.00000,3300.00000,'51\"'),
 (364,2,27,182,'.','.',500.00000,850.00000,'VARIOS'),
 (365,2,70,175,'.','.',1900.00000,2300.00000,'32\"'),
 (366,2,58,169,'.','.',400.00000,500.00000,'16MP'),
 (367,2,71,171,'VARIAS','VARIAS',150.00000,200.00000,'EMERSON'),
 (368,2,58,169,'.','.',150.00000,250.00000,'12 MP'),
 (369,2,5,175,'.','.',800.00000,1150.00000,'20\"'),
 (370,2,4,175,'.','.',2250.00000,2900.00000,'37\" VARIAS'),
 (371,2,20,175,'.','.',2700.00000,3400.00000,'40\"'),
 (372,2,69,175,'.','.',1650.00000,2300.00000,'32\"'),
 (373,2,4,217,'.','.',200.00000,500.00000,'LG-E400F'),
 (374,2,72,219,'.','.',2000.00000,3000.00000,'MAC'),
 (375,2,73,175,'.','.',2000.00000,2300.00000,'32\"'),
 (376,2,66,172,'.','.',1000.00000,1500.00000,'SA-AKX92'),
 (377,2,74,255,'VARIAS','VARIAS',700.00000,1000.00000,'MID 7033'),
 (378,2,75,219,'.','.',20000.00000,4500.00000,'MAC'),
 (379,2,75,219,'.','.',2000.00000,4500.00000,'MAC'),
 (380,2,31,250,'.','.',200.00000,400.00000,'MICROOCOMPONENTE'),
 (381,2,76,217,'.','.',300.00000,600.00000,'N295'),
 (382,9,52,256,'VARIOS','VARIOS',200.00000,300.00000,'VARIOS'),
 (383,9,77,229,'OTROS','OTROS',300.00000,600.00000,'OTROS'),
 (384,9,77,243,'OTROS','OTROS',400.00000,600.00000,'OTROS'),
 (385,2,52,228,'PULUIDORA','PULIDORA',100.00000,300.00000,'VARIOS'),
 (386,2,15,249,'VARIOS','VARIOS',1000.00000,3000.00000,'VARIOS'),
 (387,2,42,218,'BLU RAY','BLU RAY',1000.00000,1300.00000,'PHILIPS'),
 (388,2,4,175,'VARIAS','VARIAS',700.00000,1100.00000,'VARIAS'),
 (389,9,14,228,'VARIAS','VARIAS',300.00000,500.00000,'VARIAS'),
 (390,9,78,228,'VARIAS','VARIAS',300.00000,500.00000,'VARIOS'),
 (391,2,42,176,'VARIAS','VARIAS',300.00000,500.00000,'VARIAS'),
 (392,2,79,237,'OTROS','OTROS',1500.00000,2500.00000,'OTROS'),
 (393,2,58,169,'VARIAS','14 MEGAPIXELES',250.00000,350.00000,'14 MPX'),
 (394,2,8,257,'VARIAS','VARIAS',600.00000,800.00000,'BOSE'),
 (395,2,20,252,'VARIAS','VARIAS',2800.00000,3600.00000,'40\"'),
 (396,2,80,175,'OTROS','OTROS',1900.00000,2500.00000,'OTROS'),
 (397,9,52,243,'VARIAS','VARIAS',100.00000,300.00000,'TALADRO'),
 (398,2,30,219,'VARIAS','VARIAS',1000.00000,3000.00000,'ASUS'),
 (399,2,20,219,'VARIAS','VARIAS',1500.00000,3000.00000,'VARIAS'),
 (400,9,52,227,'VARIOS','VARIOS',200.00000,300.00000,'CORTADORA'),
 (401,9,8,243,'VARIAS','VARIAS',150.00000,250.00000,'CRAFTSMAN'),
 (402,2,8,249,'VARIAS','VARIAS',800.00000,1000.00000,'ALPINE'),
 (403,7,7,236,'VARIAS','VARIAS',1000.00000,1400.00000,'REDLINE'),
 (404,9,81,227,'VARIAS','VARIAS',100.00000,300.00000,'VARIAS'),
 (405,2,28,180,'VARIAS','VARIAS',900.00000,1400.00000,'PSP VITA'),
 (406,2,5,180,'VARIAS','VARIAS',1300.00000,1800.00000,'PSP VITA'),
 (407,2,10,253,'VARIOS','VARIOS',1500.00000,2200.00000,'64BG'),
 (408,2,82,169,'VARIAS','VARIAS',800.00000,1500.00000,'CANON'),
 (409,2,5,180,'VARIAS','VARIAS',850.00000,1250.00000,'PS3 1ERA GEN'),
 (410,2,29,217,'VARIOS','VARIOS',150.00000,450.00000,'VARIOS'),
 (411,9,78,243,'VARIAS','VARIAS',100.00000,200.00000,'DEWALT'),
 (412,2,7,175,'VARIAS','VARIAS',800.00000,1050.00000,'SANYO'),
 (413,2,19,217,'VARIAS','VARIAS',400.00000,800.00000,'LUMIA'),
 (414,2,5,182,'VARIAS','VARIAS',500.00000,800.00000,'VARIAS'),
 (415,2,20,252,'VARIAS','VARIAS',850.00000,1400.00000,'19\"'),
 (416,2,22,217,'VARIAS','VARIAS',400.00000,800.00000,'S100'),
 (417,2,5,252,'VARIAS','VARIAS',1800.00000,2500.00000,'32\"'),
 (418,2,51,219,'VARIAS','VARIAS',900.00000,2000.00000,'VARIAS'),
 (419,2,74,250,'VARIAS','VARIAS',100.00000,500.00000,'VARIAS'),
 (420,8,83,236,'VARIOS','VARIOS',100.00000,1000.00000,'GT'),
 (421,2,4,171,'VARIAS','VARIAS',100.00000,500.00000,'VARIAS'),
 (422,2,20,176,'VARIAS','VARIAS',100.00000,1000.00000,'21\"'),
 (423,9,84,232,'VARIAS','VARIAS',100.00000,200.00000,'MAKITA'),
 (424,2,31,175,'VARIAS','VARIAS',2500.00000,5000.00000,'32\"'),
 (425,2,20,175,'VARIAS','VARIAS',1100.00000,1500.00000,'LN22B650T6D'),
 (426,2,7,252,'VARIAS','19\"',400.00000,700.00000,'19X1'),
 (427,2,66,175,'VARIAS','PANASONIC 32\"',1000.00000,1800.00000,'TC-L32B6X'),
 (428,7,8,233,'VARIAS','VARIAS',1100.00000,2000.00000,'VARIAS'),
 (429,2,31,217,'1','1',400.00000,600.00000,'L7'),
 (430,8,85,236,'1','1',500.00000,1500.00000,'VARIAS'),
 (431,2,24,258,'1','1',500.00000,900.00000,'KAV60'),
 (432,2,24,258,'1','1',500.00000,700.00000,'D257-1850'),
 (433,8,86,236,'CROMADA','VARIAS',100.00000,300.00000,'FD18341996'),
 (434,2,28,180,'1','1',900.00000,1200.00000,'CEH-2101A'),
 (435,2,4,249,'1','1',500.00000,700.00000,'LM-U15560A'),
 (436,2,31,217,'VARIOS','VARIOS',500.00000,1000.00000,'3DP920'),
 (437,2,4,259,'1','1',500.00000,600.00000,'CM4230'),
 (438,2,29,217,'VARIOS','VARIOS',300.00000,700.00000,'XT615'),
 (439,2,20,217,'1','1',200.00000,400.00000,'GT-S5360L'),
 (440,2,54,252,'.','.',1500.00000,2000.00000,'LC-39LE44OU'),
 (441,2,79,219,'VARIOS','VARIOS',1000.00000,2000.00000,'2000'),
 (442,2,11,219,'VARIOS','VARIOS',1500.00000,2200.00000,'PRESARIO CQ42'),
 (443,2,4,175,'1','1',1500.00000,1900.00000,'32LD450-UA'),
 (444,2,19,217,'1','1',300.00000,500.00000,'610'),
 (445,2,60,169,'1','1',1500.00000,2500.00000,'D5000'),
 (446,2,27,219,'1','1',2000.00000,2500.00000,'PAVILION G4'),
 (447,2,4,217,'VARIOS','VARIOS',200.00000,600.00000,'LG L5X'),
 (448,2,20,217,'.','.',400.00000,500.00000,'SAMSUNG'),
 (449,2,28,180,'.','.',900.00000,1100.00000,'PSP'),
 (450,2,10,253,'1','1',2000.00000,2500.00000,'A1395'),
 (451,2,34,217,'1','1',500.00000,700.00000,'OT 8000'),
 (452,7,87,260,'1','1',200.00000,400.00000,'440-20'),
 (453,2,20,250,'1','1',400.00000,500.00000,'MX-E630'),
 (454,2,88,261,'1','1',500.00000,700.00000,'W1018'),
 (455,2,31,217,'VARIOS','VARIOS',100.00000,200.00000,'L3'),
 (456,2,6,180,'1','1',900.00000,1100.00000,'1439'),
 (457,2,27,219,'VARIOS','VARIOS',1500.00000,2000.00000,'PAVILION DV'),
 (458,2,28,217,'VARIOS','VARIAS',200.00000,500.00000,'E C1504'),
 (459,2,89,252,'VARIOS','VARIO',500.00000,800.00000,'LED-19X1'),
 (460,2,19,217,'1','1',400.00000,600.00000,'LUMIA 520'),
 (461,2,28,217,'1','1',200.00000,400.00000,'C1504'),
 (462,2,19,217,'.','.',700.00000,880.00000,'C3'),
 (463,2,31,175,'1','1',1000.00000,1500.00000,'26LH20'),
 (464,2,20,252,'VARIOS','VARIOS',850.00000,1200.00000,'T19C350'),
 (465,2,28,180,'VARIOS','SUPER SLIM',1000.00000,1800.00000,'CECH-4001B'),
 (466,2,20,252,'1','1',500.00000,1000.00000,'C350ND'),
 (467,2,34,217,'VARIOS','VARIOS',200.00000,400.00000,'ONE TOUCH POP'),
 (468,2,4,217,'1','1',400.00000,600.00000,'OPTIMUS BLAK'),
 (469,2,20,178,'VARIOS','VARIOS',150.00000,250.00000,'E360'),
 (470,2,20,175,'VARIOS','VARIOS',800.00000,1300.00000,'T24B350ND'),
 (471,2,4,217,'VARIOS','VARIOS',300.00000,600.00000,'L5'),
 (472,2,5,217,'VARIOS','VARIOS',500.00000,1000.00000,'XPERIA S'),
 (473,2,90,217,'VARIOS','VARIOS',200.00000,500.00000,'M4TEL'),
 (474,2,34,217,'VARIOS','VARIOS',200.00000,600.00000,'4010'),
 (475,2,91,255,'VARIOS','VARIOS',200.00000,600.00000,'VARIOS'),
 (476,2,28,180,'VARIOS','PS3 SLIM',700.00000,1000.00000,'CECH-3011A'),
 (477,2,5,217,'VARIOS','ST25A',200.00000,500.00000,'XPERIA U'),
 (478,9,81,230,'VARIOS','VARIOS',150.00000,300.00000,'BLACK&DECKER'),
 (479,2,24,219,'VARIOS','VARIOS',1000.00000,2500.00000,'ASPIRE5517'),
 (480,2,92,171,'VARIOS','.7 PIES',200.00000,280.00000,'JES70SE'),
 (481,2,31,178,'VARIOS','VARIOS',100.00000,200.00000,'DP522'),
 (482,2,29,217,'VARIOS','VARIOS',500.00000,1500.00000,'X'),
 (483,2,4,249,'VARIOS','VARIOS',700.00000,950.00000,'CM7520'),
 (484,2,79,219,'VARIOS','VARIOS',1100.00000,2000.00000,'1000'),
 (485,9,93,232,'VARIOS','VARIOS',100.00000,150.00000,'4225'),
 (486,2,69,175,'VARIOS','VARIOS',1000.00000,1500.00000,'MTV3012LCD'),
 (487,2,10,253,'VARIOS','VARIOS',2380.00000,3800.00000,'MINI'),
 (488,2,5,252,'VARIOS','VARIOS',3400.00000,4100.00000,'46EX650'),
 (489,2,66,175,'VARIOS','VARIOS',2800.00000,3400.00000,'P42X3X'),
 (490,2,34,217,'VARIOS','VARIOS',200.00000,500.00000,'5020A ONE TOUCH'),
 (491,2,94,262,'VARIAS','VARIAS',500.00000,1500.00000,'15KG'),
 (492,2,95,264,'VARIAS','VARIAS',500.00000,1500.00000,'POWER TOUCH SCREEN'),
 (493,2,10,265,'VARIOS','VARIOS',1000.00000,4000.00000,'A1278'),
 (494,2,20,258,'150GB DD','VARIOS',100.00000,600.00000,'VARIOS'),
 (495,2,34,217,'VARIOS','VARIOS',400.00000,1000.00000,'SCRIBE EASY'),
 (496,2,4,217,'VARIOS','VARIOS',500.00000,1500.00000,'3DP920'),
 (497,2,63,255,'VARIOS','VARIOS',1100.00000,1700.00000,'A1000'),
 (498,9,78,232,'VARIOS','VARIOS',200.00000,250.00000,'VARIOS'),
 (499,9,84,228,'VARIOS','VARIOS',400.00000,500.00000,'VARIOS'),
 (500,9,96,266,'VARIOS','VARIOS',100.00000,1000.00000,'CP9175'),
 (501,2,20,217,'VARIOS','VARIOS',500.00000,2500.00000,'GALAXY NOTE II'),
 (502,2,97,255,'VARIOS','VARIOS',100.00000,500.00000,'7009ME'),
 (503,2,66,169,'VARIOS','VARIOS',500.00000,1500.00000,'DMC-LZ20'),
 (504,8,98,233,'VARIOS','VARIOS',100.00000,500.00000,'VARIOS'),
 (505,2,34,217,'VARIOS','VARIOS',100.00000,1000.00000,'5035A'),
 (506,2,24,255,'VARIOS','VARIOS',500.00000,1000.00000,'B1-710'),
 (507,2,24,258,'VARIOS','VARIOS',500.00000,1500.00000,'ASPIRE'),
 (508,2,24,237,'CARIOS','VARIOS',800.00000,3500.00000,'DA220HQL'),
 (509,2,4,178,'VARIOS','VARIOS',150.00000,280.00000,'DP132'),
 (510,2,20,175,'VARIOS','VARIOS',500.00000,1200.00000,'VARIOS'),
 (511,9,99,230,'VARIOS','VARIOS',200.00000,300.00000,'ROTO-1/2A3'),
 (512,2,100,255,'VARIOS','VARIOS',100.00000,500.00000,'A-GO-GO'),
 (513,2,101,269,NULL,NULL,0.00000,0.00000,NULL);
/*!40000 ALTER TABLE `prendaselec` ENABLE KEYS */;


--
-- Definition of table `rematediario`
--

DROP TABLE IF EXISTS `rematediario`;
CREATE TABLE `rematediario` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Fecha` datetime DEFAULT NULL,
  `Status` int(1) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `rematediario`
--

/*!40000 ALTER TABLE `rematediario` DISABLE KEYS */;
/*!40000 ALTER TABLE `rematediario` ENABLE KEYS */;


--
-- Definition of table `saldos`
--

DROP TABLE IF EXISTS `saldos`;
CREATE TABLE `saldos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` date DEFAULT NULL,
  `Saldo` decimal(19,4) DEFAULT '0.0000',
  `PC` varchar(25) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `saldos`
--

/*!40000 ALTER TABLE `saldos` DISABLE KEYS */;
/*!40000 ALTER TABLE `saldos` ENABLE KEYS */;


--
-- Definition of table `salidainventario`
--

DROP TABLE IF EXISTS `salidainventario`;
CREATE TABLE `salidainventario` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` datetime DEFAULT NULL,
  `Pagado` int(11) DEFAULT '0',
  `Folio` int(10) DEFAULT '0',
  `TipoSalida` int(11) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `salidainventario`
--

/*!40000 ALTER TABLE `salidainventario` DISABLE KEYS */;
/*!40000 ALTER TABLE `salidainventario` ENABLE KEYS */;


--
-- Definition of table `sucursales`
--

DROP TABLE IF EXISTS `sucursales`;
CREATE TABLE `sucursales` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Clave` int(10) DEFAULT NULL,
  `NombreSucursal` varchar(100) DEFAULT NULL,
  `RazonSocial` varchar(100) DEFAULT NULL,
  `NombreComercial` varchar(100) DEFAULT NULL,
  `RFC` varchar(30) DEFAULT NULL,
  `Direccion` varchar(100) DEFAULT NULL,
  `Ciudad` varchar(60) DEFAULT NULL,
  `Estado` varchar(50) DEFAULT NULL,
  `Telefono` varchar(25) DEFAULT NULL,
  `Cp` int(10) DEFAULT '0',
  `Email` varchar(80) DEFAULT '',
  `DomicilioAclaraciones` varchar(200) DEFAULT NULL,
  `TelefonoAclaraciones` varchar(50) DEFAULT NULL,
  `CorreoAclaraciones` varchar(50) DEFAULT NULL,
  `ContratoRegistrado` varchar(50) DEFAULT NULL,
  `CodProfeco` varchar(50) DEFAULT NULL,
  `FechaContratoRegistrado` date DEFAULT NULL,
  `Activa` int(1) DEFAULT '0',
  `Cuenta` varchar(10) DEFAULT NULL,
  `Ip` varchar(15) DEFAULT NULL,
  `HorarioSucursal` varchar(250) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `Clave` (`Clave`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `sucursales`
--

/*!40000 ALTER TABLE `sucursales` DISABLE KEYS */;
INSERT INTO `sucursales` (`ID`,`Clave`,`NombreSucursal`,`RazonSocial`,`NombreComercial`,`RFC`,`Direccion`,`Ciudad`,`Estado`,`Telefono`,`Cp`,`Email`,`DomicilioAclaraciones`,`TelefonoAclaraciones`,`CorreoAclaraciones`,`ContratoRegistrado`,`CodProfeco`,`FechaContratoRegistrado`,`Activa`,`Cuenta`,`Ip`,`HorarioSucursal`) VALUES 
 (1,101,NULL,'LUIS FERNANDO SANTOYO RODRIGUEZ','SUC. PLATEROS','SARL670403CK8','AV. PLATEROS 239-A, ZONA CENTRO','FRESNILLO','ZACATECAS','.',99000,'','AV. DE LA CONVENCION 602-B, COL. LAS VIÑAS, AGUASCALIENTES, AGS. C.P. 20160','449-972-0489','OPERACIONES@BILLETIMAX.COM','31332-2012','7/002608-2012','2012-04-05',1,'630100','5.71.248.101',NULL);
/*!40000 ALTER TABLE `sucursales` ENABLE KEYS */;


--
-- Definition of table `tablacurp`
--

DROP TABLE IF EXISTS `tablacurp`;
CREATE TABLE `tablacurp` (
  `Indice` int(5) DEFAULT NULL,
  `Valor` varchar(2) DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tablacurp`
--

/*!40000 ALTER TABLE `tablacurp` DISABLE KEYS */;
INSERT INTO `tablacurp` (`Indice`,`Valor`) VALUES 
 (0,'0'),
 (1,'1'),
 (2,'2'),
 (3,'3'),
 (4,'4'),
 (5,'5'),
 (6,'6'),
 (7,'7'),
 (8,'8'),
 (9,'9'),
 (10,'A'),
 (11,'B'),
 (12,'C'),
 (13,'D'),
 (14,'E'),
 (15,'F'),
 (16,'G'),
 (17,'H'),
 (18,'I'),
 (19,'J'),
 (20,'K'),
 (21,'L'),
 (22,'M'),
 (23,'N'),
 (24,'Ñ'),
 (25,'O'),
 (26,'P'),
 (27,'Q'),
 (28,'R'),
 (29,'S'),
 (30,'T'),
 (31,'U'),
 (32,'V'),
 (33,'W'),
 (34,'X'),
 (35,'Y'),
 (36,'Z');
/*!40000 ALTER TABLE `tablacurp` ENABLE KEYS */;


--
-- Definition of table `tarjetaspuntos`
--

DROP TABLE IF EXISTS `tarjetaspuntos`;
CREATE TABLE `tarjetaspuntos` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `TipoTarjeta` varchar(60) NOT NULL,
  `pEmpeno` double(15,5) NOT NULL DEFAULT '0.00000',
  `pEmpenoAutos` double(15,5) NOT NULL DEFAULT '0.00000',
  `pRefrendo` double(15,5) NOT NULL DEFAULT '0.00000',
  `pRefrendoExt` double(15,5) NOT NULL DEFAULT '0.00000',
  `pDesempeno` double(15,5) NOT NULL DEFAULT '0.00000',
  `pVentas` double(15,5) NOT NULL DEFAULT '0.00000',
  `pApartados` double(15,5) NOT NULL DEFAULT '0.00000',
  `pAbonos` double(15,5) NOT NULL DEFAULT '0.00000',
  `FechaCreacion` datetime NOT NULL,
  `Activa` int(11) NOT NULL DEFAULT '1',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tarjetaspuntos`
--

/*!40000 ALTER TABLE `tarjetaspuntos` DISABLE KEYS */;
INSERT INTO `tarjetaspuntos` (`ID`,`TipoTarjeta`,`pEmpeno`,`pEmpenoAutos`,`pRefrendo`,`pRefrendoExt`,`pDesempeno`,`pVentas`,`pApartados`,`pAbonos`,`FechaCreacion`,`Activa`) VALUES 
 (1,'CLIENTE FRECUENTE',10.00000,10.00000,10.00000,0.00000,10.00000,10.00000,10.00000,10.00000,'2014-03-14 09:50:54',1);
/*!40000 ALTER TABLE `tarjetaspuntos` ENABLE KEYS */;


--
-- Definition of table `tipo`
--

DROP TABLE IF EXISTS `tipo`;
CREATE TABLE `tipo` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Descripcion` varchar(50) DEFAULT NULL,
  `Kilataje` int(1) NOT NULL DEFAULT '0',
  `Peso` int(1) NOT NULL DEFAULT '0',
  `Ordenamiento` int(1) DEFAULT '0',
  `IdTipoGarantia` int(10) DEFAULT '0',
  `IdTipoBienes` int(10) DEFAULT '0',
  `IdTipoUnidad` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=11 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tipo`
--

/*!40000 ALTER TABLE `tipo` DISABLE KEYS */;
INSERT INTO `tipo` (`ID`,`Descripcion`,`Kilataje`,`Peso`,`Ordenamiento`,`IdTipoGarantia`,`IdTipoBienes`,`IdTipoUnidad`) VALUES 
 (1,'ORO',1,1,1,7,1,2),
 (2,'ELECTRONICOS',0,0,0,13,0,1),
 (3,'RELOJES',0,0,3,8,12,1),
 (7,'OTROS',0,0,4,13,0,1),
 (6,'DOCUMENTOS',0,0,5,13,0,1),
 (8,'BICICLETA',0,0,0,13,0,1),
 (9,'HERRAMIENTA',0,0,0,13,0,1),
 (10,'PLATA',1,1,2,7,1,2);
/*!40000 ALTER TABLE `tipo` ENABLE KEYS */;


--
-- Definition of table `tipointeres`
--

DROP TABLE IF EXISTS `tipointeres`;
CREATE TABLE `tipointeres` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Descripcion` varchar(50) DEFAULT NULL,
  `Serie` int(5) DEFAULT '0',
  `Ordenamiento` int(2) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `Id` (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=10 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tipointeres`
--

/*!40000 ALTER TABLE `tipointeres` DISABLE KEYS */;
INSERT INTO `tipointeres` (`ID`,`Descripcion`,`Serie`,`Ordenamiento`) VALUES 
 (1,'TRADICIONAL',1,1),
 (3,'TRADICIONAL',2,1),
 (4,'FIJA',1,1),
 (5,'FIJA',2,1),
 (6,'TRAD COMPRA',1,1),
 (7,'TRAD COMPRA',2,1),
 (8,'COMPLETO',1,1),
 (9,'COMPLETO',2,1);
/*!40000 ALTER TABLE `tipointeres` ENABLE KEYS */;


--
-- Definition of table `tipoperiodo`
--

DROP TABLE IF EXISTS `tipoperiodo`;
CREATE TABLE `tipoperiodo` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Descripcion` varchar(50) DEFAULT NULL,
  `Periodo` int(5) DEFAULT '0',
  `PrestamoAvaluo` double(15,5) DEFAULT '0.00000',
  `Ordenamiento` int(2) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=5 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tipoperiodo`
--

/*!40000 ALTER TABLE `tipoperiodo` DISABLE KEYS */;
INSERT INTO `tipoperiodo` (`ID`,`Descripcion`,`Periodo`,`PrestamoAvaluo`,`Ordenamiento`) VALUES 
 (1,'MENSUAL',30,100.00000,1),
 (2,'QUINCENAL',15,100.00000,2),
 (3,'SEMANAL',7,100.00000,3),
 (4,'DIARIA',1,100.00000,4);
/*!40000 ALTER TABLE `tipoperiodo` ENABLE KEYS */;


--
-- Definition of table `tipoprenda`
--

DROP TABLE IF EXISTS `tipoprenda`;
CREATE TABLE `tipoprenda` (
  `ID` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `Descripcion` varchar(255) DEFAULT NULL,
  `IDTipo` int(11) DEFAULT '0',
  `Minimo` double(15,5) DEFAULT '0.00000',
  `Maximo` double(15,5) DEFAULT '0.00000',
  `Funciones` varchar(255) DEFAULT NULL,
  `Caracteristicas` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=271 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tipoprenda`
--

/*!40000 ALTER TABLE `tipoprenda` DISABLE KEYS */;
INSERT INTO `tipoprenda` (`ID`,`Descripcion`,`IDTipo`,`Minimo`,`Maximo`,`Funciones`,`Caracteristicas`) VALUES 
 (10,'ANILLO CON CIRCONIAS',1,0.00000,0.00000,NULL,NULL),
 (11,'ANILLO SOLITARIO',1,0.00000,0.00000,NULL,NULL),
 (12,'ARRACADAS',1,0.00000,0.00000,NULL,NULL),
 (13,'PAR ARETES',1,0.00000,0.00000,NULL,NULL),
 (14,'ARGOLLA',1,0.00000,0.00000,NULL,NULL),
 (15,'BRAZALETE',1,0.00000,0.00000,NULL,NULL),
 (16,'PAR BROQUELES',1,0.00000,0.00000,NULL,NULL),
 (17,'CADENA TEJIDO CARTIER',1,0.00000,0.00000,NULL,NULL),
 (18,'CADENA TEJIDO CHINO',1,0.00000,0.00000,NULL,NULL),
 (19,'CADENA TEJIDO ESLABO',1,0.00000,0.00000,NULL,NULL),
 (20,'CADENA TEJIDO ESPECI',1,0.00000,0.00000,NULL,NULL),
 (21,'DIJE',1,0.00000,0.00000,NULL,NULL),
 (22,'ESCLAVA TEJIDO CARTIER',1,0.00000,0.00000,NULL,NULL),
 (23,'ESCLAVA TEJIDO CHINO',1,0.00000,0.00000,NULL,NULL),
 (24,'ESCLAVA TEJIDO ESLABONES',1,0.00000,0.00000,NULL,NULL),
 (26,'GARGANTILLA HECHURA',1,0.00000,0.00000,NULL,NULL),
 (27,'MEDALLA SENCILLA',1,0.00000,0.00000,NULL,NULL),
 (28,'MEDALLA GRABADA',1,0.00000,0.00000,NULL,NULL),
 (35,'BROQUEL',1,0.00000,0.00000,NULL,NULL),
 (38,'PULSERA TEJIDO CARTIER',1,0.00000,0.00000,NULL,NULL),
 (39,'PULSERA TEJIDO ESLABONES',1,0.00000,0.00000,NULL,NULL),
 (40,'PULSERA TEJIDO ESPECI',1,0.00000,0.00000,NULL,NULL),
 (41,'PULSERA TEJIDO CHINO',1,0.00000,0.00000,NULL,NULL),
 (43,'CADENA TEJIDO GUCCI',1,0.00000,0.00000,NULL,NULL),
 (46,'ANILLO DE GRADUACION',1,0.00000,0.00000,NULL,NULL),
 (47,'SEMANARIO',1,0.00000,0.00000,NULL,NULL),
 (50,'CADENA TEJIDO TORSAL',1,0.00000,0.00000,NULL,NULL),
 (66,'CADENA PLANCHA',1,0.00000,0.00000,NULL,NULL),
 (81,'CADENA TROQUELADA',1,0.00000,0.00000,NULL,NULL),
 (82,'PULSERA TROQUELADA',1,0.00000,0.00000,NULL,NULL),
 (85,'ANILLO COCTEL',1,0.00000,0.00000,NULL,NULL),
 (87,'CADENA CUBANA',1,0.00000,0.00000,NULL,NULL),
 (88,'PULSERA CUBANA',1,0.00000,0.00000,NULL,NULL),
 (89,'ESCLAVA CUBANA',1,0.00000,0.00000,NULL,NULL),
 (90,'ANILLO QUINCE AÑOS',1,0.00000,0.00000,NULL,NULL),
 (91,'ANILLO INICIALES',1,0.00000,0.00000,NULL,NULL),
 (93,'CADENA ANCLA',1,0.00000,0.00000,NULL,NULL),
 (94,'PULSERA ANCLA',1,0.00000,0.00000,NULL,NULL),
 (95,'ESCLAVA ANCLA',1,0.00000,0.00000,NULL,NULL),
 (96,'ESCLAVA BARROCA',1,0.00000,0.00000,NULL,NULL),
 (97,'PULSERA BARROCA',1,0.00000,0.00000,NULL,NULL),
 (99,'ANILLO DAMA',1,0.00000,0.00000,NULL,NULL),
 (114,'ROSARIO',1,0.00000,0.00000,NULL,NULL),
 (173,'TV P/NORMAL',2,0.00000,0.00000,NULL,NULL),
 (229,'ROUTERS',9,0.00000,0.00000,NULL,NULL),
 (171,'MICROONDAS',2,0.00000,0.00000,NULL,NULL),
 (172,'MODULARES',2,0.00000,0.00000,NULL,NULL),
 (169,'CAMARA DIGITAL',2,0.00000,0.00000,NULL,NULL),
 (170,'GRABADORAS',2,0.00000,0.00000,NULL,NULL),
 (174,'TV PROYECTORES',2,0.00000,0.00000,NULL,NULL),
 (175,'TVS LCD',2,0.00000,0.00000,NULL,NULL),
 (176,'TVS PLANA',2,0.00000,0.00000,NULL,NULL),
 (177,'TVS PLASMA',2,0.00000,0.00000,NULL,NULL),
 (178,'DVD',2,0.00000,0.00000,NULL,NULL),
 (180,'VIDEOJUEGOS',2,0.00000,0.00000,NULL,NULL),
 (181,'VARIOS',1,0.00000,0.00000,NULL,NULL),
 (182,'VIDEOCAMARAS',2,0.00000,0.00000,NULL,NULL),
 (183,'ESCLAVA',1,0.00000,0.00000,NULL,NULL),
 (230,'ROTOMARTILLOS',9,0.00000,0.00000,NULL,NULL),
 (185,'CENTENARIO',1,0.00000,0.00000,NULL,NULL),
 (187,'RELOJ PARA DAMA',1,0.00000,0.00000,NULL,NULL),
 (189,'PULSO',1,0.00000,0.00000,NULL,NULL),
 (190,'CADENA',1,0.00000,0.00000,NULL,NULL),
 (191,'PULSERA',1,0.00000,0.00000,NULL,NULL),
 (192,'GARGANTILLA',1,0.00000,0.00000,NULL,NULL),
 (194,'RELOJ CABALLERO',1,0.00000,0.00000,NULL,NULL),
 (196,'ARETES',1,0.00000,0.00000,NULL,NULL),
 (198,'ANILLO CABALLERO',1,0.00000,0.00000,NULL,NULL),
 (201,'DOCUMENTOS',2,0.00000,0.00000,NULL,NULL),
 (202,'COMPUTADORAS DE ESCRITORIO',2,0.00000,0.00000,NULL,NULL),
 (203,'MAQUINARIA',6,0.00000,0.00000,NULL,NULL),
 (204,'PRENDEDOR',1,0.00000,0.00000,NULL,NULL),
 (206,'ANILLO',1,0.00000,0.00000,NULL,NULL),
 (228,'PULIDORAS',9,0.00000,0.00000,NULL,NULL),
 (227,'CORTADORAS',9,0.00000,0.00000,NULL,NULL),
 (210,'CADENA CON DIJE',1,0.00000,0.00000,NULL,NULL),
 (211,'GARGANTILLA CON DIJE',1,0.00000,0.00000,NULL,NULL),
 (225,'DVD PORTATIL',2,0.00000,0.00000,NULL,NULL),
 (226,'TALADROS',9,0.00000,0.00000,NULL,NULL),
 (216,'MP3',2,0.00000,0.00000,NULL,NULL),
 (217,'CELULARES',2,0.00000,0.00000,NULL,NULL),
 (218,'TEATRO EN CASA',2,0.00000,0.00000,NULL,NULL),
 (219,'LAPTOP',2,0.00000,0.00000,NULL,NULL),
 (220,'CRUZIFIJO',1,0.00000,0.00000,NULL,NULL),
 (221,'PEDACERA',1,0.00000,0.00000,NULL,NULL),
 (222,'APPLE',2,0.00000,0.00000,NULL,NULL),
 (223,'AUTOS',8,0.00000,0.00000,NULL,NULL),
 (224,'MOTOS',8,0.00000,0.00000,NULL,NULL),
 (231,'TALADRO INALAMBRICO',9,0.00000,0.00000,NULL,NULL),
 (232,'CALADORAS',9,0.00000,0.00000,NULL,NULL),
 (233,'BICICLETAS',8,0.00000,0.00000,NULL,NULL),
 (234,'PALM',2,0.00000,0.00000,NULL,NULL),
 (235,'INSTRUMENTOS MUSICALES',7,0.00000,0.00000,NULL,NULL),
 (236,'BICILETA',7,0.00000,0.00000,NULL,NULL),
 (237,'COMPUTADORA DE ESCRITORIO',2,0.00000,0.00000,NULL,NULL),
 (238,'CAMARA FILMADORA',2,0.00000,0.00000,NULL,NULL),
 (239,'FISTON',1,0.00000,0.00000,NULL,NULL),
 (240,'BULOVA',3,0.00000,0.00000,NULL,NULL),
 (241,'AUTOCLEP',9,0.00000,0.00000,NULL,NULL),
 (242,'IPOD TOUCH',2,0.00000,0.00000,NULL,NULL),
 (243,'TALADRO',9,0.00000,0.00000,NULL,NULL),
 (244,'IPOD',2,0.00000,0.00000,NULL,NULL),
 (245,'COMPRESOR',9,0.00000,0.00000,NULL,NULL),
 (246,'CIERRAS',9,0.00000,0.00000,NULL,NULL),
 (247,'MICROMETRO',9,0.00000,0.00000,NULL,NULL),
 (248,'ROLEX',3,0.00000,0.00000,NULL,NULL),
 (249,'STEREO',2,0.00000,0.00000,NULL,NULL),
 (250,'MINICOMPONENTE',2,0.00000,0.00000,NULL,NULL),
 (251,'BICICLETA',2,0.00000,0.00000,NULL,NULL),
 (252,'LED',2,0.00000,0.00000,NULL,NULL),
 (253,'IPAD',2,0.00000,0.00000,NULL,NULL),
 (254,'HD',2,0.00000,0.00000,NULL,NULL),
 (255,'TABLET',2,0.00000,0.00000,NULL,NULL),
 (256,'ROUTERS',9,0.00000,0.00000,NULL,NULL),
 (257,'BOCINA BLUETOOTH',2,0.00000,0.00000,NULL,NULL),
 (258,'MINI LAPTOP',2,0.00000,0.00000,NULL,NULL),
 (259,'MINICOMPONENTE',2,0.00000,0.00000,NULL,NULL),
 (260,'CHOCOMILERA',7,0.00000,0.00000,NULL,NULL),
 (261,'REFRIGERADORES',2,0.00000,0.00000,NULL,NULL),
 (262,'LAVADORA',2,0.00000,0.00000,NULL,NULL),
 (263,'CENTENARIO',1,0.00000,0.00000,NULL,NULL),
 (264,'AUTOSTEREO',2,0.00000,0.00000,NULL,NULL),
 (265,'MAC',2,0.00000,0.00000,NULL,NULL),
 (266,'AUTOSCANNER',9,0.00000,0.00000,NULL,NULL),
 (267,'GENERAL',1,0.00000,0.00000,NULL,NULL),
 (268,'GENERAL',0,0.00000,0.00000,NULL,NULL),
 (269,'GENERAL',2,0.00000,0.00000,NULL,NULL),
 (270,'GENERAL',10,0.00000,0.00000,NULL,NULL);
/*!40000 ALTER TABLE `tipoprenda` ENABLE KEYS */;


--
-- Definition of table `traspasos`
--

DROP TABLE IF EXISTS `traspasos`;
CREATE TABLE `traspasos` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` date DEFAULT NULL,
  `Folio` int(10) DEFAULT '0',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  `SucursalDestino` int(11) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `traspasos`
--

/*!40000 ALTER TABLE `traspasos` DISABLE KEYS */;
/*!40000 ALTER TABLE `traspasos` ENABLE KEYS */;


--
-- Definition of table `usuarios`
--

DROP TABLE IF EXISTS `usuarios`;
CREATE TABLE `usuarios` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Nombre` varchar(120) DEFAULT NULL,
  `Usuario` varchar(30) DEFAULT NULL,
  `Contraseña` varchar(30) DEFAULT NULL,
  `empeño` int(2) DEFAULT '0',
  `empeñoautos` int(2) DEFAULT '0',
  `desempeños` int(2) DEFAULT '0',
  `refrendos` int(2) DEFAULT '0',
  `ubicacion` int(2) DEFAULT '0',
  `ventas` int(2) DEFAULT '0',
  `busqueda` int(2) DEFAULT '0',
  `conceptos` int(2) DEFAULT '0',
  `cortecaja` int(2) DEFAULT '0',
  `balance` int(2) DEFAULT '0',
  `repfinanciero` int(2) DEFAULT '0',
  `cierresucursal` int(2) DEFAULT '0',
  `grupos` int(2) DEFAULT '0',
  `dotacion` int(2) DEFAULT '0',
  `devolucion` int(2) DEFAULT '0',
  `inventariofisico` int(2) DEFAULT '0',
  `existencias` int(2) DEFAULT '0',
  `etiquetas` int(2) DEFAULT '0',
  `exporinformacion` int(2) DEFAULT '0',
  `repcontable` int(2) DEFAULT '0',
  `repauditoria` int(2) DEFAULT '0',
  `repauxiliar` int(2) DEFAULT '0',
  `repventas` int(2) DEFAULT '0',
  `repinventarios` int(2) DEFAULT '0',
  `repvencidos` int(2) DEFAULT '0',
  `rephistorico` int(2) DEFAULT '0',
  `repempeños` int(2) DEFAULT '0',
  `repasistencia` int(2) DEFAULT '0',
  `movimientocaja` int(2) DEFAULT '0',
  `movimientobanco` int(2) DEFAULT '0',
  `transferencias` int(2) DEFAULT '0',
  `remates` int(2) DEFAULT '0',
  `gastos` int(2) DEFAULT '0',
  `parametros` int(2) DEFAULT '0',
  `capboletas` int(2) DEFAULT '0',
  `usuarios` int(2) DEFAULT '0',
  `cancelbol` int(2) DEFAULT '0',
  `repgastos` int(2) DEFAULT '0',
  `catdivisas` int(2) DEFAULT '0',
  `cotizacion` int(2) DEFAULT '0',
  `comvendiv` int(2) DEFAULT '0',
  `repdivisas` int(2) DEFAULT '0',
  `movidiv` int(2) DEFAULT '0',
  `facturacion` int(2) DEFAULT '0',
  `cotizarempeño` int(2) DEFAULT '0',
  `reporteremates` int(2) DEFAULT '0',
  `abonar` int(2) DEFAULT '0',
  `precio` int(2) DEFAULT '0',
  `modificarcorte` int(2) DEFAULT '0',
  `hacercorte` int(2) DEFAULT '0',
  `interesrefrendo` int(2) DEFAULT '0',
  `interesdesempeño` int(2) DEFAULT '0',
  `IDUsuario` int(2) DEFAULT '0',
  `AnaliClientes` int(2) DEFAULT '0',
  `RegUbicacion` int(2) DEFAULT '0',
  `RepAlmoneda` int(2) DEFAULT '0',
  `RepCierres` int(2) DEFAULT '0',
  `RepIngresos` int(2) DEFAULT '0',
  `CancelVenta` int(2) DEFAULT '0',
  `CambioVenta` int(2) DEFAULT '0',
  `PagoDemasia` int(2) DEFAULT '0',
  `RepApartado` int(2) DEFAULT '0',
  `RepUtilidad` int(2) DEFAULT '0',
  `EntradaInven` int(2) DEFAULT '0',
  `SalidaInven` int(2) DEFAULT '0',
  `Deslotifica` int(2) DEFAULT '0',
  `TrasInven` int(2) DEFAULT '0',
  `ListaPrecio` int(2) DEFAULT '0',
  `RepCompras` int(2) DEFAULT '0',
  `RepTras` int(2) DEFAULT '0',
  `Kardex` int(2) DEFAULT '0',
  `RepAnti` int(2) DEFAULT '0',
  `RepEnve` int(2) DEFAULT '0',
  `RepEnveP` int(2) DEFAULT '0',
  `RepSalida` int(2) DEFAULT '0',
  `RepAutorizaciones` int(2) DEFAULT '0',
  `RepCierreSucursal` int(2) DEFAULT '0',
  `RepPrendasSimi` int(2) DEFAULT '0',
  `RepAleatoria` int(2) DEFAULT '0',
  `RepPrendasAudi` int(2) DEFAULT '0',
  `Traspasos` int(2) DEFAULT '0',
  `Sucursales` int(2) DEFAULT '0',
  `CatTipos` int(2) DEFAULT '0',
  `CatFamilias` int(2) DEFAULT '0',
  `CatSubFamilias` int(2) DEFAULT '0',
  `CatMedios` int(2) DEFAULT '0',
  `CatCuentasGas` int(2) DEFAULT '0',
  `CancelarGas` int(2) DEFAULT '0',
  `CargosAbonos` int(2) DEFAULT '0',
  `CatClientes` int(2) DEFAULT '0',
  `MoviBoveda` int(2) DEFAULT '0',
  `MostrarApartados` int(2) DEFAULT '0',
  `ApartadosVencidos` int(2) DEFAULT '0',
  `EntradasInventario` int(2) DEFAULT '0',
  `SalidasInventario` int(2) DEFAULT '0',
  `PrecioVitrina` int(2) DEFAULT '0',
  `TipoPrenda` int(2) DEFAULT '0',
  `PreciosKilataje` int(2) DEFAULT '0',
  `TarjetaBeneficio` int(2) DEFAULT '0',
  `DescuentoVentas` int(2) DEFAULT '0',
  `RecalculoPrecios` int(2) DEFAULT '0',
  `PrestamoBoleta1` int(1) DEFAULT '0',
  `Estatus` int(2) DEFAULT '1',
  `PagosFijos` int(2) DEFAULT '0',
  `CambioPlan` int(2) DEFAULT '0',
  `CierreDivisas` int(2) DEFAULT '0',
  `RepCartera` int(2) DEFAULT '0',
  `VenCliente` int(2) DEFAULT '0',
  `EtiInven` int(2) DEFAULT '0',
  `RepDota` int(2) DEFAULT '0',
  `RepDesempenos` int(2) DEFAULT '0',
  `RepRefrendos` int(2) DEFAULT '0',
  `RepHorarios` int(2) DEFAULT '0',
  `RepPartidaBoveda` int(2) DEFAULT '0',
  `RepAseguradora` int(2) DEFAULT '0',
  `RepCancelaciones` int(2) DEFAULT '0',
  `RepEmpeProm` int(2) DEFAULT '0',
  `RepDesemProm` int(2) DEFAULT '0',
  `RepRefProm` int(2) DEFAULT '0',
  `ConTipoTasa` int(2) DEFAULT '0',
  `ConVencidos` int(2) DEFAULT '0',
  `ConStatus` int(2) DEFAULT '0',
  `PrestamoMes` int(2) DEFAULT '0',
  `Medios` int(2) DEFAULT '0',
  `ConfiguraTasas` int(2) DEFAULT '0',
  `ConfiguraDiam` int(2) DEFAULT '0',
  `Catalogos` int(2) DEFAULT '0',
  `MensajeContratos` int(2) DEFAULT '0',
  `ConexionSuc` int(2) DEFAULT '0',
  `GeneraAutoriza` int(1) DEFAULT '0',
  `CatElec` int(1) DEFAULT '0',
  `RefrendarVencidos` int(1) DEFAULT '0',
  `CancelaCierre` int(1) DEFAULT '0',
  `mld_parametros` int(2) DEFAULT '0',
  `mld_movatipicos` int(2) DEFAULT '0',
  `mld_expclientes` int(2) DEFAULT '0',
  `mld_reppormenorizado` int(2) DEFAULT '0',
  `RepIdentClientes` int(2) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM AUTO_INCREMENT=14 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `usuarios`
--

/*!40000 ALTER TABLE `usuarios` DISABLE KEYS */;
INSERT INTO `usuarios` (`ID`,`Nombre`,`Usuario`,`Contraseña`,`empeño`,`empeñoautos`,`desempeños`,`refrendos`,`ubicacion`,`ventas`,`busqueda`,`conceptos`,`cortecaja`,`balance`,`repfinanciero`,`cierresucursal`,`grupos`,`dotacion`,`devolucion`,`inventariofisico`,`existencias`,`etiquetas`,`exporinformacion`,`repcontable`,`repauditoria`,`repauxiliar`,`repventas`,`repinventarios`,`repvencidos`,`rephistorico`,`repempeños`,`repasistencia`,`movimientocaja`,`movimientobanco`,`transferencias`,`remates`,`gastos`,`parametros`,`capboletas`,`usuarios`,`cancelbol`,`repgastos`,`catdivisas`,`cotizacion`,`comvendiv`,`repdivisas`,`movidiv`,`facturacion`,`cotizarempeño`,`reporteremates`,`abonar`,`precio`,`modificarcorte`,`hacercorte`,`interesrefrendo`,`interesdesempeño`,`IDUsuario`,`AnaliClientes`,`RegUbicacion`,`RepAlmoneda`,`RepCierres`,`RepIngresos`,`CancelVenta`,`CambioVenta`,`PagoDemasia`,`RepApartado`,`RepUtilidad`,`EntradaInven`,`SalidaInven`,`Deslotifica`,`TrasInven`,`ListaPrecio`,`RepCompras`,`RepTras`,`Kardex`,`RepAnti`,`RepEnve`,`RepEnveP`,`RepSalida`,`RepAutorizaciones`,`RepCierreSucursal`,`RepPrendasSimi`,`RepAleatoria`,`RepPrendasAudi`,`Traspasos`,`Sucursales`,`CatTipos`,`CatFamilias`,`CatSubFamilias`,`CatMedios`,`CatCuentasGas`,`CancelarGas`,`CargosAbonos`,`CatClientes`,`MoviBoveda`,`MostrarApartados`,`ApartadosVencidos`,`EntradasInventario`,`SalidasInventario`,`PrecioVitrina`,`TipoPrenda`,`PreciosKilataje`,`TarjetaBeneficio`,`DescuentoVentas`,`RecalculoPrecios`,`PrestamoBoleta1`,`Estatus`,`PagosFijos`,`CambioPlan`,`CierreDivisas`,`RepCartera`,`VenCliente`,`EtiInven`,`RepDota`,`RepDesempenos`,`RepRefrendos`,`RepHorarios`,`RepPartidaBoveda`,`RepAseguradora`,`RepCancelaciones`,`RepEmpeProm`,`RepDesemProm`,`RepRefProm`,`ConTipoTasa`,`ConVencidos`,`ConStatus`,`PrestamoMes`,`Medios`,`ConfiguraTasas`,`ConfiguraDiam`,`Catalogos`,`MensajeContratos`,`ConexionSuc`,`GeneraAutoriza`,`CatElec`,`RefrendarVencidos`,`CancelaCierre`,`mld_parametros`,`mld_movatipicos`,`mld_expclientes`,`mld_reppormenorizado`,`RepIdentClientes`) VALUES 
 (1,'GERENTE SUCURSAL','gerente','gerente',1,1,1,1,0,1,1,1,1,1,1,1,1,1,0,1,1,1,0,1,1,1,1,1,1,1,1,0,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,0,1,1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,1,1,0,0,1,1,1,1,1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,0,0,0,0,0,0,0),
 (8,'SOPORTE SIGMA','soporte','sigmadesarrollo',1,1,1,1,0,1,1,1,1,1,1,1,1,1,0,1,1,1,0,1,1,1,1,1,1,1,1,0,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,0,1,1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,1,1,0,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,1,0);
/*!40000 ALTER TABLE `usuarios` ENABLE KEYS */;


--
-- Definition of table `vendedores`
--

DROP TABLE IF EXISTS `vendedores`;
CREATE TABLE `vendedores` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Nombre` varchar(50) DEFAULT NULL,
  `Apellidos` varchar(100) DEFAULT NULL,
  `Iniciales` varchar(10) DEFAULT NULL,
  `Meta` double(15,5) DEFAULT '0.00000',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `vendedores`
--

/*!40000 ALTER TABLE `vendedores` DISABLE KEYS */;
/*!40000 ALTER TABLE `vendedores` ENABLE KEYS */;


--
-- Definition of table `ventas`
--

DROP TABLE IF EXISTS `ventas`;
CREATE TABLE `ventas` (
  `ID` int(10) NOT NULL AUTO_INCREMENT,
  `Fecha` datetime DEFAULT NULL,
  `Folio` int(10) DEFAULT '0',
  `IDCliente` int(10) DEFAULT '0',
  `IVA` double(15,5) DEFAULT '0.00000',
  `Vencimiento` date DEFAULT NULL,
  `Cancelado` tinyint(1) NOT NULL DEFAULT '0',
  `Apartado` tinyint(1) NOT NULL DEFAULT '0',
  `Pagado` tinyint(1) NOT NULL DEFAULT '0',
  `Total` double(15,5) DEFAULT '0.00000',
  `PC` varchar(25) DEFAULT '',
  `Descuento` double(15,5) DEFAULT '0.00000',
  `IDUsuario` int(10) DEFAULT '0',
  `IDSucursal` int(10) DEFAULT '0',
  `OrigenCancelacion` int(2) DEFAULT '0',
  `FechaMovimiento` datetime DEFAULT NULL,
  `IDUsuarioDesc` int(10) DEFAULT '0',
  `ImporteDevolucion` double(15,5) DEFAULT '0.00000',
  `TipoVenta` int(11) DEFAULT '0',
  `IDVendedor` int(10) DEFAULT '0',
  `Efectivo` double(15,5) DEFAULT '0.00000',
  `DescuentoEfectivo` double(15,5) DEFAULT '0.00000',
  `DescuentoXPuntos` double(15,5) DEFAULT '0.00000',
  `SaldoPuntosAnterior` double(15,5) DEFAULT '0.00000',
  `PuntosUsados` double(15,5) DEFAULT '0.00000',
  `PuntosAcumulados` double(15,5) DEFAULT '0.00000',
  `SaldoPuntosActual` double(15,5) DEFAULT '0.00000',
  `IDTarjeta` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDCliente` (`IDCliente`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `ventas`
--

/*!40000 ALTER TABLE `ventas` DISABLE KEYS */;
/*!40000 ALTER TABLE `ventas` ENABLE KEYS */;

--
-- Create schema basereportes
--

CREATE DATABASE IF NOT EXISTS basereportes;
USE basereportes;

--
-- Definition of table `abonosapartados`
--

DROP TABLE IF EXISTS `abonosapartados`;
CREATE TABLE `abonosapartados` (
  `ID` int(11) DEFAULT NULL,
  `IDVenta` int(11) DEFAULT NULL,
  `Cliente` varchar(80) COLLATE latin1_general_ci DEFAULT NULL,
  `TotalAbonado` double DEFAULT NULL
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Dumping data for table `abonosapartados`
--

/*!40000 ALTER TABLE `abonosapartados` DISABLE KEYS */;
/*!40000 ALTER TABLE `abonosapartados` ENABLE KEYS */;


--
-- Definition of table `articulos`
--

DROP TABLE IF EXISTS `articulos`;
CREATE TABLE `articulos` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `IDEmpeno` int(11) DEFAULT '0',
  `Articulo` varchar(400) DEFAULT NULL,
  `Color` varchar(50) DEFAULT NULL,
  `Peso` double(15,3) DEFAULT '0.000',
  `Claridad` varchar(50) DEFAULT NULL,
  `Kilates` int(11) DEFAULT '0',
  `Avaluo` double(15,3) DEFAULT '0.000',
  `Prestamo` double(15,3) DEFAULT '0.000',
  `Tipo` int(11) DEFAULT '0',
  `Cantidad` int(11) DEFAULT '0',
  `Modelo` varchar(50) DEFAULT NULL,
  `Destino` varchar(45) DEFAULT NULL,
  `PesoPiedra` double(15,5) DEFAULT '0.00000',
  `CantidadPiedras` double(15,5) DEFAULT '0.00000',
  `PrestamoDiamante` double(15,5) DEFAULT '0.00000',
  `Puntos` double(15,5) DEFAULT '0.00000',
  `Observaciones` varchar(250) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `articulos`
--

/*!40000 ALTER TABLE `articulos` DISABLE KEYS */;
/*!40000 ALTER TABLE `articulos` ENABLE KEYS */;


--
-- Definition of table `articulosalmoneda`
--

DROP TABLE IF EXISTS `articulosalmoneda`;
CREATE TABLE `articulosalmoneda` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `IDEmpeno` int(11) DEFAULT '0',
  `Codigo` varchar(50) DEFAULT NULL,
  `Articulo` varchar(255) DEFAULT NULL,
  `Peso` double(15,3) DEFAULT '0.000',
  `Kilates` int(11) DEFAULT '0',
  `Avaluo` double(15,5) DEFAULT '0.00000',
  `Prestamo` double(15,5) DEFAULT '0.00000',
  `Tipo` int(11) DEFAULT '0',
  `Cantidad` int(11) DEFAULT '0',
  `Modelo` varchar(50) DEFAULT NULL,
  `Estado` varchar(50) DEFAULT NULL,
  `Serie` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `articulosalmoneda`
--

/*!40000 ALTER TABLE `articulosalmoneda` DISABLE KEYS */;
/*!40000 ALTER TABLE `articulosalmoneda` ENABLE KEYS */;


--
-- Definition of table `compraventa`
--

DROP TABLE IF EXISTS `compraventa`;
CREATE TABLE `compraventa` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `iddivisa` int(11) DEFAULT '0',
  `cotizacion` double(15,5) DEFAULT '0.00000',
  `compra` double(15,5) DEFAULT '0.00000',
  `venta` double(15,5) DEFAULT '0.00000',
  `interno` int(2) DEFAULT '0',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `compraventa`
--

/*!40000 ALTER TABLE `compraventa` DISABLE KEYS */;
/*!40000 ALTER TABLE `compraventa` ENABLE KEYS */;


--
-- Definition of table `cortecajaventanilla`
--

DROP TABLE IF EXISTS `cortecajaventanilla`;
CREATE TABLE `cortecajaventanilla` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Sucursal` varchar(80) DEFAULT NULL,
  `Cajero` varchar(60) DEFAULT NULL,
  `Saldo` double(15,5) DEFAULT '0.00000',
  `Debe` double(15,5) DEFAULT '0.00000',
  `Haber` double(15,5) DEFAULT '0.00000',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cortecajaventanilla`
--

/*!40000 ALTER TABLE `cortecajaventanilla` DISABLE KEYS */;
/*!40000 ALTER TABLE `cortecajaventanilla` ENABLE KEYS */;


--
-- Definition of table `cortecuentas`
--

DROP TABLE IF EXISTS `cortecuentas`;
CREATE TABLE `cortecuentas` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Cuenta` varchar(30) DEFAULT NULL,
  `Descripcion` varchar(60) DEFAULT NULL,
  `Fecha` date DEFAULT NULL,
  `Concepto` varchar(200) DEFAULT NULL,
  `Folio` int(15) unsigned DEFAULT '0',
  `Movimientos` int(11) DEFAULT '0',
  `Cargo` double(15,5) DEFAULT '0.00000',
  `Abono` double(15,5) DEFAULT '0.00000',
  `PC` varchar(50) DEFAULT NULL,
  `Saldo` double(15,5) DEFAULT '0.00000',
  `Serie` int(1) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=9 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cortecuentas`
--

/*!40000 ALTER TABLE `cortecuentas` DISABLE KEYS */;
INSERT INTO `cortecuentas` (`ID`,`Cuenta`,`Descripcion`,`Fecha`,`Concepto`,`Folio`,`Movimientos`,`Cargo`,`Abono`,`PC`,`Saldo`,`Serie`) VALUES 
 (5,'201700','EMPEÑO','2014-07-14','SALDO INICIAL',0,0,0.00000,0.00000,'INTERCAMBIOSPC',0.00000,0),
 (6,'201700','EMPEÑO','2014-07-14','Importacion Oro',1,0,314086.00000,0.00000,'INTERCAMBIOSPC',314086.00000,0),
 (7,'201700','EMPEÑO','2014-07-14','Importacion Electronicos',1,0,203177.00000,0.00000,'INTERCAMBIOSPC',517263.00000,0),
 (8,'201700','EMPEÑO','2014-07-14','Importacion Plata',1,0,9300.00000,0.00000,'INTERCAMBIOSPC',526563.00000,0);
/*!40000 ALTER TABLE `cortecuentas` ENABLE KEYS */;


--
-- Definition of table `cortedivisas`
--

DROP TABLE IF EXISTS `cortedivisas`;
CREATE TABLE `cortedivisas` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `IDDivisa` int(11) DEFAULT '0',
  `Dotacion` int(11) DEFAULT '0',
  `Retiro` int(11) DEFAULT '0',
  `Compras` int(11) DEFAULT '0',
  `Ventas` int(11) DEFAULT '0',
  `PC` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cortedivisas`
--

/*!40000 ALTER TABLE `cortedivisas` DISABLE KEYS */;
/*!40000 ALTER TABLE `cortedivisas` ENABLE KEYS */;


--
-- Definition of table `diario`
--

DROP TABLE IF EXISTS `diario`;
CREATE TABLE `diario` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Sucursal` varchar(60) DEFAULT NULL,
  `Cajero` varchar(60) DEFAULT NULL,
  `Cuenta` varchar(10) DEFAULT NULL,
  `Leyenda` varchar(60) DEFAULT NULL,
  `Importe1` double(15,5) DEFAULT '0.00000',
  `Folio1` int(15) DEFAULT '0',
  `Importe2` double(15,5) DEFAULT '0.00000',
  `Folio2` int(15) DEFAULT '0',
  `Importe3` double(15,5) DEFAULT '0.00000',
  `Folio3` int(15) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `diario`
--

/*!40000 ALTER TABLE `diario` DISABLE KEYS */;
/*!40000 ALTER TABLE `diario` ENABLE KEYS */;


--
-- Definition of table `estadocuentapuntos`
--

DROP TABLE IF EXISTS `estadocuentapuntos`;
CREATE TABLE `estadocuentapuntos` (
  `ID` int(11) unsigned NOT NULL AUTO_INCREMENT,
  `Fecha` date NOT NULL,
  `Folio` int(11) DEFAULT '0',
  `Movimiento` varchar(45) DEFAULT NULL,
  `Cargo` double(15,5) DEFAULT '0.00000',
  `Abono` double(15,5) DEFAULT '0.00000',
  `Saldo` double(15,5) DEFAULT '0.00000',
  `IDCliente` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `estadocuentapuntos`
--

/*!40000 ALTER TABLE `estadocuentapuntos` DISABLE KEYS */;
/*!40000 ALTER TABLE `estadocuentapuntos` ENABLE KEYS */;


--
-- Definition of table `existencia_divisas`
--

DROP TABLE IF EXISTS `existencia_divisas`;
CREATE TABLE `existencia_divisas` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Divisa` int(10) unsigned DEFAULT '0',
  `Compra` double(15,5) DEFAULT '0.00000',
  `Venta` double(15,5) DEFAULT '0.00000',
  `Entrada` int(10) DEFAULT '0',
  `Salida` int(10) DEFAULT '0',
  `EntradaInicial` int(10) DEFAULT '0',
  `SalidaInicial` int(10) DEFAULT '0',
  `TipoCambio` double(15,5) DEFAULT '0.00000',
  PRIMARY KEY (`ID`) USING BTREE
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `existencia_divisas`
--

/*!40000 ALTER TABLE `existencia_divisas` DISABLE KEYS */;
/*!40000 ALTER TABLE `existencia_divisas` ENABLE KEYS */;


--
-- Definition of table `horarios`
--

DROP TABLE IF EXISTS `horarios`;
CREATE TABLE `horarios` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Clave` int(11) DEFAULT '0',
  `Hora` varchar(50) COLLATE latin1_general_ci DEFAULT NULL,
  `Empeos` int(10) DEFAULT '0',
  `Refrendos` int(10) DEFAULT '0',
  `Desempeos` int(10) DEFAULT '0',
  `Reempeos` int(10) DEFAULT '0',
  `Ventas` int(10) DEFAULT '0',
  `Apartados` int(10) DEFAULT '0',
  `ComDiv` int(10) DEFAULT '0',
  `VenDiv` int(10) DEFAULT '0',
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Dumping data for table `horarios`
--

/*!40000 ALTER TABLE `horarios` DISABLE KEYS */;
/*!40000 ALTER TABLE `horarios` ENABLE KEYS */;


--
-- Definition of table `nota`
--

DROP TABLE IF EXISTS `nota`;
CREATE TABLE `nota` (
  `id` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `cliente` varchar(100) DEFAULT NULL,
  `descripcion` varchar(50) DEFAULT NULL,
  `importe` double(15,5) DEFAULT '0.00000',
  PRIMARY KEY (`id`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `nota`
--

/*!40000 ALTER TABLE `nota` DISABLE KEYS */;
/*!40000 ALTER TABLE `nota` ENABLE KEYS */;


--
-- Definition of table `opcionpagos`
--

DROP TABLE IF EXISTS `opcionpagos`;
CREATE TABLE `opcionpagos` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Vencimiento` date DEFAULT NULL,
  `Almacenaje` double(15,5) DEFAULT '0.00000',
  `Seguro` double(15,5) DEFAULT '0.00000',
  `Interes` double(15,5) DEFAULT '0.00000',
  `IDEmpeno` int(11) DEFAULT '0',
  `Prestamo` double(15,5) DEFAULT '0.00000',
  `TipoInteres` varchar(45) DEFAULT NULL,
  `ImporteIva` double(15,5) DEFAULT '0.00000',
  `FechaIni` date DEFAULT NULL,
  `Refrendo` double(15,5) DEFAULT '0.00000',
  `PC` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1 ROW_FORMAT=DYNAMIC;

--
-- Dumping data for table `opcionpagos`
--

/*!40000 ALTER TABLE `opcionpagos` DISABLE KEYS */;
/*!40000 ALTER TABLE `opcionpagos` ENABLE KEYS */;


--
-- Definition of table `repcomisiones`
--

DROP TABLE IF EXISTS `repcomisiones`;
CREATE TABLE `repcomisiones` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Abonos` double(15,5) DEFAULT '0.00000',
  `IDVenta` int(10) DEFAULT '0',
  `Folio` int(10) DEFAULT '0',
  `Fecha` datetime NOT NULL,
  `IVA` double(15,5) DEFAULT '0.00000',
  `Descuento` double(15,5) DEFAULT '0.00000',
  `Vencimiento` date DEFAULT NULL,
  `Total` double(15,5) DEFAULT '0.00000',
  `Pagado` tinyint(1) DEFAULT '0',
  `Cliente` varchar(171) DEFAULT NULL,
  `Vendedor` varchar(151) DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `repcomisiones`
--

/*!40000 ALTER TABLE `repcomisiones` DISABLE KEYS */;
/*!40000 ALTER TABLE `repcomisiones` ENABLE KEYS */;


--
-- Definition of table `repingresos`
--

DROP TABLE IF EXISTS `repingresos`;
CREATE TABLE `repingresos` (
  `Fecha` date DEFAULT NULL,
  `Intereses` double(15,5) DEFAULT '0.00000',
  `Ventas` double(15,5) DEFAULT '0.00000',
  `Apartados` double(15,5) DEFAULT '0.00000',
  `Iva` double(15,5) DEFAULT '0.00000',
  `OtrosIngre` double(15,5) DEFAULT '0.00000'
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `repingresos`
--

/*!40000 ALTER TABLE `repingresos` DISABLE KEYS */;
/*!40000 ALTER TABLE `repingresos` ENABLE KEYS */;


--
-- Definition of table `reportes`
--

DROP TABLE IF EXISTS `reportes`;
CREATE TABLE `reportes` (
  `ID` int(11) DEFAULT NULL,
  `Dia` date DEFAULT NULL,
  `Contratos` int(11) DEFAULT '0',
  `TipoPrenda` int(10) DEFAULT '0',
  `Prendas` int(11) DEFAULT '0',
  `Peso` double(15,5) DEFAULT '0.00000',
  `Avaluo` double(15,5) DEFAULT '0.00000',
  `Prestamo` double(15,5) DEFAULT '0.00000',
  `Intereses` double(15,5) DEFAULT '0.00000',
  `Abono` double(15,5) DEFAULT '0.00000',
  `IvaIntereses` double(15,5) DEFAULT '0.00000'
) ENGINE=MyISAM DEFAULT CHARSET=latin1 COLLATE=latin1_general_ci;

--
-- Dumping data for table `reportes`
--

/*!40000 ALTER TABLE `reportes` DISABLE KEYS */;
/*!40000 ALTER TABLE `reportes` ENABLE KEYS */;


--
-- Definition of table `repvencidos`
--

DROP TABLE IF EXISTS `repvencidos`;
CREATE TABLE `repvencidos` (
  `IDEmpeno` int(11) NOT NULL DEFAULT '0',
  `NumContrato` int(10) DEFAULT '0',
  `Fecha` datetime NOT NULL,
  `Vencimiento` date NOT NULL,
  `Cliente` varchar(90) DEFAULT NULL,
  `Avaluo` double(15,5) DEFAULT '0.00000',
  `Prestamo` double(15,5) DEFAULT '0.00000',
  `Serie` int(10) DEFAULT '0',
  `TipoInteres` varchar(20) DEFAULT NULL,
  `TipoTasa` varchar(20) DEFAULT NULL,
  `FechaMovimiento` date DEFAULT NULL,
  `Tel` varchar(80) DEFAULT NULL,
  `Celular` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`IDEmpeno`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `repvencidos`
--

/*!40000 ALTER TABLE `repvencidos` DISABLE KEYS */;
/*!40000 ALTER TABLE `repvencidos` ENABLE KEYS */;


--
-- Definition of table `simulador`
--

DROP TABLE IF EXISTS `simulador`;
CREATE TABLE `simulador` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Pago` int(11) DEFAULT '0',
  `Intereses` double(15,5) DEFAULT '0.00000',
  `Amortizacion` double(15,5) DEFAULT '0.00000',
  `Saldo` double(15,5) DEFAULT '0.00000',
  `Vencimiento` date DEFAULT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `simulador`
--

/*!40000 ALTER TABLE `simulador` DISABLE KEYS */;
/*!40000 ALTER TABLE `simulador` ENABLE KEYS */;

--
-- Create schema basedatos
--

CREATE DATABASE IF NOT EXISTS basedatos;
USE basedatos;

--
-- Definition of procedure `spRepComisiones`
--

DROP PROCEDURE IF EXISTS `spRepComisiones`;

DELIMITER $$

/*!50003 SET @TEMP_SQL_MODE=@@SQL_MODE, SQL_MODE='STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ $$
CREATE DEFINER=`billetimax`@`%` PROCEDURE `spRepComisiones`(IN FechaIni DATE, IN FechaFin DATE)
BEGIN SELECT SUM(abonos.Importe) AS Abonos,ventas.ID AS IDVenta,ventas.Folio,abonos.Fecha AS FechaAbono, ventas.IVA,ventas.Descuento,ventas.Vencimiento,ventas.Total,ventas.Pagado,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente, CONCAT(vendedores.Nombre,' ',vendedores.Apellidos) AS Vendedor FROM ventas left join abonos ON ventas.ID=abonos.IDVenta INNER JOIN clientes ON ventas.IDCliente=clientes.ID LEFT JOIN vendedores ON ventas.IDVendedor=vendedores.ID WHERE DATE_FORMAT(abonos.Fecha,'%Y%-%m%-%d') BETWEEN FechaIni AND FechaFin AND abonos.Cancelado=0 AND ventas.Cancelado=0 AND ventas.Apartado=1 GROUP BY ventas.ID; End $$
/*!50003 SET SESSION SQL_MODE=@TEMP_SQL_MODE */  $$

DELIMITER ;

--
-- Definition of procedure `spRepVencidos`
--

DROP PROCEDURE IF EXISTS `spRepVencidos`;

DELIMITER $$

/*!50003 SET @TEMP_SQL_MODE=@@SQL_MODE, SQL_MODE='STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ $$
CREATE DEFINER=`billetimax`@`%` PROCEDURE `spRepVencidos`(IN FechaIni DATE, IN FechaFin DATE, IN DiasEnajena INTEGER, IN TipoContrato INTEGER, IN TipoPrenda INTEGER)
BEGIN SELECT DISTINCT e.ID,e.NumContrato,e.Fecha,e.Vencimiento,CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,e.Avaluo,e.Serie,e.Prestamo,e.TipoInteres,e.TipoTasa FROM empeno e INNER JOIN clientes c ON e.IDCliente=c.ID LEFT JOIN detallesempeno de ON e.ID=de.IDEmpeno WHERE if(TipoContrato=1,(e.Serie=1 OR e.Serie=2 OR e.Serie=3),if(TipoContrato=3,e.Serie=2,de.Tipo=TipoPrenda)) AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0,DiasEnajena) DAY),'%Y%/%m%/%d') BETWEEN FechaIni AND FechaFin AND e.Cancelado=0 AND e.Destino=0 AND e.Pagado=0 ORDER BY NumContrato; END $$
/*!50003 SET SESSION SQL_MODE=@TEMP_SQL_MODE */  $$

DELIMITER ;

--
-- Definition of view `vwapartadosrematados`
--

DROP TABLE IF EXISTS `vwapartadosrematados`;
DROP VIEW IF EXISTS `vwapartadosrematados`;
CREATE ALGORITHM=UNDEFINED DEFINER=`mrayudon`@`%` SQL SECURITY DEFINER VIEW `vwapartadosrematados` AS select sum(`abonos`.`Importe`) AS `Abonos`,`ventas`.`ID` AS `ID`,`ventas`.`Fecha` AS `Fecha`,`ventas`.`FechaMovimiento` AS `FechaMovimiento`,`ventas`.`Folio` AS `Folio`,`ventas`.`IVA` AS `IVA`,`ventas`.`Vencimiento` AS `Vencimiento`,`ventas`.`Total` AS `Total`,`ventas`.`Descuento` AS `Descuento`,`ventas`.`Pagado` AS `Pagado`,`ventas`.`Cancelado` AS `Cancelado`,`ventas`.`OrigenCancelacion` AS `OrigenCancelacion`,concat(`clientes`.`Nombre`,' ',`clientes`.`Apellido`) AS `Cliente` from ((`ventas` left join `abonos` on((`ventas`.`ID` = `abonos`.`IDVenta`))) left join `clientes` on((`ventas`.`IDCliente` = `clientes`.`ID`))) where ((`ventas`.`OrigenCancelacion` = 2) and (`ventas`.`Apartado` = 1) and (`ventas`.`Cancelado` = 1) and (`abonos`.`Cancelado` = 0)) group by `ventas`.`ID`;

--
-- Definition of view `vwdetallesempeno`
--

DROP TABLE IF EXISTS `vwdetallesempeno`;
DROP VIEW IF EXISTS `vwdetallesempeno`;
CREATE ALGORITHM=UNDEFINED DEFINER=`mrayudon`@`%` SQL SECURITY DEFINER VIEW `vwdetallesempeno` AS select `detallesempeno`.`ID` AS `ID`,`detallesempeno`.`IDEmpeno` AS `IDEmpeno`,`detallesempeno`.`Cantidad` AS `Cantidad`,`detallesempeno`.`Tipo` AS `Tipo`,`detallesempeno`.`Articulo` AS `Articulo`,`detallesempeno`.`Peso` AS `PesoTotal`,`detallesempeno`.`PesoPiedras` AS `PesoPiedras`,(`detallesempeno`.`Peso` - `detallesempeno`.`PesoPiedras`) AS `PesoReal`,`detallesempeno`.`Prestamo` AS `Prestamo`,`detallesempeno`.`Avaluo` AS `Avaluo`,`detallesempeno`.`Observaciones` AS `Observaciones`,`detallesempeno`.`Estado` AS `Estado`,`detallesempeno`.`Marca` AS `Marca`,`detallesempeno`.`Modelo` AS `Modelo`,`detallesempeno`.`Serie` AS `Serie`,`detallesempeno`.`Tamano` AS `Tamano`,`detallesempeno`.`Color` AS `Color`,`tipo`.`Descripcion` AS `Tipo_DESC`,`kilatajes`.`Descripcion` AS `Kil_DESC` from ((`detallesempeno` left join `tipo` on((`detallesempeno`.`Tipo` = `tipo`.`ID`))) left join `kilatajes` on((`detallesempeno`.`Kilates` = `kilatajes`.`Clave`)));

--
-- Definition of view `vwfacturadiaria`
--

DROP TABLE IF EXISTS `vwfacturadiaria`;
DROP VIEW IF EXISTS `vwfacturadiaria`;
CREATE ALGORITHM=UNDEFINED DEFINER=`mrayudon`@`%` SQL SECURITY DEFINER VIEW `vwfacturadiaria` AS select count(`a`.`ID`) AS `NumRegistros`,`a`.`Fecha` AS `Fecha`,sum(`a`.`Importe`) AS `ImporteTotal` from `auxiliar` `a` where ((`a`.`Importe` > 0) and ((`a`.`Cuenta` = '520450') or (`a`.`Cuenta` = '670350') or (`a`.`Cuenta` = '680350') or (`a`.`Cuenta` = '690350'))) group by `a`.`Fecha` order by `a`.`Fecha`;

--
-- Definition of view `vwfacturaventas`
--

DROP TABLE IF EXISTS `vwfacturaventas`;
DROP VIEW IF EXISTS `vwfacturaventas`;
CREATE ALGORITHM=UNDEFINED DEFINER=`mrayudon`@`%` SQL SECURITY DEFINER VIEW `vwfacturaventas` AS select count(`a`.`ID`) AS `NumRegistros`,`a`.`Fecha` AS `Fecha`,sum(`a`.`Importe`) AS `ImporteTotal` from `auxiliar` `a` where ((`a`.`Importe` > 0) and (`a`.`Cuenta` = '620450') and (`a`.`Concepto` = 'Ventas')) group by `a`.`Fecha` order by `a`.`Fecha`;

--
-- Definition of view `vwpagosfijos`
--

DROP TABLE IF EXISTS `vwpagosfijos`;
DROP VIEW IF EXISTS `vwpagosfijos`;
CREATE ALGORITHM=UNDEFINED DEFINER=`mrayudon`@`%` SQL SECURITY DEFINER VIEW `vwpagosfijos` AS select `pagosfijos`.`IDEmpeno` AS `IDEmpeno`,max(`pagosfijos`.`FechaMovimiento`) AS `FechaMovimiento` from `pagosfijos` where ((`pagosfijos`.`Pagado` = 1) and (`pagosfijos`.`Cancelado` = 0)) group by `pagosfijos`.`IDEmpeno` order by `pagosfijos`.`IDEmpeno`,`pagosfijos`.`ID`;

--
-- Definition of view `vwrepapartados`
--

DROP TABLE IF EXISTS `vwrepapartados`;
DROP VIEW IF EXISTS `vwrepapartados`;
CREATE ALGORITHM=UNDEFINED DEFINER=`mrayudon`@`%` SQL SECURITY DEFINER VIEW `vwrepapartados` AS select sum(`abonos`.`Importe`) AS `Abonos`,`ventas`.`ID` AS `ID`,`ventas`.`Fecha` AS `Fecha`,`ventas`.`Folio` AS `Folio`,`ventas`.`IVA` AS `IVA`,`ventas`.`Vencimiento` AS `Vencimiento`,`ventas`.`Total` AS `Total`,`ventas`.`Descuento` AS `Descuento`,`ventas`.`Pagado` AS `Pagado`,`ventas`.`Cancelado` AS `Cancelado`,`ventas`.`OrigenCancelacion` AS `OrigenCancelacion`,concat(`clientes`.`Nombre`,' ',`clientes`.`Apellido`) AS `Cliente` from ((`ventas` left join `abonos` on((`ventas`.`ID` = `abonos`.`IDVenta`))) left join `clientes` on((`ventas`.`IDCliente` = `clientes`.`ID`))) where ((`ventas`.`Apartado` = 1) and (`abonos`.`Cancelado` = 0)) group by `ventas`.`ID`;



/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
