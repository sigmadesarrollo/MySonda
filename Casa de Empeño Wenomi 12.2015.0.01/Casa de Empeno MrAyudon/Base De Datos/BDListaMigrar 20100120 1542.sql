-- MySQL Administrator dump 1.4
--
-- ------------------------------------------------------
-- Server version	5.1.42-community


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
  `Peso` double(15,3),
  `Prestamo` double(15,5),
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
  `Direccion` varchar(70) DEFAULT NULL,
  `Colonia` varchar(120) DEFAULT NULL,
  `Municipio` varchar(120) DEFAULT NULL,
  `Estado` varchar(60) DEFAULT NULL,
  `Tel` varchar(50) DEFAULT NULL,
  `Identificacion` varchar(60) DEFAULT NULL,
  `NumeroIdentificacion` varchar(30) DEFAULT NULL,
  `IDMedio` int(10) DEFAULT '0',
  `Boletas` int(10) DEFAULT '0',
  `Notas` varchar(150) DEFAULT NULL,
  `CP` varchar(10) DEFAULT NULL,
  `Rfc` varchar(35) DEFAULT NULL,
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
) ENGINE=MyISAM AUTO_INCREMENT=6 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `configuraciontasas`
--

/*!40000 ALTER TABLE `configuraciontasas` DISABLE KEYS */;
INSERT INTO `configuraciontasas` (`ID`,`IDTipoInteres`,`IDTipoPeriodo`,`IDPlazo`,`TasaTipica`,`TasaPromocion`,`TasaPreferencial`,`PorPrestamo`,`Cat`,`Almacenaje`,`Seguro`) VALUES 
 (1,1,1,3,7.00000,7.00000,7.00000,83.00000,174.00000,7.50000,0.00000),
 (5,3,1,1,7.00000,7.00000,7.00000,83.00000,174.00000,7.50000,0.00000),
 (4,1,1,1,7.00000,7.00000,7.00000,83.00000,174.00000,7.50000,0.00000);
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
 (1,'110100','EFECTIVO','110101','EFECTIVO RECIBIDO EN CAJA'),
 (2,'110100','EFECTIVO','110150','EFECTIVO PAGADO DE CAJA'),
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
 (26,'199400','CAJA','199401','DOTACION A CAJERO'),
 (27,'199400','CAJA','199450','CIERRE DE CAJA'),
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
  `Año` int(10) DEFAULT '0',
  `Color` varchar(50) DEFAULT NULL,
  `Placas` varchar(50) DEFAULT NULL,
  `Factura` varchar(50) DEFAULT NULL,
  `Agencia` varchar(50) DEFAULT NULL,
  `NumTarjetacircu` varchar(50) DEFAULT NULL,
  `NumMotor` varchar(50) DEFAULT NULL,
  `SerieChasis` varchar(50) DEFAULT NULL,
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
 (22,22,'22K',1,5);
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
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `marcas`
--

/*!40000 ALTER TABLE `marcas` DISABLE KEYS */;
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
) ENGINE=MyISAM AUTO_INCREMENT=12 DEFAULT CHARSET=latin1;

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
 (10,'YA ES CLIENTE'),
 (11,'PASO DE LOCAL');
/*!40000 ALTER TABLE `medios` ENABLE KEYS */;


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
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `movimientos`
--

/*!40000 ALTER TABLE `movimientos` DISABLE KEYS */;
INSERT INTO `movimientos` (`ID`,`Movimiento`,`FolioBancos`,`FolioGastos`,`FolioVentas`,`FolioDepositos`,`FolioTransferencias`,`FolioCompras`,`FolioSalidaInventario`,`FolioAjustes`,`FolioBoveda`,`FolioDivisas`,`FolioNotas`,`Fecha`,`FolioAutorizacion`,`FolioInventario`,`FolioTraspasos`,`FolioBovedaDivisas`) VALUES 
 (1,1,1,1,1,1,1,1,1,1,1,1,1,'2010-01-20',1,1,1,1);
/*!40000 ALTER TABLE `movimientos` ENABLE KEYS */;


--
-- Definition of table `nacionalidad`
--

DROP TABLE IF EXISTS `nacionalidad`;
CREATE TABLE `nacionalidad` (
  `ID` int(10) unsigned NOT NULL AUTO_INCREMENT,
  `Nacionalidad` varchar(30) NOT NULL,
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `nacionalidad`
--

/*!40000 ALTER TABLE `nacionalidad` DISABLE KEYS */;
/*!40000 ALTER TABLE `nacionalidad` ENABLE KEYS */;


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
  PRIMARY KEY (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `parametros`
--

/*!40000 ALTER TABLE `parametros` DISABLE KEYS */;
INSERT INTO `parametros` (`ID`,`Datos`,`PrestamoAvaluo`,`PrestamoAvaluoAutos`,`PrestamoAvaluoElec`,`Almacenaje`,`Seguro`,`GtosVenta`,`Comision`,`IVA`,`Negociacion`,`Operacion`,`PagoMinimo`,`DiasEnajenacion`,`VenApartados`,`EngancheApartados`,`IvaVentas`,`DiasGracia`,`DiasGraciaAutos`,`PolizaSeguro`,`FechaExpedicion`,`Aseguradora`,`ImportePerdida`,`Notas`,`VenDemasia`,`ImporteAutorizacion`,`DiasGraciaApa`,`Gerente`,`CalidadEx`,`CalidadB`,`CalidadR`,`CalidadM`,`Centenario`,`DescuentoVentas`,`TipoCambioOnza`,`PrestamoAvaluoDiamante`,`8K`,`Venta8K`,`10K`,`Venta10K`,`14K`,`Venta14K`,`18K`,`Venta18K`,`22K`,`Venta22K`,`24K`,`Venta24K`,`LimiteInferior`,`LimiteSuperior`,`LimiteInferiorAutos`,`LimiteSuperiorAutos`,`DescuentoPagosFijos`,`Limite1`,`Limite2`,`VenAlmoneda`,`AbonoMinimo`,`ImpresoraDefault`,`DiasPenaliza`,`PrecioAutos`,`Cat`,`IntAnual`,`AlmAnual`) VALUES 
 (1,'2005-02-01',80.00000,60.00000,1.20000,2.50000,2.50000,25.00000,25.00000,0.00000,10.00000,6.00000,15.00000,15,2,25,0.00000,2,0,'0','2007-02-27','NO',20.00000,'EN REFRENDO Y DESEMPEÑOS EL INTERES SE COBRA POR DIA!,  GRACIAS POR SU PREFERENCIA!!',2,2500.00000,3,'TAPIA SANCHEZ JOSE MANUEL',90.00,90.00,90.00,90.00,1095.00000,5.00000,13.95000,1.20000,54.59000,176.00000,78.97000,220.00000,128.17000,310.00000,122.67000,396.00000,149.00000,475.00000,163.56000,528.00000,3000.00000,6000.00000,100000.00000,220000.00000,70.00000,10000.00000,15000.00000,4,25.00000,'',7,100.00000,0.00000,84.00000,90.00000);
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
 (13,1,14,1,147.20000,0),
 (14,1,14,2,147.20000,0),
 (15,1,14,3,147.20000,0),
 (16,1,14,4,147.20000,0),
 (17,1,1,1,184.34000,0),
 (18,1,1,2,184.34000,0),
 (19,1,1,3,184.34000,0),
 (20,1,1,4,184.34000,0),
 (26,1,2,1,257.72000,0),
 (27,1,2,2,257.72000,0),
 (29,1,2,3,257.72000,0),
 (30,1,2,4,257.72000,0),
 (31,1,3,1,309.44000,0),
 (32,1,3,2,309.44000,0),
 (33,1,3,3,309.44000,0),
 (34,1,3,4,309.44000,0),
 (76,1,21,4,442.05000,0),
 (75,1,21,3,442.05000,0),
 (74,1,21,2,442.05000,0),
 (73,1,21,1,442.05000,0),
 (168,1,22,1,397.84000,0),
 (169,1,22,2,397.84000,0),
 (170,1,22,3,397.84000,0),
 (171,1,22,4,397.84000,0),
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
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `prendaselec`
--

/*!40000 ALTER TABLE `prendaselec` DISABLE KEYS */;
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
  `RazonSocial` varchar(100) DEFAULT NULL,
  `NombreComercial` varchar(100) DEFAULT NULL,
  `RFC` varchar(30) DEFAULT NULL,
  `Direccion` varchar(100) DEFAULT NULL,
  `Ciudad` varchar(60) DEFAULT NULL,
  `Estado` varchar(50) DEFAULT NULL,
  `Telefono` varchar(25) DEFAULT NULL,
  `Cp` int(10) DEFAULT '0',
  `Activa` int(1) DEFAULT '0',
  `Cuenta` varchar(10) DEFAULT NULL,
  `Ip` varchar(15) DEFAULT NULL,
  PRIMARY KEY (`ID`),
  KEY `Clave` (`Clave`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `sucursales`
--

/*!40000 ALTER TABLE `sucursales` DISABLE KEYS */;
INSERT INTO `sucursales` (`ID`,`Clave`,`RazonSocial`,`NombreComercial`,`RFC`,`Direccion`,`Ciudad`,`Estado`,`Telefono`,`Cp`,`Activa`,`Cuenta`,`Ip`) VALUES 
 (1,104,'RUBALCAVA VALDIVIA LUIS HECTOR','FRANQUICIAS SUPERVARO','XXXXXXXXXXXX','AV 16 DE SEPTIEMBRE # 5-B, COL CENTRO','PABELLON DE ARTEAGA','AGUASCALIENTES','465-9582323',20660,1,'630100','5.71.248.101');
/*!40000 ALTER TABLE `sucursales` ENABLE KEYS */;


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
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`)
) ENGINE=MyISAM AUTO_INCREMENT=8 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tipo`
--

/*!40000 ALTER TABLE `tipo` DISABLE KEYS */;
INSERT INTO `tipo` (`ID`,`Descripcion`,`Kilataje`,`Peso`,`Ordenamiento`) VALUES 
 (1,'ORO',1,1,1),
 (2,'ELECTRONICOS',0,0,0),
 (3,'RELOJES',0,0,2),
 (7,'OTROS',0,0,4),
 (6,'DOCUMENTOS',0,0,3);
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
) ENGINE=MyISAM AUTO_INCREMENT=4 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `tipointeres`
--

/*!40000 ALTER TABLE `tipointeres` DISABLE KEYS */;
INSERT INTO `tipointeres` (`ID`,`Descripcion`,`Serie`,`Ordenamiento`) VALUES 
 (1,'TRADICIONAL',1,1),
 (3,'TRADICIONAL',2,1);
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
) ENGINE=MyISAM AUTO_INCREMENT=216 DEFAULT CHARSET=latin1;

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
 (161,'CON DIAMANTES',3,0.00000,0.00000,NULL,NULL),
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
 (184,'LAPTOP',2,0.00000,0.00000,NULL,NULL),
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
 (207,'RELOJ MIDO AC. INOX.',3,0.00000,0.00000,NULL,NULL),
 (208,'RELOJ ORO ROLEX',3,0.00000,0.00000,NULL,NULL),
 (209,'RELOJ OMEGA AC.INOX.',3,0.00000,0.00000,NULL,NULL),
 (210,'CADENA CON DIJE',1,0.00000,0.00000,NULL,NULL),
 (211,'GARGANTILLA CON DIJE',1,0.00000,0.00000,NULL,NULL),
 (212,'BOLSAS DE CELOFAN',7,0.00000,0.00000,NULL,NULL),
 (213,'CAJAS PARA JOYERIA',7,0.00000,0.00000,NULL,NULL),
 (214,'RELOJ CARTIER',3,0.00000,0.00000,NULL,NULL),
 (215,'RELOJES REPLICA',3,0.00000,0.00000,NULL,NULL);
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
  PRIMARY KEY (`ID`),
  KEY `ID` (`ID`),
  KEY `IDUsuario` (`IDUsuario`)
) ENGINE=MyISAM AUTO_INCREMENT=2 DEFAULT CHARSET=latin1;

--
-- Dumping data for table `usuarios`
--

/*!40000 ALTER TABLE `usuarios` DISABLE KEYS */;
INSERT INTO `usuarios` (`ID`,`Nombre`,`Usuario`,`Contraseña`,`empeño`,`empeñoautos`,`desempeños`,`refrendos`,`ubicacion`,`ventas`,`busqueda`,`conceptos`,`cortecaja`,`balance`,`repfinanciero`,`cierresucursal`,`grupos`,`dotacion`,`devolucion`,`inventariofisico`,`existencias`,`etiquetas`,`exporinformacion`,`repcontable`,`repauditoria`,`repauxiliar`,`repventas`,`repinventarios`,`repvencidos`,`rephistorico`,`repempeños`,`repasistencia`,`movimientocaja`,`movimientobanco`,`transferencias`,`remates`,`gastos`,`parametros`,`capboletas`,`usuarios`,`cancelbol`,`repgastos`,`catdivisas`,`cotizacion`,`comvendiv`,`repdivisas`,`movidiv`,`facturacion`,`cotizarempeño`,`reporteremates`,`abonar`,`precio`,`modificarcorte`,`hacercorte`,`interesrefrendo`,`interesdesempeño`,`IDUsuario`,`AnaliClientes`,`RegUbicacion`,`RepAlmoneda`,`RepCierres`,`RepIngresos`,`CancelVenta`,`CambioVenta`,`PagoDemasia`,`RepApartado`,`RepUtilidad`,`EntradaInven`,`SalidaInven`,`Deslotifica`,`TrasInven`,`ListaPrecio`,`RepCompras`,`RepTras`,`Kardex`,`RepAnti`,`RepEnve`,`RepEnveP`,`RepSalida`,`RepAutorizaciones`,`RepCierreSucursal`,`RepPrendasSimi`,`RepAleatoria`,`RepPrendasAudi`,`Traspasos`,`Sucursales`,`CatTipos`,`CatFamilias`,`CatSubFamilias`,`CatMedios`,`CatCuentasGas`,`CancelarGas`,`CargosAbonos`,`CatClientes`,`MoviBoveda`,`MostrarApartados`,`ApartadosVencidos`,`EntradasInventario`,`SalidasInventario`,`PrecioVitrina`,`TipoPrenda`,`PreciosKilataje`,`TarjetaBeneficio`,`DescuentoVentas`,`RecalculoPrecios`,`PrestamoBoleta1`,`Estatus`,`PagosFijos`,`CambioPlan`,`CierreDivisas`,`RepCartera`,`VenCliente`,`EtiInven`,`RepDota`,`RepDesempenos`,`RepRefrendos`,`RepHorarios`,`RepPartidaBoveda`,`RepAseguradora`,`RepCancelaciones`,`RepEmpeProm`,`RepDesemProm`,`RepRefProm`,`ConTipoTasa`,`ConVencidos`,`ConStatus`,`PrestamoMes`,`Medios`,`ConfiguraTasas`,`ConfiguraDiam`,`Catalogos`,`MensajeContratos`,`ConexionSuc`,`GeneraAutoriza`,`CatElec`,`RefrendarVencidos`) VALUES 
 (1,'GERENTE SUCURSAL','gerente','gerente',1,1,1,1,0,1,1,1,1,1,1,1,1,1,0,1,1,1,0,1,1,1,1,1,1,1,1,0,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,0,1,1,1,1,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,1,1,1,1,1,0,0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,0,0);
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
) ENGINE=MyISAM DEFAULT CHARSET=latin1;

--
-- Dumping data for table `cortecuentas`
--

/*!40000 ALTER TABLE `cortecuentas` DISABLE KEYS */;
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
CREATE DEFINER=`supervaro`@`%` PROCEDURE `spRepComisiones`(IN FechaIni DATE, IN FechaFin DATE)
BEGIN SELECT SUM(abonos.Importe) AS Abonos,ventas.ID AS IDVenta,ventas.Folio,abonos.Fecha AS FechaAbono, ventas.IVA,ventas.Descuento,ventas.Vencimiento,ventas.Total,ventas.Pagado,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente, CONCAT(vendedores.Nombre,' ',vendedores.Apellidos) AS Vendedor FROM ventas left join abonos ON ventas.ID=abonos.IDVenta INNER JOIN clientes ON ventas.IDCliente=clientes.ID LEFT JOIN vendedores ON ventas.IDVendedor=vendedores.ID WHERE DATE_FORMAT(abonos.Fecha,'%Y%-%m%-%d') BETWEEN FechaIni AND FechaFin AND abonos.Cancelado=0 AND ventas.Cancelado=0 AND ventas.Apartado=1 GROUP BY ventas.ID; End $$
/*!50003 SET SESSION SQL_MODE=@TEMP_SQL_MODE */  $$

DELIMITER ;

--
-- Definition of procedure `spRepVencidos`
--

DROP PROCEDURE IF EXISTS `spRepVencidos`;

DELIMITER $$

/*!50003 SET @TEMP_SQL_MODE=@@SQL_MODE, SQL_MODE='STRICT_TRANS_TABLES,NO_AUTO_CREATE_USER,NO_ENGINE_SUBSTITUTION' */ $$
CREATE DEFINER=`supervaro`@`%` PROCEDURE `spRepVencidos`(IN FechaIni DATE, IN FechaFin DATE, IN DiasEnajena INTEGER, IN TipoContrato INTEGER, IN TipoPrenda INTEGER)
BEGIN SELECT DISTINCT e.ID,e.NumContrato,e.Fecha,e.Vencimiento,CONCAT(c.Nombre,' ',c.Apellido) AS Cliente,e.Avaluo,e.Serie,e.Prestamo,e.TipoInteres,e.TipoTasa FROM empeno e INNER JOIN clientes c ON e.IDCliente=c.ID LEFT JOIN detallesempeno de ON e.ID=de.IDEmpeno WHERE if(TipoContrato=1,(e.Serie=1 OR e.Serie=2 OR e.Serie=3),if(TipoContrato=3,e.Serie=2,de.Tipo=TipoPrenda)) AND DATE_FORMAT(ADDDATE(e.Vencimiento,INTERVAL if(e.TipoTasa='DIARIA',0,DiasEnajena) DAY),'%Y%/%m%/%d') BETWEEN FechaIni AND FechaFin AND e.Cancelado=0 AND e.Destino=0 AND e.Pagado=0 ORDER BY NumContrato; END $$
/*!50003 SET SESSION SQL_MODE=@TEMP_SQL_MODE */  $$

DELIMITER ;

--
-- Definition of view `vwapartadosrematados`
--

DROP TABLE IF EXISTS `vwapartadosrematados`;
DROP VIEW IF EXISTS `vwapartadosrematados`;
CREATE ALGORITHM=UNDEFINED DEFINER=`supervaro`@`%` SQL SECURITY DEFINER VIEW `vwapartadosrematados` AS select sum(`abonos`.`Importe`) AS `Abonos`,`ventas`.`ID` AS `ID`,`ventas`.`Fecha` AS `Fecha`,`ventas`.`FechaMovimiento` AS `FechaMovimiento`,`ventas`.`Folio` AS `Folio`,`ventas`.`IVA` AS `IVA`,`ventas`.`Vencimiento` AS `Vencimiento`,`ventas`.`Total` AS `Total`,`ventas`.`Descuento` AS `Descuento`,`ventas`.`Pagado` AS `Pagado`,`ventas`.`Cancelado` AS `Cancelado`,`ventas`.`OrigenCancelacion` AS `OrigenCancelacion`,concat(`clientes`.`Nombre`,' ',`clientes`.`Apellido`) AS `Cliente` from ((`ventas` left join `abonos` on((`ventas`.`ID` = `abonos`.`IDVenta`))) left join `clientes` on((`ventas`.`IDCliente` = `clientes`.`ID`))) where ((`ventas`.`OrigenCancelacion` = 2) and (`ventas`.`Apartado` = 1) and (`ventas`.`Cancelado` = 1) and (`abonos`.`Cancelado` = 0)) group by `ventas`.`ID`;

--
-- Definition of view `vwdetallesempeno`
--

DROP TABLE IF EXISTS `vwdetallesempeno`;
DROP VIEW IF EXISTS `vwdetallesempeno`;
CREATE ALGORITHM=UNDEFINED DEFINER=`supervaro`@`%` SQL SECURITY DEFINER VIEW `vwdetallesempeno` AS select `detallesempeno`.`ID` AS `ID`,`detallesempeno`.`IDEmpeno` AS `IDEmpeno`,`detallesempeno`.`Cantidad` AS `Cantidad`,`detallesempeno`.`Tipo` AS `Tipo`,`detallesempeno`.`Articulo` AS `Articulo`,`detallesempeno`.`Peso` AS `Peso`,`detallesempeno`.`Prestamo` AS `Prestamo`,`detallesempeno`.`Observaciones` AS `Observaciones`,`detallesempeno`.`Estado` AS `Estado`,`detallesempeno`.`Marca` AS `Marca`,`detallesempeno`.`Modelo` AS `Modelo`,`detallesempeno`.`Serie` AS `Serie`,`detallesempeno`.`Tamano` AS `Tamano`,`detallesempeno`.`Color` AS `Color`,`tipo`.`Descripcion` AS `Tipo_DESC`,`kilatajes`.`Descripcion` AS `Kil_DESC` from ((`detallesempeno` left join `tipo` on((`detallesempeno`.`Tipo` = `tipo`.`ID`))) left join `kilatajes` on((`detallesempeno`.`Kilates` = `kilatajes`.`Clave`)));

--
-- Definition of view `vwfacturadiaria`
--

DROP TABLE IF EXISTS `vwfacturadiaria`;
DROP VIEW IF EXISTS `vwfacturadiaria`;
CREATE ALGORITHM=UNDEFINED DEFINER=`supervaro`@`%` SQL SECURITY DEFINER VIEW `vwfacturadiaria` AS select count(`a`.`ID`) AS `NumRegistros`,`a`.`Fecha` AS `Fecha`,sum(`a`.`Importe`) AS `ImporteTotal` from `auxiliar` `a` where ((`a`.`Importe` > 0) and ((`a`.`Cuenta` = '520450') or (`a`.`Cuenta` = '670350') or (`a`.`Cuenta` = '680350') or (`a`.`Cuenta` = '690350'))) group by `a`.`Fecha` order by `a`.`Fecha`;

--
-- Definition of view `vwfacturaventas`
--

DROP TABLE IF EXISTS `vwfacturaventas`;
DROP VIEW IF EXISTS `vwfacturaventas`;
CREATE ALGORITHM=UNDEFINED DEFINER=`supervaro`@`%` SQL SECURITY DEFINER VIEW `vwfacturaventas` AS select count(`a`.`ID`) AS `NumRegistros`,`a`.`Fecha` AS `Fecha`,sum(`a`.`Importe`) AS `ImporteTotal` from `auxiliar` `a` where ((`a`.`Importe` > 0) and (`a`.`Cuenta` = '620450') and (`a`.`Concepto` = 'Ventas')) group by `a`.`Fecha` order by `a`.`Fecha`;

--
-- Definition of view `vwpagosfijos`
--

DROP TABLE IF EXISTS `vwpagosfijos`;
DROP VIEW IF EXISTS `vwpagosfijos`;
CREATE ALGORITHM=UNDEFINED DEFINER=`supervaro`@`%` SQL SECURITY DEFINER VIEW `vwpagosfijos` AS select `pagosfijos`.`IDEmpeno` AS `IDEmpeno`,max(`pagosfijos`.`FechaMovimiento`) AS `FechaMovimiento` from `pagosfijos` where ((`pagosfijos`.`Pagado` = 1) and (`pagosfijos`.`Cancelado` = 0)) group by `pagosfijos`.`IDEmpeno` order by `pagosfijos`.`IDEmpeno`,`pagosfijos`.`ID`;

--
-- Definition of view `vwrepapartados`
--

DROP TABLE IF EXISTS `vwrepapartados`;
DROP VIEW IF EXISTS `vwrepapartados`;
CREATE ALGORITHM=UNDEFINED DEFINER=`supervaro`@`%` SQL SECURITY DEFINER VIEW `vwrepapartados` AS select sum(`abonos`.`Importe`) AS `Abonos`,`ventas`.`ID` AS `ID`,`ventas`.`Fecha` AS `Fecha`,`ventas`.`Folio` AS `Folio`,`ventas`.`IVA` AS `IVA`,`ventas`.`Vencimiento` AS `Vencimiento`,`ventas`.`Total` AS `Total`,`ventas`.`Descuento` AS `Descuento`,`ventas`.`Pagado` AS `Pagado`,`ventas`.`Cancelado` AS `Cancelado`,`ventas`.`OrigenCancelacion` AS `OrigenCancelacion`,concat(`clientes`.`Nombre`,' ',`clientes`.`Apellido`) AS `Cliente` from ((`ventas` left join `abonos` on((`ventas`.`ID` = `abonos`.`IDVenta`))) left join `clientes` on((`ventas`.`IDCliente` = `clientes`.`ID`))) where ((`ventas`.`Apartado` = 1) and (`abonos`.`Cancelado` = 0)) group by `ventas`.`ID`;



/*!40101 SET SQL_MODE=@OLD_SQL_MODE */;
/*!40014 SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS */;
/*!40014 SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
/*!40101 SET CHARACTER_SET_RESULTS=@OLD_CHARACTER_SET_RESULTS */;
/*!40101 SET COLLATION_CONNECTION=@OLD_COLLATION_CONNECTION */;
/*!40101 SET CHARACTER_SET_CLIENT=@OLD_CHARACTER_SET_CLIENT */;
