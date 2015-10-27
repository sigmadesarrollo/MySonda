Attribute VB_Name = "mdlLavadoDinero"

Global vTipoPago As Integer
Global MLD_INSTRUMENTO_MONETARIO As Integer

Global Const MLD_PRESTAMO As Integer = 1
Global Const MLD_METALES As Integer = 2
Global Const MLD_VEHICULOS As Integer = 3
Global Const MLD_INMUEBLES As Integer = 4

Global MLD_PRESTAMO_OPER As Integer

Global MLD_METALES_COMPRA As Integer
Global MLD_METALES_VENTA As Integer

'Global MLD_VEHICULOS_COMPRA_NUEVO As Integer
'Global MLD_VEHICULOS_VENTA_NUEVO As Integer
Global MLD_VEHICULOS_COMPRA_USADO As Integer
Global MLD_VEHICULOS_VENTA_USADO As Integer

Global MLD_INMUEBLES_OPER As Integer


Type TipoOperacion
    Id_Prestamo_Oper As Integer
    Id_Metales_Compra As Integer
    Id_Metales_Venta As Integer
    Id_Vehiculos_Compra_Usado As Integer
    Id_Vehiculos_Venta_Usado As Integer
    Id_Inmuebles_Oper As Integer
    Clave_Prestamo_Oper As Integer
    Clave_Metales_Compra As Integer
    Clave_Metales_Venta As Integer
    Clave_Vehiculos_Compra_Usado As Integer
    Clave_Vehiculos_Venta_Usado As Integer
    Clave_Inmuebles_Oper As Integer
End Type

Type TipoAlerta
    ID As Integer
    Clave As Integer
    Descripcion As String
End Type

Global vTipoOper As TipoOperacion

Public Sub InicializarAlerta(ByRef tAlerta As TipoAlerta, ByVal tModulo As Integer)
    Dim vTabla As String
    
    Select Case tModulo
        Case MLD_PRESTAMO: vTabla = "mld_prestamos_tipo_alertas"
        Case MLD_METALES: vTabla = "mld_metales_tipo_alertas"
        Case MLD_VEHICULOS: vTabla = "mld_vehiculos_tipo_alertas"
        Case MLD_INMUEBLES: vTabla = "mld_inmuebles_tipo_alertas"
    End Select
    
    tAlerta.ID = Val(SacaValor(vTabla, "Id", " WHERE RegDefault=1"))
    tAlerta.Descripcion = ""
End Sub


Public Function minusculas(Codigo As Integer) As Integer
    If Codigo <> 39 Then minusculas = Asc(LCase(Chr(Codigo)))
End Function

'MLD-MODIF.
Public Sub Mostrar_Seleccionar_Cliente(ByVal sNombre As String, ByVal sApellidoPaterno As String, ByVal sApellidoMaterno As String, ByVal frm As Form)
   Dim Seleccionar As New frmSeleccionarClientes
   If Val(SacaValor("Clientes", "COUNT(ID)", " WHERE Nombre LIKE '%" & sNombre & "%' AND Apellido LIKE '%" & Trim(sApellidoPaterno & " " & sApellidoMaterno) & "%'")) > 0 Then
      Seleccionar.Nombre = sNombre
      Seleccionar.Apellido = Trim(sApellidoPaterno & " " & sApellidoMaterno)
      Seleccionar.Show vbModal, frmMDI
      If Seleccionar.IDCliente <> 0 Then frm.Buscar Seleccionar.IDCliente
      Unload Seleccionar
   End If
End Sub



Public Sub AsignarVariablesModulo()
    
    'Instrumento Monetario por Default
    MLD_INSTRUMENTO_MONETARIO = Val(SacaValor("mld_instr_monetarios", "Id", " WHERE RegDefault=1"))
    
    
    With vTipoOper
        
        .Id_Prestamo_Oper = Val(SacaValor("mld_prestamos_tipo_operacion", "Id", " WHERE RegDefault=1"))
        .Id_Metales_Compra = Val(SacaValor("mld_metales_tipo_operacion", "Id", " WHERE Descripcion = 'Compra'"))
        .Id_Metales_Venta = Val(SacaValor("mld_metales_tipo_operacion", "Id", " WHERE Descripcion = 'Venta'"))
        .Id_Vehiculos_Compra_Usado = Val(SacaValor("mld_vehiculos_tipo_operacion", "Id", " WHERE RegDefault=1 AND Descripcion LIKE '%Compra de vehículo usado%'"))
        .Id_Vehiculos_Venta_Usado = Val(SacaValor("mld_vehiculos_tipo_operacion", "Id", " WHERE RegDefault=1 AND Descripcion LIKE '%Venta de vehículo usado%'"))
        .Id_Inmuebles_Oper = Val(SacaValor("mld_inmuebles_tipo_operacion", "Id", " WHERE RegDefault=1"))
        
        .Clave_Prestamo_Oper = Val(SacaValor("mld_prestamos_tipo_operacion", "Clave", " WHERE RegDefault=1"))
        .Clave_Metales_Compra = Val(SacaValor("mld_metales_tipo_operacion", "Clave", " WHERE Descripcion = 'Compra'"))
        .Clave_Metales_Venta = Val(SacaValor("mld_metales_tipo_operacion", "Clave", " WHERE Descripcion = 'Venta'"))
        .Clave_Vehiculos_Compra_Usado = Val(SacaValor("mld_vehiculos_tipo_operacion", "Clave", " WHERE RegDefault=1 AND Descripcion LIKE '%Compra de vehículo usado%'"))
        .Clave_Vehiculos_Venta_Usado = Val(SacaValor("mld_vehiculos_tipo_operacion", "Clave", " WHERE RegDefault=1 AND Descripcion LIKE '%Venta de vehículo usado%'"))
        .Clave_Inmuebles_Oper = Val(SacaValor("mld_inmuebles_tipo_operacion", "Clave", " WHERE RegDefault=1"))
        
    End With
        
    'MLD_PRESTAMO_OPER = Val(SacaValor("mld_prestamos_tipo_operacion", "Id", " WHERE RegDefault=1"))
    
    'MLD_METALES_COMPRA = Val(SacaValor("mld_metales_tipo_operacion", "Id", " WHERE Descripcion = 'Compra'"))
    'MLD_METALES_VENTA = Val(SacaValor("mld_metales_tipo_operacion", "Id", " WHERE Descripcion = 'Venta'"))
    
    ''MLD_VEHICULOS_COMPRA_NUEVO = SacaValor("mdl_vehiculos_tipo_operacion", "Id", " WHERE RegDefault=1")
    ''MLD_VEHICULOS_VENTA_NUEVO = SacaValor("mdl_vehiculos_tipo_operacion", "Id", " WHERE RegDefault=1")
    'MLD_VEHICULOS_COMPRA_USADO = Val(SacaValor("mld_vehiculos_tipo_operacion", "Id", " WHERE RegDefault=1 AND Descripcion LIKE '%Compra de vehículo usado%'"))
    'MLD_VEHICULOS_VENTA_USADO = Val(SacaValor("mld_vehiculos_tipo_operacion", "Id", " WHERE RegDefault=1 AND Descripcion LIKE '%Venta de vehículo usado%'"))
    
    'MLD_INMUEBLES_OPER = Val(SacaValor("mld_inmuebles_tipo_operacion", "Id", " WHERE RegDefault=1"))
    
End Sub


'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++++++++++++++++++++++++++++++++++++++      RUTINAS PARA EL LAVADO DE DINERO         ++++++++++++++++++++++++++++++++++++++++++++++++++
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

'Creo los campos
Public Sub Crear_Campos_LavadoDinero()

On Error Resume Next
    
    'If (App.Major & "." & App.Minor & "." & App.Revision) <> Trim(Regresa_Valor("MONTEPIO", "Version", "")) Then
    
        '------------------------------------------------------------------------------------
        '---- FOLIOS AVISOS ----
        dbDatos.Execute "CREATE TABLE `mld_folio_avisos` (AnoAviso int(4) DEFAULT NULL,FolioAviso int(14) DEFAULT NULL) ENGINE=MyISAM"
        
        '---- USUARIOS ----
        dbDatos.Execute "ALTER TABLE usuarios ADD COLUMN mld_parametros INT(2) DEFAULT 0, ADD COLUMN mld_movatipicos INT(2) DEFAULT 0, ADD COLUMN mld_expclientes INT(2) DEFAULT 0, ADD COLUMN mld_reppormenorizado INT(2) DEFAULT 0;"
        
        '---- CLIENTES -----
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN ApellidoPaterno VARCHAR(50) DEFAULT '' AFTER Apellido;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN ApellidoMaterno VARCHAR(50) DEFAULT '' AFTER ApellidoPaterno;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN RazonSocial VARCHAR(180) DEFAULT '' AFTER ApellidoMaterno;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN PersonaFisica TINYINT(1) DEFAULT 1 AFTER RazonSocial;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN FechaAltaRazonSocial DATE AFTER PersonaFisica;"
        
        
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN NoExterior VARCHAR(8) DEFAULT '' AFTER Direccion;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN NoInterior VARCHAR(8) DEFAULT '' AFTER NoExterior;"
        
        dbDatos.Execute "ALTER TABLE clientes ADD Column IdOcupacion INTEGER Default 0"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN IdEstadoNac INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN IdPaisNacimiento INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN IdPaisNacionalidad INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN IdTipoIdent INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER TABLE clientes ADD Column DescIdentificacionOtro VARCHAR(200) DEFAULT '';"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN FechaExpIdent DATE;"
        dbDatos.Execute "ALTER TABLE clientes ADD Column Curp VARCHAR(30) DEFAULT '' AFTER Rfc;"
        dbDatos.Execute "ALTER TABLE clientes ADD Column Email VARCHAR(50) DEFAULT '';"
        'Datos para Representante legal cuando el Cliente es Persona Moral
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN RL_Nombre VARCHAR(50) DEFAULT '';"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN RL_ApellidoPaterno VARCHAR(50) DEFAULT '';"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN RL_ApellidoMaterno VARCHAR(50) DEFAULT '';"
        dbDatos.Execute "ALTER TABLE clientes ADD Column RL_Rfc VARCHAR(30) DEFAULT '';"
        dbDatos.Execute "ALTER TABLE clientes ADD Column RL_Curp VARCHAR(30) DEFAULT '';"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN IdTipoAlerta INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER TABLE clientes ADD COLUMN DescTipoAlerta VARCHAR(3000) DEFAULT '';"
        
        
        '---- PARAMETROS ----
        dbDatos.Execute "ALTER Table parametros ADD Column IDActividadVulnerable INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER Table parametros ADD Column IdTipoGiroMercantil INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER Table parametros ADD Column IDTipoMonedaLocal INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER Table parametros ADD Column PrestamoCheque DOUBLE(15,5) DEFAULT 0.00000;"
        dbDatos.Execute "ALTER Table parametros ADD Column CompraCheque DOUBLE(15,5) DEFAULT 0.00000;"
        dbDatos.Execute "ALTER Table parametros ADD Column ImporteSalario DOUBLE(15,5) DEFAULT 0.00000;"
        dbDatos.Execute "ALTER Table parametros ADD Column ImporteUdi DOUBLE(15,6) DEFAULT 0.000000;"
        dbDatos.Execute "ALTER TABLE parametros ADD Column NumConstancia VARCHAR(30) DEFAULT '';"
        dbDatos.Execute "ALTER TABLE parametros ADD COLUMN RutaArchivosXML VARCHAR(250) DEFAULT '';"
        dbDatos.Execute "ALTER Table parametros ADD Column ImporteVSMPrestamos INTEGER(10) DEFAULT 1605;" 'ImporteVSMPrestamos
        
        '---- MOVIMIENTOS ----
        dbDatos.Execute "ALTER Table movimientos ADD Column FolioAvisosLavado INT(10) DEFAULT 1;" 'FolioAvisosLavado
        
        '---- SUCURSALES ----
        dbDatos.Execute "ALTER TABLE sucursales ADD Column Email VARCHAR(80) DEFAULT '' AFTER Cp;"
        
        '---- TIPOS -----
        dbDatos.Execute "ALTER TABLE tipo ADD COLUMN IdTipoGarantia INT(10) DEFAULT 0 AFTER Ordenamiento;"
        dbDatos.Execute "ALTER TABLE tipo ADD COLUMN IdTipoUnidad INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER TABLE tipo ADD COLUMN IdTipoBienes INT(10) DEFAULT 0 AFTER IdTipoGarantia;"
        
        '----------------------------------------------------------------------
        '------------------------------ EMPEÑOS -------------------------------
        '----------------------------------------------------------------------
        dbDatos.Execute "ALTER Table empeno ADD Column IDCoTitular INT(10) DEFAULT 0"
        dbDatos.Execute "ALTER Table empeno ADD Column Cheque TINYINT DEFAULT 0"
        dbDatos.Execute "ALTER Table empeno ADD Column SalarioMin TINYINT DEFAULT 0"
        'Valores de Unidades Monetarias
        dbDatos.Execute "ALTER Table empeno ADD Column ValorSalarioMin DOUBLE(15,5) DEFAULT 0.00000"
        dbDatos.Execute "ALTER Table empeno ADD Column ValorUDI DOUBLE(15,6) DEFAULT 0.000000"
        dbDatos.Execute "ALTER Table empeno ADD Column UltDigitosTarj VARCHAR(4) DEFAULT ''"
        dbDatos.Execute "ALTER Table empeno ADD Column IDTipoOperacion INT(10) DEFAULT 0"
        dbDatos.Execute "ALTER Table empeno ADD Column ClaveTipoOperacion INT(10) DEFAULT 0"
        dbDatos.Execute "ALTER Table empeno ADD Column IDInstrumentoMonetario INT(10) DEFAULT 0"
        dbDatos.Execute "ALTER Table empeno ADD Column IDTipoMoneda INT(10) DEFAULT 0"
        dbDatos.Execute "ALTER TABLE empeno ADD COLUMN IdTipoAlerta INT(10) DEFAULT 0;"
        dbDatos.Execute "ALTER TABLE empeno ADD COLUMN DescTipoAlerta VARCHAR(3000) DEFAULT '';"
        
        
        '---- DETALLES DE EMPEÑOS -----
        dbDatos.Execute "ALTER Table detallesempeno ADD Column IDTipoGarantia INT(10) DEFAULT 0"
        
        
        '---- DETALLES DE EMPEÑOS AUTOS -----
        dbDatos.Execute "ALTER Table detallesempenoautos ADD Column Marca VARCHAR(50) DEFAULT '' AFTER MarcayModelo;"
        dbDatos.Execute "ALTER Table detallesempenoautos ADD Column Modelo VARCHAR(50) DEFAULT '' AFTER Marca;"
        dbDatos.Execute "ALTER Table detallesempenoautos ADD Column VIN VARCHAR(30) DEFAULT '' AFTER SerieChasis;"
        dbDatos.Execute "ALTER Table detallesempenoautos ADD Column RePuVe VARCHAR(30) DEFAULT '' AFTER VIN;"
        dbDatos.Execute "ALTER Table detallesempenoautos ADD Column IDTipoGarantia INT(10) DEFAULT 0"
        dbDatos.Execute "ALTER Table detallesempenoautos ADD Column IDTipoBlindajeAutos INT(10) DEFAULT 0"
        
        
        '---- DETALLES DE EMPEÑOS INMUEBLES -----
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column IDTipoGarantia INT(10) DEFAULT 0"
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column DescInmuebleOtro VARCHAR(50) DEFAULT '';"
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column FolioCatastral VARCHAR(20) DEFAULT '';"
        
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column NoExterior VARCHAR(8) DEFAULT '' AFTER Direccion;"
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column NoInterior VARCHAR(8) DEFAULT '' AFTER NoExterior;"
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column SuperficieConstruccion DOUBLE(10,2) DEFAULT 0.00 AFTER Superficie;"
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column IdTipoBlindajeInmueble INTEGER(10) DEFAULT 0;"
        
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column FolioInstrumentoPublico VARCHAR(20) DEFAULT '';"
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column FechaInstrumentoPublico DATE DEFAULT NULL;"
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column NumNotarioPublico VARCHAR(20) DEFAULT '';"
        dbDatos.Execute "ALTER Table detallesempenoinmuebles ADD Column IdEntidadFederativa INTEGER(3) DEFAULT 0;"
        
        
'        '----------------------------------------------------------------------
'        '------------------------ VENTAS / APARTADOS --------------------------
'        '----------------------------------------------------------------------
'        dbDatos.Execute "ALTER Table ventas ADD Column ValorSalarioMin DOUBLE(15,5) DEFAULT 0.00000"
'        dbDatos.Execute "ALTER Table ventas ADD Column ValorUDI DOUBLE(15,6) DEFAULT 0.000000"
'        dbDatos.Execute "ALTER Table ventas ADD Column UltDigitosTarj VARCHAR(4) DEFAULT ''"
'        'dbDatos.Execute "ALTER Table ventas ADD Column IDTipoOperacion INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table ventas ADD Column IDInstrumentoMonetario INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table ventas ADD Column IDTipoMoneda INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER TABLE ventas ADD COLUMN IdTipoAlerta INT(10) DEFAULT 0;"
'        dbDatos.Execute "ALTER TABLE ventas ADD COLUMN DescTipoAlerta VARCHAR(3000) DEFAULT '';"
'
'
'        '---- DETALLES DE VENTAS -----
'        dbDatos.Execute "ALTER Table detallesventas ADD Column IDTipoGarantia INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table detallesventas ADD Column IDTipoOperacion INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table detallesventas ADD Column ClaveTipoOperacion INT(10) DEFAULT 0"
'
'        '----------------------------------------------------------------------
'        '------------------------ ABONOS APARTADOS --------------------------
'        '----------------------------------------------------------------------
'        dbDatos.Execute "ALTER Table abonos ADD Column ValorSalarioMin DOUBLE(15,5) DEFAULT 0.00000"
'        dbDatos.Execute "ALTER Table abonos ADD Column ValorUDI DOUBLE(15,6) DEFAULT 0.000000"
'        dbDatos.Execute "ALTER Table abonos ADD Column UltDigitosTarj VARCHAR(4) DEFAULT ''"
'        dbDatos.Execute "ALTER Table abonos ADD Column IDTipoOperacion INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table abonos ADD Column ClaveTipoOperacion INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table abonos ADD Column IDInstrumentoMonetario INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table abonos ADD Column IDTipoMoneda INT(10) DEFAULT 0"
'
'
'        '----------------------------------------------------------------------
'        '------------------------      COMPRAS       --------------------------
'        '----------------------------------------------------------------------
'        dbDatos.Execute "ALTER Table compras ADD Column ValorSalarioMin DOUBLE(15,5) DEFAULT 0.00000"
'        dbDatos.Execute "ALTER Table compras ADD Column ValorUDI DOUBLE(15,6) DEFAULT 0.000000"
'        dbDatos.Execute "ALTER Table compras ADD Column UltDigitosTarj VARCHAR(4) DEFAULT ''"
'        'dbDatos.Execute "ALTER Table compras ADD Column IDTipoOperacion INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table compras ADD Column IDInstrumentoMonetario INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table compras ADD Column IDTipoMoneda INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER TABLE compras ADD COLUMN IdTipoAlerta INT(10) DEFAULT 0;"
'        dbDatos.Execute "ALTER TABLE compras ADD COLUMN DescTipoAlerta VARCHAR(3000) DEFAULT '';"
'
'
'        '---- DETALLES DE COMPRAS -----
'        dbDatos.Execute "ALTER Table detallescompras ADD Column IDTipoGarantia INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table detallescompras ADD Column IDTipoOperacion INT(10) DEFAULT 0"
'        dbDatos.Execute "ALTER Table detallescompras ADD Column ClaveTipoOperacion INT(10) DEFAULT 0"
'
'        '------------------------------------------------------------------------------------
        
                
        dbDatos.Execute "ALTER TABLE usuarios ADD COLUMN RepIdentClientes INT(2) DEFAULT 0;"
        
        dbDatos.Execute "CREATE TABLE `basedatos`.`ocupaciones` (`ID` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT, `descripcion` VARCHAR(50), `estatus` TINYINT DEFAULT 1, `Ordenamiento` INT(2) DEFAULT 0, `Actualizar` TINYINT DEFAULT 0, PRIMARY KEY (`ID`))ENGINE = MyISAM;"
        'MOD. Catalogo de Identificaciones
        dbDatos.Execute "CREATE TABLE `basedatos`.`Identificaciones` (" & _
                        "`ID` INTEGER UNSIGNED NOT NULL AUTO_INCREMENT," & _
                        "`Identificacion` VARCHAR(80)," & _
                        "PRIMARY KEY (`ID`)" & _
                        ")" & _
                        "ENGINE = MyISAM;"
                        
        'MOD. Catalogo de Estados de Pais para CURP
        dbDatos.Execute "CREATE TABLE `estadospais` (`ID` int(10) unsigned NOT NULL AUTO_INCREMENT,`Codigo` varchar(2) DEFAULT NULL,`descripcion` varchar(50) DEFAULT NULL,`Ordenamiento` int(2) DEFAULT '0',`Actualizar` tinyint(4) DEFAULT '0',PRIMARY KEY (`ID`)) ENGINE=MyISAM"
        
        CrearTablaCURP
        
        AsignarVariablesModulo
        
        dbDatos.Execute "UPDATE empeno SET IdTipoAlerta=1, IdTipoOperacion=" & vTipoOper.Id_Prestamo_Oper & ", ClaveTipoOperacion=" & vTipoOper.Clave_Prestamo_Oper & ",IdInstrumentoMonetario=" & MLD_INSTRUMENTO_MONETARIO & " WHERE IdTipoAlerta = 0;"
        
        
        
        
    'End If
    
End Sub


Public Sub CrearTablaCURP()
    Dim Rs As New ADODB.Recordset
    Dim i As Integer
    
    dbDatos.Execute "CREATE TABLE IF NOT EXISTS tablacurp ( `Indice` int(5) DEFAULT NULL, `Valor` varchar(2) DEFAULT NULL) ENGINE=MyISAM"

    Rs.Open "SELECT COUNT(Indice) as Valor FROM tablacurp;", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not Rs.EOF Then
        
        If Rs!Valor <> 37 Then
            dbDatos.Execute "TRUNCATE TABLE tablacurp;"
            For i = 0 To 36
                If i <= 9 Then
                    dbDatos.Execute "INSERT INTO tablacurp (Indice,Valor) VALUES (" & i & ",'" & i & "');"
                Else
                    If i = 24 Then
                        dbDatos.Execute "INSERT INTO tablacurp (Indice,Valor) VALUES (" & i & ",'" & Chr(209) & "');"
                    Else
                        dbDatos.Execute "INSERT INTO tablacurp (Indice,Valor) VALUES (" & i & ",'" & Chr(55 + i + IIf(i > 24, -1, 0)) & "');"
                    End If
                End If
            Next
        End If
        
    End If
    Rs.Close
    Set Rs = Nothing
    
    
    
    dbDatos.Execute "CREATE TABLE IF NOT EXISTS `estadospais` (`ID` int(10) unsigned NOT NULL AUTO_INCREMENT,`Codigo` varchar(2) DEFAULT NULL,`Descripcion` varchar(50) DEFAULT NULL,`Ordenamiento` int(2) DEFAULT '0',`Actualizar` tinyint(4) DEFAULT '0',PRIMARY KEY (`ID`)) ENGINE=MyISAM"
    Rs.Open "SELECT COUNT(ID) as Valor FROM estadospais;", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not Rs.EOF Then
    
        If Rs!Valor <> 32 Then
            dbDatos.Execute "TRUNCATE TABLE estadospais;"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('AS','AGUASCALIENTES',1,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('BC','BAJA CALIFORNIA',2,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('BS','BAJA CALIFORNIA SUR',3,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('CC','CAMPECHE',4,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('CS','CHIAPAS',5,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('CH','CHIHUAHUA',6,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('CL','COAHUILA',7,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('CM','COLIMA',8,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('DF','DISTRITO FEDERAL',9,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('DG','DURANGO',10,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('GT','GUANAJUATO',11,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('GR','GUERRERO',12,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('HG','HIDALGO',13,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('JC','JALISCO',14,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('MC','MEXICO',15,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('MN','MICHOACAN',16,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('MS','MORELOS',17,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('NT','NAYARIT',18,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('NL','NUEVO LEON',19,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('OC','OAXACA',20,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('PL','PUEBLA',21,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('QT','QUERETARO',22,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('QR','QUINTANA ROO',23,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('SP','SAN LUIS POTOSI',24,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('SL','SINALOA',25,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('SR','SONORA',26,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('TC','TABASCO',27,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('TS','TAMAULIPAS',28,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('TL','TLAXCALA',29,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('VZ','VERACRUZ',30,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('YN','YUCATAN',31,0);"
            dbDatos.Execute "INSERT INTO estadospais (Codigo,Descripcion,Ordenamiento,Actualizar) VALUES ('ZS','ZACATECAS',32,0);"
            
        End If
    
    End If
    Rs.Close
    Set Rs = Nothing
    
    
End Sub



Public Function GenerarCURP(ByVal sNombre As String, ByVal sApellidoP As String, ByVal sApellidoM As String, ByVal Genero As String, ByVal FechaNac As Date, ByVal sEstadoNac As String) As String
    
    Dim sCURP(20) As String
    Dim sLetra(3) As String
    Dim vLetra(8) As String
    Dim sCodEdo As String
    Dim sFecha As String
    Dim aTabla(17) As String, i As Integer
    Dim cifra As Integer
    Dim DigitoVerificador As Integer
        
    vLetra(8) = IIf(Len(sNombre) >= 4, IIf(Mid(sNombre, 4, 1) <> "A", IIf(Mid(sNombre, 4, 1) <> "E", IIf(Mid(sNombre, 4, 1) <> "I", IIf(Mid(sNombre, 4, 1) <> "O", IIf(Mid(sNombre, 4, 1) <> "U", Mid(sNombre, 4, 1), "X"), "X"), "X"), "X"), "X"), "X")
    vLetra(7) = IIf(Len(sNombre) >= 3, IIf(Mid(sNombre, 3, 1) <> "A", IIf(Mid(sNombre, 3, 1) <> "E", IIf(Mid(sNombre, 3, 1) <> "I", IIf(Mid(sNombre, 3, 1) <> "O", IIf(Mid(sNombre, 3, 1) <> "U", Mid(sNombre, 3, 1), vLetra(8)), vLetra(8)), vLetra(8)), vLetra(8)), vLetra(8)), "X")
    vLetra(6) = IIf(Len(sApellidoM) >= 4, IIf(Mid(sApellidoM, 4, 1) <> "A", IIf(Mid(sApellidoM, 4, 1) <> "E", IIf(Mid(sApellidoM, 4, 1) <> "I", IIf(Mid(sApellidoM, 4, 1) <> "O", IIf(Mid(sApellidoM, 4, 1) <> "U", Mid(sApellidoM, 4, 1), "X"), "X"), "X"), "X"), "X"), "X")
    vLetra(5) = IIf(Len(sApellidoM) >= 3, IIf(Mid(sApellidoM, 3, 1) <> "A", IIf(Mid(sApellidoM, 3, 1) <> "E", IIf(Mid(sApellidoM, 3, 1) <> "I", IIf(Mid(sApellidoM, 3, 1) <> "O", IIf(Mid(sApellidoM, 3, 1) <> "U", Mid(sApellidoM, 3, 1), vLetra(6)), vLetra(6)), vLetra(6)), vLetra(6)), vLetra(6)), "X")
    vLetra(4) = IIf(Len(sApellidoP) >= 4, IIf(Mid(sApellidoP, 4, 1) <> "A", IIf(Mid(sApellidoP, 4, 1) <> "E", IIf(Mid(sApellidoP, 4, 1) <> "I", IIf(Mid(sApellidoP, 4, 1) <> "O", IIf(Mid(sApellidoP, 4, 1) <> "U", Mid(sApellidoP, 4, 1), "X"), "X"), "X"), "X"), "X"), "X")
    vLetra(3) = IIf(Len(sApellidoP) >= 3, IIf(Mid(sApellidoP, 3, 1) <> "A", IIf(Mid(sApellidoP, 3, 1) <> "E", IIf(Mid(sApellidoP, 3, 1) <> "I", IIf(Mid(sApellidoP, 3, 1) <> "O", IIf(Mid(sApellidoP, 3, 1) <> "U", Mid(sApellidoP, 3, 1), vLetra(4)), vLetra(4)), vLetra(4)), vLetra(4)), vLetra(4)), "X")
    vLetra(2) = IIf(Len(sApellidoP) >= 2, IIf(Mid(sApellidoP, 4, 1) <> "A", IIf(Mid(sApellidoP, 4, 1) <> "E", IIf(Mid(sApellidoP, 4, 1) <> "I", IIf(Mid(sApellidoP, 4, 1) <> "O", IIf(Mid(sApellidoP, 4, 1) <> "U", sLetra(2), Mid(sApellidoP, 4, 1)), Mid(sApellidoP, 4, 1)), Mid(sApellidoP, 4, 1)), Mid(sApellidoP, 4, 1)), Mid(sApellidoP, 4, 1)), "X")
    vLetra(1) = IIf(Len(sApellidoP) >= 2, IIf(Mid(sApellidoP, 3, 1) <> "A", IIf(Mid(sApellidoP, 3, 1) <> "E", IIf(Mid(sApellidoP, 3, 1) <> "I", IIf(Mid(sApellidoP, 3, 1) <> "O", IIf(Mid(sApellidoP, 3, 1) <> "U", vLetra(2), Mid(sApellidoP, 3, 1)), Mid(sApellidoP, 3, 1)), Mid(sApellidoP, 3, 1)), Mid(sApellidoP, 3, 1)), Mid(sApellidoP, 3, 1)), "X")
    
    sLetra(1) = IIf(Len(sApellidoP) >= 2, IIf(Mid(sApellidoP, 2, 1) <> "A", IIf(Mid(sApellidoP, 2, 1) <> "E", IIf(Mid(sApellidoP, 2, 1) <> "I", IIf(Mid(sApellidoP, 2, 1) <> "O", IIf(Mid(sApellidoP, 2, 1) <> "U", Mid(sApellidoP, 2, 1), vLetra(3)), vLetra(3)), vLetra(3)), vLetra(3)), vLetra(3)), "X")
    sLetra(2) = IIf(Len(sApellidoM) >= 3, IIf(Mid(sApellidoM, 2, 1) <> "A", IIf(Mid(sApellidoM, 2, 1) <> "E", IIf(Mid(sApellidoM, 2, 1) <> "I", IIf(Mid(sApellidoM, 2, 1) <> "O", IIf(Mid(sApellidoM, 2, 1) <> "U", Mid(sApellidoM, 2, 1), vLetra(5)), vLetra(5)), vLetra(5)), vLetra(5)), vLetra(5)), "X")
    sLetra(3) = IIf(Len(sNombre) >= 2, IIf(Mid(sNombre, 2, 1) <> "A", IIf(Mid(sNombre, 2, 1) <> "E", IIf(Mid(sNombre, 2, 1) <> "I", IIf(Mid(sNombre, 2, 1) <> "O", IIf(Mid(sNombre, 2, 1) <> "U", Mid(sNombre, 2, 1), vLetra(7)), vLetra(7)), vLetra(7)), vLetra(7)), vLetra(7)), "X")

    sFecha = Mid(Year(FechaNac), 3, 2) & Format(Month(FechaNac), "00") & Format(Day(FechaNac), "00")
    sCodEdo = SacaValor("estadospais", "Codigo", " WHERE descripcion='" & Trim(sEstadoNac) & "'")
    
    '-------------------------------------------------------------------------
    
    sCURP(1) = Mid(sApellidoP, 1, 1)
    sCURP(2) = IIf(Len(sApellidoP) >= 2, IIf(Mid(sApellidoP, 2, 1) <> "A", IIf(Mid(sApellidoP, 2, 1) <> "E", IIf(Mid(sApellidoP, 2, 1) <> "I", IIf(Mid(sApellidoP, 2, 1) <> "O", IIf(Mid(sApellidoP, 2, 1) <> "U", vLetra(1), Mid(sApellidoP, 2, 1)), Mid(sApellidoP, 2, 1)), Mid(sApellidoP, 2, 1)), Mid(sApellidoP, 2, 1)), Mid(sApellidoP, 2, 1)), "X")
    sCURP(3) = IIf(sApellidoM = "", "X", Mid(sApellidoM, 1, 1))
    sCURP(4) = Mid(sNombre, 1, 1)
        
    sCURP(5) = Val(Mid(sFecha, 1, 1))
    sCURP(6) = Val(Mid(sFecha, 2, 1))
    sCURP(7) = Val(Mid(sFecha, 3, 1))
    sCURP(8) = Val(Mid(sFecha, 4, 1))
    sCURP(9) = Val(Mid(sFecha, 5, 1))
    sCURP(10) = Val(Mid(sFecha, 6, 1))
        
    sCURP(11) = IIf(Genero = "MASCULINO", "H", "M")
        
    sCURP(12) = Mid(sCodEdo, 1, 1)
    sCURP(13) = Mid(sCodEdo, 2, 1)
        
    sCURP(14) = sLetra(1)
    sCURP(15) = sLetra(2)
    sCURP(16) = sLetra(3)
        
    sCURP(17) = "0"
        
    cifra = 0
    For i = 1 To 17
        cifra = cifra + (Val(SacaValor("tablacurp", "Indice", " WHERE Valor='" & sCURP(i) & "'")) * (19 - i))
    Next i
    DigitoVerificador = (cifra Mod 10) - 10
    sCURP(18) = IIf(DigitoVerificador <= -10, 0, DigitoVerificador * -1)
    
    GenerarCURP = sCURP(1) & sCURP(2) & sCURP(3) & sCURP(4) & sCURP(5) & sCURP(6) & sCURP(7) & sCURP(8) & sCURP(9) & sCURP(10) & sCURP(11) & sCURP(12) & sCURP(13) & IIf(sCURP(14) = "Ñ", "X", sCURP(14)) & IIf(sCURP(15) = "Ñ", "X", sCURP(15)) & IIf(sCURP(16) = "Ñ", "X", sCURP(16)) & sCURP(17) & sCURP(18)
        
End Function



'------------------------------------------------------------------------------------
'------ PROCEDIMIENTO DE GUARDADO DE DATOS PARA REPORTE XML LAVADO DE DINERO --------
'------------------------------------------------------------------------------------
Public Sub GuardarDatosLavadoDinero(ByVal IDMov As Long, ByVal Tabla As String, ByVal IdInstrumentoMonetario As Integer, ByVal tModulo As String, ByVal tVenta As Integer, ByVal pTipoAlerta As Integer, ByVal pDescAlerta As String)
    Dim Rs As New ADODB.Recordset
    
    '****** Parametros ******
    'IDMov: ID de registro de la tabla Empeno-Ventas-Compras-Ventas
    'ModuloSistema:  Modulo de Operacion del sistema (EMPENO-EMPENOINMUEBLE-REFRENDO-DESEMPENO-VENTAS-COMPRA-VENTAMAYORISTA)
    
    'MLD_PRESTAMO As Integer = 1
    'MLD_METALES As Integer = 2
    'MLD_VEHICULOS As Integer = 3
    'MLD_INMUEBLES As Integer = 4
    '
    '************************
    
    Dim IdTipoOperacion As Integer, IdTipoMoneda As Integer
    Dim ClaveTipoOperacion As Integer
    
    'Si la Operacion es de Prestamo
    Select Case tModulo
    
        Case MLD_PRESTAMO
            IdTipoOperacion = vTipoOper.Id_Prestamo_Oper
            ClaveTipoOperacion = vTipoOper.Clave_Prestamo_Oper
        Case MLD_METALES
            If tVenta = True Then
                IdTipoOperacion = vTipoOper.Id_Metales_Venta
                ClaveTipoOperacion = vTipoOper.Clave_Metales_Venta
            Else
                IdTipoOperacion = vTipoOper.Id_Metales_Compra
                ClaveTipoOperacion = vTipoOper.Clave_Metales_Compra
            End If
        Case MLD_VEHICULOS
            If tVenta = True Then
                IdTipoOperacion = vTipoOper.Id_Vehiculos_Venta_Usado
                ClaveTipoOperacion = vTipoOper.Clave_Vehiculos_Venta_Usado
            Else
                IdTipoOperacion = vTipoOper.Id_Vehiculos_Compra_Usado
                ClaveTipoOperacion = vTipoOper.Clave_Vehiculos_Compra_Usado
            End If
        Case MLD_INMUEBLES
            IdTipoOperacion = vTipoOper.Id_Inmuebles_Oper
            ClaveTipoOperacion = vTipoOper.Clave_Inmuebles_Oper
    End Select
    
    IdTipoMoneda = Val(Regresa_Valor_BD("IDTipoMonedaLocal"))
    
    Select Case tModulo
    
        Case MLD_PRESTAMO
            dbDatos.Execute "UPDATE " & Tabla & " SET ValorSalarioMin=" & Val(Regresa_Valor_BD("ImporteSalario")) & ", ValorUDI=" & Val(Regresa_Valor_BD("ImporteUdi")) & ", IdTipoOperacion=" & IdTipoOperacion & ",ClaveTipoOperacion=" & ClaveTipoOperacion & ", IdInstrumentoMonetario=" & IdInstrumentoMonetario & ", IdTipoMoneda=" & IdTipoMoneda & _
                            ",IdTipoAlerta=" & pTipoAlerta & " ,DescTipoAlerta='" & pDescAlerta & "'" & _
                            " WHERE Id=" & IDMov
    
        Case Else
            dbDatos.Execute "UPDATE " & Tabla & " SET ValorSalarioMin=" & Val(Regresa_Valor_BD("ImporteSalario")) & ", ValorUDI=" & Val(Regresa_Valor_BD("ImporteUdi")) & ", IdInstrumentoMonetario=" & IdInstrumentoMonetario & ", IdTipoMoneda=" & IdTipoMoneda & _
                            " WHERE Id=" & IDMov
            
            
    End Select
    
    

End Sub


Public Function LimpiaCad(texto As String) As String
      Dim cadenafinal As String
      Dim columna As Integer
      For columna = 1 To Len(texto)
          Select Case Mid$(texto, columna, 1)
                 Case "á"
                      cadenafinal = cadenafinal + "a"
                 Case "Á"
                      cadenafinal = cadenafinal + "A"
                Case "é"
                      cadenafinal = cadenafinal + "e"
                Case "É"
                      cadenafinal = cadenafinal + "E"
                Case "í"
                      cadenafinal = cadenafinal + "i"
                Case "Í"
                      cadenafinal = cadenafinal + "I"
                Case "ó"
                      cadenafinal = cadenafinal + "o"
                Case "Ó"
                      cadenafinal = cadenafinal + "O"
                Case "ú"
                      cadenafinal = cadenafinal + "u"
                Case "Ú"
                      cadenafinal = cadenafinal + "U"
                Case "ñ"
                      cadenafinal = cadenafinal + "n"
                Case "Ñ"
                      cadenafinal = cadenafinal + "N"
                'Case "'"
                '      cadenafinal = cadenafinal + " "
                Case "`"
                      cadenafinal = cadenafinal + " "
                Case "&"
                      cadenafinal = cadenafinal + " "
                Case "?"
                      cadenafinal = cadenafinal + " "
                Case "¿"
                      cadenafinal = cadenafinal + " "
                Case "¡"
                      cadenafinal = cadenafinal + " "
                Case "!"
                      cadenafinal = cadenafinal + " "
                Case "+"
                      cadenafinal = cadenafinal + " "
                'Case "-"
                '      cadenafinal = cadenafinal + " "
                Case "*"
                      cadenafinal = cadenafinal + " "
                Case "/"
                      cadenafinal = cadenafinal + " "
                Case "#"
                      cadenafinal = cadenafinal + " "
                      
                Case Else
                     cadenafinal = cadenafinal + Mid$(texto, columna, 1)
            
          End Select
                    
      Next columna
      LimpiaCad = cadenafinal
End Function
