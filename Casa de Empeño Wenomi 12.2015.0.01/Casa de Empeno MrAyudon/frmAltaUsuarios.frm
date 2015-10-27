VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmAltaUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAltaUsuarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   13515
   Begin VB.Frame FraMóduloAntilavado 
      Caption         =   "MÓDULO ANTILAVADO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   120
      TabIndex        =   136
      Top             =   2520
      Width           =   2655
      Begin VB.CheckBox chkMldMovAtipicos 
         Appearance      =   0  'Flat
         Caption         =   "Movimientos Atípicos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   141
         Top             =   720
         Width           =   2280
      End
      Begin VB.CheckBox chkMldRepPormenorizado 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Mens. Pormenorizado"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   140
         Top             =   1020
         Width           =   2550
      End
      Begin VB.CheckBox Check3 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de Divisas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   -2520
         TabIndex        =   139
         Top             =   4320
         Width           =   1815
      End
      Begin VB.CheckBox chkMldExpClientes 
         Appearance      =   0  'Flat
         Caption         =   "Expedientes de Clientes"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   138
         Top             =   1320
         Width           =   2160
      End
      Begin VB.CheckBox chkMldParametros 
         Appearance      =   0  'Flat
         Caption         =   "Configuración Parámetros"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   137
         Top             =   360
         Width           =   2265
      End
   End
   Begin VB.TextBox txtNombre 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2685
   End
   Begin VB.TextBox txtPass 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1560
      Width           =   2235
   End
   Begin VB.TextBox txtPass1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      IMEMode         =   3  'DISABLE
      Left            =   120
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox txtUsuario 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   2205
   End
   Begin VB.Frame Frame1 
      Caption         =   "PRIVILEGIOS DEL USUARIO"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8475
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   10590
      Begin VB.CheckBox chkCancelarCierre 
         Appearance      =   0  'Flat
         Caption         =   "Cancelar Cierres de Caja"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   135
         Top             =   8160
         Width           =   2310
      End
      Begin VB.CheckBox chkRefrendarVencidos 
         Appearance      =   0  'Flat
         Caption         =   "Refrendar Vencidos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   134
         Top             =   7800
         Width           =   1920
      End
      Begin VB.CheckBox chkCatElec 
         Appearance      =   0  'Flat
         Caption         =   "Catálogos Electrónicos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   133
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CheckBox chkGeneraAutoriza 
         Appearance      =   0  'Flat
         Caption         =   "Generar Autorizaciones"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   132
         Top             =   7440
         Width           =   2280
      End
      Begin VB.CheckBox chkConexionSuc 
         Appearance      =   0  'Flat
         Caption         =   "Conexión Intersucursales"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   131
         Top             =   7080
         Width           =   2280
      End
      Begin VB.CheckBox chkCatalogos 
         Appearance      =   0  'Flat
         Caption         =   "Catálogos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   130
         Top             =   3840
         Width           =   1470
      End
      Begin VB.CheckBox chkMensajeContratos 
         Appearance      =   0  'Flat
         Caption         =   "Mensajes en Contratos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   129
         Top             =   3480
         Width           =   2055
      End
      Begin VB.CheckBox chkConfiguraDiam 
         Appearance      =   0  'Flat
         Caption         =   "Configuración Precios Diam."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   128
         Top             =   2040
         Width           =   2355
      End
      Begin VB.CheckBox chkConfiguraTasas 
         Appearance      =   0  'Flat
         Caption         =   "Configuración Tasas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   127
         Top             =   1320
         Width           =   1920
      End
      Begin VB.CheckBox chkConTipoTasa 
         Appearance      =   0  'Flat
         Caption         =   "Contratos por Tipo de Tasa"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   126
         Top             =   4920
         Width           =   2490
      End
      Begin VB.CheckBox chkConVencidos 
         Appearance      =   0  'Flat
         Caption         =   "Contratos Vencidos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   125
         Top             =   5280
         Width           =   2490
      End
      Begin VB.CheckBox chkConStatus 
         Appearance      =   0  'Flat
         Caption         =   "Contratos por Status"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   124
         Top             =   5640
         Width           =   2490
      End
      Begin VB.CheckBox chkPrestamoMes 
         Appearance      =   0  'Flat
         Caption         =   "Préstamos por Mes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   123
         Top             =   6000
         Width           =   2490
      End
      Begin VB.CheckBox chkMedios 
         Appearance      =   0  'Flat
         Caption         =   "Medios de Difusión"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   122
         Top             =   6360
         Width           =   2490
      End
      Begin VB.CheckBox chkRepEmpeProm 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Empeños Promedio"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   121
         Top             =   3840
         Width           =   2355
      End
      Begin VB.CheckBox chkRepDesemProm 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Desempeños Promedio"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   120
         Top             =   4200
         Width           =   2565
      End
      Begin VB.CheckBox chkRepRefProm 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Refrendos Promedio"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   119
         Top             =   4560
         Width           =   2490
      End
      Begin VB.CheckBox chkRepCancelaciones 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Cancelaciones"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   118
         Top             =   3480
         Width           =   2070
      End
      Begin VB.CheckBox chkRepAseguradora 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Aseguradora"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   117
         Top             =   3120
         Width           =   2070
      End
      Begin VB.CheckBox chkRepPartidaBoveda 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Partidas Bóveda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   116
         Top             =   2760
         Width           =   2160
      End
      Begin VB.CheckBox chkRepHorarios 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Movimientos por Horario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   115
         Top             =   2040
         Width           =   2670
      End
      Begin VB.CheckBox chkRepDesempenos 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Desempeños"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   114
         Top             =   7800
         Width           =   2010
      End
      Begin VB.CheckBox chkRepRefrendos 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Refrendos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   113
         Top             =   8160
         Width           =   2055
      End
      Begin VB.CheckBox chkRepDota 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Dotaciones"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   112
         Top             =   3480
         Width           =   2160
      End
      Begin VB.CheckBox chkEtiInven 
         Appearance      =   0  'Flat
         Caption         =   "Etiquetas Inventario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   111
         Top             =   2760
         Width           =   2055
      End
      Begin VB.CheckBox chkVenCliente 
         Appearance      =   0  'Flat
         Caption         =   "Ventas Clientes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   110
         Top             =   6360
         Width           =   1695
      End
      Begin VB.CheckBox chkRepCartera 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Cartera"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   109
         Top             =   4560
         Width           =   1710
      End
      Begin VB.CheckBox chkCierreDivisas 
         Appearance      =   0  'Flat
         Caption         =   "Cierre de Divisas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   108
         Top             =   4200
         Width           =   1710
      End
      Begin VB.CheckBox chkCambioPlan 
         Appearance      =   0  'Flat
         Caption         =   "Cambiar Plan"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   107
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkPagosFijos 
         Appearance      =   0  'Flat
         Caption         =   "Pagos Fijos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox chkPrestamoBoleta1 
         Appearance      =   0  'Flat
         Caption         =   "Autorizar préstamo límite 1"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   102
         Top             =   6720
         Width           =   2415
      End
      Begin VB.CheckBox chkRecalculo 
         Appearance      =   0  'Flat
         Caption         =   "Recálculo de Precios"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   101
         Top             =   6360
         Width           =   1905
      End
      Begin VB.CheckBox chkDescuentoVentas 
         Appearance      =   0  'Flat
         Caption         =   "Descuento Ventas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   100
         Top             =   6000
         Width           =   1905
      End
      Begin VB.CheckBox chkTipoPrenda 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de Tipos de Prenda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   99
         Top             =   1920
         Width           =   2415
      End
      Begin VB.CheckBox chkPreciosKilataje 
         Appearance      =   0  'Flat
         Caption         =   "Configuración Precios Oro"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   98
         Top             =   1680
         Width           =   2355
      End
      Begin VB.CheckBox chkTarjetaBeneficio 
         Appearance      =   0  'Flat
         Caption         =   "Tarjeta Beneficio"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   97
         Top             =   8760
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkPrecioVitrina 
         Appearance      =   0  'Flat
         Caption         =   "Modificar Precio Vitrina"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   96
         Top             =   5640
         Width           =   2385
      End
      Begin VB.CheckBox chkMostrarApartados 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Apartados Vigentes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   95
         Top             =   7800
         Width           =   2355
      End
      Begin VB.CheckBox chkApartadosVencidos 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Apartados Vencidos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   94
         Top             =   8160
         Width           =   2355
      End
      Begin VB.CheckBox chkMoviBoveda 
         Appearance      =   0  'Flat
         Caption         =   "Movimiento Bóveda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   93
         Top             =   7440
         Width           =   1815
      End
      Begin VB.CheckBox chkCatClientes 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de clientes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   92
         Top             =   2160
         Width           =   1815
      End
      Begin VB.CheckBox chkCargosAbonos 
         Appearance      =   0  'Flat
         Caption         =   "Cargos Abonos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   91
         Top             =   8520
         Width           =   1695
      End
      Begin VB.CheckBox chkCatCuentas 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo Cuentas de Gastos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   90
         Top             =   1200
         Width           =   2385
      End
      Begin VB.CheckBox chkCatDivisas 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de Divisas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   89
         Top             =   5280
         Width           =   1815
      End
      Begin VB.CheckBox chkCatMedios 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de medios"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   88
         Top             =   1440
         Width           =   1935
      End
      Begin VB.CheckBox chkCatSubFamilias 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de subfamilias"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   87
         Top             =   8760
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkCatFamilias 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de familias"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   86
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox chkCatTipos 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de tipos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   85
         Top             =   2640
         Width           =   1935
      End
      Begin VB.CheckBox chkSucursales 
         Appearance      =   0  'Flat
         Caption         =   "Configuración Sucursales"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   84
         Top             =   2760
         Width           =   2205
      End
      Begin VB.CheckBox chkTraspasos 
         Appearance      =   0  'Flat
         Caption         =   "Traspasos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   83
         Top             =   9480
         Width           =   1215
      End
      Begin VB.CheckBox chkPrendasAuditadas 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Prendas Auditadas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   82
         Top             =   9480
         Visible         =   0   'False
         Width           =   2370
      End
      Begin VB.CheckBox chkAleatoriaSelectiva 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Aleatoria/Selectiva"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   81
         Top             =   9720
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.CheckBox chkPrendasSimilares 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Prendas Similares"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   80
         Top             =   9240
         Visible         =   0   'False
         Width           =   2205
      End
      Begin VB.CheckBox chkRepAutorizaciones 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Autorizaciones"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   79
         Top             =   2400
         Width           =   2160
      End
      Begin VB.CheckBox chkRepCierreSucursal 
         Appearance      =   0  'Flat
         Caption         =   "Reportes Cierre de sucursal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   78
         Top             =   9000
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.CheckBox chkRepSalida 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Salidas Inventario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   77
         Top             =   3840
         Width           =   2265
      End
      Begin VB.CheckBox chkRepEnveP 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Envecimiento P."
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   76
         Top             =   4920
         Width           =   2295
      End
      Begin VB.CheckBox chkRepEnve 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Envejecimiento"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   75
         Top             =   4560
         Width           =   2415
      End
      Begin VB.CheckBox chkRepAnti 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Antiguedad"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   74
         Top             =   4200
         Width           =   2055
      End
      Begin VB.CheckBox chkKardex 
         Appearance      =   0  'Flat
         Caption         =   "Kardex"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   73
         Top             =   8760
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CheckBox chkRepTras 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Traspasos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   72
         Top             =   9000
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkRepCompras 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Compras"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   71
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CheckBox chkListaPrecio 
         Appearance      =   0  'Flat
         Caption         =   "Lista de precios"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   70
         Top             =   9480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkTrasInven 
         Appearance      =   0  'Flat
         Caption         =   "Traspasos de Inventario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   9720
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CheckBox chkDeslotifica 
         Appearance      =   0  'Flat
         Caption         =   "Deslotificación"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   68
         Top             =   2040
         Width           =   1695
      End
      Begin VB.CheckBox chkSalidaInven 
         Appearance      =   0  'Flat
         Caption         =   "Salida de Inventario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   67
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox chkEntradaInven 
         Appearance      =   0  'Flat
         Caption         =   "Entrada a Inventario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   66
         Top             =   9960
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CheckBox chkRepUtilidadVen 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Utilidad Ventas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   65
         Top             =   240
         Width           =   2055
      End
      Begin VB.CheckBox chkRepApartado 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Ventas Apartado"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   64
         Top             =   7440
         Width           =   2190
      End
      Begin VB.CheckBox chkPagoDemasia 
         Appearance      =   0  'Flat
         Caption         =   "Pago de Demasias"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   63
         Top             =   6720
         Width           =   1815
      End
      Begin VB.CheckBox chkCambioVenta 
         Appearance      =   0  'Flat
         Caption         =   "Cambios Ventas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   62
         Top             =   9240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.CheckBox chkCancelVenta 
         Appearance      =   0  'Flat
         Caption         =   "Cancelar Venta"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   61
         Top             =   9960
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkRepIngresos 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Ingresos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   60
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox chkRepCierres 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Cierres"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2400
         TabIndex        =   59
         Top             =   8520
         Visible         =   0   'False
         Width           =   1935
      End
      Begin VB.CheckBox chkRepAlmoneda 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Almoneda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   58
         Top             =   3480
         Width           =   1935
      End
      Begin VB.CheckBox chkRegUbicacion 
         Appearance      =   0  'Flat
         Caption         =   "Registrar ubicación"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   2760
         Width           =   1935
      End
      Begin VB.CheckBox chkAnaliClientes 
         Appearance      =   0  'Flat
         Caption         =   "Análisis de clientes"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   8520
         Width           =   1695
      End
      Begin VB.CheckBox chkInteresrefrendo 
         Appearance      =   0  'Flat
         Caption         =   "Modificar interes refrendo"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   55
         Top             =   8520
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.CheckBox chkInteresdesempeño 
         Appearance      =   0  'Flat
         Caption         =   "Modificar interes desempeño"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   54
         Top             =   8880
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CheckBox chkHacercorte 
         Appearance      =   0  'Flat
         Caption         =   "Realizar Cierre de Caja"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   52
         Top             =   4560
         Width           =   2115
      End
      Begin VB.CheckBox chkModificarcorte 
         Appearance      =   0  'Flat
         Caption         =   "Modificar Cierre de Caja"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   51
         Top             =   4920
         Width           =   2025
      End
      Begin VB.CheckBox chkCotizar 
         Appearance      =   0  'Flat
         Caption         =   "Cotizar Empeño"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   1680
         Width           =   1695
      End
      Begin VB.CheckBox chkFacturacion 
         Appearance      =   0  'Flat
         Caption         =   "Facturación"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9960
         TabIndex        =   49
         Top             =   3120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CheckBox chkMoviDiv 
         Appearance      =   0  'Flat
         Caption         =   "Divisas Interbancarias"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   48
         Top             =   8160
         Width           =   2025
      End
      Begin VB.CheckBox chkAbono 
         Appearance      =   0  'Flat
         Caption         =   "Modificar abono ventas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   47
         Top             =   5280
         Width           =   2025
      End
      Begin VB.CheckBox chkCotizacion 
         Appearance      =   0  'Flat
         Caption         =   "Cotización Diaria"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   46
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CheckBox chkComvendiv 
         Appearance      =   0  'Flat
         Caption         =   "Compra/Venta Divisas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   45
         Top             =   6000
         Width           =   1935
      End
      Begin VB.CheckBox chkRepdivisas 
         Appearance      =   0  'Flat
         Caption         =   "Reportes Divisas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   44
         Top             =   6360
         Width           =   1935
      End
      Begin VB.CheckBox chkRepgastos 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Gastos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   43
         Top             =   1320
         Width           =   1935
      End
      Begin VB.CheckBox chkCancelarbol 
         Appearance      =   0  'Flat
         Caption         =   "Cancelacion de Movimientos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   42
         Top             =   600
         Width           =   2445
      End
      Begin VB.CheckBox chkUsuarios 
         Appearance      =   0  'Flat
         Caption         =   "Configuración Usuarios"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   39
         Top             =   2400
         Width           =   2130
      End
      Begin VB.CheckBox chkCapboletas 
         Appearance      =   0  'Flat
         Caption         =   "Captura de Contratos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   38
         Top             =   3120
         Width           =   2055
      End
      Begin VB.CheckBox chkParametros 
         Appearance      =   0  'Flat
         Caption         =   "Configuración Parámetros"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   37
         Top             =   960
         Width           =   2265
      End
      Begin VB.CheckBox chkGastos 
         Appearance      =   0  'Flat
         Caption         =   "Movimientos Gastos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   36
         Top             =   7800
         Width           =   1935
      End
      Begin VB.CheckBox chkRemates 
         Appearance      =   0  'Flat
         Caption         =   "Pasar a Almoneda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8040
         TabIndex        =   35
         Top             =   240
         Width           =   1920
      End
      Begin VB.CheckBox chkMovibancos 
         Appearance      =   0  'Flat
         Caption         =   "Movimiento Bancos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   34
         Top             =   7080
         Width           =   2085
      End
      Begin VB.CheckBox chkMovicaja 
         Appearance      =   0  'Flat
         Caption         =   "Movimientos Caja General"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   33
         Top             =   6720
         Width           =   2295
      End
      Begin VB.CheckBox chkRepEmpeños 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Empeños"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   32
         Top             =   7440
         Width           =   1575
      End
      Begin VB.CheckBox chkRephistorico 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Histórico"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   31
         Top             =   6720
         Width           =   1575
      End
      Begin VB.CheckBox chkRepvencidos 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Vencidos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   3120
         Width           =   1935
      End
      Begin VB.CheckBox chkRepinventario 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Depositaría"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   29
         Top             =   7080
         Width           =   1935
      End
      Begin VB.CheckBox chkRepventas 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Ventas Mostrador"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   7080
         Width           =   2235
      End
      Begin VB.CheckBox chkRepauxiliar 
         Appearance      =   0  'Flat
         Caption         =   "Reportes Auxiliares"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   27
         Top             =   240
         Width           =   1935
      End
      Begin VB.CheckBox chkRepauditoria 
         Appearance      =   0  'Flat
         Caption         =   "Reporte de Auditoría"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   26
         Top             =   960
         Width           =   1935
      End
      Begin VB.CheckBox chkRepcontable 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Contable"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5280
         TabIndex        =   25
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkEtiquetas 
         Appearance      =   0  'Flat
         Caption         =   "Impresión Etiquetas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   24
         Top             =   2400
         Width           =   1920
      End
      Begin VB.CheckBox chkExistencias 
         Appearance      =   0  'Flat
         Caption         =   "Existencias"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   23
         Top             =   600
         Width           =   2055
      End
      Begin VB.CheckBox chkInventario 
         Appearance      =   0  'Flat
         Caption         =   "Inventario Físico"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   22
         Top             =   1680
         Width           =   1575
      End
      Begin VB.CheckBox chkDotacion 
         Appearance      =   0  'Flat
         Caption         =   "Compras/Dotaciones Inventario"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   960
         Width           =   2580
      End
      Begin VB.CheckBox chkGrupos 
         Appearance      =   0  'Flat
         Caption         =   "Grupos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   20
         Top             =   9240
         Width           =   855
      End
      Begin VB.CheckBox chkCierresucursal 
         Appearance      =   0  'Flat
         Caption         =   "Cierre de Sucursal"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   5640
         Width           =   1815
      End
      Begin VB.CheckBox chkRepfinanciero 
         Appearance      =   0  'Flat
         Caption         =   "Reporte Activos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   5280
         Width           =   1815
      End
      Begin VB.CheckBox chkBalance 
         Appearance      =   0  'Flat
         Caption         =   "Balance"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   4920
         Width           =   1815
      End
      Begin VB.CheckBox chkCortecaja 
         Appearance      =   0  'Flat
         Caption         =   "Cierre de Caja"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   3840
         Width           =   1815
      End
      Begin VB.CheckBox chkConceptos 
         Appearance      =   0  'Flat
         Caption         =   "Catálogo de conceptos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10680
         TabIndex        =   15
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CheckBox chkBusqueda 
         Appearance      =   0  'Flat
         Caption         =   "Búsqueda de Contratos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2400
         Width           =   2115
      End
      Begin VB.CheckBox chkVentas 
         Appearance      =   0  'Flat
         Caption         =   "Ventas Mostrador"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   6000
         Width           =   1695
      End
      Begin VB.CheckBox chkRefrendos 
         Appearance      =   0  'Flat
         Caption         =   "Refrendos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   960
         Width           =   1095
      End
      Begin VB.CheckBox chkDesempeños 
         Appearance      =   0  'Flat
         Caption         =   "Desempeños"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   1455
      End
      Begin VB.CheckBox chkAutos 
         Appearance      =   0  'Flat
         Caption         =   "Empeño Automóviles"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5040
         TabIndex        =   10
         Top             =   9000
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.CheckBox chkEmpeños 
         Appearance      =   0  'Flat
         Caption         =   "Empeños"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   11340
      TabIndex        =   40
      Top             =   8670
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Limpiar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   255
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmAltaUsuarios.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosusuario 
      Height          =   240
      Left            =   2340
      TabIndex        =   41
      Top             =   945
      Width           =   330
      _ExtentX        =   582
      _ExtentY        =   423
      AlignCaption    =   4
      AutoSize        =   0   'False
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
   End
   Begin DevPowerFlatBttn.FlatBttn cmdTodos 
      Height          =   375
      Left            =   10260
      TabIndex        =   53
      Top             =   8670
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Todos"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmAltaUsuarios.frx":0110
      PictureDisabled =   "frmAltaUsuarios.frx":047A
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   12420
      TabIndex        =   103
      Top             =   8670
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmAltaUsuarios.frx":05D4
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   8160
      TabIndex        =   104
      Top             =   8670
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Aceptar"
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
      Picture         =   "frmAltaUsuarios.frx":0B26
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   375
      Left            =   9210
      TabIndex        =   105
      Top             =   8670
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Eliminar"
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
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmAltaUsuarios.frx":1078
      PictureDisabled =   "frmAltaUsuarios.frx":15CA
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   810
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Contraseña:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   7
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Verificar contraseña:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   2070
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   795
   End
End
Attribute VB_Name = "frmAltaUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim ctrl As Control
Dim Ban As Boolean
Dim Empeno As Integer, empeñoautos As Integer, desempeños As Integer, refrendos As Integer, ubicacion As Integer, Ventas As Integer, busqueda As Integer, conceptos As Integer, cortecaja As Integer, balance As Integer, repfinanciero As Integer, cierresucursal As Integer, grupos As Integer, dotacion As Integer, devolucion As Integer, inventariofisico As Integer, Existencias As Integer, etiquetas As Integer, exporinformacion As Integer, repcontable As Integer, repauditoria As Integer, repauxiliar As Integer, repventas As Integer, repinventarios As Integer, repvencidos As Integer, rephistorico As Integer, repempeños As Integer, repasistencia As Integer, movimientocaja As Integer, movimientobanco As Integer, transferencias As Integer, remates As Integer, gastos As Integer, parametros As Integer, capboletas As Integer, usuarios As Integer, cancelbol As Integer, repgastos As Integer
Dim catdivisas As Integer, Cotizacion As Integer, comvendiv As Integer, repdivisas As Integer, Movidiv As Integer, facturacion As Integer, cotizarempeño As Integer, reporteremates As Integer, Abono As Integer, Precio As Integer, modificarcorte As Integer, HacerCorte As Integer, InteresRefrendo As Integer, InteresDesempeño As Integer

Private Sub cmdAceptar_Click()
Dim rcVerifica As New ADODB.Recordset
Dim Sql As String

On Error GoTo Error

    If Valida Then
        
        If txtUsuario.Tag = "" Then
            
            rcVerifica.Open "SELECT ID FROM usuarios WHERE Usuario='" & Trim(txtUsuario.text) & "' AND Estatus=1", dbDatos, adOpenForwardOnly, adLockReadOnly
            If Not rcVerifica.BOF And Not rcVerifica.EOF Then
                
                MsgBox "El Nombre de usuario que desea registrar ya existe !!" & Chr(10) & "Seleccione otro Nombre de Usuario.", vbInformation, "Usuarios"
                txtUsuario.SetFocus
            Else
                
                Permisos
                Sql = "INSERT INTO usuarios(Nombre,Usuario,Contraseña,Empeño,empeñoautos,desempeños,refrendos,ubicacion,ventas,busqueda,conceptos,cortecaja,balance,repfinanciero,cierresucursal,grupos,dotacion,devolucion,inventariofisico,existencias,etiquetas,exporinformacion,repcontable,repauditoria,repauxiliar,repventas,repinventarios,repvencidos,rephistorico,repempeños,repasistencia,movimientocaja,movimientobanco,transferencias,remates,gastos,parametros,capboletas,usuarios,repgastos,catdivisas,cotizacion,comvendiv,repdivisas,movidiv,cancelbol,facturacion,cotizarempeño,reporteremates,abonar,precio,modificarcorte,hacercorte,interesrefrendo,interesdesempeño,AnaliClientes,RegUbicacion,RepAlmoneda,RepCierres,RepIngresos,CancelVenta,CambioVenta,PagoDemasia,RepApartado,RepUtilidad,EntradaInven,SalidaInven,Deslotifica,TrasInven,ListaPrecio,RepCompras,RepTras,Kardex,RepAnti,RepEnve,RepEnveP,RepSalida,RepCierreSucursal,RepAutorizaciones,RepPrendasSimi,RepPrendasAudi,Traspasos,RepAleatoria, " _
                   & "Sucursales,CatTipos,CatFamilias,CatSubFamilias,CatMedios,CatCuentasGas,CargosAbonos,CatClientes,MoviBoveda,MostrarApartados,ApartadosVencidos,PrecioVitrina,TipoPrenda,PreciosKilataje,TarjetaBeneficio,DescuentoVentas,RecalculoPrecios,PrestamoBoleta1,PagosFijos,CambioPlan,CierreDivisas,RepCartera,VenCliente,EtiInven,RepDota,RepDesempenos,RepRefrendos,RepHorarios,RepPartidaBoveda,RepAseguradora,RepCancelaciones,RepEmpeProm,RepDesemProm,RepRefProm,ConTipoTasa,ConVencidos,ConStatus,PrestamoMes,Medios,ConfiguraTasas,ConfiguraDiam,Catalogos,MensajeContratos,ConexionSuc,GeneraAutoriza,CatElec,RefrendarVencidos,CancelaCierre,mld_parametros,mld_movatipicos,mld_reppormenorizado,mld_expclientes) VALUES " _
                   & "('" & Trim(txtNombre.text) & "','" & txtUsuario.text & "','" & txtPass.text & "'," & Empeno & "," & empeñoautos & "," & desempeños & "," & refrendos & "," & ubicacion & "," & Ventas & "," & busqueda & "," & conceptos & "," & cortecaja & "," & balance & "," & repfinanciero & "," & cierresucursal & "," & grupos & "," & dotacion & "," & devolucion & "," & inventariofisico & "," & Existencias & "," & etiquetas & "," & exporinformacion & "," & repcontable & "," & repauditoria & "," & repauxiliar & "," & repventas & "," & repinventarios & "," & repvencidos & "," & rephistorico & "," & repempeños & "," & repasistencia & "," & movimientocaja & "," & movimientobanco & "," & transferencias & "," & remates & "," & gastos & "," & parametros & "," & capboletas & "," & usuarios & "," & repgastos & "," & chkCatDivisas.Value & "," & Cotizacion & "," & comvendiv & "," & repdivisas & "," & Movidiv & "," & cancelbol & "," & facturacion & "," & cotizarempeño & "," _
                   & reporteremates & "," & Abono & "," & Precio & "," & modificarcorte & "," & HacerCorte & "," & InteresRefrendo & "," & InteresDesempeño & "," & chkAnaliClientes.Value & "," & chkRegUbicacion.Value & "," & chkRepAlmoneda.Value & "," & chkRepCierres.Value & "," & chkRepIngresos.Value & "," & chkCancelVenta.Value & "," & chkCambioVenta.Value & "," & chkPagoDemasia.Value & "," & chkRepApartado.Value & "," & chkRepUtilidadVen.Value & "," & chkEntradaInven.Value & "," & chkSalidaInven.Value & "," & chkDeslotifica.Value & "," & chkTrasInven.Value & "," & chkListaPrecio.Value & "," & chkRepCompras.Value & "," & chkRepTras.Value & "," & chkKardex.Value & "," & chkRepAnti.Value & "," & chkRepEnve.Value & "," & chkRepEnveP.Value & "," & chkRepSalida.Value & "," & chkRepCierreSucursal.Value & "," & chkRepAutorizaciones.Value & "," & chkRepIngresos.Value & "," & chkPrendasSimilares.Value & "," & chkTraspasos.Value & "," & chkAleatoriaSelectiva.Value & "," _
                   & chkSucursales.Value & "," & chkCatTipos.Value & "," & chkCatFamilias.Value & "," & chkCatSubFamilias.Value & "," & chkCatMedios.Value & "," & chkCatCuentas.Value & "," & chkCargosAbonos.Value & "," & chkCatClientes.Value & "," & chkMoviBoveda.Value & "," & chkMostrarApartados.Value & "," & chkApartadosVencidos.Value & "," & chkPrecioVitrina.Value & "," & chkTipoPrenda.Value & "," & chkPreciosKilataje.Value & "," & chkTarjetaBeneficio.Value & "," & chkDescuentoVentas.Value & "," & chkRecalculo.Value & "," & chkPrestamoBoleta1.Value & "," & chkPagosFijos.Value & "," & chkCambioPlan.Value & "," & chkCierreDivisas.Value & "," & chkRepCartera.Value & "," & chkVenCliente.Value & "," & chkEtiInven.Value & "," & chkRepDota.Value & "," _
                   & chkRepDesempenos.Value & "," & chkRepRefrendos.Value & "," & chkRepHorarios.Value & "," & chkRepPartidaBoveda.Value & "," & chkRepAseguradora.Value & "," & chkRepCancelaciones.Value & "," & chkRepEmpeProm.Value & "," & chkRepDesemProm.Value & "," & chkRepRefProm.Value & "," & chkConTipoTasa.Value & "," & chkConVencidos.Value & "," & chkConStatus.Value & "," & chkPrestamoMes.Value & "," & chkMedios.Value & "," & chkConfiguraTasas.Value & "," & chkConfiguraDiam.Value & "," & chkCatalogos.Value & "," & chkMensajeContratos.Value & "," & chkConexionSuc.Value & "," & chkGeneraAutoriza.Value & "," & chkCatElec.Value & "," & chkRefrendarVencidos.Value & "," & chkCancelarCierre.Value & _
                   "," & chkMldParametros.Value & "," & chkMldMovAtipicos.Value & "," & chkMldRepPormenorizado.Value & "," & chkMldExpClientes.Value & ")"
                
                dbDatos.Execute Sql
            
                Deselecciona
                txtNombre.text = ""
                txtUsuario.text = ""
                txtPass.text = ""
                txtPass1.text = ""
                txtUsuario.Tag = ""
                txtNombre.SetFocus
            End If
            rcVerifica.Close
            Set rcVerifica = Nothing
            
        Else

            If MsgBox("Desea guardar los cambios realizados ??", vbQuestion + vbYesNo + vbDefaultButton1, "Usuarios") = vbYes Then
                
                Permisos
                
                Sql = "UPDATE usuarios SET Nombre='" & txtNombre.text & "',Usuario='" & txtUsuario.text & "',Contraseña='" & txtPass.text & "',empeño=" & Empeno & ",empeñoautos=" & empeñoautos & ",desempeños=" & desempeños & ",refrendos=" & refrendos & ",ubicacion=" & ubicacion & ",ventas=" & Ventas & ",busqueda=" & busqueda & ",conceptos=" & conceptos & ",cortecaja=" & cortecaja & ",balance=" & balance & ",repfinanciero=" & repfinanciero & ",cierresucursal=" & cierresucursal & ",grupos=" & grupos & ",dotacion=" & dotacion & ",devolucion=" & devolucion & ",inventariofisico=" & inventariofisico & ",existencias=" & Existencias & ",etiquetas=" & etiquetas & ",exporinformacion=" & exporinformacion & ",repcontable=" & repcontable & ",repauditoria=" & repauditoria & ",repauxiliar=" & repauxiliar & ",repventas=" & repventas & ",repinventarios=" & repinventarios & "," _
                   & "repvencidos=" & repvencidos & ",rephistorico=" & rephistorico & ",repempeños=" & repempeños & ",repasistencia=" & repasistencia & ",movimientocaja=" & movimientocaja & ",movimientobanco=" & movimientobanco & ",transferencias=" & transferencias & ",remates=" & remates & ",gastos=" & gastos & ",parametros=" & parametros & ",capboletas=" & capboletas & ",usuarios=" & usuarios & ",cancelbol=" & cancelbol & ",repgastos=" & repgastos & ",catdivisas=" & catdivisas & ",cotizacion=" & Cotizacion & ",comvendiv=" & comvendiv & ",repdivisas=" & repdivisas & ",movidiv=" & Movidiv & ",facturacion=" & facturacion & ",cotizarempeño=" & cotizarempeño & ",reporteremates=" & reporteremates & ",abonar=" & Abono & ",precio=" & Precio & ",modificarcorte=" & modificarcorte & ",hacercorte=" & HacerCorte & ",interesrefrendo=" & InteresRefrendo & ",interesdesempeño=" & InteresDesempeño & "," _
                   & "AnaliClientes=" & chkAnaliClientes.Value & ",RegUbicacion=" & chkRegUbicacion.Value & ",repalmoneda=" & chkRepAlmoneda.Value & ",repcierres=" & chkRepCierres.Value & ",repingresos=" & chkRepIngresos.Value & ",CancelVenta=" & chkCancelVenta.Value & ",CambioVenta=" & chkCambioVenta.Value & ",PagoDemasia=" & chkPagoDemasia.Value & ",RepApartado=" & chkRepApartado.Value & ",RepUtilidad=" & chkRepUtilidadVen.Value & ",EntradaInven=" & chkEntradaInven.Value & ",SalidaInven=" & chkSalidaInven.Value & ",Deslotifica=" & chkDeslotifica.Value & ",TrasInven=" & chkTrasInven.Value & ",ListaPrecio=" & chkListaPrecio.Value & ",RepCompras=" & chkRepCompras.Value & ",RepTras=" & chkRepTras.Value & ",Kardex=" & chkKardex.Value & ",RepAnti=" & chkRepAnti.Value & ",RepEnve=" & chkRepEnve.Value & ",RepEnveP=" & chkRepEnveP.Value & ",RepSalida=" & chkRepSalida.Value & "," _
                   & "RepCierreSucursal=" & chkRepCierreSucursal.Value & ",RepAutorizaciones=" & chkRepAutorizaciones.Value & ",RepPrendasSimi=" & chkPrendasSimilares.Value & ",RepPrendasAudi=" & chkPrendasAuditadas.Value & ",Traspasos=" & chkTraspasos.Value & ",RepIngresos=" & chkRepIngresos.Value & ",RepAleatoria=" & chkAleatoriaSelectiva.Value & ",Sucursales=" & chkSucursales.Value & ",CatTipos=" & chkCatTipos.Value & ",CatFamilias=" & chkCatFamilias.Value & ",CatSubFamilias=" & chkCatSubFamilias.Value & ",CatMedios=" & chkCatMedios.Value & ",CatDivisas=" & chkCatDivisas.Value & ",CatCuentasGas=" & chkCatCuentas.Value & ",CargosAbonos=" & chkCargosAbonos.Value & ",CatClientes=" & chkCatClientes.Value & ",MoviBoveda=" & chkMoviBoveda.Value & ",MostrarApartados=" & chkMostrarApartados.Value & ",ApartadosVencidos=" & chkApartadosVencidos.Value & "," _
                   & "PrecioVitrina=" & chkPrecioVitrina.Value & ",tipoprenda=" & chkTipoPrenda.Value & ",PreciosKilataje=" & chkPreciosKilataje.Value & ",TarjetaBeneficio=" & chkTarjetaBeneficio.Value & ",DescuentoVentas=" & chkDescuentoVentas.Value & ",RecalculoPrecios=" & chkRecalculo.Value & ",PrestamoBoleta1=" & chkPrestamoBoleta1.Value & ",PagosFijos=" & chkPagosFijos.Value & ",CambioPlan=" & chkCambioPlan.Value & ",CierreDivisas=" & chkCierreDivisas.Value & ",RepCartera=" & chkRepCartera.Value & ",VenCliente=" & chkVenCliente.Value & ",EtiInven=" & chkEtiInven.Value & ",RepDota=" & chkRepDota.Value & "," _
                   & "RepDesempenos=" & chkRepDesempenos.Value & ",RepRefrendos=" & chkRepRefrendos.Value & ",RepHorarios=" & chkRepHorarios.Value & ",RepPartidaBoveda=" & chkRepPartidaBoveda.Value & ",RepAseguradora=" & chkRepAseguradora.Value & ",RepCancelaciones=" & chkRepCancelaciones.Value & ",RepEmpeProm=" & chkRepEmpeProm.Value & ",RepDesemProm=" & chkRepDesemProm.Value & ",RepRefProm=" & chkRepRefProm.Value & ",ConTipoTasa=" & chkConTipoTasa.Value & ",ConVencidos=" & chkConVencidos.Value & ",ConStatus=" & chkConStatus.Value & ",PrestamoMes=" & chkPrestamoMes.Value & ",Medios=" & chkMedios.Value & ",ConfiguraTasas=" & chkConfiguraTasas.Value & ",ConfiguraDiam=" & chkConfiguraDiam.Value & ",Catalogos=" & chkCatalogos.Value & ",MensajeContratos=" & chkMensajeContratos.Value & ",ConexionSuc=" & chkConexionSuc.Value & ",Movidiv=" & chkMoviDiv.Value & ",GeneraAutoriza=" & chkGeneraAutoriza.Value & ",CatElec=" & chkCatElec.Value & "," _
                   & "RefrendarVencidos = " & chkRefrendarVencidos.Value & ",CancelaCierre=" & chkCancelarCierre.Value & _
                    ",mld_parametros=" & chkMldParametros.Value & " , mld_movatipicos=" & chkMldMovAtipicos.Value & " ,mld_expclientes=" & chkMldExpClientes.Value & " ,mld_reppormenorizado=" & chkMldRepPormenorizado.Value & _
                   " WHERE ID=" & Val(txtUsuario.Tag)
                dbDatos.Execute Sql
            
                Deselecciona
                txtUsuario.text = ""
                txtUsuario.Tag = ""
                txtNombre.text = ""
                txtPass.text = ""
                txtPass1.text = ""
                txtNombre.SetFocus
            End If
            
        End If
    
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcVerifica = Nothing
End Sub

Private Sub cmdEliminar_Click()
    
    If txtUsuario.Tag <> "" Then
    
        If MsgBox("Desea eliminar el usuario seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Usuarios") = vbYes Then
            
            dbDatos.Execute "UPDATE usuarios SET Estatus=0 WHERE ID=" & Val(txtUsuario.Tag)
            Deselecciona
            txtUsuario.Tag = ""
            txtUsuario.text = ""
            txtNombre.text = ""
            txtPass.text = ""
            txtPass1.text = ""
            txtNombre.SetFocus
        End If
    
    Else
        
        txtNombre.SetFocus
    End If
    
End Sub

Private Sub cmdLimpiar_Click()
    Deselecciona
    txtUsuario.Tag = ""
    txtNombre.text = ""
    txtUsuario.text = ""
    txtPass.text = ""
    txtPass1.text = ""
    txtNombre.SetFocus
End Sub

Private Sub cmdMosusuario_Click()
    frmMostrarUsuarios.Ver Me, txtUsuario, False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Sub Selecciona()

    For Each ctrl In Me.Controls
        
        If TypeOf ctrl Is CheckBox Then ctrl.Value = 1
    Next
End Sub

Sub Deselecciona()

    For Each ctrl In Me.Controls
        
        If TypeOf ctrl Is CheckBox Then ctrl.Value = 0
    Next
    
End Sub

Private Sub cmdTodos_Click()
    
    If Ban = False Then
        
        Ban = True
        Selecciona
    Else
    
        Ban = False
        Deselecciona
    End If
End Sub

Private Sub Form_Load()
    Ban = False
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtUsuario_GotFocus()
    Cambiar_Color True, txtUsuario
    Seleccionar_Texto txtUsuario
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtUsuario_LostFocus()
Cambiar_Color False, txtUsuario
End Sub

Private Sub txtPass_GotFocus()
    Cambiar_Color True, txtPass
    Seleccionar_Texto txtPass
End Sub

Private Sub txtPass_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPass_LostFocus()
    Cambiar_Color False, txtPass
End Sub

Private Sub txtPass1_GotFocus()
    Cambiar_Color True, txtPass1
    Seleccionar_Texto txtPass1
End Sub

Private Sub txtPass1_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPass1_LostFocus()
    If txtPass.text <> txtPass1.text Then
        
        MsgBox "Verifique su contraseña !!", vbCritical, "Usuarios"
        txtPass.SetFocus
    Else
        
        Cambiar_Color False, txtPass1
    End If
End Sub

Private Sub txtNombre_GotFocus()
    Cambiar_Color True, txtNombre
    Seleccionar_Texto txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNombre_LostFocus()
    Cambiar_Color False, txtNombre
End Sub

Function Valida() As Boolean

    Valida = True

    If txtNombre.text = "" Then
        MsgBox "Introduzca el nombre del usuario que desea registrar !!", vbInformation, "Usuarios"
        Valida = False
        txtNombre.SetFocus
        Exit Function
    End If

    If txtUsuario.text = "" Then
        MsgBox "Introduzca el nombre de usuario !!", vbInformation, "Usuarios"
        Valida = False
        txtUsuario.SetFocus
        Exit Function
    End If

    If txtPass.text = "" Then
        MsgBox "Introduzca el password !!", vbInformation, "Usuarios"
        Valida = False
        txtPass.SetFocus
        Exit Function
    End If

    If txtPass1.text = "" Then
        MsgBox "Introduzca la verificación del password !!", vbInformation, "Usuarios"
        Valida = False
        txtPass1.SetFocus
        Exit Function
    End If

    If chkEmpeños.Value = 0 And chkAutos.Value = 0 And chkDesempeños.Value = 0 And chkRefrendos.Value = 0 And chkVentas.Value = 0 And chkBusqueda.Value = 0 And chkConceptos.Value = 0 And chkCortecaja.Value = 0 And chkBalance.Value = 0 And chkRepfinanciero.Value = 0 And chkCierresucursal.Value = 0 And chkGrupos.Value = 0 And chkDotacion.Value = 0 And chkInventario.Value = 0 And chkExistencias.Value = 0 And chkEtiquetas.Value = 0 And chkRepcontable.Value = 0 And chkRepauditoria.Value = 0 And chkRepauxiliar.Value = 0 And chkVentas.Value = 0 And chkInventario.Value = 0 And chkRepvencidos.Value = 0 And chkRephistorico.Value = 0 And chkRepEmpeños.Value = 0 And chkMovicaja.Value = 0 And chkMovibancos.Value = 0 And chkRemates.Value = 0 And chkGastos.Value = 0 And chkParametros.Value = 0 And chkCapboletas.Value = 0 And chkRepgastos.Value = 0 And chkFacturacion.Value = 0 And chkCotizar.Value = 0 Then
        MsgBox "Seleccione por lo menos un permiso para el usuario !!", vbCritical, "Usuarios"
        Valida = False
        Exit Function
    End If

End Function

Public Sub Buscar(IDUsuario As Integer)
Dim rcPermisos As New ADODB.Recordset

'On Error GoTo Error

    rcPermisos.Open "SELECT * FROM usuarios WHERE ID=" & IDUsuario, dbDatos, adOpenForwardOnly, adLockReadOnly
    If Not rcPermisos.BOF And Not rcPermisos.EOF Then

        With rcPermisos
            
            Deselecciona
        
            txtNombre.text = !Nombre
            txtUsuario.text = !Usuario
            txtUsuario.Tag = !ID
            txtPass.text = !contraseña
            txtPass1.text = !contraseña
        
            chkEmpeños.Value = !empeño
            chkAutos.Value = !empeñoautos
            chkDesempeños.Value = !desempeños
            chkRefrendos.Value = !refrendos
            chkVentas.Value = !Ventas
            chkBusqueda.Value = !busqueda
            chkConceptos.Value = !conceptos
            chkCortecaja.Value = !cortecaja
            chkBalance.Value = !balance
            chkRepfinanciero.Value = !repfinanciero
            chkCierresucursal.Value = !cierresucursal
            chkGrupos.Value = !grupos
            chkDotacion.Value = !dotacion
            chkInventario.Value = !inventariofisico
            chkExistencias.Value = !Existencias
            chkEtiquetas.Value = !etiquetas
            chkRepcontable.Value = !repcontable
            chkRepauditoria.Value = !repauditoria
            chkRepauxiliar.Value = !repauxiliar
            chkRepventas.Value = !repventas
            chkRepinventario.Value = !repinventarios
            chkRepvencidos.Value = !repvencidos
            chkRephistorico.Value = !rephistorico
            chkRepEmpeños.Value = !repempeños
            chkMovicaja.Value = !movimientocaja
            chkMovibancos.Value = !movimientobanco
            chkRemates.Value = !remates
            chkGastos.Value = !gastos
            chkParametros.Value = !parametros
            chkCapboletas.Value = !capboletas
            chkUsuarios.Value = !usuarios
            chkCancelarbol.Value = !cancelbol
            chkRepgastos.Value = !repgastos
            chkCotizacion.Value = !Cotizacion
            chkComvendiv.Value = !comvendiv
            chkRepdivisas.Value = !repdivisas
            chkMoviDiv = !Movidiv
            chkFacturacion.Value = !facturacion
            chkCotizar.Value = !cotizarempeño
            chkAbono.Value = !abonar
            chkModificarcorte.Value = !modificarcorte
            chkHacercorte.Value = !HacerCorte
            chkInteresrefrendo.Value = !InteresRefrendo
            chkInteresdesempeño.Value = !InteresDesempeño
            chkAnaliClientes.Value = !analiclientes
            chkRegUbicacion.Value = !regubicacion
            chkRepAlmoneda.Value = !repalmoneda
            chkRepCierres.Value = !repcierres
            chkRepIngresos.Value = !RepIngresos
        
            chkCancelVenta.Value = !CancelVenta
            chkCambioVenta.Value = !CambioVenta
            chkPagoDemasia.Value = !PagoDemasia
            chkRepApartado.Value = !repapartado
            chkRepUtilidadVen.Value = !reputilidad
            chkEntradaInven.Value = !entradainven
            chkSalidaInven.Value = !salidainven
        
            chkDeslotifica.Value = !deslotifica
            chkTrasInven.Value = !trasinven
            chkListaPrecio.Value = !listaprecio
            chkRepCompras.Value = !repcompras
            chkRepTras.Value = !reptras
            chkKardex.Value = !kardex
            chkRepAnti.Value = !repanti
            chkRepEnve.Value = !repenve
            chkRepEnveP.Value = !repenvep
            chkRepSalida.Value = !RepSalida
        
            chkRepCierreSucursal.Value = !RepCierreSucursal
            chkRepAutorizaciones.Value = !RepAutorizaciones
            chkPrendasSimilares.Value = !RepPrendasSimi
            chkPrendasAuditadas.Value = !RepPrendasAudi
            chkTraspasos.Value = !Traspasos
            chkAleatoriaSelectiva.Value = !RepAleatoria
        
            chkSucursales.Value = !Sucursales
            chkCatTipos.Value = !CatTipos
            chkCatFamilias.Value = !CatFamilias
            chkCatSubFamilias.Value = !CatSubFamilias
            chkCatMedios.Value = !CatMedios
            chkCatDivisas.Value = !catdivisas
            chkCatCuentas.Value = !CatCuentasGas
            chkCargosAbonos.Value = !CargosAbonos
            chkCatClientes.Value = !CatClientes
            chkMoviBoveda.Value = !moviboveda
            chkMostrarApartados.Value = !MostrarApartados
            chkApartadosVencidos.Value = !ApartadosVencidos
            chkPrecioVitrina.Value = !PrecioVitrina
            chkTipoPrenda.Value = !TipoPrenda
            chkPreciosKilataje.Value = !PreciosKilataje
            chkTarjetaBeneficio.Value = !TarjetaBeneficio
            chkDescuentoVentas.Value = !DescuentoVentas
            chkRecalculo.Value = !RecalculoPrecios
            chkPrestamoBoleta1.Value = !PrestamoBoleta1
            
            chkPagosFijos.Value = !PagosFijos
            chkCambioPlan.Value = !CambioPlan
            chkCierreDivisas.Value = !CierreDivisas
            chkRepCartera.Value = !RepCartera
            
            chkVenCliente.Value = !VenCliente
            chkEtiInven.Value = !etiinven
            chkRepDota.Value = !RepDota
            
            chkRepDesempenos.Value = !RepDesempenos
            chkRepRefrendos.Value = !RepRefrendos
            chkRepHorarios.Value = !RepHorarios
            chkRepPartidaBoveda.Value = !RepPartidaBoveda
            chkRepAseguradora.Value = !RepAseguradora
            chkRepCancelaciones.Value = !RepCancelaciones
            chkRepEmpeProm.Value = !RepEmpeProm
            chkRepDesemProm.Value = !RepDesemProm
            chkRepRefProm.Value = !RepRefProm
            chkConTipoTasa.Value = !ConTipoTasa
            chkConVencidos.Value = !ConVencidos
            chkConStatus.Value = !ConStatus
            chkPrestamoMes.Value = !PrestamoMes
            chkMedios.Value = !Medios
            
            chkMoviDiv.Value = !Movidiv
            chkConfiguraTasas.Value = !ConfiguraTasas
            chkConfiguraDiam.Value = !ConfiguraDiam
            chkCatalogos.Value = !Catalogos
            chkMensajeContratos.Value = !MensajeContratos
            chkConexionSuc.Value = !conexionsuc
            chkGeneraAutoriza.Value = !GeneraAutoriza
            chkCatElec.Value = !CatElec
            chkRefrendarVencidos.Value = !RefrendarVencidos
            chkCancelarCierre.Value = !CancelaCierre
            
            '----- MLD-MODIF. -----
            chkMldParametros.Value = !mld_parametros
            chkMldMovAtipicos.Value = !mld_movatipicos
            chkMldRepPormenorizado.Value = !mld_reppormenorizado
            chkMldExpClientes.Value = !mld_expclientes

            
        End With

    End If
    rcPermisos.Close
    Set rcPermisos = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    rcPermisos.Close
    Set rcPermisos = Nothing
End Sub

Sub Permisos()

    If chkEmpeños.Value = 1 Then Empeno = 1 Else Empeno = 0
    If chkAutos.Value = 1 Then empeñoautos = 1 Else empeñoautos = 0
    If chkDesempeños.Value = 1 Then desempeños = 1 Else desempeños = 0
    If chkRefrendos.Value = 1 Then refrendos = 1 Else refrendos = 0
    If chkVentas.Value = 1 Then Ventas = 1 Else Ventas = 0
    If chkBusqueda.Value = 1 Then busqueda = 1 Else busqueda = 0
    If chkConceptos.Value = 1 Then conceptos = 1 Else conceptos = 0
    If chkCortecaja.Value = 1 Then cortecaja = 1 Else cortecaja = 0
    If chkBalance.Value = 1 Then balance = 1 Else balance = 0
    If chkRepfinanciero.Value = 1 Then repfinanciero = 1 Else repfinanciero = 0
    If chkCierresucursal.Value = 1 Then cierresucursal = 1 Else cierresucursal = 0
    If chkGrupos.Value = 1 Then grupos = 1 Else grupos = 0
    If chkDotacion.Value = 1 Then dotacion = 1 Else dotacion = 0
    If chkInventario.Value = 1 Then inventariofisico = 1 Else inventariofisico = 0
    If chkExistencias.Value = 1 Then Existencias = 1 Else Existencias = 0
    If chkEtiquetas.Value = 1 Then etiquetas = 1 Else etiquetas = 0
    If chkRepcontable.Value = 1 Then repcontable = 1 Else repcontable = 0
    If chkRepauditoria.Value = 1 Then repauditoria = 1 Else repauditoria = 0
    If chkRepauxiliar.Value = 1 Then repauxiliar = 1 Else repauxiliar = 0
    If chkRepventas.Value = 1 Then repventas = 1 Else repventas = 0
    If chkRepinventario.Value = 1 Then repinventarios = 1 Else repinventarios = 0
    If chkRepvencidos.Value = 1 Then repvencidos = 1 Else repvencidos = 0
    If chkRephistorico.Value = 1 Then rephistorico = 1 Else rephistorico = 0
    If chkRepEmpeños.Value = 1 Then repempeños = 1 Else repempeños = 0
    If chkMovicaja.Value = 1 Then movimientocaja = 1 Else movimientocaja = 0
    If chkMovibancos.Value = 1 Then movimientobanco = 1 Else movimientobanco = 0
    If chkRemates.Value = 1 Then remates = 1 Else remates = 0
    If chkGastos.Value = 1 Then gastos = 1 Else gastos = 0
    If chkParametros.Value = 1 Then parametros = 1 Else parametros = 0
    If chkCapboletas.Value = 1 Then capboletas = 1 Else capboletas = 0
    If chkUsuarios.Value = 1 Then usuarios = 1 Else usuarios = 0
    If chkCancelarbol.Value = 1 Then cancelbol = 1 Else cancelbol = 0
    If chkRepgastos.Value = 1 Then repgastos = 1 Else repgastos = 0
    If chkCotizacion.Value = 1 Then Cotizacion = 1 Else Cotizacion = 0
    If chkComvendiv.Value = 1 Then comvendiv = 1 Else comvendiv = 0
    If chkRepdivisas.Value = 1 Then repdivisas = 1 Else repdivisas = 0
    If chkMoviDiv.Value = 1 Then Movidiv = 1 Else Movidiv = 0
    If chkFacturacion.Value = 1 Then facturacion = 1 Else facturacion = 0
    If chkCotizar.Value = 1 Then cotizarempeño = 1 Else cotizarempeño = 0
    If chkAbono.Value = 1 Then Abono = 1 Else Abono = 0
    If chkModificarcorte.Value = 1 Then modificarcorte = 1 Else modificarcorte = 0
    If chkHacercorte.Value = 1 Then HacerCorte = 1 Else HacerCorte = 0
    InteresRefrendo = chkInteresrefrendo.Value
    InteresDesempeño = chkInteresdesempeño.Value
End Sub
