VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmVentaCliente 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Venta Cliente"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVentaCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7935
   ScaleWidth      =   10500
   Begin VB.TextBox txtMoratorios 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0.00"
      Top             =   6075
      Width           =   2460
   End
   Begin VB.TextBox txtPrestamo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "0.00"
      Top             =   5280
      Width           =   2460
   End
   Begin VB.TextBox txtIva 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "0.00"
      Top             =   6480
      Width           =   2460
   End
   Begin VB.TextBox txtIntereses 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "0.00"
      Top             =   5670
      Width           =   2460
   End
   Begin VB.TextBox txtTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   7980
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "0.00"
      Top             =   6885
      Width           =   2460
   End
   Begin vbAcceleratorGrid6.vbalGrid grdArticulos 
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      Top             =   1800
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   5953
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
   End
   Begin VB.TextBox txtNumContrato 
      Alignment       =   1  'Right Justify
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
      Height          =   240
      Left            =   1080
      TabIndex        =   0
      Top             =   240
      Width           =   1215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Top             =   180
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Buscar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HighlightColor  =   16777215
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmVentaCliente.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   9360
      TabIndex        =   21
      Top             =   7470
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmVentaCliente.frx":0391
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   8160
      TabIndex        =   23
      Top             =   7470
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Aceptar"
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
      Picture         =   "frmVentaCliente.frx":08E3
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   6840
      TabIndex        =   25
      Top             =   7470
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Re-Imprimir"
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
      Object.ToolTipText     =   ""
      Picture         =   "frmVentaCliente.frx":0E35
   End
   Begin VB.Label Leyenda 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   120
      TabIndex        =   28
      Top             =   7245
      Width           =   165
   End
   Begin VB.Label Total 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   30
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   720
      Left            =   2280
      TabIndex        =   27
      Top             =   7200
      Width           =   210
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Moratorios:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6540
      TabIndex        =   26
      Top             =   6105
      Width           =   1410
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Préstamo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6675
      TabIndex        =   18
      Top             =   5310
      Width           =   1260
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Total:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7035
      TabIndex        =   16
      Top             =   6930
      Width           =   930
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "I.V.A.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   7200
      TabIndex        =   15
      Top             =   6510
      Width           =   765
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Interés, Almacenaje y Seguro:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4245
      TabIndex        =   14
      Top             =   5700
      Width           =   3705
   End
   Begin VB.Label lblIdentificacion 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      TabIndex        =   12
      Top             =   1305
      Width           =   2055
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Identificación:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   3360
      TabIndex        =   11
      Top             =   1320
      Width           =   1305
   End
   Begin VB.Label lblTelefono 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   10
      Top             =   1305
      Width           =   2055
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Teléfono:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label lblDireccion 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   960
      Width           =   8055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Dirección:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   960
      Width           =   885
   End
   Begin VB.Label lblApellido 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   6
      Top             =   600
      Width           =   3495
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   4680
      TabIndex        =   4
      Top             =   600
      Width           =   885
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   765
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Contrato:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   900
   End
End
Attribute VB_Name = "frmVentaCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim IDEmpenoAnterior As Long
Dim fechaVencimiento As Date
Dim importeAlmacenaje As Double
Dim importeSeguro As Double
Dim importeInteres As Double

Private Sub cmdAceptar_Click()
Dim crImporteTotal As Double, crEfectivo As Double

'    If grdArticulos.Rows > 0 Then
        
        If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Venta Cliente") = vbYes Then
            
            crImporteTotal = CDbl(txtTotal.text)
            crEfectivo = frmEfectivo.RegresaCambio(crImporteTotal, 2)
            If crEfectivo < crImporteTotal Then Exit Sub
            CalculaCambio crEfectivo, crImporteTotal
            
            Grabar_Datos_Venta crEfectivo
            Limpiar True
            txtTotal.text = "0.00"
            
        End If
    
'    End If

End Sub

Private Sub cmdBuscar_Click()
    BuscarContrato
End Sub

Private Sub cmdImprimir_Click()
Dim Folio As Long

    Folio = frmReimpresionrecibos.ReImprimir("ventas", "Folio", " WHERE TipoVenta=" & VENTACLIENTE & " AND Folio=")
    If Folio > 0 Then
        
        Imprimir_Recibo_Venta Folio
    
    ElseIf Folio = 0 Then
        
        MsgBox "No se encontró el folio especificado !!", vbInformation, "Venta Cliente"
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    CrearEncabezados
    Poner_Flat Fl, Me.Controls, Me
    txtTotal.text = "0.00"
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtMoratorios_GotFocus()
    Seleccionar_Texto txtMoratorios
    Cambiar_Color True, txtMoratorios
End Sub

Private Sub txtMoratorios_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMoratorios_LostFocus()
    Cambiar_Color False, txtMoratorios
End Sub

Private Sub txtNumcontrato_GotFocus()
    Seleccionar_Texto txtNumContrato
    Cambiar_Color True, txtNumContrato
    
    If Leyenda.Tag = "1" Then
        
        Leyenda.Tag = ""
        Leyenda.Caption = ""
        Total.Caption = ""
        Total.ForeColor = &HFF0000
    End If
End Sub

Private Sub txtNumcontrato_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumcontrato_LostFocus()
    Cambiar_Color False, txtNumContrato
End Sub

Sub BuscarContrato()
Dim rcBusqueda As New ADODB.Recordset
Dim rcAux As New ADODB.Recordset
Dim Folio As Long

On Error GoTo Error

    If Trim(txtNumContrato.text) = "" Then
        
        MsgBox "Introduzca el número de contrato !!", vbInformation, "Venta Cliente"
        txtNumContrato.SetFocus
    Else
        
        rcBusqueda.Open "SELECT empeno.ID AS IDEmpeno,empeno.IDCliente,empeno.Iva,clientes.Nombre,clientes.Apellido,clientes.Direccion,clientes.Colonia,clientes.Municipio,clientes.Estado,clientes.Tel,clientes.Identificacion " _
                        & "FROM clientes RIGHT JOIN empeno ON clientes.ID=empeno.IDCliente WHERE empeno.Numcontrato=" & Val(txtNumContrato.text) & " AND Destino=" & D_ALMONEDA, dbDatos, adOpenForwardOnly, adLockReadOnly
        If Not rcBusqueda.BOF And Not rcBusqueda.EOF Then
            
            'Se comento para hacer la venta cliente a articulos migrados
'            rcAux.Open "SELECT ID FROM detallesentradainventario WHERE Cantidad>0 AND IDEmpeno=" & rcBusqueda!IDEmpeno, dbDatos, adOpenForwardOnly, adLockReadOnly
'            If Not rcAux.BOF And Not rcAux.EOF Then
             If Not rcBusqueda.BOF And Not rcBusqueda.EOF Then
                
                Limpiar
                With rcBusqueda
                    txtNumContrato.Tag = !IDEmpeno
                    lblNombre.Caption = !Nombre
                    lblNombre.Tag = !IDCliente
                    lblApellido.Caption = !Apellido
                    lblApellido.Tag = !Iva
                    lblDireccion.Caption = !Direccion & " " & !Colonia & " " & !Municipio & " " & !Estado
                    lblTelefono.Caption = !Tel
                    lblIdentificacion.Caption = !Identificacion
                    GeneraInteresesEmpeno !IDEmpeno
                    'Se comento
'                    MuestraArticulos !IDEmpeno
                End With
            Else
                
                GoTo NoEncontrado
            End If
'            rcAux.Close
        
        Else
        
NoEncontrado:
            MsgBox "No se encontró el contrato especificado !!", vbInformation, "Venta Cliente"
            Limpiar
            txtNumContrato.SetFocus
        End If

        rcBusqueda.Close
        Set rcBusqueda = Nothing
'        Set rcAux = Nothing
        Exit Sub
    End If

Error:
    Maneja_Error Err
    Set rcBusqueda = Nothing
    Set rcAux = Nothing
End Sub

Sub Limpiar(Optional Ban As Boolean = False)
Dim ctrl As Control
    
    grdArticulos.Clear

    For Each ctrl In Controls
        
        If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = "": ctrl.Tag = ""
        If Ban Then If TypeOf ctrl Is TextBox Then ctrl.text = "": ctrl.Tag = ""
        On Error Resume Next
        If ctrl.Name <> "Leyenda" Then ctrl.Tag = ""
    Next

End Sub

Sub CrearEncabezados()

    With grdArticulos
        .AddColumn "C1", "Cantidad", ecgHdrTextALignRight, , 70, False, , , , , , CCLSortNumeric
        .AddColumn "C2", "Código", ecgHdrTextALignLeft, , 110, , , , , , , CCLSortString
        .AddColumn "C3", "Artículo", ecgHdrTextALignLeft, , 325, , , , , , , CCLSortString
        .AddColumn "C4", "Kilates", ecgHdrTextALignCentre, , 60, , , , , , , CCLSortString
        .AddColumn "C5", "Préstamo", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C6", "Avalúo", ecgHdrTextALignRight, , 90, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C7", "Intereses", ecgHdrTextALignRight, , 90, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C8", "Almacenaje", ecgHdrTextALignRight, , 90, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C9", "Seguro", ecgHdrTextALignRight, , 90, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C10", "Iva", ecgHdrTextALignRight, , 90, False, , , , FMoneda, , CCLSortNumeric
        .AddColumn "C11", "Moratorios", ecgHdrTextALignRight, , 90, False, , , , FMoneda, , CCLSortNumeric
        
        .AddColumn "C12", "Peso", ecgHdrTextALignRight, , 90, False, , , , , , CCLSortNumeric
    End With

End Sub

Function MuestraArticulos(IDEmpeno As Long)
Dim rcArticulos As New ADODB.Recordset
Dim crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crIva As Double, AvaluoPrenda As Double, crAvaluo As Double

On Error GoTo Error
    
    '''''rcPrenda.Open "SELECT d.Cantidad,d.Peso,(d.Costo*d.Cantidad) AS Prestamo,(d.Precio*d.Cantidad) AS Avaluo,empeno.NumContrato,empeno.Fecha,empeno.ID,empeno.TipoInteres,empeno.Vencimiento,d.costo,d.ID AS IDPrenda " _
                        & "FROM empeno INNER JOIN detallesempeno ON empeno.ID=detallesempeno.IDEmpeno INNER JOIN detallesentradainventario d ON empeno.ID=d.IDEmpeno WHERE d.ID=" & grdArticulos.CellItemData(Indice, 1), dbDatos, adOpenForwardOnly, adLockOptimistic
                        
    rcArticulos.Open "SELECT d.ID AS IDArticulo,d.Costo AS Prestamo,(d.Costo*d.Cantidad) AS Prestamoo,d.Cantidad,d.Codigo,d.Descripcion,d.Kilates,d.Peso,empeno.ID,empeno.Fecha,empeno.Prestamo AS PrestamoContrato,empeno.TipoInteres,empeno.TipoTasa,empeno.Vencimiento,empeno.Folio,empeno.NumContrato,empeno.Prestamo,empeno.Operacion,empeno.Periodo,empeno.VenPeriodo " _
                    & "FROM detallesentradainventario d INNER JOIN empeno ON d.IDEmpeno=empeno.ID WHERE d.Cantidad>0 AND d.IDEmpeno=" & IDEmpeno, dbDatos, adOpenForwardOnly, adLockOptimistic
    
    If Not rcArticulos.BOF And Not rcArticulos.EOF Then
        
        With grdArticulos
            
            While Not rcArticulos.EOF
                .AddRow
                .CellText(.Rows, 1) = rcArticulos!Cantidad
                .CellItemData(.Rows, 1) = rcArticulos!IDArticulo
                .CellTextAlign(.Rows, 1) = DT_CENTER
                .CellText(.Rows, 2) = rcArticulos!Codigo
                .CellText(.Rows, 3) = rcArticulos!Cantidad & " " & rcArticulos!Descripcion
                .CellText(.Rows, 4) = SacaKilates(rcArticulos!Kilates)
                .CellItemData(.Rows, 4) = rcArticulos!Kilates
                .CellTextAlign(.Rows, 4) = DT_CENTER
                .CellText(.Rows, 5) = rcArticulos!Prestamoo
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                .CellText(.Rows, 12) = rcArticulos!Peso
                .CellTextAlign(.Rows, 12) = DT_RIGHT
                
                crAvaluo = SacaValor("detallesempeno", "Avaluo", " WHERE Codigo='" & rcArticulos!Codigo & "'")
                .CellText(.Rows, 6) = crAvaluo
                .CellTextAlign(.Rows, 6) = DT_RIGHT
                                
                'Intereses
                'crIntereses = Redondeo(GeneraIntereses(rcArticulos!Prestamoo, rcArticulos!Cantidad * crAvaluo, rcArticulos!NumContrato, IIf(rcArticulos!TipoTasa = "MENSUAL", rcArticulos!Fecha, DateAdd("D", -1, rcArticulos!Fecha)), rcArticulos!ID, "Tasa", rcArticulos!Vencimiento))
                crIntereses = GeneraIntereses(rcArticulos!id, "Tasa")
                .CellText(.Rows, 7) = crIntereses
                .CellTextAlign(.Rows, 7) = DT_RIGHT
                
                'Almacenaje
                'crAlmacenaje = Redondeo(GeneraIntereses(rcArticulos!Prestamoo, rcArticulos!Cantidad * crAvaluo, rcArticulos!NumContrato, IIf(rcArticulos!TipoTasa = "MENSUAL", rcArticulos!Fecha, DateAdd("D", -1, rcArticulos!Fecha)), rcArticulos!ID, "Almacenaje", rcArticulos!Vencimiento))
                crAlmacenaje = GeneraIntereses(rcArticulos!id, "Almacenaje")
                .CellText(.Rows, 8) = crAlmacenaje
                .CellTextAlign(.Rows, 8) = DT_RIGHT
                
                'Seguro
                'crSeguro = Redondeo(GeneraIntereses(rcArticulos!Prestamoo, rcArticulos!Cantidad * crAvaluo, rcArticulos!NumContrato, IIf(rcArticulos!TipoTasa = "MENSUAL", rcArticulos!Fecha, DateAdd("D", -1, rcArticulos!Fecha)), rcArticulos!ID, "Seguro", rcArticulos!Vencimiento))
                crSeguro = GeneraIntereses(rcArticulos!id, "Seguro")
                .CellText(.Rows, 9) = crSeguro
                .CellTextAlign(.Rows, 9) = DT_RIGHT
                                           
                'Moratorios
                crMoratorios = Redondeo(GeneraMoratorios(rcArticulos!Prestamoo, (rcArticulos!Operacion / 100), rcArticulos!Vencimiento, rcArticulos!Serie, 0))
                .CellText(.Rows, 11) = crMoratorios
                .CellTextAlign(.Rows, 11) = DT_RIGHT
                
                'IVA
                crIva = Redondeo(Regresa_Iva(crIntereses + crAlmacenaje + crSeguro + crMoratorios, IDEmpeno))
                .CellText(.Rows, 10) = crIva
                .CellTextAlign(.Rows, 10) = DT_RIGHT
            
            rcArticulos.MoveNext
            Wend
            
            'Pongo el total
            PonerTotales
            
        End With

    End If

    rcArticulos.Close
    Set rcArticulos = Nothing
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcArticulos = Nothing
End Function

Private Sub Grabar_Datos_Venta(crEfectivo As Double)
Dim crTotal As Double, crImporte As Double, crInteres As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crImporteIva As Double
Dim IDVenta As Long, Movimiento As Long, Folio As Long, Indice As Integer, IDCliente As Long, Hora As String

On Error GoTo Error
    
    crImporte = 0

'    For Indice = 1 To grdArticulos.Rows
'        crImporte = crImporte + CDbl(grdArticulos.CellText(Indice, 5))
'        crInteres = crInteres + CDbl(grdArticulos.CellText(Indice, 7))
'        crAlmacenaje = crAlmacenaje + CDbl(grdArticulos.CellText(Indice, 8))
'        crSeguro = crSeguro + CDbl(grdArticulos.CellText(Indice, 9))
'        crMoratorios = crMoratorios + CDbl(grdArticulos.CellText(Indice, 11))
'        crImporteIva = crImporteIva + CDbl(grdArticulos.CellText(Indice, 10))
'    Next Indice
    
      crImporte = CDbl(txtPrestamo.text)
      crInteres = CDbl(importeInteres)
      crMoratorios = 0
      crImporteIva = CDbl(txtIva.text)
     crSeguro = importeSeguro
     crAlmacenaje = importeAlmacenaje
     
     
    
    
    'Saco el Total
    crTotal = CDbl(txtTotal.text)
    
    'Saco el Folio
    Folio = Val(txtNumContrato.text)
    
    
    'Se comento para migracion
    'Grabo la venta
'    dbDatos.Execute "INSERT INTO ventas (Fecha,Folio,IVA,Descuento,Total,PC,IDCliente,IDUsuario,IDSucursal,IDUsuarioDesc,TipoVenta,Efectivo) VALUES ('" & _
'                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & Val(Regresa_Valor_Empeno("IVA", Val(txtNumContrato.Tag))) & ",0," & crTotal & ",'" & NombrePc & "'," & Val(lblNombre.Tag) & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",0," & VENTACLIENTE & "," & crEfectivo & ")"
'

'Se comento para migracion
    'Saco el ID de la Venta
'    IDVenta = SacaValor("ventas", "MAX(ID)")
    
    
    'Se comento para migracion
    'Grabo el detalle de la venta
'    For Indice = 1 To grdArticulos.Rows
'
'        dbDatos.Execute "INSERT INTO detallesventas (IDVenta,Codigo,Articulo,Kilates,Peso,Costo,Precio,IDArticulo,Intereses,Almacenaje,Seguro,Moratorios,ImporteIva) VALUES (" & _
'                        IDVenta & ",'" & grdArticulos.CellText(Indice, 2) & "','" & grdArticulos.CellText(Indice, 3) & "'," & grdArticulos.CellItemData(Indice, 4) & "," & grdArticulos.CellText(Indice, 12) & "," & _
'                        CDbl(grdArticulos.CellText(Indice, 5)) & "," & CDbl(grdArticulos.CellText(Indice, 5)) + CDbl(grdArticulos.CellText(Indice, 7)) + CDbl(grdArticulos.CellText(Indice, 8)) + CDbl(grdArticulos.CellText(Indice, 9)) + CDbl(grdArticulos.CellText(Indice, 11)) & "," & _
'                        grdArticulos.CellItemData(Indice, 1) & "," & CDbl(grdArticulos.CellText(Indice, 7)) & "," & CDbl(grdArticulos.CellText(Indice, 8)) & "," & CDbl(grdArticulos.CellText(Indice, 9)) & "," & CDbl(grdArticulos.CellText(Indice, 11)) & "," & CDbl(grdArticulos.CellText(Indice, 10)) & ")"
'
'        dbDatos.Execute "UPDATE detallesentradainventario SET cantidad=0,TipoSalida=" & SALIDAVENTAPIGNORANTE & " WHERE ID=" & grdArticulos.CellItemData(Indice, 1)
'
'    Next Indice
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Tomo la Hora
    Hora = Time
    
    'Grabo el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','110101'," & crImporte + crInteres + crAlmacenaje + crSeguro + crMoratorios + crImporteIva & "," & TIPO_CARGO & ",1,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabo el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','620201'," & crImporte & "," & TIPO_CARGO & ",1,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabo el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','620450'," & crImporte + crInteres + crAlmacenaje + crSeguro + crMoratorios & "," & TIPO_ABONO & ",1,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

    'Grabo el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','620350'," & crImporte & "," & TIPO_ABONO & ",1,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

    'Grabo el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                                   "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'VT03','120150'," & crImporteIva & "," & TIPO_ABONO & ",1,'Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
      
    If MsgBox("Desa imprimir recibo ??", vbQuestion + vbYesNo, "Venta Cliente") = vbYes Then
    
'        Imprimir_Recibo_Venta Folio
 dbDatos.Execute "UPDATE empeno SET FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',IDUsuarioMov=" & frmMDI.IDUsuario & ",Intereses=" & ConvMoneda(crInteres) & ",Importeiva=" & ConvMoneda(crImporteIva) & ",Pago=0,ImporteAlmacenaje=" & ConvMoneda(importeAlmacenaje) & ",ImporteSeguro=" & ConvMoneda(importeSeguro) & ",Efectivo=" & ConvMoneda(importeInteres + importeAlmacenaje + importeSeguro + crImporteIva) & " WHERE ID=" & IDEmpenoAnterior
Imprimir_Nota IDEmpenoAnterior, OD_REFRENDO, crInteres + crImporteIva, frmMDI.IDUsuario, DateAdd("D", Regresa_Valor_BD("DiasEnajenacion") + 1, fechaVencimiento)
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Sub PonerTotales()
Dim i As Integer, crPrestamo As Double, crIntereses As Double, crAlmacenaje, crSeguro As Double, crMoratorios As Double, crIva As Double
    
    crPrestamo = 0: crIntereses = 0: crAlmacenaje = 0: crSeguro = 0: crIva = 0
    
    For i = 1 To grdArticulos.Rows
        crPrestamo = crPrestamo + (Val(grdArticulos.CellText(i, 1)) * CDbl(grdArticulos.CellText(i, 5)))
        crIntereses = crIntereses + CDbl(grdArticulos.CellText(i, 7))
        crAlmacenaje = crAlmacenaje + CDbl(grdArticulos.CellText(i, 8))
        crSeguro = crSeguro + CDbl(grdArticulos.CellText(i, 9))
        crIva = crIva + CDbl(grdArticulos.CellText(i, 10))
        crMoratorios = crMoratorios + CDbl(grdArticulos.CellText(i, 11))
    Next i
    
    txtPrestamo.text = Format(crPrestamo, FMoneda)
    txtIntereses.text = Format(crIntereses + crAlmacenaje + crSeguro, FMoneda)
    txtMoratorios.text = Format(crMoratorios, FMoneda)
    txtIva.text = Format(crIva, FMoneda)
    txtTotal.text = Format(crPrestamo + crIntereses + crAlmacenaje + crSeguro + crMoratorios + crIva, FMoneda)
End Sub

Public Function Imprimir_Recibo_Venta(Folio As Long)
Dim rcAux As New ADODB.Recordset
Dim crPrestamo As Double, crIntereses As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crImporteIva As Double, Iva As Double

    rcAux.Open "SELECT SUM(Costo) AS Prestamo,SUM(Intereses) AS TotIntereses,SUM(Almacenaje) AS TotAlmacenaje,SUM(Seguro) AS TotSeguro,SUM(Moratorios) AS TotMoratorios FROM detallesventas INNER JOIN ventas ON detallesventas.IDVenta=ventas.ID WHERE ventas.Cancelado=0 AND ventas.TipoVenta=" & VENTACLIENTE & " AND ventas.Folio=" & Folio, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcAux.BOF And Not rcAux.EOF Then
        crPrestamo = rcAux!Prestamo
        crIntereses = rcAux!TotIntereses
        crAlmacenaje = rcAux!TotAlmacenaje
        crSeguro = rcAux!TotSeguro
        crMoratorios = rcAux!TotMoratorios
    End If
    rcAux.Close
    Set rcAux = Nothing
    
    'Saco el IVA
    Iva = Val(SacaValor("empeno INNER JOIN detallesentradainventario ON empeno.ID=detallesentradainventario.IDEmpeno LEFT JOIN detallesventas ON detallesentradainventario.ID=detallesventas.IDArticulo INNER JOIN ventas ON detallesventas.IDVenta=ventas.ID", "empeno.Iva", " WHERE ventas.Cancelado=0 AND ventas.TipoVenta=" & VENTACLIENTE & " AND ventas.Folio=" & Folio)) / 100
    crImporteIva = Redondeo((crIntereses + crAlmacenaje + crSeguro + crMoratorios) * Iva)
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\NotaCliente.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{ventas.Cancelado}=0 AND {ventas.Folio}=" & Folio & " AND {ventas.TipoVenta}=" & VENTACLIENTE
        .Formulas(0) = "Prestamo=" & crPrestamo & ""
        .Formulas(1) = "Intereses=" & crIntereses & ""
        .Formulas(2) = "Almacenaje=" & crAlmacenaje & ""
        .Formulas(3) = "Seguro=" & crSeguro & ""
        .Formulas(4) = "Moratorios=" & crMoratorios & ""
        .Formulas(5) = "ImporteIva=" & crImporteIva & ""
        .Formulas(6) = "IVA=" & Iva & ""
        .Formulas(7) = "ImporteLetra='" & Trim(CantidadEnLetra(CCur((crPrestamo + crIntereses + crAlmacenaje + crSeguro + crMoratorios + crImporteIva)))) & "'"
        .WindowTitle = "Nota Venta Cliente"
        .WindowState = crptMaximized
        .Action = 1
    End With
        
End Function

Public Sub GeneraInteresesEmpeno(ByVal IDEmpeno As Long)
Dim crTotal As Double, crImporte As Double, crInteres As Double, crAlmacenaje As Double, crSeguro As Double, crMoratorios As Double, crImporteIva As Double
Dim rcEmpeno As New ADODB.Recordset

rcEmpeno.Open "Select prestamoInicial,Operacion,vencimiento,serie from empeno where id=" & IDEmpeno, dbDatos, adOpenForwardOnly, adLockOptimistic
If Not rcEmpeno.EOF And Not rcEmpeno.BOF Then

IDEmpenoAnterior = IDEmpeno
fechaVencimiento = rcEmpeno!Vencimiento
crInteres = GeneraIntereses(IDEmpeno, "Tasa")
crAlmacenaje = GeneraIntereses(IDEmpeno, "Almacenaje")
crSeguro = GeneraIntereses(IDEmpeno, "Seguro")
crImporteIva = Redondeo(Regresa_Iva(crInteres + crAlmacenaje + crSeguro, IDEmpeno))

importeInteres = crInteres
importeAlmacenaje = crAlmacenaje
importeSeguro = crSeguro

    txtPrestamo.text = Format(rcEmpeno!PrestamoInicial, FMoneda)
    txtIntereses.text = Format(crInteres + crAlmacenaje + crSeguro, FMoneda)
    txtMoratorios.text = Format(0, FMoneda)
    txtIva.text = Format(crImporteIva, FMoneda)
    txtTotal.text = Format(rcEmpeno!PrestamoInicial + crInteres + crAlmacenaje + crSeguro + crImporteIva, FMoneda)
End If
End Sub

Function CalculaCambio(crEfectivo As Double, crImporte As Double) As Boolean
Dim lblLeyenda As Label, lblOperacion As Label
    
    Set lblOperacion = Total
    Set lblLeyenda = Leyenda
    
    lblOperacion.Caption = Format(crEfectivo - crImporte, FMoneda)
    lblOperacion.ForeColor = &HFF&
    lblLeyenda.Caption = "CAMBIO:"
    lblLeyenda.Tag = 1
    Abrir_Cajon
End Function


Sub Imprimir_Nota(IDEmpeno As Long, Opcion As Integer, Optional Abono As Double, Optional IDUsuarioMov As Integer, Optional Comercializacion As Date)
Dim ImprDefault As Boolean

On Error GoTo Error
    
    If Opcion = OD_REFRENDO Then
    End If
    ImprDefault = LocalizaImpresora(Regresa_Valor("Impresoras", "ImpresoraTickets", ""))
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & IIf(Opcion = OD_REFRENDO, "\Reportes\Nota.rpt", "\Reportes\NotaDesempeño.rpt")
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{empeno.ID}=" & IDEmpeno
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .Formulas(2) = "Opcion=" & Opcion & ""
        If Opcion = OD_REFRENDO Then
            .Formulas(3) = "Comercializacion='" & Format(Comercializacion, "DD-MMM-YYYY") & "'"
            .SubreportToChange = "OpcionesPagos"
            .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
            .SelectionFormula = "{opcionpagos.PC}='" & Nombre_Pc & "'"
            .DiscardSavedData = True
        End If
        
        .WindowState = crptMaximized
        .Destination = crptToPrinter
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                
        .WindowTitle = "Recibo"
        .Action = 1
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub
