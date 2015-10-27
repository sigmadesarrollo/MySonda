VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCambiosDirecto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambio Ventas"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7650
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCambiosDirecto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   7650
   Begin VB.TextBox txtCodigo1 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   13
      TabIndex        =   0
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtCodigo2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   13
      TabIndex        =   1
      Top             =   3360
      Width           =   2295
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   6480
      TabIndex        =   8
      Top             =   6840
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
      MaskColor       =   16777215
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmCambiosDirecto.frx":000C
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCambiar 
      Height          =   2055
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3625
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   17
   End
   Begin vbAcceleratorGrid6.vbalGrid grdNueva 
      Height          =   2055
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   3625
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
      DefaultRowHeight=   17
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   5280
      TabIndex        =   11
      Top             =   6840
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Guardar"
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
      Object.ToolTipText     =   ""
      Picture         =   "frmCambiosDirecto.frx":009D
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Prenda a cambiar:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   2235
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nueva Prenda:"
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
      Left            =   120
      TabIndex        =   6
      Top             =   3000
      Width           =   1800
   End
   Begin VB.Label lblPrecio2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Precio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6435
      TabIndex        =   5
      Top             =   5880
      Width           =   1005
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Diferencia:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5040
      TabIndex        =   4
      Top             =   6360
      Width           =   1140
   End
   Begin VB.Label lblNvoSaldo 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Saldo>"
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
      Left            =   6255
      TabIndex        =   3
      Top             =   6360
      Width           =   1185
   End
   Begin VB.Label lblPrecio1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "<Precio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6435
      TabIndex        =   2
      Top             =   2970
      Width           =   1005
   End
End
Attribute VB_Name = "frmCambiosDirecto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez

Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()

    If Validar_Cambio Then
        Grabar_Datos
        grdCambiar.Clear
        grdNueva.Clear
        txtCodigo1.Text = ""
        txtCodigo2.Text = ""
        SacaTotal grdCambiar, lblPrecio1
        SacaTotal grdNueva, lblPrecio2
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Limpiar
    Crear_Encabezados
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
    lblPrecio1.Caption = "0.00"
    lblPrecio2.Caption = "0.00"
    lblNvoSaldo.Caption = "0.00"
End Sub

Private Sub Crear_Encabezados()
    With grdCambiar
        .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K2", "Artículo", ecgHdrTextALignLeft, , 300, , , , , , , CCLSortString
        .AddColumn "K3", "Precio", ecgHdrTextALignRight, , 90, , , , , "###,###,###,###0.00", , CCLSortNumeric
    End With
   
    With grdNueva
        .AddColumn "K1", "Código", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K2", "Artículo", ecgHdrTextALignLeft, , 300, , , , , , , CCLSortString
        .AddColumn "K3", "Precio", ecgHdrTextALignRight, , 90, , , , , "###,###,###,###0.00", , CCLSortNumeric
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdCambiar_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If grdCambiar.SelectedRow > 0 And KeyCode = vbKeyDelete Then
        If MsgBox("Desea eliminar el articulo seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Cambio Ventas") = vbYes Then
            grdCambiar.RemoveRow grdCambiar.SelectedRow
            SacaTotal grdCambiar, lblPrecio1
            txtCodigo1.SetFocus
        End If
    End If
End Sub

Private Sub grdNueva_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)
    If grdNueva.SelectedRow > 0 And KeyCode = vbKeyDelete Then
        If MsgBox("Desea eliminar el articulo seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Cambio Ventas") = vbYes Then
            grdNueva.RemoveRow grdNueva.SelectedRow
            SacaTotal grdNueva, lblPrecio2
            txtCodigo2.SetFocus
        End If
    End If
End Sub

Private Sub txtCodigo1_GotFocus()
    Seleccionar_Texto txtCodigo1
    Cambiar_Color True, txtCodigo1
End Sub

Private Sub txtCodigo1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        
        Buscar_Codigo Trim(txtCodigo1.Text), True
        SacaTotal grdCambiar, lblPrecio1
    
    End If
End Sub

Private Sub txtCodigo1_LostFocus()
    Cambiar_Color False, txtCodigo1
End Sub

Private Sub txtCodigo2_GotFocus()
    Seleccionar_Texto txtCodigo2
    Cambiar_Color True, txtCodigo2
End Sub

Private Sub txtCodigo2_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Buscar_Codigo Trim(txtCodigo2), False
        SacaTotal grdNueva, lblPrecio2
    End If
End Sub

Private Sub txtCodigo2_LostFocus()
    Cambiar_Color False, txtCodigo2
End Sub

'Grabamos los datos del cambio
Private Sub Grabar_Datos()
Dim Movimiento As Long, Renglon As Integer, ccrEntrada As Double, ccrSalida As Double, ccrSaldo As Double

On Error GoTo error
    
    'Checo si hay Diferencia
    If Trim(lblNvoSaldo.Caption) <> "" Or Val(lblNvoSaldo.Caption) > 0 Then
        
        ccrSaldo = lblNvoSaldo.Caption
    Else
        
        ccrSaldo = 0
    End If
    
    'Regreso al inventario las prendas
    For Renglon = 1 To grdCambiar.Rows
        dbDatos.Execute "UPDATE DetallesEntradaInventario SET Cantidad=Cantidad+1 WHERE ID=" & grdCambiar.CellItemData(Renglon, 1)
        dbDatos.Execute "UPDATE DetallesVentas SET Devolucion=1 WHERE ID=" & grdCambiar.CellItemData(Renglon, 3)
        dbDatos.Execute "UPDATE Ventas SET ImporteDevolucion=" & ccrSaldo & " WHERE ID=" & grdCambiar.CellItemData(Renglon, 2)
        
        ccrEntrada = ccrEntrada + CDbl(grdCambiar.CellText(Renglon, 3))
    Next Renglon
    
    'Saco del inventario las prendas y las meto al detalle de Ventas
    For Renglon = 1 To grdNueva.Rows
        dbDatos.Execute "UPDATE DetallesEntradaInventario SET Cantidad=Cantidad-1 WHERE ID=" & grdNueva.CellItemData(Renglon, 1)
        GrabaDetalleVentas grdNueva.CellItemData(Renglon, 1), grdCambiar.CellItemData(grdCambiar.Rows, 2)
        
        ccrSalida = ccrSalida + CDbl(grdNueva.CellText(Renglon, 3))
    Next Renglon
        
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
'''''    'Grabamos el Abono
'''''    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
'''''                    "('" & Format(Date, "YYYY/MM/DD") & "','Cambio Venta'," & Movimiento & ",0,'CM01','110150'," & ccrEntrada & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'''''
'''''    'Grabo el Abono
'''''    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
'''''                    "('" & Format(Date, "YYYY/MM/DD") & "','Cambio Venta'," & Movimiento & ",0,'CM01','199450'," & ccrEntrada & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabo el Cargo
    dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','Cambio Venta'," & Movimiento & ",0,'CM01','620301'," & ccrEntrada & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
    'Si hay Diferencia grabo los movimientos Contables
    If ccrSaldo > 0 Then
        
        'Grabo el Cargo
        dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','Cambio Venta'," & Movimiento & ",0,'CM01','110101'," & ccrSaldo & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                  
        'Grabo el Cargo
        dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','Cambio Venta'," & Movimiento & ",0,'CM01','199401'," & ccrSaldo & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    End If
            
        'Grabo el Abono
        dbDatos.Execute "INSERT INTO Auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','Cambio Venta'," & Movimiento & ",0,'CM01','620350'," & ccrSalida & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                        
error:
    Maneja_Error Err
End Sub

Private Sub Buscar_Codigo(Codigo As String, Opcion As Boolean)
Dim rcInventario As New ADODB.Recordset
Dim IDVenta As Long, IDArticuloVenta As Long

On Error GoTo error

    rcInventario.Open "SELECT Codigo," & IIf(Opcion, "IDArticulo,Articulo,Precio,IDVenta,ID as ArticuloVenta", "ID as IDArticulo,Descripcion as Articulo,PrecioVitrina as Precio") & " FROM " & IIf(Opcion, "DetallesVentas", "DetallesEntradaInventario") & " where Codigo='" & Trim(Codigo) & "'" & IIf(Opcion, "", " And Cantidad>0"), dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcInventario.BOF And Not rcInventario.EOF Then
        With IIf(Opcion, grdCambiar, grdNueva)
            .AddRow
            .CellText(.Rows, 1) = rcInventario!Codigo
            .CellItemData(.Rows, 1) = rcInventario!IDArticulo
            .CellText(.Rows, 2) = rcInventario!Articulo
            If Opcion Then
                IDVenta = rcInventario!IDVenta
                IDArticuloVenta = rcInventario!ArticuloVenta
            Else
                IDVenta = 0
                IDArticuloVenta = 0
            End If
            .CellItemData(.Rows, 2) = IDVenta
            .CellText(.Rows, 3) = rcInventario!Precio
            .CellItemData(.Rows, 3) = IDArticuloVenta
            .CellTextAlign(.Rows, 3) = DT_RIGHT
        End With
    Else
        MsgBox "No se encontró el articulo especificado !!", vbInformation, "Cambios de Ventas"
    End If
    rcInventario.Close

error:
    Maneja_Error Err
    Set rcInventario = Nothing
End Sub

Private Sub Limpiar()
Dim ctrl As Control

    For Each ctrl In Controls
       If TypeOf ctrl Is Label And Mid(ctrl.Name, 1, 3) = "lbl" Then ctrl.Caption = ""
       If TypeOf ctrl Is TextBox And Mid(ctrl.Name, 1, 3) = "txt" Then ctrl.Text = ""
       If TypeOf ctrl Is vbalGrid Then ctrl.Clear
    Next
End Sub

Function SacaTotal(grid As vbalGrid, lbl As Label)
Dim i As Integer, Total As Double

    For i = 1 To grid.Rows
        Total = Total + CDbl(grid.CellText(i, 3))
    Next i
    
    lbl.Caption = Format(Total, "###,###,###,###0.00")
    SacaDiferencia
End Function

Sub SacaDiferencia()
Dim i As Integer, ccrTotal1 As Double, ccrTotal2 As Double

    For i = 1 To grdCambiar.Rows
        ccrTotal1 = ccrTotal1 + CDbl(grdCambiar.CellText(i, 3))
    Next i
    
    For i = 1 To grdNueva.Rows
        ccrTotal2 = ccrTotal2 + CDbl(grdNueva.CellText(i, 3))
    Next i
    
    lblNvoSaldo.Caption = Format(ccrTotal2 - ccrTotal1, "###,###,###,###0.00")
End Sub

Function GrabaDetalleVentas(IDArticulo As Long, IDVenta As Long)
Dim rcArticulo As New ADODB.Recordset

On Error GoTo error
    
    rcArticulo.Open "SELECT * FROM DetallesEntradaInventario WHERE ID=" & IDArticulo, dbDatos, adOpenForwardOnly, adLockOptimistic
    
    If Not rcArticulo.BOF And Not rcArticulo.EOF Then
        
        With rcArticulo
            
            dbDatos.Execute "INSERT INTO detallesventas (IDVenta,Codigo,Articulo,Kilates,peso,costo,precio,IDArticulo,Intereses,Almacenaje,Seguro) VALUES (" & _
                            IDVenta & ",'" & Trim(!Codigo) & "','" & Trim(!Descripcion) & "'," & !Kilates & "," & !Peso & "," & !costo & "," & !Precio & "," & IDArticulo & "," _
                             & !Intereses & "," & !Almacenaje & "," & !seguro & ")"
        
        End With
    
    End If
    
    rcArticulo.Close

error:
    Maneja_Error Err
    Set rcArticulo = Nothing
    
End Function

Function Validar_Cambio() As Boolean

    Validar_Cambio = True

    If grdCambiar.Rows <= 0 Or grdNueva.Rows <= 0 Then
        MsgBox "Debe capturar las prendas nuevas y a cambiar !", vbCritical, "Cambios"
        Validar_Cambio = False
    End If

    If Val(lblNvoSaldo.Caption) < 0 Then
        MsgBox "El importe de las nuevas prendas no debe ser menor al de las prendas a cambiar !", vbCritical, "Cambios"
        Validar_Cambio = False
    End If

End Function
