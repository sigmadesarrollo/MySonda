VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmSalidaInventario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Salida de Inventario"
   ClientHeight    =   7695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11835
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSalidaInventario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7695
   ScaleWidth      =   11835
   Begin VB.Frame frmSalidas 
      Caption         =   "SALIDAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7080
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11745
      Begin VB.TextBox txtConcepto 
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
         Left            =   135
         MaxLength       =   150
         TabIndex        =   2
         Top             =   1320
         Width           =   7335
      End
      Begin VB.TextBox txtCodigo 
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
         Left            =   135
         MaxLength       =   13
         TabIndex        =   1
         Top             =   585
         Width           =   2175
      End
      Begin vbAcceleratorGrid6.vbalGrid grdSalidas 
         Height          =   4755
         Left            =   60
         TabIndex        =   3
         Top             =   1785
         Width           =   11610
         _ExtentX        =   20479
         _ExtentY        =   8387
         RowMode         =   -1  'True
         GridLines       =   -1  'True
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
      Begin DevPowerFlatBttn.FlatBttn cmdMostrar 
         Height          =   285
         Index           =   1
         Left            =   2340
         TabIndex        =   4
         Top             =   585
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   503
         AlignCaption    =   4
         AutoSize        =   0   'False
         Caption         =   ". . ."
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
      Begin VB.Label lblNumSal 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   6660
         Width           =   75
      End
      Begin VB.Label lblTotalSalida 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   9750
         TabIndex        =   5
         Top             =   6660
         Width           =   525
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Concepto de salida:"
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
         Left            =   135
         TabIndex        =   13
         Top             =   960
         Width           =   2085
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
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
         Left            =   135
         TabIndex        =   12
         Top             =   225
         Width           =   840
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   60
         TabIndex        =   11
         Top             =   6570
         Width           =   11640
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7560
         TabIndex        =   10
         Top             =   6660
         Width           =   735
      End
      Begin VB.Label lblFechaSalida 
         AutoSize        =   -1  'True
         Caption         =   "<Fecha>"
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
         Left            =   9240
         TabIndex        =   9
         Top             =   240
         Width           =   960
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         Left            =   8400
         TabIndex        =   8
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000C&
         Caption         =   "Num:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   7
         Top             =   6660
         Width           =   645
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10635
      TabIndex        =   14
      Top             =   7245
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
      Picture         =   "frmSalidaInventario.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   9435
      TabIndex        =   15
      Top             =   7245
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
      Picture         =   "frmSalidaInventario.frx":055E
   End
End
Attribute VB_Name = "frmSalidaInventario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
    Screen.MousePointer = vbHourglass
    If ValidaSalida Then Grabar_Salidas
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
    frmMuestraarticulos.Ver Me, txtCodigo, True, 0, 2
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

'Inicializamos la forma
Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    frmSalidas.BorderStyle = 0
    lblFechaSalida.Caption = Format(Date, "DD/MMM/YYYY")
    CentrarForm Me, frmMDI
    Crear_Encabezados
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados
Private Sub Crear_Encabezados()

    With grdSalidas
        .AddColumn "K1", "Tipo", ecgHdrTextALignLeft, , 90, False, , , , , , CCLSortString
        .AddColumn "K2", "Codigo", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortString
        .AddColumn "K3", "Descripción", ecgHdrTextALignLeft, , 220, , , , , , , CCLSortString
        .AddColumn "K4", "Kilates", ecgHdrTextALignCentre, , 60, , , , , , , CCLSortNumeric
        .AddColumn "K5", "Peso", ecgHdrTextALignRight, , 60, , , , , "0.000", , CCLSortNumeric
        .AddColumn "K6", "Costo", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Precio", ecgHdrTextALignRight, , 80, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "Marca", ecgHdrTextALignLeft, , 79, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K9", "Modelo", ecgHdrTextALignLeft, , 79, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K10", "Cantidad", ecgHdrTextALignLeft, , 79, False, , , , , , CCLSortNumeric
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdSalidas_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If grdSalidas.Rows > 0 And grdSalidas.SelectedRow > 0 Then
        
        If KeyCode = vbKeyDelete Then
            
            If MsgBox("Desea eliminar el articulo seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Salida de Inventario") = vbYes Then
                
                grdSalidas.RemoveRow grdSalidas.SelectedRow
                TotalesSalida
                txtCodigo.SetFocus
            
            Else
                
                grdSalidas.ClearSelection
                txtCodigo.SetFocus
            End If
            
        End If
    
    End If

End Sub

Private Sub txtCodigo_GotFocus()
    Seleccionar_Texto txtCodigo
    Cambiar_Color True, txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then MuestraDatosCodigo Trim(txtCodigo.text): txtCodigo.text = ""
End Sub

Private Sub txtCodigo_LostFocus()
    Cambiar_Color False, txtCodigo
End Sub

'Imprimimos la salida del inventario
Private Sub Imprimir_Salida(Folio As Long)

    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\RepSalidaInventario.rpt"
        .Formulas(0) = "Folio=" & Folio & ""
        .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
        .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
        .Destination = crptToWindow
        .WindowState = crptMaximized
        .WindowTitle = "Salida de inventario"
        .Action = 1
    End With

End Sub

'Grabamos la salida del inventario
Private Sub Grabar_Salidas()
Dim Indice As Integer, Folio As Long, Movimiento As Long, IDSalida As Long, crImporte As Double, Hora As String
    
On Error GoTo Error
    
    'Tomo la Hora
    Hora = Time
    
    'Saco el Folio
    Folio = Regresa_Movimiento(False, "FolioSalidaInventario")
    Regresa_Movimiento True, "FolioSalidaInventario"
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Tabla SalidaInventario
    dbDatos.Execute "INSERT INTO salidainventario (Fecha,Folio,IDUsuario,IDSucursal) VALUES ('" & _
                     Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Tomo el ID de la Salida
    IDSalida = SacaValor("salidainventario", "MAX(ID)")
       
    For Indice = 1 To grdSalidas.Rows
                
        'Detalles SalidaInventario
        dbDatos.Execute "INSERT INTO detallessalida (IDSalidaInventario,IDArticulo,Codigo,Descripcion,Kilates,Precio,Costo,Peso,Tipo,Serie) VALUES (" & _
                        IDSalida & "," & grdSalidas.CellItemData(Indice, 2) & ",'" & grdSalidas.CellText(Indice, 2) & "','" & grdSalidas.CellText(Indice, 3) & "'," & Val(grdSalidas.CellItemData(Indice, 4)) & "," & ConvMoneda(grdSalidas.CellText(Indice, 7)) & "," & ConvMoneda(grdSalidas.CellText(Indice, 6)) & "," & ConvMoneda(grdSalidas.CellText(Indice, 5)) & "," & Val(grdSalidas.CellItemData(Indice, 1)) & ",'')"
        
        'Se descuenta el Articulo de Inventario
        dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=Cantidad-1,TipoSalida=" & SALIDAINVENTARIO & " WHERE ID=" & grdSalidas.CellItemData(Indice, 2)
                
    Next Indice
   
    'Tomo el Importe
    crImporte = CDbl(lblTotalSalida.Caption)
   
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                    Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'SA01','200901'," & ConvMoneda(crImporte) & "," & TIPO_CARGO & ",0,'" & Trim(txtConcepto.text) & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                    Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'SA50','620350'," & ConvMoneda(crImporte) & "," & TIPO_ABONO & ",0,'" & Trim(txtConcepto.text) & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
      
    Sleep 1000
    
    'Imprimo el reporte de las salidas
    Imprimir_Salida Folio
   
    lblNumSal.Caption = ""
    lblTotalSalida.Caption = "0.00"
    grdSalidas.Clear
    txtCodigo.text = ""
    txtConcepto.text = ""
    txtCodigo.SetFocus

Error:
    Maneja_Error Err
End Sub

'Limpiamos la forma de salida
Private Sub Limpiar_Salida()
    grdSalidas.Clear
    txtCodigo.text = ""
    txtConcepto.text = ""
    txtCodigo.SetFocus
End Sub

Public Function MuestraDatos(ID As Long)
Dim rcArticulo As New ADODB.Recordset

On Error GoTo Error

    rcArticulo.Open "SELECT d.ID,d.Codigo,d.Descripcion,d.Kilates,d.Peso,d.Precio,d.Costo,d.Cantidad,d.Tipo,tipo.Descripcion AS Tipoo,d.Marca,d.Modelo " & _
                    "FROM detallesentradainventario d LEFT JOIN tipo ON d.Tipo=tipo.ID WHERE d.ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
        
        If VerificaPrenda(rcArticulo!ID, grdSalidas) Then
        
            With grdSalidas
                
                .AddRow
                .CellText(.Rows, 1) = rcArticulo!tipoo
                .CellItemData(.Rows, 1) = rcArticulo!Tipo
                .CellTextAlign(.Rows, 1) = DT_LEFT
                
                .CellText(.Rows, 2) = rcArticulo!Codigo
                .CellItemData(.Rows, 2) = rcArticulo!ID
                .CellTextAlign(.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 3) = rcArticulo!Descripcion
                .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 4) = SacaKilates(rcArticulo!Kilates)
                .CellItemData(.Rows, 4) = rcArticulo!Kilates
                .CellTextAlign(.Rows, 4) = DT_CENTER
                
                .CellText(.Rows, 5) = rcArticulo!Peso
                .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 6) = rcArticulo!Costo
                .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 7) = rcArticulo!Precio
                .CellTextAlign(.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 8) = rcArticulo!Marca
                .CellText(.Rows, 9) = rcArticulo!Modelo
                
                .CellText(.Rows, 10) = rcArticulo!Cantidad
            End With
        
        Else
                
            MsgBox "No se pueden agregar más prenda de las que existen en el inventario !!", vbCritical, "Salida de Inventario"
        End If
        
    rcArticulo.Close
    Set rcArticulo = Nothing
    TotalesSalida
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcArticulo = Nothing
End Function

Private Sub txtConcepto_GotFocus()
    Seleccionar_Texto txtConcepto
    Cambiar_Color True, txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtConcepto_LostFocus()
    Cambiar_Color False, txtConcepto
End Sub

Function ValidaSalida() As Boolean

    ValidaSalida = True

    If txtConcepto.text = "" Then
        MsgBox "Introduzca el concepto de salida !!", vbInformation, "Salida de inventario"
        ValidaSalida = False
        txtConcepto.SetFocus
        Exit Function
    End If

    If grdSalidas.Rows = 0 Then
        MsgBox "Seleccione las prendas a dar de baja !!", vbInformation, "Salida de inventario"
        ValidaSalida = False
        Exit Function
    End If

End Function

Sub TotalesSalida()
Dim crTotal As Double, Indice As Integer
    
    crTotal = 0
    For Indice = 1 To grdSalidas.Rows
    
        crTotal = crTotal + CDbl(grdSalidas.CellText(Indice, 6))
    Next Indice
    
    lblNumSal.Caption = grdSalidas.Rows
    lblTotalSalida.Caption = Format(crTotal, FMoneda)
End Sub

Function VerificaPrenda(IDArticulo As Long, Grid As vbalGrid) As Boolean
Dim i As Integer, x As Integer, Existencia As Integer, Cantidad As Integer
    
    Cantidad = 1
    VerificaPrenda = True
    
    For i = 1 To Grid.Rows
        
        Cantidad = 1
        Existencia = CInt(Grid.CellText(i, 10))
        
        For x = 1 To Grid.Rows
            
            If Grid.CellItemData(x, 2) = IDArticulo Then
                
                Cantidad = Cantidad + 1
                If Existencia < Cantidad Then VerificaPrenda = False: Exit For Else VerificaPrenda = True
            End If
        Next x
        
    Next i
    
End Function

Public Function MuestraDatosCodigo(Codigo As String)
Dim rcArticulo As New ADODB.Recordset

On Error GoTo Error

    rcArticulo.Open "SELECT d.ID,d.Codigo,d.Descripcion,d.Kilates,d.Peso,d.Precio,d.Costo,d.Cantidad,d.Tipo,tipo.Descripcion AS Tipoo,d.Marca,d.Modelo " & _
                    "FROM detallesentradainventario d INNER JOIN tipo ON d.Tipo=tipo.ID WHERE d.Cantidad>0 AND d.Codigo='" & Codigo & "'", dbDatos, adOpenForwardOnly, adLockOptimistic
        
        If VerificaPrenda(rcArticulo!ID, grdSalidas) Then
        
            With grdSalidas
                
                .AddRow
                .CellText(.Rows, 1) = rcArticulo!tipoo
                .CellItemData(.Rows, 1) = rcArticulo!Tipo
                .CellTextAlign(.Rows, 1) = DT_LEFT
                
                .CellText(.Rows, 2) = rcArticulo!Codigo
                .CellItemData(.Rows, 2) = rcArticulo!ID
                .CellTextAlign(.Rows, 2) = DT_LEFT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 3) = rcArticulo!Descripcion
                .CellTextAlign(.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 4) = SacaKilates(rcArticulo!Kilates)
                .CellItemData(.Rows, 4) = rcArticulo!Kilates
                .CellTextAlign(.Rows, 4) = DT_CENTER
                
                .CellText(.Rows, 5) = rcArticulo!Peso
                .CellTextAlign(.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 6) = rcArticulo!Costo
                .CellTextAlign(.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 7) = rcArticulo!Precio
                .CellTextAlign(.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
                
                .CellText(.Rows, 8) = rcArticulo!Marca
                .CellText(.Rows, 9) = rcArticulo!Modelo
                
                .CellText(.Rows, 10) = rcArticulo!Cantidad
            End With
        
        Else
                
            MsgBox "No se pueden agregar más prenda de las que existen en el inventario !!", vbCritical, "Salida de Inventario"
        End If
        
    rcArticulo.Close
    Set rcArticulo = Nothing
    TotalesSalida
    Exit Function
    
Error:
    Maneja_Error Err
    Set rcArticulo = Nothing
End Function
