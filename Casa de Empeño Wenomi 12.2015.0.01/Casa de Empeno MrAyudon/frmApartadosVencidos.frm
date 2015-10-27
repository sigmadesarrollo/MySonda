VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmApartadosVencidos 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Apartados vencidos"
   ClientHeight    =   7560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmApartadosVencidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7560
   ScaleWidth      =   11310
   Begin vbAcceleratorGrid6.vbalGrid grdVencidos 
      Height          =   7065
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   12462
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
   Begin DevPowerFlatBttn.FlatBttn cmdRemate 
      Default         =   -1  'True
      Height          =   375
      Left            =   8745
      TabIndex        =   1
      Top             =   7125
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "    &Devoluciónes"
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
      Picture         =   "frmApartadosVencidos.frx":000C
      PictureDisabled =   "frmApartadosVencidos.frx":0376
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10185
      TabIndex        =   2
      Top             =   7125
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
      Picture         =   "frmApartadosVencidos.frx":04D0
   End
End
Attribute VB_Name = "frmApartadosVencidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdRemate_Click()
Dim Folio As Long, Movimiento As Long, crAbonos As Double, Cont As Long, crCosto As Double, crImporte As Double, crIva As Double, Iva As Double, Hora As String
Dim rcTmp As New ADODB.Recordset

On Error GoTo Error

    If MsgBox("Esta seguro de cancelar los apartados seleccionados ??", vbYesNo + vbDefaultButton2 + vbQuestion, "Apartados vencidos") = vbYes Then
    
        For Cont = 1 To grdVencidos.Rows - 1

            If grdVencidos.CellIcon(Cont, 1) = frmMDI.img.ItemIndex(1) Then  'si el articulo esta marcado, entonces hacemos el movimiento
                
                'Tomo el Folio
                Folio = grdVencidos.CellText(Cont, 2)
                
                'Tomo los Abonos
                crAbonos = grdVencidos.CellText(Cont, 6)
                
                'Tomo el Porcentaje de IVA
                Iva = CDbl(grdVencidos.CellText(Cont, 8)) / 100
                
                'Cancelamos la venta
                dbDatos.Execute "UPDATE ventas SET Cancelado=1,OrigenCancelacion=2,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & grdVencidos.CellItemData(Cont, 1)
                
                rcTmp.Open "SELECT IDArticulo,Costo,Precio FROM detallesventas WHERE IDVenta=" & grdVencidos.CellItemData(Cont, 1), dbDatos, adOpenForwardOnly, adLockReadOnly
                While Not rcTmp.EOF
                
                    'Regresamos los articulos
                    dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=Cantidad+1 WHERE ID=" & rcTmp!IDArticulo
                    
                    crCosto = crCosto + rcTmp!Costo
                    crImporte = crImporte + rcTmp!Precio
                rcTmp.MoveNext
                Wend
                rcTmp.Close
                Set rcTmp = Nothing
                
                'Cancelo el Abono
                dbDatos.Execute "UPDATE abonos SET Cancelado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE IDVenta=" & grdVencidos.CellItemData(Cont, 1)
                
                'Saco el Iva
                crIva = Redondeo(crImporte * (Iva))
                
                'Saco el Movimiento
                Movimiento = Regresa_Movimiento(False)
                Regresa_Movimiento True
            
                'Tomo la hora
                Hora = Time
                
                'Grabamos el cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                                Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Apartado Cancelado'," & Movimiento & "," & Folio & ",'CA01','620401'," & crImporte & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                            
                'Grabamos el cargo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                                Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Apartado Cancelado'," & Movimiento & "," & Folio & ",'CA01','620301'," & crCosto & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
                'Grabamos el abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                                Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Apartado Cancelado'," & Movimiento & "," & Folio & ",'CA50','620550'," & (crImporte + crIva) - crAbonos & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Grabamos el abono
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                                Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Apartado Cancelado'," & Movimiento & "," & Folio & ",'CA50','620250'," & crCosto & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                
                'Grabamos el abono
                'dbDatos.Execute "INSERT INTO auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                '                Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Apartado Cancelado'," & Movimiento & "," & Folio & ",'CA50','710450'," & crAbonos & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
                
                If crIva > 0 Then
                    'Grabamos el Cargo
                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES ('" & _
                                    Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','Apartado Cancelado'," & Movimiento & "," & Folio & ",'CA01','120101'," & crIva & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                End If
                
                
                'Se comento para que no refleje el dinero regresado de los abonos
                
'                If crAbonos > 0 Then
'                    'Grabamos el cargo Abonos
'                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
'                                              "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'CA50','110150'," & ConvMoneda(crAbonos) & "," & TIPO_ABONO & ",0,'Abonos Cancelados','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'                    'Grabamos el abono Abonos
'                    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
'                                              "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'CA01','620501'," & ConvMoneda(crAbonos) & "," & TIPO_CARGO & ",0,'Abonos Cancelados','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'                End If
                
            End If

        Next Cont
        
        'Imprimo los apartados
        Imprimir
        
        MsgBox "Apartados cancelados con éxito !!", vbInformation, "Apartados vencidos"
        
        'Refresco el Grid
        Cargar_Apartados
    End If
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Crear_Encabezados
    Cargar_Apartados
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

'Creamos los encabezados
Private Sub Crear_Encabezados()

    With grdVencidos
        .ImageList = frmMDI.img
        .AddColumn "K3", "Fecha", ecgHdrTextALignCentre, , 76, , , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K1", "Folio", ecgHdrTextALignCentre, , 65, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Cliente", ecgHdrTextALignLeft, , 258, , , , , , , CCLSortString
        .AddColumn "K4", "Vencimiento", ecgHdrTextALignCentre, , 72, , , , , "DD/MM/YY", , CCLSortDate
        .AddColumn "K5", "Total", ecgHdrTextALignRight, , 81, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Abonos", ecgHdrTextALignRight, , 81, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K7", "Saldo", ecgHdrTextALignRight, , 81, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K8", "IVA", ecgHdrTextALignRight, , 81, False, , , , , , CCLSortNumeric
    End With

End Sub

'Cargamos los apartados vencidos
Private Sub Cargar_Apartados()
Dim rcApartado As New ADODB.Recordset
Dim crAbonos As Double

On Error GoTo Error
    
    grdVencidos.Redraw = False
    grdVencidos.Clear
    
    rcApartado.Open "SELECT ventas.ID,ventas.Folio,ventas.Fecha,ventas.Vencimiento,ventas.Total,ventas.Descuento,ventas.Iva,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente FROM " _
                    & "ventas INNER JOIN clientes ON ventas.IDCliente=clientes.ID WHERE ventas.Pagado=0 AND ventas.Apartado=1 AND ventas.Cancelado=0 AND DATE_FORMAT(ADDDATE(Vencimiento,INTERVAL " & Val(Regresa_Valor_BD("DiasGraciaApa")) & " DAY),'%Y%/%m%/%d')<'" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
    
'''''    rcApartado.Open "SELECT ventas.ID,ventas.Folio,ventas.Fecha,ventas.Vencimiento,ventas.Total,ventas.Descuento,ventas.Iva,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente FROM " _
'''''                    & "ventas INNER JOIN clientes ON ventas.IDCliente=clientes.ID WHERE ventas.Pagado=0 AND ventas.Apartado=1 AND ventas.Cancelado=0 AND DATE_FORMAT(Fecha,'%Y%/%m%/%d')<='" & Format(Date, "YYYY/MM/DD") & "'", dbDatos, adOpenForwardOnly, adLockReadOnly
    With rcApartado
        
        While Not .EOF
            
            crAbonos = Regresa_Abonos(!ID)
            grdVencidos.AddRow
            grdVencidos.CellText(grdVencidos.Rows, 1) = !Fecha
            grdVencidos.CellItemData(grdVencidos.Rows, 1) = !ID
            grdVencidos.CellIcon(grdVencidos.Rows, 1) = frmMDI.img.ItemIndex(2)
            grdVencidos.CellTextAlign(grdVencidos.Rows, 1) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdVencidos.CellText(grdVencidos.Rows, 2) = !Folio
            grdVencidos.CellTextAlign(grdVencidos.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdVencidos.CellText(grdVencidos.Rows, 3) = !Cliente
            grdVencidos.CellTextAlign(grdVencidos.Rows, 3) = DT_LEFT Or DT_WORD_ELLIPSIS
            grdVencidos.CellText(grdVencidos.Rows, 4) = !Vencimiento
            grdVencidos.CellTextAlign(grdVencidos.Rows, 4) = DT_CENTER Or DT_WORD_ELLIPSIS
            grdVencidos.CellText(grdVencidos.Rows, 5) = (!Total - (!Total * (!Descuento / 100))) * (1 + (!Iva / 100))
            grdVencidos.CellTextAlign(grdVencidos.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdVencidos.CellText(grdVencidos.Rows, 6) = crAbonos
            grdVencidos.CellTextAlign(grdVencidos.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
            grdVencidos.CellText(grdVencidos.Rows, 7) = (!Total - (!Total * (!Descuento / 100))) * (1 + (!Iva / 100)) - crAbonos
            grdVencidos.CellTextAlign(grdVencidos.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
            
            grdVencidos.CellText(grdVencidos.Rows, 8) = !Iva
            
        .MoveNext
        Wend
        
    End With
    rcApartado.Close
    Set rcApartado = Nothing

    If grdVencidos.Rows > 0 Then
        grdVencidos.AddRow
        Poner_Totales
    End If
    
    grdVencidos.Redraw = True
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcApartado = Nothing
End Sub

Private Sub grdVencidos_Click(ByVal lRow As Long, ByVal lCol As Long)

    If lCol = 1 And lRow > 0 And lRow < grdVencidos.Rows Then grdVencidos.CellIcon(lRow, lCol) = IIf(grdVencidos.CellIcon(lRow, lCol) = frmMDI.img.ItemIndex(2), frmMDI.img.ItemIndex(1), frmMDI.img.ItemIndex(2))
End Sub

Private Sub grdVencidos_ColumnClick(ByVal lCol As Long)
    grdVencidos.RemoveRow (grdVencidos.Rows)
    Ordenar_Grid lCol, grdVencidos, 5, 6
    
    If grdVencidos.Rows > 0 Then
        grdVencidos.AddRow
        Poner_Totales
    End If
End Sub

Private Sub Poner_Totales()
Dim TotalAbo As Currency, Abonos As Currency, Saldo As Currency
Dim Renglon As Integer, Total As Integer, columna As Integer
Dim rcNota As New ADODB.Recordset
    
On Error GoTo Error
    
    'Hago la sumatoria de los totales (TotalAbo, Abonos, Saldos) desde el renglon 1 hasta el numero de renglones del GRID
    For Renglon = 1 To grdVencidos.Rows - 1
        TotalAbo = TotalAbo + CCur(grdVencidos.CellText(Renglon, 5))
        Abonos = Abonos + CCur(grdVencidos.CellText(Renglon, 6))
        Saldo = Saldo + CCur(grdVencidos.CellText(Renglon, 7))
        Total = Total + 1
    Next Renglon
         
    'En la ultima linea del GRID cargo los totales (TotalAbo, Abonos, Saldos) y cambio el color de la linea
    grdVencidos.CellText(grdVencidos.Rows, 7) = Saldo
    grdVencidos.CellTextAlign(grdVencidos.Rows, 7) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdVencidos.CellText(grdVencidos.Rows, 5) = TotalAbo
    grdVencidos.CellTextAlign(grdVencidos.Rows, 5) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdVencidos.CellText(grdVencidos.Rows, 6) = Abonos
    grdVencidos.CellTextAlign(grdVencidos.Rows, 6) = DT_RIGHT Or DT_WORD_ELLIPSIS
    grdVencidos.CellText(grdVencidos.Rows, 2) = Total
    grdVencidos.CellTextAlign(grdVencidos.Rows, 2) = DT_CENTER Or DT_WORD_ELLIPSIS

    For columna = 1 To grdVencidos.Columns
        
        grdVencidos.CellBackColor(grdVencidos.Rows, columna) = RGB(223, 208, 102)
    Next columna
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Sub Imprimir()
Dim i As Integer, strFolios As String
        
    strFolios = ""
    For i = 1 To grdVencidos.Rows
                
        If grdVencidos.CellIcon(i, 1) = frmMDI.img.ItemIndex(1) Then
            
            If strFolios = "" Then
            
                strFolios = " AND ({ventas.Folio}=" & grdVencidos.CellText(i, 2)
            Else
            
                strFolios = strFolios & " OR {ventas.Folio}=" & grdVencidos.CellText(i, 2)
            End If
            
        End If
        
    Next i
    
    If strFolios <> "" Then strFolios = strFolios & ")"
    With frmMDI.Cr
       .Reset
       .WindowShowPrintSetupBtn = True
       .WindowShowExportBtn = True
       .DiscardSavedData = True
       .ReportFileName = Path & "\Reportes\RepApartadosVencidos.rpt"
       .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
'       .SelectionFormula = "{ventas.FechaMovimiento}>=date('" & Format(Date, "YYYY,MM,DD") & "') AND {ventas.FechaMovimiento}<=date('" & Format(Date, "YYYY,MM,DD" & "'") & ")" & strFolios
       .SelectionFormula = "{ventas.FechaMovimiento}>=date('" & Format(Date, "YYYY,MM,DD") & "') AND {ventas.FechaMovimiento}<=date('" & Format(Date, "YYYY,MM,DD" & "'") & ")"
       .Formulas(0) = "Encabezado='" & "Del " & Format(Date, "dd/mmm/yyyy") & " a " & Format(Date, "dd/mmm/yyyy") & "'"
       .Formulas(1) = "Titulo='" & Sucursal.RazonSocial & "'"
       .Formulas(2) = "Subtitulo='SUCURSAL: " & Sucursal.NombreComercial & "'"
       .WindowTitle = "Reporte apartados rematados"
       .WindowState = crptMaximized
       .Destination = crptToWindow
       .Action = 1
    End With
End Sub
