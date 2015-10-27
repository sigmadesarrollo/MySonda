VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCancelaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelar Movimientos"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10965
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCancelaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7215
   ScaleWidth      =   10965
   Begin VB.TextBox Text1 
      BackColor       =   &H000000FF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   6735
      Width           =   300
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   60
      TabIndex        =   3
      Top             =   0
      Width           =   10860
      Begin VB.ComboBox cmbMovimiento 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmCancelaciones.frx":000C
         Left            =   4080
         List            =   "frmCancelaciones.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   195
         Width           =   4590
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "MOVIMIENTO A CANCELAR:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   300
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3915
      End
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCancelaciones 
      Height          =   5895
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   10398
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   9855
      TabIndex        =   1
      Top             =   6735
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
      Picture         =   "frmCancelaciones.frx":0010
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   8685
      TabIndex        =   2
      Top             =   6735
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Cancelar"
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
      Picture         =   "frmCancelaciones.frx":0562
      PictureDisabled =   "frmCancelaciones.frx":0AB4
   End
   Begin VB.Label Label2 
      Caption         =   "Movimiento Cancelado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   6765
      Width           =   2580
   End
End
Attribute VB_Name = "frmCancelaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim pIDUsuario As Integer

Public Property Let IDUsuario(Valor As Integer)
    pIDUsuario = Valor
End Property

Public Property Get IDUsuario() As Integer
    IDUsuario = pIDUsuario
End Property

Private Sub cmbMovimiento_Click()
    Saca_Movimiento
End Sub

Private Sub cmbMovimiento_GotFocus()
    Cambiar_Color True, cmbMovimiento
End Sub

Private Sub cmbMovimiento_LostFocus()
    Cambiar_Color False, cmbMovimiento
End Sub

Private Sub cmdCancelar_Click()

    Dim Movimiento As Integer, strDescripcion As String, VencimientoPago As String
    Dim crAmortizacion As Double, crPrestamoPrenda As Double, crPrestamoContrato As Double, IDEntrada As Long
    Dim rcAux As New ADODB.Recordset
    
    '***Puntos***
    Dim TarjetaPuntos As New ClienteFrecuente, DescuentoXPuntos As Currency, PuntosUsados As Long, PuntosAcumulados As Long, IDCliente As Long, IDTarjeta As Long
    Dim RefrendoImporte As Currency, DesempenoImporte As Currency, VentasImporte As Currency, ApartadoImporte As Currency, AbonoImporte As Currency

On Error GoTo Error

    With grdCancelaciones
        
        If .SelectedRow > 0 Then
            
            If .CellItemData(.SelectedRow, 3) > 0 Then .ClearSelection: Exit Sub
            If cmbMovimiento.ItemData(cmbMovimiento.ListIndex) = 3 Then If VerificaPagoFijo(.CellItemData(.SelectedRow, 2), .CellText(.SelectedRow, 6)) = False Then .ClearSelection: Exit Sub
            
            If MsgBox("Desea cancelar el movimiento seleccionado ??", vbInformation + vbYesNo + vbDefaultButton2, "Cancelar Movimientos") = vbYes Then
                
                'Tomo el motivo de la cancelación
                strDescripcion = frmMotivoCancela.Mostrar
                If strDescripcion = "" Then MsgBox "Es necesario que introduzca el motivo o razón por la cual se cancelará el movimiento !!", vbCritical, "Cancelar Movimientos": .ClearSelection: Exit Sub
                
                'Lo marco como que ya se cancelo
                .CellItemData(.SelectedRow, 3) = 1
                
                Select Case cmbMovimiento.ItemData(cmbMovimiento.ListIndex)
                Case 1
                            
                    Movimiento = 1
                    
                    IDCliente = Val(SacaValor("empeno", "IDCliente", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                    
                    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(IDCliente) Then
                        PuntosAcumulados = CLng(SacaValor("empeno", "puntosacumuladosemp", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        If PuntosAcumulados > 0 Then TarjetaPuntos.CuentaFrecuente.Redimir_Puntos EmpenoCancelacion, PuntosAcumulados, .CellText(.SelectedRow, 4), Me.IDUsuario, .CellText(.SelectedRow, 2)
                    End If
                    
                    'Cancelo en Empeños
                    dbDatos.Execute "UPDATE empeno SET Cancelado=1 WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellItemData(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND Concepto='Empeño'"
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellItemData(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                    
                Case 3
                               
                    Movimiento = 3
                    
                    '***Puntos***
                    
                    'Hago una copia del Desempeño que se cancelo
                    CopiaContrato .CellItemData(.SelectedRow, 1)
                    
                    IDCliente = Val(SacaValor("empeno", "IDCliente", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                    
                    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(IDCliente) Then
                    
                        DesempenoImporte = Val(SacaValor("empeno", "(Pago+Intereses+ImporteAlmacenaje+ImporteSeguro+ImporteMoratorios+ImportePerdida+ImporteOtros+ImporteIva)", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        PuntosUsados = Val(SacaValor("empeno", "puntosusados", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        PuntosAcumulados = Val(SacaValor("empeno", "puntosacumulados", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        DescuentoXPuntos = Val(SacaValor("empeno", "descuentoxpuntos", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        IDTarjeta = Val(SacaValor("empeno", "IDTarjeta", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        
                        If PuntosUsados > 0 Then TarjetaPuntos.CuentaFrecuente.Acumula_Desempeno_Cancelacion DesempenoCancelacion, Me.IDUsuario, PuntosUsados, DesempenoImporte, .CellText(.SelectedRow, 2), IDTarjeta
                        If PuntosAcumulados > 0 Then TarjetaPuntos.CuentaFrecuente.Redimir_Puntos DesempenoCancelacion, CLng(PuntosAcumulados), DesempenoImporte - DescuentoXPuntos, Me.IDUsuario, .CellText(.SelectedRow, 2)
                        
                    End If
                    
                    'Cancelo el Desempeño
                    dbDatos.Execute "UPDATE empeno SET Destino=0,FechaMovimiento=NULL,Pagado=0,Pago=0,Intereses=0,ImporteAlmacenaje=0,ImporteSeguro=0,ImporteMoratorios=0,ImportePerdida=0,ImporteIva=0,ImporteOtros=0,FolioNota=0,DescuentoXPuntos=0,SaldoPuntosAnterior=0,PuntosUsados=0,PuntosAcumulados=0,SaldoPuntosActual=0,IDTarjeta=0 WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellItemData(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND (Concepto='Desempeño' OR Concepto='Boleta perdida' OR Concepto = 'Redencion Puntos Desempeño') AND Iniciales='" & .CellText(.SelectedRow, 5) & "'"
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellItemData(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"

                Case 2
                                                                                           
                    Movimiento = 2
                    
                    '***Puntos***
                    
                    IDCliente = Val(SacaValor("empeno", "IDCliente", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                    
                    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(IDCliente) Then
                    
                        RefrendoImporte = Val(SacaValor("empeno", "(Pago+Intereses+ImporteAlmacenaje+ImporteSeguro+ImporteMoratorios+ImportePerdida+ImporteOtros+ImporteIva)", " WHERE NumContrato=" & CLng(.CellText(.SelectedRow, 2)) & " AND Folio=" & CLng(.CellItemData(.SelectedRow, 5)) & " AND Destino=" & OD_REFRENDO))
                        PuntosUsados = Val(SacaValor("empeno", "puntosusados", " WHERE NumContrato=" & CLng(.CellText(.SelectedRow, 2)) & " AND Folio=" & CLng(.CellItemData(.SelectedRow, 5)) & " AND Destino=" & OD_REFRENDO))
                        PuntosAcumulados = Val(SacaValor("empeno", "puntosacumulados", " WHERE NumContrato=" & CLng(.CellText(.SelectedRow, 2)) & " AND Folio=" & CLng(.CellItemData(.SelectedRow, 5)) & " AND Destino=" & OD_REFRENDO))
                        DescuentoXPuntos = Val(SacaValor("empeno", "descuentoxpuntos", " WHERE NumContrato=" & CLng(.CellText(.SelectedRow, 2)) & " AND Folio=" & CLng(.CellItemData(.SelectedRow, 5)) & " AND Destino=" & OD_REFRENDO))
                        IDTarjeta = Val(SacaValor("empeno", "IDTarjeta", " WHERE NumContrato=" & CLng(.CellText(.SelectedRow, 2)) & " AND Folio=" & CLng(.CellItemData(.SelectedRow, 5)) & " AND Destino=" & OD_REFRENDO))
                        
                        If PuntosUsados > 0 Then TarjetaPuntos.CuentaFrecuente.Acumula_Refrendos_Cancelacion RefrendoCancelacion, Me.IDUsuario, PuntosUsados, RefrendoImporte, .CellText(.SelectedRow, 2), IDTarjeta
                        If PuntosAcumulados > 0 Then TarjetaPuntos.CuentaFrecuente.Redimir_Puntos RefrendoCancelacion, CLng(PuntosAcumulados), RefrendoImporte - DescuentoXPuntos, Me.IDUsuario, .CellText(.SelectedRow, 2)
                        
                    End If
                    
                    'Tomo lo Importes
                    rcAux.Open "SELECT ID,Pago,Intereses,ImporteAlmacenaje,ImporteSeguro,ImporteMoratorios,ImportePerdida,ImporteOtros,ImporteIva,FolioNota," & _
                        "DescuentoXPuntos,SaldoPuntosAnterior,PuntosUsados,PuntosAcumulados,SaldoPuntosActual,IDTarjeta FROM empeno WHERE NumContrato=" & CLng(.CellText(.SelectedRow, 2)) & " AND Folio=" & CLng(.CellItemData(.SelectedRow, 5)) & " AND Destino=" & OD_REFRENDO, dbDatos, adOpenForwardOnly, adLockOptimistic
                    
                    'Cancelo el Refrendo
                    dbDatos.Execute "UPDATE empeno SET Cancelado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',Pago=" & rcAux!Pago & ",Intereses=" & rcAux!Intereses & ",ImporteAlmacenaje=" & rcAux!ImporteAlmacenaje & ",ImporteSeguro=" & rcAux!ImporteSeguro & ",ImporteMoratorios=" & rcAux!ImporteMoratorios & ",ImportePerdida=" & rcAux!ImportePerdida & ",ImporteOtros=" & rcAux!ImporteOtros & ",ImporteIva=" & rcAux!ImporteIva & ",FolioNota=" & rcAux!FolioNota & _
                         ",DescuentoXPuntos=" & rcAux!DescuentoXPuntos & ",SaldoPuntosAnterior=" & rcAux!SaldoPuntosAnterior & ",PuntosUsados=" & rcAux!PuntosUsados & ",PuntosAcumulados=" & rcAux!PuntosAcumulados & ",SaldoPuntosActual=" & rcAux!SaldoPuntosActual & ",IDTarjeta=" & rcAux!IDTarjeta & " WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Pongo vivo el contrato original
                    dbDatos.Execute "UPDATE empeno SET Destino=0,FolioDestino=0,FechaMovimiento=NULL,Pagado=0,Pago=0,Intereses=0,ImporteAlmacenaje=0,ImporteSeguro=0,ImporteMoratorios=0,ImportePerdida=0,ImporteOtros=0,ImporteIva=0,FolioNota=0," & _
                        "DescuentoXPuntos=0,SaldoPuntosAnterior=0,PuntosUsados=0,PuntosAcumulados=0,SaldoPuntosActual=0,IDTarjeta=0 WHERE ID=" & rcAux!ID
                    
                    rcAux.Close
                    Set rcAux = Nothing
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE (Folio=" & .CellItemData(.SelectedRow, 2) & " Or Folio=" & .CellItemData(.SelectedRow, 5) & ") AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND Movimiento=" & .CellItemData(.SelectedRow, 4) & " AND (Concepto='Refrendo' OR Concepto='Abono Refrendo' OR Concepto='Boleta perdida' OR Concepto = 'Redencion Puntos Refrendo') AND Iniciales='" & .CellText(.SelectedRow, 5) & "'"
                               
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellItemData(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                                                
                Case 6
                    
                    Movimiento = 6
                    
                    .CellItemData(.SelectedRow, 6) = 0
                    
                    'Tomo los valores
                    rcAux.Open "SELECT p.IDEmpeno,p.NumPago,p.Vencimiento,p.Pago,p.Interes,p.Almacenaje,p.Seguro,p.Amortizacion,p.Moratorios,p.Saldo,e.Prestamo FROM pagosfijos p INNER JOIN empeno e ON p.IDEmpeno=e.ID WHERE p.ID=" & CLng(.CellItemData(.SelectedRow, 1)), dbDatos, adOpenForwardOnly, adLockOptimistic
                    
                    crPrestamoContrato = rcAux!Prestamo
                    crAmortizacion = rcAux!Amortizacion
                    
                    'Creo el registro habilitado
                    dbDatos.Execute "INSERT INTO pagosfijos (IDEmpeno,NumPago,Vencimiento,Pago,Interes,Almacenaje,Seguro,Amortizacion,Saldo) VALUES (" & _
                                    rcAux!IDEmpeno & "," & rcAux!NumPago & ",'" & Format(rcAux!Vencimiento, "YYYY/MM/DD") & "'," & rcAux!Pago & "," & rcAux!Interes & "," & rcAux!Almacenaje & "," & rcAux!Seguro & "," & rcAux!Amortizacion & "," & rcAux!Saldo & ")"
                    
                    rcAux.Close
                    
                    'Cancelo el movimiento de pagos fijos
                    dbDatos.Execute "UPDATE pagosfijos SET Cancelado=1 WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Recalculo la Fecha de Vencimiento
                    If Val(.CellText(.SelectedRow, 6)) = 1 Then
                        
                        VencimientoPago = DateAdd("M", 2, CDate(.CellText(.SelectedRow, 7)))
                    Else
                    
                        VencimientoPago = DateAdd("M", 2, UltimaFecha(.CellItemData(.SelectedRow, 2), .CellText(.SelectedRow, 6)))
                    End If
                    
                    'Restablezco el préstamo y la fecha de vencimiento del contrato _
                     antes de que se registrara el pago fijo
                    dbDatos.Execute "UPDATE empeno SET Prestamo=Prestamo+" & crAmortizacion & ",Vencimiento='" & Format(CDate(VencimientoPago), "YYYY/MM/DD") & "',Destino=0,Pagado=0,FechaMovimiento=NULL WHERE ID=" & CLng(.CellItemData(.SelectedRow, 2))
                    
                    'Restablezco el préstamo de cada prenda
                    rcAux.Open "SELECT d.ID AS IDPrenda,d.Prestamo FROM detallesempeno d WHERE d.IDEmpeno=" & CLng(.CellItemData(.SelectedRow, 2)), dbDatos, adOpenForwardOnly, adLockOptimistic
                    While Not rcAux.EOF
                        
                        crPrestamoPrenda = (rcAux!Prestamo * 100) / crPrestamoContrato
                        crPrestamoPrenda = Redondeo((crPrestamoContrato + crAmortizacion) * (crPrestamoPrenda / 100))
                        dbDatos.Execute "UPDATE detallesempeno SET Prestamo=" & crPrestamoPrenda & " WHERE ID=" & rcAux!IDPrenda
                    rcAux.MoveNext
                    Wend
                    rcAux.Close
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellItemData(.SelectedRow, 4) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND Concepto='Pagos Fijos' AND Movimiento=" & CLng(.CellItemData(.SelectedRow, 5))
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                
                Case 4
                    
                    Movimiento = 4
                    
                    'Cancelo el movimiento en divisas
                    dbDatos.Execute "UPDATE divisas SET Cancelado=1 WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND (Concepto='Divisas' OR Concepto='Dotacion Divisas' OR Concepto='Retiro Divisas')"
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                
                Case 5
                    
                    Movimiento = 5
                    
                    'Cancelo el movimiento en gastos
                    dbDatos.Execute "UPDATE gastos SET Cancelado=1 WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND Concepto='Gastos'"
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                
                Case 7
                    
                    Movimiento = 7
                    
                    '***Puntos***
                    
                    IDCliente = Val(SacaValor("ventas", "IDCliente", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                    
                    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(IDCliente) Then
                    
                        VentasImporte = Val(SacaValor("ventas", "(Total-DescuentoEfectivo)", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        PuntosUsados = Val(SacaValor("ventas", "puntosusados", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        PuntosAcumulados = Val(SacaValor("ventas", "puntosacumulados", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        DescuentoXPuntos = Val(SacaValor("ventas", "descuentoxpuntos", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        IDTarjeta = Val(SacaValor("ventas", "IDTarjeta", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        
                        If PuntosUsados > 0 Then TarjetaPuntos.CuentaFrecuente.Acumula_Ventas_Cancelacion VentasCancelacion, Me.IDUsuario, PuntosUsados, VentasImporte, .CellText(.SelectedRow, 2), IDTarjeta
                        If PuntosAcumulados > 0 Then TarjetaPuntos.CuentaFrecuente.Redimir_Puntos VentasCancelacion, CLng(PuntosAcumulados), VentasImporte - DescuentoXPuntos, Me.IDUsuario, .CellText(.SelectedRow, 2)
                        
                    End If
                    
                    'Cancelo el movimiento en ventas
                    dbDatos.Execute "UPDATE ventas SET Cancelado=1,OrigenCancelacion=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Regreso las prendas al inventario
                    rcAux.Open "SELECT IDArticulo FROM detallesventas WHERE IDVenta=" & .CellItemData(.SelectedRow, 1), dbDatos, adOpenForwardOnly, adLockOptimistic
                    If Not rcAux.BOF And Not rcAux.EOF Then
                        While Not rcAux.EOF
                        
                            dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=Cantidad+1 WHERE ID=" & rcAux!IDArticulo
                        rcAux.MoveNext
                        Wend
                    End If
                    rcAux.Close
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND (Concepto='Ventas' OR Concepto = 'Redencion Puntos Ventas')"
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                
                Case 8
                    
                    Movimiento = 8
                    
                    '***Puntos***
                    
                    IDCliente = Val(SacaValor("ventas", "IDCliente", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                    
                    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(IDCliente) Then
                    
                        ApartadoImporte = Val(SacaValor("abonos", "Importe", " WHERE IDVenta=" & .CellItemData(.SelectedRow, 1) & " ORDER BY Fecha LIMIT 1"))
                        PuntosUsados = Val(SacaValor("ventas", "puntosusados", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        PuntosAcumulados = Val(SacaValor("ventas", "puntosacumulados", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        DescuentoXPuntos = Val(SacaValor("ventas", "descuentoxpuntos", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        IDTarjeta = Val(SacaValor("ventas", "IDTarjeta", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        
                        If PuntosUsados > 0 Then TarjetaPuntos.CuentaFrecuente.Acumula_Apartados_Cancelacion ApartadosCancelacion, Me.IDUsuario, PuntosUsados, ApartadoImporte, .CellText(.SelectedRow, 2), IDTarjeta
                        If PuntosAcumulados > 0 Then TarjetaPuntos.CuentaFrecuente.Redimir_Puntos ApartadosCancelacion, CLng(PuntosAcumulados), ApartadoImporte - DescuentoXPuntos, Me.IDUsuario, .CellText(.SelectedRow, 2)
                        
                    End If
                    
                    'Cancelo el movimiento en ventas de apartado
                    dbDatos.Execute "UPDATE ventas SET Cancelado=1,OrigenCancelacion=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Cancelo el Abono
                    dbDatos.Execute "UPDATE abonos SET Cancelado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE DATE_FORMAT(Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' AND IDVenta=" & .CellItemData(.SelectedRow, 1)
                    
                    'Regreso las prendas al inventario
                    rcAux.Open "SELECT IDArticulo FROM detallesventas WHERE IDVenta=" & .CellItemData(.SelectedRow, 1), dbDatos, adOpenForwardOnly, adLockOptimistic
                    If Not rcAux.BOF And Not rcAux.EOF Then
                        While Not rcAux.EOF
                        
                            dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=Cantidad+1 WHERE ID=" & rcAux!IDArticulo
                        rcAux.MoveNext
                        Wend
                    End If
                    rcAux.Close
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND (Concepto='Apartado' OR Concepto='Abonos' OR Concepto = 'Redencion Puntos Apartado')"
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                 
                Case 9
                    
                    Movimiento = 9
                    
                    '***Puntos***
                    
                    IDCliente = Val(SacaValor("ventas", "IDCliente", " WHERE ID=" & .CellItemData(.SelectedRow, 5)))
                    
                    If TarjetaPuntos.CuentaFrecuente.FindCuentaByIDCliente(IDCliente) Then
                    
                        AbonoImporte = Val(SacaValor("abonos", "Importe", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        PuntosUsados = Val(SacaValor("abonos", "puntosusados", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        PuntosAcumulados = Val(SacaValor("abonos", "puntosacumulados", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        DescuentoXPuntos = Val(SacaValor("abonos", "descuentoxpuntos", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        IDTarjeta = Val(SacaValor("abonos", "IDTarjeta", " WHERE ID=" & .CellItemData(.SelectedRow, 1)))
                        
                        If PuntosUsados > 0 Then TarjetaPuntos.CuentaFrecuente.Acumula_Abonos_Cancelacion AbonosCancelacion, Me.IDUsuario, PuntosUsados, AbonoImporte, .CellText(.SelectedRow, 3), IDTarjeta
                        If PuntosAcumulados > 0 Then TarjetaPuntos.CuentaFrecuente.Redimir_Puntos AbonosCancelacion, CLng(PuntosAcumulados), AbonoImporte - DescuentoXPuntos, Me.IDUsuario, .CellText(.SelectedRow, 3)
                        
                    End If
                    
                    'Cancelo el Abono
                    dbDatos.Execute "UPDATE abonos SET Cancelado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Desmarco si ya estaba liquidado
                    dbDatos.Execute "UPDATE ventas SET Pagado=0,FechaMovimiento=NULL WHERE ID=" & .CellItemData(.SelectedRow, 5)
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 3) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND (Concepto='Abonos' OR Concepto = 'Redencion Puntos Abonos') AND Movimiento=" & .CellItemData(.SelectedRow, 4)

                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 3) & "," & .CellText(.SelectedRow, 3) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                
                Case 10
                 
                    Movimiento = 10
                    
                    'Cancelo la Dotación
                    dbDatos.Execute "UPDATE boveda SET Cancelado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & .CellItemData(.SelectedRow, 1)

                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND (Iniciales='" & IIf(.CellItemData(.SelectedRow, 2) = 1, "DO01", "RE01") & "' OR Iniciales='" & IIf(.CellItemData(.SelectedRow, 2) = 1, "DO50", "RE50") & "')"
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"

                Case 11
                    
                    Movimiento = 11

                    'Cancelo el Movimiento
                    dbDatos.Execute "UPDATE bancos SET Cancelado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & .CellItemData(.SelectedRow, 1)

                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND (Iniciales='BA01' OR Iniciales='BA50') AND Serie=" & .CellItemData(.SelectedRow, 2)
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"

                Case 12
                
                    Movimiento = 12
                    
                    'Cancelo el Movimiento
                    dbDatos.Execute "UPDATE bancos SET Cancelado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & .CellItemData(.SelectedRow, 1)

                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND (Iniciales='TR01' OR Iniciales='TR50') AND Serie=" & .CellItemData(.SelectedRow, 2)
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                
                Case 13
                    
                    Movimiento = 13
                    
                    'Cancelo en Compras
                    dbDatos.Execute "UPDATE compras SET Cancelado=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & .CellItemData(.SelectedRow, 1)

                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND Concepto='Compras'"
                    
                    'Saco el id de la entrada
                    IDEntrada = SacaValor("entradainventario", "ID", " WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND TipoEntrada=" & ENTRADACOMPRA)
                    
                    'Doy de baja las prendas en el inventario
                    dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=0 WHERE IDEntrada=" & IDEntrada
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"
                
                Case 14
                    
                    Movimiento = 14
                    
                    'Cancelo el movimiento en ventas
                    dbDatos.Execute "UPDATE ventas SET Cancelado=1,OrigenCancelacion=1,FechaMovimiento='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "' WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    
                    'Regreso las prendas al inventario
                    rcAux.Open "SELECT IDArticulo FROM detallesventas WHERE IDVenta=" & .CellItemData(.SelectedRow, 1), dbDatos, adOpenForwardOnly, adLockOptimistic
                    While Not rcAux.EOF
                    
                        dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=Cantidad+1 WHERE ID=" & rcAux!IDArticulo
                    rcAux.MoveNext
                    Wend
                    rcAux.Close
                    
                    'Cancelo en auxiliar
                    dbDatos.Execute "UPDATE auxiliar SET Importe=0 WHERE Folio=" & .CellText(.SelectedRow, 2) & " AND Fecha='" & Format(Date, "YYYY-MM-DD") & "' AND Concepto='Ventas'"
                    
                    'Grabo el usuario que cancelo
                    dbDatos.Execute "INSERT INTO cancelaciones (Fecha,TipoMovimiento,Contrato,Folio,IDUsuario,Descripcion) VALUES ('" & _
                                    Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Movimiento & "," & .CellText(.SelectedRow, 2) & "," & .CellText(.SelectedRow, 2) & "," & Me.IDUsuario & ",'" & strDescripcion & "')"

                End Select
                
                'Lo coloreo como cancelado
                Colorea grdCancelaciones, grdCancelaciones.SelectedRow, RGB(244, 119, 66)
                MsgBox "Movimiento cancelado con éxito !!", vbInformation, "Cancelar Movimientos"
            
            End If
            
            'Quitola selección
            grdCancelaciones.ClearSelection
        
        End If
    
    End With
    Exit Sub
    
Error:
    Maneja_Error Err
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Poner_Flat Fl, Me.Controls, Me
    CargaCombo
    cmbMovimiento.ListIndex = -1
    CentrarForm Me, frmMDI
End Sub

Sub CargaCombo()
    
    cmbMovimiento.AddItem "EMPEÑOS"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 1
    
    cmbMovimiento.AddItem "DESEMPEÑOS"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 3
    
    cmbMovimiento.AddItem "REFRENDOS"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 2
    
'''''    cmbMovimiento.AddItem "PAGOS FIJOS"
'''''    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 6
    
'''''    cmbMovimiento.AddItem "DIVISAS"
'''''    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 4
    
    cmbMovimiento.AddItem "GASTOS"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 5
    
    cmbMovimiento.AddItem "VENTAS MOSTRADOR"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 7
    
    cmbMovimiento.AddItem "VENTAS CLIENTE"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 14
    
'''''    cmbMovimiento.AddItem "APARTADOS"
'''''    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 8
    
    cmbMovimiento.AddItem "ABONOS A APARTADOS"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 9
    
    cmbMovimiento.AddItem "COMPRAS"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 13
    
    cmbMovimiento.AddItem "CAJA GENERAL"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 10
    
    cmbMovimiento.AddItem "BOVEDA"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 11
    
    cmbMovimiento.AddItem "BANCOS"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 12
    
    cmbMovimiento.AddItem "CIERRE DE CAJA"
    cmbMovimiento.ItemData(cmbMovimiento.NewIndex) = 15
       
End Sub

Sub Crear_Encabezados(Opcion As Integer)
    
    With grdCancelaciones
    
        .Redraw = False
        .Clear True
        
        Select Case Opcion
        Case 1, 2, 3
        
            .AddColumn "C1", "Fecha", ecgHdrTextALignLeft, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", "Contrato", ecgHdrTextALignCentre, , 80, , , , , , , CCLSortNumeric
            .AddColumn "C3", "Cliente", ecgHdrTextALignLeft, , 330, , , , , , , CCLSortString
            .AddColumn "C4", "Préstamo", ecgHdrTextALignRight, , 120, , , , , FMoneda, , CCLSortNumeric
            .AddColumn "C5", "Iniciales", ecgHdrTextALignLeft, , 50, False, , , , , , CCLSortString
            .AddColumn "C6", "Serie", ecgHdrTextALignLeft, , 50, False, , , , , , CCLSortNumeric
        
        Case 4
        
            .AddColumn "C1", "Fecha", ecgHdrTextALignLeft, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", "Folio", ecgHdrTextALignCentre, , 80, , , , , , , CCLSortNumeric
            .AddColumn "C3", "Divisa", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortString
            .AddColumn "C4", "Tipo Cambio", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortString
            .AddColumn "C5", "Cantidad", ecgHdrTextALignRight, , 105, , , , , FMoneda, , CCLSortNumeric
            .AddColumn "C6", "Movimiento", ecgHdrTextALignCentre, , 100, , , , , FMoneda, , CCLSortNumeric
        
        Case 5
        
            .AddColumn "C1", "Fecha", ecgHdrTextALignLeft, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", "Folio", ecgHdrTextALignCentre, , 80, , , , , , , CCLSortNumeric
            .AddColumn "C3", "Cuenta", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
            .AddColumn "C4", "Concepto", ecgHdrTextALignLeft, , 255, , , , , , , CCLSortString
            .AddColumn "C5", "Importe", ecgHdrTextALignRight, , 105, , , , , FMoneda, , CCLSortNumeric
        
        Case 6
        
            .AddColumn "C1", "Fecha", ecgHdrTextALignLeft, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", "Contrato", ecgHdrTextALignCentre, , 80, , , , , , , CCLSortNumeric
            .AddColumn "C3", "Cliente", ecgHdrTextALignLeft, , 255, , , , , , , CCLSortString
            .AddColumn "C4", "Num. Pago", ecgHdrTextALignLeft, , 80, , , , , , , CCLSortString
            .AddColumn "C5", "Importe", ecgHdrTextALignRight, , 110, , , , , FMoneda, , CCLSortNumeric
            .AddColumn "C6", "", ecgHdrTextALignRight, , 110, False, , , , FMoneda, , CCLSortNumeric
            .AddColumn "C7", "", ecgHdrTextALignRight, , 110, False, , , , FMoneda, , CCLSortNumeric
            
        Case 7
            
            .AddColumn "C1", "Fecha", ecgHdrTextALignLeft, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", "Folio", ecgHdrTextALignCentre, , 80, , , , , , , CCLSortNumeric
            .AddColumn "C3", "Cliente", ecgHdrTextALignLeft, , 355, , , , , , , CCLSortString
            .AddColumn "C4", "Importe", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortString
        
        Case 8
            
            .AddColumn "C1", "Fecha", ecgHdrTextALignLeft, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", "Folio", ecgHdrTextALignCentre, , 80, , , , , , , CCLSortNumeric
            .AddColumn "C3", "Cliente", ecgHdrTextALignLeft, , 255, , , , , , , CCLSortString
            .AddColumn "C4", "Importe", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortString
            .AddColumn "C5", "Abonos", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortNumeric
        
        Case 9
            
            .AddColumn "C1", "Fecha Abono", ecgHdrTextALignCentre, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", "Fecha Venta", ecgHdrTextALignCentre, , 170, False, , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortNumeric
            .AddColumn "C3", "Folio Venta", ecgHdrTextALignCentre, , 100, , , , , , , CCLSortDate
            .AddColumn "C4", "Cliente", ecgHdrTextALignLeft, , 335, , , , , , , CCLSortString
            .AddColumn "C5", "Abono", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortString
        
        Case 10, 11, 12
            
            .AddColumn "C1", "Fecha", ecgHdrTextALignCentre, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", "Folio", ecgHdrTextALignCentre, , 90, , , , , , , CCLSortNumeric
            .AddColumn "C3", "Concepto", ecgHdrTextALignLeft, , 340, , , , , , , CCLSortString
            .AddColumn "C4", "Importe", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortNumeric
            .AddColumn "C5", "Iniciales", ecgHdrTextALignLeft, , 50, False, , , , , , CCLSortString
        
        Case 13, 14
            
            .AddColumn "C1", "Fecha", ecgHdrTextALignCentre, , 170, , , , , "DD/MM/YY HH:MM:SS AM/PM", , CCLSortDate
            .AddColumn "C2", IIf(Opcion = 13, "Folio", "Contrato"), ecgHdrTextALignCentre, , 90, , , , , , , CCLSortNumeric
            .AddColumn "C3", "Cliente", ecgHdrTextALignLeft, , 340, , , , , , , CCLSortString
            .AddColumn "C4", "Importe", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortNumeric

        End Select
        
        .Redraw = True
    End With
    
End Sub

Sub Cargar_Datos(strMovimiento As String, Opcion As Integer)
Dim rcConsulta As New ADODB.Recordset
Dim strSql As String

    With grdCancelaciones
        
        .Redraw = False
        
        Select Case Opcion
        Case 1, 2, 3
            
            strSql = "SELECT empeno.ID,empeno.Movimiento,empeno.Cancelado,empeno.Fecha,empeno.NumContrato,empeno.Folio,empeno.FolioOrigen,empeno.Serie,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente,clientes.Iniciales,empeno.Prestamo,empeno.Pago FROM empeno INNER JOIN clientes ON empeno.IDCliente=clientes.ID WHERE DATE_FORMAT(" & IIf(Opcion = OD_EMPENO Or Opcion = OD_REFRENDO, "empeno.Fecha", "empeno.FechaMovimiento") & ",'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' AND " & strMovimiento & " ORDER BY " & IIf(Opcion = OD_EMPENO Or Opcion = OD_REFRENDO, "empeno.Fecha", "empeno.FechaMovimiento") & ",empeno.NumContrato,empeno.Folio"
        
        Case 4
            
            strSql = "SELECT divisas.ID,divisas.Cancelado,divisas.Fecha,divisas.Folio,divisas.Cantidad,divisas.Importe,divisas.Tipo,monedas.Descripcion AS Desc_Divisa FROM divisas INNER JOIN monedas ON divisas.IDDivisa=monedas.Clave WHERE divisas.TipoEntrada=0 AND DATE_FORMAT(divisas.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY divisas.Fecha,divisas.Folio"
        
        Case 5
            
            strSql = "SELECT gastos.ID,gastos.Cancelado,gastos.Fecha,gastos.Folio,gastos.Concepto,gastos.Importe,cuentasgastos.Descripcion AS Cuenta FROM gastos INNER JOIN cuentasgastos ON gastos.CuentaGastos=cuentasgastos.ID WHERE DATE_FORMAT(gastos.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY gastos.Fecha,gastos.Folio"
        
        Case 6
            
            strSql = "SELECT pagosfijos.ID,pagosfijos.IDEmpeno,pagosfijos.Movimiento,pagosfijos.FechaMovimiento,pagosfijos.Cancelado,pagosfijos.NumPago,empeno.NumContrato,empeno.Folio,empeno.Fecha,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente,clientes.Iniciales,pagosfijos.Pago,pagosfijos.NumPago FROM pagosfijos INNER JOIN empeno ON pagosfijos.IDEmpeno=empeno.ID INNER JOIN clientes ON empeno.IDCliente=clientes.ID WHERE DATE_FORMAT(pagosfijos.FechaMovimiento,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY empeno.NumContrato,pagosfijos.NumPago,pagosfijos.FechaMovimiento"
        
        Case 7
            
            strSql = "SELECT ventas.ID,ventas.Fecha,ventas.Cancelado,ventas.Folio,((ventas.Total-(ventas.Total*(ventas.Descuento/100)))*(1+(ventas.IVA/100))) AS Total,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente FROM ventas LEFT JOIN clientes ON ventas.IDCliente=clientes.ID WHERE ventas.TipoVenta=0 AND DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' AND ventas.Apartado=0 AND ventas.TipoVenta=" & VENTAMOSTRADOR & " ORDER BY ventas.Fecha,ventas.Folio"
        
        Case 8
            
            strSql = "SELECT ventas.ID,ventas.Fecha,ventas.Cancelado,ventas.Folio,((ventas.Total-(ventas.Total*(ventas.Descuento/100)))*(1+(ventas.IVA/100))) AS Total,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente FROM ventas LEFT JOIN clientes ON ventas.IDCliente=clientes.ID WHERE DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' AND ventas.Apartado=1 ORDER BY ventas.Fecha,ventas.Folio"
            
        Case 9
            
            strSql = "SELECT abonos.ID,abonos.Cancelado,abonos.Fecha AS FechaAbono,abonos.Importe,abonos.Movimiento,abonos.IDVenta,ventas.Fecha AS FechaVenta,ventas.Folio,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente FROM ventas LEFT JOIN abonos ON (ventas.ID=abonos.IDVenta AND ventas.Fecha<>abonos.Fecha) LEFT JOIN clientes ON ventas.IDCliente=clientes.ID WHERE ventas.Cancelado=0 AND DATE_FORMAT(abonos.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY ventas.Fecha,ventas.Folio,abonos.Fecha"
        
        Case 10
            
            strSql = "SELECT boveda.ID,boveda.Fecha,boveda.Folio,boveda.Cancelado,boveda.Deposito,boveda.Concepto,boveda.Importe FROM boveda WHERE DATE_FORMAT(boveda.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY boveda.Fecha,boveda.Folio"
            
        Case 11
            
            strSql = "SELECT bancos.ID,bancos.Fecha,bancos.Folio,bancos.Cancelado,bancos.Deposito,bancos.Concepto,bancos.Importe FROM bancos WHERE TipoMov=0 AND DATE_FORMAT(bancos.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY bancos.Fecha,bancos.Folio"
        
        Case 12
            
            strSql = "SELECT bancos.ID,bancos.Fecha,bancos.Folio,bancos.Cancelado,bancos.Deposito,bancos.Concepto,bancos.Importe FROM bancos WHERE TipoMov=1 AND DATE_FORMAT(bancos.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY bancos.Fecha,bancos.Folio"
        
        Case 13
            
            strSql = "SELECT co.ID,co.Fecha,co.Folio,co.Cancelado,(co.Total*(1+(co.IVA/100))) AS crTotal,CONCAT(cli.Nombre,' ',cli.Apellido) AS Cliente FROM compras co LEFT JOIN clientes cli ON co.IDCliente=cli.ID WHERE DATE_FORMAT(co.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY co.Fecha,co.Folio"
        
        Case 14
            
            strSql = "SELECT ventas.ID,ventas.Fecha,ventas.Cancelado,ventas.Folio,(SUM(dv.Costo)+SUM(dv.Intereses)+SUM(dv.Almacenaje)+SUM(dv.Seguro)+SUM(dv.Moratorios)+SUM(dv.ImporteIva)) AS Total,CONCAT(clientes.Nombre,' ',clientes.Apellido) AS Cliente FROM ventas INNER JOIN detallesventas dv ON ventas.ID=dv.IDVenta LEFT JOIN clientes ON ventas.IDCliente=clientes.ID WHERE ventas.TipoVenta=" & VENTACLIENTE & " AND DATE_FORMAT(ventas.Fecha,'%Y%/%m%/%d')='" & Format(Date, "YYYY/MM/DD") & "' AND ventas.Apartado=0 GROUP BY ventas.ID ORDER BY ventas.Fecha,ventas.Folio"
        End Select
        
        rcConsulta.Open strSql, dbDatos, adOpenForwardOnly, adLockOptimistic
                              
        While Not rcConsulta.EOF
        
            .AddRow
            Select Case Opcion
            Case 1, 2, 3
                
                .CellText(.Rows, 1) = rcConsulta!Fecha
                .CellItemData(.Rows, 1) = rcConsulta!ID
                .CellText(.Rows, 2) = rcConsulta!NumContrato
                .CellItemData(.Rows, 2) = rcConsulta!Folio
                .CellTextAlign(.Rows, 2) = DT_CENTER
                .CellText(.Rows, 3) = rcConsulta!Cliente
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                .CellText(.Rows, 4) = IIf(Opcion = D_DESEMPEÑO, rcConsulta!Pago, rcConsulta!Prestamo)
                .CellItemData(.Rows, 4) = rcConsulta!Movimiento
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                .CellText(.Rows, 5) = rcConsulta!Iniciales
                .CellItemData(.Rows, 5) = rcConsulta!FolioOrigen
                .CellText(.Rows, 6) = rcConsulta!Serie
                
            Case 4
                
                .CellText(.Rows, 1) = rcConsulta!Fecha
                .CellItemData(.Rows, 1) = rcConsulta!ID
                .CellText(.Rows, 2) = rcConsulta!Folio
                .CellTextAlign(.Rows, 2) = DT_CENTER
                .CellText(.Rows, 3) = rcConsulta!Desc_Divisa
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                .CellText(.Rows, 4) = rcConsulta!Importe
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                .CellText(.Rows, 5) = rcConsulta!Cantidad
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                .CellText(.Rows, 6) = IIf(rcConsulta!Tipo = 0, "COMPRA", "VENTA")
                .CellTextAlign(.Rows, 6) = DT_CENTER
            
            Case 5
                
                .CellText(.Rows, 1) = rcConsulta!Fecha
                .CellItemData(.Rows, 1) = rcConsulta!ID
                .CellText(.Rows, 2) = rcConsulta!Folio
                .CellTextAlign(.Rows, 2) = DT_CENTER
                .CellText(.Rows, 3) = rcConsulta!Cuenta
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                .CellText(.Rows, 4) = rcConsulta!Concepto
                .CellTextAlign(.Rows, 4) = DT_LEFT
                .CellText(.Rows, 5) = rcConsulta!Importe
                .CellTextAlign(.Rows, 5) = DT_RIGHT
            
            Case 6
                
                .CellText(.Rows, 1) = rcConsulta!FechaMovimiento
                .CellItemData(.Rows, 1) = rcConsulta!ID
                .CellText(.Rows, 2) = rcConsulta!NumContrato
                .CellItemData(.Rows, 2) = rcConsulta!IDEmpeno
                .CellTextAlign(.Rows, 2) = DT_CENTER
                .CellText(.Rows, 3) = rcConsulta!Cliente
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                .CellText(.Rows, 4) = rcConsulta!NumPago & " de " & SacaValor("pagosfijos", "COUNT(ID)", " WHERE IDEmpeno=" & rcConsulta!IDEmpeno & " AND Cancelado=0")
                .CellItemData(.Rows, 4) = rcConsulta!Folio
                .CellTextAlign(.Rows, 4) = DT_CENTER
                .CellText(.Rows, 5) = rcConsulta!Pago
                .CellItemData(.Rows, 5) = rcConsulta!Movimiento
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                
                .CellText(.Rows, 6) = rcConsulta!NumPago
                .CellItemData(.Rows, 6) = 1
                .CellTextAlign(.Rows, 6) = DT_RIGHT
                
                .CellText(.Rows, 7) = rcConsulta!Fecha
                .CellTextAlign(.Rows, 7) = DT_RIGHT
            
            Case 7, 14
                
                .CellText(.Rows, 1) = rcConsulta!Fecha
                .CellItemData(.Rows, 1) = rcConsulta!ID
                
                .CellText(.Rows, 2) = rcConsulta!Folio
                .CellTextAlign(.Rows, 2) = DT_CENTER
                
                .CellText(.Rows, 3) = rcConsulta!Cliente
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                
                .CellText(.Rows, 4) = rcConsulta!Total
                .CellTextAlign(.Rows, 4) = DT_RIGHT
            
            Case 8
                
                .CellText(.Rows, 1) = rcConsulta!Fecha
                .CellItemData(.Rows, 1) = rcConsulta!ID
                
                .CellText(.Rows, 2) = rcConsulta!Folio
                .CellTextAlign(.Rows, 2) = DT_CENTER
                
                .CellText(.Rows, 3) = rcConsulta!Cliente
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                
                .CellText(.Rows, 4) = rcConsulta!Total
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                
                .CellText(.Rows, 5) = SacaValor("abonos", "SUM(Importe)", " WHERE IDVenta=" & rcConsulta!ID)
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                
            Case 9
                
                .CellText(.Rows, 1) = rcConsulta!FechaAbono
                .CellItemData(.Rows, 1) = rcConsulta!ID
                
                .CellText(.Rows, 2) = rcConsulta!FechaVenta
                .CellTextAlign(.Rows, 2) = DT_LEFT
                
                .CellText(.Rows, 3) = rcConsulta!Folio
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                .CellTextAlign(.Rows, 3) = DT_CENTER
                
                .CellText(.Rows, 4) = rcConsulta!Cliente
                .CellItemData(.Rows, 4) = rcConsulta!Movimiento
                .CellTextAlign(.Rows, 4) = DT_LEFT

                .CellText(.Rows, 5) = rcConsulta!Importe
                .CellItemData(.Rows, 5) = rcConsulta!IDVenta
                .CellTextAlign(.Rows, 5) = DT_RIGHT
            
            Case 10, 11, 12
                
                .CellText(.Rows, 1) = rcConsulta!Fecha
                .CellItemData(.Rows, 1) = rcConsulta!ID
                
                .CellText(.Rows, 2) = rcConsulta!Folio
                .CellItemData(.Rows, 2) = rcConsulta!Deposito
                .CellTextAlign(.Rows, 2) = DT_CENTER

                .CellText(.Rows, 3) = rcConsulta!Concepto
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                
                .CellText(.Rows, 4) = rcConsulta!Importe
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                
            Case 13
                    
                .CellText(.Rows, 1) = rcConsulta!Fecha
                .CellItemData(.Rows, 1) = rcConsulta!ID
                
                .CellText(.Rows, 2) = rcConsulta!Folio
                .CellTextAlign(.Rows, 2) = DT_CENTER

                .CellText(.Rows, 3) = IIf(IsNull(rcConsulta!Cliente), "PUBLICO EN GENERAL", rcConsulta!Cliente)
                .CellItemData(.Rows, 3) = rcConsulta!Cancelado
                
                .CellText(.Rows, 4) = rcConsulta!crTotal
                .CellTextAlign(.Rows, 4) = DT_RIGHT
            End Select
            
            'Sombreo el Grid
            If rcConsulta!Cancelado = 1 Then
                
                Colorea grdCancelaciones, grdCancelaciones.Rows, RGB(244, 119, 66)
            Else
                
                Colorea grdCancelaciones, grdCancelaciones.Rows, IIf(grdCancelaciones.Rows Mod 2 > 0, RGB(236, 252, 222), RGB(255, 255, 255))
            End If
        
        rcConsulta.MoveNext
        Wend
        rcConsulta.Close
        Set rcConsulta = Nothing
        
        .Redraw = True
    End With

End Sub

Sub Saca_Movimiento()
Dim strCampo As String, Movimiento As Integer
                
    strCampo = ""
    Movimiento = cmbMovimiento.ItemData(cmbMovimiento.ListIndex)
    Select Case Movimiento

    Case 1

        strCampo = "Origen=" & OD_EMPENO & " AND Destino=0"
        Movimiento = 1

    Case 3

        strCampo = "Destino=" & D_DESEMPEÑO
        Movimiento = 3

    Case 2

        strCampo = "Origen=" & OD_REFRENDO & " AND Destino=0"
        Movimiento = 2
        
    Case 15
        
        frmPasswords.DescuentoVentas = 0
        frmPasswords.PrecioVitrina = 0
        frmPasswords.Cancel = 0
        frmPasswords.Ventas = 0
        frmPasswords.HacerCorte = 0
        frmPasswords.ModificaPrecio = 0
        frmPasswords.InteresDesempeño = 0
        frmPasswords.InteresRefrendo = 0
        frmPasswords.RecalculoPrecios = 0
        frmPasswords.AutorizaPrestamo = 0
        frmPasswords.ModificaCorte = 0
        frmPasswords.CancelaCierre = 1
        
        If frmPasswords.Password(GERENTE, 1) Then
        
            frmCancelacionCierre.Cancelar
            Exit Sub
        Else
            
            Exit Sub
        End If
    End Select
    
    Crear_Encabezados Movimiento
    Cargar_Datos strCampo, Movimiento
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Sub CopiaContrato(IDEmpeno As Long)
Dim strSql1 As String, strSql2 As String, strSql3 As String
Dim rcEmpeño As New ADODB.Recordset

    With rcEmpeño
    
        'Leemos los datos del Empeno original
        .Open "SELECT empeno.* FROM empeno WHERE empeno.ID=" & IDEmpeno, dbDatos, adOpenForwardOnly, adLockOptimistic
            
        'Tomo los valores
        strSql1 = "INSERT INTO empeno (Cancelado,Fecha,Movimiento,NumContrato,Folio,Prestamo,Avaluo,Origen,Destino,Vencimiento,FolioOrigen,FechaMovimiento,Serie,Pagado,PC,IDCliente,Responsable,Valuador,Notas,Tasa,Almacenaje,Seguro,Operacion,Comision,IVA,Periodo,Venperiodo,VenAlmoneda,Tipointeres,TipoTasa,IDSucursal,IDUsuario,IDAutorizacion,NumBolsa,Ubicacion,Caja,Cajon,Fila,IDUsuarioAutoriza,TipoAutoriza,PrestamoInicial,Pago,Intereses,ImporteAlmacenaje,ImporteSeguro,ImporteMoratorios,ImportePerdida,ImporteIva,ImporteOtros,Captura,FolioNota,SaldoPuntosAnteriorEmp,PuntosAcumuladosEmp,SaldoPuntosActualEmp,IDTarjetaEmp,DescuentoXPuntos,SaldoPuntosAnterior,PuntosUsados,PuntosAcumulados,SaldoPuntosActual,IDTarjeta) VALUES "
        strSql2 = "(1,'" & Format(!Fecha, "YYYY/MM/DD HH:MM:SS") & "'," & !Movimiento & "," & !NumContrato & "," & !Folio & "," & !Prestamo & "," & !Avaluo & "," & !Origen & "," & !Destino & ",'" & Format(!Vencimiento, "YYYY/MM/DD") & "'," & !FolioOrigen & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & !Serie & "," & !Pagado & ",'" & Trim(!PC) & "'," & !IDCliente & ",'" & Trim(!Responsable) & "',"
        strSql3 = "'" & Trim(!Valuador) & "','" & Trim(!Notas) & "'," & !Tasa & "," & !Almacenaje & "," & !Seguro & "," & !Operacion & "," & !Comision & "," & !Iva & "," & !Periodo & "," & !VenPeriodo & "," & !VenAlmoneda & ",'" & Trim(!TipoInteres) & "','" & Trim(!TipoTasa) & "'," & !IDSucursal & "," & !IDUsuario & "," & !IDAutorizacion & ",'" & Trim(!NumBolsa) & "','" & Trim(!ubicacion) & "','" & Trim(!caja) & "','" & Trim(!Cajon) & "','" & Trim(!Fila) & "'," & !IDUsuarioAutoriza & "," & !TipoAutoriza & "," & !PrestamoInicial & "," & !Pago & "," & !Intereses & "," & !ImporteAlmacenaje & "," & !ImporteSeguro & "," & !ImporteMoratorios & "," & !ImportePerdida & "," & !ImporteIva & "," & !ImporteOtros & "," & !Captura & "," & !FolioNota & "," & _
            !SaldoPuntosAnteriorEmp & "," & !PuntosAcumuladosEmp & "," & !SaldoPuntosActualEmp & "," & !IDTarjetaEmp & "," & !DescuentoXPuntos & "," & !SaldoPuntosAnterior & "," & !PuntosUsados & "," & !PuntosAcumulados & "," & !SaldoPuntosActual & "," & !IDTarjeta & ")"
        
        'Cierro la conexión
        .Close
        Set rcEmpeño = Nothing
        
        'Grabo la copia del registro
        dbDatos.Execute strSql1 & strSql2 & strSql3
                 
    End With
    
End Sub

Function VerificaPagoFijo(IDEmpeno As Long, NumPago As Integer) As Boolean
Dim Renglon As Integer, Pago As Integer
    
    Pago = 0
    VerificaPagoFijo = True
    With grdCancelaciones
                    
        For Renglon = 1 To .Rows
            
            If .CellItemData(Renglon, 3) = 0 Then
                
                If .CellItemData(Renglon, 2) = IDEmpeno Then
                    
                    If NumPago < Val(.CellText(Renglon, 6)) Then
                    
                        VerificaPagoFijo = False
                        MsgBox "Es necesario que cancele los pagos posteriores al seleccionado !!", vbInformation, "Cancelaciones"
                        Exit For
                    End If
                    
                    '''''Pago = Pago + .CellItemData(Renglon, 6)
                End If
                
            End If
            
        Next Renglon
    
    End With
    
'''''    If NumPago = Pago Then VerificaPagoFijo = True Else MsgBox "Es necesario que cancele los pagos posteriores al seleccionado !!", vbInformation, "Cancelaciones"
End Function

Function UltimaFecha(IDEmpeno As Long, NumPago As Integer) As Date
Dim rcAux As New ADODB.Recordset

    rcAux.Open "SELECT MAX(Vencimiento) AS VencimientoPago FROM pagosfijos WHERE IDEmpeno=" & IDEmpeno & " AND Pagado=1 AND Cancelado=0 AND NumPago<=" & NumPago, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcAux.BOF And Not rcAux.EOF And Not IsNull(rcAux!VencimientoPago) Then
        
        UltimaFecha = rcAux!VencimientoPago
    Else
        
        UltimaFecha = SacaValor("empeno", "Fecha", " WHERE ID=" & IDEmpeno)
    End If
    rcAux.Close
    Set rcAux = Nothing
End Function
