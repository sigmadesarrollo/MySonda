VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCancelacionesVentas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cancelación de ventas"
   ClientHeight    =   3495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCancelacionesVentas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3495
   ScaleWidth      =   5850
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la venta"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   5655
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "<Total>"
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
         Left            =   1080
         TabIndex        =   11
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label lblFolio 
         AutoSize        =   -1  'True
         Caption         =   "<Folio>"
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
         Left            =   3240
         TabIndex        =   10
         Top             =   960
         Width           =   870
      End
      Begin VB.Label lblFecha 
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
         Left            =   1080
         TabIndex        =   9
         Top             =   960
         Width           =   960
      End
      Begin VB.Label lblCliente 
         AutoSize        =   -1  'True
         Caption         =   "<Cliente>"
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
         Left            =   1200
         TabIndex        =   8
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   810
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Total:"
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
         Left            =   240
         TabIndex        =   6
         Top             =   1440
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Folio:"
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
         Left            =   2520
         TabIndex        =   5
         Top             =   960
         Width           =   600
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   690
      End
   End
   Begin VB.TextBox txtNombrePago 
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
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4125
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosClave 
      Height          =   285
      Left            =   4290
      TabIndex        =   2
      Top             =   495
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
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
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4635
      TabIndex        =   12
      Top             =   3000
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
      Picture         =   "frmCancelacionesVentas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3435
      TabIndex        =   13
      Top             =   3000
      Width           =   1095
      _ExtentX        =   1931
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
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmCancelacionesVentas.frx":009D
      PictureDisabled =   "frmCancelacionesVentas.frx":0113
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Venta a cancelar:"
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
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmCancelacionesVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim Importe As Double
Dim Descuento As Double

Private Sub cmdAceptar_Click()
    Cancelar_Venta
End Sub

Private Sub Cancelar_Venta()
Dim rcConsulta As ADODB.Recordset
Dim Movimiento As Long

On Error GoTo error

If Val(txtNombrePago.Tag) = 0 Then Exit Sub
If MsgBox("Esta seguro de cancelar la venta ??", vbYesNo + vbQuestion, "Cancelación de ventas") = vbYes Then
   
    'Cancelamos la venta
    dbDatos.Execute "UPDATE ventas SET Cancelado=1,OrigenCancelacion=1,FechaMovimiento='" & Format(Date, "YYYY/MM/DD") & "' WHERE ID=" & txtNombrePago.Tag
    
    'Regresamos los articulos
    Set rcConsulta = dbDatos.Execute("select idarticulo from detallesventas where idventa=" & txtNombrePago.Tag)
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        rcConsulta.MoveFirst
        While Not rcConsulta.EOF
                dbDatos.Execute "update detallesentradainventario set cantidad=cantidad+1 where id=" & rcConsulta!IDArticulo
        rcConsulta.MoveNext
        Wend
    End If
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Grabamos el cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
              "('" & Format(Date, "YYYY/MM/DD") & "','Devolucion'," & Movimiento & "," & Val(lblFolio.Caption) & ",'DV01','620301'," & Importe & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabamos el abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
              "('" & Format(Date, "YYYY/MM/DD") & "','Devolucion'," & Movimiento & "," & Val(lblFolio.Caption) & ",'DV50','110150'," & Importe & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    'Grabamos abono 199450
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
              "('" & Format(Date, "YYYY/MM/DD") & "','Devolucion'," & Movimiento & "," & Val(lblFolio.Caption) & ",'DV50','199450'," & Importe & "," & TIPO_ABONO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
    If Descuento > 0 Then
        'Grabamos el cargo de Descuento
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "'," & Movimiento & "," & Val(lblFolio.Caption) & ",'DC06','620650'," & Descuento & "," & TIPO_ABONO & ",0,'Devolucion descuento','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    End If
    
    MsgBox "Venta cancelada con éxito !!", vbInformation, "Cancelación de ventas"
    Limpiar
End If
   
error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Private Sub cmdMosClave_Click()
    frmMostrarClientesVentas.Ver Me, txtNombrePago, False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Poner_Flat Fl, Me.Controls, Me
    Limpiar
    CentrarForm Me, frmMDI
End Sub

Private Sub Limpiar()
    lblCliente.Caption = ""
    lblFecha.Caption = ""
    lblFolio.Caption = ""
    lblTotal.Caption = ""
    txtNombrePago.Text = ""
    txtNombrePago.Tag = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtNombrePago_GotFocus()
    Seleccionar_Texto txtNombrePago
    Cambiar_Color True, txtNombrePago
End Sub

Private Sub txtNombrePago_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNombrePago_LostFocus()
    Cambiar_Color False, txtNombrePago
End Sub

'Buscamos el cliente para los pagos
Public Sub Buscar_Cliente(ID As Long)
Dim rcCliente As New ADODB.Recordset
Dim rcAbono As New ADODB.Recordset
   
    On Error GoTo error
    
    rcCliente.Open "SELECT ventas.*,concat(clientes.Nombre,' ',clientes.Apellido) as Cliente FROM ventas Left Join clientes on ventas.IDCliente=clientes.ID WHERE ventas.ID=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
   
    rcAbono.Open "SELECT Sum(Importe) AS Total FROM abonos WHERE IDVenta=" & ID, dbDatos, adOpenDynamic, adLockOptimistic
   
    With rcCliente
        txtNombrePago.Tag = !ID
        lblCliente.Caption = IIf(IsNull(!cliente), "", !cliente)
        lblFecha.Caption = Format(!Fecha, "DD/MM/YY")
        lblFolio.Caption = !Folio
        lblTotal.Caption = Format(Format(IIf(IsNull(rcAbono!Total), (!Total - !Descuento) + (((!Total - !Descuento) * !Iva) / 100), rcAbono!Total), "Currency"), "###,###,###,##0.00")
        Importe = lblTotal.Caption
        Descuento = !Descuento
    End With
   
    rcCliente.Close
    rcAbono.Close
   
error:
    Maneja_Error Err
    Set rcCliente = Nothing
    Set rcAbono = Nothing
End Sub
