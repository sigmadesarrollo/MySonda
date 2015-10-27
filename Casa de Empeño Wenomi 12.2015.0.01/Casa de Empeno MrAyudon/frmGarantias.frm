VERSION 5.00
Object = "{5F37140E-C836-11D2-BEF8-525400DFB47A}#1.1#0"; "vbalTab6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{1781610F-46E8-4DD3-922D-8DEF1A9DA567}#28.0#0"; "Credencial.ocx"
Begin VB.Form frmGarantias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Devoluciones"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6270
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGarantias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4830
   ScaleWidth      =   6270
   Begin VB.Frame Frame2 
      Caption         =   "Folio"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   4080
      TabIndex        =   6
      Top             =   0
      Width           =   1575
      Begin VB.Label lblFolio 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "<Folio>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.TextBox txtFolioFactura 
      Alignment       =   1  'Right Justify
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
      Left            =   1920
      MaxLength       =   14
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   4320
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
      Picture         =   "frmGarantias.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3840
      TabIndex        =   12
      Top             =   4320
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
      Picture         =   "frmGarantias.frx":055E
   End
   Begin vbalTabStrip6.TabControl tTab 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   960
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   5530
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CoolTabs        =   1
      Begin Credencial.usCredencial cDatosVenta 
         Height          =   2535
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4471
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AlingHeader     =   8
         AlingBody       =   0
         BodyIndent      =   10
         HeaderIndent    =   5
         HeaderText      =   "Datos de Venta"
         HeaderBackColor =   16766131
         HeightHeader    =   25
         SidePicture     =   -1  'True
         SideBackColor   =   15000804
         WidthSide       =   39
         SidePicture     =   -1  'True
         HeaderBorderBackColor=   13603685
         BackColor       =   -2147483643
      End
      Begin Credencial.usCredencial cArticulos 
         Height          =   2535
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   4471
         BeginProperty HeaderFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty BodyFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         AlingHeader     =   262144
         AlingBody       =   0
         BodyIndent      =   10
         HeaderIndent    =   5
         HeaderText      =   "Articulos"
         HeaderBackColor =   16766131
         HeightHeader    =   25
         SidePicture     =   -1  'True
         SideBackColor   =   15000804
         WidthSide       =   39
         SidePicture     =   -1  'True
         HeaderBorderBackColor=   13603685
         BackColor       =   -2147483643
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2535
         Left            =   240
         ScaleHeight     =   2535
         ScaleWidth      =   5655
         TabIndex        =   11
         Top             =   480
         Width           =   5655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo Movimiento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   4335
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   3015
         TabIndex        =   3
         Top             =   360
         Width           =   3015
         Begin VB.OptionButton opDevolucionGarantia 
            Appearance      =   0  'Flat
            Caption         =   "Devolucion Garantia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   14
            Top             =   360
            Width           =   2895
         End
         Begin VB.OptionButton opDevolucion 
            Appearance      =   0  'Flat
            Caption         =   "Devolucion Venta"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   5
            Top             =   720
            Value           =   -1  'True
            Width           =   2655
         End
         Begin VB.OptionButton opGarantia 
            Appearance      =   0  'Flat
            Caption         =   "Garantia"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   1575
         End
      End
   End
   Begin VB.Label lblLeyenda 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "Folio Factura:"
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
      TabIndex        =   0
      Top             =   360
      Width           =   1665
   End
End
Attribute VB_Name = "frmGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Fl() As New cFlatControl
Dim m_Cliente As String
Dim m_Articulo As String

Dim m_Empleado As String
Dim m_Codigo As String
Dim m_FechaVenta As Date

Private Sub cmdAceptar_Click()
   If Validar Then
      'If opGarantia.Value Or opDevolucion.Value Then
          Grabar_Datos Val(txtFolioFactura.Tag)
      'ElseIf opDevolucionGarantia.Value Then
      '    Grabar_Devolucion_Garantia Val(txtFolioFactura.Tag)
      'End If
   End If
End Sub

Private Function Validar() As Boolean
   Validar = True
   
   If Val(txtFolioFactura.Tag) = 0 Then
      MsgBox "Favor de seleccionar el articulo", vbOKOnly Or vbCritical
      Validar = False
      Exit Function
   End If
   
End Function

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    CentrarForm Me, frmMDI
    Crear_Pestanas
    Poner_Flat Fl, Me.Controls, Me
    cDatosVenta.AlingHeader = DT_WORD_ELLIPSIS Or DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    cDatosVenta.AlingBody = DT_WORD_ELLIPSIS Or DT_LEFT Or DT_WORDBREAK
    cArticulos.AlingHeader = DT_WORD_ELLIPSIS Or DT_LEFT Or DT_VCENTER Or DT_SINGLELINE
    cArticulos.AlingBody = DT_WORD_ELLIPSIS Or DT_LEFT Or DT_WORDBREAK
    lblFolio.Caption = GetFolio(False)
    
    opDevolucion_Click
    Screen.MousePointer = vbDefault
End Sub

Private Sub Crear_Pestanas()

    With tTab
        .AddTab "Cliente"
        .AddTab "Articulos"
    End With

End Sub

Private Sub opDevolucion_Click()
   Limpiar
    lblLeyenda.Caption = "Folio Factura:"
    lblFolio.Caption = GetFolio(False)
    cDatosVenta.Clear
    cArticulos.Clear
End Sub

Private Sub opDevolucionGarantia_Click()
   Limpiar
    lblLeyenda.Caption = "Folio Garantia:"
    lblFolio.Caption = GetFolio(False)
    cDatosVenta.Clear
    cArticulos.Clear
End Sub

Private Sub opGarantia_Click()
   Limpiar
    lblLeyenda.Caption = "Folio Factura:"
    lblFolio.Caption = GetFolio(False)
    cDatosVenta.Clear
    cArticulos.Clear
End Sub

Private Sub tTab_TabClick(ByVal lTab As Long)
    Select Case lTab
    
        Case 1
            cArticulos.Visible = False
            cDatosVenta.Visible = True
            
        Case 2
            cDatosVenta.Visible = False
            cArticulos.Visible = True
            
    End Select
End Sub

Private Sub txtFolioFactura_Change()
    txtFolioFactura.Tag = ""
End Sub

Private Sub txtFolioFactura_GotFocus()
    Seleccionar_Texto txtFolioFactura
    Cambiar_Color True, txtFolioFactura
End Sub

Private Sub txtFolioFactura_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    If KeyAscii = vbKeyReturn Then
        If opGarantia.Value Or opDevolucion.Value Then
            Buscar_Venta Val(txtFolioFactura.text)
        ElseIf opDevolucionGarantia.Value Then
            Buscar_Garantia Val(txtFolioFactura.text)
        End If
    End If
    'Pasar_Foco KeyAscii
End Sub

Private Sub txtFolioFactura_LostFocus()
    Cambiar_Color False, txtFolioFactura
End Sub

Private Function GetFolio(Grabar As Boolean) As Long
    If opGarantia.Value Then
        GetFolio = Regresa_Movimiento(Grabar, "FolioGarantia")
    ElseIf opDevolucion.Value Then
        GetFolio = Regresa_Movimiento(Grabar, "FolioDevolucion")
    ElseIf opDevolucionGarantia.Value Then
        GetFolio = Regresa_Movimiento(Grabar, "FolioDevolucionGarantia")
    End If
End Function

Private Sub Buscar_Garantia(Folio As Long)
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim rcDetalles As New ADODB.Recordset
    Dim Sql As String
    
    Sql = "SELECT GarantiasDevoluciones.ID,GarantiasDevoluciones.Fecha AS FechaGarantia,GarantiasDevoluciones.IDUsuario,GarantiasDevoluciones.IDVenta, " & _
          "Ventas.IDCliente,Ventas.Fecha AS FechaVenta,Ventas.Total,Ventas.Descuento  " & _
          "FROM GarantiasDevoluciones " & _
          "INNER JOIN Ventas On Ventas.ID=GarantiasDevoluciones.IDVenta " & _
          "WHERE GarantiasDevoluciones.Folio=" & Folio & " AND Garantia=1 AND Entregado=0"
    
    rc.Open Sql, dbDatos, adOpenDynamic, adLockOptimistic
    
    If Not rc.EOF Then
    
        With cDatosVenta
            .Clear
            txtFolioFactura.Tag = rc!ID
            m_Empleado = SacaValor("Usuarios", "Usuario", " WHERE ID=" & rc!IDUsuario)
            m_Cliente = SacaValor("Clientes", "CONCAT(Nombre,' ',Apellido)", " WHERE ID=" & rc!IDCliente)
            
            .Add ""
            .Add "<bold>Fecha Garantia: " & Format(rc!FechaGarantia, "DD/MMM/YY HH:MM:SS am/pm") & "</bold>"
            .Add "<bold>Usuario: " & m_Empleado & "</bold>"
            
            .Add "<bold>" & m_Cliente & "</bold>"
            .Add "<bold>No. Cedula: " & SacaValor("Clientes", "NumeroIdentificacion", " WHERE ID=" & rc!IDCliente) & "</bold>"
            .Add "Fecha de Compra: " & Format(rc!FechaVenta, "DD/MM/YYYY HH:MM:SS am/pm") & vbCrLf & _
                     "Total Venta: " & Format(rc!Total, FMoneda) & vbCrLf & _
                     "Descuento: " & Format(rc!Descuento, FMoneda)
            
        
        End With
        
        rcDetalles.Open "SELECT dv.Articulo,dv.Codigo,dv.Peso,dv.Precio,di.Modelo,di.Serie,di.Marca,dv.Kilates FROM DetallesVentas dv INNER JOIN detallesentradainventario di ON di.ID=dv.IDArticulo WHERE IDVenta=" & rc!IDVenta, dbDatos, adOpenDynamic, adLockOptimistic
        
        cArticulos.Clear
         While Not rcDetalles.EOF
            cArticulos.Add ""
            cArticulos.Add "<bold>" & rcDetalles!Articulo & "</bold>"
            cArticulos.Add "<bold>" & rcDetalles!Codigo & "</bold>"
            m_Articulo = rcDetalles!Articulo
            m_Codigo = rcDetalles!Codigo
            If rcDetalles!Peso > 0 Then
                cArticulos.Add "Kilates: " & SacaValor("Kilatajes", "Descripcion", " WHERE ID=" & rcDetalles!Kilates) & "   " & _
                               "Peso: " & rcDetalles!Peso & "Grms" & "   " & "Precio: " & Format(rcDetalles!Precio, FMoneda)
            Else
               cArticulos.Add "Serie: " & rcDetalles!Serie
               cArticulos.Add "Modelo: " & rcDetalles!Modelo
               cArticulos.Add "Marca: " & rcDetalles!Marca
            End If
            rcDetalles.MoveNext
        Wend
        
        rcDetalles.Close
        Set rcDetalles = Nothing
        
    
    Else
        MsgBox "El folio de garantia no se encuentra", vbOKOnly Or vbCritical
        Limpiar
        txtFolioFactura.SetFocus
    End If
    
    rc.Close
Error:
    Maneja_Error Err
    
    Set rc = Nothing
    Set rcDetalles = Nothing
End Sub

Private Sub Buscar_Venta(Folio As Long)
    On Error GoTo Error
    Dim rc As New ADODB.Recordset
    Dim rcDetalles As New ADODB.Recordset
    Dim Sql As String

    If Folio = 0 Then Exit Sub
    
    cDatosVenta.Clear
    cArticulos.Clear
    Sql = "SELECT v.*,di.Tipo " & _
          "FROM Ventas v " & _
          "INNER JOIN detallesventas dv ON dv.IDVenta=v.ID " & _
          "INNER JOIN detallesentradainventario di ON di.ID=dv.IDArticulo " & _
          "WHERE v.Folio=" & Folio & " AND v.TipoVenta=" & VENTAMOSTRADOR & " AND v.Cancelado=0"
          
     rc.Open Sql, dbDatos, adOpenDynamic, adLockOptimistic
     If Not rc.EOF Then
     
        If rc!Tipo <> Val(SacaValor("Tipo", "ID", " Where Descripcion = 'ORO'")) Then
          
'          If opDevolucion.Value Then
'              If rc!Estatus = StatusVentas.Vendido Then
'                  MsgBox "La Prenda debe de pasar por la garantia antes de realizar la devolucion", vbOKOnly Or vbCritical
'                  txtFolioFactura.SetFocus
'                  Exit Sub
'              End If
'          End If
      
           With cDatosVenta
              txtFolioFactura.Tag = rc!ID
              m_Empleado = SacaValor("Usuarios", "Usuario", " WHERE ID=" & rc!IDUsuario)
              m_Cliente = SacaValor("Clientes", "CONCAT(Nombre,' ',Apellido)", " WHERE ID=" & rc!IDCliente)
              .Clear
              .Add ""
              .Add "<bold>" & m_Cliente & "</bold>"
              .Add "<bold>No. Cedula: " & SacaValor("Clientes", "NumeroIdentificacion", " WHERE ID=" & rc!IDCliente) & "</bold>"
              .Add "Fecha de Compra: " & Format(rc!Fecha, "DD/MM/YYYY HH:MM:SS am/pm") & vbCrLf & _
                       "Total Venta: " & Format(rc!Total, FMoneda) & vbCrLf & _
                       "Descuento: " & Format(rc!Descuento, FMoneda)
              m_FechaVenta = rc!Fecha
           End With
                  
          
          rcDetalles.Open "SELECT dv.Articulo,dv.Codigo,dv.Peso,dv.Precio,di.Modelo,di.Serie,di.Marca,di.Kilates FROM DetallesVentas dv INNER JOIN detallesentradainventario di ON di.ID=dv.IDArticulo WHERE IDVenta=" & rc!ID, dbDatos, adOpenDynamic, adLockOptimistic
          
          cArticulos.Clear
           While Not rcDetalles.EOF
              cArticulos.Add ""
              cArticulos.Add "<bold>" & rcDetalles!Articulo & "</bold>"
              cArticulos.Add "<bold>" & rcDetalles!Codigo & "</bold>"
              m_Articulo = rcDetalles!Articulo
              m_Codigo = rcDetalles!Codigo
              If rcDetalles!Peso > 0 Then
                  cArticulos.Add "Kilates: " & SacaValor("Kilatajes", "Descripcion", " WHERE ID=" & rcDetalles!Kilates) & "   " & _
                                 "Peso: " & rcDetalles!Peso & "Grms" & "   " & "Precio: " & Format(rcDetalles!Precio, FMoneda)
              Else
                 cArticulos.Add "Serie: " & rcDetalles!Serie
                 cArticulos.Add "Modelo: " & rcDetalles!Modelo
                 cArticulos.Add "Marca: " & rcDetalles!Marca
              End If
              rcDetalles.MoveNext
          Wend
          
          rcDetalles.Close
          Set rcDetalles = Nothing
        Else
            MsgBox "No se puede realizar la " & IIf(opGarantia.Value, "garantia", "devolucion") & " de una prenda de oro", vbOKOnly Or vbCritical
            Limpiar
            txtFolioFactura.Tag = ""
            txtFolioFactura.SetFocus
        End If
   Else
       MsgBox "El numero de factura de venta no se encuentra", vbCritical Or vbOKOnly
       Limpiar
       txtFolioFactura.Tag = ""
       txtFolioFactura.SetFocus
   End If
    
    rc.Close
    
Error:
    Maneja_Error Err

End Sub

'''''Private Sub Grabar_Devolucion_Garantia(IDGarantia As Long)
'''''    On Error GoTo Error
'''''    Dim Comentarios As String
'''''    Dim Folio As Long
'''''    Dim Movimiento As Long
'''''    Dim crCosto As Currency
'''''    Dim IDArticulo As Long
'''''    Dim IDVenta As Long
'''''    Dim FolioPase As Long
'''''
'''''    If IDGarantia = 0 Then
'''''        MsgBox "Favor de seleccionar el folio de garantia", vbOKOnly Or vbInformation
'''''        Exit Sub
'''''    End If
'''''
'''''    Comentarios = frmMotivoCancela.Mostrar()
'''''    If Trim(Comentarios) <> "" Then
'''''
'''''        Folio = GetFolio(False)
'''''        GetFolio True
'''''        dbDatos.Execute "UPDATE GarantiasDevoluciones SET Entregado=1,FolioEntregado=" & Folio & ", FechaEntregado='" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "',IDUsuarioEntregado=" & frmMDI.IDUsuario & ",ComentariosEntregado='" & Comentarios & "' WHERE ID=" & IDGarantia
'''''
'''''
'''''        crCosto = SacaValor("GarantiasDevoluciones g", "dv.costo", "INNER JOIN detallesventas dv ON dv.IDVenta=g.IDVenta WHERE g.ID=" & IDGarantia)
'''''
'''''           'saco el movimiento
'''''            Movimiento = Regresa_Movimiento(False)
'''''            Regresa_Movimiento True
'''''
'''''            'grabamos los cargos
'''''            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,FechaModificacion) VALUES " & _
'''''                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Now, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'GA50','620355'," & ConvMoneda(crCosto) & "," & _
'''''                            TIPO_ABONO & ",0,'Devolucion Garantia Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "')"
'''''
'''''
'''''        IDArticulo = Val(SacaValor("garantiasdevoluciones g", "d.ID", "INNER JOIN detallesventas dv ON dv.IDVenta=g.IDVenta INNER JOIN detallesentradainventario d ON d.ID=dv.IDArticulo WHERE g.ID=" & IDGarantia))
'''''        IDVenta = Val(SacaValor("garantiasdevoluciones", "IDVenta", "WHERE ID=" & IDGarantia))
'''''
'''''        'actualizamos la garantia
'''''        dbDatos.Execute "UPDATE detallesentradainventario SET Garantias=1 WHERE ID=" & IDArticulo
'''''
'''''        'actualizamos el status en ventas
'''''        dbDatos.Execute "UPDATE ventas SET Estatus=" & StatusVentas.Entregado & " WHERE ID=" & IDVenta
'''''
'''''        FolioPase = Regresa_Movimiento(False, "FolioPasesInventario")
'''''        Regresa_Movimiento True, "FolioPasesInventario"
'''''
'''''        'grabamos el pase de inventario
'''''        dbDatos.Execute "INSERT INTO pasesinventarios (Fecha,Folio,IDArticulo,Origen,Destino,Motivo,IDUsuario) VALUES ('" & _
'''''                        Format(Now, "YYYY/MM/dd HH:MM:SS") & "'," & FolioPase & "," & IDArticulo & "," & OrigenInventarioVentas.Garantia & "," & _
'''''                        OrigenInventarioVentas.Vendido & ",'" & Comentarios & "'," & frmMDI.IDUsuario & ")"
'''''
'''''        'actualizamos el status del articulo
'''''        dbDatos.Execute "UPDATE detallesentradainventario SET TipoSalida=" & OrigenInventarioVentas.Vendido & ",Destino=" & OrigenInventarioVentas.Vendido & " WHERE ID=" & IDArticulo
'''''
'''''
'''''
'''''        Imprimir_Nota_Entrega_Garantia Folio, Comentarios
'''''
'''''        MsgBox "Articulo en garantia devuelto", vbOKOnly Or vbInformation
'''''
'''''        Limpiar
'''''    Else
'''''        MsgBox "Favor de poner los comentarios", vbOKOnly Or vbInformation
'''''    End If
'''''
'''''
'''''Error:
'''''    Maneja_Error Err
'''''
'''''
'''''End Sub

Private Sub Grabar_Datos(IDVenta As Long)
    On Error GoTo Error
    Dim Folio As Long
    Dim Comentarios As String
    Dim Movimiento As Long
    Dim crCosto As Currency
    Dim CuentaInventario As String
    Dim CuentaCosto As String
    Dim CuentaVenta As String
    Dim ID As Long
    Dim IDArticulo As Long
    Dim FolioPase As Long
    
    If IDVenta = 0 Then
        MsgBox "Favor de seleccionar la factura", vbOKOnly Or vbInformation
        txtFolioFactura.SetFocus
        Exit Sub
    End If
    
    Comentarios = frmMotivoCancela.Mostrar()
    If Trim(Comentarios) <> "" Then
    
        Folio = GetFolio(False)
        GetFolio True
        
        'crCosto = SacaValor("detallesventas", "Costo", "WHERE IDVenta=" & IDVenta)
        IDArticulo = SacaValor("detallesventas", "IDArticulo", "WHERE IDVenta=" & IDVenta)
        crCosto = SacaValor("detallesentradainventario", "Costo", "WHERE ID=" & IDArticulo)
        
        dbDatos.Execute "INSERT INTO GarantiasDevoluciones (Fecha,Folio,Garantia,IDVenta,IDUsuario,Comentarios,IDUsuarioEntregado) VALUES ('" & _
                        Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & IIf(opGarantia.Value, 1, 0) & "," & IDVenta & "," & frmMDI.IDUsuario & ",'" & Comentarios & "'," & frmMDI.IDUsuario & ")"
        'dbDatos.Execute "UPDATE Ventas SET Devolucion=1 WHERE ID=" & IDVenta
        
        
        'si es devolucion del articulo
''        If opDevolucion.Value Then
            
            'saco el movimiento
            Movimiento = Regresa_Movimiento(False)
            Regresa_Movimiento True
            
            'actualizamos la venta
            dbDatos.Execute "UPDATE Ventas SET Devolucion=1 WHERE ID=" & IDVenta
            'dbDatos.Execute "UPDATE Ventas SET Estatus=" & StatusVentas.devolucion & " WHERE ID=" & IDVenta
            dbDatos.Execute "UPDATE detallesventas SET Devolucion=1 WHERE IDVenta=" & IDVenta
            
            'crCosto = ConvMoneda(SacaValor("Ventas", "Total", " WHERE ID=" & IDVenta))
            
''            If Val(SacaValor("detallesventas dv", "d.Tipo", "INNER JOIN detallesentradainventario d ON d.ID=dv.IDArticulo WHERE dv.IDVenta = " & IDVenta)) = frmMDI.IDPrendaMisc Then
                CuentaInventario = "620301"  'Misc
                CuentaCosto = "620250"
                CuentaVenta = "620750"
''            Else
''                CuentaInventario = "620302"  'Oro
''                CuentaCosto = "620252"
''                CuentaVenta = "620752"
''            End If
                        
            'cargo cuenta de devoluciones ventas
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'DEV01','410101'," & ConvMoneda(SacaValor("Ventas", "Total+(Total*(Iva/100))", " WHERE ID=" & IDVenta)) & "," & _
                            TIPO_CARGO & ",0,'Devolucion Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                            
                            
'            'cargo a la cuenta de existencia de inventario devoluciones
'            ya no se va mas al inventario de devoluciones
'            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,FechaModificacion) VALUES " & _
'                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'DEV01','620406'," & ConvMoneda(SacaValor("Ventas", "Total", " WHERE ID=" & IDVenta)) & "," & _
'                            TIPO_CARGO & ",0,'Devolucion Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "')"
                                        
            'grabamos el abono a caja
            'la salida de dinero de caja
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'DEV50','110150'," & ConvMoneda(SacaValor("Ventas", "Total+(Total*(Iva/100))", " WHERE ID=" & IDVenta)) & "," & _
                    TIPO_ABONO & ",0,'Devolucion Factura - " & SacaValor("Ventas", "Folio", " WHERE ID=" & IDVenta) & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
                   
            FolioPase = Regresa_Movimiento(False, "FolioPasesInventario")
            Regresa_Movimiento True, "FolioPasesInventario"
            
            'grabamos el pase de inventario
            dbDatos.Execute "INSERT INTO pasesinventarios (Fecha,Folio,IDArticulo,Origen,Destino,Motivo,IDUsuario) VALUES ('" & _
                            Format(Now, "YYYY/MM/dd HH:MM:SS") & "'," & FolioPase & "," & IDArticulo & ",1,1" & _
                            ",'" & Comentarios & "'," & frmMDI.IDUsuario & ")" '" & OrigenInventarioVentas.Vendido & OrigenInventarioVentas.Vitrina &
                            
            'actualizamos el status del articulo
            dbDatos.Execute "UPDATE detallesentradainventario SET Cantidad=1,TipoSalida=0 WHERE ID=" & IDArticulo
                    
'            'abono a costos
'            'regresamos los costos
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'DEV50','" & CuentaCosto & "'," & ConvMoneda(crCosto) & "," & _
                    TIPO_ABONO & ",0,'Devolucion Factura - " & SacaValor("Ventas", "Folio", " WHERE ID=" & IDVenta) & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'
'            'el cargo de ventas
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'DEV01','620401'," & ConvMoneda(SacaValor("Ventas", "Total", " WHERE ID=" & IDVenta)) & "," & _
                    TIPO_CARGO & ",0,'Devolucion Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"


            'Iva
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'DEV01','120101'," & ConvMoneda(SacaValor("Ventas", "Total*(Iva/100)", " WHERE ID=" & IDVenta)) & "," & _
                    TIPO_CARGO & ",0,'Devolucion Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'            'disminuimos la venta
'            'de la cuenta de oro o de misc
'''            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
'''                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'DEV01','" & CuentaVenta & "'," & ConvMoneda(SacaValor("Ventas", "Total", " WHERE ID=" & IDVenta)) & "," & _
'''                    TIPO_ABONO & ",0,'Devolucion Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
'
'            'la entrada al inventario de ventas
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'DEV01','" & CuentaInventario & "'," & ConvMoneda(crCosto) & "," & _
                    TIPO_CARGO & ",0,'Devolucion Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            
            
'            'Saco el Folio
'            Folio = Regresa_Movimiento(False, "FolioInventario")
'            Regresa_Movimiento True, "FolioInventario"
'            dbDatos.Execute "INSERT INTO entradainventario (Fecha,Folio,TipoEntrada,IDUsuario,IDSucursal) VALUES ('" & _
'                            Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & Folio & "," & ENTRADADEVOLUCION & "," & frmMDI.IDUsuario & "," & Sucursal.Clave & ")"
'            ID = Val(SacaValor("entradainventario", "MAX(ID)"))
'
'            dbDatos.Execute "INSERT INTO detallesentradainventario (IDEntrada,Codigo,CodigoOriginal,SucursalOriginal,sSucursalOriginal,Tipo,Cantidad,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,FechaSaca,NumContrato,IDEmpeno,IDTablaEmpeno,UsuarioRevision,SucursalOrigen,sSucursalOrigen,TipoEntrada,PrecioVitrina,PrecioMinimo,PrecioMaximo,Reparaciones,Devoluciones,Garantias,CantidadPiedras,PesoPiedras,IDEmpleado,IDTablaEmpleado,IDAuditorRevision,IDTablaAuditorRevision,IDDetallesEmpeno,IDTablaDetallesEmpeno,IDAuditorPrecio,IDTablaAuditorPrecio) " & _
'                            "SELECT " & ID & ",Codigo,CodigoOriginal,SucursalOriginal,sSucursalOriginal,Tipo,1,Descripcion,Peso,Kilates,Precio,Costo,Estado,Marca,Modelo,Serie,Color,Tamano,TipoPrenda,Observaciones,FechaSaca,NumContrato,IDEmpeno,IDTablaEmpeno,UsuarioRevision,SucursalOrigen,sSucursalOrigen," & ENTRADADEVOLUCION & ",PrecioVitrina,PrecioMinimo,PrecioMaximo,Reparaciones,1,Garantias,CantidadPiedras,PesoPiedras,IDEmpleado,IDTablaEmpleado,IDAuditorRevision,IDTablaAuditorRevision,IDDetallesEmpeno,IDTablaDetallesEmpeno,IDAuditorPrecio,IDTablaAuditorPrecio " & _
'                            "FROM detallesentradainventario WHERE ID=" & Val(SacaValor("detallesventas", "IDArticulo", "WHERE IDVenta=" & IDVenta))
                            
                            
            
            Imprimir_Nota_Devolucion lblFolio.Caption, Comentarios, ConvMoneda(SacaValor("Ventas", "Total+(Total*(Iva/100))", " WHERE ID=" & IDVenta))
            
''        ElseIf opGarantia.Value Then
''            'saco el movimiento
''            Movimiento = Regresa_Movimiento(False)
''            Regresa_Movimiento True
''
''            'actualizamos la venta
''            dbDatos.Execute "UPDATE Ventas SET Estatus=" & StatusVentas.Garantia & " WHERE ID=" & IDVenta
''
''
''            FolioPase = Regresa_Movimiento(False, "FolioPasesInventario")
''            Regresa_Movimiento True, "FolioPasesInventario"
''
''            'grabamos los cargos entrada al inventario de garantia
''            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,FechaModificacion) VALUES " & _
''                            "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Now, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'GA01','620305'," & ConvMoneda(crCosto) & "," & _
''                            TIPO_CARGO & ",0,'Garantia Ventas','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "')"
''
''            'grabamos el pase de inventario
''            dbDatos.Execute "INSERT INTO pasesinventarios (Fecha,Folio,IDArticulo,Origen,Destino,Motivo,IDUsuario) VALUES ('" & _
''                            Format(Now, "YYYY/MM/dd HH:MM:SS") & "'," & FolioPase & "," & IDArticulo & "," & OrigenInventarioVentas.Vendido & "," & _
''                            OrigenInventarioVentas.Garantia & ",'" & Comentarios & "'," & frmMDI.IDUsuario & ")"
''
''            'actualizamos el status del articulo
''            dbDatos.Execute "UPDATE detallesentradainventario SET TipoSalida=" & OrigenInventarioVentas.Garantia & ",Destino=" & OrigenInventarioVentas.Garantia & " WHERE ID=" & IDArticulo
''
''            Imprimir_Nota_Garantia Folio, Comentarios
''
''        End If
        
        MsgBox "Movimiento Realizado", vbOKOnly Or vbInformation
        
        Limpiar
    Else
        MsgBox "Favor de poner los comentarios", vbOKOnly Or vbInformation
    End If


Error:
    Maneja_Error Err

End Sub

Private Sub Imprimir_Nota_Devolucion(Folio As Long, Motivo As String, Importe As Currency)
    Dim ImprDefault As Boolean
    'Dim Impresora As Printer
    
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
    'Regresa_Impresora Tickets, Impresora
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .ReportFileName = Path & "\Reportes\NotaDevolucion.rpt"
        
        .Formulas(0) = "Encabezado='DEVOLUCION VENTA'"
        .Formulas(1) = "Fecha='" & Format(Now, "DD/MM/YYYY") & "'"
        .Formulas(2) = "Folio=" & Folio
        .Formulas(3) = "Cliente='" & m_Cliente & "'"
        .Formulas(4) = "FolioVenta=" & txtFolioFactura.text
        .Formulas(5) = "Articulo='" & m_Articulo & "'"
        .Formulas(6) = "Usuario='" & SacaValor("Usuarios", "Usuario", "WHERE ID=" & frmMDI.IDUsuario) & "'"
        .Formulas(7) = "Caja='" & NombrePc & "'"
        
        .Formulas(8) = "RazonSocial='" & SacaValor("Sucursales", "RazonSocial", "WHERE Clave=" & Sucursal.Clave) & "'"
        .Formulas(9) = "NombreComercial='" & SacaValor("Sucursales", "NombreComercial", "WHERE Clave=" & Sucursal.Clave) & "'"
        .Formulas(10) = "Direccion='" & SacaValor("Sucursales", "Direccion", "WHERE Clave=" & Sucursal.Clave) & "'"
        .Formulas(11) = "Ciudad='" & SacaValor("Sucursales", "Ciudad", "WHERE Clave=" & Sucursal.Clave) & "'"
        .Formulas(12) = "Estado='" & SacaValor("Sucursales", "Estado", "WHERE Clave=" & Sucursal.Clave) & "'"
        .Formulas(13) = "Telefono='" & SacaValor("Sucursales", "Telefono", "WHERE Clave=" & Sucursal.Clave) & "'"
        
        .Formulas(14) = "Empleado='" & m_Empleado & "'"
        .Formulas(15) = "Motivo='" & Motivo & "'"
        .Formulas(16) = "Codigo='" & m_Codigo & "'"
        .Formulas(17) = "Importe=" & Importe
        
        .Destination = crptToWindow
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
        .WindowTitle = "Nota Garantia"
        .WindowState = crptMaximized
        .Action = 1
    End With
End Sub

'''''Private Sub Imprimir_Nota_Garantia(Folio As Long, Motivo As String)
'''''    Dim ImprDefault As Boolean
'''''    Dim Impresora As Printer
'''''
'''''    Regresa_Impresora Tickets, Impresora
'''''    With frmMDI.Cr
'''''        .Reset
'''''        .DiscardSavedData = True
'''''        .WindowShowPrintSetupBtn = True
'''''        .ReportFileName = Path & "\Reportes\NotaGarantia.rpt"
'''''
'''''        .Formulas(0) = "Encabezado='GARANTIA'"
'''''        .Formulas(1) = "Fecha='" & Format(Now, "DD/MM/YYYY") & "'"
'''''        .Formulas(2) = "Folio=" & Folio
'''''        .Formulas(3) = "Cliente='" & m_Cliente & "'"
'''''        .Formulas(4) = "FolioVenta=" & txtFolioFactura.text
'''''        .Formulas(16) = "FechaVenta='" & Format(m_FechaVenta, "DD/MM/YYYY") & "'"
'''''        .Formulas(5) = "Articulo='" & m_Articulo & "'"
'''''        .Formulas(6) = "Usuario='" & SacaValor("Usuarios", "Usuario", "WHERE ID=" & frmMDI.IDUsuario) & "'"
'''''        .Formulas(7) = "Caja='" & NombrePc & "'"
'''''
'''''        .Formulas(8) = "RazonSocial='" & SacaValor("Sucursales", "RazonSocial", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(9) = "NombreComercial='" & SacaValor("Sucursales", "NombreComercial", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(10) = "Direccion='" & SacaValor("Sucursales", "Direccion", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(11) = "Ciudad='" & SacaValor("Sucursales", "Ciudad", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(12) = "Estado='" & SacaValor("Sucursales", "Estado", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(13) = "Telefono='" & SacaValor("Sucursales", "Telefono", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''
'''''        .Formulas(14) = "Empleado='" & m_Empleado & "'"
'''''        .Formulas(15) = "Motivo='" & Motivo & "'"
'''''
'''''        .Destination = crptToWindow
'''''
'''''        AsignaImpresoraCR Impresora, crptToWindow
'''''        .WindowTitle = "Nota Garantia"
'''''        .WindowState = crptMaximized
'''''        .Action = 1
'''''    End With
'''''End Sub
'''''
'''''Private Sub Imprimir_Nota_Entrega_Garantia(Folio As Long, Motivo As String)
'''''    Dim ImprDefault As Boolean
'''''    Dim Impresora As Printer
'''''
'''''    Regresa_Impresora Tickets, Impresora
'''''    With frmMDI.Cr
'''''        .Reset
'''''        .DiscardSavedData = True
'''''        .WindowShowPrintSetupBtn = True
'''''        .ReportFileName = Path & "\Reportes\NotaEntregaGarantia.rpt"
'''''
'''''        .Formulas(0) = "Encabezado='ENTREGA GARANTIA'"
'''''        .Formulas(1) = "Fecha='" & Format(Now, "DD/MM/YYYY") & "'"
'''''        .Formulas(2) = "FolioEntrega=" & Folio
'''''        .Formulas(3) = "Cliente='" & m_Cliente & "'"
'''''        .Formulas(4) = "FolioVenta=" & txtFolioFactura.text
'''''        .Formulas(5) = "Articulo='" & m_Articulo & "'"
'''''        .Formulas(6) = "Usuario='" & SacaValor("Usuarios", "Usuario", "WHERE ID=" & frmMDI.IDUsuario) & "'"
'''''        .Formulas(7) = "Caja='" & NombrePc & "'"
'''''
'''''        .Formulas(8) = "RazonSocial='" & SacaValor("Sucursales", "RazonSocial", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(9) = "NombreComercial='" & SacaValor("Sucursales", "NombreComercial", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(10) = "Direccion='" & SacaValor("Sucursales", "Direccion", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(11) = "Ciudad='" & SacaValor("Sucursales", "Ciudad", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(12) = "Estado='" & SacaValor("Sucursales", "Estado", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''        .Formulas(13) = "Telefono='" & SacaValor("Sucursales", "Telefono", "WHERE Clave=" & Sucursal.Clave) & "'"
'''''
'''''        .Formulas(14) = "Empleado='" & m_Empleado & "'"
'''''        .Formulas(15) = "Motivo='" & Motivo & "'"
'''''        .Formulas(16) = "Folio=" & txtFolioFactura.text
'''''
'''''        .Destination = crptToWindow
'''''
'''''        AsignaImpresoraCR Impresora, crptToWindow
'''''        .WindowTitle = "Nota Garantia"
'''''        .WindowState = crptMaximized
'''''        .Action = 1
'''''    End With
'''''End Sub

Private Sub Limpiar()
    txtFolioFactura.text = ""
    txtFolioFactura.Tag = ""
    lblFolio.Caption = GetFolio(False)
    m_Cliente = ""
    m_Empleado = ""
    m_Articulo = ""
    Me.cDatosVenta.Clear
    Me.cArticulos.Clear
End Sub
