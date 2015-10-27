VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~1.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCotizacion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cotización diaria"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCotizacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6060
   ScaleWidth      =   10215
   Begin vbAcceleratorGrid6.vbalGrid grdDivisas 
      Height          =   5520
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   9737
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
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
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
      Begin VB.ComboBox cmbDivisa 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.TextBox txtPonderado 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtVenta 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4800
         TabIndex        =   2
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox txtCompra 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3360
         TabIndex        =   1
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
      Height          =   390
      Left            =   6615
      TabIndex        =   5
      Top             =   5625
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "  A&gregar"
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
      Picture         =   "frmCotizacion.frx":000C
      PictureDisabled =   "frmCotizacion.frx":0376
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   9000
      TabIndex        =   6
      Top             =   5625
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
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
      Picture         =   "frmCotizacion.frx":04D0
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   390
      Left            =   7800
      TabIndex        =   7
      Top             =   5625
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
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
      Object.ToolTipText     =   ""
      Picture         =   "frmCotizacion.frx":0A22
   End
End
Attribute VB_Name = "frmCotizacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Sub Crear_Encabezados()

    With grdDivisas
        .AddColumn "K1", "Divisa", ecgHdrTextALignLeft, , 160, , , , , , , CCLSortNumeric
        .AddColumn "K2", "Fecha", ecgHdrTextALignLeft, , 200, , , , , , , CCLSortDate
        .AddColumn "K3", "Hora", ecgHdrTextALignLeft, , 85, False, , , , , , CCLSortString
        .AddColumn "K4", "Cotización Compra", ecgHdrTextALignRight, , 147, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K5", "Cotización Venta", ecgHdrTextALignRight, , 147, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "K6", "Ponderado", ecgHdrTextALignRight, , 90, False, , , , FMoneda, , CCLSortNumeric
    End With

End Sub

Private Sub cmbDivisa_Click()

    grdDivisas.CellText(grdDivisas.SelectedRow, 1) = cmbDivisa.text
    grdDivisas.CellItemData(grdDivisas.SelectedRow, 1) = cmbDivisa.ItemData(cmbDivisa.ListIndex)
    cmbDivisa.Visible = False

End Sub

Private Sub cmbDivisa_GotFocus()
    Cambiar_Color True, cmbDivisa
End Sub

Private Sub cmbDivisa_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then cmbDivisa.Visible = False
End Sub

Private Sub cmbDivisa_LostFocus()
    Cambiar_Color False, cmbDivisa
End Sub

Private Sub cmdAceptar_Click()
Dim i As Integer, Compra As Double, Venta As Double, Ponderado As Double

    If ValidaGrid Then
    
        dbDatos.Execute "DELETE FROM cotizaciones"
    
        For i = 1 To grdDivisas.Rows

            If grdDivisas.CellText(i, 1) <> "" Then
    
                Compra = CDbl(grdDivisas.CellText(i, 4))
                Venta = CDbl(grdDivisas.CellText(i, 5))
                Ponderado = 0
        
                dbDatos.Execute "INSERT INTO cotizaciones(IDMoneda,Fecha,Hora,Compra,Venta,Ponderado,IDUsuario,IDSucursal) VALUES (" & _
                                grdDivisas.CellItemData(i, 1) & ",'" & IIf(grdDivisas.CellText(i, 3) = "", Format(Date, "YYYY/MM/DD"), Format(grdDivisas.CellText(i, 2), "YYYY/MM/DD")) & "','" & IIf(grdDivisas.CellText(i, 3) = "", Format(Time, "HH:MM:SS"), Format(grdDivisas.CellText(i, 3), "HH:MM:SS")) & "'," & Compra & "," & Venta & "," & Ponderado & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
            End If

        Next i

    End If

End Sub

Private Sub cmdAgregar_Click()
    grdDivisas.AddRow
    grdDivisas.SelectedRow = grdDivisas.Rows
    grdDivisas.CellText(grdDivisas.Rows, 2) = Format(Now, "DD/MM/YYYY HH:MM:SS AM/PM")
    grdDivisas_RequestEdit grdDivisas.Rows, 1, 0, True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Crear_Encabezados
    Carga_Datos
    Cargar_Combos "Descripcion", "monedas", cmbDivisa, , "Clave", , "Clave"
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Sub Carga_Datos()
Dim rcConsulta As New ADODB.Recordset

    rcConsulta.Open "SELECT cotizaciones.*,monedas.Descripcion FROM cotizaciones INNER JOIN monedas ON cotizaciones.IDMoneda=monedas.Clave ORDER BY cotizaciones.Fecha,cotizaciones.Hora", dbDatos, adOpenForwardOnly, adLockReadOnly
    With grdDivisas
        
        While Not rcConsulta.EOF
            .AddRow
            .CellText(.Rows, 1) = rcConsulta!Descripcion
            .CellItemData(.Rows, 1) = rcConsulta!IDMoneda
            .CellText(.Rows, 4) = rcConsulta!Compra
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            .CellText(.Rows, 5) = rcConsulta!Venta
            .CellTextAlign(.Rows, 5) = DT_RIGHT
            .CellText(.Rows, 2) = Format(rcConsulta!Fecha, "DD/MMM/YYYY") & " " & Format(rcConsulta!Hora, "HH:MM:SS AM/PM")
            .CellTextAlign(.Rows, 2) = DT_LEFT
            .CellText(.Rows, 3) = Format(rcConsulta!Hora, "HH:MM:SS AM/PM")
            .CellTextAlign(.Rows, 3) = DT_LEFT
            .CellText(.Rows, 6) = rcConsulta!Ponderado
            .CellTextAlign(.Rows, 6) = DT_RIGHT
        rcConsulta.MoveNext
        Wend
    
    End With
    rcConsulta.Close
    Set rcConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl()
End Sub

Private Sub grdDivisas_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, sText As String, obj As Object

    If grdDivisas.CellText(grdDivisas.SelectedRow, 1) <> "" And grdDivisas.CellText(grdDivisas.SelectedRow, 2) <> "" And grdDivisas.CellText(grdDivisas.SelectedRow, 3) <> "" And grdDivisas.CellText(grdDivisas.SelectedRow, 4) <> "" And grdDivisas.CellText(grdDivisas.SelectedRow, 5) <> "" And grdDivisas.CellText(grdDivisas.SelectedRow, 6) <> "" Then Exit Sub

    If lCol = 2 Or lCol = 3 Then
        cmbDivisa.Visible = False
        txtCompra.Visible = False
        txtVenta.Visible = False
        txtPonderado.Visible = False
        Exit Sub
    End If

    Select Case lCol

        Case 1: Set obj = cmbDivisa

        Case 4: Set obj = txtCompra

        Case 5: Set obj = txtVenta

    End Select
   
    grdDivisas.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight

    If Not IsMissing(grdDivisas.CellText(lRow, lCol)) Then
        sText = IIf(Not IsNull(grdDivisas.CellText(lRow, lCol)), grdDivisas.CellText(lRow, lCol), "")
    Else
        sText = ""
    End If

    If lCol <> 1 And lCol <> 2 And lCol <> 3 Then
        obj.Alignment = vbRightJustify

        If (iKeyAscii > 13) Then
            sText = Chr$(iKeyAscii) & sText
            obj.text = sText
            obj.SelStart = 1
            obj.SelLength = Len(sText)
        Else
            obj.text = sText
            obj.SelStart = 0
            obj.SelLength = Len(sText)
        End If
    End If

    If lCol = 1 Then
        obj.Move lLeft + 40, lTop + 25, lWidth - 60
    ElseIf lCol <> 2 And lCol <> 3 Then
        obj.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50
    End If

    obj.Visible = True
    obj.ZOrder
    obj.SetFocus
End Sub

Private Sub txtCompra_GotFocus()
    Seleccionar_Texto txtCompra
    Cambiar_Color True, txtCompra
End Sub

Private Sub txtCompra_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then txtCompra.Visible = False
End Sub

Private Sub txtCompra_KeyPress(KeyAscii As Integer)

    KeyAscii = Solo_Numeros(KeyAscii, 1)
    If KeyAscii = vbKeyReturn Then
        
        If txtCompra.text <> "" Then
            
            grdDivisas.CellText(grdDivisas.SelectedRow, 4) = Format(CDbl(txtCompra.text), "###,###,###,###0.00")
            grdDivisas.CellTextAlign(grdDivisas.SelectedRow, 4) = DT_RIGHT
            txtCompra.Visible = False
        Else
            
            txtCompra.Visible = False
        End If
    
    End If

End Sub

Private Sub txtCompra_LostFocus()
    Cambiar_Color False, txtCompra
    txtCompra.Visible = False
End Sub

Private Sub txtPonderado_GotFocus()
    Seleccionar_Texto txtPonderado
    Cambiar_Color True, txtPonderado
End Sub

Private Sub txtPonderado_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then txtPonderado.Visible = False
End Sub

Private Sub txtPonderado_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
    
        If txtPonderado.text <> "" Then
            
            grdDivisas.CellText(grdDivisas.SelectedRow, 6) = Format(CDbl(txtPonderado.text), "###,###,###,###0.00")
            grdDivisas.CellTextAlign(grdDivisas.SelectedRow, 6) = DT_RIGHT
            txtPonderado.Visible = False
        Else
            
            txtPonderado.Visible = False
        End If
    
    End If

    KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtPonderado_LostFocus()
    Cambiar_Color False, txtPonderado
    txtPonderado.Visible = False
End Sub

Private Sub txtVenta_GotFocus()
    Seleccionar_Texto txtVenta
    Cambiar_Color True, txtVenta
End Sub

Private Sub txtVenta_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then txtVenta.Visible = False
End Sub

Private Sub txtVenta_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then
    
        If txtVenta.text <> "" Then
            
            grdDivisas.CellText(grdDivisas.SelectedRow, 5) = Format(CDbl(txtVenta.text), "###,###,###,###0.00")
            grdDivisas.CellTextAlign(grdDivisas.SelectedRow, 5) = DT_RIGHT
            txtVenta.Visible = False
        Else
            
            txtVenta.Visible = False
        End If
    
    End If

    KeyAscii = Solo_Numeros(KeyAscii, 1)
End Sub

Private Sub txtVenta_LostFocus()
    Cambiar_Color False, txtVenta
    txtVenta.Visible = False
End Sub

Function ValidaGrid() As Boolean
Dim i As Integer

    ValidaGrid = True

    With grdDivisas

        If .Rows > 0 Then

            For i = 1 To .Rows

                If .CellText(i, 1) = "" Or .CellText(i, 4) = "" Or .CellText(i, 5) = "" Or .CellText(i, 6) = "" Then
                    
                    If .CellText(i, 1) = "" And .CellText(i, 4) = "" And .CellText(i, 5) = "" And .CellText(i, 6) = "" Then GoTo 125
                    If .CellText(i, 1) = "" Then MsgBox "Seleccione el tipo de divisa !!", vbInformation, "Cotización Diaria": ValidaGrid = False: .SelectedRow = i: grdDivisas_RequestEdit i, 1, 13, False: Exit Function
                    If .CellText(i, 4) = "" Then MsgBox "Introduzca el tipo de cambio a la compra !!", vbInformation, "Cotización Diaria": ValidaGrid = False: .SelectedRow = i: grdDivisas_RequestEdit i, 4, 13, False: Exit Function
                    If .CellText(i, 5) = "" Then MsgBox "Introduzca el tipo de cambio a la venta !!", vbInformation, "Cotización Diaria": ValidaGrid = False: .SelectedRow = i: grdDivisas_RequestEdit i, 5, 13, False: Exit Function
                    '''''If .CellText(i, 6) = "" Then MsgBox "Introduzca el tipo de cambio ponderado !!", vbInformation, "Cotización Diaria": ValidaGrid = False: .SelectedRow = i: grdDivisas_RequestEdit i, 6, 13, False: Exit Function
                
                End If
125:
            Next i

        End If

    End With

End Function
