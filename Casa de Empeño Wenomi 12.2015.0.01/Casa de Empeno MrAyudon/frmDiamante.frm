VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmDiamante 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Agregar Diamantes"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6720
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbCorte 
      Height          =   315
      ItemData        =   "frmDiamante.frx":0000
      Left            =   1020
      List            =   "frmDiamante.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   1995
   End
   Begin VB.TextBox txtPeso 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1050
      MaxLength       =   20
      TabIndex        =   4
      Top             =   900
      Width           =   1395
   End
   Begin vbAcceleratorGrid6.vbalGrid grdDiamante 
      Height          =   3510
      Left            =   30
      TabIndex        =   11
      Top             =   1695
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   6191
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
      ScrollBarStyle  =   2
      DisableIcons    =   -1  'True
   End
   Begin VB.TextBox txtAvaluo 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4110
      MaxLength       =   20
      TabIndex        =   5
      Top             =   900
      Width           =   1395
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4110
      MaxLength       =   20
      TabIndex        =   3
      Top             =   525
      Width           =   1395
   End
   Begin VB.ComboBox cmbKilates 
      Height          =   315
      ItemData        =   "frmDiamante.frx":0004
      Left            =   1020
      List            =   "frmDiamante.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   480
      Width           =   1995
   End
   Begin VB.ComboBox cmbEstado 
      Height          =   315
      ItemData        =   "frmDiamante.frx":0008
      Left            =   4080
      List            =   "frmDiamante.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   60
      Width           =   1995
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAgregar 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   1260
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Agregar"
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
      Picture         =   "frmDiamante.frx":000C
      PictureDisabled =   "frmDiamante.frx":0376
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   5505
      TabIndex        =   20
      Top             =   5580
      Width           =   1095
      _ExtentX        =   1931
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
      Picture         =   "frmDiamante.frx":04D0
   End
   Begin VB.Label lblTotAvaluo 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   1440
      TabIndex        =   19
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPrestamo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   240
      Left            =   1050
      TabIndex        =   18
      Top             =   1320
      Width           =   1395
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Préstamo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   1320
      Width           =   870
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Corte:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   510
   End
   Begin VB.Label lblCantidad 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   255
      Left            =   2640
      TabIndex        =   15
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblPuntos 
      Alignment       =   1  'Right Justify
      Caption         =   "0.00"
      Height          =   255
      Left            =   3600
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Label lblTotPrestamo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   60
      TabIndex        =   13
      Top             =   5220
      Width           =   6630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Peso Qte.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   930
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Avalúo:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   10
      Top             =   900
      Width           =   630
   End
   Begin VB.Label Label33 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   9
      Top             =   540
      Width           =   795
   End
   Begin VB.Label Label34 
      AutoSize        =   -1  'True
      Caption         =   "Calidad:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   540
      Width           =   660
   End
   Begin VB.Label Label35 
      AutoSize        =   -1  'True
      Caption         =   "Color:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3240
      TabIndex        =   7
      Top             =   120
      Width           =   480
   End
End
Attribute VB_Name = "frmDiamante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim frm As Form

Private Sub cmbCorte_Click()
    Calcular_Avaluo
End Sub

Private Sub cmbCorte_GotFocus()
    Cambiar_Color True, cmbCorte
End Sub

Private Sub cmbCorte_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbCorte_LostFocus()
    Cambiar_Color False, cmbCorte
End Sub

Private Sub cmbEstado_Click()
    Calcular_Avaluo
End Sub

Private Sub cmbEstado_GotFocus()
    Cambiar_Color True, cmbEstado
End Sub

Private Sub cmbEstado_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbEstado_LostFocus()
    Cambiar_Color False, cmbEstado
End Sub

Private Sub cmbKilates_Click()
    Calcular_Avaluo
End Sub

Private Sub cmbKilates_GotFocus()
    Cambiar_Color True, cmbKilates
End Sub

Private Sub cmbKilates_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbKilates_LostFocus()
    Cambiar_Color False, cmbKilates
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdAceptar_Click()
Dim Piedra As Double, Prestamo As Double, Avaluo As Double, i As Integer, strDescripcion

    If Val(frm.txtPesoPiedra.text) > 0 Or (Trim(frm.txtPesoPiedra.text) <> "" And Trim(frm.txtPesoPiedra.text) <> ".") Then
        
        Piedra = CDbl(frm.txtPesoPiedra.text)
    Else
        
        Piedra = 0
    End If
    
    If Val(lblTotPrestamo.Caption) > 0 Or (Trim(lblTotPrestamo.Caption) <> "" And Trim(lblTotPrestamo.Caption) <> ".") Then
        
        Prestamo = CDbl(lblTotPrestamo.Caption)
    Else
        
        Prestamo = 0
    End If
    
    If Val(lblTotAvaluo.Caption) > 0 Or (Trim(lblTotAvaluo.Caption) <> "" And Trim(lblTotAvaluo.Caption) <> ".") Then
        
        Avaluo = CDbl(lblTotAvaluo.Caption)
    Else
                
        Avaluo = 0
    End If
    
    strDescripcion = ""
    frm.lblAvaluoDiamante.Caption = Format(Avaluo, FMoneda)
    frm.lblPrestamoDiamante.Caption = Format(Prestamo, FMoneda)
        
    For i = 1 To grdDiamante.Rows
    
        strDescripcion = strDescripcion & " " & grdDiamante.CellText(i, 3) & " DMTE." & " CORTE " & grdDiamante.CellText(i, 1) & " " & grdDiamante.CellText(i, 5) & " " & grdDiamante.CellText(i, 4) & " PAT: " & Format(grdDiamante.CellText(i, 2), "0.00") & "QTE."
    Next i
    
    frm.lblPiedra.Caption = strDescripcion
    frm.lblPuntos.Caption = CDbl(lblPuntos.Caption)
    frm.lblCantidadPiedras.Caption = Val(lblCantidad.Caption)
    
    Unload Me
End Sub

Private Sub cmdAgregar_Click()

    If Completos Then

        With grdDiamante
            
            .AddRow
            .CellText(.Rows, 1) = cmbCorte.text
            .CellTextAlign(.Rows, 1) = DT_LEFT
            
            .CellText(.Rows, 2) = txtPeso.text
            .CellTextAlign(.Rows, 2) = DT_CENTER
        
            .CellText(.Rows, 3) = txtCantidad.text
            .CellTextAlign(.Rows, 3) = DT_CENTER
        
            .CellText(.Rows, 4) = cmbKilates.text
            .CellTextAlign(.Rows, 4) = DT_LEFT
        
            .CellText(.Rows, 5) = cmbEstado.text
            .CellTextAlign(.Rows, 5) = DT_LEFT
            
            .CellText(.Rows, 6) = lblPrestamo.Caption
            .CellTextAlign(.Rows, 6) = DT_RIGHT
            
            .CellText(.Rows, 7) = txtAvaluo.text
            .CellTextAlign(.Rows, 7) = DT_RIGHT
        
            CalculaTotal
        
            txtPeso.text = ""
            txtCantidad.text = ""
            txtAvaluo.text = ""
            cmbKilates.ListIndex = -1
            cmbEstado.ListIndex = -1
            cmbCorte.SetFocus
        End With

    End If

End Sub

Private Sub Form_Deactivate()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    CreaEncabezado
    Cargar_Combos "Descripcion", "kilatajes", cmbKilates, " where IDTipo= 4", "Ordenamiento"
    Cargar_Combos "Estado", "estado", cmbEstado, " where IDTipo=4", "Ordenamiento"
    Cargar_Combos "Descripcion", "cortediamantes", cmbCorte, , "Ordenamiento"
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdDiamante_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If grdDiamante.Rows > 0 And grdDiamante.SelectedRow > 0 And KeyCode = vbKeyDelete Then
        
        If MsgBox("Desea eliminar el registro seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Agregar Diamante") = vbYes Then
            
            grdDiamante.RemoveRow grdDiamante.SelectedRow
            CalculaTotal
            cmbCorte.SetFocus
        End If
    End If

End Sub

Private Sub txtCantidad_Change()
    Calcular_Avaluo
End Sub

Private Sub txtCantidad_GotFocus()
    Cambiar_Color True, txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCantidad_LostFocus()
    Cambiar_Color False, txtCantidad
End Sub

Private Sub txtAvaluo_GotFocus()
    Cambiar_Color True, txtAvaluo
End Sub

Private Sub txtAvaluo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtAvaluo_LostFocus()
    Cambiar_Color False, txtAvaluo
End Sub

'Calculamos el avaluo
Private Function Calcular_Avaluo() As Double
Dim crPrecio As Double, Cantidad As Double, PesoPiedra As Double, Puntos As Double, IDPunto As Long, PrestamoAvaluo As Double, TipoCorte As Double
Dim rcTmp As New ADODB.Recordset

On Error GoTo error

    If cmbKilates.ListIndex >= 0 And cmbEstado.ListIndex >= 0 And cmbCorte.ListIndex >= 0 Then
                
        TipoCorte = SacaValor("cortediamantes", "descuento", " WHERE ID=" & cmbCorte.ItemData(cmbCorte.ListIndex))
        
        If Val(txtCantidad.text) > 0 Or (Trim(txtCantidad.text) <> "" And Trim(txtCantidad.text) <> ".") Then
            
            Cantidad = CDbl(txtCantidad.text)
        Else
        
            Cantidad = 0
        End If
    
        If Val(txtPeso.text) > 0 Or (Trim(txtPeso.text) <> "" And Trim(txtPeso.text) <> ".") Then
            
            Puntos = CDbl(txtPeso.text)
        Else
            
            Puntos = 0
        End If
        
        IDPunto = SacaPuntos(Puntos)
    
        rcTmp.Open "SELECT Precio FROM precioskilataje WHERE IDTipo=4 AND IDKilataje = " & RegresaKilates(cmbKilates.text, "DIAMANTES") & " AND IDHechura = " & cmbEstado.ItemData(cmbEstado.ListIndex) & " AND IDRango=" & IDPunto, dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcTmp.BOF And Not rcTmp.EOF And Not IsNull(rcTmp!Precio) Then
            
            crPrecio = rcTmp!Precio
        Else
            
            crPrecio = 0
        End If
        rcTmp.Close
        
        PrestamoAvaluo = Regresa_Valor_BD("PrestamoAvaluoDiamante")
        txtAvaluo.text = Format(Redondeo(Cantidad * (Puntos * 100) * crPrecio) * (TipoCorte / 100), "###,###,###,###0.00")
        lblPrestamo.Caption = Format(Redondeo(Cantidad * (((Puntos * 100) * crPrecio) * (PrestamoAvaluo / 100)) * (TipoCorte / 100)), "###,###,###,###0.00")
    End If

error:
    Maneja_Error Err
    Set rcTmp = Nothing
   
End Function

Private Sub txtPeso_Change()
    Calcular_Avaluo
End Sub

Private Sub txtPeso_GotFocus()
    Seleccionar_Texto txtPeso
    Cambiar_Color True, txtPeso
End Sub

Private Sub txtPeso_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPeso_LostFocus()
    Cambiar_Color False, txtPeso
End Sub

Function SacaPuntos(Punto As Double) As Long
Dim rcPuntos As New ADODB.Recordset
Dim De As Double, a As Double, Posicion As Integer

On Error GoTo error

    If Punto > 0 Then
        
        rcPuntos.Open "SELECT diamantepuntos.ID,diamantepuntos.Punto FROM diamantepuntos ORDER BY diamantepuntos.ID", dbDatos, adOpenForwardOnly, adLockOptimistic
        If Not rcPuntos.BOF And Not rcPuntos.EOF Then
            
            rcPuntos.MoveFirst
            While Not rcPuntos.EOF
                Posicion = InStr(1, rcPuntos!Punto, "-")

                If Posicion > 0 Then
                    
                    De = Mid(rcPuntos!Punto, 1, (InStr(1, rcPuntos!Punto, "-")) - 1)
                    a = Mid(rcPuntos!Punto, (InStr(1, rcPuntos!Punto, "-")) + 1, Len(rcPuntos!Punto) - InStr(1, rcPuntos!Punto, "-") + 1)
                Else
                    
                    De = rcPuntos!Punto
                End If
            
                If Posicion > 0 Then
                    If Punto >= De And Punto <= a Then
                        SacaPuntos = rcPuntos!ID
                        GoTo 125
                    End If

                Else

                    If Punto >= De Then
                        SacaPuntos = rcPuntos!ID
                        GoTo 125
                    End If
                End If
            
            rcPuntos.MoveNext
            Wend
        End If

125:
        rcPuntos.Close
    End If

error:
    Maneja_Error Err
    Set rcPuntos = Nothing
End Function

Sub CreaEncabezado()

    With grdDiamante
        .AddColumn "C1", "Corte", ecgHdrTextALignLeft, , 65, , , , , , , CCLSortNumeric
        .AddColumn "C2", "Peso", ecgHdrTextALignRight, , 43, , , , , , , CCLSortNumeric
        .AddColumn "C3", "Cant.", ecgHdrTextALignRight, , 40, False, , , , , , CCLSortNumeric
        .AddColumn "C4", "Color", ecgHdrTextALignLeft, , 90, , , , , , , CCLSortNumeric
        .AddColumn "C5", "Calidad", ecgHdrTextALignLeft, , 105, , , , , , , CCLSortNumeric
        .AddColumn "C6", "Préstamo", ecgHdrTextALignRight, , 65, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C7", "Avalúo", ecgHdrTextALignRight, , 65, , , , , FMoneda, , CCLSortNumeric
    End With

End Sub

Function Completos() As Boolean

    Completos = True
    
    If cmbCorte.ListIndex = -1 Then
        Completos = False
        Exit Function
    End If
    
    If cmbKilates.ListIndex = -1 Then
        Completos = False
        Exit Function
    End If

    If cmbEstado.ListIndex = -1 Then
        Completos = False
        Exit Function
    End If

    If Val(txtCantidad.text) = 0 Or Trim(txtCantidad.text) = "" Then
        Completos = False
        Exit Function
    End If

    If Val(txtPeso.text) = 0 Or Trim(txtPeso.text) = "" Then
        Completos = False
        Exit Function
    End If

End Function

Function CalculaTotal()
Dim i As Integer, Prestamo As Double, Avaluo As Double, Puntos As Double, Cantidad As Integer

    Prestamo = 0
    For i = 1 To grdDiamante.Rows
        
        Cantidad = Val(grdDiamante.CellText(i, 3))
        Puntos = Puntos + IIf(Val(grdDiamante.CellText(i, 2)) > 0 Or Trim(grdDiamante.CellText(i, 2)) <> "", CDbl(grdDiamante.CellText(i, 2)), 0)
        Prestamo = Prestamo + CDbl(grdDiamante.CellText(i, 6))
        Avaluo = Avaluo + CDbl(grdDiamante.CellText(i, 7))
    Next i

    lblTotPrestamo.Caption = Format(Prestamo, FMoneda)
    lblTotAvaluo.Caption = Format(Avaluo, FMoneda)
    
    lblPuntos.Caption = Puntos
    lblCantidad.Caption = Cantidad
End Function

Public Sub Mostrar(formulario As Form)
    Set frm = formulario
    Me.Show vbModal
End Sub
