VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~2.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{8FD826E4-642F-44F1-BF42-64C39ED09F7F}#2.0#0"; "Linea3D.ocx"
Begin VB.Form frmPreciosDiamante 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios Diamante"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreciosDiamante.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7950
   ScaleWidth      =   5325
   Begin VB.ComboBox cmbPuntos 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   3285
   End
   Begin Line3D.ucLine3D ucLine3D10 
      Height          =   30
      Left            =   630
      Top             =   1980
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D9 
      Height          =   30
      Left            =   645
      Top             =   1665
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D8 
      Height          =   60
      Left            =   630
      Top             =   1350
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   106
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D7 
      Height          =   30
      Left            =   645
      Top             =   975
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D6 
      Height          =   1380
      Left            =   2640
      Top             =   630
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2434
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D5 
      Height          =   30
      Left            =   645
      Top             =   675
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D4 
      Height          =   1260
      Left            =   4560
      Top             =   135
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   2223
      Orientation     =   0
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D3 
      Height          =   30
      Left            =   630
      Top             =   345
      Width           =   3960
      _ExtentX        =   6985
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D2 
      Height          =   30
      Left            =   630
      Top             =   105
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   53
      LineWidth       =   2
   End
   Begin Line3D.ucLine3D ucLine3D1 
      Height          =   1875
      Left            =   630
      Top             =   105
      Width           =   30
      _ExtentX        =   53
      _ExtentY        =   3307
      Orientation     =   0
      LineWidth       =   2
   End
   Begin VB.ComboBox cmbKilataje 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1020
      Width           =   1815
   End
   Begin VB.ComboBox cmbHechura 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2745
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1020
      Width           =   1770
   End
   Begin VB.TextBox txtPrecio 
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
      Left            =   690
      TabIndex        =   3
      Top             =   1710
      Width           =   1920
   End
   Begin vbAcceleratorGrid6.vbalGrid grdPrecios 
      Height          =   5370
      Left            =   60
      TabIndex        =   4
      Top             =   2070
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   9472
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
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   7515
      Width           =   975
      _ExtentX        =   1720
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
      Picture         =   "frmPreciosDiamante.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4185
      TabIndex        =   12
      Top             =   7515
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
      Picture         =   "frmPreciosDiamante.frx":0110
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   1560
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
      Picture         =   "frmPreciosDiamante.frx":0662
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AVALÚO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1275
      TabIndex        =   6
      Top             =   1395
      Width           =   750
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Height          =   270
      Left            =   645
      TabIndex        =   11
      Top             =   1395
      Width           =   2010
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      Caption         =   "PESO QTE."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   630
      TabIndex        =   9
      Top             =   120
      Width           =   3960
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "CALIDAD"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   1185
      TabIndex        =   8
      Top             =   735
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "COLOR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   3240
      TabIndex        =   7
      Top             =   735
      Width           =   645
   End
   Begin VB.Label Label7 
      BackColor       =   &H80000013&
      Height          =   270
      Left            =   660
      TabIndex        =   10
      Top             =   705
      Width           =   3915
   End
End
Attribute VB_Name = "frmPreciosDiamante"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbHechura_GotFocus()
    cmbHechura.BackColor = &HC0FFFF
End Sub

Private Sub cmbHechura_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbHechura_LostFocus()
    cmbHechura.BackColor = vbWhite
End Sub

Private Sub cmbKilataje_GotFocus()
    cmbKilataje.BackColor = &HC0FFFF
End Sub

Private Sub cmbKilataje_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbKilataje_LostFocus()
    cmbKilataje.BackColor = vbWhite
End Sub

Private Sub cmbPuntos_GotFocus()
    Cambiar_Color True, cmbPuntos
End Sub

Private Sub cmbPuntos_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPuntos_LostFocus()
    Cambiar_Color False, cmbPuntos
End Sub

Private Sub cmdAceptar_Click()

    If Completos Then
    
        GrabaDatos
        CargaDatos
        cmbKilataje.SetFocus
    
    End If

End Sub

Private Sub cmdLimpiar_Click()
    Limpiar
    grdPrecios.ClearSelection
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Cargar_Combos "Punto", "diamantepuntos", cmbPuntos
    Cargar_Combos "Descripcion", "Kilatajes", cmbKilataje, " WHERE IDTipo=4", "Ordenamiento"
    Cargar_Combos "Estado", "Estado", cmbHechura, " WHERE IDTipo=4", "Ordenamiento"
    CrearEncabezados
    CargaDatos
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdPrecios_DblClick(ByVal lRow As Long, _
                                ByVal lCol As Long)

    '''''On Error GoTo error
    '''''
    '''''With grdPrecios
    '''''    If .Rows > 0 And .SelectedRow > 0 Then
    '''''        cmbKilataje.ListIndex = ComboInformacion(cmbKilataje, .CellItemData(.SelectedRow, 2))
    '''''        cmbHechura.ListIndex = ComboInformacion(cmbHechura, .CellItemData(.SelectedRow, 3))
    '''''        txtPrecio.Text = Format(grdPrecios.CellText(grdPrecios.SelectedRow, 4), "###,###,###0.00")
    '''''
    '''''        txtPrecio.Tag = grdPrecios.CellItemData(grdPrecios.SelectedRow, 4)
    '''''    End If
    '''''End With
    '''''
    '''''error:
    '''''    Maneja_Error Err
End Sub

Private Sub grdPrecios_KeyDown(KeyCode As Integer, Shift As Integer, bDoDefault As Boolean)

    If grdPrecios.Rows > 0 And grdPrecios.SelectedRow > 0 And KeyCode = vbKeyDelete Then
        
        If MsgBox("Desea eliminar el precio seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Precios Diamante") = vbYes Then
            
            dbDatos.Execute "DELETE FROM precioskilataje WHERE ID=" & grdPrecios.CellItemData(grdPrecios.SelectedRow, 4)
            grdPrecios.RemoveRow grdPrecios.SelectedRow
            CargaDatos
        
        Else
            
            grdPrecios.ClearSelection
        
        End If

        cmbPuntos.SetFocus
    
    End If

End Sub

Private Sub txtPrecio_GotFocus()
    Seleccionar_Texto txtPrecio
    Cambiar_Color True, txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrecio_LostFocus()
    Cambiar_Color False, txtPrecio
End Sub

Sub CrearEncabezados()

    With grdPrecios
    
        .AddColumn "C1", "Peso", ecgHdrTextALignLeft, , 60, , , , , , , CCLSortString
        .AddColumn "C2", "Calidad", ecgHdrTextALignCentre, , 93, , , , , , , CCLSortString
        .AddColumn "C3", "Color", ecgHdrTextALignCentre, , 103, , , , , , , CCLSortString
        .AddColumn "C4", "Avalúo", ecgHdrTextALignRight, , 67, , , , , FMoneda, , CCLSortNumeric

    End With

End Sub

Sub CargaDatos()
Dim rcTmp As New ADODB.Recordset

On Error GoTo Error
    
    grdPrecios.Redraw = False
    grdPrecios.Clear
    
    rcTmp.Open "SELECT diamantepuntos.Punto,PreciosKilataje.ID as IDPrecios,Tipo.Descripcion AS Tipos_Descripcion,Kilatajes.Descripcion AS Kilatajes_Descripcion, Estado.Estado AS Hechuras_Descripcion,Estado.ID as IDHechura,PreciosKilataje.Precio,PreciosKilataje.IDTipo,PreciosKilataje.IDKilataje" & " FROM PreciosKilataje Inner Join Tipo on PreciosKilataje.IDTipo=Tipo.ID Inner Join Kilatajes on Kilatajes.ID = PreciosKilataje.IDKilataje Inner Join Estado on PreciosKilataje.IDHechura=Estado.ID Inner Join diamantepuntos on PreciosKilataje.IDRango=diamantepuntos.ID where PreciosKilataje.IDTipo=4 Order By PreciosKilataje.IDRango,PreciosKilataje.IDKilataje,PreciosKilataje.IDHechura", dbDatos, adOpenForwardOnly, adLockReadOnly
    With grdPrecios
    
        While Not rcTmp.EOF
            
            .AddRow
            .CellText(.Rows, 1) = rcTmp!Punto
            .CellItemData(.Rows, 1) = rcTmp!IDTipo
            .CellText(.Rows, 2) = rcTmp!Kilatajes_Descripcion
            .CellItemData(.Rows, 2) = rcTmp!IDKilataje
            .CellTextAlign(.Rows, 2) = DT_LEFT
            .CellText(.Rows, 3) = rcTmp!Hechuras_Descripcion
            .CellItemData(.Rows, 3) = rcTmp!IDHechura
            .CellTextAlign(.Rows, 3) = DT_LEFT
            .CellText(.Rows, 4) = rcTmp!Precio
            .CellItemData(.Rows, 4) = rcTmp!IDPrecios
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            
        rcTmp.MoveNext
        Wend
    
    End With
    rcTmp.Close
    grdPrecios.Redraw = True
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Sub Limpiar()
    cmbPuntos.ListIndex = -1
    cmbKilataje.ListIndex = -1
    cmbHechura.ListIndex = -1
    txtPrecio.text = ""
    txtPrecio.Tag = ""
End Sub

Sub GrabaDatos()

    If Completos Then
        
        dbDatos.Execute "INSERT INTO precioskilataje (IDRango,IDTipo,IDKilataje,IDHechura,Precio) VALUES (" & _
                        cmbPuntos.ItemData(cmbPuntos.ListIndex) & ",4," & cmbKilataje.ItemData(cmbKilataje.ListIndex) & "," & cmbHechura.ItemData(cmbHechura.ListIndex) & "," & CDbl(txtPrecio.text) & ")"
        
        cmbKilataje.ListIndex = -1
        cmbHechura.ListIndex = -1
        txtPrecio.text = ""
        txtPrecio.SetFocus
        
    End If

End Sub

Function Completos() As Boolean

    Completos = True

    If cmbPuntos.ListIndex = -1 Then
        Completos = False
        Exit Function
    End If

    If cmbKilataje.ListIndex = -1 Then
        Completos = False
        Exit Function
    End If

    If cmbHechura.ListIndex = -1 Then
        Completos = False
        Exit Function
    End If

    If Val(txtPrecio.text) = 0 Or Trim(txtPrecio.text) = "" Then
        Completos = False
        txtPrecio.SetFocus
        Exit Function
    End If

End Function
