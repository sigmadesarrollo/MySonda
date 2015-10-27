VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatPromociones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Promociones"
   ClientHeight    =   5535
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatPromociones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5535
   ScaleWidth      =   9030
   Begin VB.TextBox txtDescripcion 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   3720
   End
   Begin VB.TextBox txtDias 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   720
   End
   Begin VB.TextBox txtPorcentaje 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   315
      Left            =   1560
      TabIndex        =   2
      Top             =   840
      Width           =   720
   End
   Begin VB.ComboBox cmbTipo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCatPromociones.frx":000C
      Left            =   1560
      List            =   "frmCatPromociones.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   1920
   End
   Begin vbAcceleratorGrid6.vbalGrid grdCatPromociones 
      Height          =   3885
      Left            =   30
      TabIndex        =   8
      Top             =   1605
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   6853
      RowMode         =   -1  'True
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      GridLineColor   =   14737632
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
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
      Left            =   7800
      TabIndex        =   6
      Top             =   1080
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
      Picture         =   "frmCatPromociones.frx":0036
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   7800
      TabIndex        =   4
      Top             =   120
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
      Object.ToolTipText     =   ""
      Picture         =   "frmCatPromociones.frx":0588
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   375
      Left            =   6480
      TabIndex        =   7
      Top             =   120
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     A&ctivar"
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
      Picture         =   "frmCatPromociones.frx":0ADA
      PictureDisabled =   "frmCatPromociones.frx":102C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdCancelar 
      Height          =   375
      Left            =   7800
      TabIndex        =   5
      Top             =   600
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
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmCatPromociones.frx":1BFE
      PictureDisabled =   "frmCatPromociones.frx":1E4D
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   240
      TabIndex        =   12
      Top             =   172
      Width           =   1230
   End
   Begin VB.Label LabelD 
      Alignment       =   1  'Right Justify
      Caption         =   "Días:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   11
      Top             =   1252
      Width           =   1230
   End
   Begin VB.Label LabelP 
      Alignment       =   1  'Right Justify
      Caption         =   "Porcentaje:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   10
      Top             =   892
      Width           =   1230
   End
   Begin VB.Label Label28 
      Alignment       =   1  'Right Justify
      Caption         =   "Tipo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   30
      Left            =   240
      TabIndex        =   9
      Top             =   532
      Width           =   1230
   End
End
Attribute VB_Name = "frmCatPromociones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbTipo_Click()
    Select Case cmbTipo.text
        Case "Porcentaje"
            LabelP.Enabled = True
            LabelD.Enabled = False
            txtPorcentaje.Enabled = True
            txtDias.Enabled = False
        Case "Dias"
            LabelP.Enabled = False
            LabelD.Enabled = True
            txtPorcentaje.Enabled = False
            txtDias.Enabled = True
        Case "Ambos"
            LabelP.Enabled = True
            LabelD.Enabled = True
            txtPorcentaje.Enabled = True
            txtDias.Enabled = True
        Case Else
            LabelP.Enabled = False
            LabelD.Enabled = False
            txtPorcentaje.Enabled = False
            txtDias.Enabled = False
    End Select
End Sub

Private Sub cmbTipo_GotFocus()
    Cambiar_Color True, cmbTipo
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
    Cambiar_Color False, cmbTipo
End Sub

Private Sub cmdAceptar_Click()
    If Val(txtDescripcion.Tag) = 0 Then
        dbDatos.Execute "INSERT INTO Promociones (IDTabla,Descripcion,Tipo,PorcentajeDescuento,DiasDescuento,Activa) VALUES (0,'" & _
                        txtDescripcion.text & "','" & Left(cmbTipo.text, 1) & "'," & Val(txtPorcentaje.text) & "," & Val(txtDias.text) & ",1)"
    ElseIf Val(txtDescripcion.Tag) > 0 Then
        dbDatos.Execute "UPDATE Promociones SET Descripcion='" & txtDescripcion.text & "',Tipo='" & Left(cmbTipo.text, 1) & "',PorcentajeDescuento=" & Val(txtPorcentaje.text) & ",DiasDescuento=" & Val(txtDias.text) & " WHERE ID=" & Val(txtDescripcion.Tag)
    End If
    cmdCancelar_Click
End Sub


Private Sub cmdEliminar_Click()
    With grdCatPromociones
        If .Rows > 0 Then
            If .SelectedRow > 0 Then
                If MsgBox("Desea " & IIf(.CellItemData(.SelectedRow, 5) = 0, "Activar", "Desactivar") & " la promoción [" & .CellText(.SelectedRow, 1) & "] ??", vbQuestion + vbYesNo + vbDefaultButton2, "Promociones") = vbYes Then
                    dbDatos.Execute "Update Promociones set Activa=" & IIf(.CellItemData(.SelectedRow, 5) = 0, 1, 0) & " WHERE ID=" & .CellItemData(.SelectedRow, 1)
                    CargarPrendas
                    cmdCancelar_Click
                End If
            End If
        End If
    End With
End Sub

Private Sub cmdCancelar_Click()
    cmbTipo.ListIndex = -1
    txtPorcentaje.text = ""
    txtDias.text = ""
    txtDescripcion.text = ""
    txtDescripcion.Tag = ""
    grdCatPromociones.ClearSelection
    CargarPrendas
    txtDescripcion.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Screen.MousePointer = vbHourglass
    LabelP.Enabled = False
    LabelD.Enabled = False
    txtPorcentaje.Enabled = False
    txtDias.Enabled = False
    cmdEliminar.Visible = False
    Crear_Encabezado
    CargarPrendas
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Sub Crear_Encabezado()

    With grdCatPromociones
        .AddColumn "C1", "Descripcion", ecgHdrTextALignLeft, , 150, , , , , , , CCLSortString
        .AddColumn "C2", "Tipo", ecgHdrTextALignLeft, , 100, , , , , , , CCLSortString
        .AddColumn "C3", "Porc.%", ecgHdrTextALignRight, , 85, , , , , FMoneda, , CCLSortNumeric
        .AddColumn "C4", "Días", ecgHdrTextALignRight, , 85, , , , , , , CCLSortNumeric
        .AddColumn "C5", "Activa", ecgHdrTextALignRight, , 50, , , , , , , CCLSortNumeric
    End With

End Sub

Sub CargarPrendas()
Dim rcTmp As New ADODB.Recordset

On Error GoTo Error

    rcTmp.Open "SELECT * FROM Promociones ORDER BY ID", dbDatos, adOpenForwardOnly, adLockReadOnly
    
    With grdCatPromociones
        
        .Redraw = False
        .Clear
        While Not rcTmp.EOF
            .AddRow
            .CellText(.Rows, 1) = rcTmp!Descripcion
            .CellItemData(.Rows, 1) = rcTmp!ID
            .CellTextAlign(.Rows, 1) = DT_LEFT
            
            .CellText(.Rows, 2) = IIf(rcTmp!Tipo = "P", "Porcentaje", IIf(rcTmp!Tipo = "D", "Dias", "Ambos"))
            .CellTextAlign(.Rows, 2) = DT_LEFT
            
            .CellText(.Rows, 3) = rcTmp!PorcentajeDescuento
            .CellTextAlign(.Rows, 3) = DT_RIGHT
            
            .CellText(.Rows, 4) = rcTmp!DiasDescuento
            .CellTextAlign(.Rows, 4) = DT_RIGHT
            
            .CellText(.Rows, 5) = IIf(rcTmp!Activa = 1, "Si", "No")
            .CellItemData(.Rows, 5) = rcTmp!Activa
            .CellTextAlign(.Rows, 5) = DT_LEFT
            rcTmp.MoveNext
        Wend
        
        .Redraw = True
    End With
    rcTmp.Close
    Set rcTmp = Nothing
    Exit Sub
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdCatPromociones_Click(ByVal lRow As Long, ByVal lCol As Long)
    With grdCatPromociones
        If .Rows > 0 Then
            If .SelectedRow > 0 Then
                cmdEliminar.Visible = True
                cmdEliminar.Caption = IIf(.CellItemData(lRow, 5) = 0, "     A&ctivar", "     &Desactivar")
            End If
        End If
    End With
End Sub

Private Sub grdCatPromociones_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    With grdCatPromociones
        If .Rows > 0 Then
            If .SelectedRow > 0 Then
                txtDescripcion.text = .CellText(lRow, 1)
                txtDescripcion.Tag = .CellItemData(lRow, 1)
                cmbTipo.text = .CellText(lRow, 2)
                txtPorcentaje.text = Format(.CellText(lRow, 3), FMoneda)
                txtDias.text = .CellText(lRow, 4)
                
                grdCatPromociones.ClearSelection
                txtDescripcion.SetFocus
            End If
        End If
    End With
End Sub

Private Sub txtDescripcion_GotFocus()
    cmdEliminar.Visible = False
    Seleccionar_Texto txtDescripcion
    Cambiar_Color True, txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescripcion_LostFocus()
    Cambiar_Color False, txtDescripcion
End Sub

Private Sub txtDias_GotFocus()
    Seleccionar_Texto txtDias
    Cambiar_Color True, txtDias
End Sub

Private Sub txtDias_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDias_LostFocus()
    Cambiar_Color False, txtDias
End Sub

Private Sub txtPorcentaje_GotFocus()
    Seleccionar_Texto txtPorcentaje
    Cambiar_Color True, txtPorcentaje
End Sub

Private Sub txtPorcentaje_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPorcentaje_LostFocus()
    txtPorcentaje.text = Format(txtPorcentaje.text, FMoneda)
    Cambiar_Color False, txtPorcentaje
End Sub

