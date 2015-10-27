VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmPrecios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Precios por kilataje"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrecios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7515
   ScaleWidth      =   5565
   Begin vbAcceleratorGrid6.vbalGrid grdPrecios 
      Height          =   5760
      Left            =   60
      TabIndex        =   9
      Top             =   1680
      Width           =   5475
      _ExtentX        =   9657
      _ExtentY        =   10160
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
   Begin VB.TextBox txtPrecio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   1260
      TabIndex        =   3
      Top             =   1320
      Width           =   960
   End
   Begin VB.ComboBox cmbHechura 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox cmbKilataje 
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   570
      Width           =   1575
   End
   Begin VB.ComboBox cmbTipo 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   1245
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   180
      Width           =   1575
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
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
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmPrecios.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   720
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
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmPrecios.frx":009C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdBorrar 
      Height          =   375
      Left            =   3000
      TabIndex        =   11
      Top             =   720
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Eliminar"
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
      Picture         =   "frmPrecios.frx":012D
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   120
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmPrecios.frx":01A3
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Precio:"
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
      Top             =   1320
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Hechura:"
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
      TabIndex        =   7
      Top             =   960
      Width           =   750
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Kilataje:"
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
      TabIndex        =   6
      Top             =   600
      Width           =   690
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Prenda:"
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
      TabIndex        =   5
      Top             =   240
      Width           =   1050
   End
End
Attribute VB_Name = "frmPrecios"
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

Private Sub cmbTipo_Click()
If cmbTipo.ListIndex > -1 Then
    Cargar_Combos "Descripcion", "Kilatajes", cmbKilataje, " where IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Ordenamiento"
    Cargar_Combos "Estado", "Estado", cmbHechura, " where IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex)
End If
End Sub

Private Sub cmbTipo_GotFocus()
cmbTipo.BackColor = &HC0FFFF
End Sub

Private Sub cmbTipo_KeyPress(KeyAscii As Integer)
Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipo_LostFocus()
cmbTipo.BackColor = vbWhite
End Sub

Private Sub cmdAceptar_Click()
If Completos Then
        
    If Val(txtPrecio.Tag) = 0 Then
        dbDatos.Execute "insert into PreciosKilataje (IDTipo,IDKilataje,IDHechura,Precio)values " _
                    & "(" & cmbTipo.ItemData(cmbTipo.ListIndex) & "," & cmbKilataje.ItemData(cmbKilataje.ListIndex) & "," & cmbHechura.ItemData(cmbHechura.ListIndex) & "," & txtPrecio.text & ")"
    Else
        If MsgBox("Desea guardar los cambios ??", vbQuestion + vbYesNo + vbDefaultButton2, "Configuración de precios") = vbYes Then
            dbDatos.Execute "update precioskilataje set IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex) & ",IDKilataje=" & cmbKilataje.ItemData(cmbKilataje.ListIndex) & ",IDHechura=" & cmbHechura.ItemData(cmbHechura.ListIndex) & ",Precio=" & CDbl(txtPrecio.text) & " where ID=" & Val(txtPrecio.Tag) & ""
        End If
    End If
    
    Limpiar
    CargaDatos
    cmbTipo.SetFocus
End If
End Sub

Private Sub cmdBorrar_Click()
If grdPrecios.SelectedRow > 0 Then
    If MsgBox("Desea eliminar el registro seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Configuración de precios") = vbYes Then
        dbDatos.Execute "delete from precioskilataje where ID=" & Val(grdPrecios.CellItemData(grdPrecios.SelectedRow, 4)) & ""
        grdPrecios.RemoveRow grdPrecios.SelectedRow
        grdPrecios.ClearSelection
        Limpiar
        cmbTipo.SetFocus
    End If
End If
End Sub

Private Sub cmdLimpiar_Click()
Limpiar
grdPrecios.ClearSelection
cmbTipo.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Inicializar
End Sub

Sub Inicializar()
Cargar_Combos "Descripcion", "Tipo", cmbTipo, " where Kilataje=1 And Peso=1"
CrearEncabezados
CargaDatos
CentrarForm Me, frmMDI
Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Quitar_Flat Fl
End Sub

Private Sub grdPrecios_DblClick(ByVal lRow As Long, ByVal lCol As Long)

On Error GoTo Error

With grdPrecios
    If .Rows > 0 And .SelectedRow > 0 Then
        cmbTipo.ListIndex = ComboInformacion(cmbTipo, .CellItemData(.SelectedRow, 1))
        cmbKilataje.ListIndex = ComboInformacion(cmbKilataje, .CellItemData(.SelectedRow, 2))
        cmbHechura.ListIndex = ComboInformacion(cmbHechura, .CellItemData(.SelectedRow, 3))
        txtPrecio.text = Format(grdPrecios.CellText(grdPrecios.SelectedRow, 4), "###,###,###0.00")
        txtPrecio.Tag = grdPrecios.CellItemData(grdPrecios.SelectedRow, 4)
    End If
End With

Error:
    Maneja_Error Err
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
    .AddColumn "C1", "Prenda", ecgHdrTextALignLeft, , 105, , , , , , , CCLSortString
    .AddColumn "C2", "Kilataje", ecgHdrTextALignLeft, , 72, , , , , , , CCLSortString
    .AddColumn "C3", "Hechura", ecgHdrTextALignLeft, , 80, , , , , , , CCLSortString
    .AddColumn "C4", "Precio", ecgHdrTextALignRight, , 78, , , , , "###,###,###0.00", , CCLSortNumeric
End With
End Sub

Function Completos() As Boolean
Dim rcTmp As ADODB.Recordset

On Error GoTo Error

    Completos = True

    'Tipo de prenda
    If cmbTipo.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de prenda !!", vbInformation, "Configuración de Precios"
        Completos = False
        cmbTipo.SetFocus
        Exit Function
    End If
    
    'Kilataje
    Set rcTmp = dbDatos.Execute("select Kilataje from tipo where ID=" & cmbTipo.ItemData(cmbTipo.ListIndex) & "")
    If rcTmp!Kilataje = 1 And cmbKilataje.ListIndex = -1 Then
        MsgBox "Seleccione el kilataje !!", vbInformation, "Configuración de Precios"
        Completos = False
        cmbKilataje.SetFocus
        Exit Function
    End If
    
'''''    'Hechura
'''''    Set rcTmp = dbDatos.Execute("select Hechura from tipos where ID=" & cmbTipo.ItemData(cmbTipo.ListIndex) & "")
'''''    If rcTmp!Hechura And cmbHechura.ListIndex = -1 Then
'''''        MsgBox "Seleccione la Hechura !!", vbInformation, "Configuración de Precios"
'''''        Completos = False
'''''        cmbHechura.SetFocus
'''''        Exit Function
'''''    End If
    
    'Precio
    If Val(txtPrecio.text) = 0 Then
        MsgBox "Introduzca el precio !!", vbInformation, "Configuración de Precios"
        Completos = False
        txtPrecio.SetFocus
        Exit Function
    End If

Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

Function CargaDatos()
Dim rcTmp As New ADODB.Recordset

On Error GoTo Error
    
    grdPrecios.Clear
    rcTmp.Open "SELECT PreciosKilataje.ID as IDPrecios,Tipo.Descripcion AS Tipos_Descripcion,Kilatajes.Descripcion AS Kilatajes_Descripcion, Estado.Estado AS Hechuras_Descripcion,Estado.ID as IDHechura,PreciosKilataje.Precio,PreciosKilataje.IDTipo,PreciosKilataje.IDKilataje" _
                                & " FROM PreciosKilataje Inner Join Tipo on PreciosKilataje.IDTipo=Tipo.ID Inner Join Kilatajes on Kilatajes.ID = PreciosKilataje.IDKilataje Inner Join Estado on PreciosKilataje.IDHechura=Estado.ID Order By PreciosKilataje.IDTipo,PreciosKilataje.IDKilataje,PreciosKilataje.IDHechura", dbDatos, adOpenDynamic, adLockOptimistic
    
    If Not rcTmp.BOF And Not rcTmp.EOF Then
        rcTmp.MoveFirst
        With grdPrecios
            While Not rcTmp.EOF
                .AddRow
                .CellText(.Rows, 1) = rcTmp!Tipos_Descripcion
                .CellItemData(.Rows, 1) = rcTmp!IDTipo
                .CellText(.Rows, 2) = rcTmp!Kilatajes_Descripcion
                .CellItemData(.Rows, 2) = rcTmp!IDKilataje
                .CellText(.Rows, 3) = rcTmp!Hechuras_Descripcion
                .CellItemData(.Rows, 3) = rcTmp!IDHechura
                .CellText(.Rows, 4) = rcTmp!Precio
                .CellItemData(.Rows, 4) = rcTmp!IDPrecios
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                
            rcTmp.MoveNext
            Wend
        End With
    End If
    rcTmp.Close
    
Error:
    Maneja_Error Err
    Set rcTmp = Nothing
End Function

Sub Limpiar()
cmbTipo.ListIndex = -1
cmbKilataje.ListIndex = -1
cmbHechura.ListIndex = -1
txtPrecio.text = ""
txtPrecio.Tag = ""
End Sub
