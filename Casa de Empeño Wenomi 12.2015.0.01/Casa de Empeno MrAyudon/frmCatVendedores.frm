VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatVendedores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de vendedores"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8340
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatVendedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4755
   ScaleWidth      =   8340
   Begin vbAcceleratorGrid6.vbalGrid grdVendedores 
      Height          =   3405
      Left            =   15
      TabIndex        =   6
      Top             =   1305
      Width           =   8280
      _ExtentX        =   14605
      _ExtentY        =   6006
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
      DisableIcons    =   -1  'True
   End
   Begin VB.TextBox txtMeta 
      Alignment       =   1  'Right Justify
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
      Height          =   255
      Left            =   1200
      MaxLength       =   12
      TabIndex        =   2
      Top             =   960
      Width           =   1455
   End
   Begin VB.TextBox txtApellidos 
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
      Height          =   255
      Left            =   1200
      MaxLength       =   100
      TabIndex        =   1
      Top             =   600
      Width           =   4575
   End
   Begin VB.TextBox txtNombre 
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
      Height          =   255
      Left            =   1200
      MaxLength       =   50
      TabIndex        =   0
      Top             =   240
      Width           =   4575
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   795
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
      Picture         =   "frmCatVendedores.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   7080
      TabIndex        =   8
      Top             =   795
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
      Picture         =   "frmCatVendedores.frx":0110
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   6000
      TabIndex        =   9
      Top             =   240
      Width           =   1035
      _ExtentX        =   1826
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
      Picture         =   "frmCatVendedores.frx":0662
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   375
      Left            =   7080
      TabIndex        =   10
      Top             =   240
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "     &Eliminar"
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
      Picture         =   "frmCatVendedores.frx":0BB4
      PictureDisabled =   "frmCatVendedores.frx":1106
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Meta:"
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
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   555
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos:"
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
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "frmCatVendedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
Dim IDVendedor As Long
    
    If Requeridos Then
        
        If Val(txtNombre.Tag) = 0 Then
            
            dbDatos.Execute "INSERT INTO vendedores (Nombre,Apellidos,Meta) VALUES ('" & _
                            Trim(txtNombre.Text) & "','" & Trim(txtApellidos.Text) & "'," & CDbl(txtMeta.Text) & ")"
                
            IDVendedor = SacaValor("vendedores", "MAX(ID)")
            
            With grdVendedores
                                        
                .Redraw = False
                
                .AddRow
                .CellText(.Rows, 1) = Trim(txtNombre.Text)
                .CellItemData(.Rows, 1) = IDVendedor
                .CellText(.Rows, 2) = Trim(txtApellidos.Text)
                .CellText(.Rows, 3) = CDbl(txtMeta.Text)
                .CellTextAlign(.Rows, 3) = DT_RIGHT
                
                Colorea grdVendedores, .Rows, IIf(.Rows Mod 2 <> 0, Trim(RGB(242, 254, 255)), Trim(RGB(255, 255, 255)))
                
                .Redraw = True
            
            End With
        
        ElseIf Val(txtNombre.Tag) > 0 Then
            
            If MsgBox("Desea guardar los cambios realizados ??", vbQuestion + vbYesNo + vbDefaultButton1, "Catálogo de vendedores") = vbYes Then
                
                dbDatos.Execute "UPDATE vendedores SET Nombre='" & Trim(txtNombre.Text) & "',Apellidos='" & Trim(txtApellidos.Text) & "',Meta=" & CDbl(txtMeta.Text) & " WHERE ID=" & Val(txtNombre.Tag)
                Cargar_Datos
                
            End If
            
        End If
        
        Limpiar
        txtNombre.SetFocus
    End If
    
End Sub

Private Sub cmdEliminar_Click()
Dim i As Integer

    With grdVendedores
        
        If .Rows > 0 Then
            
            If .SelectedRow > 0 Then
                
                If MsgBox("Desea eliminar el vendedor seleccionado ??", vbQuestion + vbYesNo + vbDefaultButton2, "Catálogo de vendedores") = vbYes Then
                    
                    dbDatos.Execute "DELETE FROM vendedores WHERE ID=" & Val(.CellItemData(.SelectedRow, 1))
                    .Redraw = False
                    .RemoveRow .SelectedRow
                    For i = 1 To .Rows
                        
                        Colorea grdVendedores, CLng(i), IIf(i Mod 2 <> 0, Trim(RGB(242, 254, 255)), Trim(RGB(255, 255, 255)))
                    Next i
                    .Redraw = True
                End If
                .ClearSelection
            End If
            
        End If
        
    End With
    
    txtNombre.SetFocus
    
End Sub

Private Sub cmdLimpiar_Click()
    Limpiar
    txtNombre.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Crear_Encabezados
    Cargar_Datos
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Sub Crear_Encabezados()
    
    With grdVendedores
        .AddColumn "C1", "Nombre", ecgHdrTextALignLeft, , 190, , , , , , , CCLSortString
        .AddColumn "C2", "Apellidos", ecgHdrTextALignLeft, , 230, , , , , , , CCLSortString
        .AddColumn "C3", "Meta", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortNumeric
    End With
    
End Sub

Sub Cargar_Datos()
Dim rcVendedores As New ADODB.Recordset
    
On Error GoTo error

    rcVendedores.Open "SELECT * FROM vendedores ORDER BY ID", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcVendedores.BOF And Not rcVendedores.EOF Then
        
        rcVendedores.MoveFirst
        With grdVendedores
            
            .Redraw = False
            .Clear
            While Not rcVendedores.EOF
                
                .AddRow
                .CellText(.Rows, 1) = rcVendedores!Nombre
                .CellItemData(.Rows, 1) = rcVendedores!ID
                .CellText(.Rows, 2) = rcVendedores!Apellidos
                .CellText(.Rows, 3) = rcVendedores!Meta
                .CellTextAlign(.Rows, 3) = DT_RIGHT
                
                Colorea grdVendedores, .Rows, IIf(.Rows Mod 2 <> 0, Trim(RGB(242, 254, 255)), Trim(RGB(255, 255, 255)))
                
            rcVendedores.MoveNext
            Wend

            .Redraw = True
            
        End With
        
    End If
    rcVendedores.Close
    Set rcVendedores = Nothing
    Exit Sub

error:
    Maneja_Error Err
    Set rcVendedores = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdVendedores_DblClick(ByVal lRow As Long, ByVal lCol As Long)
    
    With grdVendedores
        
        If .Rows > 0 Then
            
            If .SelectedRow > 0 Then
                
                txtNombre.Text = .CellText(.SelectedRow, 1)
                txtNombre.Tag = Val(.CellItemData(.SelectedRow, 1))
                txtApellidos.Text = .CellText(.SelectedRow, 2)
                txtMeta.Text = Format(.CellText(.SelectedRow, 3), FMoneda)
                .ClearSelection
                
            End If
            
        End If
        
    End With
    
End Sub

Private Sub txtApellidos_GotFocus()
    Seleccionar_Texto txtApellidos
    Cambiar_Color True, txtApellidos
End Sub

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
    KeyAscii = mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtApellidos_LostFocus()
    Cambiar_Color False, txtApellidos
End Sub

Private Sub txtMeta_GotFocus()
    Seleccionar_Texto txtMeta
    Cambiar_Color True, txtMeta
End Sub

Private Sub txtMeta_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMeta_LostFocus()
    txtMeta.Text = Format(txtMeta, FMoneda)
    Cambiar_Color False, txtMeta
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Texto txtNombre
    Cambiar_Color True, txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    KeyAscii = mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNombre_LostFocus()
    Cambiar_Color False, txtNombre
End Sub

Function Requeridos() As Boolean

    Requeridos = True
    
    If Trim(txtNombre.Text) = "" Then
        MsgBox "Introduzca el nombre !!", vbInformation, "Catálogo de vendedores"
        Requeridos = False
        txtNombre.SetFocus
        Exit Function
    End If
    
    If Trim(txtApellidos.Text) = "" Then
        MsgBox "Introduzca el apellido !!", vbInformation, "Catálogo de vendedores"
        Requeridos = False
        txtApellidos.SetFocus
        Exit Function
    End If
    
    If Trim(txtMeta.Text) = "" Then
        MsgBox "Introduzca la meta !!", vbInformation, "Catálogo de vendedores"
        Requeridos = False
        txtMeta.SetFocus
        Exit Function
    End If
    
End Function

Sub Limpiar()
    txtNombre.Text = ""
    txtNombre.Tag = ""
    txtApellidos.Text = ""
    txtMeta.Text = ""
    grdVendedores.ClearSelection
End Sub
