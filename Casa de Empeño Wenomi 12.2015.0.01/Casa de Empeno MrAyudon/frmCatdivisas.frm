VERSION 5.00
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmCatdivisas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Catálogo de divisas"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6495
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCatdivisas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5040
   ScaleWidth      =   6495
   Begin VB.TextBox txtNum 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      MaxLength       =   8
      TabIndex        =   0
      Top             =   270
      Width           =   735
   End
   Begin vbAcceleratorGrid6.vbalGrid grdDivisas 
      Height          =   3540
      Left            =   0
      TabIndex        =   3
      Top             =   1005
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   6244
      RowMode         =   -1  'True
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
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
      Begin VB.TextBox txtLimite 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.TextBox txtDivisa 
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1320
      MaxLength       =   50
      TabIndex        =   1
      Top             =   630
      Width           =   3135
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   390
      Left            =   5340
      TabIndex        =   6
      Top             =   4605
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
      Picture         =   "frmCatdivisas.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   390
      Left            =   4560
      TabIndex        =   7
      Top             =   525
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
      Picture         =   "frmCatdivisas.frx":055E
   End
   Begin vbalIml6.vbalImageList lstIcons 
      Left            =   360
      Top             =   4560
      _ExtentX        =   953
      _ExtentY        =   953
      Size            =   2296
      Images          =   "frmCatdivisas.frx":0AB0
      Version         =   131072
      KeyCount        =   2
      Keys            =   "ÿ"
   End
   Begin DevPowerFlatBttn.FlatBttn cmdEliminar 
      Height          =   390
      Left            =   4200
      TabIndex        =   8
      Top             =   4605
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   688
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
      Picture         =   "frmCatdivisas.frx":13C8
      PictureDisabled =   "frmCatdivisas.frx":191A
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Num. divisa:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
End
Attribute VB_Name = "frmCatdivisas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
Dim rcConsulta As New ADODB.Recordset

    If Trim(txtNum.Text) = "" Then
        
        MsgBox "Introduzca la clave !!", vbInformation, "Catálogo de Divisas"
        txtNum.SetFocus
        
    ElseIf Trim(txtDivisa.Text) = "" Then
        
        MsgBox "Introduzca la descripción !!", vbInformation, "Catálogo de Divisas"
        txtDivisa.SetFocus
    
    Else
        rcConsulta.Open "SELECT Clave FROM monedas WHERE clave=" & Val(txtNum.Text), dbDatos, adOpenForwardOnly, adLockOptimistic
        If rcConsulta.BOF Or rcConsulta.EOF Then
            
            dbDatos.Execute "INSERT INTO monedas (Clave,Descripcion,Maximo) VALUES (" & _
                            Val(txtNum.Text) & ",'" & Trim(txtDivisa.Text) & "',0)"
            grdDivisas.Clear False
            txtDivisa.Text = ""
            txtNum.Text = ""
            Cargar_Datos
            txtNum.SetFocus
        Else
            
            MsgBox "La clave que desea registrar ya existe !!", vbCritical, "Catálogo de divisas"
            txtNum.SetFocus
        End If
        rcConsulta.Close
        Set rcConsulta = Nothing
    End If

End Sub

Private Sub cmdEliminar_Click()

    If grdDivisas.SelectedRow > 0 And grdDivisas.Rows > 0 Then
        
        If MsgBox("Desea eliminar la divisa seleccionada ??", vbQuestion + vbYesNo + vbDefaultButton2, "Catálogo de divisas") = vbYes Then
            
            dbDatos.Execute "DELETE FROM monedas WHERE ID=" & grdDivisas.CellItemData(grdDivisas.SelectedRow, 1)
            grdDivisas.Clear False
            Cargar_Datos
            txtNum.SetFocus
        End If
    
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Crear_Encabezados
    Cargar_Datos
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Sub Crear_Encabezados()

    With grdDivisas
        .ImageList = lstIcons
        .AddColumn "K2", "Clave", ecgHdrTextALignCentre, , 70, , , , , , , CCLSortNumeric
        .AddColumn "K3", "Divisa", ecgHdrTextALignLeft, , 245, , , , , , , CCLSortString
        .AddColumn "K4", "Límite", ecgHdrTextALignRight, , 100, , , , , FMoneda, , CCLSortNumeric
    End With

End Sub

Sub Cargar_Datos()
Dim i As Integer
Dim rcConsulta As New ADODB.Recordset

    rcConsulta.Open "SELECT * FROM Monedas WHERE Descripcion<>'Moneda Nacional' ORDER BY Clave", dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then
        
        i = 1
        With grdDivisas
            rcConsulta.MoveFirst
            .Redraw = False
            While Not rcConsulta.EOF
                
                .AddRow
                
                .CellText(i, 1) = rcConsulta!clave
                .CellIcon(i, 1) = IIf(rcConsulta!Defoult = 1, lstIcons.ItemIndex(1), lstIcons.ItemIndex(2))
                .CellItemData(i, 1) = rcConsulta!ID
                .CellTextAlign(i, 1) = DT_CENTER
                                
                .CellText(i, 2) = rcConsulta!Descripcion
                
                .CellText(i, 3) = rcConsulta!Maximo
                .CellTextAlign(i, 3) = DT_RIGHT
            
            i = i + 1
            rcConsulta.MoveNext
            Wend
            .Redraw = True
        End With

    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdDivisas_Click(ByVal lRow As Long, ByVal lCol As Long)
Dim i As Integer

    If lCol = 1 And lRow > 0 Then
        
        If grdDivisas.CellIcon(lRow, lCol) = lstIcons.ItemIndex(2) Then
            
            MarcaDefault
            grdDivisas.CellIcon(lRow, lCol) = lstIcons.ItemIndex(1)
            dbDatos.Execute "UPDATE monedas SET Defoult=1 WHERE ID=" & grdDivisas.CellItemData(lRow, lCol)
        End If
    
    End If
End Sub

Private Sub grdDivisas_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
    Dim lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, sText As String, obj As Object
         
    If lCol = 1 Or lCol = 2 Then: txtLimite.Visible = False: Exit Sub

    Select Case lCol

        Case 3: Set obj = txtLimite
    End Select
   
    grdDivisas.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight

    If Not IsMissing(grdDivisas.CellText(lRow, lCol)) Then
        sText = IIf(Not IsNull(grdDivisas.CellText(lRow, lCol)), grdDivisas.CellText(lRow, lCol), "")
    Else
        sText = ""
    End If
   
    obj.Alignment = vbRightJustify
         
    If (iKeyAscii > 13) Then
        sText = Chr$(iKeyAscii) & sText
        obj.Text = sText
        obj.SelStart = 1
        obj.SelLength = Len(sText)
    Else
        obj.Text = sText
        obj.SelStart = 0
        obj.SelLength = Len(sText)
    End If

    obj.Move lLeft + 40, lTop + 25, lWidth - 60, lHeight - 50

    obj.Visible = True
    obj.ZOrder
    obj.SetFocus
End Sub

Private Sub txtDivisa_GotFocus()
    Seleccionar_Texto txtDivisa
    Cambiar_Color True, txtDivisa
End Sub

Private Sub txtDivisa_KeyPress(KeyAscii As Integer)
    KeyAscii = mayusculas(KeyAscii)
    If KeyAscii = vbKeyReturn Then cmdAceptar_Click
End Sub

Private Sub txtDivisa_LostFocus()
    Cambiar_Color False, txtDivisa
End Sub

Private Sub txtLimite_GotFocus()
    Seleccionar_Texto txtLimite
    Cambiar_Color True, txtLimite
End Sub

Private Sub txtLimite_KeyPress(KeyAscii As Integer)
    Dim Maximo As Double

    Maximo = 0
    KeyAscii = Solo_Numeros(KeyAscii, 1)

    If KeyAscii = vbKeyReturn Then
        If txtLimite.Text <> "" Then
            Maximo = txtLimite.Text
            grdDivisas.CellText(grdDivisas.SelectedRow, 3) = Format(CDbl(txtLimite.Text), "##,###")
            grdDivisas.CellTextAlign(grdDivisas.SelectedRow, 3) = DT_RIGHT
            dbDatos.Execute "update monedas set maximo=" & Maximo & " where id=" & grdDivisas.CellItemData(grdDivisas.SelectedRow, 1) & ""
            txtLimite.Visible = False
        Else
            txtLimite.Visible = False
        End If
    End If

End Sub

Private Sub txtLimite_LostFocus()
    Cambiar_Color False, txtLimite
End Sub

Private Sub txtNum_GotFocus()
    Seleccionar_Texto txtNum
    Cambiar_Color True, txtNum
End Sub

Private Sub txtNum_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNum_LostFocus()
    Cambiar_Color False, txtNum
End Sub

Sub MarcaDefault()
Dim i As Integer
    
    For i = 1 To grdDivisas.Rows
    
        grdDivisas.CellIcon(i, 1) = lstIcons.ItemIndex(2)
        dbDatos.Execute "UPDATE monedas SET Defoult=0 WHERE ID=" & grdDivisas.CellItemData(i, 1)
    Next i

End Sub
