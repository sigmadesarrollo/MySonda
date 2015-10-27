VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "VBALGR~2.OCX"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmConfiguracionINI 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion Archivo INI"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmConfiguracionINI.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   5865
   Begin VB.TextBox txtY 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   4560
      TabIndex        =   4
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox txtX 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1215
   End
   Begin VB.ComboBox cboSeccion 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   240
      Width           =   2340
   End
   Begin vbAcceleratorGrid6.vbalGrid grdConfiguracion 
      Height          =   7695
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   13573
      GridLines       =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   240
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
      Picture         =   "frmConfiguracionINI.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   3600
      TabIndex        =   6
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
      Picture         =   "frmConfiguracionINI.frx":055E
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Sección:"
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
      TabIndex        =   1
      Top             =   240
      Width           =   1005
   End
End
Attribute VB_Name = "frmConfiguracionINI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cboSeccion_Click()
    Carga_Seccion cboSeccion.List(cboSeccion.ListIndex)
End Sub

Private Sub cboSeccion_GotFocus()
    Cambiar_Color True, cboSeccion
End Sub

Private Sub cboSeccion_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cboSeccion_LostFocus()
    Cambiar_Color False, cboSeccion
End Sub

Private Sub cmdAceptar_Click()
    Dim lRow As Long
    Dim Seccion As String, KeyX As String, KeyY As String, KeyValorX As String, KeyValorY As String
    
    Seccion = cboSeccion.List(cboSeccion.ListIndex)
    
    If Len(Seccion) <> 0 And grdConfiguracion.Rows > 0 Then
    
        If MsgBox("Desea guardar los cambios realizados en la impresión ??", vbQuestion + vbYesNo + vbDefaultButton1, "Configuración INI") = vbYes Then
        
            For lRow = 1 To grdConfiguracion.Rows
                KeyX = grdConfiguracion.CellText(lRow, 1) + "X"
                KeyY = grdConfiguracion.CellText(lRow, 1) + "Y"
                
                KeyValorX = grdConfiguracion.CellText(lRow, 2)
                KeyValorY = grdConfiguracion.CellText(lRow, 3)
                
                Graba_Valor Seccion, KeyX, KeyValorX
                Graba_Valor Seccion, KeyY, KeyValorY
            Next
            
            MsgBox "Configuracion grabada con éxito!!", vbInformation, "Configuración INI"
        End If
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Carga_Secciones
    Carga_Header
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Carga_Secciones()
    Dim Secciones As String
    Dim SeccionesArray() As String, Pos As Integer
    
    Secciones = Regresa_Secciones
    SeccionesArray = Split(Secciones, vbNullChar)
    
    For Pos = 0 To (UBound(SeccionesArray) - 1)
        cboSeccion.AddItem SeccionesArray(Pos)
    Next Pos
    
    cboSeccion.ListIndex = 0
End Sub

Private Sub Carga_Header()
    With grdConfiguracion
        .AddColumn "K1", "Parametro", ecgHdrTextALignCentre, , 200, , , , , , , CCLSortString
        .AddColumn "K2", "Valor X", ecgHdrTextALignCentre, , 75, , , , , , , CCLSortString
        .AddColumn "K3", "Valor Y", ecgHdrTextALignCentre, , 75, , , , , , , CCLSortString
    End With
End Sub

Private Sub Carga_Seccion(Seccion As String)
Dim Valores As String
Dim ValoresArray() As String, Pos As Integer
Dim KeyNombre As String, KeyValor As String
    
    grdConfiguracion.Redraw = False
    grdConfiguracion.Clear
    
    txtX.Visible = False
    txtY.Visible = False
    
    Valores = Regresa_Seccion_Valores(Seccion)
    ValoresArray = Split(Valores, vbNullChar)
    
    For Pos = 0 To (UBound(ValoresArray) - 1)
        KeyNombre = Regresa_Nombre_Key(ValoresArray(Pos))
        If StrComp(KeyNombre, "NULL") <> 0 Then
            Agrega_Key Seccion, KeyNombre
        End If
    Next Pos
    
    grdConfiguracion.Redraw = True
End Sub

Function Regresa_Nombre_Key(Key As String) As String
    Dim IsCoordenada As Long
    
    IsCoordenada = InStr(1, Key, "=", vbTextCompare)
    
    If IsCoordenada > 0 Then
        If (StrComp(Mid(Key, IsCoordenada - 1, 1), "X") = 0 Or StrComp(Mid(Key, IsCoordenada - 1, 1), "Y") = 0) Then
            Regresa_Nombre_Key = Mid(Key, 1, IsCoordenada - 2)
        Else
            Regresa_Nombre_Key = "NULL"
        End If
    Else
        Regresa_Nombre_Key = "NULL"
    End If
    
End Function

Public Sub Agrega_Key(Seccion As String, KeyNombre As String)
Dim KeyValorX As String, KeyValorY As String
    
    If Busca_Key(KeyNombre) = False Then
        
        KeyValorX = Regresa_Valor(Seccion, KeyNombre + "X", 0)
        KeyValorY = Regresa_Valor(Seccion, KeyNombre + "Y", 0)
        
        If Len(KeyValorX) > 0 And Len(KeyValorY) > 0 Then
            
            grdConfiguracion.AddRow
            grdConfiguracion.CellText(grdConfiguracion.Rows, 1) = KeyNombre
            
            grdConfiguracion.CellText(grdConfiguracion.Rows, 2) = KeyValorX
            grdConfiguracion.CellTextAlign(grdConfiguracion.Rows, 2) = DT_RIGHT
            
            grdConfiguracion.CellText(grdConfiguracion.Rows, 3) = KeyValorY
            grdConfiguracion.CellTextAlign(grdConfiguracion.Rows, 3) = DT_RIGHT
            
            Colorea grdConfiguracion, grdConfiguracion.Rows, IIf(grdConfiguracion.Rows Mod 2 > 0, RGB(236, 252, 222), RGB(255, 255, 255))
            
        End If
        
    End If
    
End Sub

Function Busca_Key(KeyNombre As String) As Boolean
Dim Parametro As String, Index As Integer
    
    For Index = 1 To grdConfiguracion.Rows
        Parametro = grdConfiguracion.CellText(Index, 1)
        
        If StrComp(Parametro, KeyNombre, 1) = 0 Then
            Busca_Key = True
            Exit For
        End If
        
    Next Index
End Function

Public Function Solo_Numero(Codigo As Integer, Optional Opcion As Integer = 0) As Integer
    If (Codigo >= vbKey0 And Codigo <= vbKey9) Or Codigo = vbKeyBack Or Codigo = vbKeyReturn Or Codigo = 46 Then
        Solo_Numero = Codigo
    Else
        Solo_Numero = 0
    End If
End Function

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdConfiguracion_KeyPress(KeyAscii As Integer)
    Dim lRow As Long, lCol As Long, lLeft As Long, lTop As Long, lWidth As Long, lHeight As Long, sText As String
    
    lRow = grdConfiguracion.SelectedRow
    lCol = grdConfiguracion.SelectedCol
    
    grdConfiguracion.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
    
    If KeyAscii = vbKeyReturn Or (KeyAscii >= vbKey0 And KeyAscii <= vbKey9) Then
    
        txtX.Visible = False
        txtY.Visible = False
        
        sText = grdConfiguracion.CellFormattedText(lRow, lCol)
            
        If lCol = 2 Then
        
            If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
                txtX.text = Chr(KeyAscii)
                txtX.SelStart = 1
            Else
                txtX.text = sText
                txtX.SelStart = 0
                txtX.SelLength = Len(sText)
            End If
        
            Set txtX.Font = grdConfiguracion.CellFont(lRow, lCol)
            If grdConfiguracion.CellBackColor(lRow, lCol) = -1 Then
                txtX.BackColor = grdConfiguracion.BackColor
            Else
                txtX.BackColor = grdConfiguracion.CellBackColor(lRow, lCol)
            End If
            
            txtX.Move lLeft + grdConfiguracion.Left + 10, _
                lTop + grdConfiguracion.Top, _
                lWidth, lHeight
                
            txtX.Visible = True
            txtX.ZOrder
            txtX.SetFocus
        End If
        
        If lCol = 3 Then
            If KeyAscii >= vbKey0 And KeyAscii <= vbKey9 Then
                txtY.text = Chr(KeyAscii)
                txtY.SelStart = 1
            Else
                txtY.text = sText
                txtY.SelStart = 0
                txtY.SelLength = Len(sText)
            End If
        
            Set txtY.Font = grdConfiguracion.CellFont(lRow, lCol)
            If grdConfiguracion.CellBackColor(lRow, lCol) = -1 Then
                txtY.BackColor = grdConfiguracion.BackColor
            Else
                txtY.BackColor = grdConfiguracion.CellBackColor(lRow, lCol)
            End If
            
            txtY.Move lLeft + grdConfiguracion.Left + 10, _
                lTop + grdConfiguracion.Top, _
                lWidth, lHeight
                
            txtY.Visible = True
            txtY.ZOrder
            txtY.SetFocus
        End If
    
    End If
End Sub

Private Sub txtX_KeyPress(KeyAscii As Integer)

    If (KeyAscii = vbKeyReturn) Then
    
        If txtX.text <> "" Then
            grdConfiguracion.CellText(grdConfiguracion.SelectedRow, 2) = txtX.text
            grdConfiguracion.CellTextAlign(grdConfiguracion.SelectedRow, 2) = DT_RIGHT
        End If
        
        txtX.Visible = False
        grdConfiguracion.SetFocus
    Else
        KeyAscii = Solo_Numero(KeyAscii, 1)
    End If
    
End Sub

Private Sub txtY_KeyPress(KeyAscii As Integer)

    If (KeyAscii = vbKeyReturn) Then
    
        If txtY.text <> "" Then
            grdConfiguracion.CellText(grdConfiguracion.SelectedRow, 3) = txtY.text
            grdConfiguracion.CellTextAlign(grdConfiguracion.SelectedRow, 3) = DT_RIGHT
        End If
        
        txtY.Visible = False
        grdConfiguracion.SetFocus
    Else
        KeyAscii = Solo_Numero(KeyAscii, 1)
    End If
    
End Sub

Private Sub txtX_LostFocus()
    txtX.Visible = False
End Sub

Private Sub txtY_LostFocus()
    txtY.Visible = False
End Sub
