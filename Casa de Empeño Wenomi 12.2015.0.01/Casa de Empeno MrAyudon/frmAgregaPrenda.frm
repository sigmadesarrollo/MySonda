VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmAgregaPrenda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agregar Prenda"
   ClientHeight    =   3315
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6420
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAgregaPrenda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3315
   ScaleWidth      =   6420
   StartUpPosition =   3  'Windows Default
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
      ItemData        =   "frmAgregaPrenda.frx":000C
      Left            =   1800
      List            =   "frmAgregaPrenda.frx":000E
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1920
   End
   Begin VB.ComboBox cmbMarca 
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
      ItemData        =   "frmAgregaPrenda.frx":0010
      Left            =   4635
      List            =   "frmAgregaPrenda.frx":0012
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   135
      Width           =   1725
   End
   Begin VB.ComboBox cmbFamilia 
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
      ItemData        =   "frmAgregaPrenda.frx":0014
      Left            =   1800
      List            =   "frmAgregaPrenda.frx":0016
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   465
      Width           =   1920
   End
   Begin VB.TextBox txtModelo 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4665
      TabIndex        =   3
      Top             =   495
      Width           =   1695
   End
   Begin VB.TextBox txtCaracteristicas 
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
      Height          =   465
      Left            =   1830
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   1995
      Width           =   4530
   End
   Begin VB.TextBox txtFunciones 
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
      Height          =   465
      Left            =   1830
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   4530
   End
   Begin VB.TextBox txtDescripcion 
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
      Height          =   465
      Left            =   1830
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   900
      Width           =   4530
   End
   Begin VB.TextBox txtMinimo 
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
      Height          =   225
      Left            =   1830
      TabIndex        =   7
      Top             =   2535
      Width           =   1185
   End
   Begin VB.TextBox txtMaximo 
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
      Height          =   225
      Left            =   4080
      TabIndex        =   8
      Top             =   2535
      Width           =   1185
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5295
      TabIndex        =   10
      Top             =   2850
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
      Picture         =   "frmAgregaPrenda.frx":0018
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   4185
      TabIndex        =   9
      Top             =   2850
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
      Picture         =   "frmAgregaPrenda.frx":056A
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
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
      Left            =   90
      TabIndex        =   19
      Top             =   180
      Width           =   495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Familia:"
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
      Index           =   4
      Left            =   90
      TabIndex        =   18
      Top             =   525
      Width           =   795
   End
   Begin VB.Label Label28 
      AutoSize        =   -1  'True
      Caption         =   "Marca:"
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
      Left            =   3840
      TabIndex        =   17
      Top             =   180
      Width           =   675
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Modelo:"
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
      Index           =   3
      Left            =   3840
      TabIndex        =   16
      Top             =   525
      Width           =   780
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Características:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   15
      Top             =   1995
      Width           =   1725
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Funciones:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   1440
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Descripción:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   900
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Mínimo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   12
      Top             =   2535
      Width           =   855
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Máximo:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3120
      TabIndex        =   11
      Top             =   2535
      Width           =   930
   End
End
Attribute VB_Name = "frmAgregaPrenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim IDPrenda As Long

Private Sub cmbFamilia_Click()
Dim IDFamilia As Integer

    If cmbFamilia.text = "[1. AGREGAR]" And cmbTipo.ListIndex > -1 Then
        
        IDFamilia = 0
        IDFamilia = frmAgregaFamilia.Mostrar(cmbTipo.ItemData(cmbTipo.ListIndex))
        If IDFamilia > 0 Then
            
            cmbFamilia.Clear
            cmbFamilia.AddItem "[1. AGREGAR FAMILIA]"
            Cargar_Combos "tipoprenda.Descripcion", "tipoprenda INNER JOIN tipo ON tipoprenda.IDTipo=tipo.ID", cmbFamilia, " WHERE tipo.Kilataje=0 AND tipo.Peso=0", , False, "tipoprenda.ID"
            cmbFamilia.ListIndex = ComboInformacion(cmbFamilia, IDFamilia)
        
        Else
            
            cmbFamilia.ListIndex = -1
        End If
    
    End If
End Sub

Private Sub cmbFamilia_GotFocus()
    Cambiar_Color True, cmbFamilia
End Sub

Private Sub cmbFamilia_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbFamilia_LostFocus()
    Cambiar_Color False, cmbFamilia
End Sub

Private Sub cmbMarca_Click()
Dim IDMarca As Integer

    If cmbMarca.text = "[1. AGREGAR]" Then
        
        IDMarca = 0
        IDMarca = frmAgregaMarca.Mostrar()
        If IDMarca > 0 Then
            
            cmbMarca.Clear
            cmbMarca.AddItem "[1. AGREGAR MARCA]"
            Cargar_Combos "Descripcion", "marcas", cmbMarca, "", "Descripcion", False
            cmbMarca.ListIndex = ComboInformacion(cmbMarca, IDMarca)
        
        Else
            
            cmbMarca.ListIndex = -1
        End If
    
    End If
End Sub

Private Sub cmbMarca_GotFocus()
    Cambiar_Color True, cmbMarca
End Sub

Private Sub cmbMarca_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbMarca_LostFocus()
    Cambiar_Color False, cmbMarca
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
    
    If Completos Then
    
        dbDatos.Execute "INSERT INTO prendaselec (IDTipo,IDMarca,IDFamilia,Modelo,Minimo,Maximo,Funciones,Caracteristicas) VALUES (" & _
                        cmbTipo.ItemData(cmbTipo.ListIndex) & "," & cmbMarca.ItemData(cmbMarca.ListIndex) & "," & cmbFamilia.ItemData(cmbFamilia.ListIndex) & ",'" & Trim(txtModelo.text) & "'," & CDbl(txtMinimo.text) & "," & CDbl(txtMaximo.text) & ",'" & Trim(txtFunciones.text) & "','" & Trim(txtCaracteristicas.text) & "')"
        
        IDPrenda = SacaValor("prendaselec", "MAX(ID)")
        Unload Me
        
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    Screen.MousePointer = vbHourglass
    
    Cargar_Combos "Descripcion", "tipo", cmbTipo, " WHERE tipo.Kilataje=0 AND Peso=0", "Ordenamiento"
    
    cmbMarca.AddItem "[1. AGREGAR]"
    Cargar_Combos "Descripcion", "marcas", cmbMarca, "", "Descripcion", False
    
    cmbFamilia.AddItem "[1. AGREGAR]"
    Cargar_Combos "tipoprenda.Descripcion", "tipoprenda INNER JOIN tipo ON tipoprenda.IDTipo=tipo.ID", cmbFamilia, " WHERE tipo.Kilataje=0 AND tipo.Peso=0", , False, "tipoprenda.ID"

    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

Public Function Mostrar(Tipo As Integer, Optional Familia As Integer, Optional Marca As Integer)
    IDPrenda = -1
    cmbTipo.ListIndex = ComboInformacion(cmbTipo, Tipo)
    cmbFamilia.ListIndex = ComboInformacion(cmbFamilia, IIf(Familia = 0, -1, Familia))
    cmbMarca.ListIndex = ComboInformacion(cmbMarca, IIf(Marca = 0, -1, Marca))
    Me.Show vbModal
    frmEmpeño.MostrarCatPrendas IDPrenda
End Function

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtCaracteristicas_GotFocus()
    Seleccionar_Texto txtCaracteristicas
    Cambiar_Color True, txtCaracteristicas
End Sub

Private Sub txtCaracteristicas_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCaracteristicas_LostFocus()
    Cambiar_Color False, txtCaracteristicas
End Sub

Private Sub txtDescripcion_GotFocus()
    Seleccionar_Texto txtDescripcion
    Cambiar_Color True, txtDescripcion
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDescripcion_LostFocus()
    Cambiar_Color False, txtDescripcion
End Sub

Private Sub txtFunciones_GotFocus()
    Seleccionar_Texto txtFunciones
    Cambiar_Color True, txtFunciones
End Sub

Private Sub txtFunciones_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFunciones_LostFocus()
    Cambiar_Color False, txtFunciones
End Sub

Private Sub txtMinimo_GotFocus()
    Seleccionar_Texto txtMinimo
    Cambiar_Color True, txtMinimo
End Sub

Private Sub txtMinimo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMinimo_LostFocus()
    Cambiar_Color False, txtMinimo
End Sub

Private Sub txtMaximo_GotFocus()
    Seleccionar_Texto txtMaximo
    Cambiar_Color True, txtMaximo
End Sub

Private Sub txtMaximo_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtMaximo_LostFocus()
    Cambiar_Color False, txtMaximo
End Sub

Private Sub txtModelo_GotFocus()
    Seleccionar_Texto txtModelo
    Cambiar_Color True, txtModelo
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtModelo_LostFocus()
    Cambiar_Color False, txtModelo
End Sub

Function Completos() As Boolean

    Completos = True
    
    If cmbTipo.ListIndex = -1 Then
        MsgBox "Seleccione el tipo de prenda !!", vbInformation, "Agregar Prenda"
        Completos = False
        cmbTipo.SetFocus
        Exit Function
    End If
    
    If cmbMarca.ListIndex = -1 Then
        MsgBox "Seleccione la marca de la prenda !!", vbInformation, "Agregar Prenda"
        Completos = False
        cmbMarca.SetFocus
        Exit Function
    End If
    
    If cmbFamilia.ListIndex = -1 Then
        MsgBox "Seleccione la familia de la prenda !!", vbInformation, "Agregar Prenda"
        Completos = False
        cmbFamilia.SetFocus
        Exit Function
    End If
    
    If Trim(txtModelo.text) = "" Then
        MsgBox "Introduzca el modelo de la prenda !!", vbInformation, "Agregar Prenda"
        Completos = False
        txtModelo.SetFocus
        Exit Function
    End If
    
    If Trim(txtDescripcion.text) = "" Then
        MsgBox "Introduzca la descripción de la prenda !!", vbInformation, "Agregar Prenda"
        Completos = False
        txtDescripcion.SetFocus
        Exit Function
    End If
    
    If Trim(txtFunciones.text) = "" Then
        MsgBox "Introduzca las funciones de la prenda !!", vbInformation, "Agregar Prenda"
        Completos = False
        txtFunciones.SetFocus
        Exit Function
    End If
    
    If Trim(txtCaracteristicas.text) = "" Then
        MsgBox "Introduzca las características de la prenda !!", vbInformation, "Agregar Prenda"
        Completos = False
        txtCaracteristicas.SetFocus
        Exit Function
    End If
    
    If Trim(txtMinimo.text) = "" Or Val(txtMinimo.text) = 0 Then
        MsgBox "Introduzca el importe mínimo !!", vbInformation, "Agregar Prenda"
        Completos = False
        txtMinimo.SetFocus
        Exit Function
    End If
    
    If Trim(txtMaximo.text) = "" Or Val(txtMaximo.text) = 0 Then
        MsgBox "Introduzca el importe máximo !!", vbInformation, "Agregar Prenda"
        Completos = False
        txtMaximo.SetFocus
        Exit Function
    End If

End Function
