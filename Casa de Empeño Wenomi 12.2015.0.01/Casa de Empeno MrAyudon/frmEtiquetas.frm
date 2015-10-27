VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Object = "{0BFA85A1-F9B8-11CF-8939-444553540000}#1.0#0"; "barcode.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEtiquetas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión Etiquetas"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6555
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEtiquetas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4815
   ScaleWidth      =   6555
   Begin VB.ComboBox cmbPrenda 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEtiquetas.frx":000C
      Left            =   120
      List            =   "frmEtiquetas.frx":0022
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   4320
      Width           =   2565
   End
   Begin VB.ComboBox cmbTipoEntrada 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmEtiquetas.frx":003E
      Left            =   120
      List            =   "frmEtiquetas.frx":004B
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2565
   End
   Begin VB.TextBox txtPrecio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   960
      Left            =   3150
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   3195
      Width           =   3135
   End
   Begin VB.ComboBox cmbKilates 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmEtiquetas.frx":006B
      Left            =   1440
      List            =   "frmEtiquetas.frx":0081
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2715
      Width           =   1215
   End
   Begin VB.TextBox txtPeso 
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
      Left            =   120
      MaxLength       =   6
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.ComboBox cmbTipo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEtiquetas.frx":009F
      Left            =   120
      List            =   "frmEtiquetas.frx":00B5
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   3480
      Width           =   2565
   End
   Begin VB.TextBox txtPartida 
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
      Left            =   1440
      MaxLength       =   2
      TabIndex        =   3
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtBoleta 
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
      Left            =   120
      MaxLength       =   6
      TabIndex        =   2
      Top             =   2040
      Width           =   1095
   End
   Begin VB.TextBox txtSucursal 
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
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox txtCantidad 
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
      Left            =   5520
      MaxLength       =   3
      TabIndex        =   10
      Text            =   "1"
      Top             =   2520
      Width           =   495
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   285
      Left            =   6045
      TabIndex        =   11
      Top             =   2520
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   503
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtCantidad"
      BuddyDispid     =   196618
      OrigLeft        =   3240
      OrigTop         =   2640
      OrigRight       =   3480
      OrigBottom      =   2940
      Max             =   999
      Min             =   1
      SyncBuddy       =   -1  'True
      Wrap            =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.Frame Frame1 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1860
      Left            =   3120
      TabIndex        =   12
      Top             =   600
      Width           =   3405
      Begin BARCODELib.Barcode bcCodigo 
         Height          =   1410
         Left            =   75
         TabIndex        =   13
         Top             =   360
         Width           =   3300
         _Version        =   65536
         _ExtentX        =   5821
         _ExtentY        =   2487
         _StockProps     =   25
         Text            =   "12345678901212"
         TypeName        =   "EAN 13"
         Text            =   "12345678901212"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Borderwidth     =   0
         Borderheight    =   5
         NotchHeightInPercent=   15
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMuestraSucursal 
      Height          =   285
      Left            =   2700
      TabIndex        =   15
      Top             =   495
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   503
      AutoSize        =   0   'False
      Caption         =   "..."
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   4050
      TabIndex        =   23
      Top             =   4365
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   " &Imprimir"
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
      Object.ToolTipText     =   ""
      Picture         =   "frmEtiquetas.frx":00D1
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   5250
      TabIndex        =   24
      Top             =   4365
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
      Picture         =   "frmEtiquetas.frx":0623
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "Prenda:"
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
      Left            =   120
      TabIndex        =   25
      Top             =   3960
      Width           =   825
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Entrada:"
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
      Left            =   120
      TabIndex        =   22
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Precio"
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
      Left            =   5520
      TabIndex        =   21
      Top             =   2880
      Width           =   645
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Kilates:"
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
      Left            =   1440
      TabIndex        =   20
      Top             =   2400
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Peso:"
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
      Left            =   120
      TabIndex        =   19
      Top             =   2400
      Width           =   585
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de prenda:"
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
      Left            =   120
      TabIndex        =   18
      Top             =   3120
      Width           =   1710
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Partida:"
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
      Left            =   1440
      TabIndex        =   17
      Top             =   1680
      Width           =   825
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Boleta:"
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
      Left            =   120
      TabIndex        =   16
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Sucursal:"
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
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad de Etiquetas:"
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
      Left            =   3120
      TabIndex        =   9
      Top             =   2520
      Width           =   2385
   End
End
Attribute VB_Name = "frmEtiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. Gómez Vázquez
' Mazatlan, Sin. 11/06/02
' Modulo frmEtiquetas - frmEtiquetas.frm
' Ultima Modificacion - 11/06/02
' Modificacion Mysql - 29/12/06 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbKilates_GotFocus()
    Seleccionar_Texto cmbKilates
    Cambiar_Color True, cmbKilates
End Sub

Private Sub cmbKilates_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbKilates_LostFocus()
    Cambiar_Color False, cmbKilates
End Sub

Private Sub cmbPrenda_GotFocus()
    Cambiar_Color True, cmbPrenda
End Sub

Private Sub cmbPrenda_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbPrenda_LostFocus()
    Cambiar_Color False, cmbPrenda
End Sub

Private Sub cmbTipo_Click()
    Poner_Codigo
    
    cmbPrenda.Clear
    If cmbTipo.ListIndex > -1 Then
        
        Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda, " WHERE IDTipo=" & cmbTipo.ItemData(cmbTipo.ListIndex), "Descripcion"
    End If
    
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

Private Sub cmbTipoEntrada_Click()
    Poner_Codigo
End Sub

Private Sub cmbTipoEntrada_GotFocus()
    Seleccionar_Texto cmbTipoEntrada
    Cambiar_Color True, cmbTipoEntrada
End Sub

Private Sub cmbTipoEntrada_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbTipoEntrada_LostFocus()
    Cambiar_Color False, cmbTipoEntrada
End Sub

Private Sub cmdImprimir_Click()
Dim Indice As Integer, Cantidad As Integer, crPrecio As Double
Dim Impresora As Printer

On Error GoTo error

    If Requeridos Then
        
        crPrecio = 0
        If Val(txtPrecio.text) > 0 Or Trim(txtPrecio.text) <> "" Then
            crPrecio = CDbl(txtPrecio.text)
        End If
        Cantidad = Val(txtCantidad.text)
        
        Set Impresora = Printer
        With Impresora
                    
            .ScaleMode = vbMillimeters
            .FontBold = False
        
            For Indice = 1 To Cantidad
                
                'Imprimo el Codigo
                Impresora.PaintPicture bcCodigo.Picture, Regresa_Valor("ETIQUETAS", "CodigoX", 0), Regresa_Valor("ETIQUETAS", "CodigoY", 0), Regresa_Valor("ETIQUETAS", "Anchocodigo", 0), Regresa_Valor("ETIQUETAS", "Altocodigo", 0)
        
                'Imprimo el peso
                .Font = "Arial"
                .FontSize = 6.5
                .CurrentX = Regresa_Valor("ETIQUETAS", "PesoX", 0)
                .CurrentY = Regresa_Valor("ETIQUETAS", "PesoY", 0)
                Impresora.Print IIf(Val(txtPeso.text) > 0, Format(txtPeso.text, "##,###0.00") & " g", "")
        
                'Imprimo el Kilataje
                .CurrentX = Regresa_Valor("ETIQUETAS", "KilatesX", 0)
                .CurrentY = Regresa_Valor("ETIQUETAS", "KilatesY", 0)
                Impresora.Print cmbKilates.text
                        
                'Imprimo la prenda
                .CurrentX = Regresa_Valor("ETIQUETAS", "PrendaX", 0)
                .CurrentY = Regresa_Valor("ETIQUETAS", "PrendaY", 0)
                Impresora.Print cmbPrenda.text
            
                'Imprimo el precio
                .CurrentX = Regresa_Valor("ETIQUETAS", "PrecioX", 0)
                .CurrentY = Regresa_Valor("ETIQUETAS", "PrecioY", 0)
                Impresora.Print Format(crPrecio, FMoneda)
        
            .EndDoc
            Next Indice
        
        End With
        
    End If
    Exit Sub
    
error:
    Maneja_Error Err
End Sub

Private Sub cmdMuestraSucursal_Click()
    frmMostrarSucursales.ver Me, txtSucursal, True, True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    Cargar_Combos "Descripcion", "tipo", cmbTipo
    Cargar_Combos "Descripcion", "tipoprenda", cmbPrenda
    bcCodigo.text = "123456789012"
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtBoleta_Change()
    Poner_Codigo
End Sub

Private Sub txtBoleta_GotFocus()
    Seleccionar_Texto txtBoleta
    Cambiar_Color True, txtBoleta
End Sub

Private Sub txtBoleta_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtBoleta_LostFocus()
    Cambiar_Color False, txtBoleta
End Sub

Private Sub txtCantidad_GotFocus()
    Seleccionar_Texto txtCantidad
    Cambiar_Color True, txtCantidad
End Sub

Private Sub txtCantidad_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCantidad_LostFocus()
    Cambiar_Color False, txtCantidad
End Sub

Private Sub Poner_Codigo()
Dim Cadena As String, TipoPrenda As Integer
Dim TipoEntrada As Integer

    'Sucursal
    Cadena = Format(Val(txtSucursal.Tag), "000")
    
    'Checo el Tipo de Entrada
    If cmbTipoEntrada.ListIndex > -1 Then
        
        Select Case cmbTipoEntrada.text
        Case "ALMONEDA"
            TipoEntrada = ENTRADAEMPENO
        
        Case "COMPRA"
            TipoEntrada = ENTRADACOMPRA
        
        Case "DOTACIÓN"
            TipoEntrada = ENTRADADOTACION

        End Select
    
    Else
        TipoEntrada = 0
        
    End If
    
    'Tipo de Entrada
    Cadena = Cadena & TipoEntrada
    
    'Boleta
    Cadena = Cadena & Format(Val(txtBoleta.text), "000000")

    'Partida
    Cadena = Cadena & Format(Val(txtPartida.text), "00")

    bcCodigo.text = Cadena
   
End Sub

Private Sub txtPartida_Change()
    Poner_Codigo
End Sub

Private Sub txtPartida_GotFocus()
    Seleccionar_Texto txtPartida
    Cambiar_Color True, txtPartida
End Sub

Private Sub txtPartida_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPartida_LostFocus()
    Cambiar_Color False, txtPartida
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

Private Sub txtPrecio_GotFocus()
    Seleccionar_Texto txtPrecio
    Cambiar_Color True, txtPrecio
End Sub

Private Sub txtPrecio_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtPrecio_LostFocus()
    txtPrecio.text = Format(txtPrecio.text, FMoneda)
    Cambiar_Color False, txtPrecio
End Sub

Private Sub txtSucursal_GotFocus()
    Seleccionar_Texto txtSucursal
    Cambiar_Color True, txtSucursal
End Sub

Private Sub txtSucursal_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtSucursal_LostFocus()
    Cambiar_Color False, txtSucursal
End Sub

Private Sub UpDown1_Change()
    txtCantidad.SetFocus
    txtCantidad_GotFocus
End Sub

Public Function BuscarSucursal(IDSucursal As Long)
Dim rcConsulta As New ADODB.Recordset

On Error GoTo error

    rcConsulta.Open "SELECT Clave,NombreComercial FROM sucursales WHERE ID=" & IDSucursal, dbDatos, adOpenForwardOnly, adLockOptimistic
    If Not rcConsulta.BOF And Not rcConsulta.EOF Then

        With rcConsulta
            txtSucursal.Tag = !Clave
            txtSucursal.text = !NombreComercial
            Poner_Codigo
        End With

    End If
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Function
    
error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Function

Function Requeridos() As Boolean

    Requeridos = True

    If txtSucursal.Tag = "" Then
        MsgBox "Seleccione la sucursal !!", vbInformation, "Impresión Etiquetas"
        txtSucursal.SetFocus
        Requeridos = False
        Exit Function
    End If

    If txtBoleta.text = "" Then
        MsgBox "Introduzca el número de boleta !!", vbInformation, "Impresión Etiquetas"
        txtBoleta.SetFocus
        Requeridos = False
        Exit Function
    End If

    If txtPartida.text = "" Then
        MsgBox "Introduzca la partida !!", vbInformation, "Impresión Etiquetas"
        txtPartida.SetFocus
        Requeridos = False
        Exit Function
    End If

    If cmbTipo.ListIndex < 0 Then
        MsgBox "Seleccione el tipo de prenda !!", vbInformation, "Impresión Etiquetas"
        cmbTipo.SetFocus
        Requeridos = False
        Exit Function
    End If

    If Val(txtPrecio.text) <= 0 Then
        MsgBox "Introduzca el precio !!", vbInformation, "Impresión Etiquetas"
        txtPrecio.SetFocus
        Requeridos = False
        Exit Function
    End If

End Function
