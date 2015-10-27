VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmBovedaDivisas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja General Divisas"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBovedaDlls.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   7365
   Begin VB.TextBox txtFecha 
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
      Left            =   1200
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
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
      MaxLength       =   14
      TabIndex        =   3
      Top             =   1500
      Width           =   1575
   End
   Begin VB.OptionButton opDotacion 
      Appearance      =   0  'Flat
      Caption         =   "&Dotaci�n de divisas a cajero"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Value           =   -1  'True
      Width           =   4095
   End
   Begin VB.OptionButton opRetiro 
      Appearance      =   0  'Flat
      Caption         =   "&Retiro de divisas a cajero"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   555
      Width           =   3735
   End
   Begin VB.ComboBox cmbDivisas 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   5295
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   529
      AlignCaption    =   4
      AlignPicture    =   4
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   0
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmBovedaDlls.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6195
      TabIndex        =   6
      Top             =   2010
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
      Picture         =   "frmBovedaDlls.frx":0121
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   4995
      TabIndex        =   7
      Top             =   2010
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
      Picture         =   "frmBovedaDlls.frx":0673
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   3675
      TabIndex        =   8
      Top             =   2010
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Re-Imprimir"
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
      Picture         =   "frmBovedaDlls.frx":0BC5
   End
   Begin VB.Label lblFolio 
      AutoSize        =   -1  'True
      Caption         =   "<Folio>"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   285
      Left            =   6270
      TabIndex        =   13
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
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
      Left            =   360
      TabIndex        =   12
      Top             =   2640
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Folio:"
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
      Left            =   5550
      TabIndex        =   11
      Top             =   120
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Cantidad:"
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
      Left            =   5910
      TabIndex        =   10
      Top             =   1080
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Divisa:"
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
      TabIndex        =   9
      Top             =   1080
      Width           =   825
   End
End
Attribute VB_Name = "frmBovedaDivisas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'////////////////////////////////////////////////////////////////
'
' Sistema Montepio
' L.S.C. Juan A. G�mez V�zquez
' Mazatlan, Sin. 26/07/02
' Modulo frmBoveda - frmBoveda.frm
' Ultima Modificacion - 26/07/02
''Modificacion para Mysql 29/12/05 - L.S.C. Juan Alberto Gomez Vazquez
'////////////////////////////////////////////////////////////////

Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbDivisas_GotFocus()
    cmbDivisas.BackColor = &HC0FFFF
End Sub

Private Sub cmbDivisas_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub cmbDivisas_LostFocus()
    cmbDivisas.BackColor = vbWhite
End Sub

Private Sub cmdImprimir_Click()
Dim Folio As Long

    Folio = frmReimpresionrecibos.ReImprimir("divisas", "Folio", " WHERE TipoEntrada=2 AND Folio=")
    If Folio > 0 Then
        
        Imprimir Folio
    
    ElseIf Folio = 0 Then
        
        MsgBox "No se encontr� el folio especificado !!", vbInformation, "Caja General"
    End If

End Sub

Private Sub cmdMosFecha_Click()
    txtFecha.text = frmCalendario.Fecha(txtFecha.text)
End Sub

Private Sub cmdAceptar_Click()
    If Validar Then Grabar_Datos
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    txtFecha.text = Format(Date, "DD/MM/YY")
    lblFolio.Caption = Regresa_Movimiento(False, "FolioBovedaDivisas")
    Cargar_Combos "Descripcion", "monedas", cmbDivisas, , , , "Clave"
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtFecha_GotFocus()
    Seleccionar_Texto txtFecha
    Cambiar_Color True, txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtFecha_LostFocus()
    Cambiar_Color False, txtFecha
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
    txtCantidad.text = Format(txtCantidad.text, FMoneda)
    Cambiar_Color False, txtCantidad
End Sub

'Grabamos los datos
Private Sub Grabar_Datos()
Dim Movimiento As Long, Folio As Long, Importe As Double, Hora As String

    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton2, "Caja General Divisas") = vbYes Then
        
        'Saco el Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
        
        'Saco el Folio
        Folio = Regresa_Movimiento(False, "FolioBovedaDivisas")
        Regresa_Movimiento True, "FolioBovedaDivisas"
        
        'Tomo la Hora
        Hora = Time
        
        Importe = CDbl(txtCantidad.text)
    
        'Grabo en la tabla de Divisas
        dbDatos.Execute "INSERT INTO divisas (Folio,Fecha,IDDivisa,Importe,Cantidad,Tipo,TipoEntrada,IDUsuario,IDSucursal,Notas,PC) VALUES (" & _
                        Folio & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & cmbDivisas.ItemData(cmbDivisas.ListIndex) & ",0," & ConvMoneda(Importe) & "," & IIf(opDotacion.Value, 0, 1) & ",2," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'','" & NombrePc & "')"
                      
        'Grabamos el cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,IDDivisa) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(opDotacion.Value, "DODV01", "REDV01") & "','" & IIf(opDotacion.Value, "999401", "910901") & "'," & ConvMoneda(Importe) & "," & TIPO_CARGO & ",2,'" & IIf(opDotacion.Value, "Dotacion Divisas a Caja", "Retiro Divisas a Caja") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & cmbDivisas.ItemData(cmbDivisas.ListIndex) & ")"
                    
        'Grabamos el abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal,IDDivisa) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(opDotacion.Value, "DODV50", "REDV50") & "','" & IIf(opDotacion.Value, "910950", "999450") & "'," & ConvMoneda(Importe) & "," & TIPO_ABONO & ",2,'" & IIf(opDotacion.Value, "Dotacion Divisas a Caja", "Retiro Divisas a Caja") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & cmbDivisas.ItemData(cmbDivisas.ListIndex) & ")"
        
        'Limpio la ventana
        Limpiar
        
        'Tomo el nuevo folio
        lblFolio.Caption = Regresa_Movimiento(False, "FolioBovedaDivisas")
        
        'Saco el Recibo
        Imprimir Folio

    End If

End Sub

'Validamos que esten correctos los datos
Private Function Validar() As Boolean

    Validar = True
  
    If cmbDivisas.ListIndex < 0 Then
        MsgBox "Imposible grabar el movimiento, Datos incompletos", vbOKOnly + vbCritical
        cmbDivisas.SetFocus
        Validar = False
        Exit Function
    End If
  
    If Trim(txtCantidad.text) = "" Then
        MsgBox "Imposible grabar el movimiento, Datos incompletos", vbOKOnly + vbCritical
        txtCantidad.SetFocus
        Validar = False
        Exit Function
    End If
  
    If Not IsDate(txtFecha.text) Then
        MsgBox "Imposible de grabar el movimiento, Favor de poner una fecha v�lida", vbOKOnly + vbCritical
        Validar = False
        txtFecha.SetFocus
    End If
  
End Function

'Limpiamos los campos
Private Sub Limpiar()
    opDotacion = True
    txtCantidad.text = ""
    txtFecha.text = Format(Date, "DD/MM/YY")
    lblFolio.Caption = ""
    cmbDivisas.ListIndex = -1
End Sub

Function Imprimir(Folio As Long)
Dim Usuario As String, ImprDefault As Boolean, crImporte As Double, Operacion As Boolean
    
    'Checo si hay impresora por default
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))

    Usuario = SacaValor("usuarios", "Nombre", " WHERE ID='" & Trim(frmMDI.IDUsuario) & "'")
    Operacion = IIf(Val(SacaValor("divisas", "Tipo", " WHERE Folio=" & Folio)) = 1, True, False)
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\NotaCajaDivisas.rpt"
        .SelectionFormula = "{divisas.Folio}=" & Folio & " AND {divisas.TipoEntrada}=2"
        .Formulas(1) = "Recibido='" & IIf(Operacion, Trim(cmbDivisas.text), "BOVEDA") & " " & IIf(Operacion, Usuario, Regresa_Valor_BD("Gerente")) & "'"
        .Formulas(2) = "Enviado='" & IIf(Operacion, Trim(cmbDivisas.text), "BOVEDA") & " " & IIf(Operacion = False, Usuario, Regresa_Valor_BD("Gerente")) & "'"
        .Formulas(3) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(4) = "UsuarioRecibe='" & IIf(Operacion, Usuario, Regresa_Valor_BD("Gerente")) & "'"
        .Formulas(5) = "UsuarioEnvia='" & IIf(Operacion = False, Usuario, Regresa_Valor_BD("Gerente")) & "'"
        .Destination = crptToWindow
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                
        .WindowState = crptMaximized
        .WindowTitle = "Recibo"
        .Action = 1
    End With

End Function
