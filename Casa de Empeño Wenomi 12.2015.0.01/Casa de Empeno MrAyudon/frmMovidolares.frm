VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmMovidolares 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dotación/Retiro Divisas Bóveda"
   ClientHeight    =   2520
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
   Icon            =   "frmMovidolares.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2520
   ScaleWidth      =   6495
   Begin VB.TextBox txtDivisa 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   1560
      Width           =   3855
   End
   Begin VB.TextBox txtTipoCambio 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Top             =   3000
      Width           =   1455
   End
   Begin VB.TextBox txtCantidad 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   4860
      TabIndex        =   1
      Top             =   1560
      Width           =   1575
   End
   Begin VB.OptionButton optRetiro 
      Appearance      =   0  'Flat
      Caption         =   "Retiro de divisas a Bóveda"
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   4095
   End
   Begin VB.OptionButton optDotacion 
      Appearance      =   0  'Flat
      Caption         =   "Dotación de divisas a Bóveda"
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
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Value           =   -1  'True
      Width           =   4215
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosDivisa 
      Height          =   300
      Left            =   3990
      TabIndex        =   5
      Top             =   1560
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   529
      AlignCaption    =   4
      AutoSize        =   0   'False
      Caption         =   ". . ."
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
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   5340
      TabIndex        =   9
      Top             =   2070
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
      Picture         =   "frmMovidolares.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   4200
      TabIndex        =   10
      Top             =   2070
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
      Picture         =   "frmMovidolares.frx":055E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      Top             =   2070
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
      Picture         =   "frmMovidolares.frx":0AB0
   End
   Begin VB.Label Label2 
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
      Left            =   4800
      TabIndex        =   13
      Top             =   240
      Width           =   675
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
      Left            =   5550
      TabIndex        =   12
      Top             =   255
      Width           =   945
   End
   Begin VB.Label Label4 
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
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "T. Cambio:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   180
      TabIndex        =   7
      Top             =   2640
      Width           =   1695
   End
   Begin VB.Label Label3 
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
      Height          =   255
      Left            =   5220
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "frmMovidolares"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmdAceptar_Click()
Dim Folio As Long, Movimiento As Long, Tipo As Integer, Cantidad As Long, Cambio As Double
Dim TipoDivisa As Integer, crTotal As Double, Hora As String
        
    If Trim(txtDivisa.text) = "" Then
        
        MsgBox "Seleccione el tipo de divisa !!", vbInformation, "Dotación/Retiro Divisas Bóveda"
        
    ElseIf Trim(txtCantidad.text) = "" Then
        
        MsgBox "Introduzca la cantidad !!", vbInformation, "Dotación/Retiro Divisas Bóveda"
    Else
        
        'Saco el Folio
        Folio = Regresa_Movimiento(False, "FolioBovedaDivisas")
        Regresa_Movimiento True, "FolioBovedaDivisas"
        
        'Saco el Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
        
        Tipo = IIf(optDotacion.Value, 0, 1)
        TipoDivisa = Val(txtDivisa.Tag)
        Cantidad = txtCantidad.text
        Cambio = 0
        crTotal = Cantidad * Cambio
        Hora = Time
        
        'Grabo en la tabla de Divisas
        dbDatos.Execute "INSERT INTO divisas (Folio,Fecha,IDDivisa,Importe,Cantidad,Tipo,TipoEntrada,IDUsuario,IDSucursal,Notas,PC) VALUES (" & _
                        Folio & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & TipoDivisa & "," & ConvMoneda(Cambio) & "," & Cantidad & "," & Tipo & ",1," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'','" & NombrePc & "')"

        'Grabamos el cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,PC,IDUsuario,IDSucursal,IDDivisa) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','" & IIf(optDotacion.Value, "Dotacion Divisas", "Retiro Divisas") & "'," & Movimiento & "," & Folio & ",'" & IIf(optDotacion.Value, "BADV01", "BADV50") & "','" & IIf(optDotacion.Value, "910901", "910950") & "'," & ConvMoneda(Cantidad) & "," & TIPO_CARGO & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & TipoDivisa & ")"

        'Grabamos el abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,PC,IDUsuario,IDSucursal,IDDivisa) VALUES ('" & _
                        Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','" & IIf(optDotacion.Value, "Dotacion Divisas", "Retiro Divisas") & "'," & Movimiento & "," & Folio & ",'" & IIf(optDotacion.Value, "BADV50", "BADV01") & "','" & IIf(optDotacion.Value, "200950", "200901") & "'," & ConvMoneda(Cantidad) & "," & TIPO_ABONO & ",'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & "," & TipoDivisa & ")"
        
        'Imprimo el Ticket
        ImprimeTicket Folio
        
        txtDivisa.text = ""
        txtDivisa.Tag = ""
        txtTipoCambio.text = ""
        txtCantidad.text = ""
        lblFolio.Caption = Regresa_Movimiento(False, "FolioBovedaDivisas")
    End If

End Sub

Private Sub cmdImprimir_Click()
Dim Folio As Long

    Folio = frmReimpresionrecibos.ReImprimir("divisas", "Folio", " WHERE TipoEntrada=1 AND Folio=")
    If Folio > 0 Then
        
        ImprimeTicket Folio
    
    ElseIf Folio = 0 Then
        
        MsgBox "No se encontró el folio especificado !!", vbInformation, "Dotación/Rétiro Divisas Bóveda"
    End If

End Sub

Private Sub cmdMosDivisa_Click()
    frmMuestraDivisas.Posicion Me, Me.txtDivisa
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Sub Inicializar()
    lblFolio.Caption = Regresa_Movimiento(False, "FolioBovedaDivisas")
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
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
    txtCantidad.text = Format(txtCantidad.text, "###,###,###,###0")
    Cambiar_Color False, txtCantidad
End Sub

Private Sub txtTipoCambio_GotFocus()
    Seleccionar_Texto txtTipoCambio
    Cambiar_Color True, txtTipoCambio
End Sub

Private Sub txtTipoCambio_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtTipoCambio_LostFocus()
    Cambiar_Color False, txtTipoCambio
End Sub

Private Sub txtDivisa_GotFocus()
    Seleccionar_Texto txtDivisa
    Cambiar_Color True, txtDivisa
End Sub

Private Sub txtDivisa_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtDivisa_LostFocus()
    Cambiar_Color False, txtDivisa
End Sub

Public Sub TipoCambio(ID As Integer)
Dim rcConsulta As New ADODB.Recordset
Dim crImporte As Double, IDDivisa As Long

On Error GoTo error
    
    txtTipoCambio.text = ""
    txtDivisa.text = ""
    txtDivisa.Tag = ""
    
    rcConsulta.Open "SELECT Descripcion FROM monedas WHERE Clave=" & ID, dbDatos, adOpenForwardOnly, adLockOptimistic
        
        txtDivisa.text = rcConsulta!Descripcion
        txtDivisa.Tag = ID
    rcConsulta.Close
    Set rcConsulta = Nothing
    Exit Sub
    
error:
    Maneja_Error Err
    Set rcConsulta = Nothing
End Sub

Sub ImprimeTicket(Folio As Long)
Dim ImprDefault As Boolean

On Error GoTo error
    
    'Checo si hay impresora por default
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))

    With frmMDI.Cr
        .Reset
        .WindowShowPrintSetupBtn = True
        .WindowShowExportBtn = True
        .DiscardSavedData = True
        .ReportFileName = Path & "\Reportes\ReciboDotacionDivisas.rpt"
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .SelectionFormula = "{divisas.Folio}=" & Folio & " AND {divisas.TipoEntrada}=1"
        .Formulas(0) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(1) = "Notas='" & Regresa_Valor_BD("Notas") & "'"
        .Formulas(2) = "Gerente='" & Trim(Regresa_Valor_BD("Gerente")) & "'"
        .Destination = crptToWindow
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
        
        .WindowTitle = "Recibo"
        .WindowState = crptMaximized
        .Action = 1
    End With
    Exit Sub
    
error:
    Maneja_Error Err
End Sub
