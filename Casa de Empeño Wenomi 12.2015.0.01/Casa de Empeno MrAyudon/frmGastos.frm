VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmGastos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registrar gastos"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6510
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGastos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2310
   ScaleWidth      =   6510
   Begin VB.TextBox txtConcepto 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   120
      MaxLength       =   100
      TabIndex        =   1
      Top             =   1320
      Width           =   4215
   End
   Begin VB.ComboBox cmbCuenta 
      BackColor       =   &H00FFFFFF&
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
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.TextBox txtFecha 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtImporte 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   4560
      MaxLength       =   12
      TabIndex        =   2
      Top             =   1320
      Width           =   1575
   End
   Begin DevPowerFlatBttn.FlatBttn cmdMosFecha 
      Height          =   300
      Left            =   2040
      TabIndex        =   9
      Top             =   2520
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
      Picture         =   "frmGastos.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   5295
      TabIndex        =   4
      Top             =   1845
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
      Picture         =   "frmGastos.frx":0121
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   4095
      TabIndex        =   3
      Top             =   1845
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
      Picture         =   "frmGastos.frx":0673
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   2775
      TabIndex        =   11
      Top             =   1845
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
      Picture         =   "frmGastos.frx":0BC5
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
      Left            =   5520
      TabIndex        =   13
      Top             =   240
      Width           =   975
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
      Left            =   4680
      TabIndex        =   12
      Top             =   240
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Concepto:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   1110
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cuenta:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Importe:"
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
      Left            =   5055
      TabIndex        =   6
      Top             =   960
      Width           =   1080
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
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   795
   End
End
Attribute VB_Name = "frmGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Private Sub cmbCuenta_GotFocus()
    Cambiar_Color True, cmbCuenta
End Sub

Private Sub cmbCuenta_LostFocus()
    Cambiar_Color False, cmbCuenta
End Sub

Private Sub cmdImprimir_Click()
Dim Folio As Long

    Folio = frmReimpresionrecibos.ReImprimir("gastos", "Folio", " WHERE Folio=")
    If Folio > 0 Then
        
        Imprimir Folio
    
    ElseIf Folio = 0 Then
        
        MsgBox "No se encontró el folio especificado !!", vbInformation, "Registrar gastos"
    End If
End Sub

Private Sub cmdMosFecha_Click()
    txtFecha.text = frmCalendario.Fecha(txtFecha.text)
End Sub

Private Sub txtConcepto_GotFocus()
    Seleccionar_Texto txtConcepto
    Cambiar_Color True, txtConcepto
End Sub

Private Sub txtConcepto_KeyPress(KeyAscii As Integer)
    KeyAscii = Mayusculas(KeyAscii)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtConcepto_LostFocus()
    Cambiar_Color False, txtConcepto
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

Private Sub cmdAceptar_Click()
    
    If MsgBox("Están correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Registrar gastos") = vbYes Then
        
        If Validar Then Grabar_Datos
    
    End If

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
    lblFolio.Caption = Regresa_Movimiento(False, "FolioGastos")
    Cargar_Combos "Descripcion", "cuentasgastos", cmbCuenta
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub txtImporte_GotFocus()
    Seleccionar_Texto txtImporte
    Cambiar_Color True, txtImporte
End Sub

Private Sub txtImporte_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtImporte_LostFocus()
    txtImporte.text = Format(txtImporte.text, FMoneda)
    Cambiar_Color False, txtImporte
End Sub

'Grabamos los datos
Private Sub Grabar_Datos()
Dim Movimiento As Long, Folio As Long, Importe As Double, IDCuentaGasto As Integer, Hora As String

On Error GoTo Error
    
    'Saco el Movimiento
    Movimiento = Regresa_Movimiento(False)
    Regresa_Movimiento True
    
    'Saco el Folio
    Folio = Regresa_Movimiento(False, "FolioGastos")
    Regresa_Movimiento True, "FolioGastos"
    
    'Tomo el importe
    Importe = CDbl(txtImporte.text)
    
    'Tomo la Hora
    Hora = Time
    
    'Grabo en la cuenta de gastos
    dbDatos.Execute "INSERT INTO gastos (Folio,Fecha,Concepto,Importe,CuentaGastos,IDUsuario,IDSucursal,PC) VALUES (" & _
                    Folio & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "','" & Trim(txtConcepto.text) & "'," & ConvMoneda(Importe) & "," & cmbCuenta.ItemData(cmbCuenta.ListIndex) & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ",'" & NombrePc & "')"
                    
    'Grabamos el Cargo
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                    Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'GA01','511101'," & ConvMoneda(Importe) & "," & TIPO_CARGO & ",0,'Gastos','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
    'Grabamos el Abono
    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                    Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'GA50','110150'," & ConvMoneda(Importe) & "," & TIPO_ABONO & ",0,'Gastos','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

'''    'Grabamos el Abono
'''    dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
'''                    Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'GA50','199450'," & ConvMoneda(Importe) & "," & TIPO_ABONO & ",0,'Gastos','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
          
    'Imprimo el recibo
    Imprimir Folio
    
    Limpiar
    lblFolio.Caption = Regresa_Movimiento(False, "FolioGastos")
    Exit Sub
    
Error:
    Maneja_Error Err
    
End Sub

'Validamos que esten correctos los datos
Private Function Validar() As Boolean
    
    Validar = True
  
    If cmbCuenta.ListIndex < 0 Then
        MsgBox "Imposible grabar el gasto, Datos incompletos !!", vbOKOnly + vbCritical, "Registrar gastos"
        cmbCuenta.SetFocus
        Validar = False
        Exit Function
    End If
  
    If Trim(txtConcepto.text) = "" Then
        MsgBox "Imposible grabar el gasto, Datos incompletos !!", vbOKOnly + vbCritical, "Registrar gastos"
        txtConcepto.SetFocus
        Validar = False
        Exit Function
    End If
  
    If Trim(txtImporte.text) = "" Then
        MsgBox "Imposible grabar el gasto, Datos incompletos !!", vbOKOnly + vbCritical, "Registrar gastos"
        txtImporte.SetFocus
        Validar = False
        Exit Function
    End If
  
End Function

'Limpiamos los campos
Private Sub Limpiar()
    txtImporte.text = ""
    txtConcepto.text = ""
    cmbCuenta.ListIndex = -1
    txtFecha.text = Format(Date, "DD/MM/YY")
    lblFolio.Caption = ""
End Sub

Sub Imprimir(Folio As Long)
Dim ImprDefault As Boolean
    
    'Checo si hay impresora por default
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))

    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\NotaGastos.rpt"
        .SelectionFormula = "{gastos.Folio}=" & Folio & ""
        .Formulas(0) = "Caja='" & NombrePc & "'"
        .Destination = crptToWindow
        
        'La mando a la impresora por default
        If ImprDefault Then
            .PrinterName = strNombreImp
            .PrinterDriver = strDriverImp
            .PrinterPort = strPuertoImp
            .Destination = crptToPrinter
        End If
                
        .WindowTitle = "Recibo Gastos"
        .WindowState = crptMaximized
        .Action = 1
    End With

End Sub
