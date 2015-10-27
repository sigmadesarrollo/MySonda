VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmReimpresion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Re-Imprimir"
   ClientHeight    =   1350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3750
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReimpresion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   3750
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opContrato 
      Appearance      =   0  'Flat
      Caption         =   "Contrato"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   1335
   End
   Begin VB.OptionButton opRecibo 
      Appearance      =   0  'Flat
      Caption         =   "Recibo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1560
      TabIndex        =   3
      Top             =   120
      Width           =   1185
   End
   Begin VB.TextBox txtNumContrato 
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
      TabIndex        =   1
      Top             =   600
      Width           =   1095
   End
   Begin VB.CheckBox chkAutomovil 
      Appearance      =   0  'Flat
      Caption         =   "Autom�vil"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
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
      Top             =   1020
      Width           =   1455
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Default         =   -1  'True
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   2
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "        &Imprimir"
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
      Picture         =   "frmReimpresion.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   2505
      TabIndex        =   6
      Top             =   900
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
      Picture         =   "frmReimpresion.frx":055E
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Contrato:"
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
      TabIndex        =   0
      Top             =   600
      Width           =   1050
   End
End
Attribute VB_Name = "frmReimpresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IDEmpe�o As Long, IDNuevoEmpe�o As Long, Destino As Integer, strAbono As String, crAbono    As Double
Dim Fl() As cFlatControl

Public Function Ver() As Long

On Error Resume Next

    Me.Show vbModal
    Ver = IDEmpe�o
End Function

Private Sub cmdAceptar_Click()
    If Trim(txtNumContrato.text) = "" Then
        
        MsgBox "Introduzca el n�mero de contrato !!", vbInformation, "Re-Imprimir"
        txtNumContrato.SetFocus
    Else
        
        Buscar_Empeno txtNumContrato
    End If
End Sub

Private Sub cmdSalir_Click()
    IDEmpe�o = 0
    Unload Me
End Sub

Private Sub Form_Load()
    Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub opContrato_Click()
    Label5.Caption = "Contrato:"
End Sub

Private Sub opRecibo_Click()
    Label5.Caption = "Recibo:"
End Sub

Private Sub txtNumContrato_Change()
    txtNumContrato.Tag = ""
End Sub

Private Sub txtNumcontrato_GotFocus()
    Seleccionar_Texto txtNumContrato
    Cambiar_Color True, txtNumContrato
End Sub

Private Sub txtNumcontrato_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtNumcontrato_LostFocus()
    Cambiar_Color False, txtNumContrato
End Sub

Private Sub Buscar_Empeno(strFolio As String)
    Dim rcEmpe�o As New ADODB.Recordset
    Dim lFolio As Long, Serie As String, crPrestamo As Double, strIDUsuarioMov As String, IDUsuarioMov As Integer
    Dim Fecha As Date, Folio As Long, Movimiento As Long, Hora As String
    Dim Vencimiento As Date
    Dim crCirculando As Integer, stCirculando As String
    Dim stRentaGPS As String, stRentaSeguro As String, stRentaIva As String, stSerie As String
    Dim crRentaGPS As Double, crRentaSeguro As Double, crRentaIva As Double
    Dim elOrigen As Integer
On Error GoTo Error
    
    If chkAutomovil.Value = 0 Then
        
        Serie = "(Serie=" & SERIE_A & " OR Serie=" & SERIE_D & " OR Serie=" & SERIE_C & ")"
    Else
    
        Serie = "Serie=" & SERIE_B
    End If
  
    lFolio = strFolio
    IDEmpe�o = 0
        
    If opContrato.Value Then
    
        rcEmpe�o.Open "SELECT MAX(ID) AS IDEmpeno,Fecha FROM empeno WHERE NumContrato=" & lFolio & " AND Cancelado=0 AND " & Serie & IIf(opContrato.Value, " AND Destino=0", " AND (Destino=2 OR Destino=3)"), dbDatos, adOpenForwardOnly, adLockReadOnly
        If Not rcEmpe�o.BOF And Not rcEmpe�o.EOF And Not IsNull(rcEmpe�o!IDEmpeno) Then
            IDEmpe�o = rcEmpe�o!IDEmpeno
            Fecha = Format(rcEmpe�o!Fecha, "YYYY/MM/DD")
        End If
        rcEmpe�o.Close
        
        If IDEmpe�o > 0 Then
            If MsgBox("Se Cobrara la Reimpresion?", vbQuestion + vbYesNo, "Empe�os") = vbYes Then 'Format(Fecha, "YYYY/MM/DD") <> Format(Date, "YYYY/MM/DD")
    
                'Saco el Folio
                Folio = Regresa_Movimiento(False, "FolioReImpresiones")
                Regresa_Movimiento True, "FolioReImpresiones"
    
                'Saco el Movimiento
                Movimiento = Regresa_Movimiento(False)
                Regresa_Movimiento True
    
                Hora = Time
    
                'Cuenta ReImpresiones
                'dbDatos.Execute "INSERT INTO ReImpresion(Folio, Costo, IDEmpeno, Fecha, Tipo) VALUES(" & _
                '    Folio & "," & CDbl(Regresa_Valor_BD("ImportePerdida")) & "," & IDEmpe�o & ",'" & Format(Date, "YYYY/MM/DD") & "','" & IIf(opContrato.Value = True, "BOLETA", "RECIBO") & "')"
    
                'Cargo Efectivo
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','ReImpresion'," & Movimiento & "," & Folio & ",'RE03','110101'," & _
                    ConvMoneda(Regresa_Valor_BD("ImportePerdida")) & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
                'Cargo ReImpresion
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','ReImpresion'," & Movimiento & "," & Folio & ",'RE03','530150'," & _
                    ConvMoneda(Regresa_Valor_BD("ImportePerdida")) & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
                'Cargo Caja
                dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Concepto,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,PC,IDUsuario,IDSucursal) VALUES " & _
                    "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "','ReImpresion'," & Movimiento & "," & Folio & ",'RE03','199401'," & _
                    ConvMoneda(Regresa_Valor_BD("ImportePerdida")) & "," & TIPO_CARGO & ",0,'" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
            End If
            
            
            
            If chkAutomovil.Value <> 1 Then
                'Imprimir_Boleta_CR IDEmpe�o, True
                Imprimir_Boleta_CR_Caidas IDEmpe�o, True
                'Imprimir_Boleta_CR_Caidas IDEmpe�o, True
            Else
                crPrestamo = Regresa_Valor_Empeno("Prestamo", IDEmpe�o)
                'frmEmpe�o.Imprimir_Boleta_CR_Auto IDEmpe�o
                'Imprimir_Boleta_CR_Caidas_Autos IDEmpe�o, True
                ' Imprimir_Boleta_CR IDEmpe�o, True, , True
                 Imprimir_Boleta_CR_Caidas IDEmpe�o, False, True, False
            End If
    
        Else
            
            MsgBox "No se encontr� el contrato especificado !!", vbInformation, "Re-Imprimir"
        End If
    Else
        
        strAbono = ""
        crAbono = 0
        strIDUsuarioMov = 0
        IDUsuarioMov = 0
        IDNuevoEmpe�o = 0
        crCirculando = 0
        crRentaGPS = 0
        crRentaSeguro = 0
        crRentaIva = 0
        rcEmpe�o.Open "SELECT ID,IDEmpenoDestino FROM empeno WHERE FolioNota=" & lFolio & " AND Cancelado=0", dbDatos, adOpenForwardOnly, adLockOptimistic 'AND (Destino=0 OR Destino=" & D_DESEMPE�O & ")
        If Not rcEmpe�o.BOF And Not rcEmpe�o.EOF And Not IsNull(rcEmpe�o!ID) Then
            IDEmpe�o = rcEmpe�o!ID
            IDNuevoEmpe�o = rcEmpe�o!IDEmpenoDestino
            rcEmpe�o.Close
        
            'Tomo el Destino
            Destino = SacaValor("empeno", "Destino", " WHERE ID=" & IDEmpe�o)
             'Tomo el Destino
           ' elOrigen = SacaValor("empeno", "Origen", " WHERE ID=" & IDEmpe�o)
            
            
           
        
            If IDNuevoEmpe�o > 0 Then
                
                rcEmpe�o.Open "SELECT empeno.Prestamo,empeno.Avaluo,empeno.Fecha,empeno.TipoInteres,empeno.Serie,Vencimiento FROM empeno WHERE ID=" & IDNuevoEmpe�o, dbDatos, adOpenForwardOnly, adLockOptimistic
                    Vencimiento = rcEmpe�o!Vencimiento
                                                    
                    'Opciones de Pago
                    OpcionesPago rcEmpe�o!Prestamo, rcEmpe�o!Avaluo, rcEmpe�o!Fecha, IDNuevoEmpe�o, rcEmpe�o!TipoInteres, IIf(rcEmpe�o!Serie = SERIE_B, True, False)
                    
                rcEmpe�o.Close
                Destino = 2
            End If
            
            
            'Tomo si hay un abono
            strAbono = SacaValor("empeno", "Pago", " WHERE ID=" & IDEmpe�o)
            
            'Tomo si es Circulacion
            stCirculando = SacaValor("empeno", "Circulando", " WHERE ID=" & IDEmpe�o)
            'Tomo la serie
            stSerie = SacaValor("empeno", "Serie", " WHERE ID=" & IDEmpe�o)
            If Trim(stCirculando) <> "" Then
                crCirculando = CInt(stCirculando)
            End If
            If crCirculando = 1 Then
                'Tomo el importe por renta de GPS
                stRentaGPS = SacaValor("empeno", "ImporteRentaGPS", " WHERE ID=" & IDEmpe�o)
                If Trim(stRentaGPS) <> "" Then
                    crRentaGPS = CDbl(stRentaGPS)
                End If
                
                'Tomo el importe por renta de Seguro Auto
                stRentaSeguro = SacaValor("empeno", "ImporteSeguroAuto", " WHERE ID=" & IDEmpe�o)
                If Trim(stRentaSeguro) <> "" Then
                    crRentaSeguro = CDbl(stRentaSeguro)
                End If
                
                'Tomo el importe por renta de Seguro Auto
                stRentaIva = SacaValor("empeno", "ImporteIVAGPSSeguro", " WHERE ID=" & IDEmpe�o)
                If Trim(stRentaIva) <> "" Then
                    crRentaIva = CDbl(stRentaIva)
                End If
                
            End If
            
            
           
            
            If Trim(strAbono) <> "" Then
                crAbono = CDbl(strAbono)
            End If
            
            'Tomo el usuario que hizo el movimiento
            strIDUsuarioMov = SacaValor("empeno", "IDUsuarioMov", " WHERE ID=" & IDEmpe�o)
            
            If Trim(strIDUsuarioMov) <> "" Then
                
                IDUsuarioMov = CInt(strIDUsuarioMov)
            End If
            
            
            
            If stSerie = "2" Then
                  frmEmpe�o.Imprimir_Nota_Auto IDEmpe�o, Destino, IIf(Destino = OD_REFRENDO, crAbono, 0), IDUsuarioMov, DateAdd("D", Regresa_Valor_BD("DiasEnajenacion"), Vencimiento)

            Else
                  frmEmpe�o.Imprimir_Nota IDEmpe�o, Destino, IIf(Destino = OD_REFRENDO, crAbono, 0), IDUsuarioMov, DateAdd("D", Regresa_Valor_BD("DiasEnajenacion"), Vencimiento)

            End If
            
                      
            If crCirculando = 1 Then
                'If crCirculando = 1 Then
                   frmEmpe�o.Imprimir_Nota_GPS_Seguro IDEmpe�o, crRentaGPS, crRentaSeguro, crRentaIva, Destino, IIf(Destino = OD_REFRENDO, crAbono, 0), IDUsuarioMov, DateAdd("D", Regresa_Valor_BD("DiasEnajenacion") + 1, Vencimiento)
                'End If
            End If
        
        Else
            MsgBox "No se encontr� el recibo especificado !!", vbInformation, "Re-Imprimir"
        End If
    End If
    
           
Error:
    Maneja_Error Err
    Set rcEmpe�o = Nothing
    Unload Me

End Sub
