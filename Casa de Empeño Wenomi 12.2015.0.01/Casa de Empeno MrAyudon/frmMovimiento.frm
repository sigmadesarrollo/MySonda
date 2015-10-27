VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmMovimiento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Bóveda/Caja"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "frmMovimiento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2865
   ScaleWidth      =   7635
   Begin VB.CheckBox ChkOtros 
      Appearance      =   0  'Flat
      Caption         =   "Otros"
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
      Height          =   285
      Left            =   360
      TabIndex        =   17
      Top             =   2880
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Frame Frame2 
      Height          =   2715
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   4620
      Begin VB.OptionButton opDotCaja 
         Appearance      =   0  'Flat
         Caption         =   "DOTACIÓN A CAJA"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   2970
      End
      Begin VB.OptionButton opRetCaja 
         Appearance      =   0  'Flat
         Caption         =   "RETIRO DE CAJA"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txtImporte 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1320
         MaxLength       =   12
         TabIndex        =   0
         Text            =   "0.00"
         Top             =   2250
         Width           =   2100
      End
      Begin VB.OptionButton optDepositoo 
         Appearance      =   0  'Flat
         Caption         =   "RETIRO DE BÓVEDA"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   3015
      End
      Begin VB.OptionButton opDeposito 
         Appearance      =   0  'Flat
         Caption         =   "DOTACIÓN A BANCOS DESDE CENTRAL"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   3840
         Width           =   5280
      End
      Begin VB.OptionButton opTransferencia 
         Appearance      =   0  'Flat
         Caption         =   "RETIRO DE BANCOS A CENTRAL"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   4320
         Width           =   5055
      End
      Begin VB.OptionButton opCheque 
         Appearance      =   0  'Flat
         Caption         =   "DOTACIÓN A BÓVEDA"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3360
      End
      Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
         Height          =   375
         Left            =   3465
         TabIndex        =   1
         Top             =   2235
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
         Picture         =   "frmMovimiento.frx":000C
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "IMPORTE:"
         BeginProperty Font 
            Name            =   "Century Gothic"
            Size            =   12.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   2280
         Width           =   1125
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   2310
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
      Picture         =   "frmMovimiento.frx":055E
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   4920
      TabIndex        =   3
      Top             =   2310
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
      Picture         =   "frmMovimiento.frx":0AB0
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "FOLIO:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   5160
      TabIndex        =   18
      Top             =   240
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "FECHA:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1455
      TabIndex        =   16
      Top             =   2835
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Label lblFolio 
      AutoSize        =   -1  'True
      Caption         =   "<Folio>"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   6360
      TabIndex        =   15
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lblFecha 
      AutoSize        =   -1  'True
      Caption         =   "<FECHA>"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   2850
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "SALDO BÓVEDA"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   5160
      TabIndex        =   7
      Top             =   720
      Width           =   2145
   End
   Begin VB.Label lblBoveda 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   6600
      TabIndex        =   6
      Top             =   1080
      Width           =   570
   End
   Begin VB.Label lblBancos 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   345
      Left            =   6600
      TabIndex        =   5
      Top             =   1800
      Visible         =   0   'False
      Width           =   570
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "SALDO CAJA"
      BeginProperty Font 
         Name            =   "Century Gothic"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   345
      Left            =   5520
      TabIndex        =   4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1770
   End
End
Attribute VB_Name = "frmMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl
Dim ImporteOriginal As Double
Dim ImporteOriginal2 As Double

Private Sub Inicializar()
    Screen.MousePointer = vbHourglass
    lblFecha.Caption = Format(Date, "DD/MMM/YY")
    opCheque.Value = True
    Cargar_Montos
    CentrarForm Me, frmMDI
    Poner_Flat Fl, Me.Controls, Me
    Screen.MousePointer = vbDefault
End Sub

Private Sub ChkOtros_Click()
    txtImporte_Change
End Sub

Private Sub cmdAceptar_Click()
    
    If Validar Then
        
        If opCheque.Value Or optDepositoo.Value Then
            
            Grabar_Datos_Boveda
        Else
            
            Grabar_Datos_Caja
        End If
    
    End If
End Sub

Private Sub cmdImprimir_Click()
Dim Folio As Long, vbAnswer As Integer
    
    vbAnswer = MsgBox("Desea imprimir recibo de Caja ??" & vbCrLf & "Si selecciona No se imprimirá recibo de Bóveda.", vbQuestion + vbYesNo + vbDefaultButton2, "Bóveda/Caja") = vbYes
    Folio = frmReimpresionrecibos.ReImprimir(IIf(vbAnswer = -1, "boveda", "bancos"), "Folio", " WHERE Folio=")
    If Folio > 0 Then
        
        If vbAnswer = -1 Then
            
            Imprimir_Recibo_Caja Folio
        Else
            
            Imprimir_Recibo_Boveda Folio
        End If
        
    ElseIf Folio = 0 Then
        
        MsgBox "No se encontró el folio especificado !!", vbInformation, "Bóveda/Caja"
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Inicializar
End Sub

Private Sub Grabar_Datos()
Dim Movimiento As Long, Folio As Long, Importe As Double, Hora As String
  
    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Bóveda/Caja") = vbYes Then
        
        'Tomo la Hora
        Hora = Time
        
        'Saco el Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
        
        'Saco el Folio
        Folio = Regresa_Movimiento(False, IIf(opDeposito.Value, "FolioDepositos", "FolioTransferencias"))
        Regresa_Movimiento True, IIf(opDeposito.Value, "FolioDepositos", "FolioTransferencias")
    
        Importe = CDbl(txtImporte.text)
                          
        dbDatos.Execute "INSERT INTO bancos (Folio,Fecha,Deposito,Concepto,Importe,TipoMov,IDUsuario,IDSucursal) VALUES " _
                        & "(" & Folio & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & IIf(opDeposito.Value, 1, 0) & ",'" & IIf(opDeposito.Value, "Dotacion a bancos", "Retiro de Bancos") & "'," & ConvMoneda(Importe) & ",1," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
        'Grabamos el cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " _
                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'TR01','" & IIf(opDeposito.Value, "210101", "200901") & "'," & ConvMoneda(Importe) & "," & TIPO_CARGO & "," & IIf(opDeposito.Value, 1, 0) & ",'" & IIf(opDeposito.Value, "Dotacion a bancos", "Retiro de Bancos") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
        'Grabamos el abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " _
                        & "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'TR50','" & IIf(opDeposito.Value, "200950", "210150") & "'," & ConvMoneda(Importe) & "," & TIPO_ABONO & "," & IIf(opDeposito.Value, 1, 0) & ",'" & IIf(opDeposito.Value, "Dotacion a bancos", "Retiro de Bancos") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
        txtImporte.text = "0.00"
        Cargar_Montos
    
        lblFolio.Caption = Regresa_Movimiento(False, IIf(opDeposito.Value, "FolioDepositos", "FolioTransferencias"))
    End If

End Sub

'Validamos que esten correctos los datos
Private Function Validar() As Boolean
    
    Validar = True

    If Trim(txtImporte.text) = "" Then
        MsgBox "Imposible grabar el movimiento, Datos incompletos", vbOKOnly + vbCritical, "Bóveda/Caja"
        txtImporte.SetFocus
        Validar = False
        Exit Function
    End If

End Function

Private Sub opCheque_Click()
    lblFolio.Caption = Regresa_Movimiento(False, "FolioBancos")
    txtImporte_Change
End Sub

Private Sub opDeposito_Click()
    txtImporte_Change
    lblFolio.Caption = Regresa_Movimiento(False, "FolioDepositos")
End Sub

Private Sub opDotCaja_Click()
    lblFolio.Caption = Regresa_Movimiento(False, "FolioBoveda")
    txtImporte_Change
End Sub

Private Sub opRetCaja_Click()
    lblFolio.Caption = Regresa_Movimiento(False, "FolioBoveda")
    txtImporte_Change
End Sub

Private Sub optDepositoo_Click()
    lblFolio.Caption = Regresa_Movimiento(False, "FolioBancos")
    txtImporte_Change
End Sub

Private Sub opTransferencia_Click()
    txtImporte_Change
    lblFolio.Caption = Regresa_Movimiento(False, "FolioTransferencias")
End Sub

Private Sub txtImporte_Change()
Dim NuevoImporte As Double, Importe As Double
    
    If opDotCaja.Value Or opRetCaja.Value Then
        
        lblBoveda.Caption = Format(ImporteOriginal2, FMoneda)
        
        If opDotCaja.Value Then
        
            If txtImporte.text <> "" Then
            
                NuevoImporte = CDbl(txtImporte.text)
                lblBoveda.Caption = Format(ImporteOriginal2 - NuevoImporte, FMoneda)
            Else
                
                lblBoveda.Caption = Format(ImporteOriginal2, FMoneda)
            End If
        
        ElseIf opRetCaja.Value Then
            
            If txtImporte.text <> "" Then
                
                NuevoImporte = CDbl(txtImporte.text)
                lblBoveda.Caption = Format(ImporteOriginal2 + NuevoImporte, FMoneda)
            Else
                
                lblBoveda.Caption = Format(ImporteOriginal2, FMoneda)
            End If
        
        End If
    
    Else
        
        lblBancos.Caption = Format(ImporteOriginal, FMoneda)
        
        If optDepositoo.Value = False And ChkOtros.Value = 0 Then
         
            If txtImporte.text <> "" Then
            
                NuevoImporte = CDbl(txtImporte.text)
                Importe = ImporteOriginal2 + NuevoImporte
                lblBoveda.Caption = Format(Importe, FMoneda)
        
                NuevoImporte = CDbl(txtImporte.text)
                Importe = ImporteOriginal - NuevoImporte
                lblBancos.Caption = Format(Importe, FMoneda)
        
            Else
            
                lblBoveda.Caption = Format(ImporteOriginal2, FMoneda)
                lblBancos.Caption = Format(ImporteOriginal, FMoneda)
            End If
            
        ElseIf optDepositoo.Value = False And ChkOtros.Value = 1 Then
        
            If txtImporte.text <> "" Then
            
                NuevoImporte = CDbl(txtImporte.text)
                Importe = ImporteOriginal - NuevoImporte
                lblBancos.Caption = Format(Importe, FMoneda)
            Else
                
                lblBancos.Caption = Format(ImporteOriginal, FMoneda)
            End If
        
        ElseIf optDepositoo = True And ChkOtros.Value = 0 Then
        
            If txtImporte.text <> "" Then
                
                NuevoImporte = CDbl(txtImporte.text)
                Importe = ImporteOriginal2 - NuevoImporte
                lblBoveda.Caption = Format(Importe, FMoneda)
            Else
                
                lblBoveda.Caption = Format(ImporteOriginal2, FMoneda)
            End If
            
            If txtImporte.text <> "" Then
            
                NuevoImporte = CDbl(txtImporte.text)
                Importe = ImporteOriginal + NuevoImporte
                lblBancos.Caption = Format(Importe, FMoneda)
            Else
                
                lblBancos.Caption = Format(ImporteOriginal, FMoneda)
            End If
        
        ElseIf optDepositoo.Value = True And ChkOtros.Value = 1 Then
            
            If txtImporte.text <> "" Then
            
                NuevoImporte = CDbl(txtImporte.text)
                Importe = ImporteOriginal + NuevoImporte
                lblBancos.Caption = Format(Importe, FMoneda)
                lblBoveda.Caption = Format(ImporteOriginal2, FMoneda)
            Else
                
                lblBancos.Caption = Format(ImporteOriginal, FMoneda)
                lblBoveda.Caption = Format(ImporteOriginal2, FMoneda)
            End If
        
        End If
        
    End If
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

Sub Cargar_Montos()
Dim Cargo As Double, Abono As Double
Dim rcBD As New ADODB.Recordset

'''''    'Bancos***********************************
'''''    Cargo = 0
'''''    Abono = 0
'''''    rcBD.Open "SELECT Sum(Importe) AS Cargo FROM auxiliar WHERE Cuenta='210101'", dbDatos, adOpenForwardOnly, adLockOptimistic
'''''        Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
'''''    rcBD.Close
'''''
'''''    rcBD.Open "SELECT Sum(Importe) AS total FROM auxiliar WHERE Cuenta='210150'", dbDatos, adOpenForwardOnly, adLockOptimistic
'''''        Abono = IIf(IsNull(rcBD!Total), 0, rcBD!Total)
'''''    rcBD.Close
'''''
'''''    lblBancos.Caption = Format(Cargo - Abono, FMoneda)
'''''    lblBancos.Tag = Cargo - Abono
'''''    ImporteOriginal = Cargo - Abono
'''''    '****************************************

    'Boveda**********************************
    Cargo = 0
    Abono = 0
    rcBD.Open "SELECT Sum(Importe) AS Cargo FROM auxiliar WHERE Cuenta='110901'", dbDatos, adOpenForwardOnly, adLockOptimistic
        Cargo = IIf(IsNull(rcBD!Cargo), 0, rcBD!Cargo)
    rcBD.Close

    rcBD.Open "SELECT Sum(Importe) AS total FROM auxiliar WHERE Cuenta='110950'", dbDatos, adOpenForwardOnly, adLockOptimistic
        Abono = IIf(IsNull(rcBD!Total), 0, rcBD!Total)
    rcBD.Close
    
    lblBoveda.Caption = Format(Cargo - Abono, FMoneda)
    lblBoveda.Tag = Val(Cargo - Abono)
    ImporteOriginal2 = Cargo - Abono
    '*****************************************

    Set rcBD = Nothing
End Sub

'''''Private Sub txtImportee_Change()
'''''Dim NuevoImporte As Double, Importe As Double
'''''
''''' If optDepositoo.Value = False And ChkOtros.Value = 0 Then
'''''    If txtImporte.text <> "" Then
'''''        NuevoImporte = CDbl(txtImportee.text)
'''''        Importe = ImporteOriginal2 + NuevoImporte
'''''        lblBoveda.Caption = Format(Importe, "##,###0.00")
'''''
'''''        NuevoImporte = CDbl(txtImportee.text)
'''''        Importe = ImporteOriginal - NuevoImporte
'''''        lblBancos.Caption = Format(Importe, "##,###0.00")
'''''
'''''    Else
'''''        lblBoveda.Caption = Format(ImporteOriginal2, "##,###0.00")
'''''        lblBancos.Caption = Format(ImporteOriginal, "##,###0.00")
'''''    End If
'''''ElseIf optDepositoo.Value = False And ChkOtros.Value = 1 Then
'''''        If txtImportee.text <> "" Then
'''''            NuevoImporte = CDbl(txtImportee.text)
'''''            Importe = ImporteOriginal - NuevoImporte
'''''            lblBancos.Caption = Format(Importe, "##,###0.00")
'''''            'Exit Sub
'''''        Else
'''''            lblBancos.Caption = Format(ImporteOriginal, "##,###0.00")
'''''            'Exit Sub
'''''        End If
'''''
'''''ElseIf optDepositoo = True And ChkOtros.Value = 0 Then
'''''    If txtImportee.text <> "" Then
'''''        NuevoImporte = CDbl(txtImportee.text)
'''''        Importe = ImporteOriginal2 - NuevoImporte
'''''        lblBoveda.Caption = Format(Importe, "##,###0.00")
'''''        'Exit Sub
'''''   Else
'''''        lblBoveda.Caption = Format(ImporteOriginal2, "##,###0.00")
'''''        'Exit Sub
'''''    End If
'''''    If txtImportee.text <> "" Then
'''''        NuevoImporte = CDbl(txtImportee.text)
'''''        Importe = ImporteOriginal + NuevoImporte
'''''        lblBancos.Caption = Format(Importe, "##,###0.00")
'''''        'Exit Sub
'''''    Else
'''''        lblBancos.Caption = Format(ImporteOriginal, "##,###0.00")
'''''        'Exit Sub
'''''    End If
'''''
'''''ElseIf optDepositoo.Value = True And ChkOtros.Value = 1 Then
'''''    If txtImportee.text <> "" Then
'''''        NuevoImporte = CDbl(txtImportee.text)
'''''        Importe = ImporteOriginal + NuevoImporte
'''''        lblBancos.Caption = Format(Importe, "##,###0.00")
'''''        lblBoveda.Caption = Format(ImporteOriginal2, "##,###0.00")
'''''    Else
'''''        lblBancos.Caption = Format(ImporteOriginal, "##,###0.00")
'''''        lblBoveda.Caption = Format(ImporteOriginal2, "##,###0.00")
'''''    End If
'''''Else
'''''    If txtImportee.text <> "" Then
'''''        NuevoImporte = CDbl(txtImportee.text)
'''''        Importe = ImporteOriginal - NuevoImporte
'''''        lblBancos.Caption = Format(Importe, "##,###0.00")
'''''        'Exit Sub
'''''    Else
'''''        lblBancos.Caption = Format(ImporteOriginal, "##,###0.00")
'''''        'Exit Sub
'''''    End If
'''''End If
'''''End Sub

'Grabamos los datos
Private Sub Grabar_Datos_Boveda()
Dim Movimiento As Long, Folio As Long, Importe As Double, Hora As String

    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton1, "Bóveda/Caja") = vbYes Then
    
        'Saco el Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
    
        'Saco el Folio
        Folio = Regresa_Movimiento(False, "FolioBancos")
        Regresa_Movimiento True, "FolioBancos"
        
        'Tomo la Hora
        Hora = Time
    
        Importe = CDbl(txtImporte.text)
                  
        dbDatos.Execute "INSERT INTO bancos (Folio,Fecha,Deposito,Concepto,Importe,IDUsuario,IDSucursal) VALUES (" & _
                            Folio & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & IIf(optDepositoo.Value, 0, 1) & ",'" & IIf(opCheque.Value, "Dotacion a boveda", "Retiro de boveda") & "'," & ConvMoneda(Importe) & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
        If ChkOtros.Value = 0 Then
    
            'Grabamos el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optDepositoo.Value, "BA01", "BA01") & "','" & IIf(optDepositoo.Value, "210101", "110901") & "'," & ConvMoneda(Importe) & "," & TIPO_CARGO & "," & IIf(optDepositoo.Value, 0, 1) & ",'" & IIf(opCheque.Value, "Dotacion a boveda", "Retiro de boveda") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
            'Grabamos el abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optDepositoo.Value, "BA50", "BA50") & "','" & IIf(optDepositoo.Value, "110950", "210150") & "'," & ConvMoneda(Importe) & "," & TIPO_ABONO & "," & IIf(optDepositoo.Value, 0, 1) & ",'" & IIf(opCheque.Value, "Dotacion a boveda", "Retiro de boveda") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
        ElseIf optDepositoo.Value And ChkOtros.Value = 1 Then
    
            'Grabamos el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optDepositoo.Value, "BA01", "BA01") & "','" & IIf(optDepositoo.Value, "210101", "110001") & "'," & ConvMoneda(Importe) & "," & TIPO_CARGO & ",0,'" & IIf(opCheque.Value, "Dotacion a boveda", "Retiro de boveda") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
            'Grabamos el abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optDepositoo.Value, "BA50", "BA50") & "','" & IIf(optDepositoo.Value, "110050", "210150") & "'," & ConvMoneda(Importe) & "," & TIPO_ABONO & ",0,'" & IIf(opCheque.Value, "Dotacion a boveda", "Retiro de boveda") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"

        ElseIf ChkOtros.Value = 1 Then
    
            'Grabamos el cargo
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optDepositoo.Value, "BA01", "BA01") & "','" & IIf(optDepositoo.Value, "110001", "110001") & "'," & ConvMoneda(Importe) & "," & TIPO_CARGO & ",0,'" & IIf(opCheque.Value, "Dotacion a boveda", "Retiro de boveda") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
            'Grabamos el abono
            dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES ('" & _
                            Format(Date, "YYYY/MM/DD") & "','" & Format(Hora, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(optDepositoo.Value, "BA50", "BA50") & "','" & IIf(optDepositoo.Value, "110950", "210150") & "'," & ConvMoneda(Importe) & "," & TIPO_ABONO & ",0,'" & IIf(opCheque.Value, "Dotacion a boveda", "Retiro de boveda") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
        End If
    
        'Imprimo el Recibo
        Imprimir_Recibo_Boveda Folio
    
        txtImporte.text = "0.00"
        Cargar_Montos

        lblFolio.Caption = Regresa_Movimiento(False, "FolioBancos")
    End If

End Sub

'Grabamos los datos
Private Sub Grabar_Datos_Caja()
Dim Movimiento As Long, Folio As Long, Importe As Double

    If MsgBox("Estan correctos los datos ??", vbQuestion + vbYesNo + vbDefaultButton2, "Bóveda/Caja") = vbYes Then
        
        'Saco el Movimiento
        Movimiento = Regresa_Movimiento(False)
        Regresa_Movimiento True
        
        'Saco el Folio
        Folio = Regresa_Movimiento(False, "FolioBoveda")
        Regresa_Movimiento True, "FolioBoveda"
    
        Importe = CDbl(txtImporte.text)
    
        dbDatos.Execute "INSERT INTO boveda (Folio,Fecha,Deposito,Concepto,Importe,IDUsuario,IDSucursal) VALUES " & _
                        "(" & Folio & ",'" & Format(Now, "YYYY/MM/DD HH:MM:SS") & "'," & IIf(opDotCaja.Value, 1, 0) & ",'" & IIf(opDotCaja.Value, "Dotacion a Caja", "Retiro de Caja") & "'," & ConvMoneda(Importe) & "," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                      
'''        'Grabamos el cargo
'''        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
'''                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(opDotCaja.Value, "DO01", "RE01") & "','" & IIf(opDotCaja.Value, "199401", "110901") & "'," & ConvMoneda(Importe) & "," & TIPO_CARGO & "," & IIf(opDotCaja.Value, 1, 0) & ",'" & IIf(opDotCaja.Value, "Dotacion a Caja", "Retiro de Caja") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                    
        'Grabamos el cargo
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(opDotCaja.Value, "DO01", "RE01") & "','" & IIf(opDotCaja.Value, "110101", "110901") & "'," & ConvMoneda(Importe) & "," & TIPO_CARGO & "," & IIf(opDotCaja.Value, 1, 0) & ",'" & IIf(opDotCaja.Value, "Dotacion a Caja", "Retiro de Caja") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
                        
'''        'Grabamos el abono
'''        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
'''                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(opDotCaja.Value, "DO50", "RE50") & "','" & IIf(opDotCaja.Value, "110950", "199450") & "'," & ConvMoneda(Importe) & "," & TIPO_ABONO & "," & IIf(opDotCaja.Value, 1, 0) & ",'" & IIf(opDotCaja.Value, "Dotacion a Caja", "Retiro de Caja") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
        
        'Grabamos el abono
        dbDatos.Execute "INSERT INTO auxiliar (Fecha,Hora,Movimiento,Folio,Iniciales,Cuenta,Importe,Tipo,Serie,Concepto,PC,IDUsuario,IDSucursal) VALUES " & _
                        "('" & Format(Date, "YYYY/MM/DD") & "','" & Format(Time, "HH:MM:SS") & "'," & Movimiento & "," & Folio & ",'" & IIf(opDotCaja.Value, "DO50", "RE50") & "','" & IIf(opDotCaja.Value, "110950", "110150") & "'," & ConvMoneda(Importe) & "," & TIPO_ABONO & "," & IIf(opDotCaja.Value, 1, 0) & ",'" & IIf(opDotCaja.Value, "Dotacion a Caja", "Retiro de Caja") & "','" & NombrePc & "'," & frmMDI.IDUsuario & "," & frmMDI.IDSucursal & ")"
    
        'Saco el Recibo
        Imprimir_Recibo_Caja Folio
    
        Unload Me
    End If

End Sub

Function Imprimir_Recibo_Boveda(Folio As Long)
Dim ImprDefault As Boolean, crImporte As Double
    
    'Checo si hay impresora por default
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))
    crImporte = SacaValor("bancos", "Importe", " WHERE Cancelado=0 AND TipoMov=0 AND Folio=" & Folio)
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\NotaBoveda.rpt"
        .SelectionFormula = "{bancos.Folio}=" & Folio & " AND {bancos.TipoMov}=0 AND {bancos.Cancelado}=0"
        .Formulas(0) = "ImporteLetras='" & Trim(CantidadEnLetra(CCur(crImporte))) & "'"
        .Formulas(1) = "Caja='" & Trim(UCase(NombrePc)) & "'"
        .Formulas(2) = "Gerente='" & Trim(Regresa_Valor_BD("Gerente")) & "'"
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

Function Imprimir_Recibo_Caja(Folio As Long)
Dim Usuario As String, ImprDefault As Boolean, crImporte As Double, Operacion As Boolean
    
    'Checo si hay impresora por default
    ImprDefault = LocalizaImpresora(Regresa_Valor_BD("ImpresoraDefault"))

    Usuario = SacaValor("usuarios", "Nombre", " WHERE ID='" & Trim(frmMDI.IDUsuario) & "'")
    crImporte = SacaValor("boveda", "Importe", " WHERE Folio=" & Folio)
    Operacion = IIf(Val(SacaValor("boveda", "Deposito", " WHERE Folio=" & Folio)) = 1, True, False)
    
    With frmMDI.Cr
        .Reset
        .DiscardSavedData = True
        .WindowShowExportBtn = True
        .WindowShowPrintSetupBtn = True
        .Connect = "UID=" & USERBD & ";PWD=" & PWDBD & ";DRIVER={MySQL ODBC 3.51 Driver};SERVER=" & sServidor
        .ReportFileName = Path & "\Reportes\NotaCaja.rpt"
        .SelectionFormula = "{boveda.Folio}=" & Folio & ""
        .Formulas(0) = "ImporteLetras='" & Trim(CantidadEnLetra(CCur(crImporte))) & "'"
        .Formulas(1) = "Recibido='" & IIf(Operacion, "CAJA", "BOVEDA") & " " & IIf(Operacion, Usuario, Regresa_Valor_BD("Gerente")) & "'"
        .Formulas(2) = "Enviado='" & IIf(Operacion = False, "CAJA", "BOVEDA") & " " & IIf(Operacion = False, Usuario, Regresa_Valor_BD("Gerente")) & "'"
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
