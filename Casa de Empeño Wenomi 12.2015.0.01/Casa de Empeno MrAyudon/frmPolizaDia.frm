VERSION 5.00
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmPolizaDia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Poliza del Dia"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4380
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPolizaDia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3510
   ScaleWidth      =   4380
   Begin VB.CheckBox chkEgresos 
      Appearance      =   0  'Flat
      Caption         =   "Generar Poliza de Egresos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   2295
   End
   Begin VB.CheckBox chkIngresos 
      Appearance      =   0  'Flat
      Caption         =   "Generar Poliza de Ingresos"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox txtPolizaEgresos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   2760
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtPolizaIngresos 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   1455
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Height          =   375
      Left            =   3090
      TabIndex        =   0
      Top             =   3000
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
      Picture         =   "frmPolizaDia.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdAceptar 
      Height          =   375
      Left            =   1920
      TabIndex        =   1
      Top             =   3000
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   1
      TextColor       =   8537065
      Object.ToolTipText     =   ""
      Picture         =   "frmPolizaDia.frx":055E
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   4095
      Begin VB.OptionButton opCajaIngresos 
         Appearance      =   0  'Flat
         Caption         =   "Generar Poliza por Caja"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   720
         Width           =   2535
      End
      Begin VB.OptionButton opGlobalIngresos 
         Appearance      =   0  'Flat
         Caption         =   "Generar Poliza por Sucursal"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   1800
      Width           =   4095
      Begin VB.OptionButton opGlobalEgresos 
         Appearance      =   0  'Flat
         Caption         =   "Generar Poliza por Sucursal"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   13
         Top             =   360
         Value           =   -1  'True
         Width           =   2535
      End
      Begin VB.OptionButton opCajaEgresos 
         Appearance      =   0  'Flat
         Caption         =   "Generar Poliza por Caja"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "No. Poliza Egresos:"
      Height          =   195
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   1380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "No. Poliza Ingresos:"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   1440
   End
End
Attribute VB_Name = "frmPolizaDia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const cIngresos = 1
Private Const cEgresos = 2
Private Const cDiario = 3
Private Const cDeOrden = 4
Private Const cEstadisticas = 5

Private Const cNormal = 1
Private Const cSinAfectar = 2

Private Const cCargo = 1
Private Const cAbono = 2


Dim Fl() As New cFlatControl

Private Sub chkEgresos_Click()
   opGlobalEgresos.Enabled = (chkEgresos.Value = vbChecked)
   opCajaEgresos.Enabled = (chkEgresos.Value = vbChecked)
End Sub

Private Sub chkIngresos_Click()
   opGlobalIngresos.Enabled = (chkIngresos.Value = vbChecked)
   opCajaIngresos.Enabled = (chkIngresos.Value = vbChecked)
End Sub

Private Sub cmdAceptar_Click()
   If Validar Then CrearPoliza
End Sub

Public Function Validar() As Boolean
   Validar = True
   
   If chkIngresos.Value = vbChecked Then
      If Not IsNumeric(Me.txtPolizaIngresos.text) Then
         Validar = False
         MsgBox "Favor de poner un numero de poliza de ingreso valido", vbCritical Or vbOKOnly
         Exit Function
      End If
   End If
   
   If chkEgresos.Value = vbChecked Then
      If Not IsNumeric(Me.txtPolizaEgresos.text) Then
         Validar = False
         MsgBox "Favor de poner un numero de poliza de egreso valido", vbCritical Or vbOKOnly
         Exit Function
      End If
   End If
   
End Function

Private Sub cmdSalir_Click()
   Unload Me
End Sub

Private Sub Inicializar()
   CentrarForm Me, frmMDI
End Sub

Private Sub Form_Load()
   Inicializar
   Poner_Flat Fl, Me.Controls, Me
End Sub

Private Sub txtPolizaEgresos_GotFocus()
   Seleccionar_Texto txtPolizaEgresos
   Cambiar_Color True, txtPolizaEgresos
End Sub

Private Sub txtPolizaEgresos_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtPolizaEgresos_LostFocus()
   Cambiar_Color False, txtPolizaEgresos
End Sub

Private Sub txtPolizaIngresos_GotFocus()
   Cambiar_Color True, txtPolizaIngresos
   Seleccionar_Texto txtPolizaIngresos
End Sub

Private Sub txtPolizaIngresos_KeyPress(KeyAscii As Integer)
   Pasar_Foco KeyAscii
End Sub

Private Sub txtPolizaIngresos_LostFocus()
   Cambiar_Color False, txtPolizaIngresos
End Sub

Private Sub CrearPoliza()
   Screen.MousePointer = vbHourglass
   
   'si estan echos los cortes en todas las maquinas, se realiza la poliza diaria
   If Not ValidarCorte Then
      CrearPolizas
      MsgBox "Las polizas han sido creadas", vbInformation Or vbOKOnly
   Else
      MsgBox "Favor de realizar el corte en todas las cajas", vbInformation Or vbOKOnly
   End If
   
   Screen.MousePointer = vbDefault
End Sub

Private Sub CrearPolizas()
   
   If chkIngresos.Value = vbChecked Then CreaPolizas True
   
   If chkEgresos.Value = vbChecked Then CreaPolizas False
   
End Sub

Private Sub CreaPolizas(Ingresos As Boolean)
   On Error GoTo Error
   Dim Sql As String
   Dim rc As New ADODB.Recordset
   Dim Pcs() As String
   Dim Contador As Long
   
   Contador = 0
      
   If opGlobalIngresos.Value Then
      Sql = ""
   ElseIf opCajaIngresos.Value Then
      Sql = "SELECT DISTINCT PC FROM Auxiliar WHERE Fecha='" & Format(Date, "YYYY/MM/DD") & "' ORDER BY ID"
   End If
   
   'si sql esta vacio, tons es contabilidad global
   If Sql <> "" Then
      rc.Open Sql, dbDatos, adOpenDynamic, adLockOptimistic
      While Not rc.EOF
         DoEvents
         ReDim Preserve Pcs(Contador)
         Pcs(Contador) = rc!PC
         Contador = Contador + 1
         rc.MoveNext
      Wend
      rc.Close
   Else
      ReDim Preserve Pcs(0)
      Pcs(0) = ""
   End If
   
   If Ingresos Then
      CreaPolizasIngresos Pcs
   Else
      CreaPolizasEgresos Pcs
   End If
   
Error:
      Maneja_Error Err
      
   Set rc = Nothing

End Sub

Private Sub CreaPolizasIngresos(Pcs() As String)
   On Error GoTo Error
   Dim lArchivo As Long
   Dim Indice As Long
   Dim strArchivo As String
   Dim NoPoliza As Long
   Dim Concepto As String
   Dim Importe As Currency
   Dim CtaCargo As String
   Dim CtaAbono As String
   
   NoPoliza = Val(txtPolizaIngresos.text)
   
   
   For Indice = LBound(Pcs) To UBound(Pcs)
      DoEvents
      strArchivo = "PI-" & NoPoliza & ".txt"
   
      lArchivo = Grabar_Archivo(strArchivo)
      
      Concepto = IIf(Pcs(Indice) = "", "Poliza de ingreso de la sucursal", "Poliza de ingreso de la caja " & Pcs(Indice))
      Crear_Encabezado lArchivo, CStr(NoPoliza), Concepto, cIngresos
      
      'desempeño tradicional
      Importe = GetSaldo("Cuenta='201750' AND Concepto='Desempeño' AND (Serie=1 OR Serie=2)", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("201750") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("201701") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Desempeño Tradicional", lArchivo
      
      'abono refrendo tradicional
      Importe = GetSaldo("Cuenta='201750' AND Concepto='Abono Refrendo'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("201750") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("201701") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Abono Refrendo Tradicional", lArchivo
      
      'intereses tradicional
      Importe = GetSaldo("(Cuenta='520450' OR Cuenta='670350' OR Cuenta='680350' OR Cuenta='690350') AND (Concepto='Refrendo' OR Concepto='Desempeño')", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("520450") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("520401") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Interes Tradicional", lArchivo
      
      'iva intereses tradicional
      Importe = GetSaldo("Cuenta='120150' AND (Concepto='Refrendo' OR Concepto='Desempeño')", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("120150") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("120101") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "IVA Interes Tradicional", lArchivo
      
      'desempeño pagos fijos
      Importe = GetSaldo("Cuenta='201750' AND Concepto='Pagos Fijos'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("201750") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("201701") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Desempeño Pagos Fijos", lArchivo
      
      'intereses pagos fijos
      Importe = GetSaldo("(Cuenta='520450' OR Cuenta='670350' OR Cuenta='680350') AND Concepto='Pagos Fijos'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("520450") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("520401") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Intereses Pagos Fijos", lArchivo
      
      'intereses moratorios
      Importe = GetSaldo("Cuenta='690350' AND Concepto='Pagos Fijos'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("690350") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("690301") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Intereses Moratorios", lArchivo
            
      'iva intereses pagos fijos
      Importe = GetSaldo("Cuenta='120150' AND Concepto='Pagos Fijos'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("120150") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("120101") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "IVA Intereses Pagos Fijos", lArchivo
      
      'ventas
      Importe = GetSaldo("Cuenta='620450' AND Iniciales='VT03' AND Concepto='Ventas'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("620450") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("620401") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Ventas", lArchivo
      
      'iva ventas
      Importe = GetSaldo("Cuenta='120150' AND Concepto='Ventas'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("120150") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("120101") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "IVA Ventas", lArchivo
      
      
      'abono a apartados
      Importe = GetSaldo("Cuenta='110101' AND (Iniciales='AP03' OR Iniciales='AB05') AND (Concepto='Apartado' OR Concepto='Abonos')", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("110101") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("110150") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Abono a Apartados", lArchivo
      
      'venta divisas
      Importe = GetSaldo("Cuenta='710301' AND Iniciales='CD01' AND Serie=1", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("710301") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("710350") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Venta Divisas", lArchivo
      
      'otros cobros
      Importe = GetSaldo("(Cuenta='530150' OR Cuenta='120150') AND Concepto='Boleta perdida'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("530150") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("530150") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PI-" & NoPoliza, Importe, "Otros Cobros", lArchivo
   
      Close #lArchivo
      NoPoliza = NoPoliza + 1
   Next Indice
   
Error:
   Maneja_Error Err

End Sub

Private Sub CreaPolizasEgresos(Pcs() As String)
   On Error GoTo Error
   Dim lArchivo As Long
   Dim Indice As Long
   Dim strArchivo As String
   Dim NoPoliza As Long
   Dim Concepto As String
   Dim Importe As Currency
   Dim CtaCargo As String
   Dim CtaAbono As String
   
   NoPoliza = Val(txtPolizaEgresos.text)
   
   
   For Indice = LBound(Pcs) To UBound(Pcs)
      DoEvents
      strArchivo = "PE-" & NoPoliza & ".txt"
   
      lArchivo = Grabar_Archivo(strArchivo)
      
      Concepto = IIf(Pcs(Indice) = "", "Poliza de egresos de la sucursal", "Poliza de egresos de la caja " & Pcs(Indice))
      Crear_Encabezado lArchivo, CStr(NoPoliza), Concepto, cEgresos
            
      
'''      'empeño tradicional
'''      Importe = GetSaldo("Cuenta='199450' AND Concepto='Empeño' AND (Serie=1 OR Serie=2)", Pcs(Indice))
'''      CtaCargo = GetCuentaContpaq("199450") 'Cuenta Cargo
'''      CtaAbono = GetCuentaContpaq("199401") 'Cuenta Abono
'''      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Desempeño Tradicional", lArchivo
'''
'''
'''      'empeño fijo
'''      Importe = GetSaldo("Cuenta='199450' AND Concepto='Empeño' AND Serie=3", Pcs(Indice))
'''      CtaCargo = GetCuentaContpaq("199450") 'Cuenta Cargo
'''      CtaAbono = GetCuentaContpaq("199401") 'Cuenta Abono
'''      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Abono Refrendo Tradicional", lArchivo
      
      'empeño tradicional
      Importe = GetSaldo("Cuenta='110150' AND Concepto='Empeño' AND (Serie=1 OR Serie=2)", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("110150") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("110101") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Desempeño Tradicional", lArchivo
      
      
      'empeño fijo
      Importe = GetSaldo("Cuenta='110150' AND Concepto='Empeño' AND Serie=3", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("110150") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("110101") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Abono Refrendo Tradicional", lArchivo
      
      'Compra Divisas
      Importe = GetSaldo("Cuenta='710350' AND Iniciales='VD50' AND Serie=1", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("710350") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("710301") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Interes Tradicional", lArchivo
      
      'compra varios
      Importe = GetSaldo("Cuenta='620301' AND Iniciales='EN01'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("620350") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("620301") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "IVA Interes Tradicional", lArchivo
      
      'iva compras
      Importe = GetSaldo("Cuenta='120101'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("120150") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("120101") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Desempeño Pagos Fijos", lArchivo
      
'''      'gastos
'''      Importe = GetSaldo("Cuenta='199450' AND Iniciales='GA50'", Pcs(Indice))
'''      CtaCargo = GetCuentaContpaq("199450") 'Cuenta Cargo
'''      CtaAbono = GetCuentaContpaq("199401") 'Cuenta Abono
'''      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Intereses Pagos Fijos", lArchivo
      
      'gastos
      Importe = GetSaldo("Cuenta='110150' AND Iniciales='GA50'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("110150") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("110101") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Intereses Pagos Fijos", lArchivo
      
      'demasias
      Importe = GetSaldo("Cuenta='650201'", Pcs(Indice))
      CtaCargo = GetCuentaContpaq("650250") 'Cuenta Cargo
      CtaAbono = GetCuentaContpaq("650201") 'Cuenta Abono
      Escribe_Movimiento_Archivo CtaCargo, CtaAbono, "PE-" & NoPoliza, Importe, "Intereses Moratorios", lArchivo
   
      Close #lArchivo
      NoPoliza = NoPoliza + 1
   Next Indice
   
Error:
   Maneja_Error Err
End Sub

'Regresa la cuenta de contabilidad relacionada con la cuenta del sistema
Private Function GetCuentaContpaq(Cuenta As String) As String
   On Error GoTo Error
   Dim rc As New ADODB.Recordset
   Dim Cta As String
   
   Cta = ""
   
   rc.Open "SELECT CuentaContpaq FROM Cuentas WHERE Cuenta='" & Cuenta & "'", dbDatos, adOpenDynamic, adLockOptimistic
   
   If Not rc.EOF Then Cta = rc!CuentaContpaq
      
   GetCuentaContpaq = Cta
   
   rc.Close
Error:
   Maneja_Error Err
   
   Set rc = Nothing
End Function


'Regresa el saldo segun los filtros especificados
Private Function GetSaldo(Filtro As String, Optional PC As String = "") As Currency
   On Error GoTo Error
   Dim rc As New ADODB.Recordset
   Dim Saldo As Currency
   Dim Sql As String
   
   Saldo = 0
      
   Sql = "SELECT SUM(Importe) AS Saldo " & _
         "FROM Auxiliar " & _
         "WHERE Fecha='" & Format(Now, "yyyy/mm/dd") & "'" & IIf(PC <> "", " AND PC='" & PC & "'", "") & " " & _
         " AND " & Filtro & " " & _
         "ORDER BY ID"
      
   rc.Open Sql, dbDatos, adOpenDynamic, adLockOptimistic
   
   If Not rc.EOF Then Saldo = Val(rc!Saldo & "")
      
   GetSaldo = Saldo
   
   rc.Close
Error:
   Maneja_Error Err

   Set rc = Nothing
End Function


'valida que se hayan hecho el corte en todas las cajas
Private Function ValidarCorte() As Boolean
   On Error GoTo Error
   Dim rc As New ADODB.Recordset
   
   ValidarCorte = False
   
   rc.Open "SELECT Count(ID) AS Total FROM Auxiliar WHERE Corte=0 AND Fecha='" & Format(Now, "yyyy/mm/dd") & "'", dbDatos, adOpenDynamic, adLockOptimistic
   
   If Not rc.EOF Then ValidarCorte = rc!Total > 0
         
   rc.Close
Error:
      Maneja_Error Err
   
   Set rc = Nothing
End Function


'Creamos el archivo de la poliza
Private Function Grabar_Archivo(Archivo As String) As Long
   Dim strArchivo As String
   Dim iArchivo As Long
   
   'strArchivo = "C:\Polizas\PI" & Poliza & ".txt"
   strArchivo = App.Path & "\Polizas\" & Archivo
   
   If Dir(App.Path & "\Polizas", vbDirectory) = "" Then MkDir App.Path & "\Polizas"
   
   iArchivo = FreeFile
   Open strArchivo For Output Access Write As #iArchivo
   Grabar_Archivo = iArchivo
      
   
End Function

'Creamos el encabezado del archivo
Private Sub Crear_Encabezado(iArchivo As Long, Poliza As String, strConcepto As String, TipoPoliza As Integer)
   Dim Encabezado As String
   Dim Fecha As String
   Dim DiarioAgrupador As String
   Dim strPoliza As String
   Dim Concepto As String
   Dim Diario As String
   Dim Sistema As String
      
   strPoliza = Val(Poliza)
   strPoliza = Mid("00000000", 1, 8 - Len(strPoliza)) & strPoliza 'asignamos todos los espacios a la poliza
   
   Fecha = Format(Now, "YYYYMMDD") 'Format(Regresa_Fecha(RenIni, RenFin, Colini), "YYYYMMDD")
   Fecha = Fecha & Space(8 - Len(Fecha))
   
   Concepto = strConcepto 'Regresa_Concepto(RenIni, RenFin, Colini)
   Concepto = Concepto & Space(100 - Len(Concepto))
   
   Diario = "000"
   Sistema = "10"
   Encabezado = "P " & Fecha & " " & Format(TipoPoliza, "000") & " " & strPoliza & " " & cNormal & " " & Diario & " " & Concepto & " " & Sistema & " " & "2" & " "
   
   Print #iArchivo, Encabezado
End Sub

'Escribimos los movimientos del archivo
Public Sub Escribe_Movimiento_Archivo(CtaCargo As String, CtaAbono As String, Poliza As String, Importe As Currency, Descripcion As String, lArchivo As Long)
   Dim strCta As String
   Dim strPoliza As String
   Dim strImporte As String
   Dim strCadena As String
   Dim TipoCargoAbono As Integer
   Const Movimiento = "M"
      
   If Importe > 0 Then
      strCta = Replace(CtaCargo, "-", "")
      strCta = strCta & Space(20 - Len(strCta))
      strPoliza = Poliza
      strPoliza = strPoliza & Space(10 - Len(strPoliza))
      strImporte = Quitar_Simbolos(Format(Importe, "Currency"))
      strImporte = Space(16 - Len(strImporte)) & strImporte
      Descripcion = Mid(Descripcion, 1, 30)
      Descripcion = Descripcion & Space(30 - Len(Descripcion))
      
      'Grabamos El Cargo
      strCadena = Movimiento & " " & strCta & " " & strPoliza & " " & cCargo & " " & strImporte & Space(22) & Descripcion & " " '& Chr(10) & Chr(13)
      Print #lArchivo, strCadena
      
      'Grabamos El Abono
      strCta = Replace(CtaAbono, "-", "")
      strCta = strCta & Space(20 - Len(strCta))
      strCadena = Movimiento & " " & strCta & " " & strPoliza & " " & cAbono & " " & strImporte & Space(22) & Descripcion & " " '& Chr(10) & Chr(13)
      Print #lArchivo, strCadena
   End If
   
End Sub

'quita todo los simbolos  monetarios
Public Function Quitar_Simbolos(Cadena As String) As String
   Quitar_Simbolos = Format(Cadena, "#########0.00")
End Function
