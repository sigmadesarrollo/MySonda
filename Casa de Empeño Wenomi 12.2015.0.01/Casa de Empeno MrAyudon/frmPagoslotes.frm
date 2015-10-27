VERSION 5.00
Object = "{543749C9-8732-11D3-A204-0090275C8BC1}#1.1#0"; "vbalGrid6.ocx"
Object = "{48DC3C96-B20F-11D1-A87F-D9394DC38340}#2.11#0"; "FlatBtn6.ocx"
Begin VB.Form frmPagosFijos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pagos"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPagoslotes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7920
   ScaleWidth      =   11460
   Begin VB.Frame Frame2 
      Height          =   2730
      Left            =   120
      TabIndex        =   11
      Top             =   5040
      Width           =   3855
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2040
         TabIndex        =   37
         Top             =   2400
         Width           =   1215
      End
      Begin VB.TextBox txtCargosadicionales 
         Alignment       =   1  'Right Justify
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   2040
         TabIndex        =   21
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Interés mensual: $"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   300
         TabIndex        =   36
         Top             =   1320
         Width           =   1590
      End
      Begin VB.Label lblInteres 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   35
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label lblMoratorios 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   32
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label lblPagomensual 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label lblFecha 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lblRecibonum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Total a pagar:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   300
         Left            =   195
         TabIndex        =   17
         Top             =   2355
         Width           =   1695
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Cargos adicionales: $"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2040
         Width           =   1770
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Moratorios: $"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   765
         TabIndex        =   15
         Top             =   1680
         Width           =   1125
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Pago mensual: $"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   510
         TabIndex        =   14
         Top             =   960
         Width           =   1380
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1350
         TabIndex        =   13
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Recibo num.:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   810
         TabIndex        =   12
         Top             =   240
         Width           =   1080
      End
   End
   Begin vbAcceleratorGrid6.vbalGrid grdPagos 
      Height          =   3255
      Left            =   30
      TabIndex        =   10
      Top             =   1800
      Width           =   11385
      _ExtentX        =   20082
      _ExtentY        =   5741
      RowMode         =   -1  'True
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderFlat      =   -1  'True
      BorderStyle     =   2
      ScrollBarStyle  =   1
      DisableIcons    =   -1  'True
   End
   Begin VB.Frame Frame1 
      Height          =   1635
      Left            =   45
      TabIndex        =   4
      Top             =   45
      Width           =   10380
      Begin VB.TextBox txtNombre 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   4110
      End
      Begin VB.TextBox txtEmpresa 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   2160
         TabIndex        =   1
         Top             =   570
         Width           =   4110
      End
      Begin VB.TextBox txtTelcasa 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   930
         TabIndex        =   2
         Top             =   930
         Width           =   1575
      End
      Begin VB.TextBox txtTeloficina 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3810
         TabIndex        =   3
         Top             =   930
         Width           =   1695
      End
      Begin DevPowerFlatBttn.FlatBttn cmdBuscar 
         Height          =   210
         Left            =   6285
         TabIndex        =   5
         Top             =   240
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   370
         AlignCaption    =   4
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
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Tasa (%):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6990
         TabIndex        =   31
         Top             =   1305
         Width           =   1020
      End
      Begin VB.Label lblTasa 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8040
         TabIndex        =   30
         Top             =   1305
         Width           =   1455
      End
      Begin VB.Label lblPlazo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8040
         TabIndex        =   29
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Plazo (Meses):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6450
         TabIndex        =   28
         Top             =   960
         Width           =   1560
      End
      Begin VB.Label lblPago 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8040
         TabIndex        =   27
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Pago:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7380
         TabIndex        =   26
         Top             =   600
         Width           =   630
      End
      Begin VB.Label lblMonto 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   8040
         TabIndex        =   25
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Saldo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7320
         TabIndex        =   24
         Top             =   240
         Width           =   690
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre persona física:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   9
         Top             =   210
         Width           =   1920
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre persona moral:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   8
         Top             =   570
         Width           =   1980
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tel. Oficina:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2730
         TabIndex        =   7
         Top             =   930
         Width           =   960
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tel. Casa:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   930
         Width           =   795
      End
   End
   Begin DevPowerFlatBttn.FlatBttn cmdSalir 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   10290
      TabIndex        =   22
      Top             =   7335
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
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      TextColor       =   -2147483630
      Object.ToolTipText     =   ""
      Picture         =   "frmPagoslotes.frx":000C
   End
   Begin DevPowerFlatBttn.FlatBttn cmdImprimir 
      Height          =   375
      Left            =   7800
      TabIndex        =   23
      Top             =   7335
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "    &Imprimir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmPagoslotes.frx":009D
   End
   Begin DevPowerFlatBttn.FlatBttn cmdLimpiar 
      Height          =   375
      Left            =   6675
      TabIndex        =   33
      Top             =   7335
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   661
      AlignCaption    =   3
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "   &Limpiar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   255
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmPagoslotes.frx":01B5
   End
   Begin DevPowerFlatBttn.FlatBttn cmdReimprimir 
      Height          =   375
      Left            =   9000
      TabIndex        =   34
      Top             =   7335
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      AlignCaption    =   4
      AlignPicture    =   2
      AutoSize        =   0   'False
      Caption         =   "      &Re-imprimir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaskColor       =   16777215
      MousePointer    =   1
      PlaySounds      =   0   'False
      Object.ToolTipText     =   ""
      Picture         =   "frmPagoslotes.frx":02B9
   End
End
Attribute VB_Name = "frmPagosFijos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Fl() As cFlatControl

Dim consulta As ADODB.Recordset

Private Sub cmdBuscar_Click()
    frmMostrarCliente.ver Me, txtNombre, , , 2
End Sub

Private Sub cmdImprimir_Click()
Dim Movimiento As Double, Amortizacion As Double, Intereses As Double, Adicionales As Double, Adicionaless As String, Moratorios As Double, Moratorioss As String
Dim pagoo As String, pago As Double, monto As Double, rcSolicitud As Long
Dim rcTotal As Double, rcDiferencia As Double, rcAdelanto As Double
Dim rcConsulta As New ADODB.Recordset

    rcTotal = Format(grdPagos.CellText(grdPagos.SelectedRow, 3) + grdPagos.CellText(grdPagos.SelectedRow, 4), "##,###0.00")

    If grdPagos.Rows > 0 And grdPagos.SelectedRow > 0 Then

        Set consulta = dbDatos.Execute("select pagado,idsolicitud from recibos where id=" & grdPagos.CellItemData(grdPagos.SelectedRow, 1) & "")
        rcSolicitud = consulta!IDSolicitud
    
        If consulta!pagado = False Then

            Movimiento = Regresa_Movimiento(False, "Movimiento")
            Regresa_Movimiento True, "Movimiento"

            'Folio del recibo
            Regresa_Movimiento True, "Foliorecibo"
            rcAdelanto = Val(txtTotal.Text)

            If rcAdelanto > rcTotal Then 'PAGO + ADELANTO
    
                rcConsulta.Open "select * from recibos where idsolicitud=" & consulta!IDSolicitud, dbDatos, adOpenDynamic, adLockOptimistic
    
                While Not rcConsulta.BOF And Not rcConsulta.EOF

                    If rcConsulta!pagado = False And rcAdelanto > 0 Then
                        If rcAdelanto >= (rcConsulta!monto + rcConsulta!Interes) Then
                            rcAdelanto = rcAdelanto - (rcConsulta!monto + rcConsulta!Interes)
                            rcConsulta!pagado = True
                            rcConsulta!Folio = CLng(lblRecibonum.Caption)
                            rcConsulta.Update
                            Amortizacion = Amortizacion + rcConsulta!Amortizacion
                            Intereses = Intereses + rcConsulta!Interes
                            pago = pago + (rcConsulta!monto + rcConsulta!Interes)
                        Else
                            pago = pago + rcAdelanto
                            rcConsulta!monto = rcConsulta!monto + rcConsulta!Interes
                            rcConsulta!monto = rcConsulta!monto - rcAdelanto
                            rcConsulta!Interes = 0
                            rcAdelanto = 0
                            rcConsulta.Update
                        End If
                    End If

                    rcConsulta.MoveNext
                Wend
                monto = pago
                rcConsulta.Close
            ElseIf Val(txtTotal.Text) < rcTotal Then  'ABONO
                rcDiferencia = rcTotal - Val(txtTotal.Text)
                dbDatos.Execute "update recibos set monto=" & rcDiferencia & ", interes=0 where id=" & grdPagos.CellItemData(grdPagos.SelectedRow, 1) & ""
                monto = txtTotal.Text
                pago = txtTotal.Text
                '    'Pago******************
                '        'Abono
                '        dbDatos.Execute "insert into auxiliar(fecha,movimiento,concepto,folio,iniciales,cuenta,importe,tipo,serie,pc) values (#" & Format(Date, "mm/dd/yyyy") & "#," & movimiento & ",'Abonos'," & lblRecibonum.Caption & ",'PA50','201750'," & pago & "," & TIPO_ABONO & ",0,'" & Trim(Nombre_PC) & "')"
                '
            Else    'PAGAR EL MES SOLAMENTE
                'Amortizacion
                Amortizacion = grdPagos.CellText(grdPagos.SelectedRow, 3) - grdPagos.CellText(grdPagos.SelectedRow, 4)

                'Intereses
                Intereses = grdPagos.CellText(grdPagos.SelectedRow, 4)

                'Moratorios
                Moratorioss = lblMoratorios.Caption
                Moratorios = IIf(Trim(Moratorioss) = "", 0, Moratorioss)

                'Cargos adicionales
                Adicionaless = txtCargosadicionales.Text
                Adicionales = IIf(Trim(Adicionaless) = "", 0, Adicionaless)

                'Pago mensual
                pagoo = lblPagomensual.Caption
                pago = IIf(Trim(pagoo) = "", 0, pagoo)
        
                dbDatos.Execute "update recibos set folio=" & CLng(lblRecibonum.Caption) & ",pagado=True,fechapago=#" & Format(Date, "mm/dd/yyyy") & "#,montopagado=" & CDbl(txtTotal.Text) & ",moratorios=" & Moratorios & ",adicionales=" & Adicionales & " where id=" & grdPagos.CellItemData(grdPagos.SelectedRow, 1) & ""
                grdPagos.CellText(grdPagos.SelectedRow, 8) = txtTotal.Text
                monto = txtTotal.Text
            End If

            grdPagos.CellText(grdPagos.SelectedRow, 6) = CLng(lblRecibonum.Caption)
            '        'Pago******************
            '        'Abono
            '        dbDatos.Execute "insert into auxiliar(fecha,movimiento,concepto,folio,iniciales,cuenta,importe,tipo,serie,pc) values (#" & Format(Date, "mm/dd/yyyy") & "#," & movimiento & ",'Pago mensual'," & lblRecibonum.Caption & ",'PA50','201750'," & pago & "," & TIPO_ABONO & ",0,'" & Trim(Nombre_PC) & "')"
            '
            '        'Amortizacion******************
            '        'Cargo
            '        dbDatos.Execute "insert into auxiliar(fecha,movimiento,concepto,folio,iniciales,cuenta,importe,tipo,serie,pc) values (#" & Format(Date, "mm/dd/yyyy") & "#," & movimiento & ",'Amortizacion'," & lblRecibonum.Caption & ",'PA01','110101'," & amortizacion & "," & TIPO_CARGO & ",0,'" & Trim(Nombre_PC) & "')"
            '        'Abono
            '        dbDatos.Execute "insert into auxiliar(fecha,movimiento,concepto,folio,iniciales,cuenta,importe,tipo,serie,pc) values (#" & Format(Date, "mm/dd/yyyy") & "#," & movimiento & ",'Amortizacion'," & lblRecibonum.Caption & ",'PA50','201750'," & amortizacion & "," & TIPO_ABONO & ",0,'" & Trim(Nombre_PC) & "')"
            '
            '        'Intereses*****************
            '        'Cargo
            '        dbDatos.Execute "insert into auxiliar(fecha,movimiento,concepto,folio,iniciales,cuenta,importe,tipo,serie,pc) values (#" & Format(Date, "mm/dd/yyyy") & "#," & movimiento & ",'Intereses'," & lblRecibonum.Caption & ",'PA01','110101'," & intereses & "," & TIPO_CARGO & ",0,'" & Trim(Nombre_PC) & "')"
            '        'Abono
            '        dbDatos.Execute "insert into auxiliar(fecha,movimiento,concepto,folio,iniciales,cuenta,importe,tipo,serie,pc) values (#" & Format(Date, "mm/dd/yyyy") & "#," & movimiento & ",'Intereses'," & lblRecibonum.Caption & ",'PA50','520450'," & intereses & "," & TIPO_ABONO & ",0,'" & Trim(Nombre_PC) & "')"
            '
            '        'Cargos adicionales
            '        If adicionales > 0 Then
            '            'Cargo
            '            dbDatos.Execute "insert into auxiliar(fecha,movimiento,concepto,folio,iniciales,cuenta,importe,tipo,serie,pc) values (#" & Format(Date, "mm/dd/yyyy") & "#," & movimiento & ",'Cargos adicionales'," & lblRecibonum.Caption & ",'PA01','110101'," & adicionales & "," & TIPO_CARGO & ",0,'" & Trim(Nombre_PC) & "')"
            '            'Abono
            '            dbDatos.Execute "insert into auxiliar(fecha,movimiento,concepto,folio,iniciales,cuenta,importe,tipo,serie,pc) values (#" & Format(Date, "mm/dd/yyyy") & "#," & movimiento & ",'Cargos adicionales'," & lblRecibonum.Caption & ",'PA50','201450'," & adicionales & "," & TIPO_ABONO & ",0,'" & Trim(Nombre_PC) & "')"
            '        End If
    
            Set consulta = dbDatos.Execute("select clientes.calle,clientes.colonia,clientes.ciudad,clientes.estado,clientes.telcasa,solicitudes.numcontrato from clientes inner join solicitudes on clientes.id=solicitudes.idcliente where solicitudes.id=" & txtNombre.Tag & "")

            With frmMDI.Cr
                .Reset
                .ReportFileName = Path & "\Reportes\Pagare.rpt"
                .DiscardSavedData = True
                .Formulas(0) = "Cliente='" & Trim(txtNombre.Text) & "'"
                .Formulas(1) = "Direccion='" & Trim(consulta!calle) & "'"
                .Formulas(2) = "Colonia='" & Trim(consulta!Colonia) & "'"
                .Formulas(3) = "Ciudadyestado='" & Trim(consulta!ciudad) & ", " & Trim(consulta!Estado) & "'"
                .Formulas(4) = "Telefono='" & Trim(consulta!telcasa) & "'"
                .Formulas(5) = "Nocontrato='" & Trim(consulta!numcontrato) & "'"
                .Formulas(6) = "Concepto='" & Trim(grdPagos.CellText(grdPagos.SelectedRow, 1)) & "'"
                .Formulas(7) = "Vencimiento='" & Format(grdPagos.CellText(grdPagos.SelectedRow, 7), "dd/mmm/yyyy") & "'"
                .Formulas(8) = "Folio=" & grdPagos.CellText(grdPagos.SelectedRow, 6) & ""
                .Formulas(9) = "Pagomensual=" & pago
                .Formulas(10) = "Interesmensual=" & Intereses
                .Formulas(11) = "Interesmoratorio=" & lblMoratorios.Caption & ""
                .Formulas(12) = "Gastoscobranza=" & IIf(txtCargosadicionales.Text = "", 0, txtCargosadicionales.Text) & ""
                .Formulas(13) = "Total=" & monto
                .Formulas(14) = "Cantidad=" & CDbl(txtTotal.Text) & ""
                .Formulas(15) = "Cantidadletras='(" & Trim(CantidadEnLetra(CCur(txtTotal.Text))) & ")'"
                .Formulas(16) = "Fecha='" & Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy") & "'"
                '            .Formulas(17) = "Formapago='" & Trim(consulta!manerapago) & "'"
                .WindowShowPrintSetupBtn = True
                .WindowState = crptMaximized
                .WindowTitle = "Pagaré"
                .Action = 1
            End With
    
            consulta.Close
            Limpiar
            grdPagos.Clear
            MuestraRecibos rcSolicitud
        Else
            MsgBox "Este recibo ya ha sido liquidado anteriormente !!", vbInformation, "Pagos"
        End If
    End If

End Sub

Private Sub cmdLimpiar_Click()
    Limpiar
End Sub

Private Sub cmdReimprimir_Click()

    If grdPagos.Rows > 0 And grdPagos.SelectedRow <> 0 Then
        
        Set consulta = dbDatos.Execute("select recibos.fechapago,recibos.montopagado,recibos.adicionales,recibos.moratorios,recibos.interes,recibos.monto,recibos.folio as foliorecibo,recibos.vencimiento,recibos.pagado,recibos.concepto,clientes.calle,clientes.colonia,clientes.ciudad,clientes.estado,clientes.telcasa,solicitudes.numcontrato,formapago.descripcion as manerapago from (formapago inner join (clientes inner join solicitudes on clientes.id = solicitudes.idcliente) on formapago.Id = solicitudes.formapago) inner join recibos on solicitudes.Id = recibos.idsolicitud where recibos.id=" & grdPagos.CellItemData(grdPagos.SelectedRow, 1) & "")

        If consulta!pagado Then

            With frmMDI.Cr
                .Reset
                .ReportFileName = Path & "\Reportes\Pagare.rpt"
                .DiscardSavedData = True
                .Formulas(0) = "Cliente='" & Trim(txtNombre.Text) & "'"
                .Formulas(1) = "Direccion='" & Trim(consulta!calle) & "'"
                .Formulas(2) = "Colonia='" & Trim(consulta!Colonia) & "'"
                .Formulas(3) = "Ciudadyestado='" & Trim(consulta!ciudad) & ", " & Trim(consulta!Estado) & "'"
                .Formulas(4) = "Telefono='" & Trim(consulta!telcasa) & "'"
                .Formulas(5) = "Nocontrato='" & Trim(consulta!numcontrato) & "'"
                .Formulas(6) = "Concepto='" & Trim(consulta!concepto) & "'"
                .Formulas(7) = "Vencimiento='" & Format(consulta!Vencimiento, "dd/mmm/yyyy") & "'"
                .Formulas(8) = "Folio=" & consulta!foliorecibo & ""
                .Formulas(9) = "Pagomensual=" & consulta!monto & ""
                .Formulas(10) = "Interesmensual=" & consulta!Interes & ""
                .Formulas(11) = "Interesmoratorio=" & consulta!Moratorios & ""
                .Formulas(12) = "Gastoscobranza=" & consulta!Adicionales & ""
                .Formulas(13) = "Total=" & consulta!montopagado & ""
                .Formulas(14) = "Cantidad=" & consulta!montopagado & ""
                .Formulas(15) = "Cantidadletras='(" & Trim(CantidadEnLetra(CCur(consulta!montopagado))) & ")'"
                .Formulas(16) = "Fecha='" & Format(consulta!fechapago, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy") & "'"
                .Formulas(17) = "Formapago='" & Trim(consulta!manerapago) & "'"
                .WindowShowPrintSetupBtn = True
                .WindowState = crptMaximized
                .WindowTitle = "Pagaré"
                .Action = 1
            End With

        Else
            MsgBox "Este pagaré no puede ser re-impreso porque aún no ha sido liquidado !!", vbInformation, "Pagos"
        End If
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CrearEncabezado
    Poner_Flat Fl, Me.Controls, Me
    CentrarForm Me, frmMDI
End Sub

Sub CrearEncabezado()

    With grdPagos
        .AddColumn "C1", "Concepto", ecgHdrTextALignLeft, , 135, , , , , , , CCLSortString
        .AddColumn "C2", "Saldo a pagar", ecgHdrTextALignRight, , 90, , , , , "##,###0.00", , CCLSortNumeric
        .AddColumn "C3", "Pago", ecgHdrTextALignRight, , 80, , , , , "##,###0.00", , CCLSortNumeric
        .AddColumn "C4", "Interés", ecgHdrTextALignRight, , 80, , , , , "##,###0.00", , CCLSortNumeric
        .AddColumn "C5", "Amortización", ecgHdrTextALignRight, , 80, , , , , "##,###0.00", , CCLSortNumeric
        .AddColumn "C6", "Num. Recibo", ecgHdrTextALignRight, , 80, , , , , , , CCLSortNumeric
        .AddColumn "C7", "Fec. Vencimiento", ecgHdrTextALignRight, , 95, , , , , "dd/mmm/yyyy", , CCLSortDate
        .AddColumn "C8", "Abonos", ecgHdrTextALignRight, , 90, , , , , "##,###0.00", , CCLSortNumeric
    End With

End Sub

Public Function Buscar(IDSolicitud As Long)

    Set consulta = dbDatos.Execute("select solicitudes.*,clientes.nombreempresa,clientes.telcasa,clientes.teltrabajo,(clientes.apellidop + ' ' + clientes.apellidom + ' ' + clientes.nombre) as cliente from solicitudes inner join clientes on solicitudes.idcliente=clientes.id where solicitudes.id=" & IDSolicitud & "")

    If Not consulta.BOF And Not consulta.EOF Then
        lblMonto.Caption = "$" & Format(consulta!montocredito, "##,###0.00")
        lblPago.Caption = "$" & Format(consulta!pago, "##,###0.00")
        lblPlazo.Caption = consulta!Plazo
        lblTasa.Caption = consulta!Tasa
    
        txtNombre.Text = consulta!Cliente
        txtNombre.Tag = consulta!ID
        txtEmpresa.Text = consulta!nombreempresa
        txtTelcasa.Text = consulta!telcasa
        txtTeloficina.Text = consulta!teltrabajo
    
        grdPagos.Clear False
    
        'Muestro los recibos
        MuestraRecibos IDSolicitud
    End If

End Function

Function MuestraRecibos(IDSolicitud As Long)
Dim Interes As Double, Amortizacion As Double, Saldo As Double, i As Integer

    Set consulta = dbDatos.Execute("select recibos.montopagado,recibos.id as idrecibos,recibos.monto,recibos.amortizacion,recibos.interes,recibos.saldo,recibos.concepto,recibos.folio as foliorecibo,recibos.vencimiento as vencirecibo,solicitudes.* from recibos inner join solicitudes on recibos.idsolicitud=solicitudes.id where solicitudes.id=" & IDSolicitud & " order by recibos.id")

    If Not consulta.BOF And Not consulta.EOF Then
        consulta.MoveFirst
        i = 0

        With grdPagos
            .Redraw = False
            While Not consulta.EOF
                i = i + 1

                DoEvents
            
                .AddRow
                .CellText(.Rows, 1) = consulta!concepto
                .CellItemData(.Rows, 1) = consulta!idrecibos
                .CellTextAlign(.Rows, 1) = DT_LEFT
                .CellText(.Rows, 2) = consulta!Saldo
                .CellTextAlign(.Rows, 2) = DT_RIGHT
                .CellText(.Rows, 3) = consulta!monto
                .CellItemData(.Rows, 3) = consulta!ID
                .CellTextAlign(.Rows, 3) = DT_RIGHT
                .CellText(.Rows, 4) = consulta!Interes
                .CellTextAlign(.Rows, 4) = DT_RIGHT
                .CellText(.Rows, 5) = consulta!Amortizacion
                .CellTextAlign(.Rows, 5) = DT_RIGHT
                .CellText(.Rows, 6) = IIf(consulta!foliorecibo = 0, "", consulta!foliorecibo)
                .CellTextAlign(.Rows, 6) = DT_RIGHT
                .CellText(.Rows, 7) = consulta!vencirecibo
                .CellTextAlign(.Rows, 7) = DT_CENTER
                .CellText(.Rows, 8) = consulta!montopagado
                .CellTextAlign(.Rows, 8) = DT_RIGHT
            
                Poner_Colores grdPagos, i
                consulta.MoveNext
            Wend
            .Redraw = True
        End With

    End If

    Set consulta = Nothing
End Function

Private Sub Form_Unload(Cancel As Integer)
    Quitar_Flat Fl
End Sub

Private Sub grdPagos_Click(ByVal lRow As Long, _
                           ByVal lCol As Long)
    Dim pago As Double, Moratorios As Double, Adicionales As Double, Interes As Double

    If grdPagos.SelectedRow > 0 Then
        If grdPagos.CellText(grdPagos.SelectedRow, 6) = "" Then
            lblRecibonum.Caption = Regresa_Movimiento(False, "Foliorecibo")
            lblFecha.Caption = Format(Date, "dd/mmm/yyyy")
            lblPagomensual.Caption = Format(grdPagos.CellText(grdPagos.SelectedRow, 3), "##,###0.00")
            lblMoratorios.Caption = Format(regresa_interesmoratorio(grdPagos.CellItemData(grdPagos.SelectedRow, 1)), "##,###0.00")
            lblInteres.Caption = Format(grdPagos.CellText(grdPagos.SelectedRow, 4), "##,###0.00")
        
            pago = IIf(lblPagomensual.Caption = "", 0, lblPagomensual.Caption)
            Moratorios = IIf(lblMoratorios.Caption = "", 0, lblMoratorios.Caption)
            Adicionales = IIf(txtCargosadicionales.Text = "", 0, txtCargosadicionales.Text)
            Interes = IIf(lblInteres.Caption = "", 0, lblInteres.Caption)
        
            txtTotal.Text = Format(pago + Moratorios + Adicionales + Interes, "##,###0.00")
        Else
            Limpiar
        End If
    End If

End Sub

Private Sub txtCargosadicionales_Change()
    Dim pago As Double, Moratorios As Double, Adicionales As Double, Interes As Double

    pago = IIf(lblPagomensual.Caption = "", 0, lblPagomensual.Caption)
    Moratorios = IIf(lblMoratorios.Caption = "", 0, lblMoratorios.Caption)
    Adicionales = IIf(txtCargosadicionales.Text = "", 0, txtCargosadicionales.Text)
    Interes = IIf(lblInteres.Caption = "", 0, lblInteres.Caption)

    txtTotal.Text = Format(pago + Moratorios + Adicionales + Interes, "##,###0.00")
End Sub

Private Sub txtCargosadicionales_GotFocus()
    Seleccionar_Texto txtCargosadicionales
    Cambiar_Color True, txtCargosadicionales
End Sub

Private Sub txtCargosadicionales_KeyPress(KeyAscii As Integer)
    KeyAscii = Solo_Numeros(KeyAscii, 1)
    Pasar_Foco KeyAscii
End Sub

Private Sub txtCargosadicionales_LostFocus()
    Cambiar_Color False, txtCargosadicionales
End Sub

Private Sub txtEmpresa_GotFocus()
    Seleccionar_Texto txtEmpresa
    Cambiar_Color True, txtEmpresa
End Sub

Private Sub txtEmpresa_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtEmpresa_LostFocus()
    Cambiar_Color False, txtEmpresa
End Sub

Private Sub txtNombre_GotFocus()
    Seleccionar_Texto txtNombre
    Cambiar_Color True, txtNombre
End Sub

Private Sub txtNombre_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtNombre_LostFocus()
    Cambiar_Color False, txtNombre
End Sub

Private Sub txtTelcasa_GotFocus()
    Seleccionar_Texto txtTelcasa
    Cambiar_Color True, txtTelcasa
End Sub

Private Sub txtTelcasa_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtTelcasa_LostFocus()
    Cambiar_Color False, txtTelcasa
End Sub

Private Sub txtTeloficina_GotFocus()
    Seleccionar_Texto txtTeloficina
    Cambiar_Color True, txtTeloficina
End Sub

Private Sub txtTeloficina_KeyPress(KeyAscii As Integer)
    Pasar_Foco KeyAscii
    KeyAscii = 0
End Sub

Private Sub txtTeloficina_LostFocus()
    Cambiar_Color False, txtTeloficina
End Sub

Function regresa_interesmoratorio(idrecibo As Long) As Double
    Dim dias As Integer, Vencimiento As Date, i As Integer, Tasa As Double, Intereses As Double

    Tasa = Regresa_Valor_BD("tasa") / 100

    Set consulta = dbDatos.Execute("select * from recibos where id=" & idrecibo & "")

    If Not consulta.BOF And Not consulta.EOF Then
        Vencimiento = consulta!Vencimiento
        dias = DateDiff("d", Vencimiento, Date)

        If dias > 0 Then

            For i = 1 To dias
                Intereses = Intereses + (consulta!monto * (Tasa / 30))
            Next i

        Else
            Intereses = 0
        End If
    End If

    regresa_interesmoratorio = Intereses
End Function

Sub Limpiar()
    lblRecibonum.Caption = ""
    lblFecha.Caption = ""
    lblPagomensual.Caption = ""
    lblMoratorios.Caption = ""
    txtCargosadicionales.Text = ""
    txtTotal.Text = ""
    lblInteres.Caption = ""
End Sub
