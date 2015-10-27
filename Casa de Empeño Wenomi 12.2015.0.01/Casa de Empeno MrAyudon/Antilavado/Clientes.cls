VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'variables locales para almacenar los valores de las propiedades
Private mvarId As Long 'copia local
Private mvarNombre As String 'copia local
Private mvarApellidoPaterno As String 'copia local
Private mvarApellidoMaterno As String 'copia local
Private mvarDireccion As String 'copia local
Private mvarNoInterior As String 'copia local
Private mvarNoExterior As String 'copia local
Private mvarColonia As String 'copia local
Private mvarMunicipio As String 'copia local
Private mvarEstado As String 'copia local
Private mvarCodigoPostal As String 'copia local
Private mvarTelefono As String 'copia local
Private mvarEmail As String 'copia local
Private mvarFechaNacimiento As Date 'copia local
Private mvarMensaje As String 'copia local
Private mvarCurp As String 'copia local
Private mvarRFC As String 'copia local
Private mvarSexo As Integer 'copia local
Private mvarIDOcupacion As Long 'copia local
Private mvarIDTipoIdentificacion As Long 'copia local
Private mvarNumeroIdentificacion As String 'copia local
Private mvarIDEstadoNacimiento As Long 'copia local
Private mvarIDPaisNacimiento As Long 'copia local
Private mvarIDNacionalidad As Long 'copia local
Private mvarIDPaisNacionalidad As Long 'copia local
Private mvarFechaExpiracion As Date 'copia local
Private mvarFisicaMoral As Integer 'copia local
Private mvarRazonSocial As String 'copia local
Private mvarFechaAltaRazonSocial As Date 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarRL_Nombre As String 'copia local
Private mvarRL_ApellidoPaterno As String 'copia local
Private mvarRL_ApellidoMaterno As String 'copia local
Private mvarRL_RFC As String 'copia local
Private mvarDesIdentificacionOtro As String 'copia local
Private mvarIDTipoAlerta As Long 'copia local
Private mvarDescripcionTipoAlerta As String 'copia local
'variables locales para almacenar los valores de las propiedades
Private mvarRL_Curp As String 'copia local
Public Property Let RL_Curp(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RL_Curp = 5
    mvarRL_Curp = vData
End Property

Public Property Get RL_Curp() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RL_Curp
    RL_Curp = mvarRL_Curp
End Property

Public Property Let DescripcionTipoAlerta(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DescripcionTipoAlerta = 5
    mvarDescripcionTipoAlerta = vData
End Property

Public Property Get DescripcionTipoAlerta() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DescripcionTipoAlerta
    DescripcionTipoAlerta = mvarDescripcionTipoAlerta
End Property

Public Property Let IDTipoAlerta(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.IDTipoAlerta = 5
    mvarIDTipoAlerta = vData
End Property

Public Property Get IDTipoAlerta() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.IDTipoAlerta
    IDTipoAlerta = mvarIDTipoAlerta
End Property

Public Property Let DesIdentificacionOtro(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.DesIdentificacionOtro = 5
    mvarDesIdentificacionOtro = vData
End Property

Public Property Get DesIdentificacionOtro() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.DesIdentificacionOtro
    DesIdentificacionOtro = mvarDesIdentificacionOtro
End Property

Public Property Let RL_RFC(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RL_RFC = 5
    mvarRL_RFC = vData
End Property

Public Property Get RL_RFC() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RL_RFC
    RL_RFC = mvarRL_RFC
End Property

Public Property Let RL_ApellidoMaterno(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RL_ApellidoMaterno = 5
    mvarRL_ApellidoMaterno = vData
End Property

Public Property Get RL_ApellidoMaterno() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RL_ApellidoMaterno
    RL_ApellidoMaterno = mvarRL_ApellidoMaterno
End Property

Public Property Let RL_ApellidoPaterno(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RL_ApellidoPaterno = 5
    mvarRL_ApellidoPaterno = vData
End Property

Public Property Get RL_ApellidoPaterno() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RL_ApellidoPaterno
    RL_ApellidoPaterno = mvarRL_ApellidoPaterno
End Property

Public Property Let RL_Nombre(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RL_Nombre = 5
    mvarRL_Nombre = vData
End Property

Public Property Get RL_Nombre() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RL_Nombre
    RL_Nombre = mvarRL_Nombre
End Property

Public Property Let FechaAltaRazonSocial(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaAltaRazonSocial = 5
    mvarFechaAltaRazonSocial = vData
End Property

Public Property Get FechaAltaRazonSocial() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechaAltaRazonSocial
    FechaAltaRazonSocial = mvarFechaAltaRazonSocial
End Property

Public Property Let RazonSocial(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RazonSocial = 5
    mvarRazonSocial = vData
End Property

Public Property Get RazonSocial() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RazonSocial
    RazonSocial = mvarRazonSocial
End Property

Public Property Let FisicaMoral(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FisicaMoral = 5
    mvarFisicaMoral = vData
End Property

Public Property Get FisicaMoral() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FisicaMoral
    FisicaMoral = mvarFisicaMoral
End Property

Public Sub Limpiar()
    mvarId = 0
    mvarNombre = ""
    mvarApellidoPaterno = ""
    mvarApellidoMaterno = ""
    mvarDireccion = ""
    mvarNoInterior = ""
    mvarNoExterior = ""
    mvarColonia = ""
    mvarMunicipio = ""
    mvarEstado = ""
    mvarCodigoPostal = ""
    mvarTelefono = ""
    mvarEmail = ""
    mvarFechaNacimiento = "1900/01/01"
    mvarMensaje = ""
    mvarCurp = ""
    mvarRFC = ""
    mvarSexo = 0
    mvarIDOcupacion = 0
    mvarIDTipoIdentificacion = 0
    mvarNumeroIdentificacion = ""
    mvarIDEstadoNacimiento = 0
    mvarIDPaisNacimiento = 0
    mvarIDNacionalidad = 0
    mvarIDPaisNacionalidad = 0
    mvarFechaExpiracion = "1900/01/01"
    mvarFisicaMoral = 0
    mvarRazonSocial = ""
    mvarFechaAltaRazonSocial = "1900/01/01"
    mvarRL_Nombre = ""
    mvarRL_ApellidoPaterno = ""
    mvarRL_ApellidoMaterno = ""
    mvarRL_RFC = ""
    mvarDesIdentificacionOtro = ""
    mvarIDTipoAlerta = 0
    mvarDescripcionTipoAlerta = ""
    mvarRL_Curp = ""
End Sub

Public Function Grabar() As Boolean
    Dim SqlQry As String
    Grabar = False
    If Me.Valida = True Then
        If Me.FisicaMoral = 1 Then
            Me.RazonSocial = Me.Nombre & " " & Me.ApellidoPaterno & " " & Me.ApellidoMaterno
        Else
            Me.Nombre = Me.RazonSocial
            Me.ApellidoPaterno = ""
            Me.ApellidoMaterno = ""
        End If
        If Me.ID = 0 Then
            
            SqlQry = "INSERT INTO clientes (Iniciales,Nombre,Apellido,ApellidoPaterno,ApellidoMaterno,Direccion,NoExterior,NoInterior,Colonia,Municipio,Estado,Tel,Email,Notas," & _
                     "CP,Rfc,Curp,FecNac,Sexo,FecRegistro,IDUsuario,NumeroIdentificacion,IdOcupacion,IdEstadoNac,IdPaisNacimiento,IdTipoIdent,FechaExpIdent,IdPaisNacionalidad,RazonSocial," & _
                     "PersonaFisica,FechaAltaRazonSocial,RL_Nombre,RL_ApellidoPaterno,RL_ApellidoMaterno,RL_Rfc,RL_Curp,DescIdentificacionOtro,IdTipoAlerta,DescTipoAlerta,Identificacion) VALUES ('" & _
                     Iniciales & "','" & Trim(Me.Nombre) & "','" & Trim(Me.ApellidoPaterno & " " & Me.ApellidoMaterno) & "','" & Trim(Me.ApellidoPaterno) & "','" & _
                     Trim(Me.ApellidoMaterno) & "','" & Trim(Me.Direccion) & "','" & Trim(Me.NoExterior) & "','" & Trim(Me.NoInterior) & "','" & Trim(Me.Colonia) & "','" & _
                     Trim(Me.Municipio) & "','" & Trim(Me.Estado) & "','" & Trim(Me.Telefono) & "','" & Trim(Me.Email) & "','" & Trim(Me.Mensaje) & "','" & _
                     Trim(Me.CodigoPostal) & "','" & Trim(Me.RFC) & "','" & Trim(Me.Curp) & "','" & Format(Me.FechaNacimiento, "YYYY/MM/DD") & "'," & Me.Sexo & ",'"
            SqlQry = SqlQry & Format(Date, "YYYY/MM/DD") & "'," & Trim(frmMDI.IDUsuario) & ",'" & Me.NumeroIdentificacion & "'," & Me.IDOcupacion & "," & Me.IDEstadoNacimiento & "," & _
                     Me.IDPaisNacimiento & "," & Me.IDTipoIdentificacion & ",'" & Format(Me.FechaExpiracion, "YYYY/MM/DD") & "'," & Me.IDPaisNacionalidad & ",'" & _
                     Me.RazonSocial & "'," & Me.FisicaMoral & ",'" & Format(Me.FechaAltaRazonSocial, "YYYY/MM/DD") & "','" & Me.RL_Nombre & "','" & Me.RL_ApellidoPaterno & _
                     "','" & Me.RL_ApellidoMaterno & "','" & Me.RL_RFC & "','" & Me.RL_Curp & "','" & Me.DesIdentificacionOtro & "'," & Me.IDTipoIdentificacion & ",'" & Me.DescripcionTipoAlerta & "','" & Trim(UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE Id=" & Me.IDTipoIdentificacion))) & "')"
            
            dbDatos.Execute SqlQry
            
            'dbDatos.Execute "INSERT INTO clientes (Iniciales,Nombre,Apellido,ApellidoPaterno,ApellidoMaterno,Direccion,NoExterior,NoInterior,Colonia,Municipio,Estado,Tel,Email,Notas," & _
            '        "CP,Rfc,Curp,FecNac,Sexo,FecRegistro,IDUsuario,NumeroIdentificacion,IdOcupacion,IdEstadoNac,IdPaisNacimiento,IdTipoIdent,FechaExpIdent,IdPaisNacionalidad,RazonSocial," & _
            '        "PersonaFisica,FechaAltaRazonSocial,RL_Nombre,RL_ApellidoPaterno,RL_ApellidoMaterno,RL_Rfc,RL_Curp,DescIdentificacionOtro,IdTipoAlerta,DescTipoAlerta,Identificacion) VALUES ('" & _
            '        Iniciales & "','" & Trim(Me.Nombre) & "','" & Trim(Me.ApellidoPaterno & " " & Me.ApellidoMaterno) & "','" & Trim(Me.ApellidoPaterno) & "','" & _
            '        Trim(Me.ApellidoMaterno) & "','" & Trim(Me.Direccion) & "','" & Trim(Me.NoExterior) & "','" & Trim(Me.NoInterior) & "','" & Trim(Me.Colonia) & "','" & _
            '        Trim(Me.Municipio) & "','" & Trim(Me.Estado) & "','" & Trim(Me.Telefono) & "','" & Trim(Me.Email) & "','" & Trim(Me.Mensaje) & "','" & _
            '        Trim(Me.CodigoPostal) & "','" & Trim(Me.RFC) & "','" & Trim(Me.Curp) & "','" & Format(Me.FechaNacimiento, "YYYY/MM/DD") & "'," & Me.Sexo & ",'" & _
            '        Format(Date, "YYYY/MM/DD") & "'," & Trim(frmMDI.IDUsuario) & ",'" & Me.NumeroIdentificacion & "'," & Me.IDOcupacion & "," & Me.IDEstadoNacimiento & "," & _
            '        Me.IDPaisNacimiento & "," & Me.IDTipoIdentificacion & ",'" & Format(Me.FechaExpiracion, "YYYY/MM/DD") & "'," & Me.IDPaisNacionalidad & ",'" & _
            '        Me.RazonSocial & "'," & Me.FisicaMoral & ",'" & Format(Me.FechaAltaRazonSocial, "YYYY/MM/DD") & "','" & Me.RL_Nombre & "','" & Me.RL_ApellidoPaterno & _
            '        "','" & Me.RL_ApellidoMaterno & "','" & Me.RL_RFC & "','" & Me.RL_Curp & "','" & Me.DesIdentificacionOtro & "'," & Me.IDTipoIdentificacion & ",'" & Me.DescripcionTipoAlerta & "'," & Trim(UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE Id=" & Me.IDTipoIdentificacion))) & ")"
            Me.ID = UltimoID
        Else
        
            SqlQry = "UPDATE clientes  SET Iniciales='" & Iniciales & "',Nombre='" & Trim(Me.Nombre) & "',Apellido='" & Trim(Me.ApellidoPaterno & " " & Me.ApellidoMaterno) & _
                    "',ApellidoPaterno='" & Trim(Me.ApellidoPaterno) & "',ApellidoMaterno='" & Trim(Me.ApellidoMaterno) & "',Direccion='" & Trim(Me.Direccion) & _
                    "',NoExterior='" & Trim(Me.NoExterior) & "',NoInterior='" & Trim(Me.NoInterior) & "',Colonia='" & Trim(Me.Colonia) & "',Municipio='" & Trim(Me.Municipio) & _
                    "',Estado='" & Trim(Me.Estado) & "',Tel='" & Trim(Me.Telefono) & "',Email='" & Trim(Me.Email) & "',Notas='" & Trim(Me.Mensaje) & "',CP='" & Trim(Me.CodigoPostal) & _
                    "',Rfc='" & Trim(Me.RFC) & "',Curp='" & Trim(Me.Curp) & "',FecNac='" & Format(Me.FechaNacimiento, "YYYY/MM/DD") & "',Sexo=" & Me.Sexo & _
                    ",FecRegistro='" & Format(Date, "YYYY/MM/DD") & "',IDUsuario=" & Trim(frmMDI.IDUsuario) & ",NumeroIdentificacion='" & Me.NumeroIdentificacion
            SqlQry = SqlQry & "',IdOcupacion=" & Me.IDOcupacion & ",IdEstadoNac=" & Me.IDEstadoNacimiento & ",IdPaisNacimiento=" & Me.IDPaisNacimiento & ",IdTipoIdent=" & Me.IDTipoIdentificacion & _
                    ",FechaExpIdent='" & Format(Me.FechaExpiracion, "YYYY/MM/DD") & "',IdPaisNacionalidad=" & Me.IDPaisNacionalidad & ",RazonSocial='" & Me.RazonSocial & _
                    "',PersonaFisica=" & Me.FisicaMoral & ",FechaAltaRazonSocial='" & Format(Me.FechaAltaRazonSocial, "YYYY/MM/DD") & "',RL_Nombre='" & Trim(Me.RL_Nombre) & _
                    "',RL_ApellidoPaterno='" & Trim(Me.RL_ApellidoPaterno) & "',RL_ApellidoMaterno='" & Trim(Me.RL_ApellidoMaterno) & "',RL_Rfc='" & Trim(Me.RL_RFC) & "',RL_Curp='" & Trim(Me.RL_Curp) & _
                    "',DescIdentificacionOtro='" & Trim(Me.DesIdentificacionOtro) & "',IdTipoAlerta=" & Me.IDTipoAlerta & ",DescTipoAlerta='" & Trim(Me.DescripcionTipoAlerta) & "',Identificacion='" & Trim(UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE Id=" & Me.IDTipoIdentificacion))) & "' WHERE ID = " & Me.ID
            
            dbDatos.Execute SqlQry
            'dbDatos.Execute "UPDATE clientes  SET Iniciales='" & Iniciales & "',Nombre='" & Trim(Me.Nombre) & "',Apellido='" & Trim(Me.ApellidoPaterno & " " & Me.ApellidoMaterno) & _
            '        "',ApellidoPaterno='" & Trim(Me.ApellidoPaterno) & "',ApellidoMaterno='" & Trim(Me.ApellidoMaterno) & "',Direccion='" & Trim(Me.Direccion) & _
            '        "',NoExterior='" & Trim(Me.NoExterior) & "',NoInterior='" & Trim(Me.NoInterior) & "',Colonia='" & Trim(Me.Colonia) & "',Municipio='" & Trim(Me.Municipio) & _
            '        "',Estado='" & Trim(Me.Estado) & "',Tel='" & Trim(Me.Telefono) & "',Email='" & Trim(Me.Email) & "',Notas='" & Trim(Me.Mensaje) & "',CP='" & Trim(Me.CodigoPostal) & _
            '        "',Rfc='" & Trim(Me.RFC) & "',Curp='" & Trim(Me.Curp) & "',FecNac='" & Format(Me.FechaNacimiento, "YYYY/MM/DD") & "',Sexo=" & Me.Sexo & _
            '        ",FecRegistro='" & Format(Date, "YYYY/MM/DD") & "',IDUsuario=" & Trim(frmMDI.IDUsuario) & ",NumeroIdentificacion='" & Me.NumeroIdentificacion & _
            '        "',IdOcupacion=" & Me.IDOcupacion & ",IdEstadoNac=" & Me.IDEstadoNacimiento & ",IdPaisNacimiento=" & Me.IDPaisNacimiento & ",IdTipoIdent=" & Me.IDTipoIdentificacion & _
            '        ",FechaExpIdent='" & Format(Me.FechaExpiracion, "YYYY/MM/DD") & "',IdPaisNacionalidad=" & Me.IDPaisNacionalidad & ",RazonSocial='" & Me.RazonSocial & _
            '        "',PersonaFisica=" & Me.FisicaMoral & ",FechaAltaRazonSocial='" & Format(Me.FechaAltaRazonSocial, "YYYY/MM/DD") & "',RL_Nombre='" & Trim(Me.RL_Nombre) & _
            '        "',RL_ApellidoPaterno='" & Trim(Me.RL_ApellidoPaterno) & "',RL_ApellidoMaterno='" & Trim(Me.RL_ApellidoMaterno) & "',RL_Rfc='" & Trim(Me.RL_RFC) & "',RL_Curp='" & Trim(Me.RL_Curp) & _
            '        "',DescIdentificacionOtro='" & Trim(Me.DesIdentificacionOtro) & "',IdTipoAlerta=" & Me.IDTipoAlerta & ",DescTipoAlerta='" & Trim(Me.DescripcionTipoAlerta) & "',Identificacion='" & Trim(UCase(SacaValor("mld_tipo_identificaciones", "Descripcion", " WHERE Id=" & Me.IDTipoIdentificacion))) & "' WHERE ID = " & Me.ID
        End If
        
        Grabar = True
    End If
End Function

Private Function Iniciales() As String
    Dim Cadena As String, Nombre As String, Apellidos As String, Apellidos2 As String
   
    Nombre = Trim(Me.Nombre)
    Apellidos = Trim(Me.ApellidoPaterno)
    Apellidos2 = Trim(Me.ApellidoMaterno)
    
    Cadena = Mid(Nombre, 1, 1)
    If InStr(1, Nombre, " ") <> 0 Then Cadena = Cadena & Mid(Nombre, InStr(1, Nombre, " ") + 1, 1)
    
    Cadena = Cadena & Mid(Apellidos, 1, 1)
    If Apellidos2 <> "" Then
        Cadena = Cadena & Mid(Apellidos2, 1, 1)
    Else
        If InStr(1, Apellidos, " ") <> 0 Then Cadena = Cadena & Mid(Apellidos, InStr(1, Apellidos, " ") + 1, 1)
    End If
      
   Iniciales = Cadena
End Function

Public Function Buscar(ID As Long) As Boolean
Dim rcClientes As New ADODB.Recordset
   
On Error GoTo Error

    rcClientes.Open "SELECT * FROM clientes WHERE ID=" & ID, dbDatos, adOpenForwardOnly, adLockReadOnly
    Me.Limpiar
    With rcClientes
        Me.ID = ID
        Me.Nombre = !Nombre
        Me.ApellidoPaterno = IIf(!ApellidoPaterno = "" And !ApellidoMaterno = "", !Apellido, !ApellidoPaterno)
        Me.ApellidoMaterno = !ApellidoMaterno
        Me.Direccion = IIf(IsNull(!Direccion), "", !Direccion)
        Me.NoExterior = IIf(IsNull(!NoExterior), "", !NoExterior)
        Me.NoInterior = IIf(IsNull(!NoInterior), "", !NoInterior)
        Me.Colonia = IIf(IsNull(!Colonia), "", !Colonia)
        Me.Municipio = IIf(IsNull(!Municipio), "", !Municipio)
        Me.Estado = IIf(IsNull(!Estado), "", !Estado)
        Me.Telefono = IIf(IsNull(!Tel), "", !Tel)
        Me.CodigoPostal = IIf(IsNull(!CP), "", !CP)
        Me.Email = IIf(IsNull(!Email), "", !Email)
        Me.Mensaje = IIf(IsNull(!Notas), "", !Notas)
        Me.Curp = IIf(IsNull(!Curp), "", !Curp)
        Me.RFC = IIf(IsNull(!RFC), "", !RFC)
        Me.Sexo = IIf(IsNull(!Sexo), -1, !Sexo)
        Me.IDOcupacion = IIf(IsNull(!IDOcupacion), -1, !IDOcupacion)
        Me.IDTipoIdentificacion = IIf(IsNull(!IDTipoIdent), -1, !IDTipoIdent)
        Me.NumeroIdentificacion = IIf(IsNull(!NumeroIdentificacion), "", !NumeroIdentificacion)
        Me.FechaNacimiento = IIf(IsNull(!FecNac), "1900/01/01", !FecNac)
        Me.FechaExpiracion = IIf(IsNull(!FechaExpIdent), "1900/01/01", !FechaExpIdent)
        Me.IDNacionalidad = 0
        Me.IDEstadoNacimiento = IIf(IsNull(!IdEstadoNac), -1, !IdEstadoNac)
        Me.IDPaisNacionalidad = IIf(IsNull(!IDPaisNacionalidad), -1, IIf(!IDPaisNacionalidad = 0, -1, !IDPaisNacionalidad))
        Me.IDPaisNacimiento = IIf(IsNull(!IDPaisNacimiento), -1, IIf(!IDPaisNacimiento = 0, -1, !IDPaisNacionalidad))
        Me.FisicaMoral = IIf(IsNull(!personaFisica), 1, !personaFisica)
        Me.RazonSocial = IIf(IsNull(!RazonSocial), "", !RazonSocial)
        Me.FechaAltaRazonSocial = IIf(IsNull(!FechaAltaRazonSocial), "1900/01/01", !FechaAltaRazonSocial)
        Me.RL_Nombre = IIf(IsNull(!RL_Nombre), "", !RL_Nombre)
        Me.RL_ApellidoPaterno = IIf(IsNull(!RL_ApellidoPaterno), "", !RL_ApellidoPaterno)
        Me.RL_ApellidoMaterno = IIf(IsNull(!RL_ApellidoMaterno), "", !RL_ApellidoMaterno)
        Me.RL_RFC = IIf(IsNull(!RL_RFC), "", !RL_RFC)
        Me.RL_Curp = IIf(IsNull(!RL_RFC), "", !RL_Curp)
        Me.DesIdentificacionOtro = IIf(IsNull(!DescIdentificacionOtro), "", !DescIdentificacionOtro)
        Me.IDTipoAlerta = IIf(IsNull(!IDTipoAlerta), -1, !IDTipoAlerta)
        Me.DescripcionTipoAlerta = IIf(IsNull(!DescTipoAlerta), "", !DescTipoAlerta)
        
    End With
    rcClientes.Close
    Buscar = True
    Exit Function
Error:
    Maneja_Error Err
    Buscar = False
End Function

Public Function UltimoID() As Long
Dim rcClientes As New ADODB.Recordset
   
On Error GoTo Error

    rcClientes.Open "SELECT max(ID) AS id from clientes", dbDatos, adOpenForwardOnly, adLockReadOnly
    UltimoID = Val(rcClientes!ID)
    rcClientes.Close
    Exit Function
Error:
    Maneja_Error Err
    UltimoID = 0
End Function

Public Function Valida() As Boolean
    
    If Me.FisicaMoral < 1 Then
        Valida = False
        Exit Function
    End If
    
    If Me.FisicaMoral = 1 Then
        If Trim(Me.Nombre) = "" Then
            Valida = False
            Exit Function
        End If
          
        If Trim(Me.ApellidoPaterno) = "" Then
            Valida = False
            Exit Function
        End If
        
        If Trim(Me.ApellidoMaterno) = "" Then
            Valida = False
            Exit Function
        End If
        
        If Trim(Me.Curp) = "" Then
            Valida = False
            Exit Function
        End If
        
        If Me.Sexo < 1 Then
            Valida = False
            Exit Function
        End If
        
        If Me.IDOcupacion < 1 Then
            Valida = False
            Exit Function
        End If
    
        If Me.IDPaisNacionalidad < 1 Then
            Valida = False
            Exit Function
        End If
        
        If Me.IDEstadoNacimiento < 1 Then
            Valida = False
            Exit Function
        End If
        
        If Me.IDPaisNacimiento < 1 Then
            Valida = False
            Exit Function
        End If
    Else
        If Trim(Me.RazonSocial) = "" Then
            Valida = False
            Exit Function
        End If
        
        If Trim(Me.RL_Nombre) = "" Then
            Valida = False
            Exit Function
        End If
          
        If Trim(Me.RL_ApellidoPaterno) = "" Then
            Valida = False
            Exit Function
        End If
        
        If Trim(Me.RL_ApellidoMaterno) = "" Then
            Valida = False
            Exit Function
        End If

        If Trim(Me.RL_RFC) = "" Then
            Valida = False
            Exit Function
        End If
        
        If Trim(Me.RL_Curp) = "" Then
            Valida = False
            Exit Function
        End If
    End If
    
    If Trim(Me.Direccion) = "" Then
        Valida = False
        Exit Function
    End If
    
    If Trim(Me.NoExterior) = "" Then
        Valida = False
        Exit Function
    End If
    
    If Trim(Me.Colonia) = "" Then
        Valida = False
        Exit Function
    End If
    
    If Trim(Me.Estado) = "" Then
        Valida = False
        Exit Function
    End If
    
    If Trim(Me.Municipio) = "" Then
        Valida = False
        Exit Function
    End If
    
    If Trim(Me.CodigoPostal) = "" Then
        Valida = False
        Exit Function
    End If
    
    If Trim(Me.RFC) = "" Then
        Valida = False
        Exit Function
    End If

    If Me.IDTipoIdentificacion < 1 Then
        Valida = False
        Exit Function
    Else
        If Me.IDTipoIdentificacion >= 11 And Me.IDTipoIdentificacion <= 13 Then
            If Trim(Me.DesIdentificacionOtro) = "" Then
                Valida = False
                Exit Function
            End If
        End If
    End If
    
    If Trim(Me.NumeroIdentificacion) = "" Then
        Valida = False
        Exit Function
    End If

    
    Valida = True
End Function

Public Property Let IDNacionalidad(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.IDNacionalidad = 5
    mvarIDNacionalidad = vData
End Property

Public Property Get IDNacionalidad() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.IDNacionalidad
    IDNacionalidad = mvarIDNacionalidad
End Property

Public Property Let IDEstadoNacimiento(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.IDEstadoNacimiento = 5
    mvarIDEstadoNacimiento = vData
End Property

Public Property Get IDEstadoNacimiento() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.IDEstadoNacimiento
    IDEstadoNacimiento = mvarIDEstadoNacimiento
End Property

Public Property Let IDPaisNacimiento(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.IDEstadoNacimiento = 5
    mvarIDPaisNacimiento = vData
End Property

Public Property Get IDPaisNacimiento() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.IDEstadoNacimiento
    IDPaisNacimiento = mvarIDPaisNacimiento
End Property

Public Property Let IDPaisNacionalidad(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.IDEstadoNacimiento = 5
    mvarIDPaisNacionalidad = vData
End Property

Public Property Get IDPaisNacionalidad() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.IDEstadoNacimiento
    IDPaisNacionalidad = mvarIDPaisNacionalidad
End Property

Public Property Let NumeroIdentificacion(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NumeroIdentificacion = 5
    mvarNumeroIdentificacion = vData
End Property

Public Property Get NumeroIdentificacion() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NumeroIdentificacion
    NumeroIdentificacion = mvarNumeroIdentificacion
End Property

Public Property Let IDTipoIdentificacion(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.IDTipoIdentificacion = 5
    mvarIDTipoIdentificacion = vData
End Property

Public Property Get IDTipoIdentificacion() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.IDTipoIdentificacion
    IDTipoIdentificacion = mvarIDTipoIdentificacion
End Property

Public Property Let IDOcupacion(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.IDOcupacion = 5
    mvarIDOcupacion = vData
End Property

Public Property Get IDOcupacion() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.IDOcupacion
    IDOcupacion = mvarIDOcupacion
End Property

Public Property Let Sexo(ByVal vData As Integer)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Sexo = 5
    mvarSexo = vData
End Property

Public Property Get Sexo() As Integer
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Sexo
    Sexo = mvarSexo
End Property

Public Property Let RFC(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.RFC = 5
    mvarRFC = vData
End Property

Public Property Get RFC() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.RFC
    RFC = mvarRFC
End Property

Public Property Let Curp(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Curp = 5
    mvarCurp = vData
End Property

Public Property Get Curp() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Curp
    Curp = mvarCurp
End Property

Public Property Let Mensaje(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Mensaje = 5
    mvarMensaje = vData
End Property

Public Property Get Mensaje() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Mensaje
    Mensaje = mvarMensaje
End Property

Public Property Let FechaNacimiento(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaNacimiento = 5
    mvarFechaNacimiento = vData
End Property

Public Property Get FechaNacimiento() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechaNacimiento
    FechaNacimiento = mvarFechaNacimiento
End Property

Public Property Let FechaExpiracion(ByVal vData As Date)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.FechaExpiracion = 5
    mvarFechaExpiracion = vData
End Property

Public Property Get FechaExpiracion() As Date
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.FechaExpiracion
    FechaExpiracion = mvarFechaExpiracion
End Property

Public Property Let Email(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Email = 5
    mvarEmail = vData
End Property

Public Property Get Email() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Email
    Email = mvarEmail
End Property

Public Property Let Telefono(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Telefono = 5
    mvarTelefono = vData
End Property

Public Property Get Telefono() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Telefono
    Telefono = mvarTelefono
End Property

Public Property Let CodigoPostal(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.CodigoPostal = 5
    mvarCodigoPostal = vData
End Property

Public Property Get CodigoPostal() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.CodigoPostal
    CodigoPostal = mvarCodigoPostal
End Property

Public Property Let Estado(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Estado = 5
    mvarEstado = vData
End Property

Public Property Get Estado() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Estado
    Estado = mvarEstado
End Property

Public Property Let Municipio(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Municipio = 5
    mvarMunicipio = vData
End Property

Public Property Get Municipio() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Municipio
    Municipio = mvarMunicipio
End Property

Public Property Let Colonia(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Colonia = 5
    mvarColonia = vData
End Property

Public Property Get Colonia() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Colonia
    Colonia = mvarColonia
End Property

Public Property Let NoExterior(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NoExterior = 5
    mvarNoExterior = vData
End Property

Public Property Get NoExterior() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NoExterior
    NoExterior = mvarNoExterior
End Property

Public Property Let NoInterior(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.NoInterior = 5
    mvarNoInterior = vData
End Property

Public Property Get NoInterior() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.NoInterior
    NoInterior = mvarNoInterior
End Property

Public Property Let Direccion(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Direccion = 5
    mvarDireccion = vData
End Property

Public Property Get Direccion() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Direccion
    Direccion = mvarDireccion
End Property

Public Property Let ApellidoMaterno(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ApellidoMaterno = 5
    mvarApellidoMaterno = vData
End Property

Public Property Get ApellidoMaterno() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ApellidoMaterno
    ApellidoMaterno = mvarApellidoMaterno
End Property

Public Property Let ApellidoPaterno(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.ApellidoPaterno = 5
    mvarApellidoPaterno = vData
End Property

Public Property Get ApellidoPaterno() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.ApellidoPaterno
    ApellidoPaterno = mvarApellidoPaterno
End Property

Public Property Let Nombre(ByVal vData As String)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Nombre = 5
    mvarNombre = vData
End Property

Public Property Get Nombre() As String
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Nombre
    Nombre = mvarNombre
End Property

Public Property Let ID(ByVal vData As Long)
'se usa al asignar un valor a la propiedad, en la parte izquierda de una asignación.
'Syntax: X.Id = 5
    mvarId = vData
End Property

Public Property Get ID() As Long
'se usa al recuperar un valor de una propiedad, en la parte derecha de una asignación.
'Syntax: Debug.Print X.Id
    ID = mvarId
End Property

