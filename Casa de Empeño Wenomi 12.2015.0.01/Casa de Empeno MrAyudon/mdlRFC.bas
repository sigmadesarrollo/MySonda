Attribute VB_Name = "mdlRFC"
Option Explicit

Private Function RFCFiltraAcentos(ByVal strTexto As String) As String

'Esta rutina elimina los acentos y convierte el nombre
'a mayusculas.

strTexto = Replace(strTexto, "á", "a")
strTexto = Replace(strTexto, "é", "e")
strTexto = Replace(strTexto, "í", "i")
strTexto = Replace(strTexto, "ó", "o")
strTexto = Replace(strTexto, "ú", "u")
RFCFiltraAcentos = UCase(strTexto)

End Function

Private Function RFCApellidoCorto(ByVal strNombre As String, _
ByVal strPaterno As String, _
ByVal strMaterno As String, _
ByVal strFecha As String) As String

'Eta rutina calcula el RFC tomando en cuenta un
'apellido paterno de tres o menos letras.

RFCApellidoCorto = Left$(strPaterno, 1) & Left$(strMaterno, 1) & _
Left$(strNombre, 2) & strFecha

End Function

Private Function RFCUnApellido(ByVal strNombre As String, _
ByVal strPaterno As String, _
ByVal strMaterno As String, _
ByVal strFecha As String) As String

'Esta rutina toma en cuenta casos cuando solo se
'da un apellido, ya sea el paterno o materno.

Dim strApellido As String

Select Case True
Case Len(strPaterno) > 0 And Len(strMaterno) = 0
'Solo hay apellido paterno.
strApellido = Left$(strPaterno, 2)
Case Len(strPaterno) = 0 And Len(strMaterno) > 0
'Solo hay apellido materno.
strApellido = Left$(strMaterno, 2)
Case Else
strApellido = Left$(strNombre, 2)
End Select

'Ahora arma el RFC.
RFCUnApellido = strApellido & Left$(strNombre, 2) & strFecha

End Function

Private Sub RFCFiltraNombres(strNombre As String, _
strPaterno As String, _
strMaterno As String)


'Esta rutina elimina palabras sobrantes para el
'calculo del RFC de los tres nombres.


Dim strArPalabras() As Variant
Dim i As Integer

'Inicializa el arreglo con las palabras que no queremos.
strArPalabras = Array(".", ",", "DE ", "DEL ", "LA ", _
"LOS ", "LAS ", "Y ", "MC ", "MAC ", "VON ", "VAN ")

'Busca cada palabra en los tres nombre y eliminala
'se se encuentra.
For i = LBound(strArPalabras) To UBound(strArPalabras)
strNombre = Replace(strNombre, strArPalabras(i), "")
strPaterno = Replace(strPaterno, strArPalabras(i), "")
strMaterno = Replace(strMaterno, strArPalabras(i), "")
Next i

'Listo, ahora sigo con el nombre pila, buscando
'la presencia de Maria o Jose.

'Inicializa el arreglo con las palabras que
'queremos eliminar.
strArPalabras = Array("JOSE ", "MARIA ", "J ", "MA ")

'Haz esto solo si el nombre de pila tiene algun
'espacio.
If InStr(strNombre, " ") > 0 Then
For i = LBound(strArPalabras) To UBound(strArPalabras)
strNombre = Replace(strNombre, strArPalabras(i), "")
Next i
End If

'Por ultimo, elimina doble consonantes de los nombres
'cuando estas ocurren en las primeras dos letras del
'nombre.
Select Case Left$(strNombre, 2)
Case "CH"
strNombre = Replace(strNombre, "CH", "C", 1, 1)
Case "LL"
strNombre = Replace(strNombre, "LL", "L", 1, 1)
End Select

Select Case Left$(strPaterno, 2)
Case "CH"
strPaterno = Replace(strPaterno, "CH", "C", 1, 1)
Case "LL"
strPaterno = Replace(strPaterno, "LL", "L", 1, 1)
End Select

Select Case Left$(strMaterno, 2)
Case "CH"
strMaterno = Replace(strMaterno, "CH", "C", 1, 1)
Case "LL"
strMaterno = Replace(strMaterno, "LL", "L", 1, 1)
End Select



End Sub

Private Function RFCQuitaProhibidas(ByVal strRFC As String) As String

'Esta rutina quita cualquiera de las palabras prohibidas,
'cambiando el ultimo caracter de dicha palabra a X.

Dim strPalabras As String

'Define todas las palabras prohibidas.
strPalabras = "BUEI*BUEY*CACA*CACO*CAGA*CAGO*CAKA*CAKO*COGE*COJA*"
strPalabras = strPalabras & "KOGE*KOJO*KAKA*KULO*MAME*MAMO*MEAR*"
strPalabras = strPalabras & "MEAS*MEON*MION*COJE*COJI*COJO*CULO*"
strPalabras = strPalabras & "FETO*GUEY*JOTO*KACA*KACO*KAGA*KAGO*"
strPalabras = strPalabras & "MOCO*MULA*PEDA*PEDO*PENE*PUTA*PUTO*"
strPalabras = strPalabras & "QULO*RATA*RUIN*"

'Si alguna de estas se encuentra, cambiala.
If InStr(strPalabras, Left$(strRFC, 4) & "*") > 0 Then
'Reemplaza el cuarto caracter del RFC para eliminar
'l apalabra prohibida.
Mid(strRFC, 4, 1) = "X"
End If

RFCQuitaProhibidas = strRFC
End Function

Private Function RFCHomoclave(ByVal strNombre As String, _
ByVal strPaterno As String, _
ByVal strMaterno As String) As String

'Esta rutina calcula la homoclave, que es de dos
'caracteres. El proceso solo toma en cuenta los
'nombres de la persona.

Dim strNombreComp As String
'Dim strChars As String
'Dim strDigitos As String
Dim strCharsHc As String
'Dim strDigitos2 As String
'Dim strSeq As String
'Dim strArSeq() As String
'Dim strArSeq1() As Variant
'Dim strArSeq2() As String
Dim strChr As String
Dim i As Integer
'Dim intIdx As Integer
Dim strCadena As String
Dim intNum1 As Integer, intNum2 As Integer
'Dim intProd3 As Integer
Dim intSum As Integer
'Dim strSum As String
Dim int3 As Integer
Dim intQuo As Integer, intRem As Integer
'Dim str2Digitos As String
'Dim strHomoclave As String


'Consigue el nombre completo de la persona.
strNombreComp = strPaterno & " " & strMaterno & " " & strNombre

'Inicializa la cadena de caracteres.
'strChars = "*0123456789&\ABCDEFGHIJKLMNOPQRSTUVWXYZ"
'Y digitos.
'strDigitos = "00000102030405060708091010111213141516171819"
'strDigitos = strDigitos & "2122232425262728293233343536373839"

'Inicializa la cadena de caracteres que contiene
'los caracteres permitidos para la homoclave.
'Notese la ausencia del numero 0 y la letra o.
strCharsHc = "123456789ABCDEFGHIJKLMNPQRSTUVWXYZ"
'Y digitos.
'strDigitos2 = "000102030405060708091011121314151617181920212223"
'strDigitos2 = strDigitos2 & "24252627282930313233"

'Genera la sequencia de digitos.
' For i = 1 To Len(strChars)
' ReDim Preserve strArSeq(1 To i) As String
' strArSeq(i) = Mid$(strDigitos, i * 2 - 1, 2)
' Next i


' For i = 1 To Len(strDigitos2) Step 2
' intIdx = intIdx + 1
' ReDim Preserve strArSeq1(1 To intIdx) As Variant
' ReDim Preserve strArSeq2(1 To intIdx) As String
' strArSeq1(intIdx) = Mid$(strDigitos2, i, 2)
' strArSeq2(intIdx) = Mid$(strChars2, (i + 1) / 2, 1)
' Next i

'Inicializa la cadena con 0 para desplazar todo a
'la derecha.
strCadena = "0"

For i = 1 To Len(strNombreComp)
strChr = Mid$(strNombreComp, i, 1)
' strChr = IIf(strChr = " ", "*", strChr)

'Convierte la letra a un numero de dos
'digitos.
Select Case strChr
Case " ", "-"
strCadena = strCadena & "00"
Case "Ñ", "Ü"
strCadena = strCadena & "10"
Case "A" To "I"
strCadena = strCadena & CStr(Asc(strChr) - 54)
Case "J" To "R"
strCadena = strCadena & CStr(Asc(strChr) - 53)
Case "S" To "Z"
strCadena = strCadena & CStr(Asc(strChr) - 51)
Case "0" To "9"
'Se supone que esta linea nunca se ejecutara, pues
'un nombre no usa digitos. Aun asi, como estaba
'en el algoritmo original, lo dejo aqui.
strCadena = strCadena & Format$(strChr, "00")
End Select

' strChr = IIf(strChr = " ", "*", strChr)
' intIdx = InStr(strChars, strChr)
' If intIdx > 0 Then
' strCadena = strCadena & strArSeq(intIdx)
' Else
' strCadena = strCadena & "00"
' End If
Next i


'MsgBox strCadena
'Clipboard.Clear
'Clipboard.SetText strCadena

'Berra toda la cadena y realiza una operacion matematica
'en cada uno de los digitos.
'
'Por cada digitos se toman dos a la vez y se multiplica
'este numero por el digito de unidades del mismo numero.
'Ejemplo:
'
' Si la cadena es 01245
'
' Se comienza con el primer digito, se toman dos y luego
' se multiplica por la unidad de ese mismo numero:
'
' Primer digito = 0, los dos: 01
' Se multiplica "01" (1) por "1"
' Se acumula.
'
' Segundo digito = 1, los dos: 12
' Se multiplica "12" (12) por "2"
'
' Tercer digito = 2, los dos: 24
' Se multiplica "24" (24) por "4"
' --etc.
For i = 1 To Len(strCadena) - 1
intNum1 = Val(Mid$(strCadena, i, 2))
intNum2 = Val(Mid$(strCadena, i + 1, 1))
'intProd3 = intProd1 * intProd2
intSum = intSum + intNum1 * intNum2
'Debug.Print i, intProd1, intProd2, intSum
Next i
'MsgBox intSum

'De la suma, solo necesito los ultimos
'tres digitos. La forma mas facil de hacer
'esto en convirtiendo el numero a cadena,
'luego tomando los tres digitos de la derecha.
'strSum = CStr(intSum)
'strSum = Right$(strSum, 3)
int3 = Val(Right$(CStr(intSum), 3))

intQuo = Int(int3 / 34)
' intRem = int3 - intQuo * 34
intRem = int3 Mod 34
'La homoclave se consigue usando el
'cociente y el residuo.

'Se usa el cociente y residio para
'buscar las letras del homoclave
'dentro de la tabla de caracteres
'permitidos.
RFCHomoclave = Mid$(strCharsHc, intQuo + 1, 1) & _
Mid$(strCharsHc, intRem + 1, 1)

'Usando el cociente, se arma una cadena
'de dos digitos.
' str2Digitos = Format$(intQuo, "00")



End Function

Private Function RFCDigitoVerificador(ByVal strRFC As String) As String

'Esta rutina calcula el digito verificador. El RFC
'consta de las iniciales, los digitos de la fecha
'de nacimiento y los dos caracteres de la homoclave.

'

' Dim strDigitos As String
Dim strChars As String
' Dim strArDigitos() As String
' Dim strArChars() As Variant
Dim i As Integer, intIdx As Integer
Dim strBuffer As String
Dim intTemp As Integer
Dim strCh As String
Dim strDV As String
' Dim intProd1 As Integer
'Dim intProd3 As Integer
Dim intSumas As Integer
' Dim intContador As Integer
' Dim intQuo As Integer
' Dim intRem As Integer
Dim intDV As Integer

' strDigitos = "00010203040506070809101112131415161718192021"
' strDigitos = strDigitos & "22232425262728293031323334353637"
strChars = "0123456789ABCDEFGHIJKLMN&OPQRSTUVWXYZ*"

'Inicializa el contador.
' intContador = 13


'El RFC tiene 12 caracteres:
' 4 Letras, 6 digitos y 2 caracteres (homoclave)
'
'Barre los 12 caracteres del RFC.

For i = 1 To Len(strRFC)
strCh = Mid$(strRFC, i, 1)
strCh = IIf(strCh = " ", "*", strCh)


intIdx = InStr(strChars, strCh) - 1
'strBuffer = strBuffer & IIf(intIdx > 0, _
Mid$(strDigitos, intIdx * 2 - 1, 2), _
"00")

'intSumas = intSumas + intIdx * intContador
intSumas = intSumas + intIdx * (14 - i)
'intContador = intContador - 1

strBuffer = strBuffer & Format$(intIdx, "00")


Next i




If intSumas Mod 11 = 0 Then
strDV = "0"
Else
intDV = 11 - intSumas Mod 11
If intDV > 9 Then
strDV = "A"
Else
strDV = CStr(intDV)
End If
End If


RFCDigitoVerificador = strDV

End Function
Private Function RFCArmalo(ByVal strNombre As String, _
ByVal strPaterno As String, _
ByVal strMaterno As String, _
ByVal strFecha As String) As String

'Esta rutina arma el RFC basandose en los tres nombres
'y la fecha de nacimiento.

'Dim strArVocales() As Variant
Dim strVocales As String
Dim strLetra As String
Dim strPrimerVocal As String
Dim i As Integer, intIdx As Integer

'Inicializa la cadena de vocales.
strVocales = "AEIOU"

' strLetra = Mid$(strPaterno, 2, 1)

'Primero consigo la primera vocal del nombre
'comenzando con la segunda letra.
For i = 2 To Len(strPaterno)
If InStr(strVocales, Mid$(strPaterno, i, 1)) > 0 Then
strPrimerVocal = Mid$(strPaterno, i, 1)
Exit For
End If
Next i

' For i = 2 To Len(strPaterno)
' intIdx = InstrAr(strArVocales, Mid$(strPaterno, i, 1))
' If intIdx > 0 Then
' strLetra = strArVocales(intIdx)
' 'i = Len(strPaterno) + 8
' Exit For
' End If
' Next i


RFCArmalo = Left$(strPaterno, 1) & strPrimerVocal & Left$(strMaterno, 1) & _
Left$(strNombre, 1) & strFecha

End Function

Public Function GeneraRFC(ByVal strNombre As String, _
ByVal strPaterno As String, _
ByVal strMaterno As String, _
ByVal dteFechaNacimiento As Date) As String

' Derechos Reservados (c), 2004 Ing. Salvador Garcia Velazquez
'
' Reglas de uso:
'
' Puedes usar este algoritmo en tu aplicacion personal,
' educacional, empresarial o comercial, siempre y cuando
' este mensaje de derechos reservados este presente. Su
' uso es libre de regalias y su autor es libre de cualquier
' fallo debido al codigo o logica.
'
' Por ningun motivo se da permiso de distribuir este codigo.
' Este codigo sigue siendo propiedad exclusiva del autor.
' Las rutinas afectadas por los derechos reservados son:
'
' GeneraRFC, RFCApellidoCorto, RFCArmalo, RFCUnApellido,
' RFCDigitoVerificador, RFCFiltraAcentos, RFCFiltraNombres
' RFCHomoclave y RFCQuitaProhibidas
'
' Cualquier rutina se puede emplear independientemente, siempre
' y cuando incluya este mensaje. Para cualquier correcion, omision
' o modificacion, favor de dirigirse a sal_garcia at bigfoot punto com.


'Esta rutina genera el RFC. Datos de entrada:

'strNombre: Tipo String Nombre de pila Dato valido requerido.
'strPaterno: Tipo String Apellido paterno Por lo menos un
'strMaterno: Tipo String Apellido materno apellido es requerido.
'dteFechaNacimiento: Tipo Date


Dim strFecha As String
Dim strRFC As String
Dim strNombreOriginal As String
Dim strPaternoOriginal As String
Dim strMaternoOriginal As String

'Consigue la fecha.
strFecha = Format(dteFechaNacimiento, "yymmdd")

'El RFC se calcula a base de letras vocales
'sin acento, elimina cualquier acento dentro
'el nombre.
strNombre = RFCFiltraAcentos(strNombre)
strPaterno = RFCFiltraAcentos(strPaterno)
strMaterno = RFCFiltraAcentos(strMaterno)

'Asegura que todos los nombres esten en letras
'mayusculas.
'strNombre = UCase(Trim$(strNombre))
'strPaterno = UCase(Trim$(strPaterno))
'strMaterno = UCase(Trim$(strMaterno))

'Procede solo si existe el nombre de pila.
If Len(strNombre) > 0 Then

'Guarda el nombre original para calcular
'la homoclave.
strNombreOriginal = strNombre
strPaternoOriginal = strPaterno
strMaternoOriginal = strMaterno

'Elimina palabras sobrantes dentro de los nombres.
RFCFiltraNombres strNombre, strPaterno, strMaterno

'Toma en cuenta el nombre de pila cuando este se
'compone de un nombre mas Jose o Maria
' RFCFiltraNombrePila strNombre

'Ahora toma en cuenta nombre que comiencen con una
'doble consonante.
' RFCFiltraDobleConsonantes strNombre, strPaterno, strMaterno

'Ahora el siguiente paso es determinar como se va a
'calcular el RFC. Existen reglas:
'
' 1. Se dan los tres nombres.
' 2. Se da solo un nombre.
' 3. El apellido dado solo tiene 3 o menos letras.
Select Case True
Case Len(strPaterno) > 0 And Len(strMaterno) > 0
'Los tres nombres existen, procede.
'Determina si el apellido paterno tiene
'menos de 3 letras.
If Len(strPaterno) < 3 Then
'Calcula el RFC tomando en cuenta un apellido corto.
strRFC = RFCApellidoCorto(strNombre, strPaterno, strMaterno, strFecha)
Else
'Calcula el RFC.
strRFC = RFCArmalo(strNombre, strPaterno, strMaterno, strFecha)
End If

Case Len(strPaterno) = 0 Or Len(strMaterno) = 0
'Uno de ellos esta vacio.
strRFC = RFCUnApellido(strNombre, strPaterno, strMaterno, strFecha)

End Select

'El RFC tentativo ya esta armado. Ahora elimina
'cualquier palabra posiblemente ofensiva.
strRFC = RFCQuitaProhibidas(strRFC)

'Ya tengo el RFC, ahora solo falta la homoclave y el
'digito verificador.
strRFC = strRFC & RFCHomoclave(strNombreOriginal, strPaternoOriginal, strMaternoOriginal)

'Por ultimo, calcula el digito verificador.
GeneraRFC = strRFC & RFCDigitoVerificador(strRFC)

End If

End Function




