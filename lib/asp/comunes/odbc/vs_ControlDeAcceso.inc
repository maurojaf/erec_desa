<SCRIPT LANGUAGE="VBScript" RUNAT="Server">


' Función:    ChequeaIngresoPorRUT
' Descripción:  Chequea que el RUT y la password de ingreso al sistema estén 
'       correctas, retorna un código indicando el resultado de la validación.
' Argumentos: RUT   : RUT del cliente que está conectándose
'	      Clave : Clave de acceso al sistema del usuario
' Retorna:    1. No pudo validad la clave o el RUT no existe
'	      2. La clave expiró
'	      3. OK
'
Function ChequeaIngresoPorRUT(ByVal RUT, ByVal Clave)

  ' Encripta la clave de acceso para compararla con la de la Base de Datos
  clave_encriptada = Clave
  Usuario = RUT

  ' Se conecta a la Base de Datos
  Set objConn = Server.CreateObject("ADODB.Connection")
  On Error Resume Next
  objConn.Open Session("CadenaDeConexion")
  ' Si hay un error lo muestra y termina la ejecución del programa
  If Err.Number <> 0 Then
    Call vs_MostrarError("No se pudo conectar a la Base de Datos", true )
  End If

  ' Ejecuta la consulta que busca el usuario y la clave en la base de datos
  strSqlQuery = " SELECT * FROM usuario " & _
          " WHERE Login = '" & Usuario & _
          "' AND Password='" & clave_encriptada & "'"
  Response.Write "<PRE>strSqlQuery='" & strSqlQuery & "'"
  Set objRS = objConn.Execute(strSqlQuery)
 
  If Err Then
    If Not IsEmpty(objRS) Then
        If Not objRS Is Nothing Then
          objRS.Close
          Set objRS = Nothing
        End If
    End If
    If Not objConn Is Nothing Then
      objConn.Close
      Set obj.Conn = Nothing
    End If
    Call vs_MostrarError ("Error al obtener los datos de la Base de Datos. SQL=" & strSqlQuery, true)
  End If

  'Si no aparece el usuario, hubo un error.
  If objRS.EOF Then
    'No estaba el usuario registrado.
    ChequeaIngresoPorRUT = 1
  Else
    Session("AplicNombreDeUsuario") = TRIM(objRS("CodUsuario"))
    dif_CamClave = TRIM(objRS("dif_CamClave"))
    If CInt( dif_CamClave ) > Session("DiasDuracionPassword") Then
      ' La clave expiró
      ChequeaIngresoPorRUT = 2
    Else
      ' OK
      ChequeaIngresoPorRUT = 3
    End If
  End IF
  ' Cierra la Conexión con la Base de Datos
  objRS.Close()
  objConn.Close
  Set objConn = Nothing

End Function

' Función:    ChequeaIngreso
' Descripción:  Chequea que el usuario y la password de ingreso al sistema estén 
'       correctas, retorna un código indicando el resultado de la validación.
' Argumentos: Usuario   Código del usuario que está conectándose
'       Clave   Clave de acceso al sistema del usuario
' Retorna:    1. No pudo validad la clave o el usuario no existe
'         2. La clave expiró
'         3. OK
'
Function ChequeaIngreso(ByVal Usuario, ByVal Clave)

  ' Encripta la clave de acceso para compararla con la de la Base de Datos
  Set ObjetoEncrypt = Server.CreateObject("ControlDeAcceso.Encrypt")
 'Response.Write "<PRE>Clave='" & Clave & "'"
  clave_encriptada = ObjetoEncrypt.encrypt(Clave, "opensoft")

  ' Se conecta a la Base de Datos
  Set objConn = Server.CreateObject("ADODB.Connection")
  On Error Resume Next
  objConn.Open Session("CadenaDeConexion")
  ' Si hay un error lo muestra y termina la ejecución del programa
  If Err.Number <> 0 Then
    Call vs_MostrarError("No se pudo conectar a la Base de Datos", true )
  End If

  ' Ejecuta la consulta que busca el usuario y la clave en la base de datos
  strSqlQuery = " SELECT CodUsuario, datediff(dd, CamClave, getdate()) dif_CamClave" &_
          " FROM usuario"&_
          " WHERE CodUsuario = '" & Usuario & "'" &_
          " AND clave='" & clave_encriptada & "'"
 'Response.Write "<PRE>strSqlQuery='" & strSqlQuery & "'"
  Set objRS = objConn.Execute(strSqlQuery)
 
  If Err Then
    If Not IsEmpty(objRS) Then
        If Not objRS Is Nothing Then
          objRS.Close
          Set objRS = Nothing
        End If
    End If
    If Not objConn Is Nothing Then
      objConn.Close
      Set obj.Conn = Nothing
    End If
    Call vs_MostrarError ("Error al obtener los datos de la Base de Datos. SQL=" & strSqlQuery, true)
  End If

  'Si no aparece el usuario, hubo un error.
  If objRS.EOF Then
    'No estaba el usuario registrado.
    ChequeaIngreso = 1
  Else
    Session("AplicNombreDeUsuario") = TRIM(objRS("CodUsuario"))
    dif_CamClave = TRIM(objRS("dif_CamClave"))
    'Response.Write "<pre>dif_CamClave=" & dif_CamClave & "</pre>"
    'Response.End

    If CInt( dif_CamClave ) > Session("DiasDuracionPassword") Then
      ' La clave expiró
      ChequeaIngreso = 2
    Else
      ' OK
      ChequeaIngreso = 3
    End If
  End IF

  ' Cierra la Conexión con la Base de Datos
  objRS.Close()
  objConn.Close
  Set objConn = Nothing

End Function

' Función:    CambiaClave
' Descripción:  Cambia la clave de un usuario. Retorna un código indicando el resultado
'       del cambio.
' Argumentos: Usuario     Código del usuario que está conectándose
'       ClaveAntigua  Clave de acceso que está actualmente en la BD
'       ClaveNueva    Clave de acceso nueva
' Retorna:    1. No pudo validad la clave antigua
'         2. OK
'
Function CambiaClave(ByVal Usuario, ByVal ClaveAntigua, ByVal ClaveNueva)

  ' Encripta las claves de acceso
  Set ObjetoEncrypt = Server.CreateObject("ControlDeAcceso.Encrypt")
  clave_antigua = ObjetoEncrypt.encrypt(ClaveAntigua, "opensoft")
  clave_nueva = ObjetoEncrypt.encrypt(ClaveNueva, "opensoft")

  ' Se conecta a la Base de Datos
  Set objConn = Server.CreateObject("ADODB.Connection")
  On Error Resume Next
  objConn.Open Session("CadenaDeConexion")
  ' Si hay un error lo muestra y termina la ejecución del programa
  If Err.Number <> 0 Then
    Call vs_MostrarError("No se pudo conectar a la Base de Datos", true )
  End If

  ' Ejecuta la consulta que busca el usuario y la clave en la base de datos
  strSqlQuery = " SELECT CodUsuario, datediff(dd, CamClave, getdate()) dif_CamClave" &_
          " FROM usuario"&_
          " WHERE CodUsuario = '" & Usuario & "'" &_
          " AND clave='" & clave_antigua & "'"
  Set objRS = objConn.Execute(strSqlQuery)
  If Err Then
    If Not IsEmpty(objRS) Then
        If Not objRS Is Nothing Then
          objRS.Close
          Set objRS = Nothing
        End If
    End If
    If Not objConn Is Nothing Then
      objConn.Close
      Set obj.Conn = Nothing
    End If
    Call vs_MostrarError ("Error al obtener los datos de la Base de Datos. SQL=" & strSqlQuery, true)
  End If

  'Si no aparece el usuario, hubo un error.
  If objRS.EOF Then
    'No estaba el usuario registrado.
    CambiaClave = 1
  Else
    Session("AplicNombreDeUsuario") = TRIM(objRS("CodUsuario"))

    ' Ejecuta la instrucción que actualiza la clave de acceso
    strSqlQuery = " UPDATE usuario"&_
            " SET clave = '" & clave_nueva & "'" &_
            " WHERE CodUsuario = '" & Usuario & "'" &_
            " AND clave='" & clave_antigua & "'"
    Set objRS = objConn.Execute(strSqlQuery)
    ' Si hay un error lo muestra y termina la ejecución del programa
    If Err.Number <> 0 Then
      Call vs_MostrarError("No se pudo conectar a la Base de Datos", true )
    End If
    CambiaClave = 2
  End IF

  ' Cierra la Conexión con la Base de Datos
  objRS.Close()
  objConn.Close
  Set objConn = Nothing

End Function

' Procedimiento:	CambiaClaveSinChequeo
' Descripción:  	Cambia la clave de un usuario asumiendo que ya existe
' Argumentos: 		Usuario     Código del usuario que está conectándose
'       		ClaveNueva    Clave de acceso nueva
Sub CambiaClaveSinChequeo(ByRef objConn, ByVal Usuario, ByVal ClaveNueva)

   ' Encripta las claves de acceso
    clave_nueva = ClaveNueva

    ' Ejecuta la instrucción que actualiza la clave de acceso
    strSqlQuery = " UPDATE usuario"&_
            " SET clave = '" & clave_nueva & "'," &_
            " CamClave=getdate() " &_
            " WHERE CodUsuario = '" & Usuario & "'"
    Set objRS = objConn.Execute(strSqlQuery)
    ' Si hay un error lo muestra y termina la ejecución del programa
    If Err.Number <> 0 Then
      Call vs_MostrarError("No se pudo conectar a la Base de Datos", true )
    End If

End Sub

</SCRIPT>