<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
	<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<meta charset="utf-8">
	<link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->

	<title>VENTAS WEB</title>
</head>

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	If Trim(Request("Limpiar"))="1" Then session("IdCliente") = ""
	If Trim(Request("strRUT_DEUDOR")) <> "" Then session("IdCliente") = Trim(Request("strRUT_DEUDOR"))
	If Trim(strRUT_DEUDOR) = "" Then strRUT_DEUDOR = Trim(Request("strRUT_DEUDOR"))
	If Trim(strRUT_DEUDOR) = "" Then strRUT_DEUDOR = Trim(Request("TX_RUT"))


	''Response.write "strRUT_DEUDOR=" & strRUT_DEUDOR
%>
<%strTitulo="PANTALLA PRINCIPAL DE ASIGNACION"%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" id="datos" method="post">
<%
	strCOD_CLIENTE = request("strCOD_CLIENTE")
	strCOD_CLIENTE=session("ses_codcli")

	If Trim(Request("intGrabar"))="1" Then
		strRut = request("TX_RUT")
		intCodEjecutivo = Request("CB_EJECUTIVO")

		AbrirScg()
			If Trim(intCodEjecutivo) = "0" Then intCodEjecutivo = "NULL"

			strSql = "UPDATE CUOTA SET USUARIO_ASIG = " & intCodEjecutivo & " , FECHA_ASIGNACION = getdate() WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA where activo = 1)  "
			set rsUpdate = Conn.execute(strSql)

			strSql = "UPDATE DEUDOR SET USUARIO_ASIG = " & intCodEjecutivo & ", FECHA_ASIGNACION = getdate(), ASIGNADOR = " & session("session_idusuario") & " WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "'"

			'RESPONSE.WRITE "strSql=" & strSql
			'RESPONSE.End
			set rsUpdate2 = Conn.execute(strSql)


		CerrarScg()
		%>
		<script>
				alert('Deudor Asignado Correctamente');
		</script>
		<%
	End If

	intContacto = Trim(request("contacto"))
	if Trim(strRUT_DEUDOR) <> "" Then

		AbrirScg()
		strSql = "SELECT RUT_DEUDOR, USUARIO_ASIG FROM CUOTA "
		strSql = strSql & "WHERE COD_CLIENTE = " & strCOD_CLIENTE & " AND "
		strSql = strSql & "RUT_DEUDOR = '" & strRUT_DEUDOR & "'"

		'Response.write "<br>strSql = " & strSql
		'Response.write "<br>Conn = " & Conn
		'Response.End

		set rsDEU=Conn.execute(strSql)
		if not rsDEU.eof then
			'Response.write "<br>strSql = " & strSql
			strRUT_DEUDOR = rsDEU("RUT_DEUDOR")
			strEjeAsig = rsDEU("USUARIO_ASIG")

			strSql = "SELECT NOMBRE_DEUDOR FROM DEUDOR WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND "
			strSql = strSql & "RUT_DEUDOR = '" & strRUT_DEUDOR & "'"
			'Response.write strSql
			set rsTemp=Conn.execute(strSql)
			If not rsTemp.eof then
				strNombreAMostrar = rsTemp("NOMBRE_DEUDOR")
			Else
				strNombreAMostrar = ""
			End If

			existe = "si"
		else
			strRUT_DEUDOR = ""
			strEjeAsig = ""
			existe = "si"
		end if

		rsDEU.close
		set rsDEU=nothing
		CerrarScg()

	End If
%>
	              <table width="1000" height="335" border="0" ALIGN="CENTER">
                    <tr>
                      <td height="331" valign="top">
                      <table width="100%" border="0">
                        <tr>
                        <td class="Estilo38">
                        	<strong><font color="#009900">ID</font></strong>
                        </td>
                        <td class="Estilo38">
							<strong><font color="#009900">RUT</font></strong>
                        </td>
                        <td class="Estilo38">
							&nbsp
                        </td>
                     </tr>
					  <tr>
						<td>
						  <acronym title="IDENTIFICADOR DEL DEUDOR">
						  	<input name="rut" type="text" value="<%=rut%>" size="8" maxlength="6" onChange="//Valida_Rut(this.value,'rut')">
						  </acronym>
						</td>
						<td>
						  <acronym title="RUT">
							<input name="TX_RUT" type="text" value="<%=strRutCliente%>" size="8" maxlength="10">
						  </acronym>
						</td>
						<td>
						  <acronym title="DESPLEGAR DATOS DEL CONTACTO ">
						  	<input name="me_" type="button" id="me_" onClick="envia();" value="Buscar">
						  </acronym>
						  <acronym title="LIMPIAR PANTALLA">
						  	<input name="li_" type="button" onClick="window.navigate('asigna_manual.asp?Limpiar=1');" value="Limpiar">
						  </acronym>
						  <acronym title="IMPRIMIR FORMULARIO (CONFIGURACION HORIZONTAL)">
						  	<input name="li_" type="button" onClick="window.print();" value="Imprimir">
						  </acronym>
						  <input name="strCOD_CLIENTE" type="hidden" value="<%=strCOD_CLIENTE%>">
						  <input name="strRUT_DEUDOR" type="hidden" value="<%=strRUT_DEUDOR%>">
						  <input name="ANI" type="hidden" id="ANI" value="<%=ani%>">
						 </td>
						<td>
						<span class="Estilo50"><strong><font color="#000000"><%=strNombreAMostrar%></font></strong></span>
						</td>
					  </tr>
                        </table>

                    <%
		if (strRUT_DEUDOR <> "" and not isnull(strRUT_DEUDOR)) Then
			if existe = "si" then
				AbrirScg()
					if Trim(strEjeAsig)= "" Then
						strNomEjeAsig = "SIN ASIGNAR"
					Else
						'Response.write "strEjeAsig=" &strEjeAsig
						'Response.End
						If trim(strEjeAsig) <> "" Then
							strNomEjeAsig = TraeCampoId(Conn, "LOGIN", strEjeAsig, "USUARIO", "ID_USUARIO")
						End if
					End if
				CerrarScg()
				%>
					<font color="#009900"><strong>&nbspDATOS DEL DEUDOR</strong></font><BR>
					<table width="100%" border="0" bordercolor="#FFFFFF">
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					  <!--td width="2%">ID</td-->
					  <td width="12%">RUT</td-->
					  <td width="25%">NOMBRE</td>
					  <td width="15%">ASIGNACION ACTUAL</td>
					  <td width="20%">CAMBIAR POR</td>
					  <td width="20%">&nbsp</td>
					</tr>
					<tr bgcolor="#f6f6f6" class="Estilo8">
						<td><%=strRUT_DEUDOR%></td>
						<td><%=strNombreAMostrar%></td>
						<td><%=strNomEjeAsig%></td>
						<td>
							<select name="CB_EJECUTIVO" ID="CB_EJECUTIVO">
								<option value="0" <%if Trim(strEjeAsig)="0" then response.Write("Selected") end if%>>SELECCIONE</option>
								<%
								AbrirScg()
								strSql="SELECT ID_USUARIO,LOGIN FROM USUARIO WHERE ACTIVO = 1 AND PERFIL_COB = 1"
								If trim(intGrupo) <> "" and trim(intGrupo) <> "0" Then
									strSql = strSql & " and grupo = '" & intGrupo & "'"
								End if
								set rsEjecutivo=Conn.execute(strSql)
								if not rsEjecutivo.eof then
									do until rsEjecutivo.eof
									%>
									<option value="<%=rsEjecutivo("ID_USUARIO")%>" <%if Trim(strEjeAsig)=Trim(rsEjecutivo("ID_USUARIO")) then response.Write("selected") end if%>><%=ucase(rsEjecutivo("LOGIN"))%></option>
									<%rsEjecutivo.movenext
									loop
								end if
								rsEjecutivo.close
								set rsEjecutivo=nothing
								CerrarScg()
								%>
							</select>
        				</td>
        				<td>
						<acronym title="GRABAR GESTION">
							<input name="BT_GESTIONAR" type="button" onClick="gestionar();" value="Asignar">
						</acronym>
						</td>
					</tr>
					</table>
					<%
				end if
		end if
		%>
		</table>
	</td>
	</tr>
</table>
</form>

</body>
</html>


<script language="JavaScript" type="text/JavaScript">
function envia(){
	datos.action='asigna_manual.asp?Limpiar=1';
	datos.submit();
}

function buscar(){
	datos.action='asigna_manual.asp';
	datos.submit();
}

function gestionar(){
	if (datos.CB_EJECUTIVO.value == '') {
		alert('Debe Ingresar un ejecutivo');
		return;
	}
}

function gestionar(){
	datos.BT_GESTIONAR.disabled=true;
	datos.action='asigna_manual.asp?intGrabar=1';
	datos.submit();
}

</script>
