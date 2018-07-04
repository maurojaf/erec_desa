<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<script language="JavaScript">
	function ventanaSecundaria (URL){
		window.open(URL,"DETALLE","width=200, height=200, scrollbars=no, menubar=no, location=no, resizable=no")
	}
	</script>

<%

	Response.CodePage=65001
	Response.charset ="utf-8"
		
	inicio= request("inicio")
	termino= request("termino")

	intCliente = request("CB_CLIENTE")
	strTipoCarga = request("CB_TIPO")


	abrirscg()
		If Trim(inicio) = "" Then
			inicio = TraeFechaActual(Conn)
			inicio = "01/" & Mid(TraeFechaActual(Conn),4,10)
		End If

		If Trim(termino) = "" Then
			termino = TraeFechaActual(Conn)
		End If


	strGraba = request("strGraba")
	  If Trim(strGraba) = "S" Then


	  	strRut = Trim(Request("TX_RUT"))
		strScoring = Trim(Request("TX_SCORING"))
		strNroDoc = Trim(Request("TX_NRO_DOC"))


		If Trim(strRut) <> "" then
			if ASC(RIGHT(strRut,1)) = 10 then strRut = Mid(strRut,1,len(strRut)-1)
			''if ASC(RIGHT(strScoring,1)) = 10 then strScoring = Mid(strScoring,1,len(strScoring)-1)
			If Trim(strNroDoc) <> "" Then
				if ASC(RIGHT(strNroDoc,1)) = 10 then strNroDoc = Mid(strNroDoc,1,len(strNroDoc)-1)
			End If

			if ASC(RIGHT(strRut,1)) = 13 then strRut = Mid(strRut,1,len(strRut)-1)
			''if ASC(RIGHT(strScoring,1)) = 13 then strScoring = Mid(strScoring,1,len(strScoring)-1)
			If Trim(strNroDoc) <> "" Then
				if ASC(RIGHT(strNroDoc,1)) = 13 then strNroDoc = Mid(strNroDoc,1,len(strNroDoc)-1)
			End If
		End if

		'rESPONSE.WRITE "<br>strScoring=" & "-" & strScoring & "-"
		'rESPONSE.WRITE "<br>strRut=" & "-" & strRut & "-"
		'rESPONSE.WRITE "RIGHT=" & "-" & RIGHT(strScoring,1) & "-"
		'rESPONSE.WRITE "comp=" & "-" & ASC(RIGHT(strScoring,1)) & "-"

		vRut = split(strRut,CHR(13))
		vScoring = split(strScoring,CHR(13))
		vNroDoc = split(strNroDoc,CHR(13))

		'Response.write "<br>ASC = " & ASC(MID(strRut,11,1))

		intTamvRut=ubound(vRut)
		intTamvNroDoc=ubound(vNroDoc)
		intTamvScoring=ubound(vScoring)

		'Response.write "<br>intTamvRut = " & intTamvRut
		'Response.write "<br>intTamvNroDoc = " & intTamvNroDoc
		'Response.write "<br>intTamvScoring = " & intTamvScoring
		'Response.End


	  		For indice = 0 to intTamvRut
				if intTamvRut <> -1 Then
					strIdRut = Replace(vRut(indice), chr(13),"")
					strIdRut = Replace(vRut(indice), chr(10),"")
					
				End If
				if intTamvNroDoc <> -1 Then
					strNroDocumento = ucase(Replace(vNroDoc(indice), chr(13),""))
					strNroDocumento = ucase(Replace(vNroDoc(indice), chr(10),""))
					
				End If
				if intTamvScoring <> -1 Then
					intCodigo = ucase(Replace(vScoring(indice), chr(13),""))
					intCodigo = ucase(Replace(vScoring(indice), chr(10),""))
					
				End If
					
				If Trim(strTipoCarga) = "ELIMINAR_DEUDOR" Then

					strSql = "DELETE FROM GESTIONES_CUOTA WHERE Id_Gestion IN (SELECT ID_GESTION FROM GESTIONES WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "')"
					set rsUpdate = Conn.execute(strSql)					

					strSql = "DELETE FROM GESTIONES WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)

					strSql = "DELETE FROM CUOTA WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)
					
					strSql = "DELETE FROM DEUDOR WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
					set rsUpdate = Conn.execute(strSql)

				ElseIf Trim(strTipoCarga) = "MODIFICAR_DEUDOR_EC" Then
					strSql = "UPDATE DEUDOR SET ETAPA_COBRANZA = '" & intCodigo & "', FECHA_ESTADO_ETAPA = getdate() WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
				ElseIf Trim(strTipoCarga) = "MODIFICAR_DEUDOR_NOMBRE" Then
					strSql = "UPDATE DEUDOR SET NOMBRE_DEUDOR = '" & intCodigo & "' WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
				ElseIf Trim(strTipoCarga) = "MODIFICAR_DEUDOR_REPLEG" Then
					strSql = "UPDATE DEUDOR SET REPLEG_NOMBRE = '" & intCodigo & "' WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
				ElseIf Trim(strTipoCarga) = "MODIFICAR_DEUDOR_ADIC1" Then
					strSql = "UPDATE DEUDOR SET ADIC_1 = '" & intCodigo & "' WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
				ElseIf Trim(strTipoCarga) = "MODIFICAR_DEUDOR_ADIC2" Then
					strSql = "UPDATE DEUDOR SET ADIC_2 = '" & intCodigo & "' WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
				ElseIf Trim(strTipoCarga) = "MODIFICAR_DEUDOR_ADIC3" Then
					strSql = "UPDATE DEUDOR SET ADIC_3 = '" & intCodigo & "' WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "'"
				ElseIf Trim(strTipoCarga) = "MODIFICAR_FEC_AGEND_CUOTA" Then				
					strSql = "UPDATE CUOTA "
					strSql = strSql & " SET FECHA_AGEND_ULT_GES = '" & intCodigo & "'"
					strSql = strSql & " FROM CUOTA C INNER JOIN ESTADO_DEUDA ED ON C.ESTADO_DEUDA=ED.CODIGO"
					strSql = strSql & " WHERE COD_CLIENTE = '" & intCliente & "' AND RUT_DEUDOR = '" & strIdRut & "' AND ED.ACTIVO=1"
				End If

	   			'Response.write "<br>strSql=" & strSql
				'Response.End
				set rsUpdate = Conn.execute(strSql)
			Next
		%>
		<script>
			alert('Proceso realizado correctamente');
		</script>
		<%


	  End if

	cerrarscg()
 
	If Trim(intCodUsuario) = "" Then intCodUsuario = session("session_idusuario")

	''Response.write "intCliente=" & intCliente

	If Trim(intCliente) = "" Then intCliente = session("ses_codcli")
	%>
	<title>UTILITARIO</title>
	<style type="text/css">
	<!--
	.Estilo37 {color: #FFFFFF}
	-->
	</style>
</head>
<body>
<div class="titulo_informe">UTILITARIO</div>	
<br>
<table width="800" align="CENTER" border="0">
   <tr>
    <td valign="top">
	<BR>
	<FORM name="datos" method="post">
	<table width="50%" border="0" ALIGN="CENTER" class="intercalado">
		<thead>
		<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="25%">MANDANTE</td>
			<!--td width="25%">ASIGNACION</td-->
			<td width="25%">TIPO PROCESO</td>
		</tr>
		</thead>
		<tr>
			<td>
			<select name="CB_CLIENTE" onChange="refrescar();">
				<%
				abrirscg()
				ssql="SELECT COD_CLIENTE,DESCRIPCION FROM CLIENTE WHERE COD_CLIENTE = '" & intCliente & "'"
				set rsCLI= Conn.execute(ssql)
				if not rsCLI.eof then
					Do until rsCLI.eof%>
					<option value="<%=rsCLI("COD_CLIENTE")%>" <%if Trim(intCliente)=rsCLI("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsCLI("descripcion")%></option>
				<%
					rsCLI.movenext
					Loop
					end if
					rsCLI.close
					set rsCLI=nothing
					cerrarscg()
				%>
			</select>
			</td>
			<td>
				<select name="CB_TIPO">
					<option value="SELECCIONAR" <%If strTipoCarga = "SELECCIONAR" then response.write "SELECTED"%>>SELECCIONAR</option>
					<option value="ELIMINAR_DEUDOR" <%If strTipoCarga = "ELIMINAR_DEUDOR" then response.write "SELECTED"%>>ELIMINAR DEUDOR</option>
					<option value="MODIFICAR_DEUDOR_NOMBRE" <%If strTipoCarga = "MODIFICAR_DEUDOR_NOMBRE" then response.write "SELECTED"%>>MODIFICAR NOMBRE DEUDOR</option>
					<option value="MODIFICAR_DEUDOR_REPLEG" <%If strTipoCarga = "MODIFICAR_DEUDOR_REPLEG" then response.write "SELECTED"%>>MODIFICAR NOMBRE REP LEG</option>
					<option value="MODIFICAR_DEUDOR_ADIC1" <%If strTipoCarga = "MODIFICAR_DEUDOR_ADIC1" then response.write "SELECTED"%>>MODIFICAR ADIC_1 DEUDOR</option>
					<option value="MODIFICAR_DEUDOR_ADIC2" <%If strTipoCarga = "MODIFICAR_DEUDOR_ADIC2" then response.write "SELECTED"%>>MODIFICAR ADIC_2 DEUDOR</option>
					<option value="MODIFICAR_DEUDOR_ADIC3" <%If strTipoCarga = "MODIFICAR_DEUDOR_ADIC3" then response.write "SELECTED"%>>MODIFICAR ADIC_3 DEUDOR</option>
					<option value="MODIFICAR_FEC_AGEND_CUOTA" <%If strTipoCarga = "MODIFICAR_DEUDOR_ADIC3" then response.write "SELECTED"%>>MODIFICAR AGEND CUOTA</option>
				</select>
			</td>
		</tr>
	</table>

<table width="50%" border="1" bordercolor="#FFFFFF" ALIGN="CENTER">
	<TR>
		<TD class=hdr_i>
			Rut<BR><BR>
			<TEXTAREA NAME="TX_RUT" ROWS=30 COLS=15><%=strRut%></TEXTAREA>
		</TD>
		<TD class=hdr_i>
			Codigo o Valor<BR><BR>
			<TEXTAREA NAME="TX_SCORING" ROWS=30 COLS=25><%=strScoring%></TEXTAREA>
		</TD>
	</TR>
	<TR>
		<TD colspan="3" ALIGN="CENTER">
			<INPUT TYPE="BUTTON" class="fondo_boton_100" value="Procesar" name="B1" onClick="envia('G');return false;">
		</TD>
	</TR>
</table>
</form>
<%
		If Trim(intCliente) <> "" and Trim(intCodRemesa) <> "" then
		abrirscg()
		End if
%>




	  </td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript1.2">
function envia(intTipo)	{
		if (document.forms[0].CB_TIPO.value == 'SELECCIONAR'){
			alert('Favor seleccionar Opción');
		}
		if (document.forms[0].CB_TIPO.value == 'ELIMINAR_DEUDOR'){
					if (confirm("¿ Está REALMENTE seguro de eliminar los Deudores ingresados ? Este proceso eliminará aparte del deudor los documentos y las gestiones asociadas al deudor - cliente"))
					{
						if (confirm("¿ Está REALMENTE seguro de eliminar los deudores ingresados ?"))
						{
							if (intTipo=='G'){
										document.forms[0].action='utilitario_carga.asp?strGraba=S';
									}else{
										document.forms[0].action='utilitario_carga.asp?strRefrescar=C';
									}
							document.forms[0].submit();
						}
					}
		}
		else
		{
			if (intTipo=='G'){
						document.forms[0].action='utilitario_carga.asp?strGraba=S';
					}else{
						document.forms[0].action='utilitario_carga.asp?strRefrescar=C';
					}
			document.forms[0].submit();
		}


}

function RefrescaDatos(){
	document.forms[0].submit();
}
</script>
