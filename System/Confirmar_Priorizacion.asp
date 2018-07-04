<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

<%
	strOrigen = request("strOrigen")
	
	If strOrigen = "2" Then %>
		<!--#include file="sesion_inicio.asp"-->
	<% Else %>
		<!--#include file="sesion.asp"-->
	<% End If %>

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>

	<link rel="stylesheet" href="../css/style.css">
	<link rel="stylesheet" href="../css/style_generales_sistema.css">

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<%
	Response.CodePage=65001
	Response.charset ="utf-8" 

	intIdgestion = request("id_gestion")

	'Response.write "<br>strOrigen=" & strOrigen

	strRutDeudor = request("strRutDeudor")

	strNuevaObs = Trim(UCASE(Mid(Request("TX_OBSERVACIONES"),1,600)))

	strGraba = request("strGraba")
	strCodCliente = request("strCodCliente")
	'Response.write "<br>strCodCliente=" & strCodCliente
	'Response.write "<br>strRutDeudor=" & strRutDeudor
	intIdusuario=session("session_idusuario")

	'Response.write "<br>strRutDeudor=" & strRutDeudor

	AbrirSCG()

	If strRutDeudor <> "" then
		strNombreDeudor = TraeNombreDeudor(Conn,strRutDeudor)
	Else
		strNombreDeudor=""
	End if

	CerrarSCG()

	If Trim(request("strGraba")) = "SI" Then

		AbrirSCG()

		strSql = "SELECT ID_CUOTA FROM PRIORIZACIONES_CUOTA"
		strSql = strSql & " WHERE ID_PRIORIZACION = " & intIdGestion

		set rsTemp= Conn.execute(strSql)


		Do until rsTemp.eof

			strObjeto = "CH_" & Replace(Trim(rsTemp("ID_CUOTA")),"-","_")

			If UCASE(Request(strObjeto)) = "ON" Then

				intIdcuota = rsTemp("ID_CUOTA")

				strSql = "UPDATE PRIORIZACIONES_CUOTA SET ESTADO_PRIORIZACION = '1', FECHA_ESTADO = GETDATE(), OBSERVACION_CONF_PRIORIZACION = ' "& strNuevaObs & "', ID_USUARIO_ESTADO = " & intIdusuario
				strSql = strSql & " WHERE ESTADO_PRIORIZACION = '0' AND ID_PRIORIZACION = " & intIdGestion & " AND ID_CUOTA = " & intIdcuota

				AbrirSCG1()
				set rsUpdate=Conn1.execute(strSql)
				CerrarSCG1()

			End if

		rsTemp.movenext

		Loop
		rsTemp.close
		set rsTemp=nothing
		CerrarSCG()
		
		If strOrigen = 1 Then
			%>
			<SCRIPT>
				location.href='listado_priorizacion.asp?';
			</SCRIPT>
			<%
		Else
			%>
			<SCRIPT>
				location.href='priorizar_caso.asp?strRut=<%=strRutDeudor%>&strOrigen=1';
			</SCRIPT>
			<%
		End If
		

	End If

%>
	<title>Empresa</title>
	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo27 {color: #FFFFFF}
	.Estilo1 {
		color: #FF0000;
		font-weight: bold;
		font-family: Arial, Helvetica, sans-serif; }
	-->
	</style>

</head>
<body>
<form name="datos" method="post">
<INPUT TYPE="hidden" NAME="intOrigen" value="<%=intOrigen%>">
<INPUT TYPE="hidden" NAME="strAgendar" value="">
<div class="titulo_informe">Confirmación de Priorizaciones</div>

<table width="90%" border="0" bordercolor="#999999" align="CENTER">
    <tr>
    <td>

	<table width="100%" border="0" bordercolor="#FFFFFF" class="estilo_columnas">
		<thead>
		<tr >
			<td>MANDANTE</td>
			<td>RUT DEUDOR</td>
			<td>NOMBRE O RAZON SOCIAL:</td>
			<td>USUARIO</td>
			<td>SUCURSAL</td>
			<td>FECHA</td>
		</tr>
	</thead>
	      <tr bgcolor="#FFFFFF" class="Estilo8">
	      <td>
	      	<select name="CB_CLIENTE">
				<%
					AbrirSCG()

					ssql="SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE COD_CLIENTE = '" & strCodCliente & "' ORDER BY RAZON_SOCIAL"
					set rsTemp= Conn.execute(ssql)
					if not rsTemp.eof then
						do until rsTemp.eof%>
						<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCodCliente)=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("RAZON_SOCIAL")%></option>
						<%
						rsTemp.movenext
						loop
					end if
					rsTemp.close
					set rsTemp=nothing

					CerrarSCG()
				%>
			</select>
			</td>

			<td ALIGN="LEFT"><%=strRutDeudor%></td>
			<td><%=strNombreDeudor%><INPUT TYPE="hidden" NAME="strRutDeudor" value="<%=strRutDeudor%>"> </td>

			<td ALIGN="RIGHT"><%=session("nombre_user")	%></td>

	        <td><%=nom_sucursal%></td>
	        <td><%=DATE%></td>
	      </tr>
    </table>
	</td>
	</tr>

	<tr>
	<td>
	<div class="subtitulo_informe">> Detalle de Deuda</div>	
	<table width="100%" border="0" ALIGN="CENTER" >

	  <tr>
	    <td>
		<%
		If Trim(strRutDeudor) <> "" then

		AbrirSCG()

			strSql=" SELECT CUOTA.ID_CUOTA, RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS,ESTADO_DEUDA.ACTIVO,"
			strSql = strSql & "	CUOTA.RUT_DEUDOR, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC,"
			strSql = strSql & "	IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,isnull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS,"
			strSql = strSql & "	IsNull(CUOTA.USUARIO_ASIG,0) as USUARIO_ASIG, ESTADO_DEUDA.DESCRIPCION AS ESTADO_DEUDA,"
			strSql = strSql & "	NRO_CUOTA, SUCURSAL, NRO_DOC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,"
			strSql = strSql & "	TIPO_DOCUMENTO.NOM_TIPO_DOCUMENTO, ISNULL(CUSTODIO,'LLACRUZ') AS CUSTODIO,PRC.ESTADO_PRIORIZACION AS CONFIRMACION_PR,"
			strSql = strSql & "	PRC.OBSERVACION_CONF_PRIORIZACION AS OBS,CUOTA.HORA_AGEND_ULT_GES,PRC.OBSERVACION_CONF_PRIORIZACION, U1.LOGIN AS USUARIO_ESTADO, ISNULL((SUBSTRING(CONVERT(VARCHAR(11),PRC.FECHA_ESTADO,6),1,7) + '/ ' + SUBSTRING(CONVERT(VARCHAR(10),PRC.FECHA_ESTADO,108),1,5)),'SIN AGEND') AS FECHA_ESTADO"

			strSql = strSql & "	FROM CUOTA INNER JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
			strSql = strSql & "			   INNER JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
			strSql = strSql & "			   INNER JOIN PRIORIZACIONES_CUOTA PRC ON CUOTA.ID_CUOTA = PRC.ID_CUOTA AND PRC.ID_PRIORIZACION = " & intIdGestion
			strSql = strSql & "			   INNER JOIN PRIORIZACION PR ON PR.ID_PRIORIZACION = " & intIdGestion
			strSql = strSql & " 		   INNER JOIN USUARIO U1 ON U1.ID_USUARIO = PRC.ID_USUARIO_ESTADO"

			strSql = strSql & "	ORDER BY CUENTA, FECHA_VENC DESC"

			set rsDET=Conn.execute(strSql)
			if not rsDET.eof then
			%>
			  <table width="100%" class="intercalado" style="width:100%">
			  	<thead>
		        <tr >
		          <td>
		          	<a href="#" onClick="marcar_boxes(true);">M</a>&nbsp;&nbsp;&nbsp;
	    			<a href="#" onClick="desmarcar_boxes(true);">D</a>
	    		  </td>
	    		  <td align="CENTER">ESTADO</td>
	    		  <td align="CENTER">FECHA ESTADO</td>
	    		  <td align="CENTER">USUARIO ESTADO</td>
		          <td align="CENTER">OBS.</td>

	    		  <td align="CENTER">RUT CLIENTE</td>
	    		  <td align="CENTER">NOMBRE CLIENTE</td>
		          <td align="CENTER">NRO. DOC</td>
	    		  <td align="CENTER">CUOTA</td>
		          <td align="CENTER">F.VENCIM.</td>
		          <td align="CENTER">ANT.</td>
		          <td align="CENTER">TIPO DOC</td>
		          <td align="CENTER">CAPITAL</td>
		          <td align="CENTER">SALDO</td>
		          </tr>
		         </thead>
		         <tbody> 
				<%
				strArrConcepto = ""
				strArrID_CUOTA = ""

				Do While Not rsDET.eof

				strObsConf = Trim(rsDET("OBSERVACION_CONF_PRIORIZACION"))

				If strObsConf = "" then
					strObsConf = "SIN OBSERVACION"
				End If

				If Trim(rsDET("CONFIRMACION_PR")) <> "0" OR rsDET("ACTIVO") = 0 Then
					strConfirmada = "GESTIONADO"
				Else
					strConfirmada = "PENDIENTE"
					strObsConf = "NO CONFIRMADO"
				End If

				strArrConcepto = strArrConcepto & ";" & "CH_" & rsDET("ID_CUOTA")
				strArrID_CUOTA = strArrID_CUOTA & ";" & rsDET("ID_CUOTA")

				%>
		        <tr bordercolor="#999999" >
		          <td><INPUT TYPE="checkbox" NAME="CH_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>"></td>
		          <td><div align="left"><%=strConfirmada%></div></td>
		          <td><div align="right"><%=rsDET("FECHA_ESTADO")%></div></td>
		          <td><div align="right"><%=rsDET("USUARIO_ESTADO")%></div></td>
		          <td align="center"><img src="../imagenes/priorizar_normal.png" border="0" height="15" title="<%=strObsConf%>">

		          <td><%=rsDET("RUT_SUBCLIENTE")%></td>

				  <td class="Estilo4" title="<%=rsDET("NOMBRE_SUBCLIENTE")%>">
				  <%=Mid(rsDET("NOMBRE_SUBCLIENTE"),1,30)%></td>

		          <td><div align="right"><%=rsDET("NRO_DOC")%></div></td>
		          <td><div align="right"><%=rsDET("NRO_CUOTA")%></div></td>
		          <td><div align="right"><%=rsDET("FECHA_VENC")%></div></td>
		          <td><div align="right"><%=rsDET("ANTIGUEDAD")%></div></td>
		          <td><div align="right"><%=rsDET("NOM_TIPO_DOCUMENTO")%></div></td>
		          <td align="right" >$ <%=FN((rsDET("VALOR_CUOTA")),0)%></td>
		          <td align="right" >$ <%=FN((rsDET("SALDO")),0)%></td>
				 </tr>
				 <%

				 rsDET.movenext
				 loop

				vArrConcepto = split(strArrConcepto,";")
				vArrID_CUOTA = split(strArrID_CUOTA,";")
				intTamvConcepto = ubound(vArrConcepto)

				 %>
				</tbody>
			</table>

			<table width="100%" border="0" class="estilo_columnas">
				<thead>
				<tr>
					<td>OBSERVACIONES (Max. 300 Caract.)</td>
				</tr>
				</thead>

				<tr>
				   <td align="left">
				  <TEXTAREA NAME="TX_OBSERVACIONES" ROWS="4" COLS="70"><%=strObsNueva%></TEXTAREA>
				  </td>
				 </tr>
			 </table>

			  <%end if
			  rsDET.close
			  set rsDET=nothing
		  Else
		  end if

		CerrarSCG()%>

	    </td>
	  </tr>

		<tr>
			<TD>

				<table width="100%" border="0" bordercolor="#FFFFFF">
							<tr bordercolor="#999999" class="Estilo8">
							<td align="right">
								<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Confirmar Gestión" value="Confirmar" onClick="envia();" class="Estilo8">
								&nbsp;
								<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Volver" value="Volver" onClick="history.back();" class="Estilo8">
							</td>
							</tr>
				</table>
			</TD>
		</tr>
	</table>

<td>
<tr>
</table>

<INPUT TYPE="hidden" NAME="strGraba" value="">
</form>
</body>
</html>

<script language="JavaScript" type="text/JavaScript">
$(document).ready(function(){
	$(document).tooltip();
})

marcar_boxes(true);

function envia(){

	if (confirm("¿Está seguro de confirmar la gestión de los documentos seleccionados?, si acepta no se podrán desconfirmar los documentos confirmados."))
	{
			datos.strGraba.value='SI';
	}		datos.submit();
}

function marcar_boxes(){
		<% For i=1 TO intTamvConcepto %>
			document.forms[0].<%=vArrConcepto(i)%>.checked=true;
		<% Next %>
}

function desmarcar_boxes(){
		<% For i=1 TO intTamvConcepto %>
			document.forms[0].<%=vArrConcepto(i)%>.checked=false;
		<% Next %>
}

</script>
















