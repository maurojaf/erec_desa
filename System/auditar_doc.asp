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
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet" >	

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">


<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	strRutDeudor = request("rut")
	
	'response.write " strRut = " & strRutDeudor
	
	if trim(strRutDeudor) = "" Then
		strRutDeudor = session("session_RUT_DEUDOR") 
	End if
	
	session("session_RUT_DEUDOR") = strRutDeudor
	
	strGraba = request("strGraba")
	strCodCliente = session("ses_codcli")

	usuario=session("session_idusuario")

	AbrirSCG()

	strDocCancelados = request("TX_DOCCANCELADOS")
	strObservaciones = request("TX_OBSERVACIONES")
	strObservaciones = Trim(strObservaciones)


	If Trim(request("strGraba")) = "SI" Then

		intInteres = Request("TX_INTERES")
		intHonorarios = Request("TX_HONORARIOS")
		intProtestos = Request("TX_PROTESTOS")

		strSql = "SELECT ID_CUOTA, FECHA_ESTADO_FR , FECHA_ESTADO_NR , IsNull(FACTURA_RECEPCIONADA,2) AS FACTURA_RECEPCIONADA , IsNull(NOTIFICACION_RECEPCIONADA,2) AS NOTIFICACION_RECEPCIONADA FROM CUOTA WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE = '" & strCodCliente & "' AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) ORDER BY FECHA_VENC DESC"
		set rsCuota= Conn.execute(strSql)

		intCorrelativo = 1
		intTotalCapital = 0

		Do until rsCuota.eof

			intID_CUOTA = rsCuota("ID_CUOTA")

			intEstadoFRCuota = Trim(rsCuota("FACTURA_RECEPCIONADA"))
			intEstadoNRCuota = Trim(rsCuota("NOTIFICACION_RECEPCIONADA"))

			intEstadoFR=Trim(request("ra_FactRecep_" & Trim(intID_CUOTA)))
			intEstadoNR=Trim(request("ra_NotRecep_" & Trim(intID_CUOTA)))

			'Response.write "<br>intEstadoFR=" & intEstadoFR & " - " & "ra_FactRecep_" & Trim(intID_CUOTA)
			'Response.write "<br>intEstadoFRCuota=" & intEstadoFRCuota

			''Response.write "<br>Dif=" & (intEstadoFRCuota <> intEstadoFR)


			If (intEstadoFRCuota <> intEstadoFR) Then
				if intEstadoFR = "2" Then intEstadoFR = "NULL"
				strSql = "UPDATE CUOTA SET FECHA_ESTADO_FR = getdate(), FACTURA_RECEPCIONADA = " & intEstadoFR & " WHERE ID_CUOTA = " & intID_CUOTA
				Response.write "<br>strSql=" & strSql
				set rsUpdate=Conn.execute(strSql)
			End If

			'Response.write "<br>intEstadoNR=" & intEstadoNR  & " - " & "ra_NotRecep_" & Trim(intID_CUOTA)
			'Response.write "<br>intEstadoNRCuota=" & intEstadoNRCuota

			If (intEstadoNRCuota <> intEstadoNR) Then
				if intEstadoNR = "2" Then intEstadoNR = "NULL"
				strSql = "UPDATE CUOTA SET FECHA_ESTADO_NR = getdate(), NOTIFICACION_RECEPCIONADA = " & intEstadoNR & " WHERE ID_CUOTA = " & intID_CUOTA
				Response.write "<br>strSql=" & strSql
				set rsUpdate=Conn.execute(strSql)
			End If

			'Response.write "<br>strGestionAbono=" & strGestionAbono
			'Response.write "<br>" & strSql
			'Response.End

		rsCuota.movenext
		intCorrelativo = intCorrelativo + 1
	Loop
	rsCuota.close
	set rsCuota=nothing


		strSql = "UPDATE DEUDOR SET OBSERVACIONES_CONF_DOC = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "'"
		strSql = strSql & " WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE = '" & strCodCliente & "'"
		'Response.write "<br>strSql" & strSql
		set rsUpdate=Conn.execute(strSql)


	'Response.eND

	Response.redirect "detalle_gestiones.asp?rut=" & strRutDeudor & "&cliente=" & strCodCliente


	End If

	If Trim(request("strAgendar")) = "SI" Then

		dtmFecAgend = Request("TX_FEC_AGEND")
		strHoraAgend = Request("TX_HORA_AGEND")
		If trim(strHoraAgend)="" Then strHoraAgend = "08:00"

		'Response.write "dtmFecAgend=" & dtmFecAgend

		If dtmFecAgend <> "" Then
			dtmFecAgend = dtmFecAgend & " " &  strHoraAgend & ":00"
		End If

		strSql = "UPDATE DEUDOR SET OBSERVACIONES_CONF_DOC = '" & Mid(Request("TX_OBSERVACIONES"),1,300) & "'"
		strSql = strSql & " WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE = '" & strCodCliente & "'"
		'Response.write "<br>strSql" & strSql
		set rsUpdate=Conn.execute(strSql)

		If dtmFecAgend = "" Then
			strSql = "UPDATE DEUDOR SET FECHA_AGEND_AUD = NULL, HORA_AGEND_AUD = NULL"
			strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		Else
			strSql = "UPDATE DEUDOR SET FECHA_AGEND_AUD = '" & dtmFecAgend & "', HORA_AGEND_AUD = '" & strHoraAgend & "'"
			strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"
		End If
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)
			
		%>

		<SCRIPT>
			IrAPrincipal()
		</SCRIPT>
		<%
	End If

	If Trim(request("strLimpiar")) = "SI" Then

		'strSql = "UPDATE DEUDOR SET FECHA_AGEND_AUD = NULL, HORA_AGEND_AUD = NULL"
		'strSql = strSql & " WHERE COD_CLIENTE = '" & strCodCliente & "' AND RUT_DEUDOR = '" & strRutDeudor & "'"

		''Response.write "strSql=" & strSql
		'set rsUpdate=Conn.execute(strSql)

		strSql = "UPDATE DEUDOR SET OBSERVACIONES_CONF_DOC = NULL"
		strSql = strSql & " WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE = '" & strCodCliente & "'"
		'Response.write "<br>strSql" & strSql
		set rsUpdate=Conn.execute(strSql)


		%>

		<SCRIPT>
			IrAPrincipal()
		</SCRIPT>
		<%
	End If


%>

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo27 {color: #FFFFFF}
	.Estilo1 {
		color: #FF0000;
		font-weight: bold;
		font-family: Arial, Helvetica, sans-serif;
	-->

	}
	</style>

	<script language="JavaScript " type="text/JavaScript">

	$(document).ready(function(){

		$('#TX_FEC_AGEND').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})

	})
	function Refrescar(rut)
	{
		if(rut == '')
		{
			return
		}
				datos.action = "auditar_doc.asp?rut=" + rut + "&tipo=1";
				datos.submit();

	}

	</script>


	<link href="../css/style.css" rel="Stylesheet">
</head>
<body>
<form name="datos" method="post">
<INPUT TYPE="hidden" NAME="strAgendar" value="">

<div class="titulo_informe">Auditoria de documentos</div>
<table width="90%" border="0" bordercolor="#999999" align="CENTER">
    <tr>
    <td>

	  <%

	If strRutDeudor <> "" then
		strNombreDeudor = TraeNombreDeudor(Conn,strRutDeudor)
	Else
		strNombreDeudor=""
	End if


	%>

	<table width="100%" border="0" bordercolor="#FFFFFF"  class="estilo_columnas">
		<thead>
		<tr >
			<td>RUT DEUDOR</td>
			<td>NOMBRE O RAZON SOCIAL</td>
			<td>USUARIO</td>
			<td>FECHA</td>
		</tr>
		</thead>
	      <tr bgcolor="#FFFFFF" class="Estilo8">

			<td>
				<A HREF="principal.asp?TX_RUT=<%=strRutDeudor%>">
				<acronym title="Llevar a pantalla de selección"><%=strRutDeudor%></acronym>
				</A>
			</td>
								
			<td><%=strNombreDeudor%><INPUT TYPE="hidden" NAME="rut" value="<%=strRutDeudor%>"> </td>

			<td ALIGN="Left"><%=session("nombre_user")	%></td>

	        <td><%=DATE%></td>
	      </tr>
    </table>
	
	<td>
	<tr>

	<table width="90%" border="0" ALIGN="CENTER">
	 <tr>
	 	<td height="22">
	 	<font class="subtitulo_informe">> Detalle de Documentos</font>
	 	</td>
	</tr>
	</table>

	<table width="100%" border="0" ALIGN="CENTER">
	  <tr>
	    <td valign="top">
		<%
		If Trim(strRutDeudor) <> "" then
		abrirscg()


			strSql="SELECT OBSERVACIONES_CONF_DOC, CONVERT(VARCHAR(10),FECHA_AGEND_AUD,103) AS FECHA_AGEND_AUD, HORA_AGEND_AUD FROM DEUDOR WHERE RUT_DEUDOR = '" & strRutDeudor & "' AND COD_CLIENTE='" & strCodCliente & "'"
			set rsDeudor=Conn.execute(strSql)
			If not rsDeudor.eof then
				strObsConfDeudor = rsDeudor("OBSERVACIONES_CONF_DOC")
				dtmFechaAgend = Trim(rsDeudor("FECHA_AGEND_AUD"))
				dtmHoraAgend = Trim(rsDeudor("HORA_AGEND_AUD"))
			End If


			strSql="SELECT ID_CUOTA, RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, CONVERT(VARCHAR(10),FECHA_CREACION,103) as FECHA_CREACION , CONVERT(VARCHAR(10),FECHA_ESTADO_FR,103) as FECHA_ESTADO_FR , CONVERT(VARCHAR(10),FECHA_ESTADO_NR,103) as FECHA_ESTADO_NR , ISNULL(FACTURA_RECEPCIONADA,2) AS FACTURA_RECEPCIONADA, ISNULL(NOTIFICACION_RECEPCIONADA,2) AS NOTIFICACION_RECEPCIONADA ,DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, RUT_DEUDOR, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC, IsNull(datediff(d,FECHA_VENC,getdate()),0) as ANTIGUEDAD, NRO_DOC, IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,isnull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS,IsNull(USUARIO_ASIG,0) as USUARIO_ASIG, NRO_CUOTA, SUCURSAL , ESTADO_DEUDA, COD_REMESA, CUENTA, NRO_DOC, TIPO_DOCUMENTO, NOM_TIPO_DOCUMENTO, CUSTODIO FROM CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='"& strRutDeudor &"' AND COD_CLIENTE='"& strCodCliente &"' AND SALDO > 0 AND ESTADO_DEUDA IN ('1','7','8') AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO ORDER BY FECHA_VENC DESC"
			'response.Write(strSql)
			'response.End()
			set rsDET=Conn.execute(strSql)
			if not rsDET.eof then
			%>
			  <table width="100%" border="0" bordercolor="#FFFFFF" class="intercalado">
			  	<thead>
		        <tr bordercolor="#999999" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					<td align="CENTER">RUT CLIENTE</td>
					<td align="CENTER">NOMBRE CLIENTE</td>
					<td align="CENTER">DOC</td>
					<td align="CENTER">CUOTA</td>
					<td align="CENTER">FECHA VENCIM.</td>
					<td align="CENTER">TIPO</td>
					<td align="CENTER">CAPITAL</td>
					<td align="CENTER" colspan=4>FACTURA RECEPCEPCIONADA</td>
					<td align="CENTER">FECHA ESTADO</td>
					<td align="CENTER"  colspan=4>NOTIFICACIÓN RECEPCEPCIONADA</td>
					<td align="CENTER">FECHA ESTADO</td>
                </tr>
                 <tr bordercolor="#999999" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					<td align="CENTER" bgcolor="#FFFFFF" colspan=7>&nbsp;</td>
					<td align="CENTER">SI</td>
					<td align="CENTER">EC</td>
					<td align="CENTER">NO</td>
					<td align="CENTER">SA</td>
					<td align="CENTER" bgcolor="#FFFFFF" >&nbsp;</td>
					<td align="CENTER">SI</td>
					<td align="CENTER">EC</td>
					<td align="CENTER">NO</td>
					<td align="CENTER">SA</td>
					<td align="CENTER" bgcolor="#FFFFFF" >&nbsp;</td>
                </tr>
            	</thead>
            	<tbody>
				<%
				intSaldo = 0
				intValorCuota = 0
				total_ValorCuota = 0
				strArrConcepto = ""
				strArrID_CUOTA = ""

				Do until rsDET.eof

				strEstadoFR = Trim(rsDET("FACTURA_RECEPCIONADA"))
				strEstadoNR = Trim(rsDET("NOTIFICACION_RECEPCIONADA"))

				'Response.write "<br>strEstadoFR=" & strEstadoFR
				'Response.write "<br>strEstadoNR=" & strEstadoNR

				strArrConcepto = strArrConcepto & ";" & "CH_" & rsDET("ID_CUOTA")
				strArrID_CUOTA = strArrID_CUOTA & ";" & rsDET("ID_CUOTA")

				intSaldo = Round(session("valor_moneda") * ValNulo(rsDET("SALDO"),"N"),0)
				intValorCuota = Round(session("valor_moneda") * ValNulo(rsDET("VALOR_CUOTA"),"N"),0)



				strNroDoc = Trim(rsDET("NRO_DOC"))
				intID_CUOTA = Trim(rsDET("ID_CUOTA"))
				strNroCuota = Trim(rsDET("NRO_CUOTA"))
				strSucursal = Trim(rsDET("SUCURSAL"))
				strEstadoDeuda = Trim(rsDET("ESTADO_DEUDA"))
				strCodRemesa = Trim(rsDET("COD_REMESA"))
				intProtestoDoc = ValNulo(rsDET("GASTOS_PROTESTOS"),"N")


				dtmFechaEstadoFR = rsDET("FECHA_ESTADO_FR")
				dtmFechaEstadoNR = rsDET("FECHA_ESTADO_NR")

				If Trim(dtmFechaEstadoFR) = "" or IsNull(dtmFechaEstadoFR) Then dtmFechaEstadoFR = rsDET("FECHA_CREACION")
				If Trim(dtmFechaEstadoNR) = "" or IsNull(dtmFechaEstadoNR) Then dtmFechaEstadoNR = rsDET("FECHA_CREACION")

				strCheckNR = ""
				strCheckFR = ""
				%>
		        <tr bordercolor="#999999" >
					<td><div align="LEFT"><%=rsDET("RUT_SUBCLIENTE")%></div></td>
					<td><div align="left"><%=Mid(rsDET("NOMBRE_SUBCLIENTE"),1,25)%></div></td>
					<td><div align="right"><%=rsDET("NRO_DOC")%></div></td>
					<td><div align="right"><%=strNroCuota%></div></td>
					<td><div align="right"><%=rsDET("FECHA_VENC")%></div></td>
					<td><div align="right"><%=rsDET("NOM_TIPO_DOCUMENTO")%></div></td>
					<td align="right" >$ <%=FN((intSaldo),0)%></td>

					<td align="center">
					<% if strEstadoFR="1" then strCheckFR = "checked" Else strCheckFR = "" %>
					<input name="ra_FactRecep_<%=intID_CUOTA%>" type="radio" value="1"  <%=strCheckFR%>>
					</td>

					<% if strEstadoFR="3" then strCheckFR = "checked" Else strCheckFR = "" %>
					<td align="center" >
					<input name="ra_FactRecep_<%=intID_CUOTA%>" type="radio" value="3"  <%=strCheckFR%>>
					</td>

					<% if strEstadoFR="0" then strCheckFR = "checked" Else strCheckFR = "" %>
					<td align="center" >
					<input name="ra_FactRecep_<%=intID_CUOTA%>" type="radio" value="0" <%=strCheckFR%>>
					</td>

					<% if strEstadoFR="2" then strCheckFR = "checked" Else strCheckFR = "" %>
					<td align="center" >
					<input name="ra_FactRecep_<%=intID_CUOTA%>" type="radio" value="2" <%=strCheckFR%>>
					</td>
					<td><div align="right"><%=dtmFechaEstadoFR%></div></td>

					<td align="center" >
					<% if strEstadoNR="1" then strCheckNR = "checked" Else strCheckNR = "" %>
					<input name="ra_NotRecep_<%=intID_CUOTA%>" type="radio" value="1" <%=strCheckNR%>>
					</td>

					<td align="center" >
					<% if strEstadoNR="3" then strCheckNR = "checked" Else strCheckNR = "" %>
					<input name="ra_NotRecep_<%=intID_CUOTA%>" type="radio" value="3"  <%=strCheckNR%>>
					</td>


					<td align="center" >
					<% if strEstadoNR="0" then strCheckNR = "checked" Else strCheckNR = "" %>
					<input name="ra_NotRecep_<%=intID_CUOTA%>" type="radio" value="0" <%=strCheckNR%>>
					</td>


					<td align="center" >
					<% if strEstadoNR="2" then strCheckNR = "checked" Else strCheckNR = "" %>
					<input name="ra_NotRecep_<%=intID_CUOTA%>" type="radio" value="2" <%=strCheckNR%>>
					</td>
					<td><div align="right"><%=dtmFechaEstadoNR%></div></td>

				 </tr>
				 <%rsDET.movenext
				 Loop

					vArrConcepto = split(strArrConcepto,";")
					vArrID_CUOTA = split(strArrID_CUOTA,";")
					intTamvConcepto = ubound(vArrConcepto)

				 %>
				</tbody>
		      </table>
			  <%end if
			  rsDET.close
			  set rsDET=nothing
		  Else
		  %>
			<table width="100%" border="0" bordercolor="#FFFFFF">
			<tr bordercolor="#999999" class="Estilo8">
			<td align="center">
				Deudor no posee documentos activos
			</td>
			</tr>
			</table>
		  <%end if%>

	    </td>
	  </tr>

	</table>

	</td>
	</tr>


	<tr>
		<td height="20"  bordercolor="#999999"  class="Estilo27" align="center">

			<table width="90%" border="0" bordercolor="#FFFFFF" class="estilo_columnas" ALIGN="CENTER">
			<thead>
			<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" >
				<td>OBSERVACIONES (Max. 300 Caract.)</td>
			</tr>
			</thead>
			<tr>
				<td align="LEFT">
				&nbsp;
				<TEXTAREA NAME="TX_OBSERVACIONES" ROWS=4 COLS=70><%=strObsConfDeudor%></TEXTAREA>
				</td>
			</tr>
			 </table>

			<table width="90%" ALIGN="CENTER">
			<tr>
				<td align="left">
					<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Auditar" value="Auditar" onClick="envia();" class="Estilo8">
				</td>
				<td align="right">
					<INPUT TYPE="BUTTON" NAME="Ver Gestiones" class="fondo_boton_100" value="Ver Gestiones" onClick="ir_detalle_gestiones();" class="Estilo8">
					<INPUT TYPE="BUTTON" NAME="Limpiar" class="fondo_boton_100" value="Limpiar" onClick="LimpiarDatos();" class="Estilo8">
					<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Volver" value="Volver" onClick="history.back();" class="Estilo8">
				</td>
			</tr>
			</table>


		</td>
	</tr>

	<tr>
		<TD ALIGN="CENTER">
			<table width="90%" border="0" bordercolor="#FFFFFF" class="estilo_columnas" ALIGN="CENTER">
				<thead>
				<tr>
					<td>FECHA AGENDAMIENTO</td>
					<td colspan = "2" align="left">HORA AGEND.</td>
				</tr>
			</thead>
				<tr>
					 <td width = "200" >
						<input name="TX_FEC_AGEND" type="text" id="TX_FEC_AGEND" value="<%=dtmFechaAgend%>" size="10" maxlength="10">

					 </td>
					 <td width="100" align="left">
						<input name="TX_HORA_AGEND" type="text" id="TX_HORA_AGEND" value="<%=dtmHoraAgend%>" size="5" maxlength="5" onChange="return ValidaHora(this,this.value)">
					</td>
				<% If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then %>
					<td align="left" >
						<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="BT_AGENDAR" value="Agendar" onClick="Agendar();" class="Estilo8">
					</td>
				<% Else%>
					<td align="left" >&nbsp;</td>						
				<% End If %>
				</tr>
			</table>
		</TD>
	</tr>

</table>

<INPUT TYPE="hidden" NAME="strGraba" value="">
<INPUT TYPE="hidden" NAME="strLimpiar" value="">
</form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">

function envia(){
	if (confirm("¿ Está seguro de auditar los documentos ? "))
	{
		datos.strGraba.value='SI';
		datos.submit();
	}
}

function Agendar() {

	if (confirm("¿ Está seguro de agendar ? "))
	{
			datos.strAgendar.value='SI';
			datos.submit();
	}

}

function ir_detalle_gestiones(){

	datos.action="detalle_gestiones.asp?rut=<%=strRutDeudor%>&cliente=<%=strCodCliente%>";
	datos.submit();
}

function LimpiarDatos(){
	if (confirm("¿ Está seguro de Borrar la observación ?"))
	{
	datos.strLimpiar.value='SI';
	datos.submit();
	}
}

</script>


















