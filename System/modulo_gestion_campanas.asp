<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib2.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/Minimo.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->
	
<%
	Response.CodePage 	=65001
	Response.charset 	="utf-8"
	intCodCliente=session("ses_codcli")
	intIdFoco=Request("CB_FOCO")
	intCodCampana = Request("CB_CAMPANA")
	intEjeAsig = Request("CB_EJECUTIVO") 
	
	If Trim(Request("strBuscar")) = "S" Then
		session("Ftro_FOCOGF") = intIdFoco
		session("Ftro_Campana") = intCodCampana
		session("Ftro_Ejecutivo") = intEjeAsig
	End If
	
	If intIdFoco = "" Then intIdFoco = session("Ftro_FOCOGF")
	If intCodCampana = "" Then intCodCampana = session("Ftro_Campana")
	If intEjeAsig = "" Then intEjeAsig = session("Ftro_Ejecutivo")
	
	If intIdFoco = "" Then intIdFoco = 0
	If intCodCampana = "" Then intCodCampana = 0

	If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then
		perfilEjecutivo="0"
	Else
		intEjeAsig=session("session_idusuario")
	End If
	
	''Response.write "intEjeAsig=" & intEjeAsig

%>

<title>MODULO GESTIÓN CAMPAÑAS</title>

<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
<link href="../css/style_generales_sistema.css" rel="stylesheet">

<script language='javascript' src="../javascripts/popcalendar.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script language="JavaScript1.2">

$(document).ready(function(){
		
	$.prettyLoader();	
})

function buscar(){
	//alert("hola");
	$.prettyLoader.show(2000);
	datos.action='modulo_gestion_campanas.asp?strBuscar=S';
	datos.submit();
}

</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">

<div class="titulo_informe">MÓDULO DE GESTION DE CAMPAÑAS</div>
<br>
	<table width="90%" align="CENTER" border="0" bgcolor="#f6f6f6" class="estilo_columnas">
		<thead>
			<tr>
				<td align="left">FOCO</td>
				<td align="left">CAMPAÑA</td>
				
				  <% If perfilEjecutivo = "0" Then %>
				  	<td width="150">EJECUTIVO</td>
				  <% End If %>
				  
				<td width="60%">&nbsp;</td>
			</tr>
		</thead>
			<tr>
				<td>
					<select name="CB_FOCO" id="CB_FOCO" onChange="buscar()">
						<option value=0>SELECCIONAR</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_FOCO,NOMBRE_FOCO FROM FOCOS WHERE TIPO_FOCO=1 ORDER BY ID_FOCO ASC"
							set rsFocos=Conn.execute(strSql)
							Do While not rsFocos.eof
								If Trim(intIdFoco)=Trim(rsFocos("ID_FOCO")) Then strSelFoco = "SELECTED" Else strSelFoco = ""
								%>
								<option value="<%=rsFocos("ID_FOCO")%>" <%=strSelFoco%>> <%=rsFocos("NOMBRE_FOCO")%></option>
								<%
								rsFocos.movenext
							Loop
							rsFocos.close
							set rsFocos=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>
				<td>
					<select name="CB_CAMPANA" id="CB_CAMPANA" onChange="buscar()">
						<option value=0>SELECCIONAR</option>
						<%
						AbrirSCG()
							strSql="SELECT * FROM CAMPANA WHERE COD_CLIENTE IN ('" & intCodCliente & "')"
							set rsCampana=Conn.execute(strSql)
							Do While not rsCampana.eof
								If Trim(intCodCampana)=Trim(rsCampana("ID_CAMPANA")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsCampana("ID_CAMPANA")%>" <%=strSelCam%>> <%=rsCampana("ID_CAMPANA") & " - " & rsCampana("NOMBRE")%></option>
								<%
								rsCampana.movenext
							Loop
							rsCampana.close
							set rsCampana=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>
				
				<% If perfilEjecutivo="0" Then %>
				<td>
					<select name="CB_EJECUTIVO" id="CB_EJECUTIVO" onChange="buscar()">
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql= " SELECT DISTINCT U.ID_USUARIO, LOGIN"
							strSql= strSql & " FROM USUARIO U INNER JOIN USUARIO_CLIENTE UC ON U.ID_USUARIO = UC.ID_USUARIO AND UC.COD_CLIENTE IN ('" & intCodCliente & "')"

							strSql= strSql & " WHERE U.ACTIVO = 1 AND U.PERFIL_COB = 1"
							strSql = strSql & " AND U.PERFIL_EMP=0"
							
							set rsEjecutivo=Conn.execute(strSql)
							Do While not rsEjecutivo.eof
								If Trim(intEjeAsig)=Trim(rsEjecutivo("ID_USUARIO")) Then strSelEjecutivo = "SELECTED" Else strSelEjecutivo = ""
								%>
								<option value="<%=rsEjecutivo("ID_USUARIO")%>" <%=strSelEjecutivo%>> <%=rsEjecutivo("LOGIN")%></option>
								<%
								rsEjecutivo.movenext
							Loop
							rsEjecutivo.close
							set rsEjecutivo=nothing
						CerrarSCG()
						
						%>
					</select>
				</td>
				<% End If %>
			</tr>
	</table>
	<br/>
			
	<table  width="90%"  align="CENTER" border="0" bgcolor="#f6f6f6" class="estilo_columnas">
		<thead>
			<tr >
				<td id="rut" align="left">RUT</td>
				<td width="350">NOMBRE O RAZON SOCIAL </td>
				<td id="SALDO" align="left">SALDO</td>
				<td align="left">DOC.</td>
				<td align="center">FECHA VENC.</td>
				<td align="center">F.AGEND.</td>
				<td align="center">H.AGEND. </td>
				<td align="center">RESTANTES</td>
				<td align="center">EJECUTIVO</td>
			</tr>
		</thead>
		<tbody>
		
		<%
			AbrirSCG1()
			ssql="EXEC uspGestionCampanaSelect '"&TRIM(intCodCliente)&"'," & intIdFoco &"," & intCodCampana &",'"&intEjeAsig&"'"
			
			'Response.write "intCodCampana=" & ssql
			
			intTotalRegistros=0

			set rsCam=Conn1.execute(ssql)
			if not rsCam.eof then

				intTotalRegistros=1
				strRutDeudor = rsCam("RUT_DEUDOR")
				strNombreDeudor = rsCam("NOMBRE_DEUDOR")
				intSaldo = rsCam("SALDO")
				intTotalDoc = rsCam("TOTAL_DOC")
				dtmFechaVenc = rsCam("VENC_INFERIOR")
				dtmFecAgend = rsCam("FEC_AGEND")
				strHoraAgend = rsCam("HORA_AGEND")	
				intOrdenFoco = rsCam("ORDEN_DESP")
				strUsuarioAsig = rsCam("USUARIO_ASIG")			
								
			end if
			rsCam.close
			set rsCam=nothing
			CerrarSCG1()
			
			If intTotalRegistros=1 then
				
		%>
			<tr bgcolor="<%=strbgcolor%>" class="Estilo8">
				<td ALIGN="left">
					<a href="detalle_gestiones.asp?rut=<%=strRutDeudor%>&cliente=<%=intCodCliente%>" onclick="javascript:SetCustomer('<%=intCodCliente%>');">
						<acronym title="Llevar a pantalla de ingreso de gestión"><%=strRutDeudor%></acronym>
					</A>
				</td>																					
				<td><%=Mid(strNombreDeudor,1,30)%></td>
				<td><%=FN(intSaldo,0)%></td>				
				<td><%=FN(intTotalDoc,0)%></td>
				<td ALIGN="center"><%=dtmFechaVenc%></td>
				<td ALIGN="center"><%=dtmFecAgend%></td>
				<td ALIGN="center"><%=strHoraAgend%></td>	
				<td ALIGN="center"><%=intOrdenFoco%></td>
				<td ALIGN="center"><%=strUsuarioAsig%></td>
				
			</tr>
			
		<%	Else%>
			
			<tr bgcolor="<%=strbgcolor%>" class="Estilo8">																					
				<td colspan="9" align = "center"><h3>No Existen Casos Pendientes a Gestionar</h3></td>	
			</tr>	
			
		<%	End If%>
							
		</tbody>
			<tr class="totales">
				<td Colspan = "9" >&nbsp;</td>
			</tr>
	</table>
			
</form>
</body>
</html>
