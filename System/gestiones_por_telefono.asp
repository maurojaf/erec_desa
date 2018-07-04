<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

	<!--#include file="arch_utils.asp"-->
	<!--#include file="sesion_inicio.asp"-->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/lib.asp"-->
	
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	strRutDeudor = request("strRutDeudor")
	intIdFono = request("intIdFono")
	strFonoAgestionar = request("strFonoAgestionar")
	
	'response.Write("strRutDeudor")
%>


</head>
<body>

		<div class="titulo_informe">Detalle Gestiones Por Teléfono
 	
		</div>
		<br>
			<table width="100%" border="0" bordercolor="#FFFFFF"  class="intercalado" >
			<tr>
			<td style="font-size:13px"><b>Telefono: <%=strFonoAgestionar%></td>
			</tr>
			</table>
			<table width="100%" border="0" bordercolor="#FFFFFF" class="intercalado" align="center">
			<thead>
	        <tr >
				<td class="Estilo4">NOM. CLIENTE</td>
				<td class="Estilo4">&nbsp;</td>
				<td class="Estilo4">F.INGRESO.</td>
				<td class="Estilo4">H.INGRESO</td>
				<td class="Estilo4">GESTION</td>
				<td class="Estilo4">OBS</td>
				<td class="Estilo4">EJECUTIVO</td>
	        </tr>
	    	</thead>
	    	<tbody>
			
	     <%
			AbrirSCG1()
			
				strSql = " EXEC proc_Inf_Gestiones_por_fono '"&TRIM(strRutDeudor)&"'," & intIdFono
				
				set rsGestTel= Conn1.execute(strSql)
				
				intNumReg = 0
				
				Do until rsGestTel.eof
				
				intNumReg = intNumReg + 1
				
				Obs			= rsGestTel("OBSERVACIONES")
				intTipoGest = rsGestTel("TIPO_GESTION")

				if intTipoGest = 1 Then
					strContactado = "tel_contactado.jpg"
				Elseif intTipoGest = 2 then
					strContactado = "tel_contactado_te.jpg"
				Else
					strContactado = "tel_nocontactado.jpg"
				End If
					
	     %>

				<tr bordercolor="#FFFFFF">
				
					<td class="Estilo4"><%=rsGestTel("NOM_CLIENTE")%></td>
					<td Align="center"><img src="../imagenes/<%=strContactado%>" border="0"></td>
					<td class="Estilo4"><%=rsGestTel("FECHA_GESTION")%></td>
					<td class="Estilo4"><%=rsGestTel("HORA_GESTION")%></td>
					<td class="Estilo4"><%=rsGestTel("GESTION")%></td>
					<td class="Estilo4" title="<%=Obs%>">
						<%=Mid(Obs,1,50)%>
					</td>
					<td class="Estilo4"><%=rsGestTel("USUARIO_GESTION")%></td>
				</tr>
	      <%
	      		Response.Flush
	      		rsGestTel.movenext
			Loop
			
			CerrarSCG1()

			If intNumReg=0 then	%>
					
			<tr bgcolor="<%=strbgcolor%>" class="Estilo8">																					
				<td colspan="12" align = "center"><h3>No Existen Gestiones Asociadas al Teléfono</h3></td>	
			</tr>
			
			<%end if%>


	      	</tbody>
	        </table>
</body>
</html>


<script type="text/javascript">
	$(document).ready(function(){
		$(document).tooltip();
	})
</script>
