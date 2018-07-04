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
	
	strCodCliente   = request("strCodCliente")
	
	intIdCuota = request("intID_CUOTA")
	
	'Response.write "<br>strCodCliente=" & strCodCliente

	abrirscg()

		strSql = "SELECT C.NRO_DOC,C.NRO_CUOTA,VENC_CUOTA = CONVERT(VARCHAR(10),C.FECHA_VENC,103),DM = DATEDIFF(DAY,C.FECHA_VENC,GETDATE())"
		strSql = strSql & " FROM CUOTA C"
		strSql = strSql & " WHERE C.ID_CUOTA = " & intIdCuota
	
		set rsGestCuota=Conn.execute(strSql)
		If not rsGestCuota.eof then
			
		intNroDoc = rsGestCuota("NRO_DOC")	
		intNroCuota = rsGestCuota("NRO_CUOTA")
		strFechaVenc = rsGestCuota("VENC_CUOTA")
		intDiaMora = rsGestCuota("DM")
			
		End if
		rsGestCuota.close
		set rsGestCuota=nothing
			
	cerrarscg()

%>


</head>
<body>

		<div class="titulo_informe">Detalle Gestiones Por Documento
 	
		</div>
		<br>
			<table width="100%" border="0" bordercolor="#FFFFFF"  class="intercalado" >
			<tr>
			<td style="font-size:13px"><b>Documento <%=intNroDoc%> - Cuota <%=intNroCuota%> - Vencimiento <%=strFechaVenc%> - Día Mora (<%=intDiaMora%>)</td>
			<td  align="center"> <a href="#" onclick="exportar();return false">
			<img src="../Imagenes/Excel.gif"  >
			</a>
			</td>
			</tr>
			</table>
			<table width="100%" border="0" bordercolor="#FFFFFF" class="intercalado" align="center">
			<thead>
	        <tr >
				<td width = "70" class="Estilo4">F.ING.</td>
				<td class="Estilo4">H.ING.</td>
				<td width = "30"class="Estilo4">DMI</td>
				<td class="Estilo4">GESTION</td>
				<td class="Estilo4">F.COMP</td>
				<td class="Estilo4">&nbsp;</td>
				<td width = "65" class="Estilo4">F.AGEND</td>
				<td class="Estilo4">H.AGEND</td>
				<td class="Estilo4">OBS</td>
				<td class="Estilo4">&nbsp;</td>
				<td width = "65" class="Estilo4">MEDIO</td>
				<td class="Estilo4">EJECUTIVO</td>
	        </tr>
	    	</thead>
	    	<tbody>
			
	     <%

				strSql = "SELECT G.COD_CATEGORIA,  G.COD_SUB_CATEGORIA, G.COD_GESTION, TC.CONTACTO, G.OBSERVACIONES,GC.ID_CUOTA,G.ID_GESTION, "
				strSql = strSql & " G.FECHA_INGRESO, CONVERT(VARCHAR(5),convert(datetime, G.HORA_INGRESO), 108) HORA_INGRESO, G.FECHA_RETIRO,G.HORA_DESDE, G.HORA_HASTA, CONVERT(VARCHAR(10),G.FECHA_COMPROMISO,103) AS FECHA_COMPROMISO,"
				strSql = strSql & " FECHA_AGENDAMIENTO = CONVERT(VARCHAR(10),G.FECHA_AGENDAMIENTO,103), G.HORA_AGENDAMIENTO, DMI = DATEDIFF(DAY,C.FECHA_VENC,G.FECHA_INGRESO),C.NRO_CUOTA,VENC_CUOTA = CONVERT(VARCHAR(10),C.FECHA_VENC,103),DM = DATEDIFF(DAY,C.FECHA_VENC,GETDATE()),C.NRO_DOC,"

				strSql = strSql & " case "
				strSql = strSql & "	when G.TIPO_MEDIO_GESTION = 1 then "
				strSql = strSql & "		(SELECT UPPER(CONTACTO) AS CONTACTO FROM TELEFONO_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION) "
				strSql = strSql & "	when G.TIPO_MEDIO_GESTION = 2 then "
				strSql = strSql & "		(SELECT UPPER(CONTACTO) AS CONTACTO FROM EMAIL_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION) "
				strSql = strSql & "	when G.TIPO_MEDIO_GESTION = 3 then "
				strSql = strSql & "		(SELECT UPPER(CONTACTO) AS CONTACTO FROM DIRECCION_CONTACTO WHERE ID_CONTACTO = G.ID_CONTACTO_GESTION)  "
				strSql = strSql & " else '' end NOM_CONTACTO_GESTION "

				strSql = strSql & " , G.ID_USUARIO, DE.EMAIL EMAIL_ASOCIADO, case when DT.COD_AREA = 0 then DT.TELEFONO ELSE convert(varchar,DT.COD_AREA) +'-'+ DT.TELEFONO end TELEFONO_ASOCIADO, DD.CALLE + ' ' + DD.NUMERO  + ' ' + DD.RESTO + ' ' + DD.COMUNA DIRECCION_ASOCIADA "
                strSql = strSql & " ,DATENAME(weekday,g.FECHA_INGRESO) + ', ' + SUBSTRING(convert(varchar,CONVERT(datetime,g.FECHA_INGRESO),105),1,2) + ' de ' + DATENAME(month,g.FECHA_INGRESO)   Fecha_Tootltip "
				strSql = strSql & " FROM GESTIONES G INNER JOIN GESTIONES_CUOTA GC ON G.ID_GESTION = GC.ID_GESTION AND GC.ID_CUOTA = " & intIdCuota 
				strSql = strSql & " 				 INNER JOIN CUOTA C ON GC.ID_CUOTA = C.ID_CUOTA"
				strSql = strSql & " 				 LEFT JOIN TELEFONO_CONTACTO TC ON G.ID_CONTACTO_GESTION = TC.ID_CONTACTO"
				strSql = strSql & " 				 LEFT JOIN DEUDOR_TELEFONO DT ON G.ID_MEDIO_GESTION = DT.ID_TELEFONO"
				strSql = strSql & "					 LEFT JOIN DEUDOR_EMAIL DE ON DE.ID_EMAIL = G.ID_MEDIO_GESTION"
				strSql = strSql & "					 LEFT JOIN DEUDOR_DIRECCION DD ON DD.ID_DIRECCION = G.ID_MEDIO_GESTION"
				

				strSql = strSql & " WHERE G.COD_CLIENTE='" & strCodCliente & "'"

                strSql = strSql & " ORDER BY CAST(CONVERT(VARCHAR(10),G.FECHA_INGRESO,103)+' '+CONVERT(VARCHAR(5),convert(datetime, G.HORA_INGRESO), 108) AS DATETIME) DESC , G.ID_GESTION DESC"
				'Response.write "<br>Sql=" & strSql
				'Response.end
				AbrirSCG1()
				set rsGestCuota= Conn1.execute(strSql)
				
				intNumReg = 0
				
				Do until rsGestCuota.eof
					Obs=rsGestCuota("OBSERVACIONES")
					strContacto=rsGestCuota("CONTACTO")
					intDMI = rsGestCuota("DMI") 
					
					intNumReg = intNumReg + 1
					
					If Obs="" then
						Obs="SIN INFORMACION ADICIONAL"
					End if
					
					if trim(rsGestCuota("NOM_CONTACTO_GESTION")) <> "" AND NOT ISNULL(rsGestCuota("NOM_CONTACTO_GESTION")) then
						NOM_CONTACTO_GESTION =trim(rsGestCuota("NOM_CONTACTO_GESTION"))
					else
						NOM_CONTACTO_GESTION ="SIN CONTACTO ASOCIADO"
					end if
					
					strContacto =rsGestCuota("NOM_CONTACTO_GESTION")

					strCodGestion= rsGestCuota("COD_CATEGORIA")& "-" & rsGestCuota("COD_SUB_CATEGORIA")& "-" &rsGestCuota("COD_GESTION")

					AbrirSCG2()
					strSql = "SELECT G.COD_CATEGORIA, G.COD_SUB_CATEGORIA, G.COD_GESTION,C.DESCRIPCION + '-' + S.DESCRIPCION + '-' +  G.DESCRIPCION as DESCRIP,"
					strSql = strSql + " ISNULL(G.COMUNICA,2) AS COMUNICA, G.GESTION_MODULOS, "
					strSql = strSql + " ISNULL(G.PRIORIDAD_GTEL,0) AS PRIORIDAD_GTEL, ISNULL(G.PRIORIDAD_GMAIL,0) AS PRIORIDAD_GMAIL, ISNULL(G.PRIORIDAD_GDIR,0) AS PRIORIDAD_GDIR "
					strSql = strSql & " FROM GESTIONES_TIPO_CATEGORIA C, GESTIONES_TIPO_SUBCATEGORIA S, GESTIONES_TIPO_GESTION G"
					strSql = strSql & " WHERE C.COD_CATEGORIA = S.COD_CATEGORIA"
					strSql = strSql & " AND C.COD_CATEGORIA = G.COD_CATEGORIA"
					strSql = strSql & " AND S.COD_SUB_CATEGORIA = G.COD_SUB_CATEGORIA"
					strSql = strSql & " AND G.COD_CLIENTE='" & strCodCliente & "'"
					strSql = strSql & " AND CAST(G.COD_CATEGORIA AS VARCHAR(2)) + '-' + CAST(G.COD_SUB_CATEGORIA AS VARCHAR(2)) + '-' + CAST(G.COD_GESTION AS VARCHAR(2)) = '" & Trim(strCodGestion) & "'"

					'Response.write "strSql=" &strSql
					'Response.End
					SET rsNomGestion1=Conn2.execute(strSql)
					If Not rsNomGestion1.Eof Then
						strGestion = rsNomGestion1("DESCRIP")

						intGestionComunica = rsNomGestion1("COMUNICA")
						intGestionGtel = rsNomGestion1("PRIORIDAD_GTEL")
						intGestionGmail = rsNomGestion1("PRIORIDAD_GMAIL")
						intGestionGdir = rsNomGestion1("PRIORIDAD_GDIR")
						strTipoGestion = rsNomGestion1("GESTION_MODULOS")

					Else
						strGestion = ""
					End If
					CerrarSCG2()

					AbrirSCG2()
					If Trim(rsGestCuota("ID_USUARIO")) <> "" Then
						strLoginCobradorGest = TraeCampoId(Conn2, "LOGIN", rsGestCuota("ID_USUARIO"), "USUARIO", "ID_USUARIO")
					End If
					CerrarSCG2()
					
					

	     %>

				<tr bordercolor="#FFFFFF">
					<td class="Estilo4" title="<%=rsGestCuota("Fecha_Tootltip")%>"><%=rsGestCuota("FECHA_INGRESO")%></td>
					<td align="center" class="Estilo4"><%=rsGestCuota("HORA_INGRESO")%></td>
					<td align="center" class="Estilo4"><%=intDMI%></td>
					<td class="Estilo4"><%=strGestion%></td>
					<td class="Estilo4"><%=rsGestCuota("FECHA_COMPROMISO")%></td>
					<td class="Estilo4"><%=strConfirmaCP%></td>

					<td class="Estilo4"><%=rsGestCuota("FECHA_AGENDAMIENTO")%></td>
					<td align="center" class="Estilo4"><%=rsGestCuota("HORA_AGENDAMIENTO")%></td>
					<td class="Estilo4" title="<%=Obs%>">
						<%=Mid(Obs,1,50)%>
					</td>

					<td
						<%If intGestionComunica = 0 AND intGestionGmail = 1 Then%>
						  class="Estilo4" title="<%=strContacto%>">
						  <img src="../imagenes/Contacto.rojo.png" border="0" title="<%=NOM_CONTACTO_GESTION%>">

						  <%ElseIf intGestionComunica = 0 AND intGestionGtel = 1 Then%>
						  class="Estilo4" title="<%=strContacto%>">
						  <img src="../imagenes/Contacto.rojo.png" border="0" title="<%=NOM_CONTACTO_GESTION%>">
						  
						  <%ElseIf intGestionComunica = 0 AND intGestionGdir = 1 Then%>
						  class="Estilo4" title="<%=strContacto%>">
						  <img src="../imagenes/Contacto.rojo.png" border="0" title="<%=NOM_CONTACTO_GESTION%>">

						  <%ElseIf intGestionComunica = 1 AND intGestionGtel = 1 Then%>
						  class="Estilo4" title="<%=strContacto%>">
						  <img src="../imagenes/Contacto.azul.png" border="0" title="<%=NOM_CONTACTO_GESTION%>">

						  <%ElseIf intGestionComunica = 1 AND intGestionGmail = 1 Then%>
						  class="Estilo4" title="<%=strContacto%>">
						  <img src="../imagenes/Contacto.azul.png" border="0" title="<%=NOM_CONTACTO_GESTION%>">
						  
						  <%ElseIf intGestionComunica = 1 AND intGestionGdir = 1 Then%>
						  class="Estilo4" title="<%=strContacto%>">
						  <img src="../imagenes/Contacto.azul.png" border="0" title="<%=NOM_CONTACTO_GESTION%>">

						  <%Else%>
						   &nbsp
						<%End If%>
					</td>

					<td
					  <%If intGestionGmail = 1 Then%>
					  align= "center" class="Estilo4" title="<%=rsGestCuota("EMAIL_ASOCIADO")%>">
					  <img src="../imagenes/Arroa.png" border="0">

					  <%ElseIf intGestionGtel = 1 Then%>
					  <td class="Estilo4"><%=rsGestCuota("TELEFONO_ASOCIADO")%></td>
					  
					  <%ElseIf intGestionGdir = 1 Then%>
					  <td align= "center" class="Estilo4" title="<%=rsGestCuota("DIRECCION_ASOCIADA")%>">
					  <img src="../imagenes/mod_direccion_va.png" border="0">

					  <%Else%>
					   &nbsp;
					  <%End If%>
					</td>
					
					<td class="Estilo4"><%=UCASE(strLoginCobradorGest)%></td>
				</tr>
	      <%
	      		Response.Flush
	      		rsGestCuota.movenext
			Loop
				CerrarSCG1()

			If intNumReg=0 then	%>
					
			<tr bgcolor="<%=strbgcolor%>" class="Estilo8">																					
				<td colspan="12" align = "center"><h3>No Existen Gestiones Asociadas al Documento</h3></td>	
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
	
	function exportar()
{
	var pagina = 'gestiones_por_documento_Excel.asp?intID_CUOTA=<%=intIdCuota%>&strNroDoc=<%=strNroDoc%>'
	//alert(pagina);
	window.open(pagina,'window','params');

	
}
</script>
