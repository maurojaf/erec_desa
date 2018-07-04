<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">   

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
	intIdFocoAT=Request("CB_FOCOAT")
	intCodCampana = Request("CB_CAMPANA")
	intEjeAsig = Request("CB_EJECUTIVO") 
	intTramoMonto = Request("CB_TRMONTO")
	intTramoVenc = Request("CB_TRVENC")
	intTramoAsig = Request("CB_TRASIG")
	strListado = Request("strListado")
	intDiaSemanaList = Request("intDiaSemanaList")

	If Trim(Request("strLimpiar")) = "S" Then
		intEjeAsig = ""
		intTramoMonto = ""
		intTramoVenc = ""
		intTramoAsig = ""
	End If
	
	If intIdFocoAT = "" Then intIdFocoAT = 0
	If intCodCampana = "" Then intCodCampana = 0
	If intDiaSemanaList ="" Then intDiaSemanaList = 100
	
	If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then
		perfilEjecutivo="0"
	Else
		intEjeAsig=session("session_idusuario")
	End If
	
	'Response.write "strListado=" & strListado
	'Response.write "intEjeAsig=" & intEjeAsig
%>

<title>AGENDA</title>

<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
<link href="../css/style_generales_sistema.css" rel="stylesheet">

<script language='javascript' src="../javascripts/popcalendar.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
<script src="../Componentes/jquery.tablesorter/jquery.tablesorter.js"></script>

<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 

<script language="JavaScript1.2">
$(document).ready(function(){
	$.prettyLoader();	
})
    
function buscar(){
	$.prettyLoader.show(2000);
	datos.action='Agenda.asp?';
	datos.submit();
}
</script>

</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">

<div class="titulo_informe">AGENDA</div>
<br>
	<table width="90%" align="CENTER" border="0" bgcolor="#f6f6f6" class="estilo_columnas">
		<thead>
			<tr>
				<td align="left">FOCO</td>
				<td align="left">CAMPAÑA</td>
				<td align="left">TRAMO VENCIMIENTO</td>
				<td align="left">TRAMO MONTO</td>
				<td align="left">TRAMO ASIGNACIÓN</td>
				<td align="left" colspan=0>EJECUTIVO</td>				
			</tr>
		</thead>
			<tr>
				<td>
					<select name="CB_FOCOAT" id="CB_FOCOAT" onChange="buscar()">
						<option value=0>TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_FOCO,NOMBRE_FOCO FROM FOCOS WHERE TIPO_FOCO=1 ORDER BY ID_FOCO ASC"
							set rsFocos=Conn.execute(strSql)
							Do While not rsFocos.eof
								If Trim(intIdFocoAT)=Trim(rsFocos("ID_FOCO")) Then strSelFoco = "SELECTED" Else strSelFoco = ""
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
						<option value=100 <%if Trim(intIdFocoAT)=100 then response.Write("Selected") end if%>>SIN FOCO</option>
					</select>
				</td>
				<td>
					<select name="CB_CAMPANA" id="CB_CAMPANA" onChange="buscar()">
						<option value=0>TODAS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_CAMPANA,NOMBRE FROM CAMPANA WHERE COD_CLIENTE IN ('" & intCodCliente & "')"
							set rsCampana=Conn.execute(strSql)
							Do While not rsCampana.eof
								If Trim(intCodCampana)=Trim(rsCampana("ID_CAMPANA")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsCampana("ID_CAMPANA")%>" <%=strSelCam%>> <%=rsCampana("NOMBRE")%></option>
								<%
								rsCampana.movenext
							Loop
							rsCampana.close
							set rsCampana=nothing
						CerrarSCG()
						''Response.End
						%>
						<option value=1 <%if Trim(intCodCampana)=1 then response.Write("Selected") end if%>>SIN CAMPAÑA</option>
					</select>
				</td>
				<td>
					<select name="CB_TRVENC" id="CB_TRVENC" onChange="buscar();">
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_SVENC = ID,NOMBRE_SEGMENTO=LTRIM(NOMBRE) FROM SEGMENTACION_VENCIMIENTO WHERE COD_CLIENTE = '" & intCodCliente & "' AND GESTIONABLE=1 ORDER BY ORDEN ASC"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(intTramoVenc)=Trim(rsSel("ID_SVENC")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("ID_SVENC")%>" <%=strSelCam%>> <%=rsSel("NOMBRE_SEGMENTO")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>				
				<td>
					<select name="CB_TRMONTO" id="CB_TRMONTO" onChange="buscar();">
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_SMONTO = ID,NOMBRE_SEGMENTO=LTRIM(NOMBRE) FROM SEGMENTACION_MONTO WHERE COD_CLIENTE = ('" & intCodCliente & "') AND GESTIONABLE=1 ORDER BY ORDEN ASC"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(intTramoMonto)=Trim(rsSel("ID_SMONTO")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("ID_SMONTO")%>" <%=strSelCam%>> <%=rsSel("NOMBRE_SEGMENTO")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
						CerrarSCG()
						''Response.End
						%>
					</select>
				</td>
				<td>
					<select name="CB_TRASIG" id="CB_TRASIG" onChange="buscar();">
						<option value="">TODOS</option>
						<%
						AbrirSCG()
							strSql="SELECT ID_SASIG = ID,NOMBRE_SEGMENTO=LTRIM(NOMBRE) FROM SEGMENTACION_ASIGNACION ORDER BY ORDEN ASC"
							set rsSel=Conn.execute(strSql)
							Do While not rsSel.eof
								If Trim(intTramoAsig)=Trim(rsSel("ID_SASIG")) Then strSelCam = "SELECTED" Else strSelCam = ""
								%>
								<option value="<%=rsSel("ID_SASIG")%>" <%=strSelCam%>> <%=rsSel("NOMBRE_SEGMENTO")%></option>
								<%
								rsSel.movenext
							Loop
							rsSel.close
							set rsSel=nothing
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
			<% Else%>
				<td>&nbsp;</td>  
			<% End If %>
			
			</tr>
			
	</table>	
	<br/>
	<table width="90%" align="CENTER" border="1" bgcolor="#f6f6f6" class="estilo_columnas tablesorter intercalado">	

		<%
			AbrirSCG1()
			ssql="Exec uspAgendaSelect '"&TRIM(intCodCliente)&"'," & intIdFocoAT &","& intCodCampana &",'"&intTramoMonto&"','"&intTramoVenc&"','"&intTramoAsig&"','"&intEjeAsig&"'"
			
			'Response.write "ssql=" & ssql
			
			intTReg=0
			intIdUsuario = 0
			intTotalGeneralRutTT = 0
			intTotalGeneralDocTT = 0
			intTotalGeneralMontoTT = 0
			intTotalGeneralRutTE = 0
			intTotalGeneralDocTE = 0
			intTotalGeneralMontoTE = 0
			intTotalGeneralRutSC = 0
			intTotalGeneralDocSC = 0
			intTotalGeneralMontoSC = 0

			set rsCam=Conn1.execute(ssql)
			if not rsCam.eof then

			%>
			<thead>
				<tr>
					<td Rowspan="2" align="center">EJECUTIVO</td>
					<td Rowspan="2" align="center">AGENDAMIENTO</td>
					<td colspan="3" align="center">TIRULAR</td>
					<td colspan="3" align="center">TERCERO</td>
					<td colspan="3" align="center">SIN CONTACTO</td>
					<td colspan="3" align="center">TOTAL</td>
				</tr>
				<tr>
					<td align="center">CASOS</td>
					<td align="center">MONTO</td>
					<td align="center">DOC</td>
					<td align="center">CASOS</td>
					<td align="center">MONTO</td>
					<td align="center">DOC</td>
					<td align="center">CASOS</td>
					<td align="center">MONTO</td>
					<td align="center">DOC</td>
					<td align="center">CASOS</td>
					<td align="center">MONTO</td>
					<td align="center">DOC</td>
				</tr>
			</thead>
			<tbody>
			<%
			Do While Not rsCam.Eof
			
				intTReg = intTReg + 1
				strUsuarioAsig = rsCam("USUARIO_ASIG")	
				intEjeAsig = rsCam("ID_USUARIO")	
				intIdRegistro = rsCam("ID_TOTAL_REG_USU")
				intTipoContacto = rsCam("ID_TIPO_ULT_GEST")
				
				intDiaSemana = rsCam("DIA_SEMANA")
				if intDiaSemana=0 then strNomDiaSemana = "Vencido"
				if intDiaSemana=1 then strNomDiaSemana = "Lunes"
				if intDiaSemana=2 then strNomDiaSemana = "Martes"
				if intDiaSemana=3 then strNomDiaSemana = "Miercoles"
				if intDiaSemana=4 then strNomDiaSemana = "Jueves"
				if intDiaSemana=5 then strNomDiaSemana = "Viernes"
				if intDiaSemana=6 then strNomDiaSemana = "Sabado"
				if intDiaSemana=7 then strNomDiaSemana = "Domingo"
				if intDiaSemana=8 then strNomDiaSemana = "Futuro"
				
				intTotalRutTT = rsCam("TOTAL_RUT_TT")
				intTotalDocTT = rsCam("TOTAL_DOCUMENTOS_TT")
				intTotalMontoTT = rsCam("TOTAL_MONTO_TT")
				intTotalRutTE = rsCam("TOTAL_RUT_TE")
				intTotalDocTE = rsCam("TOTAL_DOCUMENTOS_TE")
				intTotalMontoTE = rsCam("TOTAL_MONTO_TE")
				intTotalRutSC = rsCam("TOTAL_RUT_SC")
				intTotalDocSC = rsCam("TOTAL_DOCUMENTOS_SC")
				intTotalMontoSC = rsCam("TOTAL_MONTO_SC")	

				intTotalGeneralEjecDSRut = intTotalRutTT + intTotalRutTE + intTotalRutSC
				intTotalGeneralEjecDSMonto = intTotalMontoTT + intTotalMontoTE + intTotalMontoSC
				intTotalGeneralEjecDSDoc = intTotalDocTT + intTotalDocTE + intTotalDocSC
				
				intTotalGeneralRutTT = intTotalGeneralRutTT + intTotalRutTT
				intTotalGeneralDocTT = intTotalGeneralDocTT + intTotalDocTT
				intTotalGeneralMontoTT = intTotalGeneralMontoTT + intTotalMontoTT
				
				intTotalGeneralRutTE = intTotalGeneralRutTE + intTotalRutTE
				intTotalGeneralDocTE = intTotalGeneralDocTE + intTotalDocTE
				intTotalGeneralMontoTE = intTotalGeneralMontoTE + intTotalMontoTE

				intTotalGeneralRutSC = intTotalGeneralRutSC + intTotalRutSC
				intTotalGeneralDocSC = intTotalGeneralDocSC + intTotalDocSC
				intTotalGeneralMontoSC = intTotalGeneralMontoSC + intTotalMontoSC
				
				intTotalGeneralRut = intTotalGeneralRutTT + intTotalGeneralRutTE + intTotalGeneralRutSC
				intTotalGeneralMonto = intTotalGeneralDocTT + intTotalGeneralDocTE + intTotalGeneralDocSC
				intTotalGeneralDoc = intTotalGeneralMontoTT + intTotalGeneralMontoTE + intTotalGeneralMontoSC
				
			%>
			<tr>
			<%
				if intIdUsuario <> rsCam("ID_USUARIO") then
				%>
					<td Rowspan= <%=intIdRegistro%> align="left" width="12%"><%=strUsuarioAsig%></td>	
				<%				
				end if
			%>
				<td align="center" width="8%"><%=strNomDiaSemana%></td>
				<td align="right" width="6%">
					<A HREF="Agenda.asp?strListado=S&intDiaSemanaList=<%=intDiaSemana%>&CB_EJEC_LISTADO=<%=intEjeAsig%>&intTipoContacto=1&CB_FOCOAT=<%=intIdFocoAT%>&CB_CAMPANA=<%=intCodCampana%>&CB_EJECUTIVO=<%=intEjeAsig%>&CB_TRMONTO=<%=intTramoMonto%>&CB_TRVENC=<%=intTramoVenc%>&CB_TRASIG=<%=intTramoAsig%>">
					<%=FN(intTotalRutTT,0)%>
				</td>
				<td align="right" width="6%"><%=FN(intTotalMontoTT,0)%></td>
				<td align="right" width="6%"><%=FN(intTotalDocTT,0)%></td>
				<td align="right" width="6%">
					<A HREF="Agenda.asp?strListado=S&intDiaSemanaList=<%=intDiaSemana%>&CB_EJEC_LISTADO=<%=intEjeAsig%>&intTipoContacto=2&CB_FOCOAT=<%=intIdFocoAT%>&CB_CAMPANA=<%=intCodCampana%>&CB_EJECUTIVO=<%=intEjeAsig%>&CB_TRMONTO=<%=intTramoMonto%>&CB_TRVENC=<%=intTramoVenc%>&CB_TRASIG=<%=intTramoAsig%>">
					<%=FN(intTotalRutTE,0)%>
				</td>
				<td align="right" width="6%"><%=FN(intTotalMontoTE,0)%></td>
				<td align="right" width="6%"><%=FN(intTotalDocTE,0)%></td>
				<td align="right" width="6%">
					<A HREF="Agenda.asp?strListado=S&intDiaSemanaList=<%=intDiaSemana%>&CB_EJEC_LISTADO=<%=intEjeAsig%>&intTipoContacto=3&CB_FOCOAT=<%=intIdFocoAT%>&CB_CAMPANA=<%=intCodCampana%>&CB_EJECUTIVO=<%=intEjeAsig%>&CB_TRMONTO=<%=intTramoMonto%>&CB_TRVENC=<%=intTramoVenc%>&CB_TRASIG=<%=intTramoAsig%>">
					<%=FN(intTotalRutSC,0)%>
				</td>
				<td align="right" width="6%"><%=FN(intTotalMontoSC,0)%></td>
				<td align="right" width="6%"><%=FN(intTotalDocSC,0)%></td>	
				<td align="right" width="6%">
					<A HREF="Agenda.asp?strListado=S&intDiaSemanaList=<%=intDiaSemana%>&CB_EJEC_LISTADO=<%=intEjeAsig%>&intTipoContacto=0&CB_FOCOAT=<%=intIdFocoAT%>&CB_CAMPANA=<%=intCodCampana%>&CB_EJECUTIVO=<%=intEjeAsig%>&CB_TRMONTO=<%=intTramoMonto%>&CB_TRVENC=<%=intTramoVenc%>&CB_TRASIG=<%=intTramoAsig%>">
					<%=FN(intTotalGeneralEjecDSRut,0)%>
				</td>				
				<td align="right" width="6%"><%=FN(intTotalGeneralEjecDSMonto,0)%></td>	
				<td align="right" width="6%"><%=FN(intTotalGeneralEjecDSDoc,0)%></td>				
			</tr>
			<%
	
				intIdUsuario = rsCam("ID_USUARIO")
				
			rsCam.movenext
			Loop
			end if
			rsCam.close
			set rsCam=nothing
			CerrarSCG1()
		
		If intTReg=0 then				
			
			%>
			<tr class="totales">
				<td Colspan = "15" >&nbsp;</td>
			</tr>
			
			<tr bgcolor="<%=strbgcolor%>" class="Estilo8">																					
				<td colspan="15" align = "center"><h3>No Existen Casos Casos Agendados Según los parámetros de búsqueda</h3></td>	
			</tr>
			
			<tr class="totales">
				<td Colspan = "15" >&nbsp;</td>
			</tr>
			
		<%Else%>
			<thead>	
				<tr class="totales">
					<td align="right" Colspan = "2">&nbsp;</td>
					<td align="right"><%=FN(intTotalGeneralRutTT,0)%></td>
					<td align="right"><%=FN(intTotalGeneralMontoTT,0)%></td>
					<td align="right"><%=FN(intTotalGeneralDocTT,0)%></td>
					
					<td align="right"><%=FN(intTotalGeneralRutTE,0)%></td>
					<td align="right"><%=FN(intTotalGeneralMontoTE,0)%></td>
					<td align="right"><%=FN(intTotalGeneralDocTE,0)%></td>

					<td align="right"><%=FN(intTotalGeneralRutSC,0)%></td>
					<td align="right"><%=FN(intTotalGeneralMontoSC,0)%></td>
					<td align="right"><%=FN(intTotalGeneralDocSC,0)%></td>

					<td align="right"><%=FN(intTotalGeneralRut,0)%></td>
					<td align="right"><%=FN(intTotalGeneralMonto,0)%></td>
					<td align="right"><%=FN(intTotalGeneralDoc,0)%></td>
				</tr>
					
		<%end if%>
			</thead>
		</tbody>
	</table>
	
	<%If strListado="S" and intTReg>0 Then%>
	
	<br>
	<table width="90%" align="CENTER" border="1" bgcolor="#f6f6f6" class="estilo_columnas tablesorter intercalado">
		<thead>
			<tr class="Estilo34">
				<th>&nbsp;</th>
				<th id="rut" align="left">RUT</th>
				<th>NOMBRE O RAZON SOCIAL </th>
				<th id="SALDO" align="left">SALDO</th>
				<th align="left">DOC.</th>
				<th align="center">DIA MORA</th>
				<th align="center">F.AGEND.</th>
				<th align="center">H.AGEND. </th>
				<th align="center">EJECUTIVO</th>
			</tr>
		</thead>
		<tbody>		
		<%
			AbrirSCG1()
			ssql="EXEC uspAgendaListadoSelect '"&TRIM(intCodCliente)&"'," & intIdFocoAT &","& intCodCampana &",'"&intTramoMonto&"','"&intTramoVenc&"','"&intTramoAsig&"','"&Request("CB_EJEC_LISTADO")&"',"&intDiaSemanaList&","&Request("intTipoContacto")
			
			'Response.write "ssql=" & ssql
			
			intNumReg=0

			set rsCam=Conn1.execute(ssql)
			if not rsCam.eof then
			
			Do While Not rsCam.Eof

				intNumReg= intNumReg + 1
				strRutDeudor = rsCam("RUT_DEUDOR")
				strNombreDeudor = rsCam("NOMBRE_DEUDOR")
				intSaldo = rsCam("MONTO_AGEND")
				intTotalDoc = rsCam("TOTAL_DOC_AGEND")
				intDiaMora = rsCam("DM")
				dtmFecAgend = rsCam("MIN_FEC_AGEND")
				strHoraAgend = rsCam("MIN_HORA_AGEND")
				strUsuarioAsig = rsCam("USUARIO_ASIG")	

				if strHoraAgend = "00:00" or strHoraAgend = "08:59" then strHoraAgend = "" 
				
		%>
			<tr>
				<td><%=FN(intNumReg,0)%></td>
				<td ALIGN="center"><%=strRutDeudor%></td>
				<td><%=Mid(strNombreDeudor,1,30)%></td>
				<td><%=FN(intSaldo,0)%></td>				
				<td><%=FN(intTotalDoc,0)%></td>
				<td ALIGN="center"><%=intDiaMora%></td>
				<td ALIGN="center"><%=dtmFecAgend%></td>
				<td ALIGN="center"><%=strHoraAgend%></td>	
				<td ALIGN="center"><%=strUsuarioAsig%></td>				
			</tr>				
		<%	
			rsCam.movenext
			Loop
			end if
			rsCam.close
			set rsCam=nothing
			CerrarSCG1()			
		%>
			<tr class="totales">
				<td Colspan = "15" >&nbsp;</td>
			</tr>

		</tbody>
	</table>	
	
	<%End If%>	
	
</form>
</body>
</html>
