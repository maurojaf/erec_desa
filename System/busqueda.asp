<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
    <link href="../css/style_generales_sistema.css" rel="stylesheet">
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
	<script src="../Componentes/jquery.numeric/jquery.numeric.js"></script>

	<%
	
	strOrigen = request("strOrigen")
	
	If strOrigen = "1" Then %>
		<!--#include file="sesion_inicio.asp"-->
	<% Else %>
		<!--#include file="sesion.asp"-->
	<% End If %>

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<%

	Response.CodePage=65001
	Response.charset ="utf-8"

	strCodCliente = session("ses_codcli")
	TX_RUT_REPLEGAL = TRIM(request("TX_RUT_REPLEGAL"))
	TX_NOMBRE_REPLEGAL = UCASE(TRIM(request("TX_NOMBRE_REPLEGAL")))
	TX_NOMBRE_SUBCLIENTE = UCASE(TRIM(request("TX_NOMBRE_SUBCLIENTE")))
	TX_DOCUMENTO = TRIM(request("TX_DOCUMENTO"))
	strRutSubCliente = TRIM(request("TX_RUTSUBCLIENTE"))
	TX_TELASOCIADO = UCASE(TRIM(request("TX_TELASOCIADO")))
	TX_NRO_CLIENTE = UCASE(TRIM(request("TX_NRO_CLIENTE")))
	TX_MONTOASOCIADO = Replace(TRIM(request("TX_MONTOASOCIADO")), ".", "")
	
	strRutDeudor = Trim(request("TX_RUT_DEUDOR"))
	strNombreDeudor = Trim(request("TX_NOMBRE"))

	If Trim(Request("strLimpiar")) = "S" Then
		 TX_RUT_REPLEGAL = ""
		 TX_NOMBRE_REPLEGAL = ""
		 TX_NOMBRE_SUBCLIENTE = ""
		 TX_DOCUMENTO = ""
		 strRutSubCliente = ""
		 TX_TELASOCIADO = ""
		 TX_NRO_CLIENTE = ""
		 strRutDeudor = ""
		 strNombreDeudor = ""
		 TX_MONTOASOCIADO = ""
	End If

    ssql="SELECT USA_SUBCLIENTE = ISNULL(USA_SUBCLIENTE,0), REP_LEGAL = ISNULL(USA_REPLEGAL,0), NRO_CLIENTE_DEUDOR = ISNULL(CL.NRO_CLIENTE_DEUDOR,0), TIPO_NEGOCIO = ISNULL(CL.TIPO_NEGOCIO,1) FROM CLIENTE CL WHERE CL.COD_CLIENTE = '" & strCodCliente &"'"

	abrirscg()
	set rsCli=Conn.execute(ssql)

	strUsaRep_Legal= rsCli("REP_LEGAL")
	intTipoNegocio= rsCli("TIPO_NEGOCIO")
	strUsaNro_Cliente= rsCli("NRO_CLIENTE_DEUDOR")

	cerrarscg()

	'Response.write "<br>strUsaRep_Legal= " & strUsaRep_Legal
	'Response.write "<br>strUsaNro_Cliente= " & strUsaNro_Cliente

%>
	<title>MÓDULO DE BÚSQUEDA</title>

	<style type="text/css">
	<!--
	.Estilo37 {color: #FFFFFF}
	-->
	</style>
</head>
<body>
		<div class="titulo_informe">
			BÚSQUEDA DEL DEUDOR
		</div>

<br>
<br>
<table width="90%" align="CENTER" height="420" border="0">
  <tr>
    <td height="242" valign="top">
    <form name="datos" method="post">
	
	<% If intTipoNegocio = "1" Then%>

	<table width="100%" border="0" class="estilo_columnas">
		<thead>
			<tr >
				<td>RUT DEUDOR</td>
				<td >NOMBRE DEUDOR</td>
				<td>RUT GIRADOR</td>
				<td>NOMBRE GIRADOR</td>
				<% If strUsaNro_Cliente <> "0" Then%>
				<td><%=strUsaNro_Cliente%></td>
				<%Else%>
				<td>&nbsp;</td>
				<% End If %>
			</tr>
		</thead>
			<tr>
				<td><input name="TX_RUT_DEUDOR" type="text" id="TX_RUT_DEUDOR" size="15" value="<%=strRutDeudor%>" maxlength="15"></td>
				<td><input name="TX_NOMBRE" type="text" id="TX_NOMBRE" size="30" value="<%=strNombreDeudor%>" maxlength="40"></td>
				<td><input name="TX_RUT_REPLEGAL" type="text" id="TX_RUT_REPLEGAL" size="15" value="<%=TX_RUT_REPLEGAL%>" maxlength="15"></td>
				<td><input name="TX_NOMBRE_REPLEGAL" type="text" id="TX_NOMBRE_REPLEGAL" size="30" value="<%=TX_NOMBRE_REPLEGAL%>" maxlength="40"></td>
				<% If strUsaNro_Cliente <> "0" Then%>
				<td><input name="TX_NRO_CLIENTE" type="text" id="TX_NRO_CLIENTE" size="15" value="<%=TX_NRO_CLIENTE%>" maxlength="15"></td>
				<%Else%>
				<td>&nbsp;</td>
				<% End If %>
			</tr>
		<thead>
			<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td>DOCUMENTO</td>
				<td>FONO ASOCIADO</td>
				<td Colspan = 3>MONTO DOCUMENTO</td>
			</tr>
		</thead>
			<tr>
				<td><input name="TX_DOCUMENTO" type="text" id="TX_DOCUMENTO" size="15" value="<%=TX_DOCUMENTO%>" maxlength="20"></td>
				<td><input name="TX_TELASOCIADO" type="text" id="TX_TELASOCIADO" size="12" value="<%=TX_TELASOCIADO%>" maxlength="20"></td>
				<td><input name="TX_MONTOASOCIADO" type="text" id="TX_MONTOASOCIADO" size="12" value="<%=TX_MONTOASOCIADO%>" maxlength="20"></td>
				<td>&nbsp;</td>
				<td align= "right">
				
					<input name="Limpiar" type="button" class="fondo_boton_100" value="Limpiar"  onClick="limpiar();">				
				<%If strOrigen = "" Then %>
					<input type="button" name="Submit" class="fondo_boton_100" value="Buscar" onClick="envia();">
				<% Else %>
					<input type="button" name="Submit" class="fondo_boton_100" value="Buscar" onClick="envia2();">
				<% End If %>
				</td>
			</tr>
    </table>
	
	<% ElseIf intTipoNegocio = "2" Then%>

	<table width="100%" border="0" class="estilo_columnas">
		<thead>
			<tr >
				<td>RUT DEUDOR</td>
				<td >NOMBRE DEUDOR</td>
				<td>RUT CLIENTE</td>
				<td>NOMBRE CLIENTE</td>
			</tr>
		</thead>
			<tr>
				<td><input name="TX_RUT_DEUDOR" type="text" id="TX_RUT_DEUDOR" size="15" value="<%=strRutDeudor%>" maxlength="15"></td>
				<td><input name="TX_NOMBRE" type="text" id="TX_NOMBRE" size="30" value="<%=strNombreDeudor%>" maxlength="40"></td>
				<td><input name="TX_RUTSUBCLIENTE" type="text" id="TX_RUTSUBCLIENTE" size="15" value="<%=strRutSubCliente%>" maxlength="15"></td>
				<td><input name="TX_NOMBRE_SUBCLIENTE" type="text" id="TX_NOMBRE_SUBCLIENTE" size="30" value="<%=TX_NOMBRE_SUBCLIENTE%>" maxlength="40"></td>
			</tr>
		<thead>
			<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td>DOCUMENTO</td>
				<td>FONO ASOCIADO</td>
				<td Colspan = 2>MONTO DOCUMENTO</td>
			</tr>
		</thead>
			<tr>
				<td><input name="TX_DOCUMENTO" type="text" id="TX_DOCUMENTO" size="15" value="<%=TX_DOCUMENTO%>" maxlength="20"></td>
				<td><input name="TX_TELASOCIADO" type="text" id="TX_TELASOCIADO" size="12" value="<%=TX_TELASOCIADO%>" maxlength="20"></td>
				<td><input name="TX_MONTOASOCIADO" type="text" id="TX_MONTOASOCIADO" size="12" value="<%=TX_MONTOASOCIADO%>" maxlength="20"></td>
				<td align= "right">
				
					<input name="Limpiar" type="button" class="fondo_boton_100" value="Limpiar"  onClick="limpiar();">				
				<%If strOrigen = "" Then %>
					<input type="button" name="Submit" class="fondo_boton_100" value="Buscar" onClick="envia();">
				<% Else %>
					<input type="button" name="Submit" class="fondo_boton_100" value="Buscar" onClick="envia2();">
				<% End If %>
				</td>
			</tr>
    </table>
	
	<% ElseIf intTipoNegocio = "3" Then%>

	<table width="100%" border="0" class="estilo_columnas">
		<thead>
			<tr >
				<td>RUT DEUDOR</td>
				<td Colspan = 4>NOMBRE DEUDOR</td>
			</tr>
		</thead>
			<tr>
				<td><input name="TX_RUT_DEUDOR" type="text" id="TX_RUT_DEUDOR" size="15" value="<%=strRutDeudor%>" maxlength="15"></td>
				<td><input name="TX_NOMBRE" type="text" id="TX_NOMBRE" size="30" value="<%=strNombreDeudor%>" maxlength="40"></td>
			</tr>
		<thead>
			<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td>DOCUMENTO</td>
				<td>FONO ASOCIADO</td>
				<td Colspan = 2>MONTO DOCUMENTO</td>
			</tr>
		</thead>
			<tr>
				<td><input name="TX_DOCUMENTO" type="text" id="TX_DOCUMENTO" size="15" value="<%=TX_DOCUMENTO%>" maxlength="20"></td>
				<td><input name="TX_TELASOCIADO" type="text" id="TX_TELASOCIADO" size="12" value="<%=TX_TELASOCIADO%>" maxlength="20"></td>
				<td><input name="TX_MONTOASOCIADO" type="text" id="TX_MONTOASOCIADO" size="12" value="<%=TX_MONTOASOCIADO%>" maxlength="20"></td>
				<td align= "right">
				
					<input name="Limpiar" type="button" class="fondo_boton_100" value="Limpiar"  onClick="limpiar();">				
				<%If strOrigen = "" Then %>
					<input type="button" name="Submit" class="fondo_boton_100" value="Buscar" onClick="envia();">
				<% Else %>
					<input type="button" name="Submit" class="fondo_boton_100" value="Buscar" onClick="envia2();">
				<% End If %>
				</td>
			</tr>
    </table>
	
	<% End If%>
	
		<%
		If Trim(strRutDeudor) <> "" or Trim(strNombreDeudor) <> "" or Trim(TX_DOCUMENTO) <> "" or Trim(strRutSubCliente) <> "" or Trim(TX_NOMBRE_SUBCLIENTE) <> "" or Trim(TX_TELASOCIADO) <> "" or Trim(TX_MONTOASOCIADO) <> "" or Trim(TX_NOMBRE_REPLEGAL) <> "" or Trim(TX_RUT_REPLEGAL) <> "" or Trim(TX_NRO_CLIENTE) <> "" Then

				ssql=" SELECT C.RUT_DEUDOR,MAX(D.NOMBRE_DEUDOR) AS NOMBRE_DEUDOR,"
				
			If strUsaRep_Legal = "1" Then
				ssql=ssql & " D.REPLEG_RUT, D.REPLEG_NOMBRE," 
			End If
				
				ssql=ssql & " C.RUT_SUBCLIENTE, MAX(C.NOMBRE_SUBCLIENTE) AS NOMBRE_SUBCLIENTE, (CASE WHEN SUM(ESTADO_DEUDA.ACTIVO) = 0 THEN 'NO ACTIVO' ELSE 'ACTIVO' END) AS ESTADO_DEUDOR"
				ssql=ssql & " FROM CUOTA C	INNER JOIN DEUDOR D ON C.COD_CLIENTE = D.COD_CLIENTE AND C.RUT_DEUDOR = D.RUT_DEUDOR"
				ssql=ssql & " 				INNER JOIN ESTADO_DEUDA ON C.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
				ssql=ssql & " 				LEFT JOIN DEUDOR_TELEFONO ON DEUDOR_TELEFONO.RUT_DEUDOR = C.RUT_DEUDOR"
				ssql=ssql & " WHERE C.COD_CLIENTE = '" & strCodCliente & "'"

			If Trim(strRutDeudor) <> "" Then
				ssql=ssql & " AND D.RUT_DEUDOR = '" & strRutDeudor &"'"
			End If

			If Trim(strNombreDeudor) <> "" Then
				ssql=ssql & " AND D.NOMBRE_DEUDOR LIKE '%"&strNombreDeudor&"%'"
			End If

			If Trim(TX_RUT_REPLEGAL) <> "" Then
				ssql=ssql & " AND D.REPLEG_RUT = '" & TX_RUT_REPLEGAL &"'"
			End If

			If Trim(TX_NOMBRE_REPLEGAL) <> "" Then
				ssql=ssql & " AND D.REPLEG_NOMBRE LIKE '%"&TX_NOMBRE_REPLEGAL&"%'"
			End If
			
			If Trim(strRutSubCliente) <> "" Then
				ssql=ssql & " AND  C.RUT_SUBCLIENTE = '" & strRutSubCliente &"'"
			End If

			If Trim(TX_NOMBRE_SUBCLIENTE) <> "" Then
				ssql=ssql & " AND C.NOMBRE_SUBCLIENTE LIKE '%" & TX_NOMBRE_SUBCLIENTE & "%'"
			End If

			If Trim(TX_DOCUMENTO) <> "" Then
				ssql=ssql & " AND C.NRO_DOC = '" & TX_DOCUMENTO &"'"
			End If

			If Trim(TX_NRO_CLIENTE) <> "" Then
				ssql=ssql & " AND C.NRO_CLIENTE_DEUDOR = '" & TX_NRO_CLIENTE &"'"
			End If

			If Trim(TX_TELASOCIADO) <> "" Then
				ssql=ssql & " AND DEUDOR_TELEFONO.TELEFONO_DAL = '" & TX_TELASOCIADO &"'"
			End If

			If Trim(TX_MONTOASOCIADO) <> "" Then
				ssql=ssql & " AND (C.VALOR_CUOTA = "& TX_MONTOASOCIADO &" OR C.SALDO = "& TX_MONTOASOCIADO &")"
			End If
			
				ssql=ssql & " GROUP BY C.RUT_DEUDOR,C.RUT_SUBCLIENTE"
			
			If strUsaRep_Legal = "1" Then
				ssql=ssql & ",D.REPLEG_RUT,D.REPLEG_NOMBRE"
			End If
			
				ssql=ssql & " ORDER BY NOMBRE_SUBCLIENTE ASC"

			'''Response.write "ssql=" & ssql

			AbrirSCG()
			set rsBU=Conn.execute(ssql)

				intEstadoA = 0
				intEstadoNA = 0

				do until rsBU.eof

					If rsBU("ESTADO_DEUDOR") = "ACTIVO" then
						intEstadoA = intEstadoA + 1
					Else
						intEstadoNA = intEstadoNA + 1
					End If

				rsBU.movenext
				loop
			CerrarSCG()

			AbrirSCG()
			set rsBU=Conn.execute(ssql)

			'Response.write "intEstadoA=" & intEstadoA
			'Response.write "intEstadoNA=" & intEstadoNA
			%>

				<table width="100%"  border="0" class="intercalado" style="width:100%;">
				<thead>
					<tr>
						<td colspan = "6" class="subtitulo_informe" height="30">
							> RESULTADO DE LA BÚSQUEDA - DEUDORES / CLIENTES ACTIVOS
						</td>
					</tr>

				<%if not rsBU.eof and intEstadoA > 0 then%>

					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					  <td width="10%">RUT DEUDOR</td>
					  <td <%If strUsaRep_Legal = "1" Then %> width="25%" <%Else%> width="40%" <%End If%>>NOMBRE DEUDOR</td>
					
				<%If strUsaRep_Legal = "1" Then%>
					  <td width="10%">RUT GIRADOR</td>
					  <td width="25%">NOMBRE GIRADOR</td>
				<%End If%>
	
					  <td width="10%">RUT CLIENTE</td>
					  <td>NOMBRE CLIENTE</td>
					</tr>
				</thead>
				<tbody>
					<%do until rsBU.eof

						intEstadoDeudor = rsBU("ESTADO_DEUDOR")

						If intEstadoDeudor = "ACTIVO" Then
						%>
						<tr>

						<% If strOrigen = "" Then %>
						  <td><a href="principal.asp?TX_RUT=<%=rsBU("RUT_DEUDOR")%>"><%=rsBU("RUT_DEUDOR")%></a></td>
						<% Else %>
						  <td height="22"><%=rsBU("RUT_DEUDOR")%></td>
						<% End If %>

						  <td class="Estilo4" title="<%=rsBU("NOMBRE_DEUDOR")%>">
						  <%=mid(rsBU("NOMBRE_DEUDOR"),1,40)%>
					
						<% If strUsaRep_Legal = "1" Then %>
							  <td><%=rsBU("REPLEG_RUT")%></td>

							  <td class="Estilo4" title="<%=rsBU("REPLEG_NOMBRE")%>">
							  <%=mid(rsBU("REPLEG_NOMBRE"),1,40)%>
						<% End If %>
					
						  <td><%=rsBU("RUT_SUBCLIENTE")%></td>
						  <td class="Estilo4" title="<%=rsBU("NOMBRE_SUBCLIENTE")%>">
						  <%=mid(rsBU("NOMBRE_SUBCLIENTE"),1,40)%>

						</tr>

						<%
						End If
						rsBU.movenext
						loop

				Else%>

						<td height= "20" align="center" bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8" colspan = "6"><b>NO SE ENCONTRARON COINCIDENCIAS</td>

			  <%End if
			CerrarSCG()%>
				</body>
				<thead>
					<tr>
						<td colspan = "4" class="subtitulo_informe" height="40">
						<br>	
						<br>
						<br>
							RESULTADO DE LA BÚSQUEDA - DEUDORES / CLIENTES NO ACTIVOS
						<BR>
						</td>
					</tr>

					<%AbrirSCG()

					set rsBU=Conn.execute(ssql)

						if not rsBU.eof and intEstadoNA > 0 then%>

							<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
							  <td width="10%">RUT DEUDOR</td>
							  <td <%If strUsaRep_Legal = "1" Then %> width="25%" <%Else%> width="40%" <%End If%>>NOMBRE DEUDOR</td>
							
						<%If strUsaRep_Legal = "1" Then%>
							  <td width="10%">RUT GIRADOR</td>
							  <td width="25%">NOMBRE GIRADOR</td>
						<%End If%>
			
							  <td width="10%">RUT CLIENTE</td>
							  <td>NOMBRE CLIENTE</td>
							</tr>
					</thead>
					<tbody>
								<%
								do until rsBU.eof

								intEstadoDeudor = rsBU("ESTADO_DEUDOR")

								If intEstadoDeudor = "NO ACTIVO" Then
								%>
								<tr>

									<% If strOrigen = "" Then %>
									  <td><a href="principal.asp?TX_RUT=<%=rsBU("RUT_DEUDOR")%>"><%=rsBU("RUT_DEUDOR")%></a></td>
									<% Else %>
									  <td height="22"><%=rsBU("RUT_DEUDOR")%></td>
									<% End If %>

									  <td class="Estilo4" title="<%=rsBU("NOMBRE_DEUDOR")%>">
									  <%=mid(rsBU("NOMBRE_DEUDOR"),1,40)%>
								
								<% If strUsaRep_Legal = "1" Then %>
									  <td><%=rsBU("REPLEG_RUT")%></td>

									  <td class="Estilo4" title="<%=rsBU("REPLEG_NOMBRE")%>">
									  <%=mid(rsBU("REPLEG_NOMBRE"),1,40)%>
								<% End If %>
								
									  <td><%=rsBU("RUT_SUBCLIENTE")%></td>
									  <td class="Estilo4" title="<%=rsBU("NOMBRE_SUBCLIENTE")%>">
									  <%=mid(rsBU("NOMBRE_SUBCLIENTE"),1,40)%>
								  
								</tr>

								<%
								End If
								rsBU.movenext
								loop

						  Else%>

							<tr><td height= "20" align = "center" bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8" colspan = "6"><b>NO SE ENCONTRARON COINCIDENCIAS</td></tr>

						<%End if
					rsBU.close
					set rsBU=nothing
					CerrarSCG()%>
					</tbody>
			  </table>

	<%

	End If%>
		
		<table width="100%" border="0">
			<tr height="30">
				<Td>&nbsp;</td>
			</tr>
		</table>
		
		<%
		If Trim(strRutDeudor) <> "" or Trim(strNombreDeudor) <> "" or Trim(TX_DOCUMENTO) <> "" or Trim(strRutSubCliente) <> "" or Trim(TX_NOMBRE_SUBCLIENTE) <> "" or Trim(TX_TELASOCIADO) <> "" or Trim(TX_MONTOASOCIADO) <> "" or Trim(TX_NOMBRE_REPLEGAL) <> "" or Trim(TX_RUT_REPLEGAL) <> "" or Trim(TX_NRO_CLIENTE) <> "" Then

				ssql=" SELECT C.ID_CUOTA,C.RUT_DEUDOR, CAST(ISNULL(COD_GESTION_EXTERNA,'0000') AS INT) AS COD_GESTION_EXTERNA, D.NOMBRE_DEUDOR AS NOMBRE_DEUDOR, C.RUT_SUBCLIENTE, C.NOMBRE_SUBCLIENTE AS NOMBRE_SUBCLIENTE,"
				ssql=ssql & " (CASE WHEN ESTADO_DEUDA.ACTIVO = 0 THEN 'NO ACTIVO' ELSE 'ACTIVO' END) AS ESTADO_DEUDOR, [dbo].[fun_trae_nom_estado_deuda] (C.ID_CUOTA) AS NOM_ESTADO_DEUDA,"
				ssql=ssql & " C.NRO_DOC,C.FECHA_VENC, ISNULL(DATEDIFF(D,C.FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO,ESTADO_DEUDA.DESCRIPCION AS NOM_ESTADO,"
				ssql=ssql & " C.NRO_CUOTA, C.VALOR_CUOTA, CONVERT(VARCHAR(10),FECHA_ESTADO,103) AS FECHA_ESTADO,FECHA_ESTADO AS FECHA_ESTADO_OR, ISNULL(DATEDIFF(D,C.FECHA_VENC,C.FECHA_ESTADO),0) AS TDE"
				ssql=ssql & " FROM CUOTA C	INNER JOIN DEUDOR D ON C.COD_CLIENTE = D.COD_CLIENTE AND C.RUT_DEUDOR = D.RUT_DEUDOR"
				ssql=ssql & " 				INNER JOIN ESTADO_DEUDA ON C.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"
				ssql=ssql & " 				LEFT JOIN DEUDOR_TELEFONO ON DEUDOR_TELEFONO.RUT_DEUDOR = C.RUT_DEUDOR"
				ssql=ssql & "				LEFT JOIN TIPO_DOCUMENTO ON C.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
				ssql=ssql & " WHERE C.COD_CLIENTE = '" & strCodCliente & "'"

			If Trim(strRutDeudor) <> "" Then
				ssql=ssql & " AND D.RUT_DEUDOR = '" & strRutDeudor &"'"
			End If

			If Trim(strNombreDeudor) <> "" Then
				ssql=ssql & " AND D.NOMBRE_DEUDOR LIKE '%"&strNombreDeudor&"%'"
			End If
			
			If Trim(TX_RUT_REPLEGAL) <> "" Then
				ssql=ssql & " AND D.REPLEG_RUT = '" & TX_RUT_REPLEGAL &"'"
			End If

			If Trim(TX_NOMBRE_REPLEGAL) <> "" Then
				ssql=ssql & " AND D.REPLEG_NOMBRE LIKE '%"&TX_NOMBRE_REPLEGAL&"%'"
			End If
			
			If Trim(strRutSubCliente) <> "" Then
				ssql=ssql & " AND  C.RUT_SUBCLIENTE = '" & strRutSubCliente &"'"
			End If

			If Trim(TX_NOMBRE_SUBCLIENTE) <> "" Then
				ssql=ssql & " AND C.NOMBRE_SUBCLIENTE LIKE '%" & TX_NOMBRE_SUBCLIENTE & "%'"
			End If

			If Trim(TX_DOCUMENTO) <> "" Then
				ssql=ssql & " AND C.NRO_DOC = '" & TX_DOCUMENTO &"'"
			End If

			If Trim(TX_NRO_CLIENTE) <> "" Then
				ssql=ssql & " AND C.NRO_CLIENTE_DEUDOR = '" & TX_NRO_CLIENTE &"'"
			End If

			If Trim(TX_TELASOCIADO) <> "" Then
				ssql=ssql & " AND DEUDOR_TELEFONO.TELEFONO_DAL = '" & TX_TELASOCIADO &"'"
			End If

			If Trim(TX_MONTOASOCIADO) <> "" Then
				ssql=ssql & " AND (C.VALOR_CUOTA = "& TX_MONTOASOCIADO &" OR C.SALDO = "& TX_MONTOASOCIADO &")"
			End If
			
				ssql=ssql & " GROUP BY C.ID_CUOTA,C.RUT_DEUDOR, COD_GESTION_EXTERNA, D.NOMBRE_DEUDOR, C.RUT_SUBCLIENTE, C.NOMBRE_SUBCLIENTE, ESTADO_DEUDA.ACTIVO, C.NRO_DOC, C.FECHA_VENC,NOM_TIPO_DOCUMENTO, ESTADO_DEUDA.DESCRIPCION, C.NRO_CUOTA, C.VALOR_CUOTA,C.FECHA_ESTADO  "

				ssql=ssql & " ORDER BY FECHA_ESTADO_OR DESC, NOMBRE_SUBCLIENTE ASC"

			'Response.write "ssql=" & ssql

			AbrirSCG()

			set rsBU=Conn.execute(ssql)

				intEstadoA = 0
				intEstadoNA = 0

				do until rsBU.eof

					If rsBU("ESTADO_DEUDOR") = "ACTIVO" then
						intEstadoA = intEstadoA + 1
					Else
						intEstadoNA = intEstadoNA + 1
					End If

				rsBU.movenext
				loop
			CerrarSCG()

			AbrirSCG()

			set rsBU=Conn.execute(ssql)

			'Response.write "intEstadoA=" & intEstadoA
			'Response.write "intEstadoNA=" & intEstadoNA
			%>


				<table width="100%"  border="0" class="intercalado" style="width:100%;">
				<thead>	
					<tr>
						<td colspan = "11" class="subtitulo_informe">
						<BR>
						  <strong>RESULTADO DE LA BÚSQUEDA / DOCUMENTOS ACTIVOS</strong><BR>
						<BR>
						</td>
					</tr>
				</thead>		

				<%if not rsBU.eof and intEstadoA > 0 then%>
				<thead>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
					
					<%If strOrigen = "" Then %>
					  <td>RUT DEUDOR</td>
					  <td>NOMBRE DEUDOR</td>
					<%End If%>
					
					  <td>RUT CLIENTE</td>
					  <td>NOMBRE CLIENTE</td>
					  <td>NºDOC</td>
					  <td>CUOTA</td>
					  <td>FEC.VENC.</td>
					  <td>ANT.</td>
					  <td>TIPO DOC.</td>
					  <td>CAPITAL</td>
					  <td align="center">ESTADO</td>
					  <td>F.ESTADO</td>
					  
					<%If strOrigen = "1" Then%>
					  <td>TDE</td>
					  <td>COD</td>
					<%End If%>
					
					  <td>MAS</td>
					  <td>&nbsp</td>
					</tr>
				</thead>
				<tbody>
					<%do until rsBU.eof

						intEstadoDeudor = rsBU("ESTADO_DEUDOR")
						strNomEstadoDeuda = Trim(rsBU("NOM_ESTADO_DEUDA"))
						If intEstadoDeudor = "ACTIVO" Then
						%>
						<tr>

						<%If strOrigen = "" Then%>
						  <td><a href="principal.asp?TX_RUT=<%=rsBU("RUT_DEUDOR")%>"><%=rsBU("RUT_DEUDOR")%></a></td>
						  
						  <td class="Estilo4" title="<%=rsBU("NOMBRE_DEUDOR")%>">
						  <%=mid(rsBU("NOMBRE_DEUDOR"),1,20)%>
							  
						<%End If%>

						  <td><%=rsBU("RUT_SUBCLIENTE")%></td>

						  <td class="Estilo4" title="<%=rsBU("NOMBRE_SUBCLIENTE")%>">
						  <%=mid(rsBU("NOMBRE_SUBCLIENTE"),1,20)%>
						  
						  <td><%=rsBU("NRO_DOC")%></td>
						  <td><%=rsBU("NRO_CUOTA")%></td>
						  <td><%=rsBU("FECHA_VENC")%></td>
						  <td><%=rsBU("ANTIGUEDAD")%></td>
						  <td><%=rsBU("TIPO_DOCUMENTO")%></td>
						  <td ALIGN="RIGHT"><%=FN(rsBU("VALOR_CUOTA"),0)%>&nbsp;&nbsp;&nbsp;</td>
						  <td><%=strNomEstadoDeuda%></td>
						  <td><%=rsBU("FECHA_ESTADO")%></td>
						  
						<%If strOrigen = "1" Then%>
						  <td><%=rsBU("TDE")%></td>
						  <td align = "center"><%=rsBU("COD_GESTION_EXTERNA")%></td>
						<%End If%>
						  
						   <td>
		  							<a href="javascript:ventanaMas('mas_datos_adicionales.asp?ID_CUOTA=<%=trim(rsBU("ID_CUOTA"))%>&strCodCliente=<%=strCodCliente%>&strRUT_DEUDOR=<%=trim(rsBU("RUT_DEUDOR"))%>&strNroDoc=<%=trim(rsBU("NRO_DOC"))%>&strNroCuota=<%=rsBU("NRO_CUOTA")%>')">
		  							<img src="../imagenes/Carpeta3.png" border="0"></a>
						   </td>

						  <td ALIGN="CENTER">
							  <a href="javascript:ventanaGestionesPorDoc('gestiones_por_documento.asp?intID_CUOTA=<%=trim(rsBU("ID_CUOTA"))%>&strCodCliente=<%=session("ses_codcli")%>&strNroDoc=<%=trim(rsBU("NRO_DOC"))%>&strNroCuota=<%=trim(rsBU("NRO_CUOTA"))%>')">
							  <img src="../imagenes/icon_gestiones.jpg" border="0">
							  </a>
						  </td>
						  
						</tr>

						<%
						End If
						rsBU.movenext
						loop

				Else%>

						<td height= "20"  align="center" bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8" colspan = "14"><b>NO SE ENCONTRARON COINCIDENCIAS</td>

			  <%End if%>
			  </tbody>
			</table>

			<%CerrarSCG()%>
			<table width="100%"  border="0" class="intercalado" style="width:100%;">
			<thead>	
					<tr>
						<td colspan = "11" class="subtitulo_informe">
						<BR>
						<br>
						  <strong>RESULTADO DE LA BÚSQUEDA / DOCUMENTOS NO ACTIVOS</strong><BR>
						<BR>
						</td>
					</tr>
			</thead>	
			<%AbrirSCG()

			set rsBU=Conn.execute(ssql)

				if not rsBU.eof and intEstadoNA > 0 then%>
				<thead>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				
					<%If strOrigen = "" Then %>
					  <td>RUT DEUDOR</td>
					  <td>NOMBRE DEUDOR</td>
					<%End If%>
					
					  <td>RUT CLIENTE</td>
					  <td>NOMBRE CLIENTE</td>
					  <td>NºDOC</td>
					  <td>CUOTA</td>
					  <td>FEC.VENC.</td>
					  <td>ANT.</td>
					  <td>TIPO DOC.</td>
					  <td>CAPITAL</td>
					  <td>ESTADO</td>
					  <td>F.ESTADO</td>
					  
					<%If strOrigen = "1" Then%>
					  <td>TDE</td>
					  <td>COD</td>
					<%End If%>
					
					  <td>MAS</td>
					  <td>&nbsp;</td>
					</tr>
				</thead>
				<tbody>
						<%
						do until rsBU.eof

						intEstadoDeudor = rsBU("ESTADO_DEUDOR")
						strNomEstadoDeuda = Trim(rsBU("NOM_ESTADO_DEUDA"))

						If intEstadoDeudor = "NO ACTIVO" Then
						%>
						<tr>

						<%If strOrigen = "" Then%>
						  <td><a href="principal.asp?TX_RUT=<%=rsBU("RUT_DEUDOR")%>"><%=rsBU("RUT_DEUDOR")%></a></td>
						  
						  <td class="Estilo4" title="<%=rsBU("NOMBRE_DEUDOR")%>">
						  <%=mid(rsBU("NOMBRE_DEUDOR"),1,20)%>
						<%End If%>
						
						  <td class="Estilo4" title="<%=rsBU("ID_CUOTA")%>">
						  <%=rsBU("RUT_SUBCLIENTE")%>

						  <td class="Estilo4" title="<%=rsBU("NOMBRE_SUBCLIENTE")%>">
						  <%=mid(rsBU("NOMBRE_SUBCLIENTE"),1,20)%>
						  
						  <td><%=rsBU("NRO_DOC")%></td>
						  <td><%=rsBU("NRO_CUOTA")%></td>
						  <td><%=rsBU("FECHA_VENC")%></td>
						  <td><%=rsBU("ANTIGUEDAD")%></td>
						  <td><%=rsBU("TIPO_DOCUMENTO")%></td>
						  <td ALIGN="RIGHT"><%=FN(rsBU("VALOR_CUOTA"),0)%>&nbsp;&nbsp;&nbsp;</td>
						  
						  <td><%=strNomEstadoDeuda%></td>
						  <td><%=rsBU("FECHA_ESTADO")%></td>
						  
						<%If strOrigen = "1" Then%>
						  <td><%=rsBU("TDE")%></td>
						  <td align = "center"><%=rsBU("COD_GESTION_EXTERNA")%></td>
						<% End If %>
						
						   <td>
		  							<a href="javascript:ventanaMas('mas_datos_adicionales.asp?ID_CUOTA=<%=trim(rsBU("ID_CUOTA"))%>&strCodCliente=<%=strCodCliente%>&strRUT_DEUDOR=<%=trim(rsBU("RUT_DEUDOR"))%>&strNroDoc=<%=trim(rsBU("NRO_DOC"))%>&strNroCuota=<%=rsBU("NRO_CUOTA")%>')">
		  							<img src="../imagenes/Carpeta3.png" border="0"></a>
						   </td>
						   
						  <td ALIGN="CENTER">
							  <a href="javascript:ventanaGestionesPorDoc('gestiones_por_documento.asp?intID_CUOTA=<%=trim(rsBU("ID_CUOTA"))%>&strCodCliente=<%=session("ses_codcli")%>&strNroDoc=<%=trim(rsBU("NRO_DOC"))%>&strNroCuota=<%=trim(rsBU("NRO_CUOTA"))%>')">
							  <img src="../imagenes/icon_gestiones.jpg" border="0">
							  </a>
						  </td>

						</tr>

						<%
						End If
						rsBU.movenext
						loop

				  Else%>

					<tr><td height= "20"  align = "center" bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8" colspan = "14"><b>NO SE ENCONTRARON COINCIDENCIAS</td></tr>

				<%End if

			CerrarSCG()%>
				</tbody>
			  </table>

	<%End If

	If intTipoNegocio = "1" Then
	strTipoNegocio1 = " && (datos.TX_NOMBRE_REPLEGAL.value=='') && (datos.TX_RUT_REPLEGAL.value=='')"
	End If

	If intTipoNegocio = "2" Then
	strTipoNegocio2 = " && (datos.TX_NOMBRE_SUBCLIENTE.value=='') && (datos.TX_RUTSUBCLIENTE.value=='') && (datos.TX_RUT_DEUDOR.value=='')"
	End If

	If strUsaNro_Cliente <> "0" Then
	strValNroCliente = " && (datos.TX_NRO_CLIENTE.value=='')"
	End If

		%>
   </td>
  </tr>
</table>
<br>
<br>
</form>
</body>
</html>
<script type="text/javascript">

	function LimpiaNumeros(numero) {
    
        var numero = new String(numero);
        var result = numero.replace(/\./g, "")
        /*alert(numero)
        alert(result)*/
        return result ;
    };

	function Solo_Numerico(variable) {
        Numer = parseInt(variable);
        if (isNaN(Numer)) {
            return "";
        }
        return Numer;
    }
	
	function FormatearNumero(numero) {
        var number = new String(numero);
        var result = '';
        while (number.length > 3) {
            result = '.' + number.substr(number.length - 3) + result;
            number = number.substring(0, number.length - 3);
        }
        result = number + result;
        /*alert(result);*/
        return result;

    };

	$(document).ready(function(){
		$(document).tooltip();
		
		$("#TX_MONTOASOCIADO").numeric();
		
		$("#TX_MONTOASOCIADO").val(FormatearNumero($("#TX_MONTOASOCIADO").val()));
		
		$("#TX_MONTOASOCIADO").blur(function(){
		
			var valor = $(this).val();
			
			valor = LimpiaNumeros(valor);
			
			valor = Solo_Numerico(valor);
			
			valor = FormatearNumero(valor);
		
			$(this).val(valor);
		
		});
	})
</script>


<script language="JavaScript1.2">

function envia(){
	if ((datos.TX_RUT_DEUDOR.value=='') && (datos.TX_NOMBRE.value=='') && (datos.TX_DOCUMENTO.value=='') && (datos.TX_TELASOCIADO.value=='') && (datos.TX_MONTOASOCIADO.value=='') <%=strTipoNegocio1%> <%=strTipoNegocio2%> <%=strValNroCliente%>) {
		alert('DEBE INGRESAR TEXTO DE BUSQUEDA');
	}

else
{
	datos.action='busqueda.asp';
	datos.submit();
}
}

function envia2(){
	if ((datos.TX_NOMBRE.value=='') && (datos.TX_DOCUMENTO.value=='') && (datos.TX_TELASOCIADO.value=='') && (datos.TX_MONTOASOCIADO.value=='') <%=strValRepLeg%> <%=strValCliente%> <%=strValNroCliente%>) {
		alert('DEBE INGRESAR TEXTO DE BUSQUEDA');
	}

else
{
	datos.action='busqueda.asp?strOrigen=1';
	datos.submit();
}
}
function limpiar(){
	datos.action='busqueda.asp?strLimpiar=S&strOrigen=<%=strOrigen%>';
	datos.submit();

}
function ventanaMas (URL){
window.open(URL,"DATOS","width=400, height=450, scrollbars=no, menubar=no, location=no, resizable=yes")
}

function ventanaGestionesPorDoc (URL){
	window.open(URL,"DATOS","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

</script>
