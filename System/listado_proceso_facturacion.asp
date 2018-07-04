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
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<!--#include file="../lib/lib.asp"-->

	<!--#include file="../lib/comunes/rutinas/chkFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/sondigitos.inc"-->
	<!--#include file="../lib/comunes/rutinas/formatoFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/validarFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

	<!--#include file="../lib/comunes/js_css/top_tooltip.inc" -->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	Usuario_session =Session("intCodUsuario")

	AbrirSCG()

	termino = request("termino")
	inicio = request("inicio")

	strCodCliente = REQUEST("CB_CLIENTE")
	strEstado = REQUEST("CB_ESTADO")
	dtmFechaProc = REQUEST("CB_FECHA")

	'Response.write "<br>dtmFechaProc=" & dtmFechaProc
	'Response.write "<br>strEstado=" & strEstado
%>

</head>
<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<form name="datos" method="post">
<div class="titulo_informe">LISTADO PROCESO FACTURACIÓN</div>	
<br>
<table width="90%" border="0" align="center">
  <tr height="20">
    <td style="vertical-align: top;">
		<table width="100%" border="0" class="estilo_columnas">
			<thead>
			  <tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
				<td>CLIENTE</td>
				<td>ESTADO PROCESO</td>
				<td>FECHA PROCESO</td>
				<td>&nbsp;</td>


			  </tr>
			 </thead>
			  <tr bordercolor="#999999" class="Estilo8">

				<td>

				<SELECT NAME="CB_CLIENTE" id="CB_CLIENTE">

					<option value="0">SELECCIONAR</option>
					<%
						AbrirSCG()
						ssql="SELECT COD_CLIENTE,RAZON_SOCIAL, NOMBRE_FANTASIA FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
								<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCodCliente)=Trim(rsTemp("COD_CLIENTE")) then response.Write("Selected") End If%>><%=rsTemp("NOMBRE_FANTASIA")%></option>
									<%
								rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing
						CerrarSCG()

					%>
				</SELECT>


				</td>

				<td>
					<SELECT NAME="CB_ESTADO" id="CB_ESTADO" onChange="CargaFechas(this.value,CB_CLIENTE.value);">
						<option value="0" <%If Trim(strEstado)="0" Then Response.write "SELECTED"%>>SELECCIONE</option>
						<option value="1" <%If Trim(strEstado)="1" Then Response.write "SELECTED"%>>DOCUMENTOS A VISAR</option>
						<option value="2" <%If Trim(strEstado)="2" Then Response.write "SELECTED"%>>DOCUMENTOS ENVIADOS A VISAR</option>
						<option value="3" <%If Trim(strEstado)="3" Then Response.write "SELECTED"%>>DOCUMENTOS ENVIADOS A FACTURAR</option>
					</SELECT>
				</td>
				<td>
					<SELECT NAME="CB_FECHA" id="CB_FECHA">
					</SELECT>
				</td>

				<td align = "right" >
				<input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();">
				<input Name="SubmitButton" class="fondo_boton_100" Value="Exportar" Type="BUTTON" onClick="exportar();">
				</td>
			  </tr>
		</table>
    </td>
   </tr>
   <tr><td>
	<table width="100%" border="0" bordercolor="#000000" class="intercalado" style="width:100%;">


	<%
	AbrirSCG1()
	
	strSql="SELECT FORMULA_HONORARIOS_FACT,FORMULA_HONORARIOS,FORMULA_INTERESES FROM CLIENTE WHERE COD_CLIENTE = '" & strCodCliente & "'"
	''Response.write "strSql="&strSql
	set rsDET=Conn1.execute(strSql)
	if Not rsDET.eof Then
		strNomFormHonFact = ValNulo(rsDET("FORMULA_HONORARIOS_FACT"),"C")
		strNomFormHon = ValNulo(rsDET("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsDET("FORMULA_INTERESES"),"C")
	Else
		strNomFormHon = "NO_DEFINIDA"
		strNomFormInt = "NO_DEFINIDA"
	end if
	
	CerrarSCG1()	
	
	Abrirscg()

			strSql = " SELECT "
			strSql = strSql & " (CASE WHEN ((ESTADO_FACTURA IS NULL OR ESTADO_FACTURA IN ('4','5','7')) AND CUOTA.ESTADO_DEUDA = 3 AND (CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = ''))"
			strSql = strSql & " 		  THEN '1-ENVIAR A VISAR' "
			strSql = strSql & " 		  WHEN (ESTADO_FACTURA = '1')"
			strSql = strSql & " 		  THEN '2-ENVIADO A VISAR' "
			strSql = strSql & " 		  WHEN (ESTADO_FACTURA = '2')"
			strSql = strSql & " 		  THEN '3-ENVIADO A FACTURAR'"
			strSql = strSql & " END) AS ESTADO_MODULO, "
			strSql = strSql & " (CASE WHEN ((ESTADO_FACTURA IS NULL OR ESTADO_FACTURA IN ('4','5','7')) AND CUOTA.ESTADO_DEUDA = 3 AND (CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = ''))"
			strSql = strSql & " 		  THEN 1 "
			strSql = strSql & " 		  WHEN (ESTADO_FACTURA = '1')"
			strSql = strSql & " 		  THEN 2 "
			strSql = strSql & " 		  WHEN (ESTADO_FACTURA = '2')"
			strSql = strSql & " 		  THEN 3"
			strSql = strSql & " END) AS ORDEN, "
			strSql = strSql & " CUOTA.COD_CLIENTE, CLIENTE.DESCRIPCION AS DESCLI,"
			strSql = strSql & " COUNT(CUOTA.ID_CUOTA) AS TOTAL_DOC, "
			strSql = strSql & " SUM(VALOR_CUOTA) AS VALOR_CUOTA,"
			strSql = strSql & " SUM((CASE WHEN ((ESTADO_FACTURA IS NULL OR ESTADO_FACTURA IN ('4','5','7')) AND CUOTA.ESTADO_DEUDA = 3 AND (CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = ''))"
			strSql = strSql & " 		  THEN CASE WHEN CUOTA.COD_CLIENTE = 1100 THEN dbo.fun_honorarios_Facturacion_BCI(ID_CUOTA) ELSE dbo.fun_honorarios_fact(ID_CUOTA) END"
			strSql = strSql & " 		  WHEN '2' IN (2) THEN CUOTA.MONTO_VISACION "
			strSql = strSql & " 		  WHEN '2' IN (3) THEN CUOTA.MONTO_FACTURACION "
			strSql = strSql & " END)) AS HONORARIOS_FACT"
			strSql = strSql & " FROM CUOTA  INNER JOIN CLIENTE ON CUOTA.COD_CLIENTE = CLIENTE.COD_CLIENTE "
			strSql = strSql & " 			LEFT JOIN USUARIO ON CUOTA.USUARIO_ESTADO_FACT = USUARIO.ID_USUARIO "
			strSql = strSql & " 			LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO "
			strSql = strSql & " 			LEFT JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE "
			strSql = strSql & " 			LEFT JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO "
			
			strSql = strSql & " WHERE "
			strSql = strSql & " ((ESTADO_FACTURA IS NULL OR ESTADO_FACTURA IN ('4','5','7')) AND CUOTA.ESTADO_DEUDA = 3 AND (CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = ''))"
			strSql = strSql & " OR (ESTADO_FACTURA = '1')"
			strSql = strSql & " OR (ESTADO_FACTURA = '2')"
			strSql = strSql & " AND CLIENTE.ACTIVO = 1"

			strSql = strSql & " GROUP BY CUOTA.COD_CLIENTE, CLIENTE.DESCRIPCION,"
			strSql = strSql & " (CASE WHEN ((ESTADO_FACTURA IS NULL OR ESTADO_FACTURA IN ('4','5','7')) AND CUOTA.ESTADO_DEUDA = 3 AND (CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = ''))"
			strSql = strSql & " 		  THEN '1-ENVIAR A VISAR' "
			strSql = strSql & " 		  WHEN (ESTADO_FACTURA = '1')"
			strSql = strSql & " 		  THEN '2-ENVIADO A VISAR' "
			strSql = strSql & " 		  WHEN (ESTADO_FACTURA = '2')"
			strSql = strSql & " 		  THEN '3-ENVIADO A FACTURAR'"
			strSql = strSql & " END),"
			strSql = strSql & " (CASE WHEN ((ESTADO_FACTURA IS NULL OR ESTADO_FACTURA IN ('4','5','7')) AND CUOTA.ESTADO_DEUDA = 3 AND (CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = ''))"
			strSql = strSql & " 		  THEN 1 "
			strSql = strSql & " 		  WHEN (ESTADO_FACTURA = '1')"
			strSql = strSql & " 		  THEN 2 "
			strSql = strSql & " 		  WHEN (ESTADO_FACTURA = '2')"
			strSql = strSql & " 		  THEN 3"
			strSql = strSql & " END)"
			
			strSql = strSql & " ORDER BY ORDEN ASC,DESCLI ASC "
			
			''Response.write "strSql = " & strSql
			
			if strSql <> "" then
				set rsDet=Conn.execute(strSql)

				intReg = 0
				intTotalCapital = 0
				intTotalDoc = 0
				intTotalHonorario = 0
					
				if not rsDet.eof then%>
				<thead>
					<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<td Width = "180">PROCESO</td>
						<td Width = "280">CLIENTE</td>
						<td Width = "100">TOTAL DOCUMENTOS</td>
						<td Width = "150">CAPITAL</td>
						<td Width = "150">TOTAL HONORARIO PROCESO</td>
					</tr>
				</thead>
				<tbody>
		
					<%do while not rsDet.eof
						intReg = intReg + 1
						intTotalDoc = intTotalDoc + Cdbl(rsDet("TOTAL_DOC"))
						intTotalCapital = intTotalCapital + Cdbl(rsDet("VALOR_CUOTA"))
						intTotalHonorario = intTotalHonorario + Cdbl(rsDet("HONORARIOS_FACT"))

						%>
						<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
							<td><%=Mid(rsDet("ESTADO_MODULO"),1,28)%></td>
							<td><%=Mid(rsDet("DESCLI"),1,28)%></td>
							<td ALIGN="RIGHT"><%=FN(rsDet("TOTAL_DOC"),0)%></td>
							<td ALIGN="RIGHT"><%=FN(rsDet("VALOR_CUOTA"),0)%></td>
							<td ALIGN="RIGHT"><%=FN(rsDet("HONORARIOS_FACT"),0)%></td>

						</tr>
						<%
						rsDet.movenext
					loop
				rsDet.close
				set rsDet=nothing%>
				</tbody>
				<thead>
					<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<td Colspan="2">TOTALES</td>
						<td ALIGN="RIGHT"><%=FN(intTotalDoc,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intTotalCapital,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intTotalHonorario,0)%></td>
					</tr>
				</thead>	
				<%Else%>
				<thead>								
					<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
						<td ALIGN="CENTER" Colspan = "5">NO EXISTEN PROCESOS PENDIENTES ASOCIADOS A CLIENTES</td>
					</tr>
				</thead>
				<%End if

		End if
		
	Cerrarscg()%>

  </table>
  	</td>
	</tr>
   	<tr>
	<td>
	<table width="100%" border="0" bordercolor="#000000" class="intercalado" style="width:100%;">
		<thead>
		<tr bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

		<%If strEstado = 0 Then%>
			<td colspan = "5" ALIGN="CENTER" >SELECCIONE PROCESO</td>
		<%ElseIf strEstado = 1 Then%>

			<td>&nbsp;</td>
			<td>CLIENTE</td>
			<td>RUT DEUDOR</td>
			<td>ID_CUOTA</td>
			<td>NRO_DOC</td>
			<td>FECHA_VENC</td>
			<td>FECHA_ASIG</td>
			<td>CAPITAL</td>
			<td>HONORARIO</td>
			<td>ESTADO</td>
			<td>F.ESTADO</td>

		<%ElseIf strEstado = 2 Then%>

			<td>&nbsp;</td>
			<td>CLIENTE</td>
			<td>RUT DEUDOR</td>
			<td>ID_CUOTA</td>
			<td>NRO_DOC</td>
			<td>FECHA_VENC</td>
			<td>FECHA_ASIG</td>
			<td>CAPITAL</td>
			<td>MONTO VISAR</td>
			<td>ESTADO</td>
			<td>F.ESTADO</td>
			<td>VISADOR</td>
			<td>FECHA VISACION</td>

		<%ElseIf strEstado = 3 Then%>

			<td>&nbsp;</td>
			<td>CLIENTE</td>
			<td>RUT DEUDOR</td>
			<td>ID_CUOTA</td>
			<td>NRO_DOC</td>
			<td>FECHA_VENC</td>
			<td>FECHA_ASIG</td>
			<td>CAPITAL</td>
			<td>MONTO FACTURAR</td>
			<td>ESTADO</td>
			<td>F.ESTADO</td>
			<td>USUARIO</td>
			<td>FECHA ENVIO FACT.</td>

		<%End If%>

		</tr>
		</thead>
		<tbody>
	<%

	resp="si"
	If resp="si" and strCodCliente <> "" then

	strSql = "SELECT 	ID_CUOTA, CUOTA.COD_CLIENTE, UPPER(LOGIN) AS USUARIO, NOM_TIPO_DOCUMENTO, CLIENTE.DESCRIPCION AS DESCCLI, CUOTA.RUT_DEUDOR AS RUT_DEUDOR,"
	strSql = strSql & "	NRO_DOC,FECHA_ENVIO_VISAR, MONTO_VISACION, FECHA_ENVIO_FACTURAR, MONTO_FACTURACION, NUMERO_FACTURA,"
	strSql = strSql & "	FECHA_FACTURACION ,ESTADO_FACTURA, USUARIO_ESTADO_FACT, OBSERVACION_FACTURACION, CONVERT(VARCHAR(10),CUOTA.FECHA_VENC,103) AS FECHA_VENC,"
	strSql = strSql & "	CONVERT(VARCHAR(10),CUOTA.FECHA_CREACION,103) AS FECHACREA,SUCURSAL, DEUDOR.NOMBRE_DEUDOR AS NOMDEUDOR, ESTADO_DEUDA.DESCRIPCION AS DESCRIPT,"
	strSql = strSql & "	HONORARIOS_FACT = (CASE WHEN '" & strEstado & "' IN (1) THEN dbo." & strNomFormHonFact & "(ID_CUOTA)"
	strSql = strSql & "							WHEN '" & strEstado & "' IN (2) THEN CUOTA.MONTO_VISACION"
	strSql = strSql & "							WHEN '" & strEstado & "' IN (3) THEN CUOTA.MONTO_FACTURACION"
	strSql = strSql & "					   END),"
	strSql = strSql & "	VALOR_CUOTA,CONVERT(VARCHAR(10),CUOTA.FECHA_ESTADO,103) AS FECHA_ESTADO"


	strSql = strSql & " FROM CUOTA  INNER JOIN CLIENTE ON CUOTA.COD_CLIENTE = CLIENTE.COD_CLIENTE"
	strSql = strSql & " 			LEFT JOIN USUARIO ON CUOTA.USUARIO_ESTADO_FACT = USUARIO.ID_USUARIO"
	strSql = strSql & " 			LEFT JOIN TIPO_DOCUMENTO ON CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
	strSql = strSql & " 			LEFT JOIN DEUDOR ON CUOTA.RUT_DEUDOR = DEUDOR.RUT_DEUDOR  AND CUOTA.COD_CLIENTE = DEUDOR.COD_CLIENTE"
	strSql = strSql & " 			LEFT JOIN ESTADO_DEUDA ON CUOTA.ESTADO_DEUDA = ESTADO_DEUDA.CODIGO"

	strSql = strSql & " WHERE  		(CUOTA.COD_CLIENTE =  '" & strCodCliente & "' AND ESTADO_FACTURA = '1' AND '" & Trim(strEstado) & "' IN (2) AND CAST( '" & dtmFechaProc & "' AS DATETIME) = FECHA_ENVIO_VISAR)"
	strSql = strSql & " 		OR  (CUOTA.COD_CLIENTE = '" & strCodCliente & "' AND ESTADO_FACTURA = '2' AND '" & Trim(strEstado) & "' IN (3) AND CAST( '" & dtmFechaProc & "' AS DATETIME) = FECHA_ENVIO_FACTURAR)"
	strSql = strSql & " 		OR  (CUOTA.COD_CLIENTE = '" & strCodCliente & "' AND '" & Trim(strEstado) & "' IN (1) AND (ESTADO_FACTURA IS NULL OR ESTADO_FACTURA IN ('4','5','7')) AND CUOTA.ESTADO_DEUDA = 3 AND (CUOTA.CUSTODIO IS NULL OR CUOTA.CUSTODIO = ''))"

	strSql = strSql & " ORDER BY FECHA_ESTADO ASC"
	
	'Response.write "strSql = " & strSql

	End if

		''Response.write "strSql = " & strSql

		'Response.End
	if strSql <> "" then
		AbrirSCG()
		set rsDet=Conn.execute(strSql)


		if not rsDet.eof then
			intReg = 0
			intTotalMonto = 0
			do while not rsDet.eof
				intReg = intReg + 1

				%>
				<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">

		<%
			If strEstado = 1 Then
				intTotalMonto = intTotalMonto + rsDet("HONORARIOS_FACT")
		%>

			<td><%=intReg%></td>
			<td><%=Mid(rsDet("DESCCLI"),1,30)%></td>
			<td><%=rsDet("RUT_DEUDOR")%></td>
			<td><%=rsDet("ID_CUOTA")%></td>
			<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
			<td><%=rsDet("FECHA_VENC")%></td>
			<td><%=rsDet("FECHACREA")%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("VALOR_CUOTA"),0)%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("HONORARIOS_FACT"),0)%></td>
			<td><%=Mid(rsDet("DESCRIPT"),1,28)%></td>
			<td><%=rsDet("FECHA_ESTADO")%></td>

		<%ElseIf strEstado = 2 Then
				intTotalMonto = intTotalMonto + rsDet("MONTO_VISACION")
			%>
			
			<td><%=intReg%></td>
			<td><%=Mid(rsDet("DESCCLI"),1,30)%></td>
			<td><%=rsDet("RUT_DEUDOR")%></td>
			<td><%=rsDet("ID_CUOTA")%></td>
			<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
			<td><%=rsDet("FECHA_VENC")%></td>
			<td><%=rsDet("FECHACREA")%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("VALOR_CUOTA"),0)%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("MONTO_VISACION"),0)%></td>
			<td><%=Mid(rsDet("DESCRIPT"),1,28)%></td>
			<td><%=rsDet("FECHA_ESTADO")%></td>
			<td><%=Mid(rsDet("USUARIO"),1,30)%></td>
			<td><%=rsDet("FECHA_ENVIO_VISAR")%></td>

		<%ElseIf strEstado = 3 Then
				intTotalMonto = intTotalMonto + rsDet("MONTO_FACTURACION")
			%>

			<td><%=intReg%></td>
			<td><%=Mid(rsDet("DESCCLI"),1,30)%></td>
			<td><%=rsDet("RUT_DEUDOR")%></td>
			<td><%=rsDet("ID_CUOTA")%></td>
			<td ALIGN="RIGHT"><%=rsDet("NRO_DOC")%></td>
			<td><%=rsDet("FECHA_VENC")%></td>
			<td><%=rsDet("FECHACREA")%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("VALOR_CUOTA"),0)%></td>
			<td ALIGN="RIGHT"><%=FN(rsDet("MONTO_FACTURACION"),0)%></td>
			<td><%=Mid(rsDet("DESCRIPT"),1,28)%></td>
			<td><%=rsDet("FECHA_ESTADO")%></td>
			<td><%=Mid(rsDet("USUARIO"),1,30)%></td>
			<td><%=Mid(rsDet("FECHA_ENVIO_FACTURAR"),1,28)%></td>

		<%End If%>

				</tr>
				<%
				rsDet.movenext
			loop
		CerrarSCG()
		end if
	end if
	AbrirSCG()

	%>
		</tbody>
		<thead>
		<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">

		<%If strEstado = 1 Then%>
			<td colspan=2 bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">TOTALES</td>
			<td colspan=2 ALIGN="RIGHT">TOTAL DOCUMENTOS</td>
			<td colspan=2 ALIGN="RIGHT"><%=FN(intReg,0)%></td>
			<td colspan=3 ALIGN="RIGHT">MONTO A VISAR</td>
			<td colspan=2 ALIGN="RIGHT">$<%=FN(intTotalMonto,0)%></td>
		<%ElseIf strEstado = 2 Then%>
			<td colspan=2 bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">TOTALES</td>
			<td colspan=2 ALIGN="RIGHT">DOCUMENTOS</td>
			<td colspan=2 ALIGN="RIGHT"><%=FN(intReg,0)%></td>
			<td colspan=5 ALIGN="RIGHT">MONTO ENVIADO A VISAR</td>
			<td colspan=2 ALIGN="RIGHT">$<%=FN(intTotalMonto,0)%></td>
		<%ElseIf strEstado = 3 Then%>
			<td colspan=2 bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">TOTALES</td>
			<td colspan=2 ALIGN="RIGHT">TOTAL DOCUMENTOS</td>
			<td colspan=2 ALIGN="RIGHT"><%=FN(intReg,0)%></td>
			<td colspan=5 ALIGN="RIGHT">MONTO A FACTURAR</td>
			<td colspan=2 ALIGN="RIGHT">$ <%=FN(intTotalMonto,0)%></td>
		<%End If%>

		</tr>
		</thead>
	</table>
	</td>
   </tr>
  </table>

</form>


</body>
</html>

<!--#include file="../lib/comunes/js_css/bottom_tooltip.inc" -->

<script type="text/javascript">

function envia(){

	if (datos.CB_CLIENTE.value=='0'){
		alert('Debe seleccionar un cliente para generar reporte');
	}
	else if (datos.CB_ESTADO.value=='0'){
		alert('Dede seleccionar el tipo de proceso para generar reporte');
	}
	else if (((datos.CB_ESTADO.value=='2') || (datos.CB_ESTADO.value=='3')) & (datos.CB_FECHA.value=='')){
		alert('Dede seleccionar la fecha de proceso para generar reporte');
	}
	else
	{
	datos.Submit.disabled = true;
	datos.SubmitButton.disabled = true;
	resp='si'
	document.datos.action = "listado_proceso_facturacion.asp?strBuscar=S&resp="+ resp +"";
	document.datos.submit();
	}
}

function exportar()
{
	datos.SubmitButton.disabled = true;
	datos.Submit.disabled = true;
	document.datos.action = "exp_Listado_Proceso_Fact.asp";
	document.datos.submit();
}

function CargaFechas(subCat,cat)
{
	//alert(subCat);

	var comboBox = document.getElementById('CB_FECHA');
	switch (cat)
	{
		<%
		  AbrirSCG()
			strSql="SELECT COD_CLIENTE FROM CLIENTE WHERE COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") AND ACTIVO = 1 ORDER BY COD_CLIENTE"
			set rsGestCat=Conn.execute(strSql)
			Do While not rsGestCat.eof
		%>
		case '<%=rsGestCat("COD_CLIENTE")%>':


			comboBox.options.length = 0;
				if (subCat=='0') {
					var newOption = new Option('SELECCIONE', '');
					comboBox.options[comboBox.options.length] = newOption;
				}
				if (subCat=='1') {
					var newOption = new Option('SIN FECHA PROCESO', '');
					comboBox.options[comboBox.options.length] = newOption;
				}

				if (subCat=='2') {
					var newOption = new Option('SELECCIONE', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql = "SELECT DISTINCT FECHA_ENVIO_VISAR "
					strSql = strSql & " FROM CUOTA WHERE ESTADO_FACTURA IS NOT NULL"
					strSql = strSql & " AND COD_CLIENTE = '" & rsGestCat("COD_CLIENTE") & "' AND ESTADO_FACTURA = '1'"
					'Response.write "<br>strSql=" & strSql
					set rsGestion=Conn2.execute(strSql)
					If Not rsGestion.Eof Then
						Do While Not rsGestion.Eof
							%>
								var newOption = new Option('<%=rsGestion("FECHA_ENVIO_VISAR")%>', '<%=rsGestion("FECHA_ENVIO_VISAR")%>');
								comboBox.options[comboBox.options.length] = newOption;
							<%
							rsGestion.movenext
						Loop
					Else
					%>
						var newOption = new Option('SIN FECHA ENVIO A VISAR', '');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					CerrarSCG2()
					%>
					break;
				}

				if (subCat=='3') {
					var newOption = new Option('SELECCIONE', '');
					comboBox.options[comboBox.options.length] = newOption;
					<%

					AbrirSCG2()

					strSql = "SELECT DISTINCT FECHA_ENVIO_FACTURAR"
					strSql = strSql & " FROM CUOTA WHERE ESTADO_FACTURA IS NOT NULL"
					strSql = strSql & " AND COD_CLIENTE = '" & rsGestCat("COD_CLIENTE") & "' AND ESTADO_FACTURA = '2'"
					''Response.write "<br>strSql=" & strSql
					set rsGestion=Conn2.execute(strSql)
					If Not rsGestion.Eof Then
						Do While Not rsGestion.Eof
							%>
								var newOption = new Option('<%=rsGestion("FECHA_ENVIO_FACTURAR")%>', '<%=rsGestion("FECHA_ENVIO_FACTURAR")%>');comboBox.options[comboBox.options.length] = newOption;
							<%
							rsGestion.movenext
						Loop
					Else
					%>
						var newOption = new Option('SIN FECHA ENVIO A FACTURAR', '');
						comboBox.options[comboBox.options.length] = newOption;
					<%
					End if
					CerrarSCG2()
					%>
					break;
				}

		<%
		  	rsGestCat.movenext
		  	Loop
		  	rsGestCat.close
		  	set rsGestCat=nothing
			CerrarSCG()
		%>
	}

}
function InicializaInforme()
{
		var comboBox = document.getElementById('CB_FECHA');
		comboBox.options.length = 0;
		var newOption = new Option('SELECCIONE','');
		comboBox.options[comboBox.options.length] = newOption;
}

<%If strEstado = "" then%>
InicializaInforme();
<%End If%>

<%If strEstado <> "" then%>
CargaFechas('<%=strEstado%>','<%=strCodCliente%>');
<%End If%>

<%If dtmFechaProc <> "" then%>
datos.CB_FECHA.value='<%=dtmFechaProc%>';
<%End If%>

</script>

