<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<link rel="stylesheet" href="../css/style_generales_sistema.css">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	rut = request("rut")
	cliente=request("cliente")
	strRutSubCliente=request("strRutSubCliente")
	intCodEstado = request("intCodEstado")

%>

	<title>DETALLE DE DEUDA</title>

<%
AbrirSCG()

%>
</head>
<body>
<div class="titulo_informe">Detalle Deuda <%=nombre_cliente%></div>
<br>
<table width="100%" border="0" ALIGN="CENTER">
  <tr>
    <td valign="top">
	<%
	If Trim(rut) <> "" then
	abrirscg()



	strSql="SELECT ISNULL(NRO_CLIENTE_DEUDOR,0) AS NRO_CLIENTE,ISNULL(USA_REPLEGAL,0) AS REP_LEGAL,USA_SUBCLIENTE, USA_INTERESES, USA_HONORARIOS, USA_PROTESTOS, FORMULA_HONORARIOS, FORMULA_INTERESES, IsNull(ADIC_1,'ADIC_1') as ADIC_1,IsNull(ADIC_2,'ADIC_2') as ADIC_2, IsNull(ADIC_3,'ADIC_3') as ADIC_3, IsNull(ADIC_4,'ADIC_4') as ADIC_4, IsNull(ADIC_5,'ADIC_5') as ADIC_5, IsNull(ADIC_91,'ADIC_91') as ADIC_91, IsNull(ADIC_92,'ADIC_92') as ADIC_92, IsNull(ADIC_93,'ADIC_93') as ADIC_93, IsNull(ADIC_94,'ADIC_94') as ADIC_94, IsNull(ADIC_95,'ADIC_95') as ADIC_95, USA_CUSTODIO, IsNull(COLOR_CUSTODIO,'FFFFFF') as COLOR_CUSTODIO, INTERES_MORA, COD_TIPODOCUMENTO_HON, MESES_TD_HON FROM CLIENTE WHERE COD_CLIENTE = '" & cliente & "'"
	'response.write "strSql=" & strSql
	'Response.End
	set rsDET=Conn.execute(strSql)
	if Not rsDET.eof Then

		strUsaNro_Cliente = rsDET("NRO_CLIENTE")
		strUsaRep_Legal = rsDET("REP_LEGAL")


		strUsaSubCliente = rsDET("USA_SUBCLIENTE")
		strUsaInteres = rsDET("USA_INTERESES")
		strUsaHonorarios = rsDET("USA_HONORARIOS")
		strUsaProtestos = rsDET("USA_PROTESTOS")


		strNombreAdic1 = Mid(rsDET("ADIC_1"),1,10)
		strNombreAdic2 = Mid(rsDET("ADIC_2"),1,10)
		strNombreAdic3 = Mid(rsDET("ADIC_3"),1,10)

		strNomFormHon = ValNulo(rsDET("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsDET("FORMULA_INTERESES"),"C")

		strUsaCustodio = rsDET("USA_CUSTODIO")
		intTasaMensual = ValNulo(rsDET("INTERES_MORA"),"C")
		intTipoDocHono = ValNulo(rsDET("COD_TIPODOCUMENTO_HON"),"C")
		intMesHon = ValNulo(rsDET("MESES_TD_HON"),"C")
		strColorCustodio = rsDET("COLOR_CUSTODIO")
	end if
	If intTasaMensual = "" Then
		%>
		<SCRIPT>alert('No se ha definido tasa de interes de mora, se ocupara una tasa del 2%, favor parametrizar')</SCRIPT>
		<%
		intTasaMensual = "2"
	End If

	strSql = "SELECT dbo." & strNomFormInt & "(ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(ID_CUOTA) as HONORARIOS,  CUOTA.NRO_CLIENTE_DEUDOR, ID_CUOTA, NRO_CUOTA, RUT_DEUDOR, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC, IsNull(datediff(d,FECHA_VENC,getdate()),0) as ANTIGUEDAD, NRO_DOC,IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,IsNull(USUARIO_ASIG,0) as USUARIO_ASIG, NRO_CUOTA, IsNull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS, SUCURSAL, ESTADO_DEUDA, [dbo].[fun_trae_nom_estado_deuda] (ID_CUOTA) AS NOM_ESTADO_DEUDA, COD_REMESA, CUENTA, NRO_DOC, TIPO_DOCUMENTO, CONVERT(VARCHAR(10),FECHA_ESTADO,103) AS FEC_ESTADO, ADIC_1, ADIC_2, ADIC_3, IsNull(CUSTODIO,'LLACRUZ') as CUSTODIO, NOM_TIPO_DOCUMENTO "
	strSql = strSql & " ,( "
	strSql = strSql & " SELECT COUNT(*) "
	strSql = strSql & " FROM CARGA_ARCHIVOS_CUOTA car "
	strSql = strSql & " WHERE CAR.ID_CUOTA =CUOTA.ID_CUOTA AND car.activo=1 "
	strSql = strSql & " ) CANTIDAD_DOCUMENTOS "

	strSql = strSql & " FROM CUOTA, TIPO_DOCUMENTO "
	strSql = strSql & " WHERE RUT_DEUDOR='"& rut &"' AND COD_CLIENTE='" & cliente & "' AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO "

	If trim(strRutSubCliente) <> "" Then
		strSql = strSql & " AND RUT_SUBCLIENTE = '" & strRutSubCliente & "'"
	End if

	''Response.write "strSql = " & strSql

	If Trim(intCodEstado) <> "" Then
		strSql = strSql & " AND ESTADO_DEUDA = " & intCodEstado
	End If

	strSql = strSql & " ORDER BY FECHA_VENC DESC"
		'response.Write(ssql)
		'response.End()
		set rsDET=Conn.execute(strSql)
		if not rsDET.eof then
		%>
		  <table class="intercalado">
		  	<thead>
	        <tr >
	          <td>ID</td>
	          <td>RUT</td>

	          <% If strUsaNro_Cliente <> "0" Then %>
	          <td><%=strUsaNro_Cliente%></td>
	          <% End If %>

	          <td>NRO. DOC</td>
	          <td>CUOTA</td>
	          <td>F.VENCIM.</td>
	          <td>ANTIG.</td>
	          <td>TIPO DOC</td>
	          <td>CAPITAL</td>
	          <%If Trim(strUsaInteres)="1" Then%>
	          <td>INTERES</td>
	          <%End If%>
	          <%If Trim(strUsaHonorarios)="1" Then%>
	          <td>HONORARIOS</td>
	          <%End If%>
	          <%If Trim(strUsaProtestos)="1" Then%>
	          <td>PROTESTO</td>
	          <%End If%>
	          <td>SALDO</td>
	          <td>EJECUT.</td>
	          <td>ESTADO</td>
	          <td>F.ESTADO</td>
	          <% If Trim(strUsaCustodio) = "S" Then %>
	          <td>CUSTODIO</td>
	          <% End If%>
	          <td>SEDE</td>
	          <td>MAS</td>
	          <td>&nbsp</td>
	          <td>&nbsp</td>


	        </tr>
	    	</thead>
	    	<tbody>
			<%
			intSaldo = 0
			intValorCuota = 0
			total_ValorCuota = 0
			intTasaMensual = intTasaMensual/100
			intTasaDiaria = intTasaMensual/30

			Do until rsDET.eof
				intSaldo = Round(session("valor_moneda") * ValNulo(rsDET("SALDO"),"N"),0)
				intValorCuota = Round(session("valor_moneda") * ValNulo(rsDET("VALOR_CUOTA"),"N"),0)
				intProtesto = Round(session("valor_moneda") * ValNulo(rsDET("GASTOS_PROTESTOS"),"N"),0)
				strNroDoc = Trim(rsDET("NRO_DOC"))
				strNroCuota = Trim(rsDET("NRO_CUOTA"))
				strSucursal = Trim(rsDET("SUCURSAL"))
				strEstadoDeuda = Trim(rsDET("ESTADO_DEUDA"))
				strNomEstadoDeuda = Trim(rsDET("NOM_ESTADO_DEUDA"))
				strCodRemesa = Trim(rsDET("COD_REMESA"))
				strAdic1 = Trim(rsDET("ADIC_1"))
				strAdic2 = Trim(rsDET("ADIC_2"))
				strAdic3 = Trim(rsDET("ADIC_3"))

				intAntiguedad = ValNulo(rsDET("ANTIGUEDAD"),"N")
				intIntereses = rsDET("INTERESES")
				intHonorarios = rsDET("HONORARIOS")

				''Response.write "intProtesto = " & strColorCustodio

				If Trim(strUsaCustodio) = "S" and Trim(rsDET("CUSTODIO")) <> "LLACRUZ" Then
					strBgColor="#" & strColorCustodio
				Else
					strBgColor="#FFFFFF"
				End if

				strDetCuota="mas_datos_adicionales.asp"

				intTotalPorDoc = intHonorarios + intIntereses + intProtesto + intSaldo

				%>
				<tr>
					<td><div align="right"><%=rsDET("ID_CUOTA")%></div></td>
					<td><div align="right"><%=rsDET("RUT_DEUDOR")%></div></td>

			        <% If strUsaNro_Cliente <> "0" Then %>
					<td><div align="right"><%=rsDET("NRO_CLIENTE_DEUDOR")%></div></td>
					<% End If %>

					<td><div align="right"><%=rsDET("NRO_DOC")%></div></td>
					<td><div align="right"><%=rsDET("NRO_CUOTA")%></div></td>
					<td><div align="right"><%=rsDET("FECHA_VENC")%></div></td>
					<td><div align="right"><%=rsDET("ANTIGUEDAD")%></div></td>
					<td><div align="left"><%=Mid(rsDET("NOM_TIPO_DOCUMENTO"),1,10)%></div></td>
					<td align="right" >$ <%=FN((intValorCuota),0)%></td>

					<%If Trim(strUsaInteres)="1" Then%>
					  <td align="right" >$ <%=FN((intIntereses),0)%></td>
					  <%End If%>
					  <%If Trim(strUsaHonorarios)="1" Then%>
					  <td align="right" >$ <%=FN((intHonorarios),0)%></td>
					  <%End If%>
					  <%If Trim(strUsaProtestos)="1" Then%>
					  <td align="right" >$ <%=FN((intProtesto),0)%></td>
	          		<%End If%>





					<td align="right" >$ <%=FN((intTotalPorDoc),0)%></td>
					<td align="right" >
				  <%If Not rsDET("USUARIO_ASIG")="0" Then %>
					<%=TraeCampoId(Conn, "LOGIN", rsDET("USUARIO_ASIG"), "USUARIO", "ID_USUARIO")%>
				  <%else%>
					<%="SIN ASIG."%>
				  <%End If%>
				  </td>
				  <td><%=strNomEstadoDeuda%></td>
				  <td><div align="left"><%=rsDET("FEC_ESTADO")%></div></td>
				  <% If Trim(strUsaCustodio) = "S" Then %>
					  <% If Trim(rsDET("CUSTODIO")) = "LLACRUZ" Then%>
						<td><div align="left">LLACRUZ</td>
					  <% Else%>
						<td><div align="left"><img src="../imagenes/bolita7x8.jpg" border="0">&nbsp;<%=rsDET("CUSTODIO")%></div></td>
					  <% End If%>
				  <% End If%>
				  <td><div align="left"><%=Mid(Trim(rsDET("SUCURSAL")),1,10)%></div></td>


				  <td><a href="javascript:ventanaMas('<%=strDetCuota%>?ID_CUOTA=<%=trim(rsDET("ID_CUOTA"))%>&cliente=<%=cliente%>&strRUT_DEUDOR=<%=rsDET("RUT_DEUDOR")%>&strNroDoc=<%=trim(rsDET("NRO_DOC"))%>&strNroCuota=<%=rsDET("NRO_CUOTA")%>')">
						<img src="../imagenes/Carpeta3.png" border="0"></a>

				<td ALIGN="CENTER">
				<a href="javascript:ventanaGestionesPorDoc('gestiones_por_documento.asp?intID_CUOTA=<%=trim(rsDET("ID_CUOTA"))%>&strCodCliente=<%=session("ses_codcli")%>&strNroDoc=<%=trim(rsDET("NRO_DOC"))%>&strNroCuota=<%=trim(rsDET("NRO_CUOTA"))%>')">
				<img src="../imagenes/icon_gestiones.jpg" border="0">
				</a>
				</td>

				<td>
							<%IF trim(rsDET("CANTIDAD_DOCUMENTOS"))>0 then%>
								<img src="../Imagenes/48px-Crystal_Clear_filesystem_folder_yellow.png" width="20" height="20" style="cursor:pointer;" alt="Historial documentos adjuntos" onclick="bt_ver_historial('<%=trim(rsDET("ID_CUOTA"))%>')">
							<%else%>
								<img src="../Imagenes/48px-Crystal_Clear_filesystem_folder_red.png" width="20" height="20" style="cursor:pointer;" alt="Sin Historial documentos adjuntos" onclick="bt_ver_historial('<%=trim(rsDET("ID_CUOTA"))%>')">
							<%end if%>
				</td>


				 <%

					total_saldo = total_saldo + ValNulo(intSaldo,"N")
					total_interes = total_interes + ValNulo(intIntereses,"N")
					total_protestos = total_protestos + ValNulo(intProtesto,"N")
					total_honorarios = total_honorarios + ValNulo(intHonorarios,"N")
					total_TotalPorDoc = total_TotalPorDoc + ValNulo(intTotalPorDoc,"N")

					total_ValorCuota = total_ValorCuota + intValorCuota
					total_docs = total_docs + 1

				 %>
				 </tr>
				 <%rsDET.movenext
			 Loop
			 %>
			</tbody>
		<thead>
			<tr class="totales">
				<td >&nbsp;</td>
				<td ><span class="">Docs : <%=total_docs%></span></td>

			    <% If strUsaNro_Cliente <> "0" Then %>
				<td <div align="right"><span class="Estilo28">&nbsp;</span></div></td>
				<% End If %>

				<td ><div align="right"><span class="Estilo28">&nbsp;</span></div></td>
				<td ><div align="right"><span class="Estilo27">&nbsp;</span></div></td>
				<td><div align="right"><span class="Estilo27">&nbsp;</span></div></td>
				<td ><div align="right"><span class="Estilo27">&nbsp;</span></div></td>
				<td ><div align="right"><span class="Estilo27">&nbsp;</span></div></td>
				<td ><div align="right"><span class="Estilo28">$ <%=FN(total_ValorCuota,0)%></span></div></td>

				<%If Trim(strUsaInteres)="1" Then%>
				  <td ><div align="right"><span class="Estilo28">$ <%=FN(total_interes,0)%></span></div></td>
				  <%End If%>
				  <%If Trim(strUsaHonorarios)="1" Then%>
				  <td ><div align="right"><span class="Estilo28">$ <%=FN(total_honorarios,0)%></span></div></td>
					<%End If%>
				  <%If Trim(strUsaProtestos)="1" Then%>
				  <td ><div align="right"><span class="Estilo28">$ <%=FN(total_protestos,0)%></span></div></td>
				<%End If%>

				<td ><div align="right"><span class="Estilo28">$ <%=FN(total_TotalPorDoc,0)%></span></div></td>
				<td ><div align="right"><span class="Estilo28"></span></div></td>
				<td ><div align="right"><span class="Estilo28"></span></div></td>
				<td ><div align="right"><span class="Estilo28"></span></div></td>
				<td ><div align="right"><span class="Estilo28"></span></div></td>
				<td ><div align="right"><span class="Estilo28"></span></div></td>
				<% If Trim(strUsaCustodio) = "S" Then %>
					<td ><div align="right"><span class="Estilo28"></span></div></td>
				<% End If%>
				<td ><div align="right"><span class="Estilo28"></span></div></td>
				<td ><div align="right"><span class="Estilo28"></span></div></td>

			</tr>
		</thead>
	      </table>
		  <%end if
		  rsDET.close
		  set rsDET=nothing

	  %>
	  <%end if%>
    </td>
  </tr>
  <tr>
<td ALIGN="center">
<input name="imp" type="button" class="fondo_boton_100" onClick="window.print();" value="Imprimir Ficha">
</td>

</tr>
</table>

<%
cerrarSCG()
%>
</body>
</html>

<script language="JavaScript" type="text/JavaScript">

function bt_ver_historial(ID_CUOTA)
{
	window.open('historial_documentos_biblioteca_deudor.asp?ID_CUOTA='+ID_CUOTA,"_new","width=900, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function ventanaMas (URL){
window.open(URL,"DATOS","width=840, height=450, scrollbars=no, menubar=no, location=no, resizable=yes")
}
function ventanaGestionesPorDoc (URL){
	window.open(URL,"DATOS","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}
</script>
