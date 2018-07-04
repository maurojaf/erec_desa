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
	intCodEstado = request("intCodEstado")
%>
	<title>DETALLE DE DEUDA</title>

<%
AbrirSCG()

	'cliente=rsClienteDeuda("COD_CLIENTE")
	'nombre_cliente=rsClienteDeuda("DESCRIPCION")
%>

</head>
<body>
<div class="titulo_informe">Detalle Deuda <%=nombre_cliente%> Historica</div>
<br>
<table width="1140" border="0" ALIGN="CENTER">
  <tr>
    <td valign="top">
	<%
	If Trim(rut) <> "" then
	abrirscg()



	strSql="SELECT IsNull(ADIC_1,'ADIC_1') as ADIC_1, IsNull(ADIC_2,'ADIC_2') as ADIC_2, IsNull(ADIC_3,'ADIC_3') as ADIC_3, IsNull(ADIC_4,'ADIC_4') as ADIC_4, IsNull(ADIC_5,'ADIC_5') as ADIC_5, IsNull(ADIC_91,'ADIC_91') as ADIC_91, IsNull(ADIC_92,'ADIC_92') as ADIC_92, IsNull(ADIC_93,'ADIC_93') as ADIC_93, IsNull(ADIC_94,'ADIC_94') as ADIC_94, IsNull(ADIC_95,'ADIC_95') as ADIC_95, USA_CUSTODIO, IsNull(COLOR_CUSTODIO,'FFFFFF') as COLOR_CUSTODIO, INTERES_MORA, COD_TIPODOCUMENTO_HON, MESES_TD_HON FROM CLIENTE WHERE COD_CLIENTE = '" & cliente & "'"
	'response.write "strSql=" & strSql
	'Response.End
	set rsDET=Conn.execute(strSql)
	if Not rsDET.eof Then
		strNombreAdic1 = Mid(rsDET("ADIC_1"),1,10)
		strNombreAdic2 = Mid(rsDET("ADIC_2"),1,10)
		strNombreAdic3 = Mid(rsDET("ADIC_3"),1,10)
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

	ssql=""
	ssql = "SELECT ID_CUOTA,RUT_DEUDOR, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC, IsNull(datediff(d,FECHA_VENC,getdate()),0) as ANTIGUEDAD, NRO_DOC, IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,IsNull(USUARIO_ASIG,0) as USUARIO_ASIG,NRO_CUOTA, IsNull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS, SUCURSAL , ESTADO_DEUDA, COD_REMESA, CUENTA, NRO_DOC, TIPO_DOCUMENTO, CONVERT(VARCHAR(10),FECHA_ESTADO,103) AS FEC_ESTADO, ADIC_1, ADIC_2, ADIC_3, IsNull(CUSTODIO,'') as CUSTODIO, NOM_TIPO_DOCUMENTO FROM HISTORICO_CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='"&rut&"' AND COD_CLIENTE='"&cliente&"' AND HISTORICO_CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO "
	If Trim(intCodEstado) <> "" Then
		ssql = ssql & " AND ESTADO_DEUDA = " & intCodEstado
	End If

	ssql = ssql & " ORDER BY FECHA_VENC DESC"
		'response.Write(ssql)
		'response.End()
		set rsDET=Conn.execute(ssql)
		if not rsDET.eof then
		%>
		  <table width="100%" border="0" bordercolor="#FFFFFF" class="intercalado">
		  	<thead>
	        <tr bordercolor="#999999" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
	          <td>ID</td>
	          <td>RUT</td>
	          <td>NRO. DOC</td>
	          <td>F.VENCIM.</td>
	          <td>ANTIG.</td>
	          <td>TIPO DOC</td>
	          <td>ASIG.</td>
	          <td>CAPITAL</td>
	          <td>INTERES</td>
	          <td>HONORARIOS</td>
	          <td>PROTESTO</td>
	          <td>SALDO</td>
	          <td>EJECUT.</td>
	          <td>ESTADO</td>
	          <td>F.ESTADO</td>
	          <% If Trim(strUsaCustodio) = "S" Then %>
	          <td>CUSTODIO</td>
	          <% End If%>
	          <td><%=strNombreAdic1%></td>
	          <td><%=strNombreAdic2%></td>
	          <td><%=strNombreAdic3%></td>
	          <td></td>


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
			strCodRemesa = Trim(rsDET("COD_REMESA"))
			strAdic1 = Trim(rsDET("ADIC_1"))
			strAdic2 = Trim(rsDET("ADIC_2"))
			strAdic3 = Trim(rsDET("ADIC_3"))

			''Response.write "intProtesto = " & intProtesto

			If Trim(strEstadoDeuda) = "1"  Then

				intAntiguedad = ValNulo(rsDET("ANTIGUEDAD"),"N")
				If intAntiguedad > 0 Then
					intIntereses = intTasaDiaria * intAntiguedad * intSaldo

					'If Trim(cliente) <> "1070" Then ''UMAYOR
						If Trim(intTipoDocHono) = Trim(rsDET("TIPO_DOCUMENTO")) Then
							intCMeses = rsDET("ANT_MESES")
							intCDias = rsDET("ANT_DIAS")
							intCMeses = Fix((intCDias/30))

							If Cint(intCMeses) < Cint(intMesHon) Then
								intHonorarios = GASTOS_COBRANZAS(HonorariosEspeciales1(intSaldo,intCMeses,intMesHon)) * intCMeses
							Else
								intHonorarios = GASTOS_COBRANZAS(HonorariosEspeciales1(intSaldo,intCMeses,intMesHon))
							End If

						Else
							intHonorarios = GASTOS_COBRANZAS(intSaldo)
						End If
					'Else
					'	intHonorarios = GastosCobranzasUmayor(intSaldo)
					'	''Response.write "<br>intHonorarios=" & intHonorarios
					'End If
				Else
					intIntereses = 0
					intHonorarios = 0
					intProtesto = 0
				End If
			ElseIf (Trim(strEstadoDeuda) = "7" or Trim(strEstadoDeuda) = "8") Then
				''Response.write "intSaldo = " & intSaldo

				If intSaldo <> 0 Then
					intProtesto = 0
				End If
				intIntereses = 0
				intHonorarios = 0
			Else
				intIntereses = 0
				intHonorarios = 0
				intProtesto = 0
				intSaldo = 0
			End If

			''Response.write "intProtesto = " & strColorCustodio

			If Trim(strUsaCustodio) = "S" and Trim(rsDET("CUSTODIO")) <> "" Then
				intHonorarios = 0
				strBgColor="#" & strColorCustodio
			Else
				strBgColor="#FFFFFF"
			End if




			strDetCuota="mas_datos_adicionales.asp"

			intTotalPorDoc = intHonorarios + intIntereses + intProtesto + intSaldo

			%>
	        <tr bordercolor="#999999" bgcolor="<%=strBgColor%>">
	          <!--td><div align="left">&nbsp</div></td-->
	          <td><div align="right"><%=rsDET("ID_CUOTA")%></div></td>
	          <td><div align="right"><%=rsDET("RUT_DEUDOR")%></div></td>
	           <td><div align="right"><%=rsDET("NRO_DOC")%></div></td>
	          <td><div align="right"><%=rsDET("FECHA_VENC")%></div></td>
	          <td><div align="right"><%=rsDET("ANTIGUEDAD")%></div></td>
	          <td><div align="left"><%=Mid(rsDET("NOM_TIPO_DOCUMENTO"),1,10)%></div></td>
	          <td><div align="right"><%=rsDET("COD_REMESA")%></div></td>
	          <td align="right" >$ <%=FN((intValorCuota),0)%></td>
	          <td align="right" >$ <%=FN((intIntereses),0)%></td>
	          <td align="right" >$ <%=FN((intHonorarios),0)%></td>
	          <td align="right" >$ <%=FN((intProtesto),0)%></td>
	          <td align="right" >$ <%=FN((intTotalPorDoc),0)%></td>
	          <td align="right" >
	          <%If Not rsDET("USUARIO_ASIG")="0" Then %>
			  	<%=TraeCampoId(Conn, "LOGIN", rsDET("USUARIO_ASIG"), "USUARIO", "ID_USUARIO")%>
			  <%else%>
			  	<%="SIN ASIG."%>
			  <%End If%>
			  </td>
			  <td><%=TraeCampoId(Conn, "DESCRIPCION", strEstadoDeuda, "ESTADO_DEUDA", "CODIGO")%></td>
			  <td><div align="left"><%=rsDET("FEC_ESTADO")%></div></td>
			  <% If Trim(strUsaCustodio) = "S" Then %>
				  <% If Trim(rsDET("CUSTODIO")) = "" Then%>
					<td><div align="left">&nbsp;</td>
				  <% Else%>
					<td><div align="left"><img src="../imagenes/bolita7x8.jpg" border="0">&nbsp;<%=rsDET("CUSTODIO")%></div></td>
				  <% End If%>
			  <% End If%>
			  <td><div align="left"><%=Mid(Trim(rsDET("SUCURSAL")),1,10)%></div></td>
			  <td><div align="left"><%=Mid(strAdic2,1,10)%></div></td>
			  <td><div align="left"><%=Mid(strAdic3,1,10)%></div></td>


			  <td><a href="javascript:ventanaMas('<%=strDetCuota%>?ID_CUOTA=<%=trim(rsDET("ID_CUOTA"))%>&NRO_DOC=<%=trim(rsDET("NRO_DOC"))%>&cliente=<%=cliente%>&strNroDoc=<%=strNroDoc%>&strNroCuota=<%=strNroCuota%>&strSucursal=<%=strSucursal%>&strRUT_DEUDOR=<%=rsDET("RUT_DEUDOR")%>')">VER</a></td>

			 <%

				total_saldo 		= total_saldo + ValNulo(intSaldo,"N")
				total_interes 		= total_interes + ValNulo(intIntereses,"N")
				total_protestos 	= total_protestos + ValNulo(intProtesto,"N")
				total_honorarios 	= total_honorarios + ValNulo(intHonorarios,"N")
				total_TotalPorDoc 	= total_TotalPorDoc + ValNulo(intTotalPorDoc,"N")



				total_ValorCuota = total_ValorCuota + intValorCuota
				total_docs = total_docs + 1



			 %>
			 </tr>
			 <%rsDET.movenext
			 loop
			 %>
			</tbody>
			<thead>
			<tr>
				<td bgcolor="#<%=session("COLTABBG")%>">&nbsp</td>
				<td bgcolor="#<%=session("COLTABBG")%>"><span class="Estilo13">Docs : <%=total_docs%></span></td>
				<!--td bgcolor="#<%=session("COLTABBG")%>"><span class="Estilo13"></span></td-->
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28">&nbsp</span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo27">&nbsp</span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo27">&nbsp</span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo27">&nbsp</span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo27">&nbsp</span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28">$ <%=FN(total_ValorCuota,0)%></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28">$ <%=FN(total_interes,0)%></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28">$ <%=FN(total_honorarios,0)%></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28">$ <%=FN(total_protestos,0)%></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28">$ <%=FN(total_TotalPorDoc,0)%></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28"></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28"></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28"></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28"></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28"></span></div></td>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28"></span></div></td>
				<% If Trim(strUsaCustodio) = "S" Then %>
					<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28"></span></div></td>
				<% End If%>
				<td bgcolor="#<%=session("COLTABBG2")%>"><div align="right"><span class="Estilo28"></span></div></td>

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
function ventanaMas (URL){
window.open(URL,"DATOS","width=400, height=400, scrollbars=no, menubar=no, location=no, resizable=yes")
}
</script>
