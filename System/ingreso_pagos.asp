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

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link rel="stylesheet" href="../css/style_generales_sistema.css">	

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	If Trim(Request("Limpiar"))="1" Then
		session("session_RUT_DEUDOR") = ""
		rut = ""
	End if

	rut = request("rut")

	intOrigen = request("intOrigen")

	if trim(rut) <> "" Then
		session("session_RUT_DEUDOR") = rut
	Else
		rut = session("session_RUT_DEUDOR")
	End if

	strRUT_DEUDOR=rut

	intSeq = request("intSeq")
	strGraba = request("strGraba")
	txt_FechaIni = request("txt_FechaIni")
	intSucursal="1"
	fecha= date
	strCOD_CLIENTE = request("CB_CLIENTE")
	strCOD_CLIENTE = session("ses_codcli")

	usuario=session("session_idusuario")

	AbrirSCG()

	strDocCancelados = request("TX_DOCCANCELADOS")
	strObservaciones = request("TX_OBSERVACIONES")
	strObservaciones = Trim(strObservaciones)


	If Trim(request("strGraba")) = "SI" and intOrigen = "IP" Then

		intInteres = Request("TX_INTERES")
		intHonorarios = Request("TX_HONORARIOS")
		intProtestos = Request("TX_PROTESTOS")

		strSql = "SELECT ID_CUOTA, NRO_DOC, SALDO, CUENTA, FECHA_VENC FROM CUOTA WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & strCOD_CLIENTE & "' AND SALDO > 0"
		set rsTemp= Conn.execute(strSql)

		intCorrelativo = 1
		intTotalCapital = 0
		strGestionPago=""
		strGestionAbono=""

		Do until rsTemp.eof
			strObjeto = "CH_" & Replace(Trim(rsTemp("ID_CUOTA")),"-","_")
			strObjeto1 = "TX_SALDO_" & Replace(Trim(rsTemp("ID_CUOTA")),"-","_")
			strObjeto2 = "TX_COMPROBANTE_" & Replace(Trim(rsTemp("ID_CUOTA")),"-","_")


			If UCASE(Request(strObjeto)) = "ON" Then

				intSaldoCapital = rsTemp("SALDO")
				intSaldo = Request(strObjeto1)
				strComprobante = Request(strObjeto2)
				strNroDoc = rsTemp("NRO_DOC")
				intID_CUOTA = rsTemp("ID_CUOTA")
				strCuenta = rsTemp("CUENTA")
				strFechaVenc = rsTemp("FECHA_VENC")


				''Response.write "<br>Caso = " & (Cdbl(intSaldo) >= Cdbl(intSaldoCapital))

				If (Cdbl(intSaldo) >= Cdbl(intSaldoCapital)) Then
					strGestionPago="P"
					strObservacion2 = "PAGO EN CLIENTE REALIZADO POR " & session("nombre_user")
					'strSql = "UPDATE CUOTA SET SALDO = SALDO - " & intSaldo & ", ESTADO_DEUDA = '3', FECHA_ESTADO = GETDATE() , ADIC_92 = '" & strComprobante & "', OBSERVACION = '" & strObservacion2 & "' WHERE RUT_DEUDOR = '" & rut & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "' AND NRO_DOC = '" & strNroDoc & "'"
					strSql = "UPDATE CUOTA SET SALDO = SALDO - " & intSaldo & ", ESTADO_DEUDA = '3', FECHA_ESTADO = GETDATE() , ADIC_92 = '" & strComprobante & "', OBSERVACION = '" & strObservacion2 & "' WHERE ID_CUOTA = " & intID_CUOTA
					intTotalCapital = intTotalCapital + intSaldo
					strDocumentos = strDocumentos & ","& strNroDoc

				Else
					strGestionAbono="A"
					strObservacion2 = "ABONO EN CLIENTE REALIZADO POR " & session("nombre_user")
					'strSql = "UPDATE CUOTA SET SALDO = SALDO - " & intSaldo & ", ESTADO_DEUDA = '7', FECHA_ESTADO = GETDATE() , ADIC_92 = '" & strComprobante & "', OBSERVACION = '" & strObservacion2 & "' WHERE RUT_DEUDOR = '" & rut & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "' AND NRO_DOC = '" & strNroDoc & "'"
					strSql = "UPDATE CUOTA SET SALDO = SALDO - " & intSaldo & ", ESTADO_DEUDA = '7', FECHA_ESTADO = GETDATE() , ADIC_92 = '" & strComprobante & "', OBSERVACION = '" & strObservacion2 & "' WHERE ID_CUOTA = " & intID_CUOTA
					intTotalCapitalA = intTotalCapitalA + intSaldo
					strDocumentosA = strDocumentosA & ","& strNroDoc

				End If


				'Response.write "<br>strGestionAbono=" & strGestionAbono
				'Response.write "<br>" & strSql
				'Response.End
				set rsUpdate=Conn.execute(strSql)

			End if
		rsTemp.movenext
		intCorrelativo = intCorrelativo + 1
		loop
		rsTemp.close
		set rsTemp=nothing


		If Trim(strGestionPago) = "P" or Trim(strGestionAbono) = "A" Then


			strDocumentos = Mid(strDocumentos,2,len(strDocumentos))
			strDocumentosA = Mid(strDocumentosA,2,len(strDocumentosA))



			strSql="SELECT ISNULL(ID_CAMPANA,0) as ID_CAMPANA FROM DEUDOR WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='" & strCOD_CLIENTE & "'"
			set rsDeudor = Conn.execute(strSql)
			if not rsDeudor.eof then
				intIdCampana=rsDeudor("ID_CAMPANA")
			else
				intIdCampana=0
			end if
			rsDeudor.close
			set rsDeudor=nothing

			if Trim(reagen) <> "" and Trim(reagen) <> "NULL" then
				reagen = "'" + reagen + "'"
			else
				reagen="NULL"
			end if

			if Trim(retiro) <> "" then
				retiro="'" + retiro + "'"
			else
				retiro="NULL"
			end if

			if Trim(FECHA_COMPROMISO) <> "" and Trim(FECHA_COMPROMISO) <> "NULL" then
				FECHA_COMPROMISO= "'" & FECHA_COMPROMISO & "'"
			else
				FECHA_COMPROMISO="NULL"
			end if

			if Trim(fechacancelo) <> "" and Trim(fechacancelo) <> "NULL" then
				fechacancelo= "'" & fechacancelo & "'"
			else
				fechacancelo="NULL"
			end if

			correlativo2 = 1

			If Trim(strGestionPago) = "P" Then
				categoria = 1
				subcategoria = 1
				gestion = 3

				strObservacionesGestion = "CAPITAL PAGADO : " & intTotalCapital & ", INTERES : " & intInteres & ", HONORARIOS : " & intHonorarios & ", PROTESTOS : " & intProtestos
				strObservacionesGestion = strObservacionesGestion & " DOCUMENTOS : " & strDocumentos & " OBS : " & strObservaciones


				ssql2="SELECT MAX(CORRELATIVO)+1 AS CORRELATIVO FROM GESTIONES WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='"& strCOD_CLIENTE &"'"
				set rsCOR = Conn.execute(ssql2)
				if not rsCOR.eof then
					CORRELATIVO=rsCOR("CORRELATIVO")
					if isNULL(rsCOR("CORRELATIVO")) THEN
						CORRELATIVO= "1"
					end if
				else
					CORRELATIVO= "1"
				end if
				rsCOR.close
				set rsCOR=nothing

				strSql="INSERT INTO GESTIONES ( RUT_DEUDOR,COD_CLIENTE,NRO_DOC,CORRELATIVO,COD_CATEGORIA,COD_SUB_CATEGORIA,COD_GESTION,FECHA_INGRESO,HORA_INGRESO,ID_USUARIO,FECHA_COMPROMISO,NRO_DOC_PAGO,FECHA_PAGO,OBSERVACIONES,CORRELATIVO_DATO, FECHA_RETIRO,FECHA_AGENDAMIENTO,ID_CAMPANA)"
				strSql= strSql & " VALUES ('"& rut & "','"& strCOD_CLIENTE &"','1','" & CORRELATIVO &"','"& categoria &"','"& subcategoria &"','"& gestion &"','" & date & "','"& Mid(time,1,8) &"',"& session("session_idusuario") & "," & FECHA_COMPROMISO & ",'" & comprobante & "'," & fechacancelo & ",'" & UCASE(strObservacionesGestion) & "','" & Trim(correlativo2) &"','"&   retiro & "," & reagen & "," & intIdCampana & ")"

				'Response.write "<br>strSql = " & strSql
				Conn.execute(strSql)


				strSql = "UPDATE DEUDOR SET ULTIMA_GESTION = '" & categoria &"-"& subcategoria & "-" & gestion & "'"
				strSql = strSql & " WHERE RUT_DEUDOR = '" & rut & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "'"
				'Response.write "<br>strSql = " & strSql
				'REsponse.End
				Conn.execute(strSql)


			End If

			If Trim(strGestionAbono) = "A" Then
				categoria = 1
				subcategoria = 1
				gestion = 4

				strObservacionesGestion = "CAPITAL PAGADO : " & intTotalCapitalA & ", INTERES : " & intInteres & ", HONORARIOS : " & intHonorarios & ", PROTESTOS : " & intProtestos
				strObservacionesGestion = strObservacionesGestion & " DOCUMENTOS : " & strDocumentosA & " OBS : " & strObservaciones


				ssql2="SELECT MAX(CORRELATIVO)+1 AS CORRELATIVO FROM GESTIONES WHERE RUT_DEUDOR='" & rut & "' AND COD_CLIENTE='"& strCOD_CLIENTE &"'"
				set rsCOR = Conn.execute(ssql2)
				if not rsCOR.eof then
					CORRELATIVO=rsCOR("CORRELATIVO")
					if isNULL(rsCOR("CORRELATIVO")) THEN
						CORRELATIVO= "1"
					end if
				else
					CORRELATIVO= "1"
				end if
				rsCOR.close
				set rsCOR=nothing

				strSql="INSERT INTO GESTIONES ( RUT_DEUDOR,COD_CLIENTE,NRO_DOC,CORRELATIVO,COD_CATEGORIA,COD_SUB_CATEGORIA,COD_GESTION,FECHA_INGRESO,HORA_INGRESO,ID_USUARIO,FECHA_COMPROMISO,NRO_DOC_PAGO,FECHA_PAGO,OBSERVACIONES,CORRELATIVO_DATO,FECHA_RETIRO,FECHA_AGENDAMIENTO,ID_CAMPANA)"
				strSql= strSql & " VALUES ('"& rut & "','"& strCOD_CLIENTE &"','1','" & CORRELATIVO &"','"& categoria &"','"& subcategoria &"','"& gestion &"','" & date & "','"& Mid(time,1,8) &"',"& session("session_idusuario") & "," & FECHA_COMPROMISO & ",'" & comprobante & "'," & fechacancelo & ",'" & UCASE(strObservacionesGestion) & "','" & Trim(correlativo2) &"','"&  retiro & "," & reagen & "," & intIdCampana & ")"

				'Response.write "<br>strSql = " & strSql
				Conn.execute(strSql)


				strSql = "UPDATE DEUDOR SET ULTIMA_GESTION = '" & categoria &"-"& subcategoria & "-" & gestion & "'"
				strSql = strSql & " WHERE RUT_DEUDOR = '" & rut & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "'"
				'Response.write "<br>strSql = " & strSql
				'REsponse.End
				Conn.execute(strSql)


			End If




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
		font-family: Arial, Helvetica, sans-serif; 
	}
	-->
	</style>

	<script language="JavaScript " type="text/JavaScript">

	function Refrescar(rut)
	{
		if(rut == '')
		{
			return
		}
				datos.action = "ingreso_pagos.asp?rut=" + rut + "&tipo=1";
				datos.submit();

	}

	</script>


</head>
<body>
<form name="datos" method="post">
<INPUT TYPE="hidden" NAME="intOrigen" value="<%=intOrigen%>">
<div class="titulo_informe">Módulo de Ingreso de Pagos</div>
<table width="1000" border="0" bordercolor="#999999" cellpadding="2" cellspacing="5" align="center">
  <tr>
    <td valign="top">
	  <%

	If rut <> "" then
		strNombreDeudor = TraeNombreDeudor(Conn,strRUT_DEUDOR)
	Else
		strNombreDeudor=""
	End if


	strSql="SELECT FORMULA_HONORARIOS,FORMULA_INTERESES,IsNull(ADIC_1,'ADIC_1') as ADIC_1, IsNull(ADIC_2,'ADIC_2') as ADIC_2, IsNull(ADIC_3,'ADIC_3') as ADIC_3, IsNull(ADIC_4,'ADIC_4') as ADIC_4, IsNull(ADIC_5,'ADIC_5') as ADIC_5, IsNull(ADIC_91,'ADIC_91') as ADIC_91, IsNull(ADIC_92,'ADIC_92') as ADIC_92, IsNull(ADIC_93,'ADIC_93') as ADIC_93, IsNull(ADIC_94,'ADIC_94') as ADIC_94, IsNull(ADIC_95,'ADIC_95') as ADIC_95, USA_CUSTODIO, IsNull(COLOR_CUSTODIO,'FFFFFF') as COLOR_CUSTODIO, INTERES_MORA, COD_TIPODOCUMENTO_HON, MESES_TD_HON FROM CLIENTE WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
	set rsDET=Conn.execute(strSql)
	if Not rsDET.eof Then
		strUsaCustodio = rsDET("USA_CUSTODIO")
		intTasaMensual = ValNulo(rsDET("INTERES_MORA"),"C")
		intTipoDocHono = ValNulo(rsDET("COD_TIPODOCUMENTO_HON"),"C")
		intMesHon = ValNulo(rsDET("MESES_TD_HON"),"C")
		strNomFormHon = ValNulo(rsDET("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsDET("FORMULA_INTERESES"),"C")
	end if
	If intTasaMensual = "" Then
		%>
		<SCRIPT>alert('No se ha definido tasa de interes de mora, se ocupara una tasa del 2%, favor parametrizar')</SCRIPT>
		<%
		intTasaMensual = "2"
		intTipoDocHono = ""
	End If

	%>

	<table width="840" border="0" bordercolor="#FFFFFF" class="estilo_collumnas">
		<thead>
		<tr bordercolor="#999999">
			<td>MANDANTE</td>
			<td>RUT</td>
			<td>NOMBRE O RAZON SOCIAL:</td>
			<td>USUARIO</td>
			<td>SUCURSAL</td>
			<td>FECHA</td>
			<td>&nbsp;</td>
		</tr>
		</thead>
	      <tr bgcolor="#FFFFFF" class="Estilo8">
	      <td>
	      	<select name="CB_CLIENTE">
				<%
					ssql="SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' ORDER BY RAZON_SOCIAL"
					set rsTemp= Conn.execute(ssql)
					if not rsTemp.eof then
						do until rsTemp.eof%>
						<option value="<%=rsTemp("COD_CLIENTE")%>"<%if Trim(strCOD_CLIENTE)=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("RAZON_SOCIAL")%></option>
						<%
						rsTemp.movenext
						loop
					end if
					rsTemp.close
					set rsTemp=nothing
				%>
			</select>
			</td>

			<td ALIGN="LEFT"><input name="TX_RUT" type="text" size="10" maxlength="10" onChange="Refrescar(this.value)" value="<%=rut%>"></td>
			<td><%=strNombreDeudor%><INPUT TYPE="hidden" NAME="rut" value="<%=rut%>"> </td>

			<td ALIGN="RIGHT"><%=session("nombre_user")	%></td>

	        <td><%=nom_sucursal%></td>
	        <td><%=DATE%></td>
	        <td>
				<acronym title="LIMPIAR FORMULARIO">
					<input name="li_" class="fondo_boton_100" type="button" onClick="window.navigate('ingreso_pagos.asp?Limpiar=1');" value="Limpiar">
				</acronym>
			</td>
	      </tr>
    </table>
	</td>
	</tr>
</table>



<table width="90%" border="0" bordercolor="#999999" cellpadding="2" cellspacing="5" align="center">
	<tr>
	<td>

	<table width="100%" border="0" ALIGN="CENTER" class="estilo_columnas">
	<thead>
	 <tr>
	 	<td>
	 	<font >&nbsp;Detalle de Deuda</font>
	 	</td>
	</tr>
	</thead>
	</table>

	<table width="100%" ALIGN="CENTER">
	  <tr>
	    <td valign="top">
		<%
		If Trim(rut) <> "" then
		abrirscg()
			strSql="SELECT dbo." & strNomFormInt & "(ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(ID_CUOTA) as HONORARIOS, ID_CUOTA, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, RUT_DEUDOR, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC, IsNull(datediff(d,FECHA_VENC,getdate()),0) as ANTIGUEDAD, NRO_DOC, IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,isnull(GASTOS_PROTESTOS,0) as GASTOS_PROTESTOS,IsNull(USUARIO_ASIG,0) as USUARIO_ASIG, NRO_CUOTA, SUCURSAL , ESTADO_DEUDA, COD_REMESA, CUENTA, NRO_DOC, TIPO_DOCUMENTO, NOM_TIPO_DOCUMENTO, CUSTODIO FROM CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='"& rut &"' AND COD_CLIENTE='"& strCOD_CLIENTE &"' AND SALDO > 0 AND ESTADO_DEUDA IN ('1','7','8') AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO ORDER BY CUENTA, FECHA_VENC DESC"
			'response.Write(strSql)
			'response.End()
			set rsDET=Conn.execute(strSql)
			if not rsDET.eof then
			%>
			  <table class="intercalado" STYLE="WIDTH:100%">
			  	<thead>
		        <tr >
		          <td >
		          	<a href="#" onClick="marcar_boxes(true);">M</a>&nbsp;&nbsp;&nbsp;
	    			<a href="#" onClick="desmarcar_boxes(true);">D</a>
	    		  </td>
		          <td align="CENTER">NRO. DOC</td>
		          <td align="CENTER">F.VENCIM.</td>
		          <td align="CENTER">TIPO DOC</td>
		          <td align="CENTER">CAPITAL</td>
		          <td align="CENTER">INTERES</td>
		          <td align="CENTER">HONORARIOS</td>
		          <td align="CENTER">PROTESTOS</td>
		          <td align="CENTER">PAGO</td>
		          <td align="CENTER">COMPROBANTE</td>
                </tr>
            	</thead>
            	<tbody>

				<%
				intSaldo = 0
				intValorCuota = 0
				total_ValorCuota = 0
				strArrConcepto = ""
				strArrID_CUOTA = ""

				intTasaMensual = intTasaMensual/100
				intTasaDiaria = intTasaMensual/30

				Do until rsDET.eof

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

				intHonorariosDoc = rsDET("HONORARIOS")
				intInteresDoc = rsDET("INTERESES")
				intProtestoDoc = 0


				%>
		        <tr bordercolor="#999999" >
		          <input TYPE="hidden" name="HD_HONORARIOS_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=intHonorariosDoc%>">
		          <input TYPE="hidden" name="HD_INTERES_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=intInteresDoc%>">
		          <input TYPE="hidden" name="HD_PROTESTOS_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=intProtestoDoc%>">
		          <TD><INPUT TYPE=checkbox NAME="CH_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" onClick="suma_capital(this,TX_SALDO_<%=rsDET("ID_CUOTA")%>.value,<%=Round(intHonorariosDoc,0)%>,<%=Round(intInteresDoc,0)%>,<%=Round(intProtestoDoc,0)%>);";></TD>
		          <!--td><div align="right"><%=rsDET("CUENTA")%></div></td-->
		          <td><div align="right"><%=rsDET("NRO_DOC")%></div></td>
		          <td><div align="right"><%=rsDET("FECHA_VENC")%></div></td>
		          <td><div align="right"><%=rsDET("NOM_TIPO_DOCUMENTO")%></div></td>
		          <td align="right" >$ <%=FN((intSaldo),0)%></td>
		          <td align="right" >$ <%=FN((intInteresDoc),0)%></td>
		          <td align="right" >$ <%=FN((intHonorariosDoc),0)%></td>
		          <td align="right" >$ <%=FN((intProtestoDoc),0)%></td>
		          <td align="right" ><input name="TX_SALDO_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="<%=intSaldo%>" size="10" maxlength="10" align="RIGHT"></td>
		          <td align="right" ><input name="TX_COMPROBANTE_<%=Replace(rsDET("ID_CUOTA"),"-","_")%>" type="text" value="" size="12" maxlength="20" align="RIGHT"></td>
		         <%
					total_ValorCuota = total_ValorCuota + intValorCuota
					total_gc = total_gc + clng(rsDET("GASTOS_PROTESTOS"))
					total_docs = total_docs + 1
				 %>
				 </tr>
				 <%rsDET.movenext
				 loop

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
			<table width="840" border="0" bordercolor="#FFFFFF">
			<tr bordercolor="#999999" class="Estilo8">
			<td align="center">

			Deudor no posee documentos pendientes
			</td>
			</tr>
			</table>
		  <%end if%>
	    </td>
	  </tr>


		<tr>
			<td>
				CAPITAL: <INPUT TYPE="TEXT" NAME="TX_CAPITAL" size="10">
				INTERES: <INPUT TYPE="TEXT" NAME="TX_INTERES" size="10">
				HONORARIOS: <INPUT TYPE="TEXT" NAME="TX_HONORARIOS" size="10">
				PROTESTOS: <INPUT TYPE="TEXT" NAME="TX_PROTESTOS" size="10">
			</TD>
		</tr>


		<tr>
			<td>OBSERVACIONES: <INPUT TYPE="TEXT" NAME="TX_OBSERVACIONES" size="100">

				<% if intOrigen = "IP" Then %>
					<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Guardar" value="Guardar" onClick="envia();" class="Estilo8">
				<% End if %>
				<% if intOrigen = "ID" Then %>
					<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="VerDetalle" value="Ver Detalle" onClick="VerDetalle();" class="Estilo8">
				<% End if %>
			</TD>
		</tr>

	</table>


	</td>
	</tr>



<tr>
<td>

</table>
<INPUT TYPE="hidden" NAME="strGraba" value="">
</form>

</body>
</html>
<script language="JavaScript" type="text/JavaScript">


	function envia(){
		if(datos.TX_RUT.value==''){
			alert("Debe ingresar el rut")
			datos.TX_RUT.focus();
		}
		else
		{

			<% For i=1 TO intTamvConcepto %>
				if ((document.forms[0].<%=vArrConcepto(i)%>.checked == true) && (document.forms[0].TX_COMPROBANTE_<%=vArrID_CUOTA(i)%>.value == '')) {
					alert('Debe Ingresar Comprobante para todos los documentos seleccionados');
					return;
					}
			<% Next %>


			datos.strGraba.value='SI';
			datos.submit();
		}

	}

	function VerDetalle(){
		if(datos.TX_RUT.value==''){
			alert("Debe ingresar el rut")
			datos.TX_RUT.focus();
		}
		else
		{

			<% For i=1 TO intTamvConcepto %>
				if (document.forms[0].<%=vArrConcepto(i)%>.checked == false) {
					intI=IntI+1;
					}
			<% Next %>


			datos.strGraba.value='SI';
			datos.submit();
		}

	}




	function marcar_boxes(){
			datos.TX_CAPITAL.value = 0;
			datos.TX_HONORARIOS.value = 0;
			datos.TX_INTERES.value = 0;
			datos.TX_PROTESTOS.value = 0;
		 	<% For i=1 TO intTamvConcepto %>
				document.forms[0].<%=vArrConcepto(i)%>.checked=true;
				suma_capital(document.forms[0].<%=vArrConcepto(i)%>,document.forms[0].TX_SALDO_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_HONORARIOS_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_INTERES_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_PROTESTOS_<%=vArrID_CUOTA(i)%>.value);
			<% Next %>
	}

	function desmarcar_boxes(){
			<% For i=1 TO intTamvConcepto %>
				document.forms[0].<%=vArrConcepto(i)%>.checked=false;
				suma_capital(document.forms[0].<%=vArrConcepto(i)%>,document.forms[0].TX_SALDO_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_HONORARIOS_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_INTERES_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_PROTESTOS_<%=vArrID_CUOTA(i)%>.value);
			<% Next %>
			datos.TX_CAPITAL.value = 0;
			datos.TX_HONORARIOS.value = 0;
			datos.TX_INTERES.value = 0;
			datos.TX_PROTESTOS.value = 0;

	}

	function suma_capital(objeto , intValorSaldoCapital, intValorHonorarios, intValorIntereses, intValorProtestos){
			//alert(objeto.checked);

			if (datos.TX_CAPITAL.value == '') datos.TX_CAPITAL.value = 0
			if (datos.TX_HONORARIOS.value == '') datos.TX_HONORARIOS.value = 0
			if (datos.TX_INTERES.value == '') datos.TX_INTERES.value = 0
			if (datos.TX_PROTESTOS.value == '') datos.TX_PROTESTOS.value = 0

			if (objeto.checked == true) {
				datos.TX_CAPITAL.value = eval(datos.TX_CAPITAL.value) + eval(intValorSaldoCapital);
				datos.TX_HONORARIOS.value = eval(datos.TX_HONORARIOS.value) + eval(intValorHonorarios);
				datos.TX_INTERES.value = eval(datos.TX_INTERES.value) + eval(intValorIntereses);
				datos.TX_PROTESTOS.value = eval(datos.TX_PROTESTOS.value) + eval(intValorProtestos);
			}
			else
			{
				datos.TX_CAPITAL.value = eval(datos.TX_CAPITAL.value) - eval(intValorSaldoCapital);
				datos.TX_HONORARIOS.value = eval(datos.TX_HONORARIOS.value) - eval(intValorHonorarios);
				datos.TX_INTERES.value = eval(datos.TX_INTERES.value) - eval(intValorIntereses);
				datos.TX_PROTESTOS.value = eval(datos.TX_PROTESTOS.value) - eval(intValorProtestos);
				}
		}
</script>


















