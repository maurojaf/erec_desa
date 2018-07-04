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
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	strOrigen = Request("Origen")
	
	If Trim(Request("Limpiar"))="1" Then
		session("session_RUT_DEUDOR") = ""
		rut = ""
	End if

	If Trim(Request("TX_RUT")) = "" Then
		strRUT_DEUDOR = session("session_RUT_DEUDOR")
	Else
		strRUT_DEUDOR = Trim(Request("TX_RUT"))
		session("session_RUT_DEUDOR") = strRUT_DEUDOR
	End If

	intOrigen = Request("intOrigen")

	intTipoPP = Request("CB_TIPO")

	If Trim(intTipoPP) = "RP" or Trim(intTipoPP) = "RC" or Trim(intTipoPP) = "RL" Then
		intTipoPP = "CONV"
	End If

	''Response.write "intTipoPP=" & intTipoPP

	intCliente = session("ses_codcli")

	strSede = Request("CB_SEDE")
	strCiudadSede = strSede

	intPorcDescCapital = ValNulo(Request("desc_capital"),"N")
	intPorcDescIndemComp = ValNulo(Request("desc_indemComp"),"N")
	intPorcDescHonorarios = ValNulo(Request("desc_honorarios"),"N")
	intPorcDescGastos = ValNulo(Request("desc_gastos"),"N")
	intPorcDescInteres = ValNulo(Request("desc_interes"),"N")
	intPorcDescProtestos = ValNulo(Request("desc_protestos"),"N")


	intPie = Request("pie")
	intCuotas = Request("cuotas")
	
	if Trim(Request("TX_DIAPAGO")) = "" then
		intDiaDePago = Request("PP_TX_DIAPAGO")
	else
		intDiaDePago = Request("TX_DIAPAGO")
	end if

	If Trim(Request("TX_CAPITAL")) = "" then 
		originalCapital = Request("PP_TX_CAPITAL")
	else
		originalCapital = Request("TX_CAPITAL")
	end if
	
	If Trim(Request("TX_INTERES")) = "" then 
		originalIntereses = Request("PP_TX_INTERES")
	else
		originalIntereses = Request("TX_INTERES")
	end if
	
	If Trim(Request("TX_GASTOSPROTESTOS")) = "" then 
		originalProtestos = Request("PP_TX_GASTOSPROTESTOS")
	else
		originalProtestos = Request("TX_GASTOSPROTESTOS")
	end if
	
	If Trim(Request("TX_HONORARIOS")) = "" then 
		originalHonorarios = Request("PP_TX_HONORARIOS")
	else
		originalHonorarios = Request("TX_HONORARIOS")
	end if
	
	intOriginalCapital = Round(ValNulo(originalCapital,"N"),0)
	intOriginalIndemComp = Round(ValNulo(Request("hdintIndemComp"),"N"),0)
	intOriginalIntereses = Round(ValNulo(originalIntereses,"N"),0)
	intOriginalProtestos = Round(ValNulo(originalProtestos,"N"),0)
	intOriginalHonorarios = Round(ValNulo(originalHonorarios,"N"),0)

	intOriginalGastos = Round(ValNulo(Request("hdintGastos"),"N"),0)

	intTotalDeuda = intOriginalCapital + intOriginalIntereses + intOriginalProtestos + intOriginalHonorarios

	intDescCapital = Round(intPorcDescCapital,0)
	intDescIndemComp = Round(intPorcDescIndemComp,0)
	intDescHonorarios = Round(intPorcDescHonorarios,0)
	intDescGastos = Round(intPorcDescGastos,0)
	intDescInteres = Round(intPorcDescInteres,0)
	intDescProtestos = Round(intPorcDescProtestos,0)

	intTotalDescuentos = intDescCapital + intDescIndemComp + intDescHonorarios + intDescGastos + intDescInteres
	'intTotalDescuentos = intDescCapital + intDescGastos

	intTotalConvenio = intTotalDeuda - intTotalDescuentos

	intTotDeudaCapital = Round(intOriginalCapital - intDescCapital,0)
	intTotIndemComp = Round(intOriginalIndemComp - intDescIndemComp,0)
	intTotHonorarios = Round(intOriginalHonorarios - intDescHonorarios,0)
	intTotGastos = Round(intOriginalGastos - intDescGastos,0)
	intTotIntereses = Round(intOriginalIntereses - intDescInteres,0)

	intTotProtestos = Round(intOriginalProtestos - intDescProtestos,0)

	AbrirSCG()

	strSql = "SELECT S.RAZON_SOCIAL, S.DIRECCION, S.COMUNA, S.RUT,S.TELEFONO, FOLIO_ACTUAL FROM SEDE S, CONVENIO_CORRELATIVO C WHERE S.COD_CLIENTE = '" & intCliente & "' AND S.SEDE = '" & strSede & "'"
	strSql = strSql & " AND S.RUT = C.RUT AND S.COD_CLIENTE = C.COD_CLIENTE"
	'Response.write "strSql=" & strSql
	set rsSede = Conn.execute(strSql)
	if not rsSede.eof then
		strRazonSocialSede=rsSede("RAZON_SOCIAL")
		strDireccionSede=rsSede("DIRECCION")
		strDireccionComuna=rsSede("COMUNA")
		strRutSede=rsSede("RUT")
		strTelefonoSede=rsSede("TELEFONO")
		strNroFolio=rsSede("FOLIO_ACTUAL")
	Else
		%>
		<SCRIPT>
			alert('Parámetros del convenio no han sido configurados (SEDE, CORRELATIVOS), revise configuración.');
			history.back();
		</SCRIPT>
		<%
		Response.End
	End If


	strSql=""
	strSql="SELECT USA_SUBCLIENTE,USA_INTERESES,USA_HONORARIOS,USA_PROTESTOS,FORMULA_HONORARIOS,FORMULA_INTERESES,TASA_MAX_CONV, DESCRIPCION, TIPO_INTERES FROM CLIENTE WHERE COD_CLIENTE ='" & intCliente & "'"
	set rsTasa=Conn.execute(strSql)
	if not rsTasa.eof then
		intTasaMax = rsTasa("TASA_MAX_CONV")
		strDescripcion = rsTasa("DESCRIPCION")
		strTipoInteres = rsTasa("TIPO_INTERES")
		strNomFormHon = ValNulo(rsTasa("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsTasa("FORMULA_INTERESES"),"C")

		strUsaSubCliente = rsTasa("USA_SUBCLIENTE")
		strUsaInteres = rsTasa("USA_INTERESES")
		strUsaHonorarios = rsTasa("USA_HONORARIOS")
		strUsaProtestos = rsTasa("USA_PROTESTOS")

	Else
		intTasaMax = 1
		strDescripcion = ""
		strTipoInteres = ""
	end if
	rsTasa.close
	set rsTasa=nothing



		strSql="SELECT NOMBRE_DEUDOR FROM DEUDOR WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & intCliente & "'"
		set RsDeudor=Conn.execute(strSql)
		if not RsDeudor.eof then
			strNombreDeudor = RsDeudor("NOMBRE_DEUDOR")
		end if
		RsDeudor.close
		set RsDeudor=nothing

		strSql=""
		strSql="SELECT TOP 1 CALLE,NUMERO,COMUNA,RESTO,CORRELATIVO,ESTADO FROM DEUDOR_DIRECCION WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND ESTADO <> '2' ORDER BY ESTADO DESC"
		set rsDIR=Conn.execute(strSql)
		if not rsDIR.eof then
			calle_deudor=rsDIR("Calle")
			numero_deudor=rsDIR("Numero")
			comuna_deudor=rsDIR("Comuna")
			resto_deudor=rsDIR("Resto")
			strDirDeudor = calle_deudor & " " & numero_deudor & " " & resto_deudor & " " & comuna_deudor
		end if
		rsDIR.close
		set rsDIR=nothing


		strSql=""
		strSql="SELECT TOP 1 COD_AREA,TELEFONO,CORRELATIVO,ESTADO FROM DEUDOR_TELEFONO WHERE  RUT_DEUDOR='" & strRUT_DEUDOR & "' AND ESTADO <> '2' ORDER BY ESTADO DESC"
		set rsFON=Conn.execute(strSql)
		if not rsFON.eof then
			codarea_deudor = rsFON("COD_AREA")
			Telefono_deudor = rsFON("Telefono")
			strFonoDeudor = codarea_deudor & "-" & Telefono_deudor
		end if
		rsFON.close
		set rsFON=nothing


		strSql=""
		strSql="SELECT TOP 1 RUT_DEUDOR,CORRELATIVO,FECHA_INGRESO,EMAIL,ESTADO FROM DEUDOR_EMAIL WHERE  RUT_DEUDOR='" & strRUT_DEUDOR & "' AND ESTADO <> '2' ORDER BY ESTADO DESC"
		set rsMAIL=Conn.execute(strSql)
		if not rsMAIL.eof then
			strEmail = rsMAIL("EMAIL")
		end if
		rsMAIL.close
		set rsMAIL=nothing


		If intTipoPP = "CONV" Then

			intKapitalInicial = intTotalConvenio-intPie


			If Trim(strTipoInteres)="C" Then

				intMontoConInteres = intKapitalInicial * calcula_base_exponente((1 + intTasaMax/100),intCuotas)



				if intTasaMax = "0" Then
					intValorCuota=intKapitalInicial/intCuotas
				Else
					intValorCuota=intMontoConInteres/intCuotas
				End If

				intValorCuota = Round(intValorCuota,0)



			Else
				intMontoConInteres = intKapitalInicial * (1 + ((intTasaMax/100)*intCuotas))
				if intTasaMax = "0" Then
					intValorCuota=intKapitalInicial/intCuotas
				Else
					intValorCuota=intMontoConInteres/intCuotas
				End if
				intValorCuota = Round(intValorCuota,0)
			End If


			intMonto = intValorCuota


				intMesDePago = Month(date)
				intAnoDePago = Year(date)

				'intValorCuota = intMonto


				intMontoTotal = 0
				For i=1 to intCuotas

					intTotalGastos=0
					intCont=1

					intMesDePago = intMesDePago + 1
					If intMesDePago = 13 Then
						intMesDePago = 1
						intAnoDePago = intAnoDePago + 1
					End if
					intCont = intCont + 1

					dtmFechaPago = PoneIzq(intDiaDePago,"0") & "/" & PoneIzq(intMesDePago,"0") & "/" & intAnoDePago
					intNroCuota = i

					If i=1 Then
						dtmFechaCuota1 = dtmFechaPago
					End If

					If Not Isnull(intMontoConInteres/intCuotas) Then
						intMonto = Round(intMontoConInteres/intCuotas,0)
						intMonto= intValorCuota
					End if


					If Mid(dtmFechaPago,4,2) = "02" and Cdbl(intDiaDePago) > 28 Then
						dtmFechaPago = "28/02/" & Mid(dtmFechaPago,7,4)
					End if

					If Cdbl(intDiaDePago) > 30 Then
						dtmFechaPago = "30/" & Mid(dtmFechaPago,4,2) & "/" & Mid(dtmFechaPago,7,4)
					End if
					intMontoTotal = intMontoTotal + intMonto

				Next
				dtmFechaCuotaFin = dtmFechaPago
		End If


			If Trim(Request("CB_TIPO")) = "RC" or Trim(Request("CB_TIPO")) = "RL" or Trim(Request("CB_TIPO")) = "RP" Then
				strPie = "pie"
			End If

			strFormaPago = TraeCampoId2(Conn, "DESC_FORMA_PAGO", Trim(Request("CB_FPAGO")), "CAJA_FORMA_PAGO", "ID_FORMA_PAGO")
			If Trim(strFormaPago) = "" Then strFormaPago = "NO ESPECIFICADO"

%>

<title>Plan de Pago</title>
<style type="text/css">
<!--
.Estilo1 {font-size: 14px;font-weight: bold;}
.Estilo13 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; }
.Estilo2 {font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif;}
.Estilo5 {font-size: 11px; font-weight: bold;}
.Estilo8 {font-size: 11px}
.Estilo9 {font-size: 11px; font-family: Verdana, Arial, Helvetica, sans-serif; font-weight: bold; }
.Estilo12 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 12px; font-weight: bold; }
.Estilo14 {font-size: 10px}
.Estilo15 {font-family: Verdana, Arial, Helvetica, sans-serif; font-size: 10px; font-weight: bold; }
.Estilo26 {font-size: 12px}
.Estilo28 {font-family: "Courier New", Courier, monospace}
.Estilo32 {font-size: 13px}
.Estilo22 {font-size: 13px; font-family: "Courier New", Courier, monospace;}
.Estilo36 {font-family: Verdana, Arial, Helvetica, sans-serif;font-weight: bold;}
.Estilo37 {font-family: Verdana, Arial, Helvetica, sans-serif; }
.Estilo38 {font-size: 13px; font-family: Verdana, Arial, Helvetica, sans-serif; }
.Estilo40 {	font-family: Verdana, Arial, Helvetica, sans-serif;	color: #FF0000;	font-size: 11px;font-weight: bold;}
.Estilo41 {color: #FFFFFF}
.Estilo33 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 14px}
.Estilo34 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px}

H1.SaltoDePagina {PAGE-BREAK-AFTER: always}
	.transpa {
	background-color: transparent;
	border: 1px solid #FFFFFF;
	text-align:center
	}
-->

</style>



<script language="JavaScript " type="text/JavaScript">

function Grabar()
{
	if (confirm("¿ Está seguro de grabar el <%=session("NOMBRE_CONV_PAGARE")%> ? Si aun no lo ha imprimido, presione cancelar y luego el botón imprimir."))
		{
			datos.action = "grabar_convenio.asp";
			datos.submit();
		}
}
</script>

</head>
<BODY>
<FORM NAME="datos" METHOD="POST" ACTION="grabar_convenio.asp">

<INPUT TYPE="hidden" NAME="strCOD_CLIENTE" value="<%=intCliente%>">
<INPUT TYPE="hidden" NAME="strRUT_DEUDOR" value="<%=strRUT_DEUDOR%>">
<INPUT TYPE="hidden" NAME="intTotalConvenio" value="<%=intTotalConvenio%>">
<INPUT TYPE="hidden" NAME="intTotalCapital" value="<%=intTotDeudaCapital%>">
<INPUT TYPE="hidden" NAME="intIndemComp" value="<%=intTotIndemComp%>">
<INPUT TYPE="hidden" NAME="intGastos" value="<%=intTotGastos%>">
<INPUT TYPE="hidden" NAME="intHonorarios" value="<%=intTotHonorarios%>">
<INPUT TYPE="hidden" NAME="intDescTotalCapital" value="<%=intDescCapital%>">
<INPUT TYPE="hidden" NAME="intDescIndemComp" value="<%=intDescIndemComp%>">
<INPUT TYPE="hidden" NAME="intDescGastos" value="<%=intDescGastos%>">
<INPUT TYPE="hidden" NAME="intDescHonorarios" value="<%=intDescHonorarios%>">
<INPUT TYPE="hidden" NAME="intPie" value="<%=intPie%>">
<INPUT TYPE="hidden" NAME="intCuotas" value="<%=intCuotas%>">
<INPUT TYPE="hidden" NAME="intDiaPago" value="<%=intDiaDePago%>">
<INPUT TYPE="hidden" NAME="strObservaciones" value="<%=strObservaciones%>">
<INPUT TYPE="hidden" NAME="strRutSede" value="<%=strRutSede%>">



<TABLE ALIGN="CENTER" WIDTH="800" BORDER="1" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

		<TR>
		  <TD>
		  <table width="750" border="0" cellspacing=0 cellpadding=0>
			<tr>
			  <td class="Estilo13" width=400><b><%=strRazonSocialSede%></b></td>
			  <td width=100>&nbsp;</td>
			  <td width=150 rowspan="3"  class="Estilo13" align="center">
			  <!--img border="0" width="150" height="46" src="UploadFolder/<%=session("ses_codcli")%>/logo.jpg"-->
			  <img border="0" src="../imagenes/Logos/<%=session("ses_codcli")%>/logo.jpg">
			  </td>
			</tr>
			<tr>
			  <td class="Estilo13">Dirección :<%=strDireccionSede%>&nbsp;&nbsp;<%=strDireccionComuna%></td>
			  <td>&nbsp;</td>
			</tr>
			<tr>
			  <td class="Estilo13">R.U.T: :<%=strRutSede%></td>
			  <td>&nbsp;</td>
			</tr>
			<tr>
			  <td class="Estilo13">Teléfono : <%=strTelefonoSede%></td>
			  <td>&nbsp;</td>
			  <td align="center">&nbsp;</td>
			</tr>
		  </table>
		  </TD>
 		</TR>

 		<TR>
		 	<TD>

		 		<TABLE ALIGN="CENTER" WIDTH="750" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

					<tr class="Estilo1">
						<TD align="center" class="Estilo33" colspan=2>
							PLAN DE PAGO<br><br>
						</TD>
				 		</TR>
				 		<TR class="Estilo1">
							<TD align="left" class="Estilo34" style="vertical-align: top;">
								Nombre : <%=strNombreDeudor%><BR>
								R.U.T. : <%=strRUT_DEUDOR%><BR>
								Fecha - hora : <%=now()%><BR><BR>
 							</TD>
 							<TD align="left" class="Estilo34" style="vertical-align: top;">
							  Tipo : <%= TraeCampoId2(Conn, "NOM_TIPO_PLAN_PAGO", Trim(Request("CB_TIPO")), "TIPO_PLAN_PAGO", "COD_TIPO_PLAN_PAGO") %><BR>
							  Forma de pago <%=strPie%>: <%=strFormaPago%><BR>
							  <!--Validez hasta el : <%=dateadd("d",5,date())%><BR><BR>-->

					 		</TD>
				 		</TR>
					 		<tr class="Estilo1">
				 		</TR>
				</TABLE>

	 		</TD>
 		</TR>


 		<TR>
			<TD>
				<TABLE ALIGN="CENTER" WIDTH="750" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

					<tr class="Estilo1">
						<TD colspan=4 align="LEFT" class="Estilo34">
						  <br>ANTECEDENTES DE LA DEUDA
						</TD>
					</TR>
				</TABLE>


				<table width="100%" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
				<tr class="Estilo34">
					<%If Trim(strUsaSubCliente)="1" Then%>
						<td>RUT CLIENTE</td>
						<td>NOMBRE CLIENTE</td>
					<%End If%>
					<td>NºDOC</td>
					<td>CUOTA</td>
					<td>FEC.VENC.</td>
					<td>ANT.</td>
					<td>TIPO DOC.</td>

					<!--td>ASIGNACION</td-->
					<td>CAPITAL</td>
					<%If Trim(strUsaInteres)="1" Then%>
					<td>INTERES</td>
					<%End If%>
					<%If Trim(strUsaProtestos)="1" Then%>
					<td>PROTESTOS</td>
					<%End If%>
					<%If Trim(strUsaHonorarios)="1" Then%>
					<td>HONORARIOS</td>
					<%End If%>

					<td>ABONO</td>
					<td>TOTAL</td>
				</tr>

				<%

				strSql = "SELECT RUT_SUBCLIENTE, NOMBRE_SUBCLIENTE, VALOR_CUOTA, dbo." & strNomFormInt & "(ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(ID_CUOTA) as HONORARIOS, ID_CUOTA, NRO_DOC, NRO_CUOTA, NOM_TIPO_DOCUMENTO AS TIPO_DOCUMENTO, GASTOS_PROTESTOS, CUENTA, FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD,DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, CUSTODIO, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS ,SALDO"
				strSql = strSql & " FROM CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND COD_CLIENTE='" & intCliente & "' AND SALDO > 0 AND ESTADO_DEUDA IN (SELECT ESTADO_DEUDA FROM ESTADO_DEUDA WHERE ACTIVO = 1) AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO"
				strSql = strSql & " ORDER BY CUOTA.FECHA_VENC ASC"

				set rsTemp= Conn.execute(strSql)
				''Response.write "strSql=" & strSql

				intTasaMensual = 2/100
				intTasaDiaria = intTasaMensual/30
				intCorrelativo = 1
				strArrID_CUOTA=""
				intTotSelSaldo= 0
				intTotSelCapital= 0
				intTotSelIntereses= 0
				intTotSelProtestos= 0
				intTotSelHonorarios= 0
				Do until rsTemp.eof
				
					strObjeto = "CH_" & rsTemp("ID_CUOTA")
					strObjeto1 = "TX_SALDO_" & rsTemp("ID_CUOTA")

					If UCASE(Request(strObjeto)) = "ON" Then

						'intSaldo = Request(strObjeto1)
                        intSaldo = Round(session("valor_moneda") * ValNulo(rsTemp("SALDO"),"N"),0)
						strNroDoc = rsTemp("NRO_DOC")
						strFechaVenc = rsTemp("FECHA_VENC")
						strTipoDoc = rsTemp("TIPO_DOCUMENTO")

						strNroCuota = rsTemp("NRO_CUOTA")

						intAntiguedad = ValNulo(rsTemp("ANTIGUEDAD"),"N")

						intIntereses = rsTemp("INTERESES")
						intHonorarios = rsTemp("HONORARIOS")


						intValorCapital = rsTemp("VALOR_CUOTA")
						intAbono = intValorCapital - intSaldo 

						'Response.write "intHonorarios=" & intHonorarios

						intProtestos = ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")
						strArrID_CUOTA = strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")

						intTotDoc= (intSaldo+intIntereses+intProtestos+intHonorarios)

						intTotSelSaldo= intTotSelSaldo+intSaldo
						intTotSelIntereses= intTotSelIntereses+intIntereses
						intTotSelProtestos= intTotSelProtestos+intProtestos
						intTotSelHonorarios= intTotSelHonorarios+intHonorarios

						intTotSelValorAbono= intTotSelValorAbono+intAbono
						intTotSelDoc = intTotSelDoc+intTotDoc
						intTotSelCapital = intTotSelCapital + intValorCapital


						%>
						<tr class="Estilo34">
						<%If Trim(strUsaSubCliente)="1" Then%>
							<td><%=rsTemp("RUT_SUBCLIENTE")%></td>
							<td><%=rsTemp("NOMBRE_SUBCLIENTE")%></td>
						<%End If%>
						<td><%=strNroDoc%></td>
						<td><%=strNroCuota%></td>
						<td><%=strFechaVenc%></td>
						<td><%=intAntiguedad%></td>
						<td><%=strTipoDoc%></td>
						<td ALIGN="RIGHT"><%=FN(intValorCapital,0)%></td>

						<%If Trim(strUsaInteres)="1" Then%>
						<td ALIGN="RIGHT"><%=FN(intIntereses,0)%></td>
						<%End If%>
						<%If Trim(strUsaProtestos)="1" Then%>
						<td ALIGN="RIGHT"><%=FN(intProtestos,0)%></td>
						<%End If%>
						<%If Trim(strUsaHonorarios)="1" Then%>
						<td ALIGN="RIGHT"><%=FN(intHonorarios,0)%></td>
						<%End If%>
						<td ALIGN="RIGHT"><%=FN(intAbono,0)%></td>
						<td ALIGN="RIGHT"><%=FN(intTotDoc,0)%></td>
						</tr>
						<%
					End if


				rsTemp.movenext
				intCorrelativo = intCorrelativo + 1
				loop
				rsTemp.close
				set rsTemp=nothing

				strArrID_CUOTA = Mid(strArrID_CUOTA,2,len(strArrID_CUOTA))
		%>


			<tr class="Estilo34">
				<%If Trim(strUsaSubCliente)="1" Then%>
					<td colspan=2>&nbsp;</td>
				<%End If%>
				<td colspan = 5>Totales</td>
				<td ALIGN="RIGHT"><%=FN(intTotSelCapital,0)%></td>

				<%If Trim(strUsaInteres)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelIntereses,0)%></td>
				<%End If%>
				<%If Trim(strUsaProtestos)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelProtestos,0)%></td>
				<%End If%>
				<%If Trim(strUsaHonorarios)="1" Then%>
				<td ALIGN="RIGHT"><%=FN(intTotSelHonorarios,0)%></td>
				<%End If%>

				<td ALIGN="RIGHT"><%=FN(intTotSelValorAbono,0)%></td>
				<td ALIGN="RIGHT"><%=FN(intTotSelDoc,0)%></td>
			</tr>

			<INPUT TYPE="HIDDEN" NAME="strArrID_CUOTA" VALUE="<%=strArrID_CUOTA%>">
			</table>

			<br>
					<TABLE ALIGN="CENTER" WIDTH="750" BORDER="0">
						<TR class="Estilo1">
							<TD colspan=4 align="CENTER" class="Estilo38">
							  DETALLE<br><br>
							</TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD ALIGN="RIGHT"><b>Conceptos</b></TD><TD align="right"><b>Monto Original</b></TD>
							<TD align="right"><b>Descuentos</b></TD><TD align="right"><b>Total</b></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD ALIGN="RIGHT">Capital: </TD><TD align="right">$ <%=FN(intOriginalCapital,0)%></TD>
							<TD align="right">$ <%=FN(intDescCapital,0)%></TD><TD align="right">$ <%=FN(intTotDeudaCapital,0)%></TD>
						</TR>

						<%If Trim(strUsaInteres)="1" Then%>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right">Intereses: </TD><TD align="right">$ <%=FN(intOriginalIntereses,0)%></TD>
							<TD align="right">$ <%=FN(intDescInteres,0)%></TD><TD align="right">$ <%=FN(intTotIntereses,0)%></TD>
						</TR>
						<%End If%>
						<%If Trim(strUsaProtestos)="1" Then%>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right">Protestos: </TD><TD align="right">$ <%=FN(intOriginalProtestos,0)%></TD>
							<TD align="right">$ <%=FN(intDescProtestos,0)%></TD><TD align="right">$ <%=FN(intTotProtestos,0)%></TD>
						</TR>
						<%End If%>
						<%If Trim(strUsaHonorarios)="1" Then%>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right">Honorarios: </TD><TD align="right">$ <%=FN(intOriginalHonorarios,0)%></TD>
							<TD align="right">$ <%=FN(intDescHonorarios,0)%></TD><TD align="right">$ <%=FN(intTotHonorarios,0)%></TD>
						</TR>
						<%End If%>
						<TR HEIGHT=15>
							<TD colspan=3>&nbsp</TD><TD align="right">______________</TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>Total Deuda:</TD><TD align="right"><strong>$ <%=FN(intTotalConvenio,0)%></strong></TD>
						</TR>

						<% If intTipoPP = "CONV" Then	%>

						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>Pie :</TD><TD align="right"><strong>-  $ <%=FN(intPie,0)%></strong></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD colspan=3>&nbsp</TD><TD align="right">______________</TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>Saldo:</TD><TD align="right"><strong>$ <%=FN((intTotalConvenio -  intPie),0)%></strong></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>Intereses.:</TD>	<TD align="right"><strong>$ <%=FN(intMontoTotal - (intTotalConvenio -  intPie),0)%></strong></TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>&nbsp</TD><TD align="right" class="Estilo38">&nbsp</TD>
						</TR>
						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3>&nbsp</TD><TD align="right" class="Estilo38">&nbsp</TD>
						</TR>

						<TR HEIGHT=15 class="Estilo38">
							<TD align="right" colspan=3><b>Saldo : <b></TD>
							<TD align="right"><strong>$ <%=FN(intMontoTotal,0)%></strong></TD>
						</TR>

						<% End If %>
					</table>

				<br>

				<% If intTipoPP = "CONV" Then	%>
				<TR>
				 		<TD class="Estilo38" align="CENTER">
				 			<b>DETALLE VENCIMIENTOS</b>
				 		</TD>
				</TR>
				 	 	<TR>
				 		<TD>
				 		<TABLE ALIGN= "CENTER" WIDTH="70%" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
				 						<%



										intMesDePago = Month(date)
										intAnoDePago = Year(date)

										intKapitalInicial = intTotalConvenio-intPie
										'M= C ( 1+it ) M= 3,250(1 + (0.25)(1.5/12)= 3,351.56

										'strTipoInteres = "C"

										If Trim(strTipoInteres)="C" Then
											intMontoConInteres = intKapitalInicial * calcula_base_exponente((1 + intTasaMax/100),intCuotas)
										Else
											intMontoConInteres = intKapitalInicial * (1 + ((intTasaMax/100)*intCuotas))
										End If

										For i=1 to intCuotas

										intTotalGastos=0
										intCont=1


										intMesDePago = intMesDePago + 1
										If intMesDePago = 13 Then
											intMesDePago = 1
											intAnoDePago = intAnoDePago + 1
										End if
										intCont = intCont + 1


										dtmFechaPago = PoneIzq(intDiaDePago,"0") & "/" & PoneIzq(intMesDePago,"0") & "/" & intAnoDePago
										intNroCuota = i

										'intMonto = Round(intKapitalInicial/intCuotas,0)
										If Not Isnull(intMontoConInteres/intCuotas) Then
											intMonto = Round(intMontoConInteres/intCuotas,0)
											intMonto= intValorCuota
										End if

										If Mid(dtmFechaPago,4,2) = "02" and Cdbl(intDiaDePago) > 28 Then
											dtmFechaPago = "28/02/" & Mid(dtmFechaPago,7,4)
										End if

										If Cdbl(intDiaDePago) > 30 Then
											dtmFechaPago = "30/" & Mid(dtmFechaPago,4,2) & "/" & Mid(dtmFechaPago,7,4)
										End if




										%>
										<tr class="Estilo34">
											<td>Vencimiento <%=intNroCuota%></td>
											<td><%=dtmFechaPago%></td>
											<td align="right">$ <%=FN(intMonto,0)%></td>
										</tr>

										<% Next %>
										<tr class="Estilo34">
											<td>&nbsp;</td>
											<td><b>TOTAL</b></td>
											<td align="right"><b>$ <%=FN(intMontoTotal,0)%></b></td>
										</tr>
								</TABLE>
							</TD>
				 		</TR>
<% End If %>
			<TR>
		 		<TD>
					<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0">
						<TR class="Estilo1">

							<TD align="LEFT" class="Estilo38" width="150">
								<br>
								OBSERVACIONES :
								<br>
							</TD>
							<TD align="LEFT" class="Estilo38">
							<br>
							  <% if Trim(Request("TA_OBSERVACIONES")) = "" then %>
							  SIN OBSERVACIONES
							  <% else 
								Response.Write(Request("TA_OBSERVACIONES"))
								
								end if
							  %>
							  <br>
							</TD>
						</TR>
					</TABLE>
				</TD>
			</TR>
 </TABLE>


<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
			<TR class="Estilo1">

					<TD align="LEFT" class="Estilo34">
						<% if strOrigen = "ingreso_gestion" then %>
						&nbsp;
						<% else %>
						<acronym title="Volver a plan de pagos">
							<input name="BT_PLANPAGO" class="boton_azul" type="button" id="BT_PLANPAGO" onClick="envia_plandepago();" value="Volver a Plan Pagos">
						</acronym>
						<% end if %>
				</TD>



				<TD align="RIGHT" class="Estilo34">
					<acronym title="IMPRIMIR <%=UCASE(session("NOMBRE_CONV_PAGARE"))%>">
						<input name="BT_IMPRIMIR" type="button" onClick="window.print();" value="            Imprimir              ">
					</acronym>
				</TD>

			</TR>
	</TABLE>

	<INPUT TYPE="hidden" NAME="intIntConvenio" value="<%=intMontoTotal%>">

</FORM>
&nbsp;&nbsp;
</body>
<script>
function envia_plandepago(){
datos.action='simulacion_convenio.asp?intOrigen=PP';
datos.submit();
}
</script>
</html>