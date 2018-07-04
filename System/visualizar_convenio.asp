<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

<!--#include file="arch_utils.asp"-->

<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	
	intIdConvenio = Request("intIdConvenio")

	If Trim(Request("TX_RUT")) = "" Then
		strRUT_DEUDOR = session("session_RUT_DEUDOR")
	Else
		strRUT_DEUDOR = Trim(Request("TX_RUT"))
		session("session_RUT_DEUDOR") = strRUT_DEUDOR
	End If

	intCliente = session("ses_codcli")

	AbrirSCG()

	strSql="SELECT * FROM CONVENIO_ENC WHERE ID_CONVENIO = " & intIdConvenio
	set rsConvenio=Conn.execute(strSql)
	if not rsConvenio.eof then
		strIdCliente = rsConvenio("COD_CLIENTE")
		strSede = rsConvenio("SEDE")
		strCiudadSede = strSede
		strRUT_DEUDOR = rsConvenio("RUT_DEUDOR")
		strUsrIngreso = rsConvenio("USR_INGRESO")
		strFecha = rsConvenio("FECHA_INGRESO")
		intCuotas = rsConvenio("CUOTAS")
		intDiaPago = rsConvenio("DIA_PAGO")
		strFolio = rsConvenio("FOLIO")


		intTotalConvenio = rsConvenio("TOTAL_CONVENIO")
		intCapital = rsConvenio("CAPITAL")
		intIntereses = rsConvenio("INTERESES")
		intGastos = rsConvenio("GASTOS")
		intTProtestos = rsConvenio("PROTESTOS")
		intIndemComp= rsConvenio("IC")
		intHonorarios = rsConvenio("HONORARIOS")
		intDescCapital = rsConvenio("DESC_CAPITAL")
		intDescIntereses = rsConvenio("DESC_INTERESES")
		intDescGastos = rsConvenio("DESC_GASTOS")
		intDescProtestos = rsConvenio("DESC_PROTESTOS")
		intDescIndemComp = ValNulo(rsConvenio("DESC_IC"),"N")
		intDescHonorarios = rsConvenio("DESC_HONORARIOS")
		intPie = rsConvenio("PIE")


		strSql="SELECT NOMBRE_DEUDOR FROM DEUDOR WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND COD_CLIENTE = '" & strIdCliente & "'"
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


	End If

	strSql="SELECT TASA_MAX_CONV, DESCRIPCION, TIPO_INTERES FROM CLIENTE WHERE COD_CLIENTE ='" & strIdCliente & "'"
	set rsTasa=Conn.execute(strSql)
	if not rsTasa.eof then
		intTasaMax = rsTasa("TASA_MAX_CONV")
		strDescripcion = rsTasa("DESCRIPCION")
		strTipoInteres = rsTasa("TIPO_INTERES")
	Else
		intTasaMax = 1
		strDescripcion = ""
		strTipoInteres = ""
	end if
	rsTasa.close
	set rsTasa=nothing


	strSql = "SELECT * FROM SEDE S, CONVENIO_CORRELATIVO C WHERE S.COD_CLIENTE = '" & strIdCliente & "' AND S.SEDE = '" & strSede & "'"
	strSql = strSql & " AND S.RUT = C.RUT AND S.COD_CLIENTE = C.COD_CLIENTE"
	'Response.write "strSql=" & strSql
	set rsSede = Conn.execute(strSql)
	if not rsSede.eof then
		strRazonSocialSede=rsSede("RAZON_SOCIAL")
		strDireccionSede=rsSede("DIRECCION")
		strRutSede=rsSede("RUT")
		strTelefonoSede=rsSede("TELEFONO")
		strNroFolio=rsSede("FOLIO_ACTUAL")
	End If

%>

<title>Documento sin t&iacute;tulo</title>
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
		.Estilo33 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 14px;font-weight: bold;}
		.Estilo34 {font-family: Verdana, Arial, Helvetica, sans-serif;font-size: 11px}

		H1.SaltoDePagina {PAGE-BREAK-AFTER: always}
			.transpa {
			background-color: transparent;
			border: 1px solid #FFFFFF;
			text-align:center
			}
		-->
				.align { text-align: justify }
	</style>
</head>


<BODY>

<TABLE ALIGN="CENTER" WIDTH="650" CELLSPACING=0 CELLPADDING=1>

		<TR>
		  <TD>
		  <table width="650" border="0" cellspacing=0 cellpadding=0>
			<tr>
			  <td width=150 rowspan="3"  class="Estilo13" align="center"><img border="0" src="../imagenes/Logos/<%=session("ses_codcli")%>/logo.jpg"></td>
			  <td width=100>&nbsp;</td>
			  <td></td>
			</tr>
			<tr>
			  <td></td>
			  <td>&nbsp;</td>
			  <td align="right"><span class="Estilo33"> Nº <%=strFolio%></span></td>
			</tr>
		  </table>
		  </TD>
 		</TR>

 		<TR>
		 	<TD>

		 		<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
					<tr class="Estilo1">
						<TD colspan=4 align="center" class="Estilo33">
						PAGARÉ 
						</TD>
					</TR>
					<TR>
						<td colspan=4><div align="center"><%=strMandante%></div></TD>
					</TR>
				</TABLE>

	 		</TD>
 		</TR>
 		<TR>
				<TD>

				<%
				strSql="SELECT CUOTA, TOTAL_CUOTA, CONVERT(VARCHAR(10),FECHA_PAGO,103) as FECHA_PAGO FROM CONVENIO_DET WHERE ID_CONVENIO = " & intIdConvenio & " AND CUOTA <> 0"
				'response.write(strSql)
				set rsDetConv=Conn.execute(strSql)
				If Not rsDetConv.Eof Then
					intValorCuota = rsDetConv("TOTAL_CUOTA")
				End If

				%>

				<br>
				<table width="650" border="0">
				<br>
					<tr>
						<td class="Estilo34" align="left"><p class="align">
						Debo y pagaré a la orden de Universidad Mayor, Rut. 71.500.500-K, la suma de 
						<B>$<%=FN((intTotalConvenio -  intPie),0)%></B> pesos.
						<br>
						La cantidad adecuada devengará un interés del 0,0% anual. En caso de mora
						o simple retardo en el pago, se devengará un interés mensual igual al interés
						máximo convencional que la Ley permite para operaciones de crédito de dinero 
						no reajustables, por el periodo comprendido entre la fecha de la mora o simple 
						retardo y el día de su pago.
						<br><br>
						El capital y los intereses adeudados se pagarán  en <b><%=intCuotas%></b> cuotas mensuales,
						cuyos montos y vencimientos serán los siguientes:
						<br>
						</td>
					</tr>
				</table>
				<br>
				<br>


				</TD>
	 		</TR>

	 	<TR>
	 		<TD>
	 		<TABLE ALIGN= "CENTER" WIDTH="70%" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
	 			<%
							%>
						
							<tr>
							<td>
							<table align="center" border="1" cellpadding="0" cellspacing="0">
								<tr>
									<td width="140px" align="center">  NºCuota  </td>
									<td width="200px" align="center">  Monto $  </td>
									<td width="200px" align="center">  Fecha Vencimiento  </td>	
								</tr>
							<%
							Do While Not rsDetConv.Eof
								intMontoTotal = intMontoTotal + ValNulo(rsDetConv("TOTAL_CUOTA"),"N")
							%>

							<tr>
								<td align="center"> <%=rsDetConv("CUOTA")%></td>
								<td align="center"> $ <%=FN(rsDetConv("TOTAL_CUOTA"),0)%></td>
								<td align="center"> <%=rsDetConv("FECHA_PAGO")%></td>
							</tr>
							<%
								rsDetConv.movenext
							Loop
							%>

							</table>
							</td>
							</tr>
					</TABLE>
				</TD>
	 		</TR>

	 		<TR>
			 	<TD>
					<TABLE ALIGN= "CENTER" WIDTH="90%" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

						
						<br>
						<tr>
							<td class="Estilo34" align="left"><p class="align">
							En caso de mora o simple retardo en el pago de cualquiera de las cuotas antes
							señaladas, acepto expresamente que la Universidad Mayor podrá exigir
							anticipadamente el pago íntegro y total  de la deuda insoluta, la que se 
							considerará como de plazo vencido para todos los efectos legales.
							<br><br>
							La obligación contraída en este instrumento es indivisible, para todos los 
							efectos legales, indivisibilidad que pasará a los herederos y sucesores del
							deudor, pudiendo la Universidad Mayor, demandar su cobro a cuales quiera
							de ellos, en conformidad al artículo 1.528 del Código Civil.
							<br><br>
							Libero a la Universidad Mayor de la obligación de protestar el presente pagaré en
							caso de no ser pagado a su vencimiento.
							<br><br>
							Para los efectos legales, judiciales y de eventual protesto de este documento,
							señalo domicilio en Av. Américo Vespucio Sur Nº357, y me someto a la 
							competencia y jurisdicción de los tribunales ordinarios de Justicia de la ciudad
							de Santiago comuna de Las Condes.
							<br><br>
							Para el caso de no pago del pagaré suscrito o en caso de simple retardo o
							mora en el cumplimiento de las obligaciones contraídas, el compareciente
							autoriza expresamente a Universidad Mayor para el ingreso de los datos
							personales y antecedentes del protesto o incumplimiento en un sistema de
							información comercial facultando expresamente su digitación, procesamiento y
							comunicación en línea o en cualquier otra forma y liberando desde ya a
							Universidad Mayor de dar aviso de pago o extinción de deuda, previo
							otorgamiento de constancia suficiente del pago de la deuda correspondiente,
							según lo dispuesto en la Ley 19.628.
							<br><br>

							<br>
							<br>
							<br>
							<br>
							___________________________________<br>
							Firma Suscriptor:<br><br>
							<br>
							Fecha: _____________________________<br>
							<br>
							Nombre Suscriptor: ________________________________________<br>
							<br>
							Numero Interno: ____________________<br>
							<br>
							RUT: ______________________________<br>
							<br>
							Domicilio: _________________________________________________<br>
							<br>
							Fono: _____________________________<br>
							<br>
							Celular: ___________________________<br>
							<br>
							<br><br>
							"La Universidad Mayor se encuentra exenta del impuesto de Timbres y
							Estampillas que grava la emisión de Letras de Cambios y Pagarés, y las actas
							de protesto, de acuerdo a lo establecido en los artículos 23 Nº3 y 11 inciso
							segundo del DL 3475, de fecha 29.08.1980."
							<br>
							<br>
							</td>
						</tr>
					</TABLE>
				</TD>
	 		</TR>


	 		<!--<TR class="Estilo34">
				<TD>
					<b>CERTIFICACION NOTARIAL : <br>Firmó ante mi don(ña) <%=strNombreDeudor%> , cédula de identidad Nro : <%=strRUT_DEUDOR%></b>
					<BR><b><%=UCASE(strCiudadSede)%> , <%=DATE%></b>
				</TD>
	 		</TR>-->


	 </TABLE>




	<%

	'Capital: $ FN(intTotDeudaCapital,0)
	'Intereses: $ FN(intTotIntereses,0)
	'Honorarios: FN(intTotHonorarios,0)
	'G.Protesto: FN(intTotProtestos,0)
	'Total Deuda: FN(intTotalDeuda,0)
	'Pie session("NOMBRE_CONV_PAGARE") FN(intPie,0)
	'Saldo en session("NOMBRE_CONV_PAGARE") FN(intTotalDeuda -  intPie,0)
	'Intereses session("NOMBRE_CONV_PAGARE") FN(intMontoA,0)
	'Saldo a Convenir: $ FN((intTotalDeuda -  intPie) + intMontoA,0)
	'
	'Pie session("NOMBRE_CONV_PAGARE")$ FN(intPie,0)
	'Total Pago Caja:FN(intPie,0)

	%>



	<BR>
	   	<H1 class=SaltoDePagina> </H1>
	<BR>




	<TABLE ALIGN="CENTER" WIDTH="650" BORDER="1" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

				<TR>
				  <TD>
				  <table width="650" border="0" cellspacing=0 cellpadding=0>
					<!--<tr>
					  <td class="Estilo13" width=400><b><%=strRazonSocialSede%></b></td>
					  <td width=100>&nbsp;</td>
					  <td width=150 rowspan="3"  class="Estilo13" align="center"><img border="0" width="150" height="46" src="UploadFolder/<%=session("ses_codcli")%>/logo.jpg"></td>
					</tr>
					<tr>
					  <td class="Estilo13">Dirección :<%=strDireccionSede%></td>
					  <td>&nbsp;</td>
					</tr>
					<tr>
					  <td class="Estilo13">R.U.T: :<%=strRutSede%></td>
					  <td>&nbsp;</td>
					</tr>
					<tr>
					  <td class="Estilo13">Teléfono : <%=strTelefonoSede%></td>
					  <td>&nbsp;</td>
					  <td align="center"><span class="Estilo13"><%=session("NOMBRE_CONV_PAGARE")%> : <%=strNroFolio%></span></td>
					</tr>-->
						<tr>
						  <img border="0" src="../imagenes/Logos/<%=session("ses_codcli")%>/logo.jpg"></td>
						  <td width=100>&nbsp;</td>
						  <td></td>
						</tr>
						<tr>
						  <td></td>
						  <td>&nbsp;</td>
						  <td align="right"><span class="Estilo33"> Nº <%=strFolio%></span></td>
						</tr>
				  </table>
				  </TD>
	 		</TR>

			<TR>
			 	<TD>

			 		<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

					 		<!--<tr class="Estilo1">
							 	<TD colspan=4 align="center" class="Estilo33">
						 		  ANEXO <%=UCASE(session("NOMBRE_CONV_PAGARE"))%><BR>
						 		</TD>
					 		</TR>
					 		<TR>
							 	<td colspan=4><div align="center"><%=strMandante%></div></TD>
					 		</TR>-->
							<tr class="Estilo1">
						 	<TD colspan=4 align="center" class="Estilo33">
					 		  DETALLE DEUDA<BR>
					 		</TD>
							</TR>
							<TR>
								<td colspan=4><div align="center"><%=strMandante%></div></TD>
							</TR>
					</TABLE>

		 		</TD>
	 		</TR>
	 		<TR>
				<TD>
					<TABLE ALIGN="CENTER" WIDTH="650" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>

						<tr class="Estilo1">
							<TD colspan=4 align="LEFT" class="Estilo34">
							  <br>ANTECEDENTES DE LA DEUDA
							</TD>
						</TR>
					</TABLE>


					<table width="100%" border="1" bordercolor = "#000000" cellSpacing=0 cellPadding=1>
					<tr class="Estilo34">
						<td>NRO_DOC</td>
						<td>FECHA_VENC</td>
						<td>DIAS_MORA</td>
						<td>TIPO_DOCUMENTO</td>
						<!--td>ASIGNACION</td-->
						<td>CAPITAL</td>
						<td>INTERES</td>
						<td>PROTESTOS</td>
						<td>HONORARIOS</td>
					</tr>

					<%

					strSql = "SELECT ID_CUOTA, NRO_DOC, TIPO_DOCUMENTO, VALOR_CUOTA, GASTOS_PROTESTOS, CUENTA, FECHA_VENC, ISNULL(DATEDIFF(D,FECHA_VENC,GETDATE()),0) AS ANTIGUEDAD FROM CUOTA WHERE RUT_DEUDOR='" & strRUT_DEUDOR & "' AND COD_CLIENTE='" & strIdCliente & "'" 
					strSql = strSql+ " and id_cuota in (SELECT ID_CUOTA FROM CONVENIO_CUOTA where ID_CONVENIO = "&intIdConvenio&")"
					strSql = strSql+ "ORDER BY FECHA_VENC DESC"
					
					'Response.write "strSql=" & strSql
					
					
					set rsTemp= Conn.execute(strSql)

					intTasaMensual = 2/100
					intTasaDiaria = intTasaMensual/30
					intCorrelativo = 1
					strArrID_CUOTA=""
					intTotSelSaldo= 0
					intTotSelIntereses= 0
					intTotSelProtestos= 0
					intTotSelHonorarios= 0
					Do until rsTemp.eof

							intSaldo = rsTemp("VALOR_CUOTA")
							strNroDoc = rsTemp("NRO_DOC")
							strFechaVenc = rsTemp("FECHA_VENC")
							strTipoDoc = rsTemp("TIPO_DOCUMENTO")

							intAntiguedad = ValNulo(rsTemp("ANTIGUEDAD"),"N")
							'intIntereses = intTasaDiaria * intAntiguedad * intSaldo
							'intHonorarios = GASTOS_COBRANZAS(intSaldo)
							intProtestos = ValNulo(rsTemp("GASTOS_PROTESTOS"),"N")

							strArrID_CUOTA = strArrID_CUOTA & ";" & rsTemp("ID_CUOTA")

							intTotSelSaldo= intTotSelSaldo+intSaldo
							'intTotSelIntereses= intTotSelIntereses+intIntereses
							intTotSelProtestos= intTotSelProtestos+intProtestos
							'intTotSelHonorarios= intTotSelHonorarios+intHonorarios

							%>
							<tr class="Estilo34">
							<td><%=strNroDoc%></td>
							<td><%=strFechaVenc%></td>
							<td><%=intAntiguedad%></td>
							<td><%=strTipoDoc%></td>
							<td ALIGN="RIGHT"><%=FN(intSaldo,0)%></td>
							<td ALIGN="RIGHT"><%=FN(0,0)%></td>
							<td ALIGN="RIGHT"><%=FN(intProtestos,0)%></td>
							<td ALIGN="RIGHT"><%=FN(0,0)%></td>
							</tr>
							<%

					rsTemp.movenext
					intCorrelativo = intCorrelativo + 1
					loop
					rsTemp.close
					set rsTemp=nothing

					strArrID_CUOTA = Mid(strArrID_CUOTA,2,len(strArrID_CUOTA))
			%>
				<tr class="Estilo34">
					<td colspan = 4>Totales</td>
					<td ALIGN="RIGHT"><%=FN(intTotSelSaldo,0)%></td>
					<td ALIGN="RIGHT"><%=FN(intIntereses,0)%></td>
					<td ALIGN="RIGHT"><%=FN(intTotSelProtestos,0)%></td>
					<td ALIGN="RIGHT"><%=FN(intHonorarios,0)%></td>
				</tr>

				<INPUT TYPE="HIDDEN" NAME="strArrID_CUOTA" VALUE="<%=strArrID_CUOTA%>">
				</table>

				<br>





<%
				intTotDeudaCapital = intCapital - intDescCapital
				intTotIntereses = intIntereses - intDescIntereses
				intTotProtestos = intTProtestos - intDescProtestos
				intTotHonorarios = intHonorarios - intDescHonorarios
				intTotIndemComp = intIndemComp - intDescIndemComp
				'Response.write "intTotIndemComp=" & intTotIndemComp
				'Response.write "intIndemComp=" & intIndemComp
				'Response.write "intDescIndemComp=" & intDescIndemComp
				intTotGastos = intGastos - intDescGastos

%>





						<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0">
							<TR class="Estilo1">
								<TD colspan=4 align="CENTER" class="Estilo38">
								  DETALLE <%=UCASE(session("NOMBRE_CONV_PAGARE"))%><br><br>
								</TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD ALIGN="RIGHT"><b>Conceptos</b></TD><TD align="right"><b>Monto Original</b></TD>
								<TD align="right"><b>Descuentos</b></TD><TD align="right"><b>Total</b></TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD ALIGN="RIGHT">Capital: </TD><TD align="right">$ <%=FN(intCapital,0)%></TD>
								<TD align="right">$ <%=FN(intDescCapital,0)%></TD><TD align="right">$ <%=FN(intTotDeudaCapital,0)%></TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right">Intereses: </TD><TD align="right">$ <%=FN(intIntereses,0)%></TD>
								<TD align="right">$ <%=FN(intDescInteres,0)%></TD><TD align="right">$ <%=FN(intTotIntereses,0)%></TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right">Protestos: </TD><TD align="right">$ <%=FN(intProtestos,0)%></TD>
								<TD align="right">$ <%=FN(intDescProtestos,0)%></TD><TD align="right">$ <%=FN(intTotProtestos,0)%></TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right">Honorarios: </TD><TD align="right">$ <%=FN(intHonorarios,0)%></TD>
								<TD align="right">$ <%=FN(intDescHonorarios,0)%></TD><TD align="right">$ <%=FN(intTotHonorarios,0)%></TD>
							</TR>
							<!--<TR HEIGHT=15 class="Estilo38">
								<TD align="right">Indem. Comp.: </TD><TD align="right">$ <%=FN(intIndemComp,0)%></TD>
								<TD align="right">$ <%=FN(intDescIndemComp,0)%></TD><TD align="right">$ <%=FN(intTotIndemComp,0)%></TD>
							</TR>-->
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right">Gastos Judiciales: </TD><TD align="right">$ <%=FN(intGastos,0)%></TD>
								<TD align="right">$ <%=FN(intDescGastos,0)%></TD><TD align="right">$ <%=FN(intTotGastos,0)%></TD>
							</TR>

							<TR HEIGHT=15>
								<TD colspan=3>&nbsp</TD><TD align="right">______________</TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right" colspan=3>Total Deuda:</TD><TD align="right"><strong>$ <%=FN(intTotalConvenio,0)%></strong></TD>
							</TR>
							<!--TR HEIGHT=15 class="Estilo38">
								<TD align="right">Total Descuentos:</TD><TD align="right"><strong>$ <%=FN(intTotalDeuda,0)%></strong></TD>
							</TR-->

							<TR HEIGHT=15 class="Estilo38">
								<TD align="right" colspan=3>Pie <%=session("NOMBRE_CONV_PAGARE")%>:</TD><TD align="right"><strong>-  $ <%=FN(intPie,0)%></strong></TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD colspan=3>&nbsp</TD><TD align="right">______________</TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right" colspan=3>Saldo:</TD><TD align="right"><strong>$ <%=FN((intTotalConvenio -  intPie),0)%></strong></TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right" colspan=3>Intereses <%=session("NOMBRE_CONV_PAGARE")%>.:</TD>	<TD align="right"><strong>$ <%=FN(intMontoTotal - (intTotalConvenio -  intPie),0)%></strong></TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right" colspan=3>&nbsp</TD><TD align="right" class="Estilo38">&nbsp</TD>
							</TR>
							<TR HEIGHT=15 class="Estilo38">
								<TD align="right" colspan=3>&nbsp</TD><TD align="right" class="Estilo38">&nbsp</TD>
							</TR>

							<TR HEIGHT=15 class="Estilo38">
								<TD align="right" colspan=3><b>Saldo <%=session("NOMBRE_CONV_PAGARE")%>: <b></TD>
								<TD align="right"><strong>$ <%=FN(intMontoTotal,0)%></strong></TD>
							</TR>
						</table>
					</tr>
				</TD>
	 		</TR>
	 </TABLE>

	</FORM>

	<!--<TABLE ALIGN="CENTER" WIDTH="600" BORDER="0" BORDERCOLOR = "#000000" CELLSPACING=0 CELLPADDING=1>
				<TR class="Estilo1">
					<TD align="LEFT" class="Estilo34">
						<acronym title="GRABAR <%=UCASE(session("NOMBRE_CONV_PAGARE"))%>">
							<input name="BT_GRABAR" type="button" onClick="Grabar();" value="Grabar">
						</acronym>
					</TD>
					<TD align="RIGHT" class="Estilo34">
						<acronym title="IMPRIMIR <%=UCASE(session("NOMBRE_CONV_PAGARE"))%>">
							<input name="BT_IMPRIMIR" type="button" onClick="window.print();" value="Imprimir">
						</acronym>
					</TD>

				</TR>
		</TABLE>-->

	&nbsp;&nbsp;
	</body>
</html>