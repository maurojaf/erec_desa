<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

<%
	email =Request.querystring("email")
%> 
<%if trim(email)="" then%>
	<!--#include file="sesion.asp"-->
<%end if%>

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
    <script src="../Componentes/jquery.numeric/jquery.numeric.js"></script>

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	If Trim(Request("Limpiar"))="1" Then
		session("session_RUT_DEUDOR") = ""
		strRUTDEUDOR = ""
	End if

	If Trim(Request("TX_RUT")) = "" Then
		strRUTDEUDOR = session("session_RUT_DEUDOR")
	Else
		strRUTDEUDOR = Trim(Request("TX_RUT"))
		session("session_RUT_DEUDOR") = strRUTDEUDOR
	End If

	intIdDemanda=Trim(Request("intIdDemanda"))
	intOrigen=Trim(Request("intOrigen"))
    intTipoPago = Trim(Request("CB_TIPO"))

	If Trim(intOrigen)="CO" Then
		strEstiloPrinc="style='display:block;'"
		strPagina = "generar_convenio.asp"

	ElseIf Trim(intOrigen)="PP" Then
		strEstiloPrinc="style='display:none;'"
		strPagina = "plan_pago_convenio.asp"
	End If

   strTipoPlan ="NE"
 
	AbrirSCG()
%>

<input type="hidden" id="Origen" value="<%=intOrigen%>" />
<input type="hidden" id="intInteres" value="<%=intInteres%>" />

	<style type="text/css">
	<!--	
	.Estilo29 {font-family: Arial, sans-serif}
	-->
	</style>

	<style type="text/css">
	<!--
	body{

		font-size: 11px;

	}
	.style1 {color: #ffffff}

	.style2 {color: #005083;}
	.style3 {color: #B7B7B7}
	.style4 {color: #000000}
	.Estilo1 {font-size: 12px}
	.Estilo31 {font-size: 14}
	.Estilo32 {color: #333333}
	.Estilo33 {font-size: 18px}
	-->
	 .hiddencol
        {
            display:none;
        }
	</style>
</head>

<body>
<%

	If Trim(intCOD_CLIENTE) = "" Then intCOD_CLIENTE = session("ses_codcli")
	If Trim(strRUTDEUDOR) <> "" then
		strNombreDeudor = TraeNombreDeudor(Conn,strRUTDEUDOR)
		strFonoArea = TraeUltimoFonoDeudor(Conn,strRUTDEUDOR,"COD_AREA")
		strFonoFono = TraeUltimoFonoDeudor(Conn,strRUTDEUDOR,"TELEFONO")
		strDirCalle= TraeUltimaDirDeudorSCG(Conn,strRUTDEUDOR,"CALLE")
		strDirNum = TraeUltimaDirDeudorSCG(Conn,strRUTDEUDOR,"NUMERO")
		strDirComuna = TraeUltimaDirDeudorSCG(Conn,strRUTDEUDOR,"COMUNA")
		strDirResto = TraeUltimaDirDeudorSCG(Conn,strRUTDEUDOR,"RESTO")
		strEmail = TraeUltimoEmailDeudorSCG(Conn,strRUTDEUDOR,"EMAIL")

		strTelefonoDeudor = TraeUltimoFonoDeudorSCG(Conn,strRUTDEUDOR,"COD_AREA") & "-" & TraeUltimoFonoDeudorSCG(Conn,strRUTDEUDOR,"TELEFONO")
		If Trim(strTelefonoDeudor) = "-" Then strTelefonoDeudor = "S/F"
	Else
		strNombreDeudor=""
		strFonoArea = ""
		strFonoFono = ""
		strDirCalle = ""
		strDirNum = ""
		strDirComuna = ""
		strDirResto = ""
	End if

	strSql=""
	strSql="SELECT USA_SUBCLIENTE, USA_INTERESES, USA_HONORARIOS, USA_PROTESTOS, FORMULA_HONORARIOS,FORMULA_INTERESES,PIE_PORC_CAPITAL, HON_PORC_CAPITAL, IC_PORC_CAPITAL, TASA_MAX_CONV, DESCRIPCION, RAZON_SOCIAL,INTERES_MORA, USA_CUSTODIO, COD_TIPODOCUMENTO_HON, MESES_TD_HON FROM CLIENTE WHERE COD_CLIENTE ='" & intCOD_CLIENTE & "'"
	set rsTasa=Conn.execute(strSql)
	if not rsTasa.eof then
		intTasaMax = ValNulo(rsTasa("TASA_MAX_CONV"),"N")/100
		intPorcPie = ValNulo(rsTasa("PIE_PORC_CAPITAL"),"N")/100
		intPorcHon = ValNulo(rsTasa("HON_PORC_CAPITAL"),"N")/100
		intPorcIc = ValNulo(rsTasa("IC_PORC_CAPITAL"),"N")/100

		strNomFormHon = ValNulo(rsTasa("FORMULA_HONORARIOS"),"C")
		strNomFormInt = ValNulo(rsTasa("FORMULA_INTERESES"),"C")


		strUsaSubCliente = rsTasa("USA_SUBCLIENTE")
		strUsaInteres = rsTasa("USA_INTERESES")
		strUsaHonorarios = rsTasa("USA_HONORARIOS")
		strUsaProtestos = rsTasa("USA_PROTESTOS")

       	strDescripcion = rsTasa("RAZON_SOCIAL")
		strMandante = strDescripcion
		intTasaMensual = ValNulo(rsTasa("INTERES_MORA"),"C")
		strUsaCustodio = rsTasa("USA_CUSTODIO")
		intTipoDocHono = ValNulo(rsTasa("COD_TIPODOCUMENTO_HON"),"C")
		intMesHon = ValNulo(rsTasa("MESES_TD_HON"),"C")

		If intTasaMensual = "" Then
			%>
				<SCRIPT>alert('No se ha definido tasa de interes de mora, se ocupara una tasa del 2%, favor parametrizar')</SCRIPT>
			<%
			intTasaMensual = "2"
		End If


	Else
		intTasaMax = 1
		intPorcPie = 1
		intPorcHon = 1
		intPorcIc = 1
		strDescripcion = ""
	end if
	rsTasa.close
	set rsTasa=nothing



	strSql="SELECT CUENTA, NRO_DOC, IsNull(FECHA_VENC,'01/01/1900'), IsNull(datediff(d,FECHA_VENC,getdate()),0) , TIPO_DOCUMENTO, COD_REMESA , VALOR_CUOTA, SALDO, USUARIO.LOGIN , ESTADO_DEUDA.DESCRIPCION"
	strSql=strSql & " FROM CUOTA , USUARIO , ESTADO_DEUDA WHERE RUT_DEUDOR = '" & strRUTDEUDOR & "' AND COD_CLIENTE = '" & intCOD_CLIENTE & "'"
	strSql=strSql & " AND CUOTA.USUARIO_ASIG *= USUARIO.ID_USUARIO AND CUOTA.ESTADO_DEUDA *= ESTADO_DEUDA.CODIGO"

	'response.write "<br>strSql=" & strSql
	'Response.End
	set rsDET=Conn.execute(strSql)
	intTotHonorarios=0
	intTotIndemComp=0
	intTotDeudaCapital = 0
	if not rsDET.eof then
		intColumnas = rsDET.Fields.Count - 1
		intSaldo = 0
		intValorCuota = 0
		total_ValorCuota = 0
		''rsDET.movenext
		Do until rsDET.eof
			intTotDeudaCapital = intTotDeudaCapital + Round(session("valor_moneda") * ValNulo(rsDET(7),"N"),0)
            rsDET.movenext
		Loop
	end if
	rsDET.close
	set rsDET=nothing

	If Trim(intCOD_CLIENTE)= "1000" Then
		If intIndemCompensatoriaD  > 35000 Then
			'intIndemCompensatoriaD = intIndemCompensatoriaD - 35000
			'intOtrosD = 35000
		End If
	End If

	intTotIndemComp = intIndemCompensatoriaD
	intTotHonorarios = intOtrosD
	''intTotHonorarios = 0
	intTotGastos = intGastosJudicialesD


	If Trim(intTotDeudaCapital) = "0" Then
		intTotIndemComp = 0
		intTotHonorarios = 0
		intTotGastos = 0
	End If


	    'intTotalDeuda = intTotDeudaCapital + intTotIndemComp + intTotHonorarios
		intTotalCostas = VALNULO(intTotIndemComp,"N") + VALNULO(intTotGastos,"N")

	cerrarSCG()

	'Response.write "CH_PARAM = " & UCASE(Request("CH_PARAM"))

	'If UCASE(Request("CH_PARAM")) = "ON" Then strCheckParam = "checked"  Else strCheckParam = ""
	%>

	
<FORM onSubmit="return validardatos(this);" name="datos" method="post" action="<%=strPagina%>">
<div class="titulo_informe">
			<% If Trim(intOrigen)="CO" Then %>
				RECONOCIMIENTO DE DEUDA Y <%=UCASE(session("NOMBRE_CONV_PAGARE"))%> 
			<% End If %>
			<% If Trim(intOrigen)="PP" Then %>
				PLAN DE PAGO
			<% End If %> <%=" " & strMandante%>
</div>
<TABLE ALIGN="CENTER" width="90%">
	<TR>
		<TD>
		<table width="100%" border="0" bordercolor="#FFFFFF">
	      <tr >

	        <td width="93" class="estilo_columna_individual">&nbsp;&nbsp;RUT DEUDOR&nbsp;</td>

	        <td width="10" bgcolor="#FFFFFF" border="1" bordercolor="#999999"><a href="javascript:ventanaBusqueda('Busqueda.asp?strOrigen=1&TX_RUT_DEUDOR=<%=strRUTDEUDOR%>&TX_NOMBRE=<%=strNombreDeudor%>')"><img src="../imagenes/buscar.png" border="0"></a>
	        <td width="106" class="Estilo10" bgcolor="#<%=session("COLTABBG2")%>">&nbsp;<%=strRUTDEUDOR%></td>
	        <td width="120" class="estilo_columna_individual">&nbsp;&nbsp;NOMBRE DEUDOR</td>
	        <td class="Estilo10" bgcolor="#<%=session("COLTABBG2")%>">&nbsp;<%=strNombreDeudor%></td>    
	      </tr>
	    </table>

			<table width="100%" border="1" BORDERCOLOR="#FFFFFF" class="estilo_columnas">
				<thead>
					<tr>
						<td>USUARIO</td>
						<td>FECHA</td>
						<td>SEDE</td>
						<td><% If Trim(intOrigen)="PP" Then %>TIPO PAGO<%END IF%></td>
						<td><% If Trim(intOrigen)="PP" Then %>FORMA DE PAGO PIE<%END IF%></td>
						<TD></TD>
					</TR>
				</thead>
					  <TR>
						<TD border="1" BORDERCOLOR="#FFFFFF" ALIGN="LEFT"><%=session("session_login")%>
						<INPUT TYPE="hidden" NAME="intOrigen" value="<%=intOrigen%>"></TD>

						<TD border="1" BORDERCOLOR="#FFFFFF"><%=DATE%></TD>
						<TD border="1" BORDERCOLOR="#FFFFFF">
							<select name="CB_SEDE" id="CB_SEDE">
								<option value="">SELECCIONAR</option>
								<%
								AbrirSCG()
									strSql="SELECT * FROM SEDE WHERE COD_CLIENTE = '" & intCOD_CLIENTE & "'"

									set rsSede=Conn.execute(strSql)
									Do While not rsSede.eof
									%>
									<option value="<%=rsSede("SEDE")%>"> <%=rsSede("SEDE")%></option>
									<%
									rsSede.movenext
									Loop
									rsSede.close
									set rsSede=nothing
								CerrarSCG()
								%>
							</select>
						</TD>
						<td border="1" BORDERCOLOR="#FFFFFF" >
						<% If Trim(intOrigen)="PP" Then %>
						<select name="CB_TIPO" id="CB_TIPO" onChange="cargaTipo(this.value);val_tipo();">
							<option value="">SELECCIONAR</option>
							<%
								AbrirSCG()
									strSql="SELECT  COD_TIPO_PLAN_PAGO ,NOM_TIPO_PLAN_PAGO ,CASE WHEN COD_TIPO_PLAN_PAGO ='NE' THEN  1 when  COD_TIPO_PLAN_PAGO ='CO'  then 2 ELSE 3 END ORDEN FROM TIPO_PLAN_PAGO  ORDER BY 3,2"
									set rsSede=Conn.execute(strSql)
									Do While not rsSede.eof
									%>
									<option  <%If Trim(strTipoPlan)=rsSede("COD_TIPO_PLAN_PAGO") Then Response.write "SELECTED"%> value="<%=rsSede("COD_TIPO_PLAN_PAGO")%>"> <%=rsSede("NOM_TIPO_PLAN_PAGO")%></option>
									<%
									rsSede.movenext
									Loop
									rsSede.close
									set rsSede=nothing
								CerrarSCG()
							%>

						</select>
						<%END IF%>
						</td>
						<td border="1" BORDERCOLOR="#FFFFFF" >
							<% If Trim(intOrigen)="PP" Then %>
							<select name="CB_FPAGO" id="CB_FPAGO" onChange="FORMA_PAGO();" width="100" maxlength="10">
                                <option>NO ESPECIFICADO</option>
                            </select>
                           
							<%END IF%>
						</td>												
						<TD border="1" BORDERCOLOR="#FFFFFF">							
							<input class="fondo_boton_100" name="li_" type="button" onClick="bt_limpia()" value="Limpiar">							
						</TD>
					  </TR>
	
			</table>

		</TD>
	</TR>
</TABLE>

<BR>

<div id="principal2" style="width:90%; margin:0 auto;" >
	<table  class="estilo_columnas" style="width:100%;" >
		<thead>
		<tr>
			<td colspan="10" valign="top" align="left">
				<a href="#" onClick= "marcar_boxes(true);func_porc_capital_pie();" style="color:#FFFFFF;">Marcar todos</a>&nbsp;&nbsp;&nbsp;
				<a href="#" onClick="desmarcar_boxes(true);" style="color:#FFFFFF;">Desmarcar todos</a>
			</td>
	    </tr>
		</thead>
	</TABLE>

	<table  class="intercalado" style="width:100%;" id="tbl_Procesa" >
    <thead>
			     <tr>
		          <td>&nbsp;</td>
                  <td class='hiddencol'>ID_CUOTA</td>
		          <td>NRO. DOC</td>
                  <td>CUOTA</td>
		          <td>F.VENCIM.</td>
		          <td>ANTIG.</td>
		          <td>TIPO DOC</td>
		          <td>ASIG.</td>

		          <td ALIGN="CENTER">CAPITAL</td>

		          <%If Trim(strUsaInteres)="1" Then%>
		          	<td ALIGN="CENTER">INTERES</td>
		          <%End If%>

		          <%If Trim(strUsaProtestos)="1" Then%>
		          	<td ALIGN="CENTER">PROTESTOS</td>
		          <%End If%>

		          <%If Trim(strUsaHonorarios)="1" Then%>
		          	<td ALIGN="CENTER">HONORARIOS</td>
		          <%End If%>
				  
				  <td ALIGN="CENTER">ABONO</td>

                  <td ALIGN="CENTER">SALDO</td>
                </tr>
             </thead>
             <tbody>


		<%
		If Trim(strRUTDEUDOR) <> "" then
		abrirscg()
			strSql = "SELECT NRO_CUOTA,dbo." & strNomFormInt & "(ID_CUOTA) as INTERESES, dbo." & strNomFormHon & "(ID_CUOTA) as HONORARIOS, ID_CUOTA, RUT_DEUDOR, DATEDIFF(DAY,FECHA_VENC,GETDATE()) AS ANT_DIAS, DATEDIFF(MONTH,FECHA_VENC,GETDATE()) AS ANT_MESES, CUSTODIO, IsNull(FECHA_VENC,'01/01/1900') as FECHA_VENC, IsNull(datediff(d,FECHA_VENC,getdate()),0) as ANTIGUEDAD, NRO_DOC, IsNull(VALOR_CUOTA,0) as VALOR_CUOTA,IsNull(SALDO,0) as SALDO,IsNull(USUARIO_ASIG,0) as USUARIO_ASIG, NRO_CUOTA, SUCURSAL , ESTADO_DEUDA, COD_REMESA, CUENTA, NRO_DOC, TIPO_DOCUMENTO, GASTOS_PROTESTOS, NOM_TIPO_DOCUMENTO "
			strSql = strSql & " FROM CUOTA, TIPO_DOCUMENTO WHERE RUT_DEUDOR='"& strRUTDEUDOR &"' AND COD_CLIENTE='" & intCOD_CLIENTE & "' AND SALDO > 0 AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1) AND CUOTA.TIPO_DOCUMENTO = TIPO_DOCUMENTO.COD_TIPO_DOCUMENTO ORDER BY FECHA_VENC DESC"
			'response.Write(strSql)
			'response.End()
			set rsDET=Conn.execute(strSql)

			if not rsDET.eof then
			
				intSaldo = 0
				intValorCuota = 0
				intAbono = 0
				total_ValorCuota = 0
				strArrConcepto = ""
				strArrID_CUOTA = ""

				intTasaMensual = intTasaMensual/100
				intTasaDiaria = intTasaMensual/30


				Do until rsDET.eof
					strCustodio = ValNulo(rsDET("CUSTODIO"),"C")
					If Trim(strCustodio) <> "" Then
						strCustodio="(C)"
					End If

					intSaldo = Round(session("valor_moneda") * ValNulo(rsDET("SALDO"),"N"),0)
					intProtesto = Round(session("valor_moneda") * ValNulo(rsDET("GASTOS_PROTESTOS"),"N"),0)
					intValorCuota = Round(session("valor_moneda") * ValNulo(rsDET("VALOR_CUOTA"),"N"),0)
					intAbono = intValorCuota - intSaldo 

					strNroDoc = Trim(rsDET("NRO_DOC"))
					strNroCuota = Trim(rsDET("NRO_CUOTA"))
					strSucursal = Trim(rsDET("SUCURSAL"))
					strEstadoDeuda = Trim(rsDET("ESTADO_DEUDA"))
					strCodRemesa = Trim(rsDET("COD_REMESA"))

					strArrConcepto = strArrConcepto & ";" & "CH_" & rsDET("ID_CUOTA")
					strArrID_CUOTA = strArrID_CUOTA & ";" & rsDET("ID_CUOTA")

					intAntiguedad = ValNulo(rsDET("ANTIGUEDAD"),"N")

					intIntereses = rsDET("INTERESES")
					intHonorarios = rsDET("HONORARIOS")

					intTotHonorarios = intTotHonorarios + intHonorarios
					intTotIntereses = intTotIntereses + intIntereses
					intTotProtesto = intTotProtesto + intProtesto
                    intTotalSaldo = intTotalSaldo + intSaldo 

                    intTotalDeuda =  (intTotalSaldo+  intTotProtesto +intTotIntereses +intTotHonorarios) 
                     

				%>
		        <tr >
		          <td>
                  	<input type="checkbox" checked="checked" id="CH_<%=rsDET("ID_CUOTA")%>"  name="CH_<%=rsDET("ID_CUOTA")%>" onChange ="suma_capital(this,<%=intSaldo%>,<%=Round(intHonorarios,0)%>,<%=Round(intIntereses,0)%>,<%=Round(intProtesto,0)%>)"/>
		          </td>
                  <td class='hiddencol'><%=rsDET("ID_CUOTA")%></td>
		          <td><div align="left"><%=rsDET("NRO_DOC")&" "&strCustodio%></div></td>
                  <td><div align="left"><%=rsDET("NRO_CUOTA")&" "&strCustodio%></div></td>
		          <td><div align="left"><%=rsDET("FECHA_VENC")%></div></td>
		          <td><div align="right"><%=rsDET("ANTIGUEDAD")%></div></td>
		          <td><div align="center"><%=rsDET("NOM_TIPO_DOCUMENTO")%></div></td>
		          <td><div align="center"><%=rsDET("COD_REMESA")%></div></td>
		          <td align="right">
                  <%=FormatNumber(intValorCuota,0)%>
                  </td>

		          <%If Trim(strUsaInteres)="1" Then%>
		          	<td align="right"><%=FormatNumber(Round(intIntereses,0),0)%></td>
		          <%End If%>
		          <%If Trim(strUsaProtestos)="1" Then%>
		          	<td align="right"><%=FormatNumber(Round(intProtesto,0),0)%></td>
		          <%End If%>
		          <%If Trim(strUsaHonorarios)="1" Then%>
		          	<td align="right"><%=FormatNumber(Round(intHonorarios,0),0)%></td>
		          <%End If%>				  
		          <td align="right">
                  <%=FormatNumber(intAbono,0)%>
                  </td>
                  <td align="right">
                  <%=FormatNumber(intSaldo +Round(intIntereses,0) + Round(intHonorarios,0) + Round(intProtesto,0),0) %>
                  </td>
		          <INPUT TYPE="hidden" name="HD_INTERES_<%=rsDET("ID_CUOTA")%>" value="<%=Round(intIntereses,0)%>">
		          <INPUT TYPE="hidden" name="HD_HONORARIOS_<%=rsDET("ID_CUOTA")%>" value="<%=Round(intHonorarios,0)%>">
		          <INPUT TYPE="hidden" name="HD_PROTESTOS_<%=rsDET("ID_CUOTA")%>" value="<%=Round(intProtesto,0)%>">
                  <INPUT TYPE="hidden" name="HD_CAPITAL_<%=rsDET("ID_CUOTA")%>" value="<%=intSaldo%>">
				  <INPUT TYPE="hidden" name="HD_ABONO_<%=rsDET("ID_CUOTA")%>" value="<%=intAbono%>">
		         <%
					total_ValorCuota = total_ValorCuota + intValorCuota
					total_docs = total_docs + 1
				 %>
				 </tr>
				 <%
				 	rsDET.movenext
				 loop
				 vArrConcepto = split(strArrConcepto,";")
				 vArrID_CUOTA = split(strArrID_CUOTA,";")

				 intTamvConcepto = ubound(vArrConcepto)

				 %>

				</tbody> 
			  <%end if
			  rsDET.close
			  set rsDET=nothing
		  Else
		  %>
			<tr>
			<td  colspan="10" align="center">
				Deudor no posee documentos pendientes
			</td>
			</tr>
		 <%end if%>
        <script type="text/vbscript">
            marcar_boxes(true);
        </script>
                     
         
	</table>

<INPUT TYPE="hidden" NAME="hdintCapital" value="<%=intTotDeudaCapital%>">
<INPUT TYPE="hidden" NAME="hdintIndemComp" value="<%=intTotIndemComp%>">
<INPUT TYPE="hidden" NAME="hdintGastos" value="<%=intTotGastos%>">
<INPUT TYPE="hidden" NAME="hdintHonorarios" value="<%=intTotHonorarios%>">
<INPUT TYPE="hidden" NAME="hdintHonorarios" value="<%=intTotDeudaInteres%>">


<TABLE ALIGN="CENTER" id="" BORDER="1" width="100%">
			<TR>
					<TD valign="TOP" WIDTH="35%" >
						<TABLE ALIGN="CENTER" WIDTH="100%" BORDER="0">
							<TR HEIGHT="30">
								<td COLSPAN="2" ALIGN="CENTER" class="estilo_columna_individual">
									MONTO DE DEUDA
								</td>
							</TR>

							<TR HEIGHT="30">
								<TD WIDTH="150" ALIGN="LEFT" class="Estilo22">Capital: </TD>
								<TD><input name="TX_CAPITAL" value="<%=FormatNumber(intTotalSaldo,0)%>" type="text" size="10"  readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD ALIGN="LEFT" class="Estilo22">Interes: </TD>
								<TD><input name="TX_INTERES" type="text" size="10" value="<%=FormatNumber(intTotIntereses,0)%>"  readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD align="LEFT" class="Estilo22">Gastos Judiciales: </TD>
								<TD><input name="TX_GASTOS" type="text" size="10" value="<%=0%>"    readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD align="LEFT" class="Estilo22">Gastos Protestos: </TD>
								<TD><input name="TX_GASTOSPROTESTOS" type="text" size="10" value="<%=FormatNumber(intTotProtesto,0)%>" readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD align="LEFT" class="Estilo22">Indem.Comp.: </TD>
								<TD><input name="TX_INDEM_COMP" type="text" size="10" value="<%=0%>" readonly="readonly"></TD>
							</TR>

							<TR HEIGHT="30">
								<TD align="LEFT" class="Estilo22">Honorarios : </TD>
								<TD><input name="TX_HONORARIOS" type="text" size="10" value="<%=FormatNumber(intTotHonorarios,0)%>" readonly="readonly"></TD>
							</TR>
							<TR HEIGHT="30">
								<TD>&nbsp;</TD>
								<TD>______________</TD>
							</TR>
								<TR HEIGHT="30">
									<TD align="right" class="Estilo22">Total Deuda: </TD>
									<TD><input disabled value ="<%=FormatNumber(intTotalDeuda,0) %>" name="TX_TOTALDEUDA" type="text" size="10" readonly="readonly"></TD>
								</TR>
						</TABLE>


					</TD>

					<TD valign="TOP" WIDTH="35%" >

						<TABLE ALIGN="CENTER" WIDTH="100%" BORDER="0">
						  <TR HEIGHT="30">
							<TH COLSPAN="2" align="CENTER" class="estilo_columna_individual">
								DESCUENTOS
							</TH>

						  </TR>
						  <TR HEIGHT="30">
							  <TD width="150" ALIGN="LEFT" class="Estilo23">Capital:</TD>
							  <TD>
									%<input class="porc_desc_capital" id="porc_desc_capital" name="porc_desc_capital" type="text" size="3"   onblur="func_porc_desc_capital();"  value="0" maxlength="3">
									$<input class="desc_capital" id="desc_capital" name="desc_capital" type="text" size="8"  onblur="func_descuentos(this.value,'DESCUENTO');" value="0">
							  </TD>
						  </TR>
						  <TR HEIGHT="30">
							  <TD ALIGN="LEFT" class="Estilo23">Interes:</TD>
							  <TD>
									%<input class="porc_desc_interes" id="porc_desc_interes" name="porc_desc_interes" type="text" size="3"      onblur="func_porc_desc_interes();"  value="0" maxlength="3">
									$<input class="desc_interes" id="desc_interes" name="desc_interes" type="text" size="8"     onblur="func_descuentos(this.value,'INTERES');" value="0">
							  </TD>
						  </TR>
						  <TR HEIGHT="30">
								<TD ALIGN="LEFT" class="Estilo23">Gastos Judiciales:</TD>
								<TD>
									%<input DISABLED name="porc_desc_gastos" type="text" size="3"   onblur="func_porc_desc_gastos();" value="0" maxlength="3">
									$<input DISABLED name="desc_gastos" type="text" size="8"  value="0" onblur="func_descuentos(this.value,'JUDICIAL');">
								</TD>
						  </TR>
						   <TR HEIGHT="30">
								<TD ALIGN="LEFT" class="Estilo23">Gastos Protestos:</TD>
								<TD>
									%<input DISABLED name="porc_desc_gastosprotestos" type="text" size="3"  onblur="func_porc_gastosprotestos();" value="0" maxlength="3">
									$<input DISABLED name="GASTOS_PROTESTOS" type="text" size="8" value="0"    onblur="func_descuentos(this.value,'PROTESTO');"></TD>
						  </TR>

						 <TR HEIGHT="30">
								<TD ALIGN="LEFT" class="Estilo23">IndemComp:</TD>
								<TD>
									%<input DISABLED name="porc_desc_indemComp" type="text" size="3"   onblur="func_porc_indemComp();"  value="0" maxlength="3">
									$<input DISABLED name="desc_indemComp" type="text" size="8"   value="0" onblur="func_descuentos(this.value,'INDEMCOMP');">
								</TD>
						  </TR>

						  <TR HEIGHT="30">
							<td  align="LEFT" class="Estilo23"> Honorarios:</TD>
							<TD>
								%<input class="porc_desc_honorarios" id="porc_desc_honorarios"  name="porc_desc_honorarios" type="text" size="3"    onblur="func_porc_desc_honorarios();"value="0"  maxlength="3">
								$<input class="desc_honorarios" id="desc_honorarios"  name="desc_honorarios" type="text" size="8"  value="0"    onblur="func_descuentos(this.value,'HONORARIOS');">
							</TD>
						   </TR>
							<TR HEIGHT="30">
								<TD>&nbsp;</TD>
								<TD>______________</TD>
							</TR>

						   <TR HEIGHT="30">
								<TD>Total Deuda con Descuento</TD>
								<TD>$<input name="TX_TOTALDEUDA_DESC" type="text" size="10" disabled readonly="readonly" value ="<%=FormatNumber(intTotalDeuda,0) %>" ></TD>
							</TR>

						  </TABLE>

					</TD>
					<TD valign="TOP" WIDTH="32%" >

					  <TABLE ALIGN="CENTER" BORDER="0" WIDTH="100%" valign="TOP">
					  <TR HEIGHT="30">
							<TH COLSPAN=2 ALIGN="CENTER" class="estilo_columna_individual">
								MODALIDAD DEL PAGO
							</TH>
					  </TR>
						<TR valign="TOP">
							<TD HEIGHT="30" ALIGN="LEFT" class="Estilo23">
							Pie a cancelar:$
							</TD>
						</TR>
						<TR HEIGHT="30" valign="TOP">
							<TD ALIGN="LEFT" class="Estilo23">
							Abono Deuda&nbsp;
                            <%'If Trim(intOrigen)="PP" and strTipoPlan="NE" then intPorcPie=0%>
                            <input class="porc_capital_pie" pie="porc_capital_pie"  name="porc_capital_pie" type="text" size="2" value="<%=intPorcPie*100%>" onblur="func_porc_capital_pie();" maxlength="3"  >
							$<input class="pie" pie="pie" name="pie" type="text" size="10"   
                                    
									onblur="CalculateCapitalPercentageAndRefreshAgreement();"
                                    value="<%=FormatNumber(((intTotalDeuda*(intPorcPie*100))/100),0)%>"  
                                     maxlength="10" >
                        </TD>
						</TR>
					  <TR HEIGHT="30">
						  <TD width="200" ALIGN="LEFT" class="Estilo23" >Cantidad de cuotas: 
							<select name="cuotas" size="1" style="width:50px;" id="cuotas">
									<option value="-">-</option>
									<option value="1">1</option>
									<option value="2">2</option>
									<option value="3">3</option>
									<option value="4">4</option>
									<option value="5">5</option>
									<option value="6">6</option>
									<option value="7">7</option>
									<option value="8">8</option>
									<option value="9">9</option>
									<option value="10">10</option>
									<option value="11">11</option>
									<option value="12">12</option>
									<option value="13">13</option>
									<option value="14">14</option>
									<option value="15">15</option>
									<option value="16">16</option>
									<option value="17">17</option>
									<option value="18">18</option>
									<option value="19">19</option>
									<option value="20">20</option>
									<option value="21">21</option>
									<option value="22">22</option>
									<option value="23">23</option>
									<option value="24">24</option>
									<option value="25">25</option>
									<option value="26">26</option>
									<option value="27">27</option>
									<option value="28">28</option>
									<option value="29">29</option>
									<option value="30">30</option>
									<option value="31">31</option>
									<option value="32">32</option>
									<option value="33">33</option>
									<option value="34">34</option>
									<option value="35">35</option>
									<option value="36">36</option>

							  </select>
							  &nbsp;
							  Dia de Pago: 
							  <input name="TX_DIAPAGO" type="text" value="5" size="3" maxlength="5"  onkeyUp="return ValNumero(this);" >
						  </TD>
					  </TR>
						<TR>
							<TD align="LEFT" class="Estilo22">&nbsp;</TD>
						</TR>
						<TR>
							<TD align="LEFT" class="Estilo22">&nbsp;</TD>
						</TR>
												
						<TR>
							<TD align="LEFT" class="Estilo22">Total A Convenir: </TD>
						</TR>
						<TR>
							<TD><input name="TX_TOTALCONVENIO" type="text" size="10" readonly="readonly" 
                                    value ="<%=FormatNumber(intTotalDeuda - ((intTotalDeuda*(intPorcPie*100))/100) ,0) %>" ></TD>
						</TR>

					  </TABLE>
				 </TD>
			</TR>
</TABLE>

<BR>
<BR>

 <TABLE ALIGN="CENTER" BORDER="0" style="width:100%; border-top:1px solid #000;border-left:1px solid #000;border-bottom:1px solid #000;border-right:1px solid #000;">
 		<TR>
 			<TD ALIGN="left"><br>
 				OBSERVACIONES : &nbsp;&nbsp;&nbsp;<TEXTAREA NAME="TA_OBSERVACIONES" ROWS="3" COLS="100"><%=strRut%></TEXTAREA><br>&nbsp;
 			</TD>
 			<TD ALIGN="right"><br>
				<% If Trim(intOrigen)="CO" Then %>
  					<input type="submit" name="Submit" class="fondo_boton_100" value="Ver <%=session("NOMBRE_CONV_PAGARE")%>">
  				<% End If %>
  				<% If Trim(intOrigen)="PP" Then %>
					<input type="submit" name="Submit" class="fondo_boton_100"  value="Imprimir">

					<input Name="SubmitButton" class="fondo_boton_100" Value="Exportar Excel" Type="BUTTON" onClick="exportar();">

  				<% End If %>
  				
				<br>&nbsp;
 			</TD>
 		</TR>

 </TABLE>
 </div>
 <br>
 <br>
 <br>
</FORM>

<SCRIPT LANGUAJE=JavaScript>

function roundNumber(rnum, rlength) { // Arguments: number to round, number of decimal places
  var newnumber = Math.round(rnum/10);
  return (newnumber);
}

	function suma_capital(objeto , intValorSaldoCapital, intValorHonorarios, intValorIntereses, intValorProtestos){
		
		datos.porc_capital_pie.value = <%=intPorcPie*100%>;
        var Origen = document.getElementById("Origen").value

		if (datos.TX_CAPITAL.value == '') datos.TX_CAPITAL.value = 0
		if (datos.TX_HONORARIOS.value == '') datos.TX_HONORARIOS.value = 0
		if (datos.TX_INTERES.value == '') datos.TX_INTERES.value = 0
		if (datos.TX_GASTOSPROTESTOS.value == '') datos.TX_GASTOSPROTESTOS.value = 0
		if (datos.TX_GASTOS.value == '') datos.TX_GASTOS.value = 0
		if (datos.TX_INDEM_COMP.value == '') datos.TX_INDEM_COMP.value = 0


		datos.desc_honorarios.value = 0
		datos.desc_interes.value = 0
		datos.desc_capital.value = 0
		datos.desc_gastos.value = 0
		datos.porc_desc_capital.value = 0
        datos.porc_desc_interes.value = 0
        datos.porc_desc_gastos.value = 0
        datos.porc_desc_honorarios.value = 0
        
        LimpiaNumeros();
        

       if (objeto.checked == true)
        {
        	datos.TX_CAPITAL.value = parseInt(datos.TX_CAPITAL.value) + parseInt(intValorSaldoCapital);
            datos.TX_HONORARIOS.value = parseInt(datos.TX_HONORARIOS.value) + parseInt(intValorHonorarios);
			datos.TX_INTERES.value = parseInt(datos.TX_INTERES.value) + parseInt(intValorIntereses);
			datos.TX_GASTOSPROTESTOS.value = parseInt(datos.TX_GASTOSPROTESTOS.value) + parseInt(intValorProtestos);
			datos.TX_GASTOS.value = parseInt(datos.TX_GASTOS.value);
			datos.TX_INDEM_COMP.value = parseInt(datos.TX_INDEM_COMP.value);
			datos.TX_TOTALDEUDA.value = parseInt(datos.TX_CAPITAL.value) + parseInt(datos.TX_HONORARIOS.value) + parseInt(datos.TX_INTERES.value) + parseInt(datos.TX_GASTOSPROTESTOS.value) + parseInt(datos.TX_GASTOS.value) + parseInt(datos.TX_INDEM_COMP.value);
			datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_CAPITAL.value) + parseInt(datos.TX_HONORARIOS.value) + parseInt(datos.TX_INTERES.value) + parseInt(datos.TX_GASTOSPROTESTOS.value) + parseInt(datos.TX_GASTOS.value) + parseInt(datos.TX_INDEM_COMP.value);
            datos.pie.value = (roundNumber((<%=intPorcPie%> * parseInt(datos.TX_TOTALDEUDA.value)), 0));
			datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value) - parseInt(datos.desc_indemComp.value);
            FormateaTodosNumeros();
		}
		else
		{
        
			datos.TX_CAPITAL.value = parseInt(datos.TX_CAPITAL.value) - parseInt(intValorSaldoCapital);
           	datos.TX_HONORARIOS.value = parseInt(datos.TX_HONORARIOS.value) - parseInt(intValorHonorarios);
			datos.TX_INTERES.value = parseInt(datos.TX_INTERES.value) - parseInt(intValorIntereses);
			datos.TX_GASTOSPROTESTOS.value = parseInt(datos.TX_GASTOSPROTESTOS.value) - parseInt(intValorProtestos);
			datos.TX_TOTALDEUDA.value = parseInt(datos.TX_CAPITAL.value) + parseInt(datos.TX_HONORARIOS.value) + parseInt(datos.TX_INTERES.value) + parseInt(datos.TX_GASTOSPROTESTOS.value) + parseInt(datos.TX_GASTOS.value) + parseInt(datos.TX_INDEM_COMP.value);
            datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_CAPITAL.value) + parseInt(datos.TX_HONORARIOS.value) + parseInt(datos.TX_INTERES.value) + parseInt(datos.TX_GASTOSPROTESTOS.value) + parseInt(datos.TX_GASTOS.value) + parseInt(datos.TX_INDEM_COMP.value);			
			datos.pie.value = (roundNumber((<%=intPorcPie%> * parseInt(datos.TX_TOTALDEUDA.value)), 0));
			datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value) - parseInt(datos.desc_indemComp.value);
            FormateaTodosNumeros();
        }

       

                if (datos.TX_CAPITAL.value <= 0)
                {
                    datos.TX_CAPITAL.value =  0
                    datos.TX_INTERES.value =  0
                    datos.TX_GASTOS.value =  0
                    datos.TX_GASTOSPROTESTOS.value =  0
                    datos.TX_INDEM_COMP.value =  0
                    datos.TX_HONORARIOS.value =  0
                }
        
                if ((Origen != "CO") && (Origen != "PP")) 
                {
                    datos.pie.value ="0"
                    datos.TX_TOTALCONVENIO.value="0"
                } 


        MostrarPie();
	}

	function MostrarPie()
	{
		if (datos.CB_TIPO != undefined ||
			(parseInt(datos.TX_TOTALDEUDA.value) > 0 && datos.porc_capital_pie.value == 0))
		{
			datos.porc_capital_pie.value = <%=intPorcPie*100%>;

			CalculateCapital();
			
			CalculateTotalForAgreement();
		}
        
        if ((datos.CB_TIPO != undefined) && (datos.CB_TIPO.value == "NE" || datos.CB_TIPO.value == "CO" ))
        {
                datos.porc_capital_pie.value = 0;
                datos.pie.value = 0;
                datos.TX_TOTALCONVENIO.value = 0;
        
        }

	}

function ventanaBusqueda (URL){
	window.open(URL,"DATOS3","width=1050, height=700, scrollbars=yes, menubar=no, location=no, resizable=yes")
}

function exportar()
{
	document.datos.action = "exp_Plandepago.asp";
	document.datos.submit();
}


	function func_descuentos(objeto,origen){
        

         if (datos.TX_CAPITAL.value=="")
        {
            alert("Indique Capital a calcular")
            datos.porc_desc_capital.value="";
            return false;
        }

        LimpiaNumero(objeto);
        LimpiaNumeros();
        objeto = objeto.replace(/\./g, "")

        if (!/^([0-9])*$/.test(objeto))
        {
            alert("Ingrese Solo Numeros");
            return;
        }

        var por = 0 ;

        if (origen=="DESCUENTO")
        {
            if (parseInt(datos.TX_CAPITAL.value)< parseInt(objeto))
            {
                alert("Monto Capital Descuento no debe ser superior a Capital Monto de Deuda")
                func_porc_desc_capital();
                return false;
            }
             if (parseInt(datos.TX_CAPITAL.value) > 0)
            {

                if (datos.porc_desc_capital.value != '' && datos.desc_capital.value == '')
                {
                    func_porc_desc_capital(datos.porc_desc_capital.value);
                    return false;
                }
                por =  Math.round((objeto/datos.TX_CAPITAL.value)*100);
                datos.porc_desc_capital.value= por;
            }
        }

        if (origen=="INTERES")
        {
        
        if (parseInt(datos.TX_INTERES.value)< parseInt(objeto))
            {
                alert("Monto Interes  Descuento no debe ser superior a Interes Monto de Deuda")
                func_porc_desc_interes();
                return false;
            }
               if (parseInt(datos.TX_INTERES.value) > 0)
            {

              if (datos.porc_desc_interes.value != '' && datos.desc_interes.value == '')
                {
                    func_porc_desc_interes(datos.porc_desc_interes.value);
                    return false;
                } 

            por =  Math.round((objeto/datos.TX_INTERES.value)*100);
            datos.porc_desc_interes.value= por;
            }
        }

           if (origen=="JUDICIAL")
        {
            if (parseInt(datos.TX_GASTOS.value)< parseInt(objeto))
            {
                alert("Monto Gastos Judiciales Descuento no debe ser superior a Gastos Judiciales Monto de Deuda")
                func_porc_desc_gastos();
                return false;
            }
            if (parseInt(datos.TX_GASTOS.value) > 0)
            {

              if (datos.porc_desc_gastos.value != '' && datos.desc_gastos.value == '')
                {
                    func_porc_desc_gastos(datos.porc_desc_gastos.value);
                    return false;
                } 


                por =  Math.round((objeto/datos.TX_GASTOS.value)*100);
                datos.porc_desc_gastos.value= por;
            }
        }

           if (origen=="PROTESTO")
        {
            if (parseInt(datos.TX_GASTOSPROTESTOS.value)< parseInt(objeto))
            {
                alert("Monto Gastos Protestos Descuento no debe ser superior a Gastos Protestos Monto de Deuda")
                func_porc_gastosprotestos();
                return false;
            }

             if (datos.porc_desc_gastosprotestos.value != '' && datos.GASTOS_PROTESTOS.value == '')
                {
                    func_porc_gastosprotestos(datos.porc_desc_gastosprotestos.value);
                    return false;
                } 

                if (parseInt(datos.TX_GASTOSPROTESTOS.value) > 0)
                {
                    por =  Math.round((objeto/datos.TX_GASTOSPROTESTOS.value)*100);
                    datos.porc_desc_gastosprotestos.value= por;
                }
        }


        if (origen=="INDEMCOMP")
        {
            if (parseInt(datos.TX_INDEM_COMP.value)< parseInt(objeto))
            {
                alert("Monto IndemComp Descuento no debe ser superior a IndemComp Monto de Deuda")
                func_porc_indemComp();
                return false;
            }
            
            if (parseInt(datos.TX_INDEM_COMP.value) > 0)
            {
               if (datos.porc_desc_indemComp.value != '' && datos.desc_indemComp.value == '')
                {
                    func_porc_indemComp(datos.porc_desc_indemComp.value);
                    return false;
                } 

                 por =  Math.round((objeto/datos.TX_INDEM_COMP.value)*100);
                 datos.porc_desc_indemComp.value= por;
            }
        }

          if (origen=="HONORARIOS")
        {
            if (parseInt(datos.TX_HONORARIOS.value)< parseInt(objeto))
            {
                alert("Monto Honorarios Descuento no debe ser superior a Honorarios Monto de Deuda")
                func_porc_desc_honorarios();
                return false;
            }

            if (parseInt(datos.TX_HONORARIOS.value) > 0)
            {
                
                if (datos.porc_desc_honorarios.value != '' && datos.desc_honorarios.value == '')
                {
                    func_porc_desc_honorarios(datos.porc_desc_honorarios.value);
                     MostrarPie();
                    return false;
                } 

                por =  Math.round((parseInt(objeto)/parseInt(datos.TX_HONORARIOS.value))*100);
                datos.porc_desc_honorarios.value= parseInt(por);

                MostrarPie();
            }
        }

        if (origen=="PIE") 
        {

   
            if (datos.porc_capital_pie.value != '0' && datos.porc_capital_pie.value !== '')// && datos.pie.value == '')
            {
			
                func_porc_capital_pie(datos.porc_capital_pie.value);
                return false;
            }else if (datos.pie.value == '0')
            {
             datos.porc_capital_pie.value =0;
             datos.pie.value =0;
             datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
             FormateaTodosNumeros();
             return false;
            }

        if (parseInt(datos.TX_TOTALDEUDA_DESC.value) < parseInt(objeto))
        {
        
            

            alert("Pie a cancelar no debe ser superior a Total Deuda con Descuento")
            datos.TX_TOTALCONVENIO.value = FormatearNumero(datos.TX_TOTALDEUDA.value);
            func_porc_capital_pie();
            return false;
        }
        
        }
        

		if (datos.TX_CAPITAL.value == '') datos.TX_CAPITAL.value = 0;
		if (datos.desc_interes.value == '') datos.desc_interes.value = 0;
        if (datos.desc_gastos.value == '') datos.desc_gastos.value = 0;
        if (datos.GASTOS_PROTESTOS.value == '') datos.GASTOS_PROTESTOS.value = 0;
        if (datos.desc_indemComp.value == '') datos.desc_indemComp.value = 0;
        if (datos.desc_honorarios.value == '') datos.desc_honorarios.value = 0;
        if (datos.desc_capital.value == '') datos.desc_capital.value = 0;
		datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);

		datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
		

        var Por = 0
        Por  = (datos.pie.value*100)/(datos.TX_TOTALDEUDA_DESC.value)
        datos.porc_capital_pie.value = Math.round(Por)
        
        FormateaTodosNumeros();

        MostrarPie();
	}


    

	function func_porc_desc_capital(){
		

        if (datos.TX_CAPITAL.value=="")
        {
            alert("Indique Capital a calcular")
            datos.porc_desc_capital.value="";
            return false;
        }

        LimpiaNumeros()
        

        if (!/^([0-9])*$/.test(datos.porc_desc_capital.value))
        {
            alert("% Descuento Ingrese Solo Numeros");
            datos.desc_capital.value="";
             FormateaTodosNumeros();
            return;
        }
        
        
        if (datos.porc_desc_capital.value == '' && datos.desc_capital.value != '')
        {
            func_descuentos(datos.desc_capital.value,'DESCUENTO');
            return false;
        } 

        if (datos.porc_desc_capital.value == 0) 
        {
            datos.porc_desc_capital.value =0;
            datos.desc_capital.value =0;
            datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
            datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
            FormateaTodosNumeros();
		    MostrarPie();
            return ;
        } 


		var Total = Math.round(parseInt(datos.TX_CAPITAL.value) * parseInt(datos.porc_desc_capital.value) / 100);
        

        if (parseInt(Total) <= 0) 
        {
            datos.porc_desc_capital.value = 0;
            datos.desc_capital.value = 0;
        }

		if (parseInt(datos.TX_CAPITAL.value) < parseInt(Total))
        {
            alert("Monto Capital Descuento no debe ser mayor a Capital Monto de Deuda")
            func_descuentos(datos.desc_capital.value,'DESCUENTO');
            return false;
        }

        datos.desc_capital.value  = Total
		datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
        datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);

        FormateaTodosNumeros();

		func_porc_capital_pie();
        MostrarPie();
	}

	function func_porc_desc_interes()
    {

    

         if (datos.TX_CAPITAL.value=="")
        {
            alert("Indique Interes a calcular")
            datos.porc_desc_interes.value="";
            return false;
        }
        
        LimpiaNumeros();

        if (!/^([0-9])*$/.test(datos.porc_desc_interes.value))
        {
            alert("% Interes Ingrese Solo Numeros");
            datos.porc_desc_interes.value="";
            FormateaTodosNumeros();
            return;
        }

        
          if (datos.porc_desc_interes.value == '' && datos.desc_interes.value != '')
        {
            func_descuentos(datos.desc_interes.value,'INTERES');
            return false;
        } 

        if (datos.porc_desc_interes.value == 0) 
        {
            datos.porc_desc_interes.value =0;
            datos.desc_interes.value =0;
            datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
            datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
            FormateaTodosNumeros();
		    MostrarPie();
            return ;
        } 


        var Total = Math.round(parseInt(datos.TX_INTERES.value) * parseInt(datos.porc_desc_interes.value) / 100)

       if (parseInt(Total) <= 0) 
        {
            datos.desc_interes.value = 0;
            datos.desc_interes.value=0
        }

        if (parseInt(datos.TX_INTERES.value) < parseInt(Total))
        {
            alert("Monto Interes Descuento no debe ser mayor a Interes Monto de Deuda")
            func_descuentos(datos.desc_interes.value,'INTERES');
            return false;
        }

		datos.desc_interes.value =Total

		datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);

		datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);	
		FormateaTodosNumeros();
        func_porc_capital_pie();
        MostrarPie();
	}

	function func_porc_desc_honorarios(){
        

       if (datos.TX_CAPITAL.value=="")
        {
            alert("Indique Honorarios a calcular")
            datos.porc_desc_honorarios.value="";
            return false;
        }

        
          LimpiaNumeros();
        if (!/^([0-9])*$/.test(datos.porc_desc_honorarios.value))
        {
            alert("% Honorarios Ingrese Solo Numeros");
            datos.porc_desc_honorarios.value="";
            FormateaTodosNumeros();
            return;
        }
  
       if (datos.porc_desc_honorarios.value == '' && datos.desc_honorarios.value != '')
        {
            func_descuentos(datos.desc_honorarios.value,'HONORARIOS');
            return false;
        } 

        if (datos.porc_desc_honorarios.value == 0) 
        {
            datos.porc_desc_honorarios.value =0;
            datos.desc_honorarios.value =0;
            datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
            datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);
            FormateaTodosNumeros();
		    MostrarPie();
            return ;
        } 
        

      

		var Total  = Math.round(parseInt(datos.TX_HONORARIOS.value) * parseInt(datos.porc_desc_honorarios.value) / 100);
	    
        if (parseInt(Total) <= 0) 
        {
            datos.porc_desc_honorarios.value = 0;
            datos.desc_honorarios.value=0
        }


        if (parseInt(datos.TX_HONORARIOS.value) < parseInt(Total))
        {
            alert("Monto Honorarios Descuento no debe ser mayor a Honorarios Monto de Deuda")
            func_descuentos(datos.desc_honorarios.value,'HONORARIOS');
            return false;
        }

        datos.desc_honorarios.value  = Total
        datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);

		datos.TX_TOTALDEUDA_DESC.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value);		
		
        FormateaTodosNumeros();
        func_porc_capital_pie();
        MostrarPie();
	}





	function func_porc_capital_pie()
    {

   /*
    if (datos.pie.value == '') datos.pie.value = 0;
    if (datos.porc_capital_pie.value == '') datos.porc_capital_pie.value = 0;
   */
    LimpiaNumeros();

    if (datos.porc_capital_pie.value > 100 || datos.porc_capital_pie.value < 0)
    {
        alert("% pie no Valido");
		CalculateCapitalPercentageAndRefreshAgreement();
        FormateaTodosNumeros();
        return false;
    }


        if (datos.porc_capital_pie.value == '0' && datos.pie.value != '0')
        {
            datos.porc_capital_pie.value ='0';
            datos.pie.value = '0';
            datos.TX_TOTALCONVENIO.value = (parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value));
            FormateaTodosNumeros();
            return false;
        } 


        if (datos.porc_capital_pie.value == '' && datos.pie.value != '')
        {
			//alert();
            func_descuentos(datos.desc_capital.value,'PIE');
			return false;
        } 



	    datos.pie.value = (parseInt(datos.porc_capital_pie.value) * parseInt(datos.TX_TOTALDEUDA_DESC.value))/ 100;
	    datos.pie.value = Math.round(datos.pie.value);
		
	    datos.TX_TOTALCONVENIO.value = (parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value));
		
        FormateaTodosNumeros();
        /*MostrarPie();*/
	}

	function func_actualiza_total_deuda() {
		datos.TX_TOTALDEUDA.value = parseInt(datos.TX_CAPITAL.value) + parseInt(datos.TX_HONORARIOS.value) + parseInt(datos.TX_INTERES.value) + parseInt(datos.TX_GASTOSPROTESTOS.value) + parseInt(datos.TX_INDEM_COMP.value) + parseInt(datos.TX_GASTOS.value);
		datos.pie.value = (roundNumber((<%=intPorcPie%> * parseInt(datos.TX_TOTALDEUDA.value)), 0));
		datos.TX_TOTALCONVENIO.value = parseInt(datos.TX_TOTALDEUDA.value) - parseInt(datos.pie.value) - parseInt(datos.desc_gastos.value) - parseInt(datos.desc_capital.value) - parseInt(datos.desc_interes.value) - parseInt(datos.desc_honorarios.value) - parseInt(datos.desc_indemComp.value);
	}


  function validardatos(formulario)
     {

      var cont = 0
    
     for (var i = 1; i < document.getElementById('tbl_Procesa').rows.length; i++) {
              var Id_Cuota = document.getElementById('tbl_Procesa').rows[i].cells[1].innerHTML;
                   var  chk = document.getElementById("CH_" + Id_Cuota).checked;
                    
	             if (chk == true) {
                      cont = cont + 1
                      }
	           
    }

    if (cont==0)
    {
    alert("Indique Cuota a Cancelar")
    return false;
    }


         var enviardatos=true;
      
         
         
         if(formulario.CB_SEDE.value==""){
			alert("Debe seleccionar una Sede");
			return false;
		  }

           
       var origen = document.getElementById("Origen").value;


       if (origen == "CO")
       {
             if (datos.cuotas.value=="-")
             {
         	    alert("Indique Cantidad de Cuotas");
			    return false;
             }
             if(formulario.pie.value=="" || formulario.pie.value=="0"){
				alert("Error de ingreso: Debe Ingresar el pie");
				return false;
			  }
        }

		  if(datos.CB_TIPO.value==""){
			alert("Debe seleccionar Tipo");
			return false;
		  }

		  if(datos.CB_FPAGO.value==""){
			alert("Debe seleccionar Forma de Pago");
			return false;
		  }
		  
		   if(datos.CB_TIPO.value == "RC" || datos.CB_TIPO.value == "RP" || datos.CB_TIPO.value == "RL"){

            if(formulario.pie.value=="" || formulario.pie.value=="0"){
				alert("Error de ingreso: Debe Ingresar el pie");
				return false;
			  }
			 
			  if(formulario.desc_capital.value=='-' || formulario.desc_indemComp.value=='-' || formulario.desc_honorarios.value=='-'){
				alert("Error de ingreso: Debe Ingresar los descuentos");
				return false;
			  }
			  if(formulario.cuotas.value=='-'){
				alert("Error de ingreso: Debe Ingresar cantidad de cuotas");
				return false;
			 }
		}

		return true;

	}

function val_tipo(){


   if (datos.CB_TIPO != undefined)
     {
     LimpiaNumeros();

            if(datos.CB_TIPO.value=="CO" || datos.CB_TIPO.value == "NE") {
		            datos.porc_capital_pie.value = 0;
		            datos.pie.value = 0;
		            datos.TX_DIAPAGO.value = '';
		            datos.TX_TOTALCONVENIO.value = 0;
		            datos.porc_capital_pie.disabled = true;
		            datos.pie.disabled = true;
		            datos.cuotas.disabled = true;
		            datos.TX_DIAPAGO.disabled = true;
		            datos.TX_TOTALCONVENIO.disabled = true;
		            /*OcultarFilas('OpPie');*/
	              }


	              if(datos.CB_TIPO.value == "RC" || datos.CB_TIPO.value == "RP" || datos.CB_TIPO.value == "RL"){
	  		            datos.porc_capital_pie.value = <%=intPorcPie*100%>;
	  		            datos.pie.value = '';
	  		            //datos.cuotas.value = '';
	  		            datos.TX_DIAPAGO.value = 5;
	  		            datos.TX_TOTALCONVENIO.value = '';
	  		            datos.porc_capital_pie.disabled = false;
	  		            datos.pie.disabled = false;
	  		            datos.cuotas.disabled = false;
	  		            datos.TX_DIAPAGO.disabled = false;
	  		            datos.TX_TOTALCONVENIO.disabled = false;
	  		            MostrarFilas('OpPie');
                        func_porc_capital_pie();
	              }

                  FormateaTodosNumeros();
}



}


function FormaPago2(){
	 if(datos.CB_TIPO.value == "RC" || datos.CB_TIPO.value == "RP" || datos.CB_TIPO.value == "RL"){
		datos.porc_capital_pie.value = '';
		datos.pie.value = '';
		datos.cuotas.value = '';
		datos.TX_DIAPAGO.value = '';
		datos.TX_TOTALCONVENIO.value = '';

		datos.porc_capital_pie.disabled = true;
		datos.pie.disabled = true;
		datos.cuotas.disabled = true;
		datos.TX_DIAPAGO.disabled = true;
		datos.TX_TOTALCONVENIO.disabled = true;
	  }
}

function FORMA_PAGO(){
	//alert(datos.CB_TIPO.value);
	//alert(datos.CB_FPAGO.value);
	 if(datos.CB_TIPO.value == "" || datos.CB_FPAGO.value == ""){
        OcultarFilas('principal2');
        

	  }
	 else {
		MostrarFilas('principal2');
	  }


}

function marcar_boxes(){

		desmarcar_boxes();

		<% For i=1 TO intTamvConcepto %>
			document.forms[0].<%=vArrConcepto(i)%>.checked=true; 
			suma_capital(document.forms[0].<%=vArrConcepto(i)%>,document.forms[0].HD_CAPITAL_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_HONORARIOS_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_INTERES_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_PROTESTOS_<%=vArrID_CUOTA(i)%>.value);
		<% Next %>

		datos.porc_capital_pie.value = <%=intPorcPie*100%>;

     if (datos.CB_TIPO != undefined)
     {

		 if(datos.CB_TIPO.value=="CO" || datos.CB_TIPO.value == "NE") {
			datos.porc_capital_pie.value = '0';
			datos.pie.value = '0';
			datos.TX_DIAPAGO.value = '';
			datos.TX_TOTALCONVENIO.value = '';
            datos.porc_capital_pie.value='';
		  }
      }
}

function desmarcar_boxes(){
		
		<% For i=1 TO intTamvConcepto %>
			document.forms[0].<%=vArrConcepto(i)%>.checked=false;
			suma_capital(document.forms[0].<%=vArrConcepto(i)%>,document.forms[0].HD_CAPITAL_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_HONORARIOS_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_INTERES_<%=vArrID_CUOTA(i)%>.value,document.forms[0].HD_PROTESTOS_<%=vArrID_CUOTA(i)%>.value);
		<% Next %>
		datos.TX_CAPITAL.value = 0;
		datos.TX_HONORARIOS.value = 0;
		datos.pie.value = 0;
		datos.TX_TOTALCONVENIO.value = 0;
		datos.TX_GASTOS.value = 0;
		datos.TX_INDEM_COMP.value = 0;
		datos.TX_INTERES.value = 0;
		datos.TX_GASTOSPROTESTOS.value = 0;
		datos.TX_TOTALDEUDA.value = 0;
        datos.porc_capital_pie.value = 0;
		datos.TX_TOTALDEUDA_DESC.value = 0;
}

function VisualizarPie(){

}



function cargaTipo(subCat)
{
	var comboBox = document.getElementById('CB_FPAGO');
	comboBox.options.length = 0;

	if (subCat=='NE') {
		var newOption = new Option('SELECCIONAR', '');
		comboBox.options[comboBox.options.length] = newOption;
		var newOption = new Option('NO ESPECIFICADO', 'NE');comboBox.options[comboBox.options.length] = newOption;
	}
	else
	{
			var newOption = new Option('SELECCIONAR', '');
			comboBox.options[comboBox.options.length] = newOption;
			<%
			strSql="SELECT * FROM CAJA_FORMA_PAGO WHERE PLAN_PAGO = 1"
			''Response.write "sql=" & strSql
			set rsGestion=Conn.execute(strSql)
			If Not rsGestion.Eof Then
				Do While Not rsGestion.Eof
					%>
						var newOption = new Option('<%=rsGestion("DESC_FORMA_PAGO")%>', '<%=rsGestion("ID_FORMA_PAGO")%>');comboBox.options[comboBox.options.length] = newOption;
					<%
					rsGestion.movenext
				Loop
			Else
			%>
				var newOption = new Option('SIN TIPO', '');
				comboBox.options[comboBox.options.length] = newOption;
			<%
			End if
			
            %>
	}

	/*FORMA_PAGO();*/

}

</SCRIPT>
<script type="text/javascript">
	function MostrarFilas(Fila) {
		var contac_fila ="#"+Fila
		$(contac_fila).css('display', 'block')

	}

	function OcultarFilas(Fila) {
		/*desmarcar_boxes()*/
		var contac_fila ="#"+Fila
		$(contac_fila).css('display', 'none')

	}

	$(document).ready(function(){
		
        val_tipo();

     

	})

	function bt_limpia(){
		$('#CB_SEDE').val("")
		$('#CB_TIPO').val("")
		$('#CB_FPAGO').val("")	

		datos.porc_desc_capital.value = '0'
		datos.porc_desc_interes.value = '0'
		datos.porc_desc_gastos.value = '0'
		datos.porc_desc_gastosprotestos.value = '0'
		datos.porc_desc_indemComp.value = '0'
		datos.porc_desc_honorarios.value = '0'

		datos.desc_capital.value = '0'
		datos.desc_interes.value = '0'
		datos.desc_gastos.value = '0'
		datos.GASTOS_PROTESTOS.value = '0'
		datos.desc_indemComp.value = '0'
		datos.desc_honorarios.value = '0'

		LimpiaNumeros();		
		//window.location.reload()
		//$('input[type="text"]').val("0")	
		$('#principal2').css('display', 'none')	

	}
</script>

<script type="text/javascript">

    function FormatearNumero(numero) {
        /*alert(numero);*/
        var number = new String(parseInt(numero.toString().replace(/\./g, "")));
        var result = '';
        while (number.length > 3) {
            result = '.' + number.substr(number.length - 3) + result;
            number = number.substring(0, number.length - 3);
        }
        result = number + result;
        /*alert(result);*/
        return result;

    };

    function LimpiaNumeros() {
        datos.TX_CAPITAL.value = datos.TX_CAPITAL.value.replace(/\./g, "")
        datos.TX_HONORARIOS.value = datos.TX_HONORARIOS.value.replace(/\./g, "")
        datos.TX_INTERES.value = datos.TX_INTERES.value.replace(/\./g, "")
        datos.TX_GASTOSPROTESTOS.value = datos.TX_GASTOSPROTESTOS.value.replace(/\./g, "")
        datos.TX_GASTOS.value = datos.TX_GASTOS.value.replace(/\./g, "")
        datos.TX_INDEM_COMP.value = datos.TX_INDEM_COMP.value.replace(/\./g, "")
        datos.TX_TOTALDEUDA.value = datos.TX_TOTALDEUDA.value.replace(/\./g, "")
        datos.TX_TOTALDEUDA_DESC.value = datos.TX_TOTALDEUDA_DESC.value.replace(/\./g, "")
        datos.pie.value = datos.pie.value.replace(/\./g, "")
        datos.TX_TOTALCONVENIO.value = datos.TX_TOTALCONVENIO.value.replace(/\./g, "")
        datos.desc_interes.value = datos.desc_interes.value.replace(/\./g, "")
        datos.desc_gastos.value = datos.desc_gastos.value.replace(/\./g, "")
        datos.GASTOS_PROTESTOS.value = datos.GASTOS_PROTESTOS.value.replace(/\./g, "")
        datos.desc_indemComp.value = datos.desc_indemComp.value.replace(/\./g, "")
        datos.desc_honorarios.value = datos.desc_honorarios.value.replace(/\./g, "")
        datos.desc_capital.value = datos.desc_capital.value.replace(/\./g, "")
         
    };

    function FormateaTodosNumeros() {
        datos.TX_CAPITAL.value = FormatearNumero(datos.TX_CAPITAL.value);
        datos.TX_HONORARIOS.value = FormatearNumero(datos.TX_HONORARIOS.value);
        datos.TX_INTERES.value = FormatearNumero(datos.TX_INTERES.value);
        datos.TX_GASTOSPROTESTOS.value = FormatearNumero(datos.TX_GASTOSPROTESTOS.value);
        datos.TX_GASTOS.value = FormatearNumero(datos.TX_GASTOS.value);
        datos.TX_INDEM_COMP.value = FormatearNumero(datos.TX_INDEM_COMP.value);
        datos.TX_TOTALDEUDA.value = FormatearNumero(datos.TX_TOTALDEUDA.value);
        datos.TX_TOTALDEUDA_DESC.value = FormatearNumero(datos.TX_TOTALDEUDA_DESC.value);
        datos.pie.value = FormatearNumero(datos.pie.value);
        datos.TX_TOTALCONVENIO.value = FormatearNumero(datos.TX_TOTALCONVENIO.value);
        datos.desc_interes.value = FormatearNumero(datos.desc_interes.value);
        datos.desc_gastos.value = FormatearNumero(datos.desc_gastos.value);
        datos.GASTOS_PROTESTOS.value = FormatearNumero(datos.GASTOS_PROTESTOS.value);
        datos.desc_indemComp.value = FormatearNumero(datos.desc_indemComp.value);
        datos.desc_honorarios.value = FormatearNumero(datos.desc_honorarios.value);
        datos.desc_capital.value = FormatearNumero(datos.desc_capital.value);
    };

    function Solo_Numerico(variable) {
        Numer = parseInt(variable);
        if (isNaN(Numer)) {
            return "";
        }
        return Numer;
    }
    function ValNumero(Control) {
        Control.value = LimpiaNumero(Control.value);
        Control.value = Solo_Numerico(Control.value);
    }


    function LimpiaNumero(numero) {

        var numero = new String(numero);
        var result = numero.replace(/\./g, "")
        return result;
    };
	
	
	
	
	function CalculateCapitalPercentageAndRefreshAgreement()
	{
		var pie = parseInt(LimpiaNumero($("input[name='pie']").val()));
		
		var totalDeudaConDescuento = parseInt(LimpiaNumero($("input[name='TX_TOTALDEUDA_DESC']").val()));
	
		if (pie <= totalDeudaConDescuento) {
			$("input[name='porc_capital_pie']").val(Math.round(pie / totalDeudaConDescuento * 100, 0));
			
			$("input[name='pie']").val(FormatearNumero(pie));
			
			CalculateTotalForAgreement();
		}
		else {
			alert('El monto del pie no puede ser mayor al total deuda con descuento.');
			
			CalculateCapital();
		}
	}
	
	function CalculateCapital()
	{
		var porcentajePie = parseInt(LimpiaNumero($("input[name='porc_capital_pie']").val()));
		
		var totalDeudaConDescuento = parseInt(LimpiaNumero($("input[name='TX_TOTALDEUDA_DESC']").val()));
		
		$("input[name='pie']").val(FormatearNumero(Math.round(porcentajePie / 100 * totalDeudaConDescuento, 0)));
	}
	
	function CalculateTotalForAgreement()
	{
		var totalDeuda = parseInt(LimpiaNumero($("input[name='TX_TOTALDEUDA']").val()));
		
		var pie = parseInt(LimpiaNumero($("input[name='pie']").val()));
		
		var descuentosGastosJudiciales = parseInt(LimpiaNumero($("input[name='desc_gastos']").val()));
		
		var descuentosCapital = parseInt(LimpiaNumero($("input[name='desc_capital']").val()));
		
		var descuentosIntereses = parseInt(LimpiaNumero($("input[name='desc_interes']").val()));
		
		var descuentosHonorarios = parseInt(LimpiaNumero($("input[name='desc_honorarios']").val()));
		
		var descuentosIndemComp = parseInt(LimpiaNumero($("input[name='desc_indemComp']").val()));
		
		$("input[name='TX_TOTALCONVENIO']").val(FormatearNumero(totalDeuda - pie - descuentosGastosJudiciales - descuentosCapital - descuentosIntereses - descuentosHonorarios - descuentosIndemComp));
	}

</script>

</body>
</html>



<script type="text/javascript">
    $(".porc_desc_capital").numeric();
    $(".desc_capital").numeric();
    $(".porc_desc_interes").numeric();
    $(".desc_interes").numeric();
    $(".porc_desc_honorarios").numeric();
    $(".desc_honorarios").numeric();
    $(".porc_capital_pie").numeric();
    $(".pie").numeric();
    $("#remove").click(
		function (e) {
		    e.preventDefault();
		    $(".porc_desc_capital,.desc_capital,.desc_interes,.porc_desc_interes,.porc_desc_honorarios,.desc_honorarios,.pie,.porc_capital_pie").removeNumeric();
		}
	);
	</script>