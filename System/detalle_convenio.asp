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
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	cod_caja=Session("intCodUsuario")

	AbrirSCG()
	codpago 	= request("TX_PAGO")
	strrut 		=request("TX_RUT")
	usuario 	= request("cmb_usuario")
	if usuario 	= "" then usuario = "0"
	termino 	= request("termino")
	inicio 		= request("inicio")
	resp 		= request("resp")
	If Trim(resp)="" Then resp = "si"
	if Trim(inicio) = "" Then
		inicio = TraeFechaMesActual(Conn,-1)
		inicio = "01" & Mid(inicio,3,10)
		inicio = TraeFechaActual(Conn)
	End If
	if Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If


	strTipo 	= request("strTipo")
	CLIENTE 	= REQUEST("CLIENTE")
	If CLIENTE = "" Then
		CLIENTE = session("ses_codcli")
	End If
	'Response.write "CLIENTE=" & CLIENTE
	'hoy=date

	'response.write(hoy)
%>

	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo27 {color: #FFFFFF}
	-->
	</style>

	<script language="JavaScript" src="../javascripts/cal2.js"></script>
	<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
	<script language="JavaScript" src="../javascripts/validaciones.js"></script>
	<script src="../javascripts/SelCombox.js"></script>
	<script src="../javascripts/OpenWindow.js"></script>

	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">

	<script language="JavaScript " type="text/JavaScript">
	$(document).ready(function(){

		$('#inicio').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
		$('#termino').datepicker( {changeMonth: true,changeYear: true, dateFormat: 'dd/mm/yy'})
	 
	})


	function Refrescar()
	{
		resp='no'
		datos.action = "detalle_convenio.asp?resp="+ resp +"";
		datos.submit();
	}

	function Ingresa()
	{
		with( document.datos )
		{
			action = "detalle_convenio.asp";
			submit();
		}
	}

	function envia()
	{
		//datos.TX_RUT.value='';
		//datos.TX_PAGO.value='';
		resp='si'
		datos.action = "detalle_convenio.asp?resp="+ resp +"";
		datos.submit();
	}

	function Buscar()
	{

		if (datos.TX_RUT.value == '' && datos.TX_PAGO.value =='' ) {
			alert('Debe ingresar rut o codigo del <%=session("NOMBRE_CONV_PAGARE")%>');
			return
		}
		resp='123'
		datos.action = "detalle_convenio.asp?resp="+ resp +"";
		datos.submit();
	}

	function envia_excel(URL){
		window.open(URL,"INFORMACION","width=200, height=200, scrollbars=yes, menubar=yes, location=yes, resizable=yes")
	}



	</script>


</head>
<body>
<form name="datos" method="post">
<input name="strTipo" id="strTipo" type="hidden" value="<%=strTipo%>">
<div class="titulo_informe">LISTADO DE <%=UCASE(session("NOMBRE_CONV_PAGARE"))%>S</div>
<br>
<table width="90%" height="500" border="0" align="center">
  <tr>
    <td valign="top">
	<table width="100%" border="0" bordercolor="#999999">
		<!--<tr height="20" class="Estilo8">
	        <td>RUT: </td>
			<td><INPUT TYPE="text" NAME="TX_RUT" value="<%=strrut%>"></td>
		     <td>CODIGO <%=UCASE(session("NOMBRE_CONV_PAGARE"))%>: </td>
			<td><INPUT TYPE="text" NAME="TX_PAGO" value="<%=codpago%>"></td>
			<td>
				<input type="Button" class="fondo_boton_100" name="Submit" value="Buscar" onClick="Buscar();">
			</td>
		</tr>-->
		<%if codpago <> "" then
			if sucursal = "0" then
				'sql="select ID_CONVENIO, CONVERT(varchar(10), FECHA_INGRESO, 103) AS FECHA_INGRESO,comp_ingreso, sucursal.des_suc, cliente.DESCRIPCION, RUT_DEUDOR, total_cliente, total_emp, caja_tipo_pago.desc_tipo_pago,USR_INGRESO from CONVENIO_ENC,cliente,caja_tipo_pago,sucursal where cliente.COD_CLIENTE =CONVENIO_ENC.cod_cliente and caja_tipo_pago.id_tipo_pago = CONVENIO_ENC.tipo_pago and sucursal.cod_suc = CONVENIO_ENC.sucursal and CONVENIO_ENC.ID_CONVENIO=" & codpago & " and CONVENIO_ENC.estado_caja='A' and CONVENIO_ENC.estado = 1"
			else
				'sql="select ID_CONVENIO, CONVERT(varchar(10), FECHA_INGRESO, 103) AS FECHA_INGRESO,comp_ingreso, sucursal.des_suc, cliente.DESCRIPCION, RUT_DEUDOR, total_cliente, total_emp, caja_tipo_pago.desc_tipo_pago,USR_INGRESO from CONVENIO_ENC,cliente,caja_tipo_pago,sucursal where cliente.COD_CLIENTE =CONVENIO_ENC.cod_cliente and caja_tipo_pago.id_tipo_pago = CONVENIO_ENC.tipo_pago and sucursal.cod_suc = CONVENIO_ENC.sucursal and CONVENIO_ENC.ID_CONVENIO=" & codpago & " and CONVENIO_ENC.sucursal='" & sucursal & "' and CONVENIO_ENC.estado_caja='A' and CONVENIO_ENC.estado = 1"
			end if
		end if
		%>
	</table>
	<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
		<thead>
	      <tr height="20">
		  <td>CLIENTE</td>
		 	<td>USUARIO</td>
			<td>DESDE</td>
			<td>HASTA</td>
			<Td></td>
	      </tr>
	  	</thead>
		  <tr bordercolor="#999999" class="Estilo8">
		  <td>

		<select name="CLIENTE" ID = "CLIENTE" width="15" onchange="tipopago()">
		<option value="0">SELECCIONAR</option>
		<%
		ssql="SELECT * FROM CLIENTE WHERE ACTIVO=1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ")"
		set rsCLI=Conn.execute(ssql)
		if not rsCLI.eof then
			do until rsCLI.eof
			%>
			<option value="<%=rsCLI("COD_CLIENTE")%>"
			<%if Trim(cliente)=Trim(rsCLI("COD_CLIENTE")) then
				response.Write("Selected")
			end if%>
			><%=ucase(rsCLI("DESCRIPCION"))%></option>

			<%rsCLI.movenext
			loop
		end if
		rsCLI.close
		set rsCLI=nothing
		%>
        </select>
        </td>
		  	<td>
				<SELECT NAME="cmb_usuario" id="cmb_usuario" onchange="Refrescar();">
				<option value="0">TODOS</option>
				<%
				ssql="SELECT DISTINCT ID_USUARIO, LOGIN FROM USUARIO U, CONVENIO_ENC C WHERE U.ID_USUARIO = C.USR_INGRESO"
				set rsUsu=Conn.execute(ssql)
				if not rsUsu.eof then
					do until rsUsu.eof
					%>
					<option value="<%=rsUsu("id_usuario")%>"
					<%if Trim(usuario)=Trim(rsUsu("id_usuario")) then
						response.Write("Selected")
					end if%>
					><%=ucase(rsUsu("login"))%></option>

					<%rsUsu.movenext
					loop
				end if
				rsUsu.close
				set rsUsu=nothing
				%>
				</SELECT>
			</td>
			<td><input name="inicio" readonly="true" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
			</td>
			<td><input name="termino" readonly="true" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
			<Td>
			<!--<input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onClick="envia();">-->
			</td>
			
	      </tr>
		 <thead>
	      <tr height="20">
			<td colspan="2">RUT</td>
			<td colspan="2">CODIGO <%=UCASE(session("NOMBRE_CONV_PAGARE"))%></td>
			<Td></td>
	      </tr>
	  	</thead>
		  <tr height="20" class="Estilo8">
	         <td colspan="2"><INPUT TYPE="text" NAME="TX_RUT" value="<%=strrut%>"></td>
			<td colspan="2"><INPUT TYPE="text" NAME="TX_PAGO" value="<%=codpago%>"></td>
			<td><input type="Button" class="fondo_boton_100" name="Submit" value="Buscar" onClick="envia();"></td>
		</tr>
    </table>
    <br>
	<table width="100%" border="0" id="tbl_Procesa" bordercolor="#000000" class="intercalado" style="width:100%;">
		<thead>
		<tr bordercolor="#999999"  class="Estilo13">
			<td>COD. CONV</td>
			<td>FOLIO</td>
			<td>ESTADO</td>
			<td>USU.ING.</td>
			<td>FECHA CONV</td>
			<td>CLIENTE</td>
			<td>RUT DEUDOR</td>
			<td>MONTO <%=UCASE(session("NOMBRE_CONV_PAGARE"))%></td>
			<td>MOROSIDAD</td>
			<td>&nbsp</td>
			<td>&nbsp</td>
			<td>&nbsp</td>
			<td>&nbsp</td>
            
		</tr>
		</thead>
		<tbody>
	<%

        SQL = "SELECT FOLIO,NOM_ESTADO_FOLIO,ID_CONVENIO,CONVERT(VARCHAR(10), FECHA_INGRESO, 103) AS FECHA_INGRESO,CLIENTE.DESCRIPCION, RUT_DEUDOR, TOTAL_CONVENIO, ISNULL(USR_INGRESO,0) AS USR_INGRESO "
        SQL =SQL & ",(select count(ID_PAGO) from CAJA_WEB_EMP where CAJA_WEB_EMP .ID_CONVENIO = CONVENIO_ENC.ID_CONVENIO) CantPagos "
        SQL =SQL & " FROM CONVENIO_ENC,CLIENTE,ESTADO_FOLIO WHERE CLIENTE.COD_CLIENTE = CONVENIO_ENC.COD_CLIENTE AND ESTADO_FOLIO.COD_ESTADO_FOLIO = CONVENIO_ENC.COD_ESTADO_FOLIO"
		SQL = SQL & " AND CONVENIO_ENC.COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ")"


	if resp="si" then
		
        IF CLIENTE <> "0" THEN
			SQL = SQL & " AND CONVENIO_ENC.COD_CLIENTE = '" & CLIENTE & "'"
		END IF
		IF usuario <> "0" THEN
			SQL = SQL & " AND  USR_INGRESO=" & usuario & " "
		END IF
		SQL = SQL & " AND FECHA_INGRESO between '" & inicio & " 00:00:00' and '" & termino & " 23:59:59' "
		
		If Trim(strrut)  <> "" Then
			SQL = SQL & " AND RUT_DEUDOR = '" & strrut & "'"
		End If
		If Trim(codpago)  <> "" Then
			SQL = SQL & " AND ID_CONVENIO = " & codpago
		End If
	
	end if
	
	if resp="123" then
		
		If Trim(strrut)  <> "" Then
			SQL = SQL & " AND RUT_DEUDOR = '" & strrut & "'"
		End If
		If Trim(codpago)  <> "" Then
			SQL = SQL & " AND ID_CONVENIO = " & codpago
		End If
	end if

	
	If Trim(strTipo) = "EP" Then
		SQL = SQL & " AND CONVENIO_ENC.COD_ESTADO_FOLIO IN ( 1,2 )"
	End If
		SQL = SQL & " order by ID_CONVENIO desc"
		
        'response.write(SQL)
		'response.end
	if sql <> "" then
	
		
		
		set rsDet=Conn.execute(SQL)

		if not rsDet.eof then
			do while not rsDet.eof
				ssql="SELECT * FROM USUARIO WHERE ID_USUARIO = " & rsDet("USR_INGRESO")
				set rsUsuIng=Conn.execute(ssql)
				if not rsUsuIng.eof then
					USR_INGRESO= rsUsuIng("login")
				end if
				intTotalConvenio = ValNulo(rsDet("TOTAL_CONVENIO"),"N")


                sssql = "SELECT id_convenio, sum(total_cuota) as morosidad"
				sssql = sssql & " from convenio_det where id_convenio = " & rsDet("ID_CONVENIO")
				sssql = sssql & " and pagada is null and fecha_pago < getdate()"
				sssql = sssql & "group by id_convenio order by ID_CONVENIO desc"

				set rsMorosidad=Conn.execute(sssql)


                'response.write(sssql)

				if not rsMorosidad.eof then
					usrMorosidad= rsMorosidad("morosidad")
					else
					usrMorosidad= "0"
				end if


			%>
			<tr bgcolor="#<%=session("COLTABBG2")%>" class="Estilo8">
				<td><%=rsDet("ID_CONVENIO")%></td>
				<td><%=rsDet("FOLIO")%></td>
				<td><%=rsDet("NOM_ESTADO_FOLIO")%></td>
				<td><%=USR_INGRESO%></td>
				<td><%=rsDet("FECHA_INGRESO")%></td>
				<td><%=rsDet("DESCRIPCION")%></td>
				<td><A HREF="principal.asp?TX_RUT=<%=rsDet("RUT_DEUDOR")%>"><acronym title="Llevar a pantalla de selección"><%=rsDet("RUT_DEUDOR")%></acronym></A></td>

				<TD ALIGN="RIGHT"><%=FN(intTotalConvenio,0)%></td>

				<TD ALIGN="RIGHT"><%=usrMorosidad%></td>
                <td><A HREF="caja_web.asp?CB_TIPOPAGO=CO&id_convenio=<%=rsDet("ID_CONVENIO")%>&rut=<%=rsDet("RUT_DEUDOR")%>"><acronym title="PAGO DE CUOTAS DEL CONVENIO">Pagar Cuota</acronym></A></td>
				<td><a HREF="#" onClick="javascript:ventanaConvenio('cuponera_convenios.asp?strImprime=S&intNroConvenio=<%=rsDet("ID_CONVENIO")%>')"><acronym title="IMPRIMIR CUPONERA DEL CONVENIO">Cuponera</acronym></A></td>
				<td><a HREF="#" onClick="javascript:ventanaConvenio('visualizar_convenio.asp?strRut=<%=rsDet("RUT_DEUDOR")%>&intIdConvenio=<%=rsDet("ID_CONVENIO")%>')"><acronym title="IMPRIMIR <%=UCASE(session("NOMBRE_CONV_PAGARE"))%> ORIGINAL"><%=session("NOMBRE_CONV_PAGARE")%></acronym></A></td>
				<td><A HREF="#" onClick="Reversar(<%=rsDet("ID_CONVENIO")%>,<%=rsDet("CantPagos")%>)";>Reversar</A></td>
				</tr>
			<%
			rsDet.movenext
			loop
		end if
		%>

	<%end if%>
	</tbody>
	</table>
	</td>
   </tr>
  </table>

</form>
</body>
</html>

<script language="JavaScript " type="text/JavaScript">
function Reversar(cod_convenio,CantPago)
{

   
   if (CantPago > 0)
   {
     alert("No es posible reversar un convenio que posee una cuota pagada. Si desea  reversarlo, debe reversar previamente las cuotas pagadas por caja asociadas a éste");
     return;
   }
   
	with( document.datos )
	{
		//alert("Opción deshabilitada");

		if (confirm("¿ Está seguro de reversar el <%=session("NOMBRE_CONV_PAGARE")%> ? El <%=session("NOMBRE_CONV_PAGARE")%> se eliminará completamente y la deuda será reversada, volviendo a su estado original antes del pago."))
			{
				action = "reversar_convenio.asp?cod_convenio=" + cod_convenio;
				submit();
			}
		else
			alert("Reverso del <%=session("NOMBRE_CONV_PAGARE")%> cancelado");
	}
}

function ventanaConvenio (URL){
window.open(URL,"INFORMACION","width=1000, height=500, scrollbars=yes, menubar=no, location=no, resizable=yes")
}
</script>
