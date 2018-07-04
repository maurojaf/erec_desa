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

<script language="JavaScript" src="../javascripts/cal2.js"></script>
<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
<script language="JavaScript" src="../javascripts/validaciones.js"></script>
<link href="../css/style.css" rel="stylesheet" type="text/css">
<link href="../css/style_generales_sistema.css" rel="stylesheet">
<%

Response.CodePage=65001
Response.charset ="utf-8"

	AbrirSCG()
	intCodUsuario=session("session_idusuario")
	strCuadrar = Trim(request("strCuadrar"))

	'intCodUsuario = 110
	strsql="select * from usuario where id_usuario = " & intCodUsuario & ""
	set rsUsu=Conn.execute(strsql)
	if not rsUsu.eof then
		perfil=rsUsu("perfil_caja")

		if perfil = "caja_modif" or perfil = "caja_listado" then
			IF request("cmb_sucursal") <> "" THEN
				sucursal = request("cmb_sucursal")
			ELSE
				sucursal = rsUsu("cod_suc")
			END IF
		else
			sucursal = "" 'rsUsu("cod_suc")
		end if
	end if
	'response.write(perfil)
	codpago = request("TX_pago")
	rut= session("session_RUT_DEUDOR")

	usuario=session("session_idusuario")

	termino = request("termino")
	inicio = request("inicio")
	resp = request("resp")
	GRABA = request("GRABA")

	intIdCaja = request("cmb_caja")

	strCOD_CLIENTE = session("ses_codcli")

	if Trim(inicio) = "" Then
		inicio = TraeFechaMesActual(Conn,0)
		'inicio = "01" & Mid(inicio,3,10)
	End If
	if Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If

	fecha = inicio

	''Response.write "rut = " & rut
%>
	<title>Empresa</title>

	<style type="text/css">
	<!--
	.Estilo13 {color: #FFFFFF}
	.Estilo13n {color: #000000}
	.Estilo27 {color: #FFFFFF}
	-->
	</style>

	<script language="JavaScript " type="text/JavaScript">

	function muestra_dia(){
	//alert(getCurrentDate())
	//alert("hola")
		var diferencia=DiferenciaFechas(datos.termino.value)
		//alert(diferencia)
		if(datos.termino.value!=''){
			if ((diferencia<=0)) {
				//alert('Ok')
				return true
			}else{
				alert('La fecha de apertura no puede ser posterior a la fecha actual')
				datos.termino.value = getCurrentDate();
				datos.termino.focus();
				return false;
			}
		}
	}


	function DiferenciaFechas (CadenaFecha1) {
	   var fecha_hoy = getCurrentDate() //hoy


	   //Obtiene dia, mes y año
	   var fecha1 = new fecha( CadenaFecha1 )
	   var fecha2 = new fecha(fecha_hoy)

	   //Obtiene objetos Date
	   var miFecha1 = new Date( fecha1.anio, fecha1.mes, fecha1.dia )
	   var miFecha2 = new Date( fecha2.anio, fecha2.mes, fecha2.dia )

	   //Resta fechas y redondea
	   var diferencia = miFecha1.getTime() - miFecha2.getTime()
	   var dias = Math.floor(diferencia / (1000 * 60 * 60 * 24))
	   var segundos = Math.floor(diferencia / 1000)
	   //alert ('La diferencia es de ' + dias + ' dias,\no ' + segundos + ' segundos.')

	   return dias //false
	}

	function fecha( cadena ) {

	   //Separador para la introduccion de las fechas
	   var separador = "/"

	   //Separa por dia, mes y año
	   if ( cadena.indexOf( separador ) != -1 ) {
	        var POSI_1 = 0
	        var POSI_2 = cadena.indexOf( separador, POSI_1 + 1 )
	        var POSI_3 = cadena.indexOf( separador, POSI_2 + 1 )
	        this.dia = cadena.substring( POSI_1, POSI_2 )
	        this.mes = cadena.substring( POSI_2 + 1, POSI_3 )
	        this.anio = cadena.substring( POSI_3 + 1, cadena.length )
	   } else {
	        this.dia = 0
	        this.mes = 0
	        this.anio = 0
	   }
	}

	function Refrescar()
	{
		GRABA='no'
		resp='no'

		datos.action = "apertura_caja.asp?GRABA="+ GRABA +"&resp="+ resp +"";
		datos.submit();
	}



	function Ingresa()
	{
		GRABA='si'
		resp='si'
		strCuadrar='no'

		if (datos.cmb_caja.value == '0')
		{
			alert("Debe ingresar la caja donde desea abrir");
			return;
		}
		if (confirm("¿Está seguro de realizar la apertura de caja?"))
			{
			datos.action = "apertura_caja.asp?strCuadrar="+ strCuadrar +"&GRABA="+ GRABA +"&resp="+ resp +"";
			datos.submit();
			}
		else
			alert("Proceso de apertura de caja no se ha realizado");

	}


	function Cuadrar()
	{
		//datos.TX_RUT.value='';
		//datos.TX_pago.value='';
		GRABA='no'
		resp='si'
		strCuadrar='si'
		datos.action = "apertura_caja.asp?strCuadrar="+ strCuadrar +"&GRABA="+ GRABA +"&resp="+ resp +"";
		datos.submit();
	}


	function envia()
	{
		//datos.TX_RUT.value='';
		//datos.TX_pago.value='';
		GRABA='no'
		resp='si'
		strCuadrar='no'
		datos.action = "apertura_caja.asp?strCuadrar="+ strCuadrar +"&GRABA="+ GRABA +"&resp="+ resp +"";
		datos.submit();
	}

	function envia_excel(URL){

	window.open(URL,"INFORMACION","width=200, height=200, scrollbars=yes, menubar=yes, location=yes, resizable=yes")
	}
	</script>
	
</head>
<body>

<form name="datos" method="post">
<div class="titulo_informe">APERTURA DE CAJA</div>

<table width="90%" height="500" border="0" align="center">
    <td valign="top">

	<%

	strSql="SELECT ISNULL(FECHA_APERTURA,'') AS FECHA_APERTURA, ISNULL(FECHA_CIERRE,'') AS FECHA_CIERRE FROM CAJA_WEB_EMP_CIERRE WHERE COD_USUARIO = " & usuario
	strSql= strSql & " AND CAST(CONVERT(VARCHAR(10),GETDATE(),103) AS DATETIME) <= FECHA AND FECHA < (GETDATE()-0.083) AND FECHA_CIERRE IS NULL AND (CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "' OR CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = 0)"
	'Response.write "<br>strSql=" & strSql

	set rsNoCierreCaja=Conn.execute(strSql)

	If Not rsNoCierreCaja.Eof Then

		strSql = "UPDATE CAJA_WEB_EMP_CIERRE SET FECHA_CIERRE = '" & Fecha & "', FECHA_HORA_CIERRE = GETDATE()"
		strSql = strSql & " WHERE COD_USUARIO = " & Usuario
		strSql = strSql & " AND FECHA_CIERRE IS NULL AND (CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "' OR CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = 0)"
		'Response.write (strsql)

		set rsGraba=Conn.execute(strsql)

		%>
		<SCRIPT>
			alert('La Caja se cerro automáticamente')
			location.href='apertura_caja.asp?rut=" + rut + "&tipo=1';
		</SCRIPT>
		<%
	Else

		''---REVISA QUE EL USUARIO TENGA CAJA ASIGNADA AL CLIENTE O A LLACRUZ---''

		strSql=" SELECT TOP 1 ID_CAJA"
		strSql=strSql & " FROM CAJAS_RECAUDACION_USUARIO WHERE ID_USUARIO = '" & usuario & "' AND COD_CLIENTE = '" & strCOD_CLIENTE & "'"

		set rsCajaAsig=Conn.execute(strSql)

		if not rsCajaAsig.eof then
			do until rsCajaAsig.eof

		''---REVISA QUE EL USUARIO TENGA CAJA ABIERTAS EN ESE CLIENTE O LLACRUZ---''

				strSql="SELECT TOP 1 FECHA AS FECHA, USUARIO.LOGIN AS LOGIN, CR.NOM_CAJA AS NOM_CAJA, CR.COD_CAJA AS COD_CAJA, ISNULL(FECHA_CIERRE,'') AS FECHA_CIERRE"
				strSql= strSql & " 	FROM CAJA_WEB_EMP_CIERRE INNER JOIN USUARIO ON USUARIO.ID_USUARIO = CAJA_WEB_EMP_CIERRE.COD_USUARIO"
				strSql= strSql & " 							 INNER JOIN CAJAS_RECAUDACION AS CR ON CR.ID_CAJA = CAJA_WEB_EMP_CIERRE.SUCURSAL AND CR.COD_CLIENTE = '" & strCOD_CLIENTE & "'"
				strSql= strSql & "	WHERE COD_USUARIO = " & usuario & " AND FECHA_CIERRE IS NULL AND CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "'"
				strSql= strSql & "  ORDER BY FECHA DESC"

				''Response.write "<br>strSql=" & strSql

				set rsInforme=Conn.execute(strSql)

				if not rsInforme.eof then
					do until rsInforme.eof

					%>

					<table width="100%" border="0" bordercolor="#999999" class="estilo_columnas">
						<thead>
						  <tr height="20" >
							<td colspan="2" >Informe Apertura de caja</td>
						  </tr>
						</thead>

						  <tr height="20" bordercolor="#999999" class="Estilo8">
							<td class="hdr_i" width="70%">Estado Aperuta Caja</td>
							<td class="td_t" width="55%">ABIERTA</td>
						  </tr>

						  <tr height="20" bordercolor="#999999" class="Estilo8">
							<td class="hdr_i">Día Apertura</td>
							<td class="td_t"><%=rsInforme("FECHA")%></td>
						  </tr>

						  <tr height="20" bordercolor="#999999" class="Estilo8">
							<td class="hdr_i">Caja Apertura</td>
							<td class="td_t"><%=rsInforme("NOM_CAJA")%></td>
						  </tr>

						  <tr height="20" bordercolor="#999999" class="Estilo8">
							<td class="hdr_i">Código Caja Apertura</td>
							<td class="td_t"><%=rsInforme("COD_CAJA")%></td>
						  </tr>

						  <tr height="20" bordercolor="#999999" class="Estilo8">
							<td class="hdr_i">Usuario Apertura</td>
							<td class="td_t"><%=rsInforme("LOGIN")%></td>
						  </tr>

						  <tr height="20" class="estilo_columna_individual">
							<td colspan="2">&nbsp;</td>
						  </tr>

					</table>

			<%		rsInforme.movenext
					loop
				Else

		''---COMO LA CAJA ESTA CERRADA OBTIENE INFORME DE CIERRE---''

							strSql="SELECT TOP 1 FECHA_HORA_CIERRE AS FECHA_CIERRE,FECHA AS FECHA_APERTURA, USUARIO.LOGIN AS LOGIN, CR.NOM_CAJA AS NOM_CAJA, CR.COD_CAJA AS COD_CAJA"
							strSql= strSql & " 	FROM CAJA_WEB_EMP_CIERRE INNER JOIN USUARIO ON USUARIO.ID_USUARIO = CAJA_WEB_EMP_CIERRE.COD_USUARIO"
							strSql= strSql & " 							 INNER JOIN CAJAS_RECAUDACION AS CR ON CR.ID_CAJA = CAJA_WEB_EMP_CIERRE.SUCURSAL AND CR.COD_CLIENTE = '" & strCOD_CLIENTE & "'"
							strSql= strSql & "	WHERE COD_USUARIO = " & usuario & " AND FECHA_CIERRE IS NOT NULL AND CAJA_WEB_EMP_CIERRE.CLIENTE_APERTURA = '" & strCOD_CLIENTE & "'"
							strSql= strSql & "  ORDER BY FECHA DESC"
							''Response.write "<br>strSql=" & strSql

							set rsInforme2=Conn.execute(strSql)

							if not rsInforme2.eof then
								do until rsInforme2.eof

								%>

									<table width="100%" border="0" class="estilo_columnas">
									<thead>
									  <tr height="20" >
										<td colspan="2">Informe Apertura de caja</td>
									  </tr>
									 </thead> 

									  <tr height="20" >
										<td class="hdr_i" width="70%">Estado Aperuta Caja</td>
										<td class="td_t" width="55%">CERRADA</td>
									  </tr>

									  <tr height="20" >
										<td class="hdr_i">Día y hora Apertura</td>
										<td class="td_t"><%=rsInforme2("FECHA_APERTURA")%></td>
									  </tr>

									  <tr height="20" >
										<td class="hdr_i">Día y hora Cierre</td>
										<td class="td_t"><%=rsInforme2("FECHA_CIERRE")%></td>
									  </tr>

									  <tr height="20" >
										<td class="hdr_i">Ultima Caja Abierta</td>
										<td class="td_t"><%=rsInforme2("NOM_CAJA")%></td>
									  </tr>

									  <tr height="20" >
										<td class="hdr_i">Ultimo Código Caja Abierta</td>
										<td class="td_t"><%=rsInforme2("COD_CAJA")%></td>
									  </tr>

									  <tr height="20" >
										<td class="hdr_i">Usuario Cierre</td>
										<td class="td_t"><%=rsInforme2("LOGIN")%></td>
									  </tr>

									</table>

									<table width="100%" class="estilo_columnas">
									<thead>
									  <tr height="20">
										<td colspan="2">Datos Apertura de caja</td>
									  </tr>
									 </thead>

									  <tr height="20" bordercolor="#999999" class="Estilo8"><td class="hdr_i">Cajas</td>
										<td class="td_t">
											<SELECT NAME="cmb_caja" id="cmb_caja">
												<option value="0">SELECCIONAR</option>
												<%
												strSql=" SELECT CR.ID_CAJA, COD_CAJA, (CAST(COD_CAJA AS VARCHAR(10))+'-'+NOM_CAJA) AS NOM_CAJA"
												strSql=strSql & " FROM CAJAS_RECAUDACION AS CR left JOIN CAJAS_RECAUDACION_USUARIO CRU ON CR.ID_CAJA = CRU.ID_CAJA AND CRU.ID_USUARIO = '" & usuario & "' AND (CR.COD_CLIENTE = '" & strCOD_CLIENTE & "' OR  CR.COD_CLIENTE = 0) AND CR.COD_CLIENTE = CRU.COD_CLIENTE"

												set rsCaja=Conn.execute(strSql)
												if not rsCaja.eof then
													do until rsCaja.eof
													%>
													<option value="<%=rsCaja("ID_CAJA")%>"
													<%if Trim(intIdCaja)=Trim(rsCaja("ID_CAJA")) then
														response.Write("Selected")
													end if%>
													><%=ucase(rsCaja("NOM_CAJA"))%></option>

													<%rsCaja.movenext
													loop
												end if
												rsCaja.close
												set rsCaja=nothing
												%>
											</SELECT>
											
										</td>
									  </tr>

									  <tr bordercolor="#999999" class="Estilo8">
										<td colspan=2 align="right" >
											<input type="Button" class="fondo_boton_100" name="Submit" value="Abrir" onClick="Ingresa();">
										</td>
									  </tr>

									</table>

								<%
								rsInforme2.movenext
								loop
							Else

		''---COMO NO ENCUENTRA CAJAS HISTORICAMENTE ABIERTAS Y ESTA ASIGNADO ASUME QUE EL USUARIO NO HA ABIERTO CAJAS NUNCA---''

	%>
									<table width="100%" class="estilo_columnas">
									<thead>
										  <tr height="20" bordercolor="#999999"  bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
											<td colspan=2>Informe Apertura de caja</td>
										  </tr>

										  <tr height="20" bordercolor="#999999" class="Estilo8">
											<td class="hdr_i" width="70%">Bienvenido a su primera Apertura de Caja, suerte en su trabajo!</td>
										  </tr>
									</thead>
									</table>

									<table width="100%" class="estilo_columas">
									<thead>
									  <tr height="20">
										<td colspan="2">Datos Apertura de caja</td>
									  </tr>
									</thead>
									  <tr height="20" bordercolor="#999999" class="Estilo8"><td class="hdr_i">Cajas</td>
										<td class="td_t">
											<SELECT NAME="cmb_caja" id="cmb_caja">
												<option value="0">SELECCIONAR</option>
												<%
												strSql=" SELECT CR.ID_CAJA, COD_CAJA, (CAST(COD_CAJA AS VARCHAR(10))+'-'+NOM_CAJA) AS NOM_CAJA"
												strSql=strSql & " FROM CAJAS_RECAUDACION AS CR INNER JOIN CAJAS_RECAUDACION_USUARIO CRU ON CR.ID_CAJA = CRU.ID_CAJA AND CRU.ID_USUARIO = '" & usuario & "' AND (CR.COD_CLIENTE = '" & strCOD_CLIENTE & "' OR  CR.COD_CLIENTE = 0) AND CR.COD_CLIENTE = CRU.COD_CLIENTE"
												'Response.write "strSql=" & strSql
												set rsCaja=Conn.execute(strSql)
												if not rsCaja.eof then
													do until rsCaja.eof
													%>
													<option value="<%=rsCaja("ID_CAJA")%>"
													<%if Trim(intIdCaja)=Trim(rsCaja("ID_CAJA")) then
														response.Write("Selected")
													end if%>
													><%=ucase(rsCaja("NOM_CAJA"))%></option>

													<%rsCaja.movenext
													loop
												end if
												rsCaja.close
												set rsCaja=nothing
												%>
											</SELECT>
										</td>
									  </tr>

									  <tr bordercolor="#999999" class="Estilo8">
										<td colspan=2 align="right" >
											<input type="Button" class="fondo_boton_100" name="Submit" value="Abrir" onClick="Ingresa();">
										</td>
									  </tr>

									</table>

			<%
							end if
							rsInforme2.close
							set rsInforme2=nothing

			end if
			rsInforme.close
			set rsInforme=nothing

			rsCajaAsig.movenext
		loop

		Else

			Response.Write ("<script language = ""Javascript"">" & vbCrlf)

			Response.Write (vbTab & "alert('No se puede abrir la caja, Debe solicitar al administrador que le asigne una caja para este cliente');" & vbCrlf)
			Response.Write (vbTab & "location.href='principal.asp?rut=" + rut + "&tipo=1';" & vbCrlf)

			Response.Write ("</script>")

		End if
		rsCajaAsig.close
		set rsCajaAsig=nothing
	End if

	  %>

	<%If GRABA = "si" Then
		sw=0

			strSql="SELECT COD_CLIENTE AS COD_CLIENTE"
			strSql= strSql & " 	FROM CAJAS_RECAUDACION"
			strSql= strSql & "	WHERE ID_CAJA = " & intIdCaja & " AND COD_CLIENTE = '" & strCOD_CLIENTE & "'"
			'Response.write "<br>strSql=" & strSql

			set rsApertura=Conn.execute(strSql)

				strCOD_CLIENTEApertura = rsApertura("COD_CLIENTE")

			'Response.write "<br>strCOD_CLIENTEApertura=" & strCOD_CLIENTEApertura

			If not rsApertura.EOF Then
				strsql = "INSERT INTO CAJA_WEB_EMP_CIERRE (COD_USUARIO, FECHA_APERTURA, FECHA, ASIGNACION, BOLETA_INICIAL, SUCURSAL, CLIENTE_APERTURA) VALUES (" & intCodUsuario & ",'" & fecha & "',GETDATE(),NULL,NULL," & intIdCaja & ", '" & strCOD_CLIENTEApertura & "' )"
				'Response.write (strsql)
				'Response.End
				set rsGRABA=Conn.execute(strsql)
				Response.Write ("<script language = ""Javascript"">" & vbCrlf)

				Response.Write (vbTab & "alert('Apertura de caja realizada correctamente');" & vbCrlf)

				if rut <> "" then
					Response.Write (vbTab & "location.href='caja_web.asp?rut=" + rut + "&tipo=1';" & vbCrlf)
				Else
					Response.Write (vbTab & "location.href='apertura_caja.asp?rut=" + rut + "&tipo=1';" & vbCrlf)
				End If

				Response.Write ("</script>")

			End if

	  End if
	%>
	</td>
   </tr>
  </table>

</form>
</body>
</html>

<script language="JavaScript" type="text/JavaScript">
function solonumero(valor){
     //Compruebo si es un valor numérico

 if (valor.value.length >0){
      if (isNaN(valor.value)) {
            //entonces (no es numero) devuelvo el valor cadena vacia
            ////valor.value="0";
			//alert(valor.value)
			//valor.focus();
			return ""
      }else{
            //En caso contrario (Si era un número) devuelvo el valor
			valor.value
			return valor.value
      }
	  }
}

</script>

