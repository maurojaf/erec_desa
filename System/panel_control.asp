<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/lib2.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/Minimo.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc" -->

<script language="JavaScript">
	function ventanaSecundaria (URL){
		window.open(URL,"DETALLE","width=800, height=300, scrollbars=YES, menubar=no, location=no, resizable=yes")
	}
</script>
<%

Dim PaginaActual ' en qu? pagina estamos
Dim PaginasTotales ' cu?ntas p?ginas tenemos
Dim TamPagina ' cuantos registros por pagina
Dim CuantosRegistros ' para imprimir solo el n? de registro por pagina que

strEjeAsig = Request("CB_EJECUTIVO")


dtmFec1 = Date()
dtmFec2 = Trim(dateadd("d",1,dtmFec1))
dtmFec3 = Trim(dateadd("d",1,dtmFec2))
dtmFec4 = Trim(dateadd("d",1,dtmFec3))
dtmFec5 = Trim(dateadd("d",1,dtmFec4))
dtmFec6 = Trim(dateadd("d",1,dtmFec5))
dtmFec7 = Trim(dateadd("d",1,dtmFec6))

''Response.write "dtmFec7 = "& dtmFec7


If trim(strNivelAtraso) = "" Then strNivelAtraso = "ROJOAMARILLO"
''Response.write "strNivelAtraso=" & strNivelAtraso


'MODIFICAR AQUI PARA CAMBIAR EL N? DE REGISTRO POR PAGINA
TamPagina=100

'Leemos qu? p?gina mostrar. La primera vez ser? la inicial
if Request.Querystring("pagina")="" then
	PaginaActual=1
else
	PaginaActual=CInt(Request.Querystring("pagina"))
end if


%>
<title>MODULO DE AGENDAMIENTOS</title>


<%strTitulo="MI CARTERA"%>

<script language='javascript' src="../javascripts/popcalendar.js"></script>
</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="datos" method="post">

<table width="100%" border="1" bordercolor="#FFFFFF">
	<tr>
		<TD height="20" ALIGN=LEFT class="pasos2_i">
			<B>PANEL DE CONTROL</B>
		</TD>
	</tr>
</table>


	<table width="850" align="CENTER">
		<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="60%" align="center"><%=strMensaje%></td>
		</tr>
	</table>

				<%
					strTitulo1="Lunes"
					strTitulo2="Martes"
					strTitulo3="Miercoles"
					strTitulo4="Jueves"
					strTitulo5="Viernes"
					strTitulo6="Sabado"
					strTitulo7="Domingo"
				%>


					  <table width="850" align="CENTER">
						<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
							<td>EJECUTIVO</td>
							<td colspan=2 align="CENTER">Arrastre</td>
							<td colspan=2 align="CENTER"><%=dtmFec1%></td>
							<td colspan=2 align="CENTER"><%=dtmFec2%></td>
							<td colspan=2 align="CENTER"><%=dtmFec3%></td>
							<td colspan=2 align="CENTER"><%=dtmFec4%></td>
							<td colspan=2 align="CENTER"><%=dtmFec5%></td>
							<td colspan=2 align="CENTER"><%=dtmFec6%></td>
							<td colspan=2 align="CENTER"><%=dtmFec7%></td>
						</tr>
						<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
							<td>&nbsp;</td>
							<td align="CENTER">Casos</td>
							<td align="CENTER">Doc</td>
							<td align="CENTER">Casos</td>
							<td align="CENTER">Doc</td>
							<td align="CENTER">Casos</td>
							<td align="CENTER">Doc</td>
							<td align="CENTER">Casos</td>
							<td align="CENTER">Doc</td>
							<td align="CENTER">Casos</td>
							<td align="CENTER">Doc</td>
							<td align="CENTER">Casos</td>
							<td align="CENTER">Doc</td>
							<td align="CENTER">Casos</td>
							<td align="CENTER">Doc</td>
							<td align="CENTER">Casos</td>
							<td align="CENTER">Doc</td>
						</tr>


							<%
						   		AbrirScg()

						   		strSql = "SELECT ID_USUARIO, rut_usuario, nombres_usuario, "
						   		strSql = strSql & " isnull(apellido_paterno,'') apellido_paterno , isnull(apellido_materno,'') apellido_materno, fecha_nacimiento, "
						   		strSql = strSql & " correo_electronico, telefono_contacto, perfil, LOGIN, CLAVE, "
						   		strSql = strSql & " PERFIL_ADM, perfil_cob, ACTIVO, perfil_proc, perfil_sup, "
						   		strSql = strSql & " PERFIL_CAJA, perfil_emp, PERFIL_FULL, perfil_back, "
						   		strSql = strSql & " gestionador_preventivo, anexo, observaciones "
						   		strSql = strSql & " FROM USUARIO WHERE PERFIL_COB = 1 AND ACTIVO = 1 "
								If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then
								''If TraeSiNo(session("perfil_adm"))="Si" Then

								Else
									strSql = strSql & "	AND ID_USUARIO = " & session("session_idusuario")
								End If
								set rsUsuario = Conn.execute(strSql)
								Do While Not rsUsuario.Eof
									intCodEjecutivo = rsUsuario("id_usuario")


									strSql = "SELECT COUNT(DISTINCT RUT_DEUDOR) CASOS, COUNT(DISTINCT ID_CUOTA) DOC"
									strSql = strSql & " FROM CUOTA WHERE COD_CLIENTE = '" & session("ses_codcli") & "' AND USUARIO_ASIG = " & intCodEjecutivo
									strSql = strSql & " AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
									strSql = strSql & " AND FECHA_AGEND_ULT_GES < GETDATE()"
									set rsTemp = Conn.execute(strSql)
									If not rsTemp.eof Then
										intArrastreCasos = Trim(rsTemp("CASOS"))
										intArrastreDoc = Trim(rsTemp("DOC"))
									Else
										intArrastreCasos = 0
										intArrastreDoc = 0
									End If


									strNombreEjecutivo = rsUsuario("nombres_usuario") & " " & rsUsuario("apellido_paterno")& " " & rsUsuario("apellido_materno")
										intTotal=0

										strSql = "select count(distinct ID_CUOTA) AS DOC, count(distinct RUT_DEUDOR) AS CASOS, convert(varchar(10),FECHA_AGENDAMIENTO,103) as fecha from gestiones g, gestiones_cuota gc"
										strSql = strSql & " where g.id_gestion = gc.id_gestion and ID_USUARIO = " & intCodEjecutivo & " and COD_CLIENTE = '" & session("ses_codcli") & "'"
										''strSql = strSql & " and convert(varchar(10),FECHA_AGENDAMIENTO,103) = '" & dtmFec1 & "'"
										strSql = strSql & " and FECHA_AGENDAMIENTO >= '" & dtmFec1 & "'"
										strSql = strSql & " and FECHA_AGENDAMIENTO <= '" & dtmFec7 & "'"
										strSql = strSql & " group by convert(varchar(10),FECHA_AGENDAMIENTO,103)"

										strSql = strSql & " SELECT COUNT(DISTINCT RUT_DEUDOR) CASOS, COUNT(DISTINCT ID_CUOTA) DOC"
										strSql = strSql & " FROM CUOTA WHERE USUARIO_ASIG = " & intCodEjecutivo & " AND COD_CLIENTE = '" & session("ses_codcli") & "'"
										strSql = strSql & " AND ESTADO_DEUDA IN (SELECT CODIGO FROM ESTADO_DEUDA WHERE ACTIVO = 1)"
										strSql = strSql & " AND FECHA_AGEND_ULT_GES >= '" & dtmFec1 & " 00:00:00'"
										strSql = strSql & " AND FECHA_AGEND_ULT_GES <= '" & dtmFec7 & " 23:59:59'"
										''and FECHA_AGEND_ULT_GES < getdate()

										''Response.write "strSql = " & strSql

										set rsCuenta = Conn.execute(strSql)

										intCasos1 = 0
										intDoc1 = 0
										intCasos2 = 0
										intDoc2 = 0
										intCasos3 = 0
										intDoc3 = 0
										intCasos4 = 0
										intDoc4 = 0
										intCasos5 = 0
										intDoc5 = 0
										intCasos6 = 0
										intDoc6 = 0
										intCasos7 = 0
										intDoc7 = 0


										Do While Not rsCuenta.Eof
											''Response.write "<br>dtmFec1 = " & dtmFec1

												'Response.write "<br>Comp == " & (Trim(dtmFec1) = Trim(rsCuenta("FECHA")))
												'Response.write "<br>dtmFec1 == " & dtmFec1
												''Response.write "<br>FECHA == " & Trim(rsCuenta("FECHA"))

												If (Trim(dtmFec1) = Trim(rsCuenta("FECHA"))) Then
													intCasos1 = Trim(rsCuenta("CASOS"))
													intDoc1 = Trim(rsCuenta("DOC"))
												End If
												If (Trim(dtmFec2) = Trim(rsCuenta("FECHA"))) Then
													intCasos2 = Trim(rsCuenta("CASOS"))
													intDoc2 = Trim(rsCuenta("DOC"))
												End If
												If (Trim(dtmFec3) = Trim(rsCuenta("FECHA"))) Then
													intCasos3 = Trim(rsCuenta("CASOS"))
													intDoc3 = Trim(rsCuenta("DOC"))
												End If
												If (Trim(dtmFec4) = Trim(rsCuenta("FECHA"))) Then
													intCasos4 = Trim(rsCuenta("CASOS"))
													intDoc4 = Trim(rsCuenta("DOC"))
												End If
												If (Trim(dtmFec5) = Trim(rsCuenta("FECHA"))) Then
													intCasos5 = Trim(rsCuenta("CASOS"))
													intDoc5 = Trim(rsCuenta("DOC"))
												End If
												If (Trim(dtmFec6) = Trim(rsCuenta("FECHA"))) Then
													intCasos6 = Trim(rsCuenta("CASOS"))
													intDoc6 = Trim(rsCuenta("DOC"))
												End If
												If (Trim(dtmFec7) = Trim(rsCuenta("FECHA"))) Then
													intCasos7 = Trim(rsCuenta("CASOS"))
													intDoc7 = Trim(rsCuenta("DOC"))
												End If
											rsCuenta.movenext
										Loop

									%>

									<tr>
										<td><%=strNombreEjecutivo%></td>
										<td align="right"><%=intArrastreCasos%></td>
										<td align="right"><%=intArrastreDoc%></td>
										<td align="right"><%=intCasos1%></td>
										<td align="right"><%=intDoc1%></td>
										<td align="right"><%=intCasos2%></td>
										<td align="right"><%=intDoc2%></td>
										<td align="right"><%=intCasos3%></td>
										<td align="right"><%=intDoc3%></td>
										<td align="right"><%=intCasos4%></td>
										<td align="right"><%=intDoc4%></td>
										<td align="right"><%=intCasos5%></td>
										<td align="right"><%=intDoc5%></td>
										<td align="right"><%=intCasos6%></td>
										<td align="right"><%=intDoc6%></td>
										<td align="right"><%=intCasos7%></td>
										<td align="right"><%=intDoc7%></td>
									</tr>


								<%

								intTotCasos1 = intTotCasos1 + intCasos1
								intTotCasos2 = intTotCasos2 + intCasos2
								intTotCasos3 = intTotCasos3 + intCasos3
								intTotCasos4 = intTotCasos4 + intCasos4
								intTotCasos5 = intTotCasos5 + intCasos5
								intTotCasos6 = intTotCasos6 + intCasos6
								intTotCasos7 = intTotCasos7 + intCasos7

								intTotDoc1 = intTotDoc1 + intDoc1
								intTotDoc2 = intTotDoc2 + intDoc2
								intTotDoc3 = intTotDoc3 + intDoc3
								intTotDoc4 = intTotDoc4 + intDoc4
								intTotDoc5 = intTotDoc5 + intDoc5
								intTotDoc6 = intTotDoc6 + intDoc6
								intTotDoc7 = intTotDoc7 + intDoc7

								intTotArrastreCasos = CInt(intTotArrastreCasos) + CInt(intArrastreCasos)
								intTotArrastreDoc = CInt(intTotArrastreDoc) + CInt(intArrastreDoc)



						   		rsUsuario.MoveNext

						   		Loop
						   		CerrarScg()
							%>
					<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
						<td>Totales</td>
						<td align="right"><%=intTotArrastreCasos%></td>
						<td align="right"><%=intTotArrastreDoc%></td>
						<td align="right"><%=intTotCasos1%></td>
						<td align="right"><%=intTotDoc1%></td>
						<td align="right"><%=intTotCasos2%></td>
						<td align="right"><%=intTotDoc2%></td>
						<td align="right"><%=intTotCasos3%></td>
						<td align="right"><%=intTotDoc3%></td>
						<td align="right"><%=intTotCasos4%></td>
						<td align="right"><%=intTotDoc4%></td>
						<td align="right"><%=intTotCasos5%></td>
						<td align="right"><%=intTotDoc5%></td>
						<td align="right"><%=intTotCasos6%></td>
						<td align="right"><%=intTotDoc6%></td>
						<td align="right"><%=intTotCasos7%></td>
						<td align="right"><%=intTotDoc7%></td>
					</tr>
			</table>


</form>
</body>
</html>
<script language="JavaScript1.2">

function buscar(){
	datos.action='panel_control.asp?strBuscar=S';
	datos.submit();

}

function limpiar(){
	datos.action='panel_control.asp?strLimpiar=S';
	datos.submit();

}

function IrPagina( sintAccion ) {
	if (sintAccion == 'Retroceder') {
    	self.location.href = 'panel_control.asp?pagina=<%=PaginaActual - 1%>&TX_NOMBRES=<%=strNombres%>&CB_REMESA=<%=intCodRemesa%>&CB_CLIENTE=<%=strCOD_CLIENTE%>&CB_EJECUTIVO=<%=strEjeAsig%>&CB_CAMPANA=<%=intCodCampana%>&CB_TIPOCARTERA=<%=strTipoInf%>&TX_INICIO=<%=dtmInicio%>&TX_TERMINO=<%=dtmTermino%>&CB_TIPOGESTION_TEL=<%=intGestionTel%>&CB_NTRASO=<%=strNivelAtraso%>&CB_TIPOGESTION_PRINC=<%=intGestionPrinc%>'
    }
    if (sintAccion == 'Avanzar') {
	    self.location.href = 'panel_control.asp?pagina=<%=PaginaActual + 1%>&TX_NOMBRES=<%=strNombres%>&CB_REMESA=<%=intCodRemesa%>&CB_CLIENTE=<%=strCOD_CLIENTE%>&CB_EJECUTIVO=<%=strEjeAsig%>&CB_CAMPANA=<%=intCodCampana%>&CB_TIPOCARTERA=<%=strTipoInf%>&TX_INICIO=<%=dtmInicio%>&TX_TERMINO=<%=dtmTermino%>&CB_TIPOGESTION_TEL=<%=intGestionTel%>&CB_NTRASO=<%=strNivelAtraso%>&CB_TIPOGESTION_PRINC=<%=intGestionPrinc%>'
    }

}

</script>