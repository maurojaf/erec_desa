<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
<!--#include file="sesion.asp"-->
    
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/lib.asp"-->
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>
<script language="JavaScript" type="text/JavaScript">

	function AbreArchivo(nombre){
		window.open(nombre,"INFORMACION","width=800, height=400, scrollbars=yes, menubar=no, location=no, resizable=yes");
	}
	function Terminar( sintPaginaTerminar ) {
		self.location.href = sintPaginaTerminar
	}
	function Procesar()

		{
			if (confirm("¿ Está REALMENTE seguro de Actualizar la deuda ?"))
			{
				datos.action='Man_Actualiza_UMA_FACT.asp?strProcesar=SI';
				datos.submit();
			}
		}
</script>


<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
<title>CRM Cobros</title>
<style type="text/css">
<!--body {	background-color: #cccccc;}-->
</style>
</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">

<%

'******************************
'*	INICIO CODIGO PARTICULAR  *
''******************************

strReprocesar = Request("strReprocesar")
strProcesar = Request("strProcesar")

intTotalRutCarga = Request("intTotalRutCarga")
intTotalDoc = Request("intTotalDoc")

'strCodCliente=session("ses_codcli")

if Request("archivo") <> "" then
	strArchivo=Request("archivo")
End if

strCodCliente=Request("CB_CLIENTE")


AbriRsCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

If strProcesar = "SI" Then

	'ACTUALIZA DEUDA DE DOCUMENTOS CON LOS ESTADOS SELECCIONADOS'

	strSql = "EXEC proc_Inf_Actualizacion_deuda_UMA_FACT 'INFORME_DEUDA'"

 	set rsDet=Conn.execute(strSql)

		if not rsDet.eof then
			do until rsDet.eof

				intIdEstadoAct = rsDet("ID_ESTADO_ACT")
				intID_CUOTA = rsDet("ID_CUOTA")
				intCodigo = rsDet("ESTADO_NUEVO")
				intEstadoAct = rsDet("COD_ESTADO_ACT")
				strUltimoEstado = rsDet("ULTIMO_ESTADO")
				intCodSap = rsDet("COD_SAP")

				strObjeto = "CH_" & intIdEstadoAct

				'Response.write "<BR>strObjeto =" & strObjeto
				'Response.write "<BR>strObjetor =" & UCASE(Request(strObjeto))

				If UCASE(Request(strObjeto)) = "ON" Then

					strSql = "		   UPDATE CUOTA"
					strSql = strSql &" SET FECHA_ESTADO = getdate(), ESTADO_DEUDA = "& intCodigo &","
					strSql = strSql &" OBSERVACION = (CASE WHEN "& intCodigo &" = 1 THEN 'VUELTO A ACTIVAR POR " & session("session_idusuario") & "'"
					strSql = strSql &" 				   	   WHEN "& intCodigo &" = 2 THEN 'RETIRADO POR " & session("session_idusuario") & "'"
					strSql = strSql &" 					   WHEN "& intCodigo &" = 3 THEN 'PAGO EN CLIENTE TOTAL POR " & session("session_idusuario") & "'"
					strSql = strSql &" 					   WHEN "& intCodigo &" = 5 THEN 'RETIRADO POR RESOL. POR " & session("session_idusuario") & "'"
					strSql = strSql &" 					   WHEN "& intCodigo &" = 13 THEN 'NO ASIGNABLE POR " & session("session_idusuario") & "'"
					strSql = strSql &" 					   WHEN "& intCodigo &" = 14 THEN 'FIN COBRANZA POR " & session("session_idusuario") & "'"
					strSql = strSql &" 					   END),"
					strSql = strSql &" SALDO = (CASE WHEN "& intCodigo &" = 1 THEN VALOR_CUOTA ELSE 0 END)"
					strSql = strSql &" WHERE ID_CUOTA = "& intID_CUOTA

				set rsUpdate=Conn.execute(strSql)

					strSql = "		   UPDATE CARGA_UMA_FACT"
					strSql = strSql &" SET FECHA_ULT_ESTADO_ACT = getdate(), ULT_ESTADO_ACT = '"& strUltimoEstado &"'"
					strSql = strSql &" WHERE COD_SAP = "& intCodSap &" and ULT_ESTADO_ACT IS NULL"

				set rsUpdate2=Conn.execute(strSql)

					strSql = "		   UPDATE CARGA_UMA_FACT"
					strSql = strSql &" SET FECHA_PENULT_ESTADO_ACT = FECHA_ULT_ESTADO_ACT, PENULT_ESTADO_ACT = ULT_ESTADO_ACT,FECHA_ULT_ESTADO_ACT = getdate(), ULT_ESTADO_ACT = '"& strUltimoEstado &"'"
					strSql = strSql &" WHERE COD_SAP = "& intCodSap &" and ULT_ESTADO_ACT IS NOT NULL and ULT_ESTADO_ACT <> '"& strUltimoEstado &"'"

				set rsUpdate2=Conn.execute(strSql)

				'Response.write "<BR>strSql =" & strSql

				End If

				If intIdEstadoAct = "11" Then

					strSql = "		   UPDATE CARGA_UMA_FACT"
					strSql = strSql &" SET FECHA_ULT_ESTADO_ACT = getdate(), ULT_ESTADO_ACT = '"& strUltimoEstado &"'"
					strSql = strSql &" WHERE COD_SAP = "& intCodSap &" and ULT_ESTADO_ACT IS NULL"

				set rsUpdate2=Conn.execute(strSql)

					strSql = "		   UPDATE CARGA_UMA_FACT"
					strSql = strSql &" SET FECHA_PENULT_ESTADO_ACT = FECHA_ULT_ESTADO_ACT, PENULT_ESTADO_ACT = ULT_ESTADO_ACT,FECHA_ULT_ESTADO_ACT = getdate(), ULT_ESTADO_ACT = '"& strUltimoEstado &"'"
					strSql = strSql &" WHERE COD_SAP = "& intCodSap &" and ULT_ESTADO_ACT IS NOT NULL and ULT_ESTADO_ACT <> '"& strUltimoEstado &"'"

				set rsUpdate2=Conn.execute(strSql)


				End If


				rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing

		strSql1 = "EXEC Proc_Des_Asignacion_cobradores '" & strCodCliente & "'," & session("session_idusuario")
		set rsDesAsig = Conn.execute(strSql1)

		strSql1 = "EXEC Proc_Cambia_Custodio_Deudor '" & strCodCliente & "'," & session("session_idusuario")
		set rsCambiaCustodio = Conn.execute(strSql1)

		strSql = "EXEC proc_Inf_Actualizacion_deuda_UMA_FACT 'INFORME_ESTADO'"
%>

		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700" HEIGHT = "30">
			<tr>
			<td colspan=2>PROCESO EJECUTADO CORRECTAMENTE, EL RESULTADO DE LA CARGA ES EL SIGUIENTE:</td>
			</tr>
		</table>

<%
	  set rsDet=Conn.execute(strSql)
		if not rsDet.eof then
			do until rsDet.eof

		intIdEstadoAct = rsDet("ID_ESTADO_ACT")
		intCodEstadoAct = rsDet("COD_ESTADO_ACT")
		strDescActua = Mid(rsDet("ESTADO_ACT"),1,60)
		intCuentaActua = rsDet("TOTAL")

%>

		<table border=1 width="700" border="1" cellSpacing=0 cellPadding=2  class="Estilo28">

			<tr>
			<td width="400" width="30" align="LEFT"><%=strDescActua%></td>
			<td width="30" width="30" align="right"><%=intCuentaActua%></td>
			</tr>

		</table>

<%			rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing
%>

		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700" HEIGHT = "30">

		<td colspan=1 align="right">
		<input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga_Cliente.asp');return false;">
		</td>

		</table>
<%

End If

If strArchivo <> "" Then


	Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

	strNomArchivoTerceros = "Terceros_cargados_" & Fecha & ".csv"
	terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros

	strTextoArchivoCC = ""
	strTextoArchivoCNC = ""
	strTextoArchivoCA = ""

	strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"& strCodCliente &"/" & strArchivo

	strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_ACTUALIZA_UMA_FACT]') AND type in (N'U'))"
	strSql = strSql & " DROP TABLE [TMP_ACTUALIZA_UMA_FACT]"
	Conn.Execute strSql,64


	strSql = " CREATE TABLE TMP_ACTUALIZA_UMA_FACT ("
	strSql = strSql &" 	ULTIMO_ESTADO VARCHAR(50) NOT NULL,"
	strSql = strSql &" 	ID_CUOTA INT NOT NULL,"
	strSql = strSql &" 	COD_SAP BIGINT NOT NULL,"
	strSql = strSql &"  RUT_ALUMNO VARCHAR(20) NOT NULL,"
	strSql = strSql &"  FOLIO VARCHAR(20) NULL,"
	strSql = strSql &"  VENCIMI SMALLDATETIME NOT NULL)"

	Conn.Execute strSql,64

	'response.write "Conn = " & Conn
	'response.write "strSql " & strSql

	'**********CARGA ARCHIVO************'

	strSqlFile = "BULK INSERT TMP_ACTUALIZA_UMA_FACT FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2, CODEPAGE = 'ACP')"
	Conn.Execute strSqlFile,64%>

	<form name="datos" method="post">
	<table width="990" border="1" bordercolor="#FFFFFF" bgcolor="#FFFFFF">
		<tr>
			<TD height="20" ALIGN=LEFT class="pasos2_i">
				<B>INFORME ACTUALIZACIÓN DE DEUDA MASIVA</B>
			</TD>
		</tr>
	<tr>
		<td align="center">


	<%'VERIFICA QUE TODOS LOS ID_CUOTA ASOCIADOS Y LOS DOCUMENTOS SEAN CORRESPONDIENTES A LA LLAVE DEL CLIENTE'

	strSql = "	SELECT COUNT(*) AS CANTIDAD FROM TMP_ACTUALIZA_UMA_FACT LEFT JOIN CUOTA ON TMP_ACTUALIZA_UMA_FACT.ID_CUOTA = CUOTA.ID_CUOTA"
	strSql = strSql &" 																	  AND TMP_ACTUALIZA_UMA_FACT.COD_SAP = CUOTA.NRO_CLIENTE_DOC"
	strSql = strSql &" 																	  AND CUOTA.NRO_DOC = TMP_ACTUALIZA_UMA_FACT.FOLIO"
	strSql = strSql &"																	  AND CUOTA.FECHA_VENC = CAST(TMP_ACTUALIZA_UMA_FACT.VENCIMI AS DATETIME)"
	strSql = strSql &"	WHERE CUOTA.ID_CUOTA IS NULL"


	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCuentaErroresID = rsTemp("CANTIDAD")
	Else
		intCuentaErroresID = 0
	End if

	'CUENTA LOS REGISTROS DUPLICADOS EN BASE DE ACTUALIZACION'

	strSql = "SELECT COUNT(REPETIDOS) AS REPETIDOS FROM"
	strSql = strSql &" (SELECT ROW_NUMBER() OVER (PARTITION BY COD_SAP ORDER BY COD_SAP ASC) AS REPETIDOS FROM TMP_ACTUALIZA_UMA_FACT) AS REP"
	strSql = strSql &" WHERE REPETIDOS > 1"

	set rsTemp= Conn.execute(strSql)
	if not rsTemp.eof then
		intCargaDuplicadosBase = rsTemp("REPETIDOS")
	Else
		intCargaDuplicadosBase = 0
	End if

	strSql = "EXEC proc_Inf_Actualizacion_deuda_UMA_FACT 'INFORME_ESTADO'"

		intEstadosDocyaActualizados = 0
		intEstadosPositivos = 0
		intEstadosNegativos = 0
		intEstadosVolveraActivar = 0
		intEstadosNoReconocidos = 0
		intEstadosconError = 0


	  set rsDet=Conn.execute(strSql)
		if not rsDet.eof then
			do until rsDet.eof

		'response.write "Estados=" & rsDet("COD_ESTADO_ACT")

		If rsDet("COD_ESTADO_ACT") = 1 then

		intEstadosDocyaActualizados= intEstadosDocyaActualizados + 1

		End If

		If rsDet("COD_ESTADO_ACT") = 2 then

		intEstadosPositivos = intEstadosPositivos + 1

		End If

		If rsDet("COD_ESTADO_ACT") = 3 or rsDet("COD_ESTADO_ACT") = 4 or rsDet("COD_ESTADO_ACT") = 5 then

		intEstadosNegativos = intEstadosNegativos + 1

		End If

		If rsDet("COD_ESTADO_ACT") = 6 then

		intEstadosVolveraActivar= intEstadosVolveraActivar + 1

		End If

		If rsDet("COD_ESTADO_ACT") = 7 then

		intEstadosNoReconocidos= intEstadosNoReconocidos + 1

		End If

		If rsDet("COD_ESTADO_ACT") = 9 then

		intEstadosconError= intEstadosconError + 1

		End If


			rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing

	'response.write "intEstadosDocyaActualizados=" & intEstadosDocyaActualizados
	'response.write "intEstadosPositivos=" & intEstadosPositivos
	'response.write "intEstadosNegativos=" & intEstadosNegativos
	'response.write "intEstadosVolveraActivar=" & intEstadosVolveraActivar
	'response.write "intEstadosNoReconocidos=" & intEstadosNoReconocidos
	'response.write "intEstadosconError=" & intEstadosconError

If intCuentaErroresID > 0 or intCargaDuplicadosBase > 0 or intEstadosNoReconocidos > 0 or intEstadosconError > 0 then%>

		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700" HEIGHT = "50">
			<tr>
			<td colspan=2>ACTUALIZACION CON ERRORES, FAVOR VERIFIQUE LA BASE DE CARGA, LOS ERRORES SON LOS SIGUIENTES:</td>
			</tr>
		</table>

		<table border=1 width="700" border="1" bgcolor="#<%=session("COLTABBG2")%>" cellSpacing=0 cellPadding=2  class="Estilo28">

			<tr>

			<% If intCuentaErroresID > 0 then %>
			<tr><td>ERROR DE DATOS DE LLAVE, ESTO SIGNIFICA QUE LA LLAVE UNICA (RUT ALUMNO, FOLIO, FECHA_VENCI Y EL ID DEL SISTEMA NO ESTAN CARGADOS, FAVOR REVISAR BASE DE ESTADO</td></tr>
			<% End If%>

			<% If intCargaDuplicadosBase > 0 then %>
			<tr><td>EN LA BASE DE ACTUALIZACION VIENEN CARGADOS REGISTROS DUPLICADOS, SISTEMA VALIDA LLAVE DEFINIDA</td></tr>
			<% End If%>

			<% If intEstadosNoReconocidos > 0 then %>
			<tr><td>EN LA BASE VIENEN ESTADOS DE DEUDA INGRESADOS POR USTED NO RECONOCIDOS, FAVOR COMUNIQUESE CON EL ADMINISTRADOR SI DESEA AGRGAR UN NUEVO ESTADO</td></tr>
			<% End If%>

			<% If intEstadosconError > 0 then %>
			<tr><td>ERROR 4</td></tr>
			<% End If%>

			</tr>

		</table>

<%
Else

	If intEstadosDocyaActualizados > 0 then

	AbrirScg1()
		ssrSql = "EXEC proc_Inf_Actualizacion_deuda_UMA_FACT 'INFORME_DEUDA'"

		set rsDet1=Conn.execute(ssrSql)

			if not rsDet1.eof then
				do until rsDet1.eof

					strUltimoEstado = rsDet1("ULTIMO_ESTADO")
					intCodSap = rsDet1("COD_SAP")
					intIdEstadoAct = rsDet1("ID_ESTADO_ACT")

					If intIdEstadoAct = "11" Then

						ssrSql = "		   UPDATE CARGA_UMA_FACT"
						ssrSql = ssrSql &" SET FECHA_ULT_ESTADO_ACT = getdate(), ULT_ESTADO_ACT = '"& strUltimoEstado &"'"
						ssrSql = ssrSql &" WHERE COD_SAP = "& intCodSap &" and ULT_ESTADO_ACT IS NULL"

					set rsUpdate2=Conn.execute(ssrSql)

						ssrSql = "		   UPDATE CARGA_UMA_FACT"
						ssrSql = ssrSql &" SET FECHA_PENULT_ESTADO_ACT = FECHA_ULT_ESTADO_ACT, PENULT_ESTADO_ACT = ULT_ESTADO_ACT,FECHA_ULT_ESTADO_ACT = getdate(), ULT_ESTADO_ACT = '"& strUltimoEstado &"'"
						ssrSql = ssrSql &" WHERE COD_SAP = "& intCodSap &" and ULT_ESTADO_ACT IS NOT NULL and ULT_ESTADO_ACT <> '"& strUltimoEstado &"'"

					set rsUpdate2=Conn.execute(ssrSql)

					'Response.write "<BR>ssrSql =" & ssrSql

					End If

					rsDet1.movenext
				loop
			end if
			rsDet1.close
			set rsDet1=nothing
	CerrarScg1()

	End If


	If intEstadosPositivos > 0 then

	%>
		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700">
			<tr>
			<td colspan=2><b>ESTATUS CARGA: DOCUMENTOS CON CAMBIOS POSITIVOS</b></td>
			</tr>
		</table>


	<%set rsDet=Conn.execute(strSql)
		if not rsDet.eof then
			do until rsDet.eof

		intIdEstadoAct = rsDet("ID_ESTADO_ACT")
		intCodEstadoAct = rsDet("COD_ESTADO_ACT")
		strDescActua = Mid(rsDet("ESTADO_ACT"),1,60)
		intCuentaActua = rsDet("TOTAL")

		If intCodEstadoAct = 2 then%>

		<table border=1 width="700" border="1" bgcolor="#<%=session("COLTABBG2")%>" cellSpacing=0 cellPadding=2  class="Estilo28">

			<tr>

			<td width="15" align="center"><INPUT TYPE=checkbox NAME="CH_<%=intIdEstadoAct%>"></td>
			<td width="400" width="30" align="LEFT"><%=strDescActua%></td>
			<td width="30" width="30" align="right"><%=intCuentaActua%></td>
			</tr>

		</table>

	  <%End If

			rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing%>

		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700">
			<tr>
			<td colspan=2>&nbsp;</td>
			</tr>
		</table>
		<table>
			<tr>
			<td colspan=2 bgcolor="#FFFFFF" HEIGHT = "50">&nbsp;</td>
			</tr>
		</table>

	<%End If


	If intEstadosDocyaActualizados > 0 then

	%>
		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700">
			<tr>
			<td colspan=2><b>ESTATUS CARGA: DOCUMENTOS YA ACTUALIZADOS PREVIAMENTE EN SISTEMA</b></td>

			</tr>
		</table>


	<%  set rsDet=Conn.execute(strSql)
		if not rsDet.eof then
			do until rsDet.eof

		intIdEstadoAct = rsDet("ID_ESTADO_ACT")
		intCodEstadoAct = rsDet("COD_ESTADO_ACT")
		strDescActua = Mid(rsDet("ESTADO_ACT"),1,60)
		intCuentaActua = rsDet("TOTAL")

		If intIdEstadoAct = "11" then
		strDisabled = "disabled"
		Else
		strDisabled = ""
		End If

		If intCodEstadoAct = 1 then%>

		<table border=1 width="700" border="1" bgcolor="#<%=session("COLTABBG2")%>" cellSpacing=0 cellPadding=2  class="Estilo28">

			<tr>

			<td width="15" align="center"><INPUT TYPE=checkbox NAME="CH_<%=intIdEstadoAct%>" <%=strDisabled%>></td>
			<td width="400" width="30" align="LEFT"><%=strDescActua%></td>
			<td width="30" width="30" align="right"><%=intCuentaActua%></td>

			</tr>

		</table>

	 <%End If

			rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing

	End If


	If intEstadosNegativos > 0 then

	%>
		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700">
			<tr>
			<td colspan=2><b>ESTATUS CARGA: DOCUMENTOS CON CAMBIOS NEGATIVOS</b></td>

			</tr>
		</table>


	<%set rsDet=Conn.execute(strSql)
		if not rsDet.eof then
			do until rsDet.eof

		intIdEstadoAct = rsDet("ID_ESTADO_ACT")
		intCodEstadoAct = rsDet("COD_ESTADO_ACT")
		strDescActua = Mid(rsDet("ESTADO_ACT"),1,60)
		intCuentaActua = rsDet("TOTAL")

		If (intIdEstadoAct = "31" or intIdEstadoAct = "32" or intIdEstadoAct = "51" or intIdEstadoAct = "52" or intIdEstadoAct = "53") and TraeSiNo(session("perfil_adm"))="No" then
		strDisabled = "disabled"
		Else
		strDisabled = ""
		End If

		If intCodEstadoAct = 3 or intCodEstadoAct = 4 or intCodEstadoAct = 5 then%>

		<table border=1 width="700" border="1" bgcolor="#<%=session("COLTABBG2")%>" cellSpacing=0 cellPadding=2  class="Estilo28">

			<tr>

			<td width="15" align="center"><INPUT TYPE=checkbox NAME="CH_<%=intIdEstadoAct%>" <%=strDisabled%>></td>
			<td width="400" width="30" align="LEFT"><%=strDescActua%></td>
			<td width="30" width="30" align="right"><%=intCuentaActua%></td>

			</tr>

		</table>

	<%End If

			rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing

	End If

	If intEstadosVolveraActivar > 0 then

	%>
		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700">
			<tr>
			<td colspan=2><b>ESTATUS CARGA: DOCUMENTOS QUE SE VUELVEN A ACTIVAR</b></td>

			</tr>
		</table>


	<%set rsDet=Conn.execute(strSql)
		if not rsDet.eof then
			do until rsDet.eof

		intIdEstadoAct = rsDet("ID_ESTADO_ACT")
		intCodEstadoAct = rsDet("COD_ESTADO_ACT")
		strDescActua = Mid(rsDet("ESTADO_ACT"),1,60)
		intCuentaActua = rsDet("TOTAL")

		If (intIdEstadoAct = "61" and TraeSiNo(session("perfil_adm"))="No")  or intIdEstadoAct = "62" then
		strDisabled = "disabled"
		Else
		strDisabled = ""
		End If

		If intCodEstadoAct = 6 then%>

		<table border=1 width="700" border="1" bgcolor="#<%=session("COLTABBG2")%>" cellSpacing=0 cellPadding=2  class="Estilo28">

			<tr>

			<td width="15" align="center"><INPUT TYPE=checkbox NAME="CH_<%=intIdEstadoAct%>" <%=strDisabled%>></td>
			<td width="400" width="30" align="LEFT"><%=strDescActua%></td>
			<td width="30" width="30" align="right"><%=intCuentaActua%></td>

			</tr>

		</table>

	<%End If

			rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing

	End If


	If intEstadosNoReconocidos > 0 then

	%>
		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700">
			<tr>
			<td colspan=2><b>ESTATUS CARGA: DOCUMENTOS CON ESTADO DE ACTUALIZACIÓN NO INGRESADO EN SISTEMA</b></td>

			</tr>
		</table>


	<%set rsDet=Conn.execute(strSql)
		if not rsDet.eof then
			do until rsDet.eof

		intIdEstadoAct = rsDet("ID_ESTADO_ACT")
		intCodEstadoAct = rsDet("COD_ESTADO_ACT")
		strDescActua = Mid(rsDet("ESTADO_ACT"),1,60)
		intCuentaActua = rsDet("TOTAL")

		If intIdEstadoAct = "71" then
		strDisabled = "disabled"
		Else
		strDisabled = ""
		End If

		If intCodEstadoAct = 7 then%>

		<table border=1 width="700" border="1" bgcolor="#<%=session("COLTABBG2")%>" cellSpacing=0 cellPadding=2  class="Estilo28">

			<tr>

			<td width="15" align="center"><INPUT TYPE=checkbox NAME="CH_<%=intIdEstadoAct%>" <%=strDisabled%>></td>
			<td width="400" width="30" align="LEFT"><%=strDescActua%></td>
			<td width="30" width="30" align="right"><%=intCuentaActua%></td>

			</tr>

		</table>

	<%End If

			rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing

	End If


	If intEstadosconError > 0 then

	%>
		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700">
			<tr>
			<td colspan=2><b>ESTATUS CARGA: DOCUMENTOS CON ERRORES DE ESTADO</b></td>

			</tr>
		</table>


	<%set rsDet=Conn.execute(strSql)
		if not rsDet.eof then
			do until rsDet.eof

		intIdEstadoAct = rsDet("ID_ESTADO_ACT")
		intCodEstadoAct = rsDet("COD_ESTADO_ACT")
		strDescActua = Mid(rsDet("ESTADO_ACT"),1,60)
		intCuentaActua = rsDet("TOTAL")

		If intIdEstadoAct = "91" then
		strDisabled = "disabled"
		Else
		strDisabled = ""
		End If

		If intCodEstadoAct = 7 then%>

		<table border=1 width="700" border="1" bgcolor="#<%=session("COLTABBG2")%>" cellSpacing=0 cellPadding=2  class="Estilo28">

			<tr>

			<td width="15" align="center"><INPUT TYPE=checkbox NAME="CH_<%=intIdEstadoAct%>" <%=strDisabled%>></td>
			<td width="400" width="30" align="LEFT"><%=strDescActua%></td>
			<td width="30" width="30" align="right"><%=intCuentaActua%></td>

			</tr>

		</table>

	  <%End If

			rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing

	End If

End If%>


<table border=0 width="700" height = "50" bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">

		<%If intCuentaErroresID = 0 and intCargaDuplicadosBase = 0 and intEstadosNoReconocidos = 0 and intEstadosconError = 0 then%>
		<td colspan=1 align="center" width="700" >
		<input type="BUTTON" value="Actualizar Deuda" name="terminar" onClick="Procesar();return false;">
		</td>
		<%End If%>

		<td colspan=1 align="right">
		<input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga_Cliente.asp');return false;">
		</td>

		</td></tr>
</table>
	 <%

End if



%>


				</td>
			  </tr>
			</table>
</form>
</body>
</html>

