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
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>

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
''*****************************

strProcesar = Request("strProcesar")

strCodCliente= Trim(Request("CB_CLIENTE"))
intIdUsuario = session("session_idusuario")

'Response.write "<BR>strCodCliente =" & strCodCliente

if Request("archivo") <> "" then
	strArchivo=Request("archivo")
End if

if Request("strTipoProceso") <> "" then
	strTipoProceso=Request("strTipoProceso")
End if

AbrirSCG()

''ACA DEBERIA TRAER LOS REGISTROS
dim ConnectDBQ,rsPlanilla,dbc

CerrarSCG()

If strProcesar = "SI" Then

	strSql = "EXEC proc_Inf_Actualizacion_deuda 'INFORME_ESTADO', '" & strCodCliente & "'," & intIdUsuario & ",''"

AbrirSCG()

 	set rsDet=Conn.execute(strSql)
	
		Cont= 0
		if not rsDet.eof then
			do until rsDet.eof
				
				intIdEstadoAct = rsDet("ID_ESTADO_ACT")

				strObjeto = "CH_" & intIdEstadoAct

				'Response.write "<BR>strObjeto =" & strObjeto

				If UCASE(Request(strObjeto)) = "ON" Then
				
					If Cont = 0 then
					
					strEstadosAct = CStr(intIdEstadoAct)
					
					Else
					
					strEstadosAct = strEstadosAct + "," + CStr(intIdEstadoAct)

					End If
					
				Cont= Cont + 1

				End If

				rsDet.movenext
			loop
		end if
		rsDet.close
		set rsDet=nothing
		
CerrarSCG()

	'Response.write "<BR>strEstadosAct =" & strEstadosAct

	'ACTUALIZA DEUDA DE DOCUMENTOS CON LOS ESTADOS SELECCIONADOS'
	
	If strEstadosAct = "" then
	strEstadosAct = "0"
	End IF
		
	strSql = "EXEC proc_Actualizacion_deuda 'ACTUALIZA_DEUDA', '" & strCodCliente & "','" & intIdUsuario & "','" & strEstadosAct & "'"
	'Response.write "<BR>strSql =" & strSql	
	
AbrirSCG()
	
 	set rsDet=Conn.execute(strSql)

	strSql1 = "EXEC Proc_Des_Asignacion_cobradores '" & strCodCliente & "'," & session("session_idusuario")
	'Response.write "<BR>strSql1 =" & strSql1	
	set rsDesAsig = Conn.execute(strSql1)

	strSql1 = "EXEC Proc_Cambia_Custodio_Deudor '" & strCodCliente & "'," & session("session_idusuario")
	'Response.write "<BR>strSql1 =" & strSql1	
	set rsCambiaCustodio = Conn.execute(strSql1)

CerrarSCG()	
%>
		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700" HEIGHT = "30">
			<tr>
			<td colspan=2>PROCESO EJECUTADO, EL ESTATUS ACTUAL DEL ARCHIVO DE ACTUALIZACION ES EL SIGUIENTE:</td>
			</tr>
		</table>
<%

	strSql = "EXEC proc_Inf_Actualizacion_deuda 'INFORME_ESTADO', '" & strCodCliente & "'," & intIdUsuario & ",''"

AbrirSCG()	
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

CerrarSCG()	
%>

		<table border=1 bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" width="700" HEIGHT = "30">

		<td colspan=1 align="right">
		<input type="BUTTON" value="Volver" name="terminar" onClick="Terminar('man_carga_Cliente.asp');return false;">
		</td>

		</table>
<%

Else

AbrirSCG()	

	If strArchivo <> "" Then

		Fecha= right("00"&Day(DATE()), 2) &right("00"&(Month(DATE())), 2) &Year(DATE())

		strNomArchivoTerceros = "Terceros_cargados_" & Fecha & ".csv"
		terceroCSV = request.serverVariables("APPL_PHYSICAL_PATH") & "Logs\" & strNomArchivoTerceros

		''response.write "Conn = " & "UploadFolder\"& strArchivo

		strFileDir = session("ses_ruta_sitio_Fisica")  &"/Archivo/CargaActualizaciones/"&strCodCliente &"/" & strArchivo

		strSql = " IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[TMP_ACTUALIZA_UMA]') AND type in (N'U'))"
		strSql = strSql & " DROP TABLE [TMP_ACTUALIZA_UMA]"
		Conn.Execute strSql,64

		strSql = " CREATE TABLE TMP_ACTUALIZA_UMA ("
		strSql = strSql &" 	ULTIMO_ESTADO VARCHAR(50) NOT NULL,"
		strSql = strSql &" 	ID_CUOTA INT NOT NULL,"
		strSql = strSql &" 	COD_SAP BIGINT NOT NULL,"
		strSql = strSql &"  RUT_ALUMNO VARCHAR(20) NOT NULL,"
		strSql = strSql &"  FOLIO VARCHAR(20) NULL,"
		strSql = strSql &"  VENCIMI SMALLDATETIME NOT NULL)"

		Conn.Execute strSql,64

		'response.write "Conn = " & Conn
		
		'**********CARGA ARCHIVO************'

		strSqlFile = "BULK INSERT TMP_ACTUALIZA_UMA FROM '" & strFileDir & "' with ( fieldterminator =';',ROWTERMINATOR ='\n', FIRSTROW = 2, CODEPAGE = 'ACP')"
		
		'response.write "strSqlFile " & strSqlFile
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

		strSql = "	SELECT COUNT(*) AS CANTIDAD FROM TMP_ACTUALIZA_UMA LEFT JOIN CUOTA ON TMP_ACTUALIZA_UMA.ID_CUOTA = CUOTA.ID_CUOTA"
		strSql = strSql &" 																	  AND TMP_ACTUALIZA_UMA.COD_SAP = CUOTA.NRO_CLIENTE_DOC"
		strSql = strSql &" 																	  AND CUOTA.RUT_DEUDOR = REVERSE(SUBSTRING(REVERSE(TMP_ACTUALIZA_UMA.RUT_ALUMNO),2,10))+'-'+SUBSTRING(REVERSE(TMP_ACTUALIZA_UMA.RUT_ALUMNO),1,1)"
		strSql = strSql &" 																	  AND CUOTA.NRO_DOC = TMP_ACTUALIZA_UMA.FOLIO"
		strSql = strSql &"																	  AND CUOTA.FECHA_VENC = CAST(TMP_ACTUALIZA_UMA.VENCIMI AS DATETIME)"
		strSql = strSql &"	WHERE CUOTA.ID_CUOTA IS NULL"

		set rsTemp= Conn.execute(strSql)
		if not rsTemp.eof then
			intCuentaErroresID = rsTemp("CANTIDAD")
		Else
			intCuentaErroresID = 0
		End if

		'CUENTA LOS REGISTROS DUPLICADOS EN BASE DE ACTUALIZACION'

		strSql = "SELECT COUNT(REPETIDOS) AS REPETIDOS FROM"
		strSql = strSql &" (SELECT ROW_NUMBER() OVER (PARTITION BY COD_SAP ORDER BY COD_SAP ASC) AS REPETIDOS FROM TMP_ACTUALIZA_UMA) AS REP"
		strSql = strSql &" WHERE REPETIDOS > 1"

		set rsTemp= Conn.execute(strSql)
		if not rsTemp.eof then
			intCargaDuplicadosBase = rsTemp("REPETIDOS")
		Else
			intCargaDuplicadosBase = 0
		End if

			intEstadosDocyaActualizados = 0
			intEstadosPositivos = 0
			intEstadosNegativos = 0
			intEstadosVolveraActivar = 0
			intEstadosNoReconocidos = 0
			intEstadosconError = 0

		strSql = "EXEC proc_Inf_Actualizacion_deuda 'INFORME_ESTADO', '" & strCodCliente & "'," & intIdUsuario & ",''"
		
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
		 
	<%End if

End If%>

	</td>
  </tr>
</table>

</form>
</body>
</html>

<script language="JavaScript" type="text/JavaScript">

	function Terminar( sintPaginaTerminar ) {
		self.location.href = sintPaginaTerminar
	}

	function Procesar()
		{
			if (confirm("¿ Está REALMENTE seguro de Actualizar los documentos según los estados seleccionados?"))
			{
				datos.terminar.disabled = true;
				datos.action='Man_Actualiza_UMA.asp?strProcesar=SI&CB_CLIENTE=<%=strCodCliente%>';
				datos.submit();
			}
		}
</script>