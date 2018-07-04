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
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link rel="stylesheet" href="../css/style_generales_sistema.css">

<%
strCOD_CLIENTE 		= request("cliente")
strRUT_DEUDOR 		= request("strRUT_DEUDOR")
iddeudor 			= request("Id_Deudor")
strObservaciones 	= Request("TX_OBSERVACIONES")
intOrigen 			= request("intOrigen")
Procesar 			= request("Procesar")
intCodUsuario 		= session("session_idusuario")
strTipo 			= request("strTipo")
strCOD_CLIENTE 		= session("ses_codcli")


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
--> 
}
</style>

<script language="JavaScript" src="../javascripts/cal2.js"></script>
<script language="JavaScript" src="../javascripts/cal_conf2.js"></script>
<script language="JavaScript" src="../javascripts/validaciones.js"></script>


<link href="../css/style.css" rel="Stylesheet">
</head>
<body>
<form name="datos" method="post">
<INPUT TYPE="hidden" NAME="intOrigen" value="<%=intOrigen%>">

<%
	if Trim(strTipo) = "" Then
		strTipo = "0"
	End If

AbrirSCG()

		strSql = "SELECT COUNT(RUT) AS EXISTERUT, MAX(OBSERVACION) AS OBSERVACION FROM INFORMACION_DEUDOR_SUBCLIENTE"
		strSql = strSql & " WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND RUT = '" & strRUT_DEUDOR & "' AND TIPO_INFORMACION = " & strTipo

		''Response.write "strSql=" & strSql
		set rsBusqueda=Conn.execute(strSql)

		ExisteRegistro = rsBusqueda("EXISTERUT")
		strObsDeudorInicial = rsBusqueda("OBSERVACION")

		'Response.write "ExisteRegistro=" & ExisteRegistro

	 If Trim(request("strGraba")) = "SI" and Trim(request("strTipo")) = "1" and  strObservaciones = "" Then%>

		<script language="JavaScript" type="text/JavaScript">
			alert('Debe ingresar observación para grabar información a sistema');
		</script>

	 <%ElseIf Trim(request("strGraba")) = "SI" and Trim(request("strTipo")) = "1" and  strObservaciones = strObsDeudorInicial Then%>

		<script language="JavaScript" type="text/JavaScript">
			alert('Esta ingresando la misma información en sistema, no se han realizado cambios');

			window.navigate('mas_datos_adicionales_deudor.asp?strRUT_DEUDOR=<%=strRUT_DEUDOR%>&strObservaciones=<%=strObservaciones%>&Id_Deudor=<%=Iddeudor%>&Procesar=1;');
		</script>

	 <%ElseIf Trim(request("strGraba")) = "SI" and Trim(request("strTipo")) = "1" and ExisteRegistro = 0 Then

		strSql = "INSERT INTO INFORMACION_DEUDOR_SUBCLIENTE (COD_CLIENTE,RUT_DEUDOR,RUT,TIPO_INFORMACION,OBSERVACION, FECHA_CREACION,USUARIO_CREACION)"
		strSql = strSql & " VALUES ('" & strCOD_CLIENTE & "','" & strRUT_DEUDOR & "','" & strRUT_DEUDOR & "'," & Trim(request("strTipo"))  & ",'" & Mid(strObservaciones,1,300) & "',GETDATE(),"& intCodUsuario &")"

		''Response.write "strSql=" & strSql
		set rsInsert=Conn.execute(strSql)%>

		<script language="JavaScript" type="text/JavaScript">
			alert('Información de deudor incorporada a sistema');

			window.navigate('mas_datos_adicionales_deudor.asp?strRUT_DEUDOR=<%=strRUT_DEUDOR%>&strObservaciones=<%=strObservaciones%>&Id_Deudor=<%=Iddeudor%>&Procesar=1;');
		</script>

	 <%ElseIf Trim(request("strGraba")) = "SI" and Trim(request("strTipo")) = "1" and ExisteRegistro = 1 Then

		strSql = "UPDATE INFORMACION_DEUDOR_SUBCLIENTE SET OBSERVACION = '" & Mid(strObservaciones,1,300) & "',OBSERVACION_RESPALDO = '" & Mid(strObservaciones,1,300) & "', FECHA_MODIFICACION = GETDATE(),USUARIO_MODIFICACION = "& intCodUsuario
		strSql = strSql & " WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND RUT = '" & strRUT_DEUDOR & "' AND TIPO_INFORMACION = " & Trim(request("strTipo"))
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)%>

		<script language="JavaScript" type="text/JavaScript">
			alert('Información de deudor auditada en sistema');

			window.navigate('mas_datos_adicionales_deudor.asp?strRUT_DEUDOR=<%=strRUT_DEUDOR%>&strObservaciones=<%=strObservaciones%>&Id_Deudor=<%=Iddeudor%>&Procesar=1;');
		</script>

	<%End If

	If Trim(request("strLimpia")) = "SI" Then

		strSql = "UPDATE INFORMACION_DEUDOR_SUBCLIENTE SET OBSERVACION = NULL,OBSERVACION_RESPALDO = '" & Mid(strObservaciones,1,300) & "', FECHA_LIMPIEZA = GETDATE(),USUARIO_LIMPIEZA = "& intCodUsuario
		strSql = strSql & " WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND RUT = '" & strRUT_DEUDOR & "' AND TIPO_INFORMACION = " & Trim(request("strTipo"))
		''Response.write "strSql=" & strSql
		set rsUpdate=Conn.execute(strSql)%>

		<script language="JavaScript" type="text/JavaScript">
			alert('Información de deudor borrada de sistema');

			window.navigate('mas_datos_adicionales_deudor.asp?strRUT_DEUDOR=<%=strRUT_DEUDOR%>&strObservaciones=<%=strObservaciones%>&Id_Deudor=<%=Iddeudor%>&Procesar=1;');
		</script>

	<%End If

		strSql="SELECT ISNULL(ADIC_1,'SIN INFORMACION') AS ADIC_1, ISNULL(ADIC_2,'SIN INFORMACION') AS ADIC_2, ISNULL(ADIC_3,'SIN INFORMACION') AS ADIC_3, ISNULL(FECHA_ESTADO_ETAPA,CAST('01/01/1900' AS DATETIME)) AS FECHA_ESTADO_ETAPA,FECHA_INGRESO,[dbo].[fun_trae_fecha_venc_inf_activa] (COD_CLIENTE,RUT_DEUDOR) as FVIA ,[dbo].[fun_trae_fecha_creacion_inf_activa] (COD_CLIENTE,RUT_DEUDOR) as FCIA ,[dbo].[fun_trae_fecha_ult_normalizacion] (COD_CLIENTE,RUT_DEUDOR) as FUNORM, UPPER(OBSERVACIONES_BACKOFFICE) AS OBS    FROM DEUDOR WHERE ID_DEUDOR = " & Iddeudor & ""

		''response.write "strSql=" & strSql
		'Response.End

		set rsDET=Conn.execute(strSql)

			if Not rsDET.eof Then
				strAdic1 = rsDET("ADIC_1")
				strAdic2 = rsDET("ADIC_2")
				strAdic3 = rsDET("ADIC_3")
				dtmFechaEstadoEtapa = rsDET("FECHA_ESTADO_ETAPA")
				dtmFechaCreacion = rsDET("FECHA_INGRESO")
				dtmFVIA = rsDET("FVIA")
				dtmFCIA = rsDET("FCIA")
				dtmFUNORM = rsDET("FUNORM")
			End If

		strSql="SELECT IsNull(ADIC1_DEUDOR,'ADIC_1') as ADIC_1, IsNull(ADIC2_DEUDOR,'ADIC_2') as ADIC_2, IsNull(ADIC3_DEUDOR,'ADIC_3') as ADIC_3 FROM CLIENTE WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "'"
		''response.write "strSql=" & strSql
		'Response.End

		set rsDET=Conn.execute(strSql)
		if Not rsDET.eof Then
			strNombreAdic1 = rsDET("ADIC_1")
			strNombreAdic2 = rsDET("ADIC_2")
			strNombreAdic3 = rsDET("ADIC_3")
		End If


		strSql="SELECT RUT_DEUDOR,RUT,TIPO_INFORMACION,REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(REPLACE(OBSERVACION,'Ñ','N'),'Á','A'),'É','E'),'Í','I'),'Ó','O'),'Ú','U'),CHAR(10),' '),CHAR(13),' ') AS 'OBSERVACION',FECHA_CREACION, USUARIO_CREACION FROM INFORMACION_DEUDOR_SUBCLIENTE WHERE COD_CLIENTE = '" & strCOD_CLIENTE & "' AND RUT_DEUDOR = '" & strRUT_DEUDOR & "' AND RUT = '" & strRUT_DEUDOR & "' AND TIPO_INFORMACION = 1"

		''response.write "strSql=" & strSql
		'Response.End

		set rsInfDeudor=Conn.execute(strSql)

		if Not rsInfDeudor.eof Then

			strObsDeudor = rsInfDeudor("OBSERVACION")

		End If

CerrarSCG()

%>

<table width="100%" height="167" align="center" class="intercalado">
<thead>
	    <tr>
	    	<td height="21" bordercolor="#999999" class="" colspan="4">
		    FICHA DEL DEUDOR : <%=strRUT_DEUDOR%></td>
	    </tr>
</thead>
<body>
		<tr  height="17" bordercolor="#999999">
			<td width="150"  class="columna_tipo1">CUANDO PAGA</td>
			<td width="200">&nbsp;</td>
			<td width="150"  class="columna_tipo1"><%=strNombreAdic1%></td>
			<td width="200"><%=strAdic1%></td>
		</tr>
		<tr  height="17" bordercolor="#999999">
			<td class="columna_tipo1">DONDE PAGA</td>
			<td>&nbsp;</td>
			<td  class="columna_tipo1"><%=strNombreAdic2%></td>
			<td><%=strAdic2%></td>
		</tr>
		<tr  height="17" bordercolor="#999999">
			<td  class="columna_tipo1">A QUIEN PAGA</td>
			<td>&nbsp;</td>
			<td  class="columna_tipo1"><%=strNombreAdic3%></td>
			<td><%=strAdic3%></td>
		</tr>
		<tr  height="17" bordercolor="#999999">
			<td  class="columna_tipo1">COMO PAGA</td>
			<td>&nbsp;</td>
			<td  class="columna_tipo1">Fecha Estado Etapa</td>
			<td><%=dtmFechaEstadoEtapa%></td>
		</tr>
		<tr  height="17" bordercolor="#999999">
			<td  class="columna_tipo1">COMO ATIENDE</td>
			<td>&nbsp;</td>
			<td  class="columna_tipo1">Fecha Venc. Inf Activa</td>
			<td><%=dtmFVIA%></td>
		</tr>
		<tr  height="17" bordercolor="#999999">
			<td  class="columna_tipo1">DOCUMENTOS QUE EXIGE</td>
			<td>&nbsp;</td>
			<td  class="columna_tipo1">Fecha Creación Inf Activa</td>
			<td><%=dtmFCIA%></td>
		</tr>
		<tr  height="17" bordercolor="#999999">
			<td  class="columna_tipo1">PORTALES DE CONSULTA</td>
			<td>&nbsp;</td>
			<td  class="columna_tipo1">Fecha Ultima Normalización</td>
			<td>&nbsp;</td>
		</tr>

		<tr  height="17" bordercolor="#999999">
			<td>&nbsp;</td>
			<td>&nbsp;</td>
			<td  class="columna_tipo1">Fecha Creación</td>
			<td><%=dtmFechaCreacion%></td>
		</tr>
	</body>

</table>
<br>
<table width="90%" bordercolor="#FFFFFF" class="estilo_columnas" align="center">
<thead>
		<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13" >
			<td colspan = "2" >OBSERVACION DEUDOR (Max. 300 Caracteres)</td>
		</tr>
</thead>
		<tr>
		<td colspan = "2" align="CENTER">
		<TEXTAREA NAME="TX_OBSERVACIONES" ROWS="4" COLS="90"><%=strObsDeudor%></TEXTAREA>
		</td>
		</tr>


		<tr>
		<td align="right">
		<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="PROCESAR" value="Grabar/Auditar" onClick="GrabarObsDeudor() ;" class="Estilo8">
		</td>
		<td align="left">&nbsp;
		<INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Limpiar" value="Borrar" onClick="limpiar();" class="Estilo8">
		</td>
		</tr>

</table>


<INPUT TYPE="hidden" NAME="strLimpia" value="">
<INPUT TYPE="hidden" NAME="strGraba" value="">
<INPUT TYPE="hidden" NAME="strTipo" value="">
<INPUT TYPE="hidden" NAME="strDesConfirmarTarea" value="">
<INPUT TYPE="hidden" NAME="strConfirmarTarea" value="">
<INPUT TYPE="hidden" NAME="strAgendar" value="">



</form>
</body>
</html>
<script language="JavaScript" type="text/JavaScript">

function GrabarObsDeudor() {

	if (confirm("¿ Está seguro de que desea guardar esta información? "))
	{
		datos.strGraba.value='SI';
		datos.strTipo.value='1';

		datos.submit();
	}
}

function limpiar() {

	if (confirm("¿ Está seguro de borrar la información ingresada? "))
	{

	if (confirm("¿ Está realmente seguro de borrar la información ingresada? "))
	{
			datos.strLimpia.value='SI';
			datos.strTipo.value='1';
			datos.submit();
	}

}

}

</script>


















