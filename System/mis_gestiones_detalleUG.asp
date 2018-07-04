<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">

<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
<!--#include file="../lib/comunes/rutinas/funcionesBD.inc" -->
<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
<!--#include file="../lib/lib.asp"-->

<%

	Response.CodePage=65001
	Response.charset ="utf-8"
	
hdd_cod_cliente = request("cmb_cliente")
txt_FechaIni = request("txt_FechaIni")
txt_FechaFin = request("txt_FechaFin")

'Response.write "<br>GESTION=" & request("GESTION")
'Response.End

intCodEjecutivo = request("intCodEjecutivo")
intCodCampana = request("intCodCampana")
If Trim(request("GESTION")) <> "" Then
	arrGestion = split(request("GESTION"),"-")
	COD_CATEGORIA = arrGestion(0)
	COD_SUB_CATEGORIA = arrGestion(1)
	COD_GESTION = arrGestion(2)
End If




'COD_CATEGORIA = Mid(request("GESTION"),1,1)
'COD_SUB_CATEGORIA = Mid(request("GESTION"),2,1)
'COD_GESTION = Mid(request("GESTION"),3,1)

abrirscg()


strSql = "SELECT RUT_DEUDOR,COD_CATEGORIA,COD_SUB_CATEGORIA,ISnull(COD_GESTION,0) as "
strSql = strSql & " COD_GESTION,FECHA_INGRESO,HORA_INGRESO,IsNull(FECHA_COMPROMISO,'') as FECHA_COMPROMISO,IsNull(OBSERVACIONES,'&nbsp') as OBSERVACIONES, "
strSql = strSql & " convert(varchar(2),COD_CATEGORIA) + convert(varchar(2),COD_SUB_CATEGORIA) + "
strSql = strSql & "	convert(varchar(2),IsNull(COD_GESTION,0)) as COD_GESTION, ID_USUARIO , DD.TELEFONO TELEFONO_ASOCIADO FROM GESTIONES "
strSql = strSql & "	LEFT JOIN DEUDOR_TELEFONO DD ON DD.ID_TELEFONO = GESTIONES.ID_MEDIO_GESTION "

strSql = strSql & "	WHERE COD_CLIENTE='" & hdd_cod_cliente & "'"

If Trim(COD_CATEGORIA) <> "" and Trim(COD_SUB_CATEGORIA) <> "" and Trim(COD_GESTION) <> "" Then
	strSql = strSql & "	AND  COD_CATEGORIA='" & COD_CATEGORIA & "'"
	strSql = strSql & "	AND COD_SUB_CATEGORIA='" & COD_SUB_CATEGORIA & "'"
	strSql = strSql & "	AND  COD_GESTION='" & COD_GESTION & "'"
End if

If Trim(COD_CATEGORIA) <> "" and Trim(COD_SUB_CATEGORIA) <> "" and Trim(COD_GESTION) <> "" Then
	strSql = strSql & "	AND  COD_CATEGORIA='" & COD_CATEGORIA & "'"
	strSql = strSql & "	AND COD_SUB_CATEGORIA='" & COD_SUB_CATEGORIA & "'"
	strSql = strSql & "	AND  COD_GESTION='" & COD_GESTION & "'"
End if

strSql = strSql & " AND RUT_DEUDOR IN (SELECT RUT_DEUDOR FROM DEUDOR "
strSql = strSql & "	WHERE COD_CLIENTE='" & hdd_cod_cliente & "' AND ULTIMA_GESTION = '" & request("GESTION") & "')"


'strSql = strSql & "	AND FECHA_INGRESO>='" & txt_FechaIni & "'"
'strSql = strSql & "	AND FECHA_INGRESO<='" & txt_FechaFin & "'"

If TraeSiNo(session("perfil_adm"))="Si" or TraeSiNo(session("perfil_sup"))="Si" Then
	If Trim(intCodEjecutivo) <> "" Then
			strSql = strSql & "	AND ID_USUARIO = " & intCodEjecutivo
	End If

Else
	If Trim(intCodEjecutivo) <> "" Then
		strSql = strSql & "	AND ID_USUARIO = " & intCodEjecutivo
	Else
		strSql = strSql & "	AND ID_USUARIO = " & session("session_idusuario")
	End If
End If

If Trim(intCodCampana) <> "" Then
	strSql = strSql & "	AND ID_CAMPANA = " & intCodCampana
End If

strSql = strSql & "	group by RUT_DEUDOR,COD_CATEGORIA,COD_SUB_CATEGORIA,ISnull(COD_GESTION,0), "
strSql = strSql & "	FECHA_INGRESO,HORA_INGRESO,FECHA_COMPROMISO,OBSERVACIONES ,COD_GESTION, ID_USUARIO , DD.TELEFONO  "
strSql = strSql & " order by RUT_DEUDOR,COD_CATEGORIA,COD_SUB_CATEGORIA,ISnull(COD_GESTION,0), "
strSql = strSql & " FECHA_INGRESO,HORA_INGRESO,FECHA_COMPROMISO,OBSERVACIONES,ID_USUARIO , DD.TELEFONO "
'Response.write strSql
'Response.End


SET rsGES=Conn.execute(strSql)
%>
<title>Detalle Gestiones</title>
<style type="text/css">
<!--
.Estilo37 {color: #FFFFFF}
-->
</style>
</head>
<body>
<form name="Free" method="post">
<center>

<table width="100%" border="0" cellspacing="0" cellpadding="0" class="Estilo13">
<tr  class="Estilo20">
<td width="100%" align="LEFT"><a href="javascript:history.back();">Volver</a></td>
<!--td ><a href="javascript:goExportarExcel();"><img src="../imagenes/exportarex.gif" alt=""  border="0"></a></td-->
</tr>
</table>

<table width="100%" border="1" bordercolor = "#<%=session("COLTABBG")%>" cellSpacing=0 cellPadding=2>

<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
	<td>Rut&nbsp;&nbsp;Deudor</td>
	<td>Fecha Ingreso</td>
	<td>Hora Ingreso</td>
	<td>Fono</td>
	<td>Gestión</td>
	<td>Ejecutivo</td>
	<td>Fecha Compromiso</td>
	<td>OBSERVACIONES</td>
</tr>
<%Do while not rsGES.eof

If Trim(rsges("OBSERVACIONES"))="" Then strObservaciones = "&nbsp;" Else strObservaciones = rsges("OBSERVACIONES")
If Trim(Mid(rsges("FECHA_COMPROMISO"),1,10))="01/01/1900" Then strFecComp = "&nbsp;" Else strFecComp = rsges("FECHA_COMPROMISO")

strEjecutivo = TraeCampoId(Conn, "NOMBRES_USUARIO", Trim(rsges("ID_USUARIO")), "USUARIO", "ID_USUARIO") & "-" & TraeCampoId(Conn, "APELLIDO_PATERNO", Trim(rsges("ID_USUARIO")), "USUARIO", "ID_USUARIO")

strSql = "SELECT G.COD_CATEGORIA, G.COD_SUB_CATEGORIA, G.COD_GESTION,C.DESCRIPCION + ' ' + S.DESCRIPCION + ' ' +  G.DESCRIPCION as DESCRIP"
strSql = strSql & " FROM GESTIONES_TIPO_CATEGORIA C, GESTIONES_TIPO_SUBCATEGORIA S, GESTIONES_TIPO_GESTION G"
strSql = strSql & " WHERE C.COD_CATEGORIA = S.COD_CATEGORIA"
strSql = strSql & " AND C.COD_CATEGORIA = G.COD_CATEGORIA"
strSql = strSql & " AND S.COD_SUB_CATEGORIA = G.COD_SUB_CATEGORIA"
strSql = strSql & " AND CAST(G.COD_CATEGORIA AS VARCHAR(2)) + CAST(G.COD_SUB_CATEGORIA AS VARCHAR(2)) + CAST(G.COD_GESTION AS VARCHAR(2)) = '" & rsges("COD_GESTION") & "'"

SET rsNomGestion=Conn.execute(strSql)
If Not rsNomGestion.Eof Then
	strNomGestion = rsNomGestion("DESCRIP")
Else
	strNomGestion = ""
End If

%>

<tr>
	<td class="DatosDeudorTexto" ><font class="TextoDatos">
	<A HREF="principal.asp?rut=<%=rsges("RUT_DEUDOR")%>"><acronym title="Llevar a pantalla de selección"><%=rsges("RUT_DEUDOR")%></acronym></A>
	</font></td>
	<td class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("FECHA_INGRESO") %></font></td>
	<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("HORA_INGRESO") %></font></td>
	<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("TELEFONO_ASOCIADO") %></font></td>
	<td class="DatosDeudorTexto" align="">
	<font class="TextoDatos">
	<%= Trim(strNomGestion) %>
	</font></td>
	<td class="DatosDeudorTexto"><font class="TextoDatos"><%= strEjecutivo %></font></td>
	<td class="DatosDeudorTexto"><font class="TextoDatos"><%= strFecComp %></font></td>
	<td class="DatosDeudorTexto" ><font class="TextoDatos"><%= strObservaciones %></font></td>

</tr>

<%rsGES.movenext
	Loop%>
</table>
<form>
<input type="Hidden" name="cmb_cliente">
</form>

</body>


</html>

<%
rsGES.close
set rsGES=nothing
cerrarscg()
%>

<script language="javascript">
function goPagina()
{	with( document.Free )
	{	action = goPagina.arguments[0];
		submit()
	}
}
function goCambiaMes()
{	with( document.Free)
	{	action = "../informes/cargando.asp?strEnlace='mis_gestiones.asp'";
		submit();
	}
}

function goExportarExcel()
{   with( document.Free)
    {
        open("detalle_excel.asp?cmb_cliente=<%=hdd_cod_cliente%>&COD_CATEGORIA=<%=COD_CATEGORIA%>&COD_SUB_CATEGORIA=<%=COD_SUB_CATEGORIA%>&COD_GESTION=<%=COD_GESTION%>&txt_FechaIni=<%=txt_FechaIni%>&txt_FechaFin=<%=txt_FechaFin%>");
        //submit();
    }
}
function Refrescar()
{

	if(!chkFecha(document.Free.txt_FechaIni))
	{
		document.Free.txt_FechaIni.focus()
		document.Free.txt_FechaIni.select()
		return
	}

	if(!chkFecha(document.Free.txt_FechaFin))
	{
		document.Free.txt_FechaFin.focus()
		document.Free.txt_FechaFin.select()
		return
	}

	with( document.Free )
	{
		//action = 'cons_gestion_consolidado.asp';
		action = "cargando.asp?strEnlace='cons_gestion_consolidado_intranet1.asp'";
		submit();
	}
}

</script>
