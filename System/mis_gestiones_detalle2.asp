<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="sesion.asp"-->
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

''Response.write "<br>GESTION=" & request("GESTION")


intCodEjecutivo = request("intCodEjecutivo")
intCodCampana = request("intCodCampana")
cliente = session("ses_codcli")
If Trim(request("GESTION")) <> "" Then

End If

abrirscg()
strSql = "SELECT G.RUT_DEUDOR,G.COD_CATEGORIA,G.COD_SUB_CATEGORIA,ISnull(G.COD_GESTION,0) as "
strSql = strSql & " COD_GESTION,G.FECHA_INGRESO,G.HORA_INGRESO,IsNull(G.FECHA_COMPROMISO,'') as FECHA_COMPROMISO,IsNull(G.OBSERVACIONES,'&nbsp') as OBSERVACIONES, "
strSql = strSql & " convert(varchar(2),G.COD_CATEGORIA) + convert(varchar(2),G.COD_SUB_CATEGORIA) + "
strSql = strSql & "	convert(varchar(2),IsNull(G.COD_GESTION,0)) as COD_GESTION, G.ID_USUARIO , (CASE WHEN ISNULL(DD.TELEFONO,'SIN FONO')='' THEN 'SIN FONO' ELSE ISNULL(DD.TELEFONO,'SIN FONO') END) AS TELEFONO_ASOCIADO, TC.DESCRIPCION + ' ' + TS.DESCRIPCION + ' ' + TG.DESCRIPCION AS DESCRIPCION FROM GESTIONES G, GESTIONES_TIPO_CATEGORIA TC, GESTIONES_TIPO_SUBCATEGORIA TS, GESTIONES_TIPO_GESTION TG, DEUDOR_TELEFONO DD"
strSql = strSql & " WHERE G.COD_CATEGORIA = TC.COD_CATEGORIA"
strSql = strSql & " AND G.COD_CATEGORIA = TS.COD_CATEGORIA"
strSql = strSql & " AND G.COD_SUB_CATEGORIA = TS.COD_SUB_CATEGORIA"
strSql = strSql & " AND G.COD_CATEGORIA = TG.COD_CATEGORIA"
strSql = strSql & " AND G.COD_SUB_CATEGORIA = TG.COD_SUB_CATEGORIA"
strSql = strSql & " AND G.COD_GESTION = TG.COD_GESTION"
strSql = strSql & " AND G.ID_TELEFONO_ASOCIADO = DD.ID_TELEFONO "
strSql = strSql & " AND G.COD_CLIENTE = TG.COD_CLIENTE"
strSql = strSql & " AND  (	  (G.COD_CATEGORIA = 2 AND G.COD_SUB_CATEGORIA= 2 AND G.COD_GESTION =1) "
strSql = strSql & "		   OR (G.COD_CATEGORIA = 3 AND G.COD_SUB_CATEGORIA= 2 AND G.COD_GESTION =1) "
strSql = strSql & "		   OR (G.COD_CATEGORIA = 4 AND G.COD_SUB_CATEGORIA= 1 AND G.COD_GESTION =1))"
strSql = strSql & "	AND G.COD_CLIENTE = '" & Trim(cliente) & "'"


strSql = strSql & "	AND FECHA_INGRESO>='" & txt_FechaIni & "'"
strSql = strSql & "	AND FECHA_INGRESO<='" & txt_FechaFin & "'"

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


strSql = strSql & "	group by G.RUT_DEUDOR,G.COD_CATEGORIA,G.COD_SUB_CATEGORIA,ISnull(G.COD_GESTION,0), "
strSql = strSql & "	G.FECHA_INGRESO,G.HORA_INGRESO,G.FECHA_COMPROMISO,G.OBSERVACIONES ,G.COD_GESTION, G.ID_USUARIO , DD.ID_TELEFONO,  TC.DESCRIPCION + ' ' + TS.DESCRIPCION + ' ' + TG.DESCRIPCION "
strSql = strSql & " order by G.RUT_DEUDOR,G.COD_CATEGORIA,G.COD_SUB_CATEGORIA,ISnull(G.COD_GESTION,0), "
strSql = strSql & " G.FECHA_INGRESO,G.HORA_INGRESO,G.FECHA_COMPROMISO,G.OBSERVACIONES,G.ID_USUARIO , DD.ID_TELEFONO "
''Response.write strSql
'Response.End


SET rsGES=Conn.execute(strSql)
%>
<title>Detalle Gestiones</title>
<style type="text/css">
<!--
.Estilo37 {color: #FFFFFF}
-->
</style>
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
strFoncomp= rsges("TELEFONO_ASOCIADO")
%>

<tr>
	<td class="DatosDeudorTexto" ><font class="TextoDatos">
	<A HREF="principal.asp?TX_RUT=<%=rsges("RUT_DEUDOR")%>"><acronym title="Llevar a pantalla de selección"><%=rsges("RUT_DEUDOR")%></acronym></A>
	</font></td>
	<td class="DatosDeudorTexto" ><font class="TextoDatos"><%= rsges("FECHA_INGRESO") %></font></td>
	<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("HORA_INGRESO") %></font></td>
	<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("TELEFONO_ASOCIADO") %></font></td>
	<td class="DatosDeudorTexto"><font class="TextoDatos"><%= rsges("DESCRIPCION") %></font></td>
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
