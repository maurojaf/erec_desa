<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->

	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<link href="../css/style.css" rel="stylesheet" type="text/css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet">

<%
	Response.CodePage=65001
	Response.charset ="utf-8"
%>

<TITLE>Menú</TITLE>

</HEAD>

<BODY>
<div class="titulo_informe">MÓDULO ADMINISTRACIÓN</div>
<br>
<table ALIGN="CENTER" WIDTH="600" border="0" CLASS="tabla1">
<tr BGCOLOR="#FFFFFF">
	<td class="hdr_i" width="21">
		<img src='../imagenes/boton_no_1.gif' align='absmiddle'>
	</td>
	<td class="hdr_i" width="100">
		General
	</td>
	<td class="hdr" width="21">
		<img src='../imagenes/cuadrado_t1.gif' align='absmiddle'>
	</td>
	<td class="hdr_i" width="300">
	<% If TraeSiNo(session("perfil_full"))="Si" Then %>
		<A HREF="man_Usuario.asp">
			Mantenedor de Usuarios
		</A>
	<% End If%>
	</td>
</tr>

<tr BGCOLOR="#FFFFFF">
	<td class="hdr_i" width="21">
		&nbsp
	</td>
	<td class="hdr_i" width="100">
		&nbsp
	</td>
	<td class="hdr" width="21">
		<img src='../imagenes/cuadrado_t1.gif' align='absmiddle'>
	</td>
	<td class="hdr_i" width="300">
		<A HREF="man_Cliente.asp">
			Mantenedor de Mandantes
		</A>
	</td>
</tr>


<tr BGCOLOR="#FFFFFF">
	<td class="hdr_i" width="21">
		&nbsp
	</td>
	<td class="hdr_i" width="100">
		&nbsp
	</td>
	<td class="hdr" width="21">
		<img src='../imagenes/cuadrado_t1.gif' align='absmiddle'>
	</td>
	<td class="hdr_i" width="300">
		<A HREF="man_Remesa.asp">
			Mantenedor de Asignaciones
		</A>
	</td>
</tr>

<tr BGCOLOR="#FFFFFF">
	<td class="hdr_i" width="21">
		&nbsp
	</td>
	<td class="hdr_i" width="100">
		&nbsp
	</td>
	<td class="hdr" width="21">
		<img src='../imagenes/cuadrado_t1.gif' align='absmiddle'>
	</td>
	<td class="hdr_i" width="300">
		<A HREF="man_Carga.asp">
			Modulo de cargas y actualizacion
		</A>
	</td>
</tr>

<tr BGCOLOR="#FFFFFF">
	<td class="hdr_i" width="21">
		&nbsp
	</td>
	<td class="hdr_i" width="100">
		&nbsp
	</td>
	<td class="hdr" width="21">
		<img src='../imagenes/cuadrado_t1.gif' align='absmiddle'>
	</td>
	<td class="hdr_i" width="300">
		<A HREF="man_Export.asp">
			Exportación documentos y gestiones
		</A>
	</td>
</tr>


<tr BGCOLOR="#FFFFFF">
	<td colspan=4 class="hdr_i" >
		&nbsp
	</td>
</tr>

</Table>

</body>
</html>