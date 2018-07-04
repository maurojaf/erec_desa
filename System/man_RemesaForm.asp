<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->
     <link href="../css/style_generales_sistema.css" rel="stylesheet">
<!--#include file="../lib/asp/comunes/General/MostrarRegistro.inc"-->
<!--#include file="../lib/asp/comunes/general/rutinasTraeCampo.inc"-->

<% ' Capa 1 ' %>
<!--#include file="../lib/asp/comunes/odbc/AdoVbs.inc"-->
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/asp/comunes/odbc/ObtenerRecordSet.inc"-->
<!--#include file="../lib/asp/comunes/odbc/ObtenerRegistros.inc"-->

<% ' Capa 2 ' %>
<!--#include file="../lib/asp/comunes/recordset/Remesa.inc"-->

<%
	Response.CodePage=65001
	Response.charset ="utf-8"

	sintNuevo = request("sintNuevo")
	IntId = request("COD_REMESA")
    IntIdCliente = request("COD_CLIENTE")

	'Response.Write ("***" & IntId & "***")

    If sintNuevo = 1 Then
        strFormMode="Nuevo"
        IntId=0
    Else
        strFormMode="Edit"
    End If

    AbrirSCG()

	recordset_Remesa Conn, srsRegistro, IntId, IntIdCliente
    If Not srsRegistro.Eof Then
		intIdActivo = Trim(srsRegistro("ACTIVO"))
	Else
		intIdActivo = ""
    End If

%>

<TITLE>Mantenedor de Remesas</TITLE>
<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
</HEAD>

<BODY BGCOLOR='FFFFFF'>

<SCRIPT Language=JavaScript>
function Continuar() {

	document.forms[0].submit()
    return false
}
</SCRIPT>

<style type="text/css">
	.hdr_i{
		background-color: #C9DEF2;
	}
</style>
<FORM NAME="mantenedorForm"  action="man_RemesaAction.asp" method="POST" onmouseover="highlightButton('start')" onmouseout="highlightButton('')">
	<INPUT TYPE="HIDDEN" NAME="strFormMode" VALUE="<%= strFormMode %>">



<div class="titulo_informe">MANTENEDOR DE REMESA</div>
<br>
 <table width="90%" border="0" align="center">
    	<tr>
    		<td class="subtitulo_informe">> INGRESO REMESA</td>
     </tr>
</Table>


<table width="90%" border="0" align="center">
	<tr >
		<td class="hdr_i">Codigo Remesa</td>

		<%
		If strFormMode = "Nuevo" Then
		%>
		<td class="td_t"><% general_MostrarCampo "COD_REMESA", False, Null, Null,srsRegistro %></td>
		<%
		Else

			Response.Write "<INPUT TYPE=HIDDEN NAME=COD_REMESA VALUE=""" & srsRegistro("COD_REMESA") & """>"
		%>
			<td class="td_t"><%=srsRegistro("COD_REMESA") %></td>
		<%
		End If
        %>
	</TR>
	<tr >
		<td class="hdr_i">Codigo Cliente</td>

		<%
		If strFormMode = "Nuevo" Then
		%>
		<td class="td_t">
			<select name="COD_CLIENTE">
				<option value="Seleccionar">Seleccionar</option>
				<%
				ssql="SELECT COD_CLIENTE, DESCRIPCION FROM CLIENTE ORDER BY DESCRIPCION "
				set rsTemp= Conn.execute(ssql)
				if not rsTemp.eof then
					do until rsTemp.eof%>
						<option value="<%=rsTemp("COD_CLIENTE")%>"<%if strIdCOD_CLIENTE=rsTemp("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsTemp("DESCRIPCION")%></option>
					<%
					rsTemp.movenext
					loop
				end if
				rsTemp.close
				set rsTemp=nothing

				%>
			</select>
		</td>
		<%
		Else
			Response.Write "<INPUT TYPE=HIDDEN NAME=COD_CLIENTE VALUE=""" & srsRegistro("COD_CLIENTE") & """>"
			strCliente = TraeCampoId(Conn, "DESCRIPCION", Trim(srsRegistro("COD_CLIENTE")), "CLIENTE", "COD_CLIENTE")
		%>
			<td class="td_t"><%=srsRegistro("COD_CLIENTE") %> - <%=strCliente%></td>
		<%
		End If
		%>
	</TR>
 	<tr >
		<td class="hdr_i">Nombre</td>
		<td class="td_t"><% general_MostrarCampo "NOMBRE", False, Null, Null,srsRegistro %></td>
	</TR>
	<tr >
		<td class="hdr_i">Descripcion</td>
		<td class="td_t"><% general_MostrarCampo "DESCRIPCION", False, Null, Null,srsRegistro %></td>
	</TR>
	<tr >
		<td class="hdr_i">Fecha llegada</td>
		<td class="td_t"><% general_MostrarCampo "FECHA_LLEGADA", False, Null, Null,srsRegistro %></td>
	</TR>
	<tr >
		<td class="hdr_i">Fecha carga</td>
		<td class="td_t"><% general_MostrarCampo "FECHA_CARGA", False, Null, Null,srsRegistro %></td>
	</TR>
	<tr >
		<td class="hdr_i">Activo</td>
		<td class="td_t"><% general_MostrarCampo "ACTIVO", False, Null, Null,srsRegistro %></td>
	</tr>

</table>

<table width="100%" border="0">
     <TR>
	  <td align=center  width="25%">
	   <INPUT TYPE="BUTTON" class="fondo_boton_100" value="Guardar" name="B1" onClick="Continuar();return false;">
	   <input type="BUTTON" class="fondo_boton_100" value="Terminar" name="terminar" onClick="Terminar('man_Remesa.asp');return false;"></TD>
	  </TD>
    </TR>
    <%If sintNuevo = 1 Then %>
    <TR>
    <TD align=center >
     <IMG BORDER="0" src="../imagenes/bolita.jpg" WIDTH=10>=Campo requerido
     </TD>
    </TR>
	<%End If %>
</table>

    </TD>
    </TR>
</TABLE>
</FORM>

<SCRIPT>
function Refrescar(strTipo){

	mantenedorForm.action='man_RemesaForm.asp?strTipoPropiedad=' + strTipo;
	if (mantenedorForm.strFormMode.value == 'Nuevo') {
		location.href="man_RemesaForm.asp?sintNuevo=1&strTipoPropiedad=" + strTipo;
		}
	else {
	 	mantenedorForm.submit();
	}
}
</SCRIPT>

<%CerrarSCG()%>

</BODY>
</HTML>




