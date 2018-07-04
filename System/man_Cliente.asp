<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
    <!--#include file="sesion.asp"-->

    <% ' General ' %>
    <!--#include file="arch_utils.asp"-->
    <!--#include file="../lib/asp/comunes/general/Minimo.inc"-->
    <!--#include file="../lib/asp/comunes/general/rutinasTraeCampo.inc"-->
    <!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

    <% ' Capa 1 ' %>
    <!--#include file="../lib/asp/comunes/odbc/Adovbs.inc"-->
    <!--#include file="../lib/asp/comunes/odbc/ObtenerRegistros.inc"-->
    <!--#include file="../lib/asp/comunes/odbc/BorrarRegistro.inc"-->
    <% ' Capa 2 ' %>
    <!--#include file="../lib/asp/comunes/select/Cliente.inc"-->
    <!--#include file="../lib/asp/comunes/delete/Cliente.inc"-->

    <link href="../css/style.css" rel="stylesheet" type="text/css">
    <link href="../css/style_generales_sistema.css" rel="stylesheet">

<%
    Response.CodePage=65001
    Response.charset ="utf-8"

    AbrirSCG()

    ''Obtiene los parámetros
    sintPagina  = Request("sintPagina")
    If Trim(sintPagina) <> "" Then
        intIndiceFilasIni = Request("intIndiceFilasIni")
        intIndiceFilasFin = Request("intIndiceFilasFin")
        If intIndiceFilasIni <> "" Then
            ReDim sarrBorrar( 1, intIndiceFilasFin )
            ObtenerListaBorrar sarrBorrar, intIndiceFilasIni, intIndiceFilasFin
            sstrMensajeBorrar = ""
            strMsg = ""
            For intRow = intIndiceFilasIni to intIndiceFilasFin
            	'Response.write "arrborrar = " &sarrBorrar( 0, intRow )
                If sarrBorrar( 0, intRow ) = "on" Then
                    delete_Cliente Conn, sarrBorrar(1, intRow), strMsg
                    If strMsg <> "" Then sstrMensajeBorrar = sstrMensajeBorrar & " - " & strMsg & sarrBorrar(1, intRow)
                End If
            Next
        End If
    End If

    Dim sarrRegistros, sintTotalRegistros
    select_Cliente Conn, sarrRegistros,  sintTotalRegistros


%>

<!--#include file="../lib/asp/comunes/General/CalculoFilasPorPagina.inc"-->
<TITLE>Mantenedor</TITLE>
</HEAD>

<BODY>
<!--#include file="../lib/asp/comunes/General/JavaScriptBotonesMantenedor.inc"-->

<FORM NAME="mantenedorForm" method="POST" action="man_Cliente.asp" onmouseover="highlightButton('start')" onmouseout="highlightButton('')">
<INPUT TYPE=HIDDEN NAME="sintNumFilasPagina" VALUE="<%= sintNumFilasPagina %>">
<INPUT TYPE=HIDDEN NAME="sintPagina" VALUE=<%= sintPagina  %>>
<INPUT TYPE=HIDDEN NAME="intIndiceFilasIni" VALUE="<%= intIndiceFilasIni  %>">
<INPUT TYPE=HIDDEN NAME="intIndiceFilasFin" VALUE="<%= intIndiceFilasFin  %>">
<div class="titulo_informe">MANTENEDOR DE CLIENTES</div>
<TABLE WIDTH="100%" border="0" cellspacing="0" CLASS="tabla1">
    <TR BGCOLOR="#FFFFFF" HEIGHT="30" VALIGN="MIDDLE">
        <TD ALIGN=RIGHT>
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Aceptar" VALUE="Borrar" onClick="BAceptar( this.form )">
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Nuevo" VALUE="Nuevo" onClick="IrNuevo( 'man_ClienteForm.asp' )">
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Terminar" VALUE="Terminar" onClick="IrTerminar('menuadm.asp')">
        </TD>
    </TR>
    <TR BGCOLOR="#F3F3F3" HEIGHT="3" VALIGN="MIDDLE">
        <TD COLSPAN=5></TD>
    </TR>
</TABLE>

<!--#include file="../lib/asp/comunes/general/BarraNavegacion.inc"-->

<TABLE BORDER="0" CELLSPACING="0" CELLPADDING="0" WIDTH="100%" CLASS="tabla1">
    <!--Encabezado-->
    <TR BGCOLOR="#F3F3F3">
    <TD><b>&nbsp;</TD>
    <TD><b>Id</TD>
    <TD><b>Nombre</TD>
    <TD><b>Razon Social</TD>
    <TD><b>Nombre Fantasia</TD>
    <TD><b>Rut</TD>
    <TD><b>Direccion</TD>
    <TD><b>Telefono</TD>
    <TD><b>Activo</TD>
    <TD><b>&nbsp;</TD>
    <TD WIDTH="5%"><center><b>Borrar<center></TD>
    </TR>

    <!--Datos-->
    <TR>
    <TD colspan="11" background="../imagenes/arriba.jpg">
    <img border="0" src="../imagenes/arribaizq.jpg">
    </TD>
    </TR>
<%
    For intRow = intIndiceFilasIni to intIndiceFilasFin
        If intRow Mod 2 = 0 Then
            sstrColor = "#F3F3F3"
        Else
            sstrColor = "#FFFFFF"
        End if

		If Trim(sarrRegistros(intRow).Item("COD_CLIENTE"))="" or IsNull(sarrRegistros(intRow).Item("COD_CLIENTE")) Then
			intCOD_CLIENTE=0
		Else
			intCOD_CLIENTE=sarrRegistros(intRow).Item("COD_CLIENTE")
		End if
			'Response.write "<br>intCOD_CLIENTE =" & intCOD_CLIENTE

%>
<TR BGCOLOR="<%=sstrColor%>">
	<TD>&nbsp;</TD>
    <TD>
        <INPUT TYPE=HIDDEN NAME="sstrId<%= Cstr( intRow ) %>" VALUE="<%= Trim(sarrRegistros(intRow).Item("COD_CLIENTE"))%>">
        <%= Trim(sarrRegistros(intRow).Item("COD_CLIENTE")) %>
    </TD>
    <TD>
		<A HREF="man_ClienteForm.asp?sintNuevo=0&COD_CLIENTE=<%=Trim(sarrRegistros(intRow).Item("COD_CLIENTE"))%>">
		<%= Trim(sarrRegistros(intRow).Item("DESCRIPCION"))%>
		</A>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("RAZON_SOCIAL")) %>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("NOMBRE_FANTASIA")) %>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("RUT")) %>
    </TD>
     <TD>
		<%= Trim(sarrRegistros(intRow).Item("DIRECCION")) %>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("FONO_1")) %>
    </TD>
    <TD>
		<%= TraeSiNo(Trim(sarrRegistros(intRow).Item("ACTIVO"))) %>
    </TD>
    <TD>&nbsp;</TD>
    <TD WIDTH="10%" ALIGN=CENTER>
    	<INPUT TYPE=checkbox NAME="borrar<%= intRow %>">
    </TD>
   </TR>
<%
    Next
%>

 <!--#include file="../lib/asp/comunes/general/BarraNavegacion.inc"-->

</FORM>
</BODY>
</HTML>
<SCRIPT>
function Refrescar(){
	mantenedorForm.action='man_Cliente.asp?intFiltro=1';
	mantenedorForm.submit();
}

</SCRIPT>
<%
	CerrarSCG()
%>


<!-- ---------------------------------------------------- -->
<SCRIPT LANGUAGE=VBScript RUNAT=Server>

Sub ObtenerListaBorrar ( ByRef sarrBorrar, ByRef Inicio, ByVal Fin)
  For i = Inicio to Fin
    sarrBorrar(0, i) = Request( "borrar" & i )
    sarrBorrar(1, i) = Request( "sstrId" & i)
  Next
End Sub

</SCRIPT>
