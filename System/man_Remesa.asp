<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
    <!--#include file="sesion.asp"-->
    <link href="../css/style.css" rel="stylesheet" type="text/css">
    <link href="../css/style_generales_sistema.css" rel="stylesheet">
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
<!--#include file="../lib/asp/comunes/select/Remesa.inc"-->
<!--#include file="../lib/asp/comunes/delete/Remesa.inc"-->
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
            ReDim sarrBorrar( 2, intIndiceFilasFin )
            ObtenerListaBorrar sarrBorrar, intIndiceFilasIni, intIndiceFilasFin
            sstrMensajeBorrar = ""
            strMsg = ""
            For intRow = intIndiceFilasIni to intIndiceFilasFin
            	'Response.write "arrborrar = " &sarrBorrar( 0, intRow )
                If sarrBorrar( 0, intRow ) = "on" Then
                    delete_Remesa Conn, sarrBorrar(1, intRow), sarrBorrar(2, intRow), strMsg
                    If strMsg <> "" Then sstrMensajeBorrar = sstrMensajeBorrar & " - " & strMsg & sarrBorrar(1, intRow)
                End If
            Next
        End If
    End If

    Dim sarrRegistros, sintTotalRegistros
    select_Remesa Conn, sarrRegistros,  sintTotalRegistros


%>

<!--#include file="../lib/asp/comunes/General/CalculoFilasPorPagina.inc"-->

<TITLE>Mantenedor</TITLE>

</HEAD>

<BODY>
<!--#include file="../lib/asp/comunes/General/JavaScriptBotonesMantenedor.inc"-->

<FORM NAME="mantenedorForm" method="POST" action="man_Remesa.asp" onmouseover="highlightButton('start')" onmouseout="highlightButton('')">
<INPUT TYPE=HIDDEN NAME="sintNumFilasPagina" VALUE="<%= sintNumFilasPagina %>">
<INPUT TYPE=HIDDEN NAME="sintPagina" VALUE=<%= sintPagina  %>>
<INPUT TYPE=HIDDEN NAME="intIndiceFilasIni" VALUE="<%= intIndiceFilasIni  %>">
<INPUT TYPE=HIDDEN NAME="intIndiceFilasFin" VALUE="<%= intIndiceFilasFin  %>">

<div class="titulo_informe">MANTENEDOR DE ASIGNACIONES</div>
<br>
<TABLE align="right">
  
    <TR BGCOLOR="#FFFFFF" HEIGHT="30" VALIGN="MIDDLE">
        <TD ALIGN="RIGHT">
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Aceptar" VALUE="Borrar" onClick="BAceptar( this.form )">
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Nuevo" VALUE="Nuevo" onClick="IrNuevo( 'man_RemesaForm.asp' )">
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Terminar" VALUE="Terminar" onClick="IrTerminar('menuadm.asp')">
        </TD>
    </TR>

</TABLE>

<!--#include file="../lib/asp/comunes/general/BarraNavegacion.inc"-->

<TABLE BORDER="0" CELLSPACING="0" width="100%" CELLPADDING="0"  CLASS="tabla1">
    <!--Encabezado-->
    <TR BGCOLOR="#F3F3F3">
    <TD><b>&nbsp;</TD>
    <TD><b>Id</TD>
    <TD><b>Cliente</TD>
    <TD><b>Nombre</TD>
    <TD><b>Descripción</TD>
    <TD><b>F.Llegada</TD>
    <TD><b>F.Carga</TD>
    <TD><b>Activo</TD>
    <TD><b>&nbsp;</TD>
    <TD><center><b>Borrar<center></TD>
    </TR>

    <!--Datos-->
    <TR>
    <TD colspan=11 background="../Imagenes/arriba.jpg">
    <img border="0" src="../Imagenes/arribaizq.jpg">
    </TD>
    </TR>
<%
    For intRow = intIndiceFilasIni to intIndiceFilasFin
        If intRow Mod 2 = 0 Then
            sstrColor = "#F3F3F3"
        Else
            sstrColor = "#FFFFFF"
        End if

		If Trim(sarrRegistros(intRow).Item("COD_REMESA"))="" or IsNull(sarrRegistros(intRow).Item("COD_REMESA")) Then
			intCodRemesa=0
		Else
			intCodRemesa=sarrRegistros(intRow).Item("COD_REMESA")
		End if
			'Response.write "<br>intCodRemesa =" & intCodRemesa

		strCliente = TraeCampoId(Conn, "DESCRIPCION", Trim(sarrRegistros(intRow).Item("COD_CLIENTE")), "CLIENTE", "COD_CLIENTE")

%>
<TR BGCOLOR="<%=sstrColor%>">
	<TD>&nbsp;</TD>
    <TD>
        <INPUT TYPE=HIDDEN NAME="sstrId<%= Cstr( intRow ) %>" VALUE="<%= Trim(sarrRegistros(intRow).Item("COD_REMESA"))%>">
        <INPUT TYPE=HIDDEN NAME="sstrIdCliente<%= Cstr( intRow ) %>" VALUE="<%= Trim(sarrRegistros(intRow).Item("COD_CLIENTE"))%>">
        <A HREF="man_RemesaForm.asp?sintNuevo=0&COD_REMESA=<%=Trim(sarrRegistros(intRow).Item("COD_REMESA"))%>&COD_CLIENTE=<%=Trim(sarrRegistros(intRow).Item("COD_CLIENTE"))%>">
        <%= Trim(sarrRegistros(intRow).Item("COD_REMESA")) %>
        </A>
    </TD>
    <TD>
		<A HREF="man_RemesaForm.asp?sintNuevo=0&COD_REMESA=<%=Trim(sarrRegistros(intRow).Item("COD_REMESA"))%>&COD_CLIENTE=<%=Trim(sarrRegistros(intRow).Item("COD_CLIENTE"))%>">
		<%= Trim(sarrRegistros(intRow).Item("COD_CLIENTE")) & " - " & strCliente%>
		</A>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("NOMBRE")) %>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("DESCRIPCION")) %>
    </TD>
     <TD>
		<%= Trim(sarrRegistros(intRow).Item("FECHA_LLEGADA")) %>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("FECHA_CARGA")) %>
    </TD>
    <TD>
		<%= TraeSiNo(Trim(sarrRegistros(intRow).Item("ACTIVO"))) %>
    </TD>
    <TD>&nbsp;</TD>
    <TD WIDTH="10%" ALIGN="CENTER">
    	<INPUT TYPE="checkbox" NAME="borrar<%= intRow %>">
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
	mantenedorForm.action='man_Remesa.asp?intFiltro=1';
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
    sarrBorrar(2, i) = Request( "sstrIdCliente" & i)
  Next
End Sub

</SCRIPT>
