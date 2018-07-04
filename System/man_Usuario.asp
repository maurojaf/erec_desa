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
    <!--#include file="../lib/asp/comunes/select/UsuarioQry.inc"-->
    <!--#include file="../lib/asp/comunes/delete/Usuario.inc"-->
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
                    delete_Usuario Conn, sarrBorrar(1, intRow), strMsg
                    If strMsg <> "" Then sstrMensajeBorrar = sstrMensajeBorrar & " - " & strMsg & sarrBorrar(1, intRow)
                End If
            Next
        End If
    End If

    Dim sarrRegistros, sintTotalRegistros
    select_UsuarioQry Conn, " WHERE ID_USUARIO >= 100 ", sarrRegistros,  sintTotalRegistros


%>

<!--#include file="../lib/asp/comunes/General/CalculoFilasPorPagina.inc"-->

<TITLE>Mantenedor</TITLE>
<LINK rel="stylesheet" TYPE="text/css" HREF="../css/isk_style.css">
</HEAD>

<BODY>
<!--#include file="../lib/asp/comunes/General/JavaScriptBotonesMantenedor.inc"-->

<FORM NAME="mantenedorForm" method="POST" action="man_Usuario.asp" onmouseover="highlightButton('start')" onmouseout="highlightButton('')">
<INPUT TYPE=HIDDEN NAME="sintNumFilasPagina" VALUE="<%= sintNumFilasPagina %>">
<INPUT TYPE=HIDDEN NAME="sintPagina" VALUE=<%= sintPagina  %>>
<INPUT TYPE=HIDDEN NAME="intIndiceFilasIni" VALUE="<%= intIndiceFilasIni  %>">
<INPUT TYPE=HIDDEN NAME="intIndiceFilasFin" VALUE="<%= intIndiceFilasFin  %>">
<div class="titulo_informe">MANTENEDOR DE USUARIO</div>
<br>
<TABLE WIDTH="100%" border=0 cellspacing=0 CLASS="tabla1">
    <TR BGCOLOR="#FFFFFF" HEIGHT="30" VALIGN="MIDDLE">
        <TD>

		</TD>
		<TD ALIGN=RIGHT>

        </TD>

        <TD ALIGN=RIGHT>
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Aceptar" VALUE="Borrar" onClick="BAceptar( this.form )">
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Nuevo" VALUE="Nuevo" onClick="IrNuevo( 'man_UsuarioForm.asp' )">
        <INPUT TYPE="BUTTON" class="fondo_boton_100" NAME="Terminar" VALUE="Terminar" onClick="IrTerminar('menuadm.asp')">
        </TD>
    </TR>

</TABLE>

<!--#include file="../lib/asp/comunes/general/BarraNavegacion.inc"-->

<TABLE BORDER=0 CELLSPACING=0 CELLPADDING=0 WIDTH="100%" CLASS="tabla1">
    <!--Encabezado-->
    <TR BGCOLOR="#F3F3F3">
    <TD><b>&nbsp;</TD>
    <TD><b>Id</TD>
    <TD WIDTH="200"><b>Nombre</TD>
    <TD><b>Login</TD>
    <TD><b>Rut</TD>
    <TD><b>Anexo</TD>
    <TD><b>P.Adm</TD>
    <TD><b>P.Sup</TD>
    <TD><b>P.Cob</TD>
    <TD><b>P.Caja</TD>
    <TD><b>P.Full</TD>
    <TD><b>Activo</TD>
    <TD><b>&nbsp;</TD>
    <TD><center><b>Borrar<center></TD>
    </TR>

    <!--Datos-->
    <TR>
    <TD colspan=14 background="../imagenes/arriba.jpg">
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

		If Trim(sarrRegistros(intRow).Item("ID_USUARIO"))="" or IsNull(sarrRegistros(intRow).Item("ID_USUARIO")) Then
			intID_USUARIO=0
		Else
			intID_USUARIO=sarrRegistros(intRow).Item("ID_USUARIO")
		End if
			'Response.write "<br>intID_USUARIO =" & intID_USUARIO

%>
<TR BGCOLOR="<%=sstrColor%>">
	<TD>&nbsp;</TD>
    <TD>
        <INPUT TYPE=HIDDEN NAME="sstrId<%= Cstr( intRow ) %>" VALUE="<%= Trim(sarrRegistros(intRow).Item("ID_USUARIO"))%>">
        <%= Trim(sarrRegistros(intRow).Item("ID_USUARIO")) %>
    </TD>
    <TD>
		<A HREF="man_UsuarioForm.asp?sintNuevo=0&ID_USUARIO=<%=Trim(sarrRegistros(intRow).Item("ID_USUARIO"))%>">
		<%= Trim(sarrRegistros(intRow).Item("NOMBRES_USUARIO")) & " " & Trim(sarrRegistros(intRow).Item("APELLIDO_PATERNO"))& " " & Trim(sarrRegistros(intRow).Item("APELLIDO_MATERNO"))%>
		</A>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("LOGIN")) %>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("RUT_USUARIO")) %>
    </TD>
    <TD>
		<%= Trim(sarrRegistros(intRow).Item("ANEXO")) %>
    </TD>
     <TD>
		<%= TraeSiNo(Trim(sarrRegistros(intRow).Item("PERFIL_ADM"))) %>
    </TD>
    <TD>
		<%= TraeSiNo(Trim(sarrRegistros(intRow).Item("PERFIL_SUP"))) %>
    </TD>
    <TD>
		<%= TraeSiNo(Trim(sarrRegistros(intRow).Item("PERFIL_COB"))) %>
    </TD>
    <TD>
		<%= TraeSiNo(Trim(sarrRegistros(intRow).Item("PERFIL_CAJA"))) %>
    </TD>
    <TD>
		<%= TraeSiNo(Trim(sarrRegistros(intRow).Item("PERFIL_FULL"))) %>
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
	mantenedorForm.action='man_Usuario.asp?intFiltro=1';
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
