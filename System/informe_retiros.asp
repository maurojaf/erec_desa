<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
<!--#include file="arch_utils.asp"-->
<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
<!--#include file="../lib/comunes/rutinas/rutinasFecha.inc" -->
<!--#include file="../lib/lib.asp"-->
<!--#include file="sesion.asp"-->
<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->

<script language="JavaScript">
function ventanaSecundaria (URL){
	window.open(URL,"DETALLE","width=200, height=200, scrollbars=no, menubar=no, location=no, resizable=no")
}
</script>

<%
Response.CodePage=65001
Response.charset ="utf-8"

inicio= request("inicio")
termino= request("termino")


intFechaIni = inicio
intFechaFin = termino

abrirscg()
	If Trim(inicio) = "" Then
		inicio = TraeFechaActual(Conn)
		inicio = "01/" & Mid(TraeFechaActual(Conn),4,10)
	End If

	If Trim(termino) = "" Then
		termino = TraeFechaActual(Conn)
	End If
cerrarscg()

intCliente = request("CB_CLIENTE")
intCliente=session("ses_codcli")
intOrigen = request("CB_ORIGEN")
intCodRemesa = request("CB_REMESA")
intCodUsuario = request("CB_COBRADOR")

''If Trim(intCodUsuario) = "" Then intCodUsuario = session("session_idusuario")
If Trim(intCodUsuario) = "" Then intCodUsuario = "T"

If Trim(intCliente) = "" Then intCliente = "1000"
%>
<title>INFORME RETIROS</title>

<style type="text/css">
<!--
.Estilo37 {color: #FFFFFF}
-->
</style>
<meta http-equiv="Content-Type" content="text/html;charset=utf-8" />
<table width="100%" border="1" bordercolor="#FFFFFF">
	<tr>
		<TD height="20" ALIGN=LEFT class="pasos2_i">
			<B>RETIROS Y CASTIGOS</B>
		</TD>
		<TD height="20">

		</TD>
	</tr>
</table>

<table width="800" align="CENTER" border="0">
  <tr>
    <td valign="top" background="../imagenes/fondo_coventa.jpg">
	<BR>
	<form name="datos" method="post">
	<table width="100%" border="0">
		<tr bgcolor="#<%=session("COLTABBG")%>" class="Estilo13">
			<td width="18%">MANDANTE</td>
			<td width="18%">ASIGNACION</td>
			<% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
				<td width="18%">EJECUTIVO</td>
			<% End If%>
			<td width="18%">F.INICIO</td>
			<td width="18%">F.TERMINO</td>
			<td width="10%">&nbsp</td>
		</tr>
		<tr>
			<td>
			<select name="CB_CLIENTE" onChange="refrescar();">
				<%
				abrirscg()
				ssql="SELECT COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE COD_CLIENTE = '" & intCliente & "' ORDER BY RAZON_SOCIAL"
				set rsCLI= Conn.execute(ssql)
				if not rsCLI.eof then
					do until rsCLI.eof%>
					<option value="<%=rsCLI("COD_CLIENTE")%>"<%if cint(cliente)=rsCLI("COD_CLIENTE") then response.Write("Selected") End If%>><%=rsCLI("RAZON_SOCIAL")%></option>
				<%
					rsCLI.movenext
					loop
					end if
					rsCLI.close
					set rsCLI=nothing
					cerrarscg()
				%>
        	</select>


			</td>
			<td>
				<select name="CB_REMESA">
					<option value="T">TODAS</option>
					<%
					AbrirSCG()
						strSql="SELECT * FROM REMESA WHERE COD_REMESA >= 100 and COD_CLIENTE = '" & intCliente & "'"
						set rsRemesa=Conn.execute(strSql)
						Do While not rsRemesa.eof
						If Trim(intCodRemesa)=Trim(rsRemesa("COD_REMESA")) Then strSelRem = "SELECTED" Else strSelRem = ""
						%>
						<option value="<%=rsRemesa("COD_REMESA")%>" <%=strSelRem%>> <%=rsRemesa("COD_REMESA") & " - " & rsRemesa("FECHA_ASIGNACION")%></option>
						<%
						rsRemesa.movenext
						Loop
						rsRemesa.close
						set rsRemesa=nothing
					CerrarSCG()
					''Response.End
					%>
				</select>
			</td>
			<% If TraeSiNo(session("perfil_adm")) = "Si" or TraeSiNo(session("perfil_sup")) = "Si" Then %>
				<td>
					<select name="CB_COBRADOR">
						<option value="T">TODOS</option>
						<%
						abrirscg()
						ssql="SELECT ID_USUARIO,LOGIN FROM USUARIO WHERE PERFIL_COB = 1 AND ACTIVO = 1"
						set rsTemp= Conn.execute(ssql)
						if not rsTemp.eof then
							do until rsTemp.eof%>
							<option value="<%=rsTemp("ID_USUARIO")%>"<%if Trim(intCodUsuario)=Trim(rsTemp("ID_USUARIO")) then response.Write("Selected") End If%>><%=rsTemp("LOGIN")%></option>
							<%
							rsTemp.movenext
							loop
						end if
						rsTemp.close
						set rsTemp=nothing
						cerrarscg()
						%>
					</select>
				</td>
			<% End If%>
			<td>
				<input name="inicio" type="text" id="inicio" value="<%=inicio%>" size="10" maxlength="10">
			 	<a href="javascript:showCal('Calendar7');"><img src="../Imagenes/calendario.gif" border="0">
			 </td>
			<td>
				<input name="termino" type="text" id="termino" value="<%=termino%>" size="10" maxlength="10">
				<a href="javascript:showCal('Calendar6');"><img src="../Imagenes/calendario.gif" border="0"></a>
			 </td>
			<td>
				<input type="button" name="Submit" value="Aceptar" onClick="envia();">
			</td>
		</tr>
		<tr>
			<td colspan"6">Tipo :

			<%
				If Trim(intOrigen) = "T" Then strSelTodos = "SELECTED"
				If Trim(intOrigen) = "R" Then strSelRetiro = "SELECTED"
				If Trim(intOrigen) = "S" Then strSelRetiroR = "SELECTED"
				If Trim(intOrigen) = "C" Then strSelCastigo = "SELECTED"
			%>
				<select name="CB_ORIGEN">
					<option value="T" <%=strSelTodos%>>TODOS</option>
					<option value="R" <%=strSelRetiro%>>RETIROS POR CLIENTE</option>
					<option value="S" <%=strSelRetiroR%>>RETIROS POR RESOLUCION</option>
					<option value="C" <%=strSelCastigo%>>CASTIGOS</option>
				</select>
			</td>
		</tr>
    </table>
</form>
    <%

		If Trim(intCliente) <> "" and Trim(intCodRemesa) <> "" then
		abrirscg()

%>


<table width="100%" border="0">
  <tr bgcolor="#<%=session("COLTABBG")%>">
  		<td><span class="Estilo37">FECHA:</span></td>
  		<td><span class="Estilo37">CASOS</span></td>
  		<td><span class="Estilo37">MONTO</span></td>
  		<td><span class="Estilo37">DOCS</span></td>
  		<td><span class="Estilo37">ACUM CASOS</span></td>
  		<td><span class="Estilo37">ACUM MONTO</span></td>
  		<td><span class="Estilo37">ACUM DOCS</span></td>
  	</tr>
    <%

	strSql="SELECT IsNull(datediff(d,'" & inicio & "' , '" & termino & "'),0) + 1 as DIAS"

	intFecha = inicio
	set rsFechas=Conn.execute(strSql)
	If Not rsFechas.eof Then
		intDias = rsFechas("DIAS")
		For I = 1 To intDias

			strSql = "SELECT COUNT(DISTINCT(RUT_DEUDOR)) as RUT, IsNull(COUNT(NRO_DOC),0) as FOLIO, IsNull(SUM(VALOR_CUOTA),0) as MONTO FROM CUOTA WHERE COD_CLIENTE = '" & intCliente & "'"
			If Trim(intCodRemesa) <> "T" Then
				strSql = strSql & " AND COD_REMESA = " & intCodRemesa
			End If

			'strSql = strSql & " AND SALDO = 0 AND CONVERT(VARCHAR(10),FECHA_ESTADO,103) = '" & intFecha & "' AND ESTADO_DEUDA IN "

			strSql = strSql & " AND CONVERT(VARCHAR(10),FECHA_ESTADO,103) = CAST('"&intFecha&"' AS DATETIME) AND ESTADO_DEUDA IN "

			If Trim(intOrigen) = "T" Then
				strSql = strSql & " (2,5,6)"
			End if
			If Trim(intOrigen) = "R" Then
				strSql = strSql & " (2)"
			End if
			If Trim(intOrigen) = "S" Then
				strSql = strSql & " (5)"
			End if
			If Trim(intOrigen) = "C" Then
				strSql = strSql & " (6)"
			End if

			set rsTemp= Conn.execute(strSql)
			If not rsTemp.eof then
				intCasos = rsTemp("RUT")
				intDocs = rsTemp("FOLIO")
				intMonto = rsTemp("MONTO")
			Else
				intCasos = 0
				intDocs = 0
				intMonto = 0
			End if
			rsTemp.close
			set rsTemp=nothing

			intAcumCasos = intAcumCasos + intCasos
			intAcumDocs = intAcumDocs + intDocs
			intAcumMonto = intAcumMonto + intMonto

			If intCasos <> 0 Then
				intMuestraCasos = intAcumCasos
				intMuestraDocs = intAcumDocs
				intMuestraMonto = intAcumMonto
			Else
				intMuestraCasos = 0
				intMuestraDocs = 0
				intMuestraMonto = 0
			End if

		%>

		<tr>
			<TD WIDTH="14%" ALIGN="RIGHT">
				<A HREF="detalle_retiros.asp?intFechaIni=<%=intFechaIni%>&intFechaFin=<%=intFechaFin%>&intFecha=<%=intFecha%>&intCliente=<%=intCliente%>&intOrigen=<%=intOrigen%>&intCodRemesa=<%=intCodRemesa%>&intCodUsuario=<%=intCodUsuario%>">
					<%=intFecha%>
				</A>
			</td>
			<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intCasos,0)%></td>
			<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intMonto,0)%></td>
			<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intDocs,0)%></td>
			<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intMuestraCasos,0)%></td>
			<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intMuestraMonto,0)%></td>
			<TD WIDTH="14%" ALIGN="RIGHT"><%=FN(intMuestraDocs,0)%></td>
		</tr>
		<%

			strSql = "SELECT DATEADD(day, 1, '" & intFecha & "') AS fecha"
			set rsTemp = Conn.execute(strSql)
			If Not rsTemp.eof Then
				intFecha = rsTemp("fecha")
			End if
		Next
	End If
	rsFechas.close
	set rsFechas=nothing
	''Response.End
	%>

	  <tr bgcolor="#<%=session("COLTABBG")%>">
	  		<td ALIGN="RIGHT" bgcolor="#FFFFFF">
	  			<span class="Estilo37">
					<A HREF="detalle_retiros.asp?intFechaIni=<%=intFechaIni%>&intFechaFin=<%=intFechaFin%>&intFecha=&intCliente=<%=intCliente%>&intOrigen=<%=intOrigen%>&intCodRemesa=<%=intCodRemesa%>&intCodUsuario=<%=intCodUsuario%>">
						TOTALES
					</A>
	  			</span>
	  		</td>
	  		<td ALIGN="RIGHT"><span class="Estilo37"><%=FN(intAcumCasos,0)%></span></td>
			<td ALIGN="RIGHT"><span class="Estilo37"><%=FN(intAcumMonto,0)%></span></td>
	  		<td ALIGN="RIGHT"><span class="Estilo37"><%=FN(intAcumDocs,0)%></span></td>
	  		<td ALIGN="RIGHT"><span class="Estilo37"><%=FN(intAcumCasos,0)%></span></td>
	  		<td ALIGN="RIGHT"><span class="Estilo37"><%=FN(intAcumMonto,0)%></span></td>
	  		<td ALIGN="RIGHT"><span class="Estilo37"><%=FN(intAcumDocs,0)%></span></td>
  	</tr>

    <%	cerrarscg()
end if %>
</table>

	  </td>
  </tr>
</table>






<script language="JavaScript1.2">
function envia(){
		if (datos.CB_CLIENTE.value=='0'){
			alert('DEBE SELECCIONAR UN CLIENTE');
		}else if(datos.inicio.value==''){
			alert('DEBE SELECCIONAR FECHA DE INICIO');
		}else if(datos.termino.value==''){
			alert('DEBES SELECCIONAR FECHA DE TERMINO');
		}else{
		//datos.action='cargando.asp';
		datos.action='informe_retiros.asp';
		datos.submit();
	}
}


function refrescar(){
		if (datos.CB_CLIENTE.value=='0'){
			alert('DEBE SELECCIONAR UN CLIENTE');
		}else
		{
		datos.action='informe_retiros.asp';
		datos.submit();
	}
}


function enviaexcel(){
if (datos.CB_CLIENTE.value=='0'){
alert('DEBE SELECCIONAR UN CLIENTE');
}else if(datos.inicio.value==''){
alert('DEBE SELECCIONAR FECHA DE INICIO');
}else if(datos.termino.value==''){
alert('DEBES SELECCIONAR FECHA DE TERMINO');
}else{
datos.action='informe_retiros_xls.asp';
datos.submit();
}
}



</script>
