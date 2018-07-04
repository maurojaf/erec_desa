<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
    <link href="../css/normalize.css" rel="stylesheet">
	<!--#include file="sesion.asp"-->    
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/asp/comunes/general/SoloNumeros.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<link href="../css/style_generales_sistema.css" rel="stylesheet">
</head>
<body>
<form action="" method="post" name="datos">
<input name="rut" type="hidden" id="rut" value="rut">
	  
<%
Response.CodePage=65001
Response.charset ="utf-8"

intCliente=request("intCliente")
intCodUsuario=session("session_idusuario")
intTipoTramo=request("intTipoTramo")
intTipoInforme=request("intTipoInforme")

If request("TX_ANEXO_1") = "" then intNvoTramo1 = 0 else intNvoTramo1 = request("TX_ANEXO_1")
If request("TX_ANEXO_2") = "" then intNvoTramo2 = 0 else intNvoTramo2 = request("TX_ANEXO_2")
If request("TX_ANEXO_3") = "" then intNvoTramo3 = 0 else intNvoTramo3 = request("TX_ANEXO_3")
If request("TX_ANEXO_4") = "" then intNvoTramo4 = 0 else intNvoTramo4 = request("TX_ANEXO_4")

strProcesar = request("strProcesar")

'Response.write "intTipoInforme = " & intTipoInforme

If strProcesar="S" THEN

	AbrirSCG()
	
		strSql = "SELECT * FROM TRAMOS_DEUDA" 
		strSql = strSql & " WHERE COD_CLIENTE= '" & intCliente & "' AND TIPO_TRAMOS = " & intTipoTramo
		strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

		'Response.write "strSql = " & strSql
		set rsIns=Conn.execute(strSql)
			if not rsIns.eof then
			
			AbrirSCG2()

			strSql = "UPDATE TRAMOS_DEUDA"
			strSql = strSql & " SET TRAMO = " & intNvoTramo1 & ", USUARIO_ESTADO = " & intCodUsuario
			strSql = strSql & " WHERE COD_CLIENTE = '" & intCliente & "' AND ORDEN_TRAMO=1 AND TIPO_TRAMOS = " & intTipoTramo
			
			set rsIns1=Conn2.execute(strSql)

						strSql = "UPDATE TRAMOS_DEUDA"
			strSql = strSql & " SET TRAMO = " & intNvoTramo2 & ", USUARIO_ESTADO = " & intCodUsuario
			strSql = strSql & " WHERE COD_CLIENTE = '" & intCliente & "' AND ORDEN_TRAMO=2 AND TIPO_TRAMOS = " & intTipoTramo
			
			set rsIns1=Conn2.execute(strSql)
			
			strSql = "UPDATE TRAMOS_DEUDA"
			strSql = strSql & " SET TRAMO = " & intNvoTramo3 & ", USUARIO_ESTADO = " & intCodUsuario
			strSql = strSql & " WHERE COD_CLIENTE = '" & intCliente & "' AND ORDEN_TRAMO=3 AND TIPO_TRAMOS = " & intTipoTramo
			
			set rsIns1=Conn2.execute(strSql)

			strSql = "UPDATE TRAMOS_DEUDA"
			strSql = strSql & " SET TRAMO = " & intNvoTramo4 & ", USUARIO_ESTADO = " & intCodUsuario
			strSql = strSql & " WHERE COD_CLIENTE = '" & intCliente & "' AND ORDEN_TRAMO=4 AND TIPO_TRAMOS = " & intTipoTramo
			
			set rsIns1=Conn2.execute(strSql)
	
			CerrarSCG2()

			else
			
			AbrirSCG2()
			
			strSql = "INSERT INTO TRAMOS_DEUDA" 
			strSql = strSql & " VALUES ('" & intCliente & "',1," & intTipoTramo & "," & intNvoTramo1 & "," & intCodUsuario & ",GETDATE())"	
			
			set rsIns1=Conn2.execute(strSql)

			strSql = "INSERT INTO TRAMOS_DEUDA" 
			strSql = strSql & " VALUES ('" & intCliente & "',2," & intTipoTramo & "," & intNvoTramo2 & "," & intCodUsuario & ",GETDATE())"	
			set rsIns1=Conn2.execute(strSql)

			strSql = "INSERT INTO TRAMOS_DEUDA" 
			strSql = strSql & " VALUES ('" & intCliente & "',3," & intTipoTramo & "," & intNvoTramo3 & "," & intCodUsuario & ",GETDATE())"	
			set rsIns1=Conn2.execute(strSql)

			strSql = "INSERT INTO TRAMOS_DEUDA" 
			strSql = strSql & " VALUES ('" & intCliente & "',4," & intTipoTramo & "," & intNvoTramo4 & "," & intCodUsuario & ",GETDATE())"	
			set rsIns1=Conn2.execute(strSql)			
			
			CerrarSCG2()
			
			End If
		rsIns.close
		set rsIns=nothing	
			
	CerrarSCG()

	%>
		<SCRIPT>
			//alert(<%=intTipoInforme%>);
			location.href='informe_cartera.asp?intTipoInforme=<%=intTipoInforme%>';
		</SCRIPT>	
	<%
					
End If
		
		AbrirSCG()
		
			strSql = "SELECT * FROM TRAMOS_DEUDA" 
			strSql = strSql & " WHERE COD_CLIENTE= '" & intCliente & "' AND TIPO_TRAMOS = " & intTipoTramo 
			strSql = strSql & " ORDER BY ORDEN_TRAMO ASC"

			'Response.write "strSql = " & strSql
				set rsDet=Conn.execute(strSql)

				if not rsDet.eof then
					intReg = 0
					do while not rsDet.eof
					
					intOrdenTramo=rsDet("ORDEN_TRAMO")
					intTramo=rsDet("TRAMO")
					
						If intOrdenTramo = 1 then
						intTramo_1= intTramo
						End If

						If intOrdenTramo = 2 then
						intTramo_2= intTramo
						End If
	
						If intOrdenTramo = 3 then
						intTramo_3= intTramo
						End If

						If intOrdenTramo = 4 then
						intTramo_4= intTramo
						End If

					rsDet.movenext
					loop
				end if
				rsDet.close
				set rsDet=nothing
				
		CerrarSCG%>

<Table WIDTH="550" BORDER="1" CELLPADDING=0 CELLSPACING=0 ALIGN="CENTER">
			<TR>
			 <TD>
				<div class="titulo_informe">
				
				<%If intTipoTramo=1 then%>
				
					MANTENEDOR TRAMOS DEUDA CASO
					
				<%ElseIf intTipoTramo=2 then%>
				
					MANTENEDOR TRAMOS DIAS MORA CASO
				
				<%ElseIf intTipoTramo=3 then%>
				
					MANTENEDOR TRAMOS DIAS ASIGNACION CASO

				<%ElseIf intTipoTramo=4 then%>
				
					MANTENEDOR TRAMOS DOCUMENTOS CASO
				
				<%ElseIf intTipoTramo=5 then%>
				
					MANTENEDOR TRAMOS DEUDA DOCUMENTOS

				<%ElseIf intTipoTramo=6 then%>
				
					MANTENEDOR TRAMOS DIAS MORA DOCUMENTOS

				<%ElseIf intTipoTramo=7 then%>
				
					MANTENEDOR TRAMOS DIAS ASIGNACION DOCUMENTOS
					
				<%End If%>
				
				</div>
				<br>
		<table width="100%" border="0" CLASS="tabla1">
		
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Tramo 1</Font></td>
				<td class="td_t" align = "center"><input name="TX_ANEXO_1" type="text" value="<%=intTramo_1%>" size="10" maxlength="100"><td>
			</tr>
			
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Tramo 2</Font></td>
				<td class="td_t" align = "center"><input name="TX_ANEXO_2" type="text" value="<%=intTramo_2%>" size="10" maxlength="100"><td>
			</tr>

			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Tramo 3</Font></td>
				<td class="td_t" align = "center"><input name="TX_ANEXO_3" type="text" value="<%=intTramo_3%>" size="10" maxlength="100"><td>
			</tr>
			
			<tr BGCOLOR="#FFFFFF">
				<td class="hdr_i">Tramo 4</Font></td>
				<td class="td_t" align = "center"><input name="TX_ANEXO_4" type="text" value="<%=intTramo_4%>" size="10" maxlength="100"><td>
			</tr>
		
			<tr bordercolor="#FFFFFF">
				<td align="CENTER"><input name="Submit" type="button" class="fondo_boton_100" onClick="envia();" value="Guardar Cambios Realizados"></td>
				<td align="RIGHT"><input name="Volver" type="button" class="fondo_boton_100" onClick="history.back();" value="Volver"></td>
			</tr>
			
		</table>
</table>
</form>
</body>
</html>
<script type="text/javascript">

function envia()
{
	document.datos.action = "man_carga_tramos_informes.asp?strProcesar=S&intCliente=<%=intCliente%>&intTipoTramo=<%=intTipoTramo%>&intTipoInforme=<%=intTipoInforme%>";
	document.datos.submit();
}

</script>


