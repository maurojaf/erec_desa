<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" LCID = 1034%>
		
<!--#include file="../../arch_utils.asp"-->
<!--#include file="../../../lib/comunes/rutinas/funciones.inc" -->
		
<%
	IdCampo			=	request("IdCampo")
	RutDeudor		=	request("RutDeudor")
	CodigoCliente	=	request("CodigoCliente")
	
	ID_DOMINIO_CAMPO_TEXTO = 7
		
	AbrirSCG()
		
		StrSql="SELECT  FD.Fecha, "
		StrSql= StrSql & "	U.[LOGIN] AS Usuario, "
		StrSql= StrSql & "	CASE WHEN FD.IdDominioCampo = " & ID_DOMINIO_CAMPO_TEXTO & " THEN VCT.Texto ELSE DC.Valor END AS Valor, "
		StrSql= StrSql & "	AFD.Observaciones, "
		StrSql= StrSql & "	FD.IdDominioCampo "
		StrSql= StrSql & "	FROM    dbo.FichaDeudor FD "
		StrSql= StrSql & "	INNER JOIN dbo.USUARIO U "
		StrSql= StrSql & "	ON FD.IdUsuarioModificacion = U.ID_USUARIO "
		StrSql= StrSql & "	INNER JOIN dbo.DominioCampo DC ON FD.IdDominioCampo = DC.IdDominioCampo "
		StrSql= StrSql & "	LEFT JOIN ValorCampoTexto VCT ON FD.IdFichaDeudor = VCT.IdFichaDeudor "
		StrSql= StrSql & "	LEFT JOIN dbo.AtributoFichaDeudor AFD ON FD.IdFichaDeudor = AFD.IdFichaDeudor "
		StrSql= StrSql & "	WHERE FD.CodigoCliente = '" & CodigoCliente & "' "
		StrSql= StrSql & "	AND FD.RutDeudor = '" & RutDeudor & "' "
		StrSql= StrSql & "	AND FD.IdCampo = '" & IdCampo & "' "
		StrSql= StrSql & "	ORDER BY FD.Fecha DESC "
		
		set rsHistorialCambiosCampo = Conn.execute(StrSql)
%>
<table width="100%" border="1" class="HistorialCambios">
	<% if not rsHistorialCambiosCampo.Eof then %>
	<tr>
		<td class="Title">FECHA</td>
		<td class="Title">USUARIO</td>
		<td class="Title">INFORMACI&Oacute;N ANTERIOR</td>
		<td class="Title">OBSERVACIONES</td>
	</tr>
	<%
		i = 1
		
		if rsHistorialCambiosCampo("IdDominioCampo") = ID_DOMINIO_CAMPO_TEXTO then
			valorStyle = "class=""Texto"""
		else
			valorStyle = ""
		end if
		
		while not rsHistorialCambiosCampo.Eof
		
			if i mod 2 = 0 then
				clase = "AtributoGris"
			else
				clase = "Atributo"
			end if
			
			fecha = split(rsHistorialCambiosCampo("Fecha"), " ")(0)
			
			hora = split(rsHistorialCambiosCampo("Fecha"), " ")(1)
			
			horaSinSegundos = split(hora, ":")(0) & ":" & split(hora, ":")(1)
			
			fechaToShow = fecha & " " & horaSinSegundos
		%>
	<tr class="<%=clase %>">
		<td><%=fechaToShow %></td>
		<td><%=rsHistorialCambiosCampo("Usuario") %></td>
		<td <%=valorStyle %> ><%=rsHistorialCambiosCampo("Valor") %></td>
		<td><div class="Observaciones" >
			<%
				if rsHistorialCambiosCampo("Observaciones") <> "" then
				
					observaciones = rsHistorialCambiosCampo("Observaciones")
					
					if Len(observaciones) > 20 then
						observaciones = "<label title=""" & observaciones & """>" & Mid(observaciones, 1, 20) & "...</label>"
					end if
				
					Response.Write(observaciones)
				else
					Response.Write("SIN OBSERVACIONES")
				end if
			%>
			</div>
		</td>
	</tr>
		<%
		
			i = i + 1
			rsHistorialCambiosCampo.MoveNext		
		wend
	
	else %>
	<tr>
		<td class="Title">EL CAMPO DE LA FICHA NO REGISTRA HISTORIAL DE CAMBIOS.</td>
	<tr>
	<% end if
		
	CerrarSCG()
	%>
</table>