<SCRIPT LANGUAGE="VBScript" RUNAT="Server">
' Procedimiento:	vs_MostrarError
' Descripción:	Muestra el error de acceso a la BD junto a una información
'               adicional.
' Argumentos:	strInfoAdicional	String con información adicional que sea
'									más útil al usuario
'				Terminar			Indica si terminamos o no con el procesamiento 
'									del script (true o false)
' Retorna:		Nada
'
Sub vs_MostrarError (ByVal strInfoAdicional, ByVal Terminar)
	Response.Write "	<BR>&nbsp;<BR>"
	Response.Write "	<TABLE WIDTH=100% CELLSPACING=1 CELLPADDING=2 BORDER=0>"
	Response.Write "		<TR>"
	Response.Write "			<TD ALIGN=Left BGCOLOR=Silver><FONT SIZE=-1><B>&nbsp;&nbsp;Elemento</B></FONT></TD>"
	Response.Write "			<TD WIDTH=100% ALIGN=Left BGCOLOR=Silver><FONT SIZE=-1><B>Descripci&oacute;n</B></FONT></TD>"
	Response.Write "		</TR>"
	Response.Write "		<TR>"
	Response.Write "			<TD><FONT SIZE=-1><B>&nbsp;&nbsp;Source:</B></FONT></TD>"
	Response.Write "			<TD BGCOLOR=White><FONT SIZE=-1>" & Err.Source & "</TD>"
	Response.Write "		</TR>"
	Response.Write "		<TR>"
	Response.Write "			<TD NOWRAP><FONT SIZE=-1><B>&nbsp;&nbsp;N&uacute;mero de error:</B></FONT></TD>"
	Response.Write "			<TD BGCOLOR=White><FONT SIZE=-1>" & Err.Number & "</FONT></TD>"
	Response.Write "		</TR>"
	Response.Write "		<TR>"
	Response.Write "			<TD><FONT SIZE=-1><B>&nbsp;&nbsp;Descripci&oacute;n:</B></FONT></TD>"
	Response.Write "			<TD BGCOLOR=White><FONT SIZE=-1><B>" & Server.HTMLEncode( strInfoAdicional ) & "</B><BR>" &  Server.HTMLEncode( Err.Description ) & "</FONT></TD>"
	Response.Write "		</TR>"
	Response.Write "		<TR>"
	Response.Write "			<TD COLSPAN=2><HR></TD>"
	Response.Write "		</TR>"
	Response.Write "	</TABLE>"
	Response.Write "	<BR>&nbsp;<BR>"
	
	If Terminar = true Then
		Response.End
	End If
End Sub
</SCRIPT>