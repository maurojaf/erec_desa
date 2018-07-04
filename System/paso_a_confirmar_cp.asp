<% @LCID = 1034 %>

<link href="../css/style.css" rel="stylesheet" type="text/css">

<%
id_gestion = request("id_gestion")
intCodGestConcat = request("intCodGestConcat")
rut = request("rut")
cliente = request("cliente")
dtmFecCompGest = request("dtmFecCompGest")

%>

<script languaje="javascript">
	//top.opener.document.Contenido.location = top.opener.document.location;
	//window.opener.parent.top.location.reload();
	window.parent.document.location.href='confirmar_cp.asp?id_gestion=<%=id_gestion%>&rut=<%=rut%>&cliente=<%=cliente%>&dtmFecCompGest=<%=dtmFecCompGest%>&intCodGestConcat=<%=intCodGestConcat%>&origen=1';
	//window.parent.document.location.href='confirmar_cp.asp';
	//window.parent.document.location.reload();
	//top.opener.document.location.href=theURL;
</script>














