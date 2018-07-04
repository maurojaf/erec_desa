<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"  LCID = 1034%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
   
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
	<META HTTP-EQUIV="Cache-Control" CONTENT ="no-cache">
    <meta charset="utf-8">
	
   	<!--#include file="arch_utils.asp"-->
	<!--#include file="sesion_inicio.asp"-->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/RutinasVarias.inc" -->


    <%
		Id_Usuario = session("session_idusuario")
		
		resp = request("resp")
	%>



<%
	Response.CodePage=65001
	Response.charset ="utf-8"
	AbrirSCG()
	
	inicio 			= request("inicio")
	termino 		= request("termino")
    intCod_Cliente =  request("Cod_Cliente")
	
	if intCod_Cliente ="" then  intCod_Cliente = session("ses_codcli")       
	
	int_inicio 			= split(inicio, "/")(2) & split(inicio, "/")(1) & split(inicio, "/")(0)
	int_termino 		= split(termino, "/")(2) & split(termino, "/")(1) & split(termino, "/")(0)
    intCod_Cliente =  request("Cod_Cliente")
	campana				= Replace(Replace(Replace(Request("campana"), "[", ""), "]", ""), """", "''")
	campana_cliente		= Replace(Replace(Replace(Request("campana_cliente"), "[", ""), "]", ""), """", "''")
	
	fileName ="InformeGestiones" & intCod_Cliente  & ".xls"
	Response.AddHeader "content-disposition", "attachment; filename=" & fileName
	Response.ContentType = "application/octet-stream"
	Response.Flush()
%>
	<style type="text/css">
        .hiddencol {
		
            display:none;
			
        }
		
        .span_aviso_rojo{
		
			color:#FE2E2E;
			font-size:12px;
			width:1px;
			margin-right:2cm;
			
		}
		
		.abrir_cerrar {
		
			cursor: pointer;
			display: inline-block;
			float: right;
			position: relative;
			margin-right:2cm;
		
		}
	
		.intercalado {
		
			width: 90%;
			border: 1px solid #ccc;
			margin: 0 auto;
		
		}
		
		.intercalado thead tr {
		
			background-color: #999999;
			height: 22px;
			color: #fff;
			font-weight: bold;
		
		}
		
		.intercalado thead tr td {
		
			border: 1px solid #F4F4F4;
		
		}
		
		.intercalado tbody tr:nth-child(2n + 1) {
		
			background: #FFF;
			height: 22px;
		
		}
		
		.intercalado tbody tr:nth-child(2n) {
		
			background: #F0F0F0;
			height: 22px;
		
		}
	
	</style>
</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<form name="datos" method="post">
<br>

	<%
	
			strSql = "EXEC dbo.uspInformeGestionesSelect @CodigoMandante = '" & intCod_Cliente & "', @FechaInicio = " & int_inicio & ", @FechaTermino = " & int_termino & ", @Campana = '" & campana & "', @CampanaCliente = '" & campana_cliente & "'"
			
			set rsDet = Conn.execute(strSql)
	
	%>
	
			<table border="0" class="intercalado" align="center">
				<thead>
					<tr>
						<%
					
						For Each objField in rsDet.Fields
							Response.Write "<td>" & objField.Name & "</td>"
						Next
						
						%>
					</tr>
				</thead>
				<tbody>
						<%
					
							While Not rsDet.EOF
								Response.Write "<tr>"
								For Each objField in rsDet.Fields
									Response.Write "<td>" & rsDet(objField.Name) & "</td>"
								Next
								rsDet.MoveNext
								Response.Write "</tr>"
							Wend
						%>
				</tbody>
			</table>
	
	<br>
	
</form>

</body>
</html>
<%cerrarscg()%>