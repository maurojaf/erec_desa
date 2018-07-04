<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"  LCID = 1034%>
<!DOCTYPE html>
<html lang="es">
<HEAD>
   
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
   
	<!--#include file="sesion.asp"-->
	<!--#include file="arch_utils.asp"-->
	<!--#include file="../lib/comunes/rutinas/funciones.inc" -->
	<!--#include file="../lib/comunes/rutinas/TraeCampo.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasSCG.inc" -->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->
	<!--#include file="../lib/lib.asp"-->
	<!--#include file="../lib/comunes/rutinas/chkFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/sondigitos.inc"-->
	<!--#include file="../lib/comunes/rutinas/formatoFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/validarFecha.inc"-->
	<!--#include file="../lib/comunes/rutinas/diasEnMes.inc"-->
	<!--#include file="../lib/asp/comunes/general/rutinasBooleano.inc"-->


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
    intCod_Cliente =  request("CB_CLIENTE")
	
	if intCod_Cliente ="" then  intCod_Cliente = session("ses_codcli")       

	if Trim(inicio) = "" Then
	
		inicio = DateAdd("d", -1, TraeFechaActual(Conn))
		
	End If
	
	if Trim(termino) = "" Then
	
		termino = DateAdd("d", -1, TraeFechaActual(Conn))
		
	End If

	Hoy = TraeFechaActual(Conn)
	
	int_inicio 			= split(inicio, "/")(2) & split(inicio, "/")(1) & split(inicio, "/")(0)
	int_termino 		= split(termino, "/")(2) & split(termino, "/")(1) & split(termino, "/")(0)
    intCod_Cliente		= request("CB_CLIENTE")
%>
	<link href="../css/style_multi_select.css" rel="stylesheet"> 
    <link href="../css/normalize.css" rel="stylesheet">	
	<link rel="stylesheet" type="text/css" href="../Componentes/jquery-1.9.2/css/start/jquery-ui-1.9.2.custom.css">
	<link href="../css/style_generales_sistema.css" rel="stylesheet"> 
	<link href="../Componentes/prettyLoader/css/prettyLoader.css" rel="stylesheet">
    <link href="../css/jquery.alerts.css" rel="stylesheet"> 
	<script src="../Componentes/jquery-1.9.2/js/jquery-1.8.3.js"></script>
	<script src="../Componentes/jquery-1.9.2/js/jquery-ui-1.9.2.custom.js"></script>
	<script src="../Componentes/prettyLoader/js/jquery.prettyLoader.js"></script> 
	<script src="../Componentes/jquery.alerts.mod.js"></script>
    
    <script src="../Componentes/jquery.multiselect.js"></script>
    

	<script language="JavaScript " type="text/JavaScript">

	    $(document).ready(function () {
		
			$(document).tooltip();
	        $.prettyLoader();
			$('#inicio').datepicker({ changeMonth: true, changeYear: true, dateFormat: 'dd/mm/yy' });
	        $('#termino').datepicker({ changeMonth: true, changeYear: true, dateFormat: 'dd/mm/yy' });
			
	    });
		
		function envia() {
		
			$.prettyLoader.show();
			document.datos.action = 'InformeRecupero.asp?resp=si';
			document.datos.submit();
			
		}

		function exportar() {
		
			var Cod_Cliente = $('#CB_CLIENTE').val();
			var inicio = $("#inicio").val();
			var termino = $("#termino").val();
			var pagina = 'ExporteInformeRecupero.asp?Cod_Cliente=' + Cod_Cliente + '&inicio=' + inicio + '&termino=' + termino;
			
			window.open(pagina, 'window', 'params');
			
		}
	
    </script>

	<style type="text/css">
        .hiddencol {
		
            display:none;
			
        }
		
        .span_aviso_rojo {
		
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
	</style>
</head>

<body leftmargin="0" rightmargin="0" marginwidth="0" topmargin="0" marginheight="0">
<form name="datos" method="post">
<div class="titulo_informe">INFORME RECUPERO</div>
<br>
&nbsp;<table width="90%" class="estilo_columnas" align="center">
		<thead>
	      <tr height="20" >
			<td>CLIENTE</td>
			<td class="style1">FECHA INICIO</td>
			<td class="style1">FECHA TERMINO</td>
			<td >&nbsp;</td>
		  </tr>
		</thead>
			
		  <tr >
         
        
			<td>
		
                <select name="CB_CLIENTE" id="CB_CLIENTE">
				<%
				ssql="SELECT  COD_CLIENTE,RAZON_SOCIAL FROM CLIENTE WHERE ACTIVO = 1 AND COD_CLIENTE IN (SELECT COD_CLIENTE FROM USUARIO_CLIENTE WHERE ID_USUARIO = " & session("session_idusuario") & ") ORDER BY COD_CLIENTE,RAZON_SOCIAL ASC"
				set rsTemp= Conn.execute(ssql)
				if not rsTemp.eof then
			    
                do until rsTemp.eof
                   cadena = intCOD_CLIENTE
                   y =  rsTemp("COD_CLIENTE") 
                    strChecked = ""
                    a=Split(cadena,",")
                    for each x in a
                        IF  trim(x) = trim(y)  THEN
                            strChecked = " selected "
                        END IF 
                    next
                    
                    %>
					<option  <%=strChecked%> value="<%=rsTemp("COD_CLIENTE")%>" ><%=rsTemp("RAZON_SOCIAL")%></option>
					<%

                    
					rsTemp.movenext
					loop
				end if
				rsTemp.close
				set rsTemp=nothing
				%>
				</select>
                
                
                </td>
			
			<td class="style1"><input name="inicio" type="text" id="inicio" value="<%=inicio%>" readonly="true" size="10" maxlength="10">
			</td>
			
			<td class="style1"><input name="termino" type="text" id="termino" value="<%=termino%>" readonly="true" size="10" maxlength="10">
			</td>

				<td align="right">
				 <input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onclick= "envia();">
				 <input type="Button" class="fondo_boton_100" name="Exportar" value="Exportar" onclick= "exportar();">
				 </td>
			</tr>
    </table>
	
	<%
	
		if resp = "si" then
		
			strSql = "EXEC dbo.uspInformeRecuperoSelect @CodigoMandante = '" & intCod_Cliente & "', @FechaInicio = " & int_inicio & ", @FechaTermino = " & int_termino
			
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
	
	<%
	
		end if
	
	%>
	
	<br>
	
</form>

</body>
</html>
<%cerrarscg()%>