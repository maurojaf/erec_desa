<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"  LCID = 1034%>
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
	Response.CodePage=65001
	Response.charset ="utf-8"
	AbrirSCG()
	
	inicio 			= request("inicio")
	termino 		= request("termino")
    intCod_Cliente =  request("CB_CLIENTE")

	if request("GetCampanas") = "true" then
	
		%>
		
		<select name="campana" id="campana" multiple style="font-size:12px;">
			<option value="">-- Todas --</option>
			<%
				ssql = "SELECT ID_CAMPANA, COD_CLIENTE, COD_REMESA, FECHA_CREACION, NOMBRE, DESCRIPCION, FECHA_INICIO, FECHA_TERMINO, OBSERVACION, FECHA_MODIFICACION, ID_USUARIO, ID_USUARIO_MODIFICA FROM dbo.CAMPANA WHERE COD_CLIENTE = '" & intCOD_CLIENTE & "'"
				
				set rsTemp= Conn.execute(ssql)
				
				if not rsTemp.eof then
			    
					do until rsTemp.eof
					
						%>
						<option value="<%=rsTemp("ID_CAMPANA")%>"><%=rsTemp("NOMBRE")%></option>
						<%
				
						rsTemp.movenext
						
					loop
					
				end if
				
				rsTemp.close
				
				set rsTemp=nothing
			%>
		</select>
		
		<script language="JavaScript " type="text/JavaScript">

	    $(document).ready(function () {
		
			$("#campana").multiselect();
			
		});
		
		</script>
		
		<%
		
		cerrarscg()
	
		Response.End
	
	end if
	
	if request("GetCampanasCliente") = "true" then
	
		%>
		
		<select name="campana_cliente" id="campana_cliente" multiple style="font-size:12px;">
			<option value="">-- Todas --</option>
			<%
				ssql = "SELECT ID_CAMPANA_CLIENTE ,COD_CLIENTE ,FECHA_CREACION ,NOMBRE ,DESCRIPCION ,FECHA_INICIO ,FECHA_TERMINO ,OBSERVACION ,FECHA_MODIFICACION ,ID_USUARIO ,ID_USUARIO_MODIFICA FROM dbo.CAMPANA_CLIENTE WHERE COD_CLIENTE = '" & intCOD_CLIENTE & "'"
				
				set rsTemp= Conn.execute(ssql)
				
				if not rsTemp.eof then
			    
					do until rsTemp.eof
					
						%>
						<option value="<%=rsTemp("ID_CAMPANA_CLIENTE")%>"><%=rsTemp("NOMBRE")%></option>
						<%
				
						rsTemp.movenext
						
					loop
					
				end if
				
				rsTemp.close
				
				set rsTemp=nothing
			%>
		</select>
		
		<script language="JavaScript " type="text/JavaScript">

	    $(document).ready(function () {
		
			$("#campana_cliente").multiselect();
			
		});
		
		</script>
		
		<%
		
		cerrarscg()
	
		Response.End
	
	end if

%>

<!DOCTYPE html>
<html lang="es">
<HEAD>
   
    <meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
    <meta charset="utf-8">
   <!--#include file="sesion.asp"-->
	<%
		Id_Usuario = session("session_idusuario")
		
		resp = request("resp")
	%>



<%
	if intCod_Cliente ="" then  intCod_Cliente = session("ses_codcli")       

	if Trim(inicio) = "" Then
	
		strMesActual = Month(TraeFechaActual(Conn))
		strAnoActual = Cdbl(Year(TraeFechaActual(Conn)))
		If strMesActual = 1 Then strAnoActual = strAnoActual - 1
		If strMesActual = 1 Then strMesActual = 12
		strMesActual = strMesActual - 1
		if Len(strMesActual) = 1 Then strMesActual = "0" & strMesActual
		If Trim(inicio) = "" Then inicio = "01/" & strMesActual & "/" & strAnoActual
		
	End If
	
	if Trim(termino) = "" Then
	
		termino = TraeFechaActual(Conn)
		
	End If

	Hoy = TraeFechaActual(Conn)
	
	int_inicio 			= split(inicio, "/")(2) & split(inicio, "/")(1) & split(inicio, "/")(0)
	int_termino 		= split(termino, "/")(2) & split(termino, "/")(1) & split(termino, "/")(0)
    intCod_Cliente		= request("CB_CLIENTE")
	campana				= Replace(Replace(Replace(Request("campana"), "[", ""), "]", ""), """", "''")
	campana_cliente		= Replace(Replace(Replace(Request("campana_cliente"), "[", ""), "]", ""), """", "''")
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
			$("#campana").multiselect();
			$("#campana_cliente").multiselect();
			
			$('#CB_CLIENTE').change(function(){
			
				$.ajax({

                    url: 'InformeGestiones.asp',

                    type: 'POST',

                    data: {

                        GetCampanas: 'true',
						
						CB_CLIENTE: $("#CB_CLIENTE").val()

                    },

                    success: function (data) {

                        $("#TdCampana").html(data);

                    }

                });
				
				$.ajax({

                    url: 'InformeGestiones.asp',

                    type: 'POST',

                    data: {

                        GetCampanasCliente: 'true',
						
						CB_CLIENTE: $("#CB_CLIENTE").val()

                    },

                    success: function (data) {

                        $("#TdCampanaCliente").html(data);

                    }

                });
			
			});
			
			$('#CB_CLIENTE').change();
			
	    });
		
		function envia() {
		
			$.prettyLoader.show();
			
			var campana = [];

			$('select[id="campana"] option:checked').each(function () {
				campana.push($(this).val())
			})
			
			var campana_cliente = [];

			$('select[id="campana_cliente"] option:checked').each(function () {
				campana_cliente.push($(this).val())
			})
			
			document.datos.action = 'InformeGestiones.asp?resp=si&campana=' + encodeURIComponent(JSON.stringify(campana)) + '&campana_cliente=' + encodeURIComponent(JSON.stringify(campana_cliente));
			document.datos.submit();
			
		}

		function exportar() {
		
			var Cod_Cliente = $('#CB_CLIENTE').val();
			var inicio = $("#inicio").val();
			var termino = $("#termino").val();
			
			var campana = [];

			$('select[id="campana"] option:checked').each(function () {
				campana.push($(this).val())
			})
			
			var campana_cliente = [];

			$('select[id="campana_cliente"] option:checked').each(function () {
				campana_cliente.push($(this).val())
			})
			
			var pagina = 'ExporteInformeGestiones.asp?Cod_Cliente=' + Cod_Cliente + '&inicio=' + inicio + '&termino=' + termino + '&campana=' + encodeURIComponent(JSON.stringify(campana)) + '&campana_cliente=' + encodeURIComponent(JSON.stringify(campana_cliente));
			
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
<div class="titulo_informe">INFORME GESTIONES</div>
<br>
&nbsp;<table width="90%" class="estilo_columnas" align="center">
		<thead>
	      <tr height="20" >
			<td>CLIENTE</td>
			<td class="style1">FECHA INICIO</td>
			<td class="style1">FECHA TERMINO</td>
			<td class="style1">CAMPA&Ntilde;A</td>
			<td class="style1">CAMPA&Ntilde;A CLIENTE</td>
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
			
			<td class="style1" id="TdCampana">
				<select name="campana" id="campana" multiple style="font-size:12px;">
					<option value="">-- Todas --</option>
				</select>
			</td>
			
			<td class="style1" id="TdCampanaCliente">
				<select name="campana_cliente" id="campana_cliente" multiple style="font-size:12px;">
					<option value="">-- Todas --</option>
				</select>
			</td>

				<td align="right">
				 <input type="Button" class="fondo_boton_100" name="Submit" value="Ver" onclick= "envia();">
				 <input type="Button" class="fondo_boton_100" name="Exportar" value="Exportar" onclick= "exportar();">
				 </td>
			</tr>
    </table>
	
	<%
	
		if resp = "si" then
		
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
	
	<%
	
		end if
	
	%>
	
	<br>
	
</form>

</body>
</html>
<%cerrarscg()%>